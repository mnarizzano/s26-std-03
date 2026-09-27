"""
pdf_exporter.py

Exporter producing the disassembly document in PDF format from the Guide
built by disassembly_loader. The loader already delivers the steps
linearized and ordered in guide.steps, one item per diamond (the
granularity required by URS 10.0), so this module only iterates over
guide.steps and lays the content out.

Location: src/main/pdf_exporter.py
Public entry point for the GUI:

    from pdf_exporter import export_to_pdf
    export_to_pdf(json_path, depth=spec, include_bom=bom_needed)

Self-contained: depends only on disassembly_loader and reportlab.

URS requirements covered:
  1.3  - bill of materials and required tools at the top of the document
  2.2  - validation warnings listed one by one with the node ids involved,
         so the operator sees them before starting
  4.1  - text and image for every step, action and component
  10.0 - one section per diamond (inherited from the Guide)
  12.0 - export to PDF
  19.0 - cap on the number of exported steps (max_groups)

Layout features: BOM table wrapping inside cells, colour-banded step
headers, highlighted safety instructions, clickable table of contents,
PDF bookmarks, page numbers and product name in the footer.
"""

import hashlib
from pathlib import Path
from typing import Optional

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.lib import colors
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame,
    Paragraph, Spacer, Table, TableStyle,
    ListFlowable, ListItem, Image, PageBreak,
)
from reportlab.platypus.tableofcontents import TableOfContents

from abc import ABC, abstractmethod
from dataclasses import dataclass

from disassembly_loader import Guide, Step, Severity, build_guide


@dataclass
class ExportOptions:
    """
    Rendering options for the PDF export. Disassembly depth and BOM
    inclusion are properties of the Guide, already fixed by build_guide();
    only options affecting the document layout belong here.

    Attributes:
        output_path: path of the file to generate.
        max_groups: maximum number of Steps to export (URS requirement 19.0).
        include_images: whether to embed action images (URS requirement 4.1).
        show_warnings: whether to list guide.warnings in the document.
    """
    output_path: Path
    max_groups: Optional[int] = None
    include_images: bool = True
    show_warnings: bool = True


class ExportMethod(ABC):
    """
    Contract every export format follows: one export() method taking a Guide
    and returning the Path written (URS requirement 11.0). Kept here so the
    module stays a single self-contained file.
    """

    format_name: str = "generic"
    file_extension: str = ""

    @abstractmethod
    def export(self, guide: Guide, options: ExportOptions) -> Path:
        """Generate the output file from the Guide. Returns the Path written."""
        raise NotImplementedError

    def validate_options(self, options: ExportOptions) -> None:
        """Optional hook for format-specific validity checks."""
        return None


# Document palette: single place to restyle the whole output.
_PALETTE = {
    "primary": colors.HexColor("#1f3a5f"),        # dark blue: step headers
    "primary_text": colors.white,
    "table_header": colors.HexColor("#333333"),
    "row_alt": colors.HexColor("#f2f2f2"),
    "outputs_bg": colors.HexColor("#e8f5e9"),     # light green: extracted parts
    "continues_bg": colors.HexColor("#eceff1"),   # light grey: continuation
    "safety_bg": colors.HexColor("#fff3cd"),      # light amber: safety notice
    "safety_text": colors.HexColor("#8a1f1f"),
    "muted_text": colors.HexColor("#555555"),
}

_SEVERITY_COLOR = {
    Severity.INFO: colors.HexColor("#2e6da4"),
    Severity.WARNING: colors.HexColor("#c77c02"),
    Severity.ERROR: colors.HexColor("#c9302c"),
}

# Keywords marking an instruction as a safety notice. Both English and
# Italian are matched, since source models may be written in either.
_SAFETY_KEYWORDS = (
    "warning", "caution", "danger", "precaution", "careful",
    "attenzione", "pericolo", "cautela",
)


def _is_safety_text(text: str) -> bool:
    """Return True when the text contains a safety keyword."""
    lowered = text.lower()
    return any(k in lowered for k in _SAFETY_KEYWORDS)


def _bookmark_key(text: str) -> str:
    """Build a stable unique key for bookmarks and TOC entries."""
    return "bm-" + hashlib.sha1(text.encode("utf-8")).hexdigest()[:16]


class _GuideDocTemplate(BaseDocTemplate):
    """
    Document template drawing the footer (product name and page number) and
    registering TOC entries and PDF bookmarks whenever a section or step
    heading is laid out.
    """

    def __init__(self, filename: str, product_name: str, **kwargs) -> None:
        super().__init__(filename, **kwargs)
        self._product_name = product_name
        frame = Frame(
            self.leftMargin, self.bottomMargin,
            self.width, self.height, id="main",
        )
        self.addPageTemplates([
            PageTemplate(id="main", frames=[frame], onPage=self._draw_footer)
        ])

    def _draw_footer(self, canvas, doc) -> None:
        """Draw the footer repeated on every page."""
        canvas.saveState()
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.grey)
        canvas.drawString(self.leftMargin, 1.1 * cm, self._product_name)
        canvas.drawRightString(A4[0] - self.rightMargin, 1.1 * cm, f"Page {doc.page}")
        canvas.setLineWidth(0.3)
        canvas.setStrokeColor(colors.lightgrey)
        canvas.line(self.leftMargin, 1.4 * cm, A4[0] - self.rightMargin, 1.4 * cm)
        canvas.restoreState()

    def afterFlowable(self, flowable) -> None:
        """
        Catch paragraphs styled SectionHeading (level 0) and StepTitle
        (level 1) to create the PDF bookmark, the reader outline entry and
        the clickable table-of-contents entry.
        """
        if not isinstance(flowable, Paragraph):
            return
        style_name = flowable.style.name
        if style_name not in ("SectionHeading", "StepTitle"):
            return
        text = flowable.getPlainText()
        level = 0 if style_name == "SectionHeading" else 1
        key = _bookmark_key(text)
        self.canv.bookmarkPage(key)
        self.canv.addOutlineEntry(text, key, level=level, closed=False)
        # The 4-tuple form (with key) makes the TOC entry clickable.
        self.notify("TOCEntry", (level, text, self.page, key))


class PDFExporter(ExportMethod):
    """Concrete exporter rendering a Guide as a PDF document."""

    format_name = "PDF"
    file_extension = "pdf"

    def __init__(self) -> None:
        s = getSampleStyleSheet()

        # Top-level sections (BOM, tools, warnings, procedure).
        s.add(ParagraphStyle(
            name="SectionHeading", parent=s["Heading2"],
            textColor=_PALETTE["primary"], spaceBefore=14, spaceAfter=6,
        ))
        # Step header: solid colour band.
        s.add(ParagraphStyle(
            name="StepTitle", parent=s["Heading3"],
            textColor=_PALETTE["primary_text"], backColor=_PALETTE["primary"],
            borderPadding=(4, 8, 4, 8), leftIndent=0,
            spaceBefore=14, spaceAfter=6,
        ))
        # "Disassembling" and tools lines under the step header.
        s.add(ParagraphStyle(
            name="MetaLine", parent=s["Normal"], fontName="Helvetica-Oblique",
            textColor=_PALETTE["muted_text"], spaceAfter=3,
        ))
        # Plain instruction and safety instruction.
        s.add(ParagraphStyle(name="ActionText", parent=s["Normal"]))
        s.add(ParagraphStyle(
            name="ActionSafety", parent=s["Normal"],
            backColor=_PALETTE["safety_bg"], textColor=_PALETTE["safety_text"],
            borderPadding=(2, 4, 2, 4),
        ))
        # Extracted parts and continuation boxes.
        s.add(ParagraphStyle(
            name="OutputsBox", parent=s["Normal"],
            backColor=_PALETTE["outputs_bg"], borderPadding=(4, 6, 4, 6),
            spaceBefore=4,
        ))
        s.add(ParagraphStyle(
            name="ContinuesBox", parent=s["Normal"],
            backColor=_PALETTE["continues_bg"], borderPadding=(4, 6, 4, 6),
            spaceBefore=4,
        ))
        # BOM table cells: wrapping text inside each cell.
        s.add(ParagraphStyle(name="CellText", parent=s["Normal"], fontSize=9, leading=11))
        s.add(ParagraphStyle(
            name="CellHeader", parent=s["Normal"], fontSize=9, leading=11,
            textColor=colors.white, fontName="Helvetica-Bold",
        ))
        s.add(ParagraphStyle(
            name="WarningLine", parent=s["Normal"], fontSize=9, spaceAfter=2,
        ))
        self._styles = s

    def validate_options(self, options: ExportOptions) -> None:
        """
        Normalize the destination: make sure it is a Path carrying the .pdf
        suffix, since a save dialog may return a name without extension.

        Parameters:
            options: the options to validate, adjusted in place.

        Returns:
            None.
        """
        path = Path(options.output_path)
        if path.suffix.lower() != ".pdf":
            path = path.with_suffix(".pdf")
        options.output_path = path

    # ------------------------------------------------------------------ #
    def export(self, guide: Guide, options: ExportOptions) -> Path:
        """
        Build and write the complete PDF (title, contents, BOM and tools,
        validation warnings, steps).

        Parameters:
            guide: the Guide to export.
            options: rendering options (step cap, images, warnings).

        Returns:
            the Path of the PDF written at options.output_path.
        """
        self.validate_options(options)

        # The GUI may hand over a destination whose folder does not exist yet
        # (a path typed by the user, or a subfolder of a save dialog).
        options.output_path.parent.mkdir(parents=True, exist_ok=True)

        doc = _GuideDocTemplate(
            str(options.output_path),
            product_name=guide.product.name,
            pagesize=A4,
            leftMargin=2 * cm, rightMargin=2 * cm,
            topMargin=2 * cm, bottomMargin=2 * cm,
        )

        story = []
        story += self._build_header(guide)
        story += self._build_toc()
        story.append(PageBreak())
        story += self._build_bom_and_tools(guide)
        if options.show_warnings and guide.warnings:
            story += self._build_warnings(guide)
        story.append(PageBreak())
        story += self._build_steps(guide, options)

        # multiBuild runs several passes: needed to resolve TOC page numbers.
        doc.multiBuild(story)
        # Absolute path, consistent with the other exporters: the GUI shows it
        # to the user in the success dialog.
        return options.output_path.resolve()

    # ------------------------------------------------------------------ #
    # Document sections
    # ------------------------------------------------------------------ #
    def _build_header(self, guide: Guide) -> list:
        """
        Build the document title from the product name.

        Parameters:
            guide: the Guide holding the product.

        Returns:
            list of reportlab flowables to append to the story.
        """
        s = self._styles
        return [
            Paragraph(guide.product.name, s["Title"]),
            Spacer(1, 0.3 * cm),
        ]

    def _build_toc(self) -> list:
        """
        Build the table of contents. Entries are registered by
        _GuideDocTemplate.afterFlowable while the document is laid out and
        resolved by multiBuild; each entry links to its section or step.

        Returns:
            list of reportlab flowables to append to the story.
        """
        s = self._styles
        toc = TableOfContents()
        toc.levelStyles = [
            ParagraphStyle(
                name="TOCLevel0", parent=s["Normal"], fontName="Helvetica-Bold",
                fontSize=11, leftIndent=0, spaceBefore=6,
            ),
            ParagraphStyle(
                name="TOCLevel1", parent=s["Normal"], fontSize=10,
                leftIndent=16, spaceBefore=2,
            ),
        ]
        return [Paragraph("Contents", s["SectionHeading"]), toc]

    def _build_bom_and_tools(self, guide: Guide) -> list:
        """
        Build the bill of materials table and the list of required tools,
        shown at the top of the document (URS 1.3). Every cell is a
        Paragraph so long text wraps inside the cell instead of overflowing.
        Tools come pre-aggregated in guide.tools.

        Parameters:
            guide: the Guide holding the BOM and the aggregated tools.

        Returns:
            list of reportlab flowables to append to the story.
        """
        s = self._styles
        elements = [Paragraph("Bill of Materials", s["SectionHeading"])]

        bom = guide.bill_of_materials
        if bom:
            header = [
                Paragraph("Component", s["CellHeader"]),
                Paragraph("Weight", s["CellHeader"]),
                Paragraph("Material", s["CellHeader"]),
            ]
            data = [header]
            for c in bom:
                weight = f"{c.weight:g} {c.weight_unit}" if c.weight is not None else "-"
                data.append([
                    Paragraph(c.name, s["CellText"]),
                    Paragraph(weight, s["CellText"]),
                    Paragraph(c.material or "-", s["CellText"]),
                ])

            # 6.5 + 2.5 + 8 = 17 cm, the usable width of A4 with 2 cm margins.
            table = Table(data, colWidths=[6.5 * cm, 2.5 * cm, 8 * cm], repeatRows=1)
            table.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, 0), _PALETTE["table_header"]),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, _PALETTE["row_alt"]]),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]))
            elements.append(table)
        else:
            # bom is None when the guide was built with include_bom=False.
            elements.append(Paragraph(
                "Bill of materials not requested at load time.", s["Normal"]
            ))

        elements.append(Spacer(1, 0.4 * cm))
        elements.append(Paragraph("Required tools", s["SectionHeading"]))
        if guide.tools:
            elements.append(ListFlowable(
                [ListItem(Paragraph(t, s["Normal"])) for t in guide.tools],
                bulletType="bullet",
            ))
        else:
            elements.append(Paragraph("No tools specified.", s["Normal"]))

        return elements

    def _build_warnings(self, guide: Guide) -> list:
        """
        Build the validation section, listing every finding with its
        severity and the node ids involved (URS 2.2).

        Parameters:
            guide: the Guide holding guide.warnings.

        Returns:
            list of reportlab flowables to append to the story.
        """
        s = self._styles
        elements = [
            Spacer(1, 0.4 * cm),
            Paragraph("Validation warnings", s["SectionHeading"]),
        ]
        for w in guide.warnings:
            color = _SEVERITY_COLOR.get(w.severity, colors.black)
            ids = ", ".join(str(i) for i in w.node_ids) if w.node_ids else "-"
            text = (
                f'<font color="{color.hexval()}"><b>[{w.severity.value.upper()}]</b></font> '
                f'{w.message} (nodes: {ids})'
            )
            elements.append(Paragraph(text, s["WarningLine"]))
        return elements

    def _build_steps(self, guide: Guide, options: ExportOptions) -> list:
        """
        Build the procedure section by iterating guide.steps, already
        grouped per diamond by the loader (URS 10.0), optionally capping
        how many are rendered (URS 19.0).

        Parameters:
            guide: the Guide holding the steps.
            options: rendering options, including max_groups.

        Returns:
            list of reportlab flowables to append to the story.
        """
        s = self._styles
        elements = [Paragraph("Disassembly procedure", s["SectionHeading"])]

        steps = guide.steps
        if options.max_groups is not None:
            steps = steps[: options.max_groups]

        for step in steps:
            elements += self._render_step(step, options)

        return elements

    def _render_step(self, step: Step, options: ExportOptions) -> list:
        """
        Render one Step: colour-banded header, the component being worked
        on, required tools, numbered instructions with safety highlighting
        and images (URS 4.1), the extracted parts box and the continuation
        box listing every component carrying the disassembly forward.

        Parameters:
            step: the Step to render.
            options: rendering options (e.g. include_images).

        Returns:
            list of reportlab flowables to append to the story.
        """
        s = self._styles
        elements = [Paragraph(f"{step.index}. {step.operation}", s["StepTitle"])]

        # step.input tells which piece this operation opens: on branched
        # models a flat list would otherwise be ambiguous.
        elements.append(Paragraph(f"Disassembling: {step.input.name}", s["MetaLine"]))

        if step.tools_required:
            elements.append(Paragraph(
                f"Tools: {', '.join(step.tools_required)}", s["MetaLine"]
            ))

        if step.actions:
            items = []
            for action in step.actions:
                text = action.text
                if action.tools:
                    text += f" <i>({action.tools})</i>"
                style = s["ActionSafety"] if _is_safety_text(action.text) else s["ActionText"]
                items.append(ListItem(Paragraph(text, style)))
            elements.append(ListFlowable(items, bulletType="1"))

            if options.include_images:
                for action in step.actions:
                    elements += self._render_image_or_link(action.image_path)

        if step.outputs:
            out_names = ", ".join(
                f"{o.name}{self._kept_whole_note(o)}" for o in step.outputs
            )
            elements.append(Paragraph(
                f"<b>Extracted components:</b> {out_names}", s["OutputsBox"]
            ))

        # continues_as is a tuple: a diamond may fork into several
        # composites, each opened by a later step.
        if step.continues_as:
            cont_names = ", ".join(c.name for c in step.continues_as)
            elements.append(Paragraph(
                f"<b>Continues as:</b> {cont_names}", s["ContinuesBox"]
            ))

        elements.append(Spacer(1, 0.4 * cm))
        return elements

    @staticmethod
    def _kept_whole_note(component) -> str:
        """
        Build the suffix marking a sub-assembly kept intact by a depth cut,
        including how many parts it hides (URS 3.0).

        Parameters:
            component: the output component to describe.

        Returns:
            the suffix string, empty when the component was fully taken apart.
        """
        if not component.kept_whole:
            return ""
        if component.contained_leaf_count:
            return f" <i>(kept whole, {component.contained_leaf_count} parts inside)</i>"
        return " <i>(kept whole)</i>"

    def _render_image_or_link(self, image_path: Optional[str]) -> list:
        """
        Render the image of an action. A URL is shown as a clickable link
        (never downloaded); an existing local file is embedded in the PDF.

        Parameters:
            image_path: local path or URL of the image, or None.

        Returns:
            list of reportlab flowables, empty when there is no usable image.
        """
        if not image_path:
            return []
        if image_path.startswith(("http://", "https://")):
            return [Paragraph(
                f'<link href="{image_path}">{image_path}</link>', self._styles["Normal"]
            )]
        p = Path(image_path)
        if not p.exists():
            return []
        try:
            return [Image(str(p), width=6 * cm, height=4 * cm)]
        except Exception:
            return []


def export_pdf(guide: Guide, options: ExportOptions) -> Path:
    """
    Export a Guide to PDF. This is the entry point the GUI calls; it hides
    the exporter class entirely.

    Parameters:
        guide: the Guide to export.
        options: rendering options, including the destination path.

    Returns:
        the Path of the PDF written at options.output_path.
    """
    return PDFExporter().export(guide, options)


def export_to_pdf(
    json_path=None,
    output_path=None,
    *,
    depth=None,
    include_bom: bool = False,
    guide=None,
    max_groups: Optional[int] = None,
    include_images: bool = True,
    show_warnings: bool = True,
) -> Path:
    """
    GUI-facing wrapper, matching the convention shared by all exporters of
    the project: it accepts either a model file to load or an already built
    Guide, and returns the path of the generated PDF.

    Parameters:
        json_path: path of a Builder JSON model (str or Path). May also be
            an already built Guide, which avoids reloading the same model
            once per export format.
        output_path: destination file. Defaults to the model file name with
            a .pdf suffix, or to the product name when a Guide is passed.
        depth: DepthSpec applied when loading a model. Ignored when a Guide
            is passed, whose depth was fixed at load time.
        include_bom: whether to load the bill of materials. Ignored when a
            Guide is passed.
        guide: an already built Guide, alternative to json_path.
        max_groups: cap on exported steps (URS 19.0).
        include_images: embed action images (URS 4.1).
        show_warnings: list validation warnings in the document.

    Returns:
        the Path of the PDF actually created.
    """
    source = guide if guide is not None else json_path
    if source is None:
        raise ValueError("export_to_pdf needs either json_path or guide.")

    if isinstance(source, Guide):
        built_guide = source
        default_name = built_guide.product.name
    else:
        built_guide = build_guide(str(source), depth=depth, include_bom=include_bom)
        default_name = Path(source).stem

    if output_path is None:
        output_path = Path(f"{default_name}.pdf")

    options = ExportOptions(
        output_path=Path(output_path),
        max_groups=max_groups,
        include_images=include_images,
        show_warnings=show_warnings,
    )
    return PDFExporter().export(built_guide, options)
