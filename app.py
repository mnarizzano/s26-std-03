"""
Disassembly Wizard - Desktop GUI Application (app.py)
Integrator: Alhareth Al-Haji
Description: Desktop graphical interface connecting Rubin's disassembly_loader
             with all team export modules (PDF, Word, Markdown, HTML+JS, PPTX, TXT).
"""

import os
import sys
import json
import importlib.util
from pathlib import Path
import customtkinter as ctk
from tkinter import filedialog, messagebox

# ==============================================================================
# 1. DIRECTORY SETUP & PATH RESOLUTION
# ==============================================================================
CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
if os.path.isdir(os.path.join(CURRENT_DIR, "src")):
    ROOT_DIR = CURRENT_DIR
else:
    ROOT_DIR = os.path.abspath(os.path.join(CURRENT_DIR, ".."))


SEARCH_PATHS = [
    CURRENT_DIR,
    os.path.join(ROOT_DIR, "src", "main"),
    os.path.join(ROOT_DIR, "src", "main", "word-exporter"),
    os.path.join(ROOT_DIR, "src", "main", "pptx-converter"),
    os.path.join(ROOT_DIR, "src", "main", "txt-exporter"),
    os.path.join(ROOT_DIR, "src", "main", "html-js-exporter"),
    os.path.join(ROOT_DIR, "src", "main", "md_exporter"),
    os.path.join(ROOT_DIR, "src", "main", "loader_se"),
]

for path_dir in SEARCH_PATHS:
    if os.path.isdir(path_dir) and path_dir not in sys.path:
        sys.path.insert(0, path_dir)

# ==============================================================================
# 2. MODULE RESOLUTION & INTEGRATION
# ==============================================================================
# Disassembly Loader
try:
    from disassembly_loader import build_guide, DepthSpec, DepthMode, UnparsableModelError
except ImportError:
    try:
        from loader_se import build_guide, DepthSpec, DepthMode, UnparsableModelError
    except ImportError:
        build_guide = None
        DepthSpec = None
        DepthMode = None
        UnparsableModelError = Exception

# 1. PDF Exporter (Gabriele)
try:
    from pdf_exporter import export_to_pdf
except Exception:
    export_to_pdf = None

# 2. Word Exporter (Simone)
try:
    from word_exporter import export_to_word
except Exception:
    export_to_word = None

# 3. Markdown Exporter (Emanuele)
try:
    from md_exporter import export_to_md
except Exception:
    export_to_md = None

# 4. PowerPoint Exporter (Ibrahim)
export_to_pptx = None
pptx_exporter_path = os.path.join(ROOT_DIR, "src", "main", "pptx-converter", "pptx_exporter.py")
if os.path.isfile(pptx_exporter_path):
    try:
        spec_pptx = importlib.util.spec_from_file_location("pptx_exporter_root", pptx_exporter_path)
        if spec_pptx and spec_pptx.loader:
            mod_pptx = importlib.util.module_from_spec(spec_pptx)
            spec_pptx.loader.exec_module(mod_pptx)
            export_to_pptx = getattr(mod_pptx, "export_to_pptx", None)
    except Exception:
        export_to_pptx = None

# 5. Text Exporter (Ali)
TXTExporter = None
txt_module_path = os.path.join(ROOT_DIR, "src", "main", "txt-exporter", "exporters", "txt_exporter.py")
if os.path.isfile(txt_module_path):
    try:
        spec_txt = importlib.util.spec_from_file_location("ali_txt_exporter", txt_module_path)
        if spec_txt and spec_txt.loader:
            mod_txt = importlib.util.module_from_spec(spec_txt)
            spec_txt.loader.exec_module(mod_txt)
            TXTExporter = getattr(mod_txt, "TXTExporter", None)
    except Exception:
        TXTExporter = None


def generate_interactive_html(input_json_path, output_html_path, guide=None):
    """Fallback generator for Leen's interactive HTML/JS guide if module is missing."""
    with open(input_json_path, "r", encoding="utf-8") as f:
        model_data = json.load(f)

    # Normalize IR schema structure
    ir_dict = {
        "schema_version": model_data.get("schema_version", "1.0"),
        "product": model_data.get("product", {
            "name": getattr(guide.product, "name", "Product") if guide else "Product",
            "weight": getattr(guide.product, "weight", 0) if guide else 0,
            "weight_unit": "g"
        }),
        "depth": model_data.get("depth", {"mode": "full", "keep_whole_ids": []}),
        "steps": model_data.get("steps", []),
        "warnings": model_data.get("warnings", []),
        "bill_of_materials": model_data.get("bill_of_materials", [])
    }

    # If steps are empty in JSON root, populate from guide object
    if not ir_dict["steps"] and guide and hasattr(guide, "steps"):
        for s in guide.steps:
            outs = []
            if hasattr(s, "outputs") and s.outputs:
                for o in s.outputs:
                    o_obj = o[0] if isinstance(o, (tuple, list)) else o
                    outs.append({
                        "node_id": getattr(o_obj, "node_id", 0),
                        "name": getattr(o_obj, "name", str(o_obj)),
                        "weight": getattr(o_obj, "weight", 0),
                        "weight_unit": "g",
                        "material": getattr(o_obj, "material", "1"),
                        "color": getattr(o_obj, "color", "1"),
                        "image": None
                    })
            ir_dict["steps"].append({
                "index": s.index,
                "operation": s.operation,
                "actions": [{"text": getattr(a, "text", str(a))} for a in getattr(s, "actions", [])],
                "outputs": outs,
                "continues_as": {"name": getattr(s.continues_as, "name", "")} if getattr(s, "continues_as", None) else None,
                "tools_required": getattr(s, "tools_required", [])
            })

    ir_json_str = json.dumps(ir_dict, ensure_ascii=False)
    prod_name = ir_dict["product"].get("name", "Disassembly Wizard")

    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>{prod_name} - Disassembly Wizard</title>
<style>
:root{{--bg:#0d1015;--panel:#151a22;--panel2:#1c222d;--line:#2b3441;--text:#f4f5f7;--muted:#9aabc0;--accent:#3b82f6;--success:#10b981}}
*{{box-sizing:border-box}}body{{margin:0;font-family:Inter,system-ui,sans-serif;background:var(--bg);color:var(--text)}}
.app{{min-height:100vh;display:flex}}.sidebar{{width:300px;border-right:1px solid var(--line);padding:20px;background:#10141a;height:100vh;overflow:auto}}
.brand h1{{margin:0;font-size:20px;color:var(--accent)}}.step-item{{width:100%;text-align:left;padding:12px;border-radius:10px;background:transparent;color:var(--text);border:1px solid transparent;margin-bottom:6px;cursor:pointer}}
.step-item.active{{background:var(--panel2);border-color:var(--accent)}}.content{{flex:1;padding:30px 45px;max-width:1200px;margin:0 auto}}
.progress{{height:8px;background:#1d2530;border-radius:99px;overflow:hidden;margin-bottom:20px}}.progress div{{height:100%;background:var(--accent);transition:width .3s}}
.step-layout{{display:grid;grid-template-columns:1fr;gap:20px;margin-top:20px}}.instructions-panel,.part-card{{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:20px}}
.parts-grid{{display:grid;grid-template-columns:repeat(auto-fit,minmax(260px,1fr));gap:16px;margin-top:15px}}
.btn{{border:1px solid var(--line);border-radius:10px;padding:12px 24px;font-weight:700;cursor:pointer;background:var(--accent);color:#fff}}
.btn-sec{{background:#151a22;color:var(--text)}}.nav-row{{display:flex;gap:15px;margin-top:30px}}
</style>
</head>
<body>
<div class="app">
  <aside class="sidebar">
    <div class="brand"><h1>Disassembly Guide</h1><p>{prod_name}</p></div>
    <div id="stepsList"></div>
  </aside>
  <main class="content">
    <div class="progress"><div id="progressBar"></div></div>
    <h3 id="stepCounter" style="color:var(--muted)"></h3>
    <h1 id="stepOperation"></h1>
    <div class="step-layout">
      <div class="instructions-panel">
        <h3>Instructions</h3>
        <div id="actionsList"></div>
        <p id="continuesAs" style="color:var(--muted);margin-top:15px;"></p>
      </div>
      <div>
        <h3>Recovered Components</h3>
        <div id="outputsList" class="parts-grid"></div>
      </div>
    </div>
    <div class="nav-row">
      <button class="btn btn-sec" onclick="prevStep()">Back</button>
      <button class="btn" style="margin-left:auto" onclick="nextStep()" id="nextBtn">Next</button>
    </div>
  </main>
</div>
<script>
const ir = {ir_json_str};
let current = 0;
const steps = ir.steps || [];
function render(){{
  if(!steps.length) return;
  const s = steps[current];
  document.getElementById('stepCounter').textContent = 'STEP ' + (current+1) + ' OF ' + steps.length;
  document.getElementById('stepOperation').textContent = s.operation;
  document.getElementById('progressBar').style.width = ((current+1)/steps.length*100) + '%';
  document.getElementById('actionsList').innerHTML = (s.actions||[]).map((a,i)=>'<p><b>' + (i+1) + '.</b> ' + a.text + '</p>').join('') || '<p>Perform action carefully.</p>';
  document.getElementById('continuesAs').textContent = s.continues_as ? 'Remaining Assembly: ' + s.continues_as.name : 'End of disassembly branch.';
  document.getElementById('outputsList').innerHTML = (s.outputs||[]).map(o=>'<div class="part-card"><h4>' + o.name + '</h4><p>Weight: ' + (o.weight||'—') + ' g</p></div>').join('') || '<p>No parts detached.</p>';
  document.getElementById('stepsList').innerHTML = steps.map((st,i)=>'<button class="step-item ' + (i===current?'active':'') + '" onclick="goTo(' + i + ')">Step ' + (i+1) + ': ' + st.operation + '</button>').join('');
  document.getElementById('nextBtn').textContent = current === steps.length-1 ? 'Finish' : 'Next';
}}
function nextStep(){{ if(current < steps.length-1){{ current++; render(); }} else alert('Disassembly workflow complete!'); }}
function prevStep(){{ if(current > 0){{ current--; render(); }} }}
function goTo(i){{ current = i; render(); }}
render();
</script>
</body>
</html>
"""
    with open(output_html_path, "w", encoding="utf-8") as f:
        f.write(html_content)
    return output_html_path


# Configure CustomTkinter Theme
ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")

# ==============================================================================
# 3. MAIN GUI APPLICATION CLASS
# ==============================================================================
class EliteExporterApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title("Disassembly Wizard - Desktop GUI Exporter")
        self.geometry("980x720")

        self.json_file_path = ctk.StringVar(value="")
        self.output_dir_path = ctk.StringVar(value="")
        self.selected_format = "PDF"

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=4)
        self.grid_rowconfigure(0, weight=1)

        self.create_sidebar()
        self.create_main_content()

    def create_sidebar(self):
        self.sidebar = ctk.CTkFrame(self, width=220, corner_radius=0)
        self.sidebar.grid(row=0, column=0, sticky="nsew")
        self.sidebar.grid_rowconfigure(8, weight=1)

        sidebar_title = ctk.CTkLabel(
            self.sidebar, text="EXPORT FORMATS", font=ctk.CTkFont(size=14, weight="bold")
        )
        sidebar_title.grid(row=0, column=0, padx=20, pady=(30, 20))

        formats = [
            ("PDF (.pdf)", "PDF"),
            ("Word (.docx)", "DOCX"),
            ("Markdown (.md)", "MD"),
            ("HTML + JS", "HTML+JS"),
            ("PowerPoint (.pptx)", "PPTX"),
            ("Text (.txt)", "TXT")
        ]

        self.format_buttons = {}
        for idx, (text, fmt) in enumerate(formats):
            btn = ctk.CTkButton(
                self.sidebar,
                text=text,
                fg_color="transparent",
                text_color=("gray10", "gray90"),
                hover_color=("gray70", "gray30"),
                anchor="w",
                command=lambda f=fmt: self.change_format(f)
            )
            btn.grid(row=idx + 1, column=0, padx=10, pady=5, sticky="ew")
            self.format_buttons[fmt] = btn

        self.mode_switch = ctk.CTkSwitch(self.sidebar, text="Dark Mode", command=self.toggle_theme)
        self.mode_switch.grid(row=9, column=0, padx=20, pady=20, sticky="s")
        if ctk.get_appearance_mode() == "Dark":
            self.mode_switch.select()

    def create_main_content(self):
        self.content_area = ctk.CTkFrame(self, fg_color="transparent")
        self.content_area.grid(row=0, column=1, sticky="nsew", padx=30, pady=20)
        self.content_area.grid_columnconfigure(0, weight=1)
        self.content_area.grid_rowconfigure(4, weight=1)

        self.main_title = ctk.CTkLabel(
            self.content_area, text="PDF Export Settings", font=ctk.CTkFont(size=22, weight="bold")
        )
        self.main_title.grid(row=0, column=0, sticky="w", pady=(0, 15))

        # Panel 1: Input JSON Selection
        self.input_frame = ctk.CTkFrame(self.content_area)
        self.input_frame.grid(row=1, column=0, sticky="ew", pady=5, padx=5)
        self.input_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.input_frame, text="Source Model (JSON)", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=15, pady=(8, 0), sticky="w"
        )
        self.file_entry = ctk.CTkEntry(
            self.input_frame, textvariable=self.json_file_path, placeholder_text="Select input JSON model..."
        )
        self.file_entry.grid(row=1, column=0, padx=15, pady=(5, 10), sticky="ew")
        self.upload_btn = ctk.CTkButton(self.input_frame, text="Browse Input", width=110, command=self.browse_json)
        self.upload_btn.grid(row=1, column=1, padx=(0, 15), pady=(5, 10))

        # Panel 2: Destination Output Directory
        self.output_frame = ctk.CTkFrame(self.content_area)
        self.output_frame.grid(row=2, column=0, sticky="ew", pady=5, padx=5)
        self.output_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(self.output_frame, text="Output Directory (Optional Destination)", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=15, pady=(8, 0), sticky="w"
        )
        self.output_entry = ctk.CTkEntry(
            self.output_frame, textvariable=self.output_dir_path, placeholder_text="Default: same directory as input model"
        )
        self.output_entry.grid(row=1, column=0, padx=15, pady=(5, 10), sticky="ew")
        self.output_btn = ctk.CTkButton(self.output_frame, text="Browse Output", width=110, command=self.browse_output_dir)
        self.output_btn.grid(row=1, column=1, padx=(0, 15), pady=(5, 10))

        # Panel 3: Parameters
        self.options_frame = ctk.CTkFrame(self.content_area)
        self.options_frame.grid(row=3, column=0, sticky="ew", pady=5, padx=5)
        self.options_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(self.options_frame, text="Disassembly Parameters", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=0, columnspan=2, padx=15, pady=(8, 0), sticky="w"
        )
        ctk.CTkLabel(self.options_frame, text="Disassembly Depth:").grid(row=1, column=0, padx=15, pady=10, sticky="w")
        self.depth_dropdown = ctk.CTkOptionMenu(
            self.options_frame,
            values=["Full Disassembly", "Keep Main Assemblies (KEEP_MAIN)"]
        )
        self.depth_dropdown.grid(row=1, column=1, padx=15, pady=10, sticky="w")

        self.bom_switch = ctk.CTkSwitch(self.options_frame, text="Include Bill of Materials (BoM)")
        self.bom_switch.grid(row=2, column=0, columnspan=2, padx=15, pady=(0, 10), sticky="w")

        # Panel 4: Telemetry Logs
        self.log_frame = ctk.CTkFrame(self.content_area)
        self.log_frame.grid(row=4, column=0, sticky="nsew", pady=5, padx=5)
        self.log_frame.grid_columnconfigure(0, weight=1)
        self.log_frame.grid_rowconfigure(1, weight=1)

        ctk.CTkLabel(self.log_frame, text="System Log Console & Validation Telemetry", font=ctk.CTkFont(size=12, weight="bold")).grid(
            row=0, column=0, padx=15, pady=(8, 0), sticky="w"
        )
        self.log_textbox = ctk.CTkTextbox(self.log_frame, font=ctk.CTkFont(family="Courier", size=12))
        self.log_textbox.grid(row=1, column=0, padx=10, pady=10, sticky="nsew")
        self.log_write("Status: System ready. Please select an input model file.")

        # Panel 5: Export Trigger
        self.export_btn = ctk.CTkButton(
            self.content_area, text="GENERATE EXPORT", font=ctk.CTkFont(size=14, weight="bold"),
            height=45, command=self.execute_export
        )
        self.export_btn.grid(row=5, column=0, pady=(15, 0))

        self.change_format("PDF")

    def browse_json(self):
        path = filedialog.askopenfilename(filetypes=[("JSON Files", "*.json")])
        if path:
            self.json_file_path.set(path)
            self.log_write(f"[INPUT BOUND] Selected model: {os.path.basename(path)}")

    def browse_output_dir(self):
        folder = filedialog.askdirectory()
        if folder:
            self.output_dir_path.set(folder)
            self.log_write(f"[OUTPUT DIR BOUND] Destination: {folder}")

    def change_format(self, fmt):
        self.selected_format = fmt
        self.main_title.configure(text=f"{fmt} Export Settings")
        for f, btn in self.format_buttons.items():
            if f == fmt:
                btn.configure(fg_color=ctk.ThemeManager.theme["CTkButton"]["fg_color"], text_color="white")
            else:
                btn.configure(fg_color="transparent", text_color=("gray10", "gray90"))

    def toggle_theme(self):
        ctk.set_appearance_mode("Dark" if self.mode_switch.get() == 1 else "Light")

    def log_write(self, txt):
        self.log_textbox.configure(state="normal")
        self.log_textbox.insert("end", txt + "\n")
        self.log_textbox.see("end")
        self.log_textbox.configure(state="disabled")

    def execute_export(self):
        input_path = self.json_file_path.get().strip()
        if not input_path:
            messagebox.showwarning("Missing Input", "Please select a JSON model file first.")
            return

        if not build_guide:
            self.log_write("[ERROR] Disassembly loader module could not be loaded.")
            messagebox.showerror("Error", "'disassembly_loader' module not found.")
            return

        base_name = os.path.splitext(os.path.basename(input_path))[0]
        ext_map = {
            "PDF": ".pdf",
            "DOCX": ".docx",
            "MD": ".md",
            "HTML+JS": ".html",
            "PPTX": ".pptx",
            "TXT": ".txt"
        }
        chosen_ext = ext_map.get(self.selected_format, "")

        out_dir = self.output_dir_path.get().strip()
        if out_dir and os.path.isdir(out_dir):
            target_out_path = os.path.join(out_dir, f"{base_name}_guide{chosen_ext}")
        else:
            target_out_path = os.path.join(os.path.dirname(input_path), f"{base_name}_guide{chosen_ext}")

        depth_val = self.depth_dropdown.get()
        if "Keep Main" in depth_val:
            spec = DepthSpec(mode=DepthMode.KEEP_MAIN)
        else:
            spec = DepthSpec(mode=DepthMode.FULL)

        bom_needed = bool(self.bom_switch.get())

        try:
            self.log_write(f"\n--- Loading Model: {os.path.basename(input_path)} ---")
            guide = build_guide(input_path, depth=spec, include_bom=bom_needed)
            self.log_write(f"[SUCCESS] Product: '{guide.product.name}' | Total Steps: {len(guide.steps)}")

            if guide.warnings:
                self.log_write(f"Validation warnings encountered ({len(guide.warnings)}):")
                for w in guide.warnings:
                    self.log_write(f"  - [{w.severity.upper()}] {w.rule}: {w.message}")

            self.log_write(f"[ROUTING] Exporting to {self.selected_format}...")
            generated = None

            # 1. PDF Exporter (Gabriele)
            if self.selected_format == "PDF":
                if export_to_pdf:
                    generated = export_to_pdf(input_path, output_path=target_out_path, depth=spec, include_bom=bom_needed, guide=guide)
                else:
                    messagebox.showerror("Module Error", "'pdf_exporter.py' not found.")

            # 2. Word Exporter (Simone)
            elif self.selected_format == "DOCX":
                if export_to_word:
                    generated = export_to_word(input_path, output_path=target_out_path, depth=spec, include_bom=bom_needed)
                else:
                    messagebox.showerror("Module Error", "'word_exporter.py' not found.")

            # 3. Markdown Exporter (Emanuele)
            elif self.selected_format == "MD":
                if export_to_md:
                    try:
                        generated = export_to_md(input_path, output_path=target_out_path, depth=spec, include_bom=bom_needed)
                    except TypeError:
                        generated = export_to_md(input_path, depth=spec, include_bom=bom_needed)
                else:
                    messagebox.showerror("Module Error", "'md_exporter.py' not found.")

            # 4. HTML+JS Exporter (Leen)
            elif self.selected_format == "HTML+JS":
                generated = generate_interactive_html(input_path, target_out_path, guide=guide)

            # 5. PowerPoint Exporter (Ibrahim)
            elif self.selected_format == "PPTX":
                if export_to_pptx:
                    generated = export_to_pptx(
                        json_path=input_path,
                        output_path=target_out_path,
                        guide=guide,
                        depth=spec,
                        include_bom=bom_needed
                    )
                else:
                    messagebox.showerror("Module Error", "'pptx_exporter.py' not found.")

            # 6. Text Exporter (Ali)
            elif self.selected_format == "TXT":
                if TXTExporter:
                    exporter = TXTExporter()
                    exporter.export(str(input_path), str(target_out_path))
                    generated = target_out_path
                else:
                    messagebox.showerror("Module Error", "'txt_exporter.py' (TXTExporter) not found.")

            if generated:
                self.log_write(f"[SUCCESS] File generated at:\n  {generated}")
                messagebox.showinfo("Export Successful", f"Guide successfully exported to:\n{generated}")

        except UnparsableModelError as e:
            self.log_write(f"[PARSING ERROR] {str(e)}")
            messagebox.showerror("Parsing Error", str(e))
        except Exception as e:
            self.log_write(f"[EXECUTION ERROR] {str(e)}")
            messagebox.showerror("Execution Error", str(e))


if __name__ == "__main__":
    app = EliteExporterApp()
    app.mainloop()