**Module name**: word\_exporter

**Module path**: src/main/word-exporter/word\_exporter.py

**Function name and call**: 
	export\_to\_word(json\_path, output\_path=None, depth=None, include\_bom=False)
	
**Inputs**:

- [type: string] [format: absolute path with slashes] "json\_path" : path of the json input file
- [type: string] [format: absolute path] [optional] "output\_path" : destination path for the generated docx file
- [type: DepthSpec (see the loader module)] [optional] "depth": max depth level of the disassemblying process
- [type: bool] [optional] "include\_bom": whether the bill of material has to be added (or not) to the generated docx

**Outputs**:

- [type: string] [format: absolute path with slashes] path of the generated docx file

 
