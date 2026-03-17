# AGENTS.md - Excel Dify Plugin

## Project Overview

**Type**: Dify Plugin (Python)  
**Purpose**: Convert Excel files to JSON and JSON to Excel files  
**Author**: qiangxinglin  
**Version**: 0.1.0

## Project Structure

```
excel-dify-plugin/
├── main.py                    # Plugin entry point (Plugin.run())
├── manifest.yaml              # Plugin metadata, version, author, permissions
├── requirements.txt           # Python dependencies
├── provider/
│   ├── excel_tools.py         # ToolProvider class (no credentials needed)
│   └── excel_tools.yaml       # Provider definition
├── tools/
│   ├── excel2json.py          # Excel → JSON tool implementation
│   ├── excel2json.yaml       # Excel → JSON tool definition
│   ├── json2excel.py         # JSON → Excel tool implementation
│   └── json2excel.yaml       # JSON → Excel tool definition
├── .github/workflows/
│   └── plugin-publish.yml     # CI: Package and publish on release
├── _assets/                   # Images for README
├── GUIDE.md                   # Additional documentation
└── PRIVACY.md                 # Privacy policy
```

## Dependencies

```
dify_plugin>=0.2.0,<0.3.0
numpy==1.26.4
pandas==2.2.3
openpyxl==3.1.5
xlrd==2.0.2
```

## Essential Commands

### Development/Testing
- No local test commands defined
- Install dependencies: `pip install -r requirements.txt`

### Publishing
- **Trigger**: Create a GitHub release
- **Process**: GitHub Actions workflow (`plugin-publish.yml`) automatically:
  1. Packages the plugin using `dify-plugin` CLI
  2. Creates a branch in `langgenius/dify-plugins`
  3. Opens a PR to merge the packaged `.difypkg` file

### Running Locally
- Run plugin: `python main.py` (requires Dify plugin daemon)

## Code Patterns

### Tool Implementation
Tools inherit from `dify_plugin.Tool` and implement `_invoke()`:

```python
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from typing import Any
from collections.abc import Generator

class MyTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        # Access parameters via tool_parameters dict
        yield self.create_text_message("output")
        yield self.create_blob_message(blob=bytes, meta={"mime_type": "...", "filename": "..."})
```

### Tool YAML Definition
- `identity.name`: Tool identifier
- `identity.author`: Tool author
- `parameters`: List of input parameters with types (file, string, etc.)
- `extra.python.source`: Path to Python implementation

### Provider Implementation
Simple provider without credentials:

```python
from dify_plugin import ToolProvider
from dify_plugin.errors.tool import ToolProviderCredentialValidationError

class ExcelToolsProvider(ToolProvider):
    def _validate_credentials(self, credentials: dict[str, Any]) -> None:
        # No credentials needed - pass
        pass
```

## Naming Conventions

- **Tools**: `excel2json`, `json2excel` (camelCase in YAML, snake_case in Python)
- **Provider**: `ExcelToolsProvider` (PascalCase)
- **Files**: snake_case (`excel2json.py`, `excel_tools.py`)
- **YAML files**: snake_case matching Python file names

## Key Implementation Details

### Excel2Json Tool
- Input: Excel file (via `file` parameter)
- Output: JSON string in `text` field
- Always outputs 2D array format: `[[header1, header2], [row1], [row2], ...]`
- Always extracts styles to `[styles]` key
- Supports multi-sheet Excel (object with sheet names as keys)

### Json2Excel Tool
- Input: JSON string (2D array or array of objects, object for multi-sheet)
- Optional: `filename` parameter for output filename
- Output: Blob message with Excel file
- Supports `[format]` reserved key for row heights and column widths
- Supports `[meta]` reserved key for inserting content before header
- Supports `[styles]` reserved key for cell formatting
- Supports `[styles]` reserved key for cell formatting

### [format] Feature
- `[format]` key is reserved (Excel prohibits sheet names with `[` or `]`)
- Supports `defaults` and `sheets` sections
- Allows `rowHeight`, `columnWidth`, `rowHeights`, `columnWidths`
- Column identifiers: letters ("A", "B") or 1-based integers ("1", "2")

### [meta] Feature
- `[meta]` key is reserved (Excel prohibits `[` and `]` in sheet names)
- Insert custom rows before the header (title, author, date, etc.)
- Format: array of objects mapping column letter to value
- Global meta (applies to first sheet): `{"[meta]": [{"A": "Title"}], "Sheet1": [...]}`
- Per-sheet meta: `{"[meta]": {"Sheet1": [...]}, "Sheet1": [...]}`
- Column identifiers same as `[format]`: letters ("A") or 1-based integers ("1")

### [styles] Feature
- **excel2json**: When `include_styles=true`, extracts cell formatting into `[styles]` key
- **json2excel**: Accepts `[styles]` key to apply cell formatting when generating Excel
- Maps cell coordinates (e.g., "A1", "B2") to style objects
- Supports: font (bold, italic, size, color, underline), fill (fgColor, bgColor), alignment (horizontal, vertical)
- Only non-default styles are included to keep output concise
- Enables round-trip: read Excel (with include_styles) → write back with [styles] → preserves formatting

## Important Gotchas

1. **[format], [meta], and [styles] reserved keys**: All are safe because Excel prohibits `[` and `]` in sheet names
2. **Output field**: Excel2Json outputs to `text` field (not `json`) to preserve header order
3. **String parsing**: All Excel cells are parsed as strings regardless of content
4. **Multi-sheet Excel**: Returns object with sheet names as keys when multiple sheets exist
5. **Column/Row indices**: 1-based (Excel convention), not 0-based
6. **Styles round-trip**: Use output from excel2json (with include_styles) directly as input to json2excel to preserve styles

## Dify Plugin API Reference

- `Tool`: Base class for tools
- `ToolProvider`: Base class for tool providers
- `ToolInvokeMessage`: Return messages (text, blob, json, link)
- `create_text_message(text)`: Return text output
- `create_blob_message(blob, meta)`: Return binary file
- `DifyPluginEnv(MAX_REQUEST_TIMEOUT=N)`: Configure timeout in seconds
