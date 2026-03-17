# Excel ↔ Json Converter

**Author:** qiangxinglin

**Version:** 0.1.0

**Type:** tool

**Repository** https://github.com/qiangxinglin/excel-dify-plugin

## Description

The built-in `Doc Extractor` would convert input `.xlsx` file to markdown table **string** for downstream nodes (e.g. LLM). But this does not cover all situations! This plugin provides 2 tools:
- `xlsx → json`: Read the Excel file and output the Json presentation of the data.
- `json → xlsx`: Convert the given json string (list of records) to xlsx blob.



## Usage
> [!IMPORTANT]
> Correctly configure the **`FILES_URL`** in your `docker-compose.yaml` or [`.env`](https://github.com/langgenius/dify/blob/main/docker/.env.example#L48) in advance.

![](_assets/workflow_usage.png)

## Tools

### xlsx → json

- The output is placed in the `text` output field.
- Output format is always a 2D array: `[["Header1", "Header2"], ["Row1Col1", "Row1Col2"], ...]`
- All cells are parsed as strings.
- Styles are automatically extracted to the `[styles]` key.
- If the Excel file contains multiple sheets, the output will be an object with sheet names as keys.

**Output example:**
```json
{
  "[styles]": {
    "A1": {"font": {"bold": true}, "fill": {"fgColor": "FFFF00"}},
    "B1": {"font": {"bold": true}}
  },
  "Sheet1": [
    ["项目", "当前值", "年初值"],
    ["资产", "9,880,979", "9,734,788"],
    ["负债", "9,110,145", "8,977,638"]
  ]
}
```

This format is compatible with `json2excel` - you can use the array output directly as input for JSON → Excel conversion.
{
  "[styles]": {
    "A1": {
      "font": {"bold": true, "size": 12},
      "fill": {"fgColor": "FFFF00"},
      "alignment": {"horizontal": "center"}
    },
    "B1": {
      "font": {"bold": true}
    }
  },
  "Sheet1": [
    {"Name": "John", "Age": "18"},
    {"Name": "Doe", "Age": "20"}
  ]
}
```

#### Complete Example: Preserve Styles During Conversion

This example demonstrates how to convert Excel to JSON and back while preserving all cell styles (font, fill, alignment).

**Step 1: Excel → JSON**

The tool automatically outputs 2D array format with `[styles]` key:

```json
{
  "[styles]": {
    "A1": {
      "font": {"bold": true, "size": 15},
      "fill": {"fgColor": "003679A9"},
      "alignment": {"horizontal": "center", "vertical": "center"}
    },
    "A3": {
      "font": {"bold": true, "size": 10},
      "fill": {"fgColor": "003679A9"},
      "alignment": {"horizontal": "center", "vertical": "center"}
    }
  },
  "REPORT0": [
    ["项目", "当前值", "年初值"],
    ["资产", "9,880,979", "9,734,788"],
    ["负债", "9,110,145", "8,977,638"]
  ]
}
```

**Step 2: JSON → Excel**

Use the JSON output from Step 1 directly as input. The `[styles]` key will be automatically applied to the generated Excel file.

**Supported style properties:**
- `font.bold`: boolean
- `font.italic`: boolean
- `font.size`: number
- `font.color`: hex color string (e.g., "00FFFFFF")
- `font.underline`: boolean
- `fill.fgColor`: foreground color hex string
- `fill.bgColor`: background color hex string
- `alignment.horizontal`: "left", "center", "right", etc.
- `alignment.vertical`: "top", "center", "bottom", etc.

**Note:** Only cells with non-default styles are included in the `[styles]` output to keep the JSON concise.

### json → xlsx

- The output filename can be configured, default `Converted_Data`
- If the input JSON is an object (whose values are arrays), the plugin will automatically create a multi-sheet Excel file, where each key of the object will become a sheet name.

![](_assets/workflow_run.png)
![](_assets/output_xlsx.png)

#### Format Settings

The plugin supports optional formatting for row heights and column widths using the `[format]` reserved key.

> **Note:** Excel prohibits sheet names containing these characters: `/ \ ? * : [ ]`
> Therefore, `[format]` is guaranteed to never conflict with actual sheet names.

##### `[format]` Structure

```json
{
  "[format]": {
    "defaults": {
      "rowHeight": 20,           // Default height for all rows
      "columnWidth": 15,         // Default width for all columns
      "rowHeights": {            // Specific row heights (1-based)
        "1": 30,                 // Row 1 height = 30
        "2": 25                  // Row 2 height = 25
      },
      "columnWidths": {          // Specific column widths
        "A": 25,                 // Column A width = 25
        "B": 15                  // Column B width = 15
      }
    },
    "sheets": {
      "SheetName": {             // Per-sheet overrides
        "rowHeight": 22,
        "columnWidth": 18,
        "rowHeights": {"1": 28},
        "columnWidths": {"A": 30, "B": 20}
      }
    }
  },
  "SheetName": [...]             // Actual data
}
```

##### Format Priority Rules

Settings are applied in the following order (later overrides earlier):

1. `[format].defaults.rowHeight` / `columnWidth` - Global defaults for all rows/columns
2. `[format].sheets.<name>.rowHeight` / `columnWidth` - Per-sheet defaults
3. `[format].defaults.rowHeights` / `columnWidths` - Global specific rows/columns
4. `[format].sheets.<name>.rowHeights` / `columnWidths` - Per-sheet specific rows/columns

##### Validation and Warnings

- **Unknown sheet references**: If `[format].sheets` references a sheet that doesn't exist in the data, a warning will be displayed and those configurations will be ignored. The Excel file will still be generated successfully.
- **Type errors**: If format values have incorrect types (e.g., non-dict, negative numbers), an error will be thrown and the Excel generation will fail.

##### Examples

**Single sheet with global formatting:**

```json
{
  "[format]": {
    "defaults": {
      "rowHeight": 20,
      "columnWidth": 15
    }
  },
  "Sheet1": [
    {"Name": "John", "Age": "18"},
    {"Name": "Doe", "Age": "20"}
  ]
}
```

**Multiple sheets with per-sheet formatting:**

```json
{
  "[format]": {
    "defaults": {
      "rowHeight": 18
    },
    "sheets": {
      "Employees": {
        "columnWidths": {"A": 25, "B": 15},
        "rowHeights": {"1": 30}
      },
      "Departments": {
        "columnWidths": {"A": 20}
      }
    }
  },
  "Employees": [{"Name": "John", "Department": "R&D"}],
  "Departments": [{"ID": "1", "Name": "HR"}]
}
```

**Column identifiers:**

You can use either Excel letters or 1-based numeric indexes:

```json
{
  "[format]": {
    "defaults": {
      "columnWidths": {
        "A": 25,    // Letter format
        "1": 25,    // Numeric format (same as "A")
        "B": 15,
        "2": 15     // Same as "B"
      }
    }
  },
  "Sheet1": [...]
}
```

#### [meta] Structure

The `[meta]` reserved key allows inserting custom content (title, author, date, etc.) before the header row.

> **Note:** `[meta]` is safe to use because Excel prohibits `[` and `]` in sheet names.

```json
{
  "[meta]": [
    {"A": "Report Title", "D": "Date: 2024-01-01"},
    {"A": "Author: John Doe"}
  ],
  "Sheet1": [
    {"Name": "John", "Age": "18"},
    {"Name": "Jane", "Age": "20"}
  ]
}
```

This produces:
- Row 1: "Report Title" in A, "Date: 2024-01-01" in D
- Row 2: "Author: John Doe" in A
- Row 3: Header (Name, Age)
- Row 4-5: Data rows

##### Per-sheet Meta

You can also specify meta per-sheet using an object format:

```json
{
  "[meta]": {
    "Employees": [
      {"A": "Employee Report", "D": "2024-01-01"},
      {"A": "Author: HR Team"}
    ],
    "Departments": [
      {"A": "Department Overview"}
    ]
  },
  "Employees": [...],
  "Departments": [...]
}
```

##### Combining [meta] and [format]

You can use both `[meta]` and `[format]` together:

```json
{
  "[meta]": [
    {"A": "My Report", "C": "2024-01-01"}
  ],
  "[format]": {
    "defaults": {
      "rowHeight": 20
    }
  },
  "Sheet1": [...]
}
```

##### Column Identifiers

Meta rows use the same column identifier format as `[format]`:
- Letters: `"A"`, `"B"`, `"AA"`, etc.
- 1-based integers: `"1"`, `"2"`, etc.

#### [styles] Structure

The `[styles]` key allows applying cell formatting (font, fill, alignment) when generating Excel. This enables round-trip conversion: read Excel with `include_styles=true` → output can be used as input to preserve styles.

```json
{
  "[styles]": {
    "A1": {
      "font": {"bold": true, "size": 12},
      "fill": {"fgColor": "FFFF00"},
      "alignment": {"horizontal": "center", "vertical": "center"}
    },
    "B1": {
      "font": {"bold": true, "italic": true}
    },
    "A2": {
      "fill": {"fgColor": "FFCCCC"}
    }
  },
  "Sheet1": [
    {"Name": "John", "Age": "18"},
    {"Name": "Jane", "Age": "20"}
  ]
}
```

**Style properties:**
- `font.bold`: boolean
- `font.italic`: boolean  
- `font.size`: number
- `font.color`: hex color string (e.g., "00FFFFFF", "FFFF00")
- `font.underline`: boolean
- `fill.fgColor`: foreground color hex string
- `fill.bgColor`: background color hex string
- `alignment.horizontal`: "left", "center", "right", etc.
- `alignment.vertical`: "top", "center", "bottom", etc.

**Per-sheet styles:** Use an object with sheet names as keys:

```json
{
  "[styles]": {
    "Employees": {
      "A1": {"font": {"bold": true}}
    },
    "Departments": {
      "A1": {"font": {"italic": true}}
    }
  },
  "Employees": [...],
  "Departments": [...]
}
```

**Round-trip example:** Read Excel with `include_styles=true`, then use the output JSON directly as input to preserve styles.


## Used Open sourced projects

- [pandas](https://github.com/pandas-dev/pandas), BSD 3-Clause License

## Changelog
- **0.1.0**: Refactor excel2json to use pure openpyxl (no pandas), always output 2D array, always extract styles
- **0.0.9**: Add `output_format` parameter to excel2json (records/array) and support array format input in json2excel
- **0.0.8**: Add `[styles]` support to json2excel for applying cell formatting (enables round-trip with excel2json)
- **0.0.7**: Add `include_styles` parameter to excel2json for extracting cell formatting (font, fill, alignment)
- **0.0.6**: Add `[meta]` support for inserting custom content (title, author, etc.) before header row
- **0.0.5**: Add `[format]` metadata support for controlling row heights and column widths during JSON → Excel conversion
- **0.0.4**: Add missing dependency (xlrd)
- **0.0.3**: Add multi-sheet support for Excel processing (closes #13)

## License
- Apache License 2.0


## Privacy

This plugin collects no data.

All the file transformations are completed locally. NO data is transmitted to third-party services.