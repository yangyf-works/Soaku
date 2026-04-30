# Soaku

A fast and practical tool to compare differences between two Excel files.

The name “Soaku” was inspired by Git's naming approach.

It detects:
- Added / deleted sheets
- Added / deleted columns
- Added / deleted rows
- Modified cells (with position mapping)

---

## Features
- Fast (optimized for large files, even sheets with over 20,000 rows can be compared)
- Sheet / Column / Row level diff
- Cell-level change detection
- Smart row matching (similarity-based)
- Robust error handling (no crashes on bad input)
- JSON output (easy to integrate with other tools)
---

## Usage

python main.py before.xlsx after.xlsx

---

## Output Format (JSON)
```
{
  "sheet_added": ["Sheet3"],
  "sheet_deleted": ["Sheet1"],
  "sheet_modified": {
    "Sheet2": {
      "columns": {
        "added": ["D"],
        "deleted": ["B"]
      },
      "rows": {
        "added": [10, 11],
        "deleted": [3],
        "modified": [
          ["A5", "A5"],
          ["C7", "E7"]
        ]
      }
    }
  }
}
```
