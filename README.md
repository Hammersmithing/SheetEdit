# SheetEdit

A lightweight .xlsx spreadsheet editor built with PySide6 and openpyxl. Designed for film production workflows — especially call sheets — but works as a general-purpose spreadsheet tool.

## Features

- **Full spreadsheet editing** — cell formatting, fonts, colors, borders, merges, row/column resizing
- **Google Sheets-style navigation** — direct typing, arrow keys, tab/enter movement
- **Templates** — built-in and custom templates for quick starts
- **Snippets** — save and reuse formatted cell blocks; right-click to save/insert
- **Snippet Composer** — build documents by ordering and stacking snippets (File > Build from Snippets)
- **AI-powered call sheet import** — drag-and-drop images, PDFs, or spreadsheets; extracts data via macOS Vision OCR + Ollama and fills call sheet templates
- **Editable import guide** — human-readable markdown file that controls how the AI maps data to snippets (File > Edit Import Guide)
- **Cell rules** — validation constraints (not empty, number only, max length, etc.) with visual indicators and a rules editor
- **Print preview** — portrait/landscape toggle, Cmd+P shortcut
- **Unsaved changes prompt** — warns before quitting with unsaved work

## Requirements

- macOS
- Python 3.10+
- PySide6, openpyxl
- Optional: pyobjc-framework-Vision (for OCR import), requests (for Ollama), Pillow

## Running

```bash
python sheetedit.py
```

Or install the bundled app to `/Applications/SheetEdit.app` via PyInstaller.

## Future Development

### Headless Workflow Engine

The long-term vision is a pipeline that can chain operations across multiple spreadsheets without opening any GUI — enabling full automation of linked document workflows.

**Concept:** A workflow config (JSON or YAML) that describes a sequence of operations:

```yaml
workflow: "Daily Call Sheet Build"
steps:
  - read: "crew_list.xlsx"
    extract: crew
  - read: "script_breakdown.xlsx"
    extract: scenes, cast
  - compose:
      snippets: [Header, Scenes, Cast, Crew, KeyPersonnel, Notes]
  - fill:
      guide: "import_guide.md"
      sources: [crew, scenes, cast]
  - save: "Day4_CallSheet.xlsx"
  - copy:
      from: "budget.xlsx!B3"
      to: "daily_report.xlsx!F12"
```

**Building blocks already in place:**
- Snippet system (reusable formatted cell blocks)
- Import guide (AI-driven data mapping)
- Cell rules (validation constraints)
- Cell conversion helpers (read/write cells independent of UI)

**Planned capabilities:**
- CLI mode: `sheetedit --run workflow.json`
- Cross-document cell references and data linking
- Scheduled/automated runs
- Visual workflow editor inside the app
- Batch processing across multiple files
