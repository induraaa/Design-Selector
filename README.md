# Wafer Design Picker Tool 🧩

A desktop application for loading semiconductor wafer map files, automatically detecting their format, previewing designs, and exporting selected designs into clean wafer map outputs.

Built with **Python**, **Tkinter**, **NumPy**, **Pillow**, **openpyxl**, and a fast custom XLSX reader, the tool is designed for engineers who need a quick way to isolate one or more wafer designs and convert them into export-ready maps.

---

## Overview

The Wafer Design Picker Tool loads Excel-based wafer map files and automatically detects the supported SCR map format. It normalises the grid, identifies unique design IDs, counts die per design, and displays the full wafer map visually.

Users can then:

- select one or more designs
- preview the selected output
- export the result as either:
  - **Excel workbook (.xlsx)**
  - **Unicode text file (.txt)**

Selected designs are exported as **Bin 1**, while unselected designs are marked separately for downstream processing.

---

## Key Features ✨

- Automatic detection of multiple wafer map formats
- Fast XLSX loading using direct ZIP/XML parsing
- Interactive wafer map rendering
- Separate preview for selected design output
- Multi-design selection with die counts
- Export selected designs as:
  - Excel workbook
  - text wafer map
- Grid line toggle
- zoom in / zoom out / fit view
- design summary and file statistics
- reset and refresh actions
- clean desktop-style interface

---

## Supported Formats

The tool automatically detects and supports these wafer map layouts:

### Format A
- null = `.`
- edge = `X`
- design values like `1`, `1b`, `1c`

### Format B
- null = `-`
- edge = `X`
- design values like `1`, `2`, `3`

### Format C
- row prefix = `RowData:`
- null = `___`
- design values like `001`, `002`

---

## How It Works

1. Open a wafer map Excel file
2. The program detects the file format automatically
3. The wafer grid is normalised into a common internal structure
4. All detected designs are listed with die counts
5. Select one or more designs from the design panel
6. Preview the selected output
7. Export the result as Excel or text

---

## Interface Summary 🖥️

### File Info
Displays:

- file name
- detected format
- grid dimensions
- total die count
- load time

### Designs Panel
Shows:

- all detected design IDs
- die count for each design
- checkbox selection for export

### Wafer Map
Displays the full wafer map with zoom and pan support.

### Selected Design Preview
Shows only the export view, so you can confirm exactly what will be written out.

### Selection Summary
Displays:

- total selected die
- selected design names

### Export Options
Exports selected designs as **Bin 1** into either Excel or text format.

---

## Export Behaviour 📦

### Excel Export
Creates an `.xlsx` file containing:

- wafer map sheet
- summary sheet
- selected designs marked as **1**
- edge dies preserved as `X`
- unselected designs marked as `x`

### Text Export
Creates a `.txt` file containing:

- `1` for selected designs
- `X` for edge dies
- `.` for null area
- `x` for unselected designs

Supported line endings:

- CRLF
- LF

---

## Tech Stack

- Python
- Tkinter
- NumPy
- Pillow
- openpyxl

---

## Performance Notes

A major feature of this tool is its **fast XLSX reader**, which avoids standard workbook loading for input and instead reads worksheet XML directly from the Excel file archive. This makes large wafer maps much faster to process.

---

## Running the Tool

```bash
python wafer_map_tool.py
