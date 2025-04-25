# Excel VBA Utilities – Portfolio Project

Author: Kyle Hsu  
Date: April 21, 2025


## 📌 Overview

This project contains a modular collection of reusable **Excel VBA tools** that automate formatting, exporting, and sheet operations in real-world reporting scenarios. Designed with scalability and usability in mind, these utilities aim to reduce manual tasks, improve formatting consistency, and simplify workflows for both technical and non-technical users.

Each module solves a common problem encountered in business environments such as scheduling, financial reporting, or operations.

---

## 🧩 Modules

### 🔹 1. ExportUtilities.bas
> Automates export of ranges or sheets to image and PDF formats.

- `ExportSheetAsPNG`: Export any range as a high-resolution PNG (ideal for messaging platforms like LINE).
- `ExportSheetsAsPDF`: Merge multiple sheets into a single PDF with optional timestamp and auto-opening.
- Includes folder creation, timestamping, and error handling.

---

### 🔹 2. FormattingUtilities.bas
> Applies consistent and professional formatting to Excel data.

- `ApplyHeaderFormatting`: Styles header rows with standardized font, size, color.
- `ApplyAlternatingRowColors`: Applies alternating background colors to enhance readability.
- `ApplyDateColumnFormatting`: Highlights weekends in date headers.
- `ApplyConditionalFormatting`: Highlights specified values dynamically.
- `FormatAsTable`: Creates visually structured, bordered tables with alternating row colors.
- `StandardizeColumnWidths`: Applies custom or uniform column widths.

---

### 🔹 3. SheetUtilities.bas
> Handles common sheet-level operations with automation.

- `SheetExists`: Checks if a sheet exists in the workbook.
- `CreateOrReplaceSheet`: Deletes and replaces existing sheets with optional confirmation.
- `HideUnusedRows`: Hides empty or zero-value rows based on a column range.
- `CopyRangeAcrossSheets`: Copies a source range to multiple destination sheets and cells.

---

### 🔹 4. UIEnhancements.bas
> Builds a friendly user interface inside Excel for macros.

- `CreateCommandButton`: Adds styled buttons to trigger macros.
- `CreateButtonPanel`: Automatically lays out a grid of buttons with captions and colors.
- `ApplyStandardCellFormatting`: Applies clean, uniform text and fill styling.
- `ApplyStandardBorders`: Adds consistent inside/outside borders to any range.

---

## 🎯 Real-World Problems Solved

- **Saved time** by eliminating repetitive manual formatting and export steps.
- **Improved report consistency** across monthly reports and handover files.
- **Enabled non-technical users** to execute macros via intuitive UI buttons.
- **Simplified teamwork and onboarding** by reducing Excel learning curve.

## 🧠 About the Author

Kyle Hsu  
Background in finance & computer science  
Specialized in automating Excel processes and building user-friendly tools
