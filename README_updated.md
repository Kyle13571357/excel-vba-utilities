# Excel VBA Utilities – Portfolio Project

Author: Kyle Hsu  
Date: April 21, 2025

## 📌 Overview

This project includes a modular set of **Excel VBA tools** for automating reporting workflows, formatting, data validation, and UI enhancements. Designed with clean architecture and reusability in mind, these tools are built to be used in secure and restricted enterprise environments where external scripting languages and add-ins are not permitted.

---

## 🧩 Modules

### 🔹 1. ExportUtilities.bas
Automates export of selected ranges or entire sheets.

- `ExportSheetAsPNG`: Exports selected range to PNG (ideal for messaging platforms).
- `ExportSheetsAsPDF`: Combines multiple sheets into a single PDF.
- Includes automatic folder creation, timestamping, and error handling.

---

### 🔹 2. FormattingUtilities.bas
Applies consistent and professional formatting.

- `ApplyHeaderFormatting`, `ApplyAlternatingRowColors`
- `ApplyDateColumnFormatting`, `ApplyConditionalFormatting`
- `FormatAsTable`, `StandardizeColumnWidths`

---

### 🔹 3. SheetUtilities.bas
Sheet-level utility functions.

- `SheetExists`, `CreateOrReplaceSheet`
- `HideUnusedRows`, `CopyRangeAcrossSheets`

---

### 🔹 4. UIEnhancements.bas
User interface automation for Excel.

- `CreateCommandButton`, `CreateButtonPanel`
- `ApplyStandardCellFormatting`, `ApplyStandardBorders`

---

### 🔹 5. IntervalValidationUtilities.bas
Ensures compliance of shift schedules by reconstructing dynamic formulas.

- `ResetIntervalCheckFormulas`: Entry point that validates and rebuilds two schedule check sheets (T1/T2).
- `FormatCheckSheet`: Reconstructs formulas dynamically for interval validation.
- `InsertFormatColumns`: Inserts necessary columns and applies consistent formatting.
- `ClearBorderStyles`, `SetLeftBorder`: Helper methods for visual formatting.

✅ **Why this matters**:  
Built for shared workbooks in environments where Excel protection disrupts maintainability. This module uses VBA to dynamically reconstruct formulas based on lookup values, ensuring regulatory compliance (e.g., 11-hour rest rule) without relying on locked cells that restrict flexibility.

---

## 🔧 How to Use

1. Open Excel and launch the VBA editor (`Alt + F11`)
2. Import any `.bas` module using `File → Import File…`
3. Call the module’s main procedure from your own macro or attach it to a button
4. Optional: Combine multiple modules for advanced workflows

---

## 💡 Practical Outcomes

- Reduced manual report formatting and exporting workload
- Improved accuracy in shift schedule validation against legal requirements
- Created usable macro interfaces for non-technical colleagues
- Supported compliant workflows within restricted corporate IT environments

---

## 🧠 About the Author

Kyle Hsu  
Former banking professional turned automation enthusiast  
Experienced in building robust internal tools using native VBA  
Focused on bridging the gap between process control and user-friendly design
