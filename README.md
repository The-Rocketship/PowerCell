# PowerCell📊

> **PowerCell** is a lightweight, standalone, single-file PowerShell application that brings the native **Microsoft Excel GUI experience** to any Windows environment—without requiring Microsoft Office to be installed!

![PowerCell Excel Interface](screenshot.png)

---

## 🌟 Key Features

- **Native Windows WPF GUI**: Modern Office dark green (`#107C41`) Ribbon toolbar, Formula Bar, Name Box, and interactive DataGrid with gridlines and selection highlights.
- **Single-File Distribution**: Completely self-contained in `PowerCell.ps1`. Zero external dependencies or installation needed!
- **Reactive Formula Engine**: Dynamic calculation of formulas starting with `=`, including:
  - Aggregates: `=SUM(A1:A10)`, `=AVG(B1:B10)`, `=COUNT(C1:C10)`, `=MIN(...)`, `=MAX(...)`.
  - Math Expressions: `=A1 + B2 * 1.5`, `=(C1 - C2) / 10`.
- **Multi-Format Spreadsheet Support**: Directly open, edit, and save `.csv`, `.tsv`, and `.xlsx` files.
- **Double-Click Portable Execution**: Simply double-click `powercell.cmd` or run `.\PowerCell.ps1` from any PowerShell window.
- **Interactive Editing**: Inline cell editing, Formula Bar synchronization, row/column insertion & deletion, and real-time status bar aggregates (Sum, Average, Count).

---

## 🚀 Quick Start

### Prerequisites
- **Windows 10 / 11** or **Windows Server**
- **PowerShell 5.1+** or **PowerShell 7+** (built into Windows)

### Running PowerCell

```powershell
# Clone or download the repository
cd C:\dev\PowerCell

# Launch PowerCell with a new blank spreadsheet
.\PowerCell.ps1

# Or open an existing CSV / Excel file directly
.\PowerCell.ps1 sample_data.csv
```

> 💡 **Tip:** You can also double-click `powercell.cmd` directly in File Explorer to launch the editor!

---

## 🛠️ Usage & Controls

| Action | Control / Location | Description |
| :--- | :--- | :--- |
| **Edit Cell** | Double-Click cell or select & edit in Formula Bar | Updates raw value or formula |
| **Commit Formula** | Press `Enter` in Formula Bar or Cell | Triggers live formula evaluation across grid |
| **Save File** | `💾 Save (Ctrl+S)` or Quick Access Bar | Saves changes to CSV / XLSX |
| **Open File** | `📂 Open File` | Load CSV, TSV, or XLSX spreadsheets |
| **Insert Row** | `➕ Add Row` | Inserts a new row at selected cursor position |
| **Delete Row** | `➖ Delete Row` | Deletes row at selected cursor position |
| **Insert Column** | `➕ Add Column` | Adds a new column |
| **AutoSum** | `∑ AutoSum` | Automatically generates `=SUM(col1:colN)` |
| **AutoAverage** | `x̄ Average` | Automatically generates `=AVG(col1:colN)` |

---

## 📁 Repository Structure

```
PowerCell/
├── PowerCell.ps1     # 100% Standalone Single-File Application
├── screenshot.png    # GUI Screenshot
---

## 🧪 Testing

Run the included unit tests to verify formula evaluation and format conversion:

```powershell
powershell -ExecutionPolicy Bypass -File .\tests\TestEngine.ps1
powershell -ExecutionPolicy Bypass -File .\tests\TestFormats.ps1
```

---

## 📄 License

Distributed under the [MIT License](LICENSE). Free for personal and commercial use.
