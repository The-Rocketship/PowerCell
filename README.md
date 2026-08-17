# PowerCell Excel 📊

> **PowerCell** is a high-performance, lightweight, single-file PowerShell application that delivers an authentic **Microsoft Excel GUI experience** to any Windows environment—without requiring Microsoft Office!

![PowerCell Excel Interface](screenshot.png)

---

## 🌟 Key Features

- **Streamlined Office Dark Theme Ribbon**: Authentically styled Office dark ribbon organized into logical section cards (*Clipboard*, *Font*, *Alignment*, *Number & Currencies*, *Cells*, *Editing*).
- **Multi-Currency Formatting**: Instant currency selection for **USD (`$`)**, **EUR (`€`)**, **GBP (`£`)**, and **JPY (`¥`)**, plus **Percentage (`%`)**.
- **Text Positioning & Alignment**: Full horizontal (**Left**, **Center**, **Right**) and vertical (**Top**, **Middle**, **Bottom**) cell text alignment.
- **Multi-Cell Range Selection**: Click & drag, `Shift`+Click, or `Ctrl`+Click to select ranges (e.g. `B2:D10`).
- **Live Range Aggregates**: Real-time calculation of **Sum**, **Average**, and **Count** across selected multi-cell ranges displayed in the green Office Status Bar.
- **60 FPS UI Virtualization**: High-performance WPF row and column virtualization (`VirtualizationMode="Recycling"`) for smooth scrolling and instant sub-millisecond cell rendering.
- **Font & Style Controls**: Toggle **Bold** (`B`), *Italic* (`I`), **Underline**, and alternating **Zebra Table Styles**.
- **Column Sorting & Live Filtering**:
  - **Sort Ascending (A ➔ Z)**: Sort rows dynamically by selected column.
  - **Sort Descending (Z ➔ A)**: Sort rows descending by selected column.
  - **Live Filter**: Instant search row filtering across the entire dataset.
- **Reactive Formula Engine**: Dynamic formula recalculation starting with `=`, supporting:
  - Aggregates: `=SUM(A1:A10)`, `=AVG(B1:B10)`, `=COUNT(C1:C10)`, `=MIN(...)`, `=MAX(...)`.
  - Expressions: `=A1 + B2 * 1.5`, `=(C1 - C2) / 10`.
- **Multi-Format Support**: Directly open, edit, and save `.csv`, `.tsv`, and `.xlsx` files.
- **Single-File Distribution**: Portable script in `PowerCell.ps1` with zero external dependencies. Double-click `powercell.cmd` to run immediately!

---

## 🚀 Quick Start

### Prerequisites
- **Windows 10 / 11** or **Windows Server**
- **PowerShell 5.1+** or **PowerShell 7+**

### Running PowerCell

```powershell
# Navigate to directory
cd C:\dev\PowerCell

# Launch PowerCell with a blank sheet
.\PowerCell.ps1

# Or double-click the batch shortcut
.\powercell.cmd

# Open a CSV or Excel file directly
.\PowerCell.ps1 sample_data.csv
```

---

## 🛠️ Usage & Controls

| Feature | Ribbon Group | Description |
| :--- | :--- | :--- |
| **File / Clipboard** | `Clipboard` | **Save**, **Open**, and **Export to XLSX** |
| **Font Formatting** | `Font` | Font Family, Size, **Bold** (`B`), *Italic* (`I`), **Table Zebra Style** |
| **Text Positioning** | `Alignment` | **Top**, **Mid**, **Bot** vertical align & **Left**, **Center**, **Right** horizontal align |
| **Currencies** | `Number & Currencies` | Quick currency toggle for **$ USD**, **€ EUR**, **£ GBP**, **¥ JPY**, and **%** |
| **Row Management** | `Cells` | **➕ Insert Row** and **➖ Delete Row** at selected cell position |
| **AutoSum & Sorting** | `Editing` | **∑ AutoSum**, **x̄ Average**, **Sort A➔Z / Z➔A**, and **Filter** box |
| **Range Calculation** | Status Bar | Displays live **Sum**, **Average**, and **Count** for selected cell range |

---

## 📄 License

Distributed under the [MIT License](LICENSE). Free for personal and commercial use.
