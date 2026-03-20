# 📊 Grant Financial Report Automation

A Python automation tool that consolidates university grant financial data from multiple Excel exports into a single, professionally formatted executive summary report — eliminating manual copy-paste work and reducing reporting time.

![Python](https://img.shields.io/badge/Python-3.11-blue?logo=python)
![pandas](https://img.shields.io/badge/pandas-2.x-150458?logo=pandas)
![openpyxl](https://img.shields.io/badge/openpyxl-Excel-green)
![License](https://img.shields.io/badge/License-MIT-green)

---

## 🎯 The Problem

Grant financial data at the university level is typically scattered across multiple systems:

- **Budget files** — transaction-level details (expenses, encumbrances)
- **Project files** — metadata such as PI names, sponsor names, and grant dates

Manually combining these exports for weekly reporting was time-consuming and prone to copy-paste errors, especially when Project IDs appear as text in one file and numbers in another.

---

## ✅ The Solution

This script automates the full consolidation pipeline:

1. **Ingest** — reads multiple raw Excel dumps from a `/data` directory
2. **Clean** — normalizes data types, handles text-vs-numeric Project ID mismatches
3. **Merge** — performs a left join on Project ID to unify budget and project metadata
4. **Format** — produces a polished Excel output with:
   - Professional headers and alternating row colors
   - Currency formatting for financial columns
   - Conditional logic for handling missing or incomplete data
   - Column-width auto-fit for readability

---

## 🛠️ Tech Stack

| Library | Purpose |
|---------|---------|
| `pandas` | Data ingestion, cleaning, and merging |
| `openpyxl` | Excel output formatting and styling |
| `os` / `sys` | File path handling and script execution |
| `warnings` | Suppressing non-critical pandas warnings |

---

## 🚀 Getting Started

### Prerequisites

- Python 3.9+
- pip

### Installation

```bash
git clone https://github.com/SkyMoon56/Grant-Financial-Report-Automation.git
cd Grant-Financial-Report-Automation
pip install -r requirements.txt
```

### Usage

1. Place your raw Excel exports in the `Scripts/` directory (or update the file paths in the script)
2. Run the main script:

```bash
python Scripts/report_automation.py
```

3. The formatted output will be saved as an Excel file in the same directory

---

## 📂 Project Structure

```
Grant-Financial-Report-Automation/
├── Scripts/
│   ├── report_automation.py   # Main consolidation and formatting script
│   └── (additional scripts)   # Supporting utilities
├── requirements.txt
└── README.md
```

---

## 📈 Sample Output

The generated Excel report includes:

- One row per active grant/project
- Merged columns: Project ID, PI Name, Sponsor, Start/End Dates, Budget, Expenses, Encumbrances, Balance
- Color-coded rows for quick visual scanning
- Currency-formatted financial columns (e.g., `$124,530.00`)
- Highlighted rows where data is missing or incomplete

---

## 💡 Key Technical Decisions

- **Left Join on Project ID** — preserves all budget records even when project metadata is missing, flagging gaps rather than silently dropping rows
- **Type normalization** — coerces Project ID to string before merging to prevent mismatches between numeric and text representations
- **openpyxl styling** — applies formatting post-merge for full control over cell styles without requiring xlwings or Excel itself

---

## 📝 License

MIT — free to use, fork, and adapt.

---

## 🤝 Contact

**Sky Moon** — [sky.moon7567@gmail.com](mailto:sky.moon7567@gmail.com) | [LinkedIn](https://linkedin.com/in/sky-moon/)
