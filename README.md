# Multi-File Data Comparator (CSV / Excel)

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue.svg)](https://www.python.org/)
[![GUI](https://img.shields.io/badge/GUI-CustomTkinter-blueviolet.svg)](https://github.com/TomSchimansky/CustomTkinter)
[![Data Engine](https://img.shields.io/badge/Engine-pandas%20%7C%20openpyxl-green.svg)](https://pandas.pydata.org/)

A high-performance, desktop application and CLI tool designed for multi-file data matching and cross-referencing across CSV, Excel (`.xlsx`, `.xls`), TSV, and TXT files.

Built for fast investigation workflows (e.g., cyber crime investigation, digital arrest record cross-matching, financial audits, bank account and telecom CDR reconciliation), this tool enables you to match records from a single source dataset against hundreds of target files and directories.

---

## 🌟 Key Features

* **Multi-Format Support**: Reads `.csv`, `.xlsx`, `.xls`, `.tsv`, and `.txt` files seamlessly.
* **Leading Zero Preservation**: Automatically loads data as raw strings to retain formatting on phone numbers, bank account numbers, complaint IDs, and national identifiers.
* **Flexible Key Column Matching**:
  * **Single Column Matching**: Match on any single key (e.g., `Account No` or `Phone Number`).
  * **Multi-Column Composite Keys**: Combine multiple columns (e.g., `Account No` + `Bank Name`) for composite index matching.
* **Recursive Folder Search**: Scan target folders and nested subdirectories for target files.
* **Matched File Exporter**: Optionally copy all target files that contain matches into a dedicated destination folder (`comparison_report_Matched_Files/`).
* **Rich Styled Excel Reports**: Generates professional, color-coded `.xlsx` workbooks with:
  1. **Summary Dashboard**: Overview statistics, match percentages, and execution parameters.
  2. **Matched Source Records**: Source rows with matching target file lists, match counts, and exact target values.
  3. **Unmatched Source Records**: Source rows not found in any target dataset.
  4. **Target File Statistics**: File-by-file status, total rows, and match metrics.
* **Dual Interface Options**:
  * **Modern GUI**: Sleek CustomTkinter interface with dark/light mode, real-time stats cards, and tabbed result previews.
  * **CLI & Interactive Terminal**: Automated command-line flag execution or step-by-step terminal wizard.

---

## 📂 Repository Structure

```text
Digital Arrest Scripting/
├── app.py                      # Main CustomTkinter Desktop GUI application
├── comparator_engine.py        # Core data processing, matching algorithm, & Excel generator
├── compare.py                  # Command-Line Interface (CLI) & interactive terminal mode
├── generate_demo_data.py       # Script to generate sample source & target datasets for testing
├── requirements.txt            # Python package dependencies
├── run_gui.bat                 # Windows batch launcher for the GUI app
└── sample_data/                # Created by demo data generator for testing
```

---

## ⚙️ Prerequisites & Installation

### 1. Prerequisites
* **Python 3.8** or higher installed on your system.
* Windows / macOS / Linux compatible.

### 2. Installation
Clone or download this repository, then navigate to the project directory:

```bash
cd "Digital Arrest Scripting"
```

Install the required Python packages:

```bash
pip install -r requirements.txt
```

#### Required Dependencies (`requirements.txt`):
* `pandas >= 2.0.0`
* `openpyxl >= 3.1.0`
* `xlrd >= 2.0.1`
* `customtkinter >= 5.2.0`

---

## 🚀 Quick Start Guide

### Option 1: Launch Desktop GUI (Recommended)

#### On Windows:
Double-click `run_gui.bat` or run:
```cmd
run_gui.bat
```

#### On Any Platform (Python command):
```bash
python app.py
```

---

### Option 2: Command-Line Interface (CLI)

#### Interactive Terminal Wizard:
Run `compare.py` without arguments to launch an interactive wizard:
```bash
python compare.py
```

#### Automated CLI Execution:
Pass arguments directly for automated scripts or batch processing:
```bash
python compare.py -s sample_data/source_complaints.xlsx -t sample_data/targets -k "Account No" -o my_report.xlsx -r -c
```

#### CLI Parameters:
| Flag | Long Flag | Description | Default |
| :--- | :--- | :--- | :--- |
| `-s` | `--source` | Path to source CSV/Excel file | *Required for CLI* |
| `-t` | `--target` | Path to target folder or target file | *Required for CLI* |
| `-k` | `--keys` | Comma-separated key column names (e.g. `Account No` or `Account No,Phone Number`) | *Required for CLI* |
| `-o` | `--output` | Path to output Excel report (`.xlsx`) | `comparison_report.xlsx` |
| `-r` | `--recursive` | Search target directory subfolders recursively | `False` |
| `-c` | `--copy-matched` | Copy matched target files into a separate folder | `False` |

---

## 🧪 Testing with Demo Data

Generate realistic sample cyber crime / bank / telecom datasets to test the application immediately:

1. Run the demo data generator script:
   ```bash
   python generate_demo_data.py
   ```
   This creates a `sample_data/` directory containing:
   * `sample_data/source_complaints.xlsx` (Source cyber complaints file)
   * `sample_data/targets/bank_branch_mumbai.csv`
   * `sample_data/targets/bank_branch_delhi.xlsx`
   * `sample_data/targets/telecom_records.csv`

2. Launch the app (`python app.py` or `python compare.py`).
3. Select `sample_data/source_complaints.xlsx` as **Source File**.
4. Check `Account No` or `Phone Number` as the **Key Column**.
5. Select `sample_data/targets` as the **Target Folder**.
6. Run the comparison to view the dashboard and generated report!

---

## 🖥️ Desktop GUI Walkthrough

1. **Source File Selection**: Click **Browse File** and select your primary CSV or Excel document.
2. **Key Column Selection**: Available header columns will populate automatically. Check one or multiple columns to define the matching keys.
3. **Target Folder/Files Selection**:
   * Click **📁 Browse Folder** to process a directory of target files.
   * Toggle **Search Subfolders Recursively** if target files are organized in subdirectories.
   * Alternatively, click **📄 Select Files** to select specific target files.
4. **Copy Matched Target Files**: Keep **Copy matched target files to a separate folder** checked to automatically collect all matching target files in a clean folder next to your Excel report.
5. **Run Comparison**: Click **🚀 START COMPARISON**. The multi-threaded engine will process the datasets and update the live stat cards and tabs.
6. **View Results**: Explore matched records, unmatched records, and target file stats directly inside the GUI or open the generated Excel file.

---

## 📊 Generated Excel Report Breakdown

The exported Excel file (`comparison_report.xlsx`) is auto-formatted with custom header styling, auto-adjusted column widths, and gridlines:

| Tab Name | Description |
| :--- | :--- |
| **Summary Dashboard** | Key metrics summary (Total source records, match counts, match percentage, target files checked, execution parameters). |
| **Matched Records** | Detailed list of source records that were found in target files. Contains additional columns: `Matched Target Files`, `Match Count`, and `Match Details`. |
| **Unmatched Records** | Source records that had no matches in any of the target files checked. |
| **Target Files Stats** | List of all target files processed, showing row counts, status (OK/Error), and matched source record count per file. |

---

## 🛠️ Data Matching Engine Details

* **Whitespace & Formatting Stripping**: Automatically strips outer whitespace and quotes from string values (`norm_val`).
* **Header Auto-Mapping**: Flexible case-insensitive header matching maps source key columns to target file columns even if header casing or minor punctuation differs.
* **String Preservation**: Reads all numerical values as strings using pandas `dtype=str`, avoiding loss of leading zeroes (e.g., `000123456789` is preserved instead of becoming `123456789`).

---

## 📄 License

This project is open-source and available under the standard MIT License.
