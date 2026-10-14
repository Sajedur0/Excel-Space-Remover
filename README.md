# 🧹 Excel Space Cleaner

A simple, powerful web app built with **Streamlit** that detects and removes unwanted leading and trailing spaces from your Excel and CSV files — so your data stays clean and consistent.

Leading and trailing spaces are one of the most common causes of broken lookups, failed `VLOOKUP`s, duplicate rows, and messy exports. Excel Space Cleaner finds every one of them, highlights them, and cleans them in a single click.

---

## ✨ Features

- **📂 Multiple input sources** — Upload a file from your computer, load it from a direct URL, or try it out with built-in sample data.
- **🔍 Automatic space detection** — Scans every cell and reports exactly how many cells, rows, and columns are affected.
- **👁️ Visual highlighting** — Cells containing leading or trailing spaces are highlighted in red before cleaning.
- **📋 Detailed inspection** — Expand a report to see each affected cell with its row, column, and the exact value.
- **🧹 Flexible cleaning options** — Trim, left-trim, right-trim, or remove **all** spaces (including those between words).
- **📥 Export results** — Download the cleaned data as `.xlsx` or `.csv`.
- **📊 Before/after summary** — See how many spaces were found and how many were removed.

---

## 🚀 Getting Started

### Prerequisites

- [Python](https://www.python.org/downloads/) 3.8 or newer
- `pip`

### Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/Sajedur0/Excel-Space-Remover.git
   cd Excel-Space-Remover
   ```

2. **(Optional but recommended) Create a virtual environment**

   ```bash
   python -m venv venv
   ```

   - Windows:
     ```bash
     venv\Scripts\activate
     ```
   - macOS / Linux:
     ```bash
     source venv/bin/activate
     ```

3. **Install the dependencies**

   ```bash
   pip install -r requirements.txt
   ```

---

## ▶️ Running the App

**Option 1 — One-click (Windows):** Double-click `run.bat`

**Option 2 — Command line:**

```bash
streamlit run main.py
```

The app will open automatically in your browser at [http://localhost:8501](http://localhost:8501).

---

## 📖 How to Use

1. **Load your file** — Choose an input source (computer, URL, or sample data). You can optionally pick a specific sheet.
2. **Review the analysis** — Check the metrics showing how many cells, rows, and columns contain unwanted spaces.
3. **Inspect the preview** — Affected cells are highlighted in red so you can see exactly what will change.
4. **Choose a cleaning option** and output format, then click **Clean & Download**.
5. **Download** your cleaned file.

### Cleaning Options

| Option | What it does |
| --- | --- |
| Remove leading & trailing spaces (Trim) | Strips spaces from both ends of each cell |
| Remove leading spaces only (LTrim) | Strips spaces from the start of each cell |
| Remove trailing spaces only (RTrim) | Strips spaces from the end of each cell |
| Remove all spaces | Removes every space, including those between words |

---

## 📁 Supported Formats

- `.xlsx` (Excel)
- `.xls` (Excel)
- `.csv` (Comma-separated values)

> Maximum file size: **50 MB**

---

## 🗂️ Project Structure

```
Excel-Space-Remover/
├── main.py            # Streamlit application
├── requirements.txt   # Python dependencies
├── run.bat            # Windows one-click launcher
├── README.md
└── LICENSE
```

---

## 🛠️ Tech Stack

- [Streamlit](https://streamlit.io/) — Web UI
- [pandas](https://pandas.pydata.org/) — Data processing
- [openpyxl](https://openpyxl.readthedocs.io/) — Excel read/write

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome. Feel free to fork the repository, open an issue, or submit a pull request.

---

## 📄 License

This project is licensed under the [Apache License 2.0](LICENSE).
