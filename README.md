# CFDI Extractor

A Python-based tool for accountants to extract and consolidate data from Mexican CFDI XML files (Ingreso/Egreso, Pago, Nómina).  
Supports both **desktop GUI** (PyQt5) and **web** (Streamlit) versions, with Excel export.

---

## 📂 Folder Structure

```
.
├── assets/                  # Static assets (icons, etc.)
├── src/                     # Application source code
│   ├── extractors.py        # Functions to parse XML for each CFDI type and save to Excel
│   ├── identifier.py        # Detects CFDI type from XML
│   ├── processor.py         # Routes XML to correct parser, updates counters, saves results
│   └── gui.py               # Desktop PyQt5 interface
├── test/                    # Test XMLs and ZIP samples
├── venv/                    # Python virtual environment (not committed)
├── .gitignore               # Git ignore rules
└── requirements.txt         # Python dependencies
```

---

## 🚀 Features
- Parse **Ingreso/Egreso, Pago, and Nómina** CFDI XML files.
- Automatic type detection and organized Excel sheets per type.
- Desktop GUI for offline use.
- Web interface for online processing.
- Excel output with professional headers and multiple sheets.
- Handles multiple ZIP files at once.

---

## 📦 Installation

### 1. Clone the repository
```bash
git clone https://github.com/yourusername/cfdi-extractor.git
cd cfdi-extractor
```

### 2. Create a virtual environment
```bash
python -m venv venv
```
Activate it:
- **Windows**:
  ```bash
  venv\Scripts\activate
  ```
- **macOS/Linux**:
  ```bash
  source venv/bin/activate
  ```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

---

## 🖥 Running the Desktop Version (PyQt5)
```bash
python src/gui.py
```
- Select input and output folders in the GUI.
- Choose the Excel filename.
- Click **"Run"** to process and generate the Excel file.

---

## 📤 Building the Desktop Executable
You can bundle the PyQt5 app into a standalone `.exe`:

```bash
pyinstaller src/gui.py --onefile --noconsole --icon=assets/logo.ico
```
- Output will be in `dist/`.
- Replace `assets/logo.ico` with your own icon if needed.

---
## 🧪 Testing
### Desktop:
1. Run `python src/gui.py`.
2. Use test files from `/test` folder.
3. Confirm Excel file matches expected data.

### Web:
1. Run `streamlit run app_streamlit.py`.
2. Upload the same test ZIPs.
3. Compare the Excel output with desktop output.

---
