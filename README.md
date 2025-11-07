# Invoice Processor

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
├── test/                    # Test XMLs and ZIP samples, NOT ON REPO (used for local tests only)
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
## 🧪 Usage
### Desktop:
1. Run `python src/gui.py`.
2. Use test files from `/test` folder.
3. Confirm Excel file matches expected data.
4. You should see something like this:
   
<p align="center">
  <img width="398" height="382" alt="image" src="https://github.com/user-attachments/assets/d6b55596-babd-49fe-99d5-5f83147c4511" />
</p>


## 🌐 Live Demo
Try the **web version** of Invoice Processor here:  
🔗 [**facturador.duckdns.org**](http://facturador.duckdns.org)

*(Note: performance may vary depending on server uptime and network speed.)*

---

## 🧠 To-Do List
- [X] Add support for **CFDI Retenciones**  
- [X] Implement **batch progress bar** in PyQt5 GUI  
- [ ] Add **dark mode** for Streamlit interface 🌙  
- [X] Create **Dockerfile** for easy deployment 🐳  
- [ ] Add **unit tests** for XML extractors  
- [ ] Include **user manual (PDF)** inside repository 📘  


## Next steps
- Correcting website ico
- Recreating visual interface
- Creating different testcases
