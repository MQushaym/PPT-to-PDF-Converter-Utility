# PPT to PDF Converter Utility

This project is a simple Python command-line utility to batch-convert PowerPoint files (`.ppt` and `.pptx`) into PDF files.

The main feature of this tool is its use of **LibreOffice** (in headless mode) to perform the conversion. This means:
* It does not require Microsoft PowerPoint to be installed.
* It works cross-platform (Windows, macOS, Linux).

---

## ⚙️ Requirements

This tool does not require any external Python libraries, but it has one system-level dependency:

* **Python 3.6+** (uses built-in libraries like `argparse` and `pathlib`).
* **LibreOffice** must be installed on your system.

The script will attempt to find the `soffice` executable path automatically.

---

## 🚀 Usage

1.  Ensure LibreOffice is installed on your system.
2.  Run the script from your terminal, passing the path to the folder containing your files.

### Basic Usage

This will convert files only in the specified folder (not subfolders).

```bash
python ppt_to_pdf_lo.py "/path/to/your/folder"
```
Recursive Usage
To search inside the specified folder and all its subfolders, use the --recursive flag:
```
python ppt_to_pdf_lo.py "/path/to/your/folder" --recursive
```
Manually Specifying soffice Path
If the script cannot find your LibreOffice installation automatically, you can provide the path manually:
```
python ppt_to_pdf_lo.py "/path/to/your/folder" --soffice "/custom/path/to/soffice"
```
