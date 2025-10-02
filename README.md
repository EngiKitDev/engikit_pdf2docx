# EngiKit — PDF → DOCX Converter (GUI)

A simple open-source PDF → DOCX converter with a graphical interface based on `tkinter` and using the `pdf2docx` library.  
This project is part of the **EngiKitDev** initiative — a toolkit useful for engineers and developers.

> **Note on naming:**  
> This project is independent and not affiliated with the official [`pdf2docx`](https://pypi.org/project/pdf2docx/) library.  
> To avoid confusion, the main script is named `engikit_pdf2docx.py`. The project *uses* the `pdf2docx` library under the terms of the AGPL license.

## Features
- GUI to select a PDF file and an output folder.  
- Option to convert all pages or a selected page range.  
- Progress logging inside the application window.

## Prerequisites
- Python 3.8+  
- the `pdf2docx` library (listed in `requirements.txt`)  
- On some Linux distributions you may need to install system packages for tkinter (e.g. `python3-tk`).

> **Note:** The script launches a graphical interface (tkinter). Running the GUI requires a display environment (local desktop).

## Installation (local)
1. Clone the repo:
```bash
git clone https://github.com/engikitdev/engikit_pdf2docx.git
cd pdf2docx
```
2. Create a virtual environment and install dependencies:
```bash
python -m venv venv
```
macOS/Linux
```bash
source venv/bin/activate
```
Windows (PowerShell)
```bash
.\venv\Scripts\Activate.ps1
```
```bash
pip install --upgrade pip
pip install -r requirements.txt
```
Running (GUI)
```bash
python engikit_pdf2docx.py
```
>When started, a window opens where you choose the PDF, the output folder and page range. The result will be saved as *_converted.docx.

## Files in the repo

- engikit_pdf2docx.py — main GUI script.

- requirements.txt — Python dependencies.

- LICENSE — MIT License.

- README.md — this file.

## License

MIT — see the LICENSE file.
This project also depends on the pdf2docx
