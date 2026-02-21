# Myanmar OCR Tool - User Guide

This guide explains how to use:

- `ocr_folder_ui.py` (desktop OCR UI)
- `combine_docx_master.py` (merge many DOCX files into one Master DOCX)

## 1) Requirements

- Windows with Python installed
- Tesseract OCR installed and available in `PATH`
- Myanmar language data installed in Tesseract (`mya`)

Check language list:

```powershell
tesseract --list-langs
```

You should see `mya` in the output.

For DOCX output, install:

```powershell
python -m pip install --user python-docx
```

## 2) Run the OCR UI

Option A:

- Double-click `run_ocr_ui.bat`

Option B:

```powershell
python -B ocr_folder_ui.py
```

## 3) OCR UI Fields

- **Input folder**: folder containing images (`png/jpg/jpeg/tif/tiff/bmp/webp`)
- **Output folder**: where OCR results are written
- **Language**: default `mya` (you can use `mya+eng`)
- **PSM**: page segmentation mode (default `6`)
- **Output**:
  - `txt`: plain text
  - `md`: markdown text
  - `docx`: Word document
  - `pdf`: searchable PDF (best layout preservation)
  - `hocr`: layout-rich HTML OCR format
- **Combine into a single file**:
  - For `txt/md/docx`: combines all pages into one file
  - For `pdf/hocr`: each page is generated separately
- **Filename contains**: only process files whose names include this text
- **Page range**: filters by trailing page number in filename
  - Examples: `3-10`, `25`, `40-`, `1-5,10,20-30`

## 4) Buttons

- **Start OCR**: begin batch OCR
- **Check Tesseract**: validates Tesseract and shows language list in log
- **Myanmar Book Preset**:
  - `Language = mya`
  - `PSM = 6`
  - `Output = pdf`
  - `Combine = off`

## 5) Output Naming Rules

- `txt + combine on` -> `ocr_combined.txt`
- `md + combine on` -> `ocr_combined.md`
- `docx + combine on` -> `<input_folder_name>.docx`
- `docx + combine off` -> one DOCX per image (same base filename)
- `pdf/hocr` -> one output per image

## 6) Combine Many DOCX into One Master

Use `combine_docx_master.py` to merge all `.docx` in a folder.

Basic:

```powershell
python -B combine_docx_master.py "D:\path\to\docx_folder"
```

Default output:

- `<folder_name>_Master.docx` inside that folder

Custom output:

```powershell
python -B combine_docx_master.py "D:\path\to\docx_folder" --output "D:\path\Master.docx"
```

Notes:

- Files are merged in natural order (`1,2,3,...,10,...`)
- Temporary Word lock files (`~$*.docx`) are ignored

## 7) Recommended Workflow (Myanmar Books)

1. Open `ocr_folder_ui.py`
2. Click **Myanmar Book Preset**
3. Set Input/Output folders
4. (Optional) Set `Page range` like `3-10`
5. Start OCR
6. If you created many DOCX files, merge with `combine_docx_master.py`

## 8) Troubleshooting

- **`tesseract command not found`**
  - Add Tesseract install folder to `PATH`
  - Restart terminal/UI
- **No text or weak results**
  - Try `Language = mya+eng`
  - Try different `PSM` (`4`, `6`, `11`)
  - Use clearer scans
- **DOCX error**
  - Install `python-docx`:
    - `python -m pip install --user python-docx`

