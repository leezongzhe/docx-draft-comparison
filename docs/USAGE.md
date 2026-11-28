# Usage

## Install

Install Python 3.11 or later (tested through Python 3.14) and either Microsoft Word 2013 or later, or LibreOffice. From the repository root, run:

```bat
scripts\setup.bat
```

## Graphical interface

Run `scripts\run_tool.bat`, select the two DOCX drafts, choose an output PDF path, and select a converter. `auto` tries Microsoft Word first and then LibreOffice.

The highlight legend is:

- Red: deleted
- Green: added
- Orange: replaced
- Purple outline: likely moved

## Command line

```powershell
.venv\Scripts\python.exe src\app.py --cli `
  --v1 "draft-1.docx" `
  --v2 "draft-2.docx" `
  --output "comparison.pdf" `
  --label1 "Version 1" `
  --label2 "Version 2" `
  --backend auto
```

Use `--backend word` to require Microsoft Word or `--backend libreoffice` to require LibreOffice.

## Troubleshooting

- If setup cannot find Python, install Python and ensure the `py` launcher is available.
- If conversion fails, confirm that Word or LibreOffice can open both input files.
- For Word 2013, open Word manually once and complete activation or first-run setup. The tool automatically tries both 64-bit and 32-bit Windows automation. If Windows still reports that the type library is not registered, repair the Office installation.
- Floating objects and text embedded in images are preserved visually but may not receive word highlights.
- Keep document and output files outside this repository; DOCX and PDF files are intentionally ignored.

## Privacy

Comparison happens on the local computer. The application does not upload or log document content.
