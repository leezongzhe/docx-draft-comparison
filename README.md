# DOCX Draft Comparison

A local Windows tool that compares two DOCX drafts and creates a side-by-side PDF with word-level change highlights.

The repository contains only the reusable tool, tests, and documentation. It does not contain source documents, sample documents, generated PDFs, or user output. Those file types are excluded by `.gitignore`.

## Features

- Side-by-side page rendering that preserves the original DOCX layout
- Word-level added, deleted, and replaced highlights
- Moved-passage detection
- Multi-column reading-order handling
- Table-aware matching for repeated numeric values
- Microsoft Word and LibreOffice conversion backends
- Local processing with no document upload or content logging

## Requirements

- Windows 10 or later
- Python 3.11 or later (tested through Python 3.14)
- Microsoft Word 2013 or later, or LibreOffice, for DOCX-to-PDF conversion

## Quick start

1. Run `scripts\setup.bat`.
2. Run `scripts\run_tool.bat`.
3. Choose two DOCX files and an output PDF location.

Command-line use:

```powershell
.venv\Scripts\python.exe src\app.py --cli --v1 "draft-1.docx" --v2 "draft-2.docx" --output "comparison.pdf"
```

See [docs/USAGE.md](docs/USAGE.md) for options and troubleshooting. See [docs/ARCHITECTURE.md](docs/ARCHITECTURE.md) for implementation details.

Word 2013 compatibility includes automatic retries through both 64-bit and 32-bit Windows automation, which supports the common combination of 64-bit Windows with 32-bit Office 2013.

## Repository layout

```text
src/       Application and comparison engine
scripts/   Windows setup and launcher scripts
tests/     Automated regression tests using synthetic text only
docs/      Usage and architecture documentation
```

## Privacy

Documents are read locally. The application does not send document content to a server. Keep source files and generated comparisons outside the repository, and do not force-add ignored document or output files.
