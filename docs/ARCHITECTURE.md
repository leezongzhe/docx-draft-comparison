# Architecture

## Processing pipeline

1. `src/app.py` collects file paths and options from the GUI or CLI.
2. `src/comparator.py` validates the inputs and converts each DOCX to PDF.
3. PDF text and coordinates are extracted into pages, lines, and blocks.
4. Blocks are matched using normalized text, document order, geometry, and row context.
5. A document-level sequence comparison identifies added, deleted, replaced, and moved text.
6. Original pages are rendered side by side and highlight overlays are drawn into the output PDF.

## Matching details

The matcher keeps decimal values intact, handles multi-column reading order, and uses nearby row labels and geometry to disambiguate repeated values in tables. Matching spans block and page boundaries so layout-only reflow does not become a false content change.

## Modules

- `app.py`: argument parsing, desktop interface, progress reporting, and error presentation
- `comparator.py`: conversion, extraction, alignment, highlighting, and PDF composition
- `word_to_pdf.ps1`: isolated Microsoft Word automation helper

For Word 2013 compatibility, the converter probes and retries both 64-bit and 32-bit PowerShell. This handles 32-bit Office installations on 64-bit Windows, where the Word COM type library may not be visible from the default PowerShell architecture.

## Data handling

Input documents are opened read-only and processed locally. Temporary conversion files are removed after the comparison. The repository excludes documents and generated outputs by default.
