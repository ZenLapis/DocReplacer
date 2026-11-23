# DocReplacer

Windows desktop app for preparing branded document folders. It reads the business name from the `FR.A4.1` form, makes a safe copy of the folder, replaces the name and logo across every Word document, and can produce a ZIP and one combined PDF.

- **Single folder** – pick or drop a source folder, review the detected name, enter the new name and logo, prepare.
- **Batch** – paste many names or drop many logos, edit names inline, generate all copies in one go. The PDF for one hotel builds while the next is being copied.
- **PDF merge** – combine any folder into a single PDF. Office documents are converted in parallel through a warm Word/Excel pool, and unchanged files are cached, so repeat runs take seconds.

The source folder is never edited in place. The UI is in Turkish.

## Run

Requires Windows, Python 3.12+, Microsoft Office (for PDF conversion) and the Edge WebView2 runtime.

```powershell
py -m pip install -r requirements.txt
pyw docreplacer.py
```

## Tests

```powershell
py -m unittest -v test_docreplacer
```
