# PDFREADER

Client-side shipment PDF reader for GitHub Pages.

## What it does

- Uploads a shipment PDF in the browser
- Uses `pdf.js` and `tesseract.js` to OCR each page client-side
- Totals quantity shipped by part number
- Builds:
  - a CSV with overall part totals
  - an Excel workbook with:
    - `Part Totals`
    - `PO Part Totals`
    - `Detail`
- Downloads everything as a ZIP so the browser can ask where to save it

## GitHub Pages

This repo is now designed for static hosting.

To publish it with GitHub Pages:

1. Push the repo to GitHub.
2. In GitHub, open `Settings` -> `Pages`.
3. Under `Build and deployment`, choose `Deploy from a branch`.
4. Select the `main` branch and `/ (root)` folder.
5. Save.

Your public app will then be served from your GitHub Pages URL.

## Notes

- Best results are usually in Chrome or Edge.
- OCR runs in the browser, so large PDFs can take a while.
- No server is required for the GitHub Pages version.
