# B’Groceries Product Knowledge

React JSX frontend with Vite. The original Excel workbook is preserved in the parent folder.

## Run locally

Requires Node.js 20.19+ or 22.12+.

```powershell
cd "D:\1.BGroceries\Products Knowledge System\frontend"
npm install
npm run dev
```

Open http://127.0.0.1:5173. The development server is restricted to this computer.

```powershell
npm test
npm run build
npm run preview
```

Production files are generated in `dist`. Serve this directory at the root of an HTTPS site for phone camera access. Opening `index.html` directly from disk is not supported.

## Features

- 797 products, 791 embedded workbook photos, all 8 departments.
- Camera barcode recognition with ZXing, manual entry, and keyboard-style USB scanners. Focus the barcode input or search field and scan; scanner Enter submits the lookup.
- Search by English/Khmer knowledge, product name, barcode, or item code; category and department filters; sorting and progressive loading.
- Product photos, knowledge, original USD prices, units, expiry dates, item codes, and action notes.
- Accessible native dialogs, keyboard controls, mobile/tablet/desktop layouts, and missing-image fallbacks.
- Catalog notes from the workbook's Notes sheet.
- Excel data center: validated `.xlsx` import, import preview, merge by barcode or replace catalog, all/filtered export, optional embedded pictures, and a downloadable blank template.
- Persistent imported catalogs in this browser's IndexedDB. No server upload or cross-device synchronization.
- Grid/list views, clickable expiry/photo counts, status filters, larger touch targets, and a mobile bottom navigation bar.

## Import and export Excel

1. Select **Import Excel**, then choose your original workbook or a file created from **Download a template**.
2. Review the valid/new/matching product counts and sample rows. Fix any listed errors in Excel before continuing.
3. Choose **Update matching barcodes & add new products** (default) or **Replace the catalog**, then apply. Merge updates supplied columns for matching products and preserves omitted columns. An explicitly blank cell clears that field. Missing new pictures retain the existing product's picture in either mode.
4. Select **Export Excel**, choose the entire catalog or the current search/filter results, and optionally include pictures. Select **Prepare Excel file**, then the download link. The explicit download link also supports mobile browsers that require a user gesture.

Supported file format: `.xlsx`, up to 30 MB and 5,000 product rows. `.xls`, `.xlsm`, CSV, and password-protected files are not supported. Required headers are **Barcode** and **Product Name**. Other headers match the supplied workbook/template, including the original Khmer/English multiline headers. Header order may change. Header detection searches the first 20 rows. Barcodes must be 8–14 digits and unique within an import; store barcodes and item codes as Text to preserve leading zeroes. Prices must be non-negative numbers or blank; expiry must be an Excel date or `YYYY-MM-DD`.

PNG, JPEG, and GIF pictures anchored to product rows are read. Excel's newer in-cell image format and external image links are not supported. Exports include the first picture for each product, fitted without stretching. They preserve product data and notes, not the original workbook's exact layout. Formula cells are read from saved results; formulas are never executed. Exported text is stored as literal strings.

Imports remain on this browser/device, including after a page reload. Clearing site data or using a different browser/device starts with the original catalog. Export backups to keep changes. If storage fails, the current catalog stays unchanged. Excel support is loaded only when needed. For a complete shared team catalog across devices, a backend and authenticated hosting would be a separate step.

## Data and refresh

`public/catalog.json` and `public/images/` are generated from `Japan_Product_Knowledge_Catalog.xlsx`. Prices reflect 30 October 2026, not live inventory. Six products have no embedded photo. Units and product descriptions are preserved as supplied, even where a unit appears unusual. Expiry labels use the current Bangkok date and the workbook's 45-day warning period; verify the actual package's expiry.

To replace the catalog, run `scripts/extract_catalog.py` with Python, openpyxl, and Pillow installed:

```powershell
python scripts/extract_catalog.py "path/to/Japan_Product_Knowledge_Catalog.xlsx"
npm test
npm run build
```

The Python regeneration script expects the original `Product_Knowledge` and `Notes` worksheets and column order. It updates the bundled default catalog. Browser imports use header detection instead. Editing an Excel file outside the app does not automatically refresh the app; import it again. There is no server database or login system. Imported workbooks without a `Catalog_Info` price date show an unspecified price date rather than reusing the original date.

## Verification

Eleven automated tests cover the original catalog, barcode matching, multilingual filtering, expiry boundaries, the original workbook import, Excel export/import round trips including embedded images and literal text, malformed files and invalid rows, merge/replace and partial-column behavior, filtered exports and template formatting. Browser checks cover Excel import preview/apply, persistence after reload, downloading the full image-inclusive export, widths of 320/360/390/768/1024/1440 pixels, filters, and grid/list views. The downloaded workbook was re-read and verified to contain 797 products, 791 pictures, and zero validation errors. Camera recognition and physical USB scanner input require testing on real devices; universal device compatibility is not guaranteed. iOS/Android camera scanning requires HTTPS (localhost works on the development computer). The development URL is not reachable from a phone.

## Publication

A private Sites project was registered, but uploading the catalog and publishing were blocked by automatic approval review. No app source or catalog was uploaded. The `.openai/hosting.json` identity is retained for publication after approval. Publishing transfers the catalog and its images to the hosting service. Do not make it public without choosing the intended audience.

Implementation references: [React](https://react.dev/learn/build-a-react-app-from-scratch), [Vite](https://vite.dev/guide/), [ZXing browser](https://github.com/zxing-js/browser).
