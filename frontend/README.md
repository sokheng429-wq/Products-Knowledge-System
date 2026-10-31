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

## Data and refresh

`public/catalog.json` and `public/images/` are generated from `Japan_Product_Knowledge_Catalog.xlsx`. Prices reflect 30 October 2026, not live inventory. Six products have no embedded photo. Units and product descriptions are preserved as supplied, even where a unit appears unusual. Expiry labels use the current Bangkok date and the workbook's 45-day warning period; verify the actual package's expiry.

To replace the catalog, run `scripts/extract_catalog.py` with Python, openpyxl, and Pillow installed:

```powershell
python scripts/extract_catalog.py "path/to/Japan_Product_Knowledge_Catalog.xlsx"
npm test
npm run build
```

Keep the `Product_Knowledge` and `Notes` worksheets and their current column order. The importer maps embedded pictures by their anchored workbook row and preserves barcode strings. This is a static frontend: editing Excel does not automatically refresh the app. There is no database, login system, or browser-based Excel editor.

## Verification

Five automated tests cover row/image counts and file existence, barcode normalization and UPC/EAN equivalence, duplicate handling, multilingual search with combined filters, and expiry boundaries. Production build passed. Browser checks covered desktop barcode lookup and mobile category filtering, manual barcode lookup, photos and details. Camera recognition and physical USB scanner input require testing on real devices; universal device compatibility is not guaranteed.

## Publication

A private Sites project was registered, but uploading the catalog and publishing were blocked by automatic approval review. No app source or catalog was uploaded. The `.openai/hosting.json` identity is retained for publication after approval. Publishing transfers the catalog and its images to the hosting service. Do not make it public without choosing the intended audience.

Implementation references: [React](https://react.dev/learn/build-a-react-app-from-scratch), [Vite](https://vite.dev/guide/), [ZXing browser](https://github.com/zxing-js/browser).
