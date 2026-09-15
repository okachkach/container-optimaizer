# Container Optimizer

Container Optimizer is a browser-based React app for building shipment packing
plans. It lets you import a product catalog from Excel, choose carton quantities,
generate container loads by weight and CBM capacity, adjust the generated
containers manually, and export packing lists for sharing or later editing.

The app is designed for container shipment planning where each product has
carton quantity, pieces per carton, gross weight, net weight, CBM, unit price,
marks, description, and optional photos.

## Main Features

- Import product catalogs from `.xlsx` files.
- Extract product photos embedded in Excel workbooks.
- Add products manually when they are missing from the catalog.
- Select carton quantities and build a shipment list.
- Optimize containers using configurable maximum weight and CBM limits.
- Rename containers and move products between containers after optimization.
- Add products to a specific container manually.
- Remove containers with undo support during the current session.
- Export packing lists as CSV or styled Excel workbooks.
- Export one container or all containers.
- Reopen app-generated Excel packing lists after editing them in Excel or LibreOffice.
- Save the current workspace and container history in browser IndexedDB.
- Download JSON backups for recovery or transfer to another browser/computer.

## Tech Stack

- React 19
- Create React App / `react-scripts`
- Tailwind CSS
- ExcelJS for styled Excel export/import
- SheetJS `xlsx` and JSZip for catalog parsing and image extraction
- IndexedDB for browser-local saved projects and history
- Testing Library and Jest for tests

## Project Structure

```text
src/
  App.js                  Main React app, UI, optimization flow, exports
  index.js                React entry point
  index.css               Tailwind and global styles
  packingProject.js       JSON backup validation and app-exported XLSX import
  projectStorage.js       IndexedDB persistence helpers
  App.test.js             UI behavior tests
  packingProject.test.js  Import/backup validation tests

public/
  index.html              App HTML shell
  manifest.json           PWA metadata
  pdf.worker.min.mjs      PDF worker asset kept with the app
```

## Getting Started

Install dependencies:

```bash
npm ci
```

Start the development server:

```bash
npm start
```

Create a production build:

```bash
npm run build
```

Run tests:

```bash
CI=true npm test -- --watchAll=false --runInBand
```

## Basic Workflow

1. Open the app in the browser.
2. Upload an Excel catalog or add products manually.
3. Enter the number of cartons for the products you want to ship.
4. Set the maximum container weight and CBM capacity.
5. Generate containers.
6. Review the packing result and make manual adjustments if needed.
7. Export the packing list as Excel or CSV.
8. Download a JSON backup for long-term recovery or transfer.

## Excel Catalog Import

The catalog importer reads the first worksheet of an uploaded Excel file. It
expects rows that contain product details such as marks, description, carton
quantity, pieces per carton, price, CBM, and weight. Embedded images are read
from the workbook when possible and attached to the imported catalog products.

If the catalog format changes, review `extractXlsxData` in `src/App.js` because
that function maps spreadsheet columns into app product fields.

## Packing List Excel Export And Reimport

The app can export generated containers as styled `.xlsx` packing lists. These
files are intended to be edited later in Excel or LibreOffice and then reopened
inside the app.

When editing an exported packing list, keep the app-generated structure intact:

- Keep the column headers unchanged.
- Keep the `TOTAL` row marker.
- Edit values such as `CTN`, `PCS/CTN`, `U/PRICE`, `DESCRIPTION`, `MARKS&NO`,
  `G.W.(KGS)`, `N.W.(KGS)`, and `CBM`.
- Save the workbook as `.xlsx`.
- Recalculate formulas before saving if Excel asks.

On import, the app recalculates totals from the edited rows. `T/QTY` and
`AMOUNT` are derived from carton quantities, pieces per carton, and unit price.
Gross weight, net weight, and CBM are read as row totals and converted back into
per-carton values.

## Saved History And Backups

The app stores the current workspace and generated container history in the
browser using IndexedDB. This means the data stays available after closing and
reopening the browser, as long as the browser site data is not cleared.

Important notes:

- History is local to the browser and site address.
- Clearing browser storage can remove saved work.
- JSON backups are the safest way to move work between computers or recover
  from browser storage loss.
- Imported JSON backups are validated before being restored.
- App-exported XLSX files can restore packing-list data, but JSON backups keep
  the fuller project state.

## Multi-Project Direction

The storage layer is being prepared for multiple projects, per-project history,
and all-project library backups. Until the UI integration is completed, treat
the visible app workflow as a single active project with saved history and JSON
backup support.

The intended multi-project model is:

- Each project has its own catalog, shipment, containers, capacity, names, and
  history.
- Users can switch between projects without losing saved work.
- A single-project JSON backup restores one project.
- An all-project JSON backup restores the full local library.

## Import Limits And Security

The app performs validation when importing backups and edited packing lists.
Current limits include:

- Maximum input file size: 20 MB
- Maximum embedded image size: 2 MB
- Maximum project size: 50,000 cartons

These checks reduce accidental crashes and obvious malformed imports, but they
are not a complete security boundary. Only open Excel and JSON files from
trusted sources.

## Deployment

Build the static production files:

```bash
npm run build
```

The deployable app is generated in:

```text
build/
```

You can deploy that folder to any static hosting provider, for example Netlify,
Vercel, Cloudflare Pages, GitHub Pages, or an Nginx/Apache server.

For static hosts, use these defaults:

- Build command: `npm run build`
- Publish directory: `build`
- Node version: use the version supported by the lockfile and `react-scripts`

Because the app stores data in browser IndexedDB, there is no server database to
deploy. Each browser keeps its own local data.

## Known Maintenance Notes

- `src/App.js` is large and owns many responsibilities: parsing, optimization,
  UI state, storage orchestration, and export formatting. Splitting it into
  focused modules would make future changes safer.
- Some Excel parsing logic depends on fixed column positions. A shared schema or
  header-based parser would make imports more resilient.
- Browser-local storage is convenient, but important business data should still
  be backed up as JSON files.
- The app currently has tests for backup/import helpers and core UI behavior;
  broader end-to-end coverage would be useful before major deployment changes.
