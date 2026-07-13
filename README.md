# Inventory Parser

A clean React app for parsing dealership inventory spreadsheets. Drop any `.xlsx`, `.xls`, or `.csv` file and the app reads the first row as column headers — exactly as they appear — and renders all data rows in a searchable, sortable, paginated table.

## Setup

```bash
npm install
npm start
```

The app opens at `http://localhost:3000`.

## Features

- **Drag-and-drop or click to upload** — .xlsx, .xls, .csv
- **Row 1 = headers** — every cell in the first row becomes its own named column, no transformation
- **Search** — filters across all visible columns
- **Sort** — click any column header to sort ascending/descending
- **Hide/show columns** — toggle individual columns on/off with the pill buttons
- **Pagination** — 50 rows per page
- **Fully client-side** — no data leaves the browser

## Dependencies

| Package | Purpose |
|---|---|
| `react` / `react-dom` | UI framework |
| `react-scripts` | CRA dev server + build tooling |
| `xlsx` | Parse .xlsx / .xls / .csv files in the browser |
| `lucide-react` | Icon set |

## Extending

The parsing logic lives entirely in `src/App.jsx` inside the `parseFile` callback. Once you're ready to add alias matching, the `columns` state array (the raw header strings) is the input — run your matcher against it there and update the column display names without touching the data rows.
