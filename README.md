# Survey Studio

A PowerPoint task pane add-in that loads survey/Excel data and generates D3-based charts (bar, donut, treemap), inserted directly onto slides as images.

## Features

- **Excel import** — load an `.xlsx`/`.xls` file, pick a worksheet, preview the raw data in a separate dialog window
- **Chart types** — bar, donut, and treemap, built with D3 and rendered as PNG shapes on the slide
- **Per-chart customization** — variable (column), category order, per-category colors, title, font sizes
- **Filters** — three layers, all AND-combined: global, per-slide (follows the active slide automatically), and per-chart
- **Persistence** — chart configs and filters are stored as tags on the shape/slide/presentation, so everything survives closing and reopening the file
- **Bulk update** — re-render all charts on the current slide, or across the whole presentation, in one click (e.g. after changing a filter)
- **Info overlay** — toggles a frame + ID/variable label over every chart on the slide, plus the active filters, for debugging
- **Auto-open** — optionally has the task pane open automatically when the presentation is opened (if the add-in is installed)

## Prerequisites

- [Node.js](https://nodejs.org/) (LTS or newer)
- PowerPoint (desktop or web) with support for sideloaded add-ins

## Getting started

```bash
npm install
npm start
```

`npm start` builds the add-in, starts the local HTTPS dev server, and sideloads it into PowerPoint desktop. Trust the local dev certificate once if prompted (`npx office-addin-dev-certs install`).

For PowerPoint on the web, start the dev server separately and sideload manually:

```bash
npm run dev-server
```

Then in PowerPoint on the web: **Insert → Add-ins → Upload My Add-in**, and select `manifest.xml`.

### Other scripts

| Command | Description |
| --- | --- |
| `npm run build` | Production build |
| `npm run build:dev` | Development build (no dev server) |
| `npm run lint` | Lint the source |
| `npm run validate` | Validate `manifest.xml` |
| `npm stop` | Stop a sideloaded session started with `npm start` |

## Project structure

```
manifest.xml              Add-in manifest (PowerPoint, task pane)
src/
  taskpane/                Main task pane app (React)
    components/            Tab UI: Daten, Filter, Grafiken, Diagramm
    chartToSlide.ts         D3 rendering + insert/update chart shapes
    chartRegistry.ts        Chart config type + tag (de)serialization
    excelImport.ts          Workbook parsing, filtering, frequency counts
    overlay.ts               Info overlay drawing
    bulkUpdate.ts            "Update all charts" logic
    slideFilters.ts          Per-slide filter persistence
    filterPersistence.ts     Global filter persistence
    chartIds.ts               Unique chart ID enforcement
  dialog/                   Raw-data dialog window (separate bundle, Office Dialog API)
  commands/                 Ribbon command function file (scaffold)
```

## Notes

- Inserting pictures (`PowerPoint.ShapeCollection.addPicture`) is currently only available in the **preview/beta** Office JavaScript API, so the task pane loads the beta `office.js` build. This isn't guaranteed to work on every PowerPoint build — it's been tested against PowerPoint on the web and recent desktop versions.
- All add-in state (chart configs, filters) is stored using the PowerPoint `Tags` API directly on shapes/slides/the presentation — no external backend or database.
