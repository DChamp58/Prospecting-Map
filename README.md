# Prospecting Map

A single-page **Leaflet** map that plots prospects from a spreadsheet as markers
and lets you filter them interactively. No build step and no backend — it's just
static HTML/CSS/JS, so you can open it locally or host it anywhere (GitHub Pages,
S3, etc.).

## Quick start

Because the app loads local data files (`fetch`), open it through a tiny web
server rather than `file://`:

```bash
# from the project folder
python3 -m http.server 8000
# then visit http://localhost:8000
```

It opens with built-in **sample data**. Drop your own spreadsheet on the
**Load spreadsheet** box (or click to browse) to replace it.

## Expected spreadsheet format

The app reads the format you provided. The first row must be the header row.
It automatically picks the sheet with the most rows and matches columns
case-insensitively.

| Full name | Email | Role | Company | State | Zip Code | Beckhoff Proficiency | Course name | Enrolled | Started | Completed | Score | Course progress |
|-----------|-------|------|---------|-------|----------|----------------------|-------------|----------|---------|-----------|-------|-----------------|

The shorter contact-only layout also works:

| Name | Email | Role | Company | State | Beckhoff Proficiency | Zip Code |
|------|-------|------|---------|-------|----------------------|----------|

A ready-to-use example is included: **`sample-prospects.xlsx`**.

### How markers are placed

There are no latitude/longitude columns, so each row is geocoded locally:

1. **Zip Code** → looked up in `data/zipcodes.min.json` (~42,000 US ZIP
   centroids). This is the precise case.
2. If the ZIP is blank or unknown → the **State** centroid is used
   (`data/states.min.json`) with a small offset so co-located prospects don't
   stack exactly. These are marked *(approx.)* in the popup.
3. If neither is available, the row is counted but not shown on the map (the
   summary tells you how many were skipped).

Overlapping markers are grouped with **marker clustering**; zoom in to split
them apart.

## Filtering

All filters combine with **AND** logic and are built automatically from your
data:

| Filter | Behavior |
|--------|----------|
| **Search** | Matches name, email, or company (substring). |
| **State** | Multi-select; pick one or more states. |
| **Role** | Single select. |
| **Company** | Single select. |
| **Beckhoff Proficiency** | Single select. |
| **Course name** | Single select (hidden if your data has no courses). |
| **Min course progress** | Slider 0-100%; keeps rows at or above the value. |
| **Completed courses only** | Checkbox; treats `yes`/`true`/`100%`/a date as completed. |

Markers are **colored by Beckhoff Proficiency** (see the legend in the sidebar).

## Project layout

```
index.html              Page shell + filter UI
styles.css              Styling
app.js                  Parsing, geocoding, filtering, rendering
data/zipcodes.min.json  ZIP -> [lat, lng] centroids (US)
data/states.min.json    State -> [lat, lng] centroids (fallback)
sample-prospects.xlsx   Example file in the expected format
```

## Customizing

- **Proficiency colors / labels:** edit `PROFICIENCY_COLORS` in `app.js`.
- **Column names:** the `FIELD_ALIASES` map in `app.js` lists accepted header
  spellings for each field — add aliases there if your headers differ.
- **Map start view / tiles:** see `initMap()` in `app.js`.

## Notes

- Leaflet, the marker-cluster plugin, and SheetJS load from CDNs, so the page
  needs internet access the first time (and for map tiles).
- Everything is processed in the browser; spreadsheet data never leaves the
  page.
