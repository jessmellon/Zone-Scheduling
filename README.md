# Dates Calendar

Static site that reads the shared Google Sheet `Dates` tab and renders a calendar grouped by column `E` (`Zone`).

For documentation:

- `QUICK_START.md`: short how-to guide for everyday use
- `USER_GUIDE.md`: full walkthrough of features and workflow

## Run locally

Use any simple static server from this folder. For example:

```bash
python3 -m http.server 8000
```

Then open `http://localhost:8000`.

## What it does

- Pulls live data from the Google Sheets visualization feed.
- Uses column `E` (`Zone`) as the grouping and filter dimension.
- Uses the `Date` column from the `Dates` tab for calendar placement.
- Shows month navigation, category counts, filtering, and event detail inspection.

## Files

- `index.html`: layout and UI hooks
- `styles.css`: visual design and responsive layout
- `app.js`: Google Sheet fetch, feed parsing, column mapping, and calendar rendering
