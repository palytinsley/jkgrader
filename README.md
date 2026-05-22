# JK Show and Dress Grader

Teacher-facing grading dashboard for Ken Tinsley's Junk Kouture show and dress unit at Palo Alto High School.

The frontend is a single static `index.html` designed for GitHub Pages. The backend is a Google Apps Script web app in `gas/Code.gs` that reads and writes the shared Google Sheet.

## Structure

- `index.html` - single-file vanilla HTML/CSS/JS dashboard
- `gas/Code.gs` - Apps Script API backend
- `gas/appsscript.json` - Apps Script manifest
- `gas/.clasp.json` - local clasp project binding, gitignored

## Deployment

Set the deployed GAS Web App URL in `index.html`:

```js
const GAS_URL = '';
```

Then publish the static file through GitHub Pages.

## First Run

Run the `setupSheets` action or use the `JK Grader` custom menu in the spreadsheet to create headers in empty grading tabs.
