# Family Giving Tracker

A static page that reads and writes a private Google Sheet. The repository holds no giving data: the page loads the sheet only after a Google sign-in, and the sheet's sharing settings decide who can read or edit.

## Screens

- **Decide**: proposed gifts for the year, with last year's total; approve, defer, or decline.
- **Execute**: approved gifts not yet sent, grouped by vehicle; mark sent, then confirm.
- **Budget**: totals by location, vehicle, and who directs, against the soft targets; the giving pyramid.
- **History**: gifts by organization and year; click an organization for its full history.

## Files

- `config.js`: OAuth client ID and spreadsheet ID (public identifiers).
- `app.js`: all logic. Rows are matched by Gift ID, so sorting the sheet by hand is safe.
- `index.html`, `style.css`: the page.

## Sheet layout expected

Tabs `Gifts`, `Organizations`, `Config`. Columns are found by header name, so they can be reordered but not renamed.

## Running locally

`python3 -m http.server 8000`, then open `http://localhost:8000/`. Add `?demo` to try it with invented data and no sign-in.
