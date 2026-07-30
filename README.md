# EntoData Doctor

EntoData Doctor is a free, open-source browser tool for checking and cleaning
messy specimen spreadsheets. It is designed to work with the same ordinary
Excel and CSV tables used by EntoLabel.

## Try it

1. Unzip the release.
2. Open `index.html` in a modern browser.
3. Click **Try the intentionally messy example**, or choose your own
   `.xlsx`, `.xls`, `.csv`, or `.tsv` file.

The app has no account and no database. Spreadsheet processing happens in the
browser. The original file is never overwritten.

## Version 0.2 includes

- worksheet and header-row detection;
- multilingual column mapping suggestions;
- preservation of unknown columns;
- missing and duplicate Specimen ID checks;
- missing date, locality, scientific-name, and collector checks;
- date validation and ISO date suggestions;
- time and time-range recognition, including values such as `14:00`,
  `10.00`, and `14:00-00:00`;
- latitude/longitude parsing and range checks;
- incomplete coordinate-pair checks;
- whitespace and placeholder cleanup;
- altitude-unit cleanup, including values such as `300 M`, `1'000m`, and
  `520 m a.s.l.`;
- suggested vocabulary for sex, life stage, and collecting method;
- separate safe fixes and confirm-required suggestions;
- direct editing of every current-table cell in the app;
- **Edit cell** links from individual check-up notes;
- paginated original/current table preview with all columns available;
- reset of all accepted changes;
- Excel export with `Cleaned Data`, `Issues`, and `Change log` sheets;
- flat CSV export ready for reuse in EntoLabel.

## Put it on GitHub Pages

1. Create a new public GitHub repository, for example `entodata-doctor`.
2. Upload **the contents of this folder** to the repository root.
3. Open **Settings → Pages**.
4. Under **Build and deployment**, choose **Deploy from a branch**.
5. Select the `main` branch and `/ (root)`, then save.

GitHub will show the public address when deployment finishes. No paid hosting or
API key is required. EntoData Doctor is a static browser app, so it does not
need Streamlit, Python, or a server. A Russian step-by-step copy of these
instructions is included in `PUBLISHING.md`.

## Important limits of this first version

Version 0.2 does not yet check names against an external taxonomic catalogue,
geocode localities, merge collecting events, or create a Darwin Core archive.
It never silently corrects a taxon or invents missing scientific information.

## Project files

- `index.html` — website entry point;
- `entodata-doctor.css` — visual design;
- `entodata-doctor.js` — import, diagnosis, fixing, and export logic;
- `vendor/xlsx.full.min.js` — SheetJS Community Edition spreadsheet reader.

## License

EntoData Doctor is released under the MIT License. The bundled SheetJS
Community Edition library retains its own Apache 2.0 license; see
`THIRD_PARTY_NOTICES.md`.
