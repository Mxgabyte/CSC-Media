# CSC Graphic Generator Refactor

Open `index.html` from a localhost server or hosted site. Clipboard image copy still needs a secure browser context, but card copy now falls back to downloading a PNG when the clipboard blocks image writes.

Changed:
- Split the generator into `index.html`, `styles.css`, `config.js`, and `app.js`.
- Removed unused color-sheet helper code that referenced an undefined `COLORS_URL`.
- Consolidated power ranking graphic rendering into one shared builder used by both Power Rankings and Power Article.
- Added a Load Debug panel with row counts and warnings for missing/blank sheet data.
- Added PNG download fallback when browser clipboard image copy is blocked.
- Fixed rich article copy canvas tainting by using safe logo fallbacks during file:// canvas export.

Power rankings record update:
- Added a new team records source in `config.js`: `TEAM_RECORDS_SHEET_URL`.
- The records sheet expects one worksheet per tier.
- Column A = Franchise, Column B = Team Name, Column C = Record.
- Power ranking cards now show the franchise plus the team name underneath.
- Power ranking cards now show a `CURRENT RECORD` box on the far right when a record is available.
- Cache key was bumped so the app pulls fresh data instead of using old saved cards without records.
