# CSC Graphic Generator Refactor

Open `index.html` from a localhost server or hosted site. Clipboard image copy still needs a secure browser context, but card copy now falls back to downloading a PNG when the clipboard blocks image writes.

Changed:
- Split the generator into `index.html`, `styles.css`, `config.js`, and `app.js`.
- Removed unused color-sheet helper code that referenced an undefined `COLORS_URL`.
- Consolidated power ranking graphic rendering into one shared builder used by both Power Rankings and Power Article.
- Added a Load Debug panel with row counts and warnings for missing/blank sheet data.
- Added PNG download fallback when browser clipboard image copy is blocked.
- Left the power rankings card content/design alone.


Patch note:
- Fixed rich article copy canvas tainting by inlining card images as data URLs before html2canvas exports them. If an image cannot be inlined because the browser blocks it, the export uses a safe text fallback instead of failing the full article copy.
