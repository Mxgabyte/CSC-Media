# CSC Media Hub V2.9

Fixes player hyperlink rendering.

## Changes
- Restored dash-marker hyperlinking for article body text and writer labels.
- Anything like `Mxgabyte-` links and displays as `Mxgabyte` when the name exists on the hyperlink/player sheet.
- Names with real hyphens still work, such as `-Cram--` displaying/linking as `-Cram-`.
- Added support for more hyperlink sheet tab names: Hyperlinks, Player Links, Profiles, CSC Profiles, etc.
- Added support for URL/link columns if the hyperlink sheet has actual profile URLs.
- Added `Profile names` count to Load Debug so you can immediately tell if the hyperlink sheet loaded.
- Bumped cache version to force a fresh pull.


## V2.10
- Updated the hyperlink/player sheet to `1TF6C-wP2ZFErV7o7bMF1sH0DFjsBmzMlWdeFYKd9-zo`.
- Added `External` as an accepted hyperlink/player tab name.


## V2.13
- Fixed prediction article alignment.
- Prediction writeups now match by the matchup inside the Final column N text instead of blindly using row order.
- Row-order fallback only applies to writeup cells that do not contain a matchup title.


V2.13: Fixed a JavaScript syntax error in the prediction matchup matching function that prevented the site from opening.


## V2.13
- Fixed Final tab row 1 / N1 article text being lost when Google CSV is blocked and OpenSheet treats row 1 as headers.
- Added Google Visualization JSON fallback with headers=0 so row 1 is preserved.
