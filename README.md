# CSC Media Hub V2.6

Fixes the standings-to-logo mismatch introduced when standings Column A started using full franchise names instead of abbreviations.

## What changed in V2.6

- Removed unsafe row-order fallback for standings records.
- The app no longer guesses by row position when it cannot match a franchise.
- Logo/ranking key stays tied to the abbreviation from the Power Rankings sheet.
- Team display name, franchise display name, and record still come from the standings sheet.
- Added stronger alias support in `config.js` for the current standings full franchise names.
- Added automatic abbreviation guessing for simple full names like `NA Nades` -> `NAN` or `Original Superstars` -> `OS` when the match is unique.

## Best setup

In the standings sheet:

- Column A: Franchise display name, like `NA Nades`
- Column B: Team name, like `The Decoys`
- Column C: Record, like `4-4`
- Optional Column D: Code, like `NAN`

If a logo/team ever fails to match, the cleanest fix is adding Column D with the franchise abbreviation.


## V2.7
- Standings Column D can now be named `Abbreviations` and will be used as the exact franchise/logo code.
- Cache key bumped so standings changes pull fresh.
