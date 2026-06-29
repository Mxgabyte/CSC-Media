# CSC Media Hub V2.16

## Changes

- Publish Mode now keeps a clean public viewer menu.
- Public viewer menu includes only:
  - Tier selector
  - Power Rankings / Predictions selector
- Normal visitors are forced into Publish Mode by default.
- Admin/editor mode can be unlocked with a URL query parameter.
- Publish Mode still hides utility controls, copy menus, debug, section menus, refresh/publish buttons, and sidebars.

## Admin unlock

By default, open the site with:

```text
?admin=1
```

Example:

```text
https://mxgabyte.github.io/CSC-Media/?admin=1
```

To make the admin URL less obvious, edit `config.js`:

```js
const ADMIN_UNLOCK_KEY = "yourSecretWordHere";
```

Then open:

```text
https://mxgabyte.github.io/CSC-Media/?admin=yourSecretWordHere
```

Important: this is only a UI gate. GitHub Pages is a static public site, so this does not provide real security.

## Real admin/public split later

For real access control, use one of these:

1. Keep a private admin site and a public viewer site.
2. Use Cloudflare Pages Access / Netlify Identity / another auth gate.
3. Use Google Apps Script or GitHub API to publish shared snapshots to a backend file or sheet.

The current local Publish Draft system is still browser-local. For a true public publish flow, the app needs a shared published-data source.
