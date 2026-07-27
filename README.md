# IdeaTheorem Intranet Theme

## Summary

Custom SPFx solution providing the navigation/theme application customizer and 15 web parts used on the BullWealth intranet.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.21.1-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

1. `npm install -g gulp-cli yo @microsoft/generator-sharepoint`
2. Node version 18

## Solution

| Solution                   | Author(s)   |
| --------------------------- | ----------- |
| ideatheorem-intranet-theme  | IdeaTheorem |

## Local development

1. `npm install`
2. `gulp build`
3. `gulp serve --nobrowser`
4. Access the following link for debugging:
   `https://ideatheorem0.sharepoint.com/sites/BullWealthIntranet-test/_layouts/15/workbench.aspx?debug=true&debugManifestsFile=https://localhost:4321/temp/build/manifests.js`

## How to create a new web part

1. `yo @microsoft/sharepoint`
2. Respond to the prompts

---

## Deploying to production

The live app is the App Catalog entry **"IdeaTheorem Intranet Theme"** (Product ID `e1a45c8f-413c-4de9-abc8-3611caa22b12`), used by the site `https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet`.

Because `includeClientSideAssets` is `false`, the `.sppkg` only contains manifests — the compiled JS bundles are **not** inside it and must be uploaded separately to the asset library. Both steps are required for every release.

### 1. Back up the current production build (do this before every deployment)

- Go to the App Catalog (see step 4 below to get there), select the app row, click **Download** to save the currently-live `.sppkg`.
- Go to `https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/IntranetAssets/`, select all files, and download them too.
- Keep both — they're the only way to roll back, since the App Catalog only ever stores the current version.

### 2. Bump the version

Edit `config/package-solution.json` and increment **both** version fields (`solution.version` and `solution.features[0].version`), e.g. `1.0.0.2` → `1.0.0.3`.

### 3. Build and package

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```

This produces `sharepoint/solution/ideatheorem-intranet-theme.sppkg`.

> Note: `gulp bundle --ship` can exit with code 1 purely because pre-existing lint warnings are written to stderr, even when the bundle succeeds. Check the output for `Finished 'bundle'` rather than trusting the exit code alone.

### 4. Upload the package to the App Catalog

1. Go to `https://bullwealthmanagementgro-admin.sharepoint.com` → **More features** → **Apps** → **Open** (App Catalog).
2. Deselect any selected row so the **Upload** button appears, then upload `sharepoint/solution/ideatheorem-intranet-theme.sppkg`.
3. Confirm **Replace** when prompted (same Product ID upgrade), then **Deploy**.
4. On `/sites/MrkedCapitalIntranet` → **Site Contents**, update the app if it doesn't refresh automatically.

### 5. Upload the compiled JS bundles

Upload every file in `release/assets/` (skip the `.LICENSE.txt` files; include the `fonts` folder) into the root of:

`https://bullwealthmanagementgro.sharepoint.com/sites/MrkedCapitalIntranet/IntranetAssets/`

Confirm **Replace** for any duplicates. Skipping this step leaves the new package pointing at JS files that don't exist yet, which breaks the nav and any web part whose bundle changed (404s in the browser console).

### 6. Verify

Hard refresh the live site and check the browser console for any `Failed to load resource: 404` errors on `*.js` files — that means an asset upload was missed.

### 7. Rollback

If something breaks: re-upload the backed-up `.sppkg` from step 1 the same way (Replace → Deploy), and re-upload the backed-up `IntranetAssets` files if new ones overwrote them.

## Version history

| Version | Date       | Comments                          |
| ------- | ---------- | ---------------------------------- |
| 1.0.0.3 | 2026-07-24 | Data-driven nav library dropdowns  |
| 1.0.0.2 | -          | Previous production release        |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**
