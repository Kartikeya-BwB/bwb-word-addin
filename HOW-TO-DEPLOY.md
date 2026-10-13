# Deploying the BwB Toolkit

## Where it lives

GitHub Pages serves the root of the `main` branch at `https://kartikeya-bwb.github.io/bwb-word-addin/`. The files
at the root are build output; the source is in `source/`. The add-in's manifest, the file Word installs, is at
`https://kartikeya-bwb.github.io/bwb-word-addin/manifest.xml`.

## Publishing a new version

1. In `source/`, run `npm ci`, `npm test`, then `npm run deploy`. This builds the add-in and copies it to the
   repository root.
2. Commit and merge to `main`. GitHub Pages updates within a few minutes.
3. Run `npm run check-live` in `source/`. It confirms that every page, template and icon the manifest points at
   answers.

Word picks up changes to the pane and the ribbon commands the next time the add-in loads. A change to the ribbon
itself (a new button, say) changes `manifest.xml`; Word reads it again within a day, or at once if the add-in is
removed and added again.

## Installing it in Word

### For everyone at BwB (recommended)

A Microsoft 365 administrator deploys it once and it appears for everyone chosen, on Windows, Mac and the web:

1. Microsoft 365 admin centre, Settings, Integrated apps, Upload custom apps.
2. Choose "Office Add-in", then give the manifest URL above (or upload `manifest.xml`).
3. Assign it to everyone or to a group, and finish. It can take up to 24 hours to appear.

### For one person, to try it

- **Word on the web:** Insert, Add-ins, then My Add-ins, Upload My Add-in, and choose `manifest.xml`.
- **Windows:** put `manifest.xml` in a shared network folder. In Word, File, Options, Trust Center, Trust Center
  Settings, Trusted Add-in Catalogs: add the folder's path, tick "Show in Menu" and restart Word. Then Insert,
  My Add-ins, Shared Folder, BwB Toolkit.
- **Mac:** copy `manifest.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the `wef`
  folder if needed), restart Word, then Insert, Add-ins, My Add-ins.

Once installed, the **BwB** tab appears on the ribbon. The manifest asks for it to sit after Home; Word versions that do not support tab positions put it at the end.

## Moving from version 1

Version 2 has a new add-in ID, so Word treats it as a separate add-in from version 1 ("BwB Document Assistant").
Install the BwB Toolkit, then remove version 1 (in the admin centre, or from My Add-ins). Until version 1 is
removed, its button opens the new pane, because both are served from the same address.
