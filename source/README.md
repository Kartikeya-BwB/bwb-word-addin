# BwB Document Assistant: source

This folder is the source code for the BwB Word add-in that GitHub Pages serves
from the repo root.

## Where it came from

Until October 2026 the repo held only the compiled build: one minified 212 KB
JavaScript file and no source. On 13 October 2026 the source was recovered from
that build with [REA (Reverse Engineer Anything)](https://github.com/morluto/rea)
and rewritten into the files in this folder.

**Proof that it is the same app:** building this folder reproduces all 9
deployed files byte for byte, including `assets/index-BarzYrWO.js` (SHA-256
`df524a42…`), the CSS, `index.html`, `manifest.xml` and the icons. Run
`npm run verify` to check this yourself.

How the recovery went:

1. `rea analyze-javascript-application` inventoried the deployed files and
   traced `index.html` to `assets/index-BarzYrWO.js`. Its deeper semantic pass
   ran out of memory on the bundled React code (even with an 8 GB limit), and
   REA reported that coverage as partial rather than guessing.
2. `rea recover-javascript-sources`, using the Wakaru engine, turned the
   one-line bundle into about 17,000 readable lines with the JSX restored.
   About 850 of those lines are the add-in; the rest is React and the icon
   library.
3. The toolchain was identified from evidence in the files. `favicon.svg` and
   `icons.svg` are byte-identical to the create-vite 9 React starter, the
   bundle embeds React `19.2.6`, its helper code is Rolldown's (Vite 8), and
   the icon code matches lucide-react 1.x. The upload date (11 May 2026)
   narrowed the versions.
4. The add-in code was rewritten as ordinary JSX and rebuilt. With Vite 8.0.11,
   @vitejs/plugin-react 6.0.1, React 19.2.6 and lucide-react 1.14.0 the output
   matched exactly. Changing a single character of the source changes the
   output hash, so the match is not a coincidence.

What could not be recovered: the original variable names, comments and file
layout. The names here were chosen for readability and do not affect the build
output, and the split into one file per tab is a reconstruction.

## Working on it

Needs Node.js 20.19+ or 22.12+.

```sh
cd source
npm ci          # installs the exact versions in package-lock.json
npm run dev     # previews the task pane in a browser; Word actions do nothing outside Word
npm run build   # writes dist/
npm run verify  # builds, then compares dist/ with the deployed files at the repo root
```

To deploy a change, run `npm run build`, copy everything in `source/dist/` to
the repo root (replacing the old files in `assets/`), then commit and push to
`main`. After you change the source, `verify` will list the files that differ
from the live site, which is a quick way to see what a deploy will change.

`manifest.xml` and the icons live in `source/public/` and are copied into the
build, so edit them there.

Versions are pinned exactly so the build stays reproducible. Upgrading them is
fine, but the output hashes will change.

## Layout

| File                               | What it does                                                  |
| ---------------------------------- | ------------------------------------------------------------- |
| `src/main.jsx`                     | Waits for `Office.onReady`, then renders the app              |
| `src/App.jsx`                      | Header, the four tabs, toast messages and the `runWord` helper |
| `src/components/StylesPanel.jsx`   | Applies heading and body paragraph styles                     |
| `src/components/ListsPanel.jsx`    | Applies BwB bullet and numbered list styles                   |
| `src/components/InsertPanel.jsx`   | Inserts tables, the callout paragraph and page breaks         |
| `src/components/ColoursPanel.jsx`  | Brand colour swatches with copy buttons, and the brand fonts  |
| `src/components/ui.jsx`            | Shared `Section` and `Divider` components                     |
| `scripts/verify-live.mjs`          | Compares a build with the deployed files                      |

## Issues spotted during recovery

All of these are in the add-in as deployed today. None has been fixed here, so
that the build stays identical to the live version. Items marked "check in
Word" are inferences from the code and Microsoft's documentation and have not
been tested in Word.

1. **Inserting a data table can leave two tables** (check in Word). The table
   is inserted first and styled afterwards. Commands that already ran are not
   undone when a later one fails (Microsoft's `OfficeExtension.Error` example
   shows this), so if a styling step throws, the table stays and the fallback
   in the `catch` block inserts a second one. Two lines look likely to throw:
   `cell.shadingColor = '001522'` (Microsoft documents the `#RRGGBB` form) and
   `table.style = 'TableGrid'` (the style's name is "Table Grid", and the next
   line sets `styleBuiltIn` anyway).
2. **Header text may be unreadable** (check in Word). Header cells are shaded
   navy but no font colour is set. If the BwB template's body text is navy, as
   the Colours tab says, the header text will be navy on navy.
3. **Two-Column Layout comes out entirely navy.** It is a one-row table, and
   the header shading is applied to the first row, which is the whole table.
4. **Highlight Callout does not insert a box.** It inserts a plain Normal
   paragraph and then tells the user to apply the BwB Highlight table style,
   which cannot be applied to a paragraph.
5. **Section Divider is not implemented.** It only shows a red message asking
   the user to copy the panel from the template.
6. **Most styles only work in documents based on the BwB template.**
   `BwB Bullet Lime`, `Strong1` and the other BwB styles are custom styles, so
   they only work where those styles exist. The heading names (`heading 1` and
   so on) are the English names; `styleBuiltIn` also works in other Word
   languages.
7. **The add-in ID is a placeholder.** `manifest.xml` uses
   `a1b2c3d4-e5f6-7890-abcd-ef1234567890`. Replace it with a freshly generated
   GUID before an organisation-wide rollout, because changing it later makes
   Word treat it as a different add-in.
8. **Leftovers.** `favicon.svg` and `icons.svg` are the Vite starter's logo
   and social icons and are not used. `ui.jsx` has 11 unused style entries.
   `manifest.xml` does not declare the WordApi 1.3 requirement that table
   insertion and table styles need.
