# Version 1: how the source was recovered

This is the record of how the first BwB Word add-in (the "BwB Document
Assistant") was recovered from its deployed build. Version 2 was then built on
what it found; the last section says how version 2 deals with each issue.

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

## Where the version 1 source went

Version 2 (the BwB Toolkit) replaced these files in October 2026. The recovered
version 1 source, and the `npm run verify` script that proves it rebuilds the
original deployment byte for byte, are in commit `19b6e8f`:

```sh
git checkout 19b6e8f -- source   # restores the version 1 source tree
```

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

9. **The manifest misspells a namespace** (found later, while validating
   version 2). It declares `xmlns:bt` as
   `http://schemas.microsoft.com/office/officeappbasistypes/1.0` ("basis"); the
   schema's namespace is `.../officeappbasictypes/1.0` ("basic"). Against
   Microsoft's XSDs every `bt:` element then fails validation, and Office checks
   manifests against those schemas when an add-in is installed.

## How version 2 deals with these

| Issue | In version 2 |
| ----- | ------------ |
| 1. Two tables | Tables are built as one WordprocessingML fragment and inserted in a single call, so there is nothing to fall back from. |
| 2. Header text | The BwB Standard table style sets white bold header text on navy; nothing is shaded by hand. |
| 3. Two-column layout | Replaced by real BwB layout blocks (snapshot, icon grid, key figures) built from the template's layout table styles. |
| 4. Highlight callout | Callouts are real one-cell layout tables with the green or charcoal bar, title and text. |
| 5. Section divider | New reports can open each main section with the template's navy divider page (an option on the Start tab). A divider is a Word section of its own, so it is not offered for inserting into an open document. |
| 6. Template-only styles | Any BwB style is brought into a document that lacks it, the first time it is needed, without changing the text. Headings use `styleBuiltIn`. |
| 7. Placeholder ID | Version 2 has a new, generated ID (`b0732c64-7b19-470c-8f8f-22e8d6799c0f`). |
| 8. Leftovers | Removed. The manifest declares WordApi 1.3, the lowest set version 2 needs. |
| 9. Namespace | The manifest is generated (`scripts/make-manifest.mjs`) with the correct namespace and validates against the schemas. |

