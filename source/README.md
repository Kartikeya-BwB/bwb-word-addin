# BwB Toolkit: source

The BwB Toolkit is one add-in for Word and PowerPoint that puts the BwB format one click away. It works the way
consulting firms' in-house toolbars do: a BwB ribbon tab and a task pane with the house templates, a library of
building blocks and slides, power tools for lining things up, charts, icons and photos in the brand, and a checker
that finds what to fix before a document or deck goes to QA.

GitHub Pages serves the build from the repository root. Everything here is the source.

## What it does in Word

| Where | What |
| ----- | ---- |
| **Start** | Starts any of 17 BwB document types (inception report, final report, gap analysis, policy brief, memo and so on), with the cover photo of your choice. Each comes with the cover, contents, numbered headings, the type's structure with writing prompts, and the disclaimer; long reports can open each main section with a navy divider page. Records the document type, project, deliverable, version and funder in the document, and suggests a file name in the BwB pattern. |
| **Write** | Applies BwB heading, text and list styles. Headings 1 to 3 number themselves. **Clean up** fixes the house style in the selection that has one right answer: UK spelling, dashes, dates, currency codes, "per cent", the company name, double spaces and quotation marks. |
| **Insert** | Eight building blocks with placeholders to type over, BwB data tables, captions, source lines, footnotes, contents, the disclaimer, the EU acknowledgement, an abbreviations table built from the document, Harvey balls, **charts**, **icons, flags and maps**, the **photo library** and the logo. |
| **Tools** | 23 power tools. Pictures: fit to the text width, same width, align, caption them all, alt text from captions. Tables: distribute and fit columns, set exact widths, add a total row, delete empty rows, sort, transpose, convert to and from text. Text: change case (keeping acronyms), join lines broken by a PDF, clear stray formatting, remove empty paragraphs, non-breaking spaces, symbols. A navigator with heading levels and cross-references, and review tools (track changes, accept or reject all, delete comments, review notes). |
| **Check** | A Pre-QA check in the BwB QA vocabulary (Blocker, Major, Minor, Copyedit). It covers what to clear before sending, structure, tables and figures, brand, consistency (acronyms, hyphenation), house-style wording and copyediting (repeated words, unmatched brackets and quotation marks, stray spaces, list endings). Every finding has **Go to**, and most have **Fix**; **Fix all safe items** applies the ones that change only what they name. Also the author self-check list, status stickers, field updates and a QA note to paste into the QA request. |
| **Brand** | Colours (with the text rules, and buttons to apply the text colours), type scale, house style, document furniture and icon rules. |

## What it does in PowerPoint

| Where | What |
| ----- | ---- |
| **Slides** | **New BwB deck** on the BwB template. A library of 20 BwB slides (title, agenda, divider, executive summary, key figures, chart, table, process, matrix, timeline, team, closing and more) to insert after the current slide. The storyline: every action title in order, and an executive summary built from them. Sections, an agenda built from them and section trackers on every slide. Stickers (Draft, For discussion and so on), and copying the selected slides to a new deck. |
| **Arrange** | Line up selected shapes with the selection, an anchor, the slide, the margins or the content area. Spread them out with equal gaps or an exact gap, even out gaps, make them the same size, set an exact position and size in centimetres, stack them, lay them out in a grid, tidy them into rows and columns, snap them to the BwB 12-column grid, swap two shapes, stretch an edge, nudge, and change their order or group them. The anchor is the first shape you select, or the one you set. |
| **Insert** | BwB slide elements (callout, key figures, quote, process, matrix, timeline, numbered markers, source line, BwB table), Harvey balls, status marks and dots, review notes, **charts**, **icons, flags and maps**, the **photo library** and logos placed where the template puts them. |
| **Check** | A deck check in the same QA vocabulary: what to clear before sending (review notes, stickers, placeholder text), brand (fonts, colours, sizes), layout (titles out of place, shapes off the slide or off the grid), wording and the storyline, with **Go to**, **Fix** and **Fix all safe items**. |
| **Brand** | Colours with buttons that fill the selected shapes or colour their text, the slide type scale, the slide grid and slide furniture. |

## In both

- **Ribbon.** A BwB tab in each app. Most buttons act at once; the rest open the pane where they are needed.
- **Find a tool.** Press / in the pane and type what you want to do, for example "swap", "total row" or "harvey half".
- **Charts.** Column, stacked and 100% column, bar, stacked and 100% bar, line, waterfall, Mekko, Gantt, donut and
  Harvey ball tables, in the BwB colours and type. Paste the data from Excel. PowerPoint gets a vector graphic in
  the content area; Word gets a sharp picture with a numbered caption.
- **Icons, flags and maps.** Lucide (the BwB set, 1,866 icons), the BwB sectors, Tabler (5,184 icons), 260 flags
  and 231 maps of countries and regions, in the brand colours and sizes.
- **Photos.** The BwB photo library: sectors, themes, markets and the team, taken from the BwB website.

The rules come from the BwB Document Format skill (v1.1.0). The QA vocabulary, checklists and document types come
from the BwB Deliverables Hub plan (QA Policy v2.0, Style Guide v2.0 and the document-type research of
10 October 2026).

## Working on it

Needs Node.js 22 or later.

```sh
cd source
npm ci          # exact versions from package-lock.json
npm test        # unit tests, including tests against fake Word and PowerPoint APIs
npm run build   # regenerates the ribbon icons and manifest, type-checks, then writes dist/
npm run deploy  # builds, then copies dist/ to the repository root for GitHub Pages
npm run demo    # the pane in a browser: Word, or add ?host=powerpoint for a pretend PowerPoint with a sample deck
```

To try a change in Word or PowerPoint before deploying it:

1. Run `npx office-addin-dev-certs install` once. It trusts a localhost certificate that the dev server picks up.
2. Run `npm run dev` to serve the add-in at https://localhost:3000.
3. Sideload `manifest.dev.xml`, which installs as "BwB Toolkit (dev)" next to the published add-in. In Word or
   PowerPoint on the web: Insert, Add-ins, Upload My Add-in. On Windows and Mac, see
   [HOW-TO-DEPLOY.md](../HOW-TO-DEPLOY.md).

`public/manifest.xml` and `manifest.dev.xml` are generated from `src/commands/ribbon-word.json` and
`src/commands/ribbon-powerpoint.json` by `scripts/make-manifest.mjs`, and the ribbon icons by
`scripts/make-ribbon-icons.mjs`. Edit the JSON, never the XML. The libraries are generated and committed:

| Library | Command | Output |
| ------- | ------- | ------ |
| Photos | `npm run images -- --website <the BwB website's public folder>` | `public/assets/images` |
| Tabler icons, flags and maps | `npm run icons` | `public/assets/icons` |
| Slide template, starter deck and slide library | `npm run slides` | `public/templates/bwb-slides.pptx`, `bwb-deck.pptx`, `public/assets/slides` |

## How it works

- **One add-in, two apps.** The manifest has a Word section and a PowerPoint section with their own ribbons. The
  pane asks Office which app it is in and shows that app's tabs; each tab's code loads when it is first opened.
- **New documents are built in the pane.** `src/docx` is a TypeScript port of the BwB skill's Python document
  builder. It fills the BwB base template with the cover, contents, skeleton and disclaimer, and Word opens the
  result with `createDocument().open()`. A new deck is the BwB starter deck opened with `createPresentation`.
- **Blocks are inserted as WordprocessingML** that carries the BwB styles with it. Word adds any style the
  document lacks and never overwrites one it has. Library slides go in with `insertSlidesFromBase64`.
- **Arrange computes geometry itself.** PowerPoint has no align or distribute API, so `src/geometry` works out
  every position from the shapes' frames, allowing for rotation, and writes them back in one go.
- **The checks read a snapshot.** The document or deck is read in a few batched calls; the rules are pure functions
  over that snapshot, so they are tested without Office. Fixes check that the text still reads as it did, and stop
  rather than guess if it does not.
- **The wording rules are a line-by-line port** of the skill's Python rules (`src/rules/writing.ts`), checked
  against outputs recorded from the Python on a corpus of over 700 strings.

### Word versions

Everything needs WordApi 1.3 (Word 2016 or later, Microsoft 365, Word on the web). Newer Word adds:

| Requirement set | Adds |
| --------------- | ---- |
| WordApi 1.4 | The count of comments in the check, bookmarks for cross-references |
| WordApi 1.5 | Checking which styles a document has, restyling Word's own styles, updating fields, footnotes |
| WordApi 1.6 | The count of tracked changes, accepting or rejecting them all |
| WordApiDesktop | Fitting columns to their contents, Word's own case changes and cross-references (Windows and Mac only) |

### PowerPoint versions

The deck check needs PowerPointApi 1.4, and Arrange, the elements and Go to need PowerPointApi 1.5 (current
Microsoft 365 and PowerPoint on the web). Newer PowerPoint adds:

| Requirement set | Adds |
| --------------- | ---- |
| PowerPointApi 1.8 | Order and group, BwB tables, moving a new slide next to the current one, shapes inside groups |
| PowerPointApi 1.9 | The BwB table style and cell formatting |
| PowerPointApi 1.10 | Rotation-aware arranging, alt text on inserted pictures, the slide size, copying slides to a new deck |

Without a set, the pane says what needs a newer version, or what to do by hand. The pane also needs a modern web
view: Microsoft 365 on Windows, Mac or the web, or Office 2021 or later on Windows.

## Layout

| Path | What it holds |
| ---- | ------------- |
| `src/brand` | BwB tokens, styles, document types, checklists, the block library, the slide grid, icons, photos |
| `src/docx` | The document and block builder (a port of the skill's Python builder) and the BwB template's styles |
| `src/rules` | House-style wording rules, acronym and hyphenation analysis, heading case |
| `src/check` | The Word Pre-QA rules and the engine that runs them |
| `src/word` | Everything that talks to Word; `src/word/tools` holds the Tools tab's power tools |
| `src/ppt` | Everything that talks to PowerPoint: arrange, elements, marks, format; `slides` and `check` hold the slide tools and the deck check |
| `src/geometry` | Pure alignment, distribution and layout maths for Arrange |
| `src/charts` | The chart renderer: pure TypeScript from data to a BwB SVG |
| `src/media` | Placing pictures and SVG graphics in either app |
| `src/ui` | The task pane (React): `tabs`, `ppt`, `tools`, `charts` and `library` |
| `src/commands` | The two ribbon definitions, the functions their buttons run and the tool catalogue for Find a tool |
| `scripts` | Generators for the manifest, ribbon icons and libraries; deploy and live-check scripts |
| `test/fakeWord`, `test/fakePpt` | In-memory Word and PowerPoint APIs with Office.js load and sync semantics, used by the tests and the demo |
| `docs/v1-recovery.md` | How version 1 was recovered from its deployed build, and how version 2 deals with its issues |
| `docs/research` | Research on premium Word and PowerPoint add-ins: what they offer, what they cost, what the toolkit builds instead, and the legal boundaries |
| `docs/clean-room` | The rules for building features independently, and a record for each feature built that way |

## Licences

- Lucide icons: ISC. Tabler icons: MIT. Flags: flag-icons, MIT. Maps: Natural Earth (public domain) through
  world-atlas, ISC. Each library's licence ships beside it in `public/assets`.
- **Photos:** the team photos are BwB's own. The sector, theme and market photos come from the BwB website; their
  licences are not recorded there, so check them before using a photo in a deliverable for a client.
- **Fonts.** Candara and Aptos ship with Microsoft 365. Under the BwB licensing rule the add-in never serves or
  embeds them; the pane uses them when they are installed and falls back to Segoe UI.

## Known limits and open questions

- **Tested against fakes.** Every tool is tested against in-memory fakes of the Office APIs; the PowerPoint tools
  and the Word power tools still need a pass in real Word and PowerPoint on Windows, Mac and the web.
- **The disclaimer wording** is the standard text from the BwB Document Format skill. The Deliverables Hub plan
  lists the disclaimer variants as an open compliance question, so check with the Compliance Lead before relying on
  it for anything that could be a financial promotion.
- **The checks are snapshots.** After editing, run them again; a fix on text that has changed is skipped, not
  guessed. The deck check does not read speaker notes (PowerPoint has no API for them), or text in tables, charts
  and SmartArt.
- **Keyboard shortcuts** are not set up yet. Word and current versions of PowerPoint (Windows 2601 or later, Mac
  16.105.2 or later, and the web) support custom shortcuts for add-ins; they are being added.
- **Ribbon commands** report failures in a small dialog, because a ribbon button has no pane to show a message in.
