# BwB Toolkit: source

The BwB Toolkit is a Word add-in for writing documents in the BwB format. It works the way consulting firms'
in-house toolbars do: a ribbon tab and a task pane that put the house template, a library of building blocks,
status stickers, Harvey balls and a house-style checker one click away.

GitHub Pages serves the build from the repository root. Everything here is the source.

## What it does

| Where | What |
| ----- | ---- |
| **Start** | Starts any of 17 BwB document types (inception report, final report, gap analysis, policy brief, memo and so on). Each comes with the cover, contents, numbered headings, the type's structure with writing prompts, and the disclaimer; long reports can open each main section with a navy divider page. Records the document type, project, deliverable, version and funder in the document, and suggests a file name in the BwB pattern. |
| **Write** | Applies BwB heading, text and list styles. Headings 1 to 3 number themselves. **Clean up** fixes the house style in the selection that has one right answer: UK spelling, dashes, dates, currency codes, "per cent", the company name, double spaces and quotation marks. |
| **Insert** | Eight building blocks with placeholders to type over: callout, note, key figures, snapshot, pull quote, case study, icon grid and details table. Also BwB data tables, formatting for a pasted table, numbered captions, source lines, footnotes, contents, the disclaimer, the EU acknowledgement, an abbreviations table built from the document, Harvey balls, 1,866 Lucide icons in the brand colours and sizes, and the logo. |
| **Check** | A Pre-QA check in the BwB QA vocabulary (Blocker, Major, Minor, Copyedit). It covers what to clear before sending, structure, tables and figures, brand, consistency (acronyms, hyphenation) and house-style wording. Every finding has **Go to**, and most have **Fix**; **Fix all safe items** applies the ones that change only what they name. Also the author self-check list (saved with the document), status stickers (Draft, For discussion and so on), field updates and a QA note to paste into the QA request. |
| **Brand** | Colours (with the text rules, and buttons to apply the text colours), type scale, house style, document furniture and icon rules. |
| **Ribbon** | A BwB tab with Document, Text, Insert, Review and Brand groups. Most buttons act at once; the rest open the pane where they are needed. |

The rules come from the BwB Document Format skill (v1.1.0). The QA vocabulary, checklists and document types come
from the BwB Deliverables Hub plan (QA Policy v2.0, Style Guide v2.0 and the document-type research of
10 October 2026).

## Working on it

Needs Node.js 22 or later.

```sh
cd source
npm ci          # exact versions from package-lock.json
npm test        # unit tests, including tests against a fake Word API
npm run build   # regenerates the icons and manifests, type-checks, then writes dist/
npm run deploy  # builds, then copies dist/ to the repository root for GitHub Pages
```

To try a change in Word before deploying it:

1. Run `npx office-addin-dev-certs install` once. It trusts a localhost certificate that the dev server picks up.
2. Run `npm run dev` to serve the add-in at https://localhost:3000.
3. Sideload `manifest.dev.xml`, which installs as "BwB Toolkit (dev)" next to the published add-in. In Word on
   the web: Insert, Add-ins, Upload My Add-in. On Windows and Mac, see [HOW-TO-DEPLOY.md](../HOW-TO-DEPLOY.md).

`public/manifest.xml` and `manifest.dev.xml` are generated from `src/commands/ribbon.json` by
`scripts/make-manifest.mjs`, and the ribbon icons by `scripts/make-ribbon-icons.mjs`. Edit the JSON, never the XML.

## How it works

- **New documents are built in the pane.** `src/docx` is a TypeScript port of the BwB skill's Python document
  builder. It fills the BwB base template (`public/templates/bwb-base.docx`) with the cover, contents, skeleton and
  disclaimer. Word opens the result with `createDocument().open()` (WordApi 1.3), so it works the same on Windows,
  Mac and the web.
- **Blocks are inserted as WordprocessingML.** The same builder turns a block into a fragment that carries the
  BwB styles with it. Word adds any style the document lacks and never overwrites one it has.
- **Styles arrive when they are first needed.** In a document that did not start from the template, the first BwB
  style used brings in the full BwB style set (`src/word/styles.ts`), without changing any text.
- **The check reads a snapshot.** `src/word/snapshot.ts` reads the document in a few batched calls; the rules in
  `src/check` are pure functions over that snapshot, so they are tested without Word. Fixes check that the
  paragraph still reads as it did, and stop rather than guess if it does not.
- **The wording rules are a line-by-line port** of the skill's Python rules (`src/rules/writing.ts`). A
  differential test checks the port against outputs recorded from the Python on a corpus of over 700 strings.

### Word versions

Everything needs WordApi 1.3 (Word 2016 or later, Microsoft 365, Word on the web). Newer Word adds:

| Requirement set | Adds |
| --------------- | ---- |
| WordApi 1.4 | The count of comments in the check |
| WordApi 1.5 | Checking which styles a document has, restyling Word's own styles, updating fields from the pane, footnotes |
| WordApi 1.6 | The count of tracked changes in the check |

Without them, the add-in says what to do by hand (for example Ctrl+A, then F9 to update fields).

The pane also needs a modern web view: Microsoft 365 on Windows, Mac or the web, or Office 2021 or later on Windows.
Office 2016 and 2019 bought outright run add-ins in Internet Explorer 11 or the old Edge, which cannot run it.

## Layout

| Path | What it holds |
| ---- | ------------- |
| `src/brand` | BwB tokens, styles, document types and their skeletons, checklists, the block library and icons |
| `src/docx` | The document and block builder (a port of the skill's Python builder) and the BwB template's styles |
| `src/rules` | House-style wording rules, acronym and hyphenation analysis, heading case |
| `src/check` | The Pre-QA rules and the engine that runs them |
| `src/word` | Everything that talks to Word: styles, inserting, tables, stickers, properties, the snapshot, fixes |
| `src/ui` | The task pane (React) |
| `src/commands` | The ribbon definition and the functions its buttons run |
| `scripts` | Generators for the manifests, ribbon icons and icon index; deploy and live-check scripts |
| `test/fakeWord` | An in-memory Word API with Office.js load and sync semantics, used by the tests |
| `docs/v1-recovery.md` | How version 1 was recovered from its deployed build, and how version 2 deals with its issues |

## Known limits and open questions

- **The disclaimer wording** is the standard text from the BwB Document Format skill. The Deliverables Hub plan
  lists the disclaimer variants as an open compliance question, so check with the Compliance Lead before relying on
  it for anything that could be a financial promotion.
- **Fonts.** Candara and Aptos ship with Microsoft 365. Under the BwB licensing rule the pane never serves them;
  it uses them when they are installed and falls back to Segoe UI.
- **The check is a snapshot.** After editing, run it again; a fix on a paragraph that has changed is skipped, not
  guessed.
- **Ribbon commands** report failures in a small dialog, because a ribbon button has no pane to show a message in.
