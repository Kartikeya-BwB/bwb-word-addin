# Building features independently

The BwB Toolkit offers some features that commercial add-ins also offer. Features, workflows and chart types are
free for anyone to build; the way another product expresses them (its code, screens, icons, help text, templates,
word lists and names) is not. These rules keep the Toolkit on the right side of that line, and this folder keeps a
record that each feature was built independently.

## Sources we use

- BwB's own requirements: the Style Guide, the QA Policy, the BwB Document Format skill and requests from BwB staff.
- General knowledge of how documents and slides are made, and published methods (readability formulas, chart forms
  such as waterfall, Mekko and Gantt charts that are decades old).
- Microsoft's Office JavaScript API documentation.
- Open-source code and data under MIT, ISC, BSD or Apache 2.0 licences, with their notices (see
  `public/THIRD-PARTY-NOTICES.txt`), public-domain material, and UK public sector information under the Open
  Government Licence v3.0, credited as that licence requires.

## What we never do

- Install, trial, sign up for or accept the licence of a commercial add-in in order to build the Toolkit.
- Decompile or read another product's code, inspect its files or network traffic, or open its sample files.
- Copy or trace another product's text, screenshots, icons, templates, slide designs, chart styles, help pages,
  word lists or style sheets, or ask an AI tool to reproduce them.
- Use another product's name, or a look-alike, anywhere in the Toolkit: the ribbon, the pane, file names, code,
  comments, tests, commit messages or user documentation. Features get plain, descriptive names. The one exception
  is the market research in `docs/research`, which names the products it looked at; it decides what to build, never
  how a feature looks or reads.
- Use confidential knowledge of another product.

## Record of independent development

Each feature built this way has a short file in this folder: what it does, who built it and when, the sources used
(from the list above), and any third-party data with its licence. Keep the git history intact.

## Ask before

- Publishing the Toolkit openly or giving it to anyone outside BwB.
- Using it in the United States: some software patents there may touch features such as synchronised agendas and
  suggested rewrites of wordy phrases. A patent attorney's freedom-to-operate opinion is the way to settle this.
- Using content whose licence is unclear.
