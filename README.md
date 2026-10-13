# BwB Toolkit (Word add-in)

A Word add-in for writing documents in the BwB format: a **BwB** ribbon tab and a task pane that start any of
17 BwB document types, apply BwB styles, insert building blocks (callouts, key figures, snapshots, tables,
captions, icons, Harvey balls and more), and run a Pre-QA check with one-click fixes before a document goes to QA.

- **Repo root:** what GitHub Pages serves at `https://kartikeya-bwb.github.io/bwb-word-addin/`. These files are
  build output from `source/`, so don't edit them by hand. `manifest.xml` is the file Word installs.
- **[`source/`](source/README.md):** the source code, with how it works and how to build, test and deploy it.
- **[HOW-TO-DEPLOY.md](HOW-TO-DEPLOY.md):** installing the add-in in Word, for one person or for everyone.
- **[`source/docs/v1-recovery.md`](source/docs/v1-recovery.md):** how the first version's source was recovered
  from its deployed build with REA, and what version 2 changed.
