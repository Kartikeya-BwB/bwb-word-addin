#!/usr/bin/env python3
"""Record the reference behaviour of bwb_lib/writing.py for the TypeScript port in src/rules/writing.ts.

Usage, from source/:
  python3 scripts/gen-writing-fixtures.py --skill "$SKILL"

$SKILL is the bwb-document-format skill folder (the one holding scripts/bwb_lib/writing.py).

The script reads test/fixtures/writing-corpus.json, regenerates its seeded "generated" section and writes it back, then
runs every corpus text through the Python module and writes test/fixtures/writing-expected.json:
  * scan() for each context in CONTEXTS, plus the rule and the code point span (in the protected string) of every issue
  * fix() with fragment False/True crossed with quotes False/True, for each context
  * fix() with only one rule switched on, for each rule that changes the text (checks the per-issue replacements)
  * curly()
  * Python's Unicode tables for the character classes the port emulates (\\w, \\d, \\s, isdigit, isupper ...)
Both files are deterministic: sorted keys, one entry per line, ASCII only (other characters are \\u escaped).
"""
import argparse
import hashlib
import json
import platform
import random
import re
import sys
import unicodedata
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CONTEXTS = [{}, {"title": True}, {"compact": True}]
FIX_VARIANTS = {"plain": (False, False), "quotes": (False, True), "fragment": (True, False), "fragmentQuotes": (True, True)}
# Patterns fix() applies, in its order. The TypeScript rule ids use the same names.
FIX_RULES = ["EM_DASH", "EN_NUM_RANGE", "EN_SPACED", "YEAR_HYPHEN", "PERCENT", "OUR_RE", "WORDS_RE", "YZE_RE", "ISE_RE",
             "DATE_US", "DATE_ORD", "DATE_ABBR_DMY", "DATE_ABBR_MY", "DATE_ISO", "POUND", "EURO", "CODE_GLUED",
             "BRAND_BWB", "BRAND_NAME", "DOUBLE_SPACE"]
# Patterns scan() walks, in its order. The three WARN_* names stand for the entries of WARN_ONLY_SPELLING.
SCAN_RULES = ["EM_DASH", "EN_NUM_RANGE", "EN_SPACED", "EN_OTHER", "YEAR_HYPHEN", "PERCENT", "OUR_RE", "WORDS_RE",
              "YZE_RE", "ISE_RE", "WARN_PROGRAM", "WARN_LICENSE", "WARN_DIALOG", "DATE_US", "DATE_ORD", "DATE_ABBR_DMY",
              "DATE_ABBR_MY", "DATE_ISO", "DATE_NUMERIC", "POUND", "EURO", "OTHER_SYMBOL", "CODE_GLUED", "BRAND_BWB",
              "BRAND_NAME", "DOUBLE_SPACE", "PLACEHOLDER", "FILLER"]
WARN_RULES = ["WARN_PROGRAM", "WARN_LICENSE", "WARN_DIALOG"]
NEVER = re.compile(r"(?!)")

# ------------------------------------------------------------------------------------------- seeded random corpus
GEN_SEED = 20261013
GEN_COUNT = 360
FRAGMENTS = [
    "color", "Color", "colors", "COLOR", "honorable", "favorite", "the Labor Party", "behavior", "neighbors", "colour",
    "center", "Center", "the Center for Global Development", "centered", "defense", "catalog", "gray", "skeptical",
    "enrollment", "fulfill", "jewelry", "modeling", "modeled", "labeled", "canceled", "traveling", "counselor",
    "analyze", "Analyze", "Analyze Inc.", "analyzing", "paralyzed", "catalyze", "organize", "organization",
    "Organization", "World Health Organization", "realized", "prioritizing", "size", "resized", "prize", "seize",
    "maize", "citizen", "program", "programs", "programme", "license", "licensed", "licence", "dialog", "dialogue",
    "percent", "Percent", "per cent", "2022-2025", "2022 - 2025", "FY2022-2025", "2022\u20132025", "5\u201310",
    "5 \u2013 10", "Paris \u2013 London", "North\u2013South", "\u2014", " \u2014 ", "\u2014 ", "January 5, 2024",
    "Jan. 5 2024", "Sept 15th, 2025", "5th May", "1st of June 2026", "22nd Feb. 2023", "5 Jan 2024", "15 Sept 2025",
    "Jan 2024", "Oct. 2026", "2024-01-05", "2024-13-01", "05/10/2026", "5.10.26", "\u00a35m", "\u00a3 5m",
    "\u20ac10", "\u20ac 1.2bn", "$5", "US$250m", "A$ 3", "\u00a5100", "\u20b950", "USD5m", "GBP5m", "EUR 10m",
    "KES1000", "BWB", "Bwb", "bwb", "BwB", "bWb", "bankers without boundaries", "Bankers Without Boundaries",
    "Bankers without Boundaries", "bwb.earth", "  ", "   ", "\u00b7", "TBC", "TBD", "XXX", "[insert date]",
    "[Client name]", "Lorem ipsum", "DD Month YYYY", "<<Name>>", "touch base", "circle back", "going forward",
    "Going Forward", "low-hanging fruit", "deep dive", "synergies", "\"quoted\"", "'single'", "it's", "rock 'n' roll",
    "https://bwb.earth/color-center", "www.example.org/2022-2025", "info@bwb.earth", "`color  code`",
    "[the color guide](https://x.org/color)", "{icon:leaf}", "{icon:leaf:green}", "caf\u00e9", "Z\u00fcrich",
    "Organizaci\u00f3n", "\U0001F600", "\u00b2", "\u0663", "20\u0662\u0664", "x", "and", "the", "5", "10", "2024",
    "Hello.", "Note:", "(", ")", "\"", "'", ",", ".", "!", "?", ";", "\n", "\t", "\u00a0", "-", "/", "@",
]
SEPARATORS = [" ", " ", " ", " ", "", ", ", ". ", "  ", "\t", "\n", " \u2014 ", "-", " \u2013 ", ": "]


def generate_texts(seed=GEN_SEED, count=GEN_COUNT):
    rng = random.Random(seed)
    seen, out = set(), []
    while len(out) < count:
        parts = [rng.choice(FRAGMENTS) for _ in range(rng.randint(1, 7))]
        text = parts[0]
        for p in parts[1:]:
            text += rng.choice(SEPARATORS) + p
        r = rng.random()
        if r < 0.15:
            text = text[:1].upper() + text[1:]
        elif r < 0.25:
            text += rng.choice([".", "!", "?", ":", " ", "\n"])
        elif r < 0.3:
            text = rng.choice(["\"", "(", "\u201c", "'"]) + text
        if text not in seen:
            seen.add(text)
            out.append(text)
    return out


# ------------------------------------------------------------------------------------------- recording helpers
class Recorder:
    """Stands in for a compiled pattern during scan() and logs every match finditer() yields."""

    def __init__(self, name, rx, log):
        self.name, self.rx, self.log = name, rx, log

    def finditer(self, s):
        for m in self.rx.finditer(s):
            self.log.append((self.name, m))
            yield m


def scan_with_spans(w, text, ctx):
    """scan() plus, for every issue, the rule that produced it and its code point span in the protected string."""
    log = []
    saved = {n: getattr(w, n) for n in SCAN_RULES if n not in WARN_RULES}
    saved_warn = w.WARN_ONLY_SPELLING
    try:
        for n, rx in saved.items():
            setattr(w, n, Recorder(n, rx, log))
        w.WARN_ONLY_SPELLING = [(Recorder(n, rx, log), msg) for n, (rx, msg) in zip(WARN_RULES, saved_warn)]
        issues = w.scan(text, ctx)
    finally:
        for n, rx in saved.items():
            setattr(w, n, rx)
        w.WARN_ONLY_SPELLING = saved_warn
    kept = [(n, m) for n, m in log
            if not (n == "ISE_RE" and w.ISE_EXCEPT.match(m.group(0)))
            and not (n == "BRAND_NAME" and m.group(0) == "Bankers without Boundaries")]
    if len(kept) != len(issues):
        raise SystemExit(f"internal error: {len(kept)} recorded matches for {len(issues)} issues in {text!r}")
    return [list(i) for i in issues], [[n, m.start(), m.end()] for n, m in kept]


def run_fix(w, text, ctx, fragment, quotes):
    try:
        new, n = w.fix(text, ctx, fragment=fragment, quotes=quotes)
        return [new, n]
    except Exception as e:  # noqa: BLE001 - the port mirrors Python's exceptions
        return {"error": type(e).__name__}


def single_rule_outputs(w, text, ctx):
    """fix(text, ctx) with only one rule switched on; kept only where the result differs from the text."""
    saved = {n: getattr(w, n) for n in FIX_RULES}
    out = {}
    try:
        for keep in FIX_RULES:
            for n in FIX_RULES:
                setattr(w, n, saved[n] if n == keep else NEVER)
            try:
                new, _ = w.fix(text, ctx)
            except Exception as e:  # noqa: BLE001
                out[keep] = {"error": type(e).__name__}
                continue
            if new != text:
                out[keep] = new
    finally:
        for n, rx in saved.items():
            setattr(w, n, rx)
    return out


def case_record(w, case_id, text):
    rec = {"id": case_id, "text": text, "curly": list(w.curly(text)), "results": []}
    protected = w._protect(text)[0]
    if protected != text:
        rec["protected"] = protected
    # Texts that already contain the placeholder markers make _restore misbehave; the per-rule check skips them.
    markers = "\ue000" in text or "\ue001" in text
    for ctx in CONTEXTS:
        issues, spans = scan_with_spans(w, text, ctx)
        res = {"scan": issues, "spans": spans,
               "fix": {k: run_fix(w, text, ctx, f, q) for k, (f, q) in FIX_VARIANTS.items()},
               "rules": None if markers else single_rule_outputs(w, text, ctx)}
        rec["results"].append(res)
    return rec


# ------------------------------------------------------------------------------------------- Unicode tables
def ranges(pred):
    out, start = [], None
    for cp in range(0x110000):
        if pred(cp):
            if start is None:
                start = cp
        elif start is not None:
            out.append([start, cp - 1])
            start = None
    if start is not None:
        out.append([start, 0x10FFFF])
    return out


def unicode_tables():
    w, d, s = re.compile(r"\w"), re.compile(r"\d"), re.compile(r"\s")
    ci = {}
    for letter in "abcdefghijklmnopqrstuvwxyz":
        rx = re.compile(letter, re.I)
        ci[letter] = [cp for cp in range(0x110000) if rx.match(chr(cp))]
    rx_class = re.compile("[A-Za-z]", re.I)
    samples = ["COLOR", "Color", "cOLOR", "colour", "A1", "1", "", " ", "\u0130", "\u0131", "\u017f", "\u212a", "\u00df",
               "\u01c5", "\u01c5A", "\ufb01", "\u00c0\u00c9", "\u015eENER", "MODEL\u0130NG", "FIBER", "f\u0131ber",
               "s\u212aeptic", "\u03a3", "\u0391\u03a3", "\u0391\u03a3 \u0391", "PER CENT", "BEHAV\u0130OR"]
    return {
        "unidata": unicodedata.unidata_version,
        "assigned": ranges(lambda cp: unicodedata.category(chr(cp)) != "Cn"),
        "w": ranges(lambda cp: w.match(chr(cp)) is not None),
        "d": ranges(lambda cp: d.match(chr(cp)) is not None),
        "s": ranges(lambda cp: s.match(chr(cp)) is not None),
        "isdigit": ranges(lambda cp: chr(cp).isdigit()),
        "isalnum": ranges(lambda cp: chr(cp).isalnum()),
        "isupper": ranges(lambda cp: chr(cp).isupper()),
        "decimalZeros": [cp for cp in range(0x110000) if unicodedata.decimal(chr(cp), -1) == 0],
        "ignoreCase": ci,
        "ignoreCaseAsciiClass": [cp for cp in range(0x110000) if rx_class.match(chr(cp))],
        "strings": [{"s": x, "isupper": x.isupper(), "upper": x.upper(), "lower": x.lower()} for x in samples],
    }


# ------------------------------------------------------------------------------------------- output
def dumps(x):
    return json.dumps(x, ensure_ascii=True, sort_keys=True)


def write_corpus(path, corpus):
    lines = ["{", ' "generated": {', f'  "count": {corpus["generated"]["count"]},',
             f'  "seed": {corpus["generated"]["seed"]},', '  "texts": [']
    texts = corpus["generated"]["texts"]
    lines += [f"   {dumps(t)}{',' if i < len(texts) - 1 else ''}" for i, t in enumerate(texts)]
    lines += ["  ]", " },", ' "handwritten": [']
    groups = corpus["handwritten"]
    for gi, g in enumerate(groups):
        lines.append(f'  {{"group": {dumps(g["group"])}, "texts": [')
        lines += [f"   {dumps(t)}{',' if i < len(g['texts']) - 1 else ''}" for i, t in enumerate(g["texts"])]
        lines.append(f"  ]}}{',' if gi < len(groups) - 1 else ''}")
    lines += [" ]", "}"]
    path.write_text("\n".join(lines) + "\n", encoding="ascii")


def write_expected(path, meta, cases, uni):
    lines = ["{", f' "contexts": {dumps(CONTEXTS)},', f' "fixRules": {dumps(FIX_RULES)},', f' "meta": {dumps(meta)},',
             f' "scanRules": {dumps(SCAN_RULES)},', ' "unicode": {']
    keys = sorted(uni)
    lines += [f"  {dumps(k)}: {dumps(uni[k])}{',' if i < len(keys) - 1 else ''}" for i, k in enumerate(keys)]
    lines += [" },", ' "cases": [']
    lines += [f"  {dumps(c)}{',' if i < len(cases) - 1 else ''}" for i, c in enumerate(cases)]
    lines += [" ]", "}"]
    path.write_text("\n".join(lines) + "\n", encoding="ascii")


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--skill", required=True, help="path to the bwb-document-format skill folder")
    ap.add_argument("--corpus", default=str(ROOT / "test/fixtures/writing-corpus.json"))
    ap.add_argument("--out", default=str(ROOT / "test/fixtures/writing-expected.json"))
    ap.add_argument("--no-unicode", action="store_true", help="skip the Unicode tables (faster, for quick checks)")
    args = ap.parse_args()

    module_path = Path(args.skill) / "scripts" / "bwb_lib" / "writing.py"
    if not module_path.is_file():
        raise SystemExit(f"writing.py not found under {args.skill}")
    sys.path.insert(0, str(module_path.parent.parent))
    from bwb_lib import writing as w  # noqa: E402

    corpus_path = Path(args.corpus)
    corpus = json.loads(corpus_path.read_text(encoding="utf-8"))
    corpus["generated"] = {"seed": GEN_SEED, "count": GEN_COUNT, "texts": generate_texts()}
    write_corpus(corpus_path, corpus)

    cases = []
    for g in corpus["handwritten"]:
        for i, text in enumerate(g["texts"]):
            cases.append(case_record(w, f"{g['group']}/{i}", text))
    for i, text in enumerate(corpus["generated"]["texts"]):
        cases.append(case_record(w, f"generated/{i}", text))

    meta = {"python": platform.python_version(), "unidata": unicodedata.unidata_version,
            "writingSha256": hashlib.sha256(module_path.read_bytes()).hexdigest()}
    uni = {} if args.no_unicode else unicode_tables()
    write_expected(Path(args.out), meta, cases, uni)
    n_hand = sum(len(g["texts"]) for g in corpus["handwritten"])
    print(f"{len(cases)} cases ({n_hand} hand-written, {len(corpus['generated']['texts'])} generated) "
          f"x {len(CONTEXTS)} contexts written to {args.out}")


if __name__ == "__main__":
    main()
