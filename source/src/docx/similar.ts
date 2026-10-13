/**
 * Port of Python's difflib.get_close_matches (SequenceMatcher without junk), so unknown icon names get the same
 * "did you mean" suggestions as the Python builder. Strings are compared by code point.
 */

interface Match {
  a: number
  b: number
  size: number
}

class SequenceMatcher {
  private a: string[] = []
  private b: string[] = []
  private b2j = new Map<string, number[]>()
  private fullbcount: Map<string, number> | null = null

  setSeq2(b: string): void {
    this.b = Array.from(b)
    this.fullbcount = null
    this.b2j = new Map()
    this.b.forEach((elt, i) => {
      const list = this.b2j.get(elt)
      if (list) list.push(i)
      else this.b2j.set(elt, [i])
    })
    // autojunk ("popular" elements) only applies to sequences of 200 or more items
    const n = this.b.length
    if (n >= 200) {
      const ntest = Math.floor(n / 100) + 1
      for (const [elt, idxs] of [...this.b2j]) if (idxs.length > ntest) this.b2j.delete(elt)
    }
  }

  setSeq1(a: string): void {
    this.a = Array.from(a)
  }

  private findLongestMatch(alo: number, ahi: number, blo: number, bhi: number): Match {
    const { a, b, b2j } = this
    let besti = alo
    let bestj = blo
    let bestsize = 0
    let j2len = new Map<number, number>()
    for (let i = alo; i < ahi; i++) {
      const newj2len = new Map<number, number>()
      for (const j of b2j.get(a[i] as string) ?? []) {
        if (j < blo) continue
        if (j >= bhi) break
        const k = (j2len.get(j - 1) ?? 0) + 1
        newj2len.set(j, k)
        if (k > bestsize) {
          besti = i - k + 1
          bestj = j - k + 1
          bestsize = k
        }
      }
      j2len = newj2len
    }
    while (besti > alo && bestj > blo && a[besti - 1] === b[bestj - 1]) {
      besti--
      bestj--
      bestsize++
    }
    while (besti + bestsize < ahi && bestj + bestsize < bhi && a[besti + bestsize] === b[bestj + bestsize]) bestsize++
    return { a: besti, b: bestj, size: bestsize }
  }

  private matches(): number {
    const queue: Array<[number, number, number, number]> = [[0, this.a.length, 0, this.b.length]]
    let total = 0
    for (let q = queue.pop(); q; q = queue.pop()) {
      const [alo, ahi, blo, bhi] = q
      const { a: i, b: j, size: k } = this.findLongestMatch(alo, ahi, blo, bhi)
      if (k) {
        total += k
        if (alo < i && blo < j) queue.push([alo, i, blo, j])
        if (i + k < ahi && j + k < bhi) queue.push([i + k, ahi, j + k, bhi])
      }
    }
    return total
  }

  ratio(): number {
    return ratioOf(this.matches(), this.a.length + this.b.length)
  }

  quickRatio(): number {
    if (!this.fullbcount) {
      this.fullbcount = new Map()
      for (const elt of this.b) this.fullbcount.set(elt, (this.fullbcount.get(elt) ?? 0) + 1)
    }
    const avail = new Map<string, number>()
    let matches = 0
    for (const elt of this.a) {
      const numb = avail.has(elt) ? (avail.get(elt) as number) : (this.fullbcount.get(elt) ?? 0)
      avail.set(elt, numb - 1)
      if (numb > 0) matches++
    }
    return ratioOf(matches, this.a.length + this.b.length)
  }

  realQuickRatio(): number {
    const la = this.a.length
    const lb = this.b.length
    return ratioOf(Math.min(la, lb), la + lb)
  }
}

function ratioOf(matches: number, length: number): number {
  return length ? (2.0 * matches) / length : 1.0
}

/** difflib.get_close_matches(word, possibilities, n, cutoff). */
export function getCloseMatches(word: string, possibilities: readonly string[], n = 3, cutoff = 0.6): string[] {
  if (!(n > 0)) throw new RangeError(`n must be > 0: ${n}`)
  const result: Array<[number, string]> = []
  const s = new SequenceMatcher()
  s.setSeq2(word)
  for (const x of possibilities) {
    s.setSeq1(x)
    if (s.realQuickRatio() >= cutoff && s.quickRatio() >= cutoff) {
      const r = s.ratio()
      if (r >= cutoff) result.push([r, x])
    }
  }
  // heapq.nlargest compares (score, word) tuples: higher score first, then the larger word
  result.sort((p, q) => (q[0] !== p[0] ? q[0] - p[0] : q[1] > p[1] ? 1 : q[1] < p[1] ? -1 : 0))
  return result.slice(0, n).map((r) => r[1])
}
