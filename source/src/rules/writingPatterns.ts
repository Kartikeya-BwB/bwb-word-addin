/**
 * The regular expressions and word lists of bwb_lib/writing.py, transcribed pattern by pattern.
 *
 * Every pattern keeps the structure of the Python one (same alternatives in the same order, same groups, same
 * greediness), so that JavaScript's backtracking finds the same matches. Only the character classes change: see
 * writingUnicode.ts. Python's re.I patterns are spelled out with pyIgnoreCase instead of using the i flag.
 */
import {
  PY_BOUNDARY as B,
  PY_DIGIT as D,
  PY_IGNORECASE_ASCII_LETTER,
  PY_SPACE as S,
  PY_WORD as W,
  pyIgnoreCase as ci,
  pyIgnoreCaseClass,
} from './writingUnicode.ts'

const rx = (source: string): RegExp => new RegExp(source, 'gu')
// Python's \b next to a character that must be a word character is just "no word character on the other side".
// The patterns use these cheaper forms wherever that holds, and the full PY_BOUNDARY elsewhere.
const START = `(?<![${W}])`
const END = `(?![${W}])`
/** An alternation of literals, each matched as under re.I. */
const alt = (words: string): string => words.split('|').map(ci).join('|')

export const MONTHS = [
  'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November',
  'December',
] as const
const MONTH_ALT =
  'January|February|March|April|May|June|July|August|September|October|November|December|' +
  'Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sept|Sep|Oct|Nov|Dec'
const MONTH_ABBR = 'Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sept|Sep|Oct|Nov|Dec'
const MONTH_FULL: ReadonlyMap<string, string> = new Map(MONTHS.map((m) => [m.slice(0, 3).toLowerCase(), m]))

/** month_name(): the full name for a month token taken from one of the date patterns. */
export function monthName(token: string): string {
  let t = token.toLowerCase()
  while (t.endsWith('.')) t = t.slice(0, -1)
  if (t === 'may') return 'May'
  return MONTH_FULL.get(t.slice(0, 3)) ?? token
}

// ------------------------------------------------------------------------------------------- spelling tables
const OUR =
  'color|behavior|favor|honor|labor|neighbor|harbor|flavor|humor|rumor|endeavor|odor|armor|vapor|rigor|' +
  'savor|splendor|candor|clamor|glamor|vigor|valor|ardor|fervor'
const OUR_SUFFIXES = 's|ed|ing|ful|less|able|ably|ite|ites|er|ers|ist|ists'
export const OUR_RE = rx(`${START}(${alt(OUR)})(${alt(OUR_SUFFIXES)})?${END}`)

/** US to UK spellings, in the order of the Python dict (the order matters for equal-length alternatives). */
export const WORDS: ReadonlyMap<string, string> = new Map([
  ['center', 'centre'], ['centers', 'centres'], ['centered', 'centred'], ['centering', 'centring'],
  ['centerpiece', 'centrepiece'], ['liter', 'litre'], ['liters', 'litres'], ['fiber', 'fibre'], ['fibers', 'fibres'],
  ['theater', 'theatre'], ['theaters', 'theatres'], ['defense', 'defence'], ['offense', 'offence'],
  ['pretense', 'pretence'], ['catalog', 'catalogue'], ['catalogs', 'catalogues'], ['gray', 'grey'],
  ['skeptic', 'sceptic'], ['skeptics', 'sceptics'], ['skeptical', 'sceptical'], ['skepticism', 'scepticism'],
  ['aluminum', 'aluminium'], ['enrollment', 'enrolment'], ['enroll', 'enrol'], ['enrolls', 'enrols'],
  ['fulfill', 'fulfil'], ['fulfills', 'fulfils'], ['fulfillment', 'fulfilment'], ['skillful', 'skilful'],
  ['installment', 'instalment'], ['installments', 'instalments'], ['jewelry', 'jewellery'],
  ['modeling', 'modelling'], ['modeled', 'modelled'], ['modeler', 'modeller'], ['modelers', 'modellers'],
  ['labeled', 'labelled'], ['labeling', 'labelling'], ['canceled', 'cancelled'], ['canceling', 'cancelling'],
  ['traveled', 'travelled'], ['traveling', 'travelling'], ['traveler', 'traveller'], ['travelers', 'travellers'],
  ['signaled', 'signalled'], ['signaling', 'signalling'], ['fueled', 'fuelled'], ['fueling', 'fuelling'],
  ['channeled', 'channelled'], ['channeling', 'channelling'], ['leveled', 'levelled'], ['leveling', 'levelling'],
  ['totaled', 'totalled'], ['totaling', 'totalling'], ['counselor', 'counsellor'], ['counselors', 'counsellors'],
  ['counseling', 'counselling'],
])
// sorted(WORDS, key=len, reverse=True): both sorts are stable, so equal lengths keep the dict order.
const WORDS_BY_LENGTH = [...WORDS.keys()].sort((a, b) => b.length - a.length)
export const WORDS_RE = rx(`${START}(${WORDS_BY_LENGTH.map(ci).join('|')})${END}`)

export const YZE_RE = rx(`${START}(${alt('analy|paraly|cataly')})${ci('z')}(${alt('e|es|ed|ing|er|ers')})${END}`)
export const ISE_RE = rx(
  `${START}(${PY_IGNORECASE_ASCII_LETTER}{2,}?)${ci('iz')}(${alt('e|es|ed|ing|er|ers|ation|ations|able')})${END}`,
)
/** ISE_EXCEPT.match(word); no g flag, so test() keeps no state. */
export const ISE_EXCEPT = new RegExp(
  `^(?:[${W}]*${ci('siz')}(?:${alt('e|es|ed|ing|able')})|[${W}]*${ci('prize')}[${W}]*|[${W}]*${ci('seiz')}[${W}]*|` +
    `${ci('maize')}|${ci('baize')})$`,
  'u',
)

export const WARN_PROGRAM = rx(`${START}${ci('program')}${pyIgnoreCaseClass('s')}?${END}`)
export const WARN_LICENSE = rx(`${START}${ci('license')}${pyIgnoreCaseClass('sd')}?${END}`)
export const WARN_DIALOG = rx(`${START}${ci('dialog')}${END}`)

// ------------------------------------------------------------------------------------------- patterns
export const EM_DASH = rx('[ \\t]*\\u2014[ \\t]*')
export const EN_NUM_RANGE = rx(`(?<=${D})[ \\t]*\\u2013[ \\t]*(?=${D})`)
export const EN_SPACED = rx('[ \\t]+\\u2013[ \\t]+')
export const EN_OTHER = rx(`(?<=[${W}])\\u2013(?=[${W}])`)
export const YEAR_HYPHEN = rx(
  `(?<![${W}/@.\\-])((?:19|20)${D}${D})[ \\t]*-[ \\t]*((?:19|20)${D}${D})(?![${W}/@\\-])`,
)
export const PERCENT = rx(`${START}${ci('percent')}${END}`)

const YEAR = `((?:19|20)${D}${D})`
export const DATE_US = rx(`${START}(${MONTH_ALT})\\.?[ \\t]+(${D}{1,2})(?:st|nd|rd|th)?,?[ \\t]+${YEAR}${END}`)
// The final \b can follow a full stop ("1st Jan."), so it stays a full boundary test.
export const DATE_ORD = rx(
  `${START}(${D}{1,2})(?:st|nd|rd|th)(?:[ \\t]+of)?[ \\t]+(${MONTH_ALT})\\.?(?:,?[ \\t]+${YEAR})?${B}`,
)
export const DATE_ABBR_DMY = rx(`${START}(${D}{1,2})[ \\t]+(${MONTH_ABBR})\\.?[ \\t]+${YEAR}${END}`)
export const DATE_ABBR_MY = rx(`${START}(${MONTH_ABBR})\\.?[ \\t]+${YEAR}${END}`)
export const DATE_ISO = rx(
  `(?<![${W}/_.\\-])((?:19|20)${D}${D})-(0[1-9]|1[0-2])-(0[1-9]|[12]${D}|3[01])(?![${W}/_\\-]|\\.[${W}])`,
)
export const DATE_NUMERIC = rx(`(?<![${W}/.])${D}{1,2}[/.]${D}{1,2}[/.](?:19|20)?${D}${D}(?![${W}/])`)

export const POUND = rx(`\\u00A3[ \\t]?(?=${D})`)
export const EURO = rx(`\\u20AC[ \\t]?(?=${D})`)
export const OTHER_SYMBOL = rx(
  `(?<![A-Za-z])(?:US|A|C|S|NZ|HK)?[$\\u00A5\\u20B9\\u20A6\\u20B1\\u0E3F\\u20A9\\u20AB][ \\t]?(?=${D})`,
)
export const CODE_GLUED = rx(
  `${START}(USD|GBP|EUR|CHF|JPY|CAD|AUD|NZD|INR|ZAR|KES|NGN|SGD|CNY|BRL|MXN|IDR|PHP|THB|VND|BDT|PKR|GHS|EGP|MAD|XOF)` +
    `(?=${D})`,
)

export const BRAND_BWB = rx(`(?<![@./${W}])(BWB|Bwb|bWb|BWb|bwb)(?![${W}]|\\.earth|\\.com|@)`)
export const BRAND_NAME = rx(`${START}${ci('bankers')}[ \\t]+${ci('without')}[ \\t]+${ci('boundaries')}${END}`)
export const BRAND_NAME_TEXT = 'Bankers without Boundaries'

export const DOUBLE_SPACE = rx(`(?<=[^${S}\\u00B7])[ ]{2,}(?=[^${S}\\u00B7])`)
export const PLACEHOLDER = rx(
  [
    ci('lorem ipsum'),
    `\\[(?:${alt('insert|tbc|tbd|name|date|client')})[^\\]]*\\]`,
    `${START}${ci('TBC')}${END}`,
    `${START}${ci('TBD')}${END}`,
    `${START}${ci('XX')}${ci('X')}+${END}`,
    `${START}${ci('DD Month YYYY')}${END}`,
    `${START}${ci('YYYY')}${END}`,
    '<<[^>]+>>',
  ].join('|'),
)
export const FILLER = rx(
  `${START}(${alt(
    'just to flag|circle back|touch base|loop you in|loop in|quick call|quick recap|going forward|' +
      'at the end of the day|synergies|low-hanging fruit|move the needle|deep dive',
  )})${END}`,
)
export const STRAIGHT_QUOTE = /["']/gu

/** _protect(): URLs, e-mail addresses, code spans, markdown link targets and {icon:name} tokens. */
const PROTECT = rx(
  `\\{icon:[^}]*\\}|\\]\\([^)]*\\)|https?://[^${S}]+|www\\.[^${S}]+|[^${S}]+@[^${S}]+\\.[${W}]+|\`[^\`]*\``,
)
// The e-mail alternative backtracks quadratically over long runs without spaces (in Python too). It cannot match
// without an @, so text without one uses the pattern minus that alternative: the same matches, found faster.
const PROTECT_NO_AT = rx(`\\{icon:[^}]*\\}|\\]\\([^)]*\\)|https?://[^${S}]+|www\\.[^${S}]+|\`[^\`]*\``)
/** The _protect() pattern to use for this text. */
export const protectPattern = (text: string): RegExp => (text.includes('@') ? PROTECT : PROTECT_NO_AT)
/** _restore(): a placeholder as _protect() writes it (any decimal digits, as Python's \d accepts them). */
export const RESTORE = /\uE000(\p{Nd}+)\uE001/gu
export const HOLE_OPEN = '\uE000'
export const HOLE_CLOSE = '\uE001'
