// Shared by make-ribbon-icons.mjs and make-manifest.mjs: the two ribbons and the icon each control uses, drawn as SVG.
import { readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

export const root = join(dirname(fileURLToPath(import.meta.url)), '..')
const load = (name) => JSON.parse(readFileSync(join(root, 'src', 'commands', name), 'utf8'))

/** One add-in, two hosts: each has its own ribbon, control id prefix and resource id prefix. */
export const HOSTS = [
  { key: 'word', manifestHost: 'Document', idPrefix: 'BwB', res: 'W', ribbon: load('ribbon-word.json') },
  { key: 'powerpoint', manifestHost: 'Presentation', idPrefix: 'BwBP', res: 'P', ribbon: load('ribbon-powerpoint.json') },
]

/** Every size Word desktop asks for at the common display scales; 16, 32 and 80 are required. */
export const SIZES = [16, 20, 24, 32, 40, 48, 64, 80]
export const ADDIN_SIZES = [16, 32, 64, 80, 128]

const HEX = { navy: '#022338', green: '#328840', lime: '#BED62F', charcoal: '#4A4A4C', white: '#FFFFFF' }

/** A file-safe name for an icon spec: "heading-1", "square-green-fill", "harvey-2". */
export function iconKey(spec) {
  if (typeof spec === 'string') return spec
  if ('harvey' in spec) return `harvey-${spec.harvey}`
  return [spec.name, spec.colour ?? 'navy', spec.fill ? 'fill' : '', spec.background ? `on-${spec.background}` : ''].filter(Boolean).join('-')
}

/** Every control and menu item of a ribbon, with its group, in ribbon order. */
export function allControls(ribbon) {
  const out = []
  for (const group of ribbon.groups) {
    for (const control of group.controls) {
      out.push({ ...control, group: group.id })
      for (const item of control.items ?? []) out.push({ ...item, type: 'item', group: group.id, parent: control.id })
    }
  }
  return out
}

/** Distinct icon specs used by both ribbons' groups, controls and menu items, keyed by iconKey. */
export function collectIcons() {
  const icons = new Map()
  for (const { ribbon } of HOSTS) {
    for (const group of ribbon.groups) icons.set(iconKey(group.icon), group.icon)
    for (const control of allControls(ribbon)) icons.set(iconKey(control.icon), control.icon)
  }
  return icons
}

function lucideSvg(name, colour, fill) {
  const file = join(root, 'node_modules', 'lucide-static', 'icons', `${name}.svg`)
  const svg = readFileSync(file, 'utf8')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/\s*class="[^"]*"/, '')
    .replace('stroke="currentColor"', `stroke="${HEX[colour]}"`)
  return fill ? svg.replace('fill="none"', `fill="${HEX[colour]}"`) : svg
}

/** A Harvey ball: a navy ring filled by quarters, 0 to 4. */
function harveySvg(quarters) {
  const navy = HEX.navy
  const ends = { 1: 'A9 9 0 0 1 21 12', 2: 'A9 9 0 0 1 12 21', 3: 'A9 9 0 1 1 3 12' }
  const fill = quarters >= 4
    ? `<circle cx="12" cy="12" r="9" fill="${navy}"/>`
    : quarters > 0 ? `<path d="M12 12 L12 3 ${ends[quarters]} Z" fill="${navy}"/>` : ''
  return `<svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24">${fill}` +
    `<circle cx="12" cy="12" r="9" fill="none" stroke="${navy}" stroke-width="2"/></svg>`
}

export function iconSvg(spec) {
  if (typeof spec === 'string') return lucideSvg(spec, 'navy', false)
  if ('harvey' in spec) return harveySvg(spec.harvey)
  const svg = lucideSvg(spec.name, spec.colour ?? 'navy', Boolean(spec.fill))
  if (!spec.background) return svg
  return svg.replace(/(<svg[^>]*>)/, `$1<rect width="24" height="24" fill="${HEX[spec.background]}"/>`)
}

/** The add-in's own icon: a navy tile with a white document and the BwB colour strip along the bottom. */
export function addinIconSvg() {
  const doc = lucideSvg('file-text', 'white', false).replace(/^[\s\S]*?<svg[^>]*>/, '').replace(/<\/svg>\s*$/, '')
  return `<svg xmlns="http://www.w3.org/2000/svg" width="64" height="64" viewBox="0 0 64 64">` +
    `<rect width="64" height="64" fill="${HEX.navy}"/>` +
    `<g transform="translate(14 8) scale(1.5)" fill="none" stroke="${HEX.white}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${doc}</g>` +
    `<rect x="0" y="56" width="32" height="8" fill="${HEX.green}"/>` +
    `<rect x="32" y="56" width="21" height="8" fill="${HEX.charcoal}"/>` +
    `<rect x="53" y="56" width="11" height="8" fill="${HEX.lime}"/></svg>`
}
