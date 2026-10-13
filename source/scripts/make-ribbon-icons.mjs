// Draws the ribbon icons (navy Lucide icons, with green and lime accents where a control needs one) and the
// add-in icon as PNGs in every size Word asks for. Run after changing src/commands/ribbon.json.
import { mkdirSync, rmSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { Resvg } from '@resvg/resvg-js'
import { ADDIN_SIZES, SIZES, addinIconSvg, collectIcons, iconSvg, root } from './ribbon-lib.mjs'

const png = (svg, size) => new Resvg(svg, { fitTo: { mode: 'width', value: size } }).render().asPng()

const ribbonDir = join(root, 'public', 'assets', 'ribbon')
rmSync(ribbonDir, { recursive: true, force: true })
mkdirSync(ribbonDir, { recursive: true })
const icons = collectIcons()
for (const [key, spec] of icons) {
  const svg = iconSvg(spec)
  for (const size of SIZES) writeFileSync(join(ribbonDir, `${key}-${size}.png`), png(svg, size))
}

const addinDir = join(root, 'public', 'assets', 'addin')
mkdirSync(addinDir, { recursive: true })
for (const size of ADDIN_SIZES) writeFileSync(join(addinDir, `icon-${size}.png`), png(addinIconSvg(), size))
// The first version's manifest points at these, so installs of it show the new icon until they are replaced.
for (const size of [16, 32, 80]) writeFileSync(join(root, 'public', `icon-${size}.png`), png(addinIconSvg(), size))

console.log(`${icons.size} ribbon icons in ${SIZES.length} sizes, and the add-in icon in ${ADDIN_SIZES.length} sizes.`)
