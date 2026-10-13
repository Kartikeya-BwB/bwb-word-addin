import { Image } from 'lucide-react'
import { FOOTER_AREA, SLIDE, TITLE_AREA } from '../../brand/slideGrid.ts'
import { currentHost } from '../../host.ts'
import { cmToPt, insertPictureHere } from '../../media/insert.ts'
import { fetchPng } from '../../word/marks.ts'
import { Tile } from '../components.tsx'
import { useAction } from '../useAction.ts'

interface Logo {
  id: string
  label: string
  hint: string
  src: string
  onNavy: boolean
  /** Width in Word, in cm; on a slide, in points, with an optional fixed place. */
  wordCm: number
  slide: { width: number; left?: number; top?: number }
}

const WORD_LOGOS: Logo[] = [
  { id: 'logo', label: 'BwB logo', hint: 'For white pages, 4 cm wide', src: './assets/brand/logo-black.png', onNavy: false, wordCm: 4, slide: { width: 120 } },
  { id: 'logo-white', label: 'BwB logo in white', hint: 'Only on navy panels, 4 cm wide', src: './assets/brand/logo-white.png', onNavy: true, wordCm: 4, slide: { width: 120 } },
]

const SLIDE_LOGOS: Logo[] = [
  {
    id: 'logo-footer', label: 'Logo in the footer', hint: 'Bottom left, where the BwB template puts it', src: './assets/brand/header-logo.png', onNavy: false,
    wordCm: 4, slide: { width: 84, left: FOOTER_AREA.left, top: FOOTER_AREA.top + 2 },
  },
  {
    id: 'logo-corner', label: 'Logo at the top right', hint: 'For a title or cover slide on white', src: './assets/brand/logo-black.png', onNavy: false,
    wordCm: 4, slide: { width: 110, left: SLIDE.width - 40 - 110, top: TITLE_AREA.top },
  },
  { id: 'logo-white', label: 'BwB logo in white', hint: 'Only on navy, centred on the slide', src: './assets/brand/logo-white.png', onNavy: true, wordCm: 4, slide: { width: 200 } },
]

export function LogoPanel() {
  const { busy, run } = useAction()
  const slides = currentHost() === 'powerpoint'
  const logos = slides ? SLIDE_LOGOS : WORD_LOGOS

  const insert = (logo: Logo): void =>
    void run(logo.id, async () => {
      const png = await fetchPng(logo.src)
      const widthPt = slides ? logo.slide.width : cmToPt(logo.wordCm)
      await insertPictureHere({
        bytes: png.bytes, widthPt, heightPt: (widthPt * png.height) / png.width, altText: 'Bankers without Boundaries logo',
        ...(logo.slide.left !== undefined ? { left: logo.slide.left } : {}),
        ...(logo.slide.top !== undefined ? { top: logo.slide.top } : {}),
      })
      return 'Logo inserted.'
    })

  return (
    <div className="tiles">
      {logos.map((l) => (
        <Tile key={l.id} mark={<Image size={18} strokeWidth={1.8} />} markClass={l.onNavy ? '' : 'light'}
          {...(l.onNavy ? {} : { markStyle: { color: 'var(--navy)' } })} label={busy === l.id ? 'Working…' : l.label} hint={l.hint}
          disabled={busy !== null} onClick={() => insert(l)} />
      ))}
    </div>
  )
}
