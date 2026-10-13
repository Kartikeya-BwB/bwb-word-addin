/** Picture type and pixel size from the file header, for building figure blocks from bytes (PNG, JPEG, GIF). */
import type { PictureData } from './types.ts'

export type ImageInfo = Omit<PictureData, 'data'>

function u16be(d: Uint8Array, i: number): number {
  return ((d[i] ?? 0) << 8) | (d[i + 1] ?? 0)
}

function u32be(d: Uint8Array, i: number): number {
  return (((d[i] ?? 0) << 24) >>> 0) + (((d[i + 1] ?? 0) << 16) | ((d[i + 2] ?? 0) << 8) | (d[i + 3] ?? 0))
}

/** The type and size of a PNG, JPEG or GIF image, or null for anything else. */
export function readImageInfo(d: Uint8Array): ImageInfo | null {
  if (d.length >= 24 && d[0] === 0x89 && d[1] === 0x50 && d[2] === 0x4e && d[3] === 0x47) {
    return { ext: 'png', widthPx: u32be(d, 16), heightPx: u32be(d, 20) }
  }
  if (d.length >= 10 && d[0] === 0x47 && d[1] === 0x49 && d[2] === 0x46) {
    return { ext: 'gif', widthPx: (d[6] ?? 0) | ((d[7] ?? 0) << 8), heightPx: (d[8] ?? 0) | ((d[9] ?? 0) << 8) }
  }
  if (d.length >= 4 && d[0] === 0xff && d[1] === 0xd8) {
    let i = 2
    while (i + 9 < d.length) {
      if (d[i] !== 0xff) {
        i++
        continue
      }
      const marker = d[i + 1] ?? 0
      if (marker === 0xff) {
        i++
        continue
      }
      // start-of-frame markers carry the size; C4 (huffman), C8 and CC (arithmetic coding tables) do not
      if (marker >= 0xc0 && marker <= 0xcf && marker !== 0xc4 && marker !== 0xc8 && marker !== 0xcc) {
        return { ext: 'jpg', widthPx: u16be(d, i + 7), heightPx: u16be(d, i + 5) }
      }
      if (marker === 0xd8 || (marker >= 0xd0 && marker <= 0xd7) || marker === 0x01) {
        i += 2
        continue
      }
      i += 2 + u16be(d, i + 2)
    }
  }
  return null
}
