import { describe, expect, it } from 'vitest'
import { readBytes } from '../../test/helpers/fixtures.ts'
import { readImageInfo } from './imageInfo.ts'

describe('readImageInfo', () => {
  it.each([
    ['test/fixtures/docx/media/chart.png', 'png', 480, 270],
    ['test/fixtures/docx/media/tall.png', 'png', 200, 700],
    ['test/fixtures/docx/media/photo.jpg', 'jpg', 320, 200],
    ['test/fixtures/docx/media/diagram.gif', 'gif', 120, 80],
    ['test/fixtures/docx/media/cover-landscape.jpg', 'jpg', 600, 300],
    ['public/templates/bwb-cover.jpg', 'jpg', 1400, 1563],
    ['public/assets/brand/logo-white.png', 'png', 333, 163],
  ])('reads %s', (path, ext, widthPx, heightPx) => {
    expect(readImageInfo(readBytes(...path.split('/')))).toEqual({ ext, widthPx, heightPx })
  })

  it('returns null for anything else', () => {
    expect(readImageInfo(new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>'))).toBeNull()
    expect(readImageInfo(new Uint8Array([0xff, 0xd8, 0xff]))).toBeNull()
    expect(readImageInfo(new Uint8Array())).toBeNull()
  })
})
