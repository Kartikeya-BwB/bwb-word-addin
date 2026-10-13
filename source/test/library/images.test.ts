import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import type { ImageIndex } from '../../src/brand/images.ts'
import { orientation, searchImages } from '../../src/brand/images.ts'

const PUBLIC = join(import.meta.dirname, '../../public')
const index = JSON.parse(readFileSync(join(PUBLIC, 'assets/images/index.json'), 'utf8')) as ImageIndex

describe('the image library index', () => {
  it('lists every photo with a full-size file and a thumbnail that exist', () => {
    expect(index.images.length).toBeGreaterThanOrEqual(30)
    for (const image of index.images) {
      expect(existsSync(join(PUBLIC, image.file)), image.file).toBe(true)
      expect(existsSync(join(PUBLIC, image.thumb)), image.thumb).toBe(true)
      expect(image.file.endsWith('.jpg'), image.file).toBe(true)
      expect(index.categories.map((c) => c.id)).toContain(image.category)
      expect(Math.max(image.width, image.height)).toBeLessThanOrEqual(2000)
    }
  })

  it('has unique ids, and a title and tags for every photo', () => {
    const ids = index.images.map((i) => i.id)
    expect(new Set(ids).size).toBe(ids.length)
    for (const image of index.images) {
      expect(image.title.trim(), image.id).not.toBe('')
      expect(image.tags.length, image.id).toBeGreaterThan(2)
    }
  })
})

describe('searchImages', () => {
  it('finds photos by tag and title, best matches first', () => {
    const water = searchImages(index, 'water')
    expect(water.length).toBeGreaterThan(2)
    expect(water.every((i) => i.tags.some((t) => t.includes('water')) || i.title.toLowerCase().includes('water'))).toBe(true)
    expect(searchImages(index, 'energy')[0]?.id).toBe('sector-energy')
  })

  it('keeps catalogue order without a query, and filters by category', () => {
    expect(searchImages(index, '').map((i) => i.id)).toEqual(index.images.map((i) => i.id))
    expect(searchImages(index, '', 'team').every((i) => i.category === 'team')).toBe(true)
    expect(searchImages(index, 'zzz-no-match')).toEqual([])
  })

  it('tells portrait from landscape', () => {
    expect(orientation({ width: 1920, height: 1280 })).toBe('landscape')
    expect(orientation({ width: 1331, height: 2000 })).toBe('portrait')
    expect(orientation({ width: 1000, height: 1000 })).toBe('square')
  })
})
