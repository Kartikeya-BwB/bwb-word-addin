/** Inline and anchored pictures, and the fallback header and footer parts: a port of the drawing helpers in styles.py. */
import { esc, ROOT_NS, spacer, XML_DECL } from './xml.ts'

export const STRIP_W = 6734175
export const STRIP_H = 45085
/** EMU from the page top of the colour strip on content pages. */
export const STRIP_Y_CONTENT = 10332720

function pic(id: number, name: string, descr: string, rid: string, cx: number, cy: number, crop: string): string {
  return (
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData ' +
    'uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic ' +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:nvPicPr><pic:cNvPr id="${id}" name="${name}" descr="${descr}"/>` +
    `<pic:cNvPicPr><a:picLocks noChangeAspect="1"/></pic:cNvPicPr></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rid}"/>` +
    `${crop}<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>'
  )
}

const FRAME_LOCKS =
  '<wp:cNvGraphicFramePr><a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'noChangeAspect="1"/></wp:cNvGraphicFramePr>'

export function inlinePicture(rid: string, cx: number, cy: number, did: number, name = 'Picture', descr = '', crop = ''): string {
  return (
    '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>' +
    `<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/>` +
    '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
    `<wp:docPr id="${did}" name="${esc(name)}" descr="${esc(descr)}"/>` +
    FRAME_LOCKS +
    pic(did, esc(name), esc(descr), rid, cx, cy, crop) +
    '</wp:inline></w:drawing></w:r>'
  )
}

export interface AnchorOptions {
  name?: string
  descr?: string
  behind?: boolean
  halign?: string | null
  z?: number
  crop?: string
}

export function anchoredPicture(
  rid: string,
  cx: number,
  cy: number,
  did: number,
  hrel: string,
  hoff: number,
  vrel: string,
  voff: number,
  o: AnchorOptions = {},
): string {
  const name = o.name ?? 'Picture'
  const descr = o.descr ?? ''
  const behind = o.behind ?? true
  const z = o.z ?? 251658240
  const h = o.halign
    ? `<wp:positionH relativeFrom="${hrel}"><wp:align>${o.halign}</wp:align></wp:positionH>`
    : `<wp:positionH relativeFrom="${hrel}"><wp:posOffset>${hoff}</wp:posOffset></wp:positionH>`
  return (
    '<w:r><w:rPr><w:noProof/></w:rPr><w:drawing>' +
    `<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="${z}" behindDoc="${behind ? 1 : 0}" ` +
    'locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
    h +
    `<wp:positionV relativeFrom="${vrel}"><wp:posOffset>${voff}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
    `<wp:docPr id="${did}" name="${esc(name)}" descr="${esc(descr)}"/>` +
    FRAME_LOCKS +
    pic(did, esc(name), esc(descr), rid, cx, cy, o.crop ?? '') +
    '</wp:anchor></w:drawing></w:r>'
  )
}

/** Header for the cover and dividers, added only when a base package lacks one. */
export function headerBlankXml(): string {
  return XML_DECL + `<w:hdr ${ROOT_NS}>` + spacer(20) + '</w:hdr>'
}

/** Footer for dividers, added only when a base package lacks one. */
export function footerDividerXml(): string {
  return XML_DECL + `<w:ftr ${ROOT_NS}>` + spacer(20) + '</w:ftr>'
}
