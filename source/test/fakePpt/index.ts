/**
 * A fake PowerPoint for tests: an in-memory deck behind the PowerPoint JavaScript API, with Office.js load and sync
 * semantics, in the style of test/fakeWord.
 *
 * Files: base.ts (load and sync), model.ts (data and builders), queries.ts, text.ts, shapes.ts, table.ts,
 * presentation.ts, install.ts (globals and events). Import everything from this file.
 *
 * Using it
 *
 *   const a = shape({ id: 'a', left: 40, top: 108 })
 *   const deck = createPresentation([slide([a, textBox('Hello', { id: 'b' })], { id: 's1' })])
 *   select(deck, ['a', 'b'])                       // or clickShapes(deck, ['a', 'b']) to fire the event too
 *   const fake = installFakePowerPoint(deck)       // sets PowerPoint, Office and OfficeExtension on globalThis
 *   await codeUnderTest()                          // calls PowerPoint.run(async (context) => ...)
 *   expect(a.left).toBe(...)                       // the model is plain data: read it directly
 *   uninstallFakePowerPoint()                      // in afterEach
 *
 * Semantics
 *
 * - A property is readable only after load() and a later context.sync(); reading earlier throws PropertyNotLoaded.
 *   load() takes 'a,b/c', ['a'], { select, expand } or nested option objects. collection.load('name') loads the
 *   items' name, as in Office.js; load('items') or load() loads the items' scalar properties. A navigation property
 *   loaded on its own (shape.load('fill')) loads its scalars; nested paths (items/textFrame/textRange/text) work.
 * - Writes and method calls are queued and applied, in order, on sync. The first failure rejects the sync and drops
 *   the rest of the batch. PowerPoint.run syncs once more after the batch function returns. fake.syncs() counts them.
 * - Requirement sets: model.apiVersion (default '1.10', or installFakePowerPoint(deck, { apiVersion: '1.5' })) is
 *   what Office.context.requirements.isSetSupported('PowerPointApi', v) answers to. Using a newer API than that fails
 *   the sync with code ApiNotFound: getSelected* and setSelected* (1.5), z-order, groups, placeholderFormat, addTable
 *   and Slide.moveTo (1.8), table cell formatting and the table style option (1.9), rotation, isDecorative, pageSetup,
 *   alt text, exportAsBase64Presentation and getTextFrameOrNullObject (1.10). Reading a 1.10 property loaded through
 *   load() or '*' on an older fake throws.
 * - *OrNullObject methods return null objects: isNullObject is readable after any load and sync; reading other
 *   properties of a null object, or writing to it, throws (check isNullObject first). getItem() on a missing id fails
 *   the sync with ItemNotFound.
 * - Errors are FakeOfficeError with code, message and debugInfo, like OfficeExtension.Error.
 *
 * The model (model.ts)
 *
 * - FakePresentation: slides, masters with layouts, tags, pageSetup (960 by 540), selection, apiVersion, and the
 *   records dataInserts (setSelectedDataAsync), slideInserts (insertSlidesFromBase64), slideExports
 *   (exportAsBase64Presentation), createdPresentations (PowerPoint.createPresentation), handlers and settings.
 * - FakeSlide: id, shapes (top-level, back to front: the index is the z-order position), tags, layoutId, masterId.
 * - FakeShape: id, name, type, left/top/width/height/rotation, isDecorative, geometricShapeType, connectorType,
 *   placeholderType, fill, line, textFrame (null for shapes without text), tags (keys upper case), members, table.
 *   Shape types are Office.js's: a picture is 'Image' (PowerPoint.ShapeType has no 'Picture'), an inserted SVG is
 *   'Graphic'. A group's frame is the union of its members (boundsOf); moving or resizing a group moves or scales
 *   its members. Fill and outline written to a group go to its members.
 * - FakeTextFrame: runs of text with a font each, so a write through getSubstring() keeps the formatting around it;
 *   one horizontalAlignment and one bulletVisible (paragraphFormat.bulletFormat.visible) for the whole frame;
 *   verticalAlignment, autoSizeSetting, wordWrap and the four margins. Body and content placeholders have bullets.
 * - Builders: shape() (a rectangle in PowerPoint's default style), textBox(), placeholder(type, text), picture(),
 *   line(), group(members), tableShape(values), slide(shapes), createPresentation(slides), textFrame(), run(), font().
 * - Queries (queries.ts): findShape, findSlide, activeSlide, allShapes, shapesTagged, textOf, boundsOf, locate,
 *   selectedShapes.
 * - Selection (queries.ts): select(deck, ids, slideId?), selectSlides(deck, ids), selectText(deck, shapeId, start,
 *   length). They change the model only; clickShapes(deck, ids) also fires DocumentSelectionChanged.
 *   getSelectedSlides() lists selection.slideIds, the active slide first. getSelectedShapes() lists the selected
 *   shapes in z-order, never in the order they were picked, as in PowerPoint.
 *
 * What the API covers
 *
 * - Presentation: slides (add, getItem, getItemAt, getItemOrNullObject, getCount), slideMasters and their layouts,
 *   tags, pageSetup, getSelectedShapes/Slides/TextRange(OrNullObject), setSelectedSlides, insertSlidesFromBase64
 *   (recorded; set model.slidesFromBase64 to turn the call into slides, inserted after targetSlideId or first, with
 *   the ids the test gave them), slides.exportAsBase64Presentation(ids) and getSelectedSlides()
 *   .exportAsBase64Presentation() (recorded in slideExports; the value is base64 of "fake-presentation:id,id").
 * - PowerPoint.createPresentation(base64?) records the call in createdPresentations.
 * - Slide: id, index, shapes, tags, layout, slideMaster, delete, moveTo, setSelectedShapes, applyLayout,
 *   exportAsBase64. slides.add({ layoutId }) copies the layout's placeholders.
 * - Shapes: addGeometricShape, addTextBox, addLine, addTable (values and options recorded on shape.table), addGroup,
 *   getItem(OrNullObject), getItemAt, getCount; on the selection or a group, group().
 * - Shape: id, name, type, geometry, rotation, visible, isDecorative, alt text, zOrderPosition, level, fill
 *   (setSolidColor, clear, foregroundColor, transparency, type), lineFormat, textFrame (textRange with text, font,
 *   paragraphFormat, getSubstring, setSelected), tags, placeholderFormat (type, containedType), group (shapes, ungroup),
 *   parentGroup, delete, setZOrder, getTable (values, getCellOrNullObject: text, and 1.9 formatting recorded in
 *   table.cells), getParentSlide, getTextFrameOrNullObject.
 * - Office.context: host 'PowerPoint', requirements, document.setSelectedDataAsync (records the call in dataInserts;
 *   an image or SVG becomes a selected 'Image' or 'Graphic' shape on the active slide; set model.failNextDataInsert
 *   to make it fail), addHandlerAsync and removeHandlerAsync, settings. Office.HostType, CoercionType, EventType and
 *   AsyncResultStatus, and PowerPoint enums (PowerPoint.ShapeZOrder.bringToFront is 'BringToFront').
 * - Events: fireSelectionChanged(deck) calls every DocumentSelectionChanged handler and awaits what they return;
 *   clickShapes(deck, ids) selects and fires. Selection changes made through the API do not fire the event.
 *
 * Not modelled: paragraph-level formatting (alignment and bullets are per frame), bullet styles, hyperlinks,
 * adjustments, charts, custom XML, and the formatting of table cells beyond recording it.
 */

export * from './base.ts'
export * from './model.ts'
export * from './queries.ts'
export { frameText } from './text.ts'
export { clickShapes, fireSelectionChanged, installFakePowerPoint, uninstallFakePowerPoint } from './install.ts'
export type { FakePowerPointHandle, InstallOptions } from './install.ts'
