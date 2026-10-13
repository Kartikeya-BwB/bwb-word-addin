// Writes the add-in manifests from the ribbon definitions (src/commands/ribbon-word.json and ribbon-powerpoint.json):
// public/manifest.xml for the published add-in and manifest.dev.xml for https://localhost:3000. One add-in serves Word
// and PowerPoint, with a ribbon for each. Run after changing a ribbon; never edit the XML by hand.
//   node scripts/make-manifest.mjs [--base https://example.org/path/]
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { HOSTS, SIZES, allControls, collectIcons, iconKey, root } from './ribbon-lib.mjs'

const PUBLISHED = { id: 'b0732c64-7b19-470c-8f8f-22e8d6799c0f', name: 'BwB Toolkit', base: 'https://kartikeya-bwb.github.io/bwb-word-addin/' }
const DEV = { id: 'c8da8771-79e9-4ccb-90a9-e6a5a2ae9158', name: 'BwB Toolkit (dev)', base: 'https://localhost:3000/' }
const VERSION = '2.1.0.0'
const PANE = 'BwBToolkit'
/** Menu items and collapsed groups show small icons, so they carry only the sizes the schema requires. */
const ITEM_SIZES = [16, 32, 80]

const baseArg = process.argv.indexOf('--base')
if (baseArg > 0) PUBLISHED.base = process.argv[baseArg + 1].replace(/\/?$/, '/')

const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

function check(text, max, what) {
  if (text.length > max) throw new Error(`${what} is ${text.length} characters; the limit is ${max}: ${text}`)
  if (/\u2014/.test(text)) throw new Error(`${what} has an em dash: ${text}`)
  return text
}

const GET_STARTED = {
  word: 'Open the BwB tab to start a document from 17 BwB types, apply BwB styles, blocks and power tools, and run the Pre-QA check before sending.',
  powerpoint: 'Open the BwB tab for the BwB slide library, alignment tools, Harvey balls, stickers, trackers and the deck check.',
}

function build({ id, name, base }) {
  const icons = [...collectIcons().keys()]
  const iconId = new Map(icons.map((key, i) => [key, `Ico${i + 1}`]))
  const usedImages = new Set()
  const short = new Map([['GetStarted.Title', 'The BwB tab is ready']])
  const long = new Map()
  const urls = new Map([['Url.Commands', `${base}commands.html`], ['Url.LearnMore', 'https://github.com/Kartikeya-BwB/bwb-word-addin#readme']])

  const iconXml = (spec, pad, sizes = SIZES) => {
    const key = iconId.get(iconKey(spec))
    sizes.forEach((s) => usedImages.add(`${key}.${s}`))
    return `${pad}<Icon>\n${sizes.map((s) => `${pad}  <bt:Image size="${s}" resid="${key}.${s}"/>`).join('\n')}\n${pad}</Icon>`
  }

  const actionXml = (c, pad) => {
    if (c.fn) return `${pad}<Action xsi:type="ExecuteFunction">\n${pad}  <FunctionName>${c.fn}</FunctionName>\n${pad}</Action>`
    const resid = `Url.Pane.${c.pane.replace(/[^a-z0-9]/gi, '_')}`
    urls.set(resid, `${base}index.html#${c.pane}`)
    return `${pad}<Action xsi:type="ShowTaskpane">\n${pad}  <TaskpaneId>${PANE}</TaskpaneId>\n${pad}  <SourceLocation resid="${resid}"/>\n${pad}</Action>`
  }

  const hostXml = (host) => {
    const { ribbon, idPrefix, res, manifestHost, key } = host
    const labelled = (c, pad, sizes) => {
      short.set(`L.${res}.${c.id}`, check(c.label, 125, `Label of ${c.id}`))
      long.set(`T.${res}.${c.id}`, check(c.tip, 250, `Tip of ${c.id}`))
      return `${pad}<Label resid="L.${res}.${c.id}"/>\n${pad}<Supertip>\n${pad}  <Title resid="L.${res}.${c.id}"/>\n` +
        `${pad}  <Description resid="T.${res}.${c.id}"/>\n${pad}</Supertip>\n${iconXml(c.icon, pad, sizes)}`
    }
    const controlXml = (c, pad) => {
      if (c.type === 'menu') {
        if (c.items.length > 15) throw new Error(`Menu ${c.id} has ${c.items.length} items; keep menus to 15.`)
        const items = c.items.map((it) =>
          `${pad}    <Item id="${idPrefix}.${it.id}">\n${labelled(it, `${pad}      `, ITEM_SIZES)}\n${actionXml(it, `${pad}      `)}\n${pad}    </Item>`).join('\n')
        return `${pad}<Control xsi:type="Menu" id="${idPrefix}.${c.id}">\n${labelled(c, `${pad}  `)}\n${pad}  <Items>\n${items}\n${pad}  </Items>\n${pad}</Control>`
      }
      return `${pad}<Control xsi:type="Button" id="${idPrefix}.${c.id}">\n${labelled(c, `${pad}  `)}\n${actionXml(c, `${pad}  `)}\n${pad}</Control>`
    }
    const ids = allControls(ribbon).map((c) => c.id)
    const dupes = ids.filter((x, i) => ids.indexOf(x) !== i)
    if (dupes.length) throw new Error(`Duplicate control ids in the ${key} ribbon: ${dupes.join(', ')}`)

    const pad = '              '
    const groups = ribbon.groups.map((g) => {
      if (g.controls.length > 6) throw new Error(`Group ${g.id} has ${g.controls.length} controls; Office allows 6.`)
      short.set(`L.${res}.G.${g.id}`, g.label)
      return `${pad}<Group id="${idPrefix}.G.${g.id}">\n${pad}  <Label resid="L.${res}.G.${g.id}"/>\n${iconXml(g.icon, `${pad}  `, ITEM_SIZES)}\n` +
        `${g.controls.map((c) => controlXml(c, `${pad}  `)).join('\n')}\n${pad}</Group>`
    }).join('\n')
    short.set(`L.${res}.Tab`, ribbon.tab.label)
    long.set(`GetStarted.${res}`, GET_STARTED[key])
    return `      <Host xsi:type="${manifestHost}">
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.${res}"/>
            <LearnMoreUrl resid="Url.LearnMore"/>
          </GetStarted>
          <FunctionFile resid="Url.Commands"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <CustomTab id="${ribbon.tab.id}">
${groups}
              <Label resid="L.${res}.Tab"/>
              <InsertAfter>TabHome</InsertAfter>
            </CustomTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>`
  }

  const hosts = HOSTS.map(hostXml).join('\n')
  const images = icons.flatMap((key) => SIZES.filter((s) => usedImages.has(`${iconId.get(key)}.${s}`))
    .map((s) => `        <bt:Image id="${iconId.get(key)}.${s}" DefaultValue="${base}assets/ribbon/${key}-${s}.png"/>`)).join('\n')
  const list = (m, tag) => [...m].map(([k, v]) => `        <bt:${tag} id="${k}" DefaultValue="${esc(v)}"/>`).join('\n')
  const origin = new URL(base).origin

  return `<?xml version="1.0" encoding="UTF-8"?>
<!-- Generated by scripts/make-manifest.mjs from src/commands/ribbon-word.json and ribbon-powerpoint.json. Do not edit by hand. -->
<OfficeApp
  xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xsi:type="TaskPaneApp">
  <Id>${id}</Id>
  <Version>${VERSION}</Version>
  <ProviderName>Bankers without Boundaries</ProviderName>
  <DefaultLocale>en-GB</DefaultLocale>
  <DisplayName DefaultValue="${esc(name)}"/>
  <Description DefaultValue="Word documents and PowerPoint decks in the BwB format: document types, styles, building blocks, alignment tools, icons, photos, charts and a pre-QA check."/>
  <IconUrl DefaultValue="${base}assets/addin/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="${base}assets/addin/icon-64.png"/>
  <SupportUrl DefaultValue="https://bwb.earth"/>
  <AppDomains>
    <AppDomain>${origin}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Document"/>
    <Host Name="Presentation"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="${base}index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
${hosts}
    </Hosts>
    <Resources>
      <bt:Images>
${images}
      </bt:Images>
      <bt:Urls>
${list(urls, 'Url')}
      </bt:Urls>
      <bt:ShortStrings>
${list(short, 'String')}
      </bt:ShortStrings>
      <bt:LongStrings>
${list(long, 'String')}
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`
}

// The manifest is machine-written and large, so it carries no indentation: Office caps the file's size.
const compact = (xml) => xml.replace(/^[ \t]+/gm, '')
writeFileSync(join(root, 'public', 'manifest.xml'), compact(build(PUBLISHED)))
writeFileSync(join(root, 'manifest.dev.xml'), compact(build(DEV)))
console.log(`Manifests written for ${PUBLISHED.base} and ${DEV.base}.`)
