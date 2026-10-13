// Which Office application the add-in is running in. One add-in serves Word and PowerPoint: the ribbon comes from
// the manifest's host sections, and the task pane picks its tabs from this.

export type Host = 'word' | 'powerpoint'

/** The host Office reports, or for a pane opened in a plain browser, ?host=powerpoint in the address. */
export function currentHost(): Host {
  try {
    if (typeof Office !== 'undefined' && Office.context && Office.context.host) {
      return Office.context.host === Office.HostType.PowerPoint ? 'powerpoint' : 'word'
    }
  } catch {
    // Office.context is not ready outside Office.
  }
  return new URLSearchParams(window.location.search).get('host') === 'powerpoint' ? 'powerpoint' : 'word'
}

/** True when Office is present and reports this host. A missing host (the test fakes) counts as Word. */
export function inHost(host: Host): boolean {
  if (typeof Office === 'undefined') return false
  try {
    const reported = Office.context?.host
    if (!reported) return host === 'word'
    return host === (reported === Office.HostType.PowerPoint ? 'powerpoint' : 'word')
  } catch {
    return false
  }
}
