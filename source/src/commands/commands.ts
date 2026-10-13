// The function file behind the ribbon's ExecuteFunction buttons. A ribbon command has no pane to show a message
// in, so a failure opens a small dialog; Word keeps this page alive until event.completed() is called.

import { errorMessage } from '../word/run.ts'
import { HANDLERS } from './handlers.ts'

function showNotice(message: string): Promise<void> {
  return new Promise((resolve) => {
    const url = new URL('notice.html', window.location.href)
    url.searchParams.set('m', message)
    Office.context.ui.displayDialogAsync(url.toString(), { height: 24, width: 30, displayInIframe: true }, (result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        resolve()
        return
      }
      const dialog = result.value
      const done = (): void => {
        try {
          dialog.close()
        } catch {
          // The person closed it already.
        }
        resolve()
      }
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, done)
      dialog.addEventHandler(Office.EventType.DialogEventReceived, done)
    })
  })
}

function command(handler: () => Promise<unknown>): (event: Office.AddinCommands.Event) => void {
  return (event) => {
    handler().then(
      () => event.completed(),
      async (error: unknown) => {
        await showNotice(errorMessage(error)).catch(() => undefined)
        event.completed()
      },
    )
  }
}

void Office.onReady(() => {
  for (const [name, handler] of Object.entries(HANDLERS)) {
    const fn = command(handler)
    Office.actions.associate(name, fn)
    // Older Word looks the function up on the global object instead.
    Object.assign(globalThis, { [name]: fn })
  }
})
