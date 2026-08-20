import { useEffect, useMemo, useState } from 'react'
import {
  createBrowserOfficeRuntime,
  createOfficeDocumentClient,
  type OfficeHost,
} from './office-document.js'

const hostLabels: Record<OfficeHost, string> = {
  word: 'Microsoft Word',
  excel: 'Microsoft Excel',
  powerpoint: 'Microsoft PowerPoint',
  unknown: 'Office host',
}

export function App() {
  const client = useMemo(() => createOfficeDocumentClient(createBrowserOfficeRuntime()), [])
  const [host, setHost] = useState<OfficeHost>('unknown')
  const [text, setText] = useState('')
  const [status, setStatus] = useState('Connecting to Office…')
  const [busy, setBusy] = useState(true)

  useEffect(() => {
    client
      .initialize()
      .then((activeHost) => {
        setHost(activeHost)
        setStatus(`${hostLabels[activeHost]} is ready`)
      })
      .catch((error: unknown) => {
        setStatus(error instanceof Error ? error.message : 'Office initialization failed')
      })
      .finally(() => setBusy(false))
  }, [client])

  async function readSelection() {
    setBusy(true)
    try {
      const selectedText = await client.readSelection()
      setText(selectedText)
      setStatus('Selection loaded')
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Could not read the selection')
    } finally {
      setBusy(false)
    }
  }

  async function replaceSelection() {
    setBusy(true)
    try {
      await client.replaceSelection(text)
      setStatus('Selection updated')
    } catch (error) {
      setStatus(error instanceof Error ? error.message : 'Could not update the selection')
    } finally {
      setBusy(false)
    }
  }

  return (
    <main className="taskpane">
      <header>
        <span className="eyebrow">WisWork developer module</span>
        <h1>Office Add-in Lab</h1>
        <p>Build and test shared Office.js capabilities in Word, Excel, and PowerPoint.</p>
      </header>

      <section className="status-card" aria-live="polite">
        <span className={`status-dot ${busy ? 'busy' : ''}`} />
        <div>
          <strong>{hostLabels[host]}</strong>
          <p>{status}</p>
        </div>
      </section>

      <section className="editor-card">
        <label htmlFor="selection">Selected content</label>
        <textarea
          id="selection"
          value={text}
          onChange={(event) => setText(event.target.value)}
          placeholder="Read the current selection, edit it here, then write it back."
          rows={10}
        />
        <div className="actions">
          <button type="button" className="secondary" disabled={busy} onClick={readSelection}>
            Read selection
          </button>
          <button type="button" disabled={busy} onClick={replaceSelection}>
            Replace selection
          </button>
        </div>
      </section>
    </main>
  )
}
