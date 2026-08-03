// Minimal CDP (Chrome DevTools Protocol) client over WebSocket, used to drive
// the installed GenOffice app (an Electron app launched with
// --remote-debugging-port). Renderer is sandboxed, so we operate at the DOM
// level: Runtime.evaluate for reads/writes, Page.captureScreenshot for images.
import { execFile, spawn } from 'node:child_process'
import { writeFileSync } from 'node:fs'
import { join } from 'node:path'
import WebSocket from 'ws'

export const CDP_PORT = 9222

export interface CdpTarget {
  id: string
  title: string
  url: string
  type: string
  webSocketDebuggerUrl?: string
}

/** GET http://127.0.0.1:<port>/json/version — null when no debugger is up. */
export async function cdpVersion(port = CDP_PORT): Promise<Record<string, string> | null> {
  try {
    const res = await fetch(`http://127.0.0.1:${port}/json/version`)
    if (!res.ok) return null
    return (await res.json()) as Record<string, string>
  } catch {
    return null
  }
}

/** GET http://127.0.0.1:<port>/json/list — page targets of the app. */
export async function cdpTargets(port = CDP_PORT): Promise<CdpTarget[]> {
  try {
    const res = await fetch(`http://127.0.0.1:${port}/json/list`)
    if (!res.ok) return []
    const list = (await res.json()) as CdpTarget[]
    return list.filter((t) => t.type === 'page')
  } catch {
    return []
  }
}

async function connect(wsUrl: string): Promise<WebSocket> {
  return new Promise((resolve, reject) => {
    const ws = new WebSocket(wsUrl)
    ws.once('open', () => resolve(ws))
    ws.once('error', (e) => reject(e))
  })
}

interface CdpResponse {
  id: number
  result?: { [k: string]: unknown }
  error?: { message: string }
}

/** Send one CDP command and await its result. */
export async function cdpSend(ws: WebSocket, method: string, params: Record<string, unknown> = {}): Promise<Record<string, unknown>> {
  const id = Math.floor(Math.random() * 1e9)
  const result = await new Promise<CdpResponse>((resolve, reject) => {
    const onMsg = (raw: WebSocket.RawData) => {
      const msg = JSON.parse(raw.toString()) as CdpResponse
      if (msg.id !== id) return
      ws.off('message', onMsg)
      msg.error ? reject(new Error(msg.error.message)) : resolve(msg)
    }
    ws.on('message', onMsg)
    ws.send(JSON.stringify({ id, method, params }))
  })
  return result.result ?? {}
}

export interface EvalResult {
  value?: unknown
  description?: string
  exception?: string
}

/** Evaluate JS in the page context. */
export async function cdpEval(ws: WebSocket, expression: string): Promise<EvalResult> {
  const r = await cdpSend(ws, 'Runtime.evaluate', {
    expression,
    returnByValue: true,
    awaitPromise: true,
  })
  if (r.exceptionDetails) {
    return { exception: JSON.stringify(r.exceptionDetails).slice(0, 500) }
  }
  const res = r.result as { value?: unknown; description?: string } | undefined
  return { value: res?.value, description: res?.description }
}

/** Capture a PNG screenshot of the page. */
export async function cdpScreenshot(ws: WebSocket, outPath: string): Promise<void> {
  const r = await cdpSend(ws, 'Page.captureScreenshot', { format: 'png' })
  const b64 = r.data as string
  if (!b64) throw new Error('Page.captureScreenshot returned no data')
  writeFileSync(outPath, Buffer.from(b64, 'base64'))
}

/** Open the first page target's debugger websocket. */
export async function cdpOpenPage(port = CDP_PORT): Promise<WebSocket> {
  const targets = await cdpTargets(port)
  const page = targets[0]
  if (!page?.webSocketDebuggerUrl) {
    // fetch returns only the JSON fields we typed; re-fetch raw for the ws url
    const res = await fetch(`http://127.0.0.1:${port}/json/list`)
    const all = (await res.json()) as Array<{ type: string; webSocketDebuggerUrl?: string }>
    const p = all.find((t) => t.type === 'page')
    if (!p?.webSocketDebuggerUrl) throw new Error('no page target on the CDP port')
    return connect(p.webSocketDebuggerUrl)
  }
  return connect((page as unknown as { webSocketDebuggerUrl: string }).webSocketDebuggerUrl)
}

// ─── process helpers (macOS: `open -a`, `pkill`) ───────────────────────────

export function isMac(): boolean {
  return process.platform === 'darwin'
}

export function launchGenOffice(port = CDP_PORT): Promise<void> {
  return new Promise((resolve, reject) => {
    if (isMac()) {
      execFile('open', ['-a', 'GenOffice', '--args', `--remote-debugging-port=${port}`], (err) =>
        err ? reject(err) : resolve(),
      )
    } else if (process.platform === 'win32') {
      const bin = join(
        process.env.LOCALAPPDATA ?? '',
        'Programs',
        'GenOffice',
        'GenOffice.exe',
      )
      const child = spawn(bin, [`--remote-debugging-port=${port}`], { detached: true, stdio: 'ignore' })
      child.on('error', (err) => reject(err))
      child.unref()
      resolve()
    } else {
      reject(new Error('GenOffice não suporta Linux (apenas macOS e Windows)'))
    }
  })
}

export function killGenOffice(): Promise<void> {
  const pattern = isMac() ? 'GenOffice.app/Contents/MacOS/GenOffice' : 'GenOffice'
  return new Promise((resolve) => {
    execFile('pkill', ['-f', pattern], () => resolve()) // no process = exit 1, ignore
  })
}

export async function waitForCdp(port = CDP_PORT, timeoutMs = 30_000): Promise<boolean> {
  const start = Date.now()
  while (Date.now() - start < timeoutMs) {
    if (await cdpVersion(port)) return true
    await new Promise((r) => setTimeout(r, 800))
  }
  return false
}
