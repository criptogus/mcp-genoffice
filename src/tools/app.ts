import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js'
import { execFile, execSync } from 'node:child_process'
import { existsSync } from 'node:fs'
import { join } from 'node:path'
import { z } from 'zod'
import {
  CDP_PORT,
  cdpEval,
  cdpOpenPage,
  cdpScreenshot,
  cdpTargets,
  cdpVersion,
  isMac,
  killGenOffice,
  launchGenOffice,
  waitForCdp,
} from '../cdp.js'

// Fork: o conector detecta HermesOffice (fork) primeiro, com fallback para o
// GenOffice original — assim funciona com os dois instalados.
const APP_NAMES = ['HermesOffice', 'GenOffice'] as const

function detectAppName(): string {
  for (const name of APP_NAMES) {
    if (isMac()) {
      if (existsSync(`/Applications/${name}.app`)) return name
    } else if (existsSync(join(process.env.LOCALAPPDATA ?? '', 'Programs', name))) {
      return name
    }
  }
  return APP_NAMES[0]
}

function appInstalled(): boolean {
  return APP_NAMES.some((name) => {
    if (isMac()) return existsSync(`/Applications/${name}.app`)
    return existsSync(join(process.env.LOCALAPPDATA ?? '', 'Programs', name))
  })
}

function appVersion(): string | null {
  if (!isMac()) return null
  const name = detectAppName()
  try {
    const out = execSync(
      `/usr/libexec/PlistBuddy -c "Print :CFBundleShortVersionString" /Applications/${name}.app/Contents/Info.plist`,
    )
    return out.toString().trim()
  } catch {
    return null
  }
}

export function registerAppTools(server: McpServer): void {
  server.registerTool(
    'genoffice_app_status',
    {
      title: 'GenOffice app status',
      description:
        'Check whether the GenOffice desktop app is installed and whether its CDP ' +
        'debug port is up. Returns the app bundle version, the running process ' +
        'version (when the debugger is live) and the current page title/url. ' +
        'Read-only; does not launch anything.',
      inputSchema: {},
    },
    async () => {
      const installed = appInstalled()
      const ver = appVersion()
      const version = await cdpVersion()
      const targets = await cdpTargets()
      const page = targets[0]
      return {
        content: [
          {
            type: 'text' as const,
            text:
              `App instalado: ${installed ? 'sim' : 'não'}${ver ? ` (v${ver})` : ''}\n` +
              `CDP porta ${CDP_PORT}: ${version ? 'ATIVO' : 'inativo'}\n` +
              (version
                ? `  debugger: ${version['Browser'] ?? ''}\n` +
                  (page ? `  página: "${page.title}" — ${page.url}\n` : '')
                : '  (use genoffice_app_launch para iniciar com debug)'),
          },
        ],
      }
    },
  )

  server.registerTool(
    'genoffice_app_launch',
    {
      title: 'GenOffice app launch (CDP)',
      description:
        'Launch the installed GenOffice app with the CDP debug port open so the ' +
        'agent can drive it (screenshots, DOM evaluation). If the app is already ' +
        'running WITHOUT the debug port, it is terminated first (single-instance ' +
        'lock would otherwise swallow the relaunch) — set killExisting=false to ' +
        'abort instead. Handles the auto-updater relaunch by retrying until the ' +
        'port answers. macOS uses `open -a`; other platforms launch the binary.',
      inputSchema: {
        killExisting: z
          .boolean()
          .optional()
          .describe('Kill a running instance without CDP first (default true)'),
      },
    },
    async ({ killExisting }) => {
      if (!appInstalled()) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: 'GenOffice não está instalado (procure /Applications/GenOffice.app). Instale o app primeiro.',
            },
          ],
        }
      }
      if (await cdpVersion()) {
        return {
          content: [
            { type: 'text' as const, text: `App já está rodando com CDP na porta ${CDP_PORT}.` },
          ],
        }
      }
      if (killExisting !== false) {
        await killGenOffice()
        await new Promise((r) => setTimeout(r, 1000))
      }
      // up to 3 attempts — covers auto-updater relaunch dropping the flag
      for (let attempt = 1; attempt <= 3; attempt++) {
        await launchGenOffice()
        if (await waitForCdp(CDP_PORT, 25_000)) {
          const targets = await cdpTargets()
          return {
            content: [
              {
                type: 'text' as const,
                text:
                  `App lançado com CDP na porta ${CDP_PORT} (tentativa ${attempt})\n` +
                  (targets[0] ? `página: "${targets[0].title}" — ${targets[0].url}` : ''),
              },
            ],
          }
        }
        await killGenOffice()
        await new Promise((r) => setTimeout(r, 1500))
      }
      return {
        isError: true,
        content: [
          {
            type: 'text' as const,
            text: `Não consegui abrir o CDP após 3 tentativas. O app pode ter baixado um update e relançado sem a flag — tente novamente.`,
          },
        ],
      }
    },
  )

  server.registerTool(
    'genoffice_app_open_file',
    {
      title: 'GenOffice app open file',
      description:
        'Open an Office/PDF file in the GenOffice desktop app via the macOS `open` ' +
        'command (the app registers .docx/.xlsx/.pptx/.pdf document types). The app ' +
        'does not need to be running with CDP for this — it launches normally. ' +
        'macOS only.',
      inputSchema: {
        path: z.string().describe('Absolute path to the file to open'),
      },
    },
    async ({ path }) => {
      if (!isMac()) {
        return {
          isError: true,
          content: [
            { type: 'text' as const, text: 'genoffice_app_open_file usa `open -a` — macOS only.' },
          ],
        }
      }
      if (!existsSync(path)) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `Arquivo não encontrado: ${path}. Verifique o caminho absoluto.`,
            },
          ],
        }
      }
      const appName = detectAppName()
      return await new Promise((resolve) => {
        execFile('open', ['-a', appName, path], (err) => {
          if (err) {
            resolve({
              isError: true,
              content: [{ type: 'text' as const, text: `open -a falhou: ${err.message}` }],
            })
          } else {
            resolve({
              content: [
                { type: 'text' as const, text: `Abrindo ${path} no ${appName}…` },
              ],
            })
          }
        })
      })
    },
  )

  server.registerTool(
    'genoffice_app_screenshot',
    {
      title: 'GenOffice app screenshot',
      description:
        'Capture a PNG screenshot of the GenOffice app window over CDP and save it ' +
        'to outPath. Requires the app running with the debug port (genoffice_app_launch). ' +
        'Use together with vision to inspect the current UI state.',
      inputSchema: {
        outPath: z.string().describe('Absolute path for the .png (e.g. /tmp/genoffice.png)'),
      },
    },
    async ({ outPath }) => {
      try {
        const ws = await cdpOpenPage()
        try {
          await cdpScreenshot(ws, outPath)
        } finally {
          ws.close()
        }
        return {
          content: [
            {
              type: 'text' as const,
              text: `Screenshot salvo → ${outPath} (abra com vision p/ inspecionar)`,
            },
          ],
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_app_screenshot failed: ${(err as Error).message}. ` +
                `O app está rodando com CDP? (genoffice_app_launch)`,
            },
          ],
        }
      }
    },
  )

  server.registerTool(
    'genoffice_app_eval',
    {
      title: 'GenOffice app eval (DOM)',
      description:
        'Evaluate a JavaScript expression in the GenOffice app page context over CDP ' +
        '(DOM reads/writes). Read-only by default (mutate=true required to change ' +
        'the page). Use for inspecting UI state that the read tools do not cover. ' +
        'Requires the app running with the debug port.',
      inputSchema: {
        expression: z
          .string()
          .describe('JS expression to evaluate in the page (e.g. document.title)'),
        mutate: z
          .boolean()
          .optional()
          .describe('Allow page mutation (default false — read-only)'),
      },
    },
    async ({ expression, mutate }) => {
      if (mutate) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: 'mutate=true ainda não é permitido — use as tools genoffice_* para editar arquivos.',
            },
          ],
        }
      }
      try {
        const ws = await cdpOpenPage()
        try {
          const r = await cdpEval(ws, expression)
          if (r.exception) {
            return {
              isError: true,
              content: [{ type: 'text' as const, text: `exceção: ${r.exception}` }],
            }
          }
          return {
            content: [
              {
                type: 'text' as const,
                text:
                  typeof r.value === 'string'
                    ? r.value
                    : JSON.stringify(r.value ?? r.description ?? 'undefined', null, 2).slice(0, 4000),
              },
            ],
          }
        } finally {
          ws.close()
        }
      } catch (err) {
        return {
          isError: true,
          content: [
            {
              type: 'text' as const,
              text: `genoffice_app_eval failed: ${(err as Error).message}. ` +
                `O app está rodando com CDP? (genoffice_app_launch)`,
            },
          ],
        }
      }
    },
  )
}
