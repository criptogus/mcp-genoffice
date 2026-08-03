// Engine loader: resolves a genspark-ai/genoffice clone and dynamically
// imports its pure-TS engine packages. The engines are NOT published to npm,
// so we load them from a clone — either the developer's (GENOFFICE_SRC) or an
// auto-cloned, SHA-pinned checkout in the user cache dir.
import { execSync } from 'node:child_process'
import { existsSync, mkdirSync } from 'node:fs'
import { homedir } from 'node:os'
import { join } from 'node:path'
import { pathToFileURL } from 'node:url'

/** Commit SHA the auto-clone pins to (update on each validated release). */
export const GENOFFICE_PIN = '4da673d4dfa994bd0b4a9bc43430e4a058a17c61'

export const CACHE_SRC = join(homedir(), '.cache', 'mcp-genoffice', 'src')

let resolved: string | null = null

/** Absolute path to a genoffice clone that has the engine packages. */
export function genofficeSrcDir(): string {
  if (resolved) return resolved

  const fromEnv = process.env.GENOFFICE_SRC
  if (fromEnv) {
    if (!existsSync(join(fromEnv, 'packages', 'docx-engine'))) {
      throw new Error(
        `GENOFFICE_SRC=${fromEnv} is not a genspark-ai/genoffice clone (missing packages/docx-engine). ` +
          `Point it at a checkout, or unset it to auto-clone the pinned revision into ${CACHE_SRC}.`,
      )
    }
    resolved = fromEnv
    return resolved
  }

  if (!existsSync(join(CACHE_SRC, 'packages', 'docx-engine'))) {
    ensurePinnedClone()
  }
  resolved = CACHE_SRC
  return resolved
}

/** Shallow-fetch an exact commit (works for any reachable SHA, no full history). */
function ensurePinnedClone(): void {
  mkdirSync(CACHE_SRC, { recursive: true })
  console.error(`[mcp-genoffice] cloning genspark-ai/genoffice@${GENOFFICE_PIN.slice(0, 12)} ...`)
  execSync(`git init -q "${CACHE_SRC}"`, { stdio: 'inherit' })
  execSync(`git -C "${CACHE_SRC}" remote add origin https://github.com/genspark-ai/genoffice.git`, { stdio: 'inherit' })
  execSync(`git -C "${CACHE_SRC}" fetch -q --depth 1 origin ${GENOFFICE_PIN}`, { stdio: 'inherit' })
  execSync(`git -C "${CACHE_SRC}" checkout -q FETCH_HEAD`, { stdio: 'inherit' })
  console.error('[mcp-genoffice] npm install (first run only) ...')
  execSync('npm install --no-audit --no-fund', { cwd: CACHE_SRC, stdio: 'inherit' })
}

/** Dynamically import an engine package's entry (runs its TS source via tsx). */
export async function loadEngine<T>(pkg: string): Promise<T> {
  const dir = genofficeSrcDir()
  const entry = pathToFileURL(join(dir, 'packages', pkg, 'src', 'index.ts')).href
  return (await import(entry)) as T
}
