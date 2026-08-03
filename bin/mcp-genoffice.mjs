#!/usr/bin/env node
// npm bin shim: runs the TypeScript server through the tsx loader, which the
// engines need at runtime (they are imported as TS source from a genoffice clone).
// Uses absolute paths so it works regardless of the parent process cwd.
import { spawn } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const root = join(here, '..')
const tsxCli = join(root, 'node_modules', 'tsx', 'dist', 'cli.mjs')
const entry = join(root, 'src', 'index.ts')
const child = spawn(process.execPath, [tsxCli, entry], {
  stdio: 'inherit',
})
child.on('exit', (code) => process.exit(code ?? 0))
