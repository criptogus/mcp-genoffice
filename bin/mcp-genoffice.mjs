#!/usr/bin/env node
// npm bin shim: runs the TypeScript server through the tsx loader, which the
// engines need at runtime (they are imported as TS source from a genoffice clone).
import { spawn } from 'node:child_process'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'

const here = dirname(fileURLToPath(import.meta.url))
const entry = join(here, '..', 'src', 'index.ts')
const child = spawn(process.execPath, ['--import', 'tsx', entry], {
  stdio: 'inherit',
})
child.on('exit', (code) => process.exit(code ?? 0))
