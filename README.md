# mcp-genoffice

MCP server that gives AI agents **byte-preserving, surgical editing** of Office
documents through the [GenOffice](https://github.com/genspark-ai/genoffice)
engine packages (Apache-2.0, pure TypeScript, no Electron required).

The core guarantee, inherited from GenOffice's architecture: *the original file
is the source of truth*. Only the blocks you edit are regenerated as OOXML
fragments; every untouched paragraph, style, header, comment and zip part keeps
its original bytes. Layout never breaks.

## Tools

| Tool | Description |
| --- | --- |
| `genoffice_extract_text` | Extract readable text from `.docx` / `.xlsx` / `.pptx` / `.pdf` (slides/table structure preserved) |
| `genoffice_docx_blocks` | Parse a `.docx` and list top-level blocks (index, type, style, text) |
| `genoffice_docx_patch` | Rewrite one or more paragraphs with a byte-preserving roundtrip; writes a new file, never touches the original |

More tools (pptx edit, xlsx ops, CDP control of the GenOffice app) are on the
roadmap.

## Install / run

```bash
# from source
npm install
npm start

# as a library in your MCP client (Hermes, Claude Desktop, ...)
npx -y mcp-genoffice
```

Configure in Hermes (`~/.hermes/config.yaml`):

```yaml
mcp_servers:
  genoffice:
    command: "npx"
    args: ["-y", "mcp-genoffice"]
    env:
      # optional: point at your own genoffice clone; otherwise the server
      # auto-clones the pinned revision into ~/.cache/mcp-genoffice/src
      # GENOFFICE_SRC: "/path/to/genoffice"
    timeout: 300
    connect_timeout: 120
```

## How the engines are loaded

The GenOffice engine packages are **not published to npm** and ship as
TypeScript source. The server loads them from a checkout of
`genspark-ai/genoffice` in two ways:

1. **`GENOFFICE_SRC` env var** — use your own clone (fast, dev mode).
2. **Auto-clone** (default) — a shallow, SHA-pinned checkout
   (`GENOFFICE_PIN` in `src/engine.ts`) plus `npm install` on first use.

The server runs under the `tsx` loader so the TS-source engines import cleanly.

## Development

```bash
npm install
npx tsc --noEmit          # typecheck
node tests/client.mjs     # end-to-end: spawns the server, lists tools, patches a fixture
```

The e2e test uses fixture files from a genoffice clone. By default it points at
`/tmp/genoffice`; override with `FIXTURE_SRC`. Unset `SERVER_SRC` to exercise
the auto-clone path.

## Roadmap

- [x] Headless engine mode (extract + docx blocks + docx patch)
- [ ] `genoffice_docx_patch` rich options (styles, headers, comments, watermark)
- [ ] PPTX / XLSX tools (pptx-engine, sheets sidecar)
- [ ] CDP mode: drive the installed GenOffice app (launch, open file, AI edit)
- [ ] Hermes MCP catalog entry (`optional-mcps/genoffice`) + usage skill

## License

MIT. The GenOffice engine packages are Apache-2.0 (loaded at runtime from a
user-provided or auto-cloned checkout, never bundled).
