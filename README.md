# Gartner-style Markdown to PowerPoint

Two Markdown-to-PowerPoint generators sharing the same `markdown-template.md` grammar, backed by `gartner_template.pptx`.

## md2ppt.py (recommended)

Opens `gartner_template.pptx` directly and builds each slide from its real slide layouts (placeholders, theme colors, Gartner Sans font) — the output inherits the actual corporate template rather than an approximation of it.

```bash
python3 -m venv .venv && .venv/bin/pip install -r requirements.txt
.venv/bin/python md2ppt.py markdown-template.md output.pptx
```

Supported layouts: `title`, `bullets`, `two-column`, `three-column`, `four-column`, `takeaway`, `divider`, `table`.

## md2ppt.js

A from-scratch PptxGenJS generator that hand-draws shapes styled to match the template's palette, rather than using the template file itself. Kept for a lighter, dependency-free (Node-only) path.

```bash
npm install
node md2ppt.js markdown-template.md output.pptx
```

Supported layouts: `title`, `bullets`, `two-column`, `takeaway`, `table`. `three-column`/`four-column`/`divider` fall back to the plain bullets layout.

## Markdown format

A deck is a `markdown-template.md`-style file: `---`-separated sections. The first section (before any `type:` line) is deck-level metadata; every section after that is one slide.

Deck metadata fields:

| Field | Used by |
|---|---|
| `title` | Both — doc title / default title-slide title |
| `subtitle` | Both — default title-slide subtitle |
| `presenter` | Both — doc author |
| `date` | Neither renders it on a slide; kept for author reference |
| `classification` | md2ppt.js only, as a footer stamp per slide. md2ppt.py's layouts have no footer placeholder, so it's currently unused there. |

Each slide section starts with `type:` (defaults to `bullets` if omitted), then `title:`, then a body whose shape depends on the type:

| `type` | Body format |
|---|---|
| `title` | `subtitle:` field; falls back to deck `title`/`subtitle` if omitted |
| `bullets` | `- ` / `* ` lines |
| `two-column`, `three-column`, `four-column` | N `::: column` blocks, each starting with an optional `## Heading` line followed by `- ` bullets |
| `table` | A GitHub-style pipe table (header row, `---` alignment row, data rows) |
| `takeaway` | A single sentence — the body text |
| `divider` | Title only, no body — section-break slide |

`three-column`/`four-column`/`divider` are only fully rendered by `md2ppt.py`; `md2ppt.js` falls back to a plain bullets slide for them.

## Design rules

- 16:9 slides with consistent margins
- restrained Gartner palette
- automatic text shrinking inside bounded boxes (md2ppt.js) / native placeholder autofit (md2ppt.py)
- no arbitrary Markdown positioning
