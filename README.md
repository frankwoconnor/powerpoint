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

## Design rules

- 16:9 slides with consistent margins
- restrained Gartner palette
- automatic text shrinking inside bounded boxes (md2ppt.js) / native placeholder autofit (md2ppt.py)
- no arbitrary Markdown positioning
