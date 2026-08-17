# Gartner-style Markdown to PowerPoint

A constrained Markdown-to-PowerPoint generator using PptxGenJS and the attached Gartner template as its visual reference.

## Use

```bash
npm install
node md2ppt.js markdown-template.md output.pptx
```

Supported layouts: title, bullets, two-column and takeaway. The authoring template also documents a table convention for future extension.

## Design rules

- 16:9 slides with consistent margins
- restrained Gartner palette
- fixed title, content and classification-footer regions
- automatic text shrinking inside bounded boxes
- no arbitrary Markdown positioning
- Gartner Sans can replace Arial when installed locally
