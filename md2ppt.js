#!/usr/bin/env node
const fs = require('fs');
const PptxGenJS = require('pptxgenjs');

const COLORS = {
  navy: '002869',
  cobalt: '0057AF',
  opal: '1F96FF',
  crystal: 'E0F1FF',
  onyx: '00002D',
  gray: '707996',
  white: 'FFFFFF',
};
const FONT = 'Arial';

// Parses a leading "key: value" block (YAML-lite) off the top of a string.
// Returns [metadata, remainingBody].
function parseMeta(section) {
  const meta = {};
  const lines = section.trim().split(/\n/);
  while (lines.length && /^[\w-]+:/.test(lines[0])) {
    const match = lines.shift().match(/^([\w-]+):\s*(.*)$/);
    meta[match[1]] = match[2];
  }
  return [meta, lines.join('\n').trim()];
}

// Splits the markdown deck on "---" separators into deck metadata + slides.
function parseDeck(md) {
  const sections = md.trim().split(/^---\s*$/m).map(s => s.trim()).filter(Boolean);
  let deckMeta = {};
  if (sections[0] && !sections[0].includes('\ntype:')) {
    [deckMeta] = parseMeta(sections.shift());
  }
  const slides = sections.map(section => {
    const [meta, body] = parseMeta(section);
    meta.body = body;
    return meta;
  });
  return [deckMeta, slides];
}

// Extracts "- item" / "* item" lines from a markdown body as plain strings.
function parseBullets(body) {
  return (body || '')
    .split(/\n/)
    .filter(line => /^\s*[-*]\s+/.test(line))
    .map(line => line.replace(/^\s*[-*]\s+/, ''));
}

// Converts bullet strings into pptxgenjs text-run objects.
function bulletRuns(items) {
  return items.map((text, idx) => ({
    text,
    options: { bullet: { indent: 18 }, breakLine: idx < items.length - 1 },
  }));
}

// Parses a GitHub-flavored markdown table (header, "---" alignment row, data
// rows) into { headerCells, aligns, dataRows }. Returns null if body isn't a
// pipe table.
function parseTable(body) {
  const rows = (body || '')
    .split(/\n/)
    .map(line => line.trim())
    .filter(line => line.startsWith('|'));
  if (rows.length < 2) return null;

  const splitRow = line => line.replace(/^\||\|$/g, '').split('|').map(cell => cell.trim());
  const headerCells = splitRow(rows[0]);
  const aligns = splitRow(rows[1]).map(cell => {
    const left = cell.startsWith(':');
    const right = cell.endsWith(':');
    if (left && right) return 'center';
    if (right) return 'right';
    return 'left';
  });
  const dataRows = rows.slice(2).map(splitRow);
  return { headerCells, aligns, dataRows };
}

function addFooter(slide, classification, pageNum) {
  slide.addShape('rect', {
    x: 0.55, y: 7.14, w: 12.2, h: 0.012,
    fill: { color: COLORS.gray }, line: { transparency: 100 },
  });
  slide.addText(classification || 'RESTRICTED', {
    x: 0.58, y: 7.18, w: 2, h: 0.15,
    fontFace: FONT, fontSize: 7, color: COLORS.gray, margin: 0,
  });
  slide.addText(String(pageNum), {
    x: 12.35, y: 7.16, w: 0.4, h: 0.16,
    fontFace: FONT, fontSize: 8, color: COLORS.gray, align: 'right', margin: 0,
  });
}

// Adds a standard slide (title + rule + footer) and returns it.
function addBaseSlide(pres, title, classification, pageNum, bg = COLORS.white) {
  const slide = pres.addSlide();
  slide.background = { color: bg };
  slide.addText(title || '', {
    x: 0.72, y: 0.45, w: 11.85, h: 0.72,
    fontFace: FONT, fontSize: 28, bold: true, color: COLORS.onyx, margin: 0, fit: 'shrink',
  });
  slide.addShape('rect', {
    x: 0.72, y: 1.3, w: 1, h: 0.035,
    fill: { color: COLORS.cobalt }, line: { transparency: 100 },
  });
  addFooter(slide, classification, pageNum);
  return slide;
}

function addTitleSlide(pres, slideData, deckMeta, pageNum) {
  const slide = pres.addSlide();
  slide.background = { color: COLORS.navy };
  slide.addText(slideData.title || deckMeta.title || '', {
    x: 1, y: 1.3, w: 11.2, h: 2,
    fontFace: FONT, fontSize: 35, bold: true, color: COLORS.white, margin: 0, fit: 'shrink',
  });
  slide.addText(slideData.subtitle || deckMeta.subtitle || '', {
    x: 1, y: 4, w: 10, h: 0.7,
    fontFace: FONT, fontSize: 18, color: 'B8DEFF', margin: 0, fit: 'shrink',
  });
  addFooter(slide, slideData.classification || deckMeta.classification, pageNum);
}

function addTakeawaySlide(pres, slideData, deckMeta, pageNum) {
  const slide = addBaseSlide(pres, slideData.title, slideData.classification || deckMeta.classification, pageNum, COLORS.crystal);
  slide.addShape('roundRect', {
    x: 1.15, y: 1.75, w: 11, h: 3.8,
    fill: { color: COLORS.navy }, line: { color: COLORS.navy },
  });
  slide.addText(slideData.body, {
    x: 1.7, y: 2.3, w: 9.9, h: 2.3,
    fontFace: FONT, fontSize: 28, bold: true, color: COLORS.white,
    align: 'center', valign: 'mid', fit: 'shrink',
  });
}

function addTwoColumnSlide(pres, slideData, deckMeta, pageNum) {
  const slide = addBaseSlide(pres, slideData.title, slideData.classification || deckMeta.classification, pageNum);
  const allColumns = (slideData.body || '').split(/^::: column\s*$/m).map(c => c.trim()).filter(Boolean);
  if (allColumns.length > 2) {
    console.warn(`Slide ${pageNum}: two-column layout has ${allColumns.length} "::: column" blocks; only the first 2 are rendered.`);
  }
  const columns = allColumns.slice(0, 2);
  columns.forEach((column, j) => {
    const lines = column.split(/\n/);
    const heading = (lines.shift() || '').replace(/^#+\s*/, '');
    const items = parseBullets(lines.join('\n'));
    slide.addShape('roundRect', {
      x: 0.72 + j * 6.11, y: 1.58, w: 5.78, h: 4.95,
      fill: { color: j ? COLORS.crystal : 'F4F8FC' }, line: { color: 'B8DEFF' },
    });
    slide.addText(heading, {
      x: 1 + j * 6.11, y: 1.88, w: 5.2, h: 0.5,
      fontFace: FONT, fontSize: 21, bold: true, color: COLORS.cobalt, fit: 'shrink',
    });
    slide.addText(bulletRuns(items), {
      x: 1 + j * 6.11, y: 2.55, w: 5.05, h: 3.5,
      fontFace: FONT, fontSize: 17, color: COLORS.onyx, fit: 'shrink',
    });
  });
}

function addTableSlide(pres, slideData, deckMeta, pageNum) {
  const slide = addBaseSlide(pres, slideData.title, slideData.classification || deckMeta.classification, pageNum);
  const table = parseTable(slideData.body);
  if (!table) {
    console.warn(`Slide ${pageNum}: type: table but no markdown table found in body; rendering as plain text.`);
    slide.addText(slideData.body || '', {
      x: 0.88, y: 1.65, w: 11.3, h: 4.9,
      fontFace: FONT, fontSize: 20, color: COLORS.onyx, margin: 0.05, fit: 'shrink',
    });
    return;
  }

  const { headerCells, aligns, dataRows } = table;
  const headerRow = headerCells.map((text, i) => ({
    text,
    options: { bold: true, color: COLORS.white, fill: { color: COLORS.navy }, align: aligns[i] || 'left' },
  }));
  const bodyRows = dataRows.map((cells, rowIdx) =>
    cells.map((text, i) => ({
      text,
      options: {
        color: COLORS.onyx,
        fill: { color: rowIdx % 2 ? COLORS.crystal : COLORS.white },
        align: aligns[i] || 'left',
      },
    }))
  );

  slide.addTable([headerRow, ...bodyRows], {
    x: 0.88, y: 1.65, w: 11.3,
    fontFace: FONT, fontSize: 14,
    border: { type: 'solid', color: 'B8DEFF', pt: 0.5 },
    autoPage: false,
  });
}

function addBulletsSlide(pres, slideData, deckMeta, pageNum) {
  const slide = addBaseSlide(pres, slideData.title, slideData.classification || deckMeta.classification, pageNum);
  const items = parseBullets(slideData.body);
  slide.addText(items.length ? bulletRuns(items) : slideData.body, {
    x: 0.88, y: 1.65, w: 11.3, h: 4.9,
    fontFace: FONT, fontSize: 20, color: COLORS.onyx, margin: 0.05, fit: 'shrink', breakLine: false,
  });
}

const SLIDE_BUILDERS = {
  title: addTitleSlide,
  takeaway: addTakeawaySlide,
  'two-column': addTwoColumnSlide,
  table: addTableSlide,
  bullets: addBulletsSlide,
};

async function build(inputPath, outputPath) {
  let markdown;
  try {
    markdown = fs.readFileSync(inputPath, 'utf8');
  } catch (err) {
    throw new Error(`Could not read input file "${inputPath}": ${err.message}`);
  }

  const [deckMeta, slides] = parseDeck(markdown);
  if (!slides.length) {
    throw new Error(`No slides found in "${inputPath}". Each slide must be separated by a line containing only "---".`);
  }

  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = deckMeta.presenter || '';
  pres.title = deckMeta.title || '';

  slides.forEach((slideData, idx) => {
    const pageNum = idx + 1;
    const type = slideData.type || 'bullets';
    const addSlide = SLIDE_BUILDERS[type] || addBulletsSlide;
    addSlide(pres, slideData, deckMeta, pageNum);
  });

  await pres.writeFile({ fileName: outputPath });
}

const inputPath = process.argv[2];
const outputPath = process.argv[3] || 'output.pptx';
if (!inputPath) {
  console.error('Usage: node md2ppt.js input.md output.pptx');
  process.exit(2);
}
build(inputPath, outputPath).catch(err => {
  console.error(err.message || err);
  process.exit(1);
});
