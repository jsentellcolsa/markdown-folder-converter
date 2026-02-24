#!/usr/bin/env node
/**
 * convert.js — Converts .docx and .pptx files to Markdown for VitePress.
 *
 * Usage:
 *   node convert.js [inputDir] [outputDir]
 *
 * Defaults:
 *   inputDir  = ./example-fast-process-structure
 *   outputDir = ./docs
 *
 * .docx → Markdown  (via mammoth + turndown)
 * .pptx → Markdown  (each slide becomes a ## section)
 *
 * The folder/filename structure of inputDir is mirrored in outputDir.
 * A VitePress sidebar config is written to outputDir/sidebar.ts.
 */

'use strict';

const fs = require('fs');
const path = require('path');
const mammoth = require('mammoth');
const TurndownService = require('turndown');
const { gfm } = require('turndown-plugin-gfm');
const JSZip = require('jszip');
const { parseStringPromise } = require('xml2js');

// ─── Config ──────────────────────────────────────────────────────────────────

const INPUT_DIR  = path.resolve(process.argv[2] || './example-fast-process-structure');
const OUTPUT_DIR = path.resolve(process.argv[3] || './docs');

const turndown = new TurndownService({
  headingStyle:    'atx',
  codeBlockStyle:  'fenced',
  bulletListMarker: '-',
});
turndown.use(gfm); // enables GFM tables, strikethrough, task lists

// ─── Utilities ───────────────────────────────────────────────────────────────

/** Recursively create a directory if it doesn't exist. */
function mkdirp(dir) {
  fs.mkdirSync(dir, { recursive: true });
}

/**
 * Walk a directory tree, yielding every file path.
 * @param {string} dir
 * @returns {string[]}
 */
function walk(dir) {
  const results = [];
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) {
      results.push(...walk(full));
    } else {
      results.push(full);
    }
  }
  return results;
}

/** Slugify a string for use as a VitePress link. */
function toLink(str) {
  return str.replace(/\\/g, '/');
}

/**
 * Safely get a deeply nested value from an xml2js-parsed object.
 * xml2js (explicitArray:true) wraps everything in arrays.
 */
function get(obj, ...keys) {
  let cur = obj;
  for (const k of keys) {
    if (cur == null) return undefined;
    cur = Array.isArray(cur) ? cur[0] : cur;
    cur = cur?.[k];
  }
  return Array.isArray(cur) ? cur[0] : cur;
}

/** Collect all values matching a key anywhere in an xml2js tree. */
function findAll(obj, targetKey) {
  const results = [];
  function recurse(node) {
    if (node == null || typeof node !== 'object') return;
    if (Array.isArray(node)) { node.forEach(recurse); return; }
    for (const [k, v] of Object.entries(node)) {
      if (k === targetKey) {
        [].concat(v).forEach(item => results.push(item));
      } else {
        recurse(v);
      }
    }
  }
  recurse(obj);
  return results;
}

// ─── .docx → Markdown ────────────────────────────────────────────────────────

/**
 * Convert a Word document to Markdown.
 * @param {string} filePath  Absolute path to .docx
 * @returns {Promise<string>} Markdown string
 */
async function convertDocx(filePath) {
  const { value: html, messages } = await mammoth.convertToHtml({ path: filePath });

  if (messages.length) {
    const warnings = messages.filter(m => m.type === 'warning');
    if (warnings.length) {
      console.warn(`  [docx] ${warnings.length} warning(s) in ${path.basename(filePath)}`);
    }
  }

  return turndown.turndown(html || '');
}

// ─── .pptx → Markdown ────────────────────────────────────────────────────────

/**
 * Extract plain text from an xml2js paragraph node (`a:p`).
 * Returns an object with { text, level } where level is the bullet indent (0-based).
 */
function parseParagraph(para) {
  if (!para || typeof para !== 'object') return { text: '', level: 0 };

  // Indent level from <a:pPr lvl="N"/>
  const pPr = get(para, 'a:pPr');
  const level = pPr?.['$']?.lvl ? parseInt(pPr['$'].lvl, 10) : 0;

  // Collect text from all <a:r><a:t> and <a:fld><a:t> runs
  const runs = [
    ...[].concat(para['a:r'] ?? []),
    ...[].concat(para['a:fld'] ?? []),
  ];

  const text = runs
    .map(run => {
      const tNodes = [].concat(run['a:t'] ?? []);
      return tNodes.map(t => (typeof t === 'string' ? t : t?._ ?? '')).join('');
    })
    .join('');

  return { text, level };
}

/**
 * Convert an array of paragraph objects into Markdown lines.
 * Bullet/indent levels map to nested markdown list items.
 */
function paragraphsToMarkdown(paragraphs) {
  const lines = [];
  for (const para of paragraphs) {
    const { text, level } = parseParagraph(para);
    if (!text.trim()) {
      lines.push('');
      continue;
    }
    const indent = '  '.repeat(level);
    lines.push(`${indent}- ${text}`);
  }
  // Remove runs of blank lines → single blank line
  return lines.join('\n').replace(/\n{3,}/g, '\n\n');
}

/**
 * Parse a single slide XML into { title, bodyLines, notesLines }.
 * @param {object} slideDoc  Parsed xml2js document for a slide
 * @returns {{ title: string, bodyBlocks: string[] }}
 */
function parseSlideDoc(slideDoc) {
  const spTree = get(slideDoc, 'p:sld', 'p:cSld', 'p:spTree');
  if (!spTree) return { title: '', bodyBlocks: [] };

  const shapes = [].concat(spTree['p:sp'] ?? []);
  let title = '';
  const bodyBlocks = [];

  for (const sp of shapes) {
    // Determine if this shape is a title placeholder
    const ph = get(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    const phType = ph?.['$']?.type ?? '';   // 'title', 'body', 'subTitle', etc.
    const phIdx  = ph?.['$']?.idx;

    const txBody = get(sp, 'p:txBody');
    if (!txBody) continue;

    const paragraphs = [].concat(txBody['a:p'] ?? []);
    const texts = paragraphs
      .map(p => parseParagraph(p).text)
      .filter(Boolean);

    if (!texts.length) continue;

    const isTitle = phType === 'title' || phType === 'ctrTitle' || phType === 'subTitle';
    const isBody  = phType === 'body' || phIdx != null || (!phType && !phIdx);

    if (isTitle && !title) {
      title = texts.join(' ');
    } else {
      // Convert paragraphs to markdown (preserving bullet levels)
      const block = paragraphsToMarkdown(paragraphs);
      if (block.trim()) bodyBlocks.push(block);
    }
  }

  return { title, bodyBlocks };
}

/**
 * Extract speaker-notes text from a notesSlide XML document.
 * The first shape (idx=0) in a notes slide is the slide image placeholder;
 * the second (idx=1) is the actual notes text body.
 */
function parseNotesDoc(notesDoc) {
  const spTree = get(notesDoc, 'p:notes', 'p:cSld', 'p:spTree');
  if (!spTree) return '';

  const shapes = [].concat(spTree['p:sp'] ?? []);
  const noteLines = [];

  for (const sp of shapes) {
    const ph = get(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph');
    const phIdx = ph?.['$']?.idx;
    // idx="1" is the notes text placeholder
    if (phIdx !== '1') continue;

    const txBody = get(sp, 'p:txBody');
    if (!txBody) continue;

    const paragraphs = [].concat(txBody['a:p'] ?? []);
    for (const para of paragraphs) {
      const { text } = parseParagraph(para);
      if (text.trim()) noteLines.push(text);
    }
  }

  return noteLines.join('\n');
}

/**
 * Convert a PowerPoint file to Markdown.
 * Each slide becomes a `##` section. Speaker notes appear as blockquotes.
 * @param {string} filePath  Absolute path to .pptx
 * @returns {Promise<string>} Markdown string
 */
async function convertPptx(filePath) {
  const raw  = fs.readFileSync(filePath);
  const zip  = await JSZip.loadAsync(raw);

  // Sort slide files numerically
  const slideKeys = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const n = f => parseInt(f.match(/slide(\d+)/)[1], 10);
      return n(a) - n(b);
    });

  const sections = [];

  for (let i = 0; i < slideKeys.length; i++) {
    const slideNum = i + 1;
    const slideXml = await zip.files[slideKeys[i]].async('string');
    const slideDoc = await parseStringPromise(slideXml, { explicitArray: true });

    const { title, bodyBlocks } = parseSlideDoc(slideDoc);

    // Speaker notes
    let notes = '';
    const notesKey = `ppt/notesSlides/notesSlide${slideNum}.xml`;
    if (zip.files[notesKey]) {
      const notesXml = await zip.files[notesKey].async('string');
      const notesDoc = await parseStringPromise(notesXml, { explicitArray: true });
      notes = parseNotesDoc(notesDoc);
    }

    // Build section markdown
    const heading = title ? `## ${title}` : `## Slide ${slideNum}`;
    const parts = [heading];

    if (bodyBlocks.length) parts.push(bodyBlocks.join('\n\n'));

    if (notes.trim()) {
      const noteLines = notes.split('\n').map(l => `> ${l}`).join('\n');
      parts.push(`**Notes:**\n\n${noteLines}`);
    }

    sections.push(parts.join('\n\n'));
  }

  return sections.join('\n\n---\n\n');
}

// ─── Sidebar generator ───────────────────────────────────────────────────────

/**
 * Build a VitePress sidebar config from the output directory tree.
 * Returns the TypeScript source as a string.
 * @param {string} outDir   Absolute path to output directory
 * @param {string} baseUrl  URL prefix used by VitePress (e.g. '/')
 */
function buildSidebar(outDir, baseUrl = '/') {
  function buildItems(dir, urlBase) {
    const entries = fs.readdirSync(dir, { withFileTypes: true })
      .filter(e => e.name !== 'sidebar.ts' && e.name !== 'index.md')
      .sort((a, b) => {
        // Directories first, then files
        if (a.isDirectory() !== b.isDirectory()) return a.isDirectory() ? -1 : 1;
        return a.name.localeCompare(b.name);
      });

    return entries.map(entry => {
      const fullPath = path.join(dir, entry.name);
      const urlPath  = `${urlBase}${entry.name}`;

      if (entry.isDirectory()) {
        return {
          text:      entry.name,
          collapsed: true,
          items:     buildItems(fullPath, urlPath + '/'),
        };
      }

      if (entry.name.endsWith('.md')) {
        const text = entry.name.replace(/\.md$/, '');
        return { text, link: urlPath.replace(/\.md$/, '') };
      }

      return null;
    }).filter(Boolean);
  }

  const items = buildItems(outDir, baseUrl);
  const json  = JSON.stringify(items, null, 2);

  return `// Auto-generated by convert.js — do not edit by hand.
// Place this in your VitePress config's sidebar field, e.g.:
//
//   import sidebar from './sidebar'
//   export default defineConfig({ themeConfig: { sidebar } })

export default ${json} as const;
`;
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main() {
  if (!fs.existsSync(INPUT_DIR)) {
    console.error(`Input directory not found: ${INPUT_DIR}`);
    process.exit(1);
  }

  mkdirp(OUTPUT_DIR);

  const files = walk(INPUT_DIR).filter(f => {
    const ext = path.extname(f).toLowerCase();
    return ext === '.docx' || ext === '.pptx';
  });

  if (!files.length) {
    console.log('No .docx or .pptx files found.');
    return;
  }

  console.log(`Found ${files.length} file(s). Converting…\n`);

  let ok = 0, fail = 0;

  for (const filePath of files) {
    const ext      = path.extname(filePath).toLowerCase();
    const relative = path.relative(INPUT_DIR, filePath);
    const outRel   = relative.replace(/\.(docx|pptx)$/i, '.md');
    const outPath  = path.join(OUTPUT_DIR, outRel);

    mkdirp(path.dirname(outPath));

    const baseName = path.basename(filePath);
    const title    = path.basename(filePath, ext);

    process.stdout.write(`  ${relative}  →  ${outRel} … `);

    try {
      let body;
      if (ext === '.docx') {
        body = await convertDocx(filePath);
      } else {
        body = await convertPptx(filePath);
      }

      const frontmatter = [
        '---',
        `title: "${title.replace(/"/g, '\\"')}"`,
        `source: "${baseName}"`,
        '---',
        '',
      ].join('\n');

      fs.writeFileSync(outPath, frontmatter + body + '\n', 'utf8');
      console.log('OK');
      ok++;
    } catch (err) {
      console.log(`FAILED\n    ${err.message}`);
      fail++;
    }
  }

  // Generate sidebar config
  const sidebarPath = path.join(OUTPUT_DIR, 'sidebar.ts');
  fs.writeFileSync(sidebarPath, buildSidebar(OUTPUT_DIR), 'utf8');

  console.log(`\nDone. ${ok} converted, ${fail} failed.`);
  console.log(`Sidebar config → ${sidebarPath}`);
  console.log(`Output        → ${OUTPUT_DIR}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
