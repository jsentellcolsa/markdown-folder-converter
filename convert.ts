#!/usr/bin/env tsx
/**
 * convert.ts — Converts .docx and .pptx files to Markdown for VitePress.
 *
 * Usage:
 *   tsx convert.ts [inputDir] [outputDir]
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

import fs from 'fs';
import path from 'path';
import mammoth from 'mammoth';
import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';
import JSZip from 'jszip';
import { parseStringPromise } from 'xml2js';

// ─── Config ──────────────────────────────────────────────────────────────────

const INPUT_DIR  = path.resolve(process.argv[2] ?? './example-fast-process-structure');
const OUTPUT_DIR = path.resolve(process.argv[3] ?? './docs');

const turndown = new TurndownService({
  headingStyle:     'atx',
  codeBlockStyle:   'fenced',
  bulletListMarker: '-',
});
turndown.use(gfm); // enables GFM tables, strikethrough, task lists

// ─── Utilities ───────────────────────────────────────────────────────────────

/** Recursively create a directory if it doesn't exist. */
function mkdirp(dir: string): void {
  fs.mkdirSync(dir, { recursive: true });
}

/** Walk a directory tree, returning every file path. */
function walk(dir: string): string[] {
  const results: string[] = [];
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

/**
 * Safely get a deeply nested value from an xml2js-parsed object.
 * xml2js (explicitArray:true) wraps everything in arrays.
 */
function get(obj: unknown, ...keys: string[]): unknown {
  let cur: unknown = obj;
  for (const k of keys) {
    if (cur == null) return undefined;
    if (Array.isArray(cur)) cur = (cur as unknown[])[0];
    cur = (cur as Record<string, unknown>)?.[k];
  }
  return Array.isArray(cur) ? (cur as unknown[])[0] : cur;
}

/** Collect all values matching a key anywhere in an xml2js tree. */
function findAll(obj: unknown, targetKey: string): unknown[] {
  const results: unknown[] = [];
  function recurse(node: unknown): void {
    if (node == null || typeof node !== 'object') return;
    if (Array.isArray(node)) { (node as unknown[]).forEach(recurse); return; }
    for (const [k, v] of Object.entries(node as Record<string, unknown>)) {
      if (k === targetKey) {
        ([] as unknown[]).concat(v).forEach(item => results.push(item));
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
 */
async function convertDocx(filePath: string): Promise<string> {
  const { value: html, messages } = await mammoth.convertToHtml({ path: filePath });

  const warnings = messages.filter(m => m.type === 'warning');
  if (warnings.length) {
    console.warn(`  [docx] ${warnings.length} warning(s) in ${path.basename(filePath)}`);
  }

  return turndown.turndown(html ?? '');
}

// ─── .pptx → Markdown ────────────────────────────────────────────────────────

// xml2js produces deeply dynamic structures; use a loose type for parsed nodes.
type XmlNode = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

interface ParsedParagraph {
  text: string;
  level: number;
}

interface SidebarItem {
  text: string;
  link?: string;
  collapsed?: boolean;
  items?: SidebarItem[];
}

interface SlideData {
  title: string;
  bodyBlocks: string[];
}

/**
 * Extract plain text from an xml2js paragraph node (`a:p`).
 * Returns { text, level } where level is the bullet indent (0-based).
 */
function parseParagraph(para: unknown): ParsedParagraph {
  if (!para || typeof para !== 'object') return { text: '', level: 0 };
  const p = para as XmlNode;

  // Indent level from <a:pPr lvl="N"/>
  const pPr = get(para, 'a:pPr') as XmlNode | undefined;
  const level = pPr?.['$']?.lvl ? parseInt(pPr['$'].lvl as string, 10) : 0;

  // Collect text from all <a:r><a:t> and <a:fld><a:t> runs
  const runs: XmlNode[] = [
    ...([] as XmlNode[]).concat(p['a:r'] ?? []),
    ...([] as XmlNode[]).concat(p['a:fld'] ?? []),
  ];

  const text = runs
    .map(run => {
      const tNodes: unknown[] = ([] as unknown[]).concat(run['a:t'] ?? []);
      return tNodes.map(t => (typeof t === 'string' ? t : (t as XmlNode)?._ ?? '')).join('');
    })
    .join('');

  return { text, level };
}

/**
 * Convert an array of paragraph objects into Markdown lines.
 * Bullet/indent levels map to nested markdown list items.
 */
function paragraphsToMarkdown(paragraphs: unknown[]): string {
  const lines: string[] = [];
  for (const para of paragraphs) {
    const { text, level } = parseParagraph(para);
    if (!text.trim()) {
      lines.push('');
      continue;
    }
    const indent = '  '.repeat(level);
    lines.push(`${indent}- ${text}`);
  }
  // Collapse runs of blank lines to a single blank line
  return lines.join('\n').replace(/\n{3,}/g, '\n\n');
}

/**
 * Parse a single slide XML into { title, bodyBlocks }.
 */
function parseSlideDoc(slideDoc: unknown): SlideData {
  const spTree = get(slideDoc, 'p:sld', 'p:cSld', 'p:spTree') as XmlNode | undefined;
  if (!spTree) return { title: '', bodyBlocks: [] };

  const shapes: XmlNode[] = ([] as XmlNode[]).concat(spTree['p:sp'] ?? []);
  let title = '';
  const bodyBlocks: string[] = [];

  for (const sp of shapes) {
    const ph = get(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph') as XmlNode | undefined;
    const phType: string = ph?.['$']?.type ?? '';

    const txBody = get(sp, 'p:txBody') as XmlNode | undefined;
    if (!txBody) continue;

    const paragraphs: unknown[] = ([] as unknown[]).concat(txBody['a:p'] ?? []);
    const texts = paragraphs.map(p => parseParagraph(p).text).filter(Boolean);

    if (!texts.length) continue;

    const isTitle = phType === 'title' || phType === 'ctrTitle' || phType === 'subTitle';

    if (isTitle && !title) {
      title = texts.join(' ');
    } else {
      const block = paragraphsToMarkdown(paragraphs);
      if (block.trim()) bodyBlocks.push(block);
    }
  }

  return { title, bodyBlocks };
}

/**
 * Extract speaker-notes text from a notesSlide XML document.
 * The shape with idx="1" is the actual notes text body.
 */
function parseNotesDoc(notesDoc: unknown): string {
  const spTree = get(notesDoc, 'p:notes', 'p:cSld', 'p:spTree') as XmlNode | undefined;
  if (!spTree) return '';

  const shapes: XmlNode[] = ([] as XmlNode[]).concat(spTree['p:sp'] ?? []);
  const noteLines: string[] = [];

  for (const sp of shapes) {
    const ph = get(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph') as XmlNode | undefined;
    const phIdx: string | undefined = ph?.['$']?.idx;
    if (phIdx !== '1') continue;

    const txBody = get(sp, 'p:txBody') as XmlNode | undefined;
    if (!txBody) continue;

    const paragraphs: unknown[] = ([] as unknown[]).concat(txBody['a:p'] ?? []);
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
 */
async function convertPptx(filePath: string): Promise<string> {
  const raw = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(raw);

  // Sort slide files numerically
  const slideKeys = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const slideNum = (key: string) => parseInt(key.match(/slide(\d+)/)![1], 10);
      return slideNum(a) - slideNum(b);
    });

  const sections: string[] = [];

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
 */
function buildSidebar(outDir: string, baseUrl = '/'): string {
  function buildItems(dir: string, urlBase: string): SidebarItem[] {
    const entries = fs.readdirSync(dir, { withFileTypes: true })
      .filter(e => e.name !== 'sidebar.ts' && e.name !== 'index.md')
      .sort((a, b) => {
        if (a.isDirectory() !== b.isDirectory()) return a.isDirectory() ? -1 : 1;
        return a.name.localeCompare(b.name);
      });

    return entries.flatMap((entry): SidebarItem[] => {
      const fullPath = path.join(dir, entry.name);
      const urlPath  = `${urlBase}${entry.name}`;

      if (entry.isDirectory()) {
        return [{
          text:      entry.name,
          collapsed: true,
          items:     buildItems(fullPath, urlPath + '/'),
        }];
      }

      if (entry.name.endsWith('.md')) {
        const text = entry.name.replace(/\.md$/, '');
        return [{ text, link: urlPath.replace(/\.md$/, '') }];
      }

      if (entry.name.endsWith('.xlsx')) {
        const text = entry.name.replace(/\.xlsx$/, '');
        return [{ text: `📥 ${text}`, link: urlPath }];
      }

      return [];
    });
  }

  const items = buildItems(outDir, baseUrl);
  const json  = JSON.stringify(items, null, 2);

  return `// Auto-generated by convert.ts — do not edit by hand.
// Place this in your VitePress config's sidebar field, e.g.:
//
//   import sidebar from './sidebar'
//   export default defineConfig({ themeConfig: { sidebar } })

export default ${json} as const;
`;
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  if (!fs.existsSync(INPUT_DIR)) {
    console.error(`Input directory not found: ${INPUT_DIR}`);
    process.exit(1);
  }

  mkdirp(OUTPUT_DIR);

  const allFiles = walk(INPUT_DIR).filter(f => {
    if (path.basename(f).startsWith('~$')) return false; // skip Office temp/lock files
    const ext = path.extname(f).toLowerCase();
    return ext === '.docx' || ext === '.pptx' || ext === '.xlsx';
  });

  if (!allFiles.length) {
    console.log('No .docx, .pptx, or .xlsx files found.');
    return;
  }

  const convertFiles = allFiles.filter(f => path.extname(f).toLowerCase() !== '.xlsx');
  const xlsxFiles    = allFiles.filter(f => path.extname(f).toLowerCase() === '.xlsx');

  // Copy Excel files as-is so VitePress serves them as static downloads
  for (const filePath of xlsxFiles) {
    const relative = path.relative(INPUT_DIR, filePath);
    const outPath  = path.join(OUTPUT_DIR, relative);
    mkdirp(path.dirname(outPath));
    fs.copyFileSync(filePath, outPath);
    console.log(`  ${relative}  →  ${relative} (copied)`);
  }

  if (convertFiles.length) console.log();

  console.log(`Found ${convertFiles.length} file(s) to convert${xlsxFiles.length ? `, ${xlsxFiles.length} Excel file(s) copied` : ''}.\n`);

  let ok = 0;
  let fail = 0;

  for (const filePath of convertFiles) {
    const ext      = path.extname(filePath).toLowerCase();
    const relative = path.relative(INPUT_DIR, filePath);
    const outRel   = relative.replace(/\.(docx|pptx)$/i, '.md');
    const outPath  = path.join(OUTPUT_DIR, outRel);

    mkdirp(path.dirname(outPath));

    const baseName = path.basename(filePath);
    const title    = path.basename(filePath, ext);

    process.stdout.write(`  ${relative}  →  ${outRel} … `);

    try {
      const body = ext === '.docx'
        ? await convertDocx(filePath)
        : await convertPptx(filePath);

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
      console.log(`FAILED\n    ${(err as Error).message}`);
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
