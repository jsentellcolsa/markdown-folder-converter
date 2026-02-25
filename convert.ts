#!/usr/bin/env tsx
/**
 * convert.ts — Converts .docx files to Markdown.
 *
 * Usage:
 *   tsx convert.ts [inputDir] [outputDir]
 *
 * Defaults:
 *   inputDir  = ./example-fast-process-structure
 *   outputDir = ./docs
 *
 * .docx → Markdown  (via mammoth + turndown)
 *
 * The folder/filename structure of inputDir is mirrored in outputDir.
 */

import fs from 'fs';
import path from 'path';
import mammoth from 'mammoth';
import TurndownService from 'turndown';
import { gfm } from 'turndown-plugin-gfm';

// ─── Config ──────────────────────────────────────────────────────────────────

const INPUT_DIR  = path.resolve(process.argv[2] ?? './example-fast-process-structure');
const OUTPUT_DIR = path.resolve(process.argv[3] ?? './docs');

const turndown = new TurndownService({
  headingStyle:     'atx',
  codeBlockStyle:   'fenced',
  bulletListMarker: '-',
});
turndown.use(gfm);

// ─── Utilities ───────────────────────────────────────────────────────────────

/** Makes a string safe for use as a URL path segment. */
function slugify(name: string): string {
  return name
    .toLowerCase()
    .replace(/\s+/g, '-')           // spaces → dashes
    .replace(/[^a-z0-9\-_.]/g, '-') // non-URL-safe chars → dashes
    .replace(/-{2,}/g, '-')         // collapse consecutive dashes
    .replace(/^-+|-+$/g, '');       // trim leading/trailing dashes
}

/** Applies slugify to every segment of a relative path, preserving the extension. */
function slugifyRelativePath(relPath: string): string {
  const segments = relPath.split(path.sep);
  return segments.map((seg, i) => {
    const isLast = i === segments.length - 1;
    if (isLast) {
      const ext = path.extname(seg);
      const base = path.basename(seg, ext);
      return slugify(base) + ext;
    }
    return slugify(seg);
  }).join('/');
}

function mkdirp(dir: string): void {
  fs.mkdirSync(dir, { recursive: true });
}

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

// ─── .docx → Markdown ────────────────────────────────────────────────────────

async function convertDocx(filePath: string): Promise<string> {
  const { value: html, messages } = await mammoth.convertToHtml({ path: filePath });
  const warnings = messages.filter(m => m.type === 'warning');
  if (warnings.length) {
    console.warn(`  [docx] ${warnings.length} warning(s) in ${path.basename(filePath)}`);
  }
  return turndown.turndown(html ?? '');
}

// ─── Main ─────────────────────────────────────────────────────────────────────

async function main(): Promise<void> {
  if (!fs.existsSync(INPUT_DIR)) {
    console.error(`Input directory not found: ${INPUT_DIR}`);
    process.exit(1);
  }

  if (fs.existsSync(OUTPUT_DIR)) {
    fs.rmSync(OUTPUT_DIR, { recursive: true });
  }
  mkdirp(OUTPUT_DIR);

  const docxFiles = walk(INPUT_DIR).filter(f => {
    if (path.basename(f).startsWith('~$')) return false;
    return path.extname(f).toLowerCase() === '.docx';
  });

  if (!docxFiles.length) {
    console.log('No .docx files found.');
    return;
  }

  console.log(`Found ${docxFiles.length} .docx file(s) to convert.\n`);

  let ok = 0;
  let fail = 0;

  for (const filePath of docxFiles) {
    const ext      = path.extname(filePath);
    const relative = path.relative(INPUT_DIR, filePath);
    const outRel   = slugifyRelativePath(relative.replace(/\.docx$/i, '.md'));
    const outPath  = path.join(OUTPUT_DIR, outRel);
    const title    = path.basename(filePath, ext);
    const baseName = path.basename(filePath);

    mkdirp(path.dirname(outPath));
    process.stdout.write(`  ${relative}  →  ${outRel} … `);

    try {
      const body = await convertDocx(filePath);

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

  console.log(`\nDone. ${ok} converted, ${fail} failed.`);
  console.log(`Output → ${OUTPUT_DIR}`);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
