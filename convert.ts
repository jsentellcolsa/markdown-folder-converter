#!/usr/bin/env tsx
/**
 * convert.ts — Converts .docx and .pptx files for a VitePress site.
 *
 * Usage:
 *   tsx convert.ts [inputDir] [outputDir]
 *
 * Defaults:
 *   inputDir  = ./example-fast-process-structure
 *   outputDir = ./docs
 *
 * .docx → Markdown  (via mammoth + turndown)
 * .pptx → HTML      (self-contained slide viewer with embedded images)
 *                    + a thin .md wrapper page with an <iframe>
 *
 * The folder/filename structure of inputDir is mirrored in outputDir.
 * PPTX HTML files land in outputDir/public/slides/... (served as static assets).
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
turndown.use(gfm);

// ─── Utilities ───────────────────────────────────────────────────────────────

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

function escapeHtml(str: string): string {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/** Resolve a zip-relative path (e.g. '../media/img.png' from 'ppt/slides/'). */
function resolveZipPath(fromKey: string, target: string): string {
  const dir = fromKey.split('/').slice(0, -1).join('/');
  const parts = (dir + '/' + target).split('/');
  const result: string[] = [];
  for (const p of parts) {
    if (p === '..') result.pop();
    else if (p !== '.') result.push(p);
  }
  return result.join('/');
}

function getMimeType(ext: string): string | null {
  const map: Record<string, string> = {
    png:  'image/png',
    jpg:  'image/jpeg',
    jpeg: 'image/jpeg',
    gif:  'image/gif',
    svg:  'image/svg+xml',
    webp: 'image/webp',
    bmp:  'image/bmp',
    tif:  'image/tiff',
    tiff: 'image/tiff',
  };
  return map[ext.toLowerCase()] ?? null;
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

// ─── .pptx → HTML ────────────────────────────────────────────────────────────

// xml2js produces deeply dynamic structures; use a loose type for parsed nodes.
type XmlNode = Record<string, any>; // eslint-disable-line @typescript-eslint/no-explicit-any

interface HtmlShape {
  xPct: number;
  yPct: number;
  wPct: number;
  hPct: number;
  inner: string; // HTML content
}

interface HtmlSlide {
  background: string; // CSS color or ''
  shapes: HtmlShape[];
  notes: string;
}

/** Extract the xfrm (position + size) of a shape in EMU. */
function getTransform(sp: XmlNode, slideW: number, slideH: number): HtmlShape | null {
  const xfrm = get(sp, 'p:spPr', 'a:xfrm') as XmlNode | undefined;
  if (!xfrm) return null;
  const off = get(xfrm, 'a:off') as XmlNode | undefined;
  const ext = get(xfrm, 'a:ext') as XmlNode | undefined;
  if (!off?.['$'] || !ext?.['$']) return null;

  return {
    xPct: (parseInt(off['$'].x, 10) || 0) / slideW * 100,
    yPct: (parseInt(off['$'].y, 10) || 0) / slideH * 100,
    wPct: (parseInt(ext['$'].cx, 10) || 0) / slideW * 100,
    hPct: (parseInt(ext['$'].cy, 10) || 0) / slideH * 100,
    inner: '',
  };
}

/** Convert xml2js paragraph nodes to HTML. */
function paragraphsToHtml(paragraphs: unknown[]): string {
  return paragraphs.map(para => {
    const p = para as XmlNode;

    const pPr = get(para, 'a:pPr') as XmlNode | undefined;
    const algn = pPr?.['$']?.algn;
    const alignCss = algn === 'ctr' ? 'text-align:center;'
                   : algn === 'r'   ? 'text-align:right;'
                   : algn === 'just'? 'text-align:justify;'
                   : '';

    const runs: XmlNode[] = [
      ...([] as XmlNode[]).concat(p['a:r']   ?? []),
      ...([] as XmlNode[]).concat(p['a:fld'] ?? []),
    ];

    if (!runs.length) return `<p style="margin:0;line-height:1.3;">&nbsp;</p>`;

    const runHtml = runs.map(run => {
      const tNodes: unknown[] = ([] as unknown[]).concat(run['a:t'] ?? []);
      const text = tNodes.map(t =>
        typeof t === 'string' ? t : (t as XmlNode)?._ ?? ''
      ).join('');
      if (!text) return '';

      const rPr = get(run, 'a:rPr') as XmlNode | undefined;
      const bold      = rPr?.['$']?.b === '1';
      const italic    = rPr?.['$']?.i === '1';
      const underline = rPr?.['$']?.u && rPr['$'].u !== 'none';
      const szRaw     = rPr?.['$']?.sz;
      const fontSize  = szRaw ? parseInt(szRaw, 10) / 100 : null;

      let color = '';
      const solidFill = get(rPr ?? {}, 'a:solidFill') as XmlNode | undefined;
      if (solidFill) {
        const srgb = get(solidFill, 'a:srgbClr') as XmlNode | undefined;
        if (srgb?.['$']?.val) color = '#' + srgb['$'].val;
        if (!color) {
          const sys = get(solidFill, 'a:sysClr') as XmlNode | undefined;
          if (sys?.['$']?.lastClr) color = '#' + sys['$'].lastClr;
        }
      }

      let style = '';
      if (bold)      style += 'font-weight:bold;';
      if (italic)    style += 'font-style:italic;';
      if (underline) style += 'text-decoration:underline;';
      if (fontSize)  style += `font-size:${fontSize}pt;`;
      if (color)     style += `color:${color};`;

      const safe = escapeHtml(text);
      return style ? `<span style="${style}">${safe}</span>` : safe;
    }).join('');

    return `<p style="margin:0;line-height:1.3;${alignCss}">${runHtml || '&nbsp;'}</p>`;
  }).join('');
}

/** Extract speaker-notes plain text from a notesSlide document. */
function extractNotesText(notesDoc: unknown): string {
  const spTree = get(notesDoc, 'p:notes', 'p:cSld', 'p:spTree') as XmlNode | undefined;
  if (!spTree) return '';

  const shapes: XmlNode[] = ([] as XmlNode[]).concat(spTree['p:sp'] ?? []);
  const lines: string[] = [];

  for (const sp of shapes) {
    const ph = get(sp, 'p:nvSpPr', 'p:nvPr', 'p:ph') as XmlNode | undefined;
    if (ph?.['$']?.idx !== '1') continue;

    const txBody = get(sp, 'p:txBody') as XmlNode | undefined;
    if (!txBody) continue;

    const paragraphs: unknown[] = ([] as unknown[]).concat(txBody['a:p'] ?? []);
    for (const para of paragraphs) {
      const p = para as XmlNode;
      const runs: XmlNode[] = ([] as XmlNode[]).concat(p['a:r'] ?? []);
      const text = runs
        .flatMap(r => ([] as unknown[]).concat(r['a:t'] ?? []))
        .map(t => typeof t === 'string' ? t : (t as XmlNode)?._ ?? '')
        .join('');
      if (text.trim()) lines.push(text);
    }
  }

  return lines.join('\n');
}

/** Extract the solid-fill background color of a slide (hex or ''). */
function extractSlideBackground(slideDoc: unknown): string {
  const bgPr = get(slideDoc, 'p:sld', 'p:cSld', 'p:bg', 'p:bgPr') as XmlNode | undefined;
  if (!bgPr) return '';
  const srgb = get(bgPr, 'a:solidFill', 'a:srgbClr') as XmlNode | undefined;
  if (srgb?.['$']?.val) return '#' + srgb['$'].val;
  return '';
}

/**
 * Parse one slide XML into HtmlSlide.
 * Handles text shapes (p:sp) and picture shapes (p:pic).
 */
function parseSlide(
  slideDoc: unknown,
  rels: Record<string, string>,          // rId → zip path
  mediaCache: Record<string, string>,    // zip path → data URI
  slideW: number,
  slideH: number,
): HtmlSlide {
  const spTree = get(slideDoc, 'p:sld', 'p:cSld', 'p:spTree') as XmlNode | undefined;
  const shapes: HtmlShape[] = [];

  if (spTree) {
    // ── Text shapes ──────────────────────────────────────────────
    for (const sp of ([] as XmlNode[]).concat(spTree['p:sp'] ?? [])) {
      const xfrm = getTransform(sp, slideW, slideH);
      if (!xfrm) continue;

      const txBody = get(sp, 'p:txBody') as XmlNode | undefined;
      if (!txBody) continue;

      const paragraphs: unknown[] = ([] as unknown[]).concat(txBody['a:p'] ?? []);
      const html = paragraphsToHtml(paragraphs);
      if (!html.replace(/&nbsp;/g, '').trim()) continue;

      shapes.push({ ...xfrm, inner: `<div class="tb">${html}</div>` });
    }

    // ── Picture shapes ───────────────────────────────────────────
    for (const pic of ([] as XmlNode[]).concat(spTree['p:pic'] ?? [])) {
      const xfrm = getTransform(pic, slideW, slideH);
      if (!xfrm) continue;

      const blip = get(pic, 'p:blipFill', 'a:blip') as XmlNode | undefined;
      const rId: string | undefined = blip?.['$']?.['r:embed'];
      if (!rId) continue;

      const zipPath = rels[rId];
      const dataUri = zipPath ? mediaCache[zipPath] : undefined;
      if (!dataUri) continue;

      shapes.push({
        ...xfrm,
        inner: `<img src="${dataUri}" style="width:100%;height:100%;object-fit:contain;" alt="">`,
      });
    }
  }

  return {
    background: extractSlideBackground(slideDoc),
    shapes,
    notes: '',
  };
}

/** Render all slides into a self-contained HTML presentation viewer. */
function buildPresentationHtml(
  title: string,
  slides: HtmlSlide[],
  slideW: number,
  slideH: number,
): string {
  // EMU → px at 96 DPI (914400 EMU per inch)
  const wpx = Math.round(slideW / 914400 * 96);
  const hpx = Math.round(slideH / 914400 * 96);

  const slidesHtml = slides.map((slide, i) => {
    const bg = slide.background ? `background:${slide.background};` : 'background:#fff;';
    const shapesHtml = slide.shapes.map(s => `
      <div class="sp" style="left:${s.xPct.toFixed(3)}%;top:${s.yPct.toFixed(3)}%;width:${s.wPct.toFixed(3)}%;height:${s.hPct.toFixed(3)}%;">
        ${s.inner}
      </div>`).join('');

    const notesHtml = slide.notes.trim()
      ? `<div class="notes"><b>Notes:</b> ${escapeHtml(slide.notes).replace(/\n/g, '<br>')}</div>`
      : '';

    return `
  <div class="swrap${i === 0 ? ' active' : ''}" data-i="${i}">
    <div class="svp">
      <div class="sc" style="${bg}">
        ${shapesHtml}
      </div>
    </div>
    ${notesHtml}
  </div>`;
  }).join('\n');

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>${escapeHtml(title)}</title>
<style>
*{box-sizing:border-box;margin:0;padding:0}
html,body{height:100%;overflow:hidden;font-family:Arial,sans-serif;background:#323232}
#viewer{display:flex;flex-direction:column;height:100%}
#area{flex:1;display:flex;align-items:center;justify-content:center;min-height:0;padding:10px}
.swrap{display:none;flex-direction:column;align-items:center}
.swrap.active{display:flex}
.svp{overflow:hidden;position:relative;flex-shrink:0}
.sc{position:absolute;inset:0;width:${wpx}px;height:${hpx}px;transform-origin:top left;overflow:hidden;box-shadow:0 3px 16px rgba(0,0,0,.55)}
.sp{position:absolute;overflow:hidden}
.tb{width:100%;height:100%;overflow:hidden;padding:2px}
.tb p{word-break:break-word}
.notes{margin-top:6px;padding:5px 10px;font-size:12px;color:#ddd;background:rgba(255,255,180,.12);border-left:3px solid rgba(255,255,180,.4);max-height:60px;overflow-y:auto}
#nav{display:flex;align-items:center;justify-content:center;gap:14px;background:#111;padding:7px 12px;flex-shrink:0;color:#fff;font-size:14px}
#nav button{background:#555;color:#fff;border:none;padding:5px 14px;border-radius:3px;cursor:pointer}
#nav button:hover:not(:disabled){background:#888}
#nav button:disabled{opacity:.35;cursor:default}
#ctr{min-width:70px;text-align:center}
</style>
</head>
<body>
<div id="viewer">
  <div id="area">
${slidesHtml}
  </div>
  <div id="nav">
    <button id="prev" onclick="go(-1)" disabled>&#9664; Prev</button>
    <span id="ctr">1 / ${slides.length}</span>
    <button id="next" onclick="go(1)"${slides.length <= 1 ? ' disabled' : ''}>Next &#9654;</button>
  </div>
</div>
<script>
const W=${wpx},H=${hpx};
let idx=0;
const wraps=document.querySelectorAll('.swrap');
const vps=document.querySelectorAll('.svp');
const scs=document.querySelectorAll('.sc');

function scale(){
  const a=document.getElementById('area');
  const aw=a.clientWidth-20,ah=a.clientHeight-20;
  const s=Math.min(aw/W,ah/H);
  const sw=Math.round(W*s),sh=Math.round(H*s);
  vps.forEach(v=>{v.style.width=sw+'px';v.style.height=sh+'px';});
  scs.forEach(c=>{c.style.transform='scale('+s+')';});
}

function go(d){
  wraps[idx].classList.remove('active');
  idx=Math.max(0,Math.min(wraps.length-1,idx+d));
  wraps[idx].classList.add('active');
  document.getElementById('ctr').textContent=(idx+1)+' / '+wraps.length;
  document.getElementById('prev').disabled=idx===0;
  document.getElementById('next').disabled=idx===wraps.length-1;
  scale();
}

document.addEventListener('keydown',e=>{
  if(['ArrowRight','ArrowDown','PageDown'].includes(e.key))go(1);
  if(['ArrowLeft','ArrowUp','PageUp'].includes(e.key))go(-1);
});
window.addEventListener('resize',scale);
scale();
</script>
</body>
</html>`;
}

/**
 * Convert a PowerPoint file to a self-contained HTML presentation viewer.
 */
async function convertPptxToHtml(filePath: string, title: string): Promise<string> {
  const raw = fs.readFileSync(filePath);
  const zip = await JSZip.loadAsync(raw);

  // ── Slide dimensions from presentation.xml ────────────────────
  let slideW = 9144000; // 10 inches default
  let slideH = 6858000; // 7.5 inches default

  if (zip.files['ppt/presentation.xml']) {
    const presXml = await zip.files['ppt/presentation.xml'].async('string');
    const presDoc = await parseStringPromise(presXml, { explicitArray: true });
    const sldSz = get(presDoc, 'p:presentation', 'p:sldSz') as XmlNode | undefined;
    if (sldSz?.['$']) {
      slideW = parseInt(sldSz['$'].cx, 10) || slideW;
      slideH = parseInt(sldSz['$'].cy, 10) || slideH;
    }
  }

  // ── Pre-load all media as base64 data URIs ────────────────────
  const mediaCache: Record<string, string> = {};
  for (const zipPath of Object.keys(zip.files)) {
    if (!zipPath.startsWith('ppt/media/')) continue;
    const zipFile = zip.files[zipPath];
    if (zipFile.dir) continue;
    const ext = zipPath.split('.').pop() ?? '';
    const mime = getMimeType(ext);
    if (mime) {
      const b64 = await zipFile.async('base64');
      mediaCache[zipPath] = `data:${mime};base64,${b64}`;
    }
  }

  // ── Sort slides numerically ───────────────────────────────────
  const slideKeys = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const n = (k: string) => parseInt(k.match(/slide(\d+)/)![1], 10);
      return n(a) - n(b);
    });

  const slides: HtmlSlide[] = [];

  for (let i = 0; i < slideKeys.length; i++) {
    const slideKey = slideKeys[i];
    const slideNum = i + 1;

    // ── Load slide rels (image references) ────────────────────
    const rels: Record<string, string> = {};
    const relKey = `ppt/slides/_rels/slide${slideNum}.xml.rels`;
    if (zip.files[relKey]) {
      const relsXml = await zip.files[relKey].async('string');
      const relsDoc = await parseStringPromise(relsXml, { explicitArray: true });
      const relsRoot = (relsDoc as XmlNode)?.['Relationships']?.[0] as XmlNode | undefined;
      for (const rel of ([] as XmlNode[]).concat(relsRoot?.['Relationship'] ?? [])) {
        if (rel['$']?.Id && rel['$']?.Target) {
          rels[rel['$'].Id] = resolveZipPath(slideKey, rel['$'].Target as string);
        }
      }
    }

    // ── Parse slide ───────────────────────────────────────────
    const slideXml = await zip.files[slideKey].async('string');
    const slideDoc = await parseStringPromise(slideXml, { explicitArray: true });
    const slide = parseSlide(slideDoc, rels, mediaCache, slideW, slideH);

    // ── Speaker notes ─────────────────────────────────────────
    const notesKey = `ppt/notesSlides/notesSlide${slideNum}.xml`;
    if (zip.files[notesKey]) {
      const notesXml = await zip.files[notesKey].async('string');
      const notesDoc = await parseStringPromise(notesXml, { explicitArray: true });
      slide.notes = extractNotesText(notesDoc);
    }

    slides.push(slide);
  }

  return buildPresentationHtml(title, slides, slideW, slideH);
}

// ─── Sidebar generator ───────────────────────────────────────────────────────

interface SidebarItem {
  text: string;
  link?: string;
  collapsed?: boolean;
  items?: SidebarItem[];
}

function buildSidebar(outDir: string, baseUrl = '/'): string {
  function buildItems(dir: string, urlBase: string): SidebarItem[] {
    const entries = fs.readdirSync(dir, { withFileTypes: true })
      .filter(e =>
        e.name !== 'sidebar.ts' &&
        e.name !== 'index.md' &&
        e.name !== 'public'         // skip static-asset dir
      )
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
    if (path.basename(f).startsWith('~$')) return false;
    const ext = path.extname(f).toLowerCase();
    return ext === '.docx' || ext === '.pptx' || ext === '.xlsx';
  });

  if (!allFiles.length) {
    console.log('No .docx, .pptx, or .xlsx files found.');
    return;
  }

  const convertFiles = allFiles.filter(f => path.extname(f).toLowerCase() !== '.xlsx');
  const xlsxFiles    = allFiles.filter(f => path.extname(f).toLowerCase() === '.xlsx');

  // Copy Excel files as-is
  for (const filePath of xlsxFiles) {
    const relative = path.relative(INPUT_DIR, filePath);
    const outPath  = path.join(OUTPUT_DIR, relative);
    mkdirp(path.dirname(outPath));
    fs.copyFileSync(filePath, outPath);
    console.log(`  ${relative}  →  ${relative} (copied)`);
  }

  if (convertFiles.length) console.log();

  console.log(
    `Found ${convertFiles.length} file(s) to convert` +
    (xlsxFiles.length ? `, ${xlsxFiles.length} Excel file(s) copied` : '') +
    '.\n'
  );

  let ok = 0;
  let fail = 0;

  for (const filePath of convertFiles) {
    const ext      = path.extname(filePath).toLowerCase();
    const relative = path.relative(INPUT_DIR, filePath);
    const baseName = path.basename(filePath);
    const title    = path.basename(filePath, ext);

    if (ext === '.pptx') {
      // ── PPTX → HTML viewer + .md iframe wrapper ─────────────
      const relFwd     = relative.replace(/\\/g, '/');
      const htmlRelFwd = relFwd.replace(/\.pptx$/i, '.html');
      const mdRel      = relative.replace(/\.pptx$/i, '.md');

      const htmlOutPath = path.join(OUTPUT_DIR, 'public', 'slides',
        htmlRelFwd.replace(/\//g, path.sep));
      const mdOutPath = path.join(OUTPUT_DIR, mdRel);

      mkdirp(path.dirname(htmlOutPath));
      mkdirp(path.dirname(mdOutPath));

      process.stdout.write(`  ${relative}  →  ${mdRel} + public/slides/${htmlRelFwd} … `);

      try {
        const html = await convertPptxToHtml(filePath, title);
        fs.writeFileSync(htmlOutPath, html, 'utf8');

        // URL-encode each path segment for the iframe src
        const encodedUrl = '/slides/' +
          htmlRelFwd.split('/').map(seg => encodeURIComponent(seg)).join('/');

        const mdContent = [
          '---',
          `title: "${title.replace(/"/g, '\\"')}"`,
          `source: "${baseName}"`,
          '---',
          '',
          `<div style="width:100%;height:75vh;">`,
          `  <iframe src="${encodedUrl}" style="width:100%;height:100%;border:none;" title="${escapeHtml(title)}"></iframe>`,
          `</div>`,
          '',
        ].join('\n');

        fs.writeFileSync(mdOutPath, mdContent, 'utf8');
        console.log('OK');
        ok++;
      } catch (err) {
        console.log(`FAILED\n    ${(err as Error).message}`);
        fail++;
      }

    } else {
      // ── DOCX → Markdown ──────────────────────────────────────
      const outRel  = relative.replace(/\.docx$/i, '.md');
      const outPath = path.join(OUTPUT_DIR, outRel);

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
