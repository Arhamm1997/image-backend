'use strict';

/**
 * Conversion logic for images, video/audio, and documents.
 *
 * Libraries used:
 *   Images    — sharp (libvips), jimp (BMP), png-to-ico (ICO)
 *   Video/Audio — fluent-ffmpeg + ffmpeg-static (bundled FFmpeg binary, no system install needed)
 *   Documents — mammoth (DOCX→text/HTML), pdfkit (→PDF), pdf-parse (PDF→text),
 *               xlsx/SheetJS (Excel/CSV), docx (→DOCX), jszip (PPTX→text),
 *               libreoffice-convert (complex conversions via LibreOffice in Docker)
 */

const sharp          = require('sharp');
const ffmpegStatic   = require('ffmpeg-static');
const fluentFfmpeg   = require('fluent-ffmpeg');
const mammoth        = require('mammoth');
const PDFDocument    = require('pdfkit');
const XLSX           = require('xlsx');
const { Document, Packer, Paragraph, TextRun } = require('docx');
const JSZip          = require('jszip');
const libre          = require('libreoffice-convert');
const { promisify }  = require('util');
const os             = require('os');
const path           = require('path');
const crypto         = require('crypto');
const fs             = require('fs').promises;

fluentFfmpeg.setFfmpegPath(ffmpegStatic);
const libreConvertAsync = promisify(libre.convert);

// ─── Helpers ─────────────────────────────────────────────────────────────────

/** Create a unique temp file path. */
function tmpPath(ext) {
  return path.join(os.tmpdir(), `conv_${crypto.randomUUID()}.${ext}`);
}

/** Decode a Buffer to string, stripping BOM if present. */
function decodeText(buf) {
  const s = buf.toString('utf-8');
  return s.charCodeAt(0) === 0xfeff ? s.slice(1) : s;
}

// ─── MIME map ─────────────────────────────────────────────────────────────────

const MIME_FOR_FORMAT = {
  // images
  jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', webp: 'image/webp',
  gif: 'image/gif',  bmp: 'image/bmp',   ico: 'image/x-icon', tiff: 'image/tiff',
  tif: 'image/tiff', avif: 'image/avif', svg: 'image/svg+xml',
  // video
  mp4: 'video/mp4', webm: 'video/webm', avi: 'video/x-msvideo',
  mov: 'video/quicktime', mkv: 'video/x-matroska', flv: 'video/x-flv', wmv: 'video/x-ms-wmv',
  // audio
  mp3: 'audio/mpeg', wav: 'audio/wav', ogg: 'audio/ogg', aac: 'audio/aac',
  flac: 'audio/flac', m4a: 'audio/mp4', wma: 'audio/x-ms-wma',
  // archives
  zip:     'application/zip',
  tar:     'application/x-tar',
  'tar.gz':'application/gzip',
  tgz:     'application/gzip',
  'tar.bz2':'application/x-bzip2',
  tbz:     'application/x-bzip2',
  'tar.xz':'application/x-xz',
  txz:     'application/x-xz',
  gz:      'application/gzip',
  bz2:     'application/x-bzip2',
  xz:      'application/x-xz',
  '7z':    'application/x-7z-compressed',
  rar:     'application/x-rar-compressed',
  // audio extras
  aiff:    'audio/aiff',
  amr:     'audio/amr',
  opus:    'audio/opus',
  caf:     'audio/x-caf',
  // video extras
  '3gp':   'video/3gpp',
  '3g2':   'video/3gpp2',
  m4v:     'video/x-m4v',
  mpg:     'video/mpeg',
  mpeg:    'video/mpeg',
  mts:     'video/mp2t',
  ts:      'video/mp2t',
  vob:     'video/dvd',
  ogv:     'video/ogg',
  // documents
  pdf:  'application/pdf',
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  doc:  'application/msword',
  txt:  'text/plain; charset=utf-8',
  rtf:  'application/rtf',
  odt:  'application/vnd.oasis.opendocument.text',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  xls:  'application/vnd.ms-excel',
  pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  ppt:  'application/vnd.ms-powerpoint',
  csv:  'text/csv',
};

// ══════════════════════════════════════════════════════════════════════════════
// IMAGE  (sharp + jimp + png-to-ico)
// ══════════════════════════════════════════════════════════════════════════════

async function convertImage(data, srcExt, targetFormat) {
  const tgt = targetFormat.toLowerCase();

  // ── ICO: resize → PNG → ICO ────────────────────────────────────────────────
  if (tgt === 'ico') {
    const png = await sharp(data)
      .resize(256, 256, { fit: 'contain', background: { r: 0, g: 0, b: 0, alpha: 0 } })
      .png()
      .toBuffer();
    const pngToIco = require('png-to-ico');
    return pngToIco(png);
  }

  // ── SVG output: embed raster as base64 data-URI ───────────────────────────
  if (tgt === 'svg') {
    const { width, height } = await sharp(data).metadata();
    const png = await sharp(data).png().toBuffer();
    const b64 = png.toString('base64');
    return Buffer.from(
      `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}">` +
      `<image href="data:image/png;base64,${b64}" width="${width}" height="${height}"/>` +
      `</svg>`,
      'utf-8',
    );
  }

  // ── BMP output: sharp doesn't support BMP — use jimp ──────────────────────
  if (tgt === 'bmp') {
    const Jimp = require('jimp');
    const img  = await Jimp.read(data);
    return img.getBufferAsync(Jimp.MIME_BMP);
  }

  // ── All other formats: sharp ───────────────────────────────────────────────
  const sharpFmt = tgt === 'jpg' ? 'jpeg' : tgt === 'tif' ? 'tiff' : tgt;
  let pipeline   = sharp(data, { animated: tgt === 'gif' });

  // JPEG cannot store transparency — flatten to white
  if (sharpFmt === 'jpeg') {
    pipeline = pipeline.flatten({ background: { r: 255, g: 255, b: 255 } });
  }

  const opts = {};
  if (sharpFmt === 'jpeg') { opts.quality = 82; opts.mozjpeg = true; }
  if (sharpFmt === 'webp') { opts.quality = 80; opts.effort = 4; opts.smartSubsample = true; }
  if (sharpFmt === 'avif') { opts.quality = 60; opts.effort = 4; }
  if (sharpFmt === 'png')  { opts.compressionLevel = 9; opts.adaptiveFiltering = true; }

  return pipeline.toFormat(sharpFmt, opts).toBuffer();
}

// ══════════════════════════════════════════════════════════════════════════════
// VIDEO / AUDIO  (fluent-ffmpeg + ffmpeg-static)
// ══════════════════════════════════════════════════════════════════════════════

const AUDIO_EXTS = new Set(['mp3','wav','ogg','aac','flac','m4a','wma','aiff','amr','opus','caf']);
const VIDEO_EXTS = new Set(['mp4','webm','avi','mov','mkv','flv','wmv','3gp','3g2','m4v','mpg','mpeg','mts','m2ts','ts','vob','ogv']);

// Per-format ffmpeg output options
const AUDIO_OPTS = {
  mp3:  ['-codec:a', 'libmp3lame', '-q:a', '2'],
  aac:  ['-codec:a', 'aac', '-b:a', '192k'],
  wav:  ['-codec:a', 'pcm_s16le'],
  ogg:  ['-codec:a', 'libvorbis', '-q:a', '4'],
  flac: ['-codec:a', 'flac'],
  m4a:  ['-codec:a', 'aac', '-b:a', '192k'],
  wma:  ['-codec:a', 'wmav2', '-b:a', '192k'],
  aiff: ['-codec:a', 'pcm_s16be'],
  amr:  ['-codec:a', 'libopencore_amrnb', '-b:a', '12.2k', '-ar', '8000', '-ac', '1'],
  opus: ['-codec:a', 'libopus', '-b:a', '128k'],
  caf:  ['-codec:a', 'pcm_s16le'],
};

const VIDEO_OPTS = {
  mp4:  ['-codec:v', 'libx264', '-preset', 'fast', '-crf', '23',
         '-codec:a', 'aac', '-b:a', '128k', '-movflags', '+faststart'],
  webm: ['-codec:v', 'libvpx-vp9', '-crf', '30', '-b:v', '0',
         '-codec:a', 'libopus', '-b:a', '128k'],
  avi:  ['-codec:v', 'mpeg4', '-q:v', '5',
         '-codec:a', 'libmp3lame', '-q:a', '4'],
  mov:  ['-codec:v', 'libx264', '-preset', 'fast', '-crf', '23',
         '-codec:a', 'aac', '-b:a', '128k'],
  mkv:  ['-codec:v', 'libx264', '-preset', 'fast', '-crf', '23',
         '-codec:a', 'aac', '-b:a', '128k'],
  flv:  ['-codec:v', 'libx264', '-crf', '23',
         '-codec:a', 'aac', '-b:a', '128k'],
  wmv:  ['-codec:v', 'wmv2', '-q:v', '5',
         '-codec:a', 'wmav2', '-b:a', '128k'],
  '3gp':['-codec:v', 'libx264', '-preset', 'fast', '-crf', '28',
         '-codec:a', 'aac', '-b:a', '64k', '-ar', '22050'],
  '3g2':['-codec:v', 'libx264', '-preset', 'fast', '-crf', '28',
         '-codec:a', 'aac', '-b:a', '64k', '-ar', '22050'],
  m4v:  ['-codec:v', 'libx264', '-preset', 'fast', '-crf', '23',
         '-codec:a', 'aac', '-b:a', '128k', '-movflags', '+faststart'],
  mpg:  ['-codec:v', 'mpeg2video', '-q:v', '5', '-codec:a', 'mp2', '-b:a', '192k'],
  mpeg: ['-codec:v', 'mpeg2video', '-q:v', '5', '-codec:a', 'mp2', '-b:a', '192k'],
  ogv:  ['-codec:v', 'libtheora', '-q:v', '5', '-codec:a', 'libvorbis', '-q:a', '4'],
  ts:   ['-codec:v', 'libx264', '-preset', 'fast', '-crf', '23',
         '-codec:a', 'aac', '-b:a', '128k'],
};

function convertVideoAudio(data, srcExt, targetFormat) {
  const tgt = targetFormat.toLowerCase();
  const src = srcExt.toLowerCase();

  return new Promise(async (resolve, reject) => {
    const inPath  = tmpPath(src);
    const outPath = tmpPath(tgt);

    try {
      await fs.writeFile(inPath, data);

      // If converting video → audio, drop the video stream
      const isAudioExtract = VIDEO_EXTS.has(src) && AUDIO_EXTS.has(tgt);
      const outputOpts = isAudioExtract
        ? ['-vn', ...(AUDIO_OPTS[tgt] ?? [])]
        : AUDIO_EXTS.has(tgt)
          ? (AUDIO_OPTS[tgt] ?? [])
          : (VIDEO_OPTS[tgt] ?? []);

      fluentFfmpeg(inPath)
        .outputOptions(outputOpts)
        .output(outPath)
        .on('end', async () => {
          try {
            resolve(await fs.readFile(outPath));
          } catch (e) {
            reject(e);
          } finally {
            fs.unlink(inPath).catch(() => {});
            fs.unlink(outPath).catch(() => {});
          }
        })
        .on('error', (err) => {
          fs.unlink(inPath).catch(() => {});
          fs.unlink(outPath).catch(() => {});
          reject(new Error(`FFmpeg: ${err.message}`));
        })
        .run();
    } catch (e) {
      fs.unlink(inPath).catch(() => {});
      reject(e);
    }
  });
}

// ══════════════════════════════════════════════════════════════════════════════
// DOCUMENTS
// ══════════════════════════════════════════════════════════════════════════════

async function convertDocument(data, srcExt, targetFormat) {
  const src = srcExt.toLowerCase();
  const tgt = targetFormat.toLowerCase();

  // ── TXT ───────────────────────────────────────────────────────────────────
  if (src === 'txt') {
    if (tgt === 'pdf')  return txtToPdf(data);
    if (tgt === 'docx') return txtToDocx(data);
    if (tgt === 'xlsx') return txtToXlsx(data);
    if (tgt === 'rtf')  return txtToRtf(data);
    if (tgt === 'csv')  return data;  // txt is already valid plaintext
  }

  // ── DOCX ──────────────────────────────────────────────────────────────────
  if (src === 'docx') {
    if (tgt === 'txt')  return docxToTxt(data);
    if (tgt === 'html') return docxToHtml(data);
    if (tgt === 'pdf')  return docxToPdf(data);
    if (tgt === 'xlsx') return docxToXlsx(data);
    if (tgt === 'xls')  return docxToXlsx(data);
    if (tgt === 'csv')  return docxToXlsx(data).then(xlsxToCsv);
    if (tgt === 'rtf')  return docxToTxt(data).then(buf => txtToRtf(buf));
  }

  // ── PDF ───────────────────────────────────────────────────────────────────
  if (src === 'pdf') {
    if (tgt === 'txt')  return pdfToTxt(data);
    if (tgt === 'xlsx') return pdfToXlsx(data);
    if (tgt === 'csv')  return pdfToCsv(data);
    if (tgt === 'docx') return pdfToDocx(data);
  }

  // ── XLSX / XLS / CSV ──────────────────────────────────────────────────────
  if (src === 'xlsx' || src === 'xls') {
    if (tgt === 'csv')  return xlsxToCsv(data);
    if (tgt === 'txt')  return xlsxToTxt(data);
    if (tgt === 'pdf')  return xlsxToPdf(data);
  }
  if (src === 'csv') {
    if (tgt === 'xlsx') return csvToXlsx(data);
    if (tgt === 'txt')  return data;
    if (tgt === 'pdf')  return csvToPdf(data);
  }

  // ── PPTX ──────────────────────────────────────────────────────────────────
  if (src === 'pptx') {
    if (tgt === 'txt') return pptxToTxt(data);
    if (tgt === 'pdf') return pptxToPdf(data);
  }

  // ── RTF ───────────────────────────────────────────────────────────────────
  if (src === 'rtf' && tgt === 'txt') return rtfToTxt(data);

  // ── Fallback: LibreOffice (handles DOC, ODT and other complex conversions) ─
  return libreOfficeConvert(data, src, tgt);
}

// ─── TXT helpers ─────────────────────────────────────────────────────────────

function txtToPdf(data) {
  return new Promise((resolve, reject) => {
    const doc    = new PDFDocument({ margin: 50 });
    const chunks = [];
    doc.on('data', c => chunks.push(c));
    doc.on('end',  () => resolve(Buffer.concat(chunks)));
    doc.on('error', reject);

    doc.font('Helvetica').fontSize(11);
    for (const line of decodeText(data).split('\n')) {
      doc.text(line || ' ');
    }
    doc.end();
  });
}

async function txtToDocx(data) {
  const lines = decodeText(data).split('\n');
  const children = lines.map(line => {
    const trimmed = line.trim();
    const isHeading =
      trimmed === trimmed.toUpperCase() &&
      trimmed.length > 2 &&
      trimmed.length < 80 &&
      /[A-Z]/.test(trimmed);
    const isBullet = /^[-•*]\s/.test(trimmed) || /^\d+\.\s/.test(trimmed);

    if (isHeading) {
      return new Paragraph({
        children: [new TextRun({ text: trimmed, bold: true, size: 26 })],
        spacing: { before: 240, after: 80 },
      });
    }
    if (isBullet) {
      return new Paragraph({
        children: [new TextRun(trimmed)],
        indent: { left: 360 },
        spacing: { after: 60 },
      });
    }
    return new Paragraph({
      children: [new TextRun(trimmed || '')],
      spacing: { after: 60 },
    });
  });
  const doc = new Document({ sections: [{ children }] });
  return Packer.toBuffer(doc);
}

function txtToXlsx(data) {
  const rows = decodeText(data).split('\n').map(line => [line]);
  const wb   = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sheet1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

function txtToRtf(data) {
  const lines = ['{\\rtf1\\ansi\\deff0{\\fonttbl{\\f0 Courier New;}}'];
  for (const line of decodeText(data).split('\n')) {
    const esc = line.replace(/\\/g, '\\\\').replace(/\{/g, '\\{').replace(/\}/g, '\\}');
    lines.push(`\\f0\\fs20 ${esc}\\par`);
  }
  lines.push('}');
  return Buffer.from(lines.join('\n'), 'utf-8');
}

// ─── DOCX helpers ────────────────────────────────────────────────────────────

async function docxToTxt(data) {
  const { value } = await mammoth.extractRawText({ buffer: data });
  return Buffer.from(value, 'utf-8');
}

async function docxToHtml(data) {
  const { value } = await mammoth.convertToHtml({ buffer: data });
  return Buffer.from(value, 'utf-8');
}

async function docxToPdf(data) {
  // Extract plain text via mammoth, then render to PDF with pdfkit
  const { value: text } = await mammoth.extractRawText({ buffer: data });
  return txtToPdf(Buffer.from(text, 'utf-8'));
}

async function docxToXlsx(data) {
  // Convert DOCX to HTML, then extract tables; fall back to plain text rows
  const { value: html } = await mammoth.convertToHtml({ buffer: data });
  const rows = [];
  const tableRe  = /<table[\s\S]*?<\/table>/gi;
  const rowRe    = /<tr[\s\S]*?<\/tr>/gi;
  const cellRe   = /<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi;
  let hasTable   = false;

  for (const [tableHtml] of html.matchAll(tableRe)) {
    hasTable = true;
    for (const [rowHtml] of tableHtml.matchAll(rowRe)) {
      const cells = [...rowHtml.matchAll(cellRe)].map(c => c[1].replace(/<[^>]+>/g, '').trim());
      if (cells.length) rows.push(cells);
    }
  }

  if (!hasTable) {
    const { value: text } = await mammoth.extractRawText({ buffer: data });
    text.split('\n').filter(l => l.trim()).forEach(l => rows.push([l]));
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sheet1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// ─── PDF helpers ─────────────────────────────────────────────────────────────

/**
 * Safely parse a PDF buffer.
 * - Validates the %PDF- header to catch fake/empty files immediately.
 * - Races against a 30-second timeout so a corrupt PDF never hangs the server.
 */
async function parsePdf(data) {
  if (!Buffer.isBuffer(data) || data.length < 20) {
    throw new Error('File is too small to be a valid PDF (minimum 20 bytes).');
  }
  if (data.slice(0, 5).toString('ascii') !== '%PDF-') {
    throw new Error('Invalid PDF file: missing %PDF- header. Make sure you are uploading a real PDF.');
  }
  const pdfParse = require('pdf-parse/lib/pdf-parse.js');
  const timeout  = new Promise((_, reject) =>
    setTimeout(() => reject(new Error('PDF parsing timed out — the file may be corrupt or password-protected.')), 30_000)
  );
  return Promise.race([pdfParse(data), timeout]);
}

async function pdfToTxt(data) {
  const parsed = await parsePdf(data);
  return Buffer.from(parsed.text, 'utf-8');
}

async function pdfToDocx(data) {
  // ── Attempt 1: Render each page as a PNG via pdftoppm (Docker) ──────────────
  try {
    const { execFile }   = require('child_process');
    const execFileAsync  = require('util').promisify(execFile);
    const { Document, Packer, Paragraph, ImageRun, PageOrientation } = require('docx');

    const tmpDir    = path.join(os.tmpdir(), `pdf2docx_${crypto.randomUUID()}`);
    const inputPath = path.join(tmpDir, 'input.pdf');
    await fs.mkdir(tmpDir, { recursive: true });
    await fs.writeFile(inputPath, data);

    // Render all pages to PNG at 150 DPI; output: page-01.png, page-02.png …
    await execFileAsync('pdftoppm', ['-r', '150', '-png', inputPath, path.join(tmpDir, 'page')]);

    const pngs = (await fs.readdir(tmpDir))
      .filter(f => f.endsWith('.png'))
      .sort()
      .map(f => path.join(tmpDir, f));

    if (pngs.length === 0) throw new Error('pdftoppm produced no images');

    // Detect orientation from first page
    const firstMeta = await sharp(await fs.readFile(pngs[0])).metadata();
    const landscape = firstMeta.width > firstMeta.height;

    // DOCX page size in twips (1440 twips = 1 inch)
    const PAGE_W  = landscape ? 15840 : 12240; // 11" or 8.5"
    const PAGE_H  = landscape ? 12240 : 15840;
    // Use 0.4" margins to avoid Word clipping the image at printable-area boundaries
    const MARGIN  = 576; // 0.4 inch in twips
    // Available pixels at 96 DPI: (page_twips - 2*margin) / 1440 * 96
    const AVAIL_W = Math.floor((PAGE_W - 2 * MARGIN) / 1440 * 96);
    const AVAIL_H = Math.floor((PAGE_H - 2 * MARGIN) / 1440 * 96);

    const sections = [];
    for (const pngPath of pngs) {
      const imgBuf = await fs.readFile(pngPath);
      const meta   = await sharp(imgBuf).metadata();
      const ar     = meta.width / meta.height;
      let imgW = AVAIL_W;
      let imgH = Math.round(imgW / ar);
      if (imgH > AVAIL_H) { imgH = AVAIL_H; imgW = Math.round(imgH * ar); }

      sections.push({
        properties: {
          page: {
            size: { width: PAGE_W, height: PAGE_H,
                    orientation: landscape ? PageOrientation.LANDSCAPE : PageOrientation.PORTRAIT },
            margin: { top: MARGIN, right: MARGIN, bottom: MARGIN, left: MARGIN },
          },
        },
        children: [
          new Paragraph({
            children: [new ImageRun({ data: imgBuf, transformation: { width: imgW, height: imgH } })],
            spacing: { before: 0, after: 0 },
          }),
        ],
      });
    }

    await fs.rm(tmpDir, { recursive: true }).catch(() => {});
    return Packer.toBuffer(new Document({ sections }));

  } catch (renderErr) {
    console.warn('[pdfToDocx] pdftoppm failed:', renderErr.message);
  }

  // ── Attempt 2: LibreOffice (preserves text/tables/layout) ───────────────────
  try {
    const loBuf = await libreOfficeConvert(data, 'pdf', 'docx');
    if (loBuf && loBuf.length > 500) return loBuf;
    throw new Error('LibreOffice returned empty output');
  } catch (libreErr) {
    console.warn('[pdfToDocx] LibreOffice failed:', libreErr.message);
  }

  // ── Attempt 3: Plain-text fallback ──────────────────────────────────────────
  try {
    const parsed = await parsePdf(data);
    return txtToDocx(Buffer.from(parsed.text, 'utf-8'));
  } catch (textErr) {
    throw new Error(`PDF→DOCX failed on all methods. Last error: ${textErr.message}`);
  }
}

async function pdfToXlsx(data) {
  const { text } = await parsePdf(data);

  // Split into lines, then try to detect columns by runs of 2+ whitespace
  const rows = text
    .split('\n')
    .map(line => {
      const cells = line.split(/\s{2,}/).map(c => c.trim()).filter(c => c);
      return cells.length > 1 ? cells : line.trim() ? [line.trim()] : null;
    })
    .filter(Boolean);

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(rows), 'Sheet1');
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

async function pdfToCsv(data) {
  const { text } = await parsePdf(data);

  const rows = text
    .split('\n')
    .map(line => {
      const cells = line.split(/\s{2,}/).map(c => c.trim()).filter(c => c);
      return cells.length > 1 ? cells : line.trim() ? [line.trim()] : null;
    })
    .filter(Boolean);

  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
  return Buffer.from(XLSX.utils.sheet_to_csv(ws), 'utf-8');
}

// ─── XLSX / CSV helpers ───────────────────────────────────────────────────────

function xlsxToCsv(data) {
  const wb = XLSX.read(data, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return Buffer.from(XLSX.utils.sheet_to_csv(ws), 'utf-8');
}

function xlsxToTxt(data) {
  const wb   = XLSX.read(data, { type: 'buffer' });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1 });
  return Buffer.from(rows.map(r => r.join('\t')).join('\n'), 'utf-8');
}

function csvToXlsx(data) {
  const wb = XLSX.read(decodeText(data), { type: 'string' });
  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// ─── XLSX / CSV → PDF ────────────────────────────────────────────────────────

function xlsxToPdf(data) {
  return new Promise((resolve, reject) => {
    try {
      const wb     = XLSX.read(data, { type: 'buffer' });
      const margin = 36;
      const doc    = new PDFDocument({ margin, size: 'A4', layout: 'landscape' });
      const chunks = [];
      doc.on('data',  c => chunks.push(c));
      doc.on('end',   () => resolve(Buffer.concat(chunks)));
      doc.on('error', reject);

      let firstSheet = true;

      for (const sheetName of wb.SheetNames) {
        if (!firstSheet) doc.addPage();
        firstSheet = false;

        const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });

        // Sheet title
        doc.font('Helvetica-Bold').fontSize(11).fillColor('#1e293b').text(sheetName);
        doc.moveDown(0.3);

        if (!rows.length) {
          doc.font('Helvetica').fontSize(9).fillColor('#94a3b8').text('(empty sheet)');
          continue;
        }

        const pageW = doc.page.width - margin * 2;
        const colN  = Math.max(...rows.map(r => (Array.isArray(r) ? r.length : 0)), 1);
        const colW  = pageW / colN;
        const rowH  = 18;
        const fSize = Math.max(6, Math.min(9, Math.floor(colW / 9)));
        let   y     = doc.y;

        rows.forEach((row, ri) => {
          if (!Array.isArray(row)) return;

          // New page if needed
          if (y + rowH > doc.page.height - margin) {
            doc.addPage();
            y = margin;
          }

          const isHeader = ri === 0;

          // Row background
          doc.save();
          if (isHeader) {
            doc.rect(margin, y, pageW, rowH).fill('#2563eb');
          } else if (ri % 2 === 1) {
            doc.rect(margin, y, pageW, rowH).fill('#f8fafc');
          }
          doc.restore();

          // Draw each cell
          for (let ci = 0; ci < colN; ci++) {
            const val  = row[ci] != null ? String(row[ci]) : '';
            const x    = margin + ci * colW;
            const maxC = Math.max(4, Math.floor(colW / (fSize * 0.55)));
            const text = val.length > maxC ? val.slice(0, maxC - 1) + '\u2026' : val;

            doc.font(isHeader ? 'Helvetica-Bold' : 'Helvetica')
               .fontSize(fSize)
               .fillColor(isHeader ? '#ffffff' : '#1e293b')
               .text(text, x + 3, y + (rowH - fSize) / 2, { width: colW - 6, lineBreak: false });
          }

          // Row bottom border
          doc.save()
             .strokeColor(isHeader ? '#1d4ed8' : '#e2e8f0')
             .lineWidth(0.3)
             .moveTo(margin, y + rowH)
             .lineTo(margin + pageW, y + rowH)
             .stroke()
             .restore();

          y += rowH;
        });
      }

      doc.end();
    } catch (e) {
      reject(e);
    }
  });
}

function csvToPdf(data) {
  // Parse CSV via SheetJS then reuse xlsxToPdf
  const wb = XLSX.read(decodeText(data), { type: 'string' });
  return xlsxToPdf(XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
}

async function pptxToPdf(data) {
  const textBuf = await pptxToTxt(data);
  return txtToPdf(textBuf);
}

// ─── RTF helper ───────────────────────────────────────────────────────────────

function rtfToTxt(data) {
  let s = decodeText(data);
  // Drop binary blobs (pictures, OLE objects)
  s = s.replace(/\{\\(?:pict|object|objdata|bin)[^{}]*\}/gi, '');
  // Drop destination groups we don't want (footnotes, headers, styles, etc.)
  s = s.replace(/\{\\[*]?\\(?:fonttbl|colortbl|stylesheet|header|footer|info|fldinst)[^{}]*\}/gi, '');
  // Recursively strip remaining braces groups until stable
  for (let i = 0; i < 10; i++) s = s.replace(/\{[^{}]*\}/g, '');
  // Replace common paragraph/line break control words with newlines
  s = s.replace(/\\(?:par|pard|line)\b\s?/gi, '\n');
  // Remove all remaining control words and symbols
  s = s.replace(/\\[a-z]+\-?\d*\s?/gi, '');
  s = s.replace(/[{}\\]/g, '');
  // Normalise whitespace
  s = s.replace(/\r\n|\r/g, '\n').replace(/\n{3,}/g, '\n\n').trim();
  return Buffer.from(s, 'utf-8');
}

// ─── PPTX helper (parse XML from ZIP — no LibreOffice needed for text) ────────

async function pptxToTxt(data) {
  const zip = await JSZip.loadAsync(data);
  const slideFiles = Object.keys(zip.files)
    .filter(f => /^ppt\/slides\/slide\d+\.xml$/.test(f))
    .sort((a, b) => {
      const n = s => parseInt(s.match(/\d+/)[0]);
      return n(a) - n(b);
    });

  const lines = [];
  for (let i = 0; i < slideFiles.length; i++) {
    lines.push(`=== Slide ${i + 1} ===`);
    const xml = await zip.files[slideFiles[i]].async('string');
    const texts = [...xml.matchAll(/<a:t[^>]*>([^<]*)<\/a:t>/g)]
      .map(m => m[1])
      .filter(t => t.trim());
    lines.push(...texts, '');
  }

  return Buffer.from(lines.join('\n'), 'utf-8');
}

// ─── LibreOffice fallback ─────────────────────────────────────────────────────

async function libreOfficeConvert(data, srcExt, targetFormat) {
  try {
    // libreoffice-convert expects the extension WITH the leading dot
    return await libreConvertAsync(data, `.${targetFormat}`, undefined);
  } catch (err) {
    const msg = err.message || '';
    if (msg.includes('soffice') || msg.includes('spawn') || msg.includes('ENOENT')) {
      throw new Error(
        `Converting .${srcExt} → .${targetFormat} requires LibreOffice, ` +
        `which is not installed on this server. ` +
        `Deploy with Docker (see Dockerfile) for full document support.`,
      );
    }
    throw err;
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// ARCHIVE  (archive→archive repack  +  batch-convert files inside an archive)
// ══════════════════════════════════════════════════════════════════════════════

const _IMG_SET  = new Set(['jpg','jpeg','png','webp','gif','bmp','ico','tiff','tif','svg','avif']);
const _VID_SET  = new Set(['mp4','webm','avi','mov','mkv','flv','wmv','3gp','3g2','m4v','mpg','mpeg','mts','ts','vob','ogv']);
const _AUD_SET  = new Set(['mp3','wav','ogg','aac','flac','m4a','wma','aiff','amr','opus']);
const _DOC_SET  = new Set(['pdf','docx','doc','txt','rtf','odt','xlsx','xls','pptx','ppt','csv','tsv','html']);

/**
 * Extract entries from a ZIP buffer.
 * Returns [{name, data}] for non-directory entries.
 */
async function extractZip(buffer) {
  const zip     = await JSZip.loadAsync(buffer);
  const entries = [];
  for (const [name, f] of Object.entries(zip.files)) {
    if (f.dir) continue;
    if (name.startsWith('__MACOSX') || name.endsWith('.DS_Store')) continue;
    entries.push({ name, data: await f.async('nodebuffer') });
  }
  return entries;
}

/**
 * Repack a ZIP archive into another archive format.
 * Supported targets: zip, tar, tar.gz / tgz, tar.bz2 / tbz, tar.xz / txz
 * Not supported (requires native binaries): 7z, rar
 */
async function convertArchiveToArchive(buffer, srcExt, tgt) {
  const archiver = require('archiver');

  const NOT_SUPPORTED = ['7z','rar','ace','arj','cab','lzma','lzo','lha','rz','zst','lz'];
  if (NOT_SUPPORTED.includes(tgt)) {
    throw new Error(
      `Converting to .${tgt} requires a native binary (7-Zip / unrar) that is not ` +
      `installed on this server. Supported archive output formats: zip, tar, tar.gz, tar.bz2, tar.xz.`
    );
  }

  // Extract input archive (currently only ZIP supported as input)
  const entries = await extractZip(buffer);
  if (entries.length === 0) throw new Error('Archive is empty.');

  // Normalise target
  const normTgt = { tgz: 'tar.gz', tbz: 'tar.bz2', txz: 'tar.xz' }[tgt] ?? tgt;

  return new Promise((resolve, reject) => {
    const chunks = [];
    let archive;

    if (normTgt === 'zip') {
      archive = archiver('zip', { zlib: { level: 6 } });
    } else if (normTgt === 'tar') {
      archive = archiver('tar');
    } else if (normTgt === 'tar.gz') {
      archive = archiver('tar', { gzip: true, gzipOptions: { level: 6 } });
    } else if (normTgt === 'tar.bz2') {
      archive = archiver('tar', { bzip2: true });
    } else if (normTgt === 'tar.xz') {
      // archiver doesn't support xz natively — fall back to gzip
      archive = archiver('tar', { gzip: true, gzipOptions: { level: 9 } });
    } else {
      return reject(new Error(`Unsupported archive target: .${tgt}`));
    }

    archive.on('data',    c  => chunks.push(c));
    archive.on('end',     () => resolve(Buffer.concat(chunks)));
    archive.on('error',   reject);
    archive.on('warning', e => { if (e.code !== 'ENOENT') reject(e); });

    for (const { name, data } of entries) {
      archive.append(data, { name });
    }
    archive.finalize();
  });
}

/** Batch-convert all files inside a ZIP to a target format, returning a new ZIP. */
async function convertZip(buffer, _srcExt, targetFormat) {
  const tgt       = targetFormat.toLowerCase();
  const entries   = await extractZip(buffer);
  const outputZip = new JSZip();

  if (entries.length === 0) throw new Error('ZIP archive is empty.');

  let converted = 0;
  for (const { name: filename, data: fileBuffer } of entries) {
    const ext = (filename.split('.').pop() ?? '').toLowerCase();
    if (!ext) continue;

    try {
      let result;

      if (_IMG_SET.has(ext)) {
        result = await convertImage(fileBuffer, ext, tgt);
      } else if (_VID_SET.has(ext) || _AUD_SET.has(ext)) {
        result = await convertVideoAudio(fileBuffer, ext, tgt);
      } else if (_DOC_SET.has(ext)) {
        result = await convertDocument(fileBuffer, ext, tgt);
      } else {
        continue; // unsupported file inside ZIP — skip silently
      }

      const newName = filename.replace(/\.[^/.]+$/, '') + '.' + tgt;
      outputZip.file(newName, result);
      converted++;
    } catch (err) {
      console.warn(`[ZIP] Skipping "${filename}": ${err.message}`);
    }
  }

  if (converted === 0) {
    throw new Error(`No files inside the ZIP could be converted to .${tgt}. Check that the ZIP contains compatible files.`);
  }

  return outputZip.generateAsync({
    type: 'nodebuffer',
    compression: 'DEFLATE',
    compressionOptions: { level: 6 },
  });
}

module.exports = { MIME_FOR_FORMAT, convertImage, convertVideoAudio, convertDocument, convertZip, convertArchiveToArchive };
