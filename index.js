'use strict';

/**
 * File Converter API
 *
 * Development:  node --watch index.js
 * Production:   node index.js
 * Docker:       see Dockerfile
 */

const express  = require('express');
const cors     = require('cors');
const multer   = require('multer');
const {
  MIME_FOR_FORMAT,
  convertImage, convertVideoAudio, convertDocument,
  convertZip, convertArchiveToArchive,
} = require('./converters');

const app = express();

// ─── File-type sets ───────────────────────────────────────────────────────────
const IMAGE_EXTS   = new Set(['jpg','jpeg','png','webp','gif','bmp','ico','tiff','tif','svg','avif','heic','heif']);
const VIDEO_EXTS   = new Set(['mp4','webm','avi','mov','mkv','flv','wmv','3gp','3g2','m4v','mpg','mpeg','mts','m2ts','ts','vob','ogv']);
const AUDIO_EXTS   = new Set(['mp3','wav','ogg','aac','flac','m4a','wma','aiff','amr','opus','caf']);
const DOC_EXTS     = new Set(['pdf','docx','doc','txt','rtf','odt','xlsx','xls','pptx','ppt','csv','tsv','html','htm']);
const ARCHIVE_EXTS = new Set(['zip','tar','gz','bz2','xz','7z','rar','tgz','tbz','txz','lz','lzma','lzo','zst']);

// Compound extensions that need special detection (must check before single-ext)
const COMPOUND_EXTS = ['tar.gz','tar.bz2','tar.xz','tar.lz','tar.lzma','tar.lzo','tar.z','tar.zst'];

// ─── CORS ─────────────────────────────────────────────────────────────────────
const ALLOWED_ORIGINS = [
  'http://localhost:5173',
  'http://localhost:3000',
  'http://127.0.0.1:5173',
  'https://image-frontend-black.vercel.app',
  process.env.FRONTEND_URL,
].filter(Boolean).map(o => o.replace(/\/+$/, ''));

app.use(cors({
  origin: (origin, cb) => {
    // Allow requests with no origin (curl, Postman, server-to-server)
    if (!origin) return cb(null, true);
    if (ALLOWED_ORIGINS.includes(origin)) return cb(null, true);
    cb(new Error(`CORS: origin ${origin} not allowed`));
  },
  exposedHeaders: ['Content-Disposition'],
}));

// ─── File upload (stored in memory, max 500 MB) ───────────────────────────────
const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 500 * 1024 * 1024 },
});

// ─── Helpers ──────────────────────────────────────────────────────────────────

/** Detect source extension, handling compound ones like tar.gz */
function detectSrcExt(filename) {
  const lc = filename.toLowerCase();
  for (const c of COMPOUND_EXTS) {
    if (lc.endsWith('.' + c)) return c;
  }
  return lc.includes('.') ? lc.split('.').pop() : '';
}

/** Strip the detected extension from a filename */
function stripExt(filename, ext) {
  const suffix = '.' + ext;
  if (filename.toLowerCase().endsWith(suffix)) {
    return filename.slice(0, filename.length - suffix.length);
  }
  const dot = filename.lastIndexOf('.');
  return dot > 0 ? filename.slice(0, dot) : filename;
}

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get('/health', (_req, res) => res.json({ status: 'ok' }));

app.post('/api/convert', upload.single('file'), async (req, res) => {
  // ── Validate inputs ──────────────────────────────────────────────────────
  if (!req.file) {
    return res.status(400).json({ detail: 'No file uploaded.' });
  }

  const tgt      = (req.body.target_format ?? '').toLowerCase().trim();
  const filename = req.file.originalname || 'file';
  const srcExt   = detectSrcExt(filename);

  if (!srcExt) return res.status(400).json({ detail: 'Cannot determine source format (no extension).' });
  if (!tgt)    return res.status(400).json({ detail: 'No target format specified.' });
  if (srcExt === tgt) return res.status(400).json({ detail: 'Source and target formats are the same.' });

  // ── Route to correct converter ───────────────────────────────────────────
  const isArchiveSrc = ARCHIVE_EXTS.has(srcExt) || srcExt.startsWith('tar.');
  const isArchiveTgt = ARCHIVE_EXTS.has(tgt)    || tgt.startsWith('tar.');

  let resultBuffer;
  try {
    if (IMAGE_EXTS.has(srcExt)) {
      resultBuffer = await convertImage(req.file.buffer, srcExt, tgt);
    } else if (VIDEO_EXTS.has(srcExt) || AUDIO_EXTS.has(srcExt)) {
      resultBuffer = await convertVideoAudio(req.file.buffer, srcExt, tgt);
    } else if (DOC_EXTS.has(srcExt)) {
      resultBuffer = await convertDocument(req.file.buffer, srcExt, tgt);
    } else if (isArchiveSrc) {
      if (isArchiveTgt) {
        // Archive → archive format (repack)
        resultBuffer = await convertArchiveToArchive(req.file.buffer, srcExt, tgt);
      } else {
        // Archive → convert files inside (batch)
        resultBuffer = await convertZip(req.file.buffer, srcExt, tgt);
      }
    } else {
      return res.status(400).json({ detail: `Unsupported source format: .${srcExt}` });
    }
  } catch (err) {
    const message = err?.message || 'Conversion failed.';
    return res.status(422).json({ detail: message });
  }

  // ── Send result ──────────────────────────────────────────────────────────
  const nameBase    = stripExt(filename, srcExt);
  const outMime     = MIME_FOR_FORMAT[tgt] ?? 'application/octet-stream';
  const outFilename = isArchiveSrc && !isArchiveTgt
    ? `${nameBase}_converted.zip`    // batch result always a zip
    : `${nameBase}.${tgt}`;

  res.setHeader('Content-Type', outMime);
  res.setHeader('Content-Disposition', `attachment; filename="${outFilename}"`);
  res.send(resultBuffer);
});

// ─── Start ────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 8000;
app.listen(PORT, () => {
  console.log(`File Converter API running on http://localhost:${PORT}`);
  console.log(`Allowed origins: ${ALLOWED_ORIGINS.join(', ')}`);

  // Keep-alive: ping own /health every 14 min so Render free tier never idles
  const selfUrl = process.env.RENDER_EXTERNAL_URL
    ? `${process.env.RENDER_EXTERNAL_URL.replace(/\/+$/, '')}/health`
    : `http://localhost:${PORT}/health`;

  const https = require('https');
  const http  = require('http');
  const ping  = selfUrl.startsWith('https') ? https : http;

  setInterval(() => {
    ping.get(selfUrl, (r) => { r.resume(); }).on('error', () => {});
  }, 14 * 60 * 1000);
});
