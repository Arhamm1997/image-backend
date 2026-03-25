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
const { MIME_FOR_FORMAT, convertImage, convertVideoAudio, convertDocument, convertZip } = require('./converters');

const app = express();

// ─── File-type sets ───────────────────────────────────────────────────────────
const IMAGE_EXTS = new Set(['jpg','jpeg','png','webp','gif','bmp','ico','tiff','tif','svg','avif']);
const VIDEO_EXTS = new Set(['mp4','webm','avi','mov','mkv','flv','wmv']);
const AUDIO_EXTS = new Set(['mp3','wav','ogg','aac','flac','m4a','wma']);
const DOC_EXTS   = new Set(['pdf','docx','doc','txt','rtf','odt','xlsx','xls','pptx','ppt','csv']);

// ─── CORS ─────────────────────────────────────────────────────────────────────
const ALLOWED_ORIGINS = [
  'http://localhost:5173',
  'http://localhost:3000',
  'http://127.0.0.1:5173',
  'https://image-frontend-black.vercel.app',
  process.env.FRONTEND_URL,
].filter(Boolean).map(o => o.replace(/\/+$/, ''));

app.use(cors({
  origin: ALLOWED_ORIGINS,
  exposedHeaders: ['Content-Disposition'],
}));

// ─── File upload (stored in memory, max 500 MB) ───────────────────────────────
const upload = multer({
  storage: multer.memoryStorage(),
  limits:  { fileSize: 500 * 1024 * 1024 },
});

// ─── Routes ───────────────────────────────────────────────────────────────────

app.get('/health', (_req, res) => res.json({ status: 'ok' }));

app.post('/api/convert', upload.single('file'), async (req, res) => {
  // ── Validate inputs ──────────────────────────────────────────────────────
  if (!req.file) {
    return res.status(400).json({ detail: 'No file uploaded.' });
  }

  const tgt      = (req.body.target_format ?? '').toLowerCase().trim();
  const filename = req.file.originalname || 'file';
  const srcExt   = filename.includes('.') ? filename.split('.').pop().toLowerCase() : '';

  if (!srcExt) return res.status(400).json({ detail: 'Cannot determine source format (file has no extension).' });
  if (!tgt)    return res.status(400).json({ detail: 'No target format specified.' });
  if (srcExt === tgt) return res.status(400).json({ detail: 'Source and target formats are the same.' });

  // ── Route to correct converter ───────────────────────────────────────────
  const isZipInput = srcExt === 'zip';
  let resultBuffer;
  try {
    if (IMAGE_EXTS.has(srcExt)) {
      resultBuffer = await convertImage(req.file.buffer, srcExt, tgt);
    } else if (VIDEO_EXTS.has(srcExt) || AUDIO_EXTS.has(srcExt)) {
      resultBuffer = await convertVideoAudio(req.file.buffer, srcExt, tgt);
    } else if (DOC_EXTS.has(srcExt)) {
      resultBuffer = await convertDocument(req.file.buffer, srcExt, tgt);
    } else if (isZipInput) {
      resultBuffer = await convertZip(req.file.buffer, tgt);
    } else {
      return res.status(400).json({ detail: `Unsupported source format: .${srcExt}` });
    }
  } catch (err) {
    const message = err?.message || 'Conversion failed.';
    return res.status(422).json({ detail: message });
  }

  // ── Send result ──────────────────────────────────────────────────────────
  const nameBase = filename.slice(0, filename.lastIndexOf('.'));
  const outMime     = isZipInput ? 'application/zip'                         : (MIME_FOR_FORMAT[tgt] ?? 'application/octet-stream');
  const outFilename = isZipInput ? `${nameBase}_to_${tgt}.zip`               : `${nameBase}.${tgt}`;
  res.setHeader('Content-Type', outMime);
  res.setHeader('Content-Disposition', `attachment; filename="${outFilename}"`);
  res.send(resultBuffer);
});

// ─── Start ────────────────────────────────────────────────────────────────────
const PORT = process.env.PORT || 8000;
app.listen(PORT, () => {
  console.log(`File Converter API running on http://localhost:${PORT}`);
  console.log(`Allowed origins: ${ALLOWED_ORIGINS.join(', ')}`);
});
