# ─── Build stage ──────────────────────────────────────────────────────────────
# Install Node deps inside the container so native modules (sharp) are built
# for Linux, not Windows/macOS.
FROM node:20-slim AS builder

WORKDIR /app
COPY package*.json ./
RUN npm ci --only=production

# ─── Runtime stage ────────────────────────────────────────────────────────────
FROM node:20-slim

# LibreOffice (writer / calc / impress) for complex document conversions.
# ffmpeg-static bundles its own FFmpeg binary — no system ffmpeg needed.
RUN apt-get update && apt-get install -y --no-install-recommends \
        libreoffice-writer \
        libreoffice-calc \
        libreoffice-impress \
        fonts-liberation \
        fontconfig \
        poppler-utils \
    && fc-cache -fv \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copy pre-built node_modules from builder (correct Linux binaries for sharp)
COPY --from=builder /app/node_modules ./node_modules
COPY . .

EXPOSE 8000

CMD ["node", "index.js"]
