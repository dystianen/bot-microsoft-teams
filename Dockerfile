# Menggunakan image Bun (terbaru & sangat cepat)
FROM oven/bun:1-slim

WORKDIR /app

# Copy dependency files
COPY package.json bun.lock ./

# Install menggunakan bun
RUN bun install --frozen-lockfile

# Copy source code
COPY . .

# Jalankan bot menggunakan runtime Bun
CMD ["bun", "run", "bot"]
