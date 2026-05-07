# Menggunakan image Bun
FROM oven/bun:1-slim

# Install dependencies yang dibutuhkan oleh sistem untuk menjalankan browser
RUN apt-get update && apt-get install -y \
    libnss3 \
    libnspr4 \
    libatk1.0-0 \
    libatk-bridge2.0-0 \
    libcups2 \
    libdrm2 \
    libxkbcommon0 \
    libxcomposite1 \
    libxdamage1 \
    libxext6 \
    libxfixes3 \
    librandr2 \
    libgbm1 \
    libpango-1.0-0 \
    libcairo2 \
    libasound2 \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Set environment variable agar Playwright tahu di mana browser diinstall
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

# Copy dependency files
COPY package.json bun.lock ./

# Install node modules
RUN bun install --frozen-lockfile

# Install Playwright Chromium saja (hemat size) ke folder yang sudah kita tentukan
RUN bunx playwright install chromium

# Copy source code
COPY . .

# Jalankan bot
CMD ["bun", "run", "bot"]
