FROM mcr.microsoft.com/playwright:v1.59.1-noble

RUN apt-get update && apt-get install -y unzip && rm -rf /var/lib/apt/lists/*

RUN curl -fsSL https://bun.sh/install | bash
ENV PATH="/root/.bun/bin:${PATH}"

WORKDIR /app

COPY package.json bun.lock ./
RUN bun install --frozen-lockfile

# Chromium sudah tersedia di base image playwright
# Cukup set environment variable path-nya
ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

COPY . .

CMD ["bun", "run", "bot"]