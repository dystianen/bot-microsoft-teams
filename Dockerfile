FROM oven/bun:1-slim

WORKDIR /app

ENV PLAYWRIGHT_BROWSERS_PATH=/ms-playwright

COPY package.json bun.lock ./

RUN bun install --frozen-lockfile

RUN bunx playwright install --with-deps chromium

COPY . .

CMD ["bun", "run", "bot"]