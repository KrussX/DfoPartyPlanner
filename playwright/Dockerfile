FROM mcr.microsoft.com/playwright:v1.52.0-noble

WORKDIR /app

COPY package.json package-lock.json ./
RUN npm ci --omit=dev

COPY server.js .

# Playwright image already has browsers installed, but ensure Chromium is there
RUN npx playwright install chromium

EXPOSE 3001

CMD ["node", "server.js"]
