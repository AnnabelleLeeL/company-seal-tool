FROM node:20-bookworm-slim

RUN apt-get update \
  && apt-get install -y --no-install-recommends libreoffice-writer fonts-noto-cjk fonts-liberation \
  && rm -rf /var/lib/apt/lists/*

WORKDIR /app
COPY package.json ./
COPY index.html styles.css app.js cutout.html cutout.css editor.css mobile-fix.css cutout.js smart-color.js ./
COPY vendor ./vendor
COPY server ./server

ENV PORT=10000
EXPOSE 10000
CMD ["node", "server/index.js"]
