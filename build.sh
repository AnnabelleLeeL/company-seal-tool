#!/usr/bin/env bash
set -euo pipefail
mkdir -p dist/.openai dist/server
mkdir -p dist/vendor
cp index.html styles.css app.js dist/
cp vendor/pdf.min.js vendor/pdf.worker.min.js vendor/pdf-lib.min.js dist/vendor/
cp cutout.html dist/cutout.html
cp cutout.css editor.css mobile-fix.css cutout.js smart-color.js dist/
if [[ -f .openai/hosting.json ]]; then
  cp .openai/hosting.json dist/.openai/hosting.json
fi
cp server/index.js dist/server/index.js
