#!/usr/bin/env bash
# Build SimReal-Seed-Deck-EN.pptx: pptxgenjs -> post-process (Latin font profile).
set -euo pipefail
cd "$(dirname "$0")"
export NODE_PATH="${NODE_PATH:-$(npm root -g)}"
node build_en.js
python3 tools/postprocess.py build/en/raw.pptx ../SimReal-Seed-Deck-EN.pptx en
