#!/usr/bin/env bash
# Build SimReal-BP-ZH.pptx from source: pptxgenjs -> post-process -> validate.
set -euo pipefail
cd "$(dirname "$0")"
export NODE_PATH="${NODE_PATH:-$(npm root -g)}"
node build.js
python3 tools/postprocess.py build/raw.pptx ../SimReal-BP-ZH.pptx
