#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────────────────────
# build_layer.sh
# Builds a Lambda-compatible layer zip for generate_report.py
# Uses the official AWS SAM Docker image (Amazon Linux 2023 / Python 3.12)
# to ensure native C extensions (lxml) are compiled for the Lambda runtime.
#
# Prerequisites: Docker must be running.
#
# Usage:
#   chmod +x build_layer.sh
#   ./build_layer.sh
#
# Output:
#   20260409_lambda_layer.zip  (ready to upload to AWS Lambda Layers)
#
# Layer dependencies (generate_report.py):
#   - python-docx  → from docx import Document, from docx.oxml.ns import qn
#   - lxml         → from lxml import etree
#   - pandoc is a binary dependency handled separately (e.g. a pandoc Lambda layer)
# ─────────────────────────────────────────────────────────────────────────────
set -euo pipefail

LAYER_NAME="20260409_lambda_layer"
PYTHON_RUNTIME="python3.12"
IMAGE="public.ecr.aws/sam/build-${PYTHON_RUNTIME}:latest"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

echo "→ Pulling Lambda build image: ${IMAGE}"
docker pull "${IMAGE}"

echo "→ Cleaning previous python/ build dir (if any)"
rm -rf "${SCRIPT_DIR}/python"

echo "→ Installing dependencies into python/"
docker run --rm \
  -v "${SCRIPT_DIR}":/var/task \
  -w /var/task \
  "${IMAGE}" \
  pip install \
    --quiet \
    --upgrade \
    --target python/ \
    -r layer_requirements.txt

echo "→ Zipping layer..."
cd "${SCRIPT_DIR}"
zip -r "${LAYER_NAME}.zip" python/ --quiet

echo "✓ Done → ${SCRIPT_DIR}/${LAYER_NAME}.zip"
echo "  Size: $(du -sh ${LAYER_NAME}.zip | cut -f1)"
echo ""
echo "Upload with:"
echo "  aws lambda publish-layer-version \\"
echo "    --layer-name ${LAYER_NAME} \\"
echo "    --zip-file fileb://${LAYER_NAME}.zip \\"
echo "    --compatible-runtimes ${PYTHON_RUNTIME} \\"
echo "    --compatible-architectures x86_64"
