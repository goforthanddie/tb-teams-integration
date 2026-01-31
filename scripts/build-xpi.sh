#!/usr/bin/env bash
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")/.." && pwd)"
OUT_DIR="${ROOT_DIR}/dist"
XPI_NAME="tb-teams-integration"

rm -rf "${OUT_DIR}"
mkdir -p "${OUT_DIR}"

XPI_PATH="${OUT_DIR}/${XPI_NAME}.xpi"

(
  cd "${ROOT_DIR}"
  zip -r "${XPI_PATH}" \
    manifest.json background.js \
    experiments options icons shared README.md >/dev/null
)

echo "Built: ${XPI_PATH}"
