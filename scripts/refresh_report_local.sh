#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
PROJECT_ROOT="$(cd "${SCRIPT_DIR}/.." && pwd)"

cd "${PROJECT_ROOT}"

if [[ ! -x "${PROJECT_ROOT}/.venv/bin/python" ]]; then
  echo "Missing project virtualenv at ${PROJECT_ROOT}/.venv/bin/python" >&2
  echo "Create the virtualenv or adjust the script before running a refresh." >&2
  exit 1
fi

exec "${PROJECT_ROOT}/.venv/bin/python" "${PROJECT_ROOT}/report_refresh.py" "$@"
