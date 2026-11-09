#!/usr/bin/env bash
set -euo pipefail

if ! command -v python3 >/dev/null 2>&1; then
  echo '安装需要 Python 3.9 或更高版本（只用于安装，不需要 pip 依赖）。Windows 请运行 install.ps1。' >&2
  exit 1
fi
if ! python3 -c 'import sys; sys.exit(0 if sys.version_info >= (3, 9) else 1)'; then
  echo 'Python 版本过旧，安装需要 Python 3.9 或更高版本。' >&2
  exit 1
fi
addin_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
exec python3 "$addin_root/scripts/install.py" "$@"
