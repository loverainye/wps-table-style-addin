#!/usr/bin/env bash
set -euo pipefail
addin_root="$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)"
exec bash "$addin_root/install.sh" --remove "$@"
