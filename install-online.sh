#!/usr/bin/env bash
set -euo pipefail
VERSION="1.0"
OWNER="loverainye"
REPO="wps-table-style-addin"
BASE="https://github.com/${OWNER}/${REPO}/releases/download/v${VERSION}"
TMP="$(mktemp -d)"
trap 'rm -rf "$TMP"' EXIT
case "$(uname -s)" in
  Linux*) ASSET="wps-table-style-addin-linux-${VERSION}.tar.gz" ;;
  Darwin*) ASSET="wps-table-style-addin-macos-${VERSION}.tar.gz" ;;
  *) echo '此脚本适用于 Linux 或 macOS；Windows 请运行 install-online.ps1。' >&2; exit 1 ;;
esac
curl -fL --retry 3 "${BASE}/${ASSET}" -o "${TMP}/${ASSET}"
tar -xzf "${TMP}/${ASSET}" -C "$TMP"
INSTALL="$(find "$TMP" -maxdepth 3 -name install.sh -type f | head -1)"
[[ -n "$INSTALL" ]] || { echo '安装包中缺少 install.sh。' >&2; exit 1; }
bash "$INSTALL" "$@"
