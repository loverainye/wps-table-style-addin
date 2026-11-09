#!/usr/bin/env python3
"""Optional browser UI preview for development; not required by the offline add-in."""
from functools import partial
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
import argparse

parser = argparse.ArgumentParser()
parser.add_argument('--port', type=int, default=3989)
args = parser.parse_args()
root = Path(__file__).resolve().parent.parent / 'dist'
if not (root / 'js/main.js').exists():
    raise SystemExit('请先运行 npm run build。')

class Handler(SimpleHTTPRequestHandler):
    def end_headers(self):
        self.send_header('Cache-Control', 'no-store')
        self.send_header('X-Content-Type-Options', 'nosniff')
        super().end_headers()

server = ThreadingHTTPServer(('127.0.0.1', args.port), partial(Handler, directory=str(root)))
print('可选开发预览服务；离线加载项运行无需此服务。', flush=True)
print(f'界面预览：http://127.0.0.1:{args.port}/editor.html?preview=1', flush=True)
server.serve_forever()
