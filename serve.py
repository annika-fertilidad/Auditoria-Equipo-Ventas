#!/usr/bin/env python3
"""Tiny static file server for local preview (avoids http.server __main__ getcwd issue)."""
import http.server, socketserver, os, sys

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 3456
ROOT = os.path.dirname(os.path.abspath(__file__))

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **k):
        super().__init__(*a, directory=ROOT, **k)

with socketserver.TCPServer(("", PORT), Handler) as httpd:
    print(f"Sirviendo {ROOT} en http://localhost:{PORT}")
    httpd.serve_forever()
