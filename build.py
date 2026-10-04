#!/usr/bin/env python3
"""Assemble the single-file app: build.py [state.json] [out.html]

With no arguments it writes a blank index.html (no data). Pass a backup JSON
exported from the app to bake that data into a private copy instead.
"""
import json, pathlib, sys

root = pathlib.Path(__file__).parent
css = (root / "src/app.css").read_text(encoding="utf8")
code = (root / "src/app.js").read_text(encoding="utf8")
state = json.loads(pathlib.Path(sys.argv[1]).read_text(encoding="utf8")) if len(sys.argv) > 1 else None
out = pathlib.Path(sys.argv[2]) if len(sys.argv) > 2 else root / "index.html"
blob = json.dumps(state, ensure_ascii=False, separators=(",", ":")).replace("<", "\\u003c")

out.write_text(
    '<!doctype html>\n<html lang="he" dir="rtl">\n<head>\n<meta charset="utf-8">\n'
    '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
    "<title>ניהול הוצאות חכם</title>\n"
    '<link rel="preconnect" href="https://fonts.googleapis.com">\n'
    '<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Assistant:wght@400;600;700&family=Secular+One&display=swap">\n'
    f'<style id="APPCSS">{css}</style>\n</head>\n<body>\n'
    f'<script type="application/json" id="APPSTATE">{blob}</script>\n'
    '<div id="app" dir="rtl" lang="he"></div>\n'
    f'<script id="APPCODE">{code}</script>\n</body>\n</html>',
    encoding="utf8",
)
print(f"wrote {out} ({out.stat().st_size:,} bytes, {'with data' if state else 'blank'})")
