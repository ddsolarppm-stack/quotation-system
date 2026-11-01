"""Check that index.html is safe to paste into Google Apps Script HtmlService.

Apps Script's HTML sanitizer tokenizes the page's JavaScript loosely and has broken this app twice:
  1. 4-byte UTF-8 characters (emoji) got mangled.
  2. A "//" inside a string ('https://...') was treated as a comment and the rest of the line was cut,
     which left an unterminated string -> "SyntaxError: Invalid or unexpected token".
     A quote character inside a regex literal (e.g. /["']/) can confuse its string tracking the same way.
So inside <script> we avoid: "//" except as a real comment, quotes in regex literals, and "/*" or "*/"
outside comments. Run before every push:  python3 scripts/check-gas-safe.py
"""
import re
import sys

path = sys.argv[1] if len(sys.argv) > 1 else 'index.html'
text = open(path, encoding='utf-8').read()
lines = text.split('\n')
bad = []

for no, line in enumerate(lines, 1):
    for ch in line:
        if ord(ch) > 0xFFFF or ch in '  ﻿':
            bad.append(f'{path}:{no}: U+{ord(ch):04X} {ch!r} (4-byte / invisible character)')

in_script = in_block = False
for no, line in enumerate(lines, 1):
    low = line.strip().lower()
    if low.startswith('<script'):
        in_script = True
        continue
    if low.startswith('</script'):
        in_script = False
        continue
    if not in_script:
        continue
    code = line.strip()
    if in_block:
        if '*/' in code:
            in_block = False
        continue
    if code.startswith('/*'):
        in_block = '*/' not in code
        continue
    if code.startswith('//') or code.startswith('*'):
        continue
    # a trailing "  // comment" after code is fine; any other "//" sits inside a string or regex
    body = re.sub(r'\s//\s.*$', '', line)
    if '//' in body:
        bad.append(f'{path}:{no}: "//" inside code/string -> write it as "/\\/"')
    if '/*' in body or '*/' in body:
        bad.append(f'{path}:{no}: "/*" or "*/" outside a comment')
    for rx in re.finditer(r'(?:(?<=[(=,:!&|?])|(?<=[(=,:!&|?] ))/(?![/*])((?:\\.|\[(?:\\.|[^\]])*\]|[^/\n\\])+)/[gimsuy]*', body):
        if re.search(r'[\'"`]', rx.group(1)):
            bad.append(f'{path}:{no}: quote inside regex literal {rx.group(0)[:40]!r} -> use \\x22 \\x27')

if path.endswith('.html') and not text.rstrip().endswith('</html>'):
    bad.append(f'{path}: does not end with </html>')
print('\n'.join(bad) or f'{path}: OK, safe for Apps Script')
sys.exit(1 if bad else 0)
