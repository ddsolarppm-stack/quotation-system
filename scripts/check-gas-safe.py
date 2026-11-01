"""Check that index.html is safe to paste into Google Apps Script HtmlService.

Apps Script has mangled 4-byte UTF-8 characters (emoji) in served HTML, which broke the
whole app script. Run before every push:  python3 scripts/check-gas-safe.py
"""
import sys

path = sys.argv[1] if len(sys.argv) > 1 else 'index.html'
text = open(path, encoding='utf-8').read()
bad = []
for no, line in enumerate(text.split('\n'), 1):
    for ch in line:
        if ord(ch) > 0xFFFF or ch in '  ﻿':
            bad.append(f'{path}:{no}: U+{ord(ch):04X} {ch!r}')
if not text.rstrip().endswith('</html>'):
    bad.append(f'{path}: does not end with </html>')
print('\n'.join(bad) or f'{path}: OK, safe for Apps Script')
sys.exit(1 if bad else 0)
