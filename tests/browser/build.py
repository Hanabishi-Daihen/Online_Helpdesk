"""Build a test page: index.html + a fake google.script.run (mock/head.html) + extra mock data."""
import sys, os
here = os.path.dirname(os.path.abspath(__file__)); root = os.path.join(here, '..', '..')
page = open(os.path.join(root, 'index.html'), encoding='utf-8').read()
head = open(os.path.join(here, '..', 'mock', 'head.html'), encoding='utf-8').read()
extra = open(os.path.join(here, '..', 'mock', sys.argv[2]), encoding='utf-8').read() if len(sys.argv) > 2 else ''
k = page.find('<head>') + 6
os.makedirs(os.path.join(here, '..', 'out'), exist_ok=True)
open(os.path.join(here, '..', 'out', sys.argv[1]), 'w', encoding='utf-8').write(page[:k] + head + extra + page[k:])
