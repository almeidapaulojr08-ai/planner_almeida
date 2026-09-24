# -*- coding: utf-8 -*-
"""Carimba ?v=<timestamp> em todos os <script src="js/..."> e <link href="css/..."> do index.html.

Por quê: o GitHub Pages cacheia arquivos por 10 min. Sem o carimbo, um navegador pode carregar
o index.html novo com um js/*.js velho (ou vice-versa) e quebrar. Com o carimbo, cada versão
do index.html só referencia os arquivos daquela mesma versão.

Rodar ANTES de cada commit que mexa em css/ ou js/:   python tools/stamp.py
"""
import io, os, re, time
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
P = os.path.join(ROOT, 'index.html')
s = io.open(P, encoding='utf-8', newline='').read()
stamp = time.strftime('%Y%m%d%H%M%S', time.gmtime())
s2, n = re.subn(r'((?:src|href)="(?:js|css)/[^"?]+)\?v=\d+"', r'\1?v=' + stamp + '"', s)
io.open(P, 'w', encoding='utf-8', newline='').write(s2)
print(f'{n} referências carimbadas com v={stamp}')
