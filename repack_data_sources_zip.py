# -*- coding: utf-8 -*-
"""重打包 Data_sources.zip（顶层 4 个子文件夹、无 Data_sources/ 前缀、ZIP_DEFLATED）。"""
import os
import zipfile

BASE = 'Data_sources'
OUT = 'Data_sources.zip'

entries = []
for folder in sorted(os.listdir(BASE)):
    fp = os.path.join(BASE, folder)
    if not os.path.isdir(fp):
        continue
    for root, dirs, files in os.walk(fp):
        for name in files:
            full = os.path.join(root, name)
            arc = os.path.relpath(full, BASE).replace(os.sep, '/')
            entries.append((full, arc))

print('entries:', len(entries))
with zipfile.ZipFile(OUT, 'w', zipfile.ZIP_DEFLATED) as zf:
    for i, (full, arc) in enumerate(entries, 1):
        zf.write(full, arc)
        if i % 2000 == 0:
            print(i, '/', len(entries))
print('done:', OUT, os.path.getsize(OUT), 'bytes')
