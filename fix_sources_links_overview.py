# -*- coding: utf-8 -*-
"""Pass 2 fixes:
1. 来源/Sources: added rows were restyled with the row-2 title-band fill by
   mistake; re-copy per-column styles from the nearest good data-row neighbour
   (each section has its own fill colour).
2. Add missing hyperlinks to all URL cells in column N (both files).
3. 概览/Overview "targets by category" block: remove the two leftover empty
   bordered rows (deleted categories) and close the block border at row 21.
"""
import copy
import re
import shutil

import openpyxl
from openpyxl.styles import Border
from openpyxl.worksheet.hyperlink import Hyperlink

CN = 'China_Climate_Target_Tracker_cn.xlsx'
EN = 'China_Climate_Target_Tracker_en.xlsx'
SRC_CN, SRC_EN = '来源', 'Sources'
OV_CN, OV_EN = '概览', 'Overview'

ADDED_ROWS = ([14, 15] + list(range(177, 193)) + [275, 323, 324]
              + list(range(404, 408)) + list(range(466, 470))
              + [474, 476] + list(range(480, 486)) + [505])
CODE_RE = re.compile(r'^[A-Z]+[0-9]{4}$')


def copy_style(src, dst):
    dst.font = copy.copy(src.font)
    dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border)
    dst.alignment = copy.copy(src.alignment)
    dst.number_format = src.number_format


def is_good_data_row(ws, r):
    b = ws.cell(row=r, column=2)
    return (b.value is not None and CODE_RE.match(str(b.value))
            and (b.font.name, b.font.size) == ('Arial Narrow', 11.0))


def restyle_added_rows(ws):
    added = set(ADDED_ROWS)
    fixed = 0
    for r in sorted(added):
        nb = None
        for rr in range(r - 1, 1, -1):
            if rr in added:
                continue
            if is_good_data_row(ws, rr):
                nb = rr
                break
        if nb is None:
            for rr in range(r + 1, ws.max_row + 1):
                if rr in added:
                    continue
                if is_good_data_row(ws, rr):
                    nb = rr
                    break
        if nb is None:
            raise RuntimeError(f'no good neighbour for row {r}')
        for cidx in range(1, 16):
            copy_style(ws.cell(row=nb, column=cidx), ws.cell(row=r, column=cidx))
        fixed += 1
    return fixed


def add_hyperlinks(ws):
    n = 0
    for r in range(2, ws.max_row + 1):
        c = ws.cell(row=r, column=14)
        v = c.value
        if isinstance(v, str) and v.startswith('http') and not c.hyperlink:
            c.hyperlink = Hyperlink(ref=c.coordinate, target=v)
            n += 1
    return n


def fix_overview(ws):
    # category block ends at row 21; rows 22-23 are leftover empties
    for r in (22, 23):
        for col in (17, 18):
            ws.cell(row=r, column=col).border = Border()
    for col in (17, 18):
        c = ws.cell(row=21, column=col)
        b = c.border
        c.border = Border(left=b.left, right=b.right, top=b.top,
                          bottom=openpyxl.styles.Side(style='medium'))


def main():
    for f in (CN, EN):
        shutil.copyfile(f, f.replace('.xlsx', '.pre_linksfix.bak.xlsx'))
    for f, cn in ((CN, True), (EN, False)):
        wb = openpyxl.load_workbook(f)
        ws = wb[SRC_CN if cn else SRC_EN]
        n1 = restyle_added_rows(ws)
        n2 = add_hyperlinks(ws)
        fix_overview(wb[OV_CN if cn else OV_EN])
        wb.save(f)
        print(f'{f}: restyled {n1} added source rows, added {n2} hyperlinks, overview block closed')


if __name__ == '__main__':
    main()
