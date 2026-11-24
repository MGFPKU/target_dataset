# -*- coding: utf-8 -*-
"""Fix formatting issues in China_Climate_Target_Tracker_cn/en.xlsx.

1. 说明/README sheet:
   - restore deleted B-column labels (基线/目标年份或时期/计数/目标类别/责任主体)
   - fix merge ranges (B39:B45 / B46:B62 / B63:B68) and remove stale merges
   - realign E column category descriptions after the 18->16 category
     consolidation (delete 煤电转型 + 绿色创新 descriptions, shift rest up,
     drop now-empty rows 63-64)
2. Data sheets + 来源/Sources sheet:
   - reapply original per-column styles (font/fill/border/alignment/numfmt)
     to rows/cells added or edited by scripts (Calibri 10, no fill/border)
   - normalize stray font-size-12 cells to the sheet's dominant font
"""
import copy
import shutil
from collections import Counter

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill

CN = 'China_Climate_Target_Tracker_cn.xlsx'
EN = 'China_Climate_Target_Tracker_en.xlsx'

DATA_CN = ['温室气体排放', '能源|电力', '工业', '建筑', '交通',
           '土地利用、土地利用变化和林业', '循环经济', '污染', '金融']
DATA_EN = ['GHG Emissions', 'Energy|Power', 'Industry', 'Buildings', 'Transport',
           'LULUCF', 'Circular Economy', 'Pollution', 'Finance']
README_CN, README_EN = '说明', 'README'
SRC_CN, SRC_EN = '来源', 'Sources'


def copy_style(src, dst):
    dst.font = copy.copy(src.font)
    dst.fill = copy.copy(src.fill)
    dst.border = copy.copy(src.border)
    dst.alignment = copy.copy(src.alignment)
    dst.number_format = src.number_format


def fix_readme(wb, cn):
    ws = wb[README_CN if cn else README_EN]
    bad_merges = ['B36:B42', 'B43:B61', 'B62:B67', 'C16:S16' if cn else 'C15:S15']
    existing = {str(m) for m in ws.merged_cells.ranges}
    for m in bad_merges:
        if m in existing:
            ws.unmerge_cells(m)

    labels = ({'B37': '基线', 'B38': '目标年份或时期', 'B39': '计数',
               'B46': '目标类别', 'B65': '责任主体'} if cn else
              {'B37': 'Baseline', 'B38': 'Target_Year_or_Period', 'B39': 'Count',
               'B46': 'Target_Category', 'B65': 'Accountability'})
    ref = ws['B36']
    for coord, val in labels.items():
        c = ws[coord]
        c.value = val
        copy_style(ref, c)

    # realign E column (col 5): move descriptions up after deleting the two
    # orphan entries (E56 煤电转型, E59 绿色创新) left by the 18->16 consolidation
    moves = {57: 56, 58: 57, 60: 58, 61: 59, 62: 60, 63: 61, 64: 62}
    for src, dst in moves.items():
        ws.cell(row=dst, column=5).value = ws.cell(row=src, column=5).value
    ws.cell(row=63, column=5).value = None
    ws.cell(row=64, column=5).value = None

    # rows 63-64 are now empty -> move rows 65-79 up to 63-77 (manual, keeps styles)
    for r in range(65, 80):
        for cidx in range(1, 21):
            src = ws.cell(row=r, column=cidx)
            dst = ws.cell(row=r - 2, column=cidx)
            dst.value = src.value
            copy_style(src, dst)
    for r in range(78, 80):
        for cidx in range(1, 21):
            c = ws.cell(row=r, column=cidx)
            c.value = None
            c.font = Font()
            c.fill = PatternFill()
            c.border = Border()
            c.alignment = Alignment()

    for m in ['B39:B45', 'B46:B62', 'B63:B68']:
        ws.merge_cells(m)


def fix_data_sheet(ws, ncols):
    """Normalise formatting of script-added/edited cells to the sheet template."""
    fonts = Counter()
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ncols):
        for c in row:
            if c.value is not None:
                fonts[(c.font.name, c.font.size)] += 1
    dom = fonts.most_common(1)[0][0]

    ref_row = None
    for r in range(2, ws.max_row + 1):
        c = ws.cell(row=r, column=1)
        if c.value is not None and (c.font.name, c.font.size) == dom and c.fill.patternType:
            ref_row = r
            break
    if ref_row is None:
        raise RuntimeError(f'{ws.title}: no reference row found (dom={dom})')
    refs = [ws.cell(row=ref_row, column=cidx) for cidx in range(1, ncols + 1)]

    heights = Counter(h for r in range(2, ws.max_row + 1)
                      if (h := ws.row_dimensions[r].height) is not None)
    norm_h = heights.most_common(1)[0][0] if heights else None

    n_rows, n_cells = 0, 0
    for r in range(2, ws.max_row + 1):
        chk = ws.cell(row=r, column=1)
        row_bad = chk.value is not None and (chk.font.name, chk.font.size) != dom
        if row_bad:
            for cidx in range(1, ncols + 1):
                copy_style(refs[cidx - 1], ws.cell(row=r, column=cidx))
            n_rows += 1
        else:
            for cidx in range(1, ncols + 1):
                c = ws.cell(row=r, column=cidx)
                if c.value is not None and (c.font.name, c.font.size) != dom:
                    if c.fill.patternType or c.border.left.style:
                        c.font = copy.copy(refs[cidx - 1].font)
                    else:
                        copy_style(refs[cidx - 1], c)
                    n_cells += 1
        if row_bad and norm_h and ws.row_dimensions[r].height is None:
            ws.row_dimensions[r].height = norm_h
    return n_rows, n_cells


def fix_sources_sheet(ws):
    """Normalise script-added rows in 来源/Sources (per-column template)."""
    ref_row = None
    for r in range(2, ws.max_row + 1):
        if ((ws.cell(row=r, column=2).font.name, ws.cell(row=r, column=2).font.size) == ('Arial Narrow', 11.0)
                and ws.cell(row=r, column=3).fill.patternType):
            ref_row = r
            break
    if ref_row is None:
        raise RuntimeError(f'{ws.title}: no reference row found')
    refs = {cidx: ws.cell(row=ref_row, column=cidx) for cidx in range(2, 16)}

    n_rows, n_cells = 0, 0
    for r in range(2, ws.max_row + 1):
        chk = ws.cell(row=r, column=2)
        row_bad = chk.value is not None and (chk.font.name, chk.font.size) == ('Calibri', 10.0)
        if row_bad:
            for cidx in range(1, 16):
                copy_style(refs[cidx if cidx >= 2 else 2], ws.cell(row=r, column=cidx))
            n_rows += 1
        else:
            for cidx in range(2, 16):
                c = ws.cell(row=r, column=cidx)
                if c.value is not None and (c.font.name, c.font.size) == ('Calibri', 10.0):
                    if c.fill.patternType or c.border.left.style:
                        c.font = copy.copy(refs[cidx].font)
                    else:
                        copy_style(refs[cidx], c)
                    n_cells += 1
    return n_rows, n_cells


def main():
    for f in (CN, EN):
        shutil.copyfile(f, f.replace('.xlsx', '.pre_formatfix.bak.xlsx'))
    log = []
    for f, data_names, cn in ((CN, DATA_CN, True), (EN, DATA_EN, False)):
        wb = openpyxl.load_workbook(f)
        fix_readme(wb, cn)
        log.append(f'{f} 说明/README fixed')
        for name in data_names:
            ws = wb[name]
            ncols = 13 if name in ('能源|电力', 'Energy|Power') else 12
            nr, nc = fix_data_sheet(ws, ncols)
            log.append(f'  {name}: {nr} rows restyled, {nc} cells font-fixed')
        ws = wb[SRC_CN if cn else SRC_EN]
        nr, nc = fix_sources_sheet(ws)
        log.append(f'  {SRC_CN if cn else SRC_EN}: {nr} rows restyled, {nc} cells font-fixed')
        wb.save(f)
    print('\n'.join(log))


if __name__ == '__main__':
    main()
