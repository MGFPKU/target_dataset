"""Unify percentage notation in Targets_en.xlsx: '<digit> percent' -> '<digit>%'.

- Applies to every string cell in the workbook (magnitudes, sentences, metrics, baselines).
- 'percentage' / 'percentage points' are deliberately untouched (no digit directly before).
- Idempotent. Run from target_dataset/ with plain python.
"""
import datetime
import re
import shutil
from copy import copy

import openpyxl
from openpyxl.cell.cell import MergedCell

PAT = re.compile(r"(\d)\s*percent\b", re.IGNORECASE)

CHANGELOG = (
    '2026-09-07: Unified percentage notation across the English version ("X percent" → "X%").'
)


def unify(wb: openpyxl.Workbook) -> int:
    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c, MergedCell):
                    continue
                v = c.value
                if not isinstance(v, str):
                    continue
                new_v = PAT.sub(r"\1%", v)
                if new_v != v:
                    c.value = new_v
                    changed += 1
    return changed


def add_changelog(wb: openpyxl.Workbook) -> None:
    ws = wb["README"]
    for row in ws.iter_rows(min_col=2, max_col=2):
        if row[0].value == CHANGELOG:
            return  # already present (idempotent re-run)
    ws.insert_rows(6)
    src = ws.cell(row=7, column=2)
    dst = ws.cell(row=6, column=2)
    dst.value = CHANGELOG
    dst.font = copy(src.font)
    dst.alignment = copy(src.alignment)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = src.number_format
    ws.row_dimensions[6].height = 15.0
    ws.cell(row=5, column=3).value = datetime.datetime(2026, 9, 7)


if __name__ == "__main__":
    path = "Targets_en.xlsx"
    backup = path.replace(".xlsx", ".pre_percent.bak")
    shutil.copy2(path, backup)
    print(f"backup -> {backup}")

    wb = openpyxl.load_workbook(path)
    changed = unify(wb)
    add_changelog(wb)
    wb.save(path)
    print(f"{path}: {changed} cells updated, changelog added")
