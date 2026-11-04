# -*- coding: utf-8 -*-
"""修复来源表 2026 批次行的气候关联性/行动领域列错位（31 行）：
CN 行动领域列误写英文关联性文本 → 用 EN 文件中同行的 B/D 代码回填；
EN 两列（climate_linkage / action_field）对调。
仅处理 EN col8 为 B/D 代码且 EN col9 为英文关联性文本的行，并断言 CN 侧内容匹配。
"""
import shutil
import re
import openpyxl

LINK_EN = {'Explicit mitigation', 'Mitigation-enabling', 'Economy planning incl. climate', 'Fossil-fuel'}
LINK_CN = {'明确减缓', '促进减缓', '化石能源', '经济规划（含气候）'}
CODE_RE = re.compile(r'^[A-Z](\.\d+)?(, ?[A-Z](\.\d+)?)*$')

shutil.copy('China_Climate_Target_Tracker_cn.xlsx', 'China_Climate_Target_Tracker_cn.pre_srcfix.bak.xlsx')
shutil.copy('China_Climate_Target_Tracker_en.xlsx', 'China_Climate_Target_Tracker_en.pre_srcfix.bak.xlsx')

wb_cn = openpyxl.load_workbook('China_Climate_Target_Tracker_cn.xlsx')
wb_en = openpyxl.load_workbook('China_Climate_Target_Tracker_en.xlsx')
ws_cn = wb_cn['来源']
ws_en = wb_en['Sources']

fixed, skipped = [], []
for r in range(9, ws_en.max_row + 1):
    iid = ws_en.cell(row=r, column=2).value
    if not iid:
        continue
    en8, en9 = ws_en.cell(row=r, column=8).value, ws_en.cell(row=r, column=9).value
    en8s = str(en8).strip() if en8 is not None else ''
    en9s = str(en9).strip() if en9 is not None else ''
    if CODE_RE.match(en8s) and en9s in LINK_EN:
        cn8 = ws_cn.cell(row=r, column=8).value
        cn9 = ws_cn.cell(row=r, column=9).value
        cn8s = str(cn8).strip() if cn8 is not None else ''
        cn9s = str(cn9).strip() if cn9 is not None else ''
        if cn8s not in LINK_CN or cn9s != en9s:
            skipped.append((r, iid, cn8s, cn9s, en8s, en9s))
            continue
        ws_cn.cell(row=r, column=9, value=en8s)
        ws_en.cell(row=r, column=8, value=en9s)
        ws_en.cell(row=r, column=9, value=en8s)
        fixed.append((r, iid, en8s))

wb_cn.save('China_Climate_Target_Tracker_cn.xlsx')
wb_en.save('China_Climate_Target_Tracker_en.xlsx')
print('fixed:', len(fixed))
for r, iid, code in fixed:
    print('  row', r, iid, '->', code)
if skipped:
    print('SKIPPED (mismatch):', skipped)
