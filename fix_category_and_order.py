# -*- coding: utf-8 -*-
"""Fix A-list target-category errors (CN+EN) and restore EN Energy|Power row order to match CN.

Usage:
  PYTHONIOENCODING=utf-8 python fix_category_and_order.py --check   # dry run, prints matching stats
  PYTHONIOENCODING=utf-8 python fix_category_and_order.py --apply   # reorder + fix + save (creates .bak first)
"""
import sys, re, shutil
import openpyxl

CN_FILE = 'Targets_cn.xlsx'
EN_FILE = 'Targets_en.xlsx'
CN_EP = '能源|电力'
EN_EP = 'Energy|Power'

# ---------- A-list fixes: sheet -> [(excel_row, new_category)] ----------
CN_FIXES = {
    '温室气体排放': [(75,'碳强度'),(100,'碳强度'),(98,'碳强度'),(99,'碳强度'),(125,'碳强度'),(126,'碳强度')],
    '能源|电力': [(16,'能源效率'),(18,'能源效率'),(232,'能源效率'),(189,'能源效率'),
                (124,'涉煤'),(215,'涉煤'),(216,'涉煤'),
                (294,'资源效率'),(295,'资源效率'),(296,'资源效率'),
                (214,'过渡能源'),(469,'能源效率'),(484,'过渡能源')],
    '工业': [(26,'能源效率'),(27,'能源效率'),(28,'能源效率'),(29,'能源效率'),
           (111,'能源效率'),(113,'工业转型'),(114,'工业转型'),
           (131,'能源效率'),(141,'能源效率'),(144,'能源效率'),
           (152,'资源效率'),(165,'循环利用'),(166,'循环利用')],
    '建筑': [(10,'碳强度')],
    '交通': [(46,'能源效率'),(47,'能源效率'),(64,'能源效率'),(65,'能源效率'),(66,'电气化')],
    '土地利用、土地利用变化和林业': [(16,'碳汇'),(123,'环境质量'),(171,'环境质量'),(172,'环境质量'),
                              (166,'污染控制'),(173,'环境质量')],
    '循环经济': [(70,'资源效率'),(79,'资源效率'),(88,'循环利用'),(90,'循环利用'),(96,'循环利用'),
              (97,'循环利用'),(148,'污染控制'),(157,'资源效率'),(163,'资源效率'),(158,'环境质量')],
}
EN_FIXES = {  # same rows, English category names; Energy|Power applied after reorder
    'GHG Emissions': [(75,'Carbon intensity'),(100,'Carbon intensity'),(98,'Carbon intensity'),
                      (99,'Carbon intensity'),(125,'Carbon intensity'),(126,'Carbon intensity')],
    'Energy|Power': [(16,'Energy efficiency'),(18,'Energy efficiency'),(232,'Energy efficiency'),
                     (189,'Energy efficiency'),(124,'Coal-related'),(215,'Coal-related'),(216,'Coal-related'),
                     (294,'Resource efficiency'),(295,'Resource efficiency'),(296,'Resource efficiency'),
                     (214,'Transitional energy'),(469,'Energy efficiency'),(484,'Transitional energy')],
    'Industry': [(26,'Energy efficiency'),(27,'Energy efficiency'),(28,'Energy efficiency'),(29,'Energy efficiency'),
                 (111,'Energy efficiency'),(113,'Industrial transformation'),(114,'Industrial transformation'),
                 (131,'Energy efficiency'),(141,'Energy efficiency'),(144,'Energy efficiency'),
                 (152,'Resource efficiency'),(165,'Recycling'),(166,'Recycling')],
    'Buildings': [(10,'Carbon intensity')],
    'Transport': [(46,'Energy efficiency'),(47,'Energy efficiency'),(64,'Energy efficiency'),
                  (65,'Energy efficiency'),(66,'Electrification')],
    'LULUCF': [(16,'Carbon sink'),(123,'Environmental quality'),(171,'Environmental quality'),
               (172,'Environmental quality'),(166,'Pollution control'),(173,'Environmental quality')],
    'Circular Economy': [(70,'Resource efficiency'),(79,'Resource efficiency'),(88,'Recycling'),
                         (90,'Recycling'),(96,'Recycling'),(97,'Recycling'),
                         (148,'Pollution control'),(157,'Resource efficiency'),(163,'Resource efficiency'),
                         (158,'Environmental quality')],
}

# ---------- language-independent row key helpers ----------
# manually paired rows (cn_data_idx: en_data_idx) where rowkey+col3 matching is
# ambiguous (ties or col3 phrasing differs, e.g. 基本完成 vs nearly 100%);
# every pair verified against policy-original sentences
OVERRIDES = {
    3: 41, 378: 87,
    19: 301, 20: 302, 21: 303, 22: 308, 27: 291, 30: 294, 32: 297, 33: 298, 34: 299,
    42: 471,
    59: 119, 81: 120,
    106: 461,
    157: 204, 171: 46,
    242: 80,
    244: 81, 245: 83,
    255: 165,
    257: 38,
    274: 508, 327: 536,
    295: 552,
    389: 395, 429: 396,
    397: 405, 398: 406, 400: 408, 402: 410,
    411: 420,
    442: 475,
    444: 477,
    445: 478,
    470: 500, 471: 494,
    489: 535,
    505: 288,
    507: 332, 512: 328,
    514: 232,
    550: 311, 555: 1,
}

def norm5(v):
    if isinstance(v, (int, float)):
        return ('n', int(v))
    m = {'十一五':'11th FYP','十二五':'12th FYP','十三五':'13th FYP','十四五':'14th FYP','十五五':'15th FYP'}
    return ('s', m.get(str(v).strip(), str(v).strip()))

def norm6(v):
    m = {'新目标':'n','重申目标':'r','强化或放宽目标（积极气候影响）':'s/l-POS',
         '强化或放宽目标（消极气候影响）':'s/l-NEG','强化或放宽目标（气候影响不明确）':'s/l-UNC'}
    return m.get(str(v).strip(), str(v).strip())

CN8 = {'全社会经济':'economy','行业':'sectors','不明确':'unclear','地区':'areas','城市与地区':'areas',
       '城市':'areas','企事业单位':'enterprises','企业':'enterprises','港口':'ports'}
EN8 = {'whole economy':'economy','sectors':'sectors','unclear':'unclear','areas':'areas',
       'enterprises':'enterprises','ports':'ports','cities':'areas'}

def norm8(v, is_cn):
    s = str(v).strip() if v else ''
    return (CN8 if is_cn else EN8).get(s, s)

def rowkey(r, is_cn):
    return (str(r[10]).strip(), r[0], norm5(r[5]), norm6(r[6]), norm8(r[8], is_cn))

# ---------- col3 (target value) canonicalizer ----------
NUM_RE = re.compile(r'(\d[\d,]*(?:\.\d+)?)')

# ordered longest-first unit tables: (pattern, class, factor)
CN_UNITS = [
    ('吨标准煤','tce',1),('标准煤','tce',1),('吨标煤','tce',1),('标煤','tce',1),
    ('千瓦时','kwh',1),('亿千瓦时','kwh',1e8),('千瓦','w',1e3),('兆瓦','w',1e6),
    ('吉瓦','w',1e9),('瓦','w',1),
    ('亿立方米','m3',1e8),('立方米','m3',1),
    ('平方公里','km2',1),('平方千米','km2',1),('公顷','km2',1e-2),('亩','km2',6.6667e-4),
    ('公里','km',1),('千米','km',1),
    ('吨','ton',1),
    ('蒸吨','ton',1),
    ('千克标准油','kg',1),('千克标煤','kg',1),('千克','kg',1),
    ('%','pct',1),('％','pct',1),('个百分点','pct',1),
    ('度','kwh',1),
    ('小时','h',1),('天','d',1),('日','d',1),('年','y',1),
    ('台','cnt',1),('座','cnt',1),('个','cnt',1),('家','cnt',1),('条','cnt',1),
    ('套','cnt',1),('处','cnt',1),('项','cnt',1),('艘','cnt',1),('辆','cnt',1),
    ('元','cny',1),
]
EN_UNITS = [
    ('kilowatt-hours','kwh',1),('kWh','kwh',1),('kwh','kwh',1),('kW','w',1e3),('MW','w',1e6),('GW','w',1e9),
    ('kilowatts','w',1e3),('megawatts','w',1e6),('gigawatts','w',1e9),('watts','w',1),
    ('tons of standard coal','tce',1),('tonnes of standard coal','tce',1),
    ('tons of coal equivalent','tce',1),('tonnes of coal equivalent','tce',1),
    ('standard coal','tce',1),('tce','tce',1),
    ('cubic meters','m3',1),('m3','m3',1),
    ('square kilometers','km2',1),('km2','km2',1),('hectares','km2',1e-2),('ha','km2',1e-2),
    ('mu','km2',6.6667e-4),
    ('metric tons','ton',1),('tons','ton',1),('tonnes','ton',1),('ton','ton',1),
    ('kg','kg',1),('kilograms','kg',1),
    ('kilometers','km',1),('km','km',1),
    ('%','pct',1),('percent','pct',1),('percentage points','pct',1),
    ('hours','h',1),('days','d',1),('years','y',1),
    ('sets','cnt',1),('units','cnt',1),('enterprises','cnt',1),('projects','cnt',1),
    ('yuan','cny',1),('RMB','cny',1),('CNY','cny',1),
]
CN_MULT = [('亿',1e8),('万',1e4)]
EN_MULT = [('trillion',1e12),('billion',1e9),('million',1e6),('thousand',1e3),('hundred',1e2)]
CLASS_ALIAS = {'tce':'ton'}  # coal-equivalent tons compare numerically with plain tons

def canon_col3(s, is_cn):
    if s is None:
        return frozenset()
    t = str(s)
    units = CN_UNITS if is_cn else EN_UNITS
    mults = CN_MULT if is_cn else EN_MULT
    out = set()
    for m in NUM_RE.finditer(t.replace(',', '')):
        val = float(m.group(1))
        rest = t[m.end():].lstrip()
        while True:
            hit = False
            for mp, mf in mults:
                if rest.startswith(mp):
                    val *= mf
                    rest = rest[len(mp):].lstrip()
                    hit = True
                    break
            if not hit:
                break
        cls, factor = 'cnt', 1.0
        for up, uc, uf in units:
            if rest.startswith(up):
                cls, factor = uc, uf
                break
        if 1900 <= val < 2200:
            cls = 'y'  # bare 4-digit numbers are years in both languages
        cls = CLASS_ALIAS.get(cls, cls)
        out.add((cls, round(val * factor, 3)))
    return frozenset(out)

def canon_row(r, is_cn):
    return (rowkey(r), canon_col3(r[3], is_cn))

# ---------- main ----------
def main():
    apply = '--apply' in sys.argv
    cn_wb = openpyxl.load_workbook(CN_FILE, read_only=True)
    cnr = list(cn_wb[CN_EP].iter_rows(min_row=2, values_only=True))
    en_wb_r = openpyxl.load_workbook(EN_FILE, read_only=True)
    enr = list(en_wb_r[EN_EP].iter_rows(min_row=2, values_only=True))
    cn_wb.close(); en_wb_r.close()

    # group EN rows by rowkey
    from collections import defaultdict
    eng = defaultdict(list)
    for j, r in enumerate(enr):
        eng[rowkey(r, False)].append(j)

    # for each CN row i, find matching EN row(s) by rowkey, then by col3 canon
    order = [None] * len(cnr)   # order[i] = EN data-row index for CN row i
    problems = []
    used = set()
    for i, j in OVERRIDES.items():
        assert rowkey(cnr[i], True) == rowkey(enr[j], False), f'override rowkey mismatch CN[{i}] EN[{j}]'
        assert j not in used
        order[i] = j; used.add(j)
    for i, r in enumerate(cnr):
        if order[i] is not None:
            continue
        cands = eng[rowkey(r, True)]
        cc = canon_col3(r[3], True)
        exact = [j for j in cands if canon_col3(enr[j][3], False) == cc and j not in used]
        if len(exact) == 1:
            order[i] = exact[0]; used.add(exact[0])
        else:
            problems.append((i, str(r[1]), str(r[3]), str(r[10]), cc,
                             [(j, str(enr[j][1]), str(enr[j][3]), canon_col3(enr[j][3], False))
                              for j in cands if j not in used]))

    unmatched = sum(1 for x in order if x is None)
    print(f'CN rows: {len(cnr)}, EN rows: {len(enr)}, matched: {len(cnr)-unmatched}, unresolved: {unmatched}')
    for p in problems:
        print('UNRESOLVED CN', p[0], '|', p[1], '|', p[2], '|', p[3], '|', p[4])
        for j, m, v, cj in p[5]:
            print('    cand EN', j, '|', m, '|', v, '|', cj)
    if unmatched:
        print('ABORT: unresolved rows remain')
        return

    assert sorted(order) == list(range(len(enr))), 'permutation incomplete'

    # verify full correspondence (canon check skipped for override rows: col3 phrasing differs)
    bad = 0
    for i, j in enumerate(order):
        if rowkey(cnr[i], True) != rowkey(enr[j], False):
            bad += 1
            print('MISMATCH at', i, cnr[i][1], '|', enr[j][1])
        elif i not in OVERRIDES and canon_col3(cnr[i][3], True) != canon_col3(enr[j][3], False):
            bad += 1
            print('MISMATCH at', i, cnr[i][1], '|', enr[j][1])
    print('post-match mismatches:', bad)
    if bad:
        return

    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return

    # ---- backups ----
    shutil.copy(CN_FILE, CN_FILE.replace('.xlsx', '.pre_catfix.bak'))
    shutil.copy(EN_FILE, EN_FILE.replace('.xlsx', '.pre_catfix.bak'))

    # ---- apply A-fixes to CN ----
    cn_wb = openpyxl.load_workbook(CN_FILE)
    for sheet, fixes in CN_FIXES.items():
        ws = cn_wb[sheet]
        for row, cat in fixes:
            old = ws.cell(row=row, column=8).value
            assert old == fixes_old_cn(sheet, row, cat), f'CN {sheet} row {row}: expected old cat, got {old}'
            ws.cell(row=row, column=8).value = cat
    cn_wb.save(CN_FILE)
    cn_wb.close()
    print('CN fixes applied')

    # ---- reorder EN Energy|Power + apply fixes ----
    en_wb = openpyxl.load_workbook(EN_FILE)
    ws = en_wb[EN_EP]
    n = len(enr)
    # pos[j] = current data position (0-based) of EN row j; rows currently in original order
    pos = list(range(n))
    def get_row_vals(p):
        row_idx = p + 2
        return [ws.cell(row=row_idx, column=c).value for c in range(1, 13)]
    def set_row_vals(p, vals):
        row_idx = p + 2
        for c, v in enumerate(vals, start=1):
            ws.cell(row=row_idx, column=c).value = v
    for p in range(n):
        desired = order[p]
        if pos[desired] == p:
            continue
        q = pos[desired]
        vals_p, vals_q = get_row_vals(p), get_row_vals(q)
        set_row_vals(p, vals_q); set_row_vals(q, vals_p)
        # update pos for the swapped rows
        other = next(j for j in range(n) if pos[j] == p)
        pos[other], pos[desired] = q, p
    print('EN Energy|Power reordered')

    # verify reordered values match CN keys (canon check skipped for override rows)
    bad2 = 0
    for i in range(n):
        vals = [ws.cell(row=i+2, column=c).value for c in range(1, 13)]
        if rowkey(vals, False) != rowkey(cnr[i], True):
            bad2 += 1
            print('POST-REORDER MISMATCH', i)
        elif i not in OVERRIDES and canon_col3(vals[3], False) != canon_col3(cnr[i][3], True):
            bad2 += 1
            print('POST-REORDER MISMATCH', i)
    print('post-reorder mismatches:', bad2)

    for sheet, fixes in EN_FIXES.items():
        w = en_wb[sheet]
        for row, cat in fixes:
            w.cell(row=row, column=8).value = cat
    en_wb.save(EN_FILE)
    en_wb.close()
    print('EN fixes applied and saved')

def fixes_old_cn(sheet, row, cat):
    # expected current (wrong) categories, for sanity assertion
    d = {
        '温室气体排放': {(75,'碳强度'):'碳减排',(100,'碳强度'):'碳减排',(98,'碳强度'):'碳减排',(99,'碳强度'):'碳减排',(125,'碳强度'):'碳减排',(126,'碳强度'):'碳减排'},
        '能源|电力': {(16,'能源效率'):'过渡能源',(18,'能源效率'):'过渡能源',(232,'能源效率'):'过渡能源',(189,'能源效率'):'可再生能源',
                   (124,'涉煤'):'过渡能源',(215,'涉煤'):'过渡能源',(216,'涉煤'):'过渡能源',
                   (294,'资源效率'):'过渡能源',(295,'资源效率'):'过渡能源',(296,'资源效率'):'过渡能源',
                   (214,'过渡能源'):'涉煤',(469,'能源效率'):'可再生能源',(484,'过渡能源'):'生产导向'},
        '工业': {(26,'能源效率'):'资源效率',(27,'能源效率'):'资源效率',(28,'能源效率'):'资源效率',(29,'能源效率'):'资源效率',
               (111,'能源效率'):'工业转型',(113,'工业转型'):'生产导向',(114,'工业转型'):'生产导向',
               (131,'能源效率'):'工业转型',(141,'能源效率'):'工业转型',(144,'能源效率'):'工业转型',
               (152,'资源效率'):'工业转型',(165,'循环利用'):'工业转型',(166,'循环利用'):'工业转型'},
        '建筑': {(10,'碳强度'):'碳减排'},
        '交通': {(46,'能源效率'):'技术创新',(47,'能源效率'):'技术创新',(64,'能源效率'):'技术创新',
              (65,'能源效率'):'技术创新',(66,'电气化'):'可再生能源'},
        '土地利用、土地利用变化和林业': {(16,'碳汇'):'环境质量',(123,'环境质量'):'碳汇',(171,'环境质量'):'碳汇',
                                  (172,'环境质量'):'碳汇',(166,'污染控制'):'环境质量',(173,'环境质量'):'资源效率'},
        '循环经济': {(70,'资源效率'):'污染控制',(79,'资源效率'):'循环利用',(88,'循环利用'):'污染控制',
                  (90,'循环利用'):'污染控制',(96,'循环利用'):'污染控制',(97,'循环利用'):'污染控制',
                  (148,'污染控制'):'循环利用',(157,'资源效率'):'循环利用',(163,'资源效率'):'循环利用',
                  (158,'环境质量'):'循环利用'},
    }
    return d.get(sheet, {}).get((row, cat))

if __name__ == '__main__':
    main()
