# -*- coding: utf-8 -*-
"""Fix contradictory 方向+目标值 combinations in Targets_cn.xlsx (2026-09-08).

Rows where both columns carried a comparison word (达到+不低于, 降低至+低于,
超过+不低于, 低于+低于X, 降至…以下+低于X, 超过+X以上 …) are rewritten so the
comparison lives only in 方向, with 目标值 carrying the value expression,
per the original policy sentence (政策原文 column).

Usage:
  PYTHONIOENCODING=utf-8 python fix_direction2_cn.py --check   # dry run
  PYTHONIOENCODING=utf-8 python fix_direction2_cn.py --apply   # backup + fix + save
"""
import sys, shutil
import openpyxl

CN_FILE = 'Targets_cn.xlsx'

# sheet -> [(excel_row, old_dir, new_dir, old_val, new_val)]
# None = column unchanged (still asserted in check())
DIR2_FIXES = {
    '温室气体排放': [
        (103, '实现', '超过', '超过11亿吨二氧化碳当量', '11亿吨二氧化碳当量'),
    ],
    '能源|电力': [
        (65,  '降低', '下降', '超过18%', '18%以上'),
        (191, '超过', '达到', None, None),                    # 达到1500亿立方米以上
        (210, '降至…以下', '下降至', '低于40%', '40%以下'),
        (267, '降低至', '降至', '低于58%', '58%以下'),
        (268, '降至…以下', '下降到', '58%', '58%以下'),
        (269, '控制在…以内', '下降到', None, None),           # 下降到62%以内
        (272, '降低至', '降低到', '低于65%', '65%以下'),
        (274, '降低至', '下降到', '低于58%', '58%以下'),
        (307, None, None, '低于300克标准煤/千瓦时', '300克标准煤/千瓦时'),
        (338, '控制在…以内', '控制在…以下', None, None),      # 控制在50亿吨标准煤以下
        (428, None, None, '超过30%', '30%以上'),
        (439, '达到', '超过', '超过50%', '50%'),
        (455, None, None, '超过12亿千瓦', '12亿千瓦以上'),
        (456, '超过', '达到', None, None),                    # 达到12亿千瓦以上
        (459, '提高至', '达到', '超过12亿千瓦', '12亿千瓦以上'),
        (461, None, None, '12亿千瓦', '12亿千瓦以上'),
        (463, None, None, '超过12亿千瓦', '12亿千瓦以上'),
        (464, None, None, '超过12亿千瓦', '12亿千瓦以上'),
        (497, '达到', '不低于', '不低于50%', '50%'),
        (498, '达到', '不低于', '不低于40%（综合微能网不低于60%）', '40%（综合微能网不低于60%）'),
        (517, '达到', '不低于', '不低于10%', '10%'),
    ],
    '工业': [
        (12,  '降低', '小于', '低于93千克标准煤/吨', '93千克标准煤/吨'),
        (16,  '降低', '小于', '低于114千克标准煤/吨', '114千克标准煤/吨'),
        (34,  '达到', '不低于', '不低于80%', '80%'),
        (49,  '降低至', '降至', '低于0.4吨标准煤/万元', '0.4吨标准煤/万元以下'),
        (50,  '降低', '降至', '低于0.4吨标准煤/万元', '0.4吨标准煤/万元以下'),
        (76,  None, None, '1000万千瓦以上', '1000万千瓦'),
        (115, '达到', '不低于', '不低于10%', '10%'),
        (121, None, None, '超过60%', '60%'),
        (132, None, None, '超过30%', '30%以上'),
        (134, None, None, '超过30%', '30%以上'),
    ],
    '建筑': [
        (2,  None, None, '超过1亿平方米', '1亿平方米'),
        (65, '降低至', '降低到', '低于1.35', '1.35以下'),
        (68, '降至…以下', '降至', '13%', '13%以下'),
        (80, '达到', '超过', '超过15%', '15%'),
        (96, '达到', '不低于', '不低于80%', '80%'),
    ],
    '交通': [
        (6, '降低至', '降至', '低于4.5升/百公里', '4.5升/百公里以下'),
    ],
    '土地利用、土地利用变化和林业': [
        (30,  None, None, '4.5‰以下', '4.5‰'),
        (64,  '实现', '超过', '超过23%', '23%'),
        (103, '实现', '达到', None, None),                    # 达到超过165亿立方米
        (110, '提高至', '超过', '超过143亿立方米', '143亿立方米'),
        (147, None, None, '不低于18800公顷', '18800公顷'),
        (151, None, None, '超过55%', '55%以上'),
        (163, '超过', '不低于', '不低于35%', '35%'),
        (170, None, None, '超过90%', '90%以上'),
        (171, '达到', '不低于', '不低于18%', '18%'),
        (172, '达到', '不低于', '不低于18%', '18%'),
    ],
    '循环经济': [
        (60,  None, None, '超过90%', '90%以上'),
        (123, None, None, '超过3亿吨', '3亿吨以上'),
        (144, None, None, '6700亿立方米以内', '6700亿立方米'),
        (145, None, None, '6700亿立方米以内', '6700亿立方米'),
        (180, None, None, '86%以上', '86%'),
        (183, None, None, '88%以上', '88%'),
    ],
}

def check():
    wb = openpyxl.load_workbook(CN_FILE, read_only=True)
    n = 0
    for sheet, fixes in DIR2_FIXES.items():
        ws = wb[sheet]
        for row, old_dir, new_dir, old_val, new_val in fixes:
            metric = ws.cell(row=row, column=2).value
            if old_dir is not None:
                got = ws.cell(row=row, column=3).value
                if got != old_dir:
                    print(f'ASSERT FAIL {sheet} r{row} [{metric}]: dir got {got!r} want {old_dir!r}')
                    wb.close()
                    sys.exit(1)
            if old_val is not None:
                got = ws.cell(row=row, column=4).value
                if got != old_val:
                    print(f'ASSERT FAIL {sheet} r{row} [{metric}]: val got {got!r} want {old_val!r}')
                    wb.close()
                    sys.exit(1)
            n += 1
            d = f'{old_dir} -> {new_dir}' if old_dir is not None else 'dir ok'
            v = f'{old_val} -> {new_val}' if old_val is not None else 'val ok'
            print(f'OK {sheet} r{row} [{str(metric)[:26]}] {d} | {v}')
    wb.close()
    return n

def main():
    apply = '--apply' in sys.argv
    n = check()
    print(f'{n} direction/value fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    shutil.copy(CN_FILE, CN_FILE.replace('.xlsx', '.pre_dirfix2.bak'))
    wb = openpyxl.load_workbook(CN_FILE)
    for sheet, fixes in DIR2_FIXES.items():
        ws = wb[sheet]
        for row, old_dir, new_dir, old_val, new_val in fixes:
            if new_dir is not None:
                ws.cell(row=row, column=3).value = new_dir
            if new_val is not None:
                ws.cell(row=row, column=4).value = new_val
    wb.save(CN_FILE)
    wb.close()
    print(f'{CN_FILE}: applied and saved')

if __name__ == '__main__':
    main()
