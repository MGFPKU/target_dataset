# -*- coding: utf-8 -*-
"""Fix EN->CN translation artifacts in the 方向 (direction) column of Targets_cn.xlsx.

Every fix replaces a translated phrase (e.g. 保持在…以下 = 'keep below') with the
verb used in the original Chinese policy sentence (政策原文 column), e.g. 低于 / 控制在…以内.

Usage:
  PYTHONIOENCODING=utf-8 python fix_direction_cn.py --check   # dry run
  PYTHONIOENCODING=utf-8 python fix_direction_cn.py --apply   # backup + fix + save
"""
import sys, shutil
import openpyxl

CN_FILE = 'Targets_cn.xlsx'

# sheet -> [(excel_row, old_direction, new_direction)]
DIR_FIXES = {
    '能源|电力': [
        (24, '保持在…以下', '低于'),      # 弃光率低于5%
        (26, '保持在…以下', '低于'),      # 弃光率低于5%
        (27, '保持在…以下', '低于'),      # 弃风率低于12%
        (28, '保持在…以下', '低于'),      # 弃风率低于10%
        (29, '保持在…以下', '控制在…以内'),  # 力争控制在5%左右
        (33, '保持在…以下', '低于'),      # 弃光率低于5%
        (100, '保持在…以下', '控制在…以内'), # 原油一次加工能力控制在10亿吨以内
        (105, '保持在', '稳定在'),        # 原油生产能力稳定在2亿吨左右
        (42, '选择', '遴选'),            # 遴选200家公共机构能效领跑者
        (245, '取消或推迟', '取消和推迟'),  # 取消和推迟煤电建设项目1.5亿千瓦以上
        (280, '暂停或放缓', '停建和缓建'),  # 停建和缓建煤电产能1.5亿千瓦
        (269, '降至…以下', '控制在…以内'), # 煤炭比重下降到62%以内
        (307, '降至…以下', '低于'),      # 改造后平均供电煤耗低于300克/千瓦时
        (308, '降至…以下', '低于'),      # 改造后平均供电煤耗低于310克/千瓦时
        (339, '约束在…以内', '控制在…以内'), # 能源消费总量控制在50亿吨标准煤以内
        (377, '容纳', '具备'),           # 具备5亿千瓦左右分布式新能源接入能力
        (378, '利用', '充分发挥'),        # 充分发挥…热电联产电厂的供热能力
        (398, '以…为基准', '达到'),      # 非化石能源占一次能源消费比重达到15%左右
        (433, '翻番', '倍增'),           # 实施非化石能源十年倍增行动
        (461, '致力于', '达到'),         # 风电和太阳能发电总装机容量达到12亿千瓦以上
        (5, '全面淘汰', '基本淘汰'),      # 基本淘汰10蒸吨/小时及以下燃煤锅炉
        (228, '全面淘汰', '基本淘汰'),    # 基本淘汰每小时35蒸吨以下燃煤锅炉
        (102, '逐年建设', '新增'),       # 新增原油一次管输能力1.2亿吨/年
        (455, '提升至', '达到'),         # 风电、太阳能发电总装机容量将达到12亿千瓦以上
        (362, '逐步淘汰', '淘汰'),       # 淘汰落后锅炉20万蒸吨
        (484, '保持稳定', '稳定在'),      # 原油年产量稳定在2亿吨水平
    ],
    '工业': [
        (154, '制定', '完成'),           # 到2030年完成500项以上
        (176, '选择', '遴选'),           # 遴选50家水效“领跑者”企业
        (153, '逐步淘汰', '淘汰'),        # 2014年淘汰5万台小锅炉
        (73, '发展和推广', '开发推广'),    # 力争开发推广万种绿色产品
    ],
    '建筑': [
        (2, '覆盖', '超过'),             # 地热能建筑应用面积将超过1亿平方米
        (64, '保持低于', '不超过'),       # 数据中心电能利用效率普遍不超过1.5
        (71, '占', '超过'),              # 建筑能耗中电力消费占比将超过百分之五十五
    ],
    '交通': [
        (20, '服务', '形成'),            # 形成50万辆电动汽车充电基础设施体系
        (50, '超越', '不低于'),          # 新能源汽车占比不低于30%
    ],
    '土地利用、土地利用变化和林业': [
        (12, '覆盖', '规划建设'),         # 规划建设林木种质资源保存库200处、9690公顷
        (15, '再生', '更新'),            # 规划人工更新300万公顷
        (28, '保持在…以下', '控制在…以下'), # 林业有害生物成灾率控制在4.5‰以下
        (83, '降至…以下', '控制在…以下'),  # 森林火灾受害率稳定控制在千分之一以下
        (182, '升级', '改造提升'),        # 改造提升20000公顷城市公园绿地
        (17, '保持', '规划'),            # 规划人工种草保留面积500万公顷
        (18, '保持', '达到'),            # 规划人工种草保留面积累计达到1000万公顷
        (80, '培育和管理', '规划'),       # 规划森林抚育经营7500万公顷
        (81, '培育和管理', '规划'),       # 规划森林抚育经营3500万公顷
        (123, '保持稳定', '稳定在'),      # 林业自然保护区面积占国土面积比例稳定在13%左右
        (144, '管理或修复', '治理或修复'),  # 可治理的沙化土地得到治理或修复
        (147, '创建和恢复', '营造和修复'),  # 营造和修复红树林面积18800公顷
    ],
    '循环经济': [
        (19, '获得', '力争达到'),         # 力争再生铜产量达到400万吨
        (73, '回收利用', '力争达到'),      # 力争废有色金属回收利用量达到2000万吨
        (125, '回收利用', '力争达到'),     # 力争废钢回收利用量达到3.2亿吨
        (142, '保持在…以下', '控制在…以内'), # 全国用水总量控制在6400亿立方米以内
        (41, '保持', '稳定在'),          # 全国秸秆综合利用率稳定在88%以上
        (180, '保持', '保持在…以上'),     # 秸秆综合利用率保持在86%以上
    ],
}

def check():
    wb = openpyxl.load_workbook(CN_FILE, read_only=True)
    n = 0
    for sheet, fixes in DIR_FIXES.items():
        ws = wb[sheet]
        for row, old, new in fixes:
            got = ws.cell(row=row, column=3).value
            metric = ws.cell(row=row, column=2).value
            if got != old:
                print(f'ASSERT FAIL {sheet} r{row} [{str(metric)[:30]}]: got {got!r} want old {old!r}')
                wb.close()
                sys.exit(1)
            n += 1
            print(f'OK {sheet} r{row} [{str(metric)[:28]}] {old} -> {new}')
    wb.close()
    return n

def main():
    apply = '--apply' in sys.argv
    n = check()
    print(f'{n} direction fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    shutil.copy(CN_FILE, CN_FILE.replace('.xlsx', '.pre_dirfix.bak'))
    wb = openpyxl.load_workbook(CN_FILE)
    for sheet, fixes in DIR_FIXES.items():
        ws = wb[sheet]
        for row, old, new in fixes:
            ws.cell(row=row, column=3).value = new
    wb.save(CN_FILE)
    wb.close()
    print(f'{CN_FILE}: applied and saved')

if __name__ == '__main__':
    main()
