# -*- coding: utf-8 -*-
"""2026-11-10 AP2412 补录：《水泥行业节能降碳专项行动计划》（发改环资〔2024〕733号，
2024-05-27，国家发展改革委等五部门）5 条目标写入 CN/EN 数据集，来源表 AP2412
目标标记 无目标→含目标、覆盖时期→2024-2025（该来源此前已入库 2 条目标）。

CN: 工业 +3、交通 +1、循环经济 +1、说明最后更新=2026-11-10
EN: Industry +3、Transport +1、Circular Economy +1、README Last update=2026-11-10
"""
import shutil
import datetime
import openpyxl

S1_CN = ('2024—2025年，通过实施水泥行业节能降碳改造和用能设备更新形成节能量约500万吨标准煤、'
         '减排二氧化碳约1300万吨。')
S2_CN = ('到2025年底，水泥窑使用替代燃料技术生产线比例达到30%，'
         '水泥行业替代燃料消费比例力争达到10%。')
S4_CN = '到2025年底，水泥行业清洁运输比例达到50%。'
S5_CN = '到2025年底，水泥行业综合利用废弃物总量达到8亿吨。'

S1_EN = ('During 2024–2025, the implementation of energy-saving and carbon-reduction retrofits, together with '
         'energy-consuming equipment upgrades in the cement industry, is expected to generate energy savings of '
         'approximately 5 million tons of standard coal and reduce carbon dioxide emissions by approximately '
         '13 million tons.')
S2_EN = ('By the end of 2025, the share of cement kiln production lines using alternative fuel technologies will '
         'reach 30%, and the share of alternative fuel consumption in the cement industry will strive to reach 10%.')
S4_EN = 'By the end of 2025, the share of clean transportation in the cement industry will reach 50%.'
S5_EN = ('By the end of 2025, the total amount of waste comprehensively utilized by the cement industry will reach '
         '800 million tonnes.')

NEW_DATE = datetime.datetime(2026, 11, 10, 0, 0)

# ---------- CN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_cn.xlsx')
shutil.copy('China_Climate_Target_Tracker_cn.xlsx', 'China_Climate_Target_Tracker_cn.pre_cement2412.bak.xlsx')

ws = wb['工业']
ws.append((2024, '水泥行业节能降碳改造', '形成', '约500万吨标准煤', None, '2024-2025',
           '新目标', '节能', '行业', S1_CN, 'AP2412.pdf', '工业'))
ws.append((2024, '水泥窑使用替代燃料技术生产线比例', '达到', '30%', None, 2025,
           '新目标', '非化石能源', '行业', S2_CN, 'AP2412.pdf', '工业'))
ws.append((2024, '水泥行业替代燃料消费比例', '力争达到', '10%', None, 2025,
           '新目标', '非化石能源', '行业', S2_CN, 'AP2412.pdf', '工业'))

ws = wb['交通']
ws.append((2024, '水泥行业清洁运输比例', '达到', '50%', None, 2025,
           '新目标', '绿色交通', '行业', S4_CN, 'AP2412.pdf', '交通'))

ws = wb['循环经济']
ws.append((2024, '水泥行业综合利用废弃物总量', '达到', '8亿吨', None, 2025,
           '新目标', '循环利用', '行业', S5_CN, 'AP2412.pdf', '循环经济'))

ws = wb['来源']
for r in range(9, ws.max_row + 1):
    if ws.cell(row=r, column=2).value == 'AP2412':
        ws.cell(row=r, column=10, value='2024-2025')
        ws.cell(row=r, column=11, value='含目标')
        break

wb['说明']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_cn.xlsx')
print('CN done: 工业 rows =', wb['工业'].max_row, '| 交通 =', wb['交通'].max_row,
      '| 循环经济 =', wb['循环经济'].max_row)

# ---------- EN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_en.xlsx')
shutil.copy('China_Climate_Target_Tracker_en.xlsx', 'China_Climate_Target_Tracker_en.pre_cement2412.bak.xlsx')

ws = wb['Industry']
ws.append((2024, 'energy-saving and carbon-reduction transformation of the cement industry', 'save',
           'about 5 million tons of standard coal', None, '2024-2025',
           'n', 'Energy saving', 'sectors', S1_EN, 'AP2412.pdf', 'Industry'))
ws.append((2024, 'share of cement kiln production lines using alternative fuel technologies', 'reach',
           '30%', None, 2025,
           'n', 'Non-fossil energy', 'sectors', S2_EN, 'AP2412.pdf', 'Industry'))
ws.append((2024, 'share of alternative fuel consumption in the cement industry', 'strive to reach',
           '10%', None, 2025,
           'n', 'Non-fossil energy', 'sectors', S2_EN, 'AP2412.pdf', 'Industry'))

ws = wb['Transport']
ws.append((2024, 'share of clean transportation in the cement industry', 'reach',
           '50%', None, 2025,
           'n', 'Green transportation', 'sectors', S4_EN, 'AP2412.pdf', 'Transport'))

ws = wb['Circular Economy']
ws.append((2024, 'total amount of waste comprehensively utilized by the cement industry', 'reach',
           '800 million tonnes', None, 2025,
           'n', 'Recycling', 'sectors', S5_EN, 'AP2412.pdf', 'Circular Economy'))

ws = wb['Sources']
for r in range(9, ws.max_row + 1):
    if ws.cell(row=r, column=2).value == 'AP2412':
        ws.cell(row=r, column=10, value='2024-2025')
        ws.cell(row=r, column=11, value='TAR')
        break

wb['README']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_en.xlsx')
print('EN done: Industry rows =', wb['Industry'].max_row, '| Transport =', wb['Transport'].max_row,
      '| Circular Economy =', wb['Circular Economy'].max_row)
