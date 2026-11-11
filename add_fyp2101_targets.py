# -*- coding: utf-8 -*-
"""2026-11-11 FYP2101 补录：《"十四五"工业绿色发展规划》（工信部规〔2021〕178号，
2021-11-15 印发）主要目标段 6 条缺失目标 + 2 条重申目标写入 CN/EN 数据集，
并小修：有色金属窑炉清洁生产改造目标值补"左右"后缀（CN 4000台左右 / EN about 4000 kilns，
政策原文"完成4000台左右"）。

CN: 温室气体排放 +1、污染 +1、循环经济 +3、工业 +3、说明最后更新=2026-11-11
EN: GHG Emissions +1、Pollution +1、Circular Economy +3、Industry +3、README Last update=2026-11-11
"""
import shutil
import datetime
import openpyxl

NEW_DATE = datetime.datetime(2026, 11, 11, 0, 0)

# ---------- CN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_cn.xlsx')
shutil.copy('China_Climate_Target_Tracker_cn.xlsx', 'China_Climate_Target_Tracker_cn.pre_fyp2101add.bak.xlsx')

ws = wb['温室气体排放']
ws.append((2021, '单位工业增加值二氧化碳排放量', '下降', '18%', 2020, '2025',
           '新目标', '碳强度', '行业',
           '到2025年，单位工业增加值二氧化碳排放降低18%。',
           'FYP2101.pdf', '温室气体排放'))

ws = wb['污染']
ws.append((2021, '重点行业主要污染物排放强度', '下降', '10%', 2020, '2025',
           '新目标', '污染控制', '行业',
           '到2025年，重点行业主要污染物排放强度降低10%。',
           'FYP2101.pdf', '污染'))

ws = wb['循环经济']
ws.append((2021, '大宗工业固废综合利用率', '达到', '57%', None, '2025',
           '新目标', '循环利用', '行业',
           '到2025年，大宗工业固废综合利用率达到57%。',
           'FYP2101.pdf', '循环经济'))
ws.append((2021, '主要再生资源回收利用量', '达到', '4.8亿吨', None, '2025',
           '新目标', '循环利用', '全社会经济',
           '到2025年，主要再生资源回收利用量达到4.8亿吨。',
           'FYP2101.pdf', '循环经济'))
ws.append((2021, '工业资源综合利用基地', '建设', '50个', None, '2025',
           '新目标', '循环利用', '行业',
           '到2025年，建设50个工业资源综合利用基地。',
           'FYP2101.pdf', '循环经济'))

ws = wb['工业']
ws.append((2021, '绿色环保产业产值', '达到', '11万亿元', None, '2025',
           '新目标', '生产导向', '行业',
           '到2025年，绿色环保产业产值达到11万亿元。',
           'FYP2101.pdf', '工业'))
ws.append((2021, '规模以上工业企业单位增加值能耗', '下降', '13.5%', 2020, '2025',
           '重申目标', '能源效率', '企业',
           '到2025年，规模以上工业单位增加值能耗降低13.5%。',
           'FYP2101.pdf', '工业'))
ws.append((2021, '万元工业增加值用水量', '下降', '16%', 2020, '2025',
           '重申目标', '资源效率', '行业',
           '到2025年，单位工业增加值用水量降低16%。',
           'FYP2101.pdf', '工业'))

# 小修：有色金属窑炉 4000台 -> 4000台左右
for r in range(2, wb['工业'].max_row + 1):
    if wb['工业'].cell(row=r, column=2).value == '有色金属窑炉清洁生产改造' \
            and wb['工业'].cell(row=r, column=11).value == 'FYP2101.pdf':
        wb['工业'].cell(row=r, column=4, value='4000台左右')
        break

wb['说明']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_cn.xlsx')
print('CN done: 温室气体排放 =', wb['温室气体排放'].max_row, '| 污染 =', wb['污染'].max_row,
      '| 循环经济 =', wb['循环经济'].max_row, '| 工业 =', wb['工业'].max_row)

# ---------- EN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_en.xlsx')
shutil.copy('China_Climate_Target_Tracker_en.xlsx', 'China_Climate_Target_Tracker_en.pre_fyp2101add.bak.xlsx')

ws = wb['GHG Emissions']
ws.append((2021, 'carbon dioxide emissions per unit of industrial value added', 'reduce by',
           '18%', 2020, '2025',
           'n', 'Carbon intensity', 'sectors',
           'By 2025, carbon dioxide emissions per unit of industrial value added will be reduced by 18%.',
           'FYP2101.pdf', 'Emissions'))

ws = wb['Pollution']
ws.append((2021, 'emission intensity of major pollutants in key industries', 'reduce by',
           '10%', 2020, '2025',
           'n', 'Pollution control', 'sectors',
           'By 2025, the emission intensity of major pollutants in key industries will be reduced by 10%.',
           'FYP2101.pdf', 'Pollution'))

ws = wb['Circular Economy']
ws.append((2021, 'utilization rate of bulk industrial solid waste', 'reach',
           '57%', None, '2025',
           'n', 'Recycling', 'sectors',
           'By 2025, the comprehensive utilization rate of bulk industrial solid waste will reach 57%.',
           'FYP2101.pdf', 'Circular Economy'))
ws.append((2021, 'recycling volume of major renewable resources', 'reach',
           '480 million tons', None, '2025',
           'n', 'Recycling', 'whole economy',
           'By 2025, the recycling volume of major renewable resources will reach 480 million tons.',
           'FYP2101.pdf', 'Circular Economy'))
ws.append((2021, 'industrial resource comprehensive utilization bases', 'build',
           '50', None, '2025',
           'n', 'Recycling', 'sectors',
           'By 2025, 50 industrial resource comprehensive utilization bases will be built.',
           'FYP2101.pdf', 'Circular Economy'))

ws = wb['Industry']
ws.append((2021, 'output value of green environmental protection industry', 'reach',
           '11 trillion yuan', None, '2025',
           'n', 'Production-oriented', 'sectors',
           'By 2025, the output value of the green environmental protection industry will reach 11 trillion yuan.',
           'FYP2101.pdf', 'Industry'))
ws.append((2021, 'energy consumption per unit of value added of industrial enterprises above designated size',
           'reduce by', '13.5%', 2020, '2025',
           'r', 'Energy efficiency', 'enterprises',
           'By 2025, the energy consumption per unit of value added of industrial enterprises above designated size will be reduced by 13.5%.',
           'FYP2101.pdf', 'Industry'))
ws.append((2021, 'water consumption per unit of industrial value added', 'reduce by',
           '16%', 2020, '2025',
           'r', 'Resource efficiency', 'sectors',
           'By 2025, water consumption per unit of industrial value added will be reduced by 16%.',
           'FYP2101.pdf', 'Industry'))

# 小修：non-ferrous metal kilns 4000 kilns -> about 4000 kilns
for r in range(2, wb['Industry'].max_row + 1):
    if wb['Industry'].cell(row=r, column=2).value == 'clean production transformation of non-ferrous metal kilns' \
            and wb['Industry'].cell(row=r, column=11).value == 'FYP2101.pdf':
        wb['Industry'].cell(row=r, column=4, value='about 4000 kilns')
        break

wb['README']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_en.xlsx')
print('EN done: GHG Emissions =', wb['GHG Emissions'].max_row, '| Pollution =', wb['Pollution'].max_row,
      '| Circular Economy =', wb['Circular Economy'].max_row, '| Industry =', wb['Industry'].max_row)
