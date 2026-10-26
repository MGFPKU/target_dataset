# -*- coding: utf-8 -*-
"""林业草原保护利用“十五五”规划（林规发〔2026〕23号，FYP2617）：5 条目标写入 CN/EN 数据集。

用户确认：附件1 主要指标中纳入 4 项（林草覆盖率、森林火灾受害率、草原火灾受害率、
草原综合植被盖度）+ 产业总产值（生产导向）；森林覆盖率 25.8%、森林蓄积量 224 亿m³、
自然保护地面积占比≥18% 已在数据集中由其他文件收录，不重复录入。

CN: LULUCF 表 +5 行、来源表插 FYP2617、说明表日期 2026-10-26
EN: LULUCF +5 行、Sources 插 FYP2617、README 日期 2026-10-26
"""
import shutil
import datetime
import openpyxl

S_COVER_CN = '附件1“十五五”林业草原保护利用主要指标：林草覆盖率（%）2024年56.09、2030年56.4。'
S_FIRE_CN = '附件1“十五五”林业草原保护利用主要指标：森林/草原火灾受害率（‰）≤0.9/≤2、2030年≤0.7/≤1.8。'
S_GRASS_CN = '附件1“十五五”林业草原保护利用主要指标：草原综合植被盖度（%）2024年50.52、2030年51。'
S_OUTPUT_CN = '附件1“十五五”林业草原保护利用主要指标：林草产业总产值（万亿元）11、2030年14。'

S_COVER_EN = ('Annex 1: Main indicators for forest and grassland protection and utilization during '
              'the 15th Five-Year Plan period: forest and grassland coverage rate (%) - 56.09 in 2024, 56.4 in 2030.')
S_FIRE_EN = ('Annex 1: Main indicators for forest and grassland protection and utilization during '
             'the 15th Five-Year Plan period: forest/grassland fire damage rate (per mille) - '
             'no more than 0.9/2, and no more than 0.7/1.8 in 2030.')
S_GRASS_EN = ('Annex 1: Main indicators for forest and grassland protection and utilization during '
              'the 15th Five-Year Plan period: grassland vegetation coverage (%) - 50.52 in 2024, 51 in 2030.')
S_OUTPUT_EN = ('Annex 1: Main indicators for forest and grassland protection and utilization during '
               'the 15th Five-Year Plan period: total output value of the forestry and grassland '
               'industry (trillion yuan) - 11, 14 in 2030.')

CN_SRC = [None, 'FYP2617', 2026, '林业草原保护利用“十五五”规划',
          '15th Five-Year Plan for Forest and Grassland Protection and Utilization',
          '国家林业和草原局、国家发展和改革委员会、财政部、自然资源部', '五年规划', '促进减缓', 'Mitigation-enabling',
          '十五五', '含目标', 'PDF', '中文',
          'https://www.gov.cn/zhengce/zhengceku/202607/content_7076248.htm']
EN_SRC = [None, 'FYP2617', 2026, '林业草原保护利用“十五五”规划',
          '15th Five-Year Plan for Forest and Grassland Protection and Utilization',
          'National Forestry and Grassland Administration, National Development and Reform Commission, '
          'Ministry of Finance and Ministry of Natural Resources',
          '五年规划', 'B.5', 'Mitigation-enabling', '15th FYP', 'TAR', 'PDF', 'CN',
          'https://www.gov.cn/zhengce/zhengceku/202607/content_7076248.htm']

NEW_DATE = datetime.datetime(2026, 10, 26, 0, 0)

# ---------- CN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_cn.xlsx')
ws = wb['土地利用、土地利用变化和林业']
ws.append((2026, '林草覆盖率', '达到', '56.4%', '56.09%（2024）', '2030',
           '新目标', '碳汇', '全社会经济', S_COVER_CN, 'FYP2617.pdf', '土地利用、土地利用变化和林业'))
ws.append((2026, '森林火灾受害率', '控制在…以下', '0.7‰', '0.9‰（2024）', '2030',
           '新目标', '碳汇', '全社会经济', S_FIRE_CN, 'FYP2617.pdf', '土地利用、土地利用变化和林业'))
ws.append((2026, '草原火灾受害率', '控制在…以下', '1.8‰', '2‰（2024）', '2030',
           '新目标', '碳汇', '全社会经济', S_FIRE_CN, 'FYP2617.pdf', '土地利用、土地利用变化和林业'))
ws.append((2026, '草原综合植被盖度', '达到', '51%', '50.52%（2024）', '2030',
           '新目标', '生态保护', '全社会经济', S_GRASS_CN, 'FYP2617.pdf', '土地利用、土地利用变化和林业'))
ws.append((2026, '林草产业总产值', '达到', '14万亿元', '11万亿元（2024）', '2030',
           '新目标', '生产导向', '行业', S_OUTPUT_CN, 'FYP2617.pdf', '土地利用、土地利用变化和林业'))

ws = wb['来源']
ws.insert_rows(192)
for col, val in enumerate(CN_SRC, start=1):
    ws.cell(row=192, column=col, value=val)
ws['B5'] = '="战略文件总数："&COUNTA(D9:D518)'

wb['说明']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_cn.xlsx')
print('CN done: LULUCF rows =', wb['土地利用、土地利用变化和林业'].max_row)

# ---------- EN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_en.xlsx')
ws = wb['LULUCF']
ws.append((2026, 'forest and grassland coverage rate', 'reach', '56.4%', '56.09% (2024)', '2030',
           'n', 'Carbon sink', 'whole economy', S_COVER_EN, 'FYP2617.pdf', 'LULUCF'))
ws.append((2026, 'forest fire damage rate', 'bring below', '0.7 per mille', '0.9 per mille (2024)', '2030',
           'n', 'Carbon sink', 'whole economy', S_FIRE_EN, 'FYP2617.pdf', 'LULUCF'))
ws.append((2026, 'grassland fire damage rate', 'bring below', '1.8 per mille', '2 per mille (2024)', '2030',
           'n', 'Carbon sink', 'whole economy', S_FIRE_EN, 'FYP2617.pdf', 'LULUCF'))
ws.append((2026, 'grassland vegetation coverage', 'reach', '51%', '50.52% (2024)', '2030',
           'n', 'Ecological protection', 'whole economy', S_GRASS_EN, 'FYP2617.pdf', 'LULUCF'))
ws.append((2026, 'total output value of the forestry and grassland industry', 'reach', '14 trillion yuan',
           '11 trillion yuan (2024)', '2030', 'n', 'Production-oriented', 'sectors', S_OUTPUT_EN,
           'FYP2617.pdf', 'LULUCF'))

ws = wb['Sources']
ws.insert_rows(192)
for col, val in enumerate(EN_SRC, start=1):
    ws.cell(row=192, column=col, value=val)
ws['B5'] = '="Total number of strategy documents: "&COUNTA(D9:D518)'

wb['README']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_en.xlsx')
print('EN done: LULUCF rows =', wb['LULUCF'].max_row)

# ---------- Data_sources ----------
shutil.copyfile(r'..\new_policies_raw\govcn_7076248.htm',
                'Data_sources/Filtering Task Website Source CN/FYP2617.htm')
shutil.copyfile(r'..\new_policies_raw\lycy_attach1.pdf',
                'Data_sources/Filtering Task PDF Source CN/FYP2617.pdf')
print('Data_sources: FYP2617.htm + FYP2617.pdf saved')
