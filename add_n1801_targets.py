# -*- coding: utf-8 -*-
"""2026-11-04 N1801 入库：农业农村部《农业绿色发展技术导则（2018-2030年）》（农科教发〔2018〕3号，
2018-07-02，国务院公报 2018 年第 35 号）4 条目标写入 CN/EN 数据集，来源表登记 N1801（含目标）。

CN: 土地利用、土地利用变化和林业 +1、污染 +1、温室气体排放 +1、能源|电力 +1、来源表插 N1801、说明最后更新=2026-11-04
EN: LULUCF +1、Pollution +1、GHG Emissions +1、Energy|Power +1、Sources 插 N1801、README Last update=2026-11-04
"""
import shutil
import datetime
import openpyxl

S1_CN = ('研发一批土壤改良培肥、雨养和节水灌溉、精准施肥、有害生物绿色防控、畜禽水产健康养殖和废弃物循环利用、'
         '面源污染治理和农业生态修复、轻简节本高效机械化作业、农产品收储运和加工等农业绿色生产技术，'
         '实现农田灌溉用水有效利用系数提高到0.6以上，主要作物化肥、农药利用率显著提高，'
         '农业源氮、磷污染物排放强度和负荷分别削减30%和40%以上，养殖节水源头减排20%以上，'
         '畜禽饲料转化率、水产养殖精准投喂水平较目前分别提升10%以上，农产品加工单位产值能耗较目前降低20%以上。')

S2_CN = ('形成一批主要作物绿色增产增效、种养加循环、区域低碳循环、田园综合体等农业绿色发展模式，'
         '技术模式的单位农业增加值温室气体排放强度和能耗降低30%以上，构建绿色轻简机械化种植、规模化养殖工艺模式，'
         '基本实现农业生产全程机械化，清洁化、农业废弃物全循环、农业生态服务功能大幅增强。')

S1_EN = ('Research and develop a batch of green agricultural production technologies, including soil improvement and '
         'fertility enhancement, rain-fed and water-saving irrigation, precision fertilization, green pest prevention '
         'and control, healthy livestock and aquaculture farming and waste recycling, non-point source pollution '
         'control and agricultural ecological restoration, light, simple, cost-saving and efficient mechanized '
         'operations, and collection, storage, transportation and processing of agricultural products; increase the '
         'effective utilization coefficient of farmland irrigation water to above 0.6; significantly increase the '
         'utilization rates of chemical fertilizers and pesticides for major crops; reduce the emission intensity and '
         'load of nitrogen and phosphorus pollutants from agricultural sources by more than 30% and 40%, respectively; '
         'reduce water use at source in aquaculture by more than 20%; increase the feed conversion rate of livestock '
         'and poultry and the precision feeding level of aquaculture by more than 10%, respectively, compared with '
         'current levels; and reduce energy consumption per unit of output value in agricultural product processing '
         'by more than 20% compared with current levels.')

S2_EN = ('Form a batch of agricultural green development models, such as green yield- and efficiency-enhancing models '
         'for major crops, integrated planting-breeding-processing circular models, regional low-carbon circular '
         'models and pastoral complexes; reduce the greenhouse gas emission intensity and energy consumption per unit '
         'of agricultural added value of technology models by more than 30%; build green, light and simple mechanized '
         'planting models and large-scale breeding process models; basically realize full-process mechanization and '
         'clean production of agriculture; and significantly enhance the full recycling of agricultural waste and '
         'agricultural ecological service functions.')

CN_SRC = [None, 'N1801', 2018, '农业农村部关于印发《农业绿色发展技术导则（2018-2030年）》的通知',
          'Notice of the Ministry of Agriculture and Rural Affairs on Printing and Issuing the Technical Guidelines '
          'for Green Agricultural Development (2018-2030)',
          '农业农村部', '通知', '明确减缓', 'B.5', '2018-2030', '含目标', 'HTM', '中文',
          'https://www.gov.cn/gongbao/content/2018/content_5350058.htm']

EN_SRC = [None, 'N1801', 2018, '农业农村部关于印发《农业绿色发展技术导则（2018-2030年）》的通知',
          'Notice of the Ministry of Agriculture and Rural Affairs on Printing and Issuing the Technical Guidelines '
          'for Green Agricultural Development (2018-2030)',
          'Ministry of Agriculture and Rural Affairs', '通知', 'Explicit mitigation', 'B.5', '2018-2030', 'TAR',
          'HTM', 'CN', 'https://www.gov.cn/gongbao/content/2018/content_5350058.htm']

NEW_DATE = datetime.datetime(2026, 11, 4, 0, 0)

# ---------- CN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_cn.xlsx')
shutil.copy('China_Climate_Target_Tracker_cn.xlsx', 'China_Climate_Target_Tracker_cn.pre_n1801.bak.xlsx')

ws = wb['土地利用、土地利用变化和林业']
ws.append((2018, '农田灌溉用水有效利用系数', '提高到', '0.6以上', None, 2030,
           '新目标', '资源效率', '行业', S1_CN, 'N1801.htm', '土地利用、土地利用变化和林业'))

ws = wb['污染']
ws.append((2018, '农业源氮、磷污染物排放强度和负荷', '分别削减', '30%和40%以上', None, 2030,
           '新目标', '污染控制', '行业', S1_CN, 'N1801.htm', '污染'))

ws = wb['温室气体排放']
ws.append((2018, '技术模式的单位农业增加值温室气体排放强度和能耗', '降低', '30%以上', None, 2030,
           '新目标', '碳强度', '行业', S2_CN, 'N1801.htm', '温室气体排放'))

ws = wb['能源|电力']
ws.append((2018, '农产品加工单位产值能耗', '降低', '20%以上', None, 2030,
           '新目标', '能源效率', '行业', S1_CN, 'N1801.htm', '能源|电力', '能源消费'))

ws = wb['来源']
ws.insert_rows(472)
for col, val in enumerate(CN_SRC, start=1):
    ws.cell(row=472, column=col, value=val)
ws['B5'] = '="战略文件总数："&COUNTA(D9:D519)'

wb['说明']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_cn.xlsx')
print('CN done: LULUCF rows =', wb['土地利用、土地利用变化和林业'].max_row,
      '| 污染 =', wb['污染'].max_row, '| GHG =', wb['温室气体排放'].max_row,
      '| 能源 =', wb['能源|电力'].max_row, '| 来源 =', wb['来源'].max_row)

# ---------- EN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_en.xlsx')
shutil.copy('China_Climate_Target_Tracker_en.xlsx', 'China_Climate_Target_Tracker_en.pre_n1801.bak.xlsx')

ws = wb['LULUCF']
ws.append((2018, 'effective utilization coefficient of farmland irrigation water', 'increase to', 'above 0.6',
           None, 2030, 'n', 'Resource efficiency', 'sectors', S1_EN, 'N1801.htm', 'LULUCF'))

ws = wb['Pollution']
ws.append((2018, 'emission intensity and load of nitrogen and phosphorus pollutants from agricultural sources',
           'reduce by', 'over 30% and 40%, respectively', None, 2030, 'n', 'Pollution control', 'sectors',
           S1_EN, 'N1801.htm', 'Pollution'))

ws = wb['GHG Emissions']
ws.append((2018, 'greenhouse gas emission intensity and energy consumption per unit of agricultural added value of '
                 'technology models', 'reduce by', 'over 30%', None, 2030, 'n', 'Carbon intensity', 'sectors',
           S2_EN, 'N1801.htm', 'Emissions'))

ws = wb['Energy|Power']
ws.append((2018, 'energy consumption per unit of output value in agricultural product processing', 'reduce by',
           'over 20%', None, 2030, 'n', 'Energy efficiency', 'sectors', S1_EN, 'N1801.htm', 'Energy/Power',
           'Energy consumption'))

ws = wb['Sources']
ws.insert_rows(472)
for col, val in enumerate(EN_SRC, start=1):
    ws.cell(row=472, column=col, value=val)
ws['B5'] = '="Total number of strategy documents: "&COUNTA(D9:D519)'

wb['README']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_en.xlsx')
print('EN done: LULUCF rows =', wb['LULUCF'].max_row,
      '| Pollution =', wb['Pollution'].max_row, '| GHG =', wb['GHG Emissions'].max_row,
      '| Energy =', wb['Energy|Power'].max_row, '| Sources =', wb['Sources'].max_row)
