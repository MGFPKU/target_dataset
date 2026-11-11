# -*- coding: utf-8 -*-
"""2026-11-11 DP1601 补录：《工业绿色发展规划（2016-2020年）》（工信部规〔2016〕225号，
2016-06-30 成文）22 条量化目标写入 CN/EN 数据集；来源表 DP1601 URL 更新为
工信部官网链接（原 scio.gov.cn）。

CN: 能源|电力 +1、工业 +10、循环经济 +11、说明最后更新=2026-11-11
EN: Energy|Power +1、Industry +10、Circular Economy +11、README Last update=2026-11-11

来源：专栏1"十三五"时期工业绿色发展主要指标（指导性指标）与专栏6绿色制造体系创建工程。
"""
import shutil
import datetime
import openpyxl

BOX1 = '专栏1"十三五"时期工业绿色发展主要指标：'
BOX1_EN = 'Column 1: Main indicators for green industrial development during the 13th Five-Year Plan period: '

NEW_DATE = datetime.datetime(2026, 11, 11, 0, 0)
NEW_URL = 'https://www.miit.gov.cn/jgsj/jns/wjfb/art/2020/art_d66bb56744d9433d827bdb571de9a250.html'

# ---------- CN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_cn.xlsx')
shutil.copy('China_Climate_Target_Tracker_cn.xlsx', 'China_Climate_Target_Tracker_cn.pre_dp1601add.bak.xlsx')

ws = wb['能源|电力']
ws.append((2016, '绿色低碳能源占工业能源消费量比重', '提高到', '15%', 2015, '2020',
           '新目标', '非化石能源', '行业',
           BOX1 + '绿色低碳能源占工业能源消费量比重（%）2015年12，2020年15',
           'DP1601.htm', '能源|电力', '能源消费', None, None, None, None))

ws = wb['工业']
ws.append((2016, '六大高耗能行业占工业增加值比重', '下降到', '25%', 2015, '2020',
           '新目标', '工业转型', '行业',
           BOX1 + '六大高耗能行业占工业增加值比重（%）2015年27.8，2020年25',
           'DP1601.htm', '工业'))
ws.append((2016, '绿色供应链示范企业', '建立', '100家', None, '2020',
           '新目标', '生产导向', '全社会经济',
           '专栏6绿色制造体系创建工程：在信息通信、汽车、家电、纺织等行业培育百家绿色供应链示范企业。',
           'DP1601.htm', '工业'))
ws.append((2016, '万元工业增加值用水量', '下降', '23%', 2015, '2020',
           '重申目标', '资源效率', '行业',
           BOX1 + '单位工业增加值用水量下降（%）23（累计降速）',
           'DP1601.htm', '工业'))
ws.append((2016, '绿色制造产业产值', '超过', '10万亿元', 2015, '2020',
           '重申目标', '生产导向', '行业',
           BOX1 + '绿色制造产业产值（万亿元）2015年5.3，2020年10',
           'DP1601.htm', '工业'))
ws.append((2016, '吨水泥熟料综合能耗', '降至', '105千克标准煤/吨', 2015, '2020',
           '新目标', '能源效率', '行业',
           BOX1 + '水泥熟料综合能耗（千克标准煤/吨）2015年112，2020年105',
           'DP1601.htm', '工业'))
ws.append((2016, '电解铝液交流电耗', '降至', '13200千瓦时/吨', 2015, '2020',
           '新目标', '能源效率', '行业',
           BOX1 + '电解铝液交流电耗（千瓦时/吨）2015年13350，2020年13200',
           'DP1601.htm', '工业'))
ws.append((2016, '炼油综合能耗', '降至', '63千克标准油/吨', 2015, '2020',
           '新目标', '能源效率', '行业',
           BOX1 + '炼油综合能耗（千克标准油/吨）2015年65，2020年63',
           'DP1601.htm', '工业'))
ws.append((2016, '乙烯综合能耗', '降至', '790千克标准煤/吨', 2015, '2020',
           '新目标', '能源效率', '行业',
           BOX1 + '乙烯综合能耗（千克标准煤/吨）2015年816，2020年790',
           'DP1601.htm', '工业'))
ws.append((2016, '合成氨综合能耗', '降至', '1300千克标准煤/吨', 2015, '2020',
           '新目标', '能源效率', '行业',
           BOX1 + '合成氨综合能耗（千克标准煤/吨）2015年1331，2020年1300',
           'DP1601.htm', '工业'))
ws.append((2016, '纸及纸板综合能耗', '降至', '480千克标准煤/吨', 2015, '2020',
           '新目标', '能源效率', '行业',
           BOX1 + '纸及纸板综合能耗（千克标准煤/吨）2015年530，2020年480',
           'DP1601.htm', '工业'))

ws = wb['循环经济']
ws.append((2016, '主要再生资源回收利用量', '达到', '3.5亿吨', 2015, '2020',
           '新目标', '循环利用', '全社会经济',
           BOX1 + '主要再生资源回收利用量（亿吨）2015年2.2，2020年3.5',
           'DP1601.htm', '循环经济'))
ws.append((2016, '尾矿综合利用率', '达到', '25%', 2015, '2020',
           '新目标', '循环利用', '行业',
           BOX1 + '尾矿（%）2015年22，2020年25',
           'DP1601.htm', '循环经济'))
ws.append((2016, '煤矸石综合利用率', '达到', '71%', 2015, '2020',
           '新目标', '循环利用', '不明确',
           BOX1 + '煤矸石（%）2015年68，2020年71',
           'DP1601.htm', '循环经济'))
ws.append((2016, '工业副产石膏综合利用率', '达到', '60%', 2015, '2020',
           '新目标', '循环利用', '行业',
           BOX1 + '工业副产石膏（%）2015年47，2020年60',
           'DP1601.htm', '循环经济'))
ws.append((2016, '钢铁冶炼渣综合利用率', '达到', '95%', 2015, '2020',
           '新目标', '循环利用', '行业',
           BOX1 + '钢铁冶炼渣（%）2015年79，2020年95',
           'DP1601.htm', '循环经济'))
ws.append((2016, '赤泥综合利用率', '达到', '10%', 2015, '2020',
           '新目标', '循环利用', '行业',
           BOX1 + '赤泥（%）2015年4，2020年10',
           'DP1601.htm', '循环经济'))
ws.append((2016, '再生有色金属年产量', '达到', '1800万吨', 2015, '2020',
           '新目标', '循环利用', '行业',
           BOX1 + '再生有色金属（万吨）2015年1235，2020年1800',
           'DP1601.htm', '循环经济'))
ws.append((2016, '废钢年回收利用量', '达到', '1.5亿吨', 2015, '2020',
           '新目标', '循环利用', '行业',
           BOX1 + '废钢铁（万吨）2015年8330，2020年15000',
           'DP1601.htm', '循环经济'))
ws.append((2016, '废弃电器电子产品回收利用量', '达到', '6.9亿台', 2015, '2020',
           '新目标', '循环利用', '行业',
           BOX1 + '废弃电器电子产品（亿台）2015年4，2020年6.9',
           'DP1601.htm', '循环经济'))
ws.append((2016, '废塑料年回收利用量', '达到', '2300万吨', 2015, '2020',
           '新目标', '循环利用', '全社会经济',
           BOX1 + '废塑料（国内）（万吨）2015年1800，2020年2300',
           'DP1601.htm', '循环经济'))
ws.append((2016, '废旧轮胎年回收利用量', '达到', '850万吨', 2015, '2020',
           '新目标', '循环利用', '全社会经济',
           BOX1 + '废旧轮胎（万吨）2015年550，2020年850',
           'DP1601.htm', '循环经济'))

ws = wb['来源']
for r in range(2, ws.max_row + 1):
    if ws.cell(row=r, column=2).value == 'DP1601':
        ws.cell(row=r, column=14, value=NEW_URL)
        break

wb['说明']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_cn.xlsx')
print('CN done: 能源|电力 =', wb['能源|电力'].max_row, '| 工业 =', wb['工业'].max_row,
      '| 循环经济 =', wb['循环经济'].max_row)

# ---------- EN ----------
wb = openpyxl.load_workbook('China_Climate_Target_Tracker_en.xlsx')
shutil.copy('China_Climate_Target_Tracker_en.xlsx', 'China_Climate_Target_Tracker_en.pre_dp1601add.bak.xlsx')

ws = wb['Energy|Power']
ws.append((2016, 'share of green and low-carbon energy in industrial energy consumption',
           'increase to', '15%', 2015, '2020',
           'n', 'Non-fossil energy', 'sectors',
           BOX1_EN + 'share of green and low-carbon energy in industrial energy consumption (%), 12 in 2015 and 15 in 2020.',
           'DP1601.htm', 'Energy/Power', 'Energy consumption', None, None, None, None))

ws = wb['Industry']
ws.append((2016, 'share of six energy-intensive industries in industrial added value',
           'decrease to', '25%', 2015, '2020',
           'n', 'Industrial transformation', 'sectors',
           BOX1_EN + 'share of six energy-intensive industries in industrial added value (%), 27.8 in 2015 and 25 in 2020.',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'green supply chain demonstration enterprises',
           'establish', '100 enterprises', None, '2020',
           'n', 'Production-oriented', 'whole economy',
           'Column 6: Green manufacturing system creation project: cultivate 100 green supply chain demonstration enterprises in information and communications, automobiles, home appliances, textiles and other industries.',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'water consumption per unit of industrial value added',
           'reduce by', '23%', 2015, '2020',
           'r', 'Resource efficiency', 'sectors',
           BOX1_EN + 'reduction in water consumption per unit of industrial value added (%), 23 (cumulative decrease).',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'output value of green manufacturing industry',
           'exceed', '10 trillion yuan', 2015, '2020',
           'r', 'Production-oriented', 'sectors',
           BOX1_EN + 'output value of green manufacturing industry (trillion yuan), 5.3 in 2015 and 10 in 2020.',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'comprehensive energy consumption per ton of cement clinker',
           'reduce to', '105 kg of coal equivalent per ton', 2015, '2020',
           'n', 'Energy efficiency', 'sectors',
           BOX1_EN + 'comprehensive energy consumption of cement clinker (kg of coal equivalent per ton), 112 in 2015 and 105 in 2020.',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'alternating current electricity consumption per ton of electrolytic aluminum',
           'reduce to', '13,200 kWh per ton', 2015, '2020',
           'n', 'Energy efficiency', 'sectors',
           BOX1_EN + 'alternating current electricity consumption for electrolytic aluminum (kWh per ton), 13,350 in 2015 and 13,200 in 2020.',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'comprehensive energy consumption of oil refining',
           'reduce to', '63 kg of oil equivalent per ton', 2015, '2020',
           'n', 'Energy efficiency', 'sectors',
           BOX1_EN + 'comprehensive energy consumption of oil refining (kg of oil equivalent per ton), 65 in 2015 and 63 in 2020.',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'energy consumption of ethylene',
           'reduce to', '790 kg of standard coal per ton', 2015, '2020',
           'n', 'Energy efficiency', 'sectors',
           BOX1_EN + 'comprehensive energy consumption of ethylene (kg of standard coal per ton), 816 in 2015 and 790 in 2020.',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'energy consumption of synthetic ammonia',
           'reduce to', '1300 kg of standard coal per ton', 2015, '2020',
           'n', 'Energy efficiency', 'sectors',
           BOX1_EN + 'comprehensive energy consumption of synthetic ammonia (kg of standard coal per ton), 1,331 in 2015 and 1,300 in 2020.',
           'DP1601.htm', 'Industry'))
ws.append((2016, 'energy consumption of paper and paperboard',
           'reduce to', '480 kg standard coal per ton', 2015, '2020',
           'n', 'Energy efficiency', 'sectors',
           BOX1_EN + 'comprehensive energy consumption of paper and paperboard (kg of standard coal per ton), 530 in 2015 and 480 in 2020.',
           'DP1601.htm', 'Industry'))

ws = wb['Circular Economy']
ws.append((2016, 'recycling volume of major renewable resources',
           'reach', '350 million tons', 2015, '2020',
           'n', 'Recycling', 'whole economy',
           BOX1_EN + 'recycling volume of major renewable resources (100 million tons), 2.2 in 2015 and 3.5 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'utilization rate of tailings',
           'reach', '25%', 2015, '2020',
           'n', 'Recycling', 'sectors',
           BOX1_EN + 'tailings (%), 22 in 2015 and 25 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'utilization rate of coal gangue',
           'reach', '71%', 2015, '2020',
           'n', 'Recycling', 'unclear',
           BOX1_EN + 'coal gangue (%), 68 in 2015 and 71 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'utilization rate of industrial by-product gypsum',
           'reach', '60%', 2015, '2020',
           'n', 'Recycling', 'sectors',
           BOX1_EN + 'industrial by-product gypsum (%), 47 in 2015 and 60 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'utilization rate of iron and steel smelting slag',
           'reach', '95%', 2015, '2020',
           'n', 'Recycling', 'sectors',
           BOX1_EN + 'iron and steel smelting slag (%), 79 in 2015 and 95 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'utilization rate of red mud',
           'reach', '10%', 2015, '2020',
           'n', 'Recycling', 'sectors',
           BOX1_EN + 'red mud (%), 4 in 2015 and 10 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'non-ferrous metals recycling annual output',
           'reach', '18 million tons', 2015, '2020',
           'n', 'Recycling', 'sectors',
           BOX1_EN + 'recycled non-ferrous metals (10,000 tons), 1,235 in 2015 and 1,800 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'scrap steel annual recycling and utilisation',
           'reach', '150 million tons', 2015, '2020',
           'n', 'Recycling', 'sectors',
           BOX1_EN + 'scrap steel (10,000 tons), 8,330 in 2015 and 15,000 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'recycling volume of waste electrical and electronic products',
           'reach', '690 million units', 2015, '2020',
           'n', 'Recycling', 'sectors',
           BOX1_EN + 'waste electrical and electronic products (100 million units), 4 in 2015 and 6.9 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'waste plastic annual recycling and utilization',
           'reach', '23 million tons', 2015, '2020',
           'n', 'Recycling', 'whole economy',
           BOX1_EN + 'waste plastics (domestic) (10,000 tons), 1,800 in 2015 and 2,300 in 2020.',
           'DP1601.htm', 'Circular Economy'))
ws.append((2016, 'waste tire annual recycling and utilization',
           'reach', '8.5 million tons', 2015, '2020',
           'n', 'Recycling', 'whole economy',
           BOX1_EN + 'waste tires (10,000 tons), 550 in 2015 and 850 in 2020.',
           'DP1601.htm', 'Circular Economy'))

ws = wb['Sources']
for r in range(2, ws.max_row + 1):
    if ws.cell(row=r, column=2).value == 'DP1601':
        ws.cell(row=r, column=14, value=NEW_URL)
        break

wb['README']['C5'] = NEW_DATE
wb.save('China_Climate_Target_Tracker_en.xlsx')
print('EN done: Energy|Power =', wb['Energy|Power'].max_row, '| Industry =', wb['Industry'].max_row,
      '| Circular Economy =', wb['Circular Economy'].max_row)
