"""Add the CAC dual digital-green plan target (IP2604) to Targets_cn.xlsx and Targets_en.xlsx.

Source doc: 促进数字化绿色化协同转型发展实施方案（2026-2030年）
https://www.cac.gov.cn/2026-09/04/c_1790271981772781.htm

Changes per file:
  - EN Energy|Power: insert the new target at its casefold-sorted row (469)
  - CN 能源|电力: append at the bottom (row 564)
  - Sources: insert IP2604 after IP2603 (row 460 -> new row 461)
  - 说明/README: add changelog entry at B6, refresh Last update date
  - Sources B5 COUNTA range: D9:D512 -> D9:D513

Run from target_dataset/ with plain python. Idempotent (skips if IP2604 present).
"""
import datetime
import shutil
from copy import copy

import openpyxl

CN_SENTENCE = (
    "主要目标是：到2030年，人工智能低成本、高效率优势更加彰显，深度赋能绿色低碳技术创新和产业发展，"
    "算力设施、5G基站等新兴领域用能效率大幅提升，推动算力设施等重点领域可再生能源电力消费水平"
    "达到所在省份可再生能源电力消纳责任权重水平，农业、制造业、贸易、消费、城市运行等重点领域"
    "数字化绿色化深入推进，数智技术赋能生态环境治理能力持续提升，双化协同发展模式有力支撑"
    "经济社会发展全面绿色转型和高质量发展。"
)

EN_SENTENCE = (
    "The main goal is: by 2030, the low-cost, high-efficiency advantages of artificial intelligence "
    "will be further demonstrated, deeply empowering green and low-carbon technological innovation and "
    "industrial development; the energy efficiency of emerging fields such as computing facilities and "
    "5G base stations will be greatly improved; the renewable electricity consumption level in computing "
    "facilities and other key areas will reach the level of the provincial renewable electricity "
    "consumption responsibility weight; digital and green transformation in key areas such as agriculture, "
    "manufacturing, trade, consumption, and urban operations will be further advanced; the capacity of "
    "digital and intelligent technologies to empower ecological and environmental governance will "
    "continue to improve; and the coordinated digital-green development model will strongly support the "
    "comprehensive green transformation and high-quality development of the economy and society."
)

CN_TARGET = [
    2026,
    "算力设施等重点领域可再生能源电力消费水平",
    "达到",
    "所在省份可再生能源电力消纳责任权重水平",
    None,
    2030,
    "新目标",
    "可再生能源",
    "行业",
    CN_SENTENCE,
    "IP2604.htm",
    "能源|电力",
    "能源消费",
]

EN_TARGET = [
    2026,
    "renewable electricity consumption level in computing facilities and other key areas",
    "reach",
    "the level of the provincial renewable electricity consumption responsibility weight",
    None,
    2030,
    "n",
    "Renewable energy",
    "sectors",
    EN_SENTENCE,
    "IP2604.htm",
    "Energy|Power",
    "Energy consumption",
]

SRC_URL = "https://www.cac.gov.cn/2026-09/04/c_1790271981772781.htm"
DOC_ZH = "促进数字化绿色化协同转型发展实施方案（2026-2030年）"
DOC_EN = "Implementation Plan for Promoting Coordinated Digital and Green Transformation (2026-2030)"

CN_SOURCE = [
    "IP2604", 2026, DOC_ZH, DOC_EN,
    "中央网信办等部门", "实施方案", "明确减缓", "B.3", "2026-2030",
    "含目标", "HTM", "中文", SRC_URL,
]

EN_SOURCE = [
    "IP2604", 2026, DOC_ZH, DOC_EN,
    "Cyberspace Administration of China and other departments", "实施方案",
    "Explicit mitigation", "B.3", "2026-2030",
    "TAR", "HTM", "CN", SRC_URL,
]

CHANGELOG_CN = (
    "2026-09-07：根据《促进数字化绿色化协同转型发展实施方案（2026-2030年）》（IP2604）新增1条目标："
    "算力设施等重点领域可再生能源电力消费水平达到所在省份可再生能源电力消纳责任权重水平（2030）。"
)

CHANGELOG_EN = (
    "2026-09-07: Added 1 target from the Implementation Plan for Promoting Coordinated Digital and Green "
    "Transformation (2026-2030) (IP2604): renewable electricity consumption level in computing facilities "
    "and other key areas reaching the provincial renewable electricity consumption responsibility weight (2030)."
)


def copy_style(src, dst) -> None:
    dst.font = copy(src.font)
    dst.alignment = copy(src.alignment)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = src.number_format
    dst.protection = copy(src.protection)


def write_row(ws, row_idx, values, template_row, start_col: int = 1) -> None:
    for col_idx, value in enumerate(values, start=start_col):
        cell = ws.cell(row=row_idx, column=col_idx)
        template = ws.cell(row=template_row, column=col_idx)
        copy_style(template, cell)
        cell.value = value


def add_changelog(wb, sheet_name, text) -> None:
    ws = wb[sheet_name]
    for row in ws.iter_rows(min_col=2, max_col=2):
        if row[0].value == text:
            return  # already present (idempotent re-run)
    ws.insert_rows(6)
    src = ws.cell(row=7, column=2)
    dst = ws.cell(row=6, column=2)
    copy_style(src, dst)
    dst.value = text
    ws.row_dimensions[6].height = 15.0
    ws.cell(row=5, column=3).value = datetime.datetime(2026, 9, 7)


def process(path, data_sheet, target_row_values, insert_idx, append_row, readme_sheet,
            changelog_text, source_row_values) -> None:
    backup = path.replace(".xlsx", ".pre_ip2604.bak")
    shutil.copy2(path, backup)
    print(f"backup -> {backup}")

    wb = openpyxl.load_workbook(path)
    src_ws = wb["来源" if data_sheet == "能源|电力" else "Sources"]

    # 1. Data sheet: EN inserts at sorted row; CN appends at bottom
    if insert_idx is not None:
        template = insert_idx  # styles captured from the row that will shift down
        wb[data_sheet].insert_rows(insert_idx)
        write_row(wb[data_sheet], insert_idx, target_row_values, template + 1)
        if wb[data_sheet].row_dimensions[insert_idx + 1].height:
            wb[data_sheet].row_dimensions[insert_idx].height = wb[data_sheet].row_dimensions[insert_idx + 1].height
    else:
        write_row(wb[data_sheet], append_row, target_row_values, append_row - 1)
    print(f"  {data_sheet}: target row written")

    # 2. Sources: insert IP2604 after IP2603
    exists = any(
        src_ws.cell(row=r, column=2).value == "IP2604"
        for r in range(9, src_ws.max_row + 1)
    )
    if exists:
        print("  Sources: IP2604 already present, skipping")
    else:
        anchor = max(
            r for r in range(9, src_ws.max_row + 1)
            if src_ws.cell(row=r, column=2).value == "IP2603"
        )
        src_ws.insert_rows(anchor + 1)
        write_row(src_ws, anchor + 1, source_row_values, anchor, start_col=2)
        b5 = src_ws.cell(row=5, column=2).value
        new_b5 = b5.replace("D9:D512", "D9:D513")
        if new_b5 != b5:
            src_ws.cell(row=5, column=2).value = new_b5
        print(f"  Sources: IP2604 inserted at row {anchor + 1}, B5 -> {new_b5}")

    # 3. README changelog
    add_changelog(wb, readme_sheet, changelog_text)
    print(f"  {readme_sheet}: changelog added")

    wb.save(path)
    print(f"{path}: saved")


if __name__ == "__main__":
    # EN: insert at sorted position 469 (between 'refined oil annual output' and
    # 'renewable energy share in electrolytic aluminium energy use')
    process(
        "Targets_en.xlsx",
        "Energy|Power",
        EN_TARGET,
        insert_idx=469,
        append_row=None,
        readme_sheet="README",
        changelog_text=CHANGELOG_EN,
        source_row_values=EN_SOURCE,
    )
    # CN: append at bottom row 564 (after last row 563)
    process(
        "Targets_cn.xlsx",
        "能源|电力",
        CN_TARGET,
        insert_idx=None,
        append_row=564,
        readme_sheet="说明",
        changelog_text=CHANGELOG_CN,
        source_row_values=CN_SOURCE,
    )
