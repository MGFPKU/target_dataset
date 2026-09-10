# -*- coding: utf-8 -*-
"""Merged data-revision scripts for China_Climate_Target_Tracker_cn.xlsx / China_Climate_Target_Tracker_en.xlsx (2026-09).

Each step of the 2026-09 revision process is one subcommand. The fix
subcommands run as a dry-run check by default and need --apply to write
(a .bak backup is made first). rename-categories / unify-percent /
add-ip2604 write immediately (with backups; all three are idempotent).

Usage:
  PYTHONIOENCODING=utf-8 python revisions.py <subcommand> [--apply]

Subcommands:
  compare              CN/EN consistency check (read-only report)
  data                 fix CN/EN data diffs found by compare (dry-run by default)
  category-a           A-list category fixes (CN+EN) + EN Energy|Power row order
  category-b           B-list category fixes, user-judged (CN+EN)
  direction            direction-column wording round 1
  direction2           contradictory direction+value combinations
  direction3           direction/value wording vs 政策原文 round 3
  direction3b          direction/value wording vs 政策原文 round 3b
  rename-categories    drop the trailing 目标/target suffix from category names
  restructure-taxonomy apply the Target Category Revision Rationale (18 -> 16)
  clear-changelog      remove per-revision changelog rows from 说明/README, leaving one blank row
  unify-percent        EN '<digit> percent' -> '<digit>%'
  add-ip2604           add the CAC digital-green plan target (IP2604)
  wording-fix          wording vs 政策原文 (CN first, EN mirrors CN) + EN README legend
  fix-period           来源 sheet 覆盖时期: letter codes -> period values (CN)
  wording2             AP1802 metric + 约/左右 hedge alignment + O1802 baseline (CN+EN)
  fix-double-space     EN data sheets: collapse double-space typos in metric/mag/sentence
"""
import sys, re, shutil, datetime
from collections import defaultdict, Counter
from copy import copy

import openpyxl
from openpyxl.cell.cell import MergedCell

CN_FILE = 'China_Climate_Target_Tracker_cn.xlsx'
EN_FILE = 'China_Climate_Target_Tracker_en.xlsx'
# ---------------------------------------------------------------------------
# compare: CN/EN consistency check (originally compare_cn_en.py)
# ---------------------------------------------------------------------------
CN, EN = "China_Climate_Target_Tracker_cn.xlsx", "China_Climate_Target_Tracker_en.xlsx"

CN_COLUMN_MAP = {
    "公布年份": "Announcement_Year", "指标": "Metric", "方向": "Direction",
    "目标值": "Target_Magnitude", "基线": "Baseline",
    "目标年份/时期": "Target_Year_or_Period", "计数": "Count",
    "目标类别": "Target_Category", "责任主体": "Accountability",
    "政策原文": "Sentence", "文件": "Document", "主题标签": "Topic_Label",
    "子主题标签": "Subtopic_Label",
}

CN_DIGIT = {"一": 1, "二": 2, "两": 2, "三": 3, "四": 4, "五": 5, "六": 6,
            "七": 7, "八": 8, "九": 9, "〇": 0, "零": 0}
CN_DIGIT10 = dict(CN_DIGIT); CN_DIGIT10["十"] = 10

FYP_MAP = {"十五五": 15, "十四五": 14, "十三五": 13, "十二五": 12, "十一五": 11,
           "十五": 15, "十四": 14, "十三": 13, "十二": 12, "十一": 11,
           "十": 10, "一": 1, "二": 2, "三": 3, "四": 4, "五": 5, "六": 6,
           "七": 7, "八": 8, "九": 9}

def parse_cn_int(s):
    total, cur = 0, 0
    for ch in s:
        if ch in "一二两三四五六七八九":
            cur = CN_DIGIT[ch]
        elif ch == "十":
            total += (cur or 1) * 10; cur = 0
        elif ch == "百":
            total += (cur or 1) * 100; cur = 0
        elif ch == "千":
            total += (cur or 1) * 1000; cur = 0
        elif ch == "万":
            total = (total + cur or 1) * 10000; cur = 0
    return total + cur

def cn_numeral(s):
    m = re.match(r"([一两二三四五六七八九十百千]+)(千万|百万|十万|万亿|亿|万|千|百)", s)
    if not m:
        return None
    return parse_cn_int(m.group(1)) * {"百": 1e2, "千": 1e3, "万": 1e4,
        "亿": 1e8, "万亿": 1e12, "十万": 1e5, "百万": 1e6, "千万": 1e7}[m.group(2)]

FYP_CN_RE = re.compile(
    r"(?<![一两二三四五六七八九十之])(?:十五五|十四五|十三五|十二五|十一五|十五|十四|十三|十二|十一|十|[一两二三四五六七八九])五"
    r"(?=(?:规划|期间|时期|末|初|阶段|目标)|$|[\s\"'“”（）()，,。；;])")

def norm_cn_text(s):
    vals = set()
    # digit ranges: "8—10万亿" / "1.2~1.3万"
    for m in re.finditer(r"(?<![A-Za-z0-9.])(\d+(?:\.\d+)?)[—–~-](\d+(?:\.\d+)?)\s*(万亿|亿|万)?", s):
        mult = {"万": 1e4, "亿": 1e8, "万亿": 1e12}.get(m.group(3) or "", 1)
        vals.add(round(float(m.group(1)) * mult, 6))
        vals.add(round(float(m.group(2)) * mult, 6))
    for m in re.finditer(
        r"(?<![A-Za-z0-9.,])(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)(?!\.\d)(?![\d—–~-])(?!月)\s*余?\s*"
        r"(万亿|亿|万|百|%|％|Mt|mt|多?家|多?个|多?户|多?套|多?辆|多?次|多?台|多?座)?", s):
        num = float(m.group(1).replace(",", ""))
        mult = {"百": 1e2, "万": 1e4, "亿": 1e8, "万亿": 1e12, "Mt": 1e6,
                "mt": 1e6, "%": 1, "％": 1, "家": 1, "个": 1, "户": 1, "套": 1,
                "辆": 1, "次": 1, "台": 1, "座": 1, "多家": 1, "多个": 1,
                "多户": 1, "多套": 1, "多辆": 1, "多次": 1, "多台": 1,
                "多座": 1}.get(m.group(2) or "", 1)
        vals.add(round(num * mult, 6))
    for m in re.finditer(r"[一两二三四五六七八九十百千]+(?:千万|百万|十万|万亿|亿|万|千|百)", s):
        v = cn_numeral(m.group())
        if v:
            vals.add(round(v, 6))
    # 百分之五十五 -> 55; 百分之三点六 -> 3.6
    for m in re.finditer(r"百分之([一两二三四五六七八九十百千点]+)", s):
        t = m.group(1)
        if "点" in t:
            ip, fp = t.split("点", 1)
            v = float(parse_cn_int(ip)) + float(parse_cn_int(fp)) / (10 ** len(fp))
        else:
            v = float(parse_cn_int(t))
        vals.add(v)
    # counters: 十五种 / 万种 / 百家 / 千家 (not 一次能源, not 第十一个五年,
    # not 万 in "320万户" — a standalone 万/千/百 counts only when no number
    # precedes, even with a space: "4000 万个").
    # 、-joined numerals are all counted: 一、二类 -> {1, 2}
    for m in re.finditer(
        r"(?<![0-9一两二三四五六七八九十])(?<![\d一两二三四五六七八九十]\s)"
        r"([一两二三四五六七八九十百千万]+(?:、[一两二三四五六七八九十百千万]+)*)(?:多)?(?:种|个|家|户|项|类|套|辆|台|座|倍|件)(?!五年)", s):
        for part in m.group(1).split('、'):
            vals.add(float(parse_cn_int(part)))
    if "百公里" in s:
        vals.add(100.0)
    # 万元 -> 10,000 (EN "10,000 yuan"); exclude 万亿元 / 1000万元.
    # Only when the text has another number, so bare metric names like
    # "万元工业增加值用水量" don't add a value the EN metric lacks.
    if re.search(r"(?<![\d亿])(?<![\d亿]\s)万元", s) and re.search(r"\d", s):
        vals.add(10000.0)
    for m in FYP_CN_RE.finditer(s):
        vals.add(f"FYP{FYP_MAP[m.group()]}")
    # 第十二个五年 -> FYP12
    for m in re.finditer(r"第([一两二三四五六七八九十]+)个五年", s):
        vals.add(f"FYP{parse_cn_int(m.group(1))}")
    # X千瓦 -> kW plus W- and GW-scale equivalents (mirrors EN kW/MW/GW rules).
    # \s* tolerates the dataset's spaced convention "30 万千瓦".
    for m in re.finditer(r"(?<![A-Za-z0-9.,])(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*(万亿|亿|万)?\s*千瓦(?!时)", s):
        num = float(m.group(1).replace(",", ""))
        mult = {"万": 1e4, "亿": 1e8, "万亿": 1e12}.get(m.group(2) or "", 1)
        v = num * mult
        vals.add(round(v, 6))
        vals.add(round(v * 1e3, 6))
        vals.add(round(v / 1e6, 6))
    # Chinese-numeral power units: 四千万千瓦 / 五十吉瓦
    # (the spaced-digit guard keeps "30 万千瓦" from parsing 万 as a bare 10000)
    for m in re.finditer(r"(?<![0-9一两二三四五六七八九十])(?<![\d一两二三四五六七八九十]\s)"
                         r"([一两二三四五六七八九十百千万]+)\s*千瓦(?!时)", s):
        v = float(parse_cn_int(m.group(1)))
        vals.add(round(v, 6))
        vals.add(round(v * 1e3, 6))
        vals.add(round(v / 1e6, 6))
    # X吉瓦 -> gigawatts (bare value plus W- and kW-scale)
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*吉瓦", s):
        num = float(m.group(1))
        vals.add(round(num, 6))
        vals.add(round(num * 1e9, 6))
        vals.add(round(num * 1e6, 6))
    for m in re.finditer(r"(?<![0-9一两二三四五六七八九十])(?<![\d一两二三四五六七八九十]\s)"
                         r"([一两二三四五六七八九十百千万]+)\s*吉瓦", s):
        num = float(parse_cn_int(m.group(1)))
        vals.add(round(num, 6))
        vals.add(round(num * 1e9, 6))
        vals.add(round(num * 1e6, 6))
    if "千分之一" in s:
        vals.add(0.001)
        vals.add(1.0)
    for m in re.finditer(r"(\d+(?:\.\d+)?)‰", s):
        v = float(m.group(1))
        vals.add(v)
        vals.add(round(v / 1000, 6))
    for y in years_cn(s):
        vals.discard(float(y))
    return vals

WORD_NUM = {"one": 1, "two": 2, "three": 3, "four": 4, "five": 5, "six": 6,
            "seven": 7, "eight": 8, "nine": 9, "ten": 10, "eleven": 11,
            "twelve": 12, "thirteen": 13, "fourteen": 14, "fifteen": 15,
            "sixteen": 16, "seventeen": 17, "eighteen": 18, "nineteen": 19,
            "twenty": 20}

ROMAN = {"i": 1, "ii": 2, "iii": 3, "iv": 4, "v": 5, "vi": 6, "vii": 7,
         "viii": 8, "ix": 9, "x": 10, "xi": 11, "xii": 12, "xiii": 13,
         "xiv": 14, "xv": 15, "xvi": 16, "xvii": 17, "xviii": 18,
         "xix": 19, "xx": 20}

def norm_en_text(s):
    vals = set()
    # ranges: "100-150 million tons" (also en-dash / em-dash / tilde)
    for m in re.finditer(
        r"(?<![A-Za-z0-9.])(\d+(?:\.\d+)?)[—–~-](\d+(?:\.\d+)?)\s*"
        r"(trillion|billion|million|thousand)?",
        s, re.IGNORECASE):
        w = (m.group(3) or "").lower()
        mult = {"thousand": 1e3, "million": 1e6, "billion": 1e9,
                "trillion": 1e12}.get(w, 1)
        vals.add(round(float(m.group(1)) * mult, 6))
        vals.add(round(float(m.group(2)) * mult, 6))
    for m in re.finditer(
        r"(?<![A-Za-z0-9.])(\d{1,3}(?:,\d{3})+(?:\.\d+)?|\d+(?:\.\d+)?(?!,\d{3}))"
        r"(?![\d—–~-])(?!st\b|nd\b|rd\b|th\b)"
        r"\s*(trillion|billion|million|thousand|giga|mega|percent|Mt|mt)?",
        s, re.IGNORECASE):
        num = float(m.group(1).replace(",", ""))
        w = (m.group(2) or "").lower()
        mult = {"thousand": 1e3, "million": 1e6, "billion": 1e9,
                "trillion": 1e12, "giga": 1e9, "mega": 1e6,
                "percent": 1, "mt": 1e6}.get(w, 1)
        vals.add(round(num * mult, 6))
    # parenthesized numbers mirror CN "（五）" which tokenizes to nothing
    for m in re.finditer(r"\((\d+(?:\.\d+)?)\)", s):
        vals.discard(float(m.group(1)))
    # bare numbers before power units (e.g. "600 MW"): CN attaches 万/吉瓦 to
    # the number and adds no bare value; the kW/MW/GW loops add scaled
    # equivalents instead. Hour forms (kWh) keep the bare number, like CN 千瓦时.
    for m in re.finditer(
        r"(?<![A-Za-z0-9.])(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*"
        r"(?:kilowatts?|kW|megawatts?|MW|gigawatts?|GW)\b(?![- ]?h(?:ours?)?)",
        s, re.IGNORECASE):
        vals.discard(float(m.group(1).replace(",", "")))
    # word numbers: "one million", "two hundred million"
    for m in re.finditer(
        r"\b(one|two|three|four|five|six|seven|eight|nine|ten|eleven|twelve|"
        r"thirteen|fourteen|fifteen|sixteen|seventeen|eighteen|nineteen|twenty)\s+"
        r"(hundred|thousand|million|billion|trillion)"
        r"(?:\s+(hundred|thousand|million|billion|trillion))?",
        s, re.IGNORECASE):
        mult = {"hundred": 1e2, "thousand": 1e3, "million": 1e6,
                "billion": 1e9, "trillion": 1e12}[m.group(2).lower()]
        v = WORD_NUM[m.group(1).lower()] * mult
        if m.group(3):
            v *= {"hundred": 1e2, "thousand": 1e3, "million": 1e6,
                  "billion": 1e9, "trillion": 1e12}[m.group(3).lower()]
        vals.add(round(v, 6))
    # standalone small number words: "six major categories"
    for m in re.finditer(
        r"\b(six|seven|eight|nine|ten|eleven|twelve|thirteen|fourteen|"
        r"fifteen|sixteen|seventeen|eighteen|nineteen|twenty)\b",
        s, re.IGNORECASE):
        vals.add(float(WORD_NUM[m.group(1).lower()]))
    # Roman numerals after category words: "Class I and II" (mirrors CN 一、二类).
    # The \b keeps "level in" (i = 1) out; range forms ("Class I-IV") add
    # nothing, like CN Ⅰ－Ⅳ类.
    for m in re.finditer(
        r"\b(?:[Cc]lass(?:es)?|[Cc]ategor(?:y|ies)|[Tt]ype|[Tt]ier|[Gg]rade|[Ll]evel)s?\s+([IVXivx]+)\b"
        r"(?:\s*(?:,\s*|\s+and\s+)([IVXivx]+))?(?!\s*[-–—]\s*[IVXivx]+\b)",
        s):
        for g in m.groups():
            if g and ROMAN.get(g.lower()):
                vals.add(float(ROMAN[g.lower()]))
    m = re.search(r"(\d+)(?:st|nd|rd|th) (?:FYP|Five[ -]Year)", s)
    if m:
        vals.add(f"FYP{int(m.group(1))}")
    # power units: add kW/W/GW equivalents (mirrors CN 千瓦/吉瓦 rules).
    # Optional magnitude word covers "50 million kilowatts".
    for m in re.finditer(
        r"(?<![A-Za-z0-9.])(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*"
        r"(trillion|billion|million|thousand)?\s*-?\s*(?:kilowatts?|kW)(?![- ]?h(?:ours?)?\b)",
        s, re.IGNORECASE):
        v = float(m.group(1).replace(",", "")) * {"thousand": 1e3,
            "million": 1e6, "billion": 1e9, "trillion": 1e12}.get(
            (m.group(2) or "").lower(), 1)
        vals.add(round(v, 6))
        vals.add(round(v * 1e3, 6))
        vals.add(round(v / 1e6, 6))
    for m in re.finditer(
        r"(?<![A-Za-z0-9.])(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*"
        r"(trillion|billion|million|thousand)?\s*-?\s*(?:megawatts?|MW)(?![- ]?h(?:ours?)?\b)",
        s, re.IGNORECASE):
        v = float(m.group(1).replace(",", "")) * {"thousand": 1e3,
            "million": 1e6, "billion": 1e9, "trillion": 1e12}.get(
            (m.group(2) or "").lower(), 1)
        vals.add(round(v * 1e3, 6))
        vals.add(round(v * 1e6, 6))
        vals.add(round(v / 1e3, 6))
    for m in re.finditer(
        r"(?<![A-Za-z0-9.])(\d{1,3}(?:,\d{3})+|\d+(?:\.\d+)?)\s*"
        r"(trillion|billion|million|thousand)?\s*-?\s*(?:gigawatts?|GW)(?![- ]?h(?:ours?)?\b)",
        s, re.IGNORECASE):
        v = float(m.group(1).replace(",", "")) * {"thousand": 1e3,
            "million": 1e6, "billion": 1e9, "trillion": 1e12}.get(
            (m.group(2) or "").lower(), 1)
        vals.add(round(v, 6))
        vals.add(round(v * 1e6, 6))
        vals.add(round(v * 1e9, 6))
    for m in re.finditer(r"(\d+(?:\.\d+)?)\s*per mille", s, re.IGNORECASE):
        v = float(m.group(1))
        vals.add(v)
        vals.add(round(v / 1000, 6))
    for y in years_en(s):
        vals.discard(float(y))
    return vals

YEAR_RE = r"(?<!\d)(?:19|20)\d{2}(?!\.\d)(?!\d)"

def years_cn(s):
    out = set()
    for m in re.finditer(YEAR_RE, s):
        after = s[m.end():m.end() + 2]
        if after[:1] in "万亿千百多吨米克元瓦亩辆件台套座个户家次载度" or after.startswith(("公顷",)):
            continue
        out.add(int(m.group()))
    for m in re.finditer(r"[二〇零一两二三四五六七八九]{4}(?=年)", s):
        out.add(int("".join(str(CN_DIGIT[c]) for c in m.group())))
    return out

COUNTER_WORDS = ("units", "enterprises", "companies", "parks", "factories",
                 "institutions", "vehicles", "cars", "boilers", "sets",
                 "stations", "schools", "programs", "products",
                 "cooperatives", "households", "tons", "tonnes", "hectares",
                 "kilometers", "kilometres", "kwh", "m2", "km", "meters",
                 "metres", "farms", "plants", "pieces", "projects", "hospitals",
                 "high-standard")

def years_en(s):
    out = set()
    for m in re.finditer(YEAR_RE, s):
        after = s[m.end():]
        wm = re.match(r"\s*(million|billion|thousand|trillion|giga|mega|mt)\b", after, re.I)
        if wm:
            continue
        wm2 = re.match(r"\s*([a-z]+(?:-[a-z]+)*)\b", after, re.I)
        if wm2 and wm2.group(1).lower() in COUNTER_WORDS:
            continue
        out.add(int(m.group()))
    return out

def sets_consistent(a, b):
    if a == b:
        return True
    rem_a, rem_b = set(a - b), set(b - a)
    if len(rem_a) > 2 or len(rem_b) > 2:
        return False
    for v in list(rem_a):
        if round(v / 1e3, 6) in b or round(v * 1e3, 6) in b:
            rem_a.discard(v)
    for v in list(rem_b):
        if round(v / 1e3, 6) in a or round(v * 1e3, 6) in a:
            rem_b.discard(v)
    return not rem_a and not rem_b

def fyp_end_years(s, cn):
    out = set()
    if cn:
        for m in FYP_CN_RE.finditer(s):
            out.add(1960 + (FYP_MAP[m.group()] - 1) * 5)
        for m in re.finditer(r"第([一两二三四五六七八九十]+)个五年", s):
            out.add(1960 + (parse_cn_int(m.group(1)) - 1) * 5)
    else:
        for m in re.finditer(r"(\d+)(?:st|nd|rd|th)\s+(?:FYP|Five[ -]Year)", s, re.I):
            out.add(1960 + (int(m.group(1)) - 1) * 5)
    return out

def cell_pair_ok(c, e, why):
    if c == e:
        return True
    if c is None or e is None:
        why.append(f"one side empty: CN={c!r} EN={e!r}"); return False
    if isinstance(c, (int, float)) and isinstance(e, (int, float)):
        if abs(float(c) - float(e)) > 1e-9:
            why.append(f"numeric mismatch: CN={c!r} EN={e!r}"); return False
        return True
    if isinstance(c, (int, float)) and isinstance(e, str):
        if (abs(float(c)) not in {abs(x) for x in norm_en_text(e) if not isinstance(x, str)}
                and abs(float(c)) not in years_en(e) | fyp_end_years(e, False)):
            why.append(f"CN numeric {c!r} missing in EN text {e!r}"); return False
        return True
    if isinstance(e, (int, float)) and isinstance(c, str):
        if (abs(float(e)) not in {abs(x) for x in norm_cn_text(c) if not isinstance(x, str)}
                and abs(float(e)) not in years_cn(c) | fyp_end_years(c, True)):
            why.append(f"EN numeric {e!r} missing in CN text {c!r}"); return False
        return True
    if isinstance(c, str) and isinstance(e, str):
        yc, ye = years_cn(c), years_en(e)
        if yc != ye:
            why.append(f"YEARS differ ({sorted(yc)} vs {sorted(ye)}): CN={c!r} EN={e!r}")
            return False
        sc, se = norm_cn_text(c), norm_en_text(e)
        fc = {x for x in sc if isinstance(x, str)}
        fe = {x for x in se if isinstance(x, str)}
        if fc != fe:
            why.append(f"FYP differ: CN={c!r} EN={e!r}")
            return False
        nc = {x for x in sc if not isinstance(x, str)}
        ne = {x for x in se if not isinstance(x, str)}
        if not sets_consistent(nc, ne):
            why.append(f"values differ: CN={c!r} (CN:{sorted(nc)}) "
                       f"EN={e!r} (EN:{sorted(ne)})")
            return False
        return True
    why.append(f"type mismatch: CN={type(c).__name__}:{c!r} EN={type(e).__name__}:{e!r}")
    return False

def row_tokens(row, cn):
    toks = set()
    for v in row.values():
        if isinstance(v, str):
            if cn:
                toks |= {x for x in norm_cn_text(v) if not isinstance(x, str)}
                toks |= years_cn(v)
            else:
                toks |= {x for x in norm_en_text(v) if not isinstance(x, str)}
                toks |= years_en(v)
    return toks

def mag_tokens(row, cn):
    v = row.get("Target_Magnitude")
    if isinstance(v, str):
        if cn:
            return {x for x in norm_cn_text(v) if not isinstance(x, str)} | years_cn(v)
        return {x for x in norm_en_text(v) if not isinstance(x, str)} | years_en(v)
    if isinstance(v, (int, float)):
        return {float(v)}
    return set()

def sim(a, b):
    if not a or not b:
        return 0.0
    return len(a & b) / max(len(a), len(b))

def pair_groups(g1, g2):
    """Greedy pairing of CN rows g1 with EN rows g2, magnitude-weighted."""
    t1 = [(mag_tokens(r, True), row_tokens(r, True), r) for r in g1]
    t2 = [(mag_tokens(r, False), row_tokens(r, False), r) for r in g2]
    pairs = []
    while t1 and t2:
        best = max(((10 * sim(ma, mb) + sim(a, b), i, j)
                    for i, (ma, a, _) in enumerate(t1)
                    for j, (mb, b, _) in enumerate(t2)), default=None)
        if best is None or best[0] < 1.0:
            break
        _, i, j = best
        pairs.append((t1[i][2], t2[j][2]))
        del t1[i], t2[j]
    # fallback: rows whose tokens share nothing (fully reworded metric) are
    # paired positionally when the leftover counts match — CN and EN sheets
    # keep identical row order, so the position itself identifies the pair.
    while t1 and t2 and len(t1) == len(t2):
        pairs.append((t1[0][2], t2[0][2]))
        del t1[0], t2[0]
    return pairs, [r for _, _, r in t1], [r for _, _, r in t2]

def load_rows(ws, cn):
    rows = list(ws.iter_rows(values_only=True))
    header = [str(h) if h else "" for h in rows[0]]
    cols = {h: i for i, h in enumerate(header)}
    out = []
    for row in rows[1:]:
        if all(v is None for v in row):
            continue
        if cn:
            std = {CN_COLUMN_MAP.get(h, h): row[i] for h, i in cols.items() if h}
        else:
            std = {h: row[i] for h, i in cols.items() if h}
        out.append(std)
    return out

def main_compare():
    wb_cn = openpyxl.load_workbook(CN, read_only=True, data_only=True)
    wb_en = openpyxl.load_workbook(EN, read_only=True, data_only=True)
    pairs = list(zip(wb_cn.worksheets, wb_en.worksheets))

    total = 0
    for ws_c, ws_e in pairs:
        if not ws_c["A1"].value and not ws_e["A1"].value:
            problems = []
            rc = [r for r in ws_c.iter_rows(values_only=True) if any(v is not None for v in r)]
            re_ = [r for r in ws_e.iter_rows(values_only=True) if any(v is not None for v in r)]
            n = max(len(rc), len(re_))
            for i in range(n):
                row_c = rc[i] if i < len(rc) else ()
                row_e = re_[i] if i < len(re_) else ()
                for j in range(max(len(row_c), len(row_e))):
                    c = row_c[j] if j < len(row_c) else None
                    e = row_e[j] if j < len(row_e) else None
                    if c is None and e is None:
                        continue
                    why = []
                    if not cell_pair_ok(c, e, why):
                        problems.append((i + 1, openpyxl.utils.get_column_letter(j + 1), c, e, why[0]))
            total += len(problems)
            print(f"\n[{ws_c.title} | {ws_e.title}] CN rows={len(rc)} EN rows={len(re_)} "
                  f"problems={len(problems)}")
            if len(re_) > len(rc):
                print(f"  EXTRA EN ROW (last): {[v for v in re_[-1]][:14]}")
            for r, col, c, e, why in problems[:6]:
                print(f"  row {r} col {col}: {why}")
            if len(problems) > 6:
                print(f"  ... and {len(problems)-6} more")
            continue

        cn_rows = load_rows(ws_c, cn=True)
        en_rows = load_rows(ws_e, cn=False)
        cn_groups = defaultdict(list)
        for r in cn_rows:
            cn_groups[(str(r.get("Document") or ""), str(r.get("Announcement_Year") or ""))].append(r)
        en_groups = defaultdict(list)
        for r in en_rows:
            en_groups[(str(r.get("Document") or ""), str(r.get("Announcement_Year") or ""))].append(r)

        problems = []
        unmatched = []
        unpaired = []
        for k in sorted(set(cn_groups) | set(en_groups)):
            g1, g2 = cn_groups.get(k, []), en_groups.get(k, [])
            if not g1 or not g2:
                unmatched.append((k, len(g1), len(g2)))
                continue
            paired, left1, left2 = pair_groups(g1, g2)
            for r in left1 + left2:
                unpaired.append((k, r))
            for r1, r2 in paired:
                for col_name in sorted(set(r1) | set(r2)):
                    c, e = r1.get(col_name), r2.get(col_name)
                    if c is None and e is None:
                        continue
                    why = []
                    if not cell_pair_ok(c, e, why):
                        problems.append((k, col_name, c, e, why[0]))

        total += len(problems) + len(unmatched) + len(unpaired)
        print(f"\n[{ws_c.title} | {ws_e.title}] CN={len(cn_rows)} EN={len(en_rows)} "
              f"paired groups={len(cn_groups)} problems={len(problems)} "
              f"unmatched_groups={len(unmatched)} unpaired_rows={len(unpaired)}")
        for k, n1, n2 in unmatched[:4]:
            print(f"  GROUP ONLY ONE SIDE: {k} CN={n1} EN={n2}")
        for k, r in unpaired:
            print(f"  UNPAIRED row: {k}: {r.get('Metric')!r} mag={r.get('Target_Magnitude')!r}")
        for k, col, c, e, why in problems:
            print(f"  {k} col={col}: {why}")

    print(f"\n=== TOTAL problems: {total} ===")
    wb_cn.close(); wb_en.close()
# ---------------------------------------------------------------------------
# data: fix CN/EN data diffs (originally fix_cn_en_data.py)
# ---------------------------------------------------------------------------
CN_PATH = 'China_Climate_Target_Tracker_cn.xlsx'
EN_PATH = 'China_Climate_Target_Tracker_en.xlsx'

# 1-based 列号
C = {'metric': 2, 'mag': 4, 'baseline': 5, 'sentence': 10}

# (文件, 表名, 行号, 列, 方式, 旧值, 新值)
# 方式: set=整格覆盖; replace=子串替换(必须命中); regex=正则替换(必须命中)
DATA_FIXES = [
    # ---------- GHG Emissions (EN) ----------
    ('en', 'GHG Emissions', 42, 'sentence', 'set', None,
     'From 2014 to 2015, the carbon dioxide emissions per unit of GDP will decrease by more than 4% and 3.5% respectively in two years.'),
    ('en', 'GHG Emissions', 2, 'sentence', 'set', None,
     'The energy-saving and carbon reduction transformation of key areas and industries will result in carbon dioxide emissions reduction of about 130 million tons.'),
    ('en', 'GHG Emissions', 86, 'sentence', 'set', None,
     'From 2024 to 2025, the energy-saving and carbon-reduction transformation of the steel industry will result in energy savings of about 20 million tons of standard coal and reduce carbon dioxide emissions by about 53 million tons.'),
    ('en', 'GHG Emissions', 24, 'sentence', 'set', None,
     'Carbon dioxide emissions per unit of GDP will decline by 17 percent by 2015 compared with 2010.'),
    ('en', 'GHG Emissions', 20, 'sentence', 'replace', 'period (2011-2015), China', 'period, China'),
    ('en', 'GHG Emissions', 30, 'sentence', 'replace', 'CO emission per unit of GDP 2 ', 'CO2 emission per unit of GDP '),
    ('en', 'GHG Emissions', 21, 'sentence', 'replace', 'CO emission', 'CO2 emission'),
    ('en', 'GHG Emissions', 22, 'sentence', 'replace', 'CO emission', 'CO2 emission'),
    ('en', 'GHG Emissions', 32, 'sentence', 'replace', 'CO emission', 'CO2 emission'),
    ('en', 'GHG Emissions', 116, 'mag', 'set', None,
     'effectively control trifluoromethane (HFC-23) emissions'),
    ('en', 'GHG Emissions', 35, 'sentence', 'set', None,
     'binding target of slashing carbon intensity by 18 percent during the 14th Five-Year Plan period (2021-2025).'),
    ('en', 'GHG Emissions', 72, 'sentence', 'replace',
     '9.5 percent from 2020 levels', '9.5 percent by 2030 compared with 2020 levels'),
    ('en', 'GHG Emissions', 11, 'mag', 'set', None, 'before 2030'),
    ('en', 'GHG Emissions', 88, 'sentence', 'regex',
     r'approximately 400%\s*\n\s*10000\s+tons', 'approximately 4 million tons'),

    # ---------- Energy|Power (EN) ----------
    ('en', 'Energy|Power', 16, 'sentence', 'set', None,
     'By 2020, the average coal consumption of existing coal-fired power generation units after transformation will be less than 310 grams per kilowatt-hour, of which the average coal consumption of existing units of 600,000 kilowatts and above (except air-cooled units) after transformation will be less than 300 grams per kilowatt-hour.'),
    ('en', 'Energy|Power', 427, 'mag', 'set', None, '10 percent compared with 2015'),
    ('en', 'Energy|Power', 35, 'mag', 'set', None, 'basically eliminated'),
    ('en', 'Energy|Power', 73, 'mag', 'set', None, 'basically eliminated'),
    ('en', 'Energy|Power', 278, 'metric', 'set', None, 'new coal mines (new construction and expansion)'),
    ('en', 'Energy|Power', 278, 'mag', 'set', None, '450 million tons'),
    ('en', 'Energy|Power', 69, 'sentence', 'replace',
     'In 2015, the electricity saving capacity', 'The electricity saving capacity'),
    ('en', 'Energy|Power', 299, 'sentence', 'replace', ' by 2015.', '.'),
    ('en', 'Energy|Power', 434, 'sentence', 'set', None,
     'By 2015, total national energy consumption and electricity consumption will be controlled at around 4 billion tons of standard coal and 6.15 trillion kilowatt-hours respectively.'),
    ('en', 'Energy|Power', 185, 'sentence', 'set', None,
     'During the 13th Five-Year Plan period, about 40 million kilowatts of conventional hydropower will be newly put into operation across the country, and more than 60 million kilowatts will be started, of which about 5 million kilowatts will be small hydropower.'),
    ('en', 'Energy|Power', 209, 'sentence', 'set', None,
     'By 2020, the newly installed capacity of pumped storage power stations will increase by about 17 million kilowatts, reaching about 40 million kilowatts.'),
    ('en', 'Energy|Power', 326, 'sentence', 'set', None,
     'By 2025, China’s primary crude oil processing capacity will be controlled within 1 billion tons, and about 55% of China’s total oil refining capacity should come from large refineries that have a capacity of at least 10 million tons per year.'),
    ('en', 'Energy|Power', 386, 'sentence', 'replace',
     '10 percent by 2010 compared to 2005, the extraction',
     '10 percent by 2010, and the extraction'),
    ('en', 'Energy|Power', 179, 'sentence', 'replace',
     'reach 350 million kilowatts by 2015.', 'reach 350 million kilowatts.'),
    ('en', 'Energy|Power', 470, 'sentence', 'set', None,
     'In 2024, the national non-hydropower renewable energy electricity consumption responsibility weight will reach about 18 percent, and the scale of renewable energy non-electricity utilization will exceed 60 million tons of standard coal.'),
    ('en', 'Energy|Power', 84, 'sentence', 'replace',
     'per 10,000 yuan of GDP will drop to 2.25 tons of standard coal from',
     'per 10,000 yuan of GDP (at constant 1990 prices) will drop to 2.25 tons of standard coal from'),
    ('en', 'Energy|Power', 430, 'mag', 'set', None, '5 percent compared with 2015'),
    ('en', 'Energy|Power', 430, 'sentence', 'replace',
     ' about 5 percent compared with 2015.', ' about 5 percent.'),
    ('en', 'Energy|Power', 34, 'mag', 'set', None, 'basically eliminated'),
    ('en', 'Energy|Power', 61, 'sentence', 'replace', 'By 2010, ', ''),
    ('en', 'Energy|Power', 146, 'sentence', 'replace', 'By 2010, ', ''),
    ('en', 'Energy|Power', 195, 'sentence', 'replace', 'By 2010, ', ''),
    ('en', 'Energy|Power', 91, 'sentence', 'replace', 'reduce CO emissions.', 'reduce CO2 emissions.'),
    # 能源 CN: HL1701 行 237 政策原文错填为装机句，改为与 19.4% 目标一致
    ('cn', '能源|电力', 237, 'sentence', 'set', None,
     '到2020年，力争水电占总发电量比重达到19.4%左右。'),

    # ---------- Industry (EN) ----------
    ('en', 'Industry', 84, 'sentence', 'set', None,
     'By 2020, the annual production and sales will exceed 2 million, and the cumulative production and sales will exceed 5 million electric vehicles.'),

    # ---------- Buildings (EN) ----------
    ('en', 'Buildings', 15, 'sentence', 'set', None,
     'From 2014 to 2015, the energy consumption per unit building area of public institutions across the country will decrease by 2.2 percent annually, and efforts are made to exceed the target of a 12 percent reduction during the 12th Five-Year Plan period.'),
    ('en', 'Buildings', 75, 'sentence', 'set', None,
     'By 2025, new buildings in cities and towns will fully implement green building standards'),

    # ---------- Transport (EN) ----------
    ('en', 'Transport', 32, 'sentence', 'set', None,
     'By 2015, the maximum speed of battery electric passenger vehicles and plug-in hybrid passenger vehicles shall not be less than 100 km/h, and the comprehensive pure electric driving range shall not be less than 150 km for battery electric vehicles and 50 km for plug-in hybrid vehicles, respectively.'),
    ('en', 'Transport', 33, 'sentence', 'set', None,
     'By 2015, the maximum speed of battery electric passenger vehicles and plug-in hybrid passenger vehicles shall not be less than 100 km/h, and the comprehensive pure electric driving range shall not be less than 150 km for battery electric vehicles and 50 km for plug-in hybrid vehicles, respectively.'),
    ('en', 'Transport', 50, 'mag', 'set', None, 'strive to reach the peak'),
    ('en', 'Transport', 51, 'mag', 'set', None, 'basically eliminated'),
    ('en', 'Transport', 77, 'sentence', 'replace',
     ' annually compared with 2020.', ' annually.'),

    # ---------- LULUCF ----------
    # CN 行 140: 政策原文漏写"00"（mag=2200万公顷，EN 为 22 million）
    ('cn', '土地利用、土地利用变化和林业', 140, 'sentence', 'set', None,
     '到2010年，使2200万公顷沙化土地得到有效治理。'),
    ('en', 'LULUCF', 52, 'sentence', 'set', None,
     'The Chinese government announced that, by 2020 the forest area will be increased by 40 million hectares compared with 2005.'),
    ('en', 'LULUCF', 52, 'baseline', 'set', None, 2005),
    ('cn', '土地利用、土地利用变化和林业', 52, 'baseline', 'set', None, 2005),
    ('en', 'LULUCF', 105, 'sentence', 'replace', 'm3 14 relative', 'm3 relative'),
    ('en', 'LULUCF', 79, 'sentence', 'set', None,
     'The Outline of the 14th Five-Year Plan for National Economic and Social Development of the People’s Republic of China and the Long-Range Objectives Through the Year 2035 clearly sets the quantitative target of increasing the forest coverage rate to 24.1 percent by 2025.'),
    ('en', 'LULUCF', 91, 'sentence', 'set', None,
     'At the 2009 United Nations Climate Change Summit, President Hu Jintao made a solemn promise to the world to “vigorously increase forest carbon sinks and strive to increase forest area by 40 million hectares and forest stock by 1.3 billion cubic meters by 2020 compared with 2005.”'),

    # ---------- Circular Economy ----------
    ('cn', '循环经济', 25, 'metric', 'set', None, '再生废纸年产量'),
    ('cn', '循环经济', 25, 'mag', 'set', None, '8000万吨'),
    ('cn', '循环经济', 32, 'metric', 'set', None, '大宗固体废弃物年度利用量'),
    ('cn', '循环经济', 32, 'mag', 'set', None, '约45亿吨'),
    ('en', 'Circular Economy', 27, 'sentence', 'set', None,
     'By 2015, the annual remanufacturing capacity will reach 800,000 engines, 8 million gearboxes, starters, generators, etc., and 200,000 sets of construction machinery, mining machinery, agricultural machinery, etc., and the annual output value of the remanufacturing industry will reach about 50 billion yuan.'),
    ('en', 'Circular Economy', 28, 'sentence', 'set', None,
     'By 2015, the annual remanufacturing capacity will reach 800,000 engines, 8 million gearboxes, starters, generators, etc., and 200,000 sets of construction machinery, mining machinery, agricultural machinery, etc., and the annual output value of the remanufacturing industry will reach about 50 billion yuan.'),
    ('en', 'Circular Economy', 41, 'sentence', 'set', None,
     'By 2015, the total utilization of major renewable resources will reach 266 million tons, the output value will reach 1.2 trillion yuan, and 18 million people will be employed in the recycling industry.'),
    ('en', 'Circular Economy', 118, 'sentence', 'set', None,
     'By 2020, all county towns and key towns in the country will have sewage treatment capacity, and the sewage treatment rates of cities and county towns will reach about 95 percent and 85 percent respectively.'),
    ('en', 'Circular Economy', 180, 'sentence', 'set', None,
     'By 2025, the water consumption per unit of GDP will decrease by about 16% compared with 2020.'),
    ('en', 'Circular Economy', 50, 'sentence', 'set', None,
     'By 2020, the comprehensive utilization rate of livestock and poultry manure and waste in the country will reach more than 75 percent, and the matching rate of manure and waste treatment facilities and equipment in large-scale farms will reach more than 95 percent.'),

    # ---------- Pollution (EN) ----------
    ('en', 'Pollution', 45, 'mag', 'set', None, 'maintained at the 2005 level'),
    ('en', 'Pollution', 47, 'mag', 'set', None, '2020'),
]


def main_data():
    apply = '--apply' in sys.argv
    wbs = {
        'cn': openpyxl.load_workbook(CN_PATH),
        'en': openpyxl.load_workbook(EN_PATH),
    }
    errors = []
    n_ok = 0

    def cell(wb_key, sheet, row, col_key):
        return wbs[wb_key][sheet].cell(row=row, column=C[col_key])

    for wb_key, sheet, row, col_key, kind, old, new in DATA_FIXES:
        c = cell(wb_key, sheet, row, col_key)
        cur = c.value
        tag = f'{wb_key.upper()} {sheet} r{row} {col_key}'
        if kind == 'set':
            new_val = new
        elif kind == 'replace':
            if not isinstance(cur, str) or old not in cur:
                errors.append(f'{tag}: 未命中子串 {old!r}，当前值 {cur!r}')
                continue
            new_val = cur.replace(old, new)
        elif kind == 'regex':
            if not isinstance(cur, str) or not re.search(old, cur):
                errors.append(f'{tag}: 正则未命中 {old!r}，当前值 {cur!r}')
                continue
            new_val = re.sub(old, new, cur)
        else:
            errors.append(f'{tag}: 未知方式 {kind}')
            continue
        print(f'[{tag}]')
        print(f'  - {cur!r}')
        print(f'  + {new_val!r}')
        if apply:
            c.value = new_val
        n_ok += 1

    # Sources: EN 第 496 行 E 列存在空字符串（CN 为全空），清除以对齐行数
    sc = wbs['en']['Sources'].cell(row=496, column=5)
    print(f'[EN Sources r496 baseline] {sc.value!r} -> None')
    if apply:
        sc.value = None
        n_ok += 1
    else:
        if sc.value != '':
            errors.append(f'EN Sources r496 E 当前值 {sc.value!r}，预期空字符串')

    if errors:
        print('\n== 错误 ==')
        for e in errors:
            print(e)
        sys.exit(1)

    if apply:
        wbs['cn'].save(CN_PATH)
        wbs['en'].save(EN_PATH)
        print(f'\n已应用 {n_ok} 处修改并保存。')
    else:
        print(f'\nDry-run 完成：{n_ok} 处修改待应用（加 --apply 实际写入）。')
# ---------------------------------------------------------------------------
# category-a: A-list category fixes + EN row order (originally fix_category_and_order.py)
# ---------------------------------------------------------------------------
CN_EP = '能源|电力'
EN_EP = 'Energy|Power'

# ---------- A-list fixes: sheet -> [(excel_row, new_category)] ----------
CAT_A_CN_FIXES = {
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
CAT_A_EN_FIXES = {  # same rows, English category names; Energy|Power applied after reorder
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
CAT_A_OVERRIDES = {
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
def main_category_a():
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
    for i, j in CAT_A_OVERRIDES.items():
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
        elif i not in CAT_A_OVERRIDES and canon_col3(cnr[i][3], True) != canon_col3(enr[j][3], False):
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
    for sheet, fixes in CAT_A_CN_FIXES.items():
        ws = cn_wb[sheet]
        for row, cat in fixes:
            old = ws.cell(row=row, column=8).value
            assert old == cat_a_fixes_old_cn(sheet, row, cat), f'CN {sheet} row {row}: expected old cat, got {old}'
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
        elif i not in CAT_A_OVERRIDES and canon_col3(vals[3], False) != canon_col3(cnr[i][3], True):
            bad2 += 1
            print('POST-REORDER MISMATCH', i)
    print('post-reorder mismatches:', bad2)

    for sheet, fixes in CAT_A_EN_FIXES.items():
        w = en_wb[sheet]
        for row, cat in fixes:
            w.cell(row=row, column=8).value = cat
    en_wb.save(EN_FILE)
    en_wb.close()
    print('EN fixes applied and saved')

def cat_a_fixes_old_cn(sheet, row, cat):
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
# ---------------------------------------------------------------------------
# category-b: B-list category fixes, user-judged (originally fix_category_b.py)
# ---------------------------------------------------------------------------

CAT_B_CN_FIXES = {
    '建筑': [(23, '节能')],
    '交通': [(62, '能源效率'), (74, '污染控制'), (75, '污染控制')],
    '工业': [(139, '污染控制'), (140, '污染控制'), (164, '污染控制'), (145, '节能')],
    '土地利用、土地利用变化和林业': [(12, '环境质量'), (13, '环境质量')],
    '循环经济': [(172, '能源效率'), (187, '能源效率')],
    '能源|电力': [(489, '绿色交通'), (539, '绿色交通'), (549, '绿色交通'), (550, '绿色交通'),
                (493, '涉煤'), (495, '涉煤'), (521, '涉煤')],
}
CAT_B_EN_FIXES = {
    'Buildings': [(23, 'Energy saving')],
    'Transport': [(62, 'Energy efficiency'), (74, 'Pollution control'), (75, 'Pollution control')],
    'Industry': [(139, 'Pollution control'), (140, 'Pollution control'), (164, 'Pollution control'),
                 (145, 'Energy saving')],
    'LULUCF': [(12, 'Environmental quality'), (13, 'Environmental quality')],
    'Circular Economy': [(172, 'Energy efficiency'), (187, 'Energy efficiency')],
    'Energy|Power': [(489, 'Green transportation'), (539, 'Green transportation'),
                     (549, 'Green transportation'), (550, 'Green transportation'),
                     (493, 'Coal-related'), (495, 'Coal-related'), (521, 'Coal-related')],
}

CAT_B_OLD_CN = {
    '建筑': {23: '能源效率'},
    '交通': {62: '节能', 74: '绿色交通', 75: '绿色交通'},
    '工业': {139: '工业转型', 140: '工业转型', 164: '工业转型', 145: '工业转型'},
    '土地利用、土地利用变化和林业': {12: '碳汇', 13: '碳汇'},
    '循环经济': {172: '循环利用', 187: '循环利用'},
    '能源|电力': {489: '电气化', 539: '电气化', 549: '电气化', 550: '电气化',
                493: '能源效率', 495: '能源效率', 521: '能源效率'},
}
CAT_B_OLD_EN = {
    'Buildings': {23: 'Energy efficiency'},
    'Transport': {62: 'Energy saving', 74: 'Green transportation', 75: 'Green transportation'},
    'Industry': {139: 'Industrial transformation', 140: 'Industrial transformation',
                 164: 'Industrial transformation', 145: 'Industrial transformation'},
    'LULUCF': {12: 'Carbon sink', 13: 'Carbon sink'},
    'Circular Economy': {172: 'Recycling', 187: 'Recycling'},
    'Energy|Power': {489: 'Electrification', 539: 'Electrification', 549: 'Electrification',
                     550: 'Electrification', 493: 'Energy efficiency', 495: 'Energy efficiency',
                     521: 'Energy efficiency'},
}

def check_category_b(fname, fixes, oldmap):
    wb = openpyxl.load_workbook(fname, read_only=True)
    n = 0
    for sheet, rows in fixes.items():
        ws = wb[sheet]
        for row, cat in rows:
            old = ws.cell(row=row, column=8).value
            exp = oldmap[sheet][row]
            metric = ws.cell(row=row, column=2).value
            if old != exp:
                print(f'ASSERT FAIL {fname} {sheet} r{row} [{metric}]: got {old!r} want {exp!r}')
                wb.close()
                sys.exit(1)
            n += 1
            print(f'OK {fname} {sheet} r{row} [{str(metric)[:30]}] {old} -> {cat}')
    wb.close()
    return n

def main_category_b():
    apply = '--apply' in sys.argv
    n = 0
    n += check_category_b(CN_FILE, CAT_B_CN_FIXES, CAT_B_OLD_CN)
    n += check_category_b(EN_FILE, CAT_B_EN_FIXES, CAT_B_OLD_EN)
    print(f'{n} fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    for fname, fixes in ((CN_FILE, CAT_B_CN_FIXES), (EN_FILE, CAT_B_EN_FIXES)):
        shutil.copy(fname, fname.replace('.xlsx', '.pre_bfix.bak'))
        wb = openpyxl.load_workbook(fname)
        for sheet, rows in fixes.items():
            ws = wb[sheet]
            for row, cat in rows:
                ws.cell(row=row, column=8).value = cat
        wb.save(fname)
        wb.close()
        print(f'{fname}: applied and saved')
# ---------------------------------------------------------------------------
# direction: direction-column wording round 1 (originally fix_direction_cn.py)
# ---------------------------------------------------------------------------

# sheet -> [(excel_row, old_direction, new_direction)]
DIR1_FIXES = {
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

def check_direction1():
    wb = openpyxl.load_workbook(CN_FILE, read_only=True)
    n = 0
    for sheet, fixes in DIR1_FIXES.items():
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

def main_direction1():
    apply = '--apply' in sys.argv
    n = check_direction1()
    print(f'{n} direction fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    shutil.copy(CN_FILE, CN_FILE.replace('.xlsx', '.pre_dirfix.bak'))
    wb = openpyxl.load_workbook(CN_FILE)
    for sheet, fixes in DIR1_FIXES.items():
        ws = wb[sheet]
        for row, old, new in fixes:
            ws.cell(row=row, column=3).value = new
    wb.save(CN_FILE)
    wb.close()
    print(f'{CN_FILE}: applied and saved')
# ---------------------------------------------------------------------------
# direction2: contradictory direction+value combos (originally fix_direction2_cn.py)
# ---------------------------------------------------------------------------

# sheet -> [(excel_row, old_dir, new_dir, old_val, new_val)]
# None = column unchanged (still asserted in check_direction2())
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

def check_direction2():
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

def main_direction2():
    apply = '--apply' in sys.argv
    n = check_direction2()
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
# ---------------------------------------------------------------------------
# direction3: direction/value wording round 3 (originally fix_direction3_cn.py)
# ---------------------------------------------------------------------------

# sheet -> [(excel_row, old_dir, new_dir, old_val, new_val)]
# None = column unchanged (still asserted in check_direction3())
DIR3_FIXES = {
    '交通': [
        (2, '实现', '形成', '1400万吨标准煤', '1400万吨标准煤以上'),
        (7, '降低至', '降到', '4.5升/百公里', '4.5升/百公里以下'),
        (10, '达到', '降到', '5升/百公里', None),
        (12, '降低至', '降至', '5.0升/百公里', None),
        (13, '降低至', '降至', '6.9升/百公里', None),
        (14, '降低至', '降至', '5.0升/百公里', None),
        (15, '降低至', '降到', '5.0升/百公里', None),
        (16, '降低至', '降到', '4.0升/百公里', None),
        (32, '实现', '不低于', '80%', None),
        (33, '实现', '不低于', '80%', None),
        (34, '实现', '不低于', '60%', None),
        (35, '实现', '不低于', '60%', None),
        (38, '超过', '不低于', '70%', None),
        (41, '占比达到', '达到', '72%', None),
        (59, '达到', '建成', '2800万个', None),
        (63, '提高至', '不低于', '80%', None),
        (66, '提高', '增长', '100%', None),
        (68, '提高', '年均增长', '15%以上', None),
        (69, '提高', '年均增长', '15%', '15%以上'),
        (70, '提高', '增长', '10%', None),
        (71, '提高', '增长', '10%左右', None),
        (72, '提高', '增长', '12%', None),
        (73, '提高', '增长', '12%', '12%左右'),
    ],
    '土地利用、土地利用变化和林业': [
        (2, '提高', '新增', '1250万公顷', None),
        (4, '提高', '增加', '1250万公顷', None),
        (5, '提高', '新增', '1250万公顷', None),
        (9, '实现', '增加', '5000万吨二氧化碳当量', None),
        (10, '提高', '新增', '6000万亩', '6000万亩左右'),
        (26, '修复', '新增', '1000万公顷', '超过1000万公顷'),
        (29, '降低至', '控制在…以下', '4‰', None),
        (33, '提高至', '提高到', '0.55以上', None),
        (34, '提高至', '提高到', '0.58', None),
        (35, '提高至', '提高到', '0.55以上', None),
        (38, '提高至', '提高到', '0.55以上', None),
        (41, '提高', '增加', '4000万公顷', None),
        (42, '提高', '增加', '4000万公顷', None),
        (43, '提高', '增加', '4000万公顷', None),
        (44, '提高', '新增', '4000万公顷', None),
        (45, '提高', '增加', '4000万公顷', None),
        (46, '提高', '增加', '4000万公顷', None),
        (47, '提高', '增加', '4000万公顷', None),
        (48, '提高', '增加', '4000万公顷', None),
        (50, '建设', '建成', '200个', None),
        (51, '建设', '建成', '6个', None),
        (53, '提高至', '达到', '20%', None),
        (54, '提高', '增加', '4000万公顷', None),
        (57, '提高至', '提高到', '21.66%', None),
        (58, '实现', '达到', '21.66%', None),
        (59, '提高至', '提高到', '21.66%', None),
        (60, '提高至', '提高到', '21.66%', None),
        (62, '提高至', '提高到', '23.04%', None),
        (63, '提高至', '提高到', '21.66%', None),
        (67, '提高至', '达到', '25%', '25%左右'),
        (70, '达到', None, '25%', '约25%'),
        (73, '提高至', '提高到', '24.1%', None),
        (74, '提高至', '提高到', '23.04%', None),
        (84, '降低至', '控制在…以下', '1‰（千分之一）', None),
        (86, '提高', '增加', '13亿立方米', None),
        (87, '提高', '增加', '13亿立方米', None),
        (88, '提高', '增加', '45亿立方米', '45亿立方米左右'),
        (89, '提高', '增加', '13亿立方米', None),
        (90, '提高', '增加', '45亿立方米', '约45亿立方米'),
        (92, '提高', '增加', '6亿立方米', None),
        (95, '提高', '增加', '13亿立方米', None),
        (96, '提高', '增加', '6亿立方米', None),
        (97, '提高', '增加', '6亿立方米', None),
        (98, '提高', '增加', '6亿立方米', None),
        (99, '提高', '增加', '13亿立方米', None),
        (100, '提高', '增加', '13亿立方米', None),
        (101, '提高', '增加', '13亿立方米', None),
        (102, '提高至', '增加', '6亿立方米', None),
        (105, '提高', '增加', '13亿立方米', None),
        (106, '提高', '增加', '45亿立方米', None),
        (107, '提高至', '增加到', '165亿立方米', None),
        (108, '提高至', '增加到', '超过165亿立方米', None),
        (109, '提高至', '提高到', '190亿立方米', None),
        (111, '提高', '增加', '60亿立方米', None),
        (112, '提高', '增加', '60亿立方米', None),
        (116, '提高', '增加', '6亿亩', None),
        (117, '提高', '增加', '60亿立方米', None),
        (118, '提高至', '提升至', '超过240亿立方米', None),
        (125, '围栏封育', '达到', '5000万公顷', None),
        (126, '提高', '新增', '2400万公顷', None),
        (129, '提高', '增加', '2400万公顷', None),
        (140, '种植', '新增', '5700万公顷', None),
        (150, '建设', '建成', '360个', None),
        (153, '占比达到', '达到', '16%', None),
        (156, '达到', '建成', '90%以上', None),
        (162, '达到', '不低于', '35%', None),
        (176, '达到', None, '95亿吨', '约95亿吨'),
        (177, '提高至', '提高到', '84亿吨二氧化碳（8400 MtCO₂）', None),
        (187, '提高至', '提高到', '43%', None),
        (196, '达到', '不低于', '8亿亩', None),
        (197, '实现', '不低于', '8亿亩', None),
        (198, '提高至', '增加至', '4248万公顷', None),
    ],
    '工业': [
        (2, '降低至', '降至', '13,300千瓦时/吨', None),
        (6, '降低至', '降到', '70立方米/吨', None),
        (13, '降低至', '降到', '112千克标准煤/吨', None),
        (15, '降低', '下降', '超过3个百分点', None),
        (18, '降低至', '降至', '560千克标准煤/吨及以下', None),
        (21, '降低至', '降到', '580千克标准煤/吨', None),
        (23, '降低至', '降至', '10亿吨以下', None),
        (25, '降低', '下降', '30%以上', None),
        (32, '达到', '不低于', '92%', None),
        (36, '降低至', '降到', '1050千克标准煤/吨', None),
        (37, '降低至', '降到', '330千克标准煤/吨', None),
        (38, '降低至', '降到', '300千克标准煤/吨', None),
        (39, '降低至', '降到', '86千克标准煤/吨', None),
        (40, '降低至', '降到', '1110千克标准煤/吨', None),
        (41, '降低至', '降到', '857千克标准煤/吨', None),
        (42, '降低至', '降到', '15千克标准煤/重量箱', None),
        (43, '降低至', '降到', '530千克标准煤/吨', None),
        (44, '降低至', '降到', '370千克标准煤/吨', None),
        (52, '降低', '下降', '13.5%', None),
        (57, '降低', '下降', '10%以上（5年累计）', None),
        (60, '节约', '形成', '1000万吨标准煤', '约1000万吨标准煤'),
        (61, '节约', '形成', '500万吨标准煤', '约500万吨标准煤'),
        (62, '节约', '形成', '4000万吨标准煤', '约4000万吨标准煤'),
        (63, '节约', '形成', '约2000万吨标准煤', None),
        (65, '达到', None, '15万亿元', '约15万亿元'),
        (69, '建设', '建成', '100个', None),
        (78, '提高至', '提高到', '50%以上', None),
        (79, '控制在…以内', '控制在…左右', '18亿吨', None),
        (91, '达到', None, '40%', '约40%'),
        (92, '达到', '新增', '40%', '约40%'),
        (111, '占比达到', '超过', '30%', None),
        (116, '提高至', '提高到', '90%以上', None),
        (117, '提高至', '提高到', '90%以上', None),
        (118, '提高至', '提升至', '15%以上', None),
        (119, '提高至', '提升至', '15%', None),
        (131, '提高至', '提高到', '70%', '70%以上'),
        (137, '占比达到', '达到', '15%', '15%以上'),
        (141, '提高至', '提高到', '30%', '30%以上'),
        (142, '降低', '下降', '10%（5年累计）', None),
        (143, '实现', '达到', '70%', '70%以上'),
        (146, '占比达到', '超过', '40%', None),
        (155, '降低', '下降', '10%以上', None),
        (163, '降低', '削减', '280吨/年', None),
        (167, '降低至', '降到', '4立方米', None),
        (168, '降低', '下降', '16%', '约16%'),
        (169, '降低', '下降', '16%', '约16%'),
        (170, '降低至', '降低到', '65立方米', '65立方米以下'),
        (171, '降低至', '降至', '65立方米', None),
        (174, '降低', '减少', '30%', None),
        (175, '降低', '下降', '30%', None),
    ],
    '建筑': [
        (22, '降低', '下降', '5%', None),
        (44, '提高', '增加', '超过5000万千瓦', None),
        (46, '超过', '达到', '10亿平方米', '10亿平方米以上'),
        (48, '提高', '增长', '2000万平方米', '2000万平方米以上'),
        (49, '建设', '建成', '5000万平方米', '超过5000万平方米'),
        (52, '达到', '新增', '25亿平方米', None),
        (56, '提高至', '提高到', '40%', None),
        (57, '提高至', '提高到', '40%', None),
        (58, '提高至', '达到', '620 千瓦时', None),
        (59, '降低', '下降', '3%', None),
        (60, '占比达到', '达到', '30%', None),
        (61, '占比达到', '达到', '30%', None),
        (88, '占比达到', '达到', '8%', None),
    ],
    '循环经济': [
        (2, '完成', '新增', '1505万立方米/日', None),
        (3, '达到', '新增', '1505万立方米/日', None),
        (4, '完成', '新增', '2000万立方米/日', None),
        (5, '完成', '新增', '6.01万吨/日', None),
        (7, '降低', '减少', '30%', None),
        (10, '提高', '增长', '50%', None),
        (11, '提高', '增长', '15%', None),
        (12, '提高', '增长', '30%', None),
        (15, '降低', '达到', '4亿吨/年', None),
        (46, '达到', None, '40亿吨', '约40亿吨'),
        (47, '达到', '提升至', '约45亿吨', None),
        (53, '达到', '提高至', '90%以上', None),
        (56, '降低', '下降', '30%以上', None),
        (78, '提高', None, '45%', '约45%'),
        (86, '降低', '下降', '17.84%', None),
        (93, '超过', '达到', '24%', '24%以上'),
        (95, '达到', '提高至', '75%', '75%以上'),
        (96, '达到', None, '60%', '约60%'),
        (117, '达到', '新增', '5022万立方米/日', None),
        (129, '达到', '不低于', '70%', None),
        (132, '占比达到', '达到', '35%以上', None),
        (136, '达到', '超过', '90%', None),
        (138, '达到', '超过', '30%', None),
        (151, '提高', '增加', '26亿立方米', None),
        (152, '提高', '增加', '5亿立方米', None),
        (157, '提高', '提升', '10个百分点', None),
        (161, '提高至', '提高到', '80%以上', None),
        (181, '提高至', '提高到', '80%', None),
        (194, '降低至', '降低到', '0.5吨', None),
        (195, '降低', '下降', '23%', None),
        (196, '降低', None, '16%', '约16%'),
    ],
    '污染': [
        (2, '降低', '下降', '5%', None),
        (3, '降低', '削减', '5%', None),
        (4, '提高', '新增', '30万吨', None),
        (5, '达到', '新增', '10万吨', None),
        (7, '降低', '下降', '8%', None),
        (8, '降低', '下降', '8%', None),
        (11, '降低', '削减', '10 吨', None),
        (12, '降低', '减少', '15吨', None),
        (13, '降低', '削减', '5万吨', None),
        (14, '降低', '削减', '50万吨', None),
        (15, '降低', '削减', '100万吨', None),
        (16, '降低', '削减', '15 吨', None),
        (17, '降低', '削减', '180万吨', None),
        (18, '降低', '下降', '8%', None),
        (19, '降低', '削减', '50万吨', None),
        (20, '降低', '新增', '140万吨', None),
        (21, '降低至', '减少到', '1273万吨', None),
        (25, '降低', '下降', '8%', None),
        (26, '提高', '新增', '200万吨', None),
        (28, '降低至', '下降到', '28微克/立方米', '28微克/立方米以下'),
        (29, '降低至', '下降到', '25微克/立方米', '25微克/立方米以下'),
        (31, '降低', '减少', '10%', None),
        (32, '降低', '减少', '10%', None),
        (33, '降低', '下降', '10%以上', None),
        (34, '降低', '减少', '10%', None),
        (35, '降低至', '下降到', '25微克/立方米以下', None),
        (37, '降低', '下降', '10%', '10%以上'),
        (38, '提高', '新增', '260万吨', '260万吨以上'),
        (39, '降低', '减少', '10%', None),
        (41, '降低', '下降', '7%', None),
        (43, '保持稳定', '保持在', '维持2005年水平', '2005年水平'),
        (44, '降低至', '下降至', '1.5克/千瓦时', None),
        (49, '降低', '下降', '5%以上', None),
        (50, '降低', '下降', '10%以上', None),
        (51, '降低', '下降', '10%', '10%以上'),
        (52, '降低', '减少', '25%', None),
        (54, '降低', '下降', '18%', '18%以上'),
        (55, '降低', '下降', '10%', None),
        (56, '降低', '下降', '30%以上', None),
        (57, '降低', '减少', '30%', '30%以上'),
        (60, '控制在…以内', '控制在', '2086.4万吨', None),
        (63, '降低至', '减少到', '2295万吨', None),
        (64, '降低', '减少', '240万吨', None),
        (65, '降低', '减少', '40万吨', None),
        (67, '提高', '新增', '230万吨', '230万吨以上'),
        (68, '实现', '形成', '40万吨', None),
        (69, '降低至', '下降到', '1.5克/千瓦时', None),
        (70, '降低至', '降至', '≤0.68千克/吨', '0.68千克/吨以下'),
        (71, '降低', '下降', '30%', '30%以上'),
        (73, '降低', '下降', '25%', '25%以上'),
        (75, '降低', '下降', '65%', None),
        (78, '降低', '下降', '15%以上', None),
        (79, '降低', '下降', '10%以上', None),
        (80, '降低', '下降', '10%以上', None),
        (81, '降低', '下降', '10%', '10%以上'),
    ],
    '温室气体排放': [
        (2, '实现', '减排', '1.3亿吨', '约1.3亿吨'),
        (3, '实现', '减排', '1.3亿吨', '约1.3亿吨'),
        (4, '实现', '减排', '2亿吨以上', None),
        (5, '降低', '下降', '超过15%', None),
        (17, '降低', '下降', '40%-45%', None),
        (18, '降低', '下降', '40%-45%', None),
        (19, '降低', '下降', '17%', None),
        (20, '降低', '下降', '40%-45%', None),
        (23, '降低', '下降', '40%-45%', None),
        (24, '降低', '下降', '40%-45%', None),
        (25, '降低', '下降', '17%', None),
        (26, '降低', '下降', '18%', None),
        (29, '降低', '下降', '40%-45%', None),
        (30, '降低', '下降', '18%', None),
        (32, '降低', '下降', '60%-65%', None),
        (33, '降低', '下降', '18%', None),
        (34, '降低', '下降', '60%-65%', None),
        (35, '降低', '下降', '65%以上', None),
        (37, '降低', '下降', '超过65%', None),
        (38, '降低', '下降', '18%', None),
        (39, '降低', '下降', '65%以上', None),
        (40, '降低', '下降', '18%', None),
        (41, '降低', '下降', '65%以上', None),
        (42, '降低', '下降', '40%-45%', None),
        (43, '降低', '下降', '4%', None),
        (44, '降低', '下降', '17%', None),
        (48, '降低', '下降', '18%', None),
        (51, '降低', '下降', '65%以上', None),
        (52, '降低', '下降', '60%-65%', None),
        (57, '降低', '下降', '3.5%左右', None),
        (58, '降低', '下降', '超过3.5%', '约超过3.5%'),
        (62, '降低', '下降', '22%', None),
        (63, '降低', '下降', '50%左右', None),
        (64, '降低', '下降', '21%以上', None),
        (65, '达到', '高于', '4%以上', '4%'),
        (66, '降低', '下降', '18%', '18%以上'),
        (67, '降低', '下降', '17%', '17%以上'),
        (68, '降低', '下降', '18%以上', None),
        (70, '降低', '下降', '22%', None),
        (71, '降低', '下降', '18%', None),
        (72, '降低', '下降', '18%', None),
        (73, '降低', '下降', '18%', '18%以上'),
        (76, '降低', '下降', '9.5%', '约9.5%'),
        (77, '降低', '下降', '2.6%', None),
        (78, '降低', '下降', '7%', None),
        (79, '降低', '下降', '8%', None),
        (80, '降低', '下降', '5%', None),
        (81, '降低', '下降', '3.5%', None),
        (83, '降低', '下降', '约9.5%', None),
        (84, '降低', '下降', '约9.5%', None),
        (85, '降低', '下降', '12.5%', None),
        (86, '达到', '减排', '2600万吨二氧化碳', '约2600万吨二氧化碳'),
        (87, '达到', '减排', '约1300万吨', None),
        (88, '达到', '减排', '1300万吨', '约1300万吨'),
        (89, '达到', '减排', '1.1亿吨', '约1.1亿吨'),
        (90, '达到', '减排', '5300万吨二氧化碳', '约5300万吨二氧化碳'),
        (91, '降低', '减少', '6200万吨', None),
        (92, '节约', '减少', '6亿吨二氧化碳', '约6亿吨二氧化碳'),
        (93, '节约', '减少', '约12亿吨二氧化碳', None),
        (96, '降低', '下降', '9.5%', '约9.5%'),
        (97, '降低', '下降', '约9.5%', '9.5%左右'),
        (99, '降低', '下降', '8.5%', None),
        (101, '降低', '减少', '35%', None),
        (102, '降低', '削减', '67.5%', None),
        (105, '降低', '减少', '35%', None),
        (106, '降低', '下降', '7%-10%', None),
        (108, '达到', '减少', '3000万吨二氧化碳当量', '约3000万吨二氧化碳当量'),
        (111, '达到', '减少', '5000万吨二氧化碳', '约5000万吨二氧化碳'),
        (112, '达到', '减排', '1.1亿吨二氧化碳', '约1.1亿吨二氧化碳'),
        (113, '达到', '减少', '60 Mt CO2', '约6000万吨二氧化碳'),
        (114, '降低', '下降', '40%-45%', None),
        (115, '降低', '下降', '40%-45%', None),
        (116, '降低', '下降', '40%-45%', None),
        (117, '降低', '削减', '73.2%', None),
        (118, '降低', '削减', '97.5%', None),
        (119, '降低', '削减', '67.5%', None),
        (120, '降低', '削减', '97.5%', None),
        (121, '降低', '削减', '10%', None),
        (122, '降低', '削减', '10%', None),
    ],
    '能源|电力': [
        (6, '降低', '降至', '62%', '62%以内'),
        (7, '达到', '控制在…左右', '40亿吨标准煤', None),
        (13, '降低', '下降', '10%', '10%左右'),
        (15, '建设', '建成', '100个', None),
        (18, '提高至', '提高到', '30%以上', None),
        (21, '超过', '高于', '95%', None),
        (22, '超过', '高于', '95%', None),
        (23, '超过', '高于', '95%', None),
        (29, '控制在…以内', '控制在…左右', '5%', None),
        (30, '超过', '高于', '88%', None),
        (31, '超过', '高于', '90%', None),
        (32, '超过', '达到', '95%', None),
        (35, '超过', '高于', '95%', None),
        (36, '超过', '达到', '95%', '95%以上'),
        (40, '提高至', '控制在…左右', '6.15万亿千瓦时', None),
        (45, '建设', '建成', '约1000个项目', None),
        (46, '创建', '建成', '10个示范区', None),
        (50, '降低', '下降', '10%左右', None),
        (51, '实现', '形成', '2.4亿吨标准煤', None),
        (56, '降低', '下降', '16%', None),
        (57, '降低', '下降', '16%', None),
        (59, '降低', '下降', '13.5%', None),
        (62, '降低', '下降', '13.5%', None),
        (66, '降低至', '下降到', '2.25吨标准煤（按1990年不变价计算）', None),
        (67, '降低至', '下降到', '0.869吨标准煤（按2005年价格计算）', None),
        (69, '降低至', '下降到', '1.54吨标准煤', None),
        (70, '降低', '下降', '16%', None),
        (71, '降低至', '下降至', '1.54吨标准煤（按1990年不变价计算）', None),
        (74, '降低', '下降', '5.2%左右', '5.2%'),
        (78, '降低', '下降', '16%', None),
        (79, '降低', '下降', '3.9%以上', None),
        (80, '降低', '下降', '16%', None),
        (81, '降低', '下降', '16%', None),
        (82, '降低', '下降', '13.5%', None),
        (84, '降低', '下降', '13.5%', None),
        (86, '降低', '下降', '10%左右', None),
        (89, '降低', '下降', '3.1%以上', None),
        (90, '降低', '下降', '3.4%以上', None),
        (91, '降低', '下降', '3.4%以上', None),
        (92, '降低', '下降', '3%左右', None),
        (95, '降低', '下降', '3.5%', None),
        (96, '降低', '下降', '3.5%左右', None),
        (97, '降低', '下降', '21%', '21%左右'),
        (106, '建设', '建成', '5000公里', '约5000公里'),
        (112, '降低', '下降', '10%', '10%以上'),
        (113, '超过', '达到', '6000万吨标准煤', '6000万吨标准煤以上'),
        (115, '降低', '下降', '10%', None),
        (116, '降低', '下降', '10%', None),
        (118, '降低', '下降', '10%', None),
        (119, '提高至', '提升至', '超过7000万吨标准煤', None),
        (128, '提高至', '达到', '2.5亿人', None),
        (129, '提高', '达到', '10%', '10%以上'),
        (130, '发展', '提高到', '10%', '10%左右'),
        (131, '占比达到', '提高到', '7.5%', None),
        (133, '提高至', '达到', '5600万千瓦', None),
        (134, '达到', '提高至', '10%', '10%左右'),
        (135, '占比达到', '达到', '5.3%', None),
        (136, '提高至', '提高到', '约10%', '10%左右'),
        (139, '提高至', '达到', '1565亿立方米', None),
        (140, '提高至', '增加至', '45吉瓦', None),
        (142, '提高至', '达到', '2100万千瓦', None),
        (143, '实现', '达到', '1亿千瓦', None),
        (144, '提高至', '提高到', '超过1.1亿千瓦', None),
        (145, '提高至', '达到', '6000万千瓦', None),
        (147, '提高至', '提升至', '500万千瓦', None),
        (149, '建立', '建成', '1000个示范村', None),
        (153, '提高', '提升', '4000万千瓦', '4000万千瓦以上'),
        (155, '提高', '增加', '1000万吨', None),
        (156, '提高', '增加', '2亿吨', None),
        (158, '达到', '新增', '5.5万亿立方米', None),
        (159, '开工建设', '开工', '6000万千瓦', '6000万千瓦以上'),
        (160, '达到', '年均新增', '10亿吨', '10亿吨左右'),
        (164, '实现', '达到', '2.2%', None),
        (167, '提高至', '增长至', '41亿吨', None),
        (168, '提高至', '增至', '177亿立方米', None),
        (169, '提高至', '增加至', '68亿立方米', None),
        (173, '完成', '开工', '约6000万千瓦', None),
        (174, '完成', '新增', '新增投产装机容量1亿千瓦左右', '1亿千瓦左右'),
        (175, '完成', '开工', '6000万千瓦', '约6000万千瓦'),
        (176, '完成', '新增', '约1700万千瓦', None),
        (177, '完成', '开工', '6000万千瓦', '约6000万千瓦'),
        (187, '超过', '达到', '1.8亿千瓦', '1.8亿千瓦以上'),
        (188, '实现', '形成', '5700万吨标准煤', None),
        (189, '达到', '新增', '2000万千瓦', None),
        (193, '达到', '建成', '2100万千瓦', '2100万千瓦以上'),
        (195, '开工建设', '开工', '500万千瓦', '500万千瓦左右'),
        (196, '达到', '新增', '3.5万亿立方米', None),
        (197, '发现', '新增', '12亿吨', '12亿吨左右'),
        (199, '投产', None, '3000万千瓦', '约3000万千瓦'),
        (200, '投产', None, '4000万千瓦', '约4000万千瓦'),
        (201, '发现', '新增', '4200亿立方米', None),
        (202, '提高', '新增', '1万亿立方米', None),
        (203, '建设', '建成', '16亿立方米', None),
        (204, '安装', '新增', '4000万千瓦', None),
        (205, '超过', '新增', '65亿吨', '65亿吨以上'),
        (209, '开工建设', '开工', '4.5亿吨', None),
        (219, '占比达到', '达到', '8%', '8%以上'),
        (224, '提高至', '达到', '4000万千瓦', None),
        (225, '实现', '达到', '5800万千瓦', None),
        (229, '降低', '下降', '3%', '3%以上'),
        (231, '提高', '新增', '5000万千瓦', None),
        (232, '提高至', '提高到', '70%以上', None),
        (234, '占比达到', '达到', '19.4%', '19.4%左右'),
        (235, '实现', '达到', '3.5亿千瓦', '3.5亿千瓦左右'),
        (236, '提高至', '提升至', '3.4亿千瓦', None),
        (237, '提高至', '达到', '2.9亿千瓦', None),
        (238, '实现', '达到', '3.4亿千瓦', None),
        (242, '实现', '达到', '40万千瓦', None),
        (252, '降低至', '下降到', '323克/千瓦时', None),
        (253, '降低至', '降到', '325克标准煤/千瓦时', None),
        (256, '降低至', '下降到', '63千克标准油/吨', None),
        (258, '建设', '建成', '2-3个基地', None),
        (263, '提高至', '提高到', '100亿立方米', None),
        (264, '发现', '新增', '4200亿立方米', None),
        (265, '提高', '增加', '1万亿立方米', None),
        (271, '降低至', '降低到', '约65%', '65%左右'),
        (275, '控制在…以内', '控制在…左右', '42亿吨', None),
        (281, '创建和维护', '形成', '1500万千瓦', None),
        (284, '达到', '不低于', '3.5亿千瓦', None),
        (286, '实施', None, '3.4亿千瓦', '约3.4亿千瓦'),
        (288, '实施', None, '4.2亿千瓦', '约4.2亿千瓦'),
        (290, '达到', '不低于', '3.5亿千瓦', None),
        (292, '提高至', '达到', '9.6亿千瓦', None),
        (295, '提高至', '提高到', '75%', None),
        (304, '占比达到', '达到', '28%', None),
        (309, '提高至', '提高到', '1500万千瓦', None),
        (313, '提高', '提升', '4600万千瓦', None),
        (314, '提高至', '提高到', '60%以上', None),
        (316, '提高至', '提高到', '55%以上', None),
        (318, '降低至', '降至', '6.3%', None),
        (321, '提高', None, '1.5个百分点', '约1.5个百分点'),
        (327, '达到', '新增', '约50亿吨', None),
        (328, '超过', '形成', '4000万吨', '4000万吨以上'),
        (331, '达到', None, '8200万千瓦', '约8200万千瓦'),
        (332, '建设', '建成', '200个', None),
        (333, '发展', '建成', '200个', None),
        (334, '提高至', '建成', '200个', None),
        (336, '提高至', '控制在…左右', '40亿吨标准煤', None),
        (341, '控制在…以内', '控制在…左右', '48亿吨标准煤', None),
        (343, '提高至', '提高到', '38%', None),
        (345, '实现', '形成', '800亿千瓦时', None),
        (346, '达到', '形成', '4亿吨标准煤', None),
        (352, '实现', '形成', '3亿吨标准煤', None),
        (353, '形成', None, '3亿吨标准煤', '约3亿吨标准煤'),
        (354, '降低', '下降', '2.1%', None),
        (355, '降低', '下降', '6.0%', None),
        (356, '降低', '下降', '15%', None),
        (357, '降低', '下降', '6.8%', None),
        (359, '降低', '下降', '10%', None),
        (361, '淘汰', None, '2000万千瓦', '约2000万千瓦'),
        (365, '提高', '新增', '1.3亿千瓦', None),
        (367, '提高', '增加', '3.5亿吨标准煤以上', None),
        (369, '达到', '建成', '3000个', None),
        (371, '提高', '新增', '1.3亿千瓦', '1.3亿千瓦左右'),
        (373, '提高', '开工', '6000万千瓦', '6000万千瓦以上'),
        (382, '实现', '形成', '5000万吨标准煤', '约5000万吨标准煤'),
        (383, '实现', '形成', '5000万吨标准煤', '约5000万吨标准煤'),
        (384, '降低', '下降', '5%', None),
        (385, '降低', '下降', '5%', None),
        (387, '提高至', '提高到', '10%', None),
        (388, '提高至', '提高到', '15%', '15%左右'),
        (389, '提高至', '达到', '11.4%', None),
        (390, '占比达到', '达到', '15%', '15%左右'),
        (391, '占比达到', '达到', '11.4%', None),
        (393, '提高至', '提高到', '11.4%', None),
        (399, '达到', '提高至', '15%以上', None),
        (402, '提高至', '提高到', '约15%', '15%左右'),
        (409, '达到', '提高到', '25%', '25%左右'),
        (410, '达到', '提升至', '25%左右', None),
        (417, '提高至', '提高到', '约25%', '25%左右'),
        (419, '提高至', '达到', '15%', None),
        (420, '占比达到', '达到', '11.4%', None),
        (421, '提高至', '提高到', '13%', None),
        (427, '提高至', '提高到', '18.3%', '18.3%左右'),
        (430, '提高至', '提高到', '4.5%', None),
        (431, '占比达到', '达到', '11.4%', None),
        (432, '提高至', '提高到', '11.4%', None),
        (434, '提高至', '提高到', '51.9%', '51.9%左右'),
        (435, '占比达到', '达到', '30%', None),
        (438, '占比达到', '超过', '50%', None),
        (441, '提高至', '达到', '4.7亿吨标准煤', None),
        (448, '提高', '增加', '6000亿立方米', None),
        (457, '提高至', '达到', '15.3%', None),
        (465, '实现', '达到', '12亿千瓦以上', None),
        (468, '提高至', '达到', '1亿千瓦', None),
        (492, '达到', '新增', '8000万千瓦以上', None),
        (509, '达到', '新增', '3亿千瓦以上', None),
        (510, '达到', '新增', '3.7亿千瓦以上', None),
        (511, '达到', '开工', '1亿千瓦左右', None),
        (514, '达到', '新增', '3亿千瓦以上', None),
        (519, '达到', '新增', '1亿千瓦左右', None),
        (520, '提高', '年均提升', '年均提升约1个百分点', '约1个百分点'),
        (526, '达到', '新增', '8000万千瓦以上', None),
        (533, '提高至', '提升至', '87%', None),
        (534, '提高至', '提升至', '75%', None),
        (554, '达到', '新增', '约1.6亿千瓦', None),
        (559, '达到', '新增', '约4000万千瓦', None),
        (560, '达到', '新增', '约8000万千瓦', None),
    ],
    '金融': [
        (2, '达到', None, '24万亿元人民币（基于2015年价格水平）', '约24万亿元人民币（基于2015年价格水平）'),
        (3, '达到', None, '56万亿元人民币（基于2015年价格水平）', '约56万亿元人民币（基于2015年价格水平）'),
        (6, '投入', None, '32万亿元人民币（基于2015年价格水平）', '约32万亿元人民币（基于2015年价格水平）'),
        (9, '需要', None, '32万亿元人民币（按2015年价格水平计）', '约32万亿元人民币（按2015年价格水平计）'),
        (10, '达到', '需要', '13万亿元人民币', '约13万亿元人民币'),
        (14, '达到', None, '3.9万亿元', '约3.9万亿元'),
        (15, '达到', None, '18.4万亿元', '约18.4万亿元'),
        (17, '达到', None, '57.6万亿元人民币', '约57.6万亿元人民币'),
        (18, '达到', None, '6.3万亿元人民币', '约6.3万亿元人民币'),
        (19, '达到', None, '31.1万亿元', '约31.1万亿元'),
    ],
}

def check_direction3():
    wb = openpyxl.load_workbook(CN_FILE, read_only=True)
    n = 0
    for sheet, fixes in DIR3_FIXES.items():
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
            d = f'{old_dir} -> {new_dir}' if new_dir is not None else 'dir ok'
            v = f'{old_val} -> {new_val}' if new_val is not None else 'val ok'
            print(f'OK {sheet} r{row} [{str(metric)[:26]}] {d} | {v}')
    wb.close()
    return n

def main_direction3():
    apply = '--apply' in sys.argv
    n = check_direction3()
    print(f'{n} direction/value fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    shutil.copy(CN_FILE, CN_FILE.replace('.xlsx', '.pre_dirfix3.bak'))
    wb = openpyxl.load_workbook(CN_FILE)
    for sheet, fixes in DIR3_FIXES.items():
        ws = wb[sheet]
        for row, old_dir, new_dir, old_val, new_val in fixes:
            if new_dir is not None:
                ws.cell(row=row, column=3).value = new_dir
            if new_val is not None:
                ws.cell(row=row, column=4).value = new_val
    wb.save(CN_FILE)
    wb.close()
    print(f'{CN_FILE}: applied and saved')
# ---------------------------------------------------------------------------
# direction3b: direction/value wording round 3b (originally fix_direction3b_cn.py)
# ---------------------------------------------------------------------------

# sheet -> [(excel_row, old_dir, new_dir, old_val, new_val)]
# None = column unchanged (still asserted in check_direction3b())
DIR3B_FIXES = {
    '温室气体排放': [
        (27, '降低', '下降', '3.9%以上', None),
        (46, '降低', '下降', '3.5%', '3.5%以上'),
        (69, '降低', '下降', '20%', None),
        (74, '降低', '下降', '20%', None),
        (110, '达到', '减少', '5亿吨二氧化碳', '约5亿吨二氧化碳'),
    ],
    '能源|电力': [
        (4, '达到', '建成', '20万公里', None),
        (8, '降低', '下降', '比2015年下降5%左右', '5%左右'),
        (14, '降低', '削减', '1亿吨', None),
        (34, '超过', '达到', '95%', '95%以上'),
        (49, '降低', '下降', '较2015年下降10%', '10%'),
        (61, '降低', '下降', '15%', None),
        (64, '降低', '下降', '15%', None),
        (68, '降低', '下降', '15%', None),
        (72, '降低', '下降', '20%', None),
        (73, '降低', '下降', '约20%', None),
        (75, '降低', '下降', '20%', '20%左右'),
        (76, '降低', '下降', '约20%', None),
        (83, '降低', '下降', '15%', None),
        (104, '达到', None, '2亿吨', '2亿吨左右'),
        (125, '占比达到', None, '约95%', '95%左右'),
        (150, '发展', '建成', '1000个示范村', None),
        (198, '实现', '新增', '50亿吨', '50亿吨左右'),
        (250, '淘汰', None, '2000万千瓦', '超过2000万千瓦'),
        (251, '降低', '下降', '2%', None),
        (319, '提高至', '上升至', '27%', None),
        (320, '占比达到', None, '30%', '30%左右'),
        (344, '投产', '投运', '3.55亿千瓦', None),
        (348, '相当于', None, '2100万吨二氧化硫排放', '减少二氧化硫排放2100万吨'),
        (392, '占比达到', '达到', '10%', None),
        (397, '提高至', '达到', '20%左右', None),
        (400, '提高至', '提升至', '15%', None),
        (401, '达到', '提高至', '15%', None),
        (403, '提高至', '提高到', '20%', '20%左右'),
        (412, '达到', None, '20%', '20%左右'),
        (415, '达到', None, '25%', '约25%左右'),
        (418, '提高至', '达到', '10%', None),
        (423, '提高至', '提高到', '20%左右', None),
        (437, '提高', '提高到', '10%', None),
        (467, '超过', '突破', '1亿千瓦', None),
        (543, '保持', '保持在', '90%左右', None),
    ],
    '工业': [
        (51, '降低', '下降', '20%', None),
        (98, '推广', None, '100项', '超过100项'),
        (138, '占比达到', '超过', '20%', None),
        (150, '改造', '超过', '80%以上', '80%'),
    ],
    '建筑': [
        (28, '完成', None, '3.5亿平方米', '超过3.5亿平方米'),
        (29, '提高', '增长', '2亿平方米以上', None),
        (43, '降低', '下降', '20%', None),
        (51, '提高', '增长', '两千万平方米', '2000万平方米以上'),
    ],
    '交通': [
        (18, '达到', '实现', '2000万吨商品燃油', None),
        (30, '达到', '实现', '1500万吨商品燃油', None),
        (39, '占比达到', '达到', '35%', None),
        (40, '占比达到', '达到', '20%', None),
    ],
    '土地利用、土地利用变化和林业': [
        (8, '实现', '增加', '5000万吨二氧化碳', None),
        (113, '提高至', '跃升至', '165亿立方米', None),
        (124, '修复', '改良', '2000万公顷', None),
        (127, '修复', '改良', '1500万公顷', None),
        (138, '种植', '达到', '700万公顷', None),
        (139, '种植', '达到', '350万公顷', None),
        (141, '种植', '完成', '3000万公顷', None),
        (146, '超过', '不少于', '400公里', None),
        (151, '提高至', '提高到', '55%以上', None),
        (160, '达到', '恢复', '5200万公顷', None),
        (161, '提高', '恢复', '5200万公顷', None),
        (178, '实现', None, '50%', '超过50%'),
        (179, '治理', '达', '1000万公顷以上', None),
        (180, '种植', '达到', '260亿株', '260亿株以上'),
        (181, '种植', '达到', '120亿株', '120亿株以上'),
        (184, '提高', '提高到', '40%以上', None),
        (199, '超过', '不低于', '20,000 公顷', '2万公顷'),
    ],
    '循环经济': [
        (6, '完成', None, '60100 吨/日', '6.01万吨/日'),
        (9, '达到', '替代', '13亿吨', None),
        (24, '超过', '突破', '2000万吨', None),
        (63, '占比达到', '达到', '65%', '65%左右'),
        (64, '节约', '实现', '31亿立方米', None),
        (70, '降低至', '降至', '≤3.2立方米/吨', '3.2立方米/吨及以下'),
        (198, '节约', '实现', '10亿立方米', None),
    ],
    '污染': [
        (6, '降低', '下降', '2%', None),
        (24, '降低', '下降', '2%', None),
        (27, '降低', '下降', '20%', None),
        (40, '降低', '减少', '20%', None),
        (46, '保持稳定', '保持在', None, '2005年水平'),
        (62, '降低', '下降', '2%', None),
        (66, '降低', '下降', '15%', '15%以上'),
        (76, '降低', '下降', '30%', None),
        (77, '降低', '下降', '20%', None),
    ],
}


def check_direction3b():
    wb = openpyxl.load_workbook(CN_FILE, read_only=True)
    n = 0
    errors = []
    for sheet, fixes in DIR3B_FIXES.items():
        ws = wb[sheet]
        for row, old_dir, new_dir, old_val, new_val in fixes:
            cur_dir = ws.cell(row=row, column=3).value
            cur_val = ws.cell(row=row, column=4).value
            if old_dir is not None and cur_dir != old_dir:
                errors.append(f'{sheet} r{row}: dir {cur_dir!r} != expected {old_dir!r}')
            if old_val is not None and cur_val != old_val:
                errors.append(f'{sheet} r{row}: val {cur_val!r} != expected {old_val!r}')
            if old_val is None and cur_val is not None:
                errors.append(f'{sheet} r{row}: val {cur_val!r} != expected None')
            n += 1
    wb.close()
    for e in errors:
        print('MISMATCH:', e)
    if errors:
        sys.exit(1)
    return n


def main_direction3b():
    apply = '--apply' in sys.argv
    n = check_direction3b()
    print(f'{n} direction/value fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    shutil.copy(CN_FILE, CN_FILE.replace('.xlsx', '.pre_dirfix3b.bak'))
    wb = openpyxl.load_workbook(CN_FILE)
    for sheet, fixes in DIR3B_FIXES.items():
        ws = wb[sheet]
        for row, old_dir, new_dir, old_val, new_val in fixes:
            if new_dir is not None:
                ws.cell(row=row, column=3).value = new_dir
            if new_val is not None:
                ws.cell(row=row, column=4).value = new_val
    wb.save(CN_FILE)
    wb.close()
    print(f'{CN_FILE}: applied and saved')
# ---------------------------------------------------------------------------
# rename-categories: drop 目标/target suffix (originally rename_categories.py)
# ---------------------------------------------------------------------------
CN_OLD_TO_NEW = {
    "碳减排目标": "碳减排",
    "碳强度目标": "碳强度",
    "碳汇目标": "碳汇",
    "节能目标": "节能",
    "能源效率目标": "能源效率",
    "资源效率目标": "资源效率",
    "可再生能源目标": "可再生能源",
    "电气化目标": "电气化",
    "过渡能源目标": "过渡能源",
    "涉煤目标": "涉煤",
    "生产导向目标": "生产导向",
    "工业转型目标": "工业转型",
    "技术创新目标": "技术创新",
    "污染控制目标": "污染控制",
    "环境质量目标": "环境质量",
    "循环利用目标": "循环利用",
    "绿色交通目标": "绿色交通",
    "气候金融目标": "气候金融",
}

EN_OLD_TO_NEW = {
    "Carbon intensity target": "Carbon intensity",
    "Carbon reduction target": "Carbon reduction",
    "Carbon sink target": "Carbon sink",
    "Climate finance target": "Climate finance",
    "Coal-related target": "Coal-related",
    "Electrification target": "Electrification",
    "Energy efficiency target": "Energy efficiency",
    "Energy saving target": "Energy saving",
    "Environmental quality target": "Environmental quality",
    "Green transportation target": "Green transportation",
    "Industrial transformation target": "Industrial transformation",
    "Industrial transformation targets": "Industrial transformation",
    "Pollution control target": "Pollution control",
    "Production-oriented target": "Production-oriented",
    "Production-oriented targets": "Production-oriented",
    "Recycling target": "Recycling",
    "Renewable energy target": "Renewable energy",
    "Resource efficiency target": "Resource efficiency",
    "Technology innovation target": "Technology innovation",
    "Transitional energy target": "Transitional energy",
}

RC_CHANGELOG = {
    "CN": "2026-09-07：简化目标类别名称，删除“目标”后缀（如“可再生能源目标”→“可再生能源”），中英文同步。",
    "EN": '2026-09-07: Simplified target category names by removing the "target" suffix (e.g. "Renewable energy target" → "Renewable energy"), in sync with the Chinese version.',
}


def rename_values(wb: openpyxl.Workbook, old_to_new: dict[str, str]) -> int:
    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c, MergedCell):
                    continue
                v = c.value
                if not isinstance(v, str):
                    continue
                if v.startswith("="):
                    new_v = v
                    for old, new in old_to_new.items():
                        new_v = new_v.replace(f'"{old}"', f'"{new}"')
                    if new_v != v:
                        c.value = new_v
                        changed += 1
                elif v.strip() in old_to_new:
                    c.value = old_to_new[v.strip()]
                    changed += 1
    return changed


def rc_add_changelog(wb: openpyxl.Workbook, sheet_name: str, text: str,
                     date: datetime.datetime = datetime.datetime(2026, 9, 7)) -> None:
    ws = wb[sheet_name]
    # changelog sits right below the "Last update" row (row 5); newest first
    for row in ws.iter_rows(min_col=2, max_col=2):
        if row[0].value == text:
            return  # already present (idempotent re-rc_run)
    ws.insert_rows(6)
    src = ws.cell(row=7, column=2)  # previous changelog entry, shifted down
    dst = ws.cell(row=6, column=2)
    dst.value = text
    dst.font = copy(src.font)
    dst.alignment = copy(src.alignment)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = src.number_format
    ws.cell(row=5, column=3).value = date


def rc_run(path: str, old_to_new: dict[str, str], lang: str) -> None:
    backup = path.replace(".xlsx", ".pre_category_rename.bak")
    shutil.copy2(path, backup)
    print(f"backup -> {backup}")

    wb = openpyxl.load_workbook(path)
    changed = rename_values(wb, old_to_new)
    rc_add_changelog(wb, "说明" if lang == "CN" else "README", RC_CHANGELOG[lang])
    wb.save(path)
    print(f"{path}: {changed} cells renamed, changelog added")


def main_rename_categories():
    rc_run("China_Climate_Target_Tracker_cn.xlsx", CN_OLD_TO_NEW, "CN")
    rc_run("China_Climate_Target_Tracker_en.xlsx", EN_OLD_TO_NEW, "EN")
# ---------------------------------------------------------------------------
# restructure-taxonomy: apply the Target Category Revision Rationale
# (18 -> 16 categories, CN and EN in sync, ~553 rows relabeled):
#   环境质量 -> 生态保护 | delete 技术创新 (rows -> 能源效率)
#   涉煤 + 过渡能源 -> 化石能源 (12 nuclear rows -> 非化石能源)
#   可再生能源 -> 非化石能源
# Updates the category column of all data sheets, the category list in the
# 说明/README legend, and the by-category count block of 概览/Overview.
# ---------------------------------------------------------------------------
CN_DATA_SHEETS = ['温室气体排放', '能源|电力', '工业', '建筑', '交通',
                  '土地利用、土地利用变化和林业', '循环经济', '污染', '金融']
EN_DATA_SHEETS = ['GHG Emissions', 'Energy|Power', 'Industry', 'Buildings',
                  'Transport', 'LULUCF', 'Circular Economy', 'Pollution', 'Finance']

CN_TXN = {
    'rename': {'环境质量': '生态保护', '技术创新': '能源效率', '可再生能源': '非化石能源'},
    'merge': ({'过渡能源', '涉煤'}, '化石能源', '核电', '非化石能源'),
}
EN_TXN = {
    'rename': {'Environmental quality': 'Ecological protection',
               'Technology innovation': 'Energy efficiency',
               'Renewable energy': 'Non-fossil energy'},
    'merge': ({'Transitional energy', 'Coal-related'}, 'Fossil energy',
              'nuclear', 'Non-fossil energy'),
}

CN_LEGEND_OLD = ['碳减排', '碳强度', '碳汇', '节能', '能源效率', '资源效率',
                 '可再生能源', '电气化', '过渡能源', '涉煤', '生产导向',
                 '工业转型', '技术创新', '污染控制', '环境质量', '循环利用',
                 '绿色交通', '气候金融']
CN_LEGEND_NEW = ['碳减排', '碳强度', '碳汇', '节能', '能源效率', '资源效率',
                 '非化石能源', '电气化', '化石能源', '生产导向', '工业转型',
                 '污染控制', '生态保护', '循环利用', '绿色交通', '气候金融']
EN_LEGEND_OLD = ['Carbon reduction', 'Carbon intensity', 'Carbon sink',
                 'Energy saving', 'Energy efficiency', 'Resource efficiency',
                 'Renewable energy', 'Electrification', 'Transitional energy',
                 'Coal-related', 'Production-oriented', 'Industrial transformation',
                 'Technology innovation', 'Pollution control',
                 'Environmental quality', 'Recycling', 'Green transportation',
                 'Climate finance']
EN_LEGEND_NEW = ['Carbon reduction', 'Carbon intensity', 'Carbon sink',
                 'Energy saving', 'Energy efficiency', 'Resource efficiency',
                 'Non-fossil energy', 'Electrification', 'Fossil energy',
                 'Production-oriented', 'Industrial transformation',
                 'Pollution control', 'Ecological protection', 'Recycling',
                 'Green transportation', 'Climate finance']

CN_OVERVIEW_OLD = ['碳减排', '碳强度', '碳汇', '污染控制', '节能', '能源效率',
                   '资源效率', '可再生能源', '过渡能源', '电气化', '涉煤',
                   '生产导向', '工业转型', '技术创新', '绿色交通', '环境质量',
                   '循环利用', '气候金融']
CN_OVERVIEW_NEW = ['碳减排', '碳强度', '碳汇', '污染控制', '节能', '能源效率',
                   '资源效率', '非化石能源', '化石能源', '电气化', '生产导向',
                   '工业转型', '绿色交通', '生态保护', '循环利用', '气候金融']
EN_OVERVIEW_OLD = ['Carbon reduction', 'Carbon intensity', 'Carbon sink',
                   'Pollution control', 'Energy saving', 'Energy efficiency',
                   'Resource efficiency', 'Renewable energy', 'Transitional energy',
                   'Electrification', 'Coal-related', 'Production-oriented',
                   'Industrial transformation', 'Technology innovation',
                   'Green transportation', 'Environmental quality', 'Recycling',
                   'Climate finance']
EN_OVERVIEW_NEW = ['Carbon reduction', 'Carbon intensity', 'Carbon sink',
                   'Pollution control', 'Energy saving', 'Energy efficiency',
                   'Resource efficiency', 'Non-fossil energy', 'Fossil energy',
                   'Electrification', 'Production-oriented',
                   'Industrial transformation', 'Green transportation',
                   'Ecological protection', 'Recycling', 'Climate finance']

TXN_CHANGELOG = {
    'CN': '2026-09-09：按类别修订理由书（Target Category Revision Rationale）重组目标类别（18→16类）："环境质量"→"生态保护"；删除"技术创新"（6行并入"能源效率"）；"涉煤"+"过渡能源"合并为"化石能源"（其中12条核电目标并入"非化石能源"）；"可再生能源"→"非化石能源"。中英文同步，共553行。',
    'EN': '2026-09-09: Restructured target categories per the Target Category Revision Rationale (18 → 16 categories): renamed "Environmental quality" → "Ecological protection"; removed "Technology innovation" (6 rows moved to "Energy efficiency"); merged "Coal-related" + "Transitional energy" into "Fossil energy" (12 nuclear rows moved to "Non-fossil energy"); renamed "Renewable energy" → "Non-fossil energy". CN and EN in sync; 553 rows relabeled.',
}


def txn_new_category(old, metric, txn):
    rename, (srcs, dst, nuclear_kw, nuclear_dst) = txn['rename'], txn['merge']
    if old in rename:
        return rename[old]
    if old in srcs:
        if nuclear_kw in metric.lower():
            return nuclear_dst
        return dst
    return None


def txn_find_block(ws, col, first_label):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == first_label:
            return r
    raise AssertionError(f'{ws.title}: block starting with {first_label!r} not found')


def txn_check(path, sheets, txn, legend_old, overview_old, legend_sheet):
    wb = openpyxl.load_workbook(path, read_only=True)
    stats = Counter()
    nuclear = []
    for s in sheets:
        ws = wb[s]
        for r in range(2, ws.max_row + 1):
            old = ws.cell(r, 8).value
            metric = ws.cell(r, 2).value or ''
            new = txn_new_category(old, metric, txn)
            if new:
                stats[(old, new)] += 1
                if old in txn['merge'][0] and txn['merge'][2] in metric.lower():
                    nuclear.append((s, r))
    wb.close()
    for (old, new), k in sorted(stats.items()):
        print(f'  {old} -> {new}: {k}')
    print(f'  nuclear rows rerouted to {txn["merge"][3]}: {len(nuclear)}')
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[legend_sheet]
    start = txn_find_block(ws, 3, legend_old[0])
    for i, lab in enumerate(legend_old):
        got = ws.cell(start + i, 3).value
        assert got == lab, f'legend mismatch {path} r{start+i}: {got!r} != {lab!r}'
    ov = wb['概览' if legend_sheet == '说明' else 'Overview']
    ostart = txn_find_block(ov, 17, overview_old[0])
    for i, lab in enumerate(overview_old):
        got = ov.cell(ostart + i, 17).value
        assert got == lab, f'overview mismatch {path} r{ostart+i}: {got!r} != {lab!r}'
    wb.close()
    print(f'  legend + overview blocks verified ({len(legend_old)} entries)')
    return stats


def txn_rewrite_legend(ws, old_list, new_list):
    start = txn_find_block(ws, 3, old_list[0])
    for i, lab in enumerate(old_list):
        got = ws.cell(start + i, 3).value
        assert got == lab, f'{ws.title} r{start+i}: got {got!r} want {lab!r}'
    for i, lab in enumerate(new_list):
        ws.cell(start + i, 3).value = lab
    for i in range(len(new_list), len(old_list)):
        ws.cell(start + i, 3).value = None


def txn_rewrite_overview(ws, old_list, new_list, txn):
    rename = txn['rename']
    srcs, dst, _, _ = txn['merge']
    start = txn_find_block(ws, 17, old_list[0])
    formulas = {}
    for i, lab in enumerate(old_list):
        got = ws.cell(start + i, 17).value
        assert got == lab, f'{ws.title} r{start+i}: got {got!r} want {lab!r}'
        formulas[lab] = ws.cell(start + i, 18).value
    merged = None
    for src in sorted(srcs, key=old_list.index):
        f = formulas[src].replace(f'"{src}"', f'"{dst}"')
        merged = f if merged is None else merged + '+' + f.lstrip('=')
    renamed_in_list = {new: old for old, new in rename.items()
                       if old in formulas and new not in formulas}
    out = []
    for lab in new_list:
        if lab == dst:
            out.append((lab, merged))
        elif lab in renamed_in_list:
            old = renamed_in_list[lab]
            out.append((lab, formulas[old].replace(f'"{old}"', f'"{lab}"')))
        else:
            out.append((lab, formulas[lab]))
    for i, (lab, f) in enumerate(out):
        ws.cell(start + i, 17).value = lab
        ws.cell(start + i, 18).value = f
    for i in range(len(out), len(old_list)):
        ws.cell(start + i, 17).value = None
        ws.cell(start + i, 18).value = None


def txn_apply(path, sheets, txn, legend_old, legend_new, overview_old,
              overview_new, legend_sheet, changelog_text):
    backup = path.replace('.xlsx', '.pre_taxonomy.bak')
    shutil.copy2(path, backup)
    wb = openpyxl.load_workbook(path)
    n = 0
    for s in sheets:
        ws = wb[s]
        for r in range(2, ws.max_row + 1):
            c = ws.cell(r, 8)
            new = txn_new_category(c.value, ws.cell(r, 2).value or '', txn)
            if new:
                c.value = new
                n += 1
    txn_rewrite_legend(wb[legend_sheet], legend_old, legend_new)
    txn_rewrite_overview(wb['概览' if legend_sheet == '说明' else 'Overview'],
                         overview_old, overview_new, txn)
    rc_add_changelog(wb, legend_sheet, changelog_text,
                     date=datetime.datetime(2026, 9, 9))
    wb.save(path)
    print(f'{path}: {n} category cells relabeled; legend + overview updated; '
          f'changelog added (backup {backup})')


def main_restructure_taxonomy():
    apply = '--apply' in sys.argv
    print(f'{CN_FILE}:')
    txn_check(CN_FILE, CN_DATA_SHEETS, CN_TXN, CN_LEGEND_OLD, CN_OVERVIEW_OLD, '说明')
    print(f'{EN_FILE}:')
    txn_check(EN_FILE, EN_DATA_SHEETS, EN_TXN, EN_LEGEND_OLD, EN_OVERVIEW_OLD, 'README')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    txn_apply(CN_FILE, CN_DATA_SHEETS, CN_TXN, CN_LEGEND_OLD, CN_LEGEND_NEW,
              CN_OVERVIEW_OLD, CN_OVERVIEW_NEW, '说明', TXN_CHANGELOG['CN'])
    txn_apply(EN_FILE, EN_DATA_SHEETS, EN_TXN, EN_LEGEND_OLD, EN_LEGEND_NEW,
              EN_OVERVIEW_OLD, EN_OVERVIEW_NEW, 'README', TXN_CHANGELOG['EN'])
# ---------------------------------------------------------------------------
# clear-changelog: remove the per-revision changelog rows between the
# "Last update" row and "Source information" in 说明/README, leaving
# exactly one blank row in between.
# ---------------------------------------------------------------------------
CL_SPECS = [
    (CN_FILE, '说明', '最后更新', '数据来源说明'),
    (EN_FILE, 'README', 'Last update', 'Source information'),
]


def cl_find_row(ws, text, col=2):
    for r in range(1, ws.max_row + 1):
        if ws.cell(r, col).value == text:
            return r
    return None


def cl_preview(path, legend_sheet, last_update, source_info):
    wb = openpyxl.load_workbook(path, read_only=True)
    ws = wb[legend_sheet]
    a = cl_find_row(ws, last_update)
    b = cl_find_row(ws, source_info)
    wb.close()
    assert a and b, f'{path}: {last_update!r} or {source_info!r} row not found'
    return a, b, b - a - 1


def cl_clear(path, legend_sheet, last_update, source_info):
    wb = openpyxl.load_workbook(path)
    ws = wb[legend_sheet]
    a = cl_find_row(ws, last_update)
    b = cl_find_row(ws, source_info)
    assert a and b, f'{path}: {last_update!r} or {source_info!r} row not found'
    gap, delta = b - a - 1, 1 - (b - a - 1)
    wb.close()
    if delta == 0:
        print(f'{path}: one blank row already present')
        return
    backup = path.replace('.xlsx', '.pre_clearlog.bak')
    shutil.copy2(path, backup)
    wb = openpyxl.load_workbook(path)
    ws = wb[legend_sheet]
    a = cl_find_row(ws, last_update)
    from openpyxl.worksheet.cell_range import MultiCellRange
    if delta < 0:
        ws.delete_rows(a + 1, -delta)
    else:
        ws.insert_rows(a + 1, delta)
    # insert_rows/delete_rows do not move merged ranges; shift them by hand
    rebuilt = MultiCellRange()
    for mr in ws.merged_cells.ranges:
        assert not (mr.min_row <= a < mr.max_row), f'merge {mr} overlaps gap rows'
        if mr.min_row > a:
            mr.shift(row_shift=delta)
        rebuilt.add(mr)
    ws.merged_cells = rebuilt
    wb.save(path)
    print(f'{path}: {gap} rows -> 1 blank row between r{a} and r{b} (backup {backup})')


def main_clear_changelog():
    apply = '--apply' in sys.argv
    for path, sheet, last_update, source_info in CL_SPECS:
        a, b, gap = cl_preview(path, sheet, last_update, source_info)
        print(f'{path} {sheet}: {gap} rows between r{a} and r{b} (target: 1 blank)')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    for path, sheet, last_update, source_info in CL_SPECS:
        cl_clear(path, sheet, last_update, source_info)
# ---------------------------------------------------------------------------
# unify-percent: EN '<digit> percent' -> '<digit>%' (originally unify_percent_en.py)
# ---------------------------------------------------------------------------
PERCENT_PAT = re.compile(r"(\d)\s*percent\b", re.IGNORECASE)

PERCENT_CHANGELOG = (
    '2026-09-07: Unified percentage notation across the English version ("X percent" → "X%").'
)


def unify_percent_text(wb: openpyxl.Workbook) -> int:
    changed = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c, MergedCell):
                    continue
                v = c.value
                if not isinstance(v, str):
                    continue
                new_v = PERCENT_PAT.sub(r"\1%", v)
                if new_v != v:
                    c.value = new_v
                    changed += 1
    return changed


def percent_add_changelog(wb: openpyxl.Workbook) -> None:
    ws = wb["README"]
    for row in ws.iter_rows(min_col=2, max_col=2):
        if row[0].value == PERCENT_CHANGELOG:
            return  # already present (idempotent re-run)
    ws.insert_rows(6)
    src = ws.cell(row=7, column=2)
    dst = ws.cell(row=6, column=2)
    dst.value = PERCENT_CHANGELOG
    dst.font = copy(src.font)
    dst.alignment = copy(src.alignment)
    dst.border = copy(src.border)
    dst.fill = copy(src.fill)
    dst.number_format = src.number_format
    ws.row_dimensions[6].height = 15.0
    ws.cell(row=5, column=3).value = datetime.datetime(2026, 9, 7)


def main_unify_percent():
    path = "China_Climate_Target_Tracker_en.xlsx"
    backup = path.replace(".xlsx", ".pre_percent.bak")
    shutil.copy2(path, backup)
    print(f"backup -> {backup}")

    wb = openpyxl.load_workbook(path)
    changed = unify_percent_text(wb)
    percent_add_changelog(wb)
    wb.save(path)
    print(f"{path}: {changed} cells updated, changelog added")
# ---------------------------------------------------------------------------
# add-ip2604: add the CAC digital-green plan target (originally add_ip2604_target.py)
# ---------------------------------------------------------------------------
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


def ip_add_changelog(wb, sheet_name, text) -> None:
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


def ip_process(path, data_sheet, target_row_values, insert_idx, append_row, readme_sheet,
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
    ip_add_changelog(wb, readme_sheet, changelog_text)
    print(f"  {readme_sheet}: changelog added")

    wb.save(path)
    print(f"{path}: saved")


def main_add_ip2604():
    # EN: insert at sorted position 469 (between 'refined oil annual output' and
    # 'renewable energy share in electrolytic aluminium energy use')
    ip_process(
        "China_Climate_Target_Tracker_en.xlsx",
        "Energy|Power",
        EN_TARGET,
        insert_idx=469,
        append_row=None,
        readme_sheet="README",
        changelog_text=CHANGELOG_EN,
        source_row_values=EN_SOURCE,
    )
    # CN: append at bottom row 564 (after last row 563)
    ip_process(
        "China_Climate_Target_Tracker_cn.xlsx",
        "能源|电力",
        CN_TARGET,
        insert_idx=None,
        append_row=564,
        readme_sheet="说明",
        changelog_text=CHANGELOG_CN,
        source_row_values=CN_SOURCE,
    )
# ---------------------------------------------------------------------------
# wording-fix: wording vs 政策原文 (CN first, then EN mirrors CN) 2026-09-09
# ---------------------------------------------------------------------------
WCOL = {'metric': 2, 'direction': 3, 'mag': 4, 'baseline': 5, 'sentence': 10}

# (文件, 表名, 行号, 列, 方式, 旧值, 新值)
# 方式: set=整格覆盖; replace=子串替换(必须命中)
WORDING_FIXES = [
    # ---- CN: 先确保中文与政策原文一致 ----
    # O1802 蓝天保卫战三年行动计划: 原文不带"市/省"后缀
    ('cn', '能源|电力', 8, 'metric', 'set',
     '上海市、江苏省、浙江省、安徽省及汾渭平原地区煤炭消费总量',
     '上海、江苏、浙江、安徽及汾渭平原煤炭消费总量'),
    ('cn', '能源|电力', 50, 'metric', 'replace',
     '北京市、天津市、河北省、山东省、河南省及珠三角区域',
     '北京、天津、河北、山东、河南及珠三角区域'),
    ('cn', '能源|电力', 50, 'sentence', 'replace',
     '北京市、天津市、河北省、山东省、河南省及珠三角区域',
     '北京、天津、河北、山东、河南及珠三角区域'),
    # FYP1603 原文为"中部8省"/"现役"
    ('cn', '能源|电力', 12, 'metric', 'replace', '中部八省', '中部8省'),
    ('cn', '能源|电力', 12, 'sentence', 'replace', '中部八省', '中部8省'),
    ('cn', '能源|电力', 366, 'metric', 'replace', '在役', '现役'),
    ('cn', '能源|电力', 366, 'sentence', 'replace', '在役', '现役'),
    # O1602 钢铁去产能: 原文为"1亿—1.5亿吨"(长破折号)
    ('cn', '工业', 22, 'sentence', 'replace', '1亿至1.5亿吨', '1亿—1.5亿吨'),
    ('cn', '工业', 22, 'mag', 'replace', '1亿-1.5亿吨', '1亿—1.5亿吨'),
    # HL1902 国家方案: 原文为"工业生产过程的氧化亚氮排放稳定在2005年的水平上"
    ('cn', '污染', 46, 'metric', 'set',
     '氮氧化物（NOₓ）工业排放', '氧化亚氮（N₂O）工业排放'),
    ('cn', '污染', 46, 'direction', 'set', '保持在', '稳定在'),
    ('cn', '污染', 46, 'sentence', 'replace', '二氧化氮', '氧化亚氮'),
    ('cn', '污染', 46, 'sentence', 'replace', '将保持在', '将稳定在'),
    # ---- EN: 英文以中文为准 ----
    ('en', 'Energy|Power', 49, 'mag', 'set', '10% compared with 2015', '10%'),
    ('en', 'Energy|Power', 8, 'mag', 'set', '5% compared with 2015', 'about 5%'),
    ('en', 'Energy|Power', 50, 'mag', 'set', '10%', 'about 10%'),
    ('en', 'Energy|Power', 52, 'metric', 'set',
     'proportion of large refineries in total oil refining',
     'share of refining capacity of the 10 million tonne class'),
    ('en', 'Energy|Power', 52, 'mag', 'set', '55%', 'about 55%'),
    ('en', 'Energy|Power', 244, 'mag', 'set', 'nearly 100%', 'basically completed'),
    ('en', 'Energy|Power', 244, 'sentence', 'set',
     'By the end of 2025, coal-fired boilers of 35 tons of steam per hour or less and all types of coal-fired facilities will be basically eliminated.',
     'By the end of 2025, in the plain areas of the key regions for air pollution control, bulk coal will be basically cleared, and coal-fired boilers of 35 tons of steam per hour or less and all types of coal-fired facilities will be basically eliminated.'),
    ('en', 'Pollution', 46, 'metric', 'set',
     'nitrogen oxides (NOₓ) industrial emissions',
     'nitrous oxide (N₂O) industrial emissions'),
    ('en', 'Pollution', 46, 'mag', 'set', None, '2005 level'),
    ('en', 'Pollution', 46, 'sentence', 'replace', 'NO2', 'nitrous oxide (N2O)'),
    ('en', 'Pollution', 46, 'sentence', 'replace',
     'remain stable as that in 2005', 'remain stable at the 2005 level'),
    ('en', 'Pollution', 46, 'sentence', 'replace', 'from the agriculture',
     'from agriculture'),
    ('en', 'Circular Economy', 139, 'metric', 'set',
     'standardized disposal rate of hazardous waste in Class',
     'standardized disposal rate of hazardous waste in Class I and II maintenance enterprises across the country'),
    ('en', 'Buildings', 9, 'mag', 'set', 'no more than 2000',
     'no more than 2000 units'),
]


def check_wording(fname):
    wb = openpyxl.load_workbook(fname, read_only=True)
    n = 0
    for lang, sheet, row, col, mode, old, new in WORDING_FIXES:
        if lang != ('cn' if 'cn' in fname else 'en'):
            continue
        ws = wb[sheet]
        cur = ws.cell(row=row, column=WCOL[col]).value
        if mode == 'set':
            ok = (cur == old)
        else:
            ok = isinstance(cur, str) and old in cur
        if not ok:
            print(f'ASSERT FAIL {fname} {sheet} r{row} {col}: '
                  f'got {cur!r} want contains/set {old!r}')
            wb.close()
            sys.exit(1)
        n += 1
        print(f'OK {fname} {sheet} r{row} {col} [{new[:24]!r}]')
    wb.close()
    return n


def main_wording_fix():
    apply = '--apply' in sys.argv
    n = check_wording(CN_FILE) + check_wording(EN_FILE)
    print(f'{n} wording fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    for fname in (CN_FILE, EN_FILE):
        shutil.copy(fname, fname.replace('.xlsx', '.pre_wording.bak'))
        wb = openpyxl.load_workbook(fname)
        for lang, sheet, row, col, mode, old, new in WORDING_FIXES:
            if lang != ('cn' if 'cn' in fname else 'en'):
                continue
            ws = wb[sheet]
            cell = ws.cell(row=row, column=WCOL[col])
            if mode == 'set':
                cell.value = new
            else:
                cell.value = cell.value.replace(old, new)
        # EN README legend: 'Direction' merge overran the Target value row.
        # Mirror CN 说明: single 'Direction' cell, 'Target value' merged B36:B42,
        # and the two category/accountability blocks shifted down one row.
        if 'en' in fname:
            ws = wb['README']
            for rng in ('B35:B41', 'B42:B60', 'B61:B66'):
                ws.unmerge_cells(rng)
            ws['B36'] = 'Target value'
            wb_c = openpyxl.load_workbook(CN_FILE)
            copy_style(wb_c['说明']['B36'], ws['B36'])
            wb_c.close()
            for rng in ('B36:B42', 'B43:B61', 'B62:B67'):
                ws.merge_cells(rng)
            print(f'{fname}: README legend fixed')
        wb.save(fname)
        wb.close()
        print(f'{fname}: applied and saved')


# ---------------------------------------------------------------------------
# fix-period: 来源 sheet 覆盖时期 column — letter codes -> period values
# (action-domain codes A/B.x/C.x/D.x had been mistakenly entered there)
# ---------------------------------------------------------------------------
PERIOD_FIXES = {
    14: '十五五', 15: '十五五',
    177: '十五五', 178: '十五五', 179: '十五五', 180: '十五五', 181: '十五五',
    182: '十五五', 183: '十五五', 184: '十五五', 185: '十五五', 186: '十五五',
    187: '十五五',
    270: '-', 318: '-',
    398: '2026-2030', 399: '2026-2028', 400: '2026-2028',
    459: '-', 460: '十五五',
    470: '-', 471: '-', 472: '2026-2030', 473: '-', 474: '-',
    494: '十五五',
}
PERIOD_OLD = {
    14: 'A', 15: 'A',
    177: 'B.1', 178: 'B.5, C.6', 179: 'B.4', 180: 'B.2', 181: 'B.5',
    182: 'B.1', 183: 'D.2', 184: 'A', 185: 'D.3', 186: 'B.1', 187: 'B.1',
    270: 'B.4', 318: 'B.4',
    398: 'C.4', 399: 'B.2, B.1', 400: 'B.1',
    459: 'B.1', 460: 'B.4',
    470: 'B.4', 471: 'B.1', 472: 'B.1', 473: 'B.2', 474: 'A.2',
    494: 'B.3',
}


def main_fix_period():
    apply = '--apply' in sys.argv
    wb = openpyxl.load_workbook(CN_FILE, read_only=True)
    ws = wb['来源']
    for row, old in sorted(PERIOD_OLD.items()):
        cur = ws.cell(row=row, column=10).value
        code = ws.cell(row=row, column=2).value
        if cur != old:
            print(f'ASSERT FAIL {CN_FILE} 来源 r{row} [{code}]: '
                  f'got {cur!r} want {old!r}')
            wb.close()
            sys.exit(1)
        print(f'OK {CN_FILE} 来源 r{row} [{code}] {old!r} -> {PERIOD_FIXES[row]!r}')
    wb.close()
    print(f'{len(PERIOD_FIXES)} period fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    shutil.copy(CN_FILE, CN_FILE.replace('.xlsx', '.pre_period.bak'))
    wb = openpyxl.load_workbook(CN_FILE)
    ws = wb['来源']
    for row, new in PERIOD_FIXES.items():
        ws.cell(row=row, column=10).value = new
    wb.save(CN_FILE)
    wb.close()
    print(f'{CN_FILE}: applied and saved')


# ---------------------------------------------------------------------------
# wording2: (1) AP1802 r268 metric 按原文修正; (2) 约/左右 hedge 以中文为准修正英文;
# (3) O1802 r8 baseline 与原文统一 (2026-09-09)
# ---------------------------------------------------------------------------
W2_SPECIALS = [
    # (1) AP1802 原文为"煤炭占能源消费总量比重"（CN/EN 政策原文句均已为 总量/total）
    ('cn', '能源|电力', 268, 'metric', '煤炭占一次能源消费比重',
     '煤炭占能源消费总量比重'),
    ('en', 'Energy|Power', 268, 'metric', 'proportion of coal in primary energy consumption',
     'proportion of coal in total energy consumption'),
    # (3) O1802 r8 baseline=2015 为推断值，原文无基线
    ('cn', '能源|电力', 8, 'baseline', 2015, None),
    ('en', 'Energy|Power', 8, 'baseline', 2015, None),
    # (2) GO1603 r321: 原文为"提高约1.5%"（非百分点）——中文按原文修正，英文以中文为准
    ('cn', '能源|电力', 321, 'mag', '约1.5个百分点', '约1.5%'),
    ('en', 'Energy|Power', 321, 'mag', '1.5%', 'about 1.5%'),
    # 百分点 vs %：英文以中文为准（percentage points 为既有译法，见 Energy|Power r493 等）
    ('en', 'Energy|Power', 469, 'mag', '20%', '20 percentage points'),
    ('en', 'Industry', 15, 'mag', 'more than 3%', 'more than 3 percentage points'),
    ('en', 'Transport', 49, 'mag', '5%', '5 percentage points'),
    ('en', 'Circular Economy', 128, 'mag', '3%', '3 percentage points'),
    ('en', 'Circular Economy', 157, 'mag', '10%', '10 percentage points'),
    # 能源 r106: 补 about 并顺带修正同格双空格笔误
    ('en', 'Energy|Power', 106, 'mag', '5000  kilometers', 'about 5000 kilometers'),
]

# CN 带 约/左右 而 EN 无 hedge 的行（specials 已处理的 106/321 除外），EN 前缀 'about '
HEDGE_ADD_ROWS = {
    '温室气体排放': [2, 3, 63, 76, 83, 84, 86, 88, 89, 90, 92, 96, 97,
                     108, 110, 111, 112, 113],
    '能源|电力': [13, 75, 76, 97, 104, 130, 134, 160, 175, 177, 195, 197,
                  198, 199, 200, 234, 235, 286, 288, 315, 320, 322, 331,
                  353, 361, 371, 377, 382, 383, 388, 390, 397, 403, 407,
                  409, 412, 415, 427, 434, 436],
    '工业': [60, 61, 62, 63, 65, 91, 92, 168, 169],
    '交通': [71, 73],
    '土地利用、土地利用变化和林业': [10, 67, 68, 70, 88, 90, 176],
    '循环经济': [46, 63, 78, 96, 103, 196],
    '金融': [2, 3, 6, 9, 10, 14, 15, 17, 18, 19],
}
# EN 带 hedge 而 CN 无的行，去掉 EN 前缀 hedge 词
HEDGE_STRIP_ROWS = {'能源|电力': [74]}
HEDGE_EN_TOK = r'\b(?:about|around|approximately|roughly|nearly|almost)\b'


def w2_scan(wcn, wen):
    """Scan the post-special state for CN/EN hedge asymmetry."""
    adds, strips = {}, {}
    for cns, ens in zip(CN_DATA_SHEETS, EN_DATA_SHEETS):
        wsc, wse = wcn[cns], wen[ens]
        add_rows, strip_rows = [], []
        for r in range(2, wsc.max_row + 1):
            mc = wsc.cell(row=r, column=4).value
            me = wse.cell(row=r, column=4).value
            hc = isinstance(mc, str) and ('约' in mc or '左右' in mc)
            he = isinstance(me, str) and re.search(HEDGE_EN_TOK, me) is not None
            if not hc and not he:
                continue
            dc, de = wsc.cell(row=r, column=11).value, wse.cell(row=r, column=11).value
            yc, ye = wsc.cell(row=r, column=1).value, wse.cell(row=r, column=1).value
            if dc != de or yc != ye:
                print(f'ALIGN FAIL {cns} r{r}: CN ({yc!r},{dc!r}) vs EN ({ye!r},{de!r})')
                sys.exit(1)
            if hc and not he:
                add_rows.append(r)
            elif he and not hc:
                strip_rows.append(r)
        if add_rows:
            adds[cns] = add_rows
        if strip_rows:
            strips[cns] = strip_rows
    return adds, strips


def check_wording2():
    n = 0
    for fname, lang in ((CN_FILE, 'cn'), (EN_FILE, 'en')):
        wb = openpyxl.load_workbook(fname, read_only=True)
        for l, sheet, row, col, old, new in W2_SPECIALS:
            if l != lang:
                continue
            cur = wb[sheet].cell(row=row, column=WCOL[col]).value
            if cur != old:
                print(f'ASSERT FAIL {fname} {sheet} r{row} {col}: '
                      f'got {cur!r} want {old!r}')
                wb.close()
                sys.exit(1)
            print(f'OK {fname} {sheet} r{row} {col}: {old!r} -> {new!r}')
            n += 1
        wb.close()
    wcn = openpyxl.load_workbook(CN_FILE)
    wen = openpyxl.load_workbook(EN_FILE)
    for l, sheet, row, col, old, new in W2_SPECIALS:
        (wcn if l == 'cn' else wen)[sheet].cell(row=row, column=WCOL[col]).value = new
    adds, strips = w2_scan(wcn, wen)
    wcn.close()
    wen.close()
    exp_add = {s: set(rs) for s, rs in HEDGE_ADD_ROWS.items()}
    exp_strip = {s: set(rs) for s, rs in HEDGE_STRIP_ROWS.items()}
    if {s: set(rs) for s, rs in adds.items()} != exp_add or \
            {s: set(rs) for s, rs in strips.items()} != exp_strip:
        print(f'SCAN MISMATCH: adds={adds!r} strips={strips!r}')
        sys.exit(1)
    print(f'{sum(map(len, HEDGE_ADD_ROWS.values()))} hedge-add rows + '
          f'{sum(map(len, HEDGE_STRIP_ROWS.values()))} hedge-strip rows verified')
    return n


def main_wording2():
    apply = '--apply' in sys.argv
    n = check_wording2()
    print(f'{n} wording2 fixes verified')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    for fname in (CN_FILE, EN_FILE):
        shutil.copy(fname, fname.replace('.xlsx', '.pre_wording2.bak'))
    wcn = openpyxl.load_workbook(CN_FILE)
    wen = openpyxl.load_workbook(EN_FILE)
    for l, sheet, row, col, old, new in W2_SPECIALS:
        (wcn if l == 'cn' else wen)[sheet].cell(row=row, column=WCOL[col]).value = new
    adds, strips = w2_scan(wcn, wen)
    for cns, rows in adds.items():
        wse = wen[EN_DATA_SHEETS[CN_DATA_SHEETS.index(cns)]]
        for r in rows:
            cell = wse.cell(row=r, column=4)
            cell.value = 'about ' + str(cell.value)
    for cns, rows in strips.items():
        wse = wen[EN_DATA_SHEETS[CN_DATA_SHEETS.index(cns)]]
        for r in rows:
            cell = wse.cell(row=r, column=4)
            cell.value = re.sub(r'^(?:about|around|approximately|roughly|nearly|almost)\s+',
                                '', str(cell.value))
    wcn.save(CN_FILE)
    wen.save(EN_FILE)
    wcn.close()
    wen.close()
    print(f'{CN_FILE}: applied and saved')
    print(f'{EN_FILE}: applied and saved')


# ---------------------------------------------------------------------------
# fix-double-space: collapse double-space typos in EN metric/mag/sentence cells
# (2026-09-09; CN sheets scanned clean)
# ---------------------------------------------------------------------------
DBL_ROWS = {
    ('Buildings', 2): [13, 14, 85, 86],
    ('Buildings', 4): [41, 58],
    ('Circular Economy', 2): [7, 10, 18, 21, 58, 59, 60, 62, 63, 80, 81, 92,
                              93, 99, 100, 101, 102, 108, 115, 118, 131, 136,
                              146, 153, 157, 158, 159, 160, 163, 164, 165, 171,
                              178, 185, 186, 188],
    ('Circular Economy', 4): [5, 6, 45, 57, 61, 62, 67, 68, 69, 111, 112, 194],
    ('Circular Economy', 10): [17, 97],
    ('Energy|Power', 2): [53, 59, 60, 61, 62, 63, 64, 79, 80, 81, 82, 83, 84,
                          85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 98,
                          153, 157, 158, 190, 196, 208, 252, 253, 296, 297, 298,
                          309, 310, 311, 317, 363, 374],
    ('Energy|Power', 4): [4, 16, 124, 126, 127, 211, 233, 242, 243, 306, 308,
                          317, 338, 362],
    ('Energy|Power', 10): [116],
    ('GHG Emissions', 2): [90],
    ('Industry', 2): [2, 30, 36, 38, 39, 42, 43, 45, 63, 105, 106, 166, 167],
    ('Industry', 4): [2, 4, 6, 24, 26, 27, 28, 29, 31, 49, 123, 139, 140,
                      163, 167, 171],
    ('Industry', 10): [55, 98],
    ('LULUCF', 2): [166],
    ('LULUCF', 4): [28, 29, 83, 84, 136, 147, 148, 149, 191, 199],
    ('LULUCF', 10): [27],
    ('Pollution', 4): [4, 5, 11, 12, 13, 14, 16, 19, 65, 68],
    ('Transport', 2): [17, 31, 48, 52],
}
# Industry r98: the double space hides a duplicated word ("items of  of ...")
DBL_SPECIAL = {('Industry', 10, 98): ('of of', 'of')}


def dbl_scan(fname):
    wb = openpyxl.load_workbook(fname, read_only=True)
    found = {}
    for s in EN_DATA_SHEETS:
        ws = wb[s]
        for r, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            for col, v in enumerate(row[:13], start=1):
                if isinstance(v, str) and '  ' in v:
                    found.setdefault((s, col), []).append(r)
    wb.close()
    return found


def main_fix_double_space():
    apply = '--apply' in sys.argv
    found = dbl_scan(EN_FILE)
    exp = {k: set(v) for k, v in DBL_ROWS.items()}
    if {k: set(v) for k, v in found.items()} != exp:
        print(f'SCAN MISMATCH: found={found!r}')
        sys.exit(1)
    print(f'{sum(len(v) for v in DBL_ROWS.values())} double-space cells verified '
          f'({EN_FILE})')
    if not apply:
        print('DRY RUN OK (use --apply to write)')
        return
    shutil.copy(EN_FILE, EN_FILE.replace('.xlsx', '.pre_dblspace.bak'))
    wb = openpyxl.load_workbook(EN_FILE)
    for (sheet, col), rows in DBL_ROWS.items():
        ws = wb[sheet]
        for r in rows:
            cell = ws.cell(row=r, column=col)
            v = ' '.join(str(cell.value).split())
            if (sheet, col, r) in DBL_SPECIAL:
                v = v.replace(*DBL_SPECIAL[(sheet, col, r)])
            cell.value = v
    wb.save(EN_FILE)
    wb.close()
    print(f'{EN_FILE}: applied and saved')


COMMANDS = {
    'compare': main_compare,
    'data': main_data,
    'category-a': main_category_a,
    'category-b': main_category_b,
    'direction': main_direction1,
    'direction2': main_direction2,
    'direction3': main_direction3,
    'direction3b': main_direction3b,
    'rename-categories': main_rename_categories,
    'restructure-taxonomy': main_restructure_taxonomy,
    'clear-changelog': main_clear_changelog,
    'unify-percent': main_unify_percent,
    'add-ip2604': main_add_ip2604,
    'wording-fix': main_wording_fix,
    'fix-period': main_fix_period,
    'wording2': main_wording2,
    'fix-double-space': main_fix_double_space,
}

def main():
    if len(sys.argv) < 2 or sys.argv[1] in ('-h', '--help', 'help'):
        print(__doc__)
        return
    cmd = sys.argv[1]
    if cmd not in COMMANDS:
        print(f'unknown subcommand: {cmd}\n')
        print(__doc__)
        sys.exit(1)
    sys.argv = [sys.argv[0]] + sys.argv[2:]
    COMMANDS[cmd]()

if __name__ == '__main__':
    main()
