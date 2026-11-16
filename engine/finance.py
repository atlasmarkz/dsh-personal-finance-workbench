#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""个人财务管理引擎：读取 CSV/XLSX 个人流水，做数据体检、口径推断、
收支分析（结余 / 储蓄率 / 支出结构 / 异常），并导出 xlsx 与月度报告 docx。

由 Node 侧工作台服务通过 stdin 传入 JSON 载荷，stdout 返回 JSON 结果。
命令：inspect | analyze | report | sample | selftest
"""
import sys
import os
import io
import re
import csv
import json
import math
import traceback
import datetime
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8")
    sys.stderr.reconfigure(encoding="utf-8")
except Exception:
    pass

MAX_ROWS = 200000
DETAIL_SHEET_LIMIT = 20000
TOTAL_WORDS = {"合计", "总计", "小计", "汇总", "总和", "合计：", "总计：", "total", "sum"}
TRANSFER_WORDS = ("还款", "转账", "信用卡", "花呗", "借呗", "贷款", "投资", "理财转出", "提现")


# ---------------------------------------------------------------- 基础工具

def norm_cell(value):
    if value is None:
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, float) and math.isnan(value):
        return ""
    s = str(value).strip()
    s = s.replace("\u00a0", " ").replace("\u3000", " ")
    return s


def parse_number(value):
    """把常见金额写法解析为 float；无法解析返回 None。"""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return None
        return float(value)
    s = norm_cell(value)
    if not s:
        return None
    negative = False
    if s.startswith("(") and s.endswith(")"):
        negative = True
        s = s[1:-1]
    s = s.replace(",", "").replace("，", "").replace(" ", "")
    s = s.replace("￥", "").replace("¥", "").replace("$", "").replace("元", "")
    scale = 1.0
    if s.endswith("万"):
        scale = 10000.0
        s = s[:-1]
    elif s.endswith("亿"):
        scale = 100000000.0
        s = s[:-1]
    if s.startswith("+"):
        s = s[1:]
    if not re.fullmatch(r"-?\d+(?:\.\d+)?", s):
        return None
    try:
        number = float(s) * scale
    except ValueError:
        return None
    if negative:
        number = -number
    return number


def parse_date(value):
    """把常见日期/月份写法解析为 date；无法解析返回 None。"""
    if value is None:
        return None
    if isinstance(value, datetime.datetime):
        return value.date()
    if isinstance(value, datetime.date):
        return value
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 20000 <= float(value) <= 80000:
            try:
                base = datetime.date(1899, 12, 30)
                return base + datetime.timedelta(days=int(value))
            except Exception:
                return None
        if 190001 <= int(value) <= 220012:
            return parse_date(str(int(value)))
        return None
    s = norm_cell(value)
    if not s:
        return None
    s = s.replace("年", "-").replace("月", "-").replace("日", "").replace("/", "-").replace(".", "-")
    s = re.sub(r"\s+", "", s)
    s = s.rstrip("-")
    m = re.fullmatch(r"(\d{4})-(\d{1,2})-(\d{1,2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), 1)
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})(\d{2})(\d{2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
        except ValueError:
            return None
    m = re.fullmatch(r"(\d{4})(\d{2})", s)
    if m:
        try:
            return datetime.date(int(m.group(1)), int(m.group(2)), 1)
        except ValueError:
            return None
    return None


def is_total_row(row):
    for cell in row[:4]:
        if norm_cell(cell).lower() in {w.lower() for w in TOTAL_WORDS}:
            return True
    return False


def read_csv_rows(path):
    raw = Path(path).read_bytes()
    text = None
    used = None
    for enc in ("utf-8-sig", "utf-8", "gb18030", "gbk"):
        try:
            text = raw.decode(enc)
            used = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError("无法识别文件编码，请另存为 UTF-8 或 GBK 的 CSV。")
    sample = text[:8192]
    dialect = csv.excel
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",\t;|")
    except Exception:
        pass
    rows = []
    for row in csv.reader(io.StringIO(text), dialect):
        rows.append([norm_cell(c) for c in row])
    return rows, {"encoding": used, "kind": "csv"}


def read_xlsx_rows(path, sheet):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    names = list(wb.sheetnames)
    target = sheet if sheet and sheet in names else names[0]
    ws = wb[target]
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(["" if v is None else v for v in row])
        if len(rows) > MAX_ROWS:
            break
    wb.close()
    return names, target, rows, {"kind": "xlsx"}


def load_table(path, sheet=None):
    suffix = Path(path).suffix.lower()
    if suffix in (".csv", ".txt", ".tsv"):
        rows, meta = read_csv_rows(path)
        names = [Path(path).stem]
        target = names[0]
    elif suffix in (".xlsx", ".xlsm"):
        names, target, rows, meta = read_xlsx_rows(path, sheet)
    else:
        raise ValueError("只支持 .xlsx / .xlsm / .csv 数据文件。")
    rows = [r for r in rows if any(norm_cell(c) for c in r)]
    if not rows:
        raise ValueError("文件里没有可读取的数据。")
    header_index = detect_header(rows)
    header = [norm_cell(c) for c in rows[header_index]]
    width = max(len(r) for r in rows)
    headers = []
    seen = {}
    for i in range(width):
        name = header[i] if i < len(header) else ""
        name = re.sub(r"\s+", " ", name).strip()
        if not name:
            name = "列%d" % (i + 1)
        base = name
        count = seen.get(base, 0)
        seen[base] = count + 1
        if count:
            name = "%s_%d" % (base, count + 1)
        headers.append(name)
    data = []
    for r in rows[header_index + 1:]:
        values = ["" if v is None else v for v in r]
        if len(values) < width:
            values = values + [""] * (width - len(values))
        elif len(values) > width:
            values = values[:width]
        data.append(values)
    return {
        "path": str(Path(path).resolve()),
        "sheet": target,
        "sheetNames": names,
        "headers": headers,
        "rows": data,
        "meta": meta,
        "headerRow": header_index + 1,
    }


def detect_header(rows):
    best = 0
    for i, row in enumerate(rows[:12]):
        vals = [norm_cell(c) for c in row]
        nonempty = [v for v in vals if v]
        if len(nonempty) < 2:
            continue
        numeric = sum(1 for v in nonempty if parse_number(v) is not None)
        dates = sum(1 for v in nonempty if parse_date(v) is not None)
        unique = len(set(nonempty))
        if numeric + dates <= max(1, int(len(nonempty) * 0.3)) and unique >= len(nonempty) * 0.8:
            return i
    return best


def infer_column(values, header):
    total = 0
    numeric = 0
    dates = 0
    missing = 0
    samples = []
    for v in values:
        s = norm_cell(v)
        if not s:
            missing += 1
            continue
        total += 1
        if len(samples) < 4 and s not in samples:
            samples.append(s)
        if parse_number(s) is not None:
            numeric += 1
        elif parse_date(s) is not None:
            dates += 1
    if total == 0:
        kind = "empty"
    elif numeric / total >= 0.8:
        kind = "number"
    elif dates / total >= 0.8:
        kind = "date"
    else:
        kind = "text"
    if re.search(r"日期|时间|月份|date|period|month", header, re.I) and dates > 0:
        kind = "date"
    unit = ""
    if re.search(r"万元|\(万\)|（万）", header):
        unit = "万元"
    elif re.search(r"\(元\)|（元）|金额", header):
        unit = "元"
    return {
        "type": kind,
        "numericRatio": round(numeric / total, 4) if total else 0,
        "dateRatio": round(dates / total, 4) if total else 0,
        "missing": missing,
        "missingRatio": round(missing / len(values), 4) if values else 0,
        "samples": samples,
        "unit": unit,
    }


CURRENT_HEADERS = []
CURRENT_ROWS = []


def build_profile(table):
    global CURRENT_HEADERS, CURRENT_ROWS
    CURRENT_HEADERS = table["headers"]
    CURRENT_ROWS = table["rows"]
    issues = []
    columns = []
    total = len(table["rows"])
    for i, header in enumerate(table["headers"]):
        values = [r[i] for r in table["rows"]]
        info = infer_column(values, header)
        info["index"] = i
        info["name"] = header
        if info["missingRatio"] >= 0.5 and total:
            issues.append({"level": "warn", "code": "missing", "title": "列「%s」缺失较多" % header,
                           "detail": "%.0f%% 的单元格为空，分析时可能被忽略。" % (info["missingRatio"] * 100)})
        columns.append(info)
    if total < 3:
        issues.append({"level": "high", "code": "few-rows", "title": "数据行很少",
                       "detail": "只有 %d 行数据，趋势分析参考意义有限。" % total})
    return {"columns": columns, "issues": issues}


# ---------------------------------------------------------------- 口径与记录

def suggest_mapping(columns):
    mapping = {"date": None, "amount": None, "direction": None, "category": None, "account": None, "note": None}

    def pick(patterns, kinds):
        candidates = []
        for i, col in enumerate(columns):
            if kinds and col["type"] not in kinds:
                continue
            for order, pat in enumerate(patterns):
                if re.search(pat, col["name"], re.I):
                    candidates.append((order, i))
                    break
        if not candidates:
            return None
        candidates.sort()
        return columns[candidates[0][1]]["name"]

    mapping["date"] = pick([r"^日期$", r"^时间$", r"交易日期", r"日期", r"时间", r"月份", r"^date", r"month"], ["date"])
    mapping["amount"] = pick([r"^金额$", r"收支金额", r"交易金额", r"发生额", r"金额", r"amount", r"money"], ["number"])
    mapping["direction"] = pick([r"收支方向", r"^方向$", r"收支类型", r"交易类型", r"^类型$", r"收/支", r"借贷", r"direction", r"type"], ["text"])
    mapping["category"] = pick([r"^分类$", r"类别", r"类目", r"消费类型", r"支出类别", r"用途", r"项目", r"category", r"tag"], ["text"])
    mapping["account"] = pick([r"账户", r"银行卡", r"支付方式", r"^卡号$", r"account", r"card"], ["text"])
    mapping["note"] = pick([r"备注", r"摘要", r"说明", r"商户", r"对方", r"交易说明", r"note", r"memo", r"desc"], ["text"])
    if mapping["amount"] is None:
        for col in columns:
            if col["type"] == "number":
                mapping["amount"] = col["name"]
                break
    if mapping["direction"] is None:
        for col in columns:
            if col["type"] == "text" and re.search(r"收|支|方向|类型", col["name"]):
                mapping["direction"] = col["name"]
                break
    return mapping


def parse_direction(value):
    """把收支方向文字解析为 income / expense；无法识别返回 None。"""
    s = norm_cell(value).lower()
    if not s:
        return None
    compact = re.sub(r"\s+", "", s)
    if re.search(r"支出|付款|消费|出账|借方|debit|expense|out", compact):
        return "expense"
    if re.search(r"收入|收款|进账|入账|贷方|credit|income", compact):
        return "income"
    if compact in ("-", "−", "负", "支"):
        return "expense"
    if compact in ("+", "正", "收"):
        return "income"
    return None


def pick_records(table, mapping):
    """按口径抽取个人流水，返回 records 与读取统计。"""
    names = table["headers"]
    index = {name: i for i, name in enumerate(names)}

    def col(key):
        name = mapping.get(key)
        if not name:
            return None
        return index.get(name)

    di = col("date")
    ai = col("amount")
    diri = col("direction")
    cati = col("category")
    acci = col("account")
    noti = col("note")
    if ai is None:
        raise ValueError("请先指定「金额列」，否则无法统计收支。")

    def cell(row, i):
        if i is None or i >= len(row):
            return ""
        return row[i]

    stats = {
        "rows": len(table["rows"]),
        "used": 0,
        "skippedTotal": 0,
        "skippedAmount": 0,
        "skippedDirection": 0,
        "zeroAmount": 0,
        "missingDate": 0,
        "missingCategory": 0,
        "unknownDirection": 0,
    }
    amounts = []
    for row in table["rows"]:
        value = parse_number(cell(row, ai))
        if value is not None:
            amounts.append(value)
    all_positive = all(a >= 0 for a in amounts) if amounts else True
    has_direction = diri is not None

    records = []
    for row in table["rows"]:
        if is_total_row(row):
            stats["skippedTotal"] += 1
            continue
        amount = parse_number(cell(row, ai))
        if amount is None:
            stats["skippedAmount"] += 1
            continue
        if amount == 0:
            stats["zeroAmount"] += 1
            continue
        if has_direction:
            direction = parse_direction(cell(row, diri))
            if direction is None:
                stats["unknownDirection"] += 1
                continue
            amount = abs(amount)
        elif amount < 0:
            direction = "expense"
            amount = abs(amount)
        elif all_positive:
            direction = "expense"
        else:
            direction = "income"
        date = parse_date(cell(row, di))
        if date is None:
            stats["missingDate"] += 1
        category = norm_cell(cell(row, cati))
        if not category:
            stats["missingCategory"] += 1
            category = "未分类"
        records.append({
            "date": date.strftime("%Y-%m-%d") if date else "",
            "period": date.strftime("%Y-%m") if date else "未标注日期",
            "direction": direction,
            "amount": round(amount, 2),
            "category": category,
            "account": norm_cell(cell(row, acci)),
            "note": norm_cell(cell(row, noti)),
        })
    stats["used"] = len(records)
    if not records:
        raise ValueError("没有解析出有效的收支记录，请检查「金额列」与「收支方向」的映射。")
    issues = []
    if stats["skippedTotal"]:
        issues.append({"level": "info", "code": "total-row", "title": "已跳过合计行",
                       "detail": "自动跳过 %d 行「合计 / 总计」汇总，不重复计入。" % stats["skippedTotal"]})
    if stats["skippedAmount"]:
        issues.append({"level": "warn", "code": "amount", "title": "有 %d 行金额无法识别" % stats["skippedAmount"],
                       "detail": "这些行没有计入分析，可检查是否为空值或非数字写法。"})
    if stats["unknownDirection"]:
        issues.append({"level": "warn", "code": "direction", "title": "有 %d 行收支方向无法识别" % stats["unknownDirection"],
                       "detail": "方向列按「收入 / 支出」或「收 / 支」填写最稳妥。"})
    if not has_direction and all_positive:
        issues.append({"level": "info", "code": "all-expense", "title": "未指定收支方向列，已按支出处理",
                       "detail": "流水金额全为正数且没有方向列时，默认全部计为支出；若有收入请补充方向列。"})
    if stats["missingDate"]:
        ratio = stats["missingDate"] / max(1, stats["used"])
        issues.append({"level": "warn" if ratio > 0.1 else "info", "code": "date",
                       "title": "有 %d 条记录缺少日期" % stats["missingDate"],
                       "detail": "缺少日期的记录会计入「未标注日期」，不参与月度趋势。"})
    if stats["missingCategory"]:
        ratio = stats["missingCategory"] / max(1, stats["used"])
        issues.append({"level": "warn" if ratio > 0.2 else "info", "code": "category",
                       "title": "有 %d 条记录没有分类" % stats["missingCategory"],
                       "detail": "已归入「未分类」，补充分类后支出结构会更准确。"})
    return records, stats, issues


# ---------------------------------------------------------------- 分析

def sorted_periods(periods):
    real = [p for p in periods if p and p != "未标注日期"]
    result = sorted(real)
    if "未标注日期" in periods:
        result.append("未标注日期")
    return result


def monthly_trend(records):
    by_period = {}
    for record in records:
        period = record["period"]
        bucket = by_period.setdefault(period, {"period": period, "income": 0.0, "expense": 0.0})
        if record["direction"] == "income":
            bucket["income"] += record["amount"]
        else:
            bucket["expense"] += record["amount"]
    trend = []
    for period in sorted_periods(list(by_period.keys())):
        bucket = by_period[period]
        bucket["income"] = round(bucket["income"], 2)
        bucket["expense"] = round(bucket["expense"], 2)
        bucket["net"] = round(bucket["income"] - bucket["expense"], 2)
        bucket["savingsRate"] = round(bucket["net"] / bucket["income"], 4) if bucket["income"] else None
        trend.append(bucket)
    for i, item in enumerate(trend):
        prev = trend[i - 1] if i > 0 else None
        if prev and item["period"] != "未标注日期" and prev["period"] != "未标注日期":
            item["incomeMom"] = round(item["income"] / prev["income"] - 1, 4) if prev["income"] else None
            item["expenseMom"] = round(item["expense"] / prev["expense"] - 1, 4) if prev["expense"] else None
        else:
            item["incomeMom"] = None
            item["expenseMom"] = None
    return trend


def build_structure(records, direction):
    by_category = {}
    for record in records:
        if record["direction"] != direction:
            continue
        bucket = by_category.setdefault(record["category"], {"amount": 0.0, "count": 0})
        bucket["amount"] += record["amount"]
        bucket["count"] += 1
    total = sum(b["amount"] for b in by_category.values())
    items = []
    for name, bucket in sorted(by_category.items(), key=lambda kv: kv[1]["amount"], reverse=True):
        items.append({
            "name": name,
            "amount": round(bucket["amount"], 2),
            "count": bucket["count"],
            "share": round(bucket["amount"] / total, 4) if total else 0,
        })
    head = items[:8]
    rest = items[8:]
    if rest:
        other_amount = sum(item["amount"] for item in rest)
        other_count = sum(item["count"] for item in rest)
        head.append({
            "name": "其他 %d 项" % len(rest),
            "amount": round(other_amount, 2),
            "count": other_count,
            "share": round(other_amount / total, 4) if total else 0,
        })
    for item in head:
        item["avgMonthly"] = round(item["amount"] / 1.0, 2)
    return {"items": head, "total": round(total, 2), "count": len(items)}


def detect_anomalies(records, trend, expense_structure, stats, kpis):
    anomalies = []
    months = [item for item in trend if item["period"] != "未标注日期"]
    month_count = max(1, len(months))
    avg_expense = kpis["expense"] / month_count

    expenses = [r for r in records if r["direction"] == "expense" and r["date"]]
    if expenses and avg_expense > 0:
        threshold = max(800.0, avg_expense * 0.35)
        big = sorted(expenses, key=lambda r: r["amount"], reverse=True)
        big = [r for r in big if r["amount"] >= threshold][:3]
        for record in big:
            anomalies.append({
                "level": "info",
                "code": "large-expense",
                "title": "大额单笔支出：%s 元" % format_money(record["amount"]),
                "detail": "%s · %s%s" % (record["date"], record["category"],
                                         (" · " + record["note"]) if record["note"] else ""),
            })

    by_category_month = {}
    for record in records:
        if record["direction"] != "expense" or not record["date"]:
            continue
        bucket = by_category_month.setdefault(record["category"], {})
        bucket[record["period"]] = bucket.get(record["period"], 0.0) + record["amount"]
    deviations = []
    for category, series in by_category_month.items():
        periods = sorted(series.keys())
        if len(periods) < 2:
            continue
        last = periods[-1]
        prev = periods[-2]
        previous = series.get(prev, 0.0)
        current = series.get(last, 0.0)
        if previous >= 200 and current >= 300:
            change = current / previous - 1
            if abs(change) >= 0.6:
                deviations.append((abs(change), category, last, previous, current, change))
    deviations.sort(reverse=True)
    for _, category, period, previous, current, change in deviations[:3]:
        anomalies.append({
            "level": "warn",
            "code": "category-mom",
            "title": "「%s」%s 支出环比 %s" % (category, period, format_pct(change, signed=True)),
            "detail": "上月 %s 元，本月 %s 元。若是偶发支出可忽略。" % (format_money(previous), format_money(current)),
        })

    negative_months = [item for item in months if item["net"] < 0]
    if negative_months:
        longest = []
        run = []
        for item in months:
            if item["net"] < 0:
                run.append(item["period"])
                if len(run) > len(longest):
                    longest = list(run)
            else:
                run = []
        level = "warn" if len(longest) >= 2 else "info"
        anomalies.append({
            "level": level,
            "code": "deficit",
            "title": "有 %d 个月支出大于收入" % len(negative_months),
            "detail": "发生在 %s%s。" % ("、".join(item["period"] for item in negative_months),
                                       ("；其中最长连续 %d 个月为负" % len(longest)) if len(longest) >= 2 else ""),
        })

    if kpis["savingsRate"] is not None and kpis["savingsRate"] < 0.1:
        anomalies.append({
            "level": "warn",
            "code": "low-savings",
            "title": "储蓄率偏低：%s" % format_pct(kpis["savingsRate"]),
            "detail": "结余占收入不足 10%，建议先留出应急储备再增加可选消费。"})
    elif kpis["savingsRate"] is not None and kpis["savingsRate"] >= 0.4:
        anomalies.append({
            "level": "info",
            "code": "high-savings",
            "title": "储蓄率较好：%s" % format_pct(kpis["savingsRate"]),
            "detail": "结余占比达到 40% 以上，可考虑把部分结余转入长期储备。"})

    if stats["used"] and stats["missingCategory"] / stats["used"] > 0.1:
        anomalies.append({
            "level": "info",
            "code": "category-missing",
            "title": "未分类支出占比较高",
            "detail": "有 %d 条流水没有分类，支出结构会被「未分类」拉低可读性。" % stats["missingCategory"]})

    transfer_total = 0.0
    transfer_count = 0
    for record in records:
        if record["direction"] != "expense":
            continue
        if any(word in record["category"] for word in TRANSFER_WORDS) or any(word in record["note"] for word in TRANSFER_WORDS):
            transfer_total += record["amount"]
            transfer_count += 1
    if transfer_count:
        anomalies.append({
            "level": "info",
            "code": "transfer",
            "title": "有 %d 笔疑似转账 / 还款类支出" % transfer_count,
            "detail": "合计 %s 元，这类流水通常不算消费性支出，看支出结构时可以剔除。" % format_money(transfer_total)})

    duplicates = {}
    for record in records:
        key = (record["date"], record["amount"], record["direction"], record["category"])
        duplicates[key] = duplicates.get(key, 0) + 1
    dup_count = sum(1 for count in duplicates.values() if count > 1)
    if dup_count:
        anomalies.append({
            "level": "info",
            "code": "duplicate",
            "title": "发现 %d 组疑似重复流水" % dup_count,
            "detail": "日期、金额、分类完全相同的记录，若不是重复请忽略。"})

    if expense_structure["items"] and expense_structure["items"][0]["share"] >= 0.4:
        top = expense_structure["items"][0]
        anomalies.append({
            "level": "info",
            "code": "concentration",
            "title": "支出集中在「%s」" % top["name"],
            "detail": "占总支出 %s，可优先从这里做优化。" % format_pct(top["share"])})
    if not anomalies:
        anomalies.append({"level": "info", "code": "ok", "title": "没有发现明显异常",
                          "detail": "收支结构比较平稳，可以按计划做储蓄与预算安排。"})
    return anomalies


def analyze_core(payload):
    table = load_table(payload["path"], payload.get("sheet"))
    profile = build_profile(table)
    columns = profile["columns"]
    suggested = suggest_mapping(columns)
    mapping = payload.get("mapping") or suggested
    for key in ("date", "amount", "direction", "category", "account", "note"):
        mapping.setdefault(key, suggested.get(key))
    records, stats, record_issues = pick_records(table, mapping)
    trend = monthly_trend(records)
    months = [item for item in trend if item["period"] != "未标注日期"]
    month_count = len(months)
    total_income = round(sum(r["amount"] for r in records if r["direction"] == "income"), 2)
    total_expense = round(sum(r["amount"] for r in records if r["direction"] == "expense"), 2)
    net = round(total_income - total_expense, 2)
    savings_rate = round(net / total_income, 4) if total_income else None
    if month_count >= 4:
        rates = [item["savingsRate"] for item in months if item["savingsRate"] is not None]
        if len(rates) >= 4:
            trend_direction = round(sum(rates[-3:]) / 3.0 - sum(rates[:3]) / 3.0, 4)
        else:
            trend_direction = None
    else:
        trend_direction = None
    largest = max((r for r in records if r["direction"] == "expense"), key=lambda r: r["amount"], default=None)
    dates = [r["date"] for r in records if r["date"]]
    kpis = {
        "income": total_income,
        "expense": total_expense,
        "net": net,
        "savingsRate": savings_rate,
        "months": month_count,
        "records": stats["used"],
        "avgMonthlyIncome": round(total_income / month_count, 2) if month_count else 0,
        "avgMonthlyExpense": round(total_expense / month_count, 2) if month_count else 0,
        "avgMonthlyNet": round(net / month_count, 2) if month_count else 0,
        "periodStart": min(dates) if dates else "",
        "periodEnd": max(dates) if dates else "",
        "trendDirection": trend_direction,
        "largestExpense": {"amount": largest["amount"], "category": largest["category"],
                           "date": largest["date"]} if largest else None,
    }
    expense_structure = build_structure(records, "expense")
    income_structure = build_structure(records, "income")
    unit = ""
    for column in columns:
        if column["name"] == mapping.get("amount") and column.get("unit"):
            unit = column["unit"]
    anomalies = detect_anomalies(records, trend, expense_structure, stats, kpis)
    anomalies.sort(key=lambda item: 0 if item.get("level") in ("warn", "error") else 1)
    issues = list(profile["issues"]) + record_issues
    quality = {
        "stats": stats,
        "issues": issues,
        "columns": columns,
        "dateRange": {"start": kpis["periodStart"], "end": kpis["periodEnd"]},
        "unit": unit,
    }
    return {
        "ok": True,
        "fileName": Path(payload["path"]).name,
        "path": table["path"],
        "sheet": table["sheet"],
        "sheetNames": table["sheetNames"],
        "mapping": mapping,
        "kpis": kpis,
        "trend": trend,
        "expenseStructure": expense_structure,
        "incomeStructure": income_structure,
        "anomalies": anomalies,
        "quality": quality,
    }, records


def analyze_payload(payload):
    analysis, _records = analyze_core(payload)
    return analysis


def format_money(value):
    if value is None:
        return "—"
    return "{:,.2f}".format(float(value))


def format_pct(value, digits=1, signed=False):
    if value is None:
        return "—"
    sign = "+" if signed and float(value) > 0 else ""
    return (sign + "{:." + str(digits) + "f}%").format(float(value) * 100)


# ---------------------------------------------------------------- Excel

HEADER_COLOR = "1F3864"


def style_header_row(ws, row, count):
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    fill = PatternFill("solid", fgColor=HEADER_COLOR)
    font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
    thin = Side(style="thin", color="D0D5DD")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for column in range(1, count + 1):
        cell = ws.cell(row=row, column=column)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = border


def set_widths(ws, widths):
    from openpyxl.utils import get_column_letter
    for index, width in enumerate(widths):
        ws.column_dimensions[get_column_letter(index + 1)].width = width


def body_cell(cell, bold=False):
    from openpyxl.styles import Font, Alignment, Border, Side
    thin = Side(style="thin", color="E5E7EB")
    cell.font = Font(name="微软雅黑", size=10, bold=bold)
    cell.alignment = Alignment(vertical="center")
    cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)


def money_cell(cell):
    body_cell(cell)
    cell.number_format = "#,##0.00"


def pct_cell(cell):
    body_cell(cell)
    cell.number_format = "0.0%"


def sheet_title(ws, text, span):
    from openpyxl.styles import Font, Alignment
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=span)
    cell = ws.cell(row=1, column=1, value=text)
    cell.font = Font(name="微软雅黑", size=14, bold=True, color=HEADER_COLOR)
    cell.alignment = Alignment(vertical="center")
    ws.row_dimensions[1].height = 26


def meta_row(ws, row, label, value):
    from openpyxl.styles import Font
    ws.cell(row=row, column=1, value=label).font = Font(name="微软雅黑", size=10, color="6B7280")
    cell = ws.cell(row=row, column=2, value=value)
    cell.font = Font(name="微软雅黑", size=10)
    body_cell(cell)


def make_xlsx(analysis, records, out_path, compare=None):
    from openpyxl import Workbook
    from openpyxl.styles import Font
    wb = Workbook()
    kpis = analysis["kpis"]
    trend = analysis["trend"]
    months = [item for item in trend if item["period"] != "未标注日期"]
    month_count = max(1, len(months))
    generated_at = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

    ws = wb.active
    ws.title = "概览"
    sheet_title(ws, "个人财务概览", 4)
    meta_rows = [
        ("数据源", analysis["fileName"]),
        ("工作表", analysis["sheet"]),
        ("分析期间", "%s ~ %s" % (kpis["periodStart"] or "—", kpis["periodEnd"] or "—")),
        ("记录数", "%d 条" % kpis["records"]),
        ("生成时间", generated_at),
    ]
    for offset, (label, value) in enumerate(meta_rows):
        meta_row(ws, 3 + offset, label, value)
    set_widths(ws, [18, 22, 46, 14])
    header_row = 9
    for column, text in enumerate(["指标", "数值", "说明", "口径"], start=1):
        ws.cell(row=header_row, column=column, value=text)
    style_header_row(ws, header_row, 4)
    trend_sheet = "'月度收支'"
    total_row = 5 + len(months)
    detail_sheet = "'收支明细'"
    details = records[:DETAIL_SHEET_LIMIT]
    first_detail = 5
    last_detail = first_detail + max(1, len(details)) - 1
    rows = [
        ("总收入", "=%s!B%d" % (trend_sheet, total_row), "工资、副业、理财等收入合计", "收入列求和"),
        ("总支出", "=%s!C%d" % (trend_sheet, total_row), "全部消费与固定支出合计", "支出列求和"),
        ("净结余", "=B%d-B%d" % (header_row + 1, header_row + 2), "收入 − 支出", "结余"),
        ("储蓄率", "=IFERROR(B%d/B%d,0)" % (header_row + 3, header_row + 1), "净结余 / 总收入", "比率"),
        ("月均收入", "=B%d/%d" % (header_row + 1, month_count), "总收入 / 月数", "月均"),
        ("月均支出", "=B%d/%d" % (header_row + 2, month_count), "总支出 / 月数", "月均"),
        ("月均结余", "=B%d/%d" % (header_row + 3, month_count), "净结余 / 月数", "月均"),
        ("最大单笔支出", "=IFERROR(MAX(%s!D%d:D%d),0)" % (detail_sheet, first_detail, last_detail),
         "明细中金额最大的支出（见收支明细）", "单笔"),
    ]
    for offset, (label, formula, note, rule) in enumerate(rows):
        row = header_row + 1 + offset
        ws.cell(row=row, column=1, value=label)
        value_cell = ws.cell(row=row, column=2, value=formula)
        ws.cell(row=row, column=3, value=note)
        ws.cell(row=row, column=4, value=rule)
        body_cell(ws.cell(row=row, column=1), bold=True)
        if label in ("储蓄率",):
            pct_cell(value_cell)
        else:
            money_cell(value_cell)
        body_cell(ws.cell(row=row, column=3))
        body_cell(ws.cell(row=row, column=4))

    ws2 = wb.create_sheet("月度收支")
    sheet_title(ws2, "月度收支与储蓄率", 7)
    headers = ["期间", "收入", "支出", "结余", "储蓄率", "收入环比", "支出环比"]
    for column, text in enumerate(headers, start=1):
        ws2.cell(row=4, column=column, value=text)
    style_header_row(ws2, 4, len(headers))
    for offset, item in enumerate(trend):
        row = 5 + offset
        ws2.cell(row=row, column=1, value=item["period"])
        income_cell = ws2.cell(row=row, column=2, value=item["income"])
        expense_cell = ws2.cell(row=row, column=3, value=item["expense"])
        net_cell = ws2.cell(row=row, column=4, value="=B%d-C%d" % (row, row))
        rate_cell = ws2.cell(row=row, column=5, value="=IFERROR(D%d/B%d,0)" % (row, row))
        prev = 5 + offset - 1
        if offset > 0 and item["period"] != "未标注日期" and trend[offset - 1]["period"] != "未标注日期":
            ws2.cell(row=row, column=6, value="=IFERROR(B%d/B%d-1,0)" % (row, prev))
            ws2.cell(row=row, column=7, value="=IFERROR(C%d/C%d-1,0)" % (row, prev))
        body_cell(ws2.cell(row=row, column=1))
        money_cell(income_cell)
        money_cell(expense_cell)
        money_cell(net_cell)
        pct_cell(rate_cell)
        pct_cell(ws2.cell(row=row, column=6))
        pct_cell(ws2.cell(row=row, column=7))
    total_row2 = 5 + len(trend)
    ws2.cell(row=total_row2, column=1, value="合计").font = Font(
        name="微软雅黑", size=10, bold=True)
    ws2.cell(row=total_row2, column=2, value="=SUM(B5:B%d)" % (total_row2 - 1))
    ws2.cell(row=total_row2, column=3, value="=SUM(C5:C%d)" % (total_row2 - 1))
    ws2.cell(row=total_row2, column=4, value="=B%d-C%d" % (total_row2, total_row2))
    ws2.cell(row=total_row2, column=5, value="=IFERROR(D%d/B%d,0)" % (total_row2, total_row2))
    for column in range(1, 8):
        cell = ws2.cell(row=total_row2, column=column)
        body_cell(cell, bold=True)
        if column in (2, 3, 4):
            money_cell(cell)
            cell.font = Font(name="微软雅黑", size=10, bold=True)
        if column == 5:
            pct_cell(cell)
    set_widths(ws2, [12, 14, 14, 14, 10, 12, 12])

    ws3 = wb.create_sheet("支出结构")
    sheet_title(ws3, "支出结构（按分类）", 5)
    for column, text in enumerate(["分类", "支出金额", "占比", "月均支出", "笔数"], start=1):
        ws3.cell(row=4, column=column, value=text)
    style_header_row(ws3, 4, 5)
    items = analysis["expenseStructure"]["items"]
    for offset, item in enumerate(items):
        row = 5 + offset
        ws3.cell(row=row, column=1, value=item["name"])
        amount_cell = ws3.cell(row=row, column=2, value=item["amount"])
        share_cell = ws3.cell(row=row, column=3, value="=IFERROR(B%d/$B$%d,0)" % (row, 5 + len(items)))
        ws3.cell(row=row, column=4, value=round(item["amount"] / month_count, 2))
        ws3.cell(row=row, column=5, value=item["count"])
        body_cell(ws3.cell(row=row, column=1))
        money_cell(amount_cell)
        pct_cell(share_cell)
        money_cell(ws3.cell(row=row, column=4))
        body_cell(ws3.cell(row=row, column=5))
    total_row3 = 5 + len(items)
    ws3.cell(row=total_row3, column=1, value="合计")
    ws3.cell(row=total_row3, column=2, value="=SUM(B5:B%d)" % (total_row3 - 1))
    ws3.cell(row=total_row3, column=3, value="=IFERROR(B%d/$B$%d,0)" % (total_row3, total_row3))
    ws3.cell(row=total_row3, column=5, value="=SUM(E5:E%d)" % (total_row3 - 1))
    for column in range(1, 6):
        cell = ws3.cell(row=total_row3, column=column)
        body_cell(cell, bold=True)
        if column == 2:
            money_cell(cell)
            cell.font = Font(name="微软雅黑", size=10, bold=True)
        if column == 3:
            pct_cell(cell)
    set_widths(ws3, [24, 16, 10, 14, 8])

    ws4 = wb.create_sheet("收入来源")
    sheet_title(ws4, "收入来源", 4)
    for column, text in enumerate(["来源", "收入金额", "占比", "笔数"], start=1):
        ws4.cell(row=4, column=column, value=text)
    style_header_row(ws4, 4, 4)
    income_items = analysis["incomeStructure"]["items"]
    for offset, item in enumerate(income_items):
        row = 5 + offset
        ws4.cell(row=row, column=1, value=item["name"])
        amount_cell = ws4.cell(row=row, column=2, value=item["amount"])
        share_cell = ws4.cell(row=row, column=3, value="=IFERROR(B%d/$B$%d,0)" % (row, 5 + len(income_items)))
        ws4.cell(row=row, column=4, value=item["count"])
        body_cell(ws4.cell(row=row, column=1))
        money_cell(amount_cell)
        pct_cell(share_cell)
        body_cell(ws4.cell(row=row, column=4))
    total_row4 = 5 + len(income_items)
    ws4.cell(row=total_row4, column=1, value="合计")
    ws4.cell(row=total_row4, column=2, value="=SUM(B5:B%d)" % (total_row4 - 1))
    ws4.cell(row=total_row4, column=3, value="=IFERROR(B%d/$B$%d,0)" % (total_row4, total_row4))
    ws4.cell(row=total_row4, column=4, value="=SUM(D5:D%d)" % (total_row4 - 1))
    for column in range(1, 5):
        cell = ws4.cell(row=total_row4, column=column)
        body_cell(cell, bold=True)
        if column == 2:
            money_cell(cell)
            cell.font = Font(name="微软雅黑", size=10, bold=True)
        if column == 3:
            pct_cell(cell)
    set_widths(ws4, [20, 16, 10, 8])

    ws5 = wb.create_sheet("收支明细")
    sheet_title(ws5, "收支明细", 7)
    for column, text in enumerate(["日期", "期间", "收支方向", "金额", "分类", "账户", "备注"], start=1):
        ws5.cell(row=4, column=column, value=text)
    style_header_row(ws5, 4, 7)
    ordered = details
    for offset, record in enumerate(ordered):
        row = 5 + offset
        ws5.cell(row=row, column=1, value=record["date"] or "未标注")
        ws5.cell(row=row, column=2, value=record["period"])
        ws5.cell(row=row, column=3, value="收入" if record["direction"] == "income" else "支出")
        amount_cell = ws5.cell(row=row, column=4, value=record["amount"])
        ws5.cell(row=row, column=5, value=record["category"])
        ws5.cell(row=row, column=6, value=record["account"])
        ws5.cell(row=row, column=7, value=record["note"])
        for column in range(1, 8):
            body_cell(ws5.cell(row=row, column=column))
        money_cell(amount_cell)
    set_widths(ws5, [12, 10, 10, 14, 16, 14, 30])
    ws5.freeze_panes = "A5"

    ws6 = wb.create_sheet("数据体检")
    sheet_title(ws6, "数据体检", 3)
    ws6.cell(row=4, column=1, value="指标")
    ws6.cell(row=4, column=2, value="数值")
    style_header_row(ws6, 4, 2)
    stats = analysis["quality"]["stats"]
    stat_labels = [
        ("读取行数", stats["rows"]),
        ("有效记录", stats["used"]),
        ("跳过合计行", stats["skippedTotal"]),
        ("金额无法识别", stats["skippedAmount"]),
        ("方向无法识别", stats["unknownDirection"]),
        ("缺少日期", stats["missingDate"]),
        ("没有分类", stats["missingCategory"]),
        ("金额单位", analysis["quality"].get("unit") or "按源表"),
    ]
    for offset, (label, value) in enumerate(stat_labels):
        row = 5 + offset
        ws6.cell(row=row, column=1, value=label)
        ws6.cell(row=row, column=2, value=value)
        body_cell(ws6.cell(row=row, column=1))
        body_cell(ws6.cell(row=row, column=2))
    issue_row = 5 + len(stat_labels) + 1
    ws6.cell(row=issue_row, column=1, value="级别")
    ws6.cell(row=issue_row, column=2, value="说明")
    style_header_row(ws6, issue_row, 2)
    for offset, issue in enumerate(analysis["quality"]["issues"]):
        row = issue_row + 1 + offset
        ws6.cell(row=row, column=1, value={"error": "需处理", "warn": "注意", "high": "注意"}.get(issue["level"], "提示"))
        ws6.cell(row=row, column=2, value="%s：%s" % (issue.get("title", ""), issue.get("detail", "")))
        body_cell(ws6.cell(row=row, column=1))
        body_cell(ws6.cell(row=row, column=2))
    set_widths(ws6, [20, 76])

    ws7 = wb.create_sheet("口径卡")
    sheet_title(ws7, "口径卡", 3)
    for column, text in enumerate(["字段", "对应列", "说明"], start=1):
        ws7.cell(row=4, column=column, value=text)
    style_header_row(ws7, 4, 3)
    mapping = analysis["mapping"]
    mapping_rows = [
        ("日期", mapping.get("date") or "未指定", "用于月度趋势与期间统计"),
        ("金额", mapping.get("amount") or "未指定", "正数金额，方向由「收支方向」列决定"),
        ("收支方向", mapping.get("direction") or "未指定（按金额正负）", "收入 / 支出"),
        ("分类", mapping.get("category") or "未指定", "支出结构、收入来源的分组依据"),
        ("账户", mapping.get("account") or "未指定", "可选，用于区分账户"),
        ("备注", mapping.get("note") or "未指定", "可选，原始摘要"),
    ]
    for offset, (label, value, note) in enumerate(mapping_rows):
        row = 5 + offset
        ws7.cell(row=row, column=1, value=label)
        ws7.cell(row=row, column=2, value=value)
        ws7.cell(row=row, column=3, value=note)
        for column in range(1, 4):
            body_cell(ws7.cell(row=row, column=column))
    formula_row = 5 + len(mapping_rows) + 1
    ws7.cell(row=formula_row, column=1, value="计算口径").font = Font(
        name="微软雅黑", size=11, bold=True, color=HEADER_COLOR)
    formulas = [
        "净结余 = 总收入 − 总支出",
        "储蓄率 = 净结余 ÷ 总收入",
        "支出结构占比 = 分类支出 ÷ 总支出",
        "收入环比 = 本月收入 ÷ 上月收入 − 1（支出环比同理）",
        "含「合计 / 总计 / 小计」的行会被自动跳过",
    ]
    for offset, text in enumerate(formulas):
        cell = ws7.cell(row=formula_row + 1 + offset, column=1, value=text)
        ws7.merge_cells(start_row=formula_row + 1 + offset, start_column=1, end_row=formula_row + 1 + offset, end_column=3)
        body_cell(cell)
    set_widths(ws7, [16, 26, 52])

    if compare and compare.get("previous"):
        ws8 = wb.create_sheet("复跑对比")
        sheet_title(ws8, "复跑对比", 4)
        for column, text in enumerate(["指标", "上次", "本次", "变化"], start=1):
            ws8.cell(row=4, column=column, value=text)
        style_header_row(ws8, 4, 4)
        previous = compare["previous"]
        pairs = [
            ("总收入", previous.get("income"), kpis["income"]),
            ("总支出", previous.get("expense"), kpis["expense"]),
            ("净结余", previous.get("net"), kpis["net"]),
            ("储蓄率", previous.get("savingsRate"), kpis["savingsRate"]),
        ]
        for offset, (label, old, new) in enumerate(pairs):
            row = 5 + offset
            ws8.cell(row=row, column=1, value=label)
            old_cell = ws8.cell(row=row, column=2, value=old)
            new_cell = ws8.cell(row=row, column=3, value=new)
            change_cell = ws8.cell(row=row, column=4, value="=IFERROR(C%d/B%d-1,0)" % (row, row))
            body_cell(ws8.cell(row=row, column=1))
            if label == "储蓄率":
                pct_cell(old_cell)
                pct_cell(new_cell)
            else:
                money_cell(old_cell)
                money_cell(new_cell)
            pct_cell(change_cell)
        set_widths(ws8, [16, 16, 16, 12])

    wb.calculation.fullCalcOnLoad = True
    wb.save(out_path)
    return out_path


# ---------------------------------------------------------------- Word

NUMERALS = ["一", "二", "三", "四", "五", "六", "七", "八", "九"]


def docx_defaults(document):
    from docx.shared import Pt, Cm
    from docx.oxml.ns import qn
    style = document.styles["Normal"]
    style.font.name = "微软雅黑"
    style.font.size = Pt(10.5)
    rpr = style.element.get_or_add_rPr()
    rpr.get_or_add_rFonts().set(qn("w:eastAsia"), "微软雅黑")
    style.paragraph_format.line_spacing = 1.5
    style.paragraph_format.space_after = Pt(4)
    for section in document.sections:
        section.top_margin = section.bottom_margin = Cm(2.2)
        section.left_margin = section.right_margin = Cm(2.2)


def add_section_heading(document, text):
    from docx.shared import Pt, RGBColor
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_before = Pt(14)
    paragraph.paragraph_format.space_after = Pt(6)
    run = paragraph.add_run(text)
    run.bold = True
    run.font.size = Pt(13)
    run.font.name = "微软雅黑"
    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    return paragraph


def add_para(document, text, size=10.5, bold=False, color=None, space_after=4):
    from docx.shared import Pt, RGBColor
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.space_after = Pt(space_after)
    run = paragraph.add_run(text)
    run.font.size = Pt(size)
    run.bold = bold
    if color:
        run.font.color.rgb = RGBColor(*color)
    return paragraph


def add_table(document, headers, rows, widths=None, highlight_last=False, pct_columns=(), money_columns=()):
    from docx.shared import Pt, RGBColor, Cm
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    table = document.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.autofit = False
    header_cells = table.rows[0].cells
    for i, text in enumerate(headers):
        cell = header_cells[i]
        cell.text = ""
        paragraph = cell.paragraphs[0]
        run = paragraph.add_run(text)
        run.bold = True
        run.font.size = Pt(9.5)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shading = OxmlElement("w:shd")
        shading.set(qn("w:fill"), HEADER_COLOR)
        cell._tc.get_or_add_tcPr().append(shading)
    for row_index, row in enumerate(rows):
        cells = table.add_row().cells
        for column_index, value in enumerate(row):
            cell = cells[column_index]
            cell.text = ""
            paragraph = cell.paragraphs[0]
            if isinstance(value, str):
                text = value
            elif isinstance(value, bool):
                text = "是" if value else "否"
            elif column_index in pct_columns and isinstance(value, (int, float)):
                text = format_pct(value)
            elif isinstance(value, int):
                text = str(value)
            else:
                text = format_money(value)
            run = paragraph.add_run(text)
            run.font.size = Pt(9.5)
            if highlight_last and row_index == len(rows) - 1:
                run.bold = True
    if widths:
        for row in table.rows:
            for i, width in enumerate(widths):
                if i < len(row.cells):
                    row.cells[i].width = Cm(width)
    header_tr_pr = table.rows[0]._tr.get_or_add_trPr()
    header_flag = OxmlElement("w:tblHeader")
    header_flag.set(qn("w:val"), "true")
    header_tr_pr.append(header_flag)
    for row in table.rows:
        row_tr_pr = row._tr.get_or_add_trPr()
        row_tr_pr.append(OxmlElement("w:cantSplit"))
    return table


def make_docx(analysis, out_path, compare=None):
    from docx import Document
    from docx.shared import Pt, RGBColor, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    kpis = analysis["kpis"]
    trend = [item for item in analysis["trend"] if item["period"] != "未标注日期"]
    document = Document()
    docx_defaults(document)

    title = document.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run("个人财务月度分析报告")
    run.bold = True
    run.font.size = Pt(20)
    run.font.name = "微软雅黑"
    run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)
    subtitle = document.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sub_run = subtitle.add_run("数据源：%s · 分析期间：%s 至 %s · 生成于 %s" % (
        analysis["fileName"], kpis["periodStart"] or "—", kpis["periodEnd"] or "—",
        datetime.datetime.now().strftime("%Y-%m-%d %H:%M")))
    sub_run.font.size = Pt(9)
    sub_run.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)

    add_section_heading(document, "一、结论摘要")
    savings = format_pct(kpis["savingsRate"]) if kpis["savingsRate"] is not None else "—"
    summary = "本期共 %d 条收支记录，覆盖 %d 个月：总收入 %s 元，总支出 %s 元，净结余 %s 元，储蓄率 %s。" % (
        kpis["records"], kpis["months"], format_money(kpis["income"]), format_money(kpis["expense"]),
        format_money(kpis["net"]), savings)
    add_para(document, summary)
    if kpis["savingsRate"] is not None:
        if kpis["savingsRate"] >= 0.4:
            judgement = "储蓄率处于较好水平，说明收入覆盖支出后仍有充足余量，可把结余的一部分转为长期储备。"
        elif kpis["savingsRate"] >= 0.1:
            judgement = "储蓄率处于中间水平，建议先固定一笔「先存后花」的自动转出，再安排可选消费。"
        else:
            judgement = "储蓄率偏低，收入几乎被支出吃满，建议优先压缩占比最高的支出分类，并留出应急储备。"
    else:
        judgement = "本期没有记录到收入，无法计算储蓄率；如果有工资等收入，请确认收支方向列。"
    add_para(document, judgement)
    if trend:
        recent = trend[-3:]
        recent_text = "、".join("%s 为 %s 元" % (item["period"], format_money(item["net"])) for item in recent)
        add_para(document, "最近三个月净结余：%s。大额与波动支出见第五章。" % recent_text)

    add_section_heading(document, "二、关键指标")
    kpi_rows = [
        ("总收入", kpis["income"], "工资、副业、理财等全部收入", "元"),
        ("总支出", kpis["expense"], "全部消费与固定支出", "元"),
        ("净结余", kpis["net"], "收入减支出", "元"),
        ("储蓄率", kpis["savingsRate"], "净结余占总收入比例", ""),
        ("月均收入", kpis["avgMonthlyIncome"], "总收入除以月数", "元"),
        ("月均支出", kpis["avgMonthlyExpense"], "总支出除以月数", "元"),
        ("最大单笔支出", kpis["largestExpense"]["amount"] if kpis["largestExpense"] else None,
         ("%s · %s" % (kpis["largestExpense"]["date"], kpis["largestExpense"]["category"])) if kpis["largestExpense"] else "无记录", "元"),
    ]
    table_rows = []
    for label, value, note, unit in kpi_rows:
        if label == "储蓄率":
            text = format_pct(value)
        elif value is None:
            text = "—"
        else:
            text = "%s %s" % (format_money(value), unit)
        table_rows.append((label, text, note))
    add_table(document, ["指标", "数值", "说明"], table_rows, widths=[3.4, 4.2, 9.0])

    add_section_heading(document, "三、月度收支")
    month_rows = []
    for item in trend[-12:]:
        month_rows.append((
            item["period"],
            item["income"],
            item["expense"],
            item["net"],
            item["savingsRate"],
        ))
    add_table(document, ["期间", "收入", "支出", "结余", "储蓄率"], month_rows,
              widths=[3.0, 3.6, 3.6, 3.6, 3.0], pct_columns=(4,))
    if len(trend) > 12:
        add_para(document, "（仅展示最近 12 个月，完整数据见 xlsx 的「月度收支」工作表。）", size=9,
                 color=(0x6B, 0x72, 0x80))

    add_section_heading(document, "四、支出结构")
    structure_rows = []
    for item in analysis["expenseStructure"]["items"]:
        structure_rows.append((item["name"], item["amount"], item["share"], item["count"]))
    structure_rows.append(("合计", analysis["expenseStructure"]["total"], 1.0, ""))
    add_table(document, ["分类", "支出金额", "占比", "笔数"], structure_rows,
              widths=[6.0, 4.4, 3.0, 2.8], pct_columns=(2,), highlight_last=True)
    top = analysis["expenseStructure"]["items"][0] if analysis["expenseStructure"]["items"] else None
    if top:
        add_para(document, "支出占比最高的是「%s」，占 %s，平均每月约 %s 元。" % (
            top["name"], format_pct(top["share"]), format_money(top["amount"] / max(1, kpis["months"]))))

    add_section_heading(document, "五、异常与建议")
    ordered_anomalies = sorted(analysis["anomalies"],
                               key=lambda item: 0 if item["level"] in ("warn", "error") else 1)
    for item in ordered_anomalies:
        level = {"warn": "注意", "error": "异常"}.get(item["level"], "提示")
        add_para(document, "【%s】%s：%s" % (level, item["title"], item["detail"]), size=10, space_after=3)

    if compare and compare.get("previous"):
        previous = compare["previous"]
        add_section_heading(document, "六、复跑对比")
        compare_rows = []
        for label, key in (("总收入", "income"), ("总支出", "expense"), ("净结余", "net"), ("储蓄率", "savingsRate")):
            old = previous.get(key)
            new = kpis.get(key)
            if key == "savingsRate":
                change = (new - old) if (old is not None and new is not None) else None
                compare_rows.append((label, format_pct(old), format_pct(new),
                                     ("%+.1f 个百分点" % (change * 100)) if change is not None else "—"))
            else:
                change = (new / old - 1) if (old not in (None, 0) and new is not None) else None
                compare_rows.append((label, format_money(old), format_money(new), format_pct(change) if change is not None else "—"))
        add_table(document, ["指标", "上次", "本次", "变化"], compare_rows, widths=[3.6, 3.8, 3.8, 5.0])
        add_para(document, "对比基准为同一数据文件上一次生成的分析结果。")

    add_section_heading(document, "六、口径与说明" if not (compare and compare.get("previous")) else "七、口径与说明")
    mapping = analysis["mapping"]
    mapping_parts = [
        "日期列映射到「%s」" % (mapping.get("date") or "未指定"),
        "金额列映射到「%s」" % (mapping.get("amount") or "未指定"),
    ]
    if mapping.get("direction"):
        mapping_parts.append("收支方向列映射到「%s」" % mapping["direction"])
    else:
        mapping_parts.append("没有收支方向列，按金额正负判断（负数视为支出，全为正数时默认视为支出）")
    if mapping.get("category"):
        mapping_parts.append("分类列映射到「%s」" % mapping["category"])
    add_para(document, "口径映射：" + "；".join(mapping_parts) + "。", size=10)
    add_para(document, "计算口径：净结余 = 总收入 − 总支出；储蓄率 = 净结余 ÷ 总收入；支出结构 = 分类支出 ÷ 总支出；"
                       "含「合计 / 总计 / 小计」的行自动跳过。", size=10)
    add_para(document, "数据只在本机读取和计算；本报告由「个人财务管理」工作台生成，可在 Word 中直接编辑。", size=9,
             color=(0x6B, 0x72, 0x80))

    document.save(out_path)
    return out_path


def make_report(payload):
    analysis, records = analyze_core(payload)
    out_dir = Path(payload.get("outDir") or ".")
    out_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M")
    base = Path(payload["path"]).stem
    xlsx_path = out_dir / ("%s_个人财务分析_%s.xlsx" % (base, stamp))
    docx_path = out_dir / ("%s_个人财务报告_%s.docx" % (base, stamp))
    compare = payload.get("compare")
    make_xlsx(analysis, records, str(xlsx_path), compare=compare)
    make_docx(analysis, str(docx_path), compare=compare)
    return {
        "ok": True,
        "generatedAt": datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "files": [
            {"kind": "xlsx", "path": str(xlsx_path.resolve()), "bytes": xlsx_path.stat().st_size},
            {"kind": "docx", "path": str(docx_path.resolve()), "bytes": docx_path.stat().st_size},
        ],
        "analysis": analysis,
    }


# ---------------------------------------------------------------- 示例数据

def make_sample(payload):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    explicit = payload.get("out") or payload.get("path")
    if explicit:
        out = Path(explicit)
    else:
        out_dir = Path(payload.get("outDir") or ".")
        out_dir.mkdir(parents=True, exist_ok=True)
        out = out_dir / "示例个人流水.xlsx"
    out.parent.mkdir(parents=True, exist_ok=True)
    months = ["2026-%02d" % m for m in range(1, 13)]
    expenses = {
        "2026-01": [("房租", 4500, "房租"), ("餐饮", 1280, "日常三餐"), ("餐饮", 760, "周末聚餐"),
                    ("交通", 320, "地铁公交"), ("交通", 210, "打车"), ("购物", 980, "年货"),
                    ("保险", 4800, "年度重疾险"), ("娱乐", 520, "电影演出"), ("通讯", 100, "手机套餐"),
                    ("水电燃气", 280, "一月账单")],
        "2026-02": [("房租", 4500, "房租"), ("餐饮", 1560, "春节聚餐"), ("餐饮", 890, "日常三餐"),
                    ("交通", 460, "返乡"), ("购物", 1200, "换季衣物"), ("人情", 8000, "春节红包"),
                    ("娱乐", 420, "聚会"), ("通讯", 100, "手机套餐"), ("水电燃气", 260, "二月账单")],
        "2026-03": [("房租", 4500, "房租"), ("餐饮", 1320, "日常三餐"), ("餐饮", 700, "周末聚餐"),
                    ("交通", 300, "地铁公交"), ("购物", 680, "日用"), ("教育", 800, "线上课程"),
                    ("娱乐", 580, "演出"), ("通讯", 100, "手机套餐"), ("水电燃气", 240, "三月账单")],
        "2026-04": [("房租", 4500, "房租"), ("餐饮", 1180, "日常三餐"), ("餐饮", 620, "外卖"),
                    ("交通", 280, "地铁公交"), ("购物", 540, "日用"), ("娱乐", 360, "电影"),
                    ("通讯", 100, "手机套餐"), ("水电燃气", 230, "四月账单")],
        "2026-05": [("房租", 4500, "房租"), ("餐饮", 1240, "日常三餐"), ("餐饮", 680, "周末聚餐"),
                    ("交通", 310, "地铁公交"), ("购物", 760, "夏装"), ("娱乐", 600, "演出"),
                    ("通讯", 100, "手机套餐"), ("水电燃气", 250, "五月账单")],
        "2026-06": [("房租", 4500, "房租"), ("餐饮", 1380, "日常三餐"), ("餐饮", 820, "聚餐"),
                    ("交通", 360, "打车"), ("购物", 900, "618 采购"), ("娱乐", 450, "电影"),
                    ("医疗", 6800, "体检与牙科"), ("通讯", 100, "手机套餐"), ("水电燃气", 300, "六月账单")],
        "2026-07": [("房租", 4500, "房租"), ("餐饮", 1420, "日常三餐"), ("餐饮", 760, "周末聚餐"),
                    ("交通", 420, "暑期出行"), ("购物", 820, "家居"), ("娱乐", 680, "旅行门票"),
                    ("通讯", 100, "手机套餐"), ("水电燃气", 340, "七月账单")],
        "2026-08": [("房租", 4500, "房租"), ("餐饮", 1360, "日常三餐"), ("餐饮", 700, "聚餐"),
                    ("交通", 360, "地铁公交"), ("购物", 740, "日用"), ("娱乐", 520, "演出"),
                    ("医疗", 380, "感冒就诊"), ("通讯", 100, "手机套餐"), ("水电燃气", 330, "八月账单")],
        "2026-09": [("房租", 4500, "房租"), ("餐饮", 1480, "日常三餐"), ("餐饮", 820, "聚餐"),
                    ("交通", 380, "地铁公交"), ("购物", 880, "换季衣物"), ("教育", 12000, "职业进修课程"),
                    ("娱乐", 460, "电影"), ("通讯", 100, "手机套餐"), ("水电燃气", 280, "九月账单")],
        "2026-10": [("房租", 4500, "房租"), ("餐饮", 1520, "国庆聚餐"), ("餐饮", 900, "日常三餐"),
                    ("交通", 520, "假期出行"), ("购物", 980, "家电"), ("人情", 1200, "婚礼礼金"),
                    ("娱乐", 700, "旅行"), ("通讯", 100, "手机套餐"), ("水电燃气", 240, "十月账单")],
        "2026-11": [("房租", 4500, "房租"), ("餐饮", 1440, "日常三餐"), ("餐饮", 760, "聚餐"),
                    ("交通", 340, "地铁公交"), ("购物", 6500, "双十一大件"), ("娱乐", 520, "演出"),
                    ("通讯", 100, "手机套餐"), ("水电燃气", 250, "十一月账单")],
        "2026-12": [("房租", 4500, "房租"), ("餐饮", 1580, "年末聚餐"), ("餐饮", 860, "日常三餐"),
                    ("交通", 380, "地铁公交"), ("购物", 1100, "节日礼物"), ("娱乐", 620, "跨年活动"),
                    ("通讯", 100, "手机套餐"), ("水电燃气", 270, "十二月账单")],
    }
    incomes = {
        "2026-01": [("工资", 18000, "1 月工资"), ("奖金", 12000, "年终奖")],
        "2026-02": [("工资", 18000, "2 月工资")],
        "2026-03": [("工资", 18000, "3 月工资"), ("理财收益", 420, "基金分红")],
        "2026-04": [("工资", 6000, "离职当月工资")],
        "2026-05": [("副业", 8000, "自由职业项目")],
        "2026-06": [("工资", 18000, "6 月工资"), ("副业", 2500, "兼职设计")],
        "2026-07": [("工资", 18000, "7 月工资")],
        "2026-08": [("工资", 18000, "8 月工资")],
        "2026-09": [("工资", 18000, "9 月工资"), ("副业", 2500, "兼职设计")],
        "2026-10": [("工资", 18000, "10 月工资"), ("理财收益", 650, "基金分红")],
        "2026-11": [("工资", 18000, "11 月工资")],
        "2026-12": [("工资", 18000, "12 月工资"), ("副业", 2500, "年终项目"), ("奖金", 8000, "年度奖金")],
    }
    rows = []
    for month in months:
        for category, amount, note in incomes.get(month, []):
            rows.append((month + "-10", "收入", amount, category, "招商银行", note))
        day = 5
        for category, amount, note in expenses.get(month, []):
            if category in ("房租",):
                date = month + "-05"
            else:
                day = (day % 25) + 1
                date = "%s-%02d" % (month, day)
            account = "微信" if category in ("餐饮", "交通", "购物", "娱乐") else "招商银行"
            rows.append((date, "支出", amount, category, account, note))
    rows.sort(key=lambda item: item[0])
    wb = Workbook()
    ws = wb.active
    ws.title = "个人流水"
    headers = ["日期", "收支方向", "金额", "分类", "账户", "备注"]
    ws.append(headers)
    for row in rows:
        ws.append(list(row))
    header_fill = PatternFill("solid", fgColor=HEADER_COLOR)
    for column in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=column)
        cell.fill = header_fill
        cell.font = Font(name="微软雅黑", size=10, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center")
    ws.column_dimensions["A"].width = 12
    ws.column_dimensions["B"].width = 10
    ws.column_dimensions["C"].width = 12
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 14
    ws.column_dimensions["F"].width = 24
    wb.save(str(out))
    return {"ok": True, "path": str(out.resolve()), "rows": len(rows)}


# ---------------------------------------------------------------- 入口

def payload_from_stdin():
    raw = sys.stdin.buffer.read() if hasattr(sys.stdin, "buffer") else sys.stdin.read().encode("utf-8")
    raw = raw.lstrip(b"\xef\xbb\xbf").strip()
    if not raw:
        return {}
    return json.loads(raw.decode("utf-8"))


def dispatch(command, payload):
    if command == "inspect":
        table = load_table(payload["path"], payload.get("sheet"))
        profile = build_profile(table)
        suggested = suggest_mapping(profile["columns"])
        return {
            "ok": True,
            "fileName": Path(payload["path"]).name,
            "path": table["path"],
            "sheet": table["sheet"],
            "sheetNames": table["sheetNames"],
            "rowCount": len(table["rows"]),
            "headerRow": table["headerRow"],
            "columns": profile["columns"],
            "suggestedMapping": suggested,
            "issues": profile["issues"],
            "preview": table["rows"][:5],
        }
    if command == "analyze":
        return analyze_payload(payload)
    if command == "report":
        return make_report(payload)
    if command == "sample":
        return make_sample(payload)
    if command == "selftest":
        import tempfile
        root = Path(tempfile.mkdtemp(prefix="pf-selftest-"))
        sample = make_sample({"out": str(root / "sample.xlsx")})
        inspected = dispatch("inspect", {"path": sample["path"]})
        analysis = dispatch("analyze", {"path": sample["path"]})
        report = dispatch("report", {"path": sample["path"], "outDir": str(root / "out")})
        return {"ok": True, "root": str(root), "rows": sample["rows"],
                "mapping": analysis["mapping"], "kpis": analysis["kpis"],
                "files": [file["path"] for file in report["files"]]}
    raise ValueError("未知命令：%s" % command)


def main():
    payload = payload_from_stdin()
    command = payload.get("command") or (sys.argv[1] if len(sys.argv) > 1 else "")
    try:
        result = dispatch(command, payload)
    except Exception as error:  # noqa: BLE001
        result = {"ok": False, "error": str(error), "trace": traceback.format_exc()}
    sys.stdout.write(json.dumps(result, ensure_ascii=False, default=str))
    sys.stdout.flush()


if __name__ == "__main__":
    main()
