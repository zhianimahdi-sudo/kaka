#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
intake.py — دریافت و شناسایی خودکار فایل ورودی (فایل قبلی کاربر).

کار می‌کند با:
    • فایل اکسل (xlsx/xlsm)  → استخراج کامل فرمول‌ها/نام‌ها/جداول/اعتبارسنجی‌ها
    • فایل متنی (txt/csv/tsv) → شناسایی جداکننده، تحلیل ستون‌ها، استخراج فرمول‌ها
    • پوشه             → همهٔ فایل‌های داخل آن را یکی‌یکی تحلیل می‌کند

خروجی‌ها در پوشهٔ reports/ ذخیره می‌شوند:
    <name>.report.txt   گزارش خوانا
    <name>.report.json  گزارش ساخت‌یافته (برای تطبیق با کارپوشهٔ جدید)
    <name>.table.csv    نسخهٔ یکدست‌شدهٔ جدولِ فایل متنی (اگر متنی بود)

نمونهٔ اجرا:
    PYTHONPATH=./pylibs python3 tools/intake.py                 # پیش‌فرض: /home/user/uploads
    PYTHONPATH=./pylibs python3 tools/intake.py 03.txt          # یک فایل مشخص
    PYTHONPATH=./pylibs python3 tools/intake.py /home/user/uploads
"""

from __future__ import annotations

import csv
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

DEFAULT_DIRS = ["/home/user/uploads", os.path.join(os.path.dirname(HERE), "incoming")]
OUT_DIR = os.path.join(os.path.dirname(HERE), "reports")

TEXT_EXTS = {".txt", ".csv", ".tsv", ".tab", ".dat", ".md", ".json", ".xml", ".html"}
EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}


# ────────────────────────────────────────────── کمکی‌ها

def sniff_kind(path: str) -> str:
    """نوع فایل را از پسوند و امضای بایتی تشخیص می‌دهد."""
    ext = os.path.splitext(path)[1].lower()
    if ext in EXCEL_EXTS:
        return "excel"
    if ext == ".xls":
        return "legacy-excel"
    try:
        with open(path, "rb") as fh:
            head = fh.read(8)
    except OSError:
        return "unknown"
    if head.startswith(b"PK\x03\x04"):
        return "excel-or-zip"
    if head.startswith(b"\xd0\xcf\x11\xe0"):
        return "legacy-excel"
    return "text"


def detect_delimiter(text: str) -> str:
    """پرتکرارترین جداکننده را در چند خط اول پیدا می‌کند."""
    sample = "\n".join(text.splitlines()[:40])
    counts = {d: sample.count(d) for d in ["\t", ";", ",", "|"]}
    best = max(counts, key=counts.get)
    return best if counts[best] > 0 else "\t"


def looks_numeric(value: str) -> bool:
    v = str(value).strip().replace(",", "").replace("٬", "").replace("،", "")
    v = v.replace("٫", ".").replace("−", "-")
    if not v:
        return False
    try:
        float(v)
        return True
    except ValueError:
        return False


FORMULA_HINT = re.compile(r"^\s*=|[A-Z]{1,3}\d{1,5}\s*[:=]|(SUMIFS?|COUNTIFS?|VLOOKUP|INDEX|MATCH|IFERROR|IF)\s*\(", re.I)


# ────────────────────────────────────────────── تحلیل فایل متنی

def analyze_text(path: str) -> tuple[dict, list[list[str]]]:
    raw = open(path, "rb").read()
    text = None
    for enc in ("utf-8-sig", "utf-8", "windows-1256", "cp1252"):
        try:
            text = raw.decode(enc)
            encoding = enc
            break
        except UnicodeDecodeError:
            continue
    if text is None:  # pragma: no cover
        text = raw.decode("utf-8", errors="replace")
        encoding = "utf-8 (با جانشینی خطاها)"

    delim = detect_delimiter(text)
    rows = [r for r in csv.reader(io.StringIO(text), delimiter=delim) if any(str(c).strip() for c in r)]
    width = max((len(r) for r in rows), default=0)
    rows = [r + [""] * (width - len(r)) for r in rows]

    # یافتن ردیف سرستون: اولین ردیفی که بیشتر سلول‌هایش غیرعددی و یکتا باشد
    header_idx = 0
    for i, r in enumerate(rows[:10]):
        cells = [c.strip() for c in r if c.strip()]
        if len(cells) >= 2 and sum(1 for c in cells if not looks_numeric(c)) >= max(2, len(cells) // 2):
            header_idx = i
            break
    header = [c.strip() or f"ستون {i + 1}" for i, c in enumerate(rows[header_idx])] if rows else []
    data = rows[header_idx + 1:]

    col_stats = []
    for c in range(width):
        vals = [r[c].strip() for r in data if r[c].strip()]
        num = [v for v in vals if looks_numeric(v)]
        formulas = [v for v in vals if FORMULA_HINT.search(v)]
        uniq = len(set(vals))
        examples = []
        for v in vals:
            if v not in examples:
                examples.append(v)
            if len(examples) >= 3:
                break
        col_stats.append({
            "column": c + 1,
            "header": header[c] if c < len(header) else f"ستون {c + 1}",
            "filled": len(vals),
            "unique": uniq,
            "numeric_share": round(len(num) / len(vals), 3) if vals else 0.0,
            "formula_like": len(formulas),
            "examples": examples,
        })

    formula_lines = [r for r in rows if any(FORMULA_HINT.search(c) for c in r)]
    report = {
        "file": os.path.abspath(path),
        "size_bytes": len(raw),
        "kind": "text",
        "encoding": encoding,
        "delimiter": {"\t": "Tab", ";": "نقطه‌ویرگول", ",": "کاما", "|": "لوله"}.get(delim, repr(delim)),
        "rows": len(rows),
        "columns": width,
        "header_row_index": header_idx + 1,
        "header": header,
        "column_stats": col_stats,
        "formula_like_count": len(formula_lines),
        "formula_like_sample": formula_lines[:12],
    }
    return report, rows


# ────────────────────────────────────────────── تحلیل فایل اکسل

def analyze_excel(path: str) -> dict:
    try:
        from inspect_workbook import inspect
    except ImportError:  # pragma: no cover
        return {"file": os.path.abspath(path), "kind": "excel",
                "error": "openpyxl نصب نیست؛ با PYTHONPATH=./pylibs اجرا کنید"}
    rep = inspect(path, formulas_only=False)
    keys = ["sheet_count", "sheets_order", "defined_names", "totals"]
    slim = {k: rep[k] for k in keys if k in rep}
    slim["file"] = rep["file"]
    slim["size_bytes"] = rep["size_bytes"]
    slim["kind"] = "excel"
    slim["sheets"] = {}
    for name, s in rep["sheets"].items():
        slim["sheets"][name] = {
            "dimensions": s["dimensions"], "state": s["state"], "rightToLeft": s["rightToLeft"],
            "formula_count": s["formula_count"], "formulas": s["formulas"],
            "tables": s["tables"], "data_validations": s["data_validations"],
            "conditional_formatting": s["conditional_formatting"], "charts": s["charts"],
            "merged_cells": s["merged_cells"], "number_formats_used": s["number_formats_used"],
            "input_counts": None,
        }
    return slim


# ────────────────────────────────────────────── گزارش خوانا

def render_text_report(rep: dict) -> str:
    L: list[str] = []
    p = L.append
    p("═" * 78)
    p("  گزارش دریافت فایل ورودی")
    p("═" * 78)
    p(f"فایل : {rep.get('file')}")
    p(f"حجم  : {rep.get('size_bytes', 0):,} بایت")
    p(f"نوع  : {rep.get('kind')}")

    if rep.get("kind") == "excel":
        p(f"شیت‌ها ({rep.get('sheet_count')}): {' | '.join(rep.get('sheets_order', []))}")
        t = rep.get("totals", {})
        p(f"جمع فرمول‌ها: {t.get('formulas', 0):,} · اعتبارسنجی: {t.get('validations', 0)} · "
          f"قواعد شرطی: {t.get('cf_rules', 0)} · جداول: {t.get('tables', 0)} · نمودار: {t.get('charts', 0)}")
        p("")
        p("─" * 78)
        p("■ نام‌های محدوده")
        for dn in rep.get("defined_names", []):
            p(f"   • {dn['name']:<28} ← {dn['refersTo']}")
        for name, s in rep.get("sheets", {}).items():
            p("")
            p("─" * 78)
            p(f"■ {name}  ({s['dimensions']}, RTL={'بله' if s['rightToLeft'] else 'خیر'}, "
              f"فرمول={s['formula_count']})")
            p("─" * 78)
            for tb in s["tables"]:
                p(f"   جدول: {tb['name']} [{tb['ref']}] → {' | '.join(tb['columns'])}")
            for dv in s["data_validations"]:
                p(f"   اعتبارسنجی: {dv['range']} type={dv['type']} op={dv['operator']} f1={dv['formula1']}")
            for cf in s["conditional_formatting"]:
                for r in cf["rules"]:
                    p(f"   شرطی: {cf['range']} {r['type']}"
                      f"{'/' + str(r.get('operator')) if r.get('operator') else ''} f={r.get('formula')}")
            for ch in s["charts"]:
                p(f"   نمودار: {ch.get('type')} «{ch.get('title') or '—'}»")
            if s["formulas"]:
                p(f"   ── فرمول‌ها ({len(s['formulas'])}) ──")
                for coord, f in list(s["formulas"].items())[:200]:
                    p(f"     {coord:<9} {f}")
                if len(s["formulas"]) > 200:
                    p(f"     … و {len(s['formulas']) - 200} فرمول دیگر (کامل در فایل JSON)")
        return "\n".join(L)

    # ── متن
    p(f"رمزگذاری: {rep.get('encoding')} · جداکننده: {rep.get('delimiter')}")
    p(f"ابعاد : {rep.get('rows'):,} ردیف × {rep.get('columns')} ستون · ردیف سرستون: {rep.get('header_row_index')}")
    p("")
    p("─" * 78)
    p("■ تحلیل ستون‌ها")
    p("─" * 78)
    p(f"{'#':<4}{'عنوان':<26}{'پرشده':>8}{'یکتا':>8}{'عددی':>8}{'شبیه فرمول':>12}   نمونه‌ها")
    for c in rep.get("column_stats", []):
        ex = " ، ".join(str(x)[:16] for x in c["examples"])
        p(f"{c['column']:<4}{str(c['header'])[:24]:<26}{c['filled']:>8}{c['unique']:>8}"
          f"{c['numeric_share'] * 100:>7.0f}%{c['formula_like']:>12}   {ex}")
    p("")
    p(f"■ خطوط شبیه فرمول: {rep.get('formula_like_count', 0)}")
    for line in rep.get("formula_like_sample", [])[:12]:
        p("   " + " | ".join(str(x)[:40] for x in line))
    return "\n".join(L)


# ────────────────────────────────────────────── اجرا

def intake_one(path: str) -> dict:
    kind = sniff_kind(path)
    if kind in ("excel", "excel-or-zip"):
        rep = analyze_excel(path)
    elif kind in ("text",):
        rep, rows = analyze_text(path)
    else:
        rep = {"file": os.path.abspath(path), "kind": kind, "size_bytes": os.path.getsize(path),
               "error": "این نوع فایل پشتیبانی نمی‌شود (.xls قدیمی را به xlsx ذخیره کنید)"}

    os.makedirs(OUT_DIR, exist_ok=True)
    stem = os.path.splitext(os.path.basename(path))[0] or "file"
    txt_path = os.path.join(OUT_DIR, f"{stem}.report.txt")
    json_path = os.path.join(OUT_DIR, f"{stem}.report.json")
    with open(txt_path, "w", encoding="utf-8") as fh:
        fh.write(render_text_report(rep))
    with open(json_path, "w", encoding="utf-8") as fh:
        json.dump(rep, fh, ensure_ascii=False, indent=1, default=str)

    if kind == "text":
        rep2, rows = analyze_text(path)
        csv_path = os.path.join(OUT_DIR, f"{stem}.table.csv")
        with open(csv_path, "w", encoding="utf-8-sig", newline="") as fh:
            csv.writer(fh).writerows(rows)
        rep["csv_copy"] = csv_path

    rep["report_txt"] = txt_path
    rep["report_json"] = json_path
    return rep


def main() -> int:
    args = sys.argv[1:]
    targets: list[str] = []
    if args:
        for a in args:
            if os.path.isdir(a):
                targets += [os.path.join(a, f) for f in sorted(os.listdir(a))
                            if os.path.isfile(os.path.join(a, f)) and not f.startswith(".")]
            else:
                targets.append(a)
    else:
        for d in DEFAULT_DIRS:
            if os.path.isdir(d):
                targets += [os.path.join(d, f) for f in sorted(os.listdir(d))
                            if os.path.isfile(os.path.join(d, f)) and not f.startswith(".")]

    if not targets:
        print("⚠️  هیچ فایلی برای دریافت پیدا نشد.")
        print("   مسیرهای بررسی‌شده: " + " · ".join(DEFAULT_DIRS))
        print("   راه‌های جایگزین رساندن فایل به من:")
        print("     ۱) محتوای فایل را در گفتگو بچسبانید (متن/TSV).")
        print("     ۲) فایل را در مخزن گیت‌هاب (Add file → Upload files) بگذارید؛ خودم برمی‌دارم.")
        print("     ۳) مسیر فایل را به این ابزار بدهید: python3 tools/intake.py <path>")
        return 2

    ok = 0
    for t in targets:
        print(f"\n▶ دریافت: {t}")
        try:
            rep = intake_one(t)
        except Exception as exc:
            print(f"   ❌ خطا در تحلیل: {type(exc).__name__}: {exc}")
            continue
        ok += 1
        print(f"   نوع: {rep.get('kind')}")
        if rep.get("kind") == "excel":
            print(f"   شیت‌ها: {rep.get('sheet_count')} · فرمول‌ها: {rep.get('totals', {}).get('formulas', 0):,}")
        elif rep.get("kind") == "text":
            print(f"   جدول: {rep.get('rows'):,} ردیف × {rep.get('columns')} ستون · "
                  f"خطوط شبیه فرمول: {rep.get('formula_like_count', 0)}")
        print(f"   گزارش: {rep.get('report_txt')}")
        print(f"   JSON : {rep.get('report_json')}")
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
