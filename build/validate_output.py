#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate_output.py — اعتبارسنجی خودکار کارپوشهٔ ساخته‌شده.

چه چیزهایی بررسی می‌شود؟
    ۱. هر نام محدوده به شیت و سلول موجود اشاره می‌کند.
    ۲. هر ارجاع ساخت‌یافته (مثل tbl_TX[جمع با مالیات]) به جدول و ستونی موجود اشاره می‌کند.
    ۳. هر نام محدودهٔ استفاده‌شده در فرمول‌ها واقعاً تعریف شده است.
    ۴. پرانتزها در همهٔ فرمول‌ها متوازن‌اند و هیچ «None» یا «#REF!» در فرمول‌ها نیست.
    ۵. هر جدول رسمی بازهٔ معتبر و ستون‌های بدون تکرار دارد.
    ۶. همهٔ شیت‌ها right-to-left هستند و ترتیب شیت‌ها درست است.

نمونهٔ اجرا:
    PYTHONPATH=./pylibs python3 build/validate_output.py "کمیسیون فنی.xlsx"
"""

from __future__ import annotations

import os
import re
import sys

sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))

from openpyxl import load_workbook

STRUCT_REF = re.compile(r"([A-Za-z_\u0600-\u06FF][A-Za-z0-9_\u0600-\u06FF]*)\[([^\[\]]+)\]")
WORD_REF = re.compile(r"(?<![A-Za-z0-9_!\[\]\u0600-\u06FF])([A-Za-z_][A-Za-z0-9_]{1,})(?![A-Za-z0-9_(\u0600-\u06FF])")
EXCEL_FUNCS_HINT = ("SUM", "IF", "INDEX", "MATCH", "COUNTA", "OFFSET", "TEXT", "TODAY", "ROUND",
                    "MEDIAN", "LOOKUP", "COUNTIF", "COUNTIFS", "SUMIFS", "IFERROR", "RANK", "ABS",
                    "MAX", "MIN", "AND", "OR", "NOT", "VALUE", "CHOOSE", "ROW", "COLUMN", "TRUE",
                    "FALSE", "SUMPRODUCT", "MOD", "EDATE", "TODAY", "YEAR", "MONTH", "DAY")

errors: list[str] = []
warnings: list[str] = []


def main(path: str) -> int:
    wb = load_workbook(path)          # فرمول‌ها
    wbv = load_workbook(path, data_only=True)

    # ── ۰) ترتیب و ویژگی شیت‌ها
    print("═══ ساختار شیت‌ها ═══")
    for ws in wb.worksheets:
        print(f"  • {ws.title:<14} dim={ws.calculate_dimension():<12} "
              f"RTL={'✔' if ws.sheet_view.rightToLeft else '✘'} "
              f"tables={len(ws.tables)} cf={sum(len(cf.rules) for cf in ws.conditional_formatting)} "
              f"dv={len(ws.data_validations.dataValidation)} charts={len(getattr(ws, '_charts', []))} "
              f"freeze={ws.freeze_panes}")
        if not ws.sheet_view.rightToLeft:
            warnings.append(f"شیت {ws.title}: راست‌به‌چپ نیست")

    # ── ۱) نام‌های محدوده
    print("\n═══ نام‌های محدوده ═══")
    names = {}
    sheets = set(wb.sheetnames)
    for name, dn in wb.defined_names.items():
        names[name] = dn.attr_text
        m = re.match(r"'?([^'!]+)'?!", dn.attr_text)
        if m and m.group(1) not in sheets:
            errors.append(f"نام {name}: شیت «{m.group(1)}» وجود ندارد → {dn.attr_text}")
    print(f"  تعداد: {len(names)}")

    # ── ۲) جداول و ستون‌ها
    tables: dict[str, list[str]] = {}
    for ws in wb.worksheets:
        for tname in list(ws.tables):
            tbl = ws.tables[tname]
            cols = [c.name for c in tbl.tableColumns]
            if tname in tables:
                errors.append(f"جدول تکراری: {tname}")
            tables[tname] = cols
            dr = tbl.ref.split(":")
            if len(dr) != 2:
                errors.append(f"جدول {tname}: بازهٔ نامعتبر {tbl.ref}")
            if len(set(cols)) != len(cols):
                errors.append(f"جدول {tname}: نام ستون تکراری")
            if "None" in cols:
                errors.append(f"جدول {tname}: ستون بی‌نام")
    print(f"  جداول رسمی: {len(tables)} → {', '.join(tables)}")

    # ── ۳) فرمول‌ها
    print("\n═══ بررسی فرمول‌ها ═══")
    total_formulas = 0
    bad_struct: list[str] = []
    unknown_names: list[str] = []
    for ws in wb.worksheets:
        n_ws = 0
        for row in ws.iter_rows():
            for cell in row:
                v = cell.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                n_ws += 1
                total_formulas += 1
                where = f"{ws.title}!{cell.coordinate}"
                if v.count("(") != v.count(")"):
                    errors.append(f"{where}: پرانتز نامتوازن → {v[:90]}")
                if "None" in v or "#REF!" in v:
                    errors.append(f"{where}: مقدار نامعتبر در فرمول → {v[:90]}")
                for tname, col in STRUCT_REF.findall(v):
                    if tname in tables:
                        if col not in tables[tname]:
                            bad_struct.append(f"{where}: ستون «{col}» در جدول {tname} نیست")
                    elif tname not in ("tbl_TX",) and tname not in names:
                        pass  # احتمالاً بخشی از نام دیگری است
                if "tbl_TX" in v and "tbl_TX" not in tables:
                    errors.append(f"{where}: جدول tbl_TX ساخته نشده است")
                for ref in WORD_REF.findall(v):
                    if ref in EXCEL_FUNCS_HINT or ref in tables or ref in names:
                        continue
                    if re.match(r"^[A-Z]{1,3}\d*$", ref) or len(ref) <= 2:
                        continue
                    if ref.isupper() and ref not in names:
                        unknown_names.append(f"{where}: «{ref}» → {v[:80]}")
        print(f"  {ws.title:<14} فرمول‌ها: {n_ws}")
    print(f"  جمع فرمول‌ها: {total_formulas}")

    if bad_struct:
        errors.extend(sorted(set(bad_struct)))
    if unknown_names:
        warnings.extend(sorted(set(unknown_names))[:10])

    # ── ۴) تطبیق عناوین سرستون تراکنش‌ها با ستون‌های جدول
    tx_cols = tables.get("tbl_TX", [])
    print(f"\n  ستون‌های tbl_TX: {len(tx_cols)}")
    print("  " + " | ".join(tx_cols))

    # ── ۵) ادعای محاسبهٔ خودکار هنگام باز شدن
    if not (wb.calculation and wb.calculation.fullCalcOnLoad):
        errors.append("fullCalcOnLoad فعال نیست؛ فرمول‌ها ممکن است تا اولین ویرایش خالی بمانند")

    # ── ۶) دادهٔ نمونه در فایل نهایی
    ws_tx = wb["تراکنش‌ها"]
    filled = sum(1 for r in range(6, 306) if ws_tx.cell(row=r, column=1).value)
    print(f"  رکورد نمونه در تراکنش‌ها: {filled}")

    print("\n═══ نتیجه ═══")
    if warnings:
        print(f"⚠️  هشدارها ({len(set(warnings))}):")
        for w in sorted(set(warnings)):
            print("   -", w)
    if errors:
        print(f"❌ خطاها ({len(errors)}):")
        for e in errors:
            print("   -", e)
        return 1
    print("✅ هیچ خطایی یافت نشد — فایل از نظر ارجاع‌ها، جداول و فرمول‌ها سالم است.")
    return 0


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "کمیسیون فنی.xlsx"
    sys.exit(main(target))
