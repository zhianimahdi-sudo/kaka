#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
inspect_workbook.py — استخراج کامل «منطق» یک فایل اکسل.

این ابزار برای فاز شناخت پروژهٔ بازطراحی نوشته شده است: هر فرمول، نام محدوده،
اعتبارسنجی داده، قاعدهٔ فرمت شرطی، جدول رسمی، نمودار، سلول ادغام‌شده و قالب عددی
را از فایل مبدأ بیرون می‌کشد تا هنگام طراحی فایل جدید هیچ محاسبه‌ای از دست نرود.

نمونهٔ اجرا:
    PYTHONPATH=/home/user/pylibs python3 tools/inspect_workbook.py path/to/file.xlsx
    PYTHONPATH=/home/user/pylibs python3 tools/inspect_workbook.py file.xlsx --json out.json
    PYTHONPATH=/home/user/pylibs python3 tools/inspect_workbook.py file.xlsx --formulas-only
"""

from __future__ import annotations

import argparse
import json
import os
import sys
from collections import Counter

try:
    from openpyxl import load_workbook
    from openpyxl.utils import get_column_letter
except ImportError:  # pragma: no cover
    sys.exit(
        "خطا: کتابخانهٔ openpyxl پیدا نشد.\n"
        "اجرا کنید:  PYTHONPATH=/home/user/pylibs python3 tools/inspect_workbook.py ..."
    )


# ---------------------------------------------------------------- utilities

def _rule_def(rule):
    """تبدیل یک قاعدهٔ فرمت شرطی به دیکشنری خوانا."""
    dxf = rule.dxf
    fmt = {}
    if dxf is not None:
        if dxf.font is not None:
            fmt["font"] = {
                "bold": dxf.font.b,
                "italic": dxf.font.i,
                "color": _color(dxf.font.color),
            }
        if dxf.fill is not None and dxf.fill.fgColor is not None:
            fmt["fill"] = _color(dxf.fill.fgColor)
        if dxf.border is not None:
            edges = {}
            for side in ("left", "right", "top", "bottom"):
                s = getattr(dxf.border, side, None)
                if s is not None and s.style:
                    edges[side] = s.style
            if edges:
                fmt["border"] = edges
    data = {
        "type": rule.type,
        "operator": rule.operator,
        "formula": list(rule.formula) if rule.formula else None,
        "text": rule.text if rule.type == "containsText" else None,
        "rank": rule.rank,
        "percent": rule.percent,
        "stopIfTrue": rule.stopIfTrue,
        "priority": rule.priority,
        "format": fmt or None,
    }
    if rule.type == "dataBar":
        data["dataBar"] = {
            "color": _color(rule.dataBar.color) if rule.dataBar else None,
            "minLength": getattr(rule.dataBar, "minLength", None),
            "maxLength": getattr(rule.dataBar, "maxLength", None),
        }
    if rule.type == "colorScale":
        data["colorScale"] = (
            [{"type": p.type, "value": p.val, "color": _color(p.color)} for p in rule.colorScale.cfvo]
            if rule.colorScale
            else None
        )
    if rule.type == "iconSet":
        data["iconSet"] = rule.iconSet.iconSet if rule.iconSet else None
    return {k: v for k, v in data.items() if v is not None}


def _color(color):
    if color is None:
        return None
    if getattr(color, "type", None) == "rgb":
        return color.rgb
    if getattr(color, "type", None) == "theme":
        return f"theme{color.theme}"
    if getattr(color, "type", None) == "indexed":
        return f"indexed{color.indexed}"
    return str(color)


def _dv_dict(dv):
    return {
        "range": str(dv.sqref),
        "type": dv.type,
        "operator": dv.operator,
        "formula1": dv.formula1,
        "formula2": dv.formula2,
        "allowBlank": dv.allow_blank,
        "showDropDown": getattr(dv, "showDropDown", None),
        "prompt": dv.prompt,
        "promptTitle": dv.promptTitle,
        "error": dv.error,
        "errorTitle": dv.errorTitle,
        "errorStyle": dv.errorStyle,
    }


def _chart_dict(ch):
    """خلاصهٔ یک نمودار: نوع، عنوان، محدودهٔ سری‌ها و محورها."""
    out = {"class": type(ch).__name__}
    try:
        out["type"] = ch.type if isinstance(ch.type, str) else getattr(ch.type, "v", str(ch.type))
    except Exception:
        out["type"] = None
    try:
        if ch.title is not None and ch.title.tx is not None and ch.title.tx.rich is not None:
            out["title"] = "".join(
                r.t or "" for p in ch.title.tx.rich.p for r in (p.r or [])
            ).strip()
    except Exception:
        pass
    series = []
    for s in getattr(ch, "series", []) or []:
        item = {}
        for key, ref in (("name", s.tx), ("categories", s.cat), ("values", s.val)):
            try:
                if ref is None:
                    continue
                if key == "name" and getattr(ref, "strRef", None) is not None:
                    item["name"] = ref.strRef.f
                elif getattr(ref, "numRef", None) is not None:
                    item[key] = ref.numRef.f
                elif getattr(ref, "strRef", None) is not None:
                    item[key] = ref.strRef.f
                elif getattr(ref, "f", None):
                    item[key] = ref.f
            except Exception:
                pass
        if item:
            series.append(item)
    if series:
        out["series"] = series
    out["anchor"] = getattr(getattr(ch, "anchor", None), "_from", None) and (
        f"{get_column_letter(ch.anchor._from.col + 1)}{ch.anchor._from.row + 1}"
        if hasattr(ch.anchor, "_from") and ch.anchor._from is not None
        else None
    )
    return out


# ---------------------------------------------------------------- dump core

def inspect(path: str, formulas_only: bool = False, max_rows: int = 0) -> dict:
    if not os.path.isfile(path):
        sys.exit(f"خطا: فایل پیدا نشد: {path}")
    keep_vba = path.lower().endswith((".xlsm", ".xltm"))
    wb = load_workbook(path, data_only=False, keep_vba=keep_vba)
    try:
        wb_formulas = load_workbook(path, data_only=False)
    except Exception:
        wb_formulas = wb
    try:
        wb_values = load_workbook(path, data_only=True)
    except Exception:
        wb_values = None

    report: dict = {
        "file": os.path.abspath(path),
        "size_bytes": os.path.getsize(path),
        "sheet_count": len(wb.sheetnames),
        "sheets_order": wb.sheetnames,
        "defined_names": [],
        "sheets": {},
        "input_counts": {},
    }

    # ---- نام‌های محدوده (سراسری و محلی)
    for name, dn in getattr(wb.defined_names, "items", lambda: [])():
        scope = dn.localSheetId
        report["defined_names"].append(
            {
                "name": name,
                "refersTo": dn.attr_text,
                "scope": wb.sheetnames[scope] if isinstance(scope, int) and scope < len(wb.sheetnames) else "Workbook",
                "hidden": dn.hidden,
                "comment": getattr(dn, "comment", None),
            }
        )

    global_stats = Counter()

    for ws in wb.worksheets:
        formulas = {}
        values = {}
        numbers = Counter()
        ws_values = wb_values[ws.title] if (wb_values and ws.title in wb_values.sheetnames) else None

        for row in ws.iter_rows():
            for cell in row:
                val = cell.value
                if val is None:
                    continue
                if isinstance(val, str) and val.startswith("="):
                    formulas[cell.coordinate] = val
                if ws_values is not None:
                    v = ws_values[cell.coordinate].value
                    if v is not None:
                        values[cell.coordinate] = v
                # قالب عددی/تاریخی معنادار را ثبت کن
                nf = cell.number_format
                if nf and nf != "General" and cell.value is not None:
                    numbers[nf] += 1

        sheet = {
            "state": ws.sheet_state,
            "rightToLeft": bool(ws.sheet_view.rightToLeft),
            "dimensions": ws.calculate_dimension(),
            "max_row": ws.max_row,
            "max_column": ws.max_column,
            "freeze_panes": ws.freeze_panes,
            "showGridLines": bool(ws.sheet_view.showGridLines),
            "merged_cells": [str(r) for r in ws.merged_cells.ranges],
            "protection": {"sheet": bool(ws.protection.sheet), "autoFilter": ws.protection.autoFilter},
            "print": {
                "orientation": ws.page_setup.orientation,
                "paperSize": ws.page_setup.paperSize,
                "fitToWidth": ws.page_setup.fitToWidth,
                "fitToHeight": ws.page_setup.fitToHeight,
                "scale": ws.page_setup.scale,
                "print_area": ws.print_area,
                "print_title_rows": ws.print_title_rows,
                "titles_cols": ws.print_title_cols,
                "header": ws.oddHeader.center.text if ws.oddHeader and ws.oddHeader.center else None,
                "footer": ws.oddFooter.center.text if ws.oddFooter and ws.oddFooter.center else None,
            },
            "columns": [
                {
                    "letter": get_column_letter(dim.min),
                    "width": round(dim.width, 2) if dim.width else None,
                    "hidden": bool(dim.hidden),
                }
                for dim in ws.column_dimensions.values()
                if dim.width or dim.hidden
            ],
            "row_heights": {
                str(idx): round(d.height, 2) for idx, d in ws.row_dimensions.items() if d.height
            },
            "tables": [
                {
                    "name": t.name,
                    "ref": t.ref,
                    "style": t.tableStyleInfo.name if t.tableStyleInfo else None,
                    "showRowStripes": t.tableStyleInfo.showRowStripes if t.tableStyleInfo else None,
                    "totalsRowShown": t.totalsRowShown,
                    "columns": [c.name for c in t.tableColumns],
                }
                for t in ws.tables.values()
            ],
            "data_validations": [_dv_dict(dv) for dv in ws.data_validations.dataValidation],
            "conditional_formatting": [
                {"range": str(cf.sqref), "rules": [_rule_def(r) for r in cf.rules]}
                for cf in ws.conditional_formatting
            ],
            "charts": [_chart_dict(ch) for ch in getattr(ws, "_charts", [])],
            "images": len(getattr(ws, "_images", [])),
            "hyperlinks": [{"cell": c.coordinate, "target": c.hyperlink.target, "text": c.value}
                           for row in ws.iter_rows() for c in row if c.hyperlink],
            "comments": [{"cell": c.coordinate, "text": c.comment.text}
                         for row in ws.iter_rows() for c in row if c.comment],
            "number_formats_used": dict(numbers.most_common()),
            "formula_count": len(formulas),
            "formulas": formulas,
        }
        if not formulas_only:
            # مقادیر ذخیره‌شده فقط برای ستون‌هایی که فرمول دارند مفید است (برای تست رگرسیون)
            sheet["values_of_formula_cells"] = {
                k: values.get(k, "<بدون مقدار ذخیره‌شده — فایل هرگز بازمحاسبه نشده>")
                for k in formulas
            }
            # فهرست سلول‌های ورودی (غیرفرمول) برای تشخیص «ستون‌های ورودی» هنگام بازطراحی
            inputs = {}
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if v is None or (isinstance(v, str) and v.startswith("=")):
                        continue
                    if isinstance(v, str) and len(v) > 60:
                        v = v[:60] + "…"
                    inputs[c.coordinate] = v
            sheet["input_cells"] = inputs
            report.setdefault("input_counts", {})[ws.title] = len(inputs)

        report["sheets"][ws.title] = sheet
        global_stats["formulas"] += len(formulas)
        global_stats["validations"] += len(sheet["data_validations"])
        global_stats["cf_rules"] += sum(len(c["rules"]) for c in sheet["conditional_formatting"])
        global_stats["tables"] += len(sheet["tables"])
        global_stats["charts"] += len(sheet["charts"])

    report["totals"] = dict(global_stats)
    return report


# ---------------------------------------------------------------- rendering

def render(report: dict, formulas_only: bool = False) -> str:
    L: list[str] = []
    p = L.append
    p("═" * 78)
    p("  گزارش شناخت فایل اکسل")
    p("═" * 78)
    p(f"فایل         : {report['file']}")
    p(f"حجم          : {report['size_bytes']:,} بایت")
    p(f"تعداد شیت‌ها  : {report['sheet_count']}")
    p(f"ترتیب شیت‌ها  : {' | '.join(report['sheets_order'])}")
    p(f"جمع فرمول‌ها  : {report['totals'].get('formulas', 0):,}")
    p(f"اعتبارسنجی   : {report['totals'].get('validations', 0)}")
    p(f"قواعد شرطی   : {report['totals'].get('cf_rules', 0)}")
    p(f"جداول رسمی   : {report['totals'].get('tables', 0)}")
    p(f"نمودارها     : {report['totals'].get('charts', 0)}")
    p("")

    if report["defined_names"]:
        p("─" * 78)
        p("■ نام‌های محدوده (Named Ranges)")
        p("─" * 78)
        for dn in report["defined_names"]:
            p(f"  • {dn['name']:<28} ← {dn['refersTo']}   [{dn['scope']}]")
        p("")

    for title, s in report["sheets"].items():
        p("─" * 78)
        p(f"■ شیت: {title}   ({s['state']})   dim={s['dimensions']}   "
          f"RTL={'بله' if s['rightToLeft'] else 'خیر'}   تعداد فرمول={s['formula_count']}")
        p("─" * 78)
        if s["merged_cells"]:
            p(f"  ادغام‌شده : {', '.join(s['merged_cells'])}")
        if s["freeze_panes"]:
            p(f"  Freeze    : {s['freeze_panes']}")
        if s["tables"]:
            for t in s["tables"]:
                p(f"  جدول رسمی : {t['name']} [{t['ref']}] style={t['style']} "
                  f"stripes={t['showRowStripes']} totals={t['totalsRowShown']}")
                p(f"              ستون‌ها: {' | '.join(t['columns'])}")
        if s["data_validations"]:
            p("  اعتبارسنجی داده:")
            for dv in s["data_validations"]:
                p(f"    – {dv['range']:<14} type={dv['type']} op={dv['operator']} "
                  f"f1={dv['formula1']} f2={dv['formula2']} blank={dv['allowBlank']}")
                if dv["error"]:
                    p(f"        پیام خطا: {dv['error']}")
        if s["conditional_formatting"]:
            p("  فرمت شرطی:")
            for cf in s["conditional_formatting"]:
                for r in cf["rules"]:
                    p(f"    – {cf['range']:<16} {r['type']}"
                      f"{'/' + str(r['operator']) if r.get('operator') else ''}"
                      f"  formula={r.get('formula')}  fmt={r.get('format')}")
        if s["charts"]:
            p("  نمودارها:")
            for c in s["charts"]:
                p(f"    – {c.get('type')} «{c.get('title') or 'بی‌عنوان'}» "
                  f"series={len(c.get('series', []))}")
                for ser in c.get("series", []):
                    p(f"        * name={ser.get('name')} cat={ser.get('categories')} val={ser.get('values')}")
        if s["number_formats_used"]:
            fmts = list(s["number_formats_used"].items())[:12]
            p(f"  قالب‌های عددی پرکاربرد: {', '.join(f'{k}×{v}' for k, v in fmts)}")
        if s["columns"]:
            cols = ", ".join(f"{c['letter']}={c['width']}" for c in s["columns"][:20])
            p(f"  عرض ستون‌ها: {cols}")
        if not formulas_only and s["formulas"]:
            vals = s.get("values_of_formula_cells", {})
            p(f"  ── فرمول‌ها ({s['formula_count']}) ──")
            for coord, f in s["formulas"].items():
                p(f"    {coord:<8} {f}")
                stored = vals.get(coord)
                if stored is None:
                    continue
                if isinstance(stored, str) and stored.startswith("<"):
                    continue  # قالب سلول در فایل مبدأ در دسترس نیست
                p(f"             ⇒ مقدار ذخیره‌شده: {stored}")
        p("")
    return "\n".join(L)


def main() -> None:
    ap = argparse.ArgumentParser(description="استخراج کامل منطق یک فایل اکسل")
    ap.add_argument("path", help="مسیر فایل xlsx/xlsm")
    ap.add_argument("--json", dest="json_out", help="ذخیرهٔ گزارش کامل به‌صورت JSON")
    ap.add_argument("--formulas-only", action="store_true", help="فقط فهرست فرمول‌ها را نشان بده")
    ap.add_argument("--out", dest="txt_out", help="ذخیرهٔ گزارش متنی در فایل")
    args = ap.parse_args()

    rep = inspect(args.path, formulas_only=args.formulas_only)
    text = render(rep, formulas_only=args.formulas_only)
    if args.txt_out:
        with open(args.txt_out, "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"گزارش متنی ذخیره شد: {args.txt_out}")
    else:
        print(text)
    if args.json_out:
        with open(args.json_out, "w", encoding="utf-8") as fh:
            json.dump(rep, fh, ensure_ascii=False, indent=1, default=str)
        print(f"گزارش JSON ذخیره شد: {args.json_out}  "
              f"(فرمول‌ها: {rep['totals'].get('formulas', 0)})")


if __name__ == "__main__":
    main()
