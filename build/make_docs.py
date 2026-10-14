#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
make_docs.py — تولید خودکار مستندات از فایل ساخته‌شده.

چون فرمول‌ها مستقیماً از فایل خوانده می‌شوند، این مستندات هرگز از محتوای واقعی
کارپوشه جدا نمی‌شوند:
    • docs/formula-map.md      ← زنجیرهٔ محاسبه و نقشهٔ فرمول‌ها
    • docs/values-reference.md ← مقادیر مرجع دادهٔ نمونه (برای آزمون رگرسیون)

نمونهٔ اجرا:
    PYTHONPATH=./pylibs python3 build/make_docs.py "کمیسیون فنی.xlsx"
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))

from openpyxl import load_workbook

import content as C
import model_reference as M
from theme import L


def fmt(v, nf=None):
    if isinstance(v, float):
        if nf and "0.0%" in nf:
            return f"{v * 100:,.1f}٪"
        return f"{v:,.0f}" if abs(v) >= 1000 or v == int(v) else f"{v:,.4f}"
    if isinstance(v, int):
        return f"{v:,}"
    return str(v)


# ═══════════════════════════════ نقشهٔ فرمول‌ها

MEANING = {
    ("تراکنش‌ها", "K"): "مبلغ مبنا = مبلغ کارشناسی × ضریب خدمت",
    ("تراکنش‌ها", "L"): "نرخ پایهٔ خدمت از شیت فهرست‌ها (INDEX + MATCH)",
    ("تراکنش‌ها", "M"): "نرخ مؤثر = نرخ پایه × ضریب پله (LOOKUP روی جدول پله‌ها)",
    ("تراکنش‌ها", "N"): "کمیسیون خام = مبلغ مبنا × نرخ مؤثر",
    ("تراکنش‌ها", "O"): "اعمال کف و سقف تنظیمات (و صفر کردن رکوردهای ابطال)",
    ("تراکنش‌ها", "P"): "مالیات و عوارض = کمیسیون تعدیل‌شده × نرخ مالیات",
    ("تراکنش‌ها", "Q"): "جمع با مالیات (مبلغ ناخالص قابل وصول)",
    ("تراکنش‌ها", "R"): "درصد سهم عضو (بازنویسی ردیف، در غیر این صورت سهم پیش‌فرض)",
    ("تراکنش‌ها", "S"): "مبلغ سهم عضو = جمع با مالیات × درصد سهم",
    ("تراکنش‌ها", "T"): "سهم شرکت = جمع با مالیات − سهم عضو",
    ("تراکنش‌ها", "U"): "سال شمسی از جدول تقویم (INDEX + MATCH روی کلید تاریخ)",
    ("تراکنش‌ها", "V"): "ماه شمسی از جدول تقویم",
    ("تراکنش‌ها", "W"): "کلید دوره «سال/ماه» برای تجمیع‌های ماهانه",
    ("تراکنش‌ها", "X"): "فصل شمسی از جدول تقویم",
    ("تراکنش‌ها", "Y"): "مهلت پرداخت = سری تاریخ + مهلت تنظیمات، برگردانده‌شده به تاریخ شمسی",
    ("تراکنش‌ها", "Z"): "تشخیص سپری‌شدن مهلت (رکوردهای پرداخت‌شده و ابطال کنار گذاشته می‌شوند)",
    ("اعضا", "E"): "تعداد پرونده‌های عضو (بدون ابطال) — COUNTIFS",
    ("اعضا", "F"): "جمع مبلغ کارشناسی عضو — SUMIFS",
    ("اعضا", "G"): "جمع کمیسیون خام عضو",
    ("اعضا", "H"): "جمع سهم عضو",
    ("اعضا", "I"): "جمع سهم شرکت",
    ("اعضا", "J"): "آخرین تاریخی که عضو در آن پرونده داشته است",
    ("ماشین‌حساب", "C15"): "گام ۱ — مبلغ مبنا",
    ("ماشین‌حساب", "C16"): "گام ۲ — ضریب پلهٔ اعمال‌شده",
    ("ماشین‌حساب", "C17"): "گام ۳ — نرخ پایهٔ خدمت",
    ("ماشین‌حساب", "C18"): "گام ۴ — نرخ مؤثر",
    ("ماشین‌حساب", "C19"): "گام ۵ — کمیسیون خام",
    ("ماشین‌حساب", "C20"): "گام ۶ — اثر کف (افزایش)",
    ("ماشین‌حساب", "C21"): "گام ۷ — اثر سقف (کاهش)",
    ("ماشین‌حساب", "C22"): "گام ۸ — کمیسیون پس از کف و سقف",
    ("ماشین‌حساب", "C23"): "گام ۹ — مالیات و عوارض",
    ("ماشین‌حساب", "C24"): "گام ۱۰ — جمع با مالیات",
    ("ماشین‌حساب", "C25"): "گام ۱۱ — درصد سهم عضو",
    ("ماشین‌حساب", "C26"): "گام ۱۲ — مبلغ سهم عضو",
    ("ماشین‌حساب", "C27"): "گام ۱۳ — سهم شرکت",
}


def formula_map(wb, path: str) -> str:
    out: list[str] = []
    p = out.append
    p("# نقشهٔ فرمول‌ها — کارپوشهٔ کمیسیون فنی")
    p("")
    p("> این سند **به‌صورت خودکار از خود فایل اکسل** تولید شده است؛ هر فرمولی که اینجا می‌بینید،")
    p("> دقیقاً همان چیزی است که در فایل اجرا می‌شود. برای بازتولید: `python3 build/make_docs.py`")
    p("")
    p("## ۱. فهرست شیت‌ها")
    p("")
    p("| # | شیت | وظیفه | جدول رسمی | فرمول | اعتبارسنجی | فرمت شرطی | نمودار |")
    p("|---|---|---|---|---|---|---|---|")
    for i, name in enumerate(wb.sheetnames, start=1):
        ws = wb[name]
        n_f = sum(1 for row in ws.iter_rows() for c in row
                  if isinstance(c.value, str) and c.value.startswith("="))
        p(f"| {i} | {name} | {dict((g[0], g[1]) for g in C.GUIDE_SHEETS).get(name, '—')} | "
          f"{', '.join(ws.tables) or '—'} | {n_f:,} | "
          f"{len(ws.data_validations.dataValidation) or '—'} | "
          f"{sum(len(cf.rules) for cf in ws.conditional_formatting) or '—'} | "
          f"{len(getattr(ws, '_charts', [])) or '—'} |")
    p("")

    p("## ۲. زنجیرهٔ محاسبه در شیت تراکنش‌ها")
    p("")
    p("ردیف نمونه: ردیف ۶ (اولین رکورد). همین فرمول‌ها تا ردیف ۳۰۵ کشیده شده‌اند.")
    p("")
    p("| ستون | عنوان | فرمول (ردیف ۶) | نقش |")
    p("|---|---|---|---|")
    ws = wb[L.S_TX]
    for col in "KLMNOPQRSTUVWXYZ":
        header = ws[f"{col}{L.TX_HDR}"].value
        f = ws[f"{col}{L.TX_FIRST}"].value
        meaning = MEANING.get((L.S_TX, col), "—")
        p(f"| {col} | {header} | `{f}` | {meaning} |")
    p("")

    p("## ۳. فرمول‌های شیت‌های تجمیع و کنترل")
    p("")
    for sheet in (L.S_MEM, L.S_CALC, L.S_REP, L.S_DASH, L.S_QC):
        ws = wb[sheet]
        p(f"### {sheet}")
        p("")
        p("| سلول / بازه | فرمول |")
        p("|---|---|")
        seen = set()
        shown = 0
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not (isinstance(v, str) and v.startswith("=")):
                    continue
                key = (c.column, v.split("!")[0][:40])
                if (c.column, str(v)[:60]) in seen:
                    continue
                seen.add((c.column, str(v)[:60]))
                meaning = MEANING.get((sheet, c.coordinate)) or MEANING.get((sheet, c.column_letter))
                p(f"| `{c.coordinate}`{f' — {meaning}' if meaning else ''} | `{v}` |")
                shown += 1
                if shown >= 26:
                    break
            if shown >= 26:
                break
        if shown == 0:
            p("| — | (بدون فرمول) |")
        p("")
        if shown >= 26:
            p(f"*الگوی تکرارشونده — نمایش ۲۶ فرمول یکتا؛ بقیهٔ سلول‌ها همان الگو با آدرس‌های نسبی هستند.*")
            p("")

    p("## ۴. نام‌های محدوده (Named Ranges)")
    p("")
    p("| نام | مرجع | نوع |")
    p("|---|---|---|")
    for name, dn in wb.defined_names.items():
        kind = "پویا (خودگسترش‌یابنده)" if dn.attr_text.startswith("OFFSET") else "ثابت"
        p(f"| `{name}` | `{dn.attr_text}` | {kind} |")
    p("")
    p("**قاعدهٔ پیشوندها:** `p_` پارامتر تنظیمات · `lst_` فهرست کشویی · `t_` ستون/جدول داده · "
      "`cal_` تقویم · `f_` فیلتر داشبورد · `m_` ورودی ماشین‌حساب · `tbl_`/`T_` جدول رسمی.")
    p("")
    return "\n".join(out)


# ═══════════════════════════════ مقادیر مرجع

def values_reference() -> str:
    out: list[str] = []
    p = out.append
    p("# مقادیر مرجع — دادهٔ نمونهٔ کارپوشه")
    p("")
    p("> محاسبه‌شده با **مدل مرجع پایتون** (`build/model_reference.py`) که مستقل از فایل اکسل نوشته شده است.")
    p("> آزمون خودکار `build/engine_check.py` همین مقادیر را با خروجی موتور فرمول اکسل مقایسه می‌کند.")
    p("> اگر فایل را با همین دادهٔ نمونه باز کنید، باید همین اعداد را ببینید.")
    p("")
    p("## ۱. زنجیرهٔ محاسبه ردیف‌به‌ردیف")
    p("")
    p("| ردیف | پرونده | خدمت | مبلغ کارشناسی | مبنا | نرخ مؤثر | کمیسیون خام | پس از کف/سقف | مالیات | جمع با مالیات | سهم عضو | سهم شرکت |")
    p("|---|---|---|---|---|---|---|---|---|---|---|---|")
    for i, r in enumerate(M.rows(), start=L.TX_FIRST):
        p(f"| {i} | {r['case']} | {r['service']} | {r['amount']:,} | {r['base']:,.0f} | "
          f"{r['eff_rate'] * 100:.2f}٪ | {r['gross']:,.0f} | {r['after_clamp']:,.0f} | {r['tax']:,.0f} | "
          f"{r['total']:,.0f} | {r['member_amt']:,.0f} | {r['company_amt']:,.0f} |")
    p("")
    p("## ۲. جمع‌های کل دوره")
    p("")
    t = M.totals()
    mt = M.member_totals()
    p("| شاخص | کل (همهٔ رکوردها) | فقط رکوردهای معتبر (بدون ابطال) |")
    p("|---|---|---|")
    p(f"| تعداد پرونده | {t['count']:,} | {mt['count']:,} |")
    p(f"| مبلغ کارشناسی | {t['amount']:,} | {mt['amount']:,} |")
    p(f"| کمیسیون خام | {t['gross']:,.0f} | {mt['gross']:,.0f} |")
    p(f"| جمع با مالیات | {t['total']:,.0f} | {mt['total']:,.0f} |")
    p(f"| سهم اعضا | {t['member_amt']:,.0f} | {mt['member_amt']:,.0f} |")
    p(f"| سهم شرکت | {t['company_amt']:,.0f} | {mt['company_amt']:,.0f} |")
    p("")
    p("## ۳. ماشین‌حساب سناریو (ورودی‌های پیش‌فرض شیت)")
    p("")
    p("| گام | مقدار |")
    p("|---|---|")
    for k, v in M.scenario().items():
        p(f"| {k} | {v:,.4f} |")
    p("")
    p("## ۴. مقادیر انتظاری آزمون‌های کنترل کیفیت")
    p("")
    p("| آزمون | مقدار انتظاری | وضعیت |")
    p("|---|---|---|")
    qc = M.qc_expectations()
    thresholds = {"پروندهٔ دارای سرسید سپری‌شده": "آستانه: ۳ مورد", "پوشش تقویم برای سال گزارش": "باید ≥ ۳۶۵ روز"}
    for k, v in qc.items():
        val = f"{v:,.0f}" if isinstance(v, float) else f"{v:,}"
        ok = "✅" if (v == 0 or k in ("پوشش تقویم برای سال گزارش")) else ("✅ (زیر آستانه)" if k == "پروندهٔ دارای سرسید سپری‌شده" else "❌")
        p(f"| {k} | {val} | {ok} {thresholds.get(k, '')} |")
    p("")
    p("## ۵. تجمیع بر پایهٔ عضو (بلوک ۱ شیت گزارش‌ها)")
    p("")
    p("| عضو | تعداد | مبلغ کارشناسی | کمیسیون خام | جمع با مالیات | سهم عضو | سهم شرکت |")
    p("|---|---|---|---|---|---|---|")
    for name, d in M.by_member().items():
        p(f"| {name} | {d['count']:,} | {d['amount']:,.0f} | {d['gross']:,.0f} | "
          f"{d['total']:,.0f} | {d['member_amt']:,.0f} | {d['company_amt']:,.0f} |")
    p("")
    p("## ۶. تجمیع بر پایهٔ ماه شمسی (بلوک ۲ شیت گزارش‌ها)")
    p("")
    p("| کلید دوره | تعداد | مبلغ کارشناسی | کمیسیون خام | جمع با مالیات |")
    p("|---|---|---|---|---|")
    for ym, d in M.by_month().items():
        p(f"| {ym} | {d['count']:,} | {d['amount']:,.0f} | {d['gross']:,.0f} | {d['total']:,.0f} |")
    p("")
    return "\n".join(out)


def main(path: str):
    wb = load_workbook(path)
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docs = os.path.join(root, "docs")
    os.makedirs(docs, exist_ok=True)
    fm = os.path.join(docs, "formula-map.md")
    vr = os.path.join(docs, "values-reference.md")
    with open(fm, "w", encoding="utf-8") as fh:
        fh.write(formula_map(wb, path))
    with open(vr, "w", encoding="utf-8") as fh:
        fh.write(values_reference())
    print(f"✅ مستندات ساخته شد:\n   {fm}\n   {vr}")


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "کمیسیون فنی.xlsx"
    main(target)
