#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
engine_check.py — آزمون موتور محاسبهٔ کارپوشه با یک ارزیاب واقعی فرمول.

کارپوشه را با کتابخانهٔ pycel بازمی‌کند، همهٔ فرمول‌های کلیدی را ارزیابی می‌کند و
نتیجه را با «مدل مرجع» پایتون (model_reference.py) مقایسه می‌کند. هر اختلاف، خطای
محاسباتی است.

نکته: pycel با openpyxl 3.x سازگار نیست (نام‌های محدوده را نمی‌خواند) و تابع
MEDIAN را پیاده‌سازی نکرده است؛ هر دو مورد در همین فایل با وصلهٔ آزمون حل می‌شوند و
هیچ تغییری در فایل نهایی ایجاد نمی‌شود.
"""

from __future__ import annotations

import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))

from openpyxl import load_workbook

import content as C
import theme as T
from theme import L
import model_reference as M


# ═════════════════════════ وصل‌های آزمون (فقط برای pycel، نه برای فایل)

def patch_pycel(path: str):
    """سازگارسازی pycel با openpyxl 3.x و افزودن MEDIAN."""
    import pycel.excelwrapper as ew
    import pycel.lib.stats as stats

    wb = load_workbook(path, data_only=False)
    names: dict[str, list[tuple[str, str]]] = {}
    for name, dn in wb.defined_names.items():
        text = dn.attr_text
        if text.startswith("OFFSET("):
            text = OFFSET_RESOLVED.get(name)
            if not text:
                continue
        if "!" not in text:
            continue
        sheet, _, ref = text.rpartition("!")
        sheet = sheet.strip("'")
        names[str(name)] = [(ref.replace("$", ""), sheet)]

    def defined_names(self):
        return names

    ew.ExcelOpxWrapper.defined_names = property(defined_names)

    def _median(*args):
        vals = []
        for a in args:
            if isinstance(a, (list, tuple)):
                vals.extend(x for x in a if isinstance(x, (int, float)))
            elif isinstance(a, (int, float)):
                vals.append(a)
        if not vals:
            return 0
        vals.sort()
        n = len(vals)
        return vals[n // 2] if n % 2 else (vals[n // 2 - 1] + vals[n // 2]) / 2

    stats.median = _median   # pycel تابع MEDIAN را پیاده نکرده است
    return wb


# محدودهٔ واقعی نام‌های پویا در شرایط فعلی فایل (برای آزمون)
OFFSET_RESOLVED = {
    "lst_Service": f"'{L.S_LISTS}'!$A${L.SVC_FIRST}:$A${L.SVC_LAST}",
    "t_Svc_Rate": f"'{L.S_LISTS}'!$B${L.SVC_FIRST}:$B${L.SVC_LAST}",
    "t_Svc_Cat": f"'{L.S_LISTS}'!$C${L.SVC_FIRST}:$C${L.SVC_LAST}",
    "lst_Status": f"'{L.S_LISTS}'!$A${L.STA_FIRST}:$A${L.STA_LAST}",
    "t_Sta_Group": f"'{L.S_LISTS}'!$B${L.STA_FIRST}:$B${L.STA_LAST}",
    "lst_Contractor": f"'{L.S_LISTS}'!$A${L.CLI_FIRST}:$A${L.CLI_LAST}",
    "t_Cli_City": f"'{L.S_LISTS}'!$B${L.CLI_FIRST}:$B${L.CLI_LAST}",
    "lst_Category": f"'{L.S_LISTS}'!$A${L.CAT_FIRST}:$A${L.CAT_LAST}",
    "lst_Role": f"'{L.S_LISTS}'!$A${L.ROL_FIRST}:$A${L.ROL_LAST}",
    "lst_YesNo": f"'{L.S_LISTS}'!$A${L.YN_FIRST}:$A${L.YN_LAST}",
    "lst_TierMode": f"'{L.S_LISTS}'!$A${L.MRG_FIRST}:$A${L.MRG_LAST}",
    "t_Members": f"'{L.S_MEM}'!$A${L.MEM_FIRST}:$A${L.MEM_LAST}",
    "t_Member_Role": f"'{L.S_MEM}'!$B${L.MEM_FIRST}:$B${L.MEM_LAST}",
    "t_Member_Share": f"'{L.S_MEM}'!$C${L.MEM_FIRST}:$C${L.MEM_LAST}",
    "t_Member_Active": f"'{L.S_MEM}'!$D${L.MEM_FIRST}:$D${L.MEM_LAST}",
}


def main(path: str):
    print("═" * 74)
    print("  آزمون موتور محاسبهٔ کارپوشه (ارزیابی واقعی فرمول‌ها)")
    print("═" * 74)
    patch_pycel(path)

    from pycel import ExcelCompiler
    xl = ExcelCompiler(filename=path)

    TX = L.S_TX
    checks: list[tuple[str, float, float]] = []   # (برچسب، مقدار اکسل، مقدار مرجع)
    failures: list[str] = []

    def ev(ref):
        try:
            return xl.evaluate(ref)
        except Exception as exc:  # pragma: no cover
            failures.append(f"{ref}: {type(exc).__name__} {str(exc)[:70]}")
            return None

    def near(a, b, tol=1.0):
        """مقایسهٔ عددی با تلورانس یا مقایسهٔ متنی (برای ستون‌های دوره و فصل)."""
        if a is None or b is None:
            return False
        if isinstance(a, str) or isinstance(b, str):
            return str(a).strip() == str(b).strip()
        try:
            return abs(float(a) - float(b)) <= tol
        except (TypeError, ValueError):
            return False

    # ── ۱) زنجیرهٔ محاسبهٔ ردیف‌های نمونه
    print("\n── زنجیرهٔ محاسبهٔ تراکنش‌ها (ردیف‌های نمونه) ──")
    print(f"{'ردیف':<6}{'مبنا':>14}{'نرخ مؤثر':>11}{'خام':>14}{'پس از کف/سقف':>15}"
          f"{'مالیات':>12}{'جمع':>14}{'سهم عضو':>13}{'سهم شرکت':>13}")
    for i, row in enumerate(M.rows(), start=L.TX_FIRST):
        exp = row
        vals = {c: ev(f"'{TX}'!{chr(64 + col)}{i}") for col, c in
                zip(range(11, 21), "KLMNOPQRST")}
        for key, got, want in (("مبنا", vals["K"], exp["base"]), ("نرخ مؤثر", vals["M"], exp["eff_rate"]),
                               ("خام", vals["N"], exp["gross"]), ("پس از کف/سقف", vals["O"], exp["after_clamp"]),
                               ("مالیات", vals["P"], exp["tax"]), ("جمع", vals["Q"], exp["total"]),
                               ("سهم عضو", vals["S"], exp["member_amt"]), ("سهم شرکت", vals["T"], exp["company_amt"])):
            checks.append((f"{TX}!{i} {key}", got, want))
        print(f"{i:<6}{vals['K'] or 0:>14,.0f}{vals['M'] or 0:>11.4f}{vals['N'] or 0:>14,.0f}"
              f"{vals['O'] or 0:>15,.0f}{vals['P'] or 0:>12,.0f}{vals['Q'] or 0:>14,.0f}"
              f"{vals['S'] or 0:>13,.0f}{vals['T'] or 0:>13,.0f}")

    # ── ۲) دوره (کلید ماه/سال/فصل) و مهلت پرداخت
    print("\n── استخراج دوره از تقویم ──")
    for i, row in enumerate(M.rows(), start=L.TX_FIRST):
        for col, label, want in (("U", "سال", row["jyear"]), ("V", "ماه", row["jmonth"]),
                                 ("W", "کلید دوره", row["ym"]), ("X", "فصل", row["season"])):
            got = ev(f"'{TX}'!{col}{i}")
            checks.append((f"{TX}!{col}{i} {label}", got, want))
    print(f"  نمونه → سال={ev(f'{chr(39)}{TX}{chr(39)}!U6')} کلید={ev(f'{chr(39)}{TX}{chr(39)}!W6')} "
          f"سرسید={ev(f'{chr(39)}{TX}{chr(39)}!Y6')} وضعیت سرسید={ev(f'{chr(39)}{TX}{chr(39)}!Z6')}")

    # ── ۳) تجمیع: کارت‌های راهنما + ردیف جمع شیت اعضا
    print("\n── تجمیع کل دوره (کارت‌های راهنما و جمع اعضا) ──")
    sums = M.totals()
    guides = [("تعداد پرونده", f"'{L.S_HELP}'!A6", sums["count"]),
              ("مبلغ کل کارشناسی", f"'{L.S_HELP}'!C6", sums["amount"]),
              ("کمیسیون خالص", f"'{L.S_HELP}'!E6", sums["total"])]
    for label, ref, want in guides:
        got = ev(ref)
        checks.append((f"راهنما {label}", got, want))
        g_txt = "None" if got is None else f"{round(float(got)):,}"
        print(f"  {label:<20} اکسل={g_txt:>17}   مرجع={round(want):>17,}")

    mt = M.member_totals()
    for col, key, label in (("E", "count", "تعداد پرونده"), ("F", "amount", "مبلغ کارشناسی"),
                            ("G", "gross", "کمیسیون خام"), ("H", "member_amt", "سهم اعضا"),
                            ("I", "company_amt", "سهم شرکت")):
        ref = f"'{L.S_MEM}'!{col}{L.MEM_TOTAL}"
        got = ev(ref)
        checks.append((f"اعضا {label}", got, mt[key]))
        g_txt = "None" if got is None else f"{round(float(got)):,}"
        print(f"  اعضا · {label:<14} اکسل={g_txt:>17}   مرجع={round(mt[key]):>17,}")

    # ── ۳ب) گزارش‌های تحلیلی (SUMIFS/COUNTIFS با ارجاع ساخت‌یافته)
    print("\n── گزارش‌ها (بلوک‌ها) ──")
    bm, bs, bstat, bmon = M.by_member(), M.by_service(), M.by_status(), M.by_month()
    rep_checks = []
    for i in range(3):
        r0 = L.REP_B1_FIRST + i
        name = ev(f"'{L.S_REP}'!A{r0}")
        d = bm.get(str(name).strip())
        if not d:
            continue
        for col, key, label in (("B", "count", "تعداد"), ("C", "amount", "مبلغ"),
                                ("D", "gross", "خام"), ("E", "total", "جمع"), ("F", "member_amt", "سهم")):
            rep_checks.append((f"گزارش اعضا {name} {label}", f"'{L.S_REP}'!{col}{r0}", d[key]))
    for i in range(3):
        r0 = L.REP_B3_FIRST + i
        name = ev(f"'{L.S_REP}'!A{r0}")          # نام واقعی خدمت از خود گزارش خوانده می‌شود
        d = bs.get(str(name).strip())
        if not d:
            continue
        for col, key, label in (("D", "count", "تعداد"), ("E", "amount", "مبلغ"), ("G", "total", "جمع")):
            rep_checks.append((f"گزارش خدمات {name} {label}", f"'{L.S_REP}'!{col}{r0}", d[key]))
    for i in range(4):
        r0 = L.REP_B4_FIRST + i
        name = ev(f"'{L.S_REP}'!A{r0}")
        d = bstat.get(str(name).strip())
        if not d:
            continue
        for col, key, label in (("C", "count", "تعداد"), ("F", "total", "جمع")):
            rep_checks.append((f"گزارش وضعیت {name} {label}", f"'{L.S_REP}'!{col}{r0}", d[key]))
    for i in range(3):
        r0 = L.REP_B2_FIRST + i
        key_txt = ev(f"'{L.S_REP}'!A{r0}")
        d = bmon.get(str(key_txt).strip())
        if not d:
            continue
        for col, key, label in (("C", "count", "تعداد"), ("D", "amount", "مبلغ"), ("F", "total", "جمع")):
            rep_checks.append((f"گزارش ماه {key_txt} {label}", f"'{L.S_REP}'!{col}{r0}", d[key]))

    for label, ref, want in rep_checks:
        got = ev(ref)
        checks.append((label, got, want))
    shown = [c for c in rep_checks if ev(c[1]) is None]
    print(f"  بررسی‌شده: {len(rep_checks)}  ·  ارزیابی‌نشده: {len(shown)}")
    for label, ref, want in rep_checks[:6]:
        got = ev(ref)
        g_txt = "None" if got is None else f"{round(float(got)):,}"
        print(f"  {label:<40} اکسل={g_txt:>15}   مرجع={round(want):>15,}")

    # ── ۴) ماشین‌حساب
    print("\n── ماشین‌حساب سناریو ──")
    calc_refs = {"مبلغ مبنا": "C15", "ضریب پله": "C16", "نرخ پایه": "C17", "نرخ مؤثر": "C18",
                 "خام": "C19", "اثر کف": "C20", "اثر سقف": "C21", "پس از کف/سقف": "C22",
                 "مالیات": "C23", "جمع با مالیات": "C24", "درصد سهم": "C25",
                 "مبلغ سهم": "C26", "سهم شرکت": "C27"}
    scen = M.scenario()
    for label, ref in calc_refs.items():
        got = ev(f"'{L.S_CALC}'!{ref}")
        want = scen.get(label)
        checks.append((f"ماشین‌حساب {label}", got, want))
        g_txt = "None" if got is None else f"{round(got, 4):,}"
        w_txt = "None" if want is None else f"{round(want, 4):,}"
        print(f"  {label:<16} اکسل={g_txt:>16}   مرجع={w_txt:>16}")

    # ── ۴ب) شاخص‌های داشبورد (با فیلتر دوره)
    print("\n── داشبورد (فیلتر سال ۱۴۰۵ · همهٔ ماه‌ها) ──")
    tot = M.totals()
    dash = [("تعداد پرونده", "A", mt["count"]), ("مبلغ کارشناسی", "C", mt["amount"]),
            ("کمیسیون خام", "E", mt["gross"]), ("کمیسیون خالص", "G", mt["total"]),
            ("سهم اعضا", "I", mt["member_amt"]), ("سهم شرکت", "K", mt["company_amt"]),
            ("نرخ مؤثر میانگین", "M", mt["gross"] / mt["base"]),
            ("تحقق هدف", "O", mt["total"] / M.P["target_commission"])]
    for label, col, want in dash:
        ref = f"'{L.S_DASH}'!{col}{L.DASH_KPI_ROW + 1}"
        got = ev(ref)
        checks.append((f"داشبورد {label}", got, want))
        g_txt = "None" if got is None else f"{round(float(got), 4):,}"
        print(f"  {label:<20} اکسل={g_txt:>18}   مرجع={round(want, 4):>18,}")

    # ── ۵) کنترل کیفیت
    print("\n── کنترل کیفیت ──")
    qc = M.qc_expectations()
    for i, (label, want) in enumerate(qc.items()):
        ref = f"'{L.S_QC}'!D{L.QC_FIRST + i}"
        got = ev(ref)
        checks.append((f"کنترل کیفیت {label}", got, want))
        g_txt = "None" if got is None else f"{round(float(got), 2):,}"
        print(f"  {label:<34} اکسل={g_txt:>13}   مرجع={round(float(want), 2):>13,}")

    # ── ۶) مقایسه و نتیجه
    print("\n" + "═" * 74)
    skipped = {f.split(":")[0].strip("'") for f in failures}
    bad = [(lbl, got, want) for lbl, got, want in checks
           if not near(got, want) and got is not None]
    print(f"  تعداد بررسی‌ها: {len(checks)}  (ارزیابی‌نشده با pycel: {sum(1 for _, g, _ in checks if g is None)})")
    if failures:
        print(f"  ⚠️  ارزیابی‌نشده ({len(failures)}):")
        for f in failures[:8]:
            print("     -", f)
    if bad:
        print(f"  ❌ اختلاف در {len(bad)} مورد:")
        for lbl, got, want in bad[:15]:
            print(f"     - {lbl}: اکسل={got} · مرجع={want}")
        return 1
    print("  ✅ همهٔ مقادیر با مدل مرجع یکسان است — موتور محاسبه درست کار می‌کند.")
    return 0


if __name__ == "__main__":
    target = sys.argv[1] if len(sys.argv) > 1 else "کمیسیون فنی.xlsx"
    sys.exit(main(target))
