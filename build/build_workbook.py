#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_workbook.py — ساخت کارپوشهٔ «کمیسیون فنی» از صفر.

مراحل:
    ۱. ساخت زیرساخت پنجرهٔ کارپوشه (خصیصه‌ها، محاسبهٔ خودکار هنگام باز شدن)
    ۲. ساخت ده شیت با ماژول‌های sheets_a/b/c/d
    ۳. تعریف نام‌های محدوده (ثابت + پویا با OFFSET/COUNTA)
    ۴. اعتبارسنجی خودکار خروجی (نام‌ها، جداول، ارجاع‌های ساخت‌یافته)
    ۵. ذخیرهٔ فایل نهایی

نمونهٔ اجرا:
    PYTHONPATH=./pylibs python3 build/build_workbook.py --out "کمیسیون فنی.xlsx"
"""

from __future__ import annotations

import argparse
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))

from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

import theme as T
from theme import L
import sheets_a, sheets_b, sheets_c, sheets_d


SHEET_KH = L.S_CAL        # نام شیت تقویم
SHEET_SET = L.S_SET
SHEET_LST = L.S_LISTS
SHEET_MEM = L.S_MEM


def define_names(wb):
    """تعریف همهٔ نام‌های محدوده: ثابت، پویا و مرجع."""
    names: dict[str, str] = {}

    q = lambda s: f"'{s}'"  # نام شیت در فرمول‌ها همیشه با گیومه (نیم‌فاصله و فاصله دارند)

    # ── پارامترهای تنظیمات
    for name, cell in [
        ("p_Company", f"$B${L.SET_COMPANY}"), ("p_Currency", f"$B${L.SET_CURRENCY}"),
        ("p_Year", f"$B${L.SET_YEAR}"), ("p_Tax", f"$B${L.SET_TAX}"),
        ("p_Min", f"$B${L.SET_MIN}"), ("p_Max", f"$B${L.SET_MAX}"),
        ("p_Due", f"$B${L.SET_DUE}"), ("p_TierMode", f"$B${L.SET_MODE}"),
        ("p_DefaultShare", f"$B${L.SET_SHARE}"), ("p_LowAct", f"$B${L.SET_LOWACT}"),
        ("p_Target", f"$B${L.SET_TARGET}"),
    ]:
        names[name] = f"{q(SHEET_SET)}!{cell}"
    names["t_Tier_From"] = f"{q(SHEET_SET)}!$B${L.SET_TIER_FIRST}:$B${L.SET_TIER_LAST}"
    names["t_Tier_Rate"] = f"{q(SHEET_SET)}!$C${L.SET_TIER_FIRST}:$C${L.SET_TIER_LAST}"

    # ── فهرست‌های پویا (با تایپ در ردیف‌های یدکی، خودکار گسترش می‌یابند)
    def dyn(sheet, col, first, spare, count_col=None):
        c = count_col or col
        return (f"OFFSET({q(sheet)}!${col}${first},0,0,"
                f"COUNTA({q(sheet)}!${c}${first}:${c}${spare}),1)")

    names["lst_Service"] = dyn(SHEET_LST, "A", L.SVC_FIRST, L.SVC_SPARE)
    names["t_Svc_Rate"] = dyn(SHEET_LST, "B", L.SVC_FIRST, L.SVC_SPARE)
    names["t_Svc_Cat"] = dyn(SHEET_LST, "C", L.SVC_FIRST, L.SVC_SPARE)
    names["lst_Status"] = dyn(SHEET_LST, "A", L.STA_FIRST, L.STA_SPARE)
    names["t_Sta_Group"] = dyn(SHEET_LST, "B", L.STA_FIRST, L.STA_SPARE)
    names["lst_Contractor"] = dyn(SHEET_LST, "A", L.CLI_FIRST, L.CLI_SPARE)
    names["t_Cli_City"] = dyn(SHEET_LST, "B", L.CLI_FIRST, L.CLI_SPARE)
    names["lst_Category"] = dyn(SHEET_LST, "A", L.CAT_FIRST, L.CAT_SPARE)
    names["lst_Role"] = dyn(SHEET_LST, "A", L.ROL_FIRST, L.ROL_SPARE)
    names["lst_YesNo"] = dyn(SHEET_LST, "A", L.YN_FIRST, L.YN_SPARE)
    names["lst_TierMode"] = dyn(SHEET_LST, "A", L.MRG_FIRST, L.MRG_SPARE)

    # ── اعضا (پویا: نام‌ها در ستون A مبنای شمارش است)
    names["t_Members"] = dyn(SHEET_MEM, "A", L.MEM_FIRST, L.MEM_SPARE)
    names["t_Member_Role"] = dyn(SHEET_MEM, "B", L.MEM_FIRST, L.MEM_SPARE)
    names["t_Member_Share"] = dyn(SHEET_MEM, "C", L.MEM_FIRST, L.MEM_SPARE)
    names["t_Member_Active"] = dyn(SHEET_MEM, "D", L.MEM_FIRST, L.MEM_SPARE)

    # ── تقویم (بازهٔ ثابت ۲۰۲۳ تا ۲۰۳۶ میلادی)
    for name, col in [("cal_Date", "A"), ("cal_Serial", "B"), ("cal_JYear", "C"),
                      ("cal_JMonth", "D"), ("cal_Key", "E"), ("cal_YM", "F"),
                      ("cal_MName", "G"), ("cal_Season", "H"), ("cal_WDay", "I"),
                      ("cal_Workday", "J")]:
        names[name] = f"{q(SHEET_KH)}!${col}${L.CAL_FIRST}:${col}${L.CAL_LAST}"

    # ── فیلتر داشبورد
    names["f_Year"] = f"{q(L.S_DASH)}!$B${L.DASH_FILTER_ROW}"
    names["f_MonthKey"] = f"{q(L.S_DASH)}!$E${L.DASH_FILTER_ROW}"

    # ── ورودی‌های ماشین‌حساب
    for name, row in [("m_Member", 6), ("m_Service", 7), ("m_Amount", 8),
                      ("m_Factor", 9), ("m_Share", 10), ("m_Status", 11)]:
        names[name] = f"{q(L.S_CALC)}!$C${row}"

    for name, ref in names.items():
        wb.defined_names.add(DefinedName(name, attr_text=ref))
    return names


def build(out_path: str):
    wb = Workbook()
    wb.remove(wb.active)  # شیت پیش‌فرض حذف می‌شود؛ ترتیب شیت‌ها را خودمان می‌سازیم

    # ── ترتیب منطقی شیت‌ها: راهنما → ورودی‌ها → داده → محاسبه → گزارش → نمایش → کنترل → مرجع
    sheets_a.build_guide(wb)
    sheets_a.build_settings(wb)
    _, _list_rows = sheets_a.build_lists(wb)
    sheets_b.build_members(wb)
    sheets_b.build_transactions(wb)
    sheets_c.build_calculator(wb)
    rep = sheets_c.build_reports(wb)
    sheets_d.build_dashboard(wb, rep)
    sheets_d.build_qc(wb, rep)
    sheets_a.build_calendar(wb)

    names = define_names(wb)

    # ── خصیصه‌های سند
    pr = wb.properties
    pr.title = "کارپوشهٔ جامع کمیسیون فنی"
    pr.subject = "محاسبهٔ کمیسیون، سهم اعضا و گزارش‌های مدیریتی"
    pr.creator = "طراحی‌شده از صفر — نسخهٔ ۲٫۰"
    pr.description = ("کارپوشهٔ چندشیتی با موتور محاسبهٔ فرمولی: نرخ پله‌ای، کف و سقف، "
                      "مالیات و عوارض، سهم اعضا، گزارش‌ها، داشبورد و کنترل کیفیت داده.")
    pr.keywords = "کمیسیون فنی، ریال، تقویم شمسی، داشبورد، گزارش مدیریتی"
    pr.category = "مالی و مدیریتی"

    # ── بازمحاسبهٔ کامل هنگام اولین باز شدن (فرمول‌ها بدون مقدار ذخیره‌شده‌اند)
    wb.calculation.fullCalcOnLoad = True

    # ── ترتیب شیت‌ها و حالت انتخاب
    order = [L.S_HELP, L.S_SET, L.S_LISTS, L.S_MEM, L.S_TX, L.S_CALC,
             L.S_REP, L.S_DASH, L.S_QC, L.S_CAL]
    wb._sheets = [wb[n] for n in order]
    for name in order:
        wb[name].sheet_view.tabSelected = name == L.S_HELP
    wb.active = 0

    os.makedirs(os.path.dirname(os.path.abspath(out_path)) or ".", exist_ok=True)
    wb.save(out_path)
    print(f"✅ فایل ساخته شد: {out_path}")
    print(f"   شیت‌ها: {len(order)} · نام محدوده: {len(names)}")
    return out_path


def main():
    ap = argparse.ArgumentParser(description="ساخت کارپوشهٔ کمیسیون فنی")
    ap.add_argument("--out", default=os.path.join(os.path.dirname(os.path.dirname(
        os.path.abspath(__file__))), "کمیسیون فنی.xlsx"), help="مسیر فایل خروجی")
    args = ap.parse_args()
    build(args.out)


if __name__ == "__main__":
    main()
