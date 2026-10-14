# -*- coding: utf-8 -*-
"""
sheets_a.py — شیت‌های پایه: راهنما، تنظیمات، فهرست‌ها، تقویم.

این شیت‌ها «هستهٔ مستقل از منطق فایل قبلی» هستند: تنظیمات و فهرست‌ها منبع همهٔ
پارامترها و گزینه‌های کشویی‌اند و تقویم، پل تبدیل تاریخ شمسی↔میلادی است.
"""

from __future__ import annotations

import datetime as dt

from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table

import content as C
import theme as T
from jalali import calendar_rows, gregorian_to_jalali
from theme import L, aln, band, banner, brd_bottom, brd_box, fld, fnt, heights, merge, nav_link, note_line, put, section, style_sheet, table_header, widths, zebra


# ═══════════════════════════════════════════════════════════════ شیت راهنما

def build_guide(wb):
    ws = wb.create_sheet(L.S_HELP)
    style_sheet(ws, gridlines=False, zoom=110, tab_color=L.TAB_COLORS[L.S_HELP],
                orientation="portrait", print_area="A1:H40")
    widths(ws, {"A": 6, "B": 26, "C": 34, "D": 18, "E": 14, "F": 14, "G": 14, "H": 14})

    banner(ws, "کارپوشهٔ جامع کمیسیون فنی", "راهنمای استفاده، نقشهٔ شیت‌ها و وضعیت زندهٔ کارپوشه — نسخهٔ ۲٫۰ بازطراحی‌شده", 8)

    # ── کارت‌های شاخص زنده
    r = 4
    section(ws, r, 1, 8, "وضعیت زندهٔ کارپوشه")
    r += 1
    cards = [
        ("تعداد پرونده‌ها", "=COUNTIF('تراکنش‌ها'!$F$6:$F$305,\">0\")", T.NF_QTY),
        ("مبلغ کل کارشناسی", "=SUM('تراکنش‌ها'!$F$6:$F$305)", T.NF_RIAL),
        ("کمیسیون خالص (با مالیات)", "=SUM('تراکنش‌ها'!$Q$6:$Q$305)", T.NF_RIAL),
        ("سال گزارش", "=p_Year", T.NF_INT),
    ]
    for i, (label, formula, nf) in enumerate(cards):
        c1 = 1 + i * 2
        T.kpi_card(ws, r, c1, c1 + 1, label, formula, fmt=nf,
                   fill=T.TEAL_SOFT if i % 2 == 0 else T.ACCENT, fg=T.NAVY)
    r += 3

    # ── نقشهٔ شیت‌ها
    section(ws, r, 1, 8, "نقشهٔ کارپوشه — هر شیت یک وظیفهٔ روشن")
    r += 1
    r = table_header(ws, r, 1, ["#", "شیت", "وظیفه", "نوع محتوا"], height=26)
    first_map = r
    for i, (name, duty, kind) in enumerate(C.GUIDE_SHEETS, start=1):
        put(ws, r, 1, i, f=fnt(9, True, T.MUTED), a=aln("center"))
        cell = put(ws, r, 2, name, f=fnt(10, True, T.BLUE), a=aln("right", indent=1))
        cell.hyperlink = f"#'{name}'!A1"
        put(ws, r, 3, duty, f=fnt(9), a=aln("right", indent=1))
        put(ws, r, 4, kind, f=fnt(8.5, False, T.MUTED), a=aln("center"))
        r += 1
    zebra(ws, first_map, r - 1, 1, 4)
    r += 1

    # ── گام‌های استفاده
    section(ws, r, 1, 8, "گام‌های استفاده (به همین ترتیب)")
    r += 1
    for step in C.USAGE_STEPS:
        merge(ws, r, 1, r, 8)
        put(ws, r, 1, step, f=fnt(9), a=aln("right", indent=1), b=brd_bottom())
        ws.row_dimensions[r].height = 19
        r += 1
    r += 1

    # ── راهنمای رنگ‌ها
    section(ws, r, 1, 8, "راهنمای رنگ‌ها و نمادها")
    r += 1
    r = table_header(ws, r, 1, ["#", "برچسب", "توضیح", "نمونه"], height=26)
    legend_style = {
        "input": (T.INPUT_BG, T.TEXT, T.INFO_FG),
        "calc": (T.CALC_BG, T.TEXT, T.MUTED),
        "zebra": (T.SOFT, T.TEXT, T.MUTED),
        "ok": (T.OK_BG, T.OK_FG, T.OK_FG),
        "warn": (T.WARN_BG, T.WARN_FG, T.WARN_FG),
        "bad": (T.BAD_BG, T.BAD_FG, T.BAD_FG),
        "gold": (T.GOLD, T.WHITE, T.GOLD),
    }
    first_leg = r
    for i, (label, desc, key) in enumerate(C.LEGEND, start=1):
        bg, fg, swatch = legend_style[key]
        put(ws, r, 1, i, f=fnt(9, True, T.MUTED), a=aln("center"))
        put(ws, r, 2, label, f=fnt(10, True, fg), fl=fld(bg), a=aln("center"))
        put(ws, r, 3, desc, f=fnt(9), a=aln("right", indent=1))
        merge(ws, r, 4, r, 8)
        put(ws, r, 4, "■", f=fnt(14, True, swatch), a=aln("center"))
        ws.row_dimensions[r].height = 22
        r += 1
    zebra(ws, first_leg, r - 1, 1, 3)
    r += 1

    # ── یادداشت پایانی
    note_line(ws, r, 1, 8,
              "نکتهٔ فنی: این کارپوشه همهٔ محاسبات را با فرمول زنده انجام می‌دهد؛ "
              "اگر اعداد پس از باز کردن فایل به‌روز نبودند، در اکسل کلید F9 و در برخی نسخه‌ها Ctrl+Alt+F9 را بزنید.")
    ws.row_dimensions[r].height = 30
    return ws


# ═══════════════════════════════════════════════════════════════ شیت تنظیمات

def build_settings(wb):
    ws = wb.create_sheet(L.S_SET)
    style_sheet(ws, gridlines=False, zoom=110, tab_color=L.TAB_COLORS[L.S_SET],
                freeze="A5", print_area="A1:C32")
    widths(ws, {"A": 32, "B": 22, "C": 52, "D": 3})

    banner(ws, "تنظیمات و سیاست‌های محاسباتی",
           "تنها منبع حقیقت: هر عددی که در فرمول‌ها استفاده می‌شود، از همین شیت خوانده می‌شود", 3)

    section(ws, L.SET_SEC, 1, 3, "پارامترهای عمومی")
    table_header(ws, L.SET_HDR, 1, ["پارامتر", "مقدار", "توضیح / اثر در محاسبات"], height=26)
    rows = [
        ("نام شرکت / کارپوشه", C.PARAMS["company"], "در سربرگ‌ها و گزارش‌های چاپی استفاده می‌شود", "@"),
        ("واحد پول", C.PARAMS["currency"], "فقط برچسب نمایشی اعداد است", "@"),
        ("سال گزارش", C.PARAMS["report_year"], "مبنای تجمیع در شیت «گزارش‌ها»", T.NF_INT),
        ("مالیات و عوارض (نرخ)", C.PARAMS["tax"], "ضرب در کمیسیون پس از کف/سقف", T.NF_PCT1),
        ("کف کمیسیون هر پرونده", C.PARAMS["min_commission"], "حداقل مبلغ کمیسیون (تعدیل رو به بالا)", T.NF_RIAL),
        ("سقف کمیسیون هر پرونده", C.PARAMS["max_commission"], "حداکثر مبلغ کمیسیون (تعدیل رو به پایین)", T.NF_RIAL),
        ("مهلت پرداخت (روز)", C.PARAMS["due_days"], "مبنای محاسبهٔ تاریخ سرسید در شیت تراکنش‌ها", T.NF_INT),
        ("حالت اعمال ضریب پله‌ای", C.PARAMS["tier_mode"], "«پلکانی» = استفاده از جدول پله‌ها · «ثابت» = ضریب ۱", "@"),
        ("درصد سهم پیش‌فرض عضو", C.PARAMS["default_share"], "اگر عضو سهم اختصاصی نداشته باشد", T.NF_PCT1),
        ("آستانهٔ کم‌فعالی (تعداد پرونده)", C.PARAMS["low_activity"], "مبنای برچسب «کم‌فعال» در شیت اعضا", T.NF_INT),
        ("هدف کمیسیون خالص سالانه", C.PARAMS["target_commission"], "مبنای درصد تحقق هدف در داشبورد", T.NF_RIAL),
    ]
    for i, (name, value, desc, nf) in enumerate(rows):
        rr = L.SET_COMPANY + i
        put(ws, rr, 1, name, f=fnt(10, True, T.TEXT), a=aln("right", indent=1))
        put(ws, rr, 2, value, f=fnt(11, True, T.NAVY), fl=fld(T.INPUT_BG), a=aln("center"), nf=nf)
        put(ws, rr, 3, desc, f=fnt(9, False, T.MUTED), a=aln("right", indent=1))
        ws.row_dimensions[rr].height = 21
    zebra(ws, L.SET_COMPANY, L.SET_TARGET, 1, 3, border=False)
    for rr in range(L.SET_COMPANY, L.SET_TARGET + 1):
        ws.cell(row=rr, column=2).border = brd_box(T.BLUE)

    section(ws, L.SET_TIER_SEC, 1, 3, "پله‌های ضریب نرخ (بر پایهٔ مبلغ مبنا)")
    table_header(ws, L.SET_TIER_HDR, 1, ["ردیف پله", "از مبلغ مبنا (ریال)", "ضریب نرخ"], height=26)
    for i, (frm, rate) in enumerate(C.TIERS):
        rr = L.SET_TIER_FIRST + i
        put(ws, rr, 1, f"پلهٔ {i + 1}", f=fnt(9.5, True, T.TEAL), a=aln("center"))
        put(ws, rr, 2, frm, f=fnt(10), fl=fld(T.INPUT_BG), a=aln("center"), nf=T.NF_RIAL)
        put(ws, rr, 3, rate, f=fnt(10, True, T.NAVY), fl=fld(T.INPUT_BG), a=aln("center"), nf=T.NF_FACTOR)
        ws.row_dimensions[rr].height = 21
    merge(ws, L.SET_TIER_LAST + 1, 1, L.SET_TIER_LAST + 1, 3)
    note_line(ws, L.SET_TIER_LAST + 1, 1, 3,
              "پله‌ها باید از کوچک به بزرگ مرتب باشند؛ ضریب نرخ روی «نرخ پایه» خدمت اعمال می‌شود (نرخ مؤثر).")

    section(ws, L.SET_NOTE_SEC, 1, 3, "یادداشت‌های سیاست‌گذاری")
    r = L.SET_NOTE_FIRST
    notes = [
        "۱. اگر «کف» را صفر بگذارید، هیچ حداقلی اعمال نمی‌شود؛ برای غیرفعال‌کردن سقف، عددی بزرگ (مثلاً ۹۹۹٫۹۹۹٫۹۹۹٫۹۹۹) وارد کنید.",
        "۲. کف و سقف روی «کمیسیون خام» اعمال می‌شوند و قبل از محاسبهٔ مالیات اتفاق می‌افتند.",
        "۳. ضریب خدمت (ستون G تراکنش‌ها) برای شرایطی است که یک پرونده چند برابر عملیات داشته باشد؛ مقدار پیش‌فرض ۱ است.",
        "۴. هر تغییری در این شیت، فوراً روی کل کارپوشه اثر می‌گذارد؛ آرشیوی از مقادیر قبلی نگه دارید.",
    ]
    for note in notes:
        merge(ws, r, 1, r, 3)
        put(ws, r, 1, note, f=fnt(9), a=aln("right", indent=1), b=brd_bottom())
        ws.row_dimensions[r].height = 19
        r += 1
    return ws


# ═══════════════════════════════════════════════════════════════ شیت فهرست‌ها

def build_lists(wb):
    ws = wb.create_sheet(L.S_LISTS)
    style_sheet(ws, gridlines=False, zoom=110, tab_color=L.TAB_COLORS[L.S_LISTS],
                print_area="A1:D66")
    widths(ws, {"A": 30, "B": 14, "C": 20, "D": 3})

    banner(ws, "فهرست‌های مرجع",
           "بانک گزینه‌های کشویی؛ افزودن یا ویرایش هر ردیف، بلافاصله در همهٔ شیت‌ها فعال می‌شود", 3)

    def block(sec_row, title, hdr_row, headers, data, table_name, col_spans):
        section(ws, sec_row, 1, 3, title)
        table_header(ws, hdr_row, 1, headers, height=24)
        for i, row_vals in enumerate(data):
            rr = hdr_row + 1 + i
            for j, v in enumerate(row_vals):
                nf = T.NF_PCT1 if isinstance(v, float) and 0 < v < 1 else None
                f = fnt(10, True, T.NAVY) if j == 0 else fnt(9.5)
                put(ws, rr, 1 + j, v, f=f, a=aln("right" if j == 0 else "center", indent=1 if j == 0 else 0), nf=nf)
            ws.row_dimensions[rr].height = 20
        first, last = hdr_row + 1, hdr_row + len(data)
        zebra(ws, first, last, 1, len(headers))
        tbl = Table(displayName=table_name, ref=f"A{hdr_row}:{get_column_letter(len(headers))}{last}")
        tbl.headerRowCount = 1
        tbl.totalsRowShown = False
        ws.add_table(tbl)
        return first, last

    svc_first, svc_last = block(L.SVC_SEC, "۱) انواع خدمات و نرخ پایهٔ کمیسیون",
                                L.SVC_HDR, ["نوع خدمت", "نرخ پایه", "دستهٔ خدمت"],
                                C.SERVICES, "T_Services", None)

    sta_rows = [(s, g) for s, g in zip(C.STATUSES,
               ["جاری", "جاری", "جاری", "جاری", "پایانی", "منفی", "منفی"])]
    sta_first, sta_last = block(L.STA_SEC, "۲) وضعیت پرونده‌ها",
                                L.STA_HDR, ["وضعیت", "گروه"],
                                sta_rows, "T_Status", None)

    cities = ["بوشهر", "عسلویه", "کرمان", "اصفهان", "اصفهان", "تهران", "تهران", "رشت"]
    cli_rows = list(zip(C.CONTRACTORS, cities))
    cli_first, cli_last = block(L.CLI_SEC, "۳) کارفرمایان / طرف‌های قرارداد",
                                L.CLI_HDR, ["کارفرما", "شهر"],
                                cli_rows, "T_Contractors", None)

    cat_first, cat_last = block(L.CAT_SEC, "۴) دسته‌بندی خدمات",
                                L.CAT_HDR, ["دستهٔ خدمت"],
                                [(c,) for c in C.CATEGORIES], "T_Categories", None)

    rol_first, rol_last = block(L.ROL_SEC, "۵) نقش‌های سازمانی",
                                L.ROL_HDR, ["نقش"],
                                [(r_,) for r_ in C.ROLES], "T_Roles", None)

    yn_first, yn_last = block(L.YN_SEC, "۶) مقادیر بله / خیر",
                              L.YN_HDR, ["مقدار"], [(v,) for v in C.YES_NO], "T_YesNo", None)

    tm_first, tm_last = block(L.MRG_SEC, "۷) حالت‌های اعمال ضریب نرخ",
                              L.MRG_HDR, ["حالت"], [(m,) for m in C.TIER_MODES], "T_TierMode", None)

    note_line(ws, L.MRG_SPARE + 1, 1, 3,
              "افزودن گزینهٔ جدید: کافی است در یکی از ردیف‌های یدکی زیر هر جدول تایپ کنید. "
              "محدوده‌های نام‌گذاری‌شده پویا هستند و گزینهٔ تازه بلافاصله در کشویی‌های کل کارپوشه ظاهر می‌شود.")
    note_line(ws, L.MRG_SPARE + 2, 1, 3,
              "این شیت هیچ فرمولی ندارد؛ اما مقصدِ نام‌گذاری محدوده‌ها است و ستون‌های کشویی همهٔ شیت‌ها از آن تغذیه می‌کنند.")
    return ws, {
        "svc": (svc_first, svc_last), "sta": (sta_first, sta_last),
        "cli": (cli_first, cli_last), "cat": (cat_first, cat_last),
        "rol": (rol_first, rol_last), "yn": (yn_first, yn_last),
        "tm": (tm_first, tm_last),
    }


# ═══════════════════════════════════════════════════════════════ شیت تقویم

def build_calendar(wb):
    ws = wb.create_sheet(L.S_CAL)
    style_sheet(ws, gridlines=False, zoom=100, tab_color=L.TAB_COLORS[L.S_CAL],
                freeze="A6", title_rows="5:5", print_area="A1:J200")
    widths(ws, {"A": 14, "B": 11, "C": 9, "D": 8, "E": 14, "F": 11, "G": 12, "H": 10, "I": 12, "J": 10})

    banner(ws, "تقویم مرجع شمسی",
           "پل تبدیل تاریخ: هر تاریخ شمسی در این جدول یک ردیف دارد؛ همهٔ محاسبات دوره‌ای کارپوشه به این جدول تکیه می‌کند", 10)

    headers = ["تاریخ میلادی", "سری", "سال شمسی", "ماه شمسی", "متن تاریخ شمسی",
               "کلید سال/ماه", "نام ماه", "فصل", "روز هفته", "روز کاری؟"]
    table_header(ws, L.CAL_HDR, 1, headers, height=28)

    rows = calendar_rows(2023, 2036)
    weekdays = ["دوشنبه", "سه‌شنبه", "چهارشنبه", "پنجشنبه", "جمعه", "شنبه", "یکشنبه"]
    r = L.CAL_FIRST
    for row in rows:
        put(ws, r, 1, row["gregorian"], a=aln("center"), nf=T.NF_DATE)
        put(ws, r, 2, row["serial"], a=aln("center"), nf=T.NF_INT)
        put(ws, r, 3, row["jy"], a=aln("center"), nf=T.NF_INT)
        put(ws, r, 4, row["jm"], a=aln("center"), nf=T.NF_INT)
        put(ws, r, 5, row["jtext"], f=fnt(9.5, True, T.NAVY), a=aln("center"))
        put(ws, r, 6, row["jyear_month"], a=aln("center"))
        put(ws, r, 7, row["month_name"], f=fnt(9.5), a=aln("center"))
        put(ws, r, 8, row["season"], a=aln("center"))
        put(ws, r, 9, weekdays[row["weekday"]], a=aln("center"))
        put(ws, r, 10, "بله" if row["is_workday"] else "خیر", a=aln("center"))
        r += 1
    last = r - 1

    assert last - L.CAL_FIRST + 1 == L.CAL_ROWS, "طول تقویم با مختصات ثابت هم‌خوان نیست"
    tbl = Table(displayName="T_Calendar", ref=f"A{L.CAL_HDR}:J{last}")
    tbl.headerRowCount = 1
    tbl.totalsRowShown = False
    ws.add_table(tbl)
    zebra(ws, L.CAL_FIRST, last, 1, 10)

    # آخرین روزهای هفتهٔ کاری را کمرنگ‌تر کن (نمای بصری سریع)
    ws.conditional_formatting.add(
        f"J{L.CAL_FIRST}:J{last}",
        FormulaRule(formula=[f'$J{L.CAL_FIRST}="خیر"'],
                    fill=fld(T.SOFT), font=fnt(9.5, False, T.MUTED), stopIfTrue=False))

    note_line(ws, last + 2, 1, 10,
              "این شیت پشتیبان است؛ برای پنهان‌کردن آن: راست‌کلیک روی برگه ← Hide (هیچ فرمولی خراب نمی‌شود).")
    return ws, last
