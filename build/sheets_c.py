# -*- coding: utf-8 -*-
"""
sheets_c.py — شیت‌های محاسبه و تجمیع: ماشین‌حساب و گزارش‌ها.

«ماشین‌حساب» یک سناریوی یک‌موردی را از صفر تا سهم اعضا و شرکت گام‌به‌گام محاسبه
می‌کند و «گزارش‌ها» همان موتور را به‌صورت تجمیعی روی دفتر تراکنش‌ها اجرا می‌کند.
"""

from __future__ import annotations

from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule, IconSetRule

import content as C
import theme as T
from theme import L, aln, banner, brd_box, fld, fnt, merge, note_line, put, section, style_sheet, table_header, widths, zebra


MONTH_NAMES = ["فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
               "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند"]


# ═══════════════════════════════════════════════════════════════ شیت ماشین‌حساب

MACHINE_STEPS = [
    ("۱", "مبلغ مبنا", "مبلغ کارشناسی × ضریب خدمت (پایهٔ اعمال نرخ و پله‌ها)", T.NF_RIAL),
    ("۲", "ضریب پلهٔ اعمال‌شده", "ضریب مرتبط با پلهٔ مبلغ مبنا (از شیت تنظیمات)", T.NF_FACTOR),
    ("۳", "نرخ پایهٔ خدمت", "نرخ خدمت انتخاب‌شده از شیت فهرست‌ها", T.NF_PCT1),
    ("۴", "نرخ مؤثر", "نرخ پایه × ضریب پله (عددی که واقعاً ضرب می‌شود)", T.NF_PCT1),
    ("۵", "کمیسیون خام", "مبلغ مبنا × نرخ مؤثر", T.NF_RIAL),
    ("۶", "اثر کف (مثبت = افزایش)", "اگر کمیسیون خام از کف کمتر باشد، اختلاف افزوده می‌شود", T.NF_RIAL),
    ("۷", "اثر سقف (منفی = کاهش)", "اگر کمیسیون خام از سقف بیشتر باشد، اختلاف کاسته می‌شود", T.NF_RIAL),
    ("۸", "کمیسیون پس از کف و سقف", "مقدار تعدیل‌شده (برای وضعیت ابطال، صفر)", T.NF_RIAL),
    ("۹", "مالیات و عوارض", "کمیسیون پس از کف و سقف × نرخ مالیات", T.NF_RIAL),
    ("۱۰", "جمع با مالیات", "مبلغ ناخالص قابل وصول", T.NF_RIAL),
    ("۱۱", "درصد سهم عضو", "سهم بازنویسی‌شده، در غیر این صورت سهم پیش‌فرض عضو", T.NF_PCT1),
    ("۱۲", "مبلغ سهم عضو", "جمع با مالیات × درصد سهم", T.NF_RIAL),
    ("۱۳", "سهم شرکت", "جمع با مالیات − مبلغ سهم عضو", T.NF_RIAL),
]


def build_calculator(wb):
    ws = wb.create_sheet(L.S_CALC)
    style_sheet(ws, gridlines=False, zoom=110, tab_color=L.TAB_COLORS[L.S_CALC],
                print_area="A1:H34")
    widths(ws, {"A": 26, "B": 4, "C": 24, "D": 4, "E": 22, "F": 16, "G": 16, "H": 16})

    banner(ws, "ماشین‌حساب سناریو",
           "یک پروندهٔ فرضی را پیش از ثبت واقعی بسنجید؛ منطق این شیت مو‌به‌مو همان منطق شیت تراکنش‌ها است", 8)

    # ── بخش ۱: ورودی‌ها
    section(ws, 4, 1, 8, "۱) ورودی سناریو — فقط سلول‌های آبی")
    table_header(ws, 5, 1, ["قلم ورودی", "", "مقدار", "", "توضیح"], height=24)
    merge(ws, 5, 1, 5, 2)
    merge(ws, 5, 3, 5, 4)
    merge(ws, 5, 5, 5, 8)
    inputs = [
        ("عضو مسئول", "m_Member", C.MEMBERS[0][0], "=t_Members", T.NF_TEXT, "نام عضو از شیت «اعضا»"),
        ("نوع خدمت", "m_Service", C.SERVICES[0][0], "=lst_Service", T.NF_TEXT, "نرخ پایه از همین انتخاب خوانده می‌شود"),
        ("مبلغ کارشناسی (ریال)", "m_Amount", 1_250_000_000, None, T.NF_RIAL, "مبلغ قرارداد یا برآورد کارشناسی"),
        ("ضریب خدمت", "m_Factor", 1.10, None, T.NF_FACTOR, "برای عملیات چندبرابری؛ پیش‌فرض ۱"),
        ("سهم عضو (اختیاری)", "m_Share", None, None, T.NF_PCT1, "خالی بگذارید تا سهم پیش‌فرض عضو اعمال شود"),
        ("وضعیت", "m_Status", "تأییدشده", "=lst_Status", T.NF_TEXT, "وضعیت «ابطال» کمیسیون را صفر می‌کند"),
    ]
    for i, (label, name, value, dv_src, nf, desc) in enumerate(inputs):
        r = 6 + i
        put(ws, r, 1, label, f=fnt(10, True, T.TEXT), a=aln("right", indent=1))
        put(ws, r, 3, value, f=fnt(11, True, T.NAVY), fl=fld(T.INPUT_BG), a=aln("center"), nf=nf)
        ws.cell(row=r, column=3).border = brd_box(T.BLUE)
        merge(ws, r, 5, r, 8)
        put(ws, r, 5, desc, f=fnt(9, False, T.MUTED), a=aln("right", indent=1))
        ws.row_dimensions[r].height = 22
        if dv_src:
            dv = T.mk_dv(type="list", formula1=dv_src, allow_blank=True, showDropDown=False,
                         promptTitle="انتخاب از فهرست",
                         errorTitle="مقدار نامعتبر", error="مقدار را از فهرست انتخاب کنید.")
            ws.add_data_validation(dv)
            dv.add(ws.cell(row=r, column=3))

    # ── بخش ۲: زنجیرهٔ محاسبه
    section(ws, 13, 1, 8, "۲) زنجیرهٔ محاسبهٔ گام‌به‌گام (فرمول زنده)")
    table_header(ws, 14, 1, ["گام", "عنوان", "مقدار", "", "توضیح"], height=24)
    merge(ws, 14, 1, 14, 2)
    merge(ws, 14, 3, 14, 4)
    merge(ws, 14, 5, 14, 8)

    steps_formula = {
        1: '=IF(m_Amount="","",m_Amount*IF(m_Factor="",1,m_Factor))',
        2: '=IF($C$15="","",IF(p_TierMode="پلکانی",LOOKUP($C$15,t_Tier_From,t_Tier_Rate),1))',
        3: '=IFERROR(IF(m_Service="","",INDEX(t_Svc_Rate,MATCH(m_Service,lst_Service,0))),"")',
        4: '=IF(OR($C$15="",$C$17=""),"",$C$17*$C$16)',
        5: '=IF($C$15="","",$C$15*$C$18)',
        6: '=IF($C$19="","",MAX(0,p_Min-$C$19))',
        7: '=IF($C$19="","",MIN(0,p_Max-$C$19))',
        8: '=IF($C$19="","",IF(m_Status="ابطال",0,MAX(p_Min,MIN(p_Max,$C$19))))',
        9: '=IF($C$22="","",ROUND($C$22*p_Tax,0))',
        10: '=IF($C$22="","",$C$22+$C$23)',
        11: ('=IF($C$22="","",IF(m_Share<>"",m_Share,'
             'IFERROR(INDEX(t_Member_Share,MATCH(m_Member,t_Members,0)),p_DefaultShare)))'),
        12: '=IF($C$24="","",ROUND($C$24*$C$25,0))',
        13: '=IF($C$24="","",$C$24-$C$26)',
    }
    for i, (num, title, desc, nf) in enumerate(MACHINE_STEPS):
        r = 15 + i
        put(ws, r, 1, f"{num}. {title}", f=fnt(9.5, True, T.TEXT), a=aln("right", indent=1))
        put(ws, r, 3, steps_formula[i + 1], f=fnt(10.5, True, T.TEAL if i in (10, 11, 12) else T.NAVY),
            a=aln("center"), nf=nf)
        merge(ws, r, 5, r, 8)
        put(ws, r, 5, desc, f=fnt(8.5, False, T.MUTED), a=aln("right", indent=1))
        ws.row_dimensions[r].height = 21
    zebra(ws, 15, 27, 1, 3, border=False)
    for r in range(15, 28):
        ws.cell(row=r, column=3).border = brd_box(T.LINE)

    # ── بخش ۳: خروجی
    section(ws, 29, 1, 8, "۳) خروجی سناریو")
    T.kpi_card(ws, 30, 1, 2, "کمیسیون خالص (با مالیات)", "=IF($C$24=\"\",0,$C$24)", fmt=T.NF_RIAL)
    T.kpi_card(ws, 30, 3, 4, "سهم عضو", "=IF($C$26=\"\",0,$C$26)", fmt=T.NF_RIAL, fill=T.ACCENT)
    T.kpi_card(ws, 30, 5, 6, "سهم شرکت", "=IF($C$27=\"\",0,$C$27)", fmt=T.NF_RIAL, fill=T.OK_BG, fg=T.OK_FG)
    T.kpi_card(ws, 30, 7, 8, "نرخ مؤثر", "=IF($C$18=\"\",0,$C$18)", fmt=T.NF_PCT1, fill=T.SOFT)
    note_line(ws, 33, 1, 8,
              "این شیت هیچ رکوردی در دفتر تراکنش‌ها ایجاد نمی‌کند؛ فقط برای بررسی و مذاکره پیش از ثبت است.")

    ws.conditional_formatting.add(
        "C15:C27", DataBarRule(start_type="num", start_value=0, end_type="max",
                               color=T.TEAL, showValue=True, minLength=None, maxLength=None))
    ws.conditional_formatting.add(
        "C20:C21", CellIsRule(operator="greaterThan", formula=["0"], font=fnt(10, True, T.WARN_FG)))
    return ws


# ═══════════════════════════════════════════════════════════════ شیت گزارش‌ها

def build_reports(wb):
    ws = wb.create_sheet(L.S_REP)
    style_sheet(ws, gridlines=False, zoom=100, tab_color=L.TAB_COLORS[L.S_REP],
                freeze="A6", orientation="landscape", title_rows="1:2",
                print_area=f"A1:J{L.REP_B4_TOTAL + 2}")
    widths(ws, {"A": 24, "B": 12, "C": 20, "D": 20, "E": 20, "F": 20, "G": 20,
                "H": 12, "I": 10, "J": 15})

    banner(ws, "گزارش‌های تحلیلی",
           f"تجمیع زندهٔ دفتر تراکنش‌ها برای سال گزارش (تنظیمات ← سال گزارش)؛ "
           f"رکوردهای «ابطال» در همهٔ گزارش‌ها کنار گذاشته می‌شوند", 10)

    YEAR = "p_Year"
    NOT_VOID = 'tbl_TX[وضعیت],"<>ابطال"'

    # ═══ بلوک ۱ — عملکرد اعضا
    section(ws, L.REP_B1_SEC, 1, 10, "بلوک ۱ — عملکرد اعضا (مرتب بر مبلغ خالص)")
    table_header(ws, L.REP_B1_HDR, 1,
                 ["عضو", "تعداد پرونده", "مبلغ کارشناسی", "کمیسیون خام", "جمع با مالیات",
                  "مبلغ سهم عضو", "سهم شرکت", "درصد از کل", "رتبه", "وضعیت فعالیت"], height=30)
    for i in range(10):
        r = L.REP_B1_FIRST + i
        put(ws, r, 1, f'=IFERROR(IF(INDEX(t_Members,{i + 1})="","",INDEX(t_Members,{i + 1})),"—")',
            f=fnt(9.5, True, T.TEXT), a=aln("right", indent=1))
        put(ws, r, 2, f'=IF($A{r}="","",COUNTIFS(tbl_TX[عضو],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            f=fnt(10, True, T.TEAL), a=aln("center"), nf=T.NF_INT)
        put(ws, r, 3, f'=IF($A{r}="","",SUMIFS(tbl_TX[مبلغ کارشناسی],tbl_TX[عضو],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 4, f'=IF($A{r}="","",SUMIFS(tbl_TX[کمیسیون خام],tbl_TX[عضو],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 5, f'=IF($A{r}="","",SUMIFS(tbl_TX[جمع با مالیات],tbl_TX[عضو],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            f=fnt(9.5, True, T.NAVY), a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 6, f'=IF($A{r}="","",SUMIFS(tbl_TX[مبلغ سهم عضو],tbl_TX[عضو],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 7, f'=IF($A{r}="","",SUMIFS(tbl_TX[سهم شرکت],tbl_TX[عضو],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 8, f'=IF($E{r}="","",IFERROR($E{r}/$E${L.REP_B1_TOTAL},0))',
            a=aln("center"), nf=T.NF_PCT1)
        put(ws, r, 9, f'=IF($E{r}="","",IFERROR(RANK($E{r},$E${L.REP_B1_FIRST}:$E${L.REP_B1_LAST}),""))',
            f=fnt(10, True, T.GOLD), a=aln("center"), nf=T.NF_INT)
        put(ws, r, 10, f'=IF($A{r}="","",IF($B{r}=0,"بدون فعالیت",IF($B{r}<p_LowAct,"کم‌فعال","فعال")))',
            a=aln("center"))
        ws.row_dimensions[r].height = 21
    zebra(ws, L.REP_B1_FIRST, L.REP_B1_LAST, 1, 10)

    tr = L.REP_B1_TOTAL
    put(ws, tr, 1, "جمع کل / میانگین کل", f=fnt(10.5, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    for c in range(2, 11):
        put(ws, tr, c, None, fl=fld(T.NAVY))
    for c, col in ((2, "B"), (3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G")):
        put(ws, tr, c, f'=SUM({col}{L.REP_B1_FIRST}:{col}{L.REP_B1_LAST})',
            f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"),
            nf=T.NF_INT if c == 2 else T.NF_RIAL)
    put(ws, tr, 8, "۱۰۰٪", f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    put(ws, tr, 10, f'=TEXT(TODAY(),"yyyy/mm/dd")', f=fnt(9, False, T.ACCENT), fl=fld(T.NAVY), a=aln("center"))
    ws.row_dimensions[tr].height = 24

    # ═══ بلوک ۲ — روند ماهانه
    section(ws, L.REP_B2_SEC, 1, 10, "بلوک ۲ — روند ماهانهٔ سال شمسی")
    table_header(ws, L.REP_B2_HDR, 1,
                 ["ماه (کلید)", "نام ماه", "تعداد پرونده", "مبلغ کارشناسی", "کمیسیون خام",
                  "جمع با مالیات", "مبلغ سهم اعضا", "سهم شرکت", "درصد از کل", "رشد نسبت به ماه قبل"], height=30)
    for i in range(12):
        r = L.REP_B2_FIRST + i
        put(ws, r, 1, f'=p_Year&"/"&TEXT({i + 1},"00")', f=fnt(9, False, T.MUTED), a=aln("center"))
        put(ws, r, 2, MONTH_NAMES[i], f=fnt(9.5, True, T.TEXT), a=aln("center"))
        put(ws, r, 3, f'=COUNTIFS(tbl_TX[کلید دوره],$A{r},{NOT_VOID})', f=fnt(10, True, T.TEAL),
            a=aln("center"), nf=T.NF_INT)
        put(ws, r, 4, f'=SUMIFS(tbl_TX[مبلغ کارشناسی],tbl_TX[کلید دوره],$A{r},{NOT_VOID})',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 5, f'=SUMIFS(tbl_TX[کمیسیون خام],tbl_TX[کلید دوره],$A{r},{NOT_VOID})',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 6, f'=SUMIFS(tbl_TX[جمع با مالیات],tbl_TX[کلید دوره],$A{r},{NOT_VOID})',
            f=fnt(9.5, True, T.NAVY), a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 7, f'=SUMIFS(tbl_TX[مبلغ سهم عضو],tbl_TX[کلید دوره],$A{r},{NOT_VOID})',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 8, f'=SUMIFS(tbl_TX[سهم شرکت],tbl_TX[کلید دوره],$A{r},{NOT_VOID})',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 9, f'=IFERROR($F{r}/$F${L.REP_B2_TOTAL},0)', a=aln("center"), nf=T.NF_PCT1)
        put(ws, r, 10, ("—" if i == 0 else
                               f'=IFERROR(IF($F{r - 1}=0,"—",$F{r}/$F{r - 1}-1),"—")'),
            a=aln("center"), nf=T.NF_PCT1)
        ws.row_dimensions[r].height = 21
    zebra(ws, L.REP_B2_FIRST, L.REP_B2_LAST, 1, 10)

    tr = L.REP_B2_TOTAL
    put(ws, tr, 1, "جمع سال", f=fnt(10.5, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    put(ws, tr, 2, None, fl=fld(T.NAVY))
    for c, col in ((3, "C"), (4, "D"), (5, "E"), (6, "F"), (7, "G"), (8, "H")):
        put(ws, tr, c, f'=SUM({col}{L.REP_B2_FIRST}:{col}{L.REP_B2_LAST})',
            f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"),
            nf=T.NF_INT if c == 3 else T.NF_RIAL)
    put(ws, tr, 9, "۱۰۰٪", f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    for c in (10,):
        put(ws, tr, c, None, fl=fld(T.NAVY))
    ws.row_dimensions[tr].height = 24

    # ═══ بلوک ۳ — عملکرد خدمات
    section(ws, L.REP_B3_SEC, 1, 10, "بلوک ۳ — عملکرد انواع خدمات (نرخ اسمی در برابر نرخ واقعی)")
    table_header(ws, L.REP_B3_HDR, 1,
                 ["نوع خدمت", "دسته", "نرخ پایه", "تعداد پرونده", "مبلغ کارشناسی",
                  "کمیسیون خام", "جمع با مالیات", "نرخ مؤثر واقعی", "درصد از کل", "میانگین هر پرونده"], height=30)
    for i in range(8):
        r = L.REP_B3_FIRST + i
        put(ws, r, 1, f'=IFERROR(IF(INDEX(lst_Service,{i + 1})="","",INDEX(lst_Service,{i + 1})),"—")',
            f=fnt(9.5, True, T.TEXT), a=aln("right", indent=1))
        put(ws, r, 2, f'=IF($A{r}="","",IFERROR(INDEX(t_Svc_Cat,MATCH($A{r},lst_Service,0)),"—"))',
            f=fnt(9, False, T.MUTED), a=aln("center"))
        put(ws, r, 3, f'=IF($A{r}="","",IFERROR(INDEX(t_Svc_Rate,MATCH($A{r},lst_Service,0)),0))',
            a=aln("center"), nf=T.NF_PCT1)
        put(ws, r, 4, f'=IF($A{r}="","",COUNTIFS(tbl_TX[نوع خدمت],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            f=fnt(10, True, T.TEAL), a=aln("center"), nf=T.NF_INT)
        put(ws, r, 5, f'=IF($A{r}="","",SUMIFS(tbl_TX[مبلغ کارشناسی],tbl_TX[نوع خدمت],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 6, f'=IF($A{r}="","",SUMIFS(tbl_TX[کمیسیون خام],tbl_TX[نوع خدمت],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 7, f'=IF($A{r}="","",SUMIFS(tbl_TX[جمع با مالیات],tbl_TX[نوع خدمت],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year))',
            f=fnt(9.5, True, T.NAVY), a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 8, f'=IF($A{r}="","",IFERROR($F{r}/SUMIFS(tbl_TX[مبلغ مبنا],tbl_TX[نوع خدمت],$A{r},{NOT_VOID},tbl_TX[سال شمسی],p_Year),0))',
            f=fnt(9.5, True, T.TEAL), a=aln("center"), nf=T.NF_PCT1)
        put(ws, r, 9, f'=IFERROR($G{r}/$G${L.REP_B3_TOTAL},0)', a=aln("center"), nf=T.NF_PCT1)
        put(ws, r, 10, f'=IF($A{r}="","",IFERROR($G{r}/$D{r},0))', a=aln("center"), nf=T.NF_RIAL)
        ws.row_dimensions[r].height = 21
    zebra(ws, L.REP_B3_FIRST, L.REP_B3_LAST, 1, 10)

    tr = L.REP_B3_TOTAL
    put(ws, tr, 1, "جمع کل", f=fnt(10.5, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    for c in range(2, 11):
        put(ws, tr, c, None, fl=fld(T.NAVY))
    for c, col in ((4, "D"), (5, "E"), (6, "F"), (7, "G")):
        put(ws, tr, c, f'=SUM({col}{L.REP_B3_FIRST}:{col}{L.REP_B3_LAST})',
            f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"),
            nf=T.NF_INT if c == 4 else T.NF_RIAL)
    put(ws, tr, 8, f'=IFERROR($F{tr}/E{tr},0)', f=fnt(10, True, T.WHITE), fl=fld(T.NAVY),
        a=aln("center"), nf=T.NF_PCT1)
    put(ws, tr, 9, "۱۰۰٪", f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    put(ws, tr, 10, f'=IFERROR($G{tr}/$D{tr},0)', f=fnt(10, True, T.WHITE), fl=fld(T.NAVY),
        a=aln("center"), nf=T.NF_RIAL)
    ws.row_dimensions[tr].height = 24

    # ═══ بلوک ۴ — وضعیت پرونده‌ها
    section(ws, L.REP_B4_SEC, 1, 10, "بلوک ۴ — ترکیب وضعیت پرونده‌ها (جریان نقدی)")
    table_header(ws, L.REP_B4_HDR, 1,
                 ["وضعیت", "گروه", "تعداد پرونده", "سهم از تعداد", "مبلغ کارشناسی",
                  "جمع با مالیات", "سهم از ارزش", "میانگین هر پرونده", "سهم اعضا", "سهم شرکت"], height=30)
    for i in range(7):
        r = L.REP_B4_FIRST + i
        put(ws, r, 1, f'=IFERROR(IF(INDEX(lst_Status,{i + 1})="","",INDEX(lst_Status,{i + 1})),"—")',
            f=fnt(9.5, True, T.TEXT), a=aln("right", indent=1))
        put(ws, r, 2, f'=IF($A{r}="","",IFERROR(INDEX(t_Sta_Group,MATCH($A{r},lst_Status,0)),"—"))',
            f=fnt(9, False, T.MUTED), a=aln("center"))
        put(ws, r, 3, f'=IF($A{r}="","",COUNTIFS(tbl_TX[وضعیت],$A{r},tbl_TX[سال شمسی],p_Year))',
            f=fnt(10, True, T.TEAL), a=aln("center"), nf=T.NF_INT)
        put(ws, r, 4, f'=IFERROR($C{r}/$C${L.REP_B4_TOTAL},0)', a=aln("center"), nf=T.NF_PCT1)
        put(ws, r, 5, f'=IF($A{r}="","",SUMIFS(tbl_TX[مبلغ کارشناسی],tbl_TX[وضعیت],$A{r},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 6, f'=IF($A{r}="","",SUMIFS(tbl_TX[جمع با مالیات],tbl_TX[وضعیت],$A{r},tbl_TX[سال شمسی],p_Year))',
            f=fnt(9.5, True, T.NAVY), a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 7, f'=IFERROR($F{r}/$F${L.REP_B4_TOTAL},0)', a=aln("center"), nf=T.NF_PCT1)
        put(ws, r, 8, f'=IF($A{r}="","",IFERROR($F{r}/$C{r},0))', a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 9, f'=IF($A{r}="","",SUMIFS(tbl_TX[مبلغ سهم عضو],tbl_TX[وضعیت],$A{r},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        put(ws, r, 10, f'=IF($A{r}="","",SUMIFS(tbl_TX[سهم شرکت],tbl_TX[وضعیت],$A{r},tbl_TX[سال شمسی],p_Year))',
            a=aln("center"), nf=T.NF_RIAL)
        ws.row_dimensions[r].height = 21
    zebra(ws, L.REP_B4_FIRST, L.REP_B4_LAST, 1, 10)

    tr = L.REP_B4_TOTAL
    put(ws, tr, 1, "جمع کل", f=fnt(10.5, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    for c in range(2, 11):
        put(ws, tr, c, None, fl=fld(T.NAVY))
    for c, col in ((3, "C"), (5, "E"), (6, "F"), (9, "I"), (10, "J")):
        put(ws, tr, c, f'=SUM({col}{L.REP_B4_FIRST}:{col}{L.REP_B4_LAST})',
            f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"),
            nf=T.NF_INT if c == 3 else T.NF_RIAL)
    put(ws, tr, 4, "۱۰۰٪", f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    put(ws, tr, 7, "۱۰۰٪", f=fnt(10, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    put(ws, tr, 8, f'=IFERROR($F{tr}/$C{tr},0)', f=fnt(10, True, T.WHITE), fl=fld(T.NAVY),
        a=aln("center"), nf=T.NF_RIAL)
    ws.row_dimensions[tr].height = 24

    # ── فرمت شرطی گزارش‌ها
    ws.conditional_formatting.add(
        f"H{L.REP_B1_FIRST}:H{L.REP_B1_LAST}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color=T.TEAL, showValue=True))
    ws.conditional_formatting.add(
        f"I{L.REP_B1_FIRST}:I{L.REP_B1_LAST}",
        IconSetRule("3TrafficLights1", "percent", [0, 33, 67], showValue=True, reverse=True))
    ws.conditional_formatting.add(
        f"J{L.REP_B1_FIRST}:J{L.REP_B1_LAST}",
        FormulaRule(formula=[f'$J{L.REP_B1_FIRST}="کم‌فعال"'], fill=fld(T.WARN_BG),
                    font=fnt(9.5, True, T.WARN_FG), stopIfTrue=True))
    ws.conditional_formatting.add(
        f"J{L.REP_B1_FIRST}:J{L.REP_B1_LAST}",
        FormulaRule(formula=[f'$J{L.REP_B1_FIRST}="بدون فعالیت"'], fill=fld(T.SOFT),
                    font=fnt(9.5, False, T.MUTED), stopIfTrue=True))
    ws.conditional_formatting.add(
        f"J{L.REP_B1_FIRST}:J{L.REP_B1_LAST}",
        FormulaRule(formula=[f'$J{L.REP_B1_FIRST}="فعال"'], fill=fld(T.OK_BG),
                    font=fnt(9.5, True, T.OK_FG), stopIfTrue=True))
    ws.conditional_formatting.add(
        f"J{L.REP_B2_FIRST}:J{L.REP_B2_LAST}",
        ColorScaleRule(start_type="num", start_value=-1, start_color=T.BAD_BG,
                       mid_type="num", mid_value=0, mid_color="FFFFFF",
                       end_type="num", end_value=1, end_color=T.OK_FG))
    ws.conditional_formatting.add(
        f"I{L.REP_B3_FIRST}:I{L.REP_B3_LAST}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color=T.BLUE, showValue=True))
    ws.conditional_formatting.add(
        f"C{L.REP_B4_FIRST}:C{L.REP_B4_LAST}",
        ColorScaleRule(start_type="min", start_color="FFFFFF", end_type="max", end_color=T.ACCENT))
    ws.conditional_formatting.add(
        f"A{L.REP_B4_FIRST}:J{L.REP_B4_LAST}",
        FormulaRule(formula=[f'$A{L.REP_B4_FIRST}="ابطال"'],
                    font=fnt(9.5, False, T.MUTED, strike=True), stopIfTrue=False))
    ws.conditional_formatting.add(
        f"A{L.REP_B3_FIRST}:A{L.REP_B3_LAST}",
        FormulaRule(formula=[f'$D{L.REP_B3_FIRST}=0'], font=fnt(9.5, False, T.MUTED), stopIfTrue=False))
    return ws
