# -*- coding: utf-8 -*-
"""
sheets_d.py — شیت‌های نمایش و کنترل: داشبورد و کنترل کیفیت.

داشبورد همهٔ اعداد را از دفتر تراکنش‌ها و شیت گزارش‌ها می‌خواند (هیچ محاسبهٔ تکراری
ندارد) و با فیلتر دوره کار می‌کند. شیت کنترل کیفیت، صحت داده و تطبیق جمع‌ها را با
آزمون‌های خودکار بررسی می‌کند.
"""

from __future__ import annotations

from openpyxl.chart import BarChart, DoughnutChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.marker import DataPoint
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.formatting.rule import CellIsRule, DataBarRule, FormulaRule
from openpyxl.styles import Border, Side

import content as C
import theme as T
from theme import L, aln, banner, brd_box, fld, fnt, merge, note_line, put, section, style_sheet, table_header, widths


# ═══════════════════════════════════════════════════════════════ داشبورد

def _filter_criteria():
    """
    عبارت مشترک فیلتر دوره (سال + ماه) برای همهٔ کارت‌های شاخص.

    نکتهٔ فنی: برای حالت «همهٔ ماه‌ها» از الگوی «?*» استفاده می‌شود نه «*»؛
    چون در اکسل الگوی «*» سلول‌های حاوی رشتهٔ خالی (نتایج فرمول در ردیف‌های
    استفاده‌نشده) را هم می‌شمارد، اما «?*» فقط سلول‌های دارای حداقل یک نویسه را.
    """
    return 'IF(f_MonthKey="همه","?*",f_Year&"/"&f_MonthKey)'


def build_dashboard(wb, rep):
    ws = wb.create_sheet(L.S_DASH)
    style_sheet(ws, gridlines=False, zoom=90, tab_color=L.TAB_COLORS[L.S_DASH],
                orientation="landscape", print_area=f"A1:P40")
    widths(ws, {c: 12 for c in range(1, 17)})

    banner(ws, "داشبورد مدیریتی",
           f"تصویر یک‌نگاهی عملکرد کمیسیون فنی — همهٔ اعداد از دفتر تراکنش‌ها و شیت گزارش‌ها خوانده می‌شوند", L.DASH_SPAN)

    K = _filter_criteria()
    r = L.DASH_FILTER_ROW

    # ── نوار فیلتر دوره
    put(ws, r, 1, "فیلتر دوره:", f=fnt(10, True, T.NAVY), a=aln("right", indent=1))
    merge(ws, r, 2, r, 3)
    put(ws, r, 2, C.PARAMS["report_year"], f=fnt(12, True, T.TEAL), fl=fld(T.INPUT_BG),
        a=aln("center"), nf=T.NF_INT, b=brd_box(T.BLUE))
    put(ws, r, 4, "ماه:", f=fnt(10, True, T.NAVY), a=aln("right", indent=1))
    merge(ws, r, 5, r, 6)
    put(ws, r, 5, "همه", f=fnt(12, True, T.TEAL), fl=fld(T.INPUT_BG),
        a=aln("center"), b=brd_box(T.BLUE))
    merge(ws, r, 7, r, 16)
    put(ws, r, 7, '=IF(f_MonthKey="همه","نمای سالانه — همهٔ ماه‌های سال "&f_Year,'
                  '"نمای ماهانه — ماه "&TEXT(VALUE(f_MonthKey),"00")&" از سال "&f_Year)',
        f=fnt(10, True, T.BLUE), fl=fld(T.ACCENT), a=aln("right", indent=1))
    ws.row_dimensions[r].height = 26

    dv = T.mk_dv(type="list", formula1='"همه,01,02,03,04,05,06,07,08,09,10,11,12"',
                 allow_blank=False, showDropDown=False, promptTitle="فیلتر ماه",
                 prompt="«همه» برای نمای سالانه یا شمارهٔ ماه (مثلاً 07) را انتخاب کنید.",
                 errorTitle="ماه نامعتبر", error="یک ماه یا «همه» را انتخاب کنید.")
    ws.add_data_validation(dv)
    dv.add(ws.cell(row=r, column=5))
    dv2 = T.mk_dv(type="list",
                  formula1='"1401,1402,1403,1404,1405,1406,1407,1408,1409,1410,1411,1412,1413,1414,1415"',
                  allow_blank=False, showDropDown=False, promptTitle="فیلتر سال",
                  errorTitle="سال نامعتبر", error="سال گزارش را از فهرست انتخاب کنید.")
    ws.add_data_validation(dv2)
    dv2.add(ws.cell(row=r, column=2))

    # ── کارت‌های شاخص کلیدی
    section(ws, L.DASH_KPI_ROW - 1, 1, L.DASH_SPAN, "شاخص‌های کلیدی دورهٔ انتخابی")
    kpis = [
        ("تعداد پرونده", f'=COUNTIFS(tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال")',
         T.NF_INT, T.TEAL_SOFT, T.NAVY, "پروندهٔ فعال (بدون ابطال)"),
        ("مبلغ کارشناسی", f'=SUMIFS(tbl_TX[مبلغ کارشناسی],tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال")',
         T.NF_RIAL, T.ACCENT, T.NAVY, "مجموع مبالغ پایه"),
        ("کمیسیون خام", f'=SUMIFS(tbl_TX[کمیسیون خام],tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال")',
         T.NF_RIAL, T.SOFT, T.NAVY, "پیش از کف، سقف و مالیات"),
        ("کمیسیون خالص", f'=SUMIFS(tbl_TX[جمع با مالیات],tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال")',
         T.NF_RIAL, T.TEAL_SOFT, T.TEAL, "جمع با مالیات — عدد کلیدی"),
        ("سهم اعضا", f'=SUMIFS(tbl_TX[مبلغ سهم عضو],tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال")',
         T.NF_RIAL, T.ACCENT, T.INFO_FG, "قابل پرداخت به پرسنل"),
        ("سهم شرکت", f'=SUMIFS(tbl_TX[سهم شرکت],tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال")',
         T.NF_RIAL, T.OK_BG, T.OK_FG, "درآمد ناخالص شرکت"),
        ("نرخ مؤثر میانگین",
         f'=IFERROR(SUMIFS(tbl_TX[کمیسیون خام],tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال")'
         f'/SUMIFS(tbl_TX[مبلغ مبنا],tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال"),0)',
         T.NF_PCT1, T.SOFT, T.NAVY, "کمیسیون خام ÷ مبلغ مبنا"),
        ("تحقق هدف",
         f'=IFERROR(SUMIFS(tbl_TX[جمع با مالیات],tbl_TX[کلید دوره],{K},tbl_TX[وضعیت],"<>ابطال")/p_Target,0)',
         T.NF_PCT0, T.WARN_BG, T.WARN_FG, "نسبت به هدف سالانه"),
    ]
    kr = L.DASH_KPI_ROW
    for i, (label, formula, nf, fill, fg, note) in enumerate(kpis):
        c1 = 1 + i * 2
        T.kpi_card(ws, kr, c1, c1 + 1, label, formula, fmt=nf, fill=fill, fg=fg, note=note)

    # رنگ‌بندی زندهٔ دو کارت آخر
    ws.conditional_formatting.add(
        f"A{kr + 1}:B{kr + 1}",
        CellIsRule(operator="lessThan", formula=["1"], fill=fld(T.BAD_BG), font=fnt(13, True, T.BAD_FG)))
    ws.conditional_formatting.add(
        f"O{kr + 1}:P{kr + 1}",
        FormulaRule(formula=[f'$O${kr + 1}<0.7'], fill=fld(T.BAD_BG), font=fnt(13, True, T.BAD_FG), stopIfTrue=True))
    ws.conditional_formatting.add(
        f"O{kr + 1}:P{kr + 1}",
        FormulaRule(formula=[f'AND($O${kr + 1}>=0.7,$O${kr + 1}<1)'], fill=fld(T.WARN_BG),
                    font=fnt(13, True, T.WARN_FG), stopIfTrue=True))

    # ── نمودارها
    section(ws, L.DASH_CHART_ROW - 1, 1, L.DASH_SPAN, "نمودارهای تحلیلی")
    cr = L.DASH_CHART_ROW

    def _style_axis(ch, num_fmt=T.NF_RIAL, title=None):
        ch.y_axis.numFmt = num_fmt
        ch.y_axis.majorGridlines.spPr = GraphicalProperties()
        ch.y_axis.majorGridlines.spPr.line = None
        if title:
            ch.y_axis.title = title
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.legend = None

    # ۱) ستونی — کمیسیون خالص به تفکیک عضو
    ch1 = BarChart()
    ch1.type = "col"
    ch1.style = 2
    ch1.title = "کمیسیون خالص به تفکیک عضو (ریال)"
    ch1.add_data(Reference(rep, min_col=5, min_row=L.REP_B1_FIRST, max_row=L.REP_B1_LAST), titles_from_data=False)
    ch1.set_categories(Reference(rep, min_col=1, min_row=L.REP_B1_FIRST, max_row=L.REP_B1_LAST))
    ch1.series[0].graphicalProperties.solidFill = T.TEAL
    ch1.series[0].graphicalProperties.line.solidFill = T.NAVY
    ch1.gapWidth = 60
    ch1.height, ch1.width = 8.2, 15.5
    ch1.dLbls = DataLabelList()
    ch1.dLbls.showVal = False
    _style_axis(ch1, T.NF_RIAL)
    ws.add_chart(ch1, f"A{cr}")

    # ۲) خطی — روند ماهانهٔ کمیسیون خالص
    ch2 = LineChart()
    ch2.style = 2
    ch2.title = "روند ماهانهٔ کمیسیون خالص (ریال)"
    ch2.add_data(Reference(rep, min_col=6, min_row=L.REP_B2_FIRST, max_row=L.REP_B2_LAST), titles_from_data=False)
    ch2.set_categories(Reference(rep, min_col=2, min_row=L.REP_B2_FIRST, max_row=L.REP_B2_LAST))
    ch2.series[0].graphicalProperties.line.solidFill = T.BLUE
    ch2.series[0].graphicalProperties.line.width = 28000
    ch2.series[0].smooth = False
    ch2.height, ch2.width = 8.2, 15.5
    _style_axis(ch2, T.NF_RIAL)
    ws.add_chart(ch2, f"I{cr}")

    # ۳) دوناتی — سهم انواع خدمت از کمیسیون خالص
    ch3 = DoughnutChart()
    ch3.title = "سهم انواع خدمت از کمیسیون خالص"
    ch3.add_data(Reference(rep, min_col=7, min_row=L.REP_B3_FIRST, max_row=L.REP_B3_LAST), titles_from_data=False)
    ch3.set_categories(Reference(rep, min_col=1, min_row=L.REP_B3_FIRST, max_row=L.REP_B3_LAST))
    ch3.height, ch3.width = 8.2, 15.5
    ch3.holeSize = 55
    palette = [T.NAVY, T.BLUE, T.TEAL, T.GOLD, T.OK_FG, T.WARN_FG, T.MUTED, T.BAD_FG]
    ch3.dataLabels = DataLabelList()
    ch3.dataLabels.showPercent = True
    ch3.dataLabels.showCatName = False
    pts = []
    for i, color in enumerate(palette[:8]):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=color)
        pts.append(dp)
    ch3.series[0].data_points = pts
    ws.add_chart(ch3, f"A{cr + 17}")

    # ۴) ترکیبی — تعداد پرونده (ستون) + کمیسیون خالص (خط، محور دوم)
    ch4 = BarChart()
    ch4.type = "col"
    ch4.style = 2
    ch4.title = "تعداد پرونده در برابر کمیسیون خالص (ماهانه)"
    ch4.add_data(Reference(rep, min_col=3, min_row=L.REP_B2_FIRST, max_row=L.REP_B2_LAST), titles_from_data=False)
    ch4.set_categories(Reference(rep, min_col=2, min_row=L.REP_B2_FIRST, max_row=L.REP_B2_LAST))
    ch4.series[0].graphicalProperties.solidFill = T.ACCENT
    ch4.series[0].graphicalProperties.line.solidFill = T.BLUE
    ch4.y_axis.title = "تعداد پرونده"
    ch4.y_axis.numFmt = T.NF_INT
    ch4.x_axis.delete = False
    ch4.y_axis.delete = False
    ch4.legend = None
    ch4.gapWidth = 55

    ch4l = LineChart()
    ch4l.add_data(Reference(rep, min_col=6, min_row=L.REP_B2_FIRST, max_row=L.REP_B2_LAST), titles_from_data=False)
    ch4l.series[0].graphicalProperties.line.solidFill = T.GOLD
    ch4l.series[0].graphicalProperties.line.width = 30000
    ch4l.y_axis.axId = 200
    ch4l.y_axis.title = "کمیسیون خالص (ریال)"
    ch4l.y_axis.numFmt = T.NF_RIAL
    ch4l.y_axis.majorGridlines = None
    ch4l.y_axis.crosses = "max"
    ch4l.legend = None
    ch4 += ch4l
    ch4.height, ch4.width = 8.2, 15.5
    ws.add_chart(ch4, f"I{cr + 17}")

    note_line(ws, cr + 34, 1, L.DASH_SPAN,
              "نمودارها با تغییر هر عدد به‌روز می‌شوند؛ اگر پس از باز کردن فایل نمودارها به‌روز نبودند، یک‌بار F9 بزنید. "
              "برای بزرگ‌نمایی هر نمودار، آن را انتخاب و از زبانهٔ Format تنظیم کنید.")
    return ws


# ═══════════════════════════════════════════════════════════════ کنترل کیفیت

QC_TESTS = [
    ("شمارهٔ پروندهٔ تکراری",
     "هر شمارهٔ پرونده باید یکتا باشد؛ تکرار یعنی ثبت دوبارهٔ یک پرونده.",
     '=SUMPRODUCT((COUNTIF(tbl_TX[شمارهٔ پرونده],tbl_TX[شمارهٔ پرونده])>1)'
     '*(tbl_TX[شمارهٔ پرونده]<>""))',
     0, 'row>1', "ستون A را مرتب کنید؛ ردیف‌های تکراری در تراکنش‌ها قرمز شده‌اند."),
    ("ردیف ناقص",
     "ردیفی که شمارهٔ پرونده دارد اما عضو، خدمت یا مبلغ آن خالی است.",
     '=COUNTIFS(tbl_TX[شمارهٔ پرونده],"<>")'
     '-COUNTIFS(tbl_TX[شمارهٔ پرونده],"<>",tbl_TX[عضو],"<>",tbl_TX[نوع خدمت],"<>",tbl_TX[مبلغ کارشناسی],"<>")',
     0, 'row>1', "ردیف‌های نارنجی‌رنگ در شیت تراکنش‌ها را کامل کنید."),
    ("تاریخ خارج از تقویم",
     "تاریخی که در جدول تقویم وجود ندارد (قالب اشتباه یا بیرون از بازهٔ تقویم).",
     '=COUNTIFS(tbl_TX[تاریخ],"<>",tbl_TX[مهلت پرداخت],"")',
     0, 'row>1', "تاریخ‌ها باید مثل 1405/07/22 و داخل بازهٔ شیت تقویم باشند."),
    ("خدمت بدون نرخ پایه",
     "خدمتی که در فهرست خدمات نیست، نرخ پایه‌اش خالی می‌ماند.",
     '=COUNTIFS(tbl_TX[نوع خدمت],"<>",tbl_TX[نرخ پایه],"")',
     0, 'row>1', "نام خدمت را از فهرست کشویی انتخاب کنید یا آن را در شیت فهرست‌ها اضافه کنید."),
    ("مبلغ نامعتبر",
     "مبلغ کارشناسی باید بزرگ‌تر از صفر باشد (ردیف‌های خالی محاسبه نمی‌شوند).",
     '=COUNTIFS(tbl_TX[مبلغ کارشناسی],"<=0",tbl_TX[مبلغ کارشناسی],"<>")',
     0, 'row>1', "سلول‌های مبلغ صفر یا منفی را اصلاح کنید."),
    ("ضریب خدمت خارج از بازه",
     "ضریب خدمت باید بین ۰٫۱ و ۵ باشد.",
     '=COUNTIFS(tbl_TX[ضریب خدمت],"<0.1",tbl_TX[ضریب خدمت],"<>")'
     '+COUNTIFS(tbl_TX[ضریب خدمت],">5",tbl_TX[ضریب خدمت],"<>")',
     0, 'row>1', "مقادیر پرت ستون G را بررسی کنید."),
    ("درصد سهم عضو خارج از بازه",
     "درصد سهم بازنویسی‌شده باید بین ۰ و ۱ باشد (۰٫۲۵ یعنی ۲۵ درصد).",
     '=COUNTIFS(tbl_TX[سهم عضو بازنویسی],"<0",tbl_TX[سهم عضو بازنویسی],"<>")'
     '+COUNTIFS(tbl_TX[سهم عضو بازنویسی],">1",tbl_TX[سهم عضو بازنویسی],"<>")',
     0, 'row>1', "مثلاً ۰٫۲۵ درست و ۲۵ غلط است."),
    ("تطبیق سهم‌ها با مبلغ کل",
     "جمع سهم اعضا + سهم شرکت باید برابر جمع با مالیات باشد (اختلاف باید صفر باشد).",
     '=ABS(SUM(tbl_TX[مبلغ سهم عضو])+SUM(tbl_TX[سهم شرکت])-SUM(tbl_TX[جمع با مالیات]))',
     0, 'row>1', "اگر اختلاف دیدید، نشانهٔ ویرایش دستی ستون‌های محاسباتی است."),
    ("تطبیق مالیات",
     "کمیسیون پس از کف و سقف + مالیات باید برابر جمع با مالیات باشد.",
     '=ABS(SUM(tbl_TX[کمیسیون پس از کف و سقف])+SUM(tbl_TX[مالیات و عوارض])-SUM(tbl_TX[جمع با مالیات]))',
     0, 'row>1', "ستون‌های O، P و Q تراکنش‌ها را بازبینی کنید."),
    ("رکورد ابطال‌شده در جمع‌ها",
     "پرونده‌های ابطال‌شده نباید مبلغی در جمع‌ها داشته باشند.",
     '=SUMIFS(tbl_TX[جمع با مالیات],tbl_TX[وضعیت],"ابطال")',
     0, 'row>1', "وضعیت «ابطال» باید کمیسیون را صفر کند؛ اگر صفر نیست، گزارش دهید."),
    ("پروندهٔ دارای سرسید سپری‌شده",
     "پرونده‌هایی که مهلت پرداخت‌شان گذشته و هنوز تسویه نشده‌اند (اطلاعی).",
     '=COUNTIF(tbl_TX[سرسید سپری شده],"بله")',
     3, 'row>0', "برای پیگیری: ستون Z تراکنش‌ها را فیلتر کنید روی «بله»."),
    ("پوشش تقویم برای سال گزارش",
     "سال گزارش باید در جدول تقویم موجود باشد تا محاسبات ماهانه کار کند.",
     '=COUNTIF(cal_JYear,p_Year)',
     365, 'row=1', "بازهٔ تقویم ۲۰۲۳ تا ۲۰۳۶ میلادی است؛ اگر سال گزارش بیرون این بازه است، تقویم را گسترش دهید."),
]


def build_qc(wb, rep):
    ws = wb.create_sheet(L.S_QC)
    style_sheet(ws, gridlines=False, zoom=110, tab_color=L.TAB_COLORS[L.S_QC],
                orientation="landscape", print_area="A1:G20")
    widths(ws, {"A": 5, "B": 26, "C": 46, "D": 14, "E": 12, "F": 10, "G": 44})

    banner(ws, "کنترل کیفیت داده",
           "آزمون‌های خودکار صحت، یکپارچگی و تطبیق جمع‌ها — پیش از گزارش‌گیری، این شیت را ببینید", 7)

    # ── نوار خلاصه
    merge(ws, 3, 1, 3, 7)
    put(ws, 3, 1,
        '="نتیجهٔ کنترل کیفیت:  "&COUNTIF(F'
        + str(L.QC_FIRST) + ':F' + str(L.QC_LAST) + ',"✅")&" از ' + str(L.QC_TESTS)
        + ' آزمون موفق   —   "&IF(COUNTIF(F' + str(L.QC_FIRST) + ':F' + str(L.QC_LAST)
        + ',"❌")=0,"وضعیت داده: سالم ✅","وضعیت داده: نیازمند اصلاح ❌")',
        f=fnt(12, True, T.WHITE), fl=fld(T.NAVY), a=aln("center"))
    for c in range(1, 8):
        put(ws, 3, c, fl=fld(T.NAVY))
    ws.row_dimensions[3].height = 30

    section(ws, L.QC_SEC, 1, 7, "آزمون‌های خودکار")
    table_header(ws, L.QC_HDR, 1,
                 ["#", "آزمون", "شرح و دلیل اهمیت", "مقدار", "آستانه", "نتیجه", "اقدام پیشنهادی"],
                 height=30)

    for i, (name, desc, formula, threshold, mode, action) in enumerate(QC_TESTS):
        r = L.QC_FIRST + i
        put(ws, r, 1, i + 1, f=fnt(9, True, T.MUTED), a=aln("center"))
        put(ws, r, 2, name, f=fnt(9.5, True, T.TEXT), a=aln("right", indent=1))
        put(ws, r, 3, desc, f=fnt(8.5, False, T.MUTED), a=aln("right", indent=1))
        put(ws, r, 4, formula, f=fnt(10, True, T.NAVY), a=aln("center"),
            nf=T.NF_INT if i not in (7, 8, 9) else T.NF_NUM)
        put(ws, r, 5, threshold, f=fnt(9.5, False, T.MUTED), a=aln("center"),
            nf=T.NF_INT)
        if mode == "row>0":       # بیشتر بهتر
            put(ws, r, 6, f'=IF($D{r}>=$E{r},"✅","❌")', f=fnt(12, True), a=aln("center"))
        elif mode == "row>1":
            put(ws, r, 6, f'=IF($D{r}<=$E{r},"✅","❌")', f=fnt(12, True), a=aln("center"))
        else:
            put(ws, r, 6, f'=IF($D{r}>=365,"✅","❌")', f=fnt(12, True), a=aln("center"))
        put(ws, r, 7, action, f=fnt(8.5, False, T.MUTED), a=aln("right", indent=1))
        ws.row_dimensions[r].height = 26

    from theme import zebra
    zebra(ws, L.QC_FIRST, L.QC_LAST, 1, 7, border=True)

    ws.conditional_formatting.add(
        f"F{L.QC_FIRST}:F{L.QC_LAST}",
        FormulaRule(formula=[f'$F{L.QC_FIRST}="✅"'], fill=fld(T.OK_BG),
                    font=fnt(12, True, T.OK_FG), stopIfTrue=True))
    ws.conditional_formatting.add(
        f"F{L.QC_FIRST}:F{L.QC_LAST}",
        FormulaRule(formula=[f'$F{L.QC_FIRST}="❌"'], fill=fld(T.BAD_BG),
                    font=fnt(12, True, T.BAD_FG), stopIfTrue=True))
    ws.conditional_formatting.add(
        f"D{L.QC_FIRST}:D{L.QC_LAST}",
        DataBarRule(start_type="num", start_value=0, end_type="max", color=T.WARN_FG, showValue=True))

    note_line(ws, L.QC_LAST + 2, 1, 7,
              "آزمون‌ها خواندنی و بدون ماکرو هستند: با هر تغییر در داده‌ها، نتیجه به‌روز می‌شود. "
              "آزمون «پروندهٔ دارای سرسید سپری‌شده» اطلاعی است و آستانهٔ آن ۳ مورد در نظر گرفته شده.")
    return ws
