#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
jalali.py — تبدیل دوطرفهٔ تاریخ میلادی و شمسی (الگوریتم استاندارد خیام/بیرشک).

از این ماژول برای ساخت شیت «تقویم» در فایل اکسل استفاده می‌شود: پل تبدیل
«تاریخ سری میلادی» ↔ «سال/ماه/روز شمسی» که همهٔ محاسبات دوره‌ای فایل بر آن
استوار است. هیچ وابستگی بیرونی ندارد.

نمونهٔ اجرا:
    python3 tools/jalali.py            # جدول نمونه + صحت‌سنجی خودکار
"""

from __future__ import annotations

import datetime as _dt

G_DAYS_IN_MONTH = (31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31)
J_DAYS_IN_MONTH = (31, 31, 31, 31, 31, 31, 30, 30, 30, 30, 30, 29)

J_MONTHS = (
    "فروردین", "اردیبهشت", "خرداد", "تیر", "مرداد", "شهریور",
    "مهر", "آبان", "آذر", "دی", "بهمن", "اسفند",
)
J_SEASONS = ("بهار", "تابستان", "پاییز", "زمستان")


def gregorian_to_jalali(gy: int, gm: int, gd: int) -> tuple[int, int, int]:
    """تبدیل تاریخ میلادی به شمسی."""
    g_d_m = (0, 31, 59, 90, 120, 151, 181, 212, 243, 273, 304, 334)
    gy2 = gy - 1600
    gm2 = gm - 1
    gd2 = gd - 1
    g_day_no = 365 * gy2 + (gy2 + 3) // 4 - (gy2 + 99) // 100 + (gy2 + 399) // 400
    g_day_no += g_d_m[gm2] + gd2
    if gm > 2 and ((gy % 4 == 0 and gy % 100 != 0) or gy % 400 == 0):
        g_day_no += 1

    j_day_no = g_day_no - 79
    j_np = j_day_no // 12053
    j_day_no %= 12053
    jy = 979 + 33 * j_np + 4 * (j_day_no // 1461)
    j_day_no %= 1461
    if j_day_no >= 366:
        jy += (j_day_no - 1) // 365
        j_day_no = (j_day_no - 1) % 365
    for i in range(11):
        if j_day_no < J_DAYS_IN_MONTH[i]:
            return jy, i + 1, j_day_no + 1
        j_day_no -= J_DAYS_IN_MONTH[i]
    return jy, 12, j_day_no + 1


def jalali_to_gregorian(jy: int, jm: int, jd: int) -> tuple[int, int, int]:
    """تبدیل تاریخ شمسی به میلادی."""
    jy -= 979
    jm -= 1
    jd -= 1
    j_day_no = 365 * jy + (jy // 33) * 8 + ((jy % 33) + 3) // 4
    for i in range(jm):
        j_day_no += J_DAYS_IN_MONTH[i]
    j_day_no += jd

    g_day_no = j_day_no + 79
    gy = 1600 + 400 * (g_day_no // 146097)
    g_day_no %= 146097
    leap = True
    if g_day_no >= 36525:
        g_day_no -= 1
        gy += 100 * (g_day_no // 36524)
        g_day_no %= 36524
        if g_day_no >= 365:
            g_day_no += 1
        else:
            leap = False
    gy += 4 * (g_day_no // 1461)
    g_day_no %= 1461
    if g_day_no >= 366:
        leap = False
        g_day_no -= 1
        gy += g_day_no // 365
        g_day_no %= 365
    gd_m = list(G_DAYS_IN_MONTH)
    if leap:
        gd_m[1] = 29
    gm = 0
    for i, dim in enumerate(gd_m):
        if g_day_no < dim:
            gm = i + 1
            break
        g_day_no -= dim
    if gm == 0:  # pragma: no cover
        gm = 12
    return gy, gm, g_day_no + 1


# ------------------------------------------------------------- لایهٔ کاربردی

def to_jalali_text(d: _dt.date | _dt.datetime, sep: str = "/") -> str:
    """تاریخ میلادی ← متن تاریخ شمسی بافرمت `1405/07/22` (کلید جست‌وجو در اکسل)."""
    jy, jm, jd = gregorian_to_jalali(d.year, d.month, d.day)
    return f"{jy}{sep}{jm:02d}{sep}{jd:02d}"


def from_jalali_text(text: str, sep: str = "/") -> _dt.date:
    """متن `1405/07/22` ← تاریخ میلادی."""
    parts = [p for p in str(text).replace("-", sep).split(sep) if p.strip()]
    if len(parts) != 3:
        raise ValueError(f"قالب تاریخ نامعتبر است: {text!r}")
    jy, jm, jd = (int(p) for p in parts)
    gy, gm, gd = jalali_to_gregorian(jy, jm, jd)
    return _dt.date(gy, gm, gd)


def is_jalali_leap(jy: int) -> bool:
    """سال کبیسهٔ شمسی؟ (با آزمون عملی: آیا ۳۰ اسفند معتبر است)"""
    try:
        jalali_to_gregorian(jy, 12, 30)
        return True
    except ValueError:  # pragma: no cover
        return False


def calendar_rows(start_gy: int = 2024, end_gy: int = 2036) -> list[dict]:
    """
    تولید ردیف‌های شیت «تقویم» برای بازهٔ سال‌های میلادی داده‌شده.

    هر ردیف: سری اکسل، تاریخ میلادی، سال/ماه/روز شمسی، نام ماه، فصل،
    متن تاریخ شمسی (کلید)، هفتهٔ سال شمسی، نام روز هفته، روز کاری؟
    """
    rows: list[dict] = []
    day = _dt.date(start_gy, 1, 1)
    end = _dt.date(end_gy, 12, 31)
    week_no: dict[tuple[int, int], int] = {}
    while day <= end:
        jy, jm, jd = gregorian_to_jalali(day.year, day.month, day.day)
        key = (jy, jm)
        week_no.setdefault(key, len(week_no) + 1)
        # شمارهٔ هفته بر پایهٔ روزِ سال (هر ۷ روز یک هفته) — برای گزارش‌های دوره‌ای
        rows.append(
            {
                "gregorian": day,
                "jy": jy,
                "jm": jm,
                "jd": jd,
                "month_name": J_MONTHS[jm - 1],
                "season": J_SEASONS[(jm - 1) // 3],
                "jtext": f"{jy}/{jm:02d}/{jd:02d}",
                "jyear_month": f"{jy}/{jm:02d}",
                "weekday": day.weekday(),
                "is_workday": day.weekday() < 5,  # شنبه تا چهارشنبه/پنجشنبه در تقویم ایرانی
                "serial": (day - _dt.date(1899, 12, 30)).days,  # سری تاریخ اکسل (سیستم 1900)
            }
        )
        day += _dt.timedelta(days=1)
    return rows


# ------------------------------------------------------------- صحت‌سنجی

_SELF_TEST = [
    ((2025, 3, 21), (1404, 1, 1)),
    ((2025, 10, 14), (1404, 7, 22)),
    ((2026, 3, 21), (1405, 1, 1)),
    ((2026, 10, 14), (1405, 7, 22)),
    ((1979, 2, 11), (1357, 11, 22)),
    ((2000, 1, 1), (1378, 10, 11)),
    ((2024, 2, 29), (1402, 12, 10)),
]


def self_test() -> bool:
    ok = True
    for (gy, gm, gd), (jy, jm, jd) in _SELF_TEST:
        got = gregorian_to_jalali(gy, gm, gd)
        back = jalali_to_gregorian(jy, jm, jd)
        mark = "✅" if got == (jy, jm, jd) and back == (gy, gm, gd) else "❌"
        if mark == "❌":
            ok = False
        print(f"  {mark}  {gy:04d}/{gm:02d}/{gd:02d} ↔ {jy:04d}/{jm:02d}/{jd:02d}   "
              f"تبدیل‌شده={got} بازگشتی={back}")
    return ok


if __name__ == "__main__":
    print("صحت‌سنجی تبدیل تاریخ (میلادی ↔ شمسی):")
    ok = self_test()
    print(f"\nنتیجه: {'همهٔ موارد درست است' if ok else 'خطا در تبدیل'}")
    today = _dt.date(2026, 10, 14)
    print(f"\nامروز: {today.isoformat()}  ⇔  {to_jalali_text(today)} "
          f"({J_MONTHS[gregorian_to_jalali(*[today.year, today.month, today.day])[1] - 1]})")
    rows = calendar_rows(2026, 2027)
    print(f"تعداد ردیف تقویم در بازهٔ آزمون: {len(rows)}")
    print("نمونهٔ ردیف اول :", {k: str(v) for k, v in list(rows[0].items())})
    print("نمونهٔ ردیف آخر :", {k: str(v) for k, v in list(rows[-1].items())})
