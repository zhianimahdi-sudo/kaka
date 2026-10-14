# -*- coding: utf-8 -*-
"""
model_reference.py — مدل مرجع محاسبات (پیاده‌سازی مستقل پایتون).

این ماژول همان منطق فرمول‌های کارپوشه را یک‌بار دیگر و مستقل از اکسل پیاده می‌کند تا:
    ۱. خروجی موتور اکسل با یک مرجع مستقل مقایسه شود (آزمون رگرسیون)،
    ۲. مقدار انتظاری هر ردیف نمونه در سند «مقادیر مرجع» ثبت شود.

هیچ‌کدام از فرمول‌های فایل از این فایل خوانده نمی‌شود؛ اینجا فقط ریاضیات خالص است.
"""

from __future__ import annotations

import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "tools"))

import content as C
from jalali import from_jalali_text, gregorian_to_jalali

P = C.PARAMS
SVC_RATE = {name: rate for name, rate, _cat in C.SERVICES}
SVC_CAT = {name: cat for name, _rate, cat in C.SERVICES}
MEMBER_SHARE = {name: share for name, _role, share, _active in C.MEMBERS}
TODAY = dt.date(2026, 10, 14)          # تاریخ مرجع آزمون (سرسیدها نسبت به این سنجیده می‌شوند)
EXCEL_EPOCH = dt.date(1899, 12, 30)


def excel_serial(d: dt.date) -> int:
    return (d - EXCEL_EPOCH).days


def tier_factor(base: float) -> float:
    """معادل LOOKUP(base, t_Tier_From, t_Tier_Rate) — بزرگ‌ترین پلهٔ کوچک‌تر یا مساوی."""
    if P["tier_mode"] != "پلکانی":
        return 1.0
    chosen = C.TIERS[0][1]
    for frm, rate in C.TIERS:
        if base >= frm:
            chosen = rate
    return chosen


def clamp(value: float, low: float, high: float) -> float:
    """کف و سقف با MEDIAN(p_Min, value, p_Max) در اکسل معادل است."""
    return max(low, min(high, value))


def season_of(jm: int) -> str:
    return ["بهار", "تابستان", "پاییز", "زمستان"][(jm - 1) // 3]


def compute_row(jdate: str, member: str, service: str, amount: float, factor: float | None,
                share_override: float | None, status: str) -> dict:
    """محاسبهٔ کامل یک ردیف، هم‌ارز زنجیرهٔ ستون‌های K تا Z شیت تراکنش‌ها."""
    base = amount * (factor if factor else 1)
    base_rate = SVC_RATE.get(service, 0.0)
    tf = tier_factor(base)
    eff_rate = base_rate * tf
    gross = base * eff_rate
    after_clamp = 0.0 if status == "ابطال" else clamp(gross, P["min_commission"], P["max_commission"])
    tax = round(after_clamp * P["tax"])
    total = after_clamp + tax
    pct = share_override if share_override is not None else MEMBER_SHARE.get(member, P["default_share"])
    member_amt = round(total * pct)
    company_amt = total - member_amt

    g = from_jalali_text(jdate)
    jy, jm, jd = gregorian_to_jalali(g.year, g.month, g.day)
    due = g + dt.timedelta(days=P["due_days"])
    djy, djm, djd = gregorian_to_jalali(due.year, due.month, due.day)
    if status in ("پرداخت‌شده", "ابطال"):
        overdue = "—"
    else:
        overdue = "بله" if excel_serial(due) < excel_serial(TODAY) else "خیر"

    return {
        "case": None, "jdate": jdate, "member": member, "service": service,
        "amount": amount, "factor": factor, "share_override": share_override, "status": status,
        "base": base, "base_rate": base_rate, "tier_factor": tf, "eff_rate": eff_rate,
        "gross": gross, "after_clamp": after_clamp, "tax": tax, "total": total,
        "member_pct": pct, "member_amt": member_amt, "company_amt": company_amt,
        "jyear": jy, "jmonth": jm, "jday": jd, "ym": f"{jy}/{jm:02d}", "season": season_of(jm),
        "due": f"{djy}/{djm:02d}/{djd:02d}", "overdue": overdue,
        "gregorian": g, "due_gregorian": due,
    }


def rows() -> list[dict]:
    """ردیف‌های نمونهٔ کارپوشه با محاسبات کامل."""
    out = []
    for (jdate, case, member, service, _client, amount, factor, share, status) in C.SAMPLE_TX:
        r = compute_row(jdate, member, service, amount, factor, share, status)
        r["case"] = case
        out.append(r)
    return out


def totals() -> dict:
    """جمع‌های کل دوره (معادل SUM روی ستون‌های تراکنش‌ها)."""
    rs = rows()
    keys = ("amount", "base", "gross", "after_clamp", "tax", "total", "member_amt", "company_amt")
    out = {k: sum(r[k] for r in rs) for k in keys}
    out["count"] = len(rs)
    out["count_active"] = sum(1 for r in rs if r["status"] != "ابطال")
    out["overdue"] = sum(1 for r in rs if r["overdue"] == "بله")
    return out


def by_member() -> dict:
    out: dict[str, dict] = {}
    for r in rows():
        if r["status"] == "ابطال":
            continue
        d = out.setdefault(r["member"], {"count": 0, "amount": 0.0, "gross": 0.0,
                                        "total": 0.0, "member_amt": 0.0, "company_amt": 0.0})
        d["count"] += 1
        d["amount"] += r["amount"]
        d["gross"] += r["gross"]
        d["total"] += r["total"]
        d["member_amt"] += r["member_amt"]
        d["company_amt"] += r["company_amt"]
    return out


def by_year_month() -> dict:
    out: dict[str, float] = {}
    for r in rows():
        if r["status"] == "ابطال":
            continue
        out[r["ym"]] = out.get(r["ym"], 0.0) + r["total"]
    return out


def by_service() -> dict:
    """تجمیع بر پایهٔ نوع خدمت (بدون ابطال)."""
    out: dict[str, dict] = {}
    for r in rows():
        if r["status"] == "ابطال":
            continue
        d = out.setdefault(r["service"], {"count": 0, "amount": 0.0, "gross": 0.0, "total": 0.0})
        d["count"] += 1
        d["amount"] += r["amount"]
        d["gross"] += r["gross"]
        d["total"] += r["total"]
    return out


def by_status() -> dict:
    """تجمیع بر پایهٔ وضعیت (شامل ابطال، چون گزارش وضعیت‌ها همه را می‌شمارد)."""
    out: dict[str, dict] = {}
    for r in rows():
        d = out.setdefault(r["status"], {"count": 0, "amount": 0.0, "total": 0.0})
        d["count"] += 1
        d["amount"] += r["amount"]
        d["total"] += r["total"]
    return out


def by_month() -> dict:
    """تجمیع بر پایهٔ کلید سال/ماه (بدون ابطال)."""
    out: dict[str, dict] = {}
    for r in rows():
        if r["status"] == "ابطال":
            continue
        d = out.setdefault(r["ym"], {"count": 0, "amount": 0.0, "gross": 0.0, "total": 0.0})
        d["count"] += 1
        d["amount"] += r["amount"]
        d["gross"] += r["gross"]
        d["total"] += r["total"]
    return out


def member_totals() -> dict:
    """جمع‌های ردیف «جمع کل» شیت اعضا (بدون رکوردهای ابطال)."""
    rs = [r for r in rows() if r["status"] != "ابطال"]
    return {
        "count": len(rs),
        "amount": sum(r["amount"] for r in rs),
        "base": sum(r["base"] for r in rs),
        "gross": sum(r["gross"] for r in rs),
        "total": sum(r["total"] for r in rs),
        "member_amt": sum(r["member_amt"] for r in rs),
        "company_amt": sum(r["company_amt"] for r in rs),
    }


def scenario() -> dict:
    """خروجی ماشین‌حساب با ورودی‌های پیش‌فرض شیت."""
    member = C.MEMBERS[0][0]
    service = C.SERVICES[0][0]
    amount, factor, share, status = 1_250_000_000, 1.10, None, "تأییدشده"
    r = compute_row("1405/07/22", member, service, amount, factor, share, status)
    return {
        "مبلغ مبنا": r["base"], "ضریب پله": r["tier_factor"], "نرخ پایه": r["base_rate"],
        "نرخ مؤثر": r["eff_rate"], "خام": r["gross"],
        "اثر کف": max(0.0, P["min_commission"] - r["gross"]),
        "اثر سقف": min(0.0, P["max_commission"] - r["gross"]),
        "پس از کف/سقف": r["after_clamp"], "مالیات": r["tax"], "جمع با مالیات": r["total"],
        "درصد سهم": r["member_pct"], "مبلغ سهم": r["member_amt"], "سهم شرکت": r["company_amt"],
    }


def qc_expectations() -> dict:
    """مقدار انتظاری هر آزمون شیت کنترل کیفیت (به ترتیب ردیف‌ها)."""
    rs = rows()
    t = totals()
    cal_year_days = 0
    day = dt.date(2023, 1, 1)
    while day <= dt.date(2036, 12, 31):
        if gregorian_to_jalali(day.year, day.month, day.day)[0] == P["report_year"]:
            cal_year_days += 1
        day += dt.timedelta(days=1)
    cases = [r["case"] for r in rs]
    return {
        "شمارهٔ پروندهٔ تکراری": 0,
        "ردیف ناقص": 0,
        "تاریخ خارج از تقویم": 0,
        "خدمت بدون نرخ پایه": 0,
        "مبلغ نامعتبر": 0,
        "ضریب خدمت خارج از بازه": 0,
        "درصد سهم عضو خارج از بازه": 0,
        "تطبیق سهم‌ها با مبلغ کل": abs(t["member_amt"] + t["company_amt"] - t["total"]),
        "تطبیق مالیات": abs(t["after_clamp"] + t["tax"] - t["total"]),
        "رکورد ابطال‌شده در جمع‌ها": sum(r["total"] for r in rs if r["status"] == "ابطال"),
        "پروندهٔ دارای سرسید سپری‌شده": t["overdue"],
        "پوشش تقویم برای سال گزارش": cal_year_days,
    }


if __name__ == "__main__":
    t = totals()
    print("مدل مرجع — جمع‌های کل دوره")
    for k, v in t.items():
        print(f"  {k:<16} {v:>20,}" if isinstance(v, (int, float)) else f"  {k}: {v}")
    print("\nماشین‌حساب (سناریوی پیش‌فرض):")
    for k, v in scenario().items():
        print(f"  {k:<16} {v:>18,.4f}")
    print("\nکنترل کیفیت (مقادیر انتظاری):")
    for k, v in qc_expectations().items():
        print(f"  {k:<34} {v:>10,}")
