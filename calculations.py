"""Dated IRR, currency NAV and quarterly FX hedge calculations."""

import calendar
import math
from collections import defaultdict
from datetime import date
from decimal import Decimal, ROUND_HALF_UP

from scipy.optimize import brentq


def group_cashflows(flows):
    amounts = defaultdict(list)
    for day, amount in flows:
        if not math.isfinite(amount):
            raise ValueError("cashflow amounts must be finite")
        amounts[day].append(amount)
    return [(day, math.fsum(values)) for day, values in sorted(amounts.items())]


def npv(flows, rate, when, include_today=True):
    if not math.isfinite(rate) or rate <= -1:
        raise ValueError("discount rate must be finite and greater than -1")
    return math.fsum(
        amount / (1 + rate) ** ((day - when).days / 365)
        for day, amount in flows
        if day > when or (include_today and day == when)
    )


def xirr(flows):
    flows = [(day, amount) for day, amount in group_cashflows(flows) if amount]
    if len(flows) < 2 or flows[0][1] >= 0 or any(amount < 0 for _, amount in flows[1:]):
        raise ValueError("IRR requires an initial outflow followed by receipts")

    def objective(rate):
        return npv(flows, rate, flows[0][0])

    if objective(0) == 0:
        return 0.0
    lower, upper = -0.999999, 1.0
    while objective(upper) > 0:
        upper *= 2
        if not math.isfinite(upper):
            raise ValueError("could not bracket IRR")
    if objective(lower) <= 0:
        raise ValueError("IRR is too close to -100%")
    return float(brentq(objective, lower, upper, xtol=1e-14))


def roll_dates(start, end):
    if end <= start:
        raise ValueError("exit must follow inception")
    dates, quarter = [start], 1
    while True:
        year, month = divmod(start.year * 12 + start.month - 1 + 3 * quarter, 12)
        last_day = calendar.monthrange(year, month + 1)[1]
        day = (
            last_day
            if start.day == calendar.monthrange(start.year, start.month)[1]
            else min(start.day, last_day)
        )
        following = date(year, month + 1, day)
        if following >= end:
            return [*dates, end]
        dates.append(following)
        quarter += 1


def analyse(records, fund="Fund I"):
    selected = [record for record in records if record.fund_name == fund]
    if not selected:
        raise ValueError(f"fund {fund!r} not found")
    bases = {record.base_currency for record in selected}
    if len(bases) != 1:
        raise ValueError("fund must have one base currency")
    base = bases.pop()
    fund_flows = group_cashflows([(r.cashflow_date, float(r.amount_base)) for r in selected])
    fund_rate = xirr(fund_flows)
    irr = [{"scope": "Fund", "currency": base, "annual_irr": fund_rate}]
    nav, trades, coverage = [], [], []
    for currency in sorted({record.local_currency for record in selected}):
        flows = group_cashflows(
            [
                (r.cashflow_date, float(r.amount_local))
                for r in selected
                if r.local_currency == currency
            ]
        )
        rate = xirr(flows)
        start, end = flows[0][0], flows[-1][0]
        dates = roll_dates(start, end)
        if abs(npv(flows, rate, start)) > 0.01:
            raise ValueError(f"{currency}: inception NAV is not zero")
        irr.append({"scope": "Currency", "currency": currency, "annual_irr": rate})
        for day in sorted(set(dates) | {day for day, _ in flows}):
            nav.append(
                {
                    "currency": currency,
                    "valuation_date": day,
                    "cashflow_local": dict(flows).get(day, 0),
                    "nav_pre_cashflow": 0.0 if day == start else npv(flows, rate, day),
                    "nav_post_cashflow": npv(flows, rate, day, include_today=False),
                }
            )
        if currency == base:
            continue
        for trade_date, delivery_date in zip(dates, dates[1:]):
            exposure = npv(flows, rate, trade_date, include_today=False)
            notional = Decimal(str(exposure)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            if notional <= 0:
                raise ValueError(f"{currency}: hedge exposure must be positive")
            trades.append(
                {
                    "fx_pair": f"{currency}/{base}",
                    "trade_date": trade_date,
                    "delivery_date": delivery_date,
                    "direction": "SELL",
                    "notional_currency": currency,
                    "notional_amount": notional,
                    "buy_currency": base,
                    "forward_rate": None,
                }
            )
            delivery_nav = npv(flows, rate, delivery_date)
            coverage.append(
                {
                    "currency": currency,
                    "trade_date": trade_date,
                    "delivery_date": delivery_date,
                    "coverage_at_trade": float(notional) / exposure,
                    "coverage_before_delivery": float(notional) / delivery_nav
                    if delivery_nav > 0
                    else None,
                    "unhedged_nav_before_delivery": delivery_nav - float(notional),
                }
            )
    return {
        "IRR": irr,
        "NAV Schedule": sorted(nav, key=lambda row: (row["valuation_date"], row["currency"])),
        "FX Forwards": sorted(
            trades, key=lambda row: (row["trade_date"], row["notional_currency"])
        ),
        "Hedge Coverage": sorted(coverage, key=lambda row: (row["trade_date"], row["currency"])),
    }
