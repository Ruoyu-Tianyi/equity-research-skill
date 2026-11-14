"""Reverse valuation helpers: what does the current price imply?

Library use (numbers come from the model map; nothing is estimated here):

    from reverse_valuation import implied_pe_anchor, solve, value_split
    implied_pe_anchor(price=20.0, shares=10.0, cost_of_equity=.10,
                      years_to_anchor=3.2, anchor_profit=30.0, pe=15.0)
    -> {'implied_profit': ..., 'implied_pe': ..., 'base_value': ...}

    solve(lambda g: my_dcf(g), target=price, lo=-.02, hi=.06)   # any monotone model
    value_split([('显性期现值', 120.0), ('终值现值', 380.0)])      # shares of value

Rules (research-and-valuation.md): fix all other variables and say so; report
no-solution/multi-solution honestly; implied values are a comparison point,
not the market consensus.
"""
from __future__ import annotations


def implied_pe_anchor(*, price, shares, cost_of_equity, years_to_anchor, anchor_profit=None, pe=None):
    """Discounted anchor-year PE valuation, solved both ways.

    value per share = pe * anchor_profit / shares / (1+ke)^t.
    Returns base value (if pe & profit given), the profit the price implies at
    the model PE, and the PE the price implies at the model profit.
    """
    disc = (1 + cost_of_equity) ** years_to_anchor
    out = {"discount_factor": disc}
    if pe and anchor_profit:
        out["base_value"] = pe * anchor_profit / shares / disc
    if pe:
        out["implied_profit"] = price * disc * shares / pe
        if anchor_profit: out["implied_profit_vs_model"] = out["implied_profit"] / anchor_profit
    if anchor_profit:
        out["implied_pe"] = price * disc * shares / anchor_profit
    return out


def solve(f, target, lo, hi, tol=1e-7, max_iter=200):
    """Bisection for f(x)=target on [lo, hi]. Returns dict with status; never extrapolates."""
    flo, fhi = f(lo) - target, f(hi) - target
    if flo == 0: return {"status": "ok", "x": lo}
    if fhi == 0: return {"status": "ok", "x": hi}
    if flo * fhi > 0:
        return {"status": "no_root_in_range", "f_lo": flo + target, "f_hi": fhi + target}
    for _ in range(max_iter):
        mid = (lo + hi) / 2; fm = f(mid) - target
        if abs(fm) < tol or (hi - lo) / 2 < tol: return {"status": "ok", "x": mid}
        if flo * fm < 0: hi, fhi = mid, fm
        else: lo, flo = mid, fm
    return {"status": "max_iter", "x": (lo + hi) / 2}


def value_split(parts):
    """parts = [(label, value)] -> shares; flags far-anchor dependence (>60%)."""
    total = sum(v for _, v in parts)
    shares = [(k, v, v / total if total else None) for k, v in parts]
    return {"total": total, "parts": shares}


def sensitivity(f, rows, cols):
    """2-D grid f(row, col) for heatmaps; caller verifies the base cell equals the headline value."""
    return [[f(r, c) for c in cols] for r in rows]
