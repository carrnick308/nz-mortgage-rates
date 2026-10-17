"""
Builds the "Historical mortgage rates" workbook.

The script works out every figure in Python. With VALUES_ONLY = True (the
setting for the published file) the workbook contains just the results: no
formulas, so the method can't be unhidden or copied from the file. Set it to
False to build a version with live Excel formulas instead (each formula is
stored with the same pre-calculated result).
"""

from __future__ import annotations

import calendar
from datetime import date

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

# Display names. These exact strings are also what the "Cheapest term"
# column contains and what the Workings sheet counts, so keep them in sync.
TERM_KEYS = ["floating", "6m", "1y", "2y", "3y", "4y", "5y"]
TERM_NAMES = {
    "floating": "Floating", "6m": "6 months", "1y": "1 year", "2y": "2 years",
    "3y": "3 years", "4y": "4 years", "5y": "5 years",
}

# A start year is included in the "10 years or longer" analysis once the
# data runs at least this many months from January of that year.
MIN_PERIOD_MONTHS = 120
# Periods start in January of each year from this year.
FIRST_PERIOD_YEAR = 2005

# True: publish numbers only, no formulas. False: live formulas.
VALUES_ONLY = True

BRAND = "#0F4D9C"
LIGHT = "#E8EEF7"
CHEAPEST_FILL = "#D9EAD3"
FONT = "Arial"


# ---------------------------------------------------------------- helpers

def month_end(month: str) -> date:
    y, m = map(int, month.split("-"))
    return date(y, m, calendar.monthrange(y, m)[1])


def month_start(month: str) -> date:
    y, m = map(int, month.split("-"))
    return date(y, m, 1)


def long_month(month: str) -> str:
    return month_start(month).strftime("%B %Y")


def months_between(start: str, end: str) -> int:
    """Number of months from start to end inclusive."""
    sy, sm = map(int, start.split("-"))
    ey, em = map(int, end.split("-"))
    return (ey - sy) * 12 + (em - sm) + 1


def _mean(values):
    values = [v for v in values if v is not None]
    return sum(values) / len(values) if values else None


def _put(ws, row, col, formula, fmt, value):
    """Write a calculated cell: the value alone, or the formula plus its value."""
    if not VALUES_ONLY:
        ws.write_formula(row, col, formula, fmt, value)
    elif value is None or value == "":
        ws.write_blank(row, col, None, fmt)
    elif isinstance(value, str):
        ws.write_string(row, col, value, fmt)
    else:
        ws.write_number(row, col, value, fmt)


# ---------------------------------------------------------------- figures

def cheapest(row: dict, keys: list[str]) -> str | None:
    """Cheapest term for one month. Ties go to the shorter term (the first
    in display order), matching how the original spreadsheet broke ties."""
    present = [(row[k], i) for i, k in enumerate(keys) if row.get(k) is not None]
    if not present:
        return None
    low = min(v for v, _ in present)
    return keys[min(i for v, i in present if v == low)]


def period_starts(latest: str) -> list[str]:
    starts = []
    year = FIRST_PERIOD_YEAR
    while months_between(f"{year}-01", latest) >= MIN_PERIOD_MONTHS:
        starts.append(f"{year}-01")
        year += 1
    return starts


def compute(standard: dict, special: dict) -> dict:
    """All headline figures as fractions (0.0612 = 6.12%)."""
    s_months = sorted(standard)
    latest = s_months[-1]

    def col(rows, k, months):
        return [rows[m][k] / 100 if rows[m].get(k) is not None else None for m in months]

    def overall(rows, keys):
        months = sorted(rows)
        out = {}
        for k in keys:
            vals = [v for v in col(rows, k, months) if v is not None]
            out[k] = None if not vals else {
                "average": sum(vals) / len(vals), "high": max(vals), "low": min(vals),
                "range": max(vals) - min(vals),
            }
        return out

    std_keys = [k for k in TERM_KEYS if any(standard[m].get(k) is not None for m in s_months)]
    spc_keys = [k for k in TERM_KEYS if any(special[m].get(k) is not None for m in special)]

    periods = []
    for start in period_starts(latest):
        months = [m for m in s_months if m >= start]
        counts = {k: 0 for k in std_keys}
        for m in months:
            c = cheapest(standard[m], std_keys)
            if c:
                counts[c] += 1
        periods.append({
            "start": start,
            "months": len(months),
            "average": {k: _mean(col(standard, k, months)) for k in std_keys},
            "cheapest": counts,
            "total": sum(counts.values()),
        })

    long_avg = {k: _mean([p["average"][k] for p in periods]) for k in std_keys}
    totals = {k: sum(p["cheapest"][k] for p in periods) for k in std_keys}
    grand = sum(totals.values())

    return {
        "latest": latest,
        "std_start": s_months[0],
        "spc_start": min(special) if special else None,
        "spc_latest": max(special) if special else None,
        "std_keys": std_keys,
        "spc_keys": spc_keys,
        "standard_overall": overall(standard, std_keys),
        "special_overall": overall(special, spc_keys) if special else {},
        "periods": periods,
        "long_period_average": long_avg,
        "cheapest_totals": totals,
        "cheapest_share": {k: (totals[k] / grand if grand else 0) for k in std_keys},
        "cheapest_grand_total": grand,
    }


# ---------------------------------------------------------------- workbook

def build(path, standard: dict, special: dict, figs: dict, generated: str,
          logo_path=None, source_pages: dict | None = None) -> None:
    wb = xlsxwriter.Workbook(str(path))
    wb.set_properties({
        "title": "NZ historical mortgage interest rates",
        "author": "Your Money Blueprint",
        "comments": "Data: Reserve Bank of New Zealand series B20 and B21",
    })
    # Excel recalculates everything when the file is opened.
    wb.set_calc_mode("auto")

    f = _formats(wb)
    std_keys, spc_keys = figs["std_keys"], figs["spc_keys"]

    summary = wb.add_worksheet("Summary")
    std_ws, std_ref = _data_sheet(wb, f, "Standard rates", standard, std_keys)
    spc_ws, spc_ref = _data_sheet(wb, f, "Special rates", special, spc_keys) if special else (None, None)
    work_ws, work_ref = _workings_sheet(wb, f, figs, std_ref)
    _summary_sheet(summary, f, figs, std_ref, spc_ref, work_ref, generated, logo_path,
                   source_pages or {})
    summary.activate()
    wb.close()


def _formats(wb):
    base = {"font_name": FONT, "font_size": 10}

    def fmt(**kw):
        return wb.add_format({**base, **kw})

    return {
        "title": fmt(bold=True, font_size=16, font_color=BRAND),
        "subtitle": fmt(font_size=10, font_color="#404040"),
        "link": fmt(font_color="#0563C1", underline=1),
        "h2": fmt(bold=True, font_size=12, font_color=BRAND),
        "head": fmt(bold=True, font_color="white", bg_color=BRAND, align="center",
                    valign="vcenter", border=1, border_color=BRAND, text_wrap=True),
        "head_left": fmt(bold=True, font_color="white", bg_color=BRAND, align="left",
                         valign="vcenter", border=1, border_color=BRAND),
        "label": fmt(bold=True, bg_color=LIGHT, border=1, border_color="#C9D3E3"),
        "label_wrap": fmt(bold=True, bg_color=LIGHT, border=1, border_color="#C9D3E3",
                          text_wrap=True, valign="vcenter"),
        "pct": fmt(num_format="0.00%", align="center", border=1, border_color="#C9D3E3"),
        "pct_bold": fmt(num_format="0.00%", align="center", bold=True, border=1,
                        border_color="#C9D3E3"),
        "pct0": fmt(num_format="0.0%", align="center", border=1, border_color="#C9D3E3"),
        "int": fmt(num_format="#,##0", align="center", border=1, border_color="#C9D3E3"),
        "na": fmt(align="center", font_color="#808080", border=1, border_color="#C9D3E3"),
        "note": fmt(italic=True, font_size=9, font_color="#595959", text_wrap=True,
                    valign="top"),
        "month": fmt(num_format="mmm yyyy", align="left"),
        "month_box": fmt(num_format="mmm yyyy", align="left", bold=True, bg_color=LIGHT,
                         border=1, border_color="#C9D3E3"),
        "data_pct": fmt(num_format="0.00%", align="center"),
        "data_pct_cheap": fmt(num_format="0.00%", align="center", bold=True,
                              bg_color=CHEAPEST_FILL),
        "text_c": fmt(align="center"),
        "cheap": wb.add_format({"bg_color": CHEAPEST_FILL, "bold": True}),
        "small": fmt(font_size=9, font_color="#595959"),
    }


def _data_sheet(wb, f, name, rows, keys):
    """Month-by-month data with a 'Cheapest term' column. Returns the sheet and
    a dict of cell ranges other sheets refer to."""
    ws = wb.add_worksheet(name)
    months = sorted(rows)
    n = len(months)
    first, last = 2, n + 1  # Excel row numbers of the first/last data row
    term_col = {k: 1 + i for i, k in enumerate(keys)}  # 0-based columns
    cheap_col = 1 + len(keys)

    ws.write(0, 0, "Month", f["head_left"])
    for k, c in term_col.items():
        ws.write(0, c, TERM_NAMES[k], f["head"])
    ws.write(0, cheap_col, "Cheapest term", f["head"])

    first_letter = xl_col_to_name(1)
    last_letter = xl_col_to_name(len(keys))
    for i, m in enumerate(months):
        r = i + 1
        ws.write_datetime(r, 0, month_end(m), f["month"])
        present = [rows[m][k] for k in keys if rows[m].get(k) is not None]
        low = min(present) if present else None
        for k, c in term_col.items():
            v = rows[m].get(k)
            if v is None:
                ws.write_blank(r, c, None, f["data_pct"])
            else:
                # Values-only: shade the cheapest rate(s) directly.
                cheap = VALUES_ONLY and v == low
                ws.write_number(r, c, v / 100, f["data_pct_cheap"] if cheap else f["data_pct"])
        excel_row = r + 1
        rng = f"{first_letter}{excel_row}:{last_letter}{excel_row}"
        c_key = cheapest(rows[m], keys)
        _put(ws, 
            r, cheap_col,
            f"=IFERROR(INDEX(${first_letter}$1:${last_letter}$1,MATCH(MIN({rng}),{rng},0)),\"\")",
            f["text_c"], TERM_NAMES[c_key] if c_key else "",
        )

    # Formula version: highlight the cheapest rate(s) each month by rule.
    if not VALUES_ONLY:
        ws.conditional_format(first - 1, 1, last - 1, len(keys), {
            "type": "formula",
            "criteria": f"=AND({first_letter}{first}<>\"\",{first_letter}{first}="
                        f"MIN(${first_letter}{first}:${last_letter}{first}))",
            "format": f["cheap"],
        })

    ws.freeze_panes(1, 1)
    ws.repeat_rows(0)
    ws.fit_to_pages(1, 0)
    ws.set_paper(9)
    ws.set_column(0, 0, 11)
    ws.set_column(1, len(keys), 10)
    ws.set_column(cheap_col, cheap_col, 14)
    ws.autofilter(0, 0, n, cheap_col)

    q = f"'{name}'!"
    ref = {
        "sheet": name,
        "dates": f"{q}$A${first}:$A${last}",
        "cheapest": f"{q}${xl_col_to_name(cheap_col)}${first}:${xl_col_to_name(cheap_col)}${last}",
        "term": {k: f"{q}${xl_col_to_name(c)}${first}:${xl_col_to_name(c)}${last}"
                 for k, c in term_col.items()},
    }
    return ws, ref


def _workings_sheet(wb, f, figs, std_ref):
    ws = wb.add_worksheet("Workings")
    keys = figs["std_keys"]
    periods = figs["periods"]
    nk = len(keys)
    avg_c0 = 2                 # first average column (C)
    cnt_c0 = avg_c0 + nk + 1   # first count column, after a spacer
    total_c = cnt_c0 + nk

    ws.write(0, 0, "Workings – standard rates over periods of 10 years or longer", f["title"])
    ws.write(1, 0, "Each period runs from its start month to the latest month of data. "
                   "A new start year is added automatically once it has 10 years of data.",
             f["subtitle"])

    hr1, hr2 = 3, 4  # header rows (0-based)
    ws.merge_range(hr1, 0, hr2, 0, "Period starts", f["head"])
    ws.merge_range(hr1, 1, hr2, 1, "Months", f["head"])
    ws.merge_range(hr1, avg_c0, hr1, avg_c0 + nk - 1, "Average rate over the period", f["head"])
    ws.merge_range(hr1, cnt_c0, hr1, total_c, "Months each term had the cheapest rate", f["head"])
    for i, k in enumerate(keys):
        ws.write(hr2, avg_c0 + i, TERM_NAMES[k], f["head"])
        ws.write(hr2, cnt_c0 + i, TERM_NAMES[k], f["head"])
    ws.write(hr2, total_c, "Total", f["head"])

    first = hr2 + 1
    for i, p in enumerate(periods):
        r = first + i
        start_cell = xl_rowcol_to_cell(r, 0, col_abs=True)
        ws.write_datetime(r, 0, month_start(p["start"]), f["month_box"])
        _put(ws, r, 1, f'=COUNTIFS({std_ref["dates"]},">="&{start_cell})',
                         f["int"], p["months"])
        for j, k in enumerate(keys):
            _put(ws, 
                r, avg_c0 + j,
                f'=AVERAGEIFS({std_ref["term"][k]},{std_ref["dates"]},">="&{start_cell})',
                f["pct"], p["average"][k])
            head = xl_rowcol_to_cell(hr2, cnt_c0 + j, row_abs=True)
            _put(ws, 
                r, cnt_c0 + j,
                f'=COUNTIFS({std_ref["dates"]},">="&{start_cell},{std_ref["cheapest"]},{head})',
                f["int"], p["cheapest"][k])
        _put(ws, 
            r, total_c,
            f"=SUM({xl_rowcol_to_cell(r, cnt_c0)}:{xl_rowcol_to_cell(r, cnt_c0 + nk - 1)})",
            f["int"], p["total"])
    last = first + len(periods) - 1

    # Average of every period.
    avg_r = last + 2
    ws.merge_range(avg_r, 0, avg_r, 1, "Average of all periods", f["label_wrap"])
    ws.set_row(avg_r, 26)
    for j, k in enumerate(keys):
        c = avg_c0 + j
        _put(ws, 
            avg_r, c,
            f"=AVERAGE({xl_rowcol_to_cell(first, c)}:{xl_rowcol_to_cell(last, c)})",
            f["pct_bold"], figs["long_period_average"][k])

    tot_r, share_r = last + 2, last + 3
    ws.write(tot_r, cnt_c0 - 1, "Total", f["label"])
    ws.write(share_r, cnt_c0 - 1, "Share", f["label"])
    for j, k in enumerate(keys):
        c = cnt_c0 + j
        _put(ws, 
            tot_r, c, f"=SUM({xl_rowcol_to_cell(first, c)}:{xl_rowcol_to_cell(last, c)})",
            f["int"], figs["cheapest_totals"][k])
        _put(ws, 
            share_r, c,
            f"={xl_rowcol_to_cell(tot_r, c)}/{xl_rowcol_to_cell(tot_r, total_c, row_abs=True, col_abs=True)}",
            f["pct0"], figs["cheapest_share"][k])
    _put(ws, 
        tot_r, total_c,
        f"=SUM({xl_rowcol_to_cell(first, total_c)}:{xl_rowcol_to_cell(last, total_c)})",
        f["int"], figs["cheapest_grand_total"])

    ws.merge_range(share_r + 2, 0, share_r + 2, total_c,
                   "Where two or more terms share the lowest rate in a month, the shorter term is "
                   "counted. Recent months fall inside every period, so they are counted more "
                   "than once.", f["note"])
    ws.set_row(share_r + 2, 28)
    ws.set_column(0, 0, 13)
    ws.set_column(1, 1, 8)
    ws.set_column(avg_c0, avg_c0 + nk - 1, 9.5)
    ws.set_column(cnt_c0 - 1, cnt_c0 - 1, 9)
    ws.set_column(cnt_c0, total_c, 9)
    ws.freeze_panes(first, 1)
    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.set_paper(9)

    ref = {
        "avg": {k: f"Workings!{xl_rowcol_to_cell(avg_r, avg_c0 + j, True, True)}"
                for j, k in enumerate(keys)},
        "total": {k: f"Workings!{xl_rowcol_to_cell(tot_r, cnt_c0 + j, True, True)}"
                  for j, k in enumerate(keys)},
        "share": {k: f"Workings!{xl_rowcol_to_cell(share_r, cnt_c0 + j, True, True)}"
                  for j, k in enumerate(keys)},
    }
    return ws, ref


def _summary_sheet(ws, f, figs, std_ref, spc_ref, work_ref, generated, logo_path, pages):
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 34)
    ws.set_column(1, len(TERM_KEYS), 10.5)
    ws.set_column(len(TERM_KEYS) + 1, len(TERM_KEYS) + 1, 3)

    latest = long_month(figs["latest"])
    ws.write(0, 0, "NZ historical mortgage interest rates", f["title"])
    ws.write(1, 0, f"Monthly data from {long_month(figs['std_start'])} to {latest}. "
                   f"Updated automatically each month from Reserve Bank of New Zealand data.",
             f["subtitle"])
    r = 2
    if pages.get("B20"):
        ws.write_url(r, 0, pages["B20"], f["link"], string="Source: RBNZ standard mortgage rates (B20)")
        r += 1
    if pages.get("B21"):
        ws.write_url(r, 0, pages["B21"], f["link"], string="Source: RBNZ special mortgage rates (B21)")
        r += 1
    r += 1

    def header(row):
        ws.write(row, 0, "", f["head_left"])
        for i, k in enumerate(TERM_KEYS):
            ws.write(row, 1 + i, TERM_NAMES[k], f["head"])

    def stat_rows(row, overall, ref, keys):
        for label, fn, key in [("Average", "AVERAGE", "average"), ("Highest", "MAX", "high"),
                               ("Lowest", "MIN", "low")]:
            ws.write(row, 0, label, f["label"])
            for i, k in enumerate(TERM_KEYS):
                if k in keys and overall.get(k):
                    _put(ws, row, 1 + i, f"={fn}({ref['term'][k]})",
                                     f["pct_bold"] if key == "average" else f["pct"],
                                     overall[k][key])
                else:
                    ws.write(row, 1 + i, "N/A", f["na"])
            row += 1
        ws.write(row, 0, "Range (highest – lowest)", f["label"])
        for i, k in enumerate(TERM_KEYS):
            if k in keys and overall.get(k):
                _put(ws, row, 1 + i,
                                 f"={xl_rowcol_to_cell(row - 2, 1 + i)}-{xl_rowcol_to_cell(row - 1, 1 + i)}",
                                 f["pct"], overall[k]["range"])
            else:
                ws.write(row, 1 + i, "N/A", f["na"])
        return row + 1

    def note(row, text, height=26):
        ws.merge_range(row, 0, row, len(TERM_KEYS), text, f["note"])
        ws.set_row(row, height)

    # 1. Standard rates, all months
    ws.write(r, 0, f"Standard rates – all months since {long_month(figs['std_start'])}", f["h2"])
    r += 1
    header(r)
    r = stat_rows(r + 1, figs["standard_overall"], std_ref, figs["std_keys"])
    r += 1

    # 2. Standard rates over long periods
    starts = [p["start"] for p in figs["periods"]]
    ws.write(r, 0, "Standard rates – average over periods of 10 years or longer", f["h2"])
    r += 1
    header(r)
    r += 1
    ws.write(r, 0, "Average", f["label"])
    for i, k in enumerate(TERM_KEYS):
        if k in figs["std_keys"]:
            _put(ws, r, 1 + i, f"={work_ref['avg'][k]}", f["pct_bold"],
                             figs["long_period_average"][k])
        else:
            ws.write(r, 1 + i, "N/A", f["na"])
    r += 1
    note(r, f"The average of every period starting in January, from {starts[0][:4]} to "
            f"{starts[-1][:4]}, running to {latest}. A new start year is added once it has "
            f"10 years of data. See the Workings sheet.")
    r += 2

    # 3. Which term has been cheapest
    ws.write(r, 0, "How often each term has had the cheapest standard rate", f["h2"])
    r += 1
    header(r)
    r += 1
    for label, refkey, figkey, fmt_key in [("Months cheapest", "total", "cheapest_totals", "int"),
                                           ("Share of months", "share", "cheapest_share", "pct0")]:
        ws.write(r, 0, label, f["label"])
        for i, k in enumerate(TERM_KEYS):
            if k in figs["std_keys"]:
                _put(ws, r, 1 + i, f"={work_ref[refkey][k]}", f[fmt_key], figs[figkey][k])
            else:
                ws.write(r, 1 + i, "N/A", f["na"])
        r += 1
    note(r, f"Counted across every period of 10 years or longer, starting each January "
            f"from {starts[0][:4]} to {starts[-1][:4]}. Where terms tie, the shorter term "
            f"is counted.")
    r += 2

    # 4. Special rates
    if spc_ref:
        ws.write(r, 0, f"Special rates – all months since {long_month(figs['spc_start'])}", f["h2"])
        r += 1
        header(r)
        r = stat_rows(r + 1, figs["special_overall"], spc_ref, figs["spc_keys"])
        r += 1

    # Logo and notes
    if logo_path:
        ws.insert_image(r, 0, str(logo_path), {
            "x_scale": 0.42, "y_scale": 0.42, "x_offset": 2, "y_offset": 4,
            "url": "https://www.yourmoneyblueprint.co.nz/contact",
            "tip": "Personalised, unbiased financial advice",
            "object_position": 3,
        })
        ws.write_url(r + 5, 0, "https://www.yourmoneyblueprint.co.nz/contact", f["link"],
                     string="Click the logo for personalised, unbiased financial advice")
        r += 7
    note(r, "Rates are simple averages of the advertised first-mortgage rates offered by "
            "registered banks to new borrowers, at the end of each month. Standard rates "
            "exclude special offers. These are historical averages, not a quote from any bank.",
         height=38)
    ws.write(r + 1, 0, f"File generated {generated}.", f["small"])

    ws.set_landscape()
    ws.fit_to_pages(1, 1)
    ws.set_paper(9)  # A4
