"""Excel (xlsx) export: everything in one file, with native Excel charts (graphs)."""
import io
from datetime import datetime

import xlsxwriter

EXCEL_MAX_ROWS = 1_048_575
GREEN = "#71bd44"
DARK_GREEN = "#3f7a22"


def _bar_chart(wb, kind, title, sheet, cat, val, y_title=None, x_title=None, reverse=False, num_fmt="#,##0.00"):
    ch = wb.add_chart({"type": kind})
    ch.add_series({
        "categories": cat, "values": val,
        "fill": {"color": GREEN}, "border": {"color": DARK_GREEN},
        "gap": 60,
        "data_labels": {"value": True, "num_format": num_fmt, "font": {"size": 9}},
    })
    ch.set_title({"name": title, "name_font": {"size": 13, "color": DARK_GREEN}})
    ch.set_legend({"none": True})
    cat_ax = {"num_font": {"size": 9}}
    val_ax = {"num_font": {"size": 9}, "major_gridlines": {"visible": True, "line": {"color": "#e3ebde"}}}
    if reverse:
        cat_ax["reverse"] = True        # rank 1 appears at the top
    if kind == "bar":                   # horizontal bar: y = category, x = value
        if x_title:
            val_ax["name"] = x_title
        if y_title:
            cat_ax["name"] = y_title
        ch.set_y_axis(cat_ax)
        ch.set_x_axis(val_ax)
    else:                               # column: x = category, y = value
        if x_title:
            cat_ax["name"] = x_title
        if y_title:
            val_ax["name"] = y_title
        ch.set_x_axis(cat_ax)
        ch.set_y_axis(val_ax)
    ch.set_size({"width": 760, "height": 420})
    return ch


def build_workbook(payload, package, start_date, end_date, max_rows=EXCEL_MAX_ROWS, filter_label=None):
    rows = payload["rows"]            # [package, 'YYYY-MM-DD', msisdn, revenue_of_that_day]
    # per-subscriber totals over the whole range (used for Summary, Top 10 and the breakdowns)
    subs = {}
    for r in rows:
        k = (r[0], r[2])
        subs[k] = subs.get(k, 0.0) + r[3]
    count = len(subs)
    total = round(sum(r[3] for r in rows), 2)
    avg = round(total / count, 2) if count else 0
    highest = round(max(subs.values(), default=0), 2)
    pkg_label = "All Corporate Packages" if package == "ALL" else package

    out = io.BytesIO()
    wb = xlsxwriter.Workbook(out, {"in_memory": False, "constant_memory": True})

    title = wb.add_format({"bold": True, "font_size": 16, "font_color": DARK_GREEN})
    head = wb.add_format({"bold": True, "bg_color": GREEN, "font_color": "#ffffff", "border": 1, "valign": "vcenter"})
    head_r = wb.add_format({"bold": True, "bg_color": GREEN, "font_color": "#ffffff", "border": 1, "align": "right"})
    txt = wb.add_format({"num_format": "@"})
    money = wb.add_format({"num_format": "#,##0.00"})
    integer = wb.add_format({"num_format": "#,##0"})
    lbl = wb.add_format({"bold": True, "bg_color": "#f1f8ea", "border": 1})
    val = wb.add_format({"border": 1, "align": "left"})
    val_money = wb.add_format({"border": 1, "num_format": "#,##0.00", "align": "left"})
    val_int = wb.add_format({"border": 1, "num_format": "#,##0", "align": "left"})

    # ---- 1) Subscribers (full data, first sheet). Continues on the next sheet automatically if the Excel row limit is exceeded
    date_fmt = wb.add_format({"num_format": "dd-mm-yyyy", "align": "left"})
    _dates = {}

    def to_dt(text):
        d = _dates.get(text)
        if d is None:
            d = _dates[text] = datetime.strptime(text, "%Y-%m-%d")
        return d

    parts = max(1, -(-len(rows) // max_rows))
    for part in range(parts):
        ws = wb.add_worksheet("Subscribers" if part == 0 else f"Subscribers {part + 1}")
        ws.set_column(0, 0, 38)
        ws.set_column(1, 1, 14)
        ws.set_column(2, 2, 18, txt)
        ws.set_column(3, 3, 22, money)
        ws.write_row(0, 0, ["Package", "Date", "MSISDN"], head)
        ws.write(0, 3, "Total Revenue (BDT)", head_r)
        ws.freeze_panes(1, 0)
        chunk = rows[part * max_rows:(part + 1) * max_rows]
        for i, r in enumerate(chunk, start=1):
            ws.write_string(i, 0, r[0])
            ws.write_datetime(i, 1, to_dt(r[1]), date_fmt)
            ws.write_string(i, 2, str(r[2]))   # MSISDN stored as text, so no scientific notation
            ws.write_number(i, 3, r[3], money)
        ws.autofilter(0, 0, max(len(chunk), 1), 3)

    # ---- 2) Summary (second sheet)
    ws = wb.add_worksheet("Summary")
    ws.hide_gridlines(2)
    ws.set_column(0, 0, 36)
    ws.set_column(1, 1, 46)
    ws.set_landscape()
    ws.set_paper(9)
    ws.fit_to_pages(1, 0)

    banner = wb.add_format({"bold": True, "font_size": 20, "font_color": "#ffffff", "bg_color": DARK_GREEN,
                            "valign": "vcenter", "indent": 1})
    subtitle = wb.add_format({"font_size": 11, "font_color": "#000000", "bg_color": "#e6f3da",
                              "valign": "vcenter", "indent": 1})
    th_l = wb.add_format({"bold": True, "font_color": "#ffffff", "bg_color": GREEN, "border": 1,
                          "border_color": "#cfe3c0", "valign": "vcenter", "indent": 1})
    th_r = wb.add_format({"bold": True, "font_color": "#ffffff", "bg_color": GREEN, "border": 1,
                          "border_color": "#cfe3c0", "valign": "vcenter", "indent": 1})
    k_lbl = wb.add_format({"bold": True, "bg_color": "#f1f8ea", "border": 1, "border_color": "#cfe3c0",
                           "valign": "vcenter", "indent": 1})
    k_txt = wb.add_format({"border": 1, "border_color": "#cfe3c0", "valign": "vcenter", "indent": 1, "align": "left"})
    k_int = wb.add_format({"border": 1, "border_color": "#cfe3c0", "valign": "vcenter", "indent": 1,
                           "align": "left", "num_format": "#,##0"})
    k_mon = wb.add_format({"border": 1, "border_color": "#cfe3c0", "valign": "vcenter", "indent": 1,
                           "align": "left", "num_format": "#,##0.00"})
    k_big = wb.add_format({"border": 1, "border_color": "#cfe3c0", "valign": "vcenter", "indent": 1,
                           "align": "left", "num_format": "#,##0.00", "bold": True, "font_size": 13,
                           "font_color": DARK_GREEN, "bg_color": "#f6fbf1"})
    k_big_l = wb.add_format({"bold": True, "font_size": 13, "bg_color": "#e6f3da", "border": 1,
                             "border_color": "#cfe3c0", "valign": "vcenter", "indent": 1})

    # row heights must be set before writing rows (constant_memory mode)
    ws.set_row(0, 40)
    ws.set_row(1, 24)
    ws.set_row(2, 10)
    ws.set_row(3, 24)
    for r in range(4, 13):
        ws.set_row(r, 24)

    ws.merge_range(0, 0, 0, 1, "Corporate Subscriber & Revenue Analytics", banner)
    s_txt, e_txt = start_date.strftime("%d-%m-%Y"), end_date.strftime("%d-%m-%Y")
    period = s_txt if s_txt == e_txt else f"{s_txt}  to  {e_txt}"
    ws.merge_range(1, 0, 1, 1, f"{pkg_label}   |   {period}", subtitle)

    ws.write(3, 0, "Metric", th_l)
    ws.write(3, 1, "Value", th_r)
    rows_tbl = [
        ("Package", pkg_label, k_lbl, k_txt),
        *([("Filter Applied", filter_label, k_lbl, k_txt)] if filter_label else []),
        ("Start Date", s_txt, k_lbl, k_txt),
        ("End Date", e_txt, k_lbl, k_txt),
        ("Unique Subscribers", count, k_lbl, k_int),
        ("Total Revenue (BDT)", total, k_big_l, k_big),
        ("Avg Revenue / Subscriber (BDT)", avg, k_lbl, k_mon),
        ("Highest Subscriber Revenue (BDT)", highest, k_lbl, k_mon),
        ("Data Generated At", payload.get("generated_at", ""), k_lbl, k_txt),
    ]
    for i, (k, v, fl, fv) in enumerate(rows_tbl, start=4):
        ws.write(i, 0, k, fl)
        ws.write(i, 1, v, fv)

    # ---- 3) Top 10 (table + bar chart)
    top = sorted(subs.items(), key=lambda kv: kv[1], reverse=True)[:10]   # [((package, msisdn), total)]
    ws = wb.add_worksheet("Top 10")
    ws.set_column(0, 0, 8)
    ws.set_column(1, 1, 38)
    ws.set_column(2, 2, 18, txt)
    ws.set_column(3, 3, 22)
    ws.write_row(0, 0, ["Rank", "Package", "MSISDN"], head)
    ws.write(0, 3, "Total Revenue (BDT)", head_r)
    for i, ((pkg, msisdn), tot) in enumerate(top, start=1):
        ws.write_number(i, 0, i)
        ws.write_string(i, 1, pkg)
        ws.write_string(i, 2, str(msisdn))
        ws.write_number(i, 3, round(tot, 2), money)
    if top:
        ch = _bar_chart(wb, "bar", "Top 10 Subscribers by Revenue (BDT)", "Top 10",
                        ["Top 10", 1, 2, len(top), 2], ["Top 10", 1, 3, len(top), 3], reverse=True)
        ws.insert_chart(len(top) + 3, 0, ch)

    # ---- 4) By Package (All) / Distribution (single package)
    if package == "ALL":
        g = {}
        for (pkg, _msisdn), tot in subs.items():
            d = g.setdefault(pkg, [0, 0.0])
            d[0] += 1
            d[1] += tot
        ordered = sorted(g.items(), key=lambda x: x[1][1], reverse=True)
        ws = wb.add_worksheet("By Package")
        ws.set_column(0, 0, 40)
        ws.set_column(1, 2, 22)
        ws.write(0, 0, "Package", head)
        ws.write(0, 1, "Subscribers", head_r)
        ws.write(0, 2, "Total Revenue (BDT)", head_r)
        for i, (k, (c, rv)) in enumerate(ordered, start=1):
            ws.write_string(i, 0, k)
            ws.write_number(i, 1, c, integer)
            ws.write_number(i, 2, round(rv, 2), money)
        n = min(len(ordered), 20)
        if n:
            c1 = _bar_chart(wb, "bar", "Revenue by Package (BDT)", "By Package",
                            ["By Package", 1, 0, n, 0], ["By Package", 1, 2, n, 2], reverse=True)
            c2 = _bar_chart(wb, "bar", "Subscribers by Package", "By Package",
                            ["By Package", 1, 0, n, 0], ["By Package", 1, 1, n, 1], reverse=True, num_fmt="#,##0")
            ws.insert_chart(len(ordered) + 3, 0, c1)
            ws.insert_chart(len(ordered) + 3, 5, c2)
    else:
        labels = ["0", "0 - 1", "1 - 10", "10 - 100", "100 - 1000", "1000+"]
        b = [0] * 6
        for v in subs.values():
            b[0 if v <= 0 else 1 if v < 1 else 2 if v < 10 else 3 if v < 100 else 4 if v < 1000 else 5] += 1
        ws = wb.add_worksheet("Distribution")
        ws.set_column(0, 1, 24)
        ws.write(0, 0, "Revenue Range (BDT)", head)
        ws.write(0, 1, "Subscribers", head_r)
        for i, (l, c) in enumerate(zip(labels, b), start=1):
            ws.write_string(i, 0, l)
            ws.write_number(i, 1, c, integer)
        ch = _bar_chart(wb, "column", "Revenue Distribution (Subscribers)", "Distribution",
                        ["Distribution", 1, 0, 6, 0], ["Distribution", 1, 1, 6, 1],
                        x_title="Revenue Range (BDT)", y_title="Subscribers", num_fmt="#,##0")
        ws.insert_chart(9, 0, ch)

    wb.close()
    out.seek(0)
    return out
