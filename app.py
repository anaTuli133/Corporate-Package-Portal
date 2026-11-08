import gzip
import json
import os
import re
import threading
import time
from collections import OrderedDict
from datetime import datetime

from flask import Flask, Response, jsonify, render_template, request, send_file

from export import build_workbook
from utility import fetch_packages, fetch_subscriber_revenue

app = Flask(__name__)

# ---------------- Result cache (repeating the same search is instant) ----------------
CACHE_TTL = int(os.getenv("CACHE_TTL", 300))   # seconds (default 5 min)
CACHE_MAX = int(os.getenv("CACHE_MAX", 30))    # max cached searches
_cache = OrderedDict()                          # key -> (timestamp, gzip_bytes)
_cache_lock = threading.Lock()
_key_locks = {}


def _cache_get(key):
    with _cache_lock:
        e = _cache.get(key)
        if e and time.time() - e[0] < CACHE_TTL:
            _cache.move_to_end(key)
            return e
    return None


def _cache_put(key, gz):
    with _cache_lock:
        _cache[key] = (time.time(), gz)
        _cache.move_to_end(key)
        while len(_cache) > CACHE_MAX:
            _cache.popitem(last=False)


def get_result(key, loader, refresh=False):
    """Returns (gzip_bytes, hit, age_seconds). If several users run the same search at once, the DB is queried only once."""
    if not refresh:
        e = _cache_get(key)
        if e:
            return e[1], True, int(time.time() - e[0])
    with _cache_lock:
        lock = _key_locks.setdefault(key, threading.Lock())
    with lock:
        if not refresh:
            e = _cache_get(key)
            if e:
                return e[1], True, int(time.time() - e[0])
        payload = loader()
        gz = gzip.compress(json.dumps(payload, separators=(",", ":")).encode("utf-8"), 5)
        _cache_put(key, gz)
        return gz, False, 0


# ---------------- Routes ----------------
@app.route("/")
def index():
    return render_template("portal.html")


@app.route("/api/packages")
def api_packages():
    try:
        return jsonify(ok=True, packages=fetch_packages(force=request.args.get("refresh") == "1"))
    except Exception as e:
        return jsonify(ok=False, error=str(e)), 500


def _parse(package, sd_s, ed_s):
    """Validate input. Returns (package, sd, ed) or raises ValueError(msg)."""
    package = package or "ALL"
    try:
        sd = datetime.strptime(sd_s or "", "%Y-%m-%d").date()
        ed = datetime.strptime(ed_s or "", "%Y-%m-%d").date()
    except ValueError:
        raise ValueError("Invalid date")
    if sd > ed:
        raise ValueError("Start Date cannot be after End Date")
    if ed >= datetime.now().date():
        raise ValueError("Today's data is not available yet. Please select an End Date up to yesterday")
    if package != "ALL" and package not in fetch_packages():
        raise ValueError("Unknown package")
    return package, sd, ed


def _payload(package, sd, ed, refresh=False):
    def loader():
        t0 = time.time()
        rows, total = fetch_subscriber_revenue(package, sd, ed)
        return {
            "ok": True, "rows": rows, "total_revenue": total, "count": len(rows),
            "seconds": round(time.time() - t0, 2),
            "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p"),
        }
    return get_result((package, str(sd), str(ed)), loader, refresh)


@app.route("/api/search", methods=["POST"])
def api_search():
    body = request.get_json(silent=True) or {}
    try:
        package, sd, ed = _parse(body.get("package"), body.get("start_date"), body.get("end_date"))
        gz, hit, age = _payload(package, sd, ed, bool(body.get("refresh")))
    except ValueError as e:
        return jsonify(ok=False, error=str(e)), 400
    except Exception as e:
        return jsonify(ok=False, error=f"Database error: {e}"), 500

    headers = {"X-Cache": "HIT" if hit else "MISS", "X-Cache-Age": str(age)}
    if "gzip" in request.headers.get("Accept-Encoding", ""):
        headers["Content-Encoding"] = "gzip"
        headers["Vary"] = "Accept-Encoding"
        return Response(gz, mimetype="application/json", headers=headers)
    return Response(gzip.decompress(gz), mimetype="application/json", headers=headers)


REV_LABELS = {"nonzero": "Hide zero revenue", "zero": "Only zero revenue"}


def _parse_day(text):
    """'YYYY-MM-DD' -> date, or None if it is empty/invalid."""
    try:
        return datetime.strptime((text or "").strip(), "%Y-%m-%d").date()
    except ValueError:
        return None


def _download_name(package, sd, ed, rev, q, d, ext):
    """Readable file name, e.g.
       Corporate_Sec_Prepaid_3G_01-11-2026_to_05-11-2026.xlsx
       Corporate_Sec_Prepaid_3G_04-11-2026.xlsx                 (single date searched)
       All_Corporate_Packages_04-11-2026_Zero_Revenue.xlsx"""
    pkg = "All_Corporate_Packages" if package == "ALL" else package
    pkg = re.sub(r"[^\w\-]+", "_", pkg).strip("_")
    day = _parse_day(d)
    if day:
        when = day.strftime("%d-%m-%Y")
    elif sd == ed:
        when = sd.strftime("%d-%m-%Y")
    else:
        when = f"{sd.strftime('%d-%m-%Y')}_to_{ed.strftime('%d-%m-%Y')}"
    tag = {"nonzero": "_Non_Zero_Revenue", "zero": "_Zero_Revenue"}.get(rev, "")
    if (q or "").strip():
        tag += "_Filtered"
    return f"{pkg}_{when}{tag}.{ext}"


def _apply_filters(data, rev, q, d, pk=None):
    """Same filters as the table/CSV: revenue filter, MSISDN/package search and a single date.
    Row format: [package, 'YYYY-MM-DD', msisdn, revenue]."""
    q = (q or "").strip().lower()
    d = (d or "").strip()
    pk = (pk or "").strip()
    if rev not in REV_LABELS and not q and not d and not pk:
        return data, None
    rows = data["rows"]
    if pk:
        rows = [r for r in rows if r[0] == pk]
    if rev == "nonzero":
        rows = [r for r in rows if r[3] > 0]
    elif rev == "zero":
        rows = [r for r in rows if r[3] <= 0]
    if q:
        rows = [r for r in rows if q in str(r[2]) or q in r[0].lower()]
    if d:
        rows = [r for r in rows if r[1] == d]
    parts = []
    if rev in REV_LABELS:
        parts.append(REV_LABELS[rev])
    if d:
        day = _parse_day(d)
        parts.append(f"Date: {day.strftime('%d-%m-%Y') if day else d}")
    if q:
        parts.append(f'Search: "{q}"')
    data = dict(data, rows=rows, count=len(rows), total_revenue=round(sum(r[3] for r in rows), 2))
    return data, " | ".join(parts)


@app.route("/api/export")
def api_export():
    """Full data + summary + Top 10 + breakdown + charts in a single Excel file (served from cache, so it is fast)."""
    try:
        package, sd, ed = _parse(request.args.get("package"), request.args.get("start_date"), request.args.get("end_date"))
        gz, _, _ = _payload(package, sd, ed)
        data = json.loads(gzip.decompress(gz))
        rev = request.args.get("rev", "all")
        q = request.args.get("q")
        d = request.args.get("d")
        pk = (request.args.get("pk") or "").strip()
        data, filter_label = _apply_filters(data, rev, q, d, pk)
        day = _parse_day(d)
        # if one package was picked from an "All" search, the file is treated as that single package
        package = pk or package
        # if one date was searched, the file only contains that date
        xlsx = build_workbook(data, package, day or sd, day or ed, filter_label=filter_label)
    except ValueError as e:
        return jsonify(ok=False, error=str(e)), 400
    except Exception as e:
        return jsonify(ok=False, error=f"Export error: {e}"), 500

    return send_file(
        xlsx, as_attachment=True, download_name=_download_name(package, sd, ed, rev, q, d, "xlsx"),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def _warmup():
    try:
        fetch_packages()
    except Exception:
        pass


if __name__ == "__main__":
    threading.Thread(target=_warmup, daemon=True).start()
    app.run(debug=False, host="0.0.0.0", port=int(os.getenv("PORT", 5005)), threaded=True)
