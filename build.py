"""build.py  –  Read all Weekly Cash Update*.xlsx + Monday.com data → index.html
Usage: py build.py
"""
import sys, json, datetime, re, pathlib, io, urllib.request
sys.stdout.reconfigure(encoding="utf-8")
import openpyxl

FOLDER = pathlib.Path(
    r"C:\Users\MuhammadPrax\OneDrive - Pubity Group"
    r"\Pubity - Documents\Finance - Pubity Group\Bank\Monthly Cash Update"
)
BANK_FILE = pathlib.Path(
    r"C:\Users\MuhammadPrax\OneDrive - Pubity Group"
    r"\Pubity - Documents\Finance - Pubity Group\Bank\Weekly Bank Report - Pubity.xlsx"
)
OUT   = pathlib.Path(__file__).parent / "index.html"
TODAY = datetime.date.today()

# ── helpers ─────────────────────────────────────────────────────────────
def cv(v):
    if v is None: return ""
    if isinstance(v, (datetime.datetime, datetime.date)):
        return v.strftime("%d/%m/%y")
    return str(v).replace("\n", " ").replace("\t", " ").strip()

def num(v):
    return float(v) if isinstance(v, (int, float)) else 0.0

# ── parse bank report ───────────────────────────────────────────────────
def parse_bank_report():
    """Read Bal Summary sheet → latest TOTAL + 2026 daily series."""
    try:
        with open(BANK_FILE, "rb") as f:
            buf = io.BytesIO(f.read())
    except (PermissionError, OSError) as e:
        print(f"  ! Bank report not accessible: {e}")
        return None
    try:
        wb = openpyxl.load_workbook(buf, data_only=True)
    except Exception as e:
        print(f"  ! Cannot open bank report: {e}")
        return None

    if "Bal Summary" not in wb.sheetnames:
        print("  ! 'Bal Summary' sheet not found")
        return None

    rows = list(wb["Bal Summary"].iter_rows(values_only=True))
    hdr  = rows[1]  # row 2 is the column header row
    try:
        total_col = next(i for i, v in enumerate(hdr) if v == "TOTAL")
        date_col  = next(i for i, v in enumerate(hdr) if v == "Date")
    except StopIteration:
        print("  ! Could not locate TOTAL or Date column in Bal Summary")
        return None

    latest_val  = None
    latest_date = ""
    series_2026 = []

    for r in rows[2:]:
        d = r[date_col]
        t = r[total_col]
        if not isinstance(d, (datetime.datetime, datetime.date)): continue
        if not isinstance(t, (int, float)) or t <= 0: continue
        # Skip today (and anything later): the current day's transactions are
        # incomplete until the day closes, so the chart ends at yesterday.
        d_date = d.date() if isinstance(d, datetime.datetime) else d
        if d_date >= TODAY: continue
        if latest_val is None:
            latest_val  = round(float(t), 2)
            latest_date = cv(d)
        if d.year == 2026:
            series_2026.append({"date": cv(d), "total": round(float(t), 2)})

    series_2026.reverse()  # rows are newest-first → reverse to chronological

    # Forward-fill missing calendar dates (weekends, holidays, zero-balance days that
    # were filtered out) with the last known value so Chart.js draws a smooth line.
    if series_2026:
        filled = []
        prev_val = series_2026[0]["total"]
        existing = {p["date"]: p["total"] for p in series_2026}
        first_dt = datetime.datetime.strptime(series_2026[0]["date"], "%d/%m/%y")
        last_dt  = datetime.datetime.strptime(series_2026[-1]["date"], "%d/%m/%y")
        cur_dt   = first_dt
        while cur_dt <= last_dt:
            label = cur_dt.strftime("%d/%m/%y")
            if label in existing:
                prev_val = existing[label]
            filled.append({"date": label, "total": prev_val})
            cur_dt += datetime.timedelta(days=1)
        series_2026 = filled

    return {"latest": latest_val, "latestDate": latest_date, "series2026": series_2026}

# ── parse bank payments ─────────────────────────────────────────────────
def parse_bank_payments():
    """Read PUB GBP + USD Chase sheets → list of incoming credit transactions."""
    try:
        with open(BANK_FILE, "rb") as f:
            buf = io.BytesIO(f.read())
    except (PermissionError, OSError) as e:
        print(f"  ! Bank payments not accessible: {e}")
        return []
    try:
        wb = openpyxl.load_workbook(buf, data_only=True)
    except Exception as e:
        print(f"  ! Cannot open bank file for payments: {e}")
        return []

    payments = []

    # ── PUB GBP: col18=Narrative, col19=CustRef, col21=ValueDate, col22=CreditAmt ──
    if "PUB GBP" in wb.sheetnames:
        for row in wb["PUB GBP"].iter_rows(values_only=True):
            if len(row) < 23:
                continue
            credit_amt = row[22]
            if not isinstance(credit_amt, (int, float)) or credit_amt <= 0:
                continue
            narrative = cv(row[18])
            cust_ref  = cv(row[19])
            if not narrative and not cust_ref:
                continue
            payments.append({
                "narrative": narrative,
                "custRef":   cust_ref,
                "date":      cv(row[21]),
                "amount":    round(float(credit_amt), 2),
                "currency":  "GBP",
            })

    # ── USD Chase: col0=Date, col1=Description, col4=Amount (positive=credit) ──
    if "USD Chase" in wb.sheetnames:
        for row in wb["USD Chase"].iter_rows(values_only=True):
            if len(row) < 5:
                continue
            amount = row[4]
            if not isinstance(amount, (int, float)) or amount <= 0:
                continue
            narrative = cv(row[1])
            if not narrative:
                continue
            payments.append({
                "narrative": narrative,
                "custRef":   "",
                "date":      cv(row[0]),
                "amount":    round(float(amount), 2),
                "currency":  "USD",
            })

    gbp_n = sum(1 for p in payments if p["currency"] == "GBP")
    usd_n = sum(1 for p in payments if p["currency"] == "USD")
    print(f"    Bank payments: {len(payments)} credit transactions ({gbp_n} GBP, {usd_n} USD)")
    return payments

# ── Xero (new format) parsers ────────────────────────────────────────────
def _date_from_label(s):
    """Extract '20 May 2026' / '07 May 2026' from totals/prior labels → 'dd/mm/yyyy'."""
    if not s: return ""
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)\s+(\d{4})", str(s))
    if not m: return ""
    months = {n: i for i, n in enumerate(
        ["january","february","march","april","may","june","july",
         "august","september","october","november","december"], start=1)}
    mo = months.get(m.group(2).lower(), 0)
    if not mo: return ""
    return f"{int(m.group(1)):02d}/{mo:02d}/{m.group(3)}"


# Bad debt is flagged in the workbook by filling the row Office purple
# (#7030A0; #800080 also accepted). This is the authoritative marker and is
# maintained by Finance after each bad debt review - do not infer it from the
# Notes wording, which changed to case histories in the 24.09.26 file and
# silently returned zero. The keyword check is kept as a safety net.
_BAD_DEBT_FILLS = {"FF7030A0", "FF800080"}

def _purple_row_indices(ws, max_col=12):
    """0-based indices (matching list(ws.iter_rows())) of rows filled bad-debt purple."""
    out = set()
    for i, row in enumerate(ws.iter_rows(max_col=max_col)):
        for cell in row:
            try:
                fg = cell.fill.fgColor
                if fg.type == "rgb" and str(fg.rgb).upper() in _BAD_DEBT_FILLS:
                    out.add(i)
                    break
            except Exception:
                pass
    return out


def _parse_xero_sheet(ws, entity, default_currency):
    """Parse one UK/US Aged Receivables sheet. Returns (receivables, totals, prior, collections, coll_total)."""
    rows = list(ws.iter_rows(values_only=True))
    purple = _purple_row_indices(ws)
    tot_idx = None
    prior_idx = None
    coll_hdr_idx = None
    for i, r in enumerate(rows):
        a = str(r[0] or "").strip().lower()
        if a.startswith("total (") and tot_idx is None:
            tot_idx = i
        elif a.startswith("prior period"):
            prior_idx = i
        elif a.startswith("collections "):
            coll_hdr_idx = i

    # Fallback: label cells in col A are sometimes left blank. Detect the
    # totals row structurally — first row with a blank label but numeric
    # bucket values (col E onward). The "% of Total" row (values ≈ 1) and
    # fully blank rows are skipped.
    if tot_idx is None:
        limit = coll_hdr_idx if coll_hdr_idx is not None else len(rows)
        for i in range(3, limit):
            r = rows[i]
            if str(r[0] or "").strip():
                continue
            if isinstance(r[4], (int, float)) and abs(r[4]) > 1.5:
                tot_idx = i
                break
    if tot_idx is not None and prior_idx is None:
        limit = coll_hdr_idx if coll_hdr_idx is not None else min(tot_idx + 6, len(rows))
        for i in range(tot_idx + 1, limit):
            r = rows[i]
            if str(r[0] or "").strip():
                continue
            # skip the "% of Total" row (~1) and blank spacer rows; the next
            # large positive value is the prior-period row (movement row after
            # it can be negative, so require > 0)
            if isinstance(r[4], (int, float)) and r[4] > 1.5:
                prior_idx = i
                break

    receivables = []
    if tot_idx is not None:
        for _ri, r in enumerate(rows[2:tot_idx], start=2):
            if not r[0]:
                continue
            xl_comment   = cv(r[10]) if len(r) > 10 else ""
            xl_cust_note = cv(r[11]) if len(r) > 11 else ""
            bal = num(r[4])
            receivables.append({
                "vendorNo":   "",
                "vendorName": cv(r[0]),
                "postingDate": cv(r[1]),
                "docType":    "",
                "docNo":      cv(r[3]),
                "dueDate":    cv(r[2]),
                "origAmt":    bal,
                "balance":    bal,
                "notDue":     num(r[5]),
                "d1_31":      num(r[6]),   # 1-30
                "d32_61":     num(r[7]),   # 31-60
                "d63_92":     num(r[8]),   # 61-90
                "over92":     num(r[9]),   # 90+
                "xlComment":  xl_comment,
                "xlCustNote": xl_cust_note,
                "xlCondition": "",
                "isBadDebt":  (_ri in purple)
                              or "risk of bad debt" in xl_comment.lower()
                              or "potential bad debt" in xl_comment.lower(),
                "entity":     entity,
                "currency":   default_currency,
            })

    tot_row = rows[tot_idx] if tot_idx is not None else None
    totals = {
        "date":    cv(tot_row[0]) if tot_row else "",
        "balance": num(tot_row[4]) if tot_row else 0,
        "notDue":  num(tot_row[5]) if tot_row else 0,
        "d1_31":   num(tot_row[6]) if tot_row else 0,
        "d32_61":  num(tot_row[7]) if tot_row else 0,
        "d63_92":  num(tot_row[8]) if tot_row else 0,
        "over92":  num(tot_row[9]) if tot_row else 0,
    }

    prior = None
    if prior_idx is not None:
        pr = rows[prior_idx]
        prior = {
            "date":    cv(pr[0]),
            "balance": num(pr[4]),
            "notDue":  num(pr[5]),
            "d1_31":   num(pr[6]),
            "d32_61":  num(pr[7]),
            "d63_92":  num(pr[8]),
            "over92":  num(pr[9]),
        }

    collections = []
    coll_total = 0.0
    if coll_hdr_idx is not None:
        # +1 = column header row, +2 = first data row, ends on row whose col 0 starts with "Total"
        for r in rows[coll_hdr_idx + 2:]:
            cust = cv(r[0])
            if cust.lower().startswith("total"):
                coll_total = num(r[3])
                break
            if not cust and not r[1] and not r[2]:
                # a row with no customer/date/amount but a GBP value is the
                # Total Collections row with its label cell left blank
                if num(r[3]) > 0:
                    coll_total = num(r[3])
                    break
                continue
            collections.append({
                "date":     cv(r[1]),
                "customer": cust,
                "amt":      num(r[2]),
                "amtGBP":   num(r[3]),
                "isUSD":    default_currency == "USD",
            })
    if coll_total == 0:
        coll_total = sum(c["amtGBP"] for c in collections)

    return receivables, totals, prior, collections, coll_total


def _parse_xero(wb, path):
    uk_ws = wb["UK Aged Receivables"] if "UK Aged Receivables" in wb.sheetnames else None
    us_ws = wb["US Aged Receivables"] if "US Aged Receivables" in wb.sheetnames else None

    # Exchange rate from US sheet (row labelled "USD/GBP Exchange Rate", value in col E)
    exchange = 0.0
    if us_ws:
        for r in us_ws.iter_rows(values_only=True):
            if r[0] and "exchange rate" in str(r[0]).lower():
                if isinstance(r[4], (int, float)) and r[4] > 0:
                    exchange = float(r[4])
                break
    if exchange == 0.0:
        exchange = 1.0

    uk_recv, uk_tot, uk_prior, uk_coll, uk_coll_tot = (
        _parse_xero_sheet(uk_ws, "UK", "GBP") if uk_ws else ([], {}, None, [], 0.0)
    )
    us_recv_raw, us_tot_raw, us_prior_raw, us_coll, us_coll_tot = (
        _parse_xero_sheet(us_ws, "US", "USD") if us_ws else ([], {}, None, [], 0.0)
    )

    # USD → GBP multiplier
    fx = 1.0 / exchange if exchange else 0.0

    # Convert US row values to GBP (preserve USD originals)
    us_recv = []
    for r in us_recv_raw:
        nr = dict(r)
        for k in ("balance","notDue","d1_31","d32_61","d63_92","over92","origAmt"):
            nr[k + "USD"] = r[k]
            nr[k] = round(r[k] * fx, 2)
        us_recv.append(nr)

    # Per-entity totals (entity currency)
    uk_tot_gbp = dict(uk_tot)
    us_tot_gbp = {
        "date":    us_tot_raw.get("date", ""),
        **{k: round(num(us_tot_raw.get(k, 0)) * fx, 2)
           for k in ("balance","notDue","d1_31","d32_61","d63_92","over92")},
    }

    # Group consolidated totals (in GBP)
    group_tot = {
        "date": uk_tot_gbp.get("date") or us_tot_gbp.get("date") or "",
        **{k: round(num(uk_tot_gbp.get(k, 0)) + num(us_tot_gbp.get(k, 0)), 2)
           for k in ("balance","notDue","d1_31","d32_61","d63_92","over92")},
    }

    prior_uk = uk_prior or {}

    report_date = _date_from_label(uk_tot_gbp.get("date") or us_tot_gbp.get("date") or "")
    prev_date   = _date_from_label(prior_uk.get("date", ""))

    cash_balance = None  # bank report file is the source of truth

    bad_debt = sum(r["balance"] for r in (uk_recv + us_recv) if r["isBadDebt"])

    # Weighted avg days overdue (group)
    def _pd(s):
        if not s: return None
        for fmt in ("%d/%m/%y", "%d/%m/%Y"):
            try: return datetime.datetime.strptime(s, fmt).date()
            except: pass
        return None
    ref = _pd(report_date)
    all_bal, wsum = 0.0, 0.0
    if ref:
        for r in (uk_recv + us_recv):
            if r["balance"] <= 0: continue
            due  = _pd(r["dueDate"])
            days = max(0, (ref - due).days) if due else 0
            all_bal += r["balance"]
            wsum    += r["balance"] * days
    avg_days = round(wsum / all_bal, 1) if all_bal > 0 else 0.0

    print(f"Xero format: UK={len(uk_recv)} US={len(us_recv)} "
          f"FX={exchange:.4f} report={report_date} "
          f"group=£{group_tot['balance']:,.0f}")

    return {
        "receivables":      uk_recv + us_recv,
        "totals":           group_tot,
        "totalsPrev":       prior_uk if prior_uk else uk_tot_gbp,
        "exchangeRate":     exchange,
        "fxDate":           report_date,
        "collections":      uk_coll + us_coll,
        "collectionsTotal": uk_coll_tot + us_coll_tot,
        "reportDate":       report_date,
        "prevDate":         prev_date,
        "cashBalance":      cash_balance,
        "badDebt":          bad_debt,
        "avgDaysOvr":       avg_days,
        "entities": {
            "UK": {"receivables": uk_recv, "totals": uk_tot_gbp,
                   "collections": uk_coll, "collectionsTotal": uk_coll_tot,
                   "totalsPrev": prior_uk},
            "US": {"receivables": us_recv, "totals": us_tot_gbp,
                   "collections": us_coll, "collectionsTotal": us_coll_tot,
                   "totalsPrev": {}},
        },
        "format": "xero",
    }


# ── parse one workbook ───────────────────────────────────────────────────
def parse_workbook(path):
    try:
        with open(path, "rb") as f:
            buf = io.BytesIO(f.read())
    except (PermissionError, OSError) as e:
        print(f"  ! Skipping {path.name}: {e}")
        return None
    try:
        wb = openpyxl.load_workbook(buf, data_only=True)
    except Exception as e:
        print(f"  ! Cannot open {path.name}: {e}")
        return None

    sheets = set(wb.sheetnames)
    if "UK Aged Receivables" in sheets or "US Aged Receivables" in sheets:
        return _parse_xero(wb, path)

    if "Aged Receivables" not in wb.sheetnames:
        print(f"  ! No 'Aged Receivables' sheet in {path.name}")
        return None

    ws = wb["Aged Receivables"]
    ar_rows = list(ws.iter_rows(values_only=True))

    # Detect which rows are highlighted with the bad debt purple color.
    # Two variants are used: #7030A0 (Office purple) and #800080 (web purple).
    _BAD_DEBT_ARGBS = {"FF7030A0", "FF800080"}
    bad_debt_row_indices = set()
    for row_idx, row_cells in enumerate(ws.iter_rows(max_col=13)):
        for cell in row_cells:
            try:
                fg = cell.fill.fgColor
                if fg.type == "rgb" and fg.rgb.upper() in _BAD_DEBT_ARGBS:
                    bad_debt_row_indices.add(row_idx)
                    break
            except Exception:
                pass

    # Detect two totals rows: col C contains both "total" and "lcy"
    tot_list = [
        i for i, r in enumerate(ar_rows)
        if "total" in str(r[2] or "").lower() and "lcy" in str(r[2] or "").lower()
    ]
    if len(tot_list) < 2:
        print(f"  ! Cannot detect totals rows in {path.name}")
        return None
    tot_idx, prev_tot_idx = tot_list[0], tot_list[1]

    # AR data rows (skip header at index 1, skip blank rows)
    receivables = []
    for list_idx, r in enumerate(ar_rows[2:tot_idx], start=2):
        if not r[0] and not r[1]:
            continue
        xl_comment   = cv(r[13]) if len(r) > 13 else ""  # col 14: Notes/Comments
        xl_condition = cv(r[14]) if len(r) > 14 else ""  # col 15: Condition text
        is_bad_debt  = (list_idx in bad_debt_row_indices) or ("risk of bad debt" in xl_comment.lower())
        receivables.append({
            "vendorNo":    cv(r[0]),
            "vendorName":  cv(r[1]),
            "postingDate": cv(r[2]),
            "docType":     cv(r[3]),
            "docNo":       cv(r[4]),
            "dueDate":     cv(r[5]),
            "origAmt":     num(r[6]),
            "balance":     num(r[7]),
            "notDue":      num(r[8]),
            "d1_31":       num(r[9]),
            "d32_61":      num(r[10]),
            "d63_92":      num(r[11]),
            "over92":      num(r[12]),
            "xlComment":   xl_comment,
            "xlCondition": xl_condition,
            "isBadDebt":   is_bad_debt,
        })

    def _tot(row):
        return {
            "date":    cv(row[2]),
            "balance": num(row[7]),
            "notDue":  num(row[8]),
            "d1_31":   num(row[9]),
            "d32_61":  num(row[10]),
            "d63_92":  num(row[11]),
            "over92":  num(row[12]),
        }

    totals_cur  = _tot(ar_rows[tot_idx])
    totals_prev = _tot(ar_rows[prev_tot_idx])

    # Exchange rate: first positive numeric in col D after prev_tot_idx
    exchange = 0.0
    for r in ar_rows[prev_tot_idx + 1 : prev_tot_idx + 8]:
        if isinstance(r[3], (int, float)) and r[3] > 0:
            exchange = float(r[3]); break

    m1 = re.search(r"(\d{2}/\d{2}/\d{4})", totals_cur["date"])
    m2 = re.search(r"(\d{2}/\d{2}/\d{4})", totals_prev["date"])
    report_date = m1.group(1) if m1 else ""
    prev_date   = m2.group(1) if m2 else ""

    # Collections: find "Summary of collections" header, then read every data row
    # until the Total row.  Some dates in col A are stored as text strings rather
    # than datetime objects — we accept both by keying on col B (customer) / col D (GBP).
    collections, coll_total = [], 0.0
    coll_hdr_idx = None
    for i, r in enumerate(ar_rows):
        if "summary of collections" in str(r[0] or r[1] or "").lower():
            coll_hdr_idx = i
            break

    if coll_hdr_idx is not None:
        # coll_hdr_idx   = section title row  ("Summary of collections …")
        # coll_hdr_idx+1 = column header row  ("Date | Customer | Amount | Amount GBP")
        # coll_hdr_idx+2 = first data row
        for r in ar_rows[coll_hdr_idx + 2:]:
            row_text = " ".join(str(v or "") for v in r[:4]).lower()
            if "total" in row_text:              # reached the Total row
                if isinstance(r[3], (int, float)):
                    coll_total = float(r[3])
                break
            ra = num(r[2])
            rg = num(r[3])
            if r[1] or rg > 0:                  # has a customer name or a GBP amount
                collections.append({
                    "date":     cv(r[0]),
                    "customer": cv(r[1]),
                    "amt":      ra,
                    "amtGBP":   rg,
                    "isUSD":    abs(ra - rg) > 0.01 and rg > 0,
                })
    else:
        # Fallback for files without the header
        for r in ar_rows[prev_tot_idx + 1:]:
            if isinstance(r[0], (datetime.datetime, datetime.date)):
                ra, rg = num(r[2]), num(r[3])
                collections.append({
                    "date":     cv(r[0]),
                    "customer": cv(r[1]),
                    "amt":      ra,
                    "amtGBP":   rg,
                    "isUSD":    abs(ra - rg) > 0.01 and rg > 0,
                })
            elif collections and r[0] is None and isinstance(r[3], (int, float)) and r[3] > 0:
                coll_total = float(r[3])
                break

    if coll_total == 0:
        coll_total = sum(c["amtGBP"] for c in collections)
    print(f"    Collections: {len(collections)} rows, total=£{coll_total:,.2f}")

    # Verify col14 (Notes) mapping — printed once per file so you can confirm
    notes_present = [(r["vendorName"], r["xlComment"]) for r in receivables if r["xlComment"]]
    if notes_present:
        print(f"    Col14 notes ({len(notes_present)} rows with content):")
        for name, note in notes_present:
            print(f"      {name[:30]:<30}  {note[:60]}")

    # Cash in Bank from "Cash Estimate vs Actual" sheet
    cash_balance = None
    for sname in wb.sheetnames:
        sl = sname.lower()
        if "cash estimate" in sl or ("cash" in sl and "actual" in sl):
            for r in wb[sname].iter_rows(values_only=True):
                if r[1] and "actual cash" in str(r[1]).lower() and isinstance(r[2], (int, float)):
                    cash_balance = float(r[2])
            break

    # Bad debt: rows manually highlighted with purple (#7030A0 or #800080) in Excel.
    bad_debt = sum(r["balance"] for r in receivables if r["isBadDebt"])

    # Weighted avg days overdue (used for trend chart; independent of bad debt flag).
    def _parse_date(s):
        if not s: return None
        for fmt in ("%d/%m/%y", "%d/%m/%Y"):
            try: return datetime.datetime.strptime(s, fmt).date()
            except: pass
        return None

    ref_date     = _parse_date(report_date)
    all_balance  = 0.0
    weighted_sum = 0.0
    if ref_date:
        for r in receivables:
            if r["balance"] <= 0: continue
            due  = _parse_date(r["dueDate"])
            days = max(0, (ref_date - due).days) if due else 0
            all_balance  += r["balance"]
            weighted_sum += r["balance"] * days
    avg_days_ovr = round(weighted_sum / all_balance, 1) if all_balance > 0 else 0.0

    return {
        "receivables":      receivables,
        "totals":           totals_cur,
        "totalsPrev":       totals_prev,
        "exchangeRate":     exchange,
        "fxDate":           report_date,
        "collections":      collections,
        "collectionsTotal": coll_total,
        "reportDate":       report_date,
        "prevDate":         prev_date,
        "cashBalance":      cash_balance,
        "badDebt":          bad_debt,
        "avgDaysOvr":       avg_days_ovr,
    }

# ── scan folder ──────────────────────────────────────────────────────────
def _filename_date(p):
    """Parse DD.MM.YY from filename for sorting — most recent report first."""
    m = re.search(r'(\d{2})\.(\d{2})\.(\d{2})\.xlsx$', p.name, re.IGNORECASE)
    if m:
        d, mo, y = int(m.group(1)), int(m.group(2)), 2000 + int(m.group(3))
        try:
            return datetime.date(y, mo, d)
        except ValueError:
            pass
    return datetime.date.min

files = sorted(
    list(FOLDER.glob("BI-Weekly Cash Update*.xlsx")) +
    list(FOLDER.glob("Monthly Cash Update*.xlsx")) +
    list(FOLDER.glob("Bi-Monthly Cash Update*.xlsx")),
    key=_filename_date,
    reverse=True
)
print(f"Found {len(files)} files in folder")

file_data, file_names = {}, []
for p in files:
    print(f"  Parsing {p.name}…", end=" ")
    d = parse_workbook(p)
    if d:
        file_data[p.name] = d
        file_names.append(p.name)
        cash_str = f"£{d['cashBalance']:,.0f}" if d["cashBalance"] else "n/a"
        print(f"✓  {len(d['receivables'])} AR | {len(d['collections'])} coll | cash={cash_str}")
    else:
        print("skipped")

if not file_names:
    sys.exit("ERROR: No files could be parsed.")

default_file = file_names[0]
default_rows = len(file_data[default_file]["receivables"])
print(f"\n{len(file_names)} files ready.")
print(f"  Default (main table): {default_file}")
print(f"  Rows from default file: {default_rows} AR rows (rows 3 to first empty in 'Aged Receivables' sheet)")
print(f"  Report date: {file_data[default_file]['reportDate']}")

# ── trend data — all files sorted chronologically ────────────────────────
def _parse_report_date(fname):
    rd = file_data[fname]["reportDate"]
    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try: return datetime.datetime.strptime(rd, fmt)
        except: pass
    return datetime.datetime.min

trend_data = []
for fname in sorted(file_names, key=_parse_report_date):
    d = file_data[fname]
    t = d["totals"]
    total = t["balance"]
    if total <= 0:
        continue
    bd        = d.get("badDebt", 0.0)
    over92only = max(0.0, t["over92"] - bd)
    trend_data.append({
        "date":       d["reportDate"],
        "notDue":     round(t["notDue"]   / total * 100, 1),
        "d1_31":      round(t["d1_31"]    / total * 100, 1),
        "d32_61":     round(t["d32_61"]   / total * 100, 1),
        "d63_92":     round(t["d63_92"]   / total * 100, 1),
        "over92":     round(over92only    / total * 100, 1),
        "badDebt":    round(bd            / total * 100, 1),
        "avgDaysOvr": d.get("avgDaysOvr", 0.0),
        "d1_31_abs":  round(t["d1_31"],  2),
        "d32_61_abs": round(t["d32_61"], 2),
        "d63_92_abs": round(t["d63_92"], 2),
        "over92_abs": round(over92only,  2),
        "badDebt_abs":round(bd,          2),
    })
print(f"Trend data: {len(trend_data)} points  "
      f"({trend_data[0]['date']} → {trend_data[-1]['date']})")

# ── bank report ──────────────────────────────────────────────────────────
print("\nReading bank report…", end=" ")
bank_cash = parse_bank_report()
if bank_cash:
    print(f"✓  latest=£{bank_cash['latest']:,.0f}  2026 points={len(bank_cash['series2026'])}")
else:
    print("not available")

# ── Monday.com live fetch ────────────────────────────────────────────────
# Set your Monday.com API token here (v2 Personal API token from monday.com
# → Profile picture → Developers → My Access Tokens)
MONDAY_TOKEN = "eyJhbGciOiJIUzI1NiJ9.eyJ0aWQiOjYzNTIwNjQyOCwiYWFpIjoxMSwidWlkIjo3ODgxNDc3MSwiaWFkIjoiMjAyNi0wMy0xOVQxMjo1NjoyNi44NjRaIiwicGVyIjoibWU6d3JpdGUiLCJhY3RpZCI6MTU0MzY1ODIsInJnbiI6ImV1YzEifQ.buRfF_jKfLnSqBiSwnSIDU2jOevfPvhZRjar6aLOBSo"

# Board IDs
_BOARD_BILLING  = 5092118433  # Campaign Billing and Invoices
_BOARD_OVERVIEW = 5089619640  # 2026 Pubity Group Campaign Overview

def _fetch_monday_tbi():
    """
    3-step chain to resolve deal value:
      Billing (5092118433) → board_relation_mm0y38ta
      → Campaign Overview (5089619640) → board_relation_mm0y5r2x
      → HubSpot board (1361907559) → numbers__1
    Returns (list, sync_date_str) or (None, None) on failure.
    All mirror/lookup columns are intentionally avoided — API returns unsupported type.
    """
    if not MONDAY_TOKEN:
        return None, None

    def _gql(query):
        payload = json.dumps({"query": query}).encode("utf-8")
        req = urllib.request.Request(
            "https://api.monday.com/v2",
            data=payload,
            headers={
                "Content-Type": "application/json",
                "Authorization": MONDAY_TOKEN,
                "API-Version": "2024-04",
            },
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=15) as resp:
            return json.loads(resp.read().decode("utf-8"))

    def _fmt_gbp(raw):
        """Parse a numeric string and return £X,XXX or '' on failure."""
        if not raw:
            return ""
        try:
            return "£{:,.0f}".format(float(str(raw).replace(",", "")))
        except Exception:
            return ""

    # ── Single query: billing board → status columns + Deal Value mirror column
    try:
        raw1 = _gql("""
        query {
          boards(ids: [5092118433]) {
            items_page(limit: 200) {
              items {
                id
                name
                group { id }
                column_values(ids: [
                  "color_mm0v4cv1",
                  "color_mm0vvyz3",
                  "color_mm0vj2th",
                  "lookup_mm127xxv"
                ]) {
                  id text
                  ... on MirrorValue { display_value }
                }
              }
            }
          }
        }
        """)
    except Exception as e:
        print(f"  Monday.com fetch error: {e}")
        return None, None

    if "errors" in raw1:
        print(f"  Monday.com API errors: {raw1['errors']}")
        return None, None

    billing_items = (raw1.get("data", {}).get("boards") or [{}])[0] \
                        .get("items_page", {}).get("items", [])

    GROUP_MAP = {
        "group_mm0rffvp": "Campaigns Kicked Off",
        "group_mm0w5v1k": "Delivered Campaigns",
    }

    def _sum_mirror_value(display_value):
        """Sum comma-separated numeric values from a mirror column → '£X,XXX' or ''."""
        if not display_value:
            return ""
        try:
            total = sum(float(v.strip()) for v in display_value.split(",") if v.strip())
            return _fmt_gbp(str(total)) if total else ""
        except Exception:
            return ""

    results = []
    for item in billing_items:
        group_label = GROUP_MAP.get(item.get("group", {}).get("id", ""))
        if not group_label:
            continue
        col_text = {cv["id"]: (cv.get("text") or "").strip() for cv in item["column_values"]}
        col_dv   = {cv["id"]: (cv.get("display_value") or "")  for cv in item["column_values"]}
        dv = _sum_mirror_value(col_dv.get("lookup_mm127xxv", ""))
        results.append({
            "name":      item["name"],
            "dealValue": dv,
            "group":     group_label,
            "io":        col_text.get("color_mm0v4cv1", ""),
            "ncf":       col_text.get("color_mm0vvyz3", ""),
            "lineItems": col_text.get("color_mm0vj2th", ""),
        })

    sync_date = datetime.datetime.now(datetime.timezone.utc).strftime("%d/%m/%Y %H:%M UTC")
    return results, sync_date


# ── Attempt live fetch; fall back to static snapshot ────────────────────
print("\nFetching Monday.com data…", end=" ")
_live_tbi, _live_sync = _fetch_monday_tbi()
if _live_tbi is not None:
    MONDAY_TBI       = _live_tbi
    MONDAY_SYNC_DATE = _live_sync
    print(f"✓  {len(MONDAY_TBI)} campaigns  (synced {MONDAY_SYNC_DATE})")
else:
    # Static fallback snapshot (14/03/2026) — used when token not set or API unavailable
    MONDAY_TBI = [
        # ── Campaigns Kicked Off ──────────────────────────────────────────
        {"name": "283_Peacock_TheMiniatureWife_US_1225",
         "dealValue": "£130,309", "group": "Campaigns Kicked Off",
         "io": "Received",   "ncf": "Received",    "lineItems": "Not Chased"},
        {"name": "293_Popeyes_Onepiece_US_0125",
         "dealValue": "£154,618", "group": "Campaigns Kicked Off",
         "io": "Not Chased", "ncf": "Not Chased",  "lineItems": "Not Chased"},
        {"name": "xx_wb_themummy_uk_0226",
         "dealValue": "£90,000",  "group": "Campaigns Kicked Off",
         "io": "Chased",     "ncf": "Not Chased",  "lineItems": "Not Chased"},
        {"name": "294_AT&T_TurboLive_US_0326",
         "dealValue": "£52,035",  "group": "Campaigns Kicked Off",
         "io": "Not Chased", "ncf": "Not Chased",  "lineItems": "Not Chased"},
        {"name": "xx_hersheys_itsyourhappyplace_1225",
         "dealValue": "£446,775", "group": "Campaigns Kicked Off",
         "io": "Not Chased", "ncf": "Not Chased",  "lineItems": "Not Chased"},
        # ── Delivered Campaigns ───────────────────────────────────────────
        {"name": "248_capcom_residentevil9_us_0725",
         "dealValue": "£70,000",  "group": "Delivered Campaigns",
         "io": "Received",   "ncf": "Do we need ?", "lineItems": "Not Chased"},
        {"name": "290_alp_vanheist_us_0326",
         "dealValue": "£44,601",  "group": "Delivered Campaigns",
         "io": "Not Chased", "ncf": "Not Chased",   "lineItems": "Not Chased"},
    ]
    MONDAY_SYNC_DATE = "Static snapshot (14/03/2026)"
    if not MONDAY_TOKEN:
        print("skipped (no token set — using static snapshot)")
    else:
        print("failed — using static snapshot")

# ── Chaser notes: saved edits from dashboard (dashboard takes priority) ─
NOTES_FILE = pathlib.Path(__file__).parent / "chaser-notes.json"
try:
    chaser_notes = json.loads(NOTES_FILE.read_text(encoding="utf-8")) if NOTES_FILE.exists() else {}
except Exception:
    chaser_notes = {}

# ── data payload ─────────────────────────────────────────────────────────
data_obj = {
    "files":          file_names,
    "fileData":       file_data,
    "trendData":      trend_data,
    "bankCash":       bank_cash,
    "mondayTBI":      MONDAY_TBI,
    "mondaySyncDate": MONDAY_SYNC_DATE,
    "buildDate":      TODAY.strftime("%d/%m/%Y"),
    "chaserNotes":    chaser_notes,
    "bankPayments":   parse_bank_payments(),
}
data_js = json.dumps(data_obj, ensure_ascii=False, indent=2).replace("</", "<\\/")

# ── Pubity Group logo (base64-embedded so the HTML stays self-contained) ─
import base64
_LOGO_PATH = pathlib.Path(
    r"C:\Users\MuhammadPrax\OneDrive - Pubity Group"
    r"\Pubity - Documents\Design\Logos\Pubity Group"
    r"\PG_Lockup_Yellow_Black.png"
)
try:
    logo_b64 = "data:image/png;base64," + base64.b64encode(_LOGO_PATH.read_bytes()).decode()
except Exception:
    logo_b64 = ""  # logo unavailable — img src will be blank, element hidden by CSS

# Dark-mode logo: yellow-only variant (no black text, visible on dark backgrounds)
_LOGO_DARK_PATH = pathlib.Path(
    r"C:\Users\MuhammadPrax\OneDrive - Pubity Group"
    r"\Pubity - Documents\Design\Logos\Pubity Group"
    r"\PG_Lockup_Yellow.png"
)
try:
    logo_dark_b64 = "data:image/png;base64," + base64.b64encode(_LOGO_DARK_PATH.read_bytes()).decode()
except Exception:
    logo_dark_b64 = ""  # fall back to hiding if not found

# ── vendor JS: fetched once, cached locally, inlined into the HTML ────────
# Inlining avoids Chrome's document.write intervention, which can silently
# block CDN <script src> tags when the page is rendered by the password gate.
_VENDOR_DIR = pathlib.Path(__file__).parent / "vendor"
_VENDOR_DIR.mkdir(exist_ok=True)

def _vendor_js(name, url):
    p = _VENDOR_DIR / name
    if not p.exists() or p.stat().st_size < 10_000:
        print(f"  Downloading {name} from {url} …")
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=30) as resp:
            p.write_bytes(resp.read())
    js = p.read_text(encoding="utf-8")
    # a literal </script> inside the lib would terminate our inline tag early
    return js.replace("</script", "<\\/script")

try:
    _TAILWIND_JS = _vendor_js("tailwind.js", "https://cdn.tailwindcss.com")
    _CHARTJS_JS  = _vendor_js("chart.umd.js", "https://cdn.jsdelivr.net/npm/chart.js")
except Exception as e:
    print(f"  ! Vendor JS fetch failed ({e}) — falling back to CDN script tags")
    _TAILWIND_JS = _CHARTJS_JS = ""

# ══════════════════════════════════════════════════════════════════════════
# HTML TEMPLATE
# ══════════════════════════════════════════════════════════════════════════
HTML = r"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8"/>
  <meta name="viewport" content="width=device-width,initial-scale=1.0"/>
  <title>Aged Receivables Dashboard</title>
  __VENDOR_SCRIPTS__
  <style>
    body { font-family:'Inter','Segoe UI',sans-serif; background:#f1f5f9; }
    table { border-collapse:collapse; }
    th    { white-space:nowrap; }

    /* Table row colours — exact Excel bucket colours */
    .row-not-due  { background:#ffffff; color:#1a2744; }
    .row-1-31     { background:#00B050; color:#ffffff; }
    .row-32-61    { background:#FF8C00; color:#ffffff; }
    .row-63-92    { background:#FF0000; color:#ffffff; }
    .row-over-92  { background:#C00000; color:#ffffff; }
    .row-bad-debt { background:#7030A0; color:#ffffff; }
    tr:hover td   { filter:brightness(0.92); }

    .sticky-name { position:sticky;left:0;z-index:2;background:white;box-shadow:2px 0 5px rgba(0,0,0,0.07); }
    .sticky-th   { position:sticky;left:0;z-index:3;background:#1a2744; }
    .row-not-due  .sticky-name { background:#ffffff; }
    .row-1-31     .sticky-name { background:#00B050; }
    .row-32-61    .sticky-name { background:#FF8C00; }
    .row-63-92    .sticky-name { background:#FF0000; }
    .row-over-92  .sticky-name { background:#C00000; }
    .row-bad-debt .sticky-name { background:#7030A0; }

    /* Condition badges */
    .badge-not-due  { background:#c6efce;color:#276221; }
    .badge-1-31     { background:#ffeb9c;color:#7d4e00; }
    .badge-32-61    { background:#ffe0b2;color:#7d3000; }
    .badge-63-92    { background:#ffc7ce;color:#9c0006; }
    .badge-over-92  { background:#f4b8c1;color:#5e0011; }
    .badge-bad-debt { background:#e9d5ff;color:#4b1d7a; }

    .dov-ok     { color:#00B050; }
    .dov-mild   { color:#7d4e00; }
    .dov-warn   { color:#9c0006; }
    .dov-bad    { color:#9c0006; }
    .dov-danger { color:#5e0011;font-weight:600; }

    .kpi-card { transition:transform 0.18s ease-out,filter 0.18s ease-out; will-change:transform; }
    @keyframes spin { to { transform:rotate(360deg); } }
    .kpi-card:hover { transform:translateY(-3px); filter:brightness(1.06); }
    .float-scroll { position:fixed;bottom:0;left:0;right:0;height:14px;overflow-x:auto;overflow-y:hidden;z-index:200;background:#f1f5f9;border-top:1px solid #e2e8f0;display:none;box-shadow:0 -1px 4px rgba(0,0,0,0.06); }
    .float-scroll-inner { height:1px; }
    .tbl-wrap { overflow-x:auto; }
    .sort-btn { cursor:pointer;user-select:none; }
    .sort-btn:hover { text-decoration:underline; }
    .tab-active   { border-bottom:3px solid #1a2744;color:#1a2744;font-weight:600; }
    .tab-inactive { border-bottom:3px solid transparent;color:#64748b; }
    .tab-inactive:hover { color:#1a2744; }
    /* ── Dark mode ───────────────────────────────────────────────────── */
    body.dark { background:#0f1117 !important; color:#e2e8f0; }

    body.dark .bg-white        { background:#1a1d27 !important; }
    body.dark .bg-slate-50     { background:#1e2130 !important; }
    body.dark .bg-slate-100    { background:#252942 !important; }

    body.dark .border-slate-100,
    body.dark .border-slate-200       { border-color:#2d3148 !important; }
    body.dark .border-b               { border-color:#2d3148; }

    body.dark .text-slate-800,
    body.dark .text-slate-700         { color:#e2e8f0 !important; }
    body.dark .text-slate-600         { color:#cbd5e1 !important; }
    body.dark .text-slate-500         { color:#94a3b8 !important; }
    body.dark .text-slate-400         { color:#64748b !important; }
    body.dark .text-slate-300         { color:#475569 !important; }

    body.dark .shadow-sm  { box-shadow:0 1px 4px rgba(0,0,0,0.5) !important; }

    body.dark input, body.dark select, body.dark textarea {
      background:#1e2130 !important;
      border-color:#2d3148 !important;
      color:#e2e8f0 !important;
    }
    body.dark input::placeholder { color:#64748b !important; }

    body.dark .tab-active   { border-bottom-color:#818cf8 !important; color:#818cf8 !important; }
    body.dark .tab-inactive { color:#94a3b8 !important; }
    body.dark .tab-inactive:hover { color:#e2e8f0 !important; }

    /* Sticky column in dark mode */
    body.dark .sticky-name {
      background:#1a1d27 !important;
      box-shadow:2px 0 5px rgba(0,0,0,0.4) !important;
    }

    /* Table rows in dark mode — uniform dark, left border colour set inline */
    body.dark .row-not-due,
    body.dark .row-1-31,
    body.dark .row-32-61,
    body.dark .row-63-92,
    body.dark .row-over-92,
    body.dark .row-bad-debt { background:#1a1d27 !important; }

    body.dark .row-not-due  .sticky-name,
    body.dark .row-1-31     .sticky-name,
    body.dark .row-32-61    .sticky-name,
    body.dark .row-63-92    .sticky-name,
    body.dark .row-over-92  .sticky-name,
    body.dark .row-bad-debt .sticky-name { background:#1a1d27 !important; }

    body.dark tr:hover td   { filter:brightness(1.18) !important; }

    /* Drill-down modal */
    body.dark #drill-modal > div { background:#1a1d27 !important; }
    body.dark #drill-contact-panel { background:#1e2130 !important; }
    body.dark #drill-chaser-panel  { background:#1a2035 !important; }
    body.dark #drill-chaser-body   { background:#1e2130 !important; }

    /* Inline dark-coloured headings and text */
    body.dark #hdr-title { color:#e2e8f0 !important; }
    body.dark h2 { color:#e2e8f0 !important; }
    body.dark h3 { color:#e2e8f0 !important; }

    /* Logo — show yellow-only variant in dark mode; hide it in light mode */
    #logo-dark { display:none; }
    body.dark #logo-light { display:none !important; }
    body.dark #logo-dark  { display:block !important; }

    /* Print / PDF export ──────────────────────────────────────────── */
    .print-only { display:none !important; }

    @media print {
      @page {
        size: A4 landscape;
        margin: 12mm 10mm 18mm 10mm;
        @bottom-center {
          content: "Page " counter(page) " of " counter(pages)
                   "  ·  Prepared by Finance | Pubity Group";
          font-size: 8pt;
          color: #64748b;
        }
      }

      /* Show print-only, hide interactive */
      .print-only        { display: block !important; }
      .no-print          { display: none  !important; }
      .pdf-exclude       { display: none  !important; }
      #drill-modal       { display: none  !important; }
      #panel-tbi         { display: none  !important; }
      #panel-nts         { display: none  !important; }
      #drill-contact-panel,
      #drill-chaser-panel { display: none !important; }

      /* Force AR + Collections visible even if on another tab */
      #panel-ar.hidden,
      #panel-col.hidden  { display: block !important; }

      /* Layout cleanup */
      body               { background: #fff !important; font-size: 9pt; }
      * { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
      .shadow-sm, .shadow-2xl { box-shadow: none !important; }
      .rounded-xl, .rounded-lg { border-radius: 4px !important; }
      .max-w-screen-2xl  { max-width: 100% !important; padding: 0 !important; }
      .tbl-wrap          { overflow: visible !important; }
      .sticky-name, .sticky-th { position: static !important; }
      .overflow-auto     { overflow: visible !important; }

      /* Page-break rules */
      tr                 { page-break-inside: avoid; }
      thead              { display: table-header-group; }
      tfoot              { display: table-footer-group; }
      h2                 { page-break-after: avoid; }
      .pdf-section       { page-break-inside: avoid; }
      .pdf-page-break    { page-break-before: always; }

      /* Chart / canvas sizing */
      canvas             { max-width: 100% !important; }
      #avgDaysChart      { height: 90px   !important; }

      /* Grid: force 2-col in print (charts | kpis+comparison) */
      .xl\:grid-cols-3   { grid-template-columns: 2fr 1fr !important; }
      .xl\:col-span-2    { grid-column: span 1 !important; }
      .xl\:col-span-1    { grid-column: span 1 !important; }

      /* Section dividers for multi-section print */
      .pdf-section-title {
        font-size: 10pt; font-weight: 700; color: #0e2841;
        padding: 6px 0 4px; border-bottom: 1.5px solid #156082;
        margin: 14px 0 8px; letter-spacing: 0.02em;
      }
    }
  </style>
</head>
<body class="min-h-screen">

<!-- PDF PRINT HEADER (hidden on screen, shown in print) -->
<div class="print-only" id="pdf-print-hdr" style="padding:0 0 10px;border-bottom:2.5px solid #156082;margin-bottom:16px">
  <div style="display:flex;align-items:center;justify-content:space-between">
    <div style="display:flex;align-items:center;gap:14px">
      <img src="__LOGO__" style="height:32px;width:auto" onerror="this.style.display='none'">
      <div>
        <div style="font-size:16pt;font-weight:700;color:#0e2841;line-height:1.2">Aged Accounts Receivable</div>
        <div style="font-size:9pt;color:#64748b;margin-top:2px">
          Pubity Group &nbsp;·&nbsp; As at <span id="pdf-hdr-date" style="font-weight:600;color:#0e2841"></span>
        </div>
      </div>
    </div>
    <div style="text-align:right;font-size:8.5pt;color:#64748b;line-height:1.6">
      <div>For internal use only</div>
    </div>
  </div>
</div>

<!-- HEADER -->
<div class="bg-white border-b border-slate-200 shadow-sm no-print">
  <div class="max-w-screen-2xl mx-auto px-6 py-4 flex flex-wrap items-center gap-4 justify-between">
    <div class="flex items-center gap-4">
      <img id="logo-light" src="__LOGO__" alt="Pubity Group" class="h-10 w-auto" style="__LOGO_HIDE__">
      <img id="logo-dark"  src="__LOGO_DARK__" alt="Pubity Group" class="h-10 w-auto" style="__LOGO_DARK_HIDE__">
      <div>
      <h1 id="hdr-title" class="text-2xl font-bold" style="color:#1a2744">Aged Accounts Receivable</h1>
      <p class="text-sm text-slate-500 mt-0.5">
        As at <span id="hdr-date" class="font-semibold text-slate-700"></span>
        &nbsp;|&nbsp; Built <span id="hdr-built" class="text-slate-400"></span>
        &nbsp;|&nbsp; USD/GBP <span id="hdr-fx" class="font-mono text-slate-600"></span>
        <span id="hdr-fx-date" class="text-slate-400 text-xs"></span>
      </p>
      </div>
    </div>
    <div class="flex gap-3 items-center flex-wrap">
      <div class="flex items-center gap-2">
        <label class="text-xs text-slate-500 font-medium whitespace-nowrap">Weekly file:</label>
        <select id="file-picker"
                class="px-3 py-1.5 text-sm border border-slate-200 rounded-lg bg-white text-slate-700 font-medium shadow-sm">
        </select>
      </div>
      <div id="entity-toggle-wrap" class="hidden items-center gap-2">
        <label class="text-xs text-slate-500 font-medium whitespace-nowrap">Entity:</label>
        <div class="inline-flex rounded-lg overflow-hidden border border-slate-200 shadow-sm bg-white">
          <button data-entity="Group" onclick="setEntity('Group')"
                  class="px-3 py-1.5 text-xs font-semibold bg-slate-900 text-white">Group</button>
          <button data-entity="UK" onclick="setEntity('UK')"
                  class="px-3 py-1.5 text-xs font-semibold text-slate-600 hover:bg-slate-50 border-l border-slate-200">UK</button>
          <button data-entity="US" onclick="setEntity('US')"
                  class="px-3 py-1.5 text-xs font-semibold text-slate-600 hover:bg-slate-50 border-l border-slate-200">US</button>
        </div>
      </div>
      <button id="dark-toggle" onclick="toggleDarkMode()" title="Switch to dark mode"
              class="p-2 rounded-lg bg-slate-100 hover:bg-slate-200 text-slate-600 transition">
        <svg class="w-5 h-5" fill="currentColor" viewBox="0 0 20 20"><path d="M17.293 13.293A8 8 0 016.707 2.707a8.001 8.001 0 1010.586 10.586z"/></svg>
      </button>
      <button onclick="window.print()"
              class="px-4 py-2 text-sm bg-slate-100 hover:bg-slate-200 rounded-lg text-slate-700 font-medium transition">
        Print
      </button>
      <button onclick="exportPDF()"
              class="px-4 py-2 text-sm rounded-lg font-medium transition flex items-center gap-2"
              style="background:#1a2744;color:#fff">
        <svg class="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24">
          <path stroke-linecap="round" stroke-linejoin="round" stroke-width="2"
            d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414A1 1 0 0119 9.414V19a2 2 0 01-2 2z"/>
        </svg>
        Export PDF
      </button>
    </div>
  </div>
</div>

<div class="max-w-screen-2xl mx-auto px-6 py-6">

  <!-- KPI CARDS -->
  <div id="kpi-row" class="mb-4" style="display:flex;flex-wrap:nowrap;gap:12px"></div>

  <!-- CASH IN BANK BANNER -->
  <div id="cash-banner" class="mb-6"></div>

  <!-- CHARTS + COMPARISON -->
  <div class="grid grid-cols-1 xl:grid-cols-3 gap-4 mb-6">
    <div class="xl:col-span-2 flex flex-col gap-4">

      <!-- Donut chart -->
      <div class="bg-white rounded-xl border border-slate-200 p-5">
        <h2 id="donut-title" style="font-size:13px;font-weight:600;color:#1a2744;margin-bottom:16px"></h2>
        <div class="flex justify-center" style="max-width:340px;margin:0 auto">
          <div id="donutChart" style="width:100%;position:relative"></div>
        </div>
      </div>

      <!-- avg days overdue trend -->
      <div class="bg-white rounded-xl border border-slate-200 p-5">
        <h2 style="font-size:13px;font-weight:600;color:#1a2744;margin-bottom:4px">Average Days Overdue — Trend</h2>
        <p id="avg-days-subtitle" class="text-xs text-slate-400 mb-4">
          Weighted average: sum(balance × max(0, days overdue)) / total balance · Current counts as 0 days
        </p>
        <canvas id="avgDaysChart" height="110"></canvas>
      </div>

      <!-- bucket drill-down -->
      <div class="bg-white rounded-xl border border-slate-200 p-5 pdf-exclude">
        <div class="flex items-center justify-between flex-wrap gap-2 mb-1">
          <h2 id="bucket-title" style="font-size:13px;font-weight:600;color:#1a2744"></h2>
          <select id="bucket-picker" onchange="renderBucketDrilldown(this.value)"
                  class="px-3 py-1.5 text-sm border border-slate-200 rounded-lg bg-white text-slate-700 no-print">
            <option value="d1_31_abs">1–30 Days</option>
            <option value="d32_61_abs">31–60 Days</option>
            <option value="d63_92_abs">61–90 Days</option>
            <option value="over92_abs">90+ Days</option>
            <option value="badDebt_abs">Risk of Bad Debt</option>
          </select>
        </div>
        <p id="bucket-subtitle" class="text-xs text-slate-400 mb-3"></p>
        <canvas id="bucketChart" height="110"></canvas>
      </div>

    </div>
    <div class="xl:col-span-1 flex flex-col gap-4">
      <div class="bg-white rounded-xl border border-slate-200 p-5">
        <h2 style="font-size:13px;font-weight:600;color:#1a2744;margin-bottom:12px">Period Comparison</h2>
        <div id="comparison-table" class="text-sm"></div>
      </div>
    </div>
  </div>

  <!-- CASH IN BANK 2026 -->
  <div class="bg-white rounded-xl border border-slate-200 mb-6 overflow-hidden pdf-exclude">
    <div class="px-5 py-3 flex items-center justify-between flex-wrap gap-3" style="background:#1a2744">
      <div>
        <h2 style="font-size:13px;font-weight:600;color:#ffffff">Cash in Bank — 2026</h2>
        <p style="font-size:11px;color:rgba(255,255,255,0.6);margin-top:2px">Daily total GBP balance &middot; Pubity Group all accounts combined</p>
      </div>
      <div id="cash-hdr-val" class="text-right"></div>
    </div>
    <div class="p-5">
      <canvas id="cashChart" height="65"></canvas>
    </div>
  </div>

  <!-- TABS -->
  <div class="bg-white rounded-xl border border-slate-200 mb-6">
    <div class="flex border-b border-slate-200 px-4 no-print">
      <button onclick="showTab('ar')"  id="tab-ar"  class="tab-active  px-4 py-3 text-sm transition">Aged Receivables</button>
      <button onclick="showTab('col')" id="tab-col" class="tab-inactive px-4 py-3 text-sm transition">Collections</button>
      <button onclick="showTab('tbi')" id="tab-tbi" class="tab-inactive px-4 py-3 text-sm transition">To Be Invoiced</button>
      <button onclick="showTab('nts')" id="tab-nts" class="tab-inactive px-4 py-3 text-sm transition">Notes</button>
    </div>

    <!-- AGED RECEIVABLES -->
    <div id="panel-ar">
      <div class="print-only pdf-section-title" style="margin:14px 12px 0">Aged Receivables</div>
      <div class="px-4 py-3 flex flex-wrap gap-3 border-b border-slate-100 bg-slate-50 no-print">
        <input id="search" type="text" placeholder="Search vendor, doc no…"
               oninput="applyFilters()"
               class="px-3 py-1.5 text-sm border border-slate-200 rounded-lg bg-white w-56"/>
        <select id="fCond" onchange="applyFilters()"
                class="px-3 py-1.5 text-sm border border-slate-200 rounded-lg bg-white">
          <option value="">All conditions</option>
          <option>Current</option>
          <option>≤30 Days Overdue</option>
          <option>31–60 Days Overdue</option>
          <option>61–90 Days Overdue</option>
          <option>90+ Days Overdue</option>
          <option>Risk of Bad Debt</option>
        </select>
        <button onclick="clearFilters()"
                class="px-3 py-1.5 text-sm text-slate-500 hover:text-slate-700 underline">Clear</button>
        <span id="rowCount" class="text-xs text-slate-400 self-center"></span>
      </div>
      <div class="tbl-wrap" id="ar-wrap">
        <table class="w-full text-xs" id="mainTable">
          <thead style="background:#1a2744;color:#ffffff;text-align:left">
            <tr>
              <th class="px-3 py-2.5 sticky-th sort-btn" onclick="sortBy('vendorName')" style="min-width:180px">Vendor Name ↕</th>
              <th class="px-3 py-2.5 sort-btn" onclick="sortBy('postingDate')">Posting Date ↕</th>
              <th class="px-3 py-2.5">Document No.</th>
              <th class="px-3 py-2.5 sort-btn" onclick="sortBy('dueDate')">Due Date ↕</th>
              <th class="px-3 py-2.5 text-right sort-btn" onclick="sortBy('_daysOvr')">Days Overdue ↕</th>
              <th class="px-3 py-2.5 text-right sort-btn" onclick="sortBy('origAmt')">Orig. Amount ↕</th>
              <th class="px-3 py-2.5 text-right sort-btn" onclick="sortBy('balance')">Balance ↕</th>
              <th class="px-2 py-2.5 text-right">Current</th>
              <th class="px-2 py-2.5 text-right">1–30d</th>
              <th class="px-2 py-2.5 text-right">31–60d</th>
              <th class="px-2 py-2.5 text-right">61–90d</th>
              <th class="px-2 py-2.5 text-right">90+d</th>
              <th class="px-3 py-2.5">Condition</th>
            </tr>
          </thead>
          <tbody id="arBody"></tbody>
          <tfoot id="arFoot" class="bg-slate-700 text-white font-bold text-xs"></tfoot>
        </table>
      </div>
    </div>

    <!-- COLLECTIONS -->
    <div id="panel-col" class="hidden">
      <div class="print-only pdf-section-title pdf-page-break" style="margin:14px 12px 0">Collections Received</div>
      <div class="px-4 py-3 border-b border-slate-100 bg-slate-50 no-print">
        <h3 class="text-sm font-semibold text-slate-700">Collections received</h3>
      </div>
      <div class="tbl-wrap">
        <table class="w-full text-xs">
          <thead style="background:#1a2744;color:#ffffff;text-align:left">
            <tr>
              <th class="px-3 py-2.5">Date</th>
              <th class="px-3 py-2.5">Customer</th>
              <th class="px-3 py-2.5 text-right">Amount</th>
              <th class="px-3 py-2.5 text-right">Amount (GBP)</th>
            </tr>
          </thead>
          <tbody id="collBody"></tbody>
          <tfoot id="collFoot" class="bg-slate-700 text-white font-bold text-xs"></tfoot>
        </table>
      </div>
    </div>

    <!-- TO BE INVOICED — Monday.com -->
    <div id="panel-tbi" class="hidden">
      <div class="px-4 py-3 border-b border-slate-100 bg-slate-50 flex items-center justify-between flex-wrap gap-2">
        <div>
          <h3 class="text-sm font-semibold text-slate-700">
            To Be Invoiced
            <span class="ml-2 px-2 py-0.5 text-xs bg-indigo-100 text-indigo-700 rounded-full font-normal">monday.com</span>
          </h3>
          <p class="text-xs text-slate-500 mt-0.5">Board: 2026 Pubity Group Campaign Overview</p>
        </div>
        <div class="flex items-center gap-3">
          <div class="text-xs text-slate-400">
            Last synced: <span id="monday-sync" class="font-medium text-slate-600"></span>
          </div>
          <button id="tbi-refresh-btn" onclick="refreshTBIFromMonday()"
            class="flex items-center gap-1.5 px-3 py-1.5 text-xs font-medium rounded-md bg-indigo-600 text-white hover:bg-indigo-700 active:bg-indigo-800 transition-colors no-print"
            style="line-height:1">
            <svg id="tbi-refresh-icon" xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"></polyline><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"></path></svg>
            <span id="tbi-refresh-label">Refresh</span>
          </button>
        </div>
      </div>
      <div id="tbi-error" class="hidden px-4 py-2 text-xs text-red-700 bg-red-50 border-b border-red-200"></div>
      <div id="tbi-content"></div>
    </div>

    <!-- NOTES -->
    <div id="panel-nts" class="hidden p-5 space-y-6">

      <!-- Chaser Log -->
      <div>
        <div class="flex items-center justify-between mb-3">
          <h3 class="text-sm font-bold text-slate-700 flex items-center gap-2">
            <span class="w-2 h-2 rounded-full bg-amber-500 inline-block"></span>
            Chaser Log
            <span class="text-xs font-normal text-slate-400 ml-1">(all overdue — oldest chased first)</span>
          </h3>
          <button onclick="exportNotes()"
                  class="px-3 py-1.5 text-xs bg-slate-700 text-white rounded hover:bg-slate-800 font-medium">
            Export Notes JSON
          </button>
        </div>
        <div class="tbl-wrap">
          <table class="w-full text-xs border border-slate-200 rounded-lg overflow-hidden">
            <thead class="bg-slate-100 text-slate-600">
              <tr>
                <th class="px-3 py-2 text-left">Customer</th>
                <th class="px-3 py-2 text-left">Days OD</th>
                <th class="px-3 py-2 text-right">Balance</th>
                <th class="px-3 py-2 text-left" style="min-width:240px">Notes</th>
              </tr>
            </thead>
            <tbody id="chaser-log-body"></tbody>
          </table>
        </div>
      </div>

    </div>

  </div><!-- /tabs -->

  <!-- LEGEND -->
  <div class="bg-white rounded-xl shadow-sm border border-slate-100 px-5 py-4 no-print">
    <p class="text-xs font-semibold text-slate-500 uppercase tracking-wide mb-3">Condition Legend — colors from Excel (calculated live from today's date)</p>
    <div class="flex flex-wrap gap-4 text-xs">
      <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded inline-block border border-slate-200" style="background:#ffffff"></span>Current</span>
      <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded inline-block" style="background:#00B050"></span>≤30 Days — Send reminder</span>
      <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded inline-block" style="background:#FFC000"></span>31–60 Days — Follow up</span>
      <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded inline-block" style="background:#FF0000"></span>61–90 Days — Chase actively</span>
      <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded inline-block" style="background:#C00000"></span>90+ Days — Chase actively</span>
      <span class="flex items-center gap-1.5"><span class="w-3 h-3 rounded inline-block" style="background:#7030A0"></span>Risk of Bad Debt — Escalate</span>
    </div>
  </div>

</div><!-- /container -->

<script>
const DATA = __DATA__;

// ── Dark mode ───────────────────────────────────────────────────────────
function _isDark() { return document.body.classList.contains("dark"); }

function _cg() {  // chart color palette — dark or light
  return _isDark()
    ? { grid:"#2d3148", tick:"#94a3b8", legend:"#94a3b8", border:"#1a1d27" }
    : { grid:"#f1f5f9", tick:"#64748b", legend:"#64748b", border:"#ffffff" };
}

function _updateDarkBtn() {
  const btn = document.getElementById("dark-toggle");
  if (!btn) return;
  if (_isDark()) {
    btn.title = "Switch to light mode";
    btn.innerHTML = `<svg class="w-5 h-5" fill="currentColor" viewBox="0 0 20 20">
      <path fill-rule="evenodd" d="M10 2a1 1 0 011 1v1a1 1 0 11-2 0V3a1 1 0 011-1zm4 8a4 4 0 11-8 0 4 4 0 018 0zm-.464 4.95l.707.707a1 1 0 001.414-1.414l-.707-.707a1 1 0 00-1.414 1.414zm2.12-10.607a1 1 0 010 1.414l-.706.707a1 1 0 11-1.414-1.414l.707-.707a1 1 0 011.414 0zM17 11a1 1 0 100-2h-1a1 1 0 100 2h1zm-7 4a1 1 0 011 1v1a1 1 0 11-2 0v-1a1 1 0 011-1zM5.05 6.464A1 1 0 106.465 5.05l-.708-.707a1 1 0 00-1.414 1.414l.707.707zm1.414 8.486l-.707.707a1 1 0 01-1.414-1.414l.707-.707a1 1 0 011.414 1.414zM4 11a1 1 0 100-2H3a1 1 0 000 2h1z" clip-rule="evenodd"/>
      </svg>`;
  } else {
    btn.title = "Switch to dark mode";
    btn.innerHTML = `<svg class="w-5 h-5" fill="currentColor" viewBox="0 0 20 20">
      <path d="M17.293 13.293A8 8 0 016.707 2.707a8.001 8.001 0 1010.586 10.586z"/>
      </svg>`;
  }
}

function toggleDarkMode() {
  const isDark = document.body.classList.toggle("dark");
  try { localStorage.setItem("pgDarkMode", isDark ? "1" : "0"); } catch {}
  _updateDarkBtn();
  // Re-render all color-dependent elements
  renderDonut();
  renderAvgDaysChart();
  renderBucketDrilldown(document.getElementById("bucket-picker")?.value || "d1_31_abs");
  renderCashLine();
  renderCashCard();
}

function initDarkMode() {
  try {
    if (localStorage.getItem("pgDarkMode") === "1") {
      document.body.classList.add("dark");
    }
  } catch {}
  _updateDarkBtn();
}

// ── state ──────────────────────────────────────────────────────────────
let curFile = DATA.files[0];
let CURRENT_ENTITY = "Group";   // "Group" | "UK" | "US" — only meaningful for Xero files
let D = _viewForFile(curFile);
let filtered = [];

function _viewForFile(file) {
  const base = DATA.fileData[file];
  if (!base || !base.entities || CURRENT_ENTITY === "Group") return base;
  const e = base.entities[CURRENT_ENTITY] || {};
  return {
    ...base,
    receivables:      e.receivables || [],
    totals:           e.totals && Object.keys(e.totals).length ? e.totals : base.totals,
    totalsPrev:       (e.totalsPrev && Object.keys(e.totalsPrev).length) ? e.totalsPrev : (base.totalsPrev || {}),
    collections:      e.collections || [],
    collectionsTotal: e.collectionsTotal || 0,
  };
}

function setEntity(entity) {
  CURRENT_ENTITY = entity;
  D = _viewForFile(curFile);
  // Restyle toggle buttons
  document.querySelectorAll("[data-entity]").forEach(btn => {
    const active = btn.dataset.entity === CURRENT_ENTITY;
    btn.className = "px-3 py-1.5 text-xs font-semibold " +
      (active ? "bg-slate-900 text-white" : "text-slate-600 hover:bg-slate-50") +
      (btn.dataset.entity !== "Group" ? " border-l border-slate-200" : "");
  });
  renderAll();
}

let sortKey = "vendorNo", sortAsc = true;
let chartDonut   = null;
let chartAvgDays = null;
let chartBucket  = null;
let chartCash    = null;

// ── date utils ─────────────────────────────────────────────────────────
function parseDMY(s) {
  if (!s) return null;
  const p = s.split('/');
  if (p.length !== 3) return null;
  let [d, m, y] = p.map(Number);
  if (y < 100) y += 2000;
  return new Date(y, m - 1, d);
}
function calcDaysOvr(ds) {
  const due = parseDMY(ds);
  if (!due) return null;
  const today = new Date(); today.setHours(0,0,0,0); due.setHours(0,0,0,0);
  return Math.floor((today - due) / 86400000);
}

// ── condition / action ─────────────────────────────────────────────────
// getCondition: days-only buckets (does NOT include Risk of Bad Debt).
function getCondition(days) {
  if (days === null) return "";
  if (days > 90)  return "90+ Days Overdue";
  if (days > 60)  return "61–90 Days Overdue";
  if (days > 30)  return "31–60 Days Overdue";
  if (days > 0)   return "≤30 Days Overdue";
  return "Current";
}
// getConditionForRow: authoritative — isBadDebt flag from Excel row color takes priority.
function getConditionForRow(r) {
  if (r.isBadDebt) return "Risk of Bad Debt";
  return getCondition(calcDaysOvr(r.dueDate));
}
// ── format utils ───────────────────────────────────────────────────────
const gbp  = n => n === 0 ? "—" : "£" + n.toLocaleString("en-GB",{minimumFractionDigits:2,maximumFractionDigits:2});
const gbpF = n => "£" + n.toLocaleString("en-GB",{minimumFractionDigits:2,maximumFractionDigits:2});
const gbpR = n => n === 0 ? "—" : "£" + Math.round(n).toLocaleString("en-GB");
const esc  = s => String(s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;");
const pct  = (p,t) => t ? (p/t*100).toFixed(1)+"%" : "0.0%";

function rowClass(r) {
  const cond = getConditionForRow(r);
  return {"Risk of Bad Debt":"row-bad-debt","90+ Days Overdue":"row-over-92",
          "61–90 Days Overdue":"row-63-92","31–60 Days Overdue":"row-32-61",
          "≤30 Days Overdue":"row-1-31","Current":"row-not-due"}[cond] || "";
}
function condBadge(r) {
  const cond = getConditionForRow(r);
  const cls = {"Current":"badge-not-due","≤30 Days Overdue":"badge-1-31",
               "31–60 Days Overdue":"badge-32-61","61–90 Days Overdue":"badge-63-92",
               "90+ Days Overdue":"badge-over-92","Risk of Bad Debt":"badge-bad-debt"}[cond] || "";
  return cond ? `<span class="px-2 py-0.5 rounded-full text-xs font-semibold ${cls}">${esc(cond)}</span>` : "";
}
// ── Chaser notes ────────────────────────────────────────────────────────
const _NOTES_LS = "pgChaserNotes_v1";

function _lsNotes() {
  try { return JSON.parse(localStorage.getItem(_NOTES_LS)) || {}; } catch { return {}; }
}

function saveNote(docNo, text) {
  const n = _lsNotes();
  n[docNo] = { note: text, updatedAt: new Date().toISOString() };
  localStorage.setItem(_NOTES_LS, JSON.stringify(n));
}

function getNoteObj(r) {
  const ls = _lsNotes()[r.docNo];
  if (ls) return ls;
  const built = (DATA.chaserNotes || {})[r.docNo];
  if (built) return built;
  if (r.xlComment) return { note: r.xlComment, updatedAt: null };
  return null;
}


function notesCell(r) {
  const obj = getNoteObj(r);
  const text = obj ? obj.note : "";
  const ts = obj && obj.updatedAt
    ? ` <span class="text-slate-300 text-xs ml-1">${new Date(obj.updatedAt).toLocaleDateString("en-GB",{day:"2-digit",month:"2-digit",year:"2-digit"})}</span>`
    : "";
  const display = text.length > 50 ? text.slice(0, 50) + "…" : text;
  const inner = display
    ? `<span class="text-slate-700">${esc(display)}</span>`
    : `<span class="text-slate-300 italic">Add note…</span>`;
  return `<span class="cursor-pointer hover:underline" title="${esc(text)}" onclick="editNote(this,'${esc(r.docNo)}')">${inner}</span>${ts}`;
}

function editNote(el, docNo) {
  const cur = _lsNotes()[docNo]?.note
    || (DATA.chaserNotes || {})[docNo]?.note
    || D.receivables.find(r => r.docNo === docNo)?.xlComment
    || "";
  const inp = document.createElement("input");
  inp.type = "text"; inp.value = cur;
  inp.className = "border border-blue-400 rounded px-1 py-0.5 text-xs w-52 outline-none";
  const cell = el.closest("td");
  cell.innerHTML = "";
  cell.appendChild(inp);
  inp.focus(); inp.select();
  function commit() {
    saveNote(docNo, inp.value.trim());
    renderAR(_currentData());
    renderChaserLog();
  }
  inp.addEventListener("blur", commit);
  inp.addEventListener("keydown", e => {
    if (e.key === "Enter") { e.preventDefault(); inp.blur(); }
    if (e.key === "Escape") { e.preventDefault(); renderAR(_currentData()); renderChaserLog(); }
  });
}

function editNoteLog(el, docNo) { editNote(el, docNo); }

// ── Customer-level notes (per vendor name, separate from invoice notes) ──
const _CUST_NOTES_KEY = "pgCustNotes_v1";
function _custNotesStore() {
  try { return JSON.parse(localStorage.getItem(_CUST_NOTES_KEY) || "{}"); } catch { return {}; }
}
function _getCustNote(vendor) {
  // Priority: localStorage override → Excel xlCustNote → empty
  const local = _custNotesStore()[vendor];
  if (local !== undefined && local !== "") return local;
  const r = D.receivables.find(x => x.vendorName === vendor && (x.xlCustNote || "").trim());
  return r ? r.xlCustNote.trim() : "";
}
function _saveCustNote(vendor, text) {
  const s = _custNotesStore();
  if (text.trim()) s[vendor] = text; else delete s[vendor];
  try { localStorage.setItem(_CUST_NOTES_KEY, JSON.stringify(s)); } catch {}
}

function editNoteDrill(el, docNo) {
  const cur = _lsNotes()[docNo]?.note
    || (DATA.chaserNotes || {})[docNo]?.note
    || D.receivables.find(r => r.docNo === docNo)?.xlComment
    || "";
  const ta = document.createElement("textarea");
  ta.value = cur; ta.rows = 3;
  ta.className = "border border-blue-400 rounded px-2 py-1 text-xs w-full outline-none resize-none focus:ring-1 focus:ring-blue-400";
  const cell = el.closest("td");
  cell.innerHTML = "";
  cell.appendChild(ta);
  ta.focus(); ta.select();
  function _render(val) {
    cell.innerHTML = val
      ? `<span class="cursor-pointer hover:underline text-slate-700" onclick="editNoteDrill(this,'${esc(docNo)}')">${esc(val)}</span>`
      : `<span class="cursor-pointer hover:underline text-slate-300 italic" onclick="editNoteDrill(this,'${esc(docNo)}')">Click to add note…</span>`;
  }
  function commit() {
    const val = ta.value.trim();
    saveNote(docNo, val);
    _render(val);
    renderChaserLog();
  }
  ta.addEventListener("blur", commit);
  ta.addEventListener("keydown", e => {
    if (e.key === "Enter" && !e.shiftKey) { e.preventDefault(); ta.blur(); }
    if (e.key === "Escape") { e.preventDefault(); _render(cur); }
  });
}

function _currentData() {
  if (filtered.length) return filtered;
  return D.receivables
    .map(r => ({...r, _daysOvr: calcDaysOvr(r.dueDate) ?? -9999}))
    .sort((a, b) => a.vendorName.localeCompare(b.vendorName));
}

function exportNotes() {
  const data = _lsNotes();
  const blob = new Blob([JSON.stringify(data, null, 2)], {type: "application/json"});
  const a = document.createElement("a"); a.href = URL.createObjectURL(blob);
  a.download = "chaser-notes.json"; a.click();
}
function daysCell(days) {
  if (days === null) return '<span class="text-slate-400">—</span>';
  if (days <= 0)  return `<span class="dov-ok">Due in ${Math.abs(days)}d</span>`;
  if (days <= 30) return `<span class="dov-mild">${days}d</span>`;
  if (days <= 60) return `<span class="dov-warn">${days}d</span>`;
  if (days <= 90) return `<span class="dov-bad">${days}d</span>`;
  return `<span class="dov-danger">${days}d</span>`;
}

// ── file picker ────────────────────────────────────────────────────────
function buildFilePicker() {
  const sel = document.getElementById("file-picker");
  DATA.files.forEach((f, i) => {
    const o = document.createElement("option");
    o.value = f;
    o.textContent = f.replace(/^(BI-Weekly|Monthly) Cash Update\s*/i,"").replace(/\.xlsx$/i,"");
    if (i === 0) o.selected = true;
    sel.appendChild(o);
  });
  sel.addEventListener("change", e => {
    curFile = e.target.value;
    D = _viewForFile(curFile);
    filtered = [];
    sortKey = "vendorNo"; sortAsc = true;
    _updateEntityToggleVisibility();
    renderAll();
  });
  _updateEntityToggleVisibility();
}

function _updateEntityToggleVisibility() {
  const wrap = document.getElementById("entity-toggle-wrap");
  if (!wrap) return;
  const hasEntities = !!(DATA.fileData[curFile] && DATA.fileData[curFile].entities);
  wrap.classList.toggle("hidden", !hasEntities);
  wrap.classList.toggle("flex",   hasEntities);
}

// ── KPI cards ──────────────────────────────────────────────────────────
// Colors read directly from Excel "Aged Receivables" sheet Color column (R):
//   #00B050 = ≤30 days overdue (green)
//   #FFC000 = 31–60 days overdue (amber)
//   #FF0000 = >60 days overdue (red)
//   #7030A0 = potential bad debt (purple — Office purple; also detects #800080)
//   #156082 = header / total (theme accent1)
//   #0F9ED5 = cash (theme accent4)
// KPI card colours — taken directly from Excel cell fills (Aged Receivables sheet)
// Not Due has no Excel fill → slate-grey (#475569) used as neutral "safe" colour
// Canonical Excel bucket colours — applied to KPI cards, donut, table, badges
const XL = {
  header:  { bg:"#1a2744", fg:"#ffffff" },
  notDue:  { bg:"#ffffff", fg:"#1a2744" },  // white, no border
  d1_31:   { bg:"#00B050", fg:"#ffffff" },  // Excel: FF00B050  green
  d32_61:  { bg:"#FF8C00", fg:"#ffffff" },  // orange, white text
  d63_92:  { bg:"#FF0000", fg:"#ffffff" },  // Excel: FFFF0000  red
  over92:  { bg:"#C00000", fg:"#ffffff" },  // dark red (92+)
  badDebt: { bg:"#7030A0", fg:"#ffffff" },  // Excel: FF7030A0  purple
  cash:    { bg:"#0369a1", fg:"#ffffff" },
};

function renderKPI() {
  const t = D.totals;
  const badDebt = D.receivables
    .filter(r => r.isBadDebt)
    .reduce((s,r) => s + r.balance, 0);

  // Cash in Bank lives in its own banner (cash-banner), not in the KPI row.
  // Risk of Bad Debt is a subset of 92+ Days — shown as a standalone card with a
  // clear disclaimer so management can see it without mistaking it for extra exposure.
  const cards = [
    {label:"Total Balance",    val:t.balance, sub:D.receivables.length+" invoices", ...XL.header, always:true},
    {label:"Current",          val:t.notDue,  ...XL.notDue},
    {label:"1–30 Days",        val:t.d1_31,   ...XL.d1_31},
    {label:"31–60 Days",       val:t.d32_61,  ...XL.d32_61, always:true},
    {label:"61–90 Days",       val:t.d63_92,  ...XL.d63_92, always:true},
    {label:"90+ Days",         val:t.over92,  ...XL.over92},
    {label:"Risk of Bad Debt", val:badDebt,   ...XL.badDebt, always:true,
     note:"Included within 90+ Days above"},
  ];

  document.getElementById("kpi-row").innerHTML = cards.map(c => {
    // isColored: white-text cards (coloured bg). !isColored: dark-text "Not Due" white card.
    const isColored = c.fg === "#ffffff";
    const valCol  = isColored ? "#ffffff"                : "#1a2744";
    const lblCol  = isColored ? "rgba(255,255,255,0.62)" : "rgba(26,39,68,0.48)";
    const subCol  = isColored ? "rgba(255,255,255,0.58)" : "rgba(26,39,68,0.44)";
    const noteCol = isColored ? "rgba(255,255,255,0.44)" : "rgba(26,39,68,0.34)";
    const divCol  = isColored ? "rgba(255,255,255,0.14)" : "rgba(26,39,68,0.08)";
    // Elevation — coloured: inner top highlight + two-layer drop shadow.
    // White card: no border, just a soft ambient shadow.
    const cardStyle = isColored
      ? `box-shadow:inset 0 1px 0 rgba(255,255,255,0.18),0 1px 3px rgba(0,0,0,0.16),0 5px 12px rgba(0,0,0,0.14),0 14px 30px rgba(0,0,0,0.08);`
      : `box-shadow:0 1px 3px rgba(0,0,0,0.05),0 4px 12px rgba(0,0,0,0.07),0 14px 28px rgba(0,0,0,0.04);`;
    const valHtml = `<p style="font-size:clamp(18px,1.55vw,24px);font-weight:600;margin:10px 0 0;line-height:1;color:${valCol};white-space:nowrap;font-variant-numeric:tabular-nums;letter-spacing:-0.02em">${gbpF(c.val)}</p>`;
    const sub  = c.sub || pct(c.val, t.balance) + " of total";
    const note = c.note
      ? `<p style="font-size:10px;font-weight:400;margin-top:4px;font-style:italic;color:${noteCol};line-height:1.4">${c.note}</p>` : "";
    return `<div class="kpi-card" style="flex:1;background:${c.bg};border-radius:12px;padding:22px 18px 18px;${cardStyle}min-width:0;overflow:hidden;text-align:left">
      <p style="font-size:10px;font-weight:500;text-transform:uppercase;letter-spacing:0.11em;color:${lblCol};line-height:1;margin:0">${c.label}</p>
      ${valHtml}
      <div style="border-top:1px solid ${divCol};margin:12px 0 9px"></div>
      <p style="font-size:11px;font-weight:400;margin:0;color:${subCol};line-height:1.3">${sub}</p>
      ${note}
    </div>`;
  }).join("");
}

// ── donut chart ────────────────────────────────────────────────────────
function renderDonut() {
  if (chartDonut) { chartDonut.destroy(); chartDonut = null; }
  const t = D.totals;
  const badDebt = D.receivables.filter(r => r.isBadDebt).reduce((s,r) => s + r.balance, 0);

  // 5 true ring segments — 92+ Days uses its FULL value (no splitting).
  // Risk of Bad Debt is shown as a bracket annotation outside the ring.
  // 32-61 Days is excluded when £0 (filtered below).
  const segs = [
    { label: "Current",         val: t.notDue,  color: "#D9D9D9" },  // Excel light grey
    { label: "1\u201330 Days",  val: t.d1_31,   color: "#00B050" },  // Excel green
    { label: "31\u201360 Days", val: t.d32_61,  color: "#FFC000" },  // Excel amber — removed when £0
    { label: "61\u201390 Days", val: t.d63_92,  color: "#FF0000" },  // Excel red
    { label: "90+ Days",        val: t.over92,  color: "#C00000", is92: true },  // Excel dark red
  ].filter(s => s.val > 0.01);

  const total = segs.reduce((s, sg) => s + sg.val, 0);

  document.getElementById("donut-title").textContent =
    "Receivables Distribution \u2014 " + D.reportDate;

  const cx = 120, cy = 120, R = 108, r = 64;
  const GAP = 0.018;

  function pt(a, rad) {
    return [cx + rad * Math.cos(a - Math.PI / 2), cy + rad * Math.sin(a - Math.PI / 2)];
  }
  function arcPath(sa, ea, outerR, innerR) {
    const [x1,y1] = pt(sa, outerR), [x2,y2] = pt(ea, outerR);
    const [x3,y3] = pt(ea, innerR), [x4,y4] = pt(sa, innerR);
    const lg = ea - sa > Math.PI ? 1 : 0, f = v => v.toFixed(3);
    return `M${f(x1)},${f(y1)} A${outerR},${outerR} 0 ${lg} 1 ${f(x2)},${f(y2)}` +
           ` L${f(x3)},${f(y3)} A${innerR},${innerR} 0 ${lg} 0 ${f(x4)},${f(y4)} Z`;
  }

  let angle = 0;
  const built = segs.map((seg, i) => {
    const sweep = (seg.val / total) * 2 * Math.PI;
    const sa = angle + GAP / 2, ea = angle + sweep - GAP / 2;
    angle += sweep;
    return { ...seg, sa, ea, i, path: arcPath(sa, ea, R, r) };
  });

  // ── Purple outer band for Risk of Bad Debt ───────────────────────────
  // Filled arc band sitting just outside the ring, spanning the angular
  // proportion of 92+ Days that represents Risk of Bad Debt.
  const OVL = 9;    // band thickness in SVG px
  let annotHTML = "";
  const seg92 = built.find(s => s.is92);
  if (seg92 && badDebt > 0.01 && t.over92 > 0.01) {
    const bdFrac = badDebt / t.over92;
    const osa    = seg92.sa;
    const oea    = seg92.sa + bdFrac * (seg92.ea - seg92.sa);
    annotHTML =
      `<path d="${arcPath(osa, oea, R + OVL, R)}" fill="#7030A0" stroke="white" stroke-width="1"` +
      ` style="cursor:pointer"` +
      ` onmouseenter="donutBdHover(event)" onmouseleave="donutLeave()" onmousemove="donutMove(event)"/>`;
  }

  const cg = _cg();
  const isDk = _isDark();
  const textFill = isDk ? "#e2e8f0" : "#0f172a";
  const subFill  = isDk ? "#94a3b8" : "#64748b";

  const pathsHTML = built.map(s =>
    `<path data-i="${s.i}" d="${s.path}" fill="${s.color}" stroke="white" stroke-width="1.5"` +
    ` style="cursor:pointer;transition:opacity .12s"` +
    ` onmouseenter="donutHover(event,${s.i})" onmouseleave="donutLeave()" onmousemove="donutMove(event)"/>`
  ).join("\n    ");

  const tipHTML =
    `<div id="donut-tip" style="display:none;position:absolute;background:rgba(15,23,42,0.92);` +
    `color:#f1f5f9;padding:6px 10px;border-radius:6px;font-size:11px;pointer-events:none;` +
    `white-space:nowrap;z-index:50;box-shadow:0 2px 8px rgba(0,0,0,0.35)"></div>`;

  // overflow:visible allows the bracket (SPINE=R+16=124 from centre) to paint
  // just outside the 240×240 viewBox without any coordinate-system surgery
  const svgHTML =
    `<svg id="donut-svg" viewBox="0 0 240 240"` +
    ` style="width:100%;display:block;overflow:visible">` +
    `\n    ${pathsHTML}` +
    (annotHTML ? `\n    ${annotHTML}` : "") +
    `\n    <text x="${cx}" y="${cy-9}" text-anchor="middle" font-family="system-ui,sans-serif"` +
    ` font-size="10.5" fill="${subFill}" font-weight="500">Total</text>` +
    `\n    <text x="${cx}" y="${cy+9}" text-anchor="middle" font-family="system-ui,sans-serif"` +
    ` font-size="12.5" fill="${textFill}" font-weight="700">${gbpF(total)}</text>` +
    `\n  </svg>`;

  const legendHTML = _donutLegend(built, total, badDebt, cg.legend);

  const el = document.getElementById("donutChart");
  el.innerHTML = tipHTML + svgHTML + `<div style="margin-top:10px">${legendHTML}</div>`;

  window._donutSegs    = built;
  window._donutTotal   = total;
  window._donutBadDebt = badDebt;
  window._donut92val   = t.over92;
}

function _donutLegend(segs, total, bd, legendCol) {
  const col = legendCol || "#475569";
  function shortLabel(label) {
    return label
      .replace("1\u201331 Days",  "1\u201331d")
      .replace("32\u201361 Days", "32\u201361d")
      .replace("63\u201392 Days", "63\u201392d")
      .replace("90+ Days",        "90+d");
  }
  function compactGbp(v) {
    if (v >= 1e6) return "\u00a3" + (v / 1e6).toFixed(2) + "M";
    if (v >= 1e3) return "\u00a3" + (v / 1e3).toFixed(1) + "K";
    return "\u00a3" + v.toFixed(0);
  }
  let html = '<div style="display:flex;flex-wrap:wrap;gap:5px 12px;justify-content:center;font-size:11px;line-height:1.6">';
  for (const s of segs) {
    const pct = (s.val / total * 100).toFixed(1);
    html +=
      `<div style="display:flex;align-items:center;gap:4px;color:${col}">` +
      `<span style="width:10px;height:10px;border-radius:2px;background:${s.color};` +
      `flex-shrink:0;display:inline-block"></span>` +
      `<span>${shortLabel(s.label)}&nbsp;&nbsp;${compactGbp(s.val)}&nbsp;(${pct}%)</span></div>`;
    if (s.is92 && bd > 0.01) {
      const bdPct = (bd / total * 100).toFixed(1);
      html +=
        `<div style="display:flex;align-items:center;gap:4px;color:#7030A0;flex-basis:100%;justify-content:center">` +
        `<span style="width:10px;height:10px;border-radius:2px;background:#7030A0;` +
        `flex-shrink:0;display:inline-block"></span>` +
        `<span>&nbsp;&nbsp;&#x2514;&nbsp;Risk of Bad Debt` +
        `&nbsp;&nbsp;${compactGbp(bd)}&nbsp;(${bdPct}%)</span></div>`;
    }
  }
  html += '</div>';
  return html;
}

function donutHover(evt, i) {
  const seg = (window._donutSegs || [])[i];
  if (!seg) return;
  const pct = (seg.val / (window._donutTotal || 1) * 100).toFixed(1);
  const tip = document.getElementById("donut-tip");
  if (!tip) return;
  tip.innerHTML = `<strong>${seg.label}</strong><br>${gbpF(seg.val)}&nbsp;&nbsp;(${pct}% of total)`;
  tip.style.display = "block";
  donutMove(evt);
  document.querySelectorAll("#donutChart path").forEach((p, j) => {
    p.style.opacity = j === i ? "1" : "0.38";
  });
}

function donutBdHover(evt) {
  const bd    = window._donutBadDebt || 0;
  const total = window._donutTotal   || 1;
  const v92   = window._donut92val   || 1;
  const tip   = document.getElementById("donut-tip");
  if (!tip) return;
  const bdK   = bd >= 1e6
    ? "\u00a3" + (bd / 1e6).toFixed(2) + "M"
    : "\u00a3" + (bd / 1e3).toFixed(1) + "K";
  const pct   = (bd / total * 100).toFixed(1);
  const pct92 = (bd / v92   * 100).toFixed(1);
  tip.innerHTML =
    `Risk of Bad Debt &mdash; ${bdK} (${pct}% of total &middot; ${pct92}% of 90+ Days)`;
  tip.style.display = "block";
  donutMove(evt);
}

function donutLeave() {
  const tip = document.getElementById("donut-tip");
  if (tip) tip.style.display = "none";
  document.querySelectorAll("#donutChart path").forEach(p => { p.style.opacity = "1"; });
}

function donutMove(evt) {
  const tip = document.getElementById("donut-tip");
  if (!tip || tip.style.display === "none") return;
  const box = document.getElementById("donutChart").getBoundingClientRect();
  let x = evt.clientX - box.left + 14;
  let y = evt.clientY - box.top  - 12;
  if (x + 230 > box.width) x = evt.clientX - box.left - 210;
  if (y < 0) y = 4;
  tip.style.left = x + "px";
  tip.style.top  = y + "px";
}

// ── avg days overdue trend (all files, rendered once) ─────────────────
function renderAvgDaysChart() {
  if (chartAvgDays) { chartAvgDays.destroy(); chartAvgDays = null; }
  const td = DATA.trendData;
  if (!td || !td.length) return;

  const labels = td.map(p => p.date);
  const values = td.map(p => p.avgDaysOvr);

  // Subtitle: current vs previous
  if (td.length >= 2) {
    const cur  = values[values.length - 1];
    const prv  = values[values.length - 2];
    const diff = +(cur - prv).toFixed(1);
    const arrow = diff > 0 ? "\u25b2" : diff < 0 ? "\u25bc" : "\u2014";
    const col   = diff > 0 ? "color:#dc2626" : diff < 0 ? "color:#16a34a" : "color:#64748b";
    document.getElementById("avg-days-subtitle").innerHTML =
      `Current: <strong>${cur.toFixed(1)} days</strong> &nbsp;` +
      `<span style="${col};font-weight:600">${arrow} ${Math.abs(diff).toFixed(1)}d vs prev week</span>`;
  }

  // Segment coloring: red >60, amber 30-60, green <30
  function segColor(ctx) {
    const y = ctx.p1.parsed.y;
    if (y > 60) return "#FF0000";
    if (y > 30) return "#FFC000";
    return "#00B050";
  }

  const refLines = [
    { y: 30, color: "#00B050", label: "30d" },
    { y: 60, color: "#FFC000", label: "60d" },
    { y: 90, color: "#FF0000", label: "90d" },
  ];

  chartAvgDays = new Chart(document.getElementById("avgDaysChart"), {
    type: "line",
    data: {
      labels,
      datasets: [
        {
          label: "Avg Days Overdue",
          data:  values,
          borderColor:     "#156082",
          backgroundColor: "transparent",
          fill:    false,
          tension: 0.3,
          pointRadius: 3,
          pointHoverRadius: 6,
          borderWidth: 2.5,
          segment: { borderColor: ctx => segColor(ctx) },
          pointBackgroundColor: values.map(v => v > 60 ? "#FF0000" : v > 30 ? "#FFC000" : "#00B050"),
        },
        ...refLines.map(r => ({
          label: r.label + " ref",
          data:  labels.map(() => r.y),
          borderColor: r.color,
          borderWidth: 1,
          borderDash:  [5, 4],
          pointRadius: 0,
          fill: false,
          tension: 0,
        }))
      ]
    },
    options: {
      responsive: true,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: {
          display: true,
          position: "bottom",
          labels: {
            font: { size: 10 }, padding: 10, color: _cg().legend,
            usePointStyle: true,
            filter: item => item.datasetIndex === 0,
          }
        },
        tooltip: {
          callbacks: {
            label: ctx => ctx.datasetIndex === 0
              ? `  Avg days overdue: ${ctx.parsed.y.toFixed(1)}d`
              : null,
            filter: item => item.datasetIndex === 0,
          }
        }
      },
      scales: {
        y: {
          min: 0,
          ticks: { callback: v => v + "d", font: { size: 10 }, color: _cg().tick },
          grid:  { color: _cg().grid }
        },
        x: {
          ticks: { font: { size: 9 }, maxRotation: 45, maxTicksLimit: 17, color: _cg().tick },
          grid:  { display: false }
        }
      }
    }
  });
}

// ── bucket drill-down (all files, re-rendered on picker change) ────────
function renderBucketDrilldown(key) {
  if (chartBucket) { chartBucket.destroy(); chartBucket = null; }
  const td = DATA.trendData;
  if (!td || !td.length) return;

  const keyMap = {
    "d1_31_abs":   { label: "1\u201330 Days",   pctKey: "d1_31",   color: "#FFC000" },
    "d32_61_abs":  { label: "31\u201360 Days",  pctKey: "d32_61",  color: "#FFC000" },
    "d63_92_abs":  { label: "61\u201390 Days",  pctKey: "d63_92",  color: "#FF0000" },
    "over92_abs":  { label: "90+ Days",         pctKey: "over92",  color: "#C00000" },
    "badDebt_abs": { label: "Risk of Bad Debt", pctKey: "badDebt", color: "#7030A0" },
  };
  const meta = keyMap[key] || { label: key, pctKey: key, color: "#94a3b8" };

  // Keep picker in sync
  const picker = document.getElementById("bucket-picker");
  if (picker) picker.value = key;

  document.getElementById("bucket-title").textContent =
    `Bucket Drill-Down \u2014 ${meta.label} (% of total balance)`;

  const labels = td.map(p => p.date);
  const values = td.map(p => p[meta.pctKey] || 0);

  // Subtitle: current % vs previous week
  if (td.length >= 2) {
    const cur  = values[values.length - 1];
    const prv  = values[values.length - 2];
    const diff = +(cur - prv).toFixed(1);
    const arrow = diff > 0 ? "\u25b2" : diff < 0 ? "\u25bc" : "\u2014";
    const col   = diff > 0 ? "color:#dc2626" : diff < 0 ? "color:#16a34a" : "color:#64748b";
    document.getElementById("bucket-subtitle").innerHTML =
      `Current: <strong>${cur.toFixed(1)}% of total</strong> &nbsp;` +
      `<span style="${col};font-weight:600">${arrow} ${Math.abs(diff).toFixed(1)}pp vs prev week</span>`;
  }

  chartBucket = new Chart(document.getElementById("bucketChart"), {
    type: "line",
    data: {
      labels,
      datasets: [{
        label:           meta.label,
        data:            values,
        borderColor:     meta.color,
        backgroundColor: meta.color + "18",
        fill:            true,
        tension:         0.35,
        pointRadius:     3,
        pointHoverRadius:6,
        borderWidth:     2,
      }]
    },
    options: {
      responsive: true,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => `  ${meta.label}: ${ctx.parsed.y.toFixed(1)}% of total`
          }
        }
      },
      scales: {
        y: {
          min: 0,
          ticks: { callback: v => v + "%", font: { size: 10 }, color: _cg().tick },
          grid:  { color: _cg().grid }
        },
        x: {
          ticks: { font: { size: 9 }, maxRotation: 45, maxTicksLimit: 17, color: _cg().tick },
          grid:  { display: false }
        }
      }
    }
  });
}

// ── cash banner (full-width strip between KPI row and charts) ───────────
function renderCashCard() {
  const el = document.getElementById("cash-banner");
  if (!DATA.bankCash) {
    el.innerHTML = "";   // hide entirely if bank file unavailable
    return;
  }
  const b = DATA.bankCash;
  const bg     = _isDark() ? "#1a1d27" : "#ffffff";
  const border = _isDark() ? "#2d3148"  : "#e2e8f0";
  const valCol = _isDark() ? "#e2e8f0"  : "#1a2744";
  const subCol = _isDark() ? "#94a3b8"  : "#64748b";
  const accent = "#0369a1";
  el.innerHTML = `
    <div style="background:${bg};border:1px solid ${border};border-left:4px solid ${accent};border-radius:10px;padding:16px 20px;display:flex;align-items:center;gap:20px">
      <svg style="width:32px;height:32px;flex-shrink:0;color:${accent};opacity:0.8"
           fill="none" stroke="currentColor" viewBox="0 0 24 24">
        <path stroke-linecap="round" stroke-linejoin="round" stroke-width="1.5"
          d="M3 10h18M7 15h1m4 0h1m-7 4h12a3 3 0 003-3V8a3 3 0 00-3-3H6a3 3 0 00-3 3v8a3 3 0 003 3z"/>
      </svg>
      <div>
        <p style="font-size:10px;font-weight:600;text-transform:uppercase;letter-spacing:0.07em;color:${subCol}">Cash in Bank</p>
        <p style="font-size:26px;font-weight:700;line-height:1.15;color:${valCol};margin-top:4px">${gbpF(b.latest)}</p>
        <p style="font-size:11px;margin-top:4px;color:${subCol}">As at ${esc(b.latestDate)} &middot; all accounts (GBP) &middot; Treasury metric &mdash; separate from receivables</p>
      </div>
    </div>`;
}

// ── cash line chart (2026 daily, rendered once) ─────────────────────────
function renderCashLine() {
  if (chartCash) { chartCash.destroy(); chartCash = null; }
  if (!DATA.bankCash || !DATA.bankCash.series2026.length) return;

  const b      = DATA.bankCash;
  const series = b.series2026;  // already chronological

  // Populate header value badge
  const hdrEl = document.getElementById("cash-hdr-val");
  if (hdrEl) hdrEl.innerHTML = `
    <p class="text-xs text-white" style="opacity:0.65">As at ${esc(b.latestDate)}</p>
    <p class="text-xl font-bold text-white">${gbpF(b.latest)}</p>`;

  chartCash = new Chart(document.getElementById("cashChart"), {
    type: "line",
    data: {
      labels: series.map(p => p.date),
      datasets: [{
        label: "Total Cash",
        data:  series.map(p => p.total),
        borderColor:     "#156082",
        backgroundColor: "rgba(21,96,130,0.07)",
        fill:     true,
        tension:  0.35,
        pointRadius:      2,
        pointHoverRadius: 5,
        borderWidth: 2,
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: ctx => ctx[0].label,
            label: ctx => `  Total: ${gbpF(ctx.parsed.y)}`
          }
        }
      },
      scales: {
        y: {
          ticks: { callback: v => "£" + (v / 1e6).toFixed(2) + "M", font:{size:10}, color: _cg().tick },
          grid:  { color: _cg().grid }
        },
        x: {
          ticks: { font:{size:9}, maxRotation:45, maxTicksLimit:20, color: _cg().tick },
          grid:  { display: false }
        }
      }
    }
  });
}

// ── period comparison ──────────────────────────────────────────────────
function renderComparison() {
  const c = D.totals, p = D.totalsPrev;
  function deltaHtml(cur, prv, hib) {
    const d = cur - prv;
    if (Math.abs(d) < 0.01) return '<span class="text-slate-400">—</span>';
    const good = hib ? d > 0 : d < 0;
    const col  = good ? "text-green-600" : "text-red-600";
    return `<span class="${col} font-semibold text-xs whitespace-nowrap">${d>0?"▲":"▼"} ${gbpF(Math.abs(d))}</span>`;
  }
  const rows = [
    ["Balance",    c.balance, p.balance, false],
    ["Current",    c.notDue,  p.notDue,  true ],
    ["1–30 Days",  c.d1_31,   p.d1_31,   false],
    ["31–60 Days", c.d32_61,  p.d32_61,  false],
    ["61–90 Days", c.d63_92,  p.d63_92,  false],
    ["90+ Days",   c.over92,  p.over92,  false],
  ];
  let html = `
    <div class="grid grid-cols-4 gap-1 text-xs font-semibold text-slate-500 border-b pb-2 mb-1">
      <span>Metric</span>
      <span class="text-right">${D.reportDate}</span>
      <span class="text-right text-slate-400">${D.prevDate}</span>
      <span class="text-right">Δ</span>
    </div>`;
  html += rows.map(([l,cur,prv,hib]) => `
    <div class="grid grid-cols-4 gap-1 items-center py-1 border-b border-slate-50">
      <span class="text-slate-500 text-xs">${l}</span>
      <span class="text-right font-semibold text-slate-800 text-xs">${gbpF(cur)}</span>
      <span class="text-right text-slate-400 text-xs">${gbpF(prv)}</span>
      <span class="text-right">${deltaHtml(cur,prv,hib)}</span>
    </div>`).join("");
  html += `
    <div class="mt-3 bg-blue-50 rounded-lg px-3 py-2">
      <p class="text-xs text-blue-700 font-semibold">Collections received</p>
      <p class="text-lg font-bold text-blue-800 mt-0.5">${gbpF(D.collectionsTotal)}</p>
    </div>
    <div class="text-xs text-slate-400 mt-2">
      USD/GBP ${D.exchangeRate.toFixed(6)} <span class="ml-1">(as at ${D.fxDate})</span>
    </div>`;
  document.getElementById("comparison-table").innerHTML = html;
}

// ── AR table ───────────────────────────────────────────────────────────
function _bucketBorder(r) {
  if (r.isBadDebt) return "#7030A0";
  if (r.over92  > 0.01) return "#C00000";
  if (r.d63_92  > 0.01) return "#FF0000";
  if (r.d32_61  > 0.01) return "#FF8C00";
  if (r.d1_31   > 0.01) return "#00B050";
  return "#00B050";
}

function renderAR(data) {
  document.getElementById("arBody").innerHTML = data.map((r, idx) => {
    const days   = calcDaysOvr(r.dueDate);
    const cond   = getConditionForRow(r);
    const altBg  = idx % 2 === 1 ? "#f8fafc" : "#ffffff";
    const bdrCol = _bucketBorder(r);
    return `
    <tr class="${rowClass(r)} border-b border-slate-100" style="background:${altBg}">
      <td class="px-3 py-1.5 font-medium sticky-name" style="min-width:180px;background:${altBg};border-left:3px solid ${bdrCol}"><span class="cursor-pointer text-slate-800 hover:text-blue-600 hover:underline" data-vendor="${esc(r.vendorName)}" onclick="openDrill(this)">${esc(r.vendorName)}</span></td>
      <td class="px-3 py-1.5 text-slate-600 whitespace-nowrap">${esc(r.postingDate)}</td>
      <td class="px-3 py-1.5 font-mono text-slate-600 whitespace-nowrap">${esc(r.docNo)}</td>
      <td class="px-3 py-1.5 text-slate-600 whitespace-nowrap">${esc(r.dueDate)}</td>
      <td class="px-3 py-1.5 text-right">${daysCell(days)}</td>
      <td class="px-3 py-1.5 text-right text-slate-700">${gbpF(r.origAmt)}</td>
      <td class="px-3 py-1.5 text-right font-semibold text-slate-800">${gbpF(r.balance)}</td>
      <td class="px-2 py-1.5 text-right text-green-700">${gbpR(r.notDue)}</td>
      <td class="px-2 py-1.5 text-right text-yellow-700">${gbpR(r.d1_31)}</td>
      <td class="px-2 py-1.5 text-right text-orange-600">${gbpR(r.d32_61)}</td>
      <td class="px-2 py-1.5 text-right text-red-600">${gbpR(r.d63_92)}</td>
      <td class="px-2 py-1.5 text-right font-semibold" style="color:${r.isBadDebt ? '#4b1d7a' : '#be123c'}">${gbpR(r.over92)}</td>
      <td class="px-3 py-1.5 whitespace-nowrap">${condBadge(r)}</td>
    </tr>`;
  }).join("");
  document.getElementById("rowCount").textContent = `${data.length} of ${D.receivables.length} rows`;
}

function renderARFoot() {
  const t = D.totals;
  document.getElementById("arFoot").innerHTML = `
    <tr>
      <td class="px-3 py-2.5 sticky-name" style="background:#1a2744;color:#ffffff;font-weight:600">Total — ${D.reportDate}</td>
      <td colspan="5"></td>
      <td class="px-3 py-2.5 text-right">${gbpF(t.balance)}</td>
      <td class="px-2 py-2.5 text-right">${gbpR(t.notDue)}</td>
      <td class="px-2 py-2.5 text-right">${gbpR(t.d1_31)}</td>
      <td class="px-2 py-2.5 text-right">${gbpR(t.d32_61)}</td>
      <td class="px-2 py-2.5 text-right">${gbpR(t.d63_92)}</td>
      <td class="px-2 py-2.5 text-right">${gbpR(t.over92)}</td>
      <td></td>
    </tr>`;
}

function applyFilters() {
  const q = document.getElementById("search").value.toLowerCase();
  const fC = document.getElementById("fCond").value;
  filtered = D.receivables.map(r => ({...r, _daysOvr: calcDaysOvr(r.dueDate) ?? -9999}))
    .filter(r => {
      const cond = getConditionForRow(r);
      const mq = !q || r.vendorName.toLowerCase().includes(q) ||
                       r.vendorNo.toLowerCase().includes(q)   ||
                       r.docNo.toLowerCase().includes(q);
      return mq && (!fC || cond === fC);
    })
    .sort((a, b) => a.vendorName.localeCompare(b.vendorName));
  renderAR(filtered);
}

function clearFilters() {
  document.getElementById("search").value = "";
  document.getElementById("fCond").value  = "";
  applyFilters();
}

function sortBy(key) {
  if (sortKey === key) sortAsc = !sortAsc; else { sortKey = key; sortAsc = true; }
  const data = filtered.length ? filtered : D.receivables.map(r => ({...r, _daysOvr: calcDaysOvr(r.dueDate) ?? -9999}));
  data.sort((a, b) => {
    const va = a[key] ?? "", vb = b[key] ?? "";
    if (typeof va === "number") return sortAsc ? va - vb : vb - va;
    return sortAsc ? String(va).localeCompare(String(vb)) : String(vb).localeCompare(String(va));
  });
  renderAR(data);
}

// ── collections ────────────────────────────────────────────────────────
function renderCollections() {
  const rows = D.collections;
  document.getElementById("collBody").innerHTML = rows.map((c, i) => {
    const amtDisp = c.isUSD
      ? `${c.amt.toLocaleString("en-GB",{minimumFractionDigits:2})} <span class="text-slate-400 text-xs">(USD)</span>`
      : gbpF(c.amt);
    return `<tr class="${i%2===0?"bg-white":"bg-slate-50"} border-b border-slate-100">
      <td class="px-3 py-2 text-slate-600 whitespace-nowrap">${esc(c.date)}</td>
      <td class="px-3 py-2 font-medium text-slate-800">${esc(c.customer)}</td>
      <td class="px-3 py-2 text-right text-slate-600">${amtDisp}</td>
      <td class="px-3 py-2 text-right font-semibold text-blue-700">${gbpF(c.amtGBP)}</td>
    </tr>`;
  }).join("");
  document.getElementById("collFoot").innerHTML = `
    <tr>
      <td class="px-3 py-2.5" colspan="3">Total</td>
      <td class="px-3 py-2.5 text-right">${gbpF(D.collectionsTotal)}</td>
    </tr>`;
}

// ── To Be Invoiced (Monday.com) ────────────────────────────────────────
function renderTBI() {
  document.getElementById("monday-sync").textContent = DATA.mondaySyncDate;

  // Badge colors matching Monday.com exactly:
  //   Received   = #00c875 (green)   Not Chased = #fdab3d (orange)
  //   Chased     = #579bfc (blue)    blank      = grey dash
  //   Do we need ? = #c4c4c4 (grey)
  function mondayBadge(val) {
    if (!val) return '<span class="text-slate-400">—</span>';
    const styles = {
      "Received":     "color:#fff;background:#00c875",
      "Not Chased":   "color:#fff;background:#fdab3d",
      "Chased":       "color:#fff;background:#579bfc",
      "Do we need ?": "color:#555;background:#c4c4c4",
    };
    const s = styles[val] || "color:#555;background:#e2e8f0";
    return `<span class="px-2 py-0.5 rounded text-xs font-semibold whitespace-nowrap" style="${s}">${esc(val)}</span>`;
  }

  if (!DATA.mondayTBI || DATA.mondayTBI.length === 0) {
    document.getElementById("tbi-content").innerHTML =
      `<div class="p-8 text-center text-slate-500">
        <p class="text-lg font-semibold">No campaigns in Delivered Campaigns</p>
      </div>`;
    return;
  }

  function parseDV(str) {
    if (!str) return 0;
    return parseFloat(str.replace(/[£,]/g, "")) || 0;
  }
  function fmtDV(n) { return "£" + n.toLocaleString("en-GB"); }

  // Build rows grouped, with per-group subtotal rows
  const groupOrder = ["Campaigns Kicked Off", "Delivered Campaigns"];
  let grandTotal = 0;
  let rowIdx = 0;
  const bodyRows = groupOrder.map(grp => {
    const items = DATA.mondayTBI.filter(t => t.group === grp);
    if (!items.length) return "";
    const subtotal = items.reduce((s, t) => s + parseDV(t.dealValue), 0);
    grandTotal += subtotal;
    const header = `
    <tr class="bg-slate-200">
      <td colspan="5" class="px-3 py-1.5 text-xs font-semibold text-slate-600 uppercase tracking-wide">${esc(grp)}</td>
    </tr>`;
    const dataRows = items.map(t => {
      const bg = rowIdx++ % 2 === 0 ? "bg-white" : "bg-slate-50";
      const dv = t.dealValue
        ? `<span class="font-semibold text-slate-700">${esc(t.dealValue)}</span>`
        : '<span class="text-slate-400">—</span>';
      return `
    <tr class="${bg} border-b border-slate-100">
      <td class="px-3 py-2.5 font-medium text-slate-800" style="min-width:280px">${esc(t.name)}</td>
      <td class="px-3 py-2.5">${dv}</td>
      <td class="px-3 py-2.5">${mondayBadge(t.io)}</td>
      <td class="px-3 py-2.5">${mondayBadge(t.ncf)}</td>
      <td class="px-3 py-2.5">${mondayBadge(t.lineItems)}</td>
    </tr>`;
    }).join("");
    const subtotalRow = `
    <tr class="bg-slate-100 border-t border-slate-300 font-semibold">
      <td class="px-3 py-2 text-slate-600 text-xs">Subtotal</td>
      <td class="px-3 py-2 text-slate-700">${fmtDV(subtotal)}</td>
      <td colspan="3"></td>
    </tr>`;
    return header + dataRows + subtotalRow;
  }).join("");

  const totalRow = `
    <tr class="bg-slate-700 text-white font-semibold">
      <td class="px-3 py-2.5">Grand Total</td>
      <td class="px-3 py-2.5">${fmtDV(grandTotal)}</td>
      <td colspan="3"></td>
    </tr>`;
  const rows = bodyRows;

  document.getElementById("tbi-content").innerHTML = `
    <div class="tbl-wrap">
      <table class="w-full text-xs">
        <thead class="bg-slate-700 text-white text-left">
          <tr>
            <th class="px-3 py-2.5" style="min-width:280px">Campaign Name</th>
            <th class="px-3 py-2.5">Deal Value</th>
            <th class="px-3 py-2.5">Insertion Order</th>
            <th class="px-3 py-2.5">New Customer Form</th>
            <th class="px-3 py-2.5">Line Items</th>
          </tr>
        </thead>
        <tbody>${rows}${totalRow}</tbody>
      </table>
    </div>`;
}

// ── Monday.com live refresh (browser-side) ─────────────────────────────
const MONDAY_API_TOKEN = "__MONDAY_TOKEN__";

async function _mondayGraphQL(query) {
  if (!MONDAY_API_TOKEN) throw new Error("No Monday.com API token configured. Set MONDAY_TOKEN in build.py and rebuild.");
  const resp = await fetch("https://api.monday.com/v2", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "Authorization": MONDAY_API_TOKEN,
      "API-Version": "2024-04",
    },
    body: JSON.stringify({ query }),
  });
  if (!resp.ok) throw new Error(`Monday.com API HTTP ${resp.status}`);
  const json = await resp.json();
  if (json.errors && json.errors.length) throw new Error(json.errors.map(e => e.message).join("; "));
  return json.data;
}

async function refreshTBIFromMonday() {
  const btn   = document.getElementById("tbi-refresh-btn");
  const label = document.getElementById("tbi-refresh-label");
  const icon  = document.getElementById("tbi-refresh-icon");
  const errEl = document.getElementById("tbi-error");

  btn.disabled = true;
  label.textContent = "Refreshing…";
  icon.style.animation = "spin 0.8s linear infinite";
  errEl.classList.add("hidden");
  errEl.textContent = "";

  try {
    // Single query: billing board → status cols + Deal Value mirror column (lookup_mm127xxv)
    function fmtGbp(raw) {
      const n = parseFloat(String(raw || "").replace(/,/g, ""));
      return isNaN(n) ? "" : "£" + n.toLocaleString("en-GB", {maximumFractionDigits: 0});
    }
    function sumMirrorValue(displayValue) {
      if (!displayValue) return "";
      try {
        const total = displayValue.split(",").reduce((s, v) => s + (parseFloat(v.trim()) || 0), 0);
        return total ? fmtGbp(String(total)) : "";
      } catch(e) { return ""; }
    }

    const data1 = await _mondayGraphQL(`
      query {
        boards(ids: [5092118433]) {
          items_page(limit: 200) {
            items {
              id name
              group { id }
              column_values(ids: [
                "color_mm0v4cv1",
                "color_mm0vvyz3",
                "color_mm0vj2th",
                "lookup_mm127xxv"
              ]) {
                id text
                ... on MirrorValue { display_value }
              }
            }
          }
        }
      }
    `);

    const GROUP_MAP = {
      "group_mm0rffvp": "Campaigns Kicked Off",
      "group_mm0w5v1k": "Delivered Campaigns",
    };

    const billingItems = data1?.boards?.[0]?.items_page?.items ?? [];
    const tbi = [];

    for (const it of billingItems) {
      const groupLabel = GROUP_MAP[it.group?.id];
      if (!groupLabel) continue;
      const colText = {}, colDV = {};
      for (const cv of (it.column_values || [])) {
        colText[cv.id] = (cv.text || "").trim();
        colDV[cv.id]   = cv.display_value || "";
      }
      tbi.push({
        name:      it.name,
        dealValue: sumMirrorValue(colDV["lookup_mm127xxv"]),
        group:     groupLabel,
        io:        colText["color_mm0v4cv1"] || "",
        ncf:       colText["color_mm0vvyz3"] || "",
        lineItems: colText["color_mm0vj2th"] || "",
      });
    }

    DATA.mondayTBI = tbi;
    const now = new Date();
    const pad = n => String(n).padStart(2, "0");
    DATA.mondaySyncDate = `${pad(now.getDate())}/${pad(now.getMonth()+1)}/${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())} (local)`;

    renderTBI();
    label.textContent = "Refreshed ✓";
    setTimeout(() => { label.textContent = "Refresh"; btn.disabled = false; }, 2000);
  } catch (err) {
    errEl.textContent = "Refresh failed: " + err.message;
    errEl.classList.remove("hidden");
    label.textContent = "Refresh";
    btn.disabled = false;
  }
  icon.style.animation = "";
}

// ── chaser log (notes tab) ─────────────────────────────────────────────
function renderChaserLog() {
  const today = new Date();
  today.setHours(0,0,0,0);

  const rows = D.receivables
    .map(r => {
      const days = calcDaysOvr(r.dueDate) || 0;
      const obj  = getNoteObj(r);
      const note = obj ? obj.note : "";
      return { r, days, note, obj };
    })
    .filter(({ days, r }) => days > 0 || r.isBadDebt)
    .sort((a, b) => b.days - a.days);

  document.getElementById("chaser-log-body").innerHTML = rows.map(({ r, days, note, obj }, i) => {
    const noNote = !note;
    const flag   = days > 30 && noNote;
    const bg     = flag ? "bg-red-50" : i % 2 === 0 ? "bg-white" : "bg-slate-50";
    const display = note.length > 60 ? note.slice(0,60) + "…" : note;
    const noteHtml = `<span class="cursor-pointer hover:underline ${note ? "text-slate-700" : "text-slate-300 italic"}"
      title="${esc(note)}" onclick="editNote(this,'${esc(r.docNo)}')">${note ? esc(display) : "Add note…"}</span>`;
    return `<tr class="${bg} border-b border-slate-100">
      <td class="px-3 py-2 font-medium text-slate-800 ${flag?"text-red-700":""}">${esc(r.vendorName)}</td>
      <td class="px-3 py-2">${daysCell(days)}</td>
      <td class="px-3 py-2 text-right font-semibold text-slate-700">${gbpF(r.balance)}</td>
      <td class="px-3 py-2">${noteHtml}</td>
    </tr>`;
  }).join("") || `<tr><td colspan="4" class="px-3 py-4 text-center text-slate-400">No overdue items</td></tr>`;
}

// ── tabs ───────────────────────────────────────────────────────────────
function showTab(name) {
  ["ar","col","tbi","nts"].forEach(t => {
    document.getElementById("panel-" + t).classList.add("hidden");
    document.getElementById("tab-"   + t).className = "tab-inactive px-4 py-3 text-sm transition";
  });
  document.getElementById("panel-" + name).classList.remove("hidden");
  document.getElementById("tab-"   + name).className = "tab-active px-4 py-3 text-sm transition";
}

// ── PDF Export ──────────────────────────────────────────────────────────
async function exportPDF() {
  // 1. Capture chart canvases as data-URI images
  function chartImg(id) {
    const c = document.getElementById(id);
    if (!c) return "";
    // Canvas-based chart (Chart.js)
    if (c instanceof HTMLCanvasElement) {
      try { return c.toDataURL("image/png"); } catch(e) { return ""; }
    }
    // SVG-based container (e.g. donut built with pure SVG)
    const svg = c.querySelector("svg");
    if (svg) {
      try {
        const clone = svg.cloneNode(true);
        clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
        const xml = new XMLSerializer().serializeToString(clone);
        return "data:image/svg+xml;base64," + btoa(unescape(encodeURIComponent(xml)));
      } catch(e) { return ""; }
    }
    return "";
  }
  const imgDonut  = chartImg("donutChart");
  const imgDays   = chartImg("avgDaysChart");
  const imgBucket = chartImg("bucketChart");
  const imgCash   = chartImg("cashChart");

  // 2. Logo
  const logoEl  = document.querySelector('img[alt="Pubity Group"]');
  const logoSrc = logoEl ? logoEl.src : "";

  // 3. Helpers — ALL currency uses &pound; HTML entity (avoids any unicode escape issue)
  const _esc = s => String(s ?? "").replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
  const _gbp = v => typeof v === "number"
    ? "&pound;" + v.toLocaleString("en-GB", {minimumFractionDigits:2, maximumFractionDigits:2})
    : "&pound;0.00";
  const _pct = (v, tot) => tot > 0 ? (v / tot * 100).toFixed(1) + "%" : "0%";
  const _today = new Date();
  const _daysOvr = due => {
    if (!due) return null;
    const parts = String(due).split("/");
    if (parts.length < 3) return null;
    const [d, m, y] = parts.map(Number);
    return Math.round((_today - new Date(y < 100 ? 2000 + y : y, m - 1, d)) / 86400000);
  };
  const _rowStyle = r => {
    if (r.isBadDebt) return { bg:"#7030A0", fg:"#ffffff" };
    const days = _daysOvr(r.dueDate);
    if (days === null || days <= 0) return { bg:"#ffffff", fg:"#1a2744" };
    if (days <= 31) return { bg:"#00B050", fg:"#ffffff" };
    if (days <= 61) return { bg:"#FF8C00", fg:"#ffffff" };
    if (days <= 92) return { bg:"#FF0000", fg:"#ffffff" };
    return { bg:"#C00000", fg:"#ffffff" };
  };

  // 4. KPI cards — plain ASCII labels, sub as raw HTML (not _esc'd) so &middot; renders correctly
  const t       = D.totals;
  const badDebt = D.receivables.filter(r => r.isBadDebt).reduce((s,r) => s + r.balance, 0);
  const kpiDefs = [
    {label:"Total Balance",    val:t.balance,              sub:D.receivables.length+" invoices",                   bg:"#1a2744", fg:"#ffffff"},
    {label:"Current",          val:t.notDue,               sub:_pct(t.notDue,t.balance)+" of total",               bg:"#ffffff", fg:"#1a2744"},
    {label:"1-30 Days",        val:t.d1_31,                sub:_pct(t.d1_31,t.balance)+" of total",                bg:"#00B050", fg:"#ffffff"},
    {label:"31-60 Days",       val:t.d32_61,               sub:_pct(t.d32_61,t.balance)+" of total",               bg:"#FF8C00", fg:"#ffffff"},
    {label:"61-90 Days",       val:t.d63_92,               sub:_pct(t.d63_92,t.balance)+" of total",               bg:"#FF0000", fg:"#ffffff"},
    {label:"90+ Days",         val:t.over92,               sub:_pct(t.over92,t.balance)+" of total",               bg:"#C00000", fg:"#ffffff"},
    {label:"Risk of Bad Debt", val:badDebt,                sub:"Subset of 90+ Days",                               bg:"#7030A0", fg:"#ffffff"},
  ];
  if (DATA.bankCash) {
    kpiDefs.push({label:"Cash in Bank", val:DATA.bankCash.latest,
      sub:"As at "+_esc(DATA.bankCash.latestDate), bg:"#0369a1", fg:"#ffffff"});
  }
  // sub is already safe HTML — do NOT run through _esc again
  const kpiHtml = kpiDefs.map(c => `
    <div style="flex:1;min-width:0;background:${c.bg};border:1.5px solid ${c.bg==="#ffffff"?"#e2e8f0":c.bg};border-radius:7px;padding:7px 8px;overflow:hidden;text-align:left">
      <div style="font-size:5.5pt;font-weight:700;text-transform:uppercase;letter-spacing:0.05em;color:${c.fg};opacity:0.80">${_esc(c.label)}</div>
      <div style="font-size:8pt;font-weight:700;color:${c.fg};margin-top:3px;word-break:break-all">${_gbp(c.val)}</div>
      <div style="font-size:5.5pt;color:${c.fg};opacity:0.65;margin-top:2px">${c.sub}</div>
    </div>`).join("");

  // 5. Period comparison — plain ASCII labels, HTML entities for arrows
  const p = D.totalsPrev;
  const _delta = (cur, prv, hib) => {
    const d = cur - prv;
    if (Math.abs(d) < 0.01) return '<span style="color:#94a3b8">&mdash;</span>';
    const col = (hib ? d > 0 : d < 0) ? "#16a34a" : "#dc2626";
    return `<span style="color:${col};font-weight:700">${d>0?"&#9650;":"&#9660;"} ${_gbp(Math.abs(d))}</span>`;
  };
  const compHtml = [
    ["Balance",  t.balance, p.balance, false],
    ["Current",  t.notDue,  p.notDue,  true ],
    ["1-30d",    t.d1_31,   p.d1_31,   false],
    ["31-60d",   t.d32_61,  p.d32_61,  false],
    ["61-90d",   t.d63_92,  p.d63_92,  false],
    ["90+d",     t.over92,  p.over92,  false],
  ].map(([l,cur,prv,hib],i) => `
    <tr style="background:${i%2===0?"#f8fafc":"#ffffff"}">
      <td style="padding:4px 7px;font-size:8pt;color:#64748b">${l}</td>
      <td style="padding:4px 7px;font-size:8pt;text-align:right;font-weight:600;color:#0f172a">${_gbp(cur)}</td>
      <td style="padding:4px 7px;font-size:8pt;text-align:right;color:#94a3b8">${_gbp(prv)}</td>
      <td style="padding:4px 7px;font-size:8pt;text-align:right">${_delta(cur,prv,hib)}</td>
    </tr>`).join("");

  // 6. Collections rows — USD amounts shown as plain number (USD)
  const collHtml = D.collections.map((c,i) => {
    const amtCell = c.isUSD
      ? `${c.amt.toLocaleString("en-GB",{minimumFractionDigits:2})} (USD)`
      : _gbp(c.amt);
    return `<tr style="background:${i%2===0?"#f8fafc":"#ffffff"}">
      <td style="padding:3px 7px;font-size:8pt;color:#64748b;white-space:nowrap">${_esc(c.date)}</td>
      <td style="padding:3px 7px;font-size:8pt;font-weight:500;color:#0f172a">${_esc(c.customer)}</td>
      <td style="padding:3px 7px;font-size:8pt;text-align:right;color:#64748b">${amtCell}</td>
      <td style="padding:3px 7px;font-size:8pt;text-align:right;font-weight:700;color:#1d4ed8">${_gbp(c.amtGBP)}</td>
    </tr>`;
  }).join("");

  // 7. Cash card — &middot; used directly in HTML
  const cashCardHtml = DATA.bankCash ? `
    <p style="font-size:9pt;color:#0f172a;margin-bottom:4px">
      <strong style="font-size:11pt">${_gbp(DATA.bankCash.latest)}</strong>
      &nbsp;&mdash;&nbsp; as at ${_esc(DATA.bankCash.latestDate)}, all accounts (GBP)
    </p>` : `<p style="color:#94a3b8;font-style:italic;font-size:9pt">Bank report not available</p>`;

  // 8. AR rows — clean alternating white rows, coloured text per bucket (matching dashboard)
  const arHtml = D.receivables.map((r, i) => {
    const days   = _daysOvr(r.dueDate);
    const rowBg  = i % 2 === 0 ? "#ffffff" : "#f8fafc";
    const daysDisp = days === null ? "" : days <= 0
      ? `<span style="color:#15803d;font-weight:600">Current</span>`
      : days > 90 ? `<span style="color:#be123c;font-weight:700">${days}d</span>`
      : days > 60 ? `<span style="color:#dc2626;font-weight:600">${days}d</span>`
      : days > 30 ? `<span style="color:#ea580c;font-weight:600">${days}d</span>`
      : `<span style="color:#ca8a04;font-weight:600">${days}d</span>`;
    const notDueC = "#15803d";  // green
    const d131C   = "#a16207";  // yellow-700
    const d3261C  = "#c2410c";  // orange-700
    const d6392C  = "#dc2626";  // red-600
    const over92C = r.isBadDebt ? "#6b21a8" : "#be123c";  // purple for bad debt, rose for 92+
    return `<tr style="background:${rowBg};border-bottom:1px solid #e2e8f0">
      <td style="padding:2px 5px;font-size:7.5pt;font-weight:600;overflow:hidden;max-width:0;color:#0f172a">${_esc(r.vendorName)}</td>
      <td style="padding:2px 5px;font-size:7pt;color:#64748b;font-family:monospace">${_esc(r.docNo)}</td>
      <td style="padding:2px 5px;font-size:7pt;color:#64748b;white-space:nowrap">${_esc(r.dueDate)}</td>
      <td style="padding:2px 5px;font-size:7.5pt;text-align:right">${daysDisp}</td>
      <td style="padding:2px 5px;font-size:7.5pt;text-align:right;color:#475569">${_gbp(r.origAmt)}</td>
      <td style="padding:2px 5px;font-size:7.5pt;text-align:right;font-weight:700;color:#0f172a">${_gbp(r.balance)}</td>
      <td style="padding:2px 5px;font-size:7.5pt;text-align:right;color:${r.notDue>0.01?notDueC:"#cbd5e1"}">${r.notDue>0.01?_gbp(r.notDue):"—"}</td>
      <td style="padding:2px 5px;font-size:7.5pt;text-align:right;color:${r.d1_31>0.01?d131C:"#cbd5e1"}">${r.d1_31>0.01?_gbp(r.d1_31):"—"}</td>
      <td style="padding:2px 5px;font-size:7.5pt;text-align:right;color:${r.d32_61>0.01?d3261C:"#cbd5e1"}">${r.d32_61>0.01?_gbp(r.d32_61):"—"}</td>
      <td style="padding:2px 5px;font-size:7.5pt;text-align:right;color:${r.d63_92>0.01?d6392C:"#cbd5e1"}">${r.d63_92>0.01?_gbp(r.d63_92):"—"}</td>
      <td style="padding:2px 5px;font-size:7.5pt;text-align:right;font-weight:700;color:${r.over92>0.01?over92C:"#cbd5e1"}">${r.over92>0.01?_gbp(r.over92):"—"}</td>
    </tr>`;
  }).join("");

  // 9a. Chaser Log rows for PDF
  const today_cl = new Date(); today_cl.setHours(0,0,0,0);
  function _clDays(ds) {
    if (!ds) return null;
    const parts = ds.split("/").map(Number);
    const due = new Date(parts[2], parts[1]-1, parts[0]); due.setHours(0,0,0,0);
    return Math.floor((today_cl - due) / 86400000);
  }
  const _chaserRows = D.receivables
    .map(r => {
      const days = _clDays(r.dueDate) || 0;
      const noteObj = getNoteObj(r);
      const note = noteObj ? noteObj.note : "";
      return { r, days, note };
    })
    .filter(({days, r}) => days > 0 || r.isBadDebt)
    .sort((a, b) => b.days - a.days);

  const chaserLogHtml = _chaserRows.map(({r, days, note}, i) => {
    const bg = r.isBadDebt ? "#f3e8ff" : days > 90 ? "#fee2e2"
             : days > 60 ? "#fef9c3" : days > 30 ? "#fff7ed" : (i%2===0 ? "#fff" : "#f8fafc");
    const cond = r.isBadDebt ? "Risk of Bad Debt"
      : days > 90 ? "90+ Days" : days > 60 ? "61-90 Days"
      : days > 30 ? "31-60 Days" : "≤30 Days";
    return `<tr style="background:${bg};border-bottom:1px solid rgba(0,0,0,0.07)">
      <td style="padding:2px 5px;font-size:7.5pt;font-weight:600;overflow:hidden;max-width:0">${_esc(r.vendorName)}</td>
      <td style="padding:2px 5px;font-size:7pt;text-align:right">${days > 0 ? days+"d" : "—"}</td>
      <td style="padding:2px 5px;font-size:7pt;text-align:right;font-weight:700">${_gbp(r.balance)}</td>
      <td style="padding:2px 5px;font-size:7pt">${_esc(cond)}</td>
      <td style="padding:2px 5px;font-size:7pt;color:#334155">${_esc(note)}</td>
    </tr>`;
  }).join("") || `<tr><td colspan="5" style="padding:6px;color:#94a3b8;font-style:italic">No overdue items</td></tr>`;

  // 9. TBI section — plain text status, no badges
  function _parseDV(str) {
    if (!str) return 0;
    return parseFloat(String(str).replace(/[^0-9.]/g, "")) || 0;
  }
  const tbiGroups = ["Campaigns Kicked Off", "Delivered Campaigns"];
  let tbiGrand = 0, tbiRowIdx = 0;
  const tbiBodyHtml = (DATA.mondayTBI && DATA.mondayTBI.length) ? tbiGroups.flatMap(grp => {
    const items = DATA.mondayTBI.filter(x => x.group === grp);
    if (!items.length) return [];
    const sub = items.reduce((s,x) => s + _parseDV(x.dealValue), 0);
    tbiGrand += sub;
    const hdrRow = `<tr style="background:#dde3ec"><td colspan="5" style="padding:5px 7px;font-size:7.5pt;font-weight:700;color:#1e2d40;text-transform:uppercase;letter-spacing:0.04em">${_esc(grp)}</td></tr>`;
    const dataRows = items.map(x => {
      const bg = tbiRowIdx++ % 2 === 0 ? "#f8fafc" : "#ffffff";
      const dv = x.dealValue ? `<strong>&pound;${String(x.dealValue).replace(/[^0-9.,]/g,"")}</strong>` : "&mdash;";
      return `<tr style="background:${bg};border-bottom:1px solid #e2e8f0">
        <td style="padding:3px 7px;font-size:8pt;font-weight:500;color:#0f172a">${_esc(x.name)}</td>
        <td style="padding:3px 7px;font-size:8pt;color:#374151">${dv}</td>
        <td style="padding:3px 7px;font-size:8pt;color:#475569">${x.io ? _esc(x.io) : "&mdash;"}</td>
        <td style="padding:3px 7px;font-size:8pt;color:#475569">${x.ncf ? _esc(x.ncf) : "&mdash;"}</td>
        <td style="padding:3px 7px;font-size:8pt;color:#475569">${x.lineItems ? _esc(x.lineItems) : "&mdash;"}</td>
      </tr>`;
    }).join("");
    const subRow = `<tr style="background:#f1f5f9;border-top:1.5px solid #94a3b8">
      <td style="padding:4px 7px;font-size:7.5pt;font-weight:700;color:#374151">Subtotal</td>
      <td style="padding:4px 7px;font-size:7.5pt;font-weight:700;color:#0f172a">&pound;${sub.toLocaleString("en-GB")}</td>
      <td colspan="3"></td>
    </tr>`;
    return [hdrRow + dataRows + subRow];
  }).join("") : `<tr><td colspan="5" style="padding:16px;text-align:center;color:#94a3b8;font-style:italic">No campaign data available</td></tr>`;
  const tbiTotalRow = (DATA.mondayTBI && DATA.mondayTBI.length) ? `
    <tr style="background:#1e293b;color:#ffffff">
      <td style="padding:6px 7px;font-size:8.5pt;font-weight:700">Grand Total</td>
      <td style="padding:6px 7px;font-size:8.5pt;font-weight:700">&pound;${tbiGrand.toLocaleString("en-GB")}</td>
      <td colspan="3"></td>
    </tr>` : "";

  // 10. Today string for footer/header
  const todayStr = _today.toLocaleDateString("en-GB", {day:"2-digit",month:"long",year:"numeric"});

  // 11. Build complete self-contained HTML — @charset UTF-8, Arial font, &pound; everywhere
  const html = `<!DOCTYPE html>
<html lang="en"><head>
<meta charset="UTF-8">
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<title>Aged Receivables - ${_esc(D.reportDate)}</title>
<style>
@charset "UTF-8";
@page {
  size: A4 portrait;
  margin: 14mm 12mm 18mm 12mm;
  @bottom-right {
    content: "Page " counter(page) " of " counter(pages);
    font-size: 7.5pt; color: #64748b; font-family: Arial, sans-serif;
  }
  @bottom-left {
    content: "Prepared by Finance | Pubity Group | ${todayStr}";
    font-size: 7.5pt; color: #64748b; font-family: Arial, sans-serif;
  }
}
* { box-sizing:border-box; margin:0; padding:0; -webkit-print-color-adjust:exact; print-color-adjust:exact; }
body { font-family:Arial,Helvetica,sans-serif; font-size:9pt; color:#0f172a; background:#fff; }
h2 { font-size:8.5pt; font-weight:700; color:#0e2841; padding:5px 0 4px; border-bottom:1.5px solid #156082; margin:12px 0 7px; letter-spacing:0.03em; text-transform:uppercase; page-break-after:avoid; break-after:avoid; }
table { width:100%; border-collapse:collapse; }
thead th { background:#0e2841; color:#fff; font-size:7pt; padding:4px 5px; text-align:left; font-weight:700; }
thead th.r { text-align:right; }
tfoot td { background:#374151; color:#fff; font-size:6.5pt; font-weight:700; padding:4px 3px; }
tfoot td.r { text-align:right; }
tr { page-break-inside:avoid; }
thead { display:table-header-group; }
tfoot { display:table-row-group; }
.two-col { display:grid; grid-template-columns:1fr 1fr; gap:14px; }
.nb { page-break-inside:avoid; }
.np { page-break-before:always; }
img { max-width:100%; height:auto; }
</style>
</head><body>

<div style="display:flex;align-items:center;justify-content:space-between;padding-bottom:10px;border-bottom:2.5px solid #156082;margin-bottom:12px">
  <div style="display:flex;align-items:center;gap:12px">
    ${logoSrc ? `<img src="${logoSrc}" style="height:28px;width:auto">` : ""}
    <div>
      <div style="font-size:14pt;font-weight:700;color:#0e2841;line-height:1.2">Aged Accounts Receivable</div>
      <div style="font-size:8pt;color:#64748b;margin-top:2px">
        Pubity Group &nbsp;&middot;&nbsp; As at <strong style="color:#0e2841">${_esc(D.reportDate)}</strong>
        &nbsp;&middot;&nbsp; USD/GBP <strong style="color:#0e2841">${D.exchangeRate.toFixed(4)}</strong>
        <span style="color:#94a3b8">(as at ${_esc(D.fxDate)})</span>
      </div>
    </div>
  </div>
  <div style="text-align:right;font-size:7.5pt;color:#64748b;line-height:1.7">
    <div>For internal use only</div>
    <div style="color:#94a3b8;margin-top:1px">Printed ${_esc(todayStr)}</div>
  </div>
</div>

<div style="display:flex;flex-wrap:nowrap;gap:6px;margin-bottom:14px">${kpiHtml}</div>

<h2>Receivables Distribution</h2>
<div class="nb" style="margin-bottom:10px">
  <div style="font-size:7.5pt;font-weight:600;color:#0e2841;margin-bottom:4px">By bucket &mdash; ${_esc(D.reportDate)}</div>
  ${imgDonut ? `<img src="${imgDonut}" style="width:100%;max-height:280px;object-fit:contain">` : `<p style="color:#94a3b8">Chart not available</p>`}
</div>

<h2>Average Days Overdue Trend</h2>
<div class="nb" style="margin-bottom:10px">
  ${imgDays ? `<img src="${imgDays}" style="width:100%;max-height:220px;object-fit:contain">` : `<p style="color:#94a3b8">Chart not available</p>`}
</div>

<h2>Period Comparison</h2>
<div class="nb" style="margin-bottom:12px">
  <table>
    <thead><tr>
      <th>Metric</th>
      <th class="r">${_esc(D.reportDate)}</th>
      <th class="r" style="color:#c5d0de">${_esc(D.prevDate)}</th>
      <th class="r">Change</th>
    </tr></thead>
    <tbody>${compHtml}</tbody>
    <tfoot><tr>
      <td colspan="4" style="background:#dbeafe;color:#1d4ed8;font-weight:600;font-size:7.5pt">
        Collections received this period: ${_gbp(D.collectionsTotal)}
      </td>
    </tr></tfoot>
  </table>
</div>

<div class="nb" style="margin-bottom:14px">
  <h2>Cash in Bank</h2>
  ${cashCardHtml}
  ${imgCash ? `<img src="${imgCash}" style="width:100%;max-height:170px;object-fit:contain;margin-top:8px;display:block">` : ""}
</div>

<h2 class="np">Aged Receivables Detail &mdash; ${_esc(D.reportDate)}</h2>
<table style="table-layout:fixed">
  <colgroup>
    <col style="width:24%"><col style="width:9%">
    <col style="width:8%"><col style="width:6%"><col style="width:9%">
    <col style="width:9%"><col style="width:7%"><col style="width:7%">
    <col style="width:7%"><col style="width:7%"><col style="width:7%">
  </colgroup>
  <thead><tr>
    <th>Customer</th><th>Doc No</th>
    <th>Due Date</th><th class="r">Days OD</th><th class="r">Orig Amt</th>
    <th class="r">Balance</th><th class="r">Current</th><th class="r">1-30d</th>
    <th class="r">31-60d</th><th class="r">61-90d</th><th class="r">90+d</th>
  </tr></thead>
  <tbody>${arHtml}</tbody>
  <tfoot><tr>
    <td colspan="5">Total &mdash; ${_esc(D.reportDate)}</td>
    <td class="r">${_gbp(t.balance)}</td><td class="r">${_gbp(t.notDue)}</td>
    <td class="r">${_gbp(t.d1_31)}</td><td class="r">${_gbp(t.d32_61)}</td>
    <td class="r">${_gbp(t.d63_92)}</td><td class="r">${_gbp(t.over92)}</td>
  </tr></tfoot>
</table>

<h2 class="np">To Be Invoiced &mdash; Campaign Billing</h2>
<table style="table-layout:fixed">
  <colgroup>
    <col style="width:40%"><col style="width:14%">
    <col style="width:16%"><col style="width:16%"><col style="width:14%">
  </colgroup>
  <thead><tr>
    <th>Campaign Name</th><th>Deal Value</th>
    <th>Insertion Order</th><th>New Customer Form</th><th>Line Items</th>
  </tr></thead>
  <tbody>${tbiBodyHtml}${tbiTotalRow}</tbody>
</table>


</body></html>`;

  // 12. Open in new window and print
  const w = window.open("", "_blank");
  if (!w) { alert("Popup blocked - allow popups for this page and try again."); return; }
  w.document.open("text/html", "replace");
  w.document.write(html);
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 800);
}

// ── render all (called on init and file switch) ────────────────────────
function renderAll() {
  document.getElementById("hdr-date").textContent  = D.reportDate;
  document.getElementById("hdr-built").textContent = DATA.buildDate;
  document.getElementById("hdr-fx").textContent    = D.exchangeRate.toFixed(4);
  document.getElementById("hdr-fx-date").textContent = `(as at ${D.fxDate})`;

  renderKPI();
  renderDonut();
  renderCashCard();
  renderComparison();
  filtered = D.receivables.map(r => ({...r, _daysOvr: calcDaysOvr(r.dueDate) ?? -9999}))
    .sort((a, b) => String(a.vendorNo ?? "").localeCompare(String(b.vendorNo ?? "")));
  document.getElementById("search").value = "";
  document.getElementById("fCond").value  = "";
  renderAR(filtered);
  renderARFoot();
  renderCollections();
  renderChaserLog();
}

// ── init ───────────────────────────────────────────────────────────────
initDarkMode();
buildFilePicker();
renderAll();
renderTBI();          // Monday.com — doesn't change with file switch
renderCashLine();     // Bank report — doesn't change with file switch
renderAvgDaysChart();            // All-files avg days overdue — doesn't change with file switch
renderBucketDrilldown("d1_31_abs"); // Default bucket — doesn't change with file switch

// ── Floating scrollbar for AR table ─────────────────────────────────────
(function() {
  const wrap = document.getElementById('ar-wrap');
  if (!wrap) return;
  const bar   = document.createElement('div');
  bar.className = 'float-scroll no-print';
  const inner = document.createElement('div');
  inner.className = 'float-scroll-inner';
  bar.appendChild(inner);
  document.body.appendChild(bar);
  function syncWidth() {
    const tbl = wrap.querySelector('table');
    if (tbl) inner.style.width = tbl.scrollWidth + 'px';
  }
  function updateVisibility() {
    const r = wrap.getBoundingClientRect();
    const inView = r.top < window.innerHeight && r.bottom > 0;
    bar.style.display = (inView && r.bottom > window.innerHeight) ? 'block' : 'none';
  }
  let lock = false;
  wrap.addEventListener('scroll', () => { if (lock) return; lock=true; bar.scrollLeft=wrap.scrollLeft; lock=false; });
  bar.addEventListener('scroll',  () => { if (lock) return; lock=true; wrap.scrollLeft=bar.scrollLeft; lock=false; });
  window.addEventListener('scroll', updateVisibility, { passive:true });
  window.addEventListener('resize', () => { syncWidth(); updateVisibility(); });
  syncWidth();
  updateVisibility();
})();
</script>

<!-- ── Customer Drill-Down Modal ──────────────────────────────────────── -->
<div id="drill-modal"
     class="fixed inset-0 z-50 hidden items-start justify-center p-4 pt-8"
     style="background:rgba(15,20,35,0.65)"
     onclick="if(event.target===this)closeDrill()">
  <div class="bg-white rounded-xl shadow-2xl w-full max-w-5xl max-h-[92vh] flex flex-col">

    <!-- header -->
    <div class="flex items-start justify-between px-6 py-4 border-b border-slate-200 flex-shrink-0">
      <div class="flex-1 min-w-0">
        <h2 id="drill-title" class="text-lg font-bold text-slate-800 truncate"></h2>
        <div id="drill-summary" class="flex flex-wrap gap-3 mt-2"></div>
        <!-- action buttons -->
        <!-- Customer notes -->
        <div class="mt-3">
          <label class="block text-xs font-semibold text-slate-500 mb-1">Customer Notes</label>
          <textarea id="drill-cust-note" rows="2" placeholder="Add customer-level notes here…"
            class="w-full border border-slate-200 rounded-lg px-3 py-2 text-xs text-slate-700 resize-none focus:outline-none focus:ring-2 focus:ring-blue-400 bg-slate-50"></textarea>
        </div>
        <div class="flex gap-2 mt-2">
          <button id="drill-contact-btn" onclick="drillContactFind()"
            class="px-3 py-1.5 rounded-lg text-xs font-medium border border-slate-300 text-slate-600 hover:bg-slate-50 flex items-center gap-1.5">
            <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M16 12a4 4 0 10-8 0 4 4 0 008 0zm0 0v1.5a2.5 2.5 0 005 0V12a9 9 0 10-9 9m4.5-1.206a8.959 8.959 0 01-4.5 1.207"/></svg>
            Find AP Contact
          </button>
          <button id="drill-chaser-btn" onclick="drillDraftChaser()"
            class="px-3 py-1.5 rounded-lg text-xs font-medium border border-slate-300 text-slate-600 hover:bg-slate-50 flex items-center gap-1.5">
            <svg class="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"/></svg>
            Draft Chaser Email
          </button>
        </div>
      </div>
      <button onclick="closeDrill()"
              class="ml-4 text-slate-400 hover:text-slate-700 text-2xl leading-none font-light flex-shrink-0">&times;</button>
    </div>

    <!-- contact finder panel -->
    <div id="drill-contact-panel" class="hidden px-6 py-3 border-b border-slate-100 bg-slate-50 text-xs">
      <div id="drill-contact-results" class="mb-3"></div>
      <label class="block text-slate-500 mb-1 font-medium">Saved contact notes (editable):</label>
      <textarea id="drill-contact-notes" rows="2" placeholder="Type or paste contact details to save for next time…"
        class="w-full border border-slate-200 rounded px-2 py-1.5 text-xs text-slate-700 resize-none focus:outline-none focus:ring-1 focus:ring-blue-400"
        oninput="drillSaveContactNotes()"></textarea>
      <p class="text-slate-400 mt-1.5 italic">Contact details sourced from web — please verify before use.</p>
    </div>

    <!-- chaser draft panel -->
    <div id="drill-chaser-panel" class="hidden px-6 py-3 border-b border-slate-100 bg-blue-50 text-xs">
      <div class="flex items-center justify-between mb-2">
        <span class="font-semibold text-slate-700">Draft Chaser Email</span>
        <button onclick="drillCopyChaser()"
          class="px-3 py-1 rounded text-xs font-medium bg-blue-600 text-white hover:bg-blue-700 flex items-center gap-1">
          <svg class="w-3 h-3" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M8 5H6a2 2 0 00-2 2v12a2 2 0 002 2h10a2 2 0 002-2v-1M8 5a2 2 0 002 2h2a2 2 0 002-2M8 5a2 2 0 012-2h2a2 2 0 012 2m0 0h2a2 2 0 012 2v3m2 4H10m0 0l3-3m-3 3l3 3"/></svg>
          Copy to Clipboard
        </button>
      </div>
      <div id="drill-chaser-subject" class="font-medium text-slate-700 mb-1 select-all"></div>
      <textarea id="drill-chaser-body" rows="12" readonly
        class="w-full border border-slate-200 rounded px-2 py-1.5 text-xs text-slate-700 font-mono resize-none focus:outline-none select-all bg-white"></textarea>
    </div>

    <!-- invoice table -->
    <div class="overflow-auto p-4">
      <table class="w-full text-xs">
        <thead style="background:#1a2744;color:#ffffff;text-align:left;position:sticky;top:0;z-index:10">
          <tr>
            <th class="px-3 py-2.5">Document No</th>
            <th class="px-3 py-2.5 whitespace-nowrap">Invoice Date</th>
            <th class="px-3 py-2.5 whitespace-nowrap">Due Date</th>
            <th class="px-3 py-2.5 text-right whitespace-nowrap">Orig. Amount</th>
            <th class="px-3 py-2.5 text-right whitespace-nowrap">Current Balance</th>
            <th class="px-3 py-2.5 text-right whitespace-nowrap">Days OD</th>
            <th class="px-3 py-2.5">Bucket</th>
            <th class="px-3 py-2.5" style="min-width:220px">Notes</th>
          </tr>
        </thead>
        <tbody id="drill-body"></tbody>
      </table>
    </div>

  </div>
</div>

<script>
// ── Customer drill-down ─────────────────────────────────────────────────
function _parseReportDate(s) {
  if (!s) return null;
  const m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (!m) return null;
  let [,d,mo,y] = m; if (y.length===2) y="20"+y;
  return new Date(parseInt(y), parseInt(mo)-1, parseInt(d));
}

function _drillBucket(isBadDebt, daysOvr) {
  if (isBadDebt)      return { label:"Risk of Bad Debt", bg:"#7030A0", fg:"#fff" };
  if (daysOvr===null) return { label:"—",                bg:"#f8fafc", fg:"#64748b" };
  if (daysOvr > 90)   return { label:"90+ Days",         bg:"#C00000", fg:"#fff" };
  if (daysOvr > 60)   return { label:"61–90 Days",       bg:"#FF0000", fg:"#fff" };
  if (daysOvr > 30)   return { label:"31–60 Days",       bg:"#FFC000", fg:"#1a2744" };
  if (daysOvr > 0)    return { label:"1–30 Days",        bg:"#FFC000", fg:"#1a2744" };
  return               { label:"Current",         bg:"#00B050", fg:"#fff" };
}

// SVG sparkline — kept for potential future use but not shown in current table
function _sparkline(trend) {
  const vals = trend.map(t => t.b);
  if (vals.length === 0) return '<span class="text-slate-300">—</span>';
  const W=90, H=26, pad=3;
  if (vals.length === 1) {
    return `<svg width="${W}" height="${H}"><circle cx="${W/2}" cy="${H/2}" r="3" fill="#94a3b8"/></svg>`;
  }
  const mn=Math.min(...vals), mx=Math.max(...vals), rng=mx-mn||1;
  const pts = vals.map((v,i) => {
    const x = pad + (i/(vals.length-1))*(W-pad*2);
    const y = (H-pad) - ((v-mn)/rng)*(H-pad*2);
    return [x.toFixed(1), y.toFixed(1)];
  });
  const ptStr = pts.map(p=>p.join(",")).join(" ");
  const [lx,ly] = pts[pts.length-1];
  const reducing = vals[vals.length-1] < vals[0] * 0.99;
  const flat     = Math.abs(vals[vals.length-1]-vals[0]) / (vals[0]||1) < 0.01;
  const col = reducing ? "#16a34a" : flat ? "#f59e0b" : "#dc2626";
  const label = reducing ? "↓" : flat ? "→" : "↑";
  const labelCol = col;
  return `<span title="${reducing?"Reducing":flat?"Flat":"Increasing"}">` +
    `<svg width="${W}" height="${H}" style="vertical-align:middle">` +
    `<polyline points="${ptStr}" fill="none" stroke="${col}" stroke-width="1.8" stroke-linejoin="round" stroke-linecap="round"/>` +
    `<circle cx="${lx}" cy="${ly}" r="2.5" fill="${col}"/>` +
    `</svg>` +
    `<span style="color:${labelCol};font-weight:700;font-size:11px;margin-left:2px">${label}</span></span>`;
}

// ── Contact cache (localStorage) ───────────────────────────────────────
const _CONTACT_CACHE_KEY = "pgContactCache_v1";
function _contactCache() {
  try { return JSON.parse(localStorage.getItem(_CONTACT_CACHE_KEY) || "{}"); } catch { return {}; }
}
function _saveContactCache(cache) {
  try { localStorage.setItem(_CONTACT_CACHE_KEY, JSON.stringify(cache)); } catch {}
}

// ── Drill-down state ────────────────────────────────────────────────────
let _drillVendor = "";
let _drillInvoices = [];

function openDrill(el) {
  _drillVendor = el.dataset.vendor;

  // Only rows from the currently selected file
  const today = new Date(); today.setHours(0,0,0,0);
  const rows = (D.receivables || []).filter(r => r.vendorName === _drillVendor);

  _drillInvoices = rows.map(r => {
    const due = _parseReportDate(r.dueDate);
    const daysOvr = due ? Math.max(0, Math.floor((today - due) / 86400000)) : null;
    // Drill-down always shows the Excel note directly — no localStorage override.
    // This keeps the drill-down in sync with the current xlsx regardless of any
    // previously saved notes. Edits made here are saved to localStorage and appear
    // in the Chaser Log tab.
    const note = (r.xlComment || "").trim();
    return { docNo: r.docNo, invoiceDate: r.postingDate, dueDate: r.dueDate,
             origAmt: r.origAmt, currentBalance: r.balance,
             isBadDebt: r.isBadDebt, daysOvr, note };
  }).sort((a, b) => b.currentBalance - a.currentBalance);

  // ── Summary stats ──────────────────────────────────────────────────
  const totalBal = _drillInvoices.reduce((s, inv) => s + inv.currentBalance, 0);
  const invCount = _drillInvoices.length;
  const allDue   = _drillInvoices.map(inv => _parseReportDate(inv.dueDate)).filter(Boolean);
  const oldestDue = allDue.length
    ? new Date(Math.min(...allDue)).toLocaleDateString("en-GB", {day:"2-digit",month:"2-digit",year:"2-digit"})
    : "—";
  // Customer-level note (per vendor, stored separately from invoice notes)
  const custNote = _getCustNote(_drillVendor);

  document.getElementById("drill-title").textContent = _drillVendor;
  document.getElementById("drill-summary").innerHTML = [
    `<span class="bg-slate-100 rounded-lg px-3 py-1.5"><span class="text-slate-500">Balance</span> <strong class="text-slate-800 ml-1">${gbpF(totalBal)}</strong></span>`,
    `<span class="bg-slate-100 rounded-lg px-3 py-1.5"><span class="text-slate-500">Invoices</span> <strong class="text-slate-800 ml-1">${invCount}</strong></span>`,
    `<span class="bg-slate-100 rounded-lg px-3 py-1.5"><span class="text-slate-500">Oldest due</span> <strong class="text-slate-800 ml-1">${oldestDue}</strong></span>`,
  ].join("");

  // Render customer notes field
  const cnEl = document.getElementById("drill-cust-note");
  cnEl.value = custNote;
  cnEl.oninput = () => _saveCustNote(_drillVendor, cnEl.value);

  // ── Invoice table ──────────────────────────────────────────────────
  document.getElementById("drill-body").innerHTML = _drillInvoices.map((inv, i) => {
    const bucket = _drillBucket(inv.isBadDebt, inv.daysOvr);
    const bg   = i % 2 === 0 ? "#ffffff" : "#f8fafc";
    const dOvr = inv.daysOvr !== null && inv.daysOvr > 0
      ? `<span class="font-semibold" style="color:#dc2626">${inv.daysOvr}d</span>`
      : `<span class="text-slate-400">${inv.daysOvr === 0 ? "0" : "—"}</span>`;
    return `<tr style="background:${bg}" class="border-b border-slate-100 align-middle">
      <td class="px-3 py-2.5 font-mono text-slate-700 whitespace-nowrap">${esc(inv.docNo)}</td>
      <td class="px-3 py-2.5 whitespace-nowrap text-slate-600">${esc(inv.invoiceDate)}</td>
      <td class="px-3 py-2.5 whitespace-nowrap text-slate-600">${esc(inv.dueDate)}</td>
      <td class="px-3 py-2.5 text-right text-slate-600">${gbpF(inv.origAmt)}</td>
      <td class="px-3 py-2.5 text-right font-semibold text-slate-800">${gbpF(inv.currentBalance)}</td>
      <td class="px-3 py-2.5 text-right">${dOvr}</td>
      <td class="px-3 py-2.5 whitespace-nowrap"><span class="px-2 py-0.5 rounded text-xs font-semibold" style="background:${bucket.bg};color:${bucket.fg}">${bucket.label}</span></td>
      <td class="px-3 py-2.5" style="max-width:380px;word-break:break-word">
        <span class="cursor-pointer hover:underline ${inv.note ? 'text-slate-700' : 'text-slate-300 italic'}"
          onclick="editNoteDrill(this,'${esc(inv.docNo)}')">${inv.note ? esc(inv.note) : 'Click to add note…'}</span>
      </td>
    </tr>`;
  }).join("") || `<tr><td colspan="8" class="px-3 py-6 text-center text-slate-400">No invoices found for this file</td></tr>`;

  // ── Reset action panels ─────────────────────────────────────────────
  document.getElementById("drill-contact-panel").classList.add("hidden");
  document.getElementById("drill-chaser-panel").classList.add("hidden");

  // Pre-load cached contact notes
  const cache = _contactCache();
  document.getElementById("drill-contact-notes").value = cache[_drillVendor] || "";

  const m = document.getElementById("drill-modal");
  m.classList.remove("hidden"); m.classList.add("flex");
}

function closeDrill() {
  const m = document.getElementById("drill-modal");
  m.classList.add("hidden"); m.classList.remove("flex");
}

// ── Contact Finder ──────────────────────────────────────────────────────
// Search result cache: { vendorName: { ts, domain, logo, description, emails } }
const _SEARCH_CACHE_KEY = "pgContactSearch_v1";
function _searchCache() {
  try { return JSON.parse(localStorage.getItem(_SEARCH_CACHE_KEY) || "{}"); } catch { return {}; }
}
function _saveSearchCache(cache) {
  try { localStorage.setItem(_SEARCH_CACHE_KEY, JSON.stringify(cache)); } catch {}
}

async function drillContactFind() {
  const panel = document.getElementById("drill-contact-panel");
  const wasHidden = panel.classList.contains("hidden");
  panel.classList.toggle("hidden", !wasHidden);
  if (!wasHidden) return;  // toggled off — nothing to do

  // Restore saved notes
  const notes = _contactCache();
  document.getElementById("drill-contact-notes").value = notes[_drillVendor] || "";

  const resultsEl = document.getElementById("drill-contact-results");

  // Return cached search if fresh (< 7 days)
  const sc = _searchCache();
  const cached = sc[_drillVendor];
  if (cached && (Date.now() - cached.ts) < 7 * 86400000) {
    _renderContactResults(resultsEl, cached, true);
    return;
  }

  // Show loading state
  resultsEl.innerHTML =
    `<div class="flex items-center gap-2 text-slate-500 py-1">
       <svg class="animate-spin w-3.5 h-3.5 text-blue-500" fill="none" viewBox="0 0 24 24">
         <circle class="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" stroke-width="4"></circle>
         <path class="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v8H4z"></path>
       </svg>
       Searching for <strong class="text-slate-700 ml-1">${esc(_drillVendor)}</strong>…
     </div>`;

  const result = { ts: Date.now(), domain: null, logo: null, description: null, emails: [] };

  // ── 1. Clearbit Autocomplete — company domain & logo ──────────────────
  try {
    const cbResp = await Promise.race([
      fetch(`https://autocomplete.clearbit.com/v1/companies/suggest?query=${encodeURIComponent(_drillVendor)}`),
      new Promise((_, rej) => setTimeout(() => rej(new Error("timeout")), 5000))
    ]);
    if (cbResp.ok) {
      const cbData = await cbResp.json();
      if (cbData && cbData.length > 0) {
        // Best match: first result whose name loosely matches
        const vendor = _drillVendor.toLowerCase();
        const match  = cbData.find(c => {
          const n = (c.name || "").toLowerCase();
          return n.includes(vendor.split(/\s+/)[0]) || vendor.includes(n.split(/\s+/)[0]);
        }) || cbData[0];
        result.domain = match.domain || null;
        result.logo   = match.logo   || null;
        result.cbName = match.name   || null;
      }
    }
  } catch (_) {}

  // ── 2. DuckDuckGo Instant Answer — description ───────────────────────
  try {
    const ddgQ = encodeURIComponent(_drillVendor + " company");
    const ddgResp = await Promise.race([
      fetch(`https://api.duckduckgo.com/?q=${ddgQ}&format=json&no_html=1&skip_disambig=1`),
      new Promise((_, rej) => setTimeout(() => rej(new Error("timeout")), 5000))
    ]);
    if (ddgResp.ok) {
      const ddg = await ddgResp.json();
      if (ddg.AbstractText) result.description = ddg.AbstractText.slice(0, 220);
      // Pick up the official website if Clearbit missed it
      if (!result.domain && ddg.AbstractURL) {
        try { result.domain = new URL(ddg.AbstractURL).hostname.replace(/^www\./, ""); } catch (_) {}
      }
      // Scan related topics for email-like patterns
      const allText = [ddg.AbstractText, ...(ddg.RelatedTopics||[]).map(t=>t.Text||"")].join(" ");
      const emailMatches = allText.match(/[\w.+-]+@[\w.-]+\.[a-z]{2,}/gi) || [];
      result.emails = [...new Set(emailMatches)].slice(0, 3);
    }
  } catch (_) {}

  // Cache and render
  sc[_drillVendor] = result;
  _saveSearchCache(sc);
  _renderContactResults(resultsEl, result, false);
}

function _renderContactResults(el, r, fromCache) {
  const cacheNote = fromCache
    ? `<span class="text-slate-400 ml-2">(cached ${new Date(r.ts).toLocaleDateString("en-GB")}
        — <button class="underline text-blue-500 hover:text-blue-700" onclick="_clearContactCache()">refresh</button>)</span>`
    : "";
  let html = `<div class="space-y-2">`;

  // Company identity row
  if (r.domain || r.logo || r.cbName) {
    html += `<div class="flex items-center gap-2">`;
    if (r.logo) html += `<img src="${esc(r.logo)}" class="w-6 h-6 rounded object-contain border border-slate-100" onerror="this.style.display='none'">`;
    if (r.cbName) html += `<span class="font-semibold text-slate-700">${esc(r.cbName)}</span>`;
    html += `</div>`;
  }

  // Website
  if (r.domain) {
    html += `<div class="flex items-center gap-1.5 text-slate-600">
      <svg class="w-3.5 h-3.5 text-blue-400 flex-shrink-0" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M21 12a9 9 0 01-9 9m9-9a9 9 0 00-9-9m9 9H3m9 9a9 9 0 01-9-9m9 9c1.657 0 3-4.03 3-9s-1.343-9-3-9m0 18c-1.657 0-3-4.03-3-9s1.343-9 3-9"/></svg>
      <strong>Website:</strong>
      <span class="font-mono">${esc(r.domain)}</span>
      <span class="text-slate-400 ml-1">— try
        <span class="font-mono">${esc(r.domain)}/contact</span> or
        <span class="font-mono">${esc(r.domain)}/accounts-payable</span>
      </span>
    </div>`;
  }

  // AP email guesses from domain
  if (r.domain) {
    const guesses = [`ap@${r.domain}`, `accounts.payable@${r.domain}`, `finance@${r.domain}`, `invoices@${r.domain}`];
    html += `<div class="flex items-start gap-1.5 text-slate-600">
      <svg class="w-3.5 h-3.5 text-blue-400 flex-shrink-0 mt-0.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M3 8l7.89 5.26a2 2 0 002.22 0L21 8M5 19h14a2 2 0 002-2V7a2 2 0 00-2-2H5a2 2 0 00-2 2v10a2 2 0 002 2z"/></svg>
      <div><strong>Likely AP emails</strong> <span class="text-slate-400">(unverified):</span><br>
        ${guesses.map(e => `<span class="font-mono select-all">${esc(e)}</span>`).join(" &nbsp;·&nbsp; ")}
      </div>
    </div>`;
  }

  // Emails found in DDG text
  if (r.emails && r.emails.length) {
    html += `<div class="flex items-center gap-1.5 text-slate-600">
      <svg class="w-3.5 h-3.5 text-green-500 flex-shrink-0" fill="currentColor" viewBox="0 0 20 20"><path d="M2.003 5.884L10 9.882l7.997-3.998A2 2 0 0016 4H4a2 2 0 00-1.997 1.884z"/><path d="M18 8.118l-8 4-8-4V14a2 2 0 002 2h12a2 2 0 002-2V8.118z"/></svg>
      <div><strong>Found in web:</strong>
        ${r.emails.map(e => `<span class="font-mono select-all text-green-700">${esc(e)}</span>`).join(", ")}
      </div>
    </div>`;
  }

  // Description
  if (r.description) {
    html += `<div class="text-slate-500 leading-relaxed border-t border-slate-200 pt-2 mt-1">${esc(r.description)}</div>`;
  }

  // Nothing found
  if (!r.domain && !r.description && !(r.emails && r.emails.length)) {
    html += `<p class="text-slate-400 italic">No company data found automatically. Use the notes box below to save contact details manually.</p>`;
  }

  html += `<div class="text-slate-400 mt-1">${cacheNote}</div></div>`;
  el.innerHTML = html;
}

function _clearContactCache() {
  const sc = _searchCache();
  delete sc[_drillVendor];
  _saveSearchCache(sc);
  // Re-trigger search
  document.getElementById("drill-contact-panel").classList.add("hidden");
  drillContactFind();
}

function drillSaveContactNotes() {
  const cache = _contactCache();
  cache[_drillVendor] = document.getElementById("drill-contact-notes").value;
  _saveContactCache(cache);
}

// ── Draft Chaser Email ──────────────────────────────────────────────────
function drillDraftChaser() {
  const panel = document.getElementById("drill-chaser-panel");
  const wasHidden = panel.classList.contains("hidden");
  panel.classList.toggle("hidden", !wasHidden);
  if (!wasHidden) return;  // just toggled off — don't regenerate

  const activeInvs = _drillInvoices.filter(inv => inv.currentBalance > 0);
  const totalBal   = activeInvs.reduce((s, inv) => s + inv.currentBalance, 0);
  const totalStr   = totalBal.toLocaleString("en-GB", {minimumFractionDigits:2, maximumFractionDigits:2});

  // Determine overdue phrase from the worst invoice
  const maxDays = Math.max(...activeInvs.map(inv => inv.daysOvr || 0));
  let overduePhrase;
  if (maxDays >= 90)      overduePhrase = "remain outstanding despite previous contact";
  else if (maxDays >= 30) overduePhrase = "are significantly overdue";
  else if (maxDays >= 1)  overduePhrase = "are now overdue";
  else                    overduePhrase = "are coming up for payment";

  const subject = `Payment Due — ${_drillVendor}`;

  const lines = activeInvs.map(inv => {
    const bal = inv.currentBalance.toLocaleString("en-GB",{minimumFractionDigits:2,maximumFractionDigits:2});
    return `  ${inv.docNo}  |  Due: ${inv.dueDate}  |  £${bal}`;
  }).join("\n");

  const body =
`Hi ,

Hope you're well. Just wanted to check in on the below invoice(s) which ${overduePhrase} — could you let me know the status or when we can expect payment?

${lines}

Total: £${totalStr}

Let me know if you need anything from our side.

Thanks,
Muhammad`;

  document.getElementById("drill-chaser-subject").textContent = "Subject: " + subject;
  document.getElementById("drill-chaser-body").value = body;
}

function drillCopyChaser() {
  const subject = document.getElementById("drill-chaser-subject").textContent;
  const body    = document.getElementById("drill-chaser-body").value;
  const full    = subject + "\n\n" + body;
  navigator.clipboard.writeText(full).then(() => {
    const btn = document.querySelector("#drill-chaser-panel button");
    const orig = btn.innerHTML;
    btn.textContent = "Copied!";
    btn.style.background = "#16a34a";
    setTimeout(() => { btn.innerHTML = orig; btn.style.background = ""; }, 1800);
  });
}
</script>
</body>
</html>
"""

if _TAILWIND_JS and _CHARTJS_JS:
    _vendor_tags = ("<script>" + _TAILWIND_JS + "</script>\n"
                    "  <script>" + _CHARTJS_JS + "</script>")
else:
    _vendor_tags = ('<script src="https://cdn.tailwindcss.com"></script>\n'
                    '  <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>')

html = (HTML
    .replace("__VENDOR_SCRIPTS__", _vendor_tags)
    .replace("__DATA__", data_js)
    .replace("__LOGO__", logo_b64)
    .replace("__LOGO_HIDE__", "" if logo_b64 else "display:none")
    .replace("__LOGO_DARK__", logo_dark_b64)
    .replace("__LOGO_DARK_HIDE__", "" if logo_dark_b64 else "display:none")
    .replace("__MONDAY_TOKEN__", MONDAY_TOKEN)
)
OUT.write_text(html, encoding="utf-8")
print(f"\n✓  Written {OUT}  ({OUT.stat().st_size:,} bytes)")
print(f"   Files loaded : {len(file_names)}")
print(f"   Default file : {file_names[0]}")
