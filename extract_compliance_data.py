"""
Cross-references every form in the Blueprints folder (the full August list)
against what's actually finished in Compliance Reports, and extracts
everything the local dashboard needs: Dashboard KPIs, per-store results
(region/manager/rep/compliance), and the Not Captured list. Forms with no
finished report yet show up with status "pending". Re-run any time to
refresh -- it never touches the source files.
"""
import openpyxl
import glob
import shutil
import gzip
import os
import re
import csv
import json
from datetime import datetime

BLUEPRINT_FOLDER = r"C:\Users\CarinPillay\OneDrive - Meridian Group\Meridian Nexus - Documents\Capture\Capture\Blueprints"
REPORTS_FOLDER = r"C:\Users\CarinPillay\OneDrive - Meridian Group\Meridian Nexus - Documents\Capture\Capture\Compliance Reports"
FORMS_CONFIG = r"C:\Users\CarinPillay\OneDrive - Meridian Group\Meridian Nexus - Documents\Capture\Capture\Aspen Report Automation - Handover\forms-config.csv"
CALL_CYCLE_PATH = r"C:\Users\CarinPillay\OneDrive - Meridian Group\Meridian Nexus - Documents\Stock Fix\Call Cycle master - Stock Fix.xlsx"
OUT_PATH = os.path.join(os.path.dirname(__file__), "compliance_data.json")

# Folder inside the SharePoint-synced Meridian Nexus library that the dashboard
# reads from via Graph. Writing here is all the "publish" step needs -- OneDrive
# syncs it up, and auth.js reads it back with the signed-in user's token, so the
# data never sits publicly next to index.html.
PUBLISH_FOLDER = r"C:\Users\CarinPillay\OneDrive - Meridian Group\Meridian Nexus - Documents\Capture\Capture\Dashboard"

# Per-form store-level question answers. All forms together are ~870k rows,
# far too much for one payload, so each form gets its own file that the
# dashboard fetches only when that form is opened.
QUESTIONS_SUBFOLDER = "questions"


def load_output_name_overrides():
    """FormWorkbook -> OutputFileName from forms-config.csv. A blueprint's
    OutputFileName can be renamed (e.g. to dodge Windows' 260-char path
    limit) without matching the blueprint's own filename anymore, so this
    is the authoritative link between the two, not filename similarity."""
    overrides = {}
    if not os.path.exists(FORMS_CONFIG):
        return overrides
    with open(FORMS_CONFIG, encoding="utf-8-sig", newline="") as f:
        for row in csv.DictReader(f):
            workbook = (row.get("FormWorkbook") or "").strip()
            output = (row.get("OutputFileName") or "").strip()
            if workbook and output:
                overrides[normkey(workbook.replace(".xlsx", ""))] = normkey(
                    output.replace("_Store_Compliance_Report.xlsx", "").replace(".xlsx", "")
                )
    return overrides


def load_banner_lookup():
    """Banner per store, from the Call Cycle Master. The compliance reports have
    no banner field, and the store names are not reliably banner-prefixed
    ("BEL AIR SUPERSPAR"), so this is the authoritative source.

    The reports' "Store Code" is the GeoRep code (MER####), which matches the
    Call Cycle Master's GEO REP STORE CODE column -- not its STORE CODE column.
    The sheet has one row per store/resource, so codes repeat; banner is
    consistent across them and last-write-wins is fine.

    Returns ({geo_code: banner}, {upper_store_name: banner}); the name map is a
    fallback for the ~1% of rows whose code is missing from the master."""
    by_code, by_name = {}, {}
    if not os.path.exists(CALL_CYCLE_PATH):
        print(f"  WARNING: Call Cycle Master not found at {CALL_CYCLE_PATH}")
        print("           store rows will have banner=None")
        return by_code, by_name
    wb = openpyxl.load_workbook(CALL_CYCLE_PATH, data_only=True, read_only=True)
    try:
        ws = wb["Call Cycle Master"]
        rows = ws.iter_rows(values_only=True)
        hdr = next(rows)
        ix = {h: i for i, h in enumerate(hdr) if h}
        for col in ("BANNER", "GEO REP STORE CODE", "STORE NAME"):
            if col not in ix:
                print(f"  WARNING: Call Cycle Master has no '{col}' column; banner disabled")
                return {}, {}
        for r in rows:
            banner = r[ix["BANNER"]]
            if not banner:
                continue
            banner = str(banner).strip()
            geo = r[ix["GEO REP STORE CODE"]]
            name = r[ix["STORE NAME"]]
            if geo:
                by_code[str(geo).strip()] = banner
            if name:
                by_name[str(name).strip().upper()] = banner
    finally:
        wb.close()
    return by_code, by_name


def parse_period(period_line):
    """Pull the survey window out of the Dashboard A3 line, e.g.
    "Survey period: 2026-08-01 to 2026-08-31 | 8 Yes/No | ...". Returns
    (start, end) as ISO date strings, or (None, None) if absent."""
    m = re.search(r"(\d{4}-\d{2}-\d{2})\s*to\s*(\d{4}-\d{2}-\d{2})", period_line or "")
    return (m.group(1), m.group(2)) if m else (None, None)


def classify(period_line, label):
    m = re.search(r"(\d+) Yes/No \| (\d+) KPI/other \| (\d+) photo", period_line or "")
    yes_no_q, kpi_q, photo_q = (int(m.group(1)), int(m.group(2)), int(m.group(3))) if m else (None, None, None)
    label = (label or "").strip()
    if label == "Compliance N/A":
        kind = "price"
    elif label == "Target Compliance":
        kind = "sos"
    else:
        kind = "audit"
    mixed = kind == "audit" and yes_no_q is not None and kpi_q is not None and kpi_q >= yes_no_q and yes_no_q > 0
    return kind, mixed, yes_no_q, kpi_q, photo_q


def guess_client(name):
    m = re.match(r"^([A-Za-z0-9&' ]+?)\s*-", name)
    return (m.group(1).strip() if m else name.split()[0]).strip()


def normkey(name):
    return re.sub(r"[^a-z0-9]", "", name.lower())


def yes_no_slots(question_rows, yes_no_q):
    """Yes/No checks a fully-completed store should answer -- the per-store
    denominator. Prefers the Question Summary (authoritative, and excludes
    questions that never ran), falls back to the period line's count, then to 1
    for the single-question reports that have no Question Summary sheet."""
    scored = 0
    for q in (question_rows or []):
        y, n = q.get("yes"), q.get("no")
        try:
            if y is not None and n is not None and (float(y) + float(n)) > 0:
                scored += 1
        except (TypeError, ValueError):
            continue
    if scored:
        return scored
    if yes_no_q:
        return int(yes_no_q)
    return 1


def store_compliance(row, ci_compliance, ci_answer):
    """Per-store compliance. Multi-question reports carry a rate in "Store
    Compliance"; single-question reports carry the answer instead, which is a
    rate of 1 or 0. Anything else (not captured, refused, blank) stays None so
    it is excluded from averages rather than counted as a zero."""
    if ci_compliance is not None:
        return row[ci_compliance]
    if ci_answer is None:
        return None
    val = str(row[ci_answer] or "").strip().lower()
    if val == "yes":
        return 1
    if val == "no":
        return 0
    return None


def extract_store_results(ws, banner_by_code=None, banner_by_name=None):
    start = None
    hdr = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), start=1):
        if r and r[0] == "Store Code":
            start = i + 1
            hdr = r
            break
    if start is None:
        return []
    idx = {h: j for j, h in enumerate(hdr) if h}
    ci_region = idx.get("Region")
    ci_manager = idx.get("Assigned Manager")
    ci_rep = idx.get("Assigned Rep")
    ci_captured = idx.get("Captured")
    ci_compliance = idx.get("Store Compliance")
    ci_imgurl = idx.get("Image URL Source")
    # Single-question reports have no "Store Compliance" column -- with one
    # question there is no rate to average, so the per-store result is the
    # answer itself ("Latest Answer"/"Status"), and "No Reason" records why a
    # No happened. 14 of 117 reports are this shape.
    # Raw per-store check counts. Compliance counts every unanswered check as a
    # No, so the store's own rate (Yes / Answered) is not enough: a store that
    # answered 2 of 8 checks with both Yes is 25% compliant, not 100%.
    ci_yes_checks = idx.get("Yes Checks")
    ci_answered_checks = idx.get("Answered Checks")
    ci_answer = idx.get("Latest Answer")
    if ci_answer is None:
        ci_answer = idx.get("Status")
    ci_noreason = idx.get("No Reason")
    ci_store = idx.get("Store Code")
    ci_storename = idx.get("Store Name")
    rows = []
    for row in ws.iter_rows(min_row=start, values_only=True):
        if row[ci_store] is None:
            continue
        store_code = row[ci_store]
        store_name = row[ci_storename] if ci_storename is not None else None
        banner = None
        if banner_by_code:
            banner = banner_by_code.get(str(store_code).strip())
        if banner is None and banner_by_name and store_name:
            banner = banner_by_name.get(str(store_name).strip().upper())
        rows.append({
            "store": store_code,
            "storeName": store_name,
            "banner": banner,
            "region": row[ci_region] if ci_region is not None else None,
            "manager": row[ci_manager] if ci_manager is not None else None,
            "rep": row[ci_rep] if ci_rep is not None else None,
            "captured": row[ci_captured] if ci_captured is not None else None,
            "compliance": store_compliance(row, ci_compliance, ci_answer),
            "yesChecks": row[ci_yes_checks] if ci_yes_checks is not None else None,
            "answeredChecks": row[ci_answered_checks] if ci_answered_checks is not None else None,
            "answer": row[ci_answer] if ci_answer is not None else None,
            "noReason": row[ci_noreason] if ci_noreason is not None else None,
            "imageUrl": row[ci_imgurl] if ci_imgurl is not None else None,
        })
    return rows


def extract_question_summary(ws):
    """Per-question results from the Question Summary sheet: one row per question
    per form (~1,200 rows across the whole set). Only Yes/No questions carry a
    compliance percentage; Numeric KPI questions have a target instead, and
    Supporting/Other questions are informational, so compliance is left None
    rather than being faked as 0."""
    start = None
    hdr = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True), start=1):
        if r and r[0] == "Question Group":
            start = i + 1
            hdr = r
            break
    if start is None:
        return []
    idx = {h: j for j, h in enumerate(hdr) if h}
    rows = []
    for row in ws.iter_rows(min_row=start, values_only=True):
        question = row[idx["Question"]] if "Question" in idx else None
        if not question:
            continue
        def val(name):
            j = idx.get(name)
            return row[j] if j is not None else None
        rows.append({
            "group": val("Question Group"),
            "type": val("Question Type"),
            "question": question,
            "target": val("Target Stores"),
            "answered": val("Answered"),
            "yes": val("Yes"),
            "no": val("No"),
            "compliance": val("Yes/No Compliance"),
            "numericTarget": val("Numeric Target"),
            # Numeric KPI questions have no Yes/No rate; their result is the
            # mean answer across stores ("Average"), plus a target-met rate
            # where the question defines a target.
            "average": val("Average"),
            "metTarget": val("Met Target"),
            "targetCompliance": val("Target Compliance"),
        })
    return rows


def extract_question_results(ws):
    """Store-level answers, one row per store per question, from the Question
    Results sheet. Kept out of the main payload (~870k rows across all forms)
    and written per form instead."""
    start = None
    hdr = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=8, values_only=True), start=1):
        if r and r[0] == "Store Code":
            start = i + 1
            hdr = r
            break
    if start is None:
        return []
    idx = {h: j for j, h in enumerate(hdr) if h}

    def col(name):
        return idx.get(name)

    ci = {k: col(v) for k, v in {
        "store": "Store Code", "storeName": "Store Name", "region": "Region",
        "manager": "Assigned Manager", "rep": "Assigned Rep", "captured": "Captured",
        "group": "Question Group", "question": "Question", "type": "Question Type",
        "answer": "Answer", "isYes": "Is Yes", "isNo": "Is No",
        "target": "Numeric Target", "targetMet": "Target Met", "image": "Image URL Source",
    }.items()}

    rows = []
    for row in ws.iter_rows(min_row=start, values_only=True):
        if row[ci["store"]] is None and row[ci["question"]] is None:
            continue
        rec = {}
        for key, j in ci.items():
            rec[key] = row[j] if j is not None else None
        rows.append(rec)
    return rows


def group_slot_totals(question_rows, question_results, store_rows):
    """Exact Yes-answer and check-slot totals per region and per banner.

    Computed here rather than in the browser because the store-level answers are
    ~870k rows across 117 forms: aggregating them client-side would mean
    downloading every per-form file. The result is ~10 regions and ~60 banners
    per form, which is negligible in the payload.

    Slots are (targeted stores in the group) x (scored Yes/No questions), so a
    store that was never visited still occupies its slots and counts as No --
    the same rule as the form-level figure. Yes comes from the store-level
    answers, which is the only source that carries region and banner.
    """
    # A check is anything a store can pass or fail: a Yes/No question, or a
    # Numeric KPI question that carries a target. SOS forms are entirely the
    # latter (18 KPI questions, no Yes/No), so counting only Yes/No left every
    # region with zero slots and a blank percentage.
    scored = set()
    for q in (question_rows or []):
        y, n = q.get("yes"), q.get("no")
        try:
            if y is not None and n is not None and (float(y) + float(n)) > 0:
                scored.add(q.get("question"))
                continue
        except (TypeError, ValueError):
            pass
        tc, nt = q.get("targetCompliance"), q.get("numericTarget")
        if (tc is not None and tc != "") or (nt is not None and nt != ""):
            scored.add(q.get("question"))
    n_scored = len(scored)

    # Region/banner per store, and the targeted store counts per group.
    region_of, banner_of = {}, {}
    region_target, banner_target = {}, {}
    region_captured, banner_captured = {}, {}
    for r in store_rows or []:
        code = r.get("store")
        reg = r.get("region") or "Unassigned"
        ban = r.get("banner") or "Unknown"
        region_of[code] = reg
        banner_of[code] = ban
        region_target[reg] = region_target.get(reg, 0) + 1
        banner_target[ban] = banner_target.get(ban, 0) + 1
        if r.get("captured") == "Yes":
            region_captured[reg] = region_captured.get(reg, 0) + 1
            banner_captured[ban] = banner_captured.get(ban, 0) + 1

    region_yes, banner_yes = {}, {}
    if n_scored:
        for row in question_results or []:
            if row.get("question") not in scored:
                continue
            # Passing a check is either answering Yes or meeting the target.
            if not (_truthy(row.get("isYes")) or _truthy(row.get("targetMet"))):
                continue
            code = row.get("store")
            reg = region_of.get(code, row.get("region") or "Unassigned")
            ban = banner_of.get(code, "Unknown")
            region_yes[reg] = region_yes.get(reg, 0) + 1
            banner_yes[ban] = banner_yes.get(ban, 0) + 1

    def build(targets, captures, yeses):
        out = {}
        for key, target in targets.items():
            out[key] = {
                "target": target,
                "captured": captures.get(key, 0),
                "yes": yeses.get(key, 0),
                "slots": target * n_scored,
            }
        return out

    return build(region_target, region_captured, region_yes), \
           build(banner_target, banner_captured, banner_yes)


def _truthy(v):
    """Excel exports Is Yes as TRUE/True/1/"Yes" depending on the writer."""
    if v is None:
        return False
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v == 1
    return str(v).strip().lower() in ("true", "yes", "1", "y")


def slug(name):
    """Filename-safe, stable per form so the dashboard can build the URL from
    the form name without a lookup table."""
    return re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")


def extract_not_captured(ws):
    start = None
    hdr = None
    for i, r in enumerate(ws.iter_rows(min_row=1, max_row=6, values_only=True), start=1):
        if r and r[0] == "Store Code":
            start = i + 1
            hdr = r
            break
    if start is None:
        return []
    idx = {h: j for j, h in enumerate(hdr) if h}
    rows = []
    for row in ws.iter_rows(min_row=start, values_only=True):
        if row[idx.get("Store Code", 0)] is None:
            continue
        rows.append({
            "store": row[idx.get("Store Code")],
            "storeName": row[idx.get("Store Name")] if "Store Name" in idx else None,
            "region": row[idx.get("Region")] if "Region" in idx else None,
            "manager": row[idx.get("Assigned Manager")] if "Assigned Manager" in idx else None,
            "rep": row[idx.get("Assigned Rep")] if "Assigned Rep" in idx else None,
        })
    return rows


def main():
    pending_question_files = []
    banner_by_code, banner_by_name = load_banner_lookup()
    print(f"Banner lookup: {len(banner_by_code)} store codes, {len(banner_by_name)} store names")
    output_overrides = load_output_name_overrides()
    blueprint_names = []
    for path in sorted(glob.glob(os.path.join(BLUEPRINT_FOLDER, "*.xlsx"))):
        name = os.path.basename(path)
        if name.startswith("~$"):
            continue
        blueprint_names.append(name.replace(".xlsx", ""))

    report_files = {}
    for path in glob.glob(os.path.join(REPORTS_FOLDER, "*.xlsx")):
        name = os.path.basename(path)
        if name.startswith("~$"):
            continue
        clean = name.replace("_Store_Compliance_Report.xlsx", "").replace(".xlsx", "")
        report_files[normkey(clean)] = path

    results = []
    errors = []
    for blueprint_name in blueprint_names:
        key = output_overrides.get(normkey(blueprint_name), normkey(blueprint_name))
        report_path = report_files.get(key)

        if report_path is None:
            results.append({
                "name": blueprint_name, "client": guess_client(blueprint_name),
                "status": "pending", "kind": None, "mixed": False, "complianceLabel": None,
                "target": None, "captured": None, "coverage": None, "compliance": None,
                "notCaptured": None, "yesNoQ": None, "kpiQ": None, "photoQ": None, "yesNoSlots": None,
                "regionTotals": {}, "bannerTotals": {},
                "periodStart": None, "periodEnd": None,
                "modifiedAt": None, "storeRows": [], "questionRows": [],
                "questionFile": None, "questionResultCount": 0,
            })
            continue

        try:
            wb = openpyxl.load_workbook(report_path, data_only=True, read_only=True)
            if "Dashboard" not in wb.sheetnames:
                errors.append({"name": blueprint_name, "error": "No Dashboard sheet"})
                continue
            ws = wb["Dashboard"]
            period_line = ws["A3"].value
            label = ws["J5"].value
            kind, mixed, yes_no_q, kpi_q, photo_q = classify(period_line, label)
            period_start, period_end = parse_period(period_line)
            question_rows = extract_question_summary(wb["Question Summary"]) if "Question Summary" in wb.sheetnames else []
            store_rows = extract_store_results(wb["Store Results"], banner_by_code, banner_by_name) if "Store Results" in wb.sheetnames else []
            not_captured_rows = extract_not_captured(wb["Not Captured"]) if "Not Captured" in wb.sheetnames else []
            question_results = extract_question_results(wb["Question Results"]) if "Question Results" in wb.sheetnames else []
            region_totals, banner_totals = group_slot_totals(question_rows, question_results, store_rows)

            results.append({
                "name": blueprint_name, "client": guess_client(blueprint_name),
                "status": "complete", "kind": kind, "mixed": mixed, "complianceLabel": label,
                "target": ws["A6"].value, "captured": ws["D6"].value, "coverage": ws["G6"].value,
                "compliance": ws["J6"].value, "notCaptured": ws["G10"].value,
                "yesNoQ": yes_no_q, "kpiQ": kpi_q, "photoQ": photo_q,
                "yesNoSlots": yes_no_slots(question_rows, yes_no_q),
                "regionTotals": region_totals, "bannerTotals": banner_totals,
                "periodStart": period_start, "periodEnd": period_end,
                "modifiedAt": datetime.fromtimestamp(os.path.getmtime(report_path)).isoformat(timespec="seconds"),
                "storeRows": store_rows,
                "questionRows": question_rows,
                "questionFile": (slug(blueprint_name) + ".json") if question_results else None,
                "questionResultCount": len(question_results),
            })
            if question_results:
                pending_question_files.append((slug(blueprint_name) + ".json", question_results))
        except PermissionError:
            results.append({
                "name": blueprint_name, "client": guess_client(blueprint_name),
                "status": "locked", "kind": None, "mixed": False, "complianceLabel": None,
                "target": None, "captured": None, "coverage": None, "compliance": None,
                "notCaptured": None, "yesNoQ": None, "kpiQ": None, "photoQ": None, "yesNoSlots": None,
                "regionTotals": {}, "bannerTotals": {},
                "periodStart": None, "periodEnd": None,
                "modifiedAt": None, "storeRows": [], "questionRows": [],
                "questionFile": None, "questionResultCount": 0,
            })
        except Exception as e:
            errors.append({"name": blueprint_name, "error": str(e)[:200]})

    out = {
        "generatedAt": datetime.now().isoformat(timespec="seconds"),
        "campaigns": results,
        "errors": errors,
    }
    # Written gzipped. The payload compresses ~90%, and neither python's
    # http.server nor SharePoint/Graph applies transport compression to a .json
    # file, so compressing the file itself is the only way to get the win on
    # both paths. The dashboard inflates it with DecompressionStream.
    def write_json_gz(path, obj):
        with gzip.open(path, "wt", encoding="utf-8", compresslevel=6) as f:
            json.dump(obj, f, separators=(",", ":"), default=str)

    write_json_gz(OUT_PATH + ".gz", out)

    # Per-form question files, written next to the payload in both locations.
    local_q_dir = os.path.join(os.path.dirname(__file__), QUESTIONS_SUBFOLDER)
    os.makedirs(local_q_dir, exist_ok=True)
    for fname, rows in pending_question_files:
        write_json_gz(os.path.join(local_q_dir, fname + ".gz"), {"rows": rows})

    # Publish to the SharePoint-synced folder. Non-fatal: if OneDrive is not
    # available the local copy is still written, so the dashboard keeps working
    # locally and the operator sees exactly why publishing was skipped.
    try:
        os.makedirs(os.path.join(PUBLISH_FOLDER, QUESTIONS_SUBFOLDER), exist_ok=True)
        shutil.copy2(OUT_PATH + ".gz", os.path.join(PUBLISH_FOLDER, "compliance_data.json.gz"))
        for fname, _ in pending_question_files:
            shutil.copy2(os.path.join(local_q_dir, fname + ".gz"),
                         os.path.join(PUBLISH_FOLDER, QUESTIONS_SUBFOLDER, fname + ".gz"))
        print(f"Published to {PUBLISH_FOLDER}")
    except Exception as e:
        print(f"  WARNING: could not publish to SharePoint folder: {e}")
        print("           local copy written; dashboard will work on localhost only")

    complete = sum(1 for r in results if r["status"] == "complete")
    pending = sum(1 for r in results if r["status"] == "pending")
    locked = sum(1 for r in results if r["status"] == "locked")
    gz_mb = os.path.getsize(OUT_PATH + ".gz") / 1048576
    print(f"Wrote {OUT_PATH}.gz  ({gz_mb:.1f} MB gzipped)")
    print(f"  {len(results)} total forms | {complete} complete, {pending} pending, {locked} locked, {len(errors)} errored")
    q_rows = sum(len(res.get("questionRows") or []) for res in results)
    print(f"  {q_rows} question rows extracted")
    qr_total = sum(len(r) for _, r in pending_question_files)
    q_bytes = sum(os.path.getsize(os.path.join(local_q_dir, f + ".gz")) for f, _ in pending_question_files)
    print(f"  {qr_total} store-level answers across {len(pending_question_files)} per-form file(s), "
          f"{q_bytes / 1048576:.1f} MB gzipped total")
    all_rows = [r for res in results for r in res["storeRows"]]
    with_banner = sum(1 for r in all_rows if r.get("banner"))
    if all_rows:
        print(f"  {with_banner}/{len(all_rows)} store rows matched a banner "
              f"({with_banner / len(all_rows) * 100:.1f}%)")
    for e in errors:
        print("   -", e["name"], "-", e["error"])


if __name__ == "__main__":
    main()
