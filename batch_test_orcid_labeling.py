"""
batch_test_orcid_labeling.py — Batch test for ORCID-based author labeling.

Scans a folder for <jid>_<aid>_CLN_AS.docx (or _AS.docx) + <jid>_<aid>.json pairs.
Uses labelAuthorsNER to parse authors, then rebuilds the AU paragraph with:
  - Author names (original formatting)
  - Superscript affiliations
  - [INSERT ORCID iD LOGO] in RED BOLD after authors who have ORCID in JSON
  - [GQ: N] query callout in RED BOLD at the end (if ORCID authors exist)
  - Commas / "and" separators

No [AU] labels. No aulabel style. Just the ORCID insertion text.

Usage:
    python batch_test_orcid_labeling.py "V:\\AU_TEST"
    python batch_test_orcid_labeling.py "V:\\AU_TEST" --output "V:\\AU_TEST\\results"

Output:
    - <jid>_<aid>_*_LABELED.docx  (labeled docx files)
    - test_report.html            (HTML report)
"""

import sys
import os
import re
import json
import glob
import traceback
from datetime import datetime

from docx import Document

from labelAuthorsNER import (
    AuthorLabeler,
    AuthorEntity,
    LabelResult,
    extract_au_runs,
    load_json_authors,
    is_degree_token,
    write_labeled_docx,
    ORCID_INSERT_TEXT,
)


# ============================================================================
# File discovery
# ============================================================================

def find_test_pairs(folder):
    pairs = []
    for pattern_suffix in ["*_CLN_AS.docx", "*_AS.docx"]:
        pat = os.path.join(folder, pattern_suffix)
        for docx_path in sorted(glob.glob(pat)):
            basename = os.path.basename(docx_path)
            match = re.match(r'^(.+?)(?:_CLN)?_AS\.docx$', basename, re.IGNORECASE)
            if not match:
                continue
            jid_aid = match.group(1)
            if any(j == jid_aid for _, _, j in pairs):
                continue
            json_name = f"{jid_aid}.json"
            json_path = os.path.join(folder, json_name)
            if os.path.exists(json_path):
                pairs.append((docx_path, json_path, jid_aid))
            else:
                pairs.append((docx_path, None, jid_aid))
    return pairs


def get_au_text_from_docx(docx_path):
    try:
        doc = Document(docx_path)
        for p in doc.paragraphs:
            if p.style.name in ("AU", "AU0"):
                return p.text
    except Exception:
        pass
    return ""


# ============================================================================
# Process one file pair
# ============================================================================

def process_one(docx_path, json_path, output_folder):
    result = {
        "input_file": os.path.basename(docx_path),
        "validation": "FAIL",
        "input_au_line": "",
        "output_au_line": "",
        "json_count": 0,
        "output_count": 0,
        "orcid_count": 0,
        "details": [],
        "error": "",
    }

    try:
        # 1. Read original AU line
        input_au = get_au_text_from_docx(docx_path)
        result["input_au_line"] = input_au

        # 2. Load JSON
        if not (json_path and os.path.exists(json_path)):
            result["validation"] = "NO JSON"
            result["error"] = "JSON file not found"
            return result

        with open(json_path) as f:
            jdata = json.load(f)
        json_authors = jdata.get("authors_info", {})
        json_count = len(json_authors)
        result["json_count"] = json_count

        # Count ORCID authors in JSON
        json_orcid_count = sum(1 for info in json_authors.values()
                               if info.get("orcid", "").strip())

        # 3. Use write_labeled_docx from labelAuthorsNER (production function)
        base = os.path.basename(docx_path)
        name, ext = os.path.splitext(base)
        labeled_name = f"{name}_LABELED{ext}"
        labeled_path = os.path.join(output_folder, labeled_name)

        ok, labeled_path, label_result = write_labeled_docx(
            docx_path, json_path, output_path=labeled_path
        )

        result["output_count"] = len(label_result.authors)
        result["details"] = label_result.match_details

        if not ok or not label_result.authors:
            result["validation"] = "FAIL"
            result["error"] = "No authors parsed or no AU paragraph"
            return result

        # Count how many parsed authors got ORCID from JSON
        parsed_orcid_count = sum(1 for a in label_result.authors if a.orcid)
        result["orcid_count"] = parsed_orcid_count

        # 4. Read back and verify
        output_au = get_au_text_from_docx(labeled_path)
        result["output_au_line"] = output_au

        # 5. Validate
        errors = []
        warnings = []

        # CRITICAL: Original text must be preserved exactly
        out_stripped = output_au.replace(ORCID_INSERT_TEXT, "")
        out_stripped = re.sub(r'\[GQ: \d+\]', '', out_stripped)
        if out_stripped != input_au:
            errors.append("Original text NOT preserved — characters added or removed")

        # Count check
        if len(label_result.authors) != json_count:
            warnings.append(f"Author count: parsed {len(label_result.authors)} vs JSON {json_count}")

        # ORCID insertion count vs JSON ORCID count
        orcid_inserts_in_output = output_au.count(ORCID_INSERT_TEXT)
        if orcid_inserts_in_output != json_orcid_count:
            warnings.append(f"ORCID inserts: {orcid_inserts_in_output} in docx vs {json_orcid_count} in JSON")

        # Degree leak
        for a in label_result.authors:
            if is_degree_token(a.full_name):
                warnings.append(f"Degree '{a.full_name}' misidentified as author")

        if errors:
            result["validation"] = "FAIL"
            result["error"] = "; ".join(errors)
        elif warnings:
            result["validation"] = "WARN"
            result["error"] = "; ".join(warnings)
        else:
            result["validation"] = "PASS"

    except Exception as e:
        result["validation"] = "ERROR"
        result["error"] = f"{type(e).__name__}: {str(e)}"
        result["details"].append(traceback.format_exc())

    return result


# ============================================================================
# HTML Report
# ============================================================================

def escape_html(text):
    if not text:
        return ""
    return (text.replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def highlight_output(text):
    """Highlight [INSERT ORCID iD LOGO] in red and [GQ:N] in red."""
    if not text:
        return ""
    escaped = escape_html(text)
    # Highlight ORCID insert text
    escaped = escaped.replace(
        escape_html(ORCID_INSERT_TEXT),
        f'<span style="color:#FF0000;font-weight:bold;">{escape_html(ORCID_INSERT_TEXT)}</span>'
    )
    # Highlight [GQ: N]
    escaped = re.sub(
        r'\[GQ:\s*\d+\]',
        lambda m: f'<span style="color:#FF0000;font-weight:bold;">{m.group()}</span>',
        escaped
    )
    return escaped


def generate_html_report(results, input_folder, output_path):
    pass_count = sum(1 for r in results if r["validation"] == "PASS")
    warn_count = sum(1 for r in results if r["validation"] == "WARN")
    fail_count = sum(1 for r in results if r["validation"] == "FAIL")
    error_count = sum(1 for r in results if r["validation"] == "ERROR")
    total = len(results)
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    html = f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<title>ORCID Author Labeling Test Report</title>
<style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{ font-family: 'Segoe UI', Tahoma, sans-serif; background: #f5f7fa; color: #333; padding: 20px; }}
    .header {{ background: linear-gradient(135deg, #1a237e, #283593); color: white; padding: 24px 32px; border-radius: 10px; margin-bottom: 20px; }}
    .header h1 {{ font-size: 22px; margin-bottom: 8px; }}
    .header .meta {{ font-size: 13px; opacity: 0.85; }}
    .summary {{ display: flex; gap: 16px; margin-bottom: 20px; }}
    .summary .card {{ flex: 1; padding: 16px 20px; border-radius: 8px; color: white; font-weight: 600; text-align: center; }}
    .card.total {{ background: #455a64; }}
    .card.pass {{ background: #2e7d32; }}
    .card.warn {{ background: #f57f17; }}
    .card.fail {{ background: #c62828; }}
    .card.error {{ background: #e65100; }}
    .card .num {{ font-size: 28px; display: block; }}
    .card .label {{ font-size: 12px; text-transform: uppercase; letter-spacing: 1px; }}
    table {{ width: 100%; border-collapse: collapse; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 1px 4px rgba(0,0,0,0.1); }}
    thead {{ background: #263238; color: white; }}
    th {{ padding: 12px 14px; text-align: left; font-size: 12px; text-transform: uppercase; letter-spacing: 0.5px; white-space: nowrap; }}
    td {{ padding: 10px 14px; border-bottom: 1px solid #e0e0e0; font-size: 13px; vertical-align: top; }}
    tr:hover {{ background: #f5f5f5; }}
    .au-line {{ max-width: 480px; max-height: 80px; overflow: auto; font-size: 12px; line-height: 1.5; word-break: break-word; white-space: pre-wrap; }}
    .badge {{ display: inline-block; padding: 3px 10px; border-radius: 4px; font-size: 11px; font-weight: 700; text-transform: uppercase; }}
    .badge.pass {{ background: #e8f5e9; color: #2e7d32; }}
    .badge.warn {{ background: #fff8e1; color: #f57f17; }}
    .badge.fail {{ background: #ffebee; color: #c62828; }}
    .badge.error {{ background: #fff3e0; color: #e65100; }}
    .badge.nojson {{ background: #fce4ec; color: #880e4f; }}
    .count-match {{ color: #2e7d32; font-weight: 600; }}
    .count-mismatch {{ color: #c62828; font-weight: 600; }}
    .error-text {{ color: #c62828; font-size: 11px; margin-top: 4px; }}
    .warn-text {{ color: #f57f17; font-size: 11px; margin-top: 4px; }}
    .details-toggle {{ cursor: pointer; color: #1565c0; font-size: 11px; text-decoration: underline; }}
    .details-content {{ display: none; margin-top: 6px; padding: 8px; background: #f5f5f5; border-radius: 4px; font-size: 11px; font-family: monospace; max-height: 200px; overflow: auto; white-space: pre-wrap; }}
    .orcid-tag {{ background: #fff3e0; color: #e65100; padding: 1px 5px; border-radius: 3px; font-size: 11px; font-weight: 600; }}
</style>
<script>
function toggleDetails(id) {{ var el=document.getElementById(id); el.style.display=el.style.display==='none'?'block':'none'; }}
</script>
</head>
<body>
<div class="header">
    <h1>ORCID Author Labeling Test Report</h1>
    <div class="meta">
        Module: <strong>labelAuthorsNER</strong> &nbsp;|&nbsp;
        Input: <strong>{escape_html(input_folder)}</strong> &nbsp;|&nbsp;
        Generated: <strong>{timestamp}</strong> &nbsp;|&nbsp;
        Insert text: <strong style="color:#ffcdd2;">{escape_html(ORCID_INSERT_TEXT)}</strong>
    </div>
</div>
<div class="summary">
    <div class="card total"><span class="num">{total}</span><span class="label">Total</span></div>
    <div class="card pass"><span class="num">{pass_count}</span><span class="label">Pass</span></div>
    <div class="card warn"><span class="num">{warn_count}</span><span class="label">Warn</span></div>
    <div class="card fail"><span class="num">{fail_count}</span><span class="label">Fail</span></div>
    <div class="card error"><span class="num">{error_count}</span><span class="label">Error</span></div>
</div>
<table>
<thead><tr>
    <th>No.</th><th>Input File</th><th>Result</th>
    <th>Input Author Line</th><th>Output Author Line</th>
    <th>Json Count</th><th>Output Count</th>
</tr></thead>
<tbody>
"""

    for idx, r in enumerate(results, 1):
        v = r["validation"]
        badge_cls = {"PASS":"pass","WARN":"warn","FAIL":"fail","ERROR":"error"}.get(v, "nojson")
        badge = f'<span class="badge {badge_cls}">{escape_html(v)}</span>'

        # ORCID count tag
        oc = r.get("orcid_count", 0)
        orcid_tag = f' <span class="orcid-tag">ORCID:{oc}</span>' if oc > 0 else ''

        error_html = ""
        if r["error"]:
            cls = "warn-text" if v == "WARN" else "error-text"
            error_html = f'<div class="{cls}">{escape_html(r["error"])}</div>'

        details_html = ""
        if r["details"]:
            did = f"d{idx}"
            details_html = (f'<span class="details-toggle" onclick="toggleDetails(\'{did}\')">details</span>'
                            f'<div class="details-content" id="{did}">{escape_html(chr(10).join(r["details"]))}</div>')

        jc = r["json_count"]
        outc = r["output_count"]
        cc = "count-match" if jc == outc and jc > 0 else "count-mismatch" if jc > 0 else ""

        input_html = f'<div class="au-line">{escape_html(r["input_au_line"])}</div>'
        output_html = f'<div class="au-line">{highlight_output(r["output_au_line"])}</div>'

        html += f"""<tr>
<td>{idx}</td>
<td><strong>{escape_html(r["input_file"])}</strong>{orcid_tag}{error_html}{details_html}</td>
<td>{badge}</td>
<td>{input_html}</td>
<td>{output_html}</td>
<td style="text-align:center">{jc}</td>
<td style="text-align:center" class="{cc}">{outc}</td>
</tr>\n"""

    html += """</tbody></table></body></html>"""

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(html)
    return output_path


# ============================================================================
# Main
# ============================================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python batch_test_orcid_labeling.py <input_folder> [--output <output_folder>]")
        print()
        print("Scans for *_AS.docx + *.json pairs.")
        print("Inserts [INSERT ORCID iD LOGO] in RED BOLD for authors with ORCID.")
        print("Produces labeled docx files and test_report.html.")
        sys.exit(1)

    input_folder = sys.argv[1]
    output_folder = input_folder
    if "--output" in sys.argv:
        oi = sys.argv.index("--output")
        if oi + 1 < len(sys.argv):
            output_folder = sys.argv[oi + 1]

    if not os.path.isdir(input_folder):
        print(f"ERROR: '{input_folder}' is not a directory")
        sys.exit(1)

    os.makedirs(output_folder, exist_ok=True)
    pairs = find_test_pairs(input_folder)
    if not pairs:
        print(f"No test files found in '{input_folder}'")
        sys.exit(1)

    print(f"Found {len(pairs)} test file(s) in '{input_folder}'")
    print(f"Output: '{output_folder}'")
    print()

    results = []
    for i, (docx_path, json_path, jid_aid) in enumerate(pairs, 1):
        tag = jid_aid.replace("_", " ")
        print(f"[{i}/{len(pairs)}] {tag} ... ", end="", flush=True)
        r = process_one(docx_path, json_path, output_folder)
        results.append(r)
        oc = r.get("orcid_count", 0)
        orcid_str = f" orcid={oc}" if oc else ""
        print(f"{r['validation']} ({r['output_count']}/{r['json_count']}{orcid_str})")

    report_path = os.path.join(output_folder, "test_report.html")
    generate_html_report(results, input_folder, report_path)

    pc = sum(1 for r in results if r["validation"] == "PASS")
    wc = sum(1 for r in results if r["validation"] == "WARN")
    fc = sum(1 for r in results if r["validation"] == "FAIL")
    ec = sum(1 for r in results if r["validation"] == "ERROR")
    print()
    print("=" * 60)
    print(f"  TOTAL: {len(results)}  |  PASS: {pc}  |  WARN: {wc}  |  FAIL: {fc}  |  ERROR: {ec}")
    print(f"  Report: {report_path}")
    print("=" * 60)


if __name__ == "__main__":
    main()
