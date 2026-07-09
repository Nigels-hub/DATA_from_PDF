#!/usr/bin/env python3
"""
prepare_subgroup_data.py
─────────────────────────────────────────────────────────────────────────────
Step 2 in the pipeline (runs after extract_tables.py, before
interaction_pvalues_app.R).

Reads all "*by Subgroups*.xlsx" files from extracted_tables/, extracts:
  - title            : full table title (cell A1)
  - table_number     : 4-digit numeric component from the table ID
  - domain           : "Efficacy" (1000–1999), "Safety" (2000–2999),
                       "PRO"      (3000–3999)
  - endpoint_short   : short endpoint label derived from the title.
                       Edit ENDPOINT_SHORT_MAP below to add or override any
                       auto-generated value.
  - category         : subgroup category  (e.g. "Sex", "Age Group")
  - subgroup_values  : the levels compared (e.g. "Male, Female")
  - interaction_p    : interaction p-value (numeric)

Output: subgroup_interaction_pvalues.csv  (in the current working directory)

Usage:
    python3 prepare_subgroup_data.py [TABLES_DIR] [OUTPUT_CSV]

Defaults:
    TABLES_DIR  = extracted_tables/
    OUTPUT_CSV  = subgroup_interaction_pvalues.csv
"""

import sys
import re
import csv
from pathlib import Path

import openpyxl


# ─────────────────────────────────────────────────────────────────────────────
# Short endpoint name mapping
#
# Keys are substrings matched case-insensitively against the full table title.
# Values are the short labels that will appear in the Shiny app.
# The FIRST matching entry wins, so order matters (more specific entries first).
# Add / edit entries here to customise labels for any endpoint.
# ─────────────────────────────────────────────────────────────────────────────
ENDPOINT_SHORT_MAP: list[tuple[str, str]] = [
    # ── Efficacy (table numbers 1000–1999) ───────────────────────────────────
    ("Overall Survival",                         "OS"),
    ("Progression-Free Survival (IRC)",          "PFS \u2013 IRC"),
    ("Progression-Free Survival (INV)",          "PFS \u2013 INV"),

    # ── Safety top-level (table numbers 2000–2999) ───────────────────────────
    ("TEAEs by Subgroups - Safety",              "Any TEAE"),
    ("Severe TEAEs by Subgroups - Safety",       "Severe TEAE (overall)"),
    ("TESAEs by Subgroups - Safety",             "TESAE (overall)"),
    ("Permanent Treatment Discontinuation due to TEAEs by Subgroups",
                                                 "DC due to TEAE (overall)"),
    # AESI-level safety entries (keep before generic PT entries)
    ("TEAEs leading to Death by Subgroups - Abdominal Pain (AESI)",      "Death \u2013 Abd. Pain (AESI)"),
    ("TEAEs leading to Death by Subgroups - Anemia (AESI)",              "Death \u2013 Anemia (AESI)"),
    ("TEAEs leading to Death by Subgroups - Hypersensitivity",           "Death \u2013 Hypersens. (AESI)"),
    ("TEAEs leading to Death by Subgroups - Infusion Related",           "Death \u2013 IRR (AESI)"),
    ("TEAEs leading to Death by Subgroups - Nausea (AESI)",              "Death \u2013 Nausea (AESI)"),
    ("TEAEs leading to Death by Subgroups - Neutropenia (AESI)",         "Death \u2013 Neutropenia (AESI)"),
    ("TEAEs leading to Death by Subgroups - Vomiting (AESI)",            "Death \u2013 Vomiting (AESI)"),
    ("TEAEs leading to Permanent Treatment Discontinuation by Subgroups - Abdominal Pain (AESI)",  "DC \u2013 Abd. Pain (AESI)"),
    ("TEAEs leading to Permanent Treatment Discontinuation by Subgroups - Anemia",                 "DC \u2013 Anemia (AESI)"),
    ("TEAEs leading to Permanent Treatment Discontinuation by Subgroups - Hypersensitivity",       "DC \u2013 Hypersens. (AESI)"),
    ("TEAEs leading to Permanent Treatment Discontinuation by Subgroups - Infusion Related",       "DC \u2013 IRR (AESI)"),
    ("TEAEs leading to Permanent Treatment Discontinuation by Subgroups - Nausea",                 "DC \u2013 Nausea (AESI)"),
    ("TEAEs leading to Permanent Treatment Discontinuation by Subgroups - Neutropenia",            "DC \u2013 Neutropenia (AESI)"),
    ("TEAEs leading to Permanent Treatment Discontinuation by Subgroups - Vomiting",               "DC \u2013 Vomiting (AESI)"),
    ("Severe TEAEs by Subgroups - Abdominal Pain (AESI)",                "Severe \u2013 Abd. Pain (AESI)"),
    ("Severe TEAEs by Subgroups - Anemia (AESI)",                        "Severe \u2013 Anemia (AESI)"),
    ("Severe TEAEs by Subgroups - Hypersensitivity Reactions (AESI)",    "Severe \u2013 Hypersens. (AESI)"),
    ("Severe TEAEs by Subgroups - Infusion Related Reaction (AESI)",     "Severe \u2013 IRR (AESI)"),
    ("Severe TEAEs by Subgroups - Nausea (AESI)",                        "Severe \u2013 Nausea (AESI)"),
    ("Severe TEAEs by Subgroups - Neutropenia (AESI)",                   "Severe \u2013 Neutropenia (AESI)"),
    ("Severe TEAEs by Subgroups - Vomiting (AESI)",                      "Severe \u2013 Vomiting (AESI)"),
    ("TESAEs by Subgroups - Abdominal Pain (AESI)",                      "TESAE \u2013 Abd. Pain (AESI)"),
    ("TESAEs by Subgroups - Anemia (AESI)",                              "TESAE \u2013 Anemia (AESI)"),
    ("TESAEs by Subgroups - Hypersensitivity Reactions (AESI)",          "TESAE \u2013 Hypersens. (AESI)"),
    ("TESAEs by Subgroups - Infusion Related Reaction (AESI)",           "TESAE \u2013 IRR (AESI)"),
    ("TESAEs by Subgroups - Nausea (AESI)",                              "TESAE \u2013 Nausea (AESI)"),
    ("TESAEs by Subgroups - Neutropenia (AESI)",                         "TESAE \u2013 Neutropenia (AESI)"),
    ("TESAEs by Subgroups - Vomiting (AESI)",                            "TESAE \u2013 Vomiting (AESI)"),
    ("Non-Severe TEAEs by Subgroups - Abdominal Pain (AESI)",            "Non-Sev \u2013 Abd. Pain (AESI)"),
    ("Non-Severe TEAEs by Subgroups - Anemia (AESI)",                    "Non-Sev \u2013 Anemia (AESI)"),
    ("Non-Severe TEAEs by Subgroups - Hypersensitivity Reactions (AESI)", "Non-Sev \u2013 Hypersens. (AESI)"),
    ("Non-Severe TEAEs by Subgroups - Infusion Related Reaction (AESI)", "Non-Sev \u2013 IRR (AESI)"),
    ("Non-Severe TEAEs by Subgroups - Nausea (AESI)",                    "Non-Sev \u2013 Nausea (AESI)"),
    ("Non-Severe TEAEs by Subgroups - Neutropenia (AESI)",               "Non-Sev \u2013 Neutropenia (AESI)"),
    ("Non-Severe TEAEs by Subgroups - Vomiting (AESI)",                  "Non-Sev \u2013 Vomiting (AESI)"),
    ("TEAEs by Subgroups - Abdominal Pain (AESI)",                       "TEAE \u2013 Abd. Pain (AESI)"),
    ("TEAEs by Subgroups - Anemia (AESI)",                               "TEAE \u2013 Anemia (AESI)"),
    ("TEAEs by Subgroups - Hypersensitivity Reactions (AESI)",           "TEAE \u2013 Hypersens. (AESI)"),
    ("TEAEs by Subgroups - Infusion Related Reaction (AESI)",            "TEAE \u2013 IRR (AESI)"),
    ("TEAEs by Subgroups - Nausea (AESI)",                               "TEAE \u2013 Nausea (AESI)"),
    ("TEAEs by Subgroups - Neutropenia (AESI)",                          "TEAE \u2013 Neutropenia (AESI)"),
    ("TEAEs by Subgroups - Vomiting (AESI)",                             "TEAE \u2013 Vomiting (AESI)"),
    # PT-level safety subgroup tables (keep after AESI to avoid false matches)
    ("Severe TEAEs by Subgroups - Gastrointestinal Disorders",           "Severe \u2013 GI Disorders"),
    ("Severe TEAEs by Subgroups - Nausea",                               "Severe \u2013 Nausea (PT)"),
    ("Severe TEAEs by Subgroups - Vomiting",                             "Severe \u2013 Vomiting (PT)"),
    ("Severe TEAEs by Subgroups - Asthenia",                             "Severe \u2013 Asthenia"),
    ("Severe TEAEs by Subgroups - Metabolism",                           "Severe \u2013 Metabolism"),
    ("Severe TEAEs by Subgroups - Hypoalbuminaemia",                     "Severe \u2013 Hypoalb."),
    ("TEAEs by Subgroups - Gastrointestinal Disorders",                  "TEAE \u2013 GI Disorders"),
    ("TEAEs by Subgroups - Thrombocytopenia",                            "TEAE \u2013 Thrombocytopenia"),
    ("TEAEs by Subgroups - Tachycardia",                                 "TEAE \u2013 Tachycardia"),
    ("TEAEs by Subgroups - Nausea",                                      "TEAE \u2013 Nausea (PT)"),
    ("TEAEs by Subgroups - Salivary Hypersecretion",                     "TEAE \u2013 Salivary Hypersecr."),
    ("TEAEs by Subgroups - Vomiting",                                    "TEAE \u2013 Vomiting (PT)"),
    ("TEAEs by Subgroups - Malaise",                                     "TEAE \u2013 Malaise"),
    ("TEAEs by Subgroups - Oedema Peripheral",                           "TEAE \u2013 Oedema Periph."),
    ("TEAEs by Subgroups - Flank Pain",                                  "TEAE \u2013 Flank Pain"),
    ("TEAEs by Subgroups - Oral Candidiasis",                            "TEAE \u2013 Oral Candidiasis"),
    ("TEAEs by Subgroups - Metabolism And Nutrition Disorders",          "TEAE \u2013 Metabolism"),
    ("TEAEs by Subgroups - Decreased Appetite",                          "TEAE \u2013 Dec. Appetite"),
    ("TEAEs by Subgroups - Hypoalbuminaemia",                            "TEAE \u2013 Hypoalb."),
    ("TEAEs by Subgroups - Hypocalcaemia",                               "TEAE \u2013 Hypocalc."),
    ("TEAEs by Subgroups - Oropharyngeal Pain",                          "TEAE \u2013 Oropharyngeal Pain"),
    ("TEAEs by Subgroups - Dysuria",                                     "TEAE \u2013 Dysuria"),
    ("TEAEs by Subgroups - Muscular Weakness",                           "TEAE \u2013 Muscular Weakness"),
    ("TEAEs by Subgroups - Blood Bilirubin Increased",                   "TEAE \u2013 Bili. Increased"),
    ("TESAEs by Subgroups",                                              "TESAE"),
    ("Severe TEAEs by Subgroups",                                        "Severe TEAE"),

    # ── PRO (table numbers 3000–3999) ────────────────────────────────────────
    # Time-to-first-deterioration subgroup tables
    ("Time to First Deterioration of Global Health Status",              "TTFD: GHS"),
    ("Time to First Deterioration of Physical Functioning",              "TTFD: Physical Funct."),
    ("Time to First Deterioration of Role Functioning",                  "TTFD: Role Funct."),
    ("Time to First Deterioration of Emotional Functioning",             "TTFD: Emotional Funct."),
    ("Time to First Deterioration of Cognitive Functioning",             "TTFD: Cognitive Funct."),
    ("Time to First Deterioration of Social Functioning",                "TTFD: Social Funct."),
    ("Time to First Deterioration of Fatigue",                           "TTFD: Fatigue"),
    ("Time to First Deterioration of Nausea and Vomiting",              "TTFD: Nausea & Vomiting"),
    ("Time to First Deterioration of Pain and Discomfort",               "TTFD: Pain & Discomfort"),
    ("Time to First Deterioration of Pain",                              "TTFD: Pain"),
    ("Time to First Deterioration of Dyspnoea",                          "TTFD: Dyspnoea"),
    ("Time to First Deterioration of Insomnia",                          "TTFD: Insomnia"),
    ("Time to First Deterioration of Appetite Loss",                     "TTFD: Appetite Loss"),
    ("Time to First Deterioration of Constipation",                      "TTFD: Constipation"),
    ("Time to First Deterioration of Diarrhoea",                         "TTFD: Diarrhoea"),
    ("Time to First Deterioration of Financial Difficulties",            "TTFD: Financial Diff."),
    ("Time to First Deterioration of Dysphagia",                         "TTFD: Dysphagia"),
    ("Time to First Deterioration of Eating Restrictions",               "TTFD: Eating Restrictions"),
    ("Time to First Deterioration of Reflux",                            "TTFD: Reflux"),
    ("Time to First Deterioration of Odynophagia",                       "TTFD: Odynophagia"),
    ("Time to First Deterioration of Anxiety",                           "TTFD: Anxiety"),
    ("Time to First Deterioration of Eating in Front of Others",         "TTFD: Eating w/ Others"),
    ("Time to First Deterioration of Dry Mouth",                         "TTFD: Dry Mouth"),
    ("Time to First Deterioration of Trouble with Taste",                "TTFD: Taste"),
    ("Time to First Deterioration of Body Image",                        "TTFD: Body Image"),
    ("Time to First Deterioration of Trouble Swallowing Saliva",         "TTFD: Swallowing Saliva"),
    ("Time to First Deterioration of Choked",                            "TTFD: Choking"),
    ("Time to First Deterioration of Trouble with Coughing",             "TTFD: Coughing"),
    ("Time to First Deterioration of Trouble Talking",                   "TTFD: Trouble Talking"),
    ("Time to First Deterioration of Weight Loss",                       "TTFD: Weight Loss"),
    ("Time to First Deterioration of Hair Loss",                         "TTFD: Hair Loss"),
    ("Time to First Deterioration of PI01 - Pain Intensity",             "TTFD: Pain Intensity"),
    ("Time to First Deterioration of Visual Analog Scale",               "TTFD: EQ-5D VAS"),
]


# ─────────────────────────────────────────────────────────────────────────────
# Domain classification
# ─────────────────────────────────────────────────────────────────────────────

TABLE_NUM_RE = re.compile(
    r"(?:Table|Tabelle)\s+[\w]+\.[\w]+\.(\d{4})",
    re.IGNORECASE,
)


def extract_table_number(title: str) -> int | None:
    m = TABLE_NUM_RE.search(title)
    return int(m.group(1)) if m else None


def assign_domain(table_num: int | None) -> str:
    if table_num is None:
        return "Unknown"
    if 1000 <= table_num <= 1999:
        return "Efficacy"
    if 2000 <= table_num <= 2999:
        return "Safety"
    if 3000 <= table_num <= 3999:
        return "PRO"
    return "Unknown"


def derive_endpoint_short(title: str) -> str:
    """Return short endpoint label from ENDPOINT_SHORT_MAP, or auto-derive."""
    for key, short in ENDPOINT_SHORT_MAP:
        if key.lower() in title.lower():
            return short
    # Fallback: extract the most descriptive part of the title
    m = re.search(
        r"(?:of\s+|[-\u2013]\s+)([^(\u2013\-]+?)(?:\s+\(MID|\s+by\s+Subgroup|$)",
        title,
        re.IGNORECASE,
    )
    if m:
        return m.group(1).strip()[:60]
    return title[:60]


# ─────────────────────────────────────────────────────────────────────────────
# Excel parser
# ─────────────────────────────────────────────────────────────────────────────

def parse_subgroup_file(path: Path) -> list[dict]:
    """Parse a single 'by Subgroups' xlsx and return a list of row dicts."""
    try:
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        ws = wb.active
        mat = [
            [
                str(cell.value).strip() if cell.value is not None else ""
                for cell in row
            ]
            for row in ws.iter_rows()
        ]
        wb.close()
    except Exception as exc:
        print(f"    WARNING: could not read {path.name}: {exc}")
        return []

    if not mat:
        return []

    def cell(r: int, c: int) -> str | None:
        if r < 0 or r >= len(mat):
            return None
        row = mat[r]
        if c < 0 or c >= len(row):
            return None
        v = row[c]
        return v if v and v != "None" else None

    nr = len(mat)
    nc = max((len(r) for r in mat), default=0)

    # Title is always cell (0, 0)
    title = cell(0, 0) or ""
    if not title:
        return []

    # Find the "Subgroup" header row
    subg_row: int | None = None
    for r in range(nr):
        if cell(r, 0) == "Subgroup":
            subg_row = r
            break
    if subg_row is None:
        return []

    # Find interaction p-value column: rightmost column header containing
    # "interaction" anywhere in the header block
    int_col: int | None = None
    for c in range(nc - 1, -1, -1):
        for r in range(subg_row + 1):
            v = cell(r, c)
            if v and "interaction" in v.lower():
                int_col = c
                break
        if int_col is not None:
            break
    if int_col is None:
        return []

    table_num  = extract_table_number(title)
    domain     = assign_domain(table_num)
    ep_short   = derive_endpoint_short(title)

    results: list[dict] = []
    cat_parts: list[str] = []
    sub_vals:  list[str] = []
    pval: float | None   = None

    def flush() -> None:
        if pval is not None and cat_parts:
            results.append(
                {
                    "title":           title,
                    "table_number":    table_num,
                    "domain":          domain,
                    "endpoint_short":  ep_short,
                    "category":        " ".join(cat_parts),
                    "subgroup_values": ", ".join(sub_vals),
                    "interaction_p":   pval,
                }
            )

    for i in range(subg_row + 1, nr):
        c0 = cell(i, 0)
        c1 = cell(i, 1)

        if c0 is None:
            continue

        # Stop at footnotes section
        if re.match(r"^(Notes|Abbreviations)", c0, re.IGNORECASE):
            break

        if c1 is None:
            # Category-header row (or multi-line continuation)
            if sub_vals:
                flush()
                cat_parts = [c0]
                sub_vals  = []
                pval      = None
            else:
                cat_parts.append(c0)
        else:
            # Data row: subgroup level
            sub_vals.append(c0)
            if pval is None:
                raw = cell(i, int_col)
                if raw:
                    try:
                        pval = float(raw)
                    except ValueError:
                        pass

    flush()  # save final category
    return results


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main() -> None:
    tables_dir = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("extracted_tables")
    output_csv = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("subgroup_interaction_pvalues.csv")

    files = sorted(tables_dir.glob("*[Ss]ubgroup*.xlsx"))
    if not files:
        print(f"No 'by Subgroups' xlsx files found in {tables_dir}/")
        return

    print(f"Found {len(files)} subgroup files in {tables_dir}/\n")

    all_rows: list[dict] = []
    for f in files:
        rows = parse_subgroup_file(f)
        print(f"  {f.name[:80]}: {len(rows)} record(s)")
        all_rows.extend(rows)

    if not all_rows:
        print("\nNo data extracted. Check that the xlsx files match the expected format.")
        return

    fieldnames = [
        "title", "table_number", "domain", "endpoint_short",
        "category", "subgroup_values", "interaction_p",
    ]

    with output_csv.open("w", newline="", encoding="utf-8") as fh:
        writer = csv.DictWriter(fh, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(all_rows)

    print(f"\nWrote {len(all_rows)} rows → {output_csv}")


if __name__ == "__main__":
    main()
