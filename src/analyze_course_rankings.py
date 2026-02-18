#!/usr/bin/env python3
"""
Deterministic, beginner-friendly elective-course ranking workflow.

Goal of this script:
- Read the Qualtrics Excel export in this repository.
- Rank ONLY elective courses using their 1-5 rating columns.
- Ignore all core-course rank-order columns.
- Produce deterministic outputs in outputs/ for local use and GitHub Actions artifacts.

Important survey notes:
- Qualtrics exports include metadata/header rows at the top.
- In this workbook:
  * Row 1 = machine column names / metadata
  * Row 2 = human-readable question text
  * Row 3 = import metadata JSON
  * Row 4+ = actual responses
"""

from __future__ import annotations

import csv
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


# -------------------------------------------------------------------
# Configuration
# -------------------------------------------------------------------
DATA_FILE = Path("data/Grad Program Exit Survey Data 2024.xlsx")
OUTPUT_DIR = Path("outputs")
RANKING_CSV = OUTPUT_DIR / "course_ranking.csv"
SUMMARY_TXT = OUTPUT_DIR / "analysis_summary.txt"
FIGURE_SVG = OUTPUT_DIR / "course_ranking.svg"

EXCEL_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def excel_column_to_index(cell_reference: str) -> int:
    """Convert Excel column letters (A, B, AA, AB...) to 1-based index numbers."""
    letters = "".join(ch for ch in cell_reference if ch.isalpha())
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch.upper()) - ord("A") + 1)
    return value


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """Load Excel shared strings table used by text cells."""
    shared_path = "xl/sharedStrings.xml"
    if shared_path not in zf.namelist():
        return []

    root = ET.fromstring(zf.read(shared_path))
    strings: list[str] = []
    for si in root.findall("a:si", EXCEL_NS):
        text = "".join((t.text or "") for t in si.findall(".//a:t", EXCEL_NS))
        strings.append(text)
    return strings


def read_first_sheet_rows(xlsx_path: Path) -> list[dict[int, str]]:
    """
    Read the first worksheet and return rows as dictionaries:
      {column_index: cell_text, ...}
    """
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = read_shared_strings(zf)
        sheet_root = ET.fromstring(zf.read("xl/worksheets/sheet1.xml"))

        rows: list[dict[int, str]] = []
        for row in sheet_root.findall("a:sheetData/a:row", EXCEL_NS):
            parsed: dict[int, str] = {}
            for cell in row.findall("a:c", EXCEL_NS):
                ref = cell.attrib.get("r", "")
                col_index = excel_column_to_index(ref)

                value_node = cell.find("a:v", EXCEL_NS)
                if value_node is None:
                    continue

                raw = value_node.text or ""
                if cell.attrib.get("t") == "s":
                    parsed[col_index] = shared_strings[int(raw)]
                else:
                    parsed[col_index] = raw

            rows.append(parsed)

    return rows


def safe_float(value: str | None) -> float | None:
    """Convert text to float safely; return None for blank/non-numeric values."""
    if value is None:
        return None
    text = str(value).strip()
    if text == "":
        return None
    try:
        return float(text)
    except ValueError:
        return None


def extract_course_name_from_rating_header(question_text: str) -> str:
    """
    Extract a concise course name from the long Qualtrics rating header.

    We split on ' - ' and keep the final segment, e.g.:
      "Rate ACC 6150 ... - ACC 6150 Information Systems Audit"
      -> "ACC 6150 Information Systems Audit"
    """
    if " - " in question_text:
        return question_text.split(" - ")[-1].strip()
    return question_text.strip()


def build_elective_only_ranking(rows: list[dict[int, str]]) -> list[dict[str, str | float | int]]:
    """
    Build ranking ONLY from elective 1-5 rating columns.

    Deterministic sort rules:
      1) mean_rating descending (higher is better)
      2) n_responses descending
      3) course alphabetical ascending (A->Z)

    Also enforces:
    - ignore missing ratings when computing means and counts
    - drop any elective column with all missing values
    """
    if len(rows) < 4:
        raise ValueError("Expected Qualtrics header + data rows, but file appears too short.")

    # Row 1 (index 0) is Qualtrics metadata and is removed from analysis context.
    # We intentionally use row 2 (index 1) for human-readable question text.
    question_row = rows[1]

    # Actual response records start at Excel row 4 (index 3).
    data_rows = rows[3:]

    # IMPORTANT: Only include elective rating columns.
    # We identify them by the survey prompt starting with "Rate ".
    elective_columns: dict[int, str] = {}
    for col_idx, question_text in question_row.items():
        if question_text.strip().startswith("Rate "):
            elective_columns[col_idx] = extract_course_name_from_rating_header(question_text)

    # Gather numeric ratings for each elective course.
    ratings_by_course: dict[str, list[float]] = {course: [] for course in elective_columns.values()}

    for row in data_rows:
        for col_idx, course in elective_columns.items():
            rating = safe_float(row.get(col_idx))
            if rating is not None:
                ratings_by_course[course].append(rating)

    # Build summary rows and drop courses with all-missing ratings.
    ranking_rows: list[dict[str, str | float | int]] = []
    for course, ratings in ratings_by_course.items():
        if not ratings:
            continue

        mean_rating = sum(ratings) / len(ratings)
        ranking_rows.append(
            {
                "course": course,
                "mean_rating": round(mean_rating, 4),
                "n_responses": len(ratings),
            }
        )

    # Deterministic sort with the requested tie-break rules.
    ranking_rows.sort(key=lambda r: (-float(r["mean_rating"]), -int(r["n_responses"]), str(r["course"])))

    # Assign final rank numbers after sorting.
    for idx, row in enumerate(ranking_rows, start=1):
        row["rank"] = idx

    return ranking_rows


def write_ranking_csv(rows: list[dict[str, str | float | int]], output_path: Path) -> None:
    """Write elective ranking CSV with exactly the requested columns and order."""
    fieldnames = ["course", "mean_rating", "n_responses", "rank"]
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)


def write_summary(rows: list[dict[str, str | float | int]], output_path: Path) -> None:
    """Write a short plain-English summary documenting elective-only logic and tie-break rules."""
    lines = [
        "Elective Course Ranking Summary",
        "=" * 32,
        "This analysis ranks ONLY elective courses (1-5 rating columns).",
        "Core rank-order course fields were ignored.",
        "",
        "Deterministic tie-break rules:",
        "1) mean_rating descending (higher is better)",
        "2) n_responses descending (higher is better)",
        "3) course alphabetical ascending (A->Z)",
        "",
        f"Total ranked electives: {len(rows)}",
        "",
        "Top ranked electives:",
    ]

    for row in rows[:5]:
        lines.append(
            f"  #{row['rank']}: {row['course']} | "
            f"mean_rating={row['mean_rating']} | n_responses={row['n_responses']}"
        )

    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def write_svg_chart(rows: list[dict[str, str | float | int]], output_path: Path) -> None:
    """Optional SVG chart of elective ranking sorted by mean_rating."""
    top_rows = rows
    width = 1200
    height = 120 + 36 * len(top_rows)
    left = 420
    right = 80
    top = 70
    row_h = 30
    bar_h = 18
    chart_w = width - left - right

    def esc(t: str) -> str:
        return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    parts = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        '<rect x="0" y="0" width="100%" height="100%" fill="white"/>',
        '<text x="30" y="35" font-family="Arial" font-size="24" font-weight="bold">Elective Course Ranking by Mean Rating</text>',
        '<text x="30" y="55" font-family="Arial" font-size="13" fill="#444">Scale shown from 0 to 5 mean rating</text>',
    ]

    for tick in [0, 1, 2, 3, 4, 5]:
        x = left + int(chart_w * (tick / 5.0))
        parts.append(f'<line x1="{x}" y1="{top-8}" x2="{x}" y2="{top + row_h*len(top_rows)}" stroke="#e2e2e2"/>')
        parts.append(f'<text x="{x}" y="{top + row_h*len(top_rows) + 20}" text-anchor="middle" font-size="12" font-family="Arial" fill="#666">{tick}</text>')

    for i, row in enumerate(top_rows):
        y = top + i * row_h
        mean_rating = float(row["mean_rating"])
        w = int(chart_w * (mean_rating / 5.0))
        label = esc(f"#{row['rank']} {row['course']}")

        parts.append(f'<text x="{left-10}" y="{y+14}" text-anchor="end" font-size="12" font-family="Arial">{label}</text>')
        parts.append(f'<rect x="{left}" y="{y}" width="{w}" height="{bar_h}" fill="#2f6db2"/>')
        parts.append(f'<text x="{left+w+8}" y="{y+14}" font-size="12" font-family="Arial">{mean_rating:.3f}</text>')

    parts.append("</svg>")
    output_path.write_text("\n".join(parts), encoding="utf-8")


def main() -> None:
    print("=" * 70)
    print("Starting elective-only deterministic ranking analysis...")
    print("=" * 70)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print(f"[1/6] Reading dataset from: {DATA_FILE}")
    rows = read_first_sheet_rows(DATA_FILE)
    print(f"      Parsed {len(rows)} worksheet rows.")

    print("[2/6] Removing Qualtrics metadata context and selecting elective rating columns only...")
    ranking_rows = build_elective_only_ranking(rows)
    print(f"      Built ranking rows for {len(ranking_rows)} elective courses.")

    print(f"[3/6] Writing CSV output: {RANKING_CSV}")
    write_ranking_csv(ranking_rows, RANKING_CSV)

    print(f"[4/6] Writing text summary: {SUMMARY_TXT}")
    write_summary(ranking_rows, SUMMARY_TXT)

    print(f"[5/6] Writing ranking figure: {FIGURE_SVG}")
    write_svg_chart(ranking_rows, FIGURE_SVG)

    print("[6/6] Console summary")
    print("-" * 70)
    print("Ranking scope: ELECTIVE courses only (core courses ignored).")
    print("Tie-breakers: mean_rating desc, n_responses desc, course A->Z.")
    if ranking_rows:
        print("Top 5 electives:")
        for row in ranking_rows[:5]:
            print(
                f"  #{row['rank']}: {row['course']} | "
                f"mean_rating={row['mean_rating']} | n_responses={row['n_responses']}"
            )
    else:
        print("No elective rating data found.")
    print("-" * 70)
    print("Done. Outputs are available in outputs/ (CSV, summary, and SVG figure).")


if __name__ == "__main__":
    main()
