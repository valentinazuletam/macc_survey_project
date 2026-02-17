#!/usr/bin/env python3
"""
Beginner-friendly analysis script for ranking MAcc courses from an exported Qualtrics survey.

Why this script exists:
- The input file is an Excel workbook (.xlsx) stored in this repository.
- We want a deterministic (repeatable) workflow that always produces the same ranking outputs.
- We want to keep dependencies minimal, so this script uses only Python's standard library.

What this script does at a high level:
1) Reads the first worksheet from the Excel file by directly parsing the underlying XML files.
2) Finds survey columns related to:
   - "rank order" (core courses dragged into preference order)
   - "Rate ... on a scale from 1-5" (elective course ratings)
3) Cleans numeric values and calculates per-course average scores.
4) Converts both score types to a shared normalized 0-1 scale so they can be compared.
5) Builds one combined ranking table and saves it to outputs/.
6) Draws one simple SVG bar chart from the ranking and saves it to outputs/.
7) Prints a short, plain-English console summary.
"""

from __future__ import annotations

import csv
import re
import zipfile
from pathlib import Path
import xml.etree.ElementTree as ET


# -----------------------------
# Configuration (easy to edit)
# -----------------------------
DATA_FILE = Path("data/Grad Program Exit Survey Data 2024.xlsx")
OUTPUT_DIR = Path("outputs")
RANKING_CSV = OUTPUT_DIR / "course_ranking.csv"
SUMMARY_TXT = OUTPUT_DIR / "analysis_summary.txt"
FIGURE_SVG = OUTPUT_DIR / "course_ranking.svg"

# XML namespace used in Excel .xlsx files.
EXCEL_NS = {"a": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def excel_column_to_index(cell_reference: str) -> int:
    """Convert Excel column letters (e.g., A, B, AA) into a 1-based column number."""
    letters = "".join(ch for ch in cell_reference if ch.isalpha())
    value = 0
    for ch in letters:
        value = value * 26 + (ord(ch.upper()) - ord("A") + 1)
    return value


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    """
    Read shared strings from the workbook.

    Excel stores many text values in a central sharedStrings table. Cells then reference
    string indexes. We load this once so we can decode text cells quickly.
    """
    shared_strings_path = "xl/sharedStrings.xml"
    if shared_strings_path not in zf.namelist():
        return []

    shared_root = ET.fromstring(zf.read(shared_strings_path))
    strings: list[str] = []
    for string_item in shared_root.findall("a:si", EXCEL_NS):
        # A shared string can be split across multiple <t> tags; join them.
        combined_text = "".join((t.text or "") for t in string_item.findall(".//a:t", EXCEL_NS))
        strings.append(combined_text)
    return strings


def read_first_sheet_rows(xlsx_path: Path) -> list[dict[int, str]]:
    """
    Read rows from the first worksheet as a list of dictionaries.

    Each row is represented as:
        {column_index: cell_text_value, ...}

    Note: We intentionally keep this simple and deterministic for a single-sheet survey export.
    """
    with zipfile.ZipFile(xlsx_path) as zf:
        shared_strings = read_shared_strings(zf)

        # This dataset uses sheet1.xml. We read that directly.
        sheet_xml_path = "xl/worksheets/sheet1.xml"
        sheet_root = ET.fromstring(zf.read(sheet_xml_path))

        rows: list[dict[int, str]] = []
        for row in sheet_root.findall("a:sheetData/a:row", EXCEL_NS):
            parsed_row: dict[int, str] = {}

            for cell in row.findall("a:c", EXCEL_NS):
                cell_ref = cell.attrib.get("r", "")
                col_index = excel_column_to_index(cell_ref)

                cell_type = cell.attrib.get("t")
                value_node = cell.find("a:v", EXCEL_NS)
                if value_node is None:
                    # Blank cell: ignore.
                    continue

                raw_value = value_node.text or ""
                if cell_type == "s":
                    # Shared string index -> actual text.
                    string_index = int(raw_value)
                    text_value = shared_strings[string_index]
                else:
                    text_value = raw_value

                parsed_row[col_index] = text_value

            rows.append(parsed_row)

    return rows


def safe_float(value: str) -> float | None:
    """Convert a value to float safely; return None if conversion fails."""
    if value is None:
        return None
    stripped = str(value).strip()
    if stripped == "":
        return None
    try:
        return float(stripped)
    except ValueError:
        return None


def extract_course_name(question_text: str) -> str:
    """
    Pull a readable course name/code from a long survey header.

    Strategy:
    - If the header contains " - ...", use the text after the last dash.
    - Otherwise, use the full question text.
    """
    if " - " in question_text:
        return question_text.split(" - ")[-1].strip()
    return question_text.strip()


def build_course_ranking(rows: list[dict[int, str]]) -> list[dict[str, str | float | int]]:
    """
    Build one combined ranking table from core ranking columns + elective rating columns.

    Important survey structure assumptions (based on this workbook):
    - Row 2 contains human-readable question text.
    - Data begins at row 4 (Qualtrics export convention: rows 1-3 are metadata/header rows).
    """
    if len(rows) < 4:
        raise ValueError("Workbook did not contain expected header/data rows.")

    # Row 2 in Excel corresponds to index 1 in zero-based Python list.
    question_row = rows[1]

    # Data rows start at Excel row 4 -> index 3.
    data_rows = rows[3:]

    # Identify relevant columns by question text pattern.
    core_columns: dict[int, str] = {}
    elective_columns: dict[int, str] = {}

    for col_idx, question_text in question_row.items():
        lowered = question_text.lower()
        if "place each macc core course into rank order" in lowered:
            core_columns[col_idx] = extract_course_name(question_text)
        elif question_text.strip().startswith("Rate "):
            elective_columns[col_idx] = extract_course_name(question_text)

    # Containers for values per course.
    core_values: dict[str, list[float]] = {course: [] for course in core_columns.values()}
    elective_values: dict[str, list[float]] = {course: [] for course in elective_columns.values()}

    # Collect all numeric responses from each relevant column.
    for row in data_rows:
        for col_idx, course_name in core_columns.items():
            numeric_value = safe_float(row.get(col_idx, ""))
            if numeric_value is not None:
                core_values[course_name].append(numeric_value)

        for col_idx, course_name in elective_columns.items():
            numeric_value = safe_float(row.get(col_idx, ""))
            if numeric_value is not None:
                elective_values[course_name].append(numeric_value)

    # Determine max rank value dynamically from observed core rankings.
    observed_core_numbers = [v for values in core_values.values() for v in values]
    max_core_rank = int(max(observed_core_numbers)) if observed_core_numbers else 1

    combined_rows: list[dict[str, str | float | int]] = []

    # Convert core "rank" into preference score where higher is better.
    # Example with max_core_rank=8:
    # rank 1 -> score 8 (best), rank 8 -> score 1 (least preferred)
    for course_name, values in core_values.items():
        if not values:
            continue
        converted_scores = [(max_core_rank + 1) - value for value in values]
        avg_raw = sum(converted_scores) / len(converted_scores)
        avg_normalized = avg_raw / max_core_rank

        combined_rows.append(
            {
                "course": course_name,
                "source_type": "core_rank_order",
                "responses": len(values),
                "average_raw_score": round(avg_raw, 4),
                "average_normalized_score": round(avg_normalized, 4),
            }
        )

    # Elective ratings are already in "higher is better" direction (1-5 scale).
    elective_max = 5.0
    for course_name, values in elective_values.items():
        if not values:
            continue
        avg_raw = sum(values) / len(values)
        avg_normalized = avg_raw / elective_max

        combined_rows.append(
            {
                "course": course_name,
                "source_type": "elective_rating_1_to_5",
                "responses": len(values),
                "average_raw_score": round(avg_raw, 4),
                "average_normalized_score": round(avg_normalized, 4),
            }
        )

    # Final sorting for rank order: highest normalized score first, then responses, then name.
    combined_rows.sort(
        key=lambda row: (
            row["average_normalized_score"],
            row["responses"],
            row["course"],
        ),
        reverse=True,
    )

    # Assign 1-based ranking position.
    for i, row in enumerate(combined_rows, start=1):
        row["overall_rank"] = i

    return combined_rows


def write_ranking_csv(ranking_rows: list[dict[str, str | float | int]], output_path: Path) -> None:
    """Write ranking table to CSV."""
    fieldnames = [
        "overall_rank",
        "course",
        "source_type",
        "responses",
        "average_raw_score",
        "average_normalized_score",
    ]
    with output_path.open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(ranking_rows)


def write_summary_text(ranking_rows: list[dict[str, str | float | int]], output_path: Path) -> None:
    """Write a plain-English text summary."""
    total_courses = len(ranking_rows)
    top_five = ranking_rows[:5]

    lines = [
        "MAcc Exit Survey Course Ranking Summary",
        "=" * 40,
        f"Total ranked course entries: {total_courses}",
        "",
        "Top 5 ranked courses (by normalized score):",
    ]

    for row in top_five:
        lines.append(
            f"  #{row['overall_rank']}: {row['course']} "
            f"[{row['source_type']}] - "
            f"normalized score {row['average_normalized_score']} "
            f"from {row['responses']} responses"
        )

    output_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def xml_escape(text: str) -> str:
    """Escape text for safe insertion into XML/SVG."""
    return (
        text.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
        .replace('"', "&quot;")
        .replace("'", "&apos;")
    )


def write_svg_bar_chart(ranking_rows: list[dict[str, str | float | int]], output_path: Path) -> None:
    """
    Create one deterministic SVG figure showing the top 10 ranked entries.

    We use SVG instead of external plotting libraries to keep runtime dependencies minimal.
    """
    top_rows = ranking_rows[:10]

    width = 1200
    height = 620
    left_margin = 430
    right_margin = 80
    top_margin = 70
    row_height = 45
    bar_height = 28
    chart_width = width - left_margin - right_margin

    background = '<rect x="0" y="0" width="100%" height="100%" fill="white" />'

    elements = [
        f'<svg xmlns="http://www.w3.org/2000/svg" width="{width}" height="{height}">',
        background,
        '<text x="40" y="40" font-size="24" font-family="Arial" font-weight="bold">'
        "Top Ranked MAcc Courses (Normalized Scores)</text>",
        '<text x="40" y="60" font-size="14" font-family="Arial" fill="#444">'
        "Higher bar means stronger average student preference/rating</text>",
    ]

    # Axis line
    axis_y_start = top_margin - 10
    axis_y_end = top_margin + row_height * len(top_rows)
    elements.append(
        f'<line x1="{left_margin}" y1="{axis_y_start}" x2="{left_margin}" y2="{axis_y_end}" stroke="#555" stroke-width="1" />'
    )

    # Reference grid (0.0 to 1.0)
    for tick in [0.0, 0.25, 0.5, 0.75, 1.0]:
        x = left_margin + int(chart_width * tick)
        elements.append(
            f'<line x1="{x}" y1="{axis_y_start}" x2="{x}" y2="{axis_y_end}" stroke="#e0e0e0" stroke-width="1" />'
        )
        elements.append(
            f'<text x="{x}" y="{axis_y_end + 20}" font-size="12" font-family="Arial" text-anchor="middle" fill="#666">{tick:.2f}</text>'
        )

    for idx, row in enumerate(top_rows):
        y = top_margin + idx * row_height
        score = float(row["average_normalized_score"])
        bar_width = int(chart_width * score)

        label_text = f"#{row['overall_rank']} {row['course']}"
        label_text = xml_escape(label_text)

        elements.append(
            f'<text x="{left_margin - 10}" y="{y + 18}" font-size="13" font-family="Arial" text-anchor="end">{label_text}</text>'
        )

        elements.append(
            f'<rect x="{left_margin}" y="{y}" width="{bar_width}" height="{bar_height}" fill="#2f6db2" />'
        )

        elements.append(
            f'<text x="{left_margin + bar_width + 8}" y="{y + 18}" font-size="12" font-family="Arial" fill="#222">{score:.3f}</text>'
        )

    elements.append("</svg>")
    output_path.write_text("\n".join(elements), encoding="utf-8")


def main() -> None:
    """Run the full pipeline and print user-friendly console updates."""
    print("=" * 70)
    print("Starting deterministic course ranking analysis...")
    print("=" * 70)

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print(f"[1/6] Reading dataset from: {DATA_FILE}")
    rows = read_first_sheet_rows(DATA_FILE)
    print(f"      Done. Parsed {len(rows)} rows from worksheet XML.")

    print("[2/6] Identifying ranking/rating columns and cleaning numeric responses...")
    ranking_rows = build_course_ranking(rows)
    print(f"      Done. Computed comparable scores for {len(ranking_rows)} course entries.")

    print(f"[3/6] Writing ranking table to CSV: {RANKING_CSV}")
    write_ranking_csv(ranking_rows, RANKING_CSV)

    print(f"[4/6] Writing plain-English summary text: {SUMMARY_TXT}")
    write_summary_text(ranking_rows, SUMMARY_TXT)

    print(f"[5/6] Creating one SVG figure from ranking: {FIGURE_SVG}")
    write_svg_bar_chart(ranking_rows, FIGURE_SVG)

    print("[6/6] Console summary")
    print("-" * 70)
    if ranking_rows:
        best = ranking_rows[0]
        print(
            f"Top ranked entry: #{best['overall_rank']} {best['course']} "
            f"({best['source_type']}, normalized score={best['average_normalized_score']}, responses={best['responses']})"
        )
        print("Top 5 entries:")
        for row in ranking_rows[:5]:
            print(
                f"  #{row['overall_rank']}: {row['course']} | "
                f"norm={row['average_normalized_score']} | responses={row['responses']}"
            )
    else:
        print("No ranking rows were created (no valid rating/ranking values found).")
    print("-" * 70)
    print("Analysis complete. Outputs are ready in the outputs/ directory.")


if __name__ == "__main__":
    main()
