"""Evaluation for canvas-enrollment-ppt-gsheet.

Blocking checks: Enrollment_Overview.xlsx and Enrollment_Overview.pptx.
Blocking: the required shared spreadsheet and its course data.
"""
import argparse
import os
import sys
import re
import openpyxl
from pptx import Presentation


from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[4]))
from utils.evaluation.spreadsheets import load_calculated_workbook
from utils.evaluation.artifact_guards import require_files

def num_close(a, b, tol=1.0):
    if a is None and b is None:
        return True
    if a is None or b is None:
        return False
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return str(a).strip().lower() == str(b).strip().lower()


def str_match(a, b):
    if a is None or b is None:
        return a is None and b is None
    return str(a).strip().lower() == str(b).strip().lower()


def load_sheet_rows(wb, sheet_name):
    for name in wb.sheetnames:
        if name.strip().lower() == sheet_name.strip().lower():
            return [[cell.value for cell in row] for row in wb[name].iter_rows()]
    return None


def slide_records(slide):
    """Keep visible labels and values associated across text, tables and charts."""
    records, single_lines = [], []
    for shape in slide.shapes:
        if shape.has_text_frame:
            lines = [line.strip() for line in shape.text.splitlines() if line.strip()]
            records.extend(lines)
            if len(lines) == 1:
                single_lines.append((shape, lines[0]))
        if shape.has_table:
            records.extend(" | ".join(cell.text for cell in row.cells) for row in shape.table.rows)
        if shape.has_chart:
            chart = shape.chart
            categories = [str(category.label) for category in chart.plots[0].categories]
            for series in chart.series:
                values = list(series.values)
                total = sum(value for value in values if value is not None)
                percentages = any(plot.has_data_labels and plot.data_labels.show_percentage
                                  for plot in chart.plots)
                for category, value in zip(categories, values):
                    record = f"{category}: {value}"
                    if percentages and total and value is not None:
                        record += f" ({round(value / total * 100, 1)}%)"
                    records.append(record)
    # Slide tools can put labels, codes and values in separate text boxes.
    # Join a visual row, never text from unrelated rows elsewhere on the slide.
    for shape, _ in single_lines:
        aligned = [(other.left, text) for other, text in single_lines
                   if abs(other.top - shape.top) <= 914400 * 0.15]
        if len(aligned) > 1:
            records.append(" | ".join(text for _, text in sorted(aligned)))
    return [record.casefold() for record in records]


def has_number(text, expected, tolerance=0.005):
    # English thousands separators, decimal commas and nonbreaking spaces are
    # presentation choices. Preserve signs: -22 is not an acceptable count of 22.
    pattern = r"(?<![\w.])[+-]?(?:\d{1,3}(?:[, \u00a0\u202f]\d{3})+(?:[.,]\d+)?|\d+(?:[.,]\d+)?)(?![\w.])"
    for token in re.findall(pattern, text):
        token = re.sub(r"[ \u00a0\u202f]", "", token)
        if re.fullmatch(r"[+-]?\d{1,3}(?:,\d{3})+(?:\.\d+)?", token):
            token = token.replace(",", "")
        else:
            token = token.replace(",", ".")
        if num_close(token, expected, tolerance):
            return True
    return False


def labeled_number(records, labels, expected, tolerance=0.005):
    for record in records:
        normalized = record.replace("_", " ")
        for label in labels:
            for match in re.finditer(r"(?<!\w)" + re.escape(label) + r"(?!\w)", normalized):
                # Restrict numbers to this labelled entry. A later "Teachers:"
                # or semicolon starts another entry even within one text box.
                value_text = normalized[match.end():]
                value_text = re.split(r";|(?<!\w)[^\W\d_][\w ]{1,45}\s*:", value_text, maxsplit=1)[0]
                if has_number(value_text, expected, tolerance):
                    return True
    return False


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")
    all_errors = []
    critical_errors = []

    require_files(args.agent_workspace, ("Enrollment_Overview.xlsx", "Enrollment_Overview.pptx"))

    # ---- Check Excel ----
    agent_excel = os.path.join(args.agent_workspace, "Enrollment_Overview.xlsx")
    gt_excel = os.path.join(gt_dir, "Enrollment_Overview.xlsx")

    if not os.path.exists(agent_excel):
        all_errors.append("Agent output Enrollment_Overview.xlsx not found")
    elif not os.path.exists(gt_excel):
        all_errors.append("Groundtruth Enrollment_Overview.xlsx not found")
    else:
        agent_wb = load_calculated_workbook(agent_excel)
        gt_wb = load_calculated_workbook(gt_excel)

        # Check Enrollment Details
        print("  Checking Enrollment Details...")
        a_rows = load_sheet_rows(agent_wb, "Enrollment Details")
        g_rows = load_sheet_rows(gt_wb, "Enrollment Details")
        if a_rows is None:
            all_errors.append("Sheet 'Enrollment Details' not found in agent output")
        elif g_rows is None:
            all_errors.append("Sheet 'Enrollment Details' not found in groundtruth")
        else:
            required_headers = ["course_name", "course_code", "total_enrollments",
                                "students", "teachers", "tas", "active", "completed"]
            actual_headers = [str(v).strip().lower() for v in a_rows[0]]
            if actual_headers[:8] != required_headers:
                critical_errors.append("CRITICAL: Enrollment Details missing or reordered required columns")
            a_data = a_rows[1:] if len(a_rows) > 1 else []
            g_data = g_rows[1:] if len(g_rows) > 1 else []
            if len(a_data) != len(g_data):
                all_errors.append(f"Enrollment Details row count: agent={len(a_data)}, expected={len(g_data)}")

            # Match by course code (col 1)
            a_lookup = {}
            for row in a_data:
                if row and len(row) > 1 and row[1] is not None:
                    a_lookup[str(row[1]).strip().lower()] = row
            if len(a_lookup) != len(a_data):
                critical_errors.append("CRITICAL: Duplicate or missing course codes")
            # CRITICAL: all required courses must be present and Total_Enrollments
            # must be correct within a tight tolerance (verifies per-course
            # extraction from Canvas, not just two loose columns).
            for g_row in g_data:
                if not g_row or len(g_row) < 2 or g_row[1] is None:
                    continue
                key = str(g_row[1]).strip().lower()
                a_row = a_lookup.get(key)
                if a_row is None:
                    critical_errors.append(f"CRITICAL: Missing course: {g_row[0]} ({g_row[1]})")
                    continue
                if len(a_row) < 8:
                    critical_errors.append(f"CRITICAL: {key} is missing required enrollment values")
                    continue
                if not str_match(a_row[0], g_row[0]):
                    critical_errors.append(f"CRITICAL: {key}.Course_Name is incorrect")
                # Col 2: Total_Enrollments (CRITICAL, tight tol)
                if len(a_row) > 2 and len(g_row) > 2:
                    if not num_close(a_row[2], g_row[2], 0):
                        critical_errors.append(f"CRITICAL: {key}.Total_Enrollments: {a_row[2]} vs {g_row[2]} (exact)")
                # Col 3: Students (non-critical, tight tol)
                if len(a_row) > 3 and len(g_row) > 3:
                    if not num_close(a_row[3], g_row[3], 0):
                        all_errors.append(f"{key}.Students: {a_row[3]} vs {g_row[3]} (exact)")
                # Cols 4-7: Teachers, TAs, Active, Completed (non-critical, tight tol)
                for ci, cname in ((4, "Teachers"), (5, "TAs"), (6, "Active"), (7, "Completed")):
                    if len(a_row) > ci and len(g_row) > ci:
                        if not num_close(a_row[ci], g_row[ci], 0):
                            all_errors.append(f"{key}.{cname}: {a_row[ci]} vs {g_row[ci]} (exact)")
            if not all_errors and not critical_errors:
                print("    PASS")

        # Check Summary sheet
        print("  Checking Summary...")
        a_rows = load_sheet_rows(agent_wb, "Summary")
        g_rows = load_sheet_rows(gt_wb, "Summary")
        prev_errors = len(all_errors)
        if a_rows is None:
            all_errors.append("Sheet 'Summary' not found in agent output")
        elif g_rows is None:
            all_errors.append("Sheet 'Summary' not found in groundtruth")
        else:
            a_data = a_rows[1:] if len(a_rows) > 1 else []
            g_data = g_rows[1:] if len(g_rows) > 1 else []

            # Metrics that are CRITICAL deliverables (core aggregates / ranking /
            # rounding rule). Each maps metric-key -> tolerance for numerics.
            critical_numeric = {
                "total_enrollments": 0,
                "avg_enrollment_per_course": 0.005,
            }
            critical_string = {"largest_course", "smallest_course"}

            a_lookup = {}
            for row in a_data:
                if row and row[0] is not None:
                    a_lookup[str(row[0]).strip().lower()] = row
            for g_row in g_data:
                if not g_row or g_row[0] is None:
                    continue
                key = str(g_row[0]).strip().lower()
                a_row = a_lookup.get(key)
                if a_row is None:
                    # A missing core metric is critical; others non-critical.
                    if key in critical_numeric or key in critical_string:
                        critical_errors.append(f"CRITICAL: Missing summary metric: {g_row[0]}")
                    else:
                        all_errors.append(f"Missing summary metric: {g_row[0]}")
                    continue
                if len(a_row) > 1 and len(g_row) > 1:
                    # For string values, use str_match; for numbers, use num_close
                    try:
                        float(g_row[1])
                        if key in critical_numeric:
                            if not num_close(a_row[1], g_row[1], critical_numeric[key]):
                                critical_errors.append(
                                    f"CRITICAL: Summary.{key}: {a_row[1]} vs {g_row[1]} (tol={critical_numeric[key]})")
                        elif not num_close(a_row[1], g_row[1], 0):
                            all_errors.append(f"Summary.{key}: {a_row[1]} vs {g_row[1]} (exact)")
                    except (TypeError, ValueError):
                        if key in critical_string:
                            if not str_match(a_row[1], g_row[1]):
                                critical_errors.append(f"CRITICAL: Summary.{key}: {a_row[1]} vs {g_row[1]}")
                        elif not str_match(a_row[1], g_row[1]):
                            all_errors.append(f"Summary.{key}: {a_row[1]} vs {g_row[1]}")
            new_errors = len(all_errors) - prev_errors
            if new_errors == 0:
                print("    PASS")

    # ---- Check PowerPoint ----
    agent_ppt = os.path.join(args.agent_workspace, "Enrollment_Overview.pptx")
    if not os.path.exists(agent_ppt):
        all_errors.append("Agent output Enrollment_Overview.pptx not found")
    else:
        print("  Checking Enrollment_Overview.pptx...")
        prs = Presentation(agent_ppt)
        slides = list(prs.slides)
        if len(slides) < 4:
            critical_errors.append(f"CRITICAL: PPT has {len(slides)} slides, expected at least 4")
        else:
            records = [slide_records(slide) for slide in slides]
            texts = [" ".join(items) for items in records]
            required_titles = ["course enrollment overview", "enrollment summary",
                               "top 5 courses by enrollment", "enrollment distribution"]
            for index, title in enumerate(required_titles):
                if title not in texts[index]:
                    critical_errors.append(f"CRITICAL: PPT slide {index + 1} missing title '{title}'")
            if "academic dashboard report - 2026-03-06" not in texts[0]:
                critical_errors.append("CRITICAL: PPT missing required subtitle")

            gt_book = load_calculated_workbook(gt_excel)
            details = load_sheet_rows(gt_book, "Enrollment Details")[1:]
            metrics = {str(row[0]).lower(): row[1]
                       for row in load_sheet_rows(gt_book, "Summary")[1:]}
            labels = {
                "total_courses": ("total courses", "courses", "всего курсов", "число курсов", "курсов"),
                "total_enrollments": ("total enrollments", "enrollments", "всего зачислений", "число зачислений"),
                "total_students": ("total students", "students", "всего студентов", "число студентов"),
                "avg_enrollment_per_course": ("avg enrollment per course", "average enrollment per course", "average enrollment", "среднее число зачислений", "среднее"),
            }
            for metric, aliases in labels.items():
                if not labeled_number(records[1], aliases, metrics[metric]):
                    critical_errors.append(f"CRITICAL: PPT summary missing or incorrect {metric}")
            for metric in ("largest_course", "smallest_course"):
                if str(metrics[metric]).lower() not in texts[1]:
                    critical_errors.append(f"CRITICAL: PPT summary missing {metric}")
            for row in sorted(details, key=lambda row: float(row[2]), reverse=True)[:5]:
                if not any(str(row[0]).casefold() in record and str(row[1]).casefold() in record
                           and has_number(record, row[2], 0) for record in records[2]):
                    critical_errors.append(f"CRITICAL: PPT Top-5 missing course data for {row[1]}")
            total = sum(float(row[2]) for row in details)
            for index, label in ((3, "students"), (4, "teachers"), (5, "tas"),
                                 (6, "active"), (7, "completed")):
                count = sum(float(row[index]) for row in details)
                percentage = round(count / total * 100, 1)
                aliases = {"students": ("students", "студенты"),
                           "teachers": ("teachers", "преподаватели"),
                           "tas": ("tas", "teaching assistants", "ассистенты"),
                           "active": ("active", "активные"),
                           "completed": ("completed", "завершённые", "завершенные")}
                if not labeled_number(records[3], aliases[label], percentage, 0.005):
                    critical_errors.append(f"CRITICAL: PPT distribution missing {label} percentage")

        if not any("ppt" in e.lower() or "slide" in e.lower() for e in all_errors + critical_errors):
            print("    PASS")

    # The shared spreadsheet is a required deliverable, including course values.
    print("  Checking 'Enrollment Dashboard' Google Sheet...")
    import psycopg2
    conn = None
    try:
        conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"),
                                port=int(os.environ.get("PGPORT", "5432")),
                                dbname=os.environ.get("PGDATABASE", "cowork_gym"), user="eigent", password="camel")
        cur = conn.cursor()
        cur.execute("SELECT id FROM gsheet.spreadsheets WHERE lower(trim(title)) = %s",
                    ("enrollment dashboard",))
        ss = cur.fetchone()
        if not ss:
            raise ValueError("GSheet 'Enrollment Dashboard' spreadsheet not found")
        cur.execute("SELECT id FROM gsheet.sheets WHERE spreadsheet_id = %s AND lower(trim(title)) = %s",
                    (ss[0], "course data"))
        sh = cur.fetchone()
        if not sh:
            raise ValueError("GSheet 'Course Data' sheet not found")
        cur.execute("SELECT row_index, col_index, value FROM gsheet.cells "
                    "WHERE spreadsheet_id = %s AND sheet_id = %s ORDER BY row_index, col_index",
                    (ss[0], sh[0]))
        grid = {}
        for row_index, col_index, value in cur.fetchall():
            grid.setdefault(row_index, {})[col_index] = value
        if not grid:
            raise ValueError("GSheet 'Course Data' has no cells")
        header_row = min(grid)
        header = {str(value).strip().lower(): index
                  for index, value in grid[header_row].items()}
        required = ["course_name", "course_code", "total_enrollments", "students",
                    "teachers", "tas", "active", "completed"]
        if set(required) - set(header):
            raise ValueError("GSheet 'Course Data' missing required columns")
        lookup = {str(row.get(header["course_code"], "")).strip().lower(): row
                  for index, row in grid.items() if index != header_row}
        for expected in load_sheet_rows(load_calculated_workbook(gt_excel), "Enrollment Details")[1:]:
            key = str(expected[1]).strip().lower()
            row = lookup.get(key)
            if row is None or not str_match(row.get(header["course_name"]), expected[0]):
                raise ValueError(f"GSheet missing or incorrect course {key}")
            for index, column in enumerate(required[2:], 2):
                if not num_close(row.get(header[column]), expected[index], 0):
                    raise ValueError(f"GSheet {key}.{column} is incorrect")
        cur.close()
    except ValueError as exc:
        critical_errors.append(f"CRITICAL: {exc}")
    finally:
        if conn is not None:
            conn.close()

    # ---- CRITICAL gate: any critical failure => immediate FAIL ----
    if critical_errors:
        print(f"\n=== RESULT: FAIL ({len(critical_errors)} CRITICAL errors) ===")
        for e in critical_errors[:15]:
            print(f"  {e}")
        sys.exit(1)

    if all_errors:
        print(f"\n=== RESULT: FAIL ({len(all_errors)} errors) ===")
        for e in all_errors[:15]:
            print(f"  {e}")
        sys.exit(1)
    else:
        print("\n=== RESULT: PASS ===")
        sys.exit(0)


if __name__ == "__main__":
    main()
