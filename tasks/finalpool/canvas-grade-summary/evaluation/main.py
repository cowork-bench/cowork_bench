"""Evaluation for canvas-grade-summary."""
import argparse
import os
import sys
import openpyxl


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


def normalized(value):
    return str(value).strip().casefold() if value is not None else ""


def validate_course_rows(rows, expected):
    """Validate every requested field; omitted columns cannot shrink the score."""
    if not rows:
        return ["Grade Summary is missing"]
    headers = [normalized(value) for value in rows[0]]
    required = [normalized(value) for value in expected[0]]
    if any(name not in headers for name in required):
        return ["Grade Summary is missing required columns"]
    indices = [headers.index(name) for name in required]
    actual = [[row[i] if i < len(row) else None for i in indices]
              for row in rows[1:] if any(value is not None and str(value).strip() for value in row)]
    wanted = {normalized(row[0]): row for row in expected[1:]}
    lookup = {normalized(row[0]): row for row in actual}
    errors = []
    if len(actual) != len(wanted) or set(lookup) != set(wanted):
        errors.append("Grade Summary must contain each course exactly once")
    for name, target in wanted.items():
        if name not in lookup:
            continue
        for i, label in enumerate(required[1:], 1):
            if not num_close(lookup[name][i], target[i], 0 if i == 1 else 0.005):
                errors.append(f"{name}.{label}: {lookup[name][i]} != {target[i]}")
    try:
        averages = [float(row[2]) for row in actual]
        if averages != sorted(averages, reverse=True):
            errors.append("Grade Summary must be sorted by Avg_Score descending")
    except (ValueError, TypeError):
        errors.append("Grade Summary contains nonnumeric scores")
    return errors


def validate_summary(rows, expected):
    actual = {normalized(row[0]): row[1] for row in rows or [] if len(row) > 1}
    errors = []
    for key in ("total_courses", "highest_avg_course", "overall_avg_score"):
        target = expected[key][1]
        value = actual.get(key)
        good = (str_match(value, target) if key == "highest_avg_course" else
                num_close(value, target, 0 if key == "total_courses" else 0.005))
        if not good:
            errors.append(f"Summary.{key}: {value} != {target}")
    return errors


def check_shared_data(expected_courses, expected_summary):
    """The prompt requests shared summary data but does not prescribe tab names.

    Either complete course statistics or the three aggregate summary metrics
    are a meaningful copy; arbitrary two-row content is not.
    """
    import psycopg2
    conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"),
                            port=int(os.environ.get("PGPORT", "5432")),
                            dbname=os.environ.get("PGDATABASE", "cowork_gym"),
                            user="eigent", password="camel")
    try:
        cur = conn.cursor()
        cur.execute("SELECT id FROM gsheet.spreadsheets WHERE lower(trim(title)) = %s",
                    ("grade summary report",))
        spreadsheets = cur.fetchall()
        for (spreadsheet_id,) in spreadsheets:
            cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s",
                        (spreadsheet_id,))
            for sheet_id, _ in cur.fetchall():
                cur.execute("SELECT row_index, col_index, value FROM gsheet.cells "
                            "WHERE spreadsheet_id = %s AND sheet_id = %s ORDER BY row_index, col_index",
                            (spreadsheet_id, sheet_id))
                grid = {}
                for ri, ci, value in cur.fetchall():
                    grid.setdefault(ri, {})[ci] = value
                rows = [[row.get(ci) for ci in range(max(row) + 1)]
                        for _, row in sorted(grid.items())]
                if not validate_summary(rows, expected_summary) or not validate_course_rows(rows, expected_courses):
                    cur.close()
                    return []
        return ["Grade Summary Report must contain correct shared summary or course statistics"]
    finally:
        conn.close()


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False)
    parser.add_argument("--groundtruth_workspace", required=False)
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    task_root = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))
    gt_dir = args.groundtruth_workspace or os.path.join(task_root, "groundtruth_workspace")

    agent_file = os.path.join(args.agent_workspace, "Canvas_Grade_Summary.xlsx")
    gt_file = os.path.join(gt_dir, "Canvas_Grade_Summary.xlsx")

    if not os.path.exists(agent_file):
        print(f"FAIL: Agent output not found: {agent_file}")
        sys.exit(1)
    if not os.path.exists(gt_file):
        print(f"FAIL: Groundtruth not found: {gt_file}")
        sys.exit(1)

    require_files(args.agent_workspace, ("Canvas_Grade_Summary.xlsx",))
    agent_wb = load_calculated_workbook(agent_file)
    gt_wb = load_calculated_workbook(gt_file)

    # ---- Load Grade Summary rows ----
    a_gs = load_sheet_rows(agent_wb, "Grade Summary")
    g_gs = load_sheet_rows(gt_wb, "Grade Summary")

    # ---- Load Summary rows ----
    a_sum = load_sheet_rows(agent_wb, "Summary")
    g_sum = load_sheet_rows(gt_wb, "Summary")

    g_sum_lookup = {}
    if g_sum:
        for row in g_sum:
            if row and row[0] is not None:
                g_sum_lookup[str(row[0]).strip().lower()] = row

    strict_errors = validate_course_rows(a_gs, g_gs)
    strict_errors.extend(validate_summary(a_sum, g_sum_lookup))
    strict_errors.extend(check_shared_data(g_gs, g_sum_lookup))
    if strict_errors:
        print("\n=== RESULT: FAIL ===")
        for error in strict_errors[:20]:
            print(error)
        sys.exit(1)
    # Required values and ordering have all been checked, with no partial-credit gate.
    print("\n=== RESULT: PASS ===")
    sys.exit(0)


if __name__ == "__main__":
    main()
