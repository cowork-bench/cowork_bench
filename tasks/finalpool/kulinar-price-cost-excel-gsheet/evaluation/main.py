"""
Evaluation for kulinar-price-cost-excel-gsheet.

The agent picks 5 kulinar dishes from 5 distinct categories, matches each
recipe's ingredients to grocery_prices.json (RU names, RUB prices), and builds
Recipe_Cost_Analysis.xlsx with 3 sheets (Dish Costs / Ingredient Prices /
Budget Summary) plus a Google Sheet "Cafeteria Menu Cost Analysis".

We do NOT hardcode which dishes the agent chose. Every semantic check
RECOMPUTES the expected value from the agent's own rows / the price list /
the kulinar recipe set, so any valid choice of 5 dishes passes.

Critical (semantic) checks: any failure => overall FAIL regardless of accuracy.
Otherwise pass threshold: accuracy >= 70%.
"""
import argparse
import json
import os
import sys

import openpyxl
import psycopg2

from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[4]))
from utils.evaluation.spreadsheets import load_calculated_workbook, UnsupportedExternalWorkbook

DB_CONFIG = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": "cowork_gym",
    "user": "eigent",
    "password": "camel",
}

# Source of truth for kulinar recipes (globally seeded MCP data).
KULINAR_JSON = os.path.join(
    os.path.dirname(__file__),
    "..", "..", "..", "..",
    "local_servers", "kulinar-mcp", "src", "data", "all_recipes.json",
)

CRITICAL_CHECKS = {
    "Excel file exists",
    "Excel file readable",
    "Dish Costs has required columns",
    "Dish Costs sheet exists",
    "Ingredient Prices sheet exists",
    "Ingredient Prices exactly reproduce used ingredients, units and source dishes",
    "Budget Summary sheet exists",
    "Dish Costs has 5 data rows",
    "Dishes are real kulinar recipes from 5 categories",
    "Ingredient_List matches the recipe ingredients with known prices",
    "Cost_Per_Serving == round(Total/4) for every dish",
    "Estimated_Total_Cost matches sum of matched ingredient prices",
    "Budget Summary aggregates correct",
    "Ingredient Prices: Price_Per_Unit matches grocery_prices.json",
    "Google Sheet 'Cafeteria Menu Cost Analysis' exists",
    "Google Sheet reproduces all Dish Costs values",
}

PASS_COUNT = 0
FAIL_COUNT = 0
FAILED_NAMES = []


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        FAILED_NAMES.append(name)
        msg = f": {detail[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")


def num_close(a, b, tol=0.5):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def norm_name(s):
    return str(s).strip().lower() if s is not None else ""


def load_kulinar():
    with open(os.path.normpath(KULINAR_JSON), encoding="utf-8") as f:
        recipes = json.load(f)
    by_name = {r["name"]: r for r in recipes}
    name_to_cat = {r["name"]: r["category"] for r in recipes}
    return by_name, name_to_cat


def load_prices(agent_workspace):
    """Use the immutable task input, never the agent-writable copy as an oracle."""
    path = os.path.join(os.path.dirname(__file__), "..", "initial_workspace", "grocery_prices.json")
    with open(path, encoding="utf-8") as f:
        rows = json.load(f)["prices"]
    return {r["ingredient"]: r for r in rows}


def sheet_rows(ws):
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], []
    header = [str(c).strip() if c is not None else "" for c in rows[0]]
    data = [r for r in rows[1:] if r and any(c is not None for c in r)]
    return header, data


def col_idx(header, *aliases):
    low = [h.lower() for h in header]
    for a in aliases:
        if a.lower() in low:
            return low.index(a.lower())
    return None


def find_sheet(wb, *keywords):
    for name in wb.sheetnames:
        nl = name.lower()
        if all(k in nl for k in keywords):
            return name
    return None


def check_excel(agent_workspace):
    print("\n=== Check 1: Excel File ===")
    by_name, name_to_cat = load_kulinar()
    price_map = load_prices(agent_workspace)

    excel_path = os.path.join(agent_workspace, "Recipe_Cost_Analysis.xlsx")
    if not os.path.isfile(excel_path):
        record("Excel file exists", False, f"Not found: {excel_path}")
        return
    record("Excel file exists", True)

    try:
        wb = load_calculated_workbook(excel_path)
    except (ImportError, RuntimeError, UnsupportedExternalWorkbook):
        raise
    except Exception as e:
        record("Excel file readable", False, str(e))
        return
    record("Excel file readable", True)
    record("Excel has 3 sheets", len(wb.sheetnames) >= 3, f"Sheets: {wb.sheetnames}")

    # ---------------- Dish Costs ----------------
    dc_name = find_sheet(wb, "dish") or find_sheet(wb, "cost")
    dishes = []  # (name, ingredient_list[str], total, cps)
    if not dc_name:
        record("Dish Costs sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Dish Costs sheet exists", True)
        ws = wb[dc_name]
        header, data = sheet_rows(ws)
        record("Dish Costs has 5 data rows", len(data) == 5, f"Found {len(data)} rows")

        ci_name = col_idx(header, "Dish_Name")
        ci_ing = col_idx(header, "Ingredient_List")
        ci_tot = col_idx(header, "Estimated_Total_Cost")
        ci_cps = col_idx(header, "Cost_Per_Serving")
        record("Dish Costs has required columns",
               None not in (ci_name, ci_ing, ci_tot, ci_cps),
               f"Header: {header}")

        if None not in (ci_name, ci_ing, ci_tot, ci_cps):
            for r in data[:5]:
                dishes.append((r[ci_name], r[ci_ing], r[ci_tot], r[ci_cps]))

            # (a) dishes are real kulinar recipes from 5 distinct categories
            real = [d[0] for d in dishes if d[0] in by_name]
            cats = {name_to_cat[n] for n in real}
            record("Dishes are real kulinar recipes from 5 categories",
                   len(real) == 5 and len(cats) == 5,
                   f"real={len(real)}/5, categories={sorted(cats)}")

            mismatches = []
            for dish_name, ingredient_list, _, _ in dishes:
                if dish_name not in by_name:
                    mismatches.append(str(dish_name))
                    continue
                expected_ingredients = {
                    norm_name(item["name"]) for item in by_name[dish_name]["ingredients"]
                    if item["name"] in price_map
                }
                actual_ingredients = {norm_name(item) for item in str(ingredient_list or "").split(",")
                                      if norm_name(item)}
                if actual_ingredients != expected_ingredients:
                    mismatches.append(str(dish_name))
            record("Ingredient_List matches the recipe ingredients with known prices",
                   len(dishes) == 5 and not mismatches,
                   f"Wrong ingredient lists: {mismatches}")

            # (b) Cost_Per_Serving == round(total/4, 2) recomputed from the row
            cps_ok = all(num_close(cps, round(float(tot) / 4, 2), 0.5)
                         for (_, _, tot, cps) in dishes
                         if tot is not None and cps is not None)
            record("Cost_Per_Serving == round(Total/4) for every dish", cps_ok,
                   f"rows={[(d[0], d[2], d[3]) for d in dishes]}")

            # (c) Estimated_Total_Cost == sum of matched ingredient prices.
            # Recompute from the recipe + price list; matching = recipe ingredient
            # whose name is a price-list key. Tolerant to small rounding.
            if price_map:
                tot_ok = True
                detail = ""
                for (dn, _, tot, _) in dishes:
                    if dn not in by_name or tot is None:
                        continue
                    expected = round(sum(
                        price_map[i["name"]]["price_per_unit"]
                        for i in by_name[dn]["ingredients"]
                        if i["name"] in price_map
                    ), 2)
                    if not num_close(tot, expected, 1.0):
                        tot_ok = False
                        detail = f"{dn}: got {tot}, expected {expected}"
                        break
                record("Estimated_Total_Cost matches sum of matched ingredient prices",
                       tot_ok, detail)
            else:
                record("Estimated_Total_Cost matches sum of matched ingredient prices",
                       False, "price list not found")

    # ---------------- Ingredient Prices ----------------
    ip_name = find_sheet(wb, "ingredient") or find_sheet(wb, "price")
    if not ip_name:
        record("Ingredient Prices sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Ingredient Prices sheet exists", True)
        ws = wb[ip_name]
        header, data = sheet_rows(ws)
        record("Ingredient Prices has >= 10 rows", len(data) >= 10, f"Found {len(data)} rows")

        ci_in = col_idx(header, "Ingredient_Name")
        ci_pr = col_idx(header, "Price_Per_Unit")
        if None in (ci_in, ci_pr) or not price_map:
            record("Ingredient Prices: Price_Per_Unit matches grocery_prices.json",
                   False, f"header={header}, prices_loaded={bool(price_map)}")
        else:
            ok = True
            detail = ""
            checked = 0
            for r in data:
                iname = str(r[ci_in]).strip() if r[ci_in] is not None else ""
                if iname in price_map:
                    checked += 1
                    if not num_close(r[ci_pr], price_map[iname]["price_per_unit"], 0.01):
                        ok = False
                        detail = f"{iname}: got {r[ci_pr]}, expected {price_map[iname]['price_per_unit']}"
                        break
            record("Ingredient Prices: Price_Per_Unit matches grocery_prices.json",
                   ok and checked >= 5, detail or f"checked {checked} matched ingredients")

        expected_usage = {}
        for dish_name, _, _, _ in dishes:
            if dish_name in by_name:
                for item in by_name[dish_name]["ingredients"]:
                    if item["name"] in price_map:
                        expected_usage.setdefault(norm_name(item["name"]), set()).add(norm_name(dish_name))
        unit_col = col_idx(header, "Unit")
        source_col = col_idx(header, "Source_Dish")
        complete = None not in (ci_in, ci_pr, unit_col, source_col) and bool(expected_usage)
        observed = set()
        source_prices = {norm_name(name): value for name, value in price_map.items()}
        if complete:
            for row in data:
                ingredient = norm_name(row[ci_in])
                if ingredient in observed or ingredient not in expected_usage:
                    complete = False
                    continue
                observed.add(ingredient)
                reference = source_prices[ingredient]
                sources = {norm_name(value) for value in str(row[source_col] or "").split(",") if norm_name(value)}
                complete = (complete and norm_name(row[unit_col]) == norm_name(reference["unit"])
                            and num_close(row[ci_pr], reference["price_per_unit"], 0.01)
                            and sources == expected_usage[ingredient])
        record("Ingredient Prices exactly reproduce used ingredients, units and source dishes",
               complete and observed == set(expected_usage),
               f"missing={sorted(set(expected_usage) - observed)}")

    # ---------------- Budget Summary ----------------
    bs_name = find_sheet(wb, "budget") or find_sheet(wb, "summary")
    if not bs_name:
        record("Budget Summary sheet exists", False, f"Sheets: {wb.sheetnames}")
    else:
        record("Budget Summary sheet exists", True)
        ws = wb[bs_name]
        _, data = sheet_rows(ws)
        kv = {}
        for r in data:
            if len(r) >= 2 and r[0] is not None:
                kv[str(r[0]).strip()] = r[1]
        needed = {"Total_Cost", "Average_Cost_Per_Dish", "Cheapest_Dish", "Most_Expensive_Dish"}
        record("Budget Summary has all 4 metrics", needed.issubset(kv.keys()),
               f"keys={list(kv.keys())}")

        if dishes and needed.issubset(kv.keys()):
            try:
                costs = [(d[0], float(d[2])) for d in dishes if d[2] is not None]
            except (TypeError, ValueError):
                costs = []
            if len(costs) == 5:
                exp_total = round(sum(c for _, c in costs), 2)
                exp_avg = round(exp_total / 5, 2)
                minimum = min(cost for _, cost in costs)
                maximum = max(cost for _, cost in costs)
                exp_cheap = {norm_name(name) for name, cost in costs if cost == minimum}
                exp_exp = {norm_name(name) for name, cost in costs if cost == maximum}
                ok = (num_close(kv["Total_Cost"], exp_total, 1.0)
                      and num_close(kv["Average_Cost_Per_Dish"], exp_avg, 0.5)
                      and norm_name(kv["Cheapest_Dish"]) in exp_cheap
                      and norm_name(kv["Most_Expensive_Dish"]) in exp_exp)
                record("Budget Summary aggregates correct", ok,
                       f"got total={kv['Total_Cost']} avg={kv['Average_Cost_Per_Dish']} "
                       f"cheap={kv['Cheapest_Dish']} exp={kv['Most_Expensive_Dish']}; "
                       f"expected total={exp_total} avg={exp_avg} cheap={exp_cheap} exp={exp_exp}")
            else:
                record("Budget Summary aggregates correct", False,
                       "could not parse 5 dish costs")
        else:
            record("Budget Summary aggregates correct", False,
                   "missing metrics or no dish rows")

    wb.close()
    return dishes


def check_gsheet(dishes):
    print("\n=== Check 2: Google Sheet ===")
    try:
        conn = psycopg2.connect(**DB_CONFIG)
        cur = conn.cursor()
        cur.execute("SELECT id, title FROM gsheet.spreadsheets")
        spreadsheets = cur.fetchall()

        if not spreadsheets:
            record("Google Sheet 'Cafeteria Menu Cost Analysis' exists", False, "No spreadsheets")
            cur.close(); conn.close()
            return

        target = None
        for ss_id, ss_title in spreadsheets:
            tl = (ss_title or "").lower()
            if "cafeteria" in tl or ("cost" in tl and "analysis" in tl):
                target = (ss_id, ss_title)
                break
        record("Google Sheet 'Cafeteria Menu Cost Analysis' exists", target is not None,
               f"Titles: {[s[1] for s in spreadsheets]}")
        if not target:
            target = spreadsheets[0]

        cur.execute("""
            SELECT c.sheet_id, c.row_index, c.col_index, c.value
            FROM gsheet.cells c
            JOIN gsheet.sheets s ON c.sheet_id = s.id
            WHERE s.spreadsheet_id = %s
            ORDER BY c.sheet_id, c.row_index, c.col_index
        """, (target[0],))
        cells = cur.fetchall()
        record("Google Sheet has data", len(cells) >= 5, f"Found {len(cells)} cells")

        # Match complete rows, preserving their sheet/row/column coordinates.
        # Accept reordered columns and either 0- or 1-based cell coordinates.
        sheets = {}
        for sheet_id, ri, ci, value in cells:
            sheets.setdefault(sheet_id, {}).setdefault(ri, {})[ci] = value
        required = ("Dish_Name", "Ingredient_List", "Estimated_Total_Cost", "Cost_Per_Serving")
        matched = False
        for rows in sheets.values():
            for header_row, header in rows.items():
                columns = {norm_name(value): ci for ci, value in header.items()}
                if not all(norm_name(name) in columns for name in required):
                    continue
                ci = [columns[norm_name(name)] for name in required]
                actual = [tuple(row.get(col) for col in ci)
                          for ri, row in sorted(rows.items()) if ri > header_row
                          and any(row.get(col) not in (None, "") for col in ci)]
                if len(actual) != len(dishes or []):
                    continue
                remaining = list(actual)
                for expected in dishes or []:
                    for index, row in enumerate(remaining):
                        same_ingredients = (
                            {norm_name(v) for v in str(row[1] or "").split(",")}
                            == {norm_name(v) for v in str(expected[1] or "").split(",")}
                        )
                        if (norm_name(row[0]) == norm_name(expected[0]) and same_ingredients
                                and num_close(row[2], expected[2], 0.01)
                                and num_close(row[3], expected[3], 0.01)):
                            remaining.pop(index)
                            break
                    else:
                        break
                else:
                    matched = bool(dishes) and not remaining
                if matched:
                    break
            if matched:
                break
        record("Google Sheet reproduces all Dish Costs values", matched,
               "Expected the headers and all five complete rows from Dish Costs")

        cur.close(); conn.close()
    except Exception as e:
        record("Google Sheet 'Cafeteria Menu Cost Analysis' exists", False, str(e))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    dishes = check_excel(args.agent_workspace)
    check_gsheet(dishes or [])

    total = PASS_COUNT + FAIL_COUNT
    if total == 0:
        print("\nFAIL: No checks were performed.")
        sys.exit(1)

    accuracy = PASS_COUNT / total * 100
    print(f"\n=== Results: {PASS_COUNT}/{total} passed ({accuracy:.1f}%) ===")

    critical_failed = [n for n in FAILED_NAMES if n in CRITICAL_CHECKS]
    if critical_failed:
        print(f"CRITICAL FAILURES: {len(critical_failed)}")
        for n in critical_failed:
            print(f"  - {n}")

    result = {
        "total_passed": PASS_COUNT,
        "total_checks": total,
        "accuracy": accuracy,
        "critical_failed": critical_failed,
    }
    if args.res_log_file:
        with open(args.res_log_file, "w") as f:
            json.dump(result, f, indent=2)

    if critical_failed:
        print("FAIL (critical check failed)")
        sys.exit(1)
    if accuracy >= 70:
        print("PASS")
        sys.exit(0)
    print("FAIL")
    sys.exit(1)


if __name__ == "__main__":
    main()
