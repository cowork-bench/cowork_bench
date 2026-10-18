"""Evaluation for moex-market-news-teamly.

Critical checks (CRITICAL_CHECKS): any of them failing => task FAIL even if the
overall accuracy is high. They verify the SUBSTANCE of the digest:
  - Excel News Articles row count == COUNT(*) FROM moex.news exactly.
  - Excel Summary Unique_Tickers == COUNT(DISTINCT symbol) FROM moex.news.
  - Excel Summary Top_Publisher == the actual top publisher from moex.news.
  - Teamly page exists (RU/EN title) AND body mentions every seeded ticker.
  - Email to portfolio-team@investco.com, subject 'Weekly Market News Digest',
    body mentions the correct total article count and the top publisher.

Expected values are read live from moex.news (read-only), never hardcoded.
"""
import argparse
from collections import Counter
from datetime import date, datetime
import re
import json
import os
import sys

import openpyxl
import psycopg2

DB = {
    "host": os.environ.get("PGHOST", "localhost"),
    "port": 5432,
    "dbname": os.environ.get("PGDATABASE", "cowork_gym"),
    "user": "eigent",
    "password": "camel",
}

CRITICAL_CHECKS = {
    "News Articles row count == moex.news (exact)",
    "News article metadata and ordering match source",
    "Teamly per-ticker counts and publishers match source",
    "Email sender and covered tickers match task",
    "Email subject is 'Weekly Market News Digest'",
    "Summary Total_Articles",
    "Summary Unique_Tickers == moex.news (exact)",
    "Summary Top_Publisher == actual top publisher",
    "Teamly digest page mentions every ticker",
    "Email body mentions total count and top publisher",
}

PASS_COUNT = 0
FAIL_COUNT = 0
CRITICAL_FAILED = []

TICKERS = ["SBER.ME", "GAZP.ME", "LKOH.ME", "TCSG.ME", "MGNT.ME", "MTSS.ME"]


def record(name, passed, detail=""):
    global PASS_COUNT, FAIL_COUNT, CRITICAL_FAILED
    if passed:
        PASS_COUNT += 1
        print(f"  [PASS] {name}")
    else:
        FAIL_COUNT += 1
        msg = f": {str(detail)[:300]}" if detail else ""
        print(f"  [FAIL] {name}{msg}")
        if name in CRITICAL_CHECKS:
            CRITICAL_FAILED.append(name)


def num_close(a, b, tol=2):
    try:
        return abs(float(a) - float(b)) <= tol
    except (TypeError, ValueError):
        return False


def get_expected_data():
    """Compute expected values from read-only moex.news."""
    articles = get_news_articles()
    publishers = Counter(row[2] for row in articles)
    top = sorted(publishers, key=lambda name: (-publishers[name], name))[0] if publishers else ""
    return len(articles), len({row[0] for row in articles}), top


def get_news_articles():
    """Read source records, matching get_moex_finance_news's STORY filter."""
    conn = psycopg2.connect(**DB)
    try:
        cur = conn.cursor()
        cur.execute("SELECT symbol, data FROM moex.news")
        articles = []
        for symbol, data in cur.fetchall():
            if isinstance(data, str):
                data = json.loads(data)
            content = (data or {}).get("content", {})
            if content.get("contentType") == "STORY":
                articles.append((symbol, content.get("title", ""),
                                 content.get("provider", {}).get("displayName", ""),
                                 normalized_date(content.get("pubDate", ""))))
        return articles
    finally:
        conn.close()


def normalized_date(value):
    # The task asks for publication dates, not a particular timestamp format.
    if isinstance(value, (datetime, date)):
        return value.date().isoformat() if isinstance(value, datetime) else value.isoformat()
    text = str(value or "").strip()
    try:
        return datetime.fromisoformat(text.replace("Z", "+00:00")).date().isoformat()
    except ValueError:
        try:
            return datetime.strptime(text, "%d.%m.%Y").date().isoformat()
        except ValueError:
            return text


def mentions_count(text, count):
    if re.search(rf"(?<!\d){count}(?!\d)", text):
        return True
    words = {2: ("два", "две", "two"), 3: ("три", "three"),
             4: ("четыре", "four"), 18: ("восемнадцать", "eighteen")}
    return any(re.search(rf"\b{word}\b", text, re.IGNORECASE) for word in words.get(count, ()))


def address_contains(value, wanted):
    return wanted.lower() in re.findall(r"[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}", str(value).lower())


def check_excel(agent_workspace, expected):
    """Check Market_News_Digest.xlsx."""
    print("\n=== Checking Market_News_Digest.xlsx ===")
    total_articles, unique_tickers, top_publisher = expected
    articles = get_news_articles()

    agent_file = os.path.join(agent_workspace, "Market_News_Digest.xlsx")
    if not os.path.isfile(agent_file):
        record("Excel file exists", False, f"Not found: {agent_file}")
        record("News article metadata and ordering match source", False, "no file")
        # cascade-fail criticals tied to excel
        record("News Articles row count == moex.news (exact)", False, "no file")
        record("Summary Unique_Tickers == moex.news (exact)", False, "no file")
        record("Summary Top_Publisher == actual top publisher", False, "no file")
        return
    record("Excel file exists", True)

    try:
        wb = openpyxl.load_workbook(agent_file, data_only=True)
    except Exception as e:
        record("Excel readable", False, str(e))
        record("News article metadata and ordering match source", False, "unreadable")
        record("News Articles row count == moex.news (exact)", False, "unreadable")
        record("Summary Unique_Tickers == moex.news (exact)", False, "unreadable")
        record("Summary Top_Publisher == actual top publisher", False, "unreadable")
        return
    record("Excel readable", True)

    # News Articles sheet
    news_sheet = None
    for name in wb.sheetnames:
        if "news" in name.lower() and "article" in name.lower():
            news_sheet = wb[name]
            break
    if news_sheet is None:
        record("Sheet 'News Articles' exists", False, f"Sheets: {wb.sheetnames}")
        record("News article metadata and ordering match source", False, "no sheet")
        record("News Articles row count == moex.news (exact)", False, "no sheet")
    else:
        record("Sheet 'News Articles' exists", True)
        rows = [r for r in news_sheet.iter_rows(min_row=2, values_only=True)
                if r and any(c is not None and str(c).strip() for c in r)]
        # Row count alone cannot establish that the actual news was extracted.
        record("News Articles row count == moex.news (exact)",
               len(rows) == total_articles,
               f"Expected {total_articles}, got {len(rows)}")
        headers = [str(c.value or "").strip().lower() for c in news_sheet[1]]
        required = ["ticker", "title", "publisher", "published_date"]
        columns_ok = all(name in headers for name in required)
        actual = []
        if columns_ok:
            indices = [headers.index(name) for name in required]
            for row in rows:
                values = [row[i] if i < len(row) else None for i in indices]
                actual.append((str(values[0] or "").strip().upper(),
                               str(values[1] or "").strip(), str(values[2] or "").strip(),
                               normalized_date(values[3])))
        ordered = sorted(sorted(actual, key=lambda r: r[3], reverse=True), key=lambda r: r[0])
        record("News article metadata and ordering match source",
               columns_ok and Counter(actual) == Counter(articles) and actual == ordered,
               "Check titles, publishers, dates, ticker coverage, duplicates and sort order")

    # Summary sheet
    sum_sheet = None
    for name in wb.sheetnames:
        if "summary" in name.lower():
            sum_sheet = wb[name]
            break
    if sum_sheet is None:
        record("Sheet 'Summary' exists", False, f"Sheets: {wb.sheetnames}")
        record("Summary Unique_Tickers == moex.news (exact)", False, "no sheet")
        record("Summary Top_Publisher == actual top publisher", False, "no sheet")
        return
    record("Sheet 'Summary' exists", True)

    summary = {}
    for row in sum_sheet.iter_rows(min_row=2, values_only=True):
        if row and row[0]:
            summary[str(row[0]).strip().lower()] = row[1]

    total_val = unique_val = top_val = None
    for key, val in summary.items():
        if "total" in key and "article" in key:
            total_val = val
        elif "unique" in key and "ticker" in key:
            unique_val = val
        elif "top" in key and "publisher" in key:
            top_val = val

    record("Summary Total_Articles", num_close(total_val, total_articles, 0),
           f"Expected {total_articles}, got {total_val}")
    # CRITICAL: exact distinct ticker count.
    record("Summary Unique_Tickers == moex.news (exact)",
           str(unique_val).strip() == str(unique_tickers),
           f"Expected {unique_tickers}, got {unique_val}")
    # CRITICAL: top publisher exact (case-insensitive, substring tolerant).
    tp_ok = bool(top_publisher) and top_publisher.lower() in str(top_val or "").lower()
    record("Summary Top_Publisher == actual top publisher", tp_ok,
           f"Expected '{top_publisher}', got '{top_val}'")


def check_teamly():
    """Teamly digest page exists with every seeded ticker in the body."""
    print("\n=== Checking Teamly Page ===")
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT id, title, body FROM teamly.pages
        WHERE title ILIKE '%market%' OR title ILIKE '%рыночн%'
           OR title ILIKE '%рынк%' OR title ILIKE '%аналитик%'
           OR title ILIKE '%дайджест%' OR title ILIKE '%news%'
        ORDER BY id DESC
    """)
    pages = cur.fetchall()
    cur.close()
    conn.close()

    pages = [p for p in pages if (p[1] or "").strip().casefold() == "центр рыночной аналитики"]
    record("Teamly digest page found", len(pages) >= 1,
           f"Found {len(pages)} pages: {[p[1] for p in pages]}")

    if not pages:
        record("Teamly digest page mentions every ticker", False, "no page")
        record("Teamly per-ticker counts and publishers match source", False, "no page")
        return

    # Pick the page covering the most tickers.
    best_body = ""
    best_cov = -1
    for _id, _title, body in pages:
        bl = (body or "").lower()
        cov = sum(1 for t in TICKERS if t.lower() in bl)
        if cov > best_cov:
            best_cov = cov
            best_body = bl

    missing = [t for t in TICKERS if t.lower() not in best_body]
    # CRITICAL: per-ticker digest built.
    record("Teamly digest page mentions every ticker", not missing,
           f"Missing tickers: {missing}; body head: {best_body[:200]}")
    articles = get_news_articles()
    per_ticker = {}
    for ticker, _title, publisher, _date in articles:
        per_ticker.setdefault(ticker, Counter())[publisher] += 1
    parts = re.split(r"(SBER\.ME|GAZP\.ME|LKOH\.ME|TCSG\.ME|MGNT\.ME|MTSS\.ME)",
                     best_body, flags=re.IGNORECASE)
    valid = set()
    for i in range(1, len(parts), 2):
        ticker, text = parts[i].upper(), parts[i + 1]
        publishers = per_ticker.get(ticker, Counter())
        count = sum(publishers.values())
        top = max(publishers.values(), default=0)
        if mentions_count(text, count) and any(
                publisher.casefold() in text.casefold()
                for publisher, n in publishers.items() if n == top):
            valid.add(ticker)
    record("Teamly per-ticker counts and publishers match source",
           set(per_ticker) <= valid, f"missing correct per-ticker summaries: {set(per_ticker) - valid}")



def check_email(expected):
    """Email digest with correct subject, recipient, and body content."""
    print("\n=== Checking Email ===")
    total_articles, _unique, top_publisher = expected
    conn = psycopg2.connect(**DB)
    cur = conn.cursor()
    cur.execute("""
        SELECT subject, to_addr, body_text, from_addr FROM email.messages
        WHERE subject ILIKE '%news%' OR subject ILIKE '%digest%'
           OR subject ILIKE '%дайджест%' OR subject ILIKE '%новост%'
    """)
    rows = cur.fetchall()
    cur.close()
    conn.close()

    record("Email digest exists", len(rows) > 0, "No matching email found")

    candidates = [row for row in rows if address_contains(row[1], "portfolio-team@investco.com")]
    def correct_message(row):
        subject, _recipient, body, sender = row
        body = str(body or "").casefold()
        return (str(subject).strip().casefold() == "weekly market news digest"
                and address_contains(sender, "research@investco.com")
                and all(t.casefold() in body for t in TICKERS)
                and mentions_count(body, total_articles)
                and top_publisher.casefold() in body)
    target = next((row for row in candidates if correct_message(row)),
                  candidates[0] if candidates else None)
    record("Email to portfolio-team@investco.com", target is not None,
           f"To addresses: {[str(r[1]) for r in rows]}")

    if target is None:
        record("Email sender and covered tickers match task", False, "no target email")
        record("Email body mentions total count and top publisher", False,
               "no email to portfolio-team@investco.com")
        return

    subj, _to, body, sender = target
    record("Email subject is 'Weekly Market News Digest'",
           str(subj).strip().casefold() == "weekly market news digest",
           f"subject: {subj}")

    bl = str(body or "").lower()
    total_ok = mentions_count(bl, total_articles)
    top_ok = bool(top_publisher) and top_publisher.lower() in bl
    record("Email sender and covered tickers match task",
           address_contains(sender, "research@investco.com") and all(t.lower() in bl for t in TICKERS),
           "Expected sender research@investco.com and all six covered tickers")
    # CRITICAL: body summary content verified.
    record("Email body mentions total count and top publisher",
           total_ok and top_ok,
           f"total({total_articles})={total_ok}, top('{top_publisher}')={top_ok}; "
           f"body head: {bl[:200]}")


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--agent_workspace", required=False, default=".")
    parser.add_argument("--groundtruth_workspace", required=False, default=".")
    parser.add_argument("--launch_time", required=False)
    parser.add_argument("--res_log_file", required=False)
    args = parser.parse_args()

    expected = get_expected_data()
    print(f"Expected: total={expected[0]}, unique_tickers={expected[1]}, "
          f"top_publisher='{expected[2]}'")

    check_excel(args.agent_workspace, expected)
    check_teamly()
    check_email(expected)

    total = PASS_COUNT + FAIL_COUNT
    pct = 100.0 * PASS_COUNT / total if total else 0
    print(f"\n=== SUMMARY ===")
    print(f"  Passed: {PASS_COUNT}/{total} ({pct:.1f}%)")

    if CRITICAL_FAILED:
        print(f"  CRITICAL FAIL: {CRITICAL_FAILED}")
        print("  Overall: FAIL")
        sys.exit(1)
    if pct < 70.0:
        print("  Overall: FAIL (accuracy < 70%)")
        sys.exit(1)
    print("  Overall: PASS")
    sys.exit(0)


if __name__ == "__main__":
    main()
