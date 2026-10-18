"""Fail-closed validation of required artifacts before scoring dependent checks.

Used by the tasks listed in KNOWN_ISSUES: missing/unreadable outputs must not
remove their semantic checks from the accuracy denominator. Does not grade their
content or change weights. Format parsers also reject a corrupt renamed file.
"""
from pathlib import Path
import os
from zipfile import BadZipFile
from xml.etree.ElementTree import ParseError


REQUIRED_FILES = {
    "academic-llm-reasoning-ru-arxiv-word-pptx": [
        "LLM_Reasoning_Review.docx",
        "LLM_Reasoning_Slides.pptx"
    ],
    "arxiv-latex-reasoning-gsheet": [
        "Reasoning_Methods_Review.docx"
    ],
    "arxiv-method-benchmark-tracker": [
        "Method_Benchmark.xlsx"
    ],
    "arxiv-reading-plan-excel-gcal-email": [
        "Reading_Plan.xlsx"
    ],
    "arxiv-research-tracker-teamly": [
        "Transformer_Research_Tracker.xlsx"
    ],
    "arxiv-rlhf-conference-prep-ru-pptx-gcal-email": [
        "RLHF_Conference_Report.pptx"
    ],
    "canvas-assignment-deadline-word-gcal": [
        "Assignment_Deadlines_FFF2013J.xlsx",
        "Assignment_Schedule_FFF2013J.docx"
    ],
    "canvas-assignment-workload-excel": [
        "Workload_Report.xlsx"
    ],
    "canvas-exam-prep-scheduler": [
        "Exam_Prep.xlsx"
    ],
    "canvas-grade-distribution-word": [
        "Grade_Report.docx",
        "Grade_Data.xlsx"
    ],
    "canvas-grade-equity-excel-word-gcal": [
        "Grade_Equity_Analysis.xlsx",
        "Equity_Report.docx"
    ],
    "canvas-grades-gsheet-pdf-email": [
        "Grade_Dashboard_Reference.xlsx"
    ],
    "canvas-late-submission-gcal": [
        "Late_Submissions.xlsx"
    ],
    "canvas-quiz-analysis-email": [
        "Quiz_Performance.xlsx"
    ],
    "canvas-quiz-analysis-gsheet-email": [
        "Quiz_Performance_Summary.docx"
    ],
    "canvas-quiz-item-analysis-word-gcal-email": [
        "Quiz_Item_Analysis.xlsx",
        "Quiz_Analysis_Report.docx"
    ],
    "canvas-quiz-performance-excel": [
        "Quiz_Performance.xlsx"
    ],
    "canvas-sf-skills-gap-excel-word-email": [
        "Skills_Gap_Report.xlsx",
        "Skills_Gap_Brief.docx"
    ],
    "canvas-ta-workload-excel-email": [
        "TA_Workload_Report.xlsx"
    ],
    "clickhouse-hr-salary-benchmark-forms-excel-email": [
        "Salary_Analysis.xlsx"
    ],
    "clickhouse-sales-territory-realign-ppt-gsheet-email": [
        "Territory_Analysis.xlsx",
        "Territory_Review.pptx"
    ],
    "fetch-canvas-performance-excel-word-email": [
        "Performance_Report.xlsx",
        "Performance_Analysis.docx"
    ],
    "fetch-insales-inventory-forecast-excel-gcal-email": [
        "Inventory_Forecast_Report.xlsx"
    ],
    "fetch-insales-market-analysis-excel-ppt-email": [
        "Competitive_Analysis.xlsx",
        "Strategy_Presentation.pptx"
    ],
    "fetch-insales-supplier-restock-excel-word-gcal": [
        "Restocking_Plan.xlsx",
        "Procurement_Report.docx"
    ],
    "fetch-sf-sales-forecast-ppt-gcal": [
        "Sales_Forecast_Data.xlsx",
        "Q2_Sales_Forecast.pptx"
    ],
    "fetch-terminal-data-pipeline": [
        "Data_Pipeline_Report.xlsx"
    ],
    "forms-clickhouse-feedback": [
        "Support_Analysis.xlsx",
        "Improvement_Report.docx"
    ],
    "gcal-clickhouse-deadline-tracker": [
        "SLA_Compliance_Audit.xlsx"
    ],
    "insales-category-performance-ppt": [
        "Category_Review.pptx",
        "Category_Data.xlsx"
    ],
    "insales-customer-retention-email": [
        "VIP_Customer_Report.xlsx"
    ],
    "insales-fetch-supply-chain-excel-word-gcal": [
        "Supply_Chain_Optimization.xlsx",
        "Supply_Chain_Report.docx"
    ],
    "insales-product-bundle-excel-ppt-gcal": [
        "Bundle_Analysis.xlsx",
        "Bundle_Presentation.pptx"
    ],
    "insales-product-launch-dashboard": [
        "Product_Review.xlsx",
        "Q1_Product_Review.pptx"
    ],
    "insales-refund-root-cause-excel-word-email": [
        "Refund_Analysis.xlsx",
        "Root_Cause_Report.docx"
    ],
    "insales-review-analysis-word-gform": [
        "Review_Analysis.docx"
    ],
    "insales-shipping-performance-gsheet-word-email": [
        "Shipping_Performance_Report.docx"
    ],
    "inventory-supply-chain-optimization": [
        "Inventory_Forecast_Report.xlsx",
        "Supply_Chain_Risk_Assessment.docx"
    ],
    "kulinar-grocery-budget-planner": [
        "Meal_Budget.xlsx"
    ],
    "kulinar-nutrition-benchmark": [
        "Nutrition_Benchmark.xlsx",
        "Nutrition_Report.docx"
    ],
    "kulinar-survey-analysis": [
        "Preference_Analysis.xlsx"
    ],
    "kulinar-teamly-recipe-kb": [
        "Recipe_Comparison.xlsx"
    ],
    "kulinar-wellness-excel-gcal-email": [
        "Wellness_Meal_Plan.xlsx"
    ],
    "market-competitive-intelligence-report": [
        "Competitive_Analysis.xlsx",
        "Competitive_Report.docx"
    ],
    "moex-earnings-report-word-teamly": [
        "Earnings_Report.docx"
    ],
    "moex-market-summary-teamly": [
        "Market_Overview.xlsx"
    ],
    "moex-peer-comparison-excel-ppt-email": [
        "Peer_Comparison.xlsx",
        "Investor_Presentation.pptx"
    ],
    "moex-portfolio-analysis-excel-word-email": [
        "Portfolio_Analysis.xlsx",
        "Investment_Memo.docx"
    ],
    "moex-portfolio-gsheet-pdf-gcal-email": [
        "Portfolio_Dashboard_Reference.xlsx"
    ],
    "moex-stock-comparison-excel-gcal": [
        "Stock_Comparison_Report.xlsx",
        "Stock_Analysis_Report.docx"
    ],
    "pw-sf-benchmark-excel-word-email": [
        "Benchmark_Analysis.xlsx",
        "Benchmark_Report.docx"
    ],
    "pw-sf-hr-dept-efficiency-excel-gcal": [
        "Hr_Dept_Efficiency_Report.xlsx"
    ],
    "rzd-clickhouse-sales-trip-excel-word-email": [
        "Sales_Trip_Plan.xlsx",
        "Sales_Trip_Brief.docx"
    ],
    "scholarly-arxiv-survey-word-teamly-gcal": [
        "LLM_Reasoning_Survey.docx"
    ],
    "terminal-canvas-kulinar-gsheet-teamly-email": [
        "Nutrition_Academic_Study.xlsx"
    ]
}

REQUIRED_GSHEETS = {
    "arxiv-latex-reasoning-gsheet": ("Reasoning Methods Comparison", ("Papers",)),
    "arxiv-lit-review-gsheet": ("Prompt Engineering Literature Review", ("Paper Comparison", "Technique Analysis")),
    "canvas-quiz-analysis-gsheet-email": ("Quiz Performance Tracker", ("Quiz Scores",)),
    "insales-shipping-performance-gsheet-word-email": ("Shipping Performance Dashboard", ()),
    "insales-vip-customer-gsheet-gcal-email": ("VIP Customer Tracker", ()),
    "kulinar-meal-plan-calendar": ("План обедов команды на неделю", ("Меню недели", "Список покупок")),
}


def require_files(workspace, filenames):
    """Validate all mandatory files, independently of downstream early returns."""
    for filename in filenames:
        path = Path(workspace) / filename
        if not path.is_file():
            print(f"FAIL: required artifact {filename}: file does not exist")
            raise SystemExit(1)
        # Dependency failures are evaluator errors, never model verdicts.
        if path.suffix == ".xlsx":
            import openpyxl
            from openpyxl.utils.exceptions import InvalidFileException
            def parse():
                book = openpyxl.load_workbook(path, read_only=False)
                book.close()
            format_errors = (InvalidFileException,)
        elif path.suffix == ".docx":
            from docx import Document
            from docx.opc.exceptions import PackageNotFoundError
            parse = lambda: Document(path)
            format_errors = (PackageNotFoundError,)
        elif path.suffix == ".pptx":
            from pptx import Presentation
            from pptx.exc import PackageNotFoundError
            parse = lambda: Presentation(path)
            format_errors = (PackageNotFoundError,)
        else:
            raise ValueError(f"unsupported required artifact format: {path.suffix}")
        try:
            from lxml.etree import XMLSyntaxError
        except ImportError:
            format_errors += (ParseError,)
        else:
            format_errors += (ParseError, XMLSyntaxError)
        try:
            parse()
        except (BadZipFile, KeyError, ValueError) + format_errors as exc:
            print(f"FAIL: required artifact {filename}: {exc}")
            raise SystemExit(1) from exc


def require_gsheet(title, required_tabs):
    """Check the required shared artifact even if its scoring branch returns early."""
    import psycopg2
    conn = None
    try:
        conn = psycopg2.connect(host=os.environ.get("PGHOST", "localhost"),
                                port=int(os.environ.get("PGPORT", "5432")),
                                dbname=os.environ.get("PGDATABASE", "cowork_gym"),
                                user="eigent", password="camel")
        cur = conn.cursor()
        cur.execute("SELECT id FROM gsheet.spreadsheets WHERE lower(trim(title)) = %s",
                    (title.strip().lower(),))
        spreadsheets = cur.fetchall()
        for (spreadsheet_id,) in spreadsheets:
            cur.execute("SELECT id, title FROM gsheet.sheets WHERE spreadsheet_id = %s",
                        (spreadsheet_id,))
            tabs = cur.fetchall()
            selected = [row for row in tabs if not required_tabs or
                        row[1].strip().lower() in {t.lower() for t in required_tabs}]
            if not selected or (required_tabs and {t.lower() for t in required_tabs} -
                                {t.strip().lower() for _, t in selected}):
                continue
            populated = True
            for sheet_id, _ in selected:
                cur.execute("SELECT COUNT(DISTINCT row_index) FROM gsheet.cells "
                            "WHERE spreadsheet_id = %s AND sheet_id = %s",
                            (spreadsheet_id, sheet_id))
                if cur.fetchone()[0] < 2:
                    populated = False
            if populated:
                cur.close()
                return
        print(f"FAIL: required Google Sheet {title}: spreadsheet or required populated tabs missing")
        raise SystemExit(1)
    finally:
        if conn is not None:
            conn.close()


def require_task_artifacts(evaluator_path, workspace):
    task = Path(evaluator_path).resolve().parent.parent.name
    require_files(workspace, REQUIRED_FILES.get(task, ()))
    if task in REQUIRED_GSHEETS:
        require_gsheet(*REQUIRED_GSHEETS[task])
