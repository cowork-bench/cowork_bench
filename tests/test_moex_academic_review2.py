"""Skeptical second-review regressions; portable CPU artifact probes only."""
import contextlib
import importlib.util
import io
import json
from pathlib import Path
import re
import subprocess
import types
import unittest

from docx import Document
from pptx import Presentation

ROOT = Path(__file__).resolve().parents[1]


def load_helper(filename, name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "tests" / filename)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


helpers = load_helper("test_academic_integrity.py", "academic_review2_helpers")


def original_module(ref="40b8508"):
    path = helpers.TASK / "evaluation/main.py"
    relative = str(path.relative_to(ROOT))
    source = subprocess.check_output(["git", "show", ref + ":" + relative], cwd=ROOT, text=True)
    module = types.ModuleType("academic_base_review2")
    module.__file__ = str(path)
    exec(compile(source, str(path), "exec"), module.__dict__)
    return module


def outcome(grader, helper, function, metadata=None):
    grader.FAILED_NAMES.clear()
    with contextlib.redirect_stdout(io.StringIO()):
        if function == "check_word_count":
            passed, total = grader.check_word_count(helper.root)
        else:
            passed, total = getattr(grader, function)(helper.root, metadata or helper.gt)
    return {"passed": passed, "total": total, "failed_names": list(grader.FAILED_NAMES)}


def update_counts(helper, lexical=False):
    doc = Document(helper.root / "LLM_Reasoning_Review.docx")
    doc_text = "\n".join([p.text for p in doc.paragraphs] +
                         [c.text for t in doc.tables for r in t.rows for c in r.cells])
    ppt = Presentation(helper.root / "LLM_Reasoning_Slides.pptx")
    ppt_text = []
    for slide in ppt.slides:
        for shape in slide.shapes:
            if shape.has_text_frame:
                ppt_text.append(shape.text)
            if shape.has_table:
                ppt_text.extend(c.text for r in shape.table.rows for c in r.cells)
    count = (lambda text: len(re.findall(r"\b\w+\b", text))) if lexical else (lambda text: len(text.split()))
    counts = [count(doc_text), count("\n".join(ppt_text))]
    (helper.root / "word_count.txt").write_text(
        f"LLM_Reasoning_Review.docx: {counts[0]} words\n"
        f"LLM_Reasoning_Slides.pptx: {counts[1]} words\n")
    return counts


class AcademicReview2Tests(unittest.TestCase):
    def setUp(self):
        self.h = helpers.AcademicIntegrityTests("test_source_based_review_and_slides_pass")
        self.h.setUp()
        self.addCleanup(self.h.doCleanups)
        self.h.artifacts()

    def test_wrong_benchmark_value_pair_rejected(self):
        path = self.h.root / "LLM_Reasoning_Review.docx"
        doc = Document(path)
        doc.tables[0].rows[1].cells[3].text = "GSM8K: 999%; MATH: 56.6%"
        doc.save(path)
        result = outcome(self.h.g, self.h, "check_word_doc")
        self.assertIn("Summary rows match the five source papers", result["failed_names"])

    def test_valid_alternative_seeded_benchmark_is_accepted(self):
        path = self.h.root / "LLM_Reasoning_Review.docx"
        doc = Document(path)
        doc.tables[0].rows[1].cells[3].text = "SVAMP: 79.0%"
        doc.save(path)
        result = outcome(self.h.g, self.h, "check_word_doc")
        self.assertEqual(result["failed_names"], [])

    def test_common_plain_lexical_word_counter_is_accepted(self):
        update_counts(self.h, lexical=True)
        result = outcome(self.h.g, self.h, "check_word_count")
        self.assertEqual(result["failed_names"], [])

    def test_result_associations_in_prose_and_reordered_numbers(self):
        fact = {"benchmark": "GSM8K", "value": 56.6}
        self.assertTrue(self.h.g.has_result("GSM8K accuracy: 56,6%", fact))
        self.assertTrue(self.h.g.has_result("56.6% on GSM8K", fact))
        self.assertFalse(self.h.g.has_result("GSM8K: 999%; MATH: 56.6%", fact))
        self.assertFalse(self.h.g.has_result("MATH: 56.6%, GSM8K: 999%", fact))

    def test_valid_alternative_slide_results_are_accepted(self):
        path = self.h.root / "LLM_Reasoning_Slides.pptx"
        slides = Presentation(path)
        slides.slides[2].placeholders[1].text = "chain-of-thought: SVAMP 79.0%; self-consistency: SVAMP 86.6%"
        slides.save(path)
        self.h.g.FAILED_NAMES.clear()
        with contextlib.redirect_stdout(io.StringIO()):
            self.h.g.check_pptx(self.h.root, self.h.gt)
        self.assertEqual(self.h.g.FAILED_NAMES, [])

    def test_wide_benchmark_comparison_preserves_column_association(self):
        from pptx.util import Inches
        slides = Presentation()
        slide = slides.slides.add_slide(slides.slide_layouts[6])
        table = slide.shapes.add_table(2, 3, Inches(0), Inches(0), Inches(8), Inches(2)).table
        for row, values in zip(table.rows, [["Method", "GSM8K", "MATH"], ["CoT", "56.6%", "999%"]]):
            for cell, value in zip(row.cells, values):
                cell.text = value
        fact = {"benchmark": "GSM8K", "value": 56.6}
        self.assertTrue(self.h.g.table_has_result(table, fact))
        table.cell(1, 1).text, table.cell(1, 2).text = "999%", "56.6%"
        self.assertFalse(self.h.g.table_has_result(table, fact))

    def test_competitor_served_archive_matches_expected_json_files(self):
        import tarfile
        root = ROOT / "tasks/finalpool/teamly-fetch-competitor/files"
        with tarfile.open(root / "mock_api.tar.gz") as archive:
            for path in sorted((root / "mock_api/api").rglob("*.json")):
                relative = str(path.relative_to(root / "mock_api"))
                member = next(name for name in archive.getnames() if name.endswith(relative))
                self.assertEqual(json.loads(archive.extractfile(member).read()), json.loads(path.read_text()))

    def test_spelled_out_counts_accept_cases_without_substring_matches(self):
        moex = load_helper("test_moex_teamly_integrity.py", "moex_review2_count_helpers")
        g = moex.load("moex-market-news-teamly")
        for text, count in [("три статьи", 3), ("из трёх статей", 3), ("из трех статей", 3),
                            ("three articles", 3), ("из двух статей", 2),
                            ("из четырёх статей", 4), ("о восемнадцати статьях", 18)]:
            self.assertTrue(g.mentions_count(text, count), text)
        self.assertFalse(g.mentions_count("тринадцать статей", 3))
        self.assertFalse(g.mentions_count("из 113 статей", 3))

    def test_competitor_detail_column_permutation_is_rejected(self):
        moex = load_helper("test_moex_teamly_integrity.py", "moex_review2_columns_helpers")
        h = moex.CompetitorTests("test_source_fields_and_reordered_lists_pass")
        h.setUp()
        try:
            h.workbook()
            import openpyxl
            path = h.root / "Competitor_Analysis.xlsx"
            wb = openpyxl.load_workbook(path)
            for row in wb["Detailed Analysis"]:
                row[1].value, row[4].value = row[4].value, row[1].value
            wb.save(path)
            with contextlib.redirect_stdout(io.StringIO()):
                self.assertFalse(h.g.check_excel(h.root))
            self.assertIn("Detailed Analysis column order matches task", h.g.FAILED_NAMES)
        finally:
            h.doCleanups()

    def test_source_alternative_comparison_table_is_accepted(self):
        path = self.h.root / "LLM_Reasoning_Slides.pptx"
        slides = Presentation(path)
        table = next(shape.table for shape in slides.slides[5].shapes if shape.has_table)
        for row in table.rows:
            for cell in row.cells:
                cell.text = ""
        for index, values in enumerate([["Method", "Benchmark", "Accuracy"],
                                        ["CoT", "SVAMP", "79.0%"],
                                        ["Self-consistency", "SVAMP", "86.6%"]]):
            for cell, value in zip(table.rows[index].cells, values):
                cell.text = value
        slides.save(path)
        self.h.g.FAILED_NAMES.clear()
        with contextlib.redirect_stdout(io.StringIO()):
            self.h.g.check_pptx(self.h.root, self.h.gt)
        self.assertEqual(self.h.g.FAILED_NAMES, [])


def baseline_comparison():
    h = helpers.AcademicIntegrityTests("test_source_based_review_and_slides_pass")
    h.setUp()
    try:
        base, prv1, results = original_module(), original_module("0c8d3dc"), {}
        relative = str((helpers.TASK / "groundtruth_workspace/expected_results.json").relative_to(ROOT))
        old_meta = json.loads(subprocess.check_output(["git", "show", "0c8d3dc:" + relative], cwd=ROOT, text=True))
        new_meta = h.gt
        for name in ["cross_benchmark_wrong_value", "valid_svamp_alternative", "plain_lexical_word_counts", "valid_svamp_slide_results"]:
            h.gt = old_meta
            h.artifacts()
            h.gt = new_meta
            if name == "plain_lexical_word_counts":
                counts = update_counts(h, lexical=True)
                function = "check_word_count"
            elif name == "valid_svamp_slide_results":
                path = h.root / "LLM_Reasoning_Slides.pptx"
                slides = Presentation(path)
                slides.slides[2].placeholders[1].text = "chain-of-thought: SVAMP 79.0%; self-consistency: SVAMP 86.6%"
                slides.save(path)
                function, counts = "check_pptx", None
            else:
                path = h.root / "LLM_Reasoning_Review.docx"
                doc = Document(path)
                doc.tables[0].rows[1].cells[3].text = (
                    "GSM8K: 999%; MATH: 56.6%" if name == "cross_benchmark_wrong_value" else "SVAMP: 79.0%")
                doc.save(path)
                function, counts = "check_word_doc", None
            results[name] = {"base_40b8508": outcome(base, h, function, old_meta),
                             "pr_v1_0c8d3dc": outcome(prv1, h, function, old_meta),
                             "working_tree": outcome(h.g, h, function)}
            if counts:
                results[name]["reported_counts"] = counts
            else:
                update_counts(h)
            # Also run the real main() verdict path, with matching source
            # metadata for each historical version and honest artifact counts.
            full = {}
            for label, grader, metadata in [("base_40b8508", base, old_meta),
                                             ("pr_v1_0c8d3dc", prv1, old_meta),
                                             ("working_tree", h.g, new_meta)]:
                groundtruth = h.root / ("reference_" + label)
                groundtruth.mkdir(exist_ok=True)
                (groundtruth / "expected_results.json").write_text(json.dumps(metadata))
                report = h.root / (label + "_report.json")
                grader.FAILED_NAMES.clear()
                args = types.SimpleNamespace(agent_workspace=str(h.root), groundtruth_workspace=str(groundtruth),
                                             res_log_file=str(report), launch_time=None)
                with contextlib.redirect_stdout(io.StringIO()):
                    try:
                        grader.main(args)
                    except SystemExit as exc:
                        code = exc.code
                full[label] = {"exit_code": code, "report": json.loads(report.read_text())}
            results[name]["full_native_verdicts"] = full
        return results
    finally:
        h.doCleanups()


if __name__ == "__main__":
    unittest.main()
