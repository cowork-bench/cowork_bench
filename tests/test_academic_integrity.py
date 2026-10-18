"""Generate small research artifacts in tempdirs; no binary fixtures or network."""
import contextlib
import importlib.util
import io
import json
from pathlib import Path
from types import SimpleNamespace
import tempfile
import unittest

from docx import Document
from pptx import Presentation
from pptx.util import Inches

ROOT = Path(__file__).resolve().parents[1]
TASK = ROOT / "tasks/finalpool/academic-llm-reasoning-ru-arxiv-word-pptx"


class AcademicIntegrityTests(unittest.TestCase):
    def setUp(self):
        spec = importlib.util.spec_from_file_location("academic_grader", TASK / "evaluation/main.py")
        self.g = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(self.g)
        self.gt = json.loads((TASK / "groundtruth_workspace/expected_results.json").read_text())
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.root = Path(self.tmp.name)

    @staticmethod
    def fact_text(facts):
        return "; ".join(f"{f['benchmark']}: {f['value']}%" for f in facts)

    def artifacts(self, fabricated=False, extra_row=False, extra_slide=False, bad_year=False):
        papers = self.gt["papers_used"]
        facts = self.gt["review_doc"]["required_result_facts"]
        doc = Document()
        doc.add_heading("Обзор литературы: Достижения в рассуждениях больших языковых моделей", 0)
        for heading in ["Введение", "Методы Chain-of-Thought", "Древовидные методы рассуждения", "Верификация и Reward-модели", "Сводная таблица", "Заключение"]:
            doc.add_heading(heading, 1)
            doc.add_paragraph("Обзор методов chain-of-thought, self-consistency, tree of thoughts, RAP, process supervision, MCTS.")
        doc.add_paragraph(self.fact_text(facts) if not fabricated else "GSM8K MATH: 999%")
        table = doc.add_table(rows=1, cols=4)
        for cell, heading in zip(table.rows[0].cells, self.gt["review_doc"]["required_table_columns"]):
            cell.text = heading
        for i, paper in enumerate(papers):
            values = ([f"FAKE PAPER {i}", "1900", "FAKE METHOD", "999% on every benchmark"] if fabricated else
                      [paper["title"], "1900" if bad_year and i == 0 else str(paper["years"][0]), paper["method_aliases"][0], self.fact_text(paper["results"])])
            for cell, value in zip(table.add_row().cells, values):
                cell.text = value
        if extra_row:
            for cell in table.add_row().cells:
                cell.text = "EXTRA"
        doc.save(self.root / "LLM_Reasoning_Review.docx")
        slides = Presentation()
        texts = [
            ("Достижения в рассуждениях LLM", "Обзор литературы 2022–2023"),
            ("Обзор", "\n".join(p["title"] for p in papers)),
            ("Chain-of-Thought Prompting", "chain-of-thought; self-consistency; " + self.fact_text(facts[:2])),
            ("Tree-Based Reasoning", "tree of thoughts; RAP; MCTS; " + self.fact_text(facts[2:4])),
            ("Process Supervision", "process supervision; " + self.fact_text(facts[4:])),
            ("Сравнение производительности", "Таблица результатов"),
            ("Заключение", "Сопоставлены рассуждения, поиск и проверка шагов."),
        ]
        for i, (title, body) in enumerate(texts):
            slide = slides.slides.add_slide(slides.slide_layouts[1])
            slide.shapes.title.text = title if not fabricated or i == 0 else "FAKE"
            slide.placeholders[1].text = body if not fabricated or i == 0 else "FAKE"
        if not fabricated:
            table = slides.slides[5].shapes.add_table(6, 3, Inches(0.1), Inches(1.5), Inches(9), Inches(4)).table
            for cell, text in zip(table.rows[0].cells, ["Method", "Benchmark", "Accuracy"]):
                cell.text = text
            for row, paper in zip(list(table.rows)[1:], papers):
                fact = paper["results"][0]
                for cell, text in zip(row.cells, [paper["method_aliases"][0], fact["benchmark"], str(fact["value"]) + "%"]):
                    cell.text = text
        if extra_slide:
            slides.slides.add_slide(slides.slide_layouts[1])
        slides.save(self.root / "LLM_Reasoning_Slides.pptx")
        # Independent whitespace extraction as a typical user terminal script.
        doc_text = "\n".join([p.text for p in doc.paragraphs] +
                             [c.text for t in doc.tables for r in t.rows for c in r.cells])
        ppt_text = []
        for slide in slides.slides:
            for shape in slide.shapes:
                if shape.has_text_frame:
                    ppt_text.extend(p.text for p in shape.text_frame.paragraphs)
                if shape.has_table:
                    ppt_text.extend(c.text for r in shape.table.rows for c in r.cells)
        (self.root / "word_count.txt").write_text(
            f"LLM_Reasoning_Review.docx: {len(doc_text.split())} words\n"
            f"LLM_Reasoning_Slides.pptx: {len(' '.join(ppt_text).split())} words\n")

    def grade(self):
        report = self.root / "report.json"
        with contextlib.redirect_stdout(io.StringIO()), self.assertRaises(SystemExit) as exit:
            self.g.main(SimpleNamespace(agent_workspace=str(self.root), groundtruth_workspace=str(TASK / "groundtruth_workspace"), res_log_file=str(report)))
        return exit.exception.code, json.loads(report.read_text())

    def test_source_based_review_and_slides_pass(self):
        self.artifacts()
        code, report = self.grade()
        self.assertEqual(code, 0, report)
        self.assertEqual(report["critical_failed"], [])

    def test_fabricated_papers_and_results_rejected(self):
        self.artifacts(fabricated=True)
        code, report = self.grade()
        self.assertEqual(code, 1)
        self.assertIn("Summary rows match the five source papers", report["critical_failed"])

    def test_extra_table_row_rejected(self):
        self.artifacts(extra_row=True)
        code, report = self.grade()
        self.assertEqual(code, 1)
        self.assertIn("Summary table has exactly 5 data rows and 4 columns", report["critical_failed"])

    def test_eighth_slide_rejected(self):
        self.artifacts(extra_slide=True)
        code, report = self.grade()
        self.assertEqual(code, 1)
        self.assertIn("PPTX has exactly 7 slides", report["critical_failed"])

    def test_wrong_paper_year_rejected(self):
        self.artifacts(bad_year=True)
        code, report = self.grade()
        self.assertEqual(code, 1)
        self.assertIn("Summary rows match the five source papers", report["critical_failed"])

    def test_fabricated_word_counts_rejected(self):
        self.artifacts()
        (self.root / "word_count.txt").write_text("LLM_Reasoning_Review.docx: 999999 words\nLLM_Reasoning_Slides.pptx: 999999 words\n")
        code, report = self.grade()
        self.assertEqual(code, 1)
        self.assertIn("Word counts match extracted artifact text", report["critical_failed"])

    def test_missing_documents_rejected(self):
        with contextlib.redirect_stdout(io.StringIO()):
            self.g.check_word_doc(self.root, self.gt)
            self.g.check_pptx(self.root, self.gt)
        self.assertIn("Word document readable", self.g.FAILED_NAMES)
        self.assertIn("Presentation readable", self.g.FAILED_NAMES)

    def test_term_and_decimal_variants_are_accepted(self):
        self.assertTrue(self.g.mentions("Chain–of–Thought", "chain-of-thought"))
        self.assertTrue(self.g.has_result("GSM8K: 56,6%", {"benchmark": "GSM8K", "value": 56.6}))
        self.assertTrue(self.g.has_result("GSM8K: 0.566", {"benchmark": "GSM8K", "value": 56.6}))
        self.assertFalse(self.g.has_result("GSM8K: 999%", {"benchmark": "GSM8K", "value": 56.6}))


if __name__ == "__main__":
    unittest.main()
