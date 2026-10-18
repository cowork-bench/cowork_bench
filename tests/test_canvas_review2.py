"""Second-review contract tests using real Office artifacts and explicit SQL results."""
import contextlib
import io
import re
import shutil
import sys
import types
import unittest
from pathlib import Path
from unittest.mock import patch

import openpyxl
from openpyxl.worksheet.formula import ArrayFormula
from pptx import Presentation
from pptx.util import Inches

from tests import test_canvas_integrity as canvas_tests
from tests.test_canvas_integrity import POOL, evaluator
from utils.evaluation.artifact_guards import require_gsheet


class Cursor:
    def __init__(self, sheets):
        self.sheets = sheets

    def execute(self, sql, params=None):
        if 'FROM gsheet.spreadsheets' in sql:
            self.result = [('shared-report',)]
        elif 'FROM gsheet.sheets' in sql:
            selected = [(i, name) for i, (name, _) in self.sheets.items()
                        if len(params) == 1 or name.lower() == params[1]]
            self.result = selected if 'SELECT id, title' in sql else [(i,) for i, _ in selected]
        elif 'COUNT(DISTINCT' in sql:
            rows = self.sheets[params[1]][1]
            self.result = [(sum(any(v is not None and str(v).strip() for v in row) for row in rows),)]
        elif 'FROM gsheet.cells' in sql:
            self.result = [(ri, ci, str(value) if value is not None else '')
                           for ri, row in enumerate(self.sheets[params[1]][1])
                           for ci, value in enumerate(row)]
        else:
            raise AssertionError(sql)

    def fetchall(self):
        return self.result

    def fetchone(self):
        return self.result[0] if self.result else None

    def close(self):
        pass


class CanvasReview2Tests(unittest.TestCase):
    setUp = canvas_tests.CanvasIntegrityTests.setUp
    prepare = canvas_tests.CanvasIntegrityTests.prepare
    run_enrollment = canvas_tests.CanvasIntegrityTests.run_evaluator

    def database(self, sheets):
        cursor = Cursor(sheets)
        connection = types.SimpleNamespace(cursor=lambda: cursor, close=lambda: None)
        return patch.dict(sys.modules, {'psycopg2': types.SimpleNamespace(connect=lambda **kw: connection)})

    def run_grade(self, sheets):
        task = 'canvas-grade-summary'
        module = evaluator(task)
        argv = ['eval', '--agent_workspace', str(self.workspace), '--groundtruth_workspace',
                str(POOL / task / 'groundtruth_workspace')]
        with self.database(sheets), patch.object(sys, 'argv', argv), contextlib.redirect_stdout(io.StringIO()):
            with self.assertRaises(SystemExit) as result:
                module.main()
        return result.exception.code

    def grade_sheets(self):
        book = openpyxl.load_workbook(POOL / 'canvas-grade-summary/groundtruth_workspace/Canvas_Grade_Summary.xlsx')
        return {10: ('Summary', list(book['Summary'].values))}

    def test_grade_correct_data_with_extra_blank_tab_passes(self):
        self.prepare('canvas-grade-summary')
        sheets = self.grade_sheets()
        sheets[11] = ('Sheet1', [])
        self.assertEqual(self.run_grade(sheets), 0)

    def test_shared_guard_allows_blank_extra_but_rejects_empty_required(self):
        with self.database({10: ('Data', [['Metric'], ['Value']]), 11: ('Sheet1', [])}):
            require_gsheet('Report', ())
            require_gsheet('Report', ('Data',))
            with contextlib.redirect_stdout(io.StringIO()), self.assertRaises(SystemExit):
                require_gsheet('Report', ('Data', 'Sheet1'))
        with self.database({10: ('Data', [[''], [None]])}), contextlib.redirect_stdout(io.StringIO()):
            with self.assertRaises(SystemExit):
                require_gsheet('Report', ())

    def test_grade_shared_garbage_or_wrong_aggregate_fails(self):
        self.prepare('canvas-grade-summary')
        for rows in ([[''], ['']], [['junk'], ['unrelated']],
                     [('Total_Courses', 22), ('Highest_Avg_Course', 'wrong'), ('Overall_Avg_Score', 75.4)]):
            with self.subTest(rows=rows):
                self.assertEqual(self.run_grade({10: ('Summary', rows)}), 1)

    def test_grade_complete_shared_course_table_is_valid(self):
        self.prepare('canvas-grade-summary')
        book = openpyxl.load_workbook(self.workspace / 'Canvas_Grade_Summary.xlsx')
        self.assertEqual(self.run_grade({10: ('Any tab name', list(book['Grade Summary'].values))}), 0)

    def test_grade_missing_numeric_columns_and_wrong_cell_fail(self):
        self.prepare('canvas-grade-summary')
        path = self.workspace / 'Canvas_Grade_Summary.xlsx'
        original = path.read_bytes()
        for mutation in ('delete', 'wrong'):
            path.write_bytes(original)
            book = openpyxl.load_workbook(path)
            if mutation == 'delete':
                book['Grade Summary'].delete_cols(2, 4)
            else:
                book['Grade Summary']['D2'] = -100
            book.save(path)
            self.assertEqual(self.run_grade(self.grade_sheets()), 1)

    @unittest.skipUnless(shutil.which('soffice') or shutil.which('libreoffice') or
                         Path('/Applications/LibreOffice.app/Contents/MacOS/soffice').is_file(), 'requires Calc')
    def test_array_formula_passes_but_calculation_error_fails(self):
        self.prepare('canvas-grade-summary')
        path = self.workspace / 'Canvas_Grade_Summary.xlsx'
        for formula, expected in ((ArrayFormula(ref='B2', text="=COUNTA('Grade Summary'!A2:A23)"), 0), ('=1/0', 1)):
            book = openpyxl.load_workbook(path)
            book['Summary']['B2'] = formula
            book.save(path)
            self.assertEqual(self.run_grade(self.grade_sheets()), expected)

    def edit_ppt(self, transform):
        path = self.workspace / 'Enrollment_Overview.pptx'
        prs = Presentation(path)
        for index, slide in enumerate(prs.slides):
            for shape in slide.shapes:
                if shape.has_text_frame:
                    shape.text = transform(index, shape.text)
        prs.save(path)

    def test_english_thousands_and_table_top5_pass(self):
        task = 'canvas-enrollment-ppt-gsheet'
        self.prepare(task)
        self.edit_ppt(lambda i, text: re.sub(r'32663|32593|2501|2369|2296|2285|2241',
                                           lambda match: f'{int(match[0]):,}', text))
        self.assertEqual(self.run_enrollment(task), 0)
        path = self.workspace / 'Enrollment_Overview.pptx'
        prs = Presentation(path)
        slide = prs.slides[2]
        for shape in slide.shapes:
            if shape.has_text_frame and 'Top 5 Courses' not in shape.text:
                shape.text = ''
        book = openpyxl.load_workbook(self.workspace / 'Enrollment_Overview.xlsx')
        top = sorted(list(book['Enrollment Details'].values)[1:], key=lambda row: row[2], reverse=True)[:5]
        table = slide.shapes.add_table(5, 3, Inches(.3), Inches(1), Inches(9), Inches(4)).table
        for ri, row in enumerate(top):
            for ci, value in enumerate(row[:3]):
                table.cell(ri, ci).text = str(value)
        prs.save(path)
        self.assertEqual(self.run_enrollment(task), 0)

    def test_swapped_ppt_values_fail(self):
        task = 'canvas-enrollment-ppt-gsheet'
        for slide_index, left, right in ((2, '2501', '2369'), (3, '69.2%', '30.8%')):
            with self.subTest(slide=slide_index):
                self.prepare(task)
                self.edit_ppt(lambda i, text: text.replace(left, 'SWAP').replace(right, left).replace('SWAP', right)
                              if i == slide_index else text)
                self.assertEqual(self.run_enrollment(task), 1)

    def test_ppt_same_line_labels_cannot_swap_values(self):
        module = evaluator('canvas-enrollment-ppt-gsheet')
        self.assertFalse(module.labeled_number(['active: 30.8%; completed: 69.2%'], ('active',), 69.2))
        self.assertFalse(module.has_number('-22', 22))

    def test_enrollment_wrong_names_and_inflated_total_fail(self):
        task = 'canvas-enrollment-ppt-gsheet'
        for mutation in ('names', 'total'):
            self.prepare(task)
            path = self.workspace / 'Enrollment_Overview.xlsx'
            book = openpyxl.load_workbook(path)
            if mutation == 'names':
                for row in book['Enrollment Details'].iter_rows(min_row=2):
                    row[0].value = 'NONEXISTENT COURSE'
            else:
                book['Summary']['B2'] = 72
            book.save(path)
            self.assertEqual(self.run_enrollment(task), 1)


if __name__ == '__main__':
    unittest.main()
