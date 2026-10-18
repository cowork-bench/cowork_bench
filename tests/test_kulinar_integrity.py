"""Portable CPU regressions for kulinar graders; PostgreSQL SELECTs are fixtures."""
import contextlib
import importlib.util
import io
import json
from pathlib import Path
import sys
import tempfile
import types
import unittest
from unittest.mock import patch

import openpyxl
from docx import Document

ROOT = Path(__file__).resolve().parents[1]
TASKS = ROOT / 'tasks' / 'finalpool'
PRICE = 'kulinar-price-cost-excel-gsheet'
TERMINAL = 'kulinar-terminal-menu-generator'


def load_grader(task, connect=None):
    spec = importlib.util.spec_from_file_location('kulinar_test_grader', TASKS / task / 'evaluation/main.py')
    module = importlib.util.module_from_spec(spec)
    with patch.dict(sys.modules, {'psycopg2': types.SimpleNamespace(connect=connect)}):
        spec.loader.exec_module(module)
    return module


def run_grader(module, workspace):
    output = io.StringIO()
    with patch.object(sys, 'argv', ['main.py', '--agent_workspace', str(workspace)]), contextlib.redirect_stdout(output):
        try:
            module.main()
        except SystemExit as exc:
            return exc.code, output.getvalue()
    return 0, output.getvalue()


class KulinarIntegrityTests(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.workspace = Path(self.tmp.name)
        self.recipes = json.loads((ROOT / 'local_servers/kulinar-mcp/src/data/all_recipes.json').read_text())
        self.by_name = {r['name']: r for r in self.recipes}
        self.price_names = ['Салат Оливье', 'Икра кабачковая', 'Борщ', 'Бефстроганов', 'Картофельное пюре']

    def price_fixture(self, names=None, zero_prices=False, names_only=False, reverse_columns=False):
        names = names or self.price_names
        prices = json.loads((TASKS / PRICE / 'initial_workspace/grocery_prices.json').read_text())
        if zero_prices:
            for entry in prices['prices']:
                entry['price_per_unit'] = 0
        (self.workspace / 'grocery_prices.json').write_text(json.dumps(prices, ensure_ascii=False))
        price_map = {entry['ingredient']: entry for entry in prices['prices']}
        wb = openpyxl.Workbook()
        dc = wb.active
        dc.title = 'Dish Costs'
        headers = ['Dish_Name', 'Ingredient_List', 'Estimated_Total_Cost', 'Cost_Per_Serving']
        dc.append(headers)
        rows = [headers]
        used = {}
        costs = []
        for name in names:
            matched = [i['name'] for i in self.by_name[name]['ingredients'] if i['name'] in price_map]
            total = sum(price_map[n]['price_per_unit'] for n in matched)
            costs.append((name, total))
            row = [name, ', '.join(matched), total, round(total / 4, 2)]
            dc.append(row)
            rows.append(row)
            for ingredient in matched:
                used.setdefault(ingredient, []).append(name)
        ip = wb.create_sheet('Ingredient Prices')
        ip.append(['Ingredient_Name', 'Unit', 'Price_Per_Unit', 'Source_Dish'])
        for name, dishes in used.items():
            ip.append([name, price_map[name]['unit'], price_map[name]['price_per_unit'], ', '.join(dishes)])
        bs = wb.create_sheet('Budget Summary')
        bs.append(['Metric', 'Value'])
        for row in [('Total_Cost', sum(c[1] for c in costs)),
                    ('Average_Cost_Per_Dish', round(sum(c[1] for c in costs) / 5, 2)),
                    ('Cheapest_Dish', min(costs, key=lambda x: x[1])[0]),
                    ('Most_Expensive_Dish', max(costs, key=lambda x: x[1])[0])]:
            bs.append(row)
        wb.save(self.workspace / 'Recipe_Cost_Analysis.xlsx')
        if names_only:
            rows = [[name] for name in names]
        if reverse_columns:
            rows = [list(reversed(row)) for row in rows]
        self.cells = [('sheet', ri, ci, str(value)) for ri, row in enumerate(rows, 1)
                      for ci, value in enumerate(row, 1)]
        owner = self

        class Cursor:
            def execute(self, sql, params=None):
                if 'FROM gsheet.spreadsheets' in sql:
                    self.rows = [('book', 'Cafeteria Menu Cost Analysis')]
                elif 'FROM gsheet.cells' in sql:
                    self.rows = owner.cells
                else:
                    raise AssertionError('Unexpected SQL: ' + sql)

            def fetchall(self):
                return self.rows

            def close(self):
                pass

        class Connection:
            def cursor(self):
                return Cursor()

            def close(self):
                pass

        return load_grader(PRICE, lambda **kwargs: Connection())

    def assert_result(self, grader, expected):
        code, output = run_grader(grader, self.workspace)
        self.assertEqual(code, expected, output)

    def test_complete_price_report_passes(self):
        self.assert_result(self.price_fixture(), 0)

    def test_reordered_google_columns_pass(self):
        self.assert_result(self.price_fixture(reverse_columns=True), 0)

    def test_three_categories_fail(self):
        names = ['Салат Оливье', 'Винегрет', 'Икра кабачковая', 'Грибы маринованные', 'Борщ']
        self.assert_result(self.price_fixture(names), 1)

    def test_agent_changed_prices_cannot_change_oracle(self):
        self.assert_result(self.price_fixture(zero_prices=True), 1)

    def test_google_sheet_names_without_financial_columns_fail(self):
        self.assert_result(self.price_fixture(names_only=True), 1)

    def test_google_sheet_wrong_price_fails(self):
        grader = self.price_fixture()
        self.cells = [(s, r, c, '0' if r == 2 and c == 3 else value) for s, r, c, value in self.cells]
        self.assert_result(grader, 1)

    def rewrite_ingredient_lists(self, transform):
        path = self.workspace / 'Recipe_Cost_Analysis.xlsx'
        wb = openpyxl.load_workbook(path)
        for row in wb['Dish Costs'].iter_rows(min_row=2):
            row[1].value = transform(row[1].value)
        wb.save(path)
        self.cells = [(sheet, ri, ci, transform(value) if ri > 1 and ci == 2 else value)
                      for sheet, ri, ci, value in self.cells]

    def test_consistently_fabricated_ingredient_lists_fail(self):
        grader = self.price_fixture()
        self.rewrite_ingredient_lists(lambda value: 'Несуществующий продукт')
        self.assert_result(grader, 1)

    def test_reordered_matched_ingredient_lists_pass(self):
        grader = self.price_fixture()
        self.rewrite_ingredient_lists(lambda value: ', '.join(reversed(value.split(', '))))
        self.assert_result(grader, 0)

    def test_corrupt_price_workbook_fails(self):
        grader = self.price_fixture()
        (self.workspace / 'Recipe_Cost_Analysis.xlsx').write_bytes(b'not a zip')
        self.assert_result(grader, 1)

    def terminal_fixture(self, single_entity=False, wrong_details=False, table=False):
        names = ['Винегрет', 'Борщ', 'Бефстроганов', 'Картофельное пюре', 'Сырники']
        selected = [self.by_name[name] for name in names]
        for filename in ['all_recipes.json', 'filtered_recipes.json']:
            (self.workspace / filename).write_text(json.dumps(selected, ensure_ascii=False))
        prefs = ('Сложность не выше 3; разнообразие категорий, хотя бы один суп и овощное блюдо; '
                 'не повторять категорию в соседние дни.')
        doc = Document()
        doc.add_heading('Меню обедов на неделю', 0)
        for day, recipe in zip(['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница'], selected):
            doc.add_heading(day, 1)
            doc.add_paragraph('Блюдо: ' + recipe['name'])
            fields = [('Категория', recipe['category']), ('Сложность', str(recipe['difficulty'])),
                      ('Ингредиенты', '; '.join(i['text_quantity'] for i in recipe['ingredients']))]
            if wrong_details:
                fields = [('Категория', 'напиток'), ('Сложность', '99'), ('Ингредиенты', 'только вода')]
            if table == 'horizontal':
                grid = doc.add_table(rows=2, cols=3)
                for ci, (key, value) in enumerate(fields):
                    grid.cell(0, ci).text = key
                    grid.cell(1, ci).text = value
            elif table:
                grid = doc.add_table(rows=0, cols=2)
                for key, value in fields:
                    cells = grid.add_row().cells
                    cells[0].text, cells[1].text = key, value
            else:
                for key, value in fields:
                    doc.add_paragraph(key + ': ' + value)
        doc.add_heading('Пищевые предпочтения', 1)
        doc.add_paragraph(prefs)
        doc.save(self.workspace / 'Weekly_Menu.docx')
        entities = [{'type': 'entity', 'name': 'Меню и предпочтения', 'entityType': 'meal_plan',
                     'observations': [prefs] + names}]
        if not single_entity:
            entities.append({'type': 'entity', 'name': 'Контекст', 'entityType': 'context', 'observations': ['Обеды']})
        (self.workspace / 'memory').mkdir()
        (self.workspace / 'memory/memory.json').write_text('\n'.join(json.dumps(e, ensure_ascii=False) for e in entities))
        return load_grader(TERMINAL)

    def test_terminal_complete_report_passes(self):
        self.assert_result(self.terminal_fixture(), 0)

    def test_terminal_native_single_entity_passes(self):
        self.assert_result(self.terminal_fixture(single_entity=True), 0)

    def test_terminal_table_content_passes(self):
        self.assert_result(self.terminal_fixture(table=True), 0)

    def test_terminal_horizontal_table_content_passes(self):
        self.assert_result(self.terminal_fixture(table='horizontal'), 0)

    def test_terminal_wrong_reported_recipe_details_fail(self):
        self.assert_result(self.terminal_fixture(wrong_details=True), 1)


if __name__ == '__main__':
    unittest.main()
