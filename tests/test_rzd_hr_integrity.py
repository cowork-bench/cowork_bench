"""Read-only seed fixtures and isolated CPU regressions for RZD/HR graders."""
import contextlib
from datetime import datetime, timezone
import gzip
import importlib.util
import io
import json
from pathlib import Path
import shutil
import sys
import tempfile
import types
import unittest
from unittest.mock import patch

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
TASKS = ROOT / 'tasks/finalpool'
HR = 'rzd-hr1c-training-trip-kazan-strict-excel-email-gcal'
CATERING = 'rzd-kulinar-team-trip-spb-catering-excel-gcal'
GOLD = TASKS / HR / 'groundtruth_workspace/Training_Travel_Report.xlsx'


def seed_employees():
    employees = []
    active = False
    with gzip.open(ROOT / 'db/init.sql.gz', 'rt') as stream:
        for line in stream:
            if line.startswith('COPY hr1c_data."HR__PUBLIC__EMPLOYEES" '):
                active = True
                continue
            if active:
                if line.startswith('\\.'):
                    break
                values = line.rstrip('\n').split('\t')
                if values[7] in ('Продажи', 'Маркетинг') and int(values[4]) >= 3:
                    employees.append((int(values[0]), values[1], values[7], int(values[4]),
                                      float(values[3]), values[11]))
    return employees


class RzdHrIntegrityTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.eligible = seed_employees()
        cls.selected = sorted(cls.eligible, key=lambda r: (-r[3], r[0]))[:5]

    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self.tmp.cleanup)
        self.workspace = Path(self.tmp.name)
        self.events = [
            ('Корпоративный тренинг — выезд в Казань', datetime(2026, 3, 17, 5, 55), datetime(2026, 3, 17, 11, 50)),
            ('Корпоративный тренинг — возврат в Москву', datetime(2026, 3, 17, 18, 30), datetime(2026, 3, 17, 23, 55)),
        ]
        self.recipient = self.selected[0][5]
        self.executed_queries = []

    def grader(self, task=HR):
        owner = self

        class Cursor:
            def execute(self, sql, params=None):
                owner.executed_queries.append(sql)
                low = sql.lower()
                if 'hr1c_data' in low:
                    # Emulate the ranked SQL result. Assert the query contract too,
                    # so a fixture cannot conceal removal of ORDER/LIMIT.
                    owner.assertIn('order by "years_experience" desc, "employee_id" asc', low)
                    owner.assertIn('limit 5', low)
                    self.rows = owner.selected
                elif 'gcal.events' in low:
                    self.rows = owner.events
                elif task == CATERING and 'email.' in low:
                    self.rows = [('План поездки', '752А план Excel Санкт-Петербург 15 участников 82500 рублей')]
                elif 'select to_addr' in low:
                    self.rows = [(json.dumps([owner.recipient]),)]
                elif 'email.' in low:
                    self.rows = [('manager@company.ru', 'training@hr.company.ru', 'Сводка поездки',
                                  '716Г 717Г 17.03.2026 5 участников 45000 рублей')]
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

        spec = importlib.util.spec_from_file_location('rzd_test_grader', TASKS / task / 'evaluation/main.py')
        module = importlib.util.module_from_spec(spec)
        with patch.dict(sys.modules, {'psycopg2': types.SimpleNamespace(connect=lambda **kw: Connection())}):
            spec.loader.exec_module(module)
        return module

    def assert_result(self, expected, task=HR):
        grader = self.grader(task)
        output = io.StringIO()
        with patch.object(sys, 'argv', ['main.py', '--agent_workspace', str(self.workspace)]), contextlib.redirect_stdout(output):
            with self.assertRaises(SystemExit) as result:
                grader.main()
        self.assertEqual(result.exception.code, expected, output.getvalue())

    def gold(self):
        target = self.workspace / GOLD.name
        shutil.copyfile(GOLD, target)
        return target

    def test_hr_gold_and_required_calendar_pass(self):
        self.gold()
        self.assert_result(0)

    def test_hr_utc_timestamptz_passes_same_local_windows(self):
        self.gold()
        self.events = [
            (self.events[0][0], datetime(2026, 3, 17, 2, 55, tzinfo=timezone.utc), datetime(2026, 3, 17, 8, 50, tzinfo=timezone.utc)),
            (self.events[1][0], datetime(2026, 3, 17, 15, 30, tzinfo=timezone.utc), datetime(2026, 3, 17, 20, 55, tzinfo=timezone.utc)),
        ]
        self.assert_result(0)

    def test_hr_missing_workbook_fails(self):
        self.assert_result(1)

    def test_hr_corrupt_workbook_fails(self):
        (self.workspace / GOLD.name).write_bytes(b'not an xlsx')
        self.assert_result(1)

    def replace_participants(self, people):
        path = self.gold()
        workbook = openpyxl.load_workbook(path)
        for ri, person in enumerate(people, 2):
            for ci, value in enumerate(person + ('Да',), 1):
                workbook['Employees'].cell(ri, ci, value)
        for index, person in enumerate(people):
            workbook['Travel_Plan'].cell(index + 2, 1, person[1])
            workbook['Travel_Plan'].cell(index + 7, 1, person[1])
        workbook.save(path)

    def test_hr_wrong_top_five_fails(self):
        self.replace_participants(sorted(self.eligible, key=lambda r: (r[3], r[0]))[:5])
        self.assert_result(1)

    def test_hr_duplicate_employee_fails(self):
        self.replace_participants([self.selected[0]] * 5)
        self.assert_result(1)

    def test_hr_reordered_correct_participants_pass(self):
        self.replace_participants(list(reversed(self.selected)))
        self.assert_result(0)

    def test_hr_travel_rows_for_one_person_fail(self):
        path = self.gold()
        workbook = openpyxl.load_workbook(path)
        for row in workbook['Travel_Plan'].iter_rows(min_row=2):
            row[0].value = self.selected[0][1]
        workbook.save(path)
        self.assert_result(1)

    def test_hr_wrong_calendar_windows_fail(self):
        self.gold()
        self.events = [
            ('Unrelated morning event', datetime(2026, 3, 17, 0, 1), datetime(2026, 3, 17, 0, 2)),
            ('Unrelated evening event', datetime(2026, 3, 17, 23, 58), datetime(2026, 3, 17, 23, 59)),
        ]
        self.assert_result(1)

    def test_hr_wrong_arrival_time_fails(self):
        self.gold()
        self.events[0] = (self.events[0][0], self.events[0][1], datetime(2026, 3, 17, 12, 50))
        self.assert_result(1)

    def test_hr_nonselected_recipient_fails(self):
        self.gold()
        self.recipient = next(person[5] for person in self.eligible if person not in self.selected)
        self.assert_result(1)

    def test_hr_substring_recipient_does_not_match(self):
        self.gold()
        self.recipient = 'other-' + self.selected[0][5]
        self.assert_result(1)

    def catering_fixture(self):
        self.events = [('Отправление', datetime(2026, 3, 10, 6, 0)),
                       ('Прибытие и заселение', datetime(2026, 3, 10, 11, 0)),
                       ('Командный ужин', datetime(2026, 3, 10, 18, 0))]
        recipes = {r['name']: r for r in json.loads(
            (ROOT / 'local_servers/kulinar-mcp/src/data/all_recipes.json').read_text())}
        workbook = openpyxl.Workbook()
        travel = workbook.active
        travel.title = 'Travel'
        travel.append(['Train_No', 'Direction', 'Departure', 'Arrival', 'Duration_Min',
                       'Seat_Class', 'Per_Person_RUB', 'Total_Budget_RUB'])
        travel.append(['752А', 'Москва → Санкт-Петербург', '06:50', '10:50', 240, 'Эконом', 5500, 82500])
        menu = workbook.create_sheet('Menu')
        menu.append(['Course_Type', 'Dish_Name', 'Cook_Time_Min', 'Servings_Yield', 'Key_Ingredients', 'Notes'])
        for name in ['Винегрет', 'Икра кабачковая', 'Бефстроганов', 'Котлеты домашние', 'Курица в сметане', 'Борщ']:
            recipe = recipes[name]
            course = 'Закуска' if recipe['category'] in {'закуска', 'салат'} else recipe['category']
            menu.append([course, name, recipe['total_time_minutes'], recipe['servings'],
                         ', '.join(item['name'] for item in recipe['ingredients']), ''])
        timeline = workbook.create_sheet('Timeline')
        timeline.append(['Time', 'Activity', 'Location', 'Responsible_Person', 'Notes'])
        for clock, activity, location in [('06:00', 'Дорога к вокзалу', 'Москва'), ('06:30', 'Посадка', 'Вокзал'),
                ('11:00', 'Прибытие', 'СПб'), ('12:00', 'Заселение', 'Гостиница'),
                ('18:00', 'Ужин', 'Ресторан'), ('20:30', 'Тимбилдинг', 'СПб')]:
            timeline.append([clock, activity, location, 'Команда', ''])
        workbook.save(self.workspace / 'Team_Trip_Plan.xlsx')

    def test_catering_existing_valid_fixture_passes(self):
        self.catering_fixture()
        self.assert_result(0, CATERING)

    def test_catering_corrupt_workbook_fails(self):
        self.catering_fixture()
        (self.workspace / 'Team_Trip_Plan.xlsx').write_bytes(b'not an xlsx workbook')
        self.assert_result(1, CATERING)

    def test_catering_missing_workbook_fails(self):
        self.catering_fixture()
        (self.workspace / 'Team_Trip_Plan.xlsx').unlink()
        self.assert_result(1, CATERING)


if __name__ == '__main__':
    unittest.main()
