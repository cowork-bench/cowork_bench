"""Independent CPU review probes. Actual source seed + fixtures, SQL adapters only."""
from pathlib import Path
from datetime import datetime
import contextlib,gzip,io,json,shutil,subprocess,sys,tempfile,types,unittest
from unittest.mock import patch
import openpyxl
ROOT=Path(__file__).resolve().parents[1]
POOL=ROOT/'tasks/finalpool'
HR='rzd-hr1c-training-trip-kazan-strict-excel-email-gcal'
CAT='rzd-kulinar-team-trip-spb-catering-excel-gcal'
PRICE='kulinar-price-cost-excel-gsheet'
SEED=[]
with gzip.open(ROOT/'db/init.sql.gz','rt') as f:
    active=False
    for line in f:
        if line.startswith('COPY hr1c_data.'):active='EMPLOYEES' in line
        elif active:
            if line.startswith('\\.'):break
            r=line.rstrip('\n').split('\t')
            if r[7] in {'Продажи','Маркетинг'} and int(r[4])>=3:
                SEED.append((int(r[0]),r[1],r[7],int(r[4]),float(r[3]),r[11]))
SELECTED=sorted(SEED,key=lambda r:(-r[3],r[0]))[:5]
RECIPES={r['name']:r for r in json.loads((ROOT/'local_servers/kulinar-mcp/src/data/all_recipes.json').read_text())}
PRICES=json.loads((POOL/PRICE/'initial_workspace/grocery_prices.json').read_text())
PMAP={r['ingredient']:r for r in PRICES['prices']}
class Cursor:
    def __init__(self,mode,cells):self.mode,self.cells=mode,cells
    def execute(self,sql,params=None):self.sql,self.params=sql.lower(),params
    def close(self):pass
    def fetchall(self):
        q=self.sql
        if self.mode=='hr':
            if 'hr1c_data' in q:
                return SELECTED if 'limit 5' in q else [(r[1],r[2],r[3],r[5]) for r in SEED]
            if 'gcal.events' in q:return [('Корпоративный тренинг — выезд в Казань',datetime(2026,3,17,5,55),datetime(2026,3,17,11,50)),('Корпоративный тренинг — возврат в Москву',datetime(2026,3,17,18,30),datetime(2026,3,17,23,55))]
            if 'select to_addr' in q:return [(json.dumps([SELECTED[0][5]]),)]
            if 'email.' in q:return [('hrmanager@company.ru',json.dumps(['training@hr.company.ru']),'Сводка','716Г 717Г 17.03.2026 5 участников 45000 рублей')]
        if self.mode=='catering':
            if 'gcal.events' in q:return [('Тимбилдинг — отправление',datetime(2026,3,10,6,0)),('Прибытие и заселение',datetime(2026,3,10,11,0)),('Командный ужин',datetime(2026,3,10,18,0))]
            if 'email.' in q:return [('План поездки','752А план Excel Санкт-Петербург, 15 человек, 82500 рублей, приложен Team_Trip_Plan.xlsx')]
        if self.mode=='price':
            if 'gsheet.spreadsheets' in q:return [('book','Cafeteria Menu Cost Analysis')]
            if 'gsheet.cells' in q:return self.cells if 'row_index' in q else [(r[3],) for r in self.cells]
        raise AssertionError((self.mode,q,self.params))
class Connection:
    def __init__(self,mode,cells):self.mode,self.cells=mode,cells
    def cursor(self):return Cursor(self.mode,self.cells)
    def close(self):pass

def run(task,workspace,mode,cells=None,base=False):
    path=POOL/task/'evaluation/main.py'
    src=subprocess.check_output(['git','show',f'40b8508:tasks/finalpool/{task}/evaluation/main.py'],cwd=ROOT,text=True) if base else path.read_text()
    module=types.ModuleType('review2_grader');module.__file__=str(path)
    db=types.SimpleNamespace(connect=lambda **kw:Connection(mode,cells))
    with patch.dict(sys.modules,{'psycopg2':db}):exec(compile(src,str(path),'exec'),module.__dict__)
    buf=io.StringIO()
    argv=['main.py','--agent_workspace',str(workspace),'--groundtruth_workspace',str(POOL/task/'groundtruth_workspace')]
    with patch.dict(sys.modules,{'psycopg2':db}),patch.object(sys,'argv',argv),contextlib.redirect_stdout(buf):
        try:module.main();code=0
        except SystemExit as exc:code=exc.code
    log=buf.getvalue()
    return code,log

def hr_fixture(ws):shutil.copyfile(POOL/HR/'groundtruth_workspace/Training_Travel_Report.xlsx',ws/'Training_Travel_Report.xlsx')

def price_fixture(ws,names=None):
    names=names or ['Салат Оливье','Икра кабачковая','Борщ','Бефстроганов','Картофельное пюре']
    (ws/'grocery_prices.json').write_text(json.dumps(PRICES,ensure_ascii=False))
    wb=openpyxl.Workbook();dish=wb.active;dish.title='Dish Costs';header=['Dish_Name','Ingredient_List','Estimated_Total_Cost','Cost_Per_Serving'];dish.append(header)
    matrix=[header];used={};costs=[]
    for name in names:
        ingredients=[i['name'] for i in RECIPES[name]['ingredients'] if i['name'] in PMAP]
        total=sum(PMAP[i]['price_per_unit'] for i in ingredients);costs.append((name,total))
        row=[name,', '.join(ingredients),total,round(total/4,2)];dish.append(row);matrix.append(row)
        for ingredient in ingredients:used.setdefault(ingredient,[]).append(name)
    ip=wb.create_sheet('Ingredient Prices');ip.append(['Ingredient_Name','Unit','Price_Per_Unit','Source_Dish'])
    for ingredient,dishes in used.items():ip.append([ingredient,PMAP[ingredient]['unit'],PMAP[ingredient]['price_per_unit'],', '.join(dishes)])
    bs=wb.create_sheet('Budget Summary');bs.append(['Metric','Value'])
    for row in [('Total_Cost',sum(c[1] for c in costs)),('Average_Cost_Per_Dish',round(sum(c[1] for c in costs)/5,2)),('Cheapest_Dish',min(costs,key=lambda r:r[1])[0]),('Most_Expensive_Dish',max(costs,key=lambda r:r[1])[0])]:bs.append(row)
    wb.save(ws/'Recipe_Cost_Analysis.xlsx')
    return [('sheet',ri,ci,str(value)) for ri,row in enumerate(matrix,1) for ci,value in enumerate(row,1)]


class Review2Tests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.workspace = Path(self.temp.name)

    def check(self, task, mode, expected, cells=None):
        code, log = run(task, self.workspace, mode, cells)
        self.assertEqual(code, expected, log)

    def test_hr_seed_positive_and_excel_time_cells(self):
        hr_fixture(self.workspace)
        self.check(HR, 'hr', 0)
        path = self.workspace/'Training_Travel_Report.xlsx'
        book = openpyxl.load_workbook(path)
        for row in book['Travel_Plan'].iter_rows(min_row=2):
            for cell in (row[3], row[4]):
                cell.value = datetime.strptime(cell.value, '%H:%M').time()
                cell.number_format = 'hh:mm'
        book.save(path)
        self.check(HR, 'hr', 0)

    def test_hr_missing_mandatory_budget_sheet_fails(self):
        hr_fixture(self.workspace)
        path=self.workspace/'Training_Travel_Report.xlsx';book=openpyxl.load_workbook(path)
        del book['Budget_Summary'];book.save(path)
        self.check(HR, 'hr', 1)

    def test_hr_changed_required_journey_fields_fail(self):
        for column, value in [(2, '718Г'), (4, '00:00'), (5, '00:01'), (6, 'Бизнес'), (7, 1)]:
            with self.subTest(column=column):
                hr_fixture(self.workspace)
                path=self.workspace/'Training_Travel_Report.xlsx';book=openpyxl.load_workbook(path)
                book['Travel_Plan'].cell(2,column,value);book.save(path)
                self.check(HR, 'hr', 1)

    def test_price_source_positive_and_missing_required_sheets(self):
        for sheet in (None, 'Ingredient Prices', 'Budget Summary'):
            with self.subTest(sheet=sheet):
                cells=price_fixture(self.workspace)
                if sheet:
                    path=self.workspace/'Recipe_Cost_Analysis.xlsx';book=openpyxl.load_workbook(path)
                    del book[sheet];book.save(path)
                self.check(PRICE, 'price', 1 if sheet else 0, cells)

    def test_all_tied_minimum_and_maximum_choices_pass(self):
        fixtures=[('Most_Expensive_Dish',
                   ['Щи из квашеной капусты','Жаркое в горшочках','Икра кабачковая','Картофельное пюре','Винегрет'],
                   ['Щи из квашеной капусты','Жаркое в горшочках']),
                  ('Cheapest_Dish',
                   ['Гречневая каша','Квас домашний','Икра кабачковая','Винегрет','Наполеон'],
                   ['Гречневая каша','Квас домашний'])]
        for metric,names,choices in fixtures:
            for choice in choices:
                with self.subTest(metric=metric,choice=choice):
                    cells=price_fixture(self.workspace,names)
                    path=self.workspace/'Recipe_Cost_Analysis.xlsx';book=openpyxl.load_workbook(path)
                    for row in book['Budget Summary'].iter_rows(min_row=2):
                        if row[0].value==metric:row[1].value=choice
                    book.save(path)
                    self.check(PRICE, 'price', 0, cells)

    def test_price_incorrect_extremum_fails(self):
        cells=price_fixture(self.workspace)
        path=self.workspace/'Recipe_Cost_Analysis.xlsx';book=openpyxl.load_workbook(path)
        for row in book['Budget Summary'].iter_rows(min_row=2):
            if row[0].value=='Cheapest_Dish':row[1].value='Бефстроганов'
        book.save(path)
        self.check(PRICE, 'price', 1, cells)

    def test_hr_budget_values_must_belong_to_the_correct_labels(self):
        for cell, value in [('B2', 99), ('C2', 1), ('D2', 1), ('A2', 'Итого')]:
            with self.subTest(cell=cell):
                hr_fixture(self.workspace)
                path=self.workspace/'Training_Travel_Report.xlsx';book=openpyxl.load_workbook(path)
                book['Budget_Summary'][cell]=value;book.save(path)
                self.check(HR, 'hr', 1)

    def test_ingredient_price_table_requires_complete_source_associations(self):
        for mutation in ['missing', 'unit', 'source', 'duplicate']:
            with self.subTest(mutation=mutation):
                cells=price_fixture(self.workspace)
                path=self.workspace/'Recipe_Cost_Analysis.xlsx';book=openpyxl.load_workbook(path)
                sheet=book['Ingredient Prices']
                if mutation=='missing':sheet.delete_rows(2)
                elif mutation=='unit':sheet['B2']='incorrect unit'
                elif mutation=='source':sheet['D2']='Борщ'
                else:
                    for column in range(1,5):sheet.cell(3,column,sheet.cell(2,column).value)
                book.save(path)
                self.check(PRICE, 'price', 1, cells)

    def terminal(self):
        from docx import Document
        names=['Борщ','Икра кабачковая','Котлеты домашние','Картофельное пюре','Морс клюквенный']
        records=json.loads(json.dumps([RECIPES[name] for name in names],ensure_ascii=False))
        doc=Document();doc.add_heading('Меню обедов на неделю',0)
        for day,recipe in zip(['Понедельник','Вторник','Среда','Четверг','Пятница'],records):
            doc.add_heading(day,1);doc.add_paragraph('Блюдо: '+recipe['name'])
            doc.add_paragraph('Категория: '+recipe['category']);doc.add_paragraph('Сложность: '+str(recipe['difficulty']))
            doc.add_paragraph('Ингредиенты: '+', '.join(item['name'] for item in recipe['ingredients']))
        doc.add_heading('Пищевые предпочтения',1)
        doc.add_paragraph('Сложность не выше 3; разнообразие категорий; один суп и овощное блюдо; без повторений.')
        doc.save(self.workspace/'Weekly_Menu.docx')
        for filename in ('all_recipes.json','filtered_recipes.json'):
            (self.workspace/filename).write_text(json.dumps(records,ensure_ascii=False))
        (self.workspace/'memory').mkdir(exist_ok=True)
        (self.workspace/'memory/memory.json').write_text(json.dumps({'entities':[{
            'name':'dietary_preferences','observations':['сложность <=3, категории разные, суп и овощи']+names}]},ensure_ascii=False))
        return records

    def test_terminal_required_saved_pool_is_valid_and_present(self):
        task='kulinar-terminal-menu-generator'
        self.terminal();self.check(task,'terminal',0)
        (self.workspace/'all_recipes.json').unlink()
        self.check(task,'terminal',1)

    def test_terminal_source_fields_and_ingredients_must_be_complete(self):
        task='kulinar-terminal-menu-generator'
        for mutation in ('missing_field','category','ingredients'):
            with self.subTest(mutation=mutation):
                self.terminal()
                path=self.workspace/'all_recipes.json';records=json.loads(path.read_text())
                if mutation=='missing_field':del records[0]['category']
                elif mutation=='category':records[0]['category']='fabricated'
                else:records[0]['ingredients'].pop()
                path.write_text(json.dumps(records,ensure_ascii=False))
                self.check(task,'terminal',1)

    def test_terminal_filter_keeps_all_and_only_source_records_at_difficulty_three_or_lower(self):
        task='kulinar-terminal-menu-generator'
        records=self.terminal()
        names={recipe['name'] for recipe in records}
        easy=next(recipe for recipe in RECIPES.values() if recipe['difficulty']<=3 and recipe['name'] not in names)
        hard=next(recipe for recipe in RECIPES.values() if recipe['difficulty']>3)
        (self.workspace/'all_recipes.json').write_text(json.dumps(records+[easy,hard],ensure_ascii=False))
        (self.workspace/'filtered_recipes.json').write_text(json.dumps(records+[easy],ensure_ascii=False))
        self.check(task,'terminal',0)
        (self.workspace/'filtered_recipes.json').write_text(json.dumps(records,ensure_ascii=False))
        self.check(task,'terminal',1)

    def test_terminal_ingredient_name_arrays_are_a_valid_record_representation(self):
        task='kulinar-terminal-menu-generator'
        records=self.terminal()
        for recipe in records:recipe['ingredients']=[item['name'] for item in recipe['ingredients']]
        for filename in ('all_recipes.json','filtered_recipes.json'):
            (self.workspace/filename).write_text(json.dumps(records,ensure_ascii=False))
        self.check(task,'terminal',0)

    def catering(self):
        book=openpyxl.Workbook();travel=book.active;travel.title='Travel'
        travel.append(['Train_No','Direction','Departure','Arrival','Duration_Min','Seat_Class','Per_Person_RUB','Total_Budget_RUB'])
        travel.append(['752А','Москва → Санкт-Петербург','06:50','10:50',240,'Эконом',5500,82500])
        menu=book.create_sheet('Menu')
        menu.append(['Course_Type','Dish_Name','Cook_Time_Min','Servings_Yield','Key_Ingredients','Notes'])
        for name in ['Винегрет','Икра кабачковая','Бефстроганов','Котлеты домашние','Курица в сметане','Борщ']:
            recipe=RECIPES[name];course='Закуска' if recipe['category'] in {'закуска','салат'} else recipe['category']
            menu.append([course,name,recipe['total_time_minutes'],recipe['servings'],
                         ', '.join(item['name'] for item in recipe['ingredients']),''])
        timeline=book.create_sheet('Timeline');timeline.append(['Time','Activity','Location','Responsible_Person','Notes'])
        for clock,activity,location in [('06:00','Дорога к вокзалу','Москва'),('06:30','Посадка','Вокзал'),
                ('11:00','Прибытие','СПб'),('12:00','Заселение','Гостиница'),('18:00','Ужин','Ресторан'),('20:30','Тимбилдинг','СПб')]:
            timeline.append([clock,activity,location,'Команда',''])
        path=self.workspace/'Team_Trip_Plan.xlsx';book.save(path);return path

    def test_catering_complete_source_positive(self):
        self.catering()
        self.check(CAT, 'catering', 0)

    def test_catering_fabricated_or_missing_mandatory_values_fail(self):
        for sheet,cell,value in [('Menu','B2','Fabricated Dish'),('Travel','H2',1),('Travel','D2','00:00')]:
            with self.subTest(sheet=sheet,cell=cell):
                path=self.catering();book=openpyxl.load_workbook(path);book[sheet][cell]=value;book.save(path)
                self.check(CAT, 'catering', 1)
        path=self.catering();book=openpyxl.load_workbook(path);book['Timeline'].delete_rows(7);book.save(path)
        self.check(CAT, 'catering', 1)

    @unittest.skipUnless(shutil.which('soffice') or shutil.which('libreoffice') or
                         Path('/Applications/LibreOffice.app/Contents/MacOS/soffice').is_file(),
                         'requires real LibreOffice Calc')
    def test_formula_outputs_remain_valid_in_all_three_graders(self):
        hr_fixture(self.workspace)
        path=self.workspace/'Training_Travel_Report.xlsx';book=openpyxl.load_workbook(path)
        book['Travel_Plan']['G2']='=4000+500';book['Budget_Summary']['D4']='=SUM(D2:D3)';book.save(path)
        self.check(HR, 'hr', 0)
        cells=price_fixture(self.workspace)
        path=self.workspace/'Recipe_Cost_Analysis.xlsx';book=openpyxl.load_workbook(path)
        book['Dish Costs']['C2']='=500+185';book['Dish Costs']['D2']='=ROUND(C2/4,2)';book.save(path)
        self.check(PRICE, 'price', 0, cells)
        path=self.catering();book=openpyxl.load_workbook(path)
        book['Travel']['G2']='=5000+500';book['Travel']['H2']='=G2*15';book.save(path)
        self.check(CAT, 'catering', 0)


if __name__ == '__main__':
    unittest.main()
