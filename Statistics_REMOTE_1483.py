import csv
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.styles import Font, Border, Side
import pdfkit
from jinja2 import Environment, FileSystemLoader
import prettytable
from prettytable import PrettyTable
import re

testStr = 'Ветка develop'

tax = {
    "True": "Без вычета налогов",
    "False": "С вычетом налогов"
}
translation = ['Название', 'Описание', 'Навыки', 'Опыт работы', 'Премиум-вакансия', 'Компания', 'Оклад',
               'Название региона', 'Дата публикации вакансии']
bools = {
    "True": "Да",
    "False": "Нет"
}
experience = {
    "noExperience": "Нет опыта",
    "between1And3": "От 1 года до 3 лет",
    "between3And6": "От 3 до 6 лет",
    "moreThan6": "Более 6 лет"
}
currencies = {
    "AZN": "Манаты",
    "BYR": "Белорусские рубли",
    "EUR": "Евро",
    "GEL": "Грузинский лари",
    "KGS": "Киргизский сом",
    "KZT": "Тенге",
    "RUR": "Рубли",
    "UAH": "Гривны",
    "USD": "Доллары",
    "UZS": "Узбекский сум"
}
currency_to_rub = {
    "AZN": 35.68,
    "BYR": 23.91,
    "EUR": 59.90,
    "GEL": 21.74,
    "KGS": 0.76,
    "KZT": 0.13,
    "RUR": 1,
    "UAH": 1.64,
    "USD": 60.66,
    "UZS": 0.0055,
}


def sort_exp(vac):
    s = re.findall(r'\d*\.\d+|\d+', vac.experience_id)
    return 0 if len(s) == 0 else int(s[0])


sorts = {
    'Название': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: vac.name),
    'Описание': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: vac.description),
    'Навыки': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: len(vac.key_skills)),
    'Опыт работы': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: sort_exp(vac)),
    'Премиум-вакансия': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: vac.premium),
    'Компания': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: vac.employer_name),
    'Оклад': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: vac.salary.salary_currency_to_ruble),
    'Название региона': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: vac.area_name),
    'Дата публикации вакансии': lambda vac, is_rev: vac.sort(reverse=is_rev, key=lambda vac: vac.published_at)
}
functions = {
    'Название': lambda vac, val: vac.name == val,
    'Опыт работы': lambda vac, val: vac.experience_id == val,
    'Описание': lambda vac, val: vac.description == val,
    'Дата публикации вакансии': lambda vac, val: f'{vac.published_at[8:10]}.{vac.published_at[5:7]}.{vac.published_at[:4]}' == val,
    'Премиум-вакансия': lambda vac, val: vac.premium == val,
    'Название региона': lambda vac, val: vac.area_name == val,
    'Компания': lambda vac, val: vac.employer_name == val,
    'Идентификатор валюты оклада': lambda vac, val: vac.salary.salary_currency == val,
    'Навыки': lambda vac, val: all(x in vac.key_skills for x in val.split(', ')),
    'Оклад': lambda vac, val: float(vac.salary.salary_from) <= float(val) <= float(vac.salary.salary_to)
}


class Vacancy:
    def __init__(self, vac):
        self.name = vac['name']
        self.salary_from = vac['salary_from']
        self.salary_to = vac['salary_to']
        self.salary_currency = vac['salary_currency']
        self.area_name = vac['area_name']
        self.published_at = vac['published_at']
        self.year = self.published_at[:4]
        self.salary = (float(self.salary_from) + float(self.salary_to)) / 2 * currency_to_rub[self.salary_currency]
        if len(vac) > 6:
            self.employer_name = vac['employer_name']
            self.description = vac['description']
            self.key_skills = vac['key_skills'].split('_')
            self.experience_id = experience[vac['experience_id']]
            self.premium = bools[vac['premium']]
            self.salary_gross = bools[vac['salary_gross']]

    def make_salary(self):
        sal_from = '{0:,}'.format(int(float(self.salary_from))).replace(',', ' ')
        sal_to = '{0:,}'.format(int(float(self.salary_to))).replace(',', ' ')
        return f'{sal_from} - {sal_to} ({self.salary_currency}) ({self.salary_gross})'

    def __str__(self):
        return [self.name, self.description, '\n'.join(self.key_skills), self.experience_id, self.premium,
                self.employer_name, self.make_salary(), self.area_name,
                f'{self.published_at[8:10]}.{self.published_at[5:7]}.{self.published_at[:4]}']


class DataSet:
    def __init__(self, f_name):
        self.file_name = f_name
        self.vacancies_objects = [Vacancy(obj) for obj in self.CSV_parser(self.file_name)]
        self.vac_amount = len(self.vacancies_objects)
        self.sal_by_years = {}
        self.sal_by_years_for_prof = {}
        self.sal_by_city = {}
        self.amount_by_years = {}
        self.amount_prof_by_years = {}
        self.amount_by_city = {}

    @staticmethod
    def strRefactor(str):
        str = re.sub(r"<[^>]*>", '', str)
        str = str.replace('\n', '_')
        str = ' '.join(str.split())
        return str

    @staticmethod
    def CSV_parser(csv_file):
        file = open(csv_file, encoding='utf-8-sig')
        csv_reader = csv.reader(file)
        titles = next(csv_reader)
        vacancies = [dict(zip(titles, [DataSet.strRefactor(s) for s in x])) for x in csv_reader if
                     '' not in x and len(x) == len(titles)]
        return vacancies

    @staticmethod
    def year_counter(sal, amount):
        for key in sal:
            sal[key] = int(sal[key] / amount[key]) if amount[key] != 0 else 0

    def city_counter(self, sal, amount):
        lst = []
        for key in sal:
            sal[key] = int(sal[key] / amount[key]) if amount[key] != 0 else 0
            amount[key] = round(amount[key] / self.vac_amount, 4) if self.vac_amount != 0 else 0
            if amount[key] < 0.01:
                lst.append(key)
        for key in lst:
            del sal[key]
            del amount[key]

    def make(self, prof_name):
        for vac in self.vacancies_objects:
            city = vac.area_name
            year = int(vac.year)
            if city not in self.sal_by_city:
                self.sal_by_city[city] = 0
                self.amount_by_city[city] = 0
            if year not in self.sal_by_years:
                self.sal_by_years[year] = 0
                self.amount_by_years[year] = 0
                self.sal_by_years_for_prof[year] = 0
                self.amount_prof_by_years[year] = 0
            if vac.name.find(prof_name) >= 0:
                self.sal_by_years_for_prof[year] += vac.salary
                self.amount_prof_by_years[year] += 1

            self.sal_by_city[city] += vac.salary
            self.amount_by_city[city] += 1
            self.sal_by_years[year] += vac.salary
            self.amount_by_years[year] += 1

        self.year_counter(self.sal_by_years, self.amount_by_years)
        self.year_counter(self.sal_by_years_for_prof, self.amount_prof_by_years)
        self.city_counter(self.sal_by_city, self.amount_by_city)

        self.sal_by_city = dict(sorted(self.sal_by_city.items(), key=lambda val: val[1], reverse=True)[:10])
        self.amount_by_city = dict(sorted(self.amount_by_city.items(), key=lambda val: val[1], reverse=True)[:10])


class Table:
    def __init__(self, f_name):
        self.f_name = f_name
        self.filter = input('Введите параметр фильтрации: ')
        self.sort_type = input('Введите параметр сортировки: ')
        self.is_rev_sort = input('Обратный порядок сортировки (Да / Нет): ')
        self.boarders = input('Введите диапазон вывода: ')
        self.need_titles = input('Введите требуемые столбцы: ')
        self.filter = self.param_fixer(self.filter, 'param')
        self.sort_type = self.param_fixer(self.sort_type, 'sort')
        self.is_rev_sort = self.param_fixer(self.is_rev_sort, 'rev')

        self.vacancies = DataSet(self.f_name)
        self.titles = translation

        self.boarders = self.boarders.split() if self.boarders else '0'
        self.need_titles = self.need_titles.split(', ') if self.need_titles else 'all'

    @staticmethod
    def param_fixer(p, type):
        if type == 'param':
            if not p:
                return 'nothing'
            if ':' not in p:
                print('Формат ввода некорректен')
                quit()
            p = p.split(': ')
            if p[0] not in translation and p[0] != 'Идентификатор валюты оклада':
                print('Параметр поиска некорректен')
                quit()
        if type == 'sort':
            if not p:
                return 'nothing'
            if p not in translation:
                print('Параметр сортировки некорректен')
                quit()
        if type == 'rev':
            if p not in ['Да', 'Нет', '']:
                print('Порядок сортировки задан некорректно')
                quit()
        return p

    def get_filtered(self, vac):
        if self.filter == 'nothing':
            return vac
        return functions[self.filter[0]](vac, self.filter[1])

    def sort_vac(self):
        if self.sort_type == 'nothing':
            return self.vacancies.vacancies_objects
        is_reverse = True if self.is_rev_sort == 'Да' else False
        sorts[self.sort_type](self.vacancies.vacancies_objects, is_reverse)
        return self.vacancies.vacancies_objects

    def print_table(self):
        table = PrettyTable()
        titles = self.titles.copy()
        titles.insert(0, '№')
        table.field_names = titles
        table.max_width = 20
        table.hrules = prettytable.ALL
        table.align = "l"

        if len(self.boarders) == 1:
            start = int(self.boarders[0]) - 1 if int(self.boarders[0]) != 0 else 0
            end = len(self.vacancies.vacancies_objects)
        else:
            start = int(self.boarders[0]) - 1 if int(self.boarders[0]) != 0 else 0
            end = int(self.boarders[1]) - 1 if int(self.boarders[1]) != 0 else 0
        if self.need_titles == 'all':
            self.need_titles = titles
        self.need_titles.insert(0, '№')
        i = 0
        for vac in self.sort_vac():
            row = vac.__str__()
            if not self.get_filtered(vac):
                continue
            for r in range(len(row)):
                if len(row[r]) > 100:
                    row[r] = f"{row[r][:100]}..."
            row.insert(0, i + 1)
            table.add_row(row)
            i += 1
        if i == 0:
            print('Ничего не найдено')
        else:
            print(table.get_string(start=start, end=end, fields=self.need_titles))


class Report:
    def __init__(self, f_name, prof_name):
        self.f_name = f_name
        self.prof_name = prof_name
        self.data_set = DataSet(self.f_name)
        self.data_set.make(self.prof_name)
        self.workbook = openpyxl.Workbook()
        self.years = [x for x in self.data_set.sal_by_years]
        self.fig, self.ax = plt.subplots(nrows=2, ncols=2)

    def print_data(self):
        print('Динамика уровня зарплат по годам:', self.data_set.sal_by_years)
        print('Динамика количества вакансий по годам:', self.data_set.amount_by_years)
        print('Динамика уровня зарплат по годам для выбранной профессии:', self.data_set.sal_by_years_for_prof)
        print('Динамика количества вакансий по годам для выбранной профессии:',
              self.data_set.amount_prof_by_years)
        print('Уровень зарплат по городам (в порядке убывания):', self.data_set.sal_by_city)
        print('Доля вакансий по городам (в порядке убывания):', self.data_set.amount_by_city)

    def generate_pdf(self):
        headers1, headers2, headers3, t1_data, t2_data, t3_data = self.make_data()
        env = Environment(loader=FileSystemLoader('.'))
        template = env.get_template("main.html")
        pdf_template = template.render(
            {'prof': self.prof_name, 'table1Headers': headers1, 'table1Data': t1_data, 'table2Headers': headers2,
             'table2Data': t2_data, 'table3Headers': headers3, 'table3Data': t3_data})
        options = {
            'enable-local-file-access': None
        }
        config = pdfkit.configuration(wkhtmltopdf=r'D:\wkhtmltopdf\bin\wkhtmltopdf.exe')
        pdfkit.from_string(pdf_template, 'report.pdf', configuration=config, options=options)

    def generate_image(self):
        labels = self.years
        sal = [x for x in self.data_set.sal_by_years.values()]
        job_sal = [x for x in self.data_set.sal_by_years_for_prof.values()]
        sal_count = [x for x in self.data_set.amount_by_years.values()]
        job_sal_count = [x for x in self.data_set.amount_prof_by_years.values()]
        cities = [self.make_transfer(x) for x in self.data_set.sal_by_city]
        sal_by_cities = [x for x in self.data_set.sal_by_city.values()]
        cities_percent = ['Другие'] + [x for x in self.data_set.amount_by_city]
        job_percent = [x for x in self.data_set.amount_by_city.values()]

        job_percent.insert(0, 1 - sum(job_percent))
        x = np.arange(len(labels))
        y = np.arange(len(cities))
        width = 0.35

        self.ax[0, 0].bar(x - width / 2, sal, width, label='средняя з/п')
        self.ax[0, 0].bar(x + width / 2, job_sal, width, label=f'з/п {self.prof_name}')
        self.ax[0, 1].bar(x - width / 2, sal_count, width, label='Количество вакансий')
        self.ax[0, 1].bar(x + width / 2, job_sal_count, width, label=f'Количество вакансий\n{self.prof_name}')
        self.ax[1, 0].barh(y, sal_by_cities)
        self.ax[1, 1].pie(job_percent, labels=cities_percent, startangle=40, textprops={'fontsize': 6})

        self.ax[0, 0].set_title('Уровень зарплат по годам', fontsize=10)
        self.ax[0, 1].set_title('Количество вакансий по годам', fontsize=10)
        self.ax[1, 0].set_title('Уровень зарплат по городам', fontsize=10)
        self.ax[1, 1].set_title('Доля вакансий по городам', fontsize=10)
        self.ax[0, 0].set_xticks(x, labels, fontsize=8)
        self.ax[0, 0].tick_params(axis='y', labelsize=8)
        self.ax[0, 1].set_xticks(x, labels, fontsize=8)
        self.ax[0, 1].tick_params(axis='y', labelsize=8)
        self.ax[1, 0].set_yticks(y, cities, fontsize=6)
        self.ax[1, 0].tick_params(axis='x', labelsize=8)

        self.ax[0, 0].tick_params(axis='x', labelrotation=90)
        self.ax[0, 1].tick_params(axis='x', labelrotation=90)
        self.ax[0, 0].legend(fontsize=8)
        self.ax[0, 1].legend(fontsize=8)
        self.ax[0, 0].grid(axis='y')
        self.ax[0, 1].grid(axis='y')
        self.ax[1, 0].grid(axis='x')
        self.ax[1, 0].invert_yaxis()
        self.fig.tight_layout()
        plt.show()
        self.fig.savefig('graph.png')

    def generate_excel(self):
        ws1 = self.workbook.active
        ws1.title = 'Статистика по годам'
        ws2 = self.workbook.create_sheet('Статистика по городам')

        f = Font(bold=True)
        sd = Side(border_style='thin', color='FF000000')
        b = Border(left=sd, right=sd, top=sd, bottom=sd)

        ws1['A1'] = 'Год'
        ws1['B1'] = 'Средняя зарплата'
        ws1['C1'] = 'Количество вакансий'
        ws1['D1'] = f'Средняя зарплата - {self.prof_name}'
        ws1['E1'] = f'Количество вакансий - {self.prof_name}'

        ws2['A1'] = ws2['D1'] = 'Город'
        ws2['B1'] = 'Уровень зарплат'
        ws2['E1'] = 'Доля вакансий'

        self.create_values(self.data_set.sal_by_years, ws1, 'A')
        self.create_values(self.data_set.sal_by_years.values(), ws1, 'B')
        self.create_values(self.data_set.amount_by_years.values(), ws1, 'C')
        self.create_values(self.data_set.sal_by_years_for_prof.values(), ws1, 'D')
        self.create_values(self.data_set.amount_prof_by_years.values(), ws1, 'E')
        self.create_values(self.data_set.sal_by_city, ws2, 'A')
        self.create_values(self.data_set.sal_by_city.values(), ws2, 'B')
        self.create_values(self.data_set.amount_by_city, ws2, 'D')
        self.create_values(self.data_set.amount_by_city.values(), ws2, 'E')

        row1 = ws1['1']
        row1_2 = ws2['1']
        col_e = ws2['E2':'E11']
        for r in row1:
            r.font = f
        for r in row1_2:
            r.font = f
        for r in col_e:
            for s in r:
                s.number_format = '0.00%'
        for r in ws1:
            for s in r:
                s.border = b
        for r in ws2:
            for s in r:
                if s.value:
                    s.border = b
        self.correct_rows(ws1)
        self.correct_rows(ws2)
        ws2.column_dimensions['C'].width = 2
        self.workbook.save('rep.xlsx')

    def make_data(self):
        headers1 = ['Год', 'Средняя зарплата', f'Средняя зарплата - {self.prof_name}', 'Количество вакансий',
                    f'Количество вакансий - {self.prof_name}']
        headers2 = ['Город', 'Уровень зарплат']
        headers3 = ['Город', 'Доля вакансий']
        t1_data = []
        t2_data = []
        t3_data = []

        for year in self.data_set.sal_by_years:
            temp = [year, self.data_set.sal_by_years[year], self.data_set.sal_by_years_for_prof[year],
                    self.data_set.amount_by_years[year], self.data_set.amount_prof_by_years[year]]
            t1_data.append(temp)
        for city in self.data_set.sal_by_city:
            temp = [city, self.data_set.sal_by_city[city]]
            t2_data.append(temp)
        for city in self.data_set.amount_by_city:
            temp = [city, str(round(self.data_set.amount_by_city[city] * 100, 2)).replace('.', ',') + '%']
            t3_data.append(temp)
        return headers1, headers2, headers3, t1_data, t2_data, t3_data

    @staticmethod
    def make_transfer(str):
        str = str.replace('-', '-\n')
        str = str.replace(' ', '\n')
        return str

    @staticmethod
    def correct_rows(ws):
        dims = {}
        for row in ws.rows:
            for cell in row:
                if cell.value:
                    dims[cell.column_letter] = max((dims.get(cell.column_letter, 0), len(str(cell.value)) + 2))
        for col, value in dims.items():
            ws.column_dimensions[col].width = value

    @staticmethod
    def create_values(val, cell, letter):
        for i, year in enumerate(val, start=2):
            n = f'{letter}{i}'
            cell[n] = year


class Input:
    def __init__(self):
        self.request = input('Вывести вакансии или статистику?')
        # self.f_name = input('Введите название файла: ')
        self.f_name = 'vacancies_big.csv'
        if self.request.casefold()[:4] == 'стат':
            # self.prof_name = input('Введите название профессии: ')
            self.prof_name = 'Программист'
            self.report = Report(self.f_name, self.prof_name)
            self.report.generate_pdf()
            self.report.print_data()
        else:
            self.table = Table(self.f_name)
            self.table.print_table()


x = Input()
