import csv
import matplotlib.pyplot as plt
import numpy as np
import openpyxl
from openpyxl.styles import Font, Border, Side
import pdfkit
from jinja2 import Environment, FileSystemLoader

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
    def CSV_parser(csv_file):
        file = open(csv_file, encoding='utf-8-sig')
        csv_reader = csv.reader(file)
        titles = next(csv_reader)
        vacancies = [dict(zip(titles, x)) for x in csv_reader if '' not in x and len(x) == len(titles)]
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


class Report:
    def __init__(self, f_name, prof_name):
        self.f_name = f_name
        self.prof_name = prof_name
        self.data_set = DataSet(self.f_name)
        self.data_set.make(self.prof_name)
        self.workbook = openpyxl.Workbook()
        self.years = [x for x in self.data_set.sal_by_years]
        self.fig, self.ax = plt.subplots(nrows=2, ncols=2)

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
        # self.f_name = input('Введите название файла: ')
        # self.prof_name = input('Введите название профессии: ')
        self.f_name = 'vac.csv'
        self.prof_name = 'Программист'
        self.report = Report(self.f_name, self.prof_name)

    def print(self):
        print('Динамика уровня зарплат по годам:', self.report.data_set.sal_by_years)
        print('Динамика количества вакансий по годам:', self.report.data_set.amount_by_years)
        print('Динамика уровня зарплат по годам для выбранной профессии:', self.report.data_set.sal_by_years_for_prof)
        print('Динамика количества вакансий по годам для выбранной профессии:',
              self.report.data_set.amount_prof_by_years)
        print('Уровень зарплат по городам (в порядке убывания):', self.report.data_set.sal_by_city)
        print('Доля вакансий по городам (в порядке убывания):', self.report.data_set.amount_by_city)


x = Input()
x.report.generate_pdf()
