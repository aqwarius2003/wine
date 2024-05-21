import pandas as pd
import datetime
from collections import defaultdict

from http.server import HTTPServer, SimpleHTTPRequestHandler

from jinja2 import Environment, FileSystemLoader, select_autoescape

foundation_date = 1920

df = pd.read_excel(io='wine.xlsx', sheet_name='Лист1', na_values=' ', keep_default_na=False)
wines_dic = df.to_dict(orient='records')

alco_in_category = defaultdict(list)

for wine in wines_dic:
    category = wine["Категория"]
    alco_in_category[category].append(wine)


def age_word(age):
    diglast = age % 10
    if age in range(5, 20):
        return 'лет'
    elif 1 in (age, diglast):
        return 'год'
    elif {age, diglast} & {2, 3, 4}:
        return 'года'
    else:
        return 'лет'


current_year = datetime.datetime.now().year
age = current_year - foundation_date

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')

rendered_page = template.render(
    age=f'{age} age_word(age)',
    alco_in_category=alco_in_category
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()
