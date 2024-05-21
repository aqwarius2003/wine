import pandas as pd
import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape

foundation_date = 1920


def word_age(age):
    diglast = age % 10
    if (age % 100) in range(11, 15):
        return 'лет'
    elif diglast == 1:
        return 'год'
    elif diglast in [2, 3, 4]:
        return 'года'
    else:
        return 'лет'


def main():
    df = pd.read_excel(io='wine.xlsx', sheet_name='Лист1', na_values=' ', keep_default_na=False)
    wines = df.to_dict(orient='records')

    wine_categories = defaultdict(list)

    for wine in wines:
        category = wine["Категория"]
        wine_categories[category].append(wine)

    current_year = datetime.datetime.now().year
    age = current_year - foundation_date

    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )

    template = env.get_template('template.html')

    rendered_page = template.render(
        age=f'{age} {word_age(age)}',
        wine_categories=wine_categories
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
