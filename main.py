import pandas as pd
import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import argparse
import os
from dotenv import load_dotenv
import click
from pathlib import Path
from urllib.parse import urlparse

FOUNDATION_DATE = 1920


def fetch_year_word(age):
    diglast = age % 10
    if (age % 100) in range(11, 15):
        return 'лет'
    elif diglast == 1:
        return 'год'
    elif diglast in [2, 3, 4]:
        return 'года'
    else:
        return 'лет'


def create_parser():
    description = '''Принимает Exel  файл с базой товаров.'''
    parser = argparse.ArgumentParser(description=description)
    h = 'Введите путь и название файла с расширением xlsx или xls'
    parser.add_argument('-f', '--exel_file', help=h, default='wine.xlsx', type=str)
    return parser


def load_sheet(sheet_path, sheet_name):
    if sheet_path.startswith("https://"):
        sheet_link = sheet_path
        parsed_url = urlparse(sheet_link)
        file_id = parsed_url.path.split('/')[3]
        df = pd.read_csv(f'https://docs.google.com/spreadsheets/d/{file_id}/export?format=csv')
        print("путь: ", parsed_url, '\n',file_id)
    else:
        excel_file = Path(sheet_path, sheet_name)
        print("путь: ", excel_file)
        df = pd.read_excel(excel_file, na_values=' ', sheet_name='Лист1', engine='openpyxl', keep_default_na=False)
    print(df)
    return df


@click.command()
@click.option('--sheet_path',
              prompt="введите путь к локальному файлу или вставьте ссылку таблицы Google Sheets",
              default="", required=True,
              help='Путь к файлу: (если в корневой папке проекта - нажми Enter)')
@click.option('--sheet_name', default="wine.xlsx", prompt="Название файла: ",
              required=True, help="введите название таблицы")
def main(sheet_path, sheet_name):
    df = load_sheet(sheet_path, sheet_name)

    wines = df.to_dict(orient='records')

    wine_categories = defaultdict(list)

    for wine in wines:
        category = wine["Категория"]
        wine_categories[category].append(wine)

    current_year = datetime.datetime.now().year
    age = current_year - FOUNDATION_DATE

    load_dotenv()
    template_path = os.getenv('TEMPLATE_PATH', default='.')
    template_extensions = os.getenv(
        'TEMPLATE_EXTENSIONS', default=['html', 'xml']
    )
    host_address = os.getenv('HOST_ADDRESS', default='0.0.0.0')
    host_port = int(os.getenv('HOST_PORT', default=8000))
    template_name = os.getenv('TEMPLATE_NAME', default='template.html')
    env = Environment(
        loader=FileSystemLoader(template_path),
        autoescape=select_autoescape(template_extensions)
    )
    template = env.get_template(template_name)

    rendered_page = template.render(
        age=f'{age} {fetch_year_word(age)}',
        wine_categories=wine_categories
    )

    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    server = HTTPServer((host_address, host_port), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
