import pandas as pd
import datetime
from collections import defaultdict
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape
import os
from dotenv import load_dotenv
import click
from pathlib import Path

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


def load_local_sheet(table_path):
    excel_file = Path(table_path)
    df = pd.read_excel(excel_file, na_values=' ', sheet_name='Лист1', engine='openpyxl', keep_default_na=False)
    return df


@click.command()
@click.option('--table_path',
              prompt="Введите путь и название локального файла с расширением xlsx или xls",
              default="wine.xlsx", required=True,
              help='Путь к каталогу файла и имя файла')
def main(table_path):
    try:
        df = load_local_sheet(table_path)
    except FileNotFoundError:
        print("Файл не найден, проверьте имя или путь к файлу")
        return
    except pd.errors.ParserError:
        print("Ошибка парсинга файла")
        return

    if df is not None and not df.empty:
        wines = df.to_dict(orient='records')
    else:
        print('Убедитесь в правильности пути и перезапустите программу')
        exit()

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
