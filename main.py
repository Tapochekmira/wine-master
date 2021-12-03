import collections
import datetime
import pandas
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def age_caption(age):
    if (age % 10 == 1) and (age != 11) and (age != 111):
        return "год"
    elif (
            (age % 10 > 1)
            and (age % 10 < 5)
            and (age != 12)
            and (age != 13)
            and (age != 14)
    ):
        return "года"
    else:
        return 'лет'


def main():
    env = Environment(
        loader=FileSystemLoader('.'),
        autoescape=select_autoescape(['html', 'xml'])
    )
    template = env.get_template('template.html')

    # Считаем возраст винодельни и ставим правильный вариант год/года/лет
    year_of_foundation = 1920
    year_now = datetime.datetime.now().year
    age = year_now - year_of_foundation
    winery_age = f'{age} {age_caption(age)}'

    # Считываем из файла инфу про вина
    wines = pandas.read_excel(
        'wine.xlsx',
        na_values='None',
        keep_default_na=False
    ).to_dict(orient='records')

    grouped_wines = collections.defaultdict(list)
    for wine in wines:
        grouped_wines[wine['Категория']].append(wine)

    # Отправляем в шаблон инфу
    rendered_page = template.render(
        years=winery_age,
        grouped_wines=grouped_wines,
    )
    with open('index.html', 'w', encoding="utf8") as file:
        file.write(rendered_page)

    # Запускаем сервер
    server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
    server.serve_forever()


if __name__ == '__main__':
    main()
