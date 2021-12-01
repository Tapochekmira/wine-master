import datetime
import pandas
import collections


from pprint import pprint
from http.server import HTTPServer, SimpleHTTPRequestHandler
from jinja2 import Environment, FileSystemLoader, select_autoescape


def the_ending(age):
    if (age % 10 == 1) and (age != 11) and (age != 111):
        return "год"
    elif (age % 10 > 1) and (age % 10 < 5) and (age != 12) and (age != 13) and (age != 14):
        return "года"
    else:
        return 'лет'

env = Environment(
    loader=FileSystemLoader('.'),
    autoescape=select_autoescape(['html', 'xml'])
)

template = env.get_template('template.html')


# Считаем возраст винодельни и ставим правильный вариант год/года/лет
year_of_foundation = 1920
year_now = datetime.datetime.now().year
age = year_now - year_of_foundation
ending = the_ending(age)
final = f'{age} {ending}'

# Считываем из файла инфу про вина
wines_pandas = pandas.read_excel('wine3.xlsx', na_values='None', keep_default_na=False)
wines = wines_pandas.to_dict(orient='records')

headers = wines_pandas.columns.ravel()
wines_dict = collections.defaultdict(list)

for wine in wines:
    wines_dict[wine[headers[0]]].append(wine)

categories = sorted(list(wines_dict.keys()))


rendered_page = template.render(
    headers=headers,
    year=final,
    wines=wines_dict,
    categories=categories
)

with open('index.html', 'w', encoding="utf8") as file:
    file.write(rendered_page)

server = HTTPServer(('0.0.0.0', 8000), SimpleHTTPRequestHandler)
server.serve_forever()