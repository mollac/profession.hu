# coding=utf-8
import sys
import requests
import xlsxwriter
from bs4 import BeautifulSoup

if sys.platform == 'linux':
    excelFile = r'/home/molla/Asztal/profession.xlsx'
else:
    excelFile = r'C:\Users\admin\Desktop\profession.xlsx'

pageNum = 1
links = []
jobs = []
firms = []
addresses = []
salaries = []
print('Adatgyűjtés: ', end='', flush=True)

while True:
    url = f"https://www.profession.hu/allasok/gyor-moson-sopron/{pageNum},0,31"
    print('.', end='', flush=True)
    page = requests.get(url)
    soup = BeautifulSoup(page.content, 'html.parser')
    all_card = soup.find_all(class_="card")

    for card in all_card:
        job = card.find(class_="job-card__title").get_text().strip()
        jobs.append(job.splitlines()[0])

        link = card.select('h2 a')
        links.append(link[0]['href'])

        firm = card.find(class_='job-card__company-name').get_text().strip().replace('"', '')
        firms.append(firm)

        address = card.find(class_='job-card__company-address').get_text().strip()
        addresses.append(address)

        salary = card.select('.bonus_salary > dd:nth-child(2)')
        if salary:
            salary = salary[0].text
        else:
            salary = ''
        salaries.append(salary)

    next_btn = soup.find(class_='next')
    if not next_btn:
        break
    pageNum += 1
print(' kész!')

print(f'Excel fájl létrehozása: {excelFile} ->', end=' ')
headers = ['Munkakör', 'Hírdető', 'Cím', 'Bér', 'Link']

wb = xlsxwriter.Workbook(excelFile)
ws = wb.add_worksheet()
ws.set_column('A:C', 35)
ws.set_column('D:E', 12)
ws.set_default_row(20)
ws.set_zoom(100)
row = 1

cf = wb.add_format()
cf.set_font_size(9)
cf.set_align('vcenter')
cf.set_bold(True)

for i in range(len(headers)):
    ws.write_string(0, i, headers[i], cf)

for x in range(len(jobs)):
    cf = wb.add_format()
    cf.set_font_size(9)
    cf.set_align('vcenter')

    ws.write_string(row, 0, jobs[x], cf)
    ws.write_string(row, 1, firms[x], cf)
    ws.write_string(row, 2, addresses[x], cf)
    ws.write_string(row, 3, salaries[x], cf)
    ws.write_url(row, 4, links[x], cf, string="Megnézem")
    row += 1
wb.close()
print('kész!')
