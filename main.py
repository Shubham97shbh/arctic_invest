from bs4 import BeautifulSoup as bs
import requests
import pandas as pd
import api
import openpyxl
import lxml
import json
import wkhtmltopdf
from wkhtmltopdf.main import WKhtmlToPdf
import pdfkit

URL = 'https://brevets-patents.ic.gc.ca/opic-cipo/cpd/eng/patent'
Params = {'type': 'number_search', 'tabs1Index': 'tabs1_1'}

data = pd.read_excel('data/input_sheet.xlsx')
data = pd.DataFrame(data)


# Saving the dat for new file to work

def save_f(result, app_no):
    with open(f'output/html/CA_{app_no}.html', 'w') as file:
        file.write(result)
    try:
        with open(f'output/html/CA_{app_no}.html') as file:
            pdfkit.from_file(file, f"output/pdf_html/CA_{app_no}.pdf")
    except:
        pass


def save_j(result_j):
    with open('output/excel/result.json', 'w') as file:
        file.write(str(result_j))


def web_scraper(data, Params):
    result_json = {}
    for i in data['Application no.']:
        URL_d = f"https://brevets-patents.ic.gc.ca/opic-cipo/cpd/eng/patent/{i}/summary.html"
        response = requests.get(url=URL_d, params=Params).text
        result_html = bs(response, 'lxml')
        try:
            new_json = data_check(result_html)
            result_json[i] = new_json
        except:
            pass

        save_f(result_html.prettify(), i)

    # coverting to json file to xlxl file
    # df_json = pd.read_json(str(result_json))
    # df_json.to_excel('DATAFILE.xlsx')
    save_j(result_json)


def data_check(result):
    val = result.find('div', id="tabs").find('div', {'class': "tgl-panel"})
    # converting different params as a json format

    ipc = val.findAll('span', {'class': "IPC-LEVEL-A IPC-VALUE-I"})
    re_j = {'IPC': [i.text for i in ipc]}
    # multi line
    finds = ['inventors', 'owners', 'applicants']
    for j in finds:
        re_j[j] = [v.text for v in val.find('td', {'headers': j}).find('li').findAll('strong')]
    # single line
    finds = ['agent', "associateAgent", "issued", "filingDate", "pubDate", "examDate", "lic", "lang"]
    for j in finds:
        try:
            re_j[j] = val.find('td', {'headers': j}).find('strong').text.replace('\n', '').replace('\t', '').replace(
                ' ', '')
        except:
            re_j[j] = val.find('td', {'headers': j}).find('strong').text

    finds = val.find('table', {'class': "table table-bordered col-lg-12", 'id': "pctTable"}).find('td', {
        'headers': "pct"}).find('strong').text
    re_j['pctTable'] = finds

    return re_j


if __name__ == '__main__':
    web_scraper(data, Params)
