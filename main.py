import os
import datetime
import re

from bs4 import BeautifulSoup
from loguru import logger

from typing import List, Optional, Dict

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Table
from RPA.FileSystem import FileSystem
from RPA.PDF import PDF

from config import agency


browser_lib = Selenium()
file_system_lib = FileSystem()


def get_amounts_for_each_agency() -> List[Dict]:
    browser_lib.click_element_when_visible('partial link:DIVE IN')
    browser_lib.wait_until_element_is_visible('id:agency-tiles-widget')
    agency_names = browser_lib.get_webelements('css:div#agency-tiles-widget span.h4.w200')
    agency_amounts = browser_lib.get_webelements('css:div#agency-tiles-widget span.h1.w900')
    agencies = list(zip(agency_names, agency_amounts))
    content = [{'Agency name': agency[0].text, 'Agency amount': agency[1].text} for agency in agencies]
    return content


def write_excel_worksheet_agencies(path: str, worksheet: str, content: Optional[List[Dict]]) -> None:
    lib = Files()
    lib.create_workbook(path)
    try:
        lib.create_worksheet(worksheet, content, header=True)
    finally:
        lib.save_workbook()
        lib.close_workbook()


def select_one_of_the_agencies(agency_name: str, url: str) -> None:
    agency_info = browser_lib.get_element_attribute(f'partial link:{agency_name}', 'outerHTML')
    link = agency_info.split('>')[0].split('"')[1].lstrip('/')
    browser_lib.go_to(url + link)


def get_agency_individual_investments_table() -> str:
    delta = datetime.timedelta(seconds=20)
    browser_lib.wait_until_page_contains_element('id:investments-table-widget', delta)
    browser_lib.select_from_list_by_value('name:investments-table-object_length', '-1')
    browser_lib.wait_until_page_does_not_contain_element(
        'css:#investments-table-object_paginate > span > a:nth-child(2)', delta)
    html = browser_lib.get_element_attribute('css:table#investments-table-object', 'innerHTML')
    return html


def scrape_agency_individual_investments_table(html: str) -> Table:
    soup = BeautifulSoup(html, 'lxml')
    headers = []
    for cell in soup.find('thead').select('th'):
        header = cell.find('div', class_='dataTables_sizing').text
        headers.append(header)
    table_rows = []
    for table_row in soup.find('tbody').select('tr'):
        cells = table_row.find_all('td')
        cell_values = []
        for cell in cells:
            cell_values.append(cell.text.strip())
        cell_values_dict = dict(zip(headers, cell_values))
        table_rows.append(cell_values_dict)
    return Table(table_rows)


def add_excel_worksheet_table(path: str, worksheet: str, content: Optional[Table]) -> None:
    lib = Files()
    lib.open_workbook(path)
    try:
        lib.create_worksheet(worksheet, content, header=True)
    finally:
        lib.save_workbook()
        lib.close_workbook()


def download_business_case_pdf(html: str, url: str) -> None:
    # files = file_system_lib.list_files_in_directory('output/')
    # paths = [f'output/{f.name}' for f in files if 'excel' not in f'{f.name}']
    # file_system_lib.remove_files(*paths, missing_ok=True)

    soup = BeautifulSoup(html, 'lxml')
    links = [url + link.get('href').lstrip('/') for link in soup.find_all('a')]
    for link in links:
        file_name = link.split('/')[-1].rstrip("'")
        file_system_lib.remove_file(f'output/{file_name}.pdf', missing_ok=True)
        browser_lib.go_to(link)
        browser_lib.wait_until_page_contains_element('link:Download Business Case PDF')
        browser_lib.click_element_if_visible('link:Download Business Case PDF')
        file_system_lib.wait_until_created(path=f'output/{file_name}.pdf', timeout=20.0)
        logger.debug(f"File {file_name}.pdf was downloaded.")
        browser_lib.go_back()


def remove_the_duplicate_files_from_the_folder() -> None:
    file_path = os.path.abspath('output/')
    file_list = os.listdir('output/')
    for file_name in file_list:
        if 'excel' in file_name:
            continue
        if "pdf" not in file_name:
            os.remove(os.path.join(file_path, file_name))
        pattern = r'\ \(\d{1,9}\)'
        if re.search(pattern, file_name) is None:
            continue
        original_file_name = re.sub(pattern=pattern, repl='', string=file_name)
        if not os.path.exists(os.path.join(file_path, original_file_name)):
            continue  # do not remove files which have no original
        os.remove(os.path.join(file_path, file_name))


def extract_data_from_pdf() -> List:
    file_list = [file_name for file_name in os.listdir('output/') if 'excel' not in file_name]
    pdf_values = []
    for file_name in file_list:
        page = PDF().get_text_from_pdf(source_path=f'output/{file_name}', pages=1).get(1)
        values = page.replace(f'\n', ' ').split(
            'Name of this Investment: ')[-1].split(
            "Section B")[0].split(
            '2. Unique Investment Identifier (UII): ')
        pdf_values.append(values[::-1])
    return pdf_values


def compare_values(pdf_values: list, content: Optional[Table]) -> None:
    p = len(pdf_values)
    k = 0
    n = len(content.to_list(with_index=False))
    for i in range(p):
        for j in range(n):
            pdf_value = pdf_values[i]
            web_value = list(content.get_row(j).values())
            if set(pdf_value).issubset(web_value):
                msg = f'Unique Investment Identifier (UII): {pdf_value[0]} is equal UII: {web_value[0]}, ' \
                      f'Name of this Investment: "{pdf_value[1]}" is equal Investment Title: "{web_value[2]}"'
                logger.info(msg)
                k += 1
    logger.info(f"{k} out of {p} values are equal.")


def main():
    url = 'https://itdashboard.gov/'
    file_system_lib.create_directory(path='output/', exist_ok=True)
    browser_lib.set_download_directory(directory=os.path.abspath('output/'), download_pdf=True)
    browser_lib.open_available_browser(url)
    logger.add("log.log", format="{time} {level} {message}", level="INFO",
               rotation="1 MB", compression="zip")

    try:
        content = get_amounts_for_each_agency()
        write_excel_worksheet_agencies('output/excel.xlsx', 'Agencies', content)
        select_one_of_the_agencies(agency, url)
        html = get_agency_individual_investments_table()
        content = scrape_agency_individual_investments_table(html)
        add_excel_worksheet_table('output/excel.xlsx', 'Table', content)
        download_business_case_pdf(html, url)
        remove_the_duplicate_files_from_the_folder()
        pdf_values = extract_data_from_pdf()
        compare_values(pdf_values, content)
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
