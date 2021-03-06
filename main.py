import os
import datetime
import re
from collections import namedtuple
from typing import Dict, List, Optional

from loguru import logger
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
    names = browser_lib.get_webelements('css:div#agency-tiles-widget span.h4.w200')
    amounts = browser_lib.get_webelements('css:div#agency-tiles-widget span.h1.w900')
    content = [{'Agency name': name.text, 'Agency amount': amount.text} for name, amount in zip(names, amounts)]
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


def get_agency_individual_investments_table():
    delta = datetime.timedelta(seconds=20)
    browser_lib.wait_until_page_contains_element('id:investments-table-object', delta)
    browser_lib.select_from_list_by_value('name:investments-table-object_length', '-1')
    browser_lib.wait_until_page_does_not_contain_element(
        'css:#investments-table-object_paginate > span > a:nth-child(2)', delta)
    table = browser_lib.get_webelement('id:investments-table-object')
    thead = table.find_element_by_tag_name('thead').find_elements_by_tag_name('th')
    headers = []
    for th in thead:
        header = th.find_element_by_tag_name('div').get_attribute('innerHTML')
        headers.append(header)
    tbody = table.find_element_by_tag_name('tbody').find_elements_by_tag_name('tr')
    rows = []
    for tr in tbody:
        tds = tr.find_elements_by_tag_name('td')
        row = [td.text for td in tds]
        rows.append(row)
    content = [dict(zip(headers, row)) for row in rows]
    return Table(content)


def add_excel_worksheet_table(path: str, worksheet: str, content: Optional[Table]) -> None:
    lib = Files()
    lib.open_workbook(path)
    try:
        lib.create_worksheet(worksheet, content, header=True)
    finally:
        lib.save_workbook()
        lib.close_workbook()


def download_business_case_pdf() -> None:
    links = browser_lib.get_webelements('css:table#investments-table-object > tbody > tr > td > a')
    urls = [element.get_attribute('href') for element in links]
    names = [element.text + '.pdf' for element in links]
    Files = namedtuple('Files', 'name url')
    files = [Files(name, url) for name, url in zip(names, urls)]
    for file in files:
        path = os.path.join('output', file.name)
        file_system_lib.remove_file(path, missing_ok=True)
        browser_lib.go_to(file.url)
        browser_lib.wait_until_page_contains_element('link:Download Business Case PDF')
        browser_lib.click_element_if_visible('link:Download Business Case PDF')
        file_system_lib.wait_until_created(path, timeout=60.0)
        logger.debug(f"File {file.name} was downloaded.")
        browser_lib.go_back()


def remove_the_duplicate_files_from_the_folder(folder: str) -> None:
    file_list = os.listdir(folder)
    for file_name in file_list:
        if 'excel' in file_name:
            continue
        if "pdf" not in file_name:
            os.remove(os.path.join(folder, file_name))
        pattern = r'\([0-9]+\).pdf'
        if re.search(pattern, file_name) is None:
            continue
        os.remove(os.path.join(folder, file_name))


def extract_data_from_pdf() -> List:
    file_list = [file_name for file_name in os.listdir('output') if 'excel' not in file_name]
    pdf_values = []
    for file_name in file_list:
        page = PDF().get_text_from_pdf(source_path=os.path.join('output', file_name), pages=1).get(1)
        values = page.replace(f'\n', ' ').split(
            'Name of this Investment: ')[-1].split(
            "Section B")[0].split(
            '2. Unique Investment Identifier (UII): ')
        pdf_values.append(values[::-1])
    return pdf_values


def compare_values(pdf_values: list, content: Optional[Table]) -> None:
    p = len(pdf_values)
    n = len(content.to_list(with_index=False))
    k = 0
    for i in range(n):
        web_value = content.get_row(index=i, columns=[0, 2], as_list=True)
        if web_value in pdf_values:
            msg = f'Unique Investment Identifier (UII): {web_value[0]} is equal UII, ' \
                  f'Name of this Investment: "{web_value[1]}" is equal Investment Title'
            logger.info(msg)
            k += 1
    logger.info(f"{k} out of {p} values are equal.")


def main():
    url = 'https://itdashboard.gov/'
    file_system_lib.create_directory('output', exist_ok=True)
    output = os.path.abspath('output')
    excel_path = os.path.join('output', 'excel.xlsx')
    browser_lib.set_download_directory(output, download_pdf=True)
    browser_lib.open_available_browser(url)
    logger.add("log.log", format="{time} {level} {message}", level="INFO",
               rotation="1 MB", compression="zip")

    try:
        content = get_amounts_for_each_agency()
        write_excel_worksheet_agencies(excel_path, 'Agencies', content)
        select_one_of_the_agencies(agency, url)
        content = get_agency_individual_investments_table()
        add_excel_worksheet_table(excel_path, 'Table', content)
        download_business_case_pdf()
        remove_the_duplicate_files_from_the_folder(output)
        pdf_values = extract_data_from_pdf()
        compare_values(pdf_values, content)
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
