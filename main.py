import os
import datetime
from typing import List, Optional

from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Tables import Table

from SeleniumLibrary.errors import ElementNotFound
from bs4 import BeautifulSoup


browser_lib = Selenium()


def get_amounts_for_each_agency(url) -> List:
    browser_lib.open_available_browser(url)
    browser_lib.set_browser_implicit_wait(datetime.timedelta(seconds=5))

    locators = [f'//*[@id="agency-tiles-2-widget"]/div/div[{i}]/div[{j}]/div/div/div/div[1]/a' \
                for i in range(1, 10) for j in range(1, 4)]

    amounts = []
    for locator in locators:
        try:
            agency_name = browser_lib.get_webelement(locator + '/span[1]').get_attribute('innerHTML')
            amount = browser_lib.get_webelement(locator + '/span[2]').get_attribute('innerHTML')
            amounts.append((agency_name, amount))
        except ElementNotFound:
            return amounts


def write_excel_worksheet_agencies(path: str, worksheet: str, content: Optional[list]) -> None:
    lib = Files()
    lib.create_workbook(path)
    try:
        lib.create_worksheet(worksheet, content)
    finally:
        lib.save_workbook()
        lib.close_workbook()


def select_one_of_the_agencies(agency: str, url: str) -> int:
    locators = [f'//*[@id="agency-tiles-2-widget"]/div/div[{i}]/div[{j}]/div/div/div/div[1]'
                for i in range(1, 10) for j in range(1, 4)]

    browser_lib.set_browser_implicit_wait(datetime.timedelta(seconds=5))

    for locator in locators:
        try:
            web_element = browser_lib.get_webelement(locator).get_attribute('innerHTML')
            if agency in web_element:
                link = web_element.split('>')[0].split('"')[1].lstrip('/')
                return browser_lib.open_available_browser(url + link)
        except ElementNotFound:
            pass


def get_agency_individual_investments_table() -> str:
    delta = datetime.timedelta(seconds=30)

    browser_lib.set_browser_implicit_wait(delta)
    browser_lib.select_from_list_by_value('name:investments-table-object_length', '-1')
    browser_lib.wait_until_page_does_not_contain_element(
        'xpath://*[@id="investments-table-object_paginate"]/span/a[2]', delta)

    return browser_lib.get_webelement('id:investments-table-object').get_attribute('outerHTML')


def scrape_agency_individual_investments_table(html: str) -> Table:
    soup = BeautifulSoup(html, "html.parser")

    table_rows = []
    for table_row in soup.select('tr'):
        cells = table_row.find_all('td')
        if len(cells) > 0:
            cell_values = []
            for cell in cells:
                cell_values.append(cell.text.strip())
            table_rows.append(cell_values)

    return Table(table_rows)


def add_excel_worksheet_table(path: str, worksheet: str, content: Optional[Table]) -> None:
    lib = Files()
    lib.open_workbook(path)
    try:
        lib.create_worksheet(worksheet, content)
    finally:
        lib.save_workbook()
        lib.close_workbook()


def download_business_case_pdf(html: str, url: str) -> None:
    """ WIP: Don't download a bunch of files at once :) """
    soup = BeautifulSoup(html, "html.parser")
    links = [url + link.get('href').lstrip('/') for link in soup.find_all('a')]

    browser_lib.set_download_directory(directory=os.path.abspath('output/'), download_pdf=True)
    for link in links:
        browser_lib.open_available_browser(link)
        browser_lib.wait_until_page_contains_element('link:Download Business Case PDF')
        browser_lib.click_link('link:Download Business Case PDF')


def main():
    url = 'https://itdashboard.gov/'
    try:
        content = get_amounts_for_each_agency(url)
        write_excel_worksheet_agencies('output/excel.xlsx', 'Agencies', content)
        select_one_of_the_agencies('National Science Foundation', url)
        html = get_agency_individual_investments_table()
        content = scrape_agency_individual_investments_table(html)
        add_excel_worksheet_table('output/excel.xlsx', 'Table', content)
        download_business_case_pdf(html, url)
    finally:
        browser_lib.close_all_browsers()


if __name__ == "__main__":
    main()
