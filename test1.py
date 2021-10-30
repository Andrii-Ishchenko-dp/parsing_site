from RPA.Browser.Selenium import Selenium
import time
import xlwt
from xlrd import open_workbook
from xlutils.copy import copy


browser_lib = Selenium()


def open_the_website(url):
    browser_lib.open_available_browser(url)


def click_button():

    info = "class:btn-lg-2x"
    browser_lib.click_element(info)
    time.sleep(5)


def write_file():
    book = xlwt.Workbook(encoding="utf-8")
    sheet1 = book.add_sheet("Agencies")
    sheet2 = book.add_sheet("Info")
    st=0
    st_1=0
    info = 'xpath://*[@id="agency-tiles-widget"]/div/div/div/div/div/div/div[1]/a/span'
    result = [i.text for i in browser_lib.find_elements(info)]
    for h in result[0::2]:
        sheet1.write(st, 0, h)
        book.save("test_sheet1.xls")
        st+=1
    for m in result[1::2]:
        sheet1.write(st_1, 1, m)
        book.save("test_sheet1.xls")
        st_1+=1
    time.sleep(2)


def search_depart(depart):

    browser_lib.click_link(depart)
    time.sleep(10)

def selector_click():
    element='xpath://*[@id="investments-table-object_length"]/label/select'
    element1='xpath://*[@id="investments-table-object_length"]/label/select/option[4]'
    browser_lib.click_element(element)
    time.sleep(2)
    browser_lib.click_element(element1)
    time.sleep(30)


def click_down():
    info = 'xpath://*[@id="investments-table-object"]/tbody'
    st=0
    result = [i.text for i in browser_lib.find_elements(info)]
    myString = ''.join(result).split('\n')

    old_book= open_workbook('test_sheet1.xls')
    book1=copy(old_book)

    for h in myString:
        book1.get_sheet("Info").write(st, 0, h)
        st += 1
        book1.save("test_sheet2.xls")



def safe_links():
    element = 'class:left.sorting_2'
    elements= browser_lib.find_elements(element)

    for i in elements:
        browser_lib.click_link(i)
        time.sleep(2)

def main():
    try:
        open_the_website("https://itdashboard.gov/")
        click_button()
        write_file()
        search_depart('Department of Defense')
        selector_click()
        click_down()
        safe_links()

    finally:
        browser_lib.close_all_browsers()





if __name__ == "__main__":
    main()




