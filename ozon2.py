import time
import random

import pandas as pd
from tkinter import *
from openpyxl import load_workbook
from fake_useragent import UserAgent
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, InvalidArgumentException, NoSuchElementException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium_stealth import stealth


def main(path_to_table, last_string, begin_string):
    ua = UserAgent()
    options = webdriver.ChromeOptions()
    options.add_argument("start-maximized")
    options.add_argument("--headless")
    options.add_argument(f"user-agent={ua.random}")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    # укажите путь до chromedriver
    driver = webdriver.Chrome(options=options, executable_path=r"ChromeDriver/chromedriver")

    stealth(driver,
            languages=["en-US", "en"],
            vendor="Google Inc.",
            platform="Win32",
            webgl_vendor="Intel Inc.",
            renderer="Intel Iris OpenGL Engine",
            fix_hairline=True,
            )

    df = pd.read_excel(path_to_table, index_col=False)
    excel_data = [list(df.columns.values)]
    excel_data += df.values.tolist()

    print(excel_data)
    for i in range(begin_string - 1, last_string):
        link = excel_data[i][3]
        driver.get(link)
        time.sleep(1)
        driver.implicitly_wait(20)
        try:
            price = driver.find_element(By.CSS_SELECTOR, ".m9q").text.replace(" ", "").replace("₽", "")
            excel_data[i][5] = price
            print(price)
        except NoSuchElementException:
            print(link)
    res = pd.DataFrame(excel_data)
    res.to_excel(path_to_table, index=False, index_label=False, header=False)
    print(excel_data)
    driver.quit()

# main("1.xlsx", 10, 2)
top = Tk()
top.title("Парсинг ozon.ru")
top.geometry("400x300")
path_to_table_frame = Frame(top)
path_to_table_frame.place(relx=0.15, relwidth=0.7, relheight=0.2)
path_to_table_description = Label(path_to_table_frame, text="Путь до таблицы")
path_to_table_description.pack(side=LEFT)
path_to_table_input = Entry(path_to_table_frame)
path_to_table_input.pack(side=RIGHT)

table_links_index_frame = Frame(top)
table_links_index_frame.place(relx=0.15, rely=0.15, relwidth=0.7, relheight=0.1)
links_column_description = Label(table_links_index_frame, text="Начальная строка")
links_column_description.pack(side=LEFT)
links_column_input = Entry(table_links_index_frame)
links_column_input.pack(side=RIGHT)

table_price_index_frame = Frame(top)
table_price_index_frame.place(relx=0.15, rely=0.25, relwidth=0.7, relheight=0.1)
price_column_description = Label(table_price_index_frame, text="Последняя строка")
price_column_description.pack(side=LEFT)
price_column_input = Entry(table_price_index_frame)
price_column_input.pack(side=RIGHT)

start_scraping_button_frame = Frame(top)
start_scraping_button_frame.place(rely=0.8, relx=0.2, relwidth=0.6, relheight=0.1)
start_scraping = Button(start_scraping_button_frame, text="Начать парсинг", command=lambda: main(path_to_table_input.get(), int(price_column_input.get()), int(links_column_input.get())))
start_scraping.place(relx=0.1, rely=0.1, relheight=0.8, relwidth=0.8)
top.mainloop()