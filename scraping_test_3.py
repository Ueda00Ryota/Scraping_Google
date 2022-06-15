import time
from requests import options
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import chromedriver_binary
from bs4 import BeautifulSoup
import pandas as pd
import numpy as np
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.support.ui import WebDriverWait
import PySimpleGUI as sg


class Chrome_search:
    def __init__(self):
        self.url = "https://www.google.co.jp/search"
        self.options = Options()
        # self.options.add_argument('--headless')
        self.options.add_argument("--disable-dev-shm-usage")
        sg.theme("Dark Blue 3")
        self.layout = [
            [sg.Text("検索ワード: "), sg.InputText(key="-NAME-")],
            [sg.Text("取得件数: "), sg.InputText(key="-COUNT-")],
            [sg.Button("実行", key="-SUBMIT-")],
        ]
        self.window = sg.Window("検索ワードを入力", self.layout, size=(300, 150))
        while True:
            event, values = self.window.read()
            if event == "-SUBMIT-":
                self.search_word = values["-NAME-"]
                self.search_num = int(values["-COUNT-"])
                break
            if event == sg.WIN_CLOSED:
                break

    def search(self):
        driver = webdriver.Chrome(ChromeDriverManager().install(), options=self.options)
        driver.get(self.url)
        search = driver.find_element_by_name("q")
        search.send_keys(self.search_word)
        search.submit()
        time.sleep(1)
        title_list = []
        url_list = []
        description_list = []
        html = driver.page_source.encode("utf-8")
        soup = BeautifulSoup(html, "html.parser")
        link_elem01 = soup.select(".yuRUbf > a")
        if self.search_num <= len(link_elem01):
            for i in range(self.search_num):
                url_text = link_elem01[i].get("href").replace("/url?q=", "")
                url_list.append(url_text)
        elif self.search_num > len(link_elem01):
            for i in range(len(link_elem01)):
                url_text = link_elem01[i].get("href").replace("/url?q=", "")
                url_list.append(url_text)

        time.sleep(1)
        for i in range(len(url_list)):
            driver.get(url_list[i])
            html2 = driver.page_source.encode("utf-8")
            soup2 = BeautifulSoup(html2, "html.parser")
            title_list.append(driver.title)
            try:
                description = driver.find_element_by_xpath(
                    ("//meta[@name='description']")
                ).get_attribute("content")
                description_list.append(description)
            except:
                description_list.append("")
            driver.back()
            time.sleep(0.3)

        search_ranking = np.arange(1, len(url_list) + 1)

        my_list = {
            "url": url_list,
            "ranking": search_ranking,
            "title": title_list,
            "description": description_list,
        }
        my_file = pd.DataFrame(my_list)
        driver.quit()
        my_file.to_excel(
            self.search_word + ".xlsx", self.search_word, startcol=0, startrow=1
        )
        df = pd.read_excel(self.search_word + ".xlsx")
        return df


if __name__ == "__main__":
    se = Chrome_search()
    df = se.search()
    df.head()
