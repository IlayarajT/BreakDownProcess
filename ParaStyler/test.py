import time

from bs4 import BeautifulSoup
from lxml import etree
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager


class GetArticleId:
    def __init__(self):
        self.chrome_driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()))

    def process_table(self, table_input):
        table_soup = BeautifulSoup(table_input, features='lxml')
        table_content = str(table_soup)
        with open("c:/test/table.xml", "a") as file:
            file.write(table_content)
        return table_soup

    def login_smart(self):
        smart_login_page = "https://journals.sageapps.com/smart/login.aspx"
        driver = self.chrome_driver
        try:
            driver.get(smart_login_page)
            driver.implicitly_wait(3)
            username = driver.find_element(By.ID, "ctl00_SmartMasterContent_rtbuserlogin")
            password = driver.find_element(By.ID, "ctl00_SmartMasterContent_rtbpasswd")
            login = driver.find_element(By.XPATH,
                                        '//form//span[@class="RadButton RadButton_Web20_SegoeMod rbSkinnedButton"]')
            username.send_keys("pecandm")
            password.send_keys("654user@c&M")
            login.click()
            time.sleep(3)
            driver.get("https://journals.sageapps.com/smart/MaintainArticle.aspx?articleid=1126051")
            authors_button = driver.find_element(By.XPATH, "//span[contains(text(),'Authors')]")
            authors_button.click()
            time.sleep(3)
            authors_table = driver.find_element(By.XPATH,
                                                "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody")
            authors_text = authors_table.get_attribute('innerHTML')
            authors_soup = BeautifulSoup(authors_text, features='lxml')
            for span in authors_soup.find_all("table", {'summary': 'combobox'}):
                span.decompose()
            authors_dom = etree.HTML(str(authors_soup))
            for au_id in authors_dom.xpath('/html/body/tr/td/a/img/@id'):
                author_window = driver.find_element(By.XPATH, f"//img[@id='{au_id}']/parent::a")
                author_window.click()
                time.sleep(3)
                active_window = driver.switch_to.active_element
                close_button = active_window.find_element(By.XPATH, "//a[@class='rwCloseButton']")
                driver.switch_to.default_content()
                driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, "iframe[name='rwinauthor']"));
                author_table = driver.find_element(By.XPATH, "//legend[normalize-space()='Author Name']/parent::fieldset/table/tbody")
                aff_table = driver.find_element(By.XPATH, "//legend[normalize-space()='Author Address']/parent::fieldset/table/tbody")
                info_table = driver.find_element(By.XPATH, "//legend[normalize-space()='Author Info']/parent::fieldset/table/tbody")
                author_table = self.process_table(author_table.get_attribute('innerHTML'))
                aff_table = self.process_table(aff_table.get_attribute('innerHTML'))
                info_table = self.process_table(info_table.get_attribute('innerHTML'))
                # print(author_table)
                # print(aff_table)
                # print(info_table)
                # print("\n\n")
                driver.switch_to.default_content()
                close_button.click()
            driver.quit()
        except Exception as e:
            driver.quit()
            print(e)


getArtId = GetArticleId()
getArtId.login_smart()