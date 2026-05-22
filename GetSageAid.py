import os
import re
import time

import yaml
from selenium import webdriver
from selenium.common.exceptions import TimeoutException, NoAlertPresentException, UnexpectedAlertPresentException, \
    NoSuchElementException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait

from loadconfig import getconfig


class GetArticleId:
    def __init__(self):
        self.configFolder, self.breakDownConfig = getconfig()
        chrome_driver_path = os.path.join(self.configFolder, 'SupportingFiles', 'chromedriver.exe')
        self.driver = webdriver.Chrome(chrome_driver_path)
        self.load_credentials()

    def load_credentials(self):
        sage_yaml = os.path.join(self.configFolder, 'config', 'GetSageAid.yaml')
        with open(sage_yaml, "r") as stream:
            self.sage_credentials = yaml.safe_load(stream)
        self.uk_user = self.sage_credentials['UK']['USERNAME']
        self.uk_pass = self.sage_credentials['UK']['PASSWORD']
        self.us_user = self.sage_credentials['US']['USERNAME']
        self.us_pass = self.sage_credentials['US']['PASSWORD']

    def handle_alert(self):
        try:
            alert = self.driver.switch_to.alert
            alert.dismiss()  # or alert.accept() depending on your needs
            print("Unexpected alert dismissed.")
        except NoAlertPresentException:
            print("No alert present.")

    def login_page(self, user_name, pass_word, retries=3):
        driver = self.driver
        while retries > 0:
            try:
                self.handle_alert()  # Check and handle any existing alerts
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.ID, "ctl00_SmartMasterContent_rtbuserlogin")))
                username = driver.find_element(By.ID, "ctl00_SmartMasterContent_rtbuserlogin")
                password = driver.find_element(By.ID, "ctl00_SmartMasterContent_rtbpasswd")
                login_button = driver.find_element(By.XPATH,
                                                   '//form//span[@class="RadButton RadButton_Web20_SegoeMod rbSkinnedButton"]')
                username.send_keys(user_name)
                password.send_keys(pass_word)
                login_button.click()
                time.sleep(6)
                self.handle_alert()  # Check and handle any alerts after login
                return True  # If login succeeds, exit loop
            except (NoSuchElementException, TimeoutException, UnexpectedAlertPresentException) as e:
                print(f"Error during login: {e}")
                retries -= 1
                time.sleep(5)  # Wait before retrying
        return None

    def search_article(self, ms_no):
        driver = self.driver
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "TopMenu1_ArticleSearchEdit")))
            article_search = driver.find_element(By.ID, "TopMenu1_ArticleSearchEdit")
            article_search.send_keys(ms_no)
            search_button = driver.find_element(By.XPATH, '//form//span/input[@id="ctl00_TopMenu1_rbsearch_input"]')
            search_button.click()
            time.sleep(3)
            WebDriverWait(driver, 10).until(EC.presence_of_element_located(
                (By.XPATH, "//form//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']")))
            article_table = driver.find_element(By.XPATH,
                                                "//form//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']")
            return article_table.text
        except (NoSuchElementException, TimeoutException, UnexpectedAlertPresentException) as e:
            print(f"Error during article search: {e}")
            self.handle_alert()  # Handle unexpected alert
            return None

    def login_smart(self, ms_no):
        smart_login_page = "https://journals.sageapps.com/smart/login.aspx"
        driver = self.driver

        try:
            # Attempt login with UK credentials
            driver.get(smart_login_page)
            driver.implicitly_wait(3)
            if not self.login_page(self.uk_user, self.uk_pass):  # Correct call with arguments
                print("Failed to login with UK credentials.")
                return None

            # Handle locale selection
            uk_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, '//form//span/input[@id="TopMenu1_LocaleEdit_LocaleRBEdit_1"]')))
            uk_button.click()

            # Search for article
            table_text = self.search_article(ms_no)

            if re.search("No matching records were found", table_text, re.IGNORECASE):
                # Attempt login with US credentials
                driver.get(smart_login_page)
                if not self.login_page(self.us_user, self.us_pass):  # Correct call with arguments
                    print("Failed to login with US credentials.")
                    return None

                # Handle locale selection
                us_button = WebDriverWait(driver, 10).until(EC.element_to_be_clickable(
                    (By.XPATH, '//form//span/input[@id="TopMenu1_LocaleEdit_LocaleRBEdit_0"]')))
                us_button.click()

                # Search for article
                table_text = self.search_article(ms_no)

            # Extract Article ID
            if not re.search("No matching records were found", table_text, re.IGNORECASE):
                article_table = driver.find_element(By.XPATH,
                                                    '//table[@id="ctl00_SmartMasterContent_ArticleGrid_ctl00"]/tbody/tr/td[3]')
                article_id = article_table.text
            else:
                article_id = None

            return article_id

        except Exception as e:
            print(f"An error occurred: {e}")
            self.handle_alert()  # Handle unexpected alert if it causes failure
            return None

        finally:
            driver.quit()
