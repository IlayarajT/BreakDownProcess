import json
import os
import re
import time
import logging

import yaml
from bs4 import BeautifulSoup
from lxml import etree

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException

from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager

from loadconfig import getconfig


SMART_LOGIN_URL = "https://journals.sageapps.com/smart/login.aspx"

log = logging.getLogger(__name__)

class GetJournalInfo:
    # ==============================================================
    # INIT
    # ==============================================================

    def __init__(self):
        self.configFolder, self.breakDownConfig = getconfig()

        info_yaml = os.path.join(
            self.configFolder, "config", "createArticleInfo.yaml"
        )
        with open(info_yaml, "r") as stream:
            self.sage_details = yaml.safe_load(stream)

        self.journal_json = os.path.join(
            self.configFolder,
            "SupportingFiles",
            "sageJournalInfo.json",
        )

        self.driver = self._create_driver()

        self.uk_user = self.sage_details["UK"]["USERNAME"]
        self.uk_pass = self.sage_details["UK"]["PASSWORD"]
        self.us_user = self.sage_details["US"]["USERNAME"]
        self.us_pass = self.sage_details["US"]["PASSWORD"]

        self.journal_tags = self.sage_details["journal"]
        self.article_tags = self.sage_details["article"]   # preserved
        self.author_tags = self.sage_details["author"]     # preserved

    # ==============================================================
    # DRIVER CREATION (CHROME → FIREFOX FALLBACK)
    # ==============================================================

    def _create_driver(self):
        view_web_page = self.sage_details.get("view_web_page", True)
        headless = not view_web_page

        # ---------- Try Chrome ----------
        try:
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_experimental_option(
                "excludeSwitches", ["enable-logging"]
            )
            if headless:
                chrome_options.add_argument("--headless=new")
                chrome_options.add_argument("--disable-gpu")

            log.info("Launching Chrome WebDriver")
            return webdriver.Chrome(
                service=ChromeService(ChromeDriverManager().install()),
                options=chrome_options,
            )

        except Exception as chrome_error:
            log.warning("Chrome failed, falling back to Firefox")
            log.warning(chrome_error)

        # ---------- Fallback Firefox ----------
        firefox_options = webdriver.FirefoxOptions()
        if headless:
            firefox_options.add_argument("-headless")

        log.info("Launching Firefox WebDriver")
        return webdriver.Firefox(
            service=FirefoxService(GeckoDriverManager().install()),
            options=firefox_options,
        )

    # ==============================================================
    # SAFE WAIT (DOM READY + IFRAME)
    # ==============================================================

    def wait_for(self, by, value, timeout=20):
        wait = WebDriverWait(self.driver, timeout)

        # Wait for DOM ready
        wait.until(
            lambda d: d.execute_script("return document.readyState") == "complete"
        )

        # Try default context
        try:
            return wait.until(EC.presence_of_element_located((by, value)))
        except TimeoutException:
            pass

        # Try inside iframes
        for iframe in self.driver.find_elements(By.TAG_NAME, "iframe"):
            try:
                self.driver.switch_to.frame(iframe)
                elem = wait.until(EC.presence_of_element_located((by, value)))
                return elem
            except TimeoutException:
                self.driver.switch_to.default_content()

        raise TimeoutException(f"Element not found: {value}")

    # ==============================================================
    # LOGIN
    # ==============================================================

    def login_page(self, user_name, pass_word):
        self.wait_for(
            By.ID, "ctl00_SmartMasterContent_rtbuserlogin"
        ).send_keys(user_name)

        self.wait_for(
            By.ID, "ctl00_SmartMasterContent_rtbpasswd"
        ).send_keys(pass_word)

        self.wait_for(
            By.XPATH,
            "//form//span[contains(@class,'RadButton')]",
        ).click()

        # Diagnostic logging (safe)
        log.info("URL after login: %s", self.driver.current_url)
        log.info("Title after login: %s", self.driver.title)

        content = self.wait_for(
            By.XPATH,
            "//span[contains(normalize-space(.), 'Content')]",
        )

        submenu = self.wait_for(
            By.XPATH,
            "//span[contains(normalize-space(.), 'Maintain Journal/Issue/Article')]",
        )

        actions = ActionChains(self.driver)
        actions.move_to_element(content).move_to_element(submenu).perform()
        submenu.click()

    # ==============================================================
    # SCRAPE TAGS
    # ==============================================================

    def create_dic(self, tags, journal_info, jid):
        for tag, tag_info in tags.items():
            xpath = tag_info["tag"]
            tag_type = tag_info["type"]

            element = self.wait_for(By.XPATH, xpath)

            if tag_type == "text":
                journal_info[jid][tag] = element.text
            elif tag_type == "attrib":
                journal_info[jid][tag] = element.get_attribute(
                    tag_info["attrib"]
                )
            elif tag_type == "checked":
                journal_info[jid][tag] = element.is_selected()

        return journal_info

    # ==============================================================
    # MAIN COLLECTION
    # ==============================================================

    def create_journal_json(self):
        journal_info = {}

        try:
            for loc in ("IN", "UK", "US"):
                self.driver.get(SMART_LOGIN_URL)

                if loc == "US":
                    self.login_page(self.us_user, self.us_pass)
                    locale_index = 0
                elif loc == "UK":
                    self.login_page(self.uk_user, self.uk_pass)
                    locale_index = 1
                else:
                    self.login_page(self.uk_user, self.uk_pass)
                    locale_index = 2

                self.wait_for(
                    By.XPATH,
                    f"//input[@id='SmartMasterContent_rbllocale_LocaleRBEdit_{locale_index}']",
                ).click()

                time.sleep(5)
                journal_info = self.get_jrn_info(journal_info, loc)

            with open(self.journal_json, "w") as outfile:
                json.dump(journal_info, outfile, indent=4)

            log.info("Journal info collection completed")

        finally:
            self.driver.quit()

    # ==============================================================
    # JOURNAL LIST EXTRACTION
    # ==============================================================

    def get_jrn_info(self, journal_info, loc):
        table = self.wait_for(
            By.XPATH,
            "//table[@id='SmartMasterContent_dlJournalList']/tbody",
        )
        html = table.get_attribute("innerHTML")

        soup = BeautifulSoup(html, "lxml")
        for t in soup.find_all("table", {"summary": "combobox"}):
            t.decompose()

        dom = etree.HTML(str(soup))

        for row in dom.xpath("/html/body/tr"):
            for link in row.xpath("./td/a"):
                journal_id = link.text
                journal_info[journal_id] = {"journal_loc": loc}

                btn_id = link.attrib["id"]
                self.wait_for(By.ID, btn_id).click()
                time.sleep(1)

                journal_info = self.create_dic(
                    self.journal_tags, journal_info, journal_id
                )

                self.driver.back()
                time.sleep(1)

        return journal_info


# jrnInfo = GetJournalInfo()
# jrnInfo.create_journal_json()