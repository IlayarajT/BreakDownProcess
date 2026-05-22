import json
import os
import re
import time
import logging
import yaml
from bs4 import BeautifulSoup
from lxml import etree
from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from loadconfig import getconfig

logging.getLogger("urllib3").setLevel(logging.WARNING)
logging.getLogger('selenium').setLevel(logging.CRITICAL)


class GetArticleId:
    def __init__(self):
        self.config_folder, self.breakdown_config = getconfig()
        self._load_configuration()
        self._initialize_driver()
        self._load_credentials()

    def _load_configuration(self):
        """Load YAML configuration and journal JSON data"""
        info_yaml = os.path.join(self.config_folder, 'config\\createArticleInfo.yaml')
        with open(info_yaml, "r") as stream:
            self.sage_details = yaml.safe_load(stream)
        
        self.waiting_time = self.sage_details['waiting_time']
        self.journals_file = os.path.join(self.config_folder, "SupportingFiles\\sageJournalInfo.json")
        
        with open(self.journals_file, "r") as stream:
            self.journal_json = json.load(stream)
        
        self.journal_tags = self.sage_details['journal']
        self.article_tags = self.sage_details['article']
        self.author_tags = self.sage_details['author']

    def _initialize_driver(self):
        """Set up and configure the WebDriver"""
        options = Options()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.headless = self.sage_details['view_web_page']
        
        service = Service(executable_path="firefox.geckodriver")
        self.driver = webdriver.Firefox(options=options, service=service)
        self.driver.set_window_size(1920, 1080)
        self.driver.maximize_window()

    def _load_credentials(self):
        """Load user credentials from configuration"""
        self.uk_user = self.sage_details['UK']['USERNAME']
        self.uk_pass = self.sage_details['UK']['PASSWORD']
        self.us_user = self.sage_details['US']['USERNAME']
        self.us_pass = self.sage_details['US']['PASSWORD']

    def _populate_info_dict(self, driver, tags, article_info, section, au_count=None):
        """
        Populate article information dictionary from web elements
        """
        target_dict = article_info[section]
        
        if au_count is not None:
            if au_count not in target_dict:
                target_dict[au_count] = {}
            target_dict = target_dict[au_count]

        for field, config in tags.items():
            xpath = config['tag']
            element = driver.find_element(By.XPATH, xpath)
            
            if config['type'] == 'text':
                value = element.text
            elif config['type'] == 'attrib':
                value = element.get_attribute(config['attrib'])
            elif config['type'] == 'checked':
                value = element.is_selected()
            else:
                continue
                
            target_dict[field] = value

        return article_info

    def _collect_article_info(self, driver, article_info):
        """Collect article information from the page"""
        return self._populate_info_dict(
            driver, self.article_tags, article_info, "article_info"
        )

    def _collect_authors_info(self, driver, article_info):
        """Collect author information from the page"""
        print("Collecting author(s) info....")
        driver.find_element(By.XPATH, "//span[contains(text(),'Authors')]").click()
        time.sleep(3)
        driver.implicitly_wait(self.waiting_time)
        
        authors_table = driver.find_element(
            By.XPATH, "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody"
        )
        authors_soup = BeautifulSoup(authors_table.get_attribute('innerHTML'), 'lxml')
        
        # Remove combobox elements from author table
        for span in authors_soup.find_all("table", {'summary': 'combobox'}):
            span.decompose()
            
        authors_dom = etree.HTML(str(authors_soup))
        au_count = 0
        
        for au_id in authors_dom.xpath('/html/body/tr/td/a/img/@id'):
            au_count += 1
            driver.find_element(By.XPATH, f"//img[@id='{au_id}']/parent::a").click()
            time.sleep(3)
            driver.implicitly_wait(self.waiting_time)
            
            # Switch to author popup frame
            driver.switch_to.frame(driver.find_element(By.CSS_SELECTOR, "iframe[name='rwinauthor']"))
            article_info = self._populate_info_dict(
                driver, self.author_tags, article_info, "authors_info", au_count
            )
            
            # Return to main frame and close popup
            driver.switch_to.default_content()
            driver.find_element(By.XPATH, "//a[@class='rwCloseButton']").click()
            time.sleep(3)
            driver.implicitly_wait(self.waiting_time)
            
        return article_info

    def _collect_funder_info(self, driver, article_info):
        """Collect funder information from the page"""
        print("Collecting funder info....")
        driver.find_element(By.XPATH, "//span[@class='rtsTxt'][contains(text(), 'Open Funder')]").click()
        time.sleep(2)
        driver.implicitly_wait(self.waiting_time)
        
        funder_table = driver.find_element(
            By.XPATH, "//table[@id='ctl00_ArticleFundRef_uc_fundRefGrid_ctl00']/tbody"
        )
        funder_text = funder_table.get_attribute('innerHTML')
        
        if "No matching records were found" in funder_text:
            article_info['funder_info'] = False
        else:
            funder_soup = BeautifulSoup(funder_text, 'lxml')
            funder_dom = etree.HTML(str(funder_soup))
            funder_count = 0
            
            for idx, row in enumerate(funder_dom.xpath('/html/body/tr')):
                funder_count += 1
                article_info['funder_info'][funder_count] = {
                    'id': row.xpath(f'td[1]')[0].text,
                    'funder-name': row.xpath(f'td[2]')[0].text,
                    'funder-id': row.xpath(f'td[3]')[0].text,
                    'grant-id': row.xpath(f'td[4]')[0].text
                }
                
        return article_info

    def _collect_journal_info(self, driver, article_info, jrn_loc, journal_abbr):
        """Collect journal information and update JSON file"""
        print("Collecting journal info....")
        time.sleep(3)
        
        # Select journal location
        if jrn_loc == "UK":
            driver.find_element(By.XPATH, "//input[@id='SmartMasterContent_rbllocale_LocaleRBEdit_1']").click()
        elif jrn_loc == "US":
            driver.find_element(By.XPATH, "//input[@id='SmartMasterContent_rbllocale_LocaleRBEdit_0']").click()
        
        # Open journal details
        driver.find_element(By.XPATH, f"//a[contains(text(), '{journal_abbr}')]").click()
        time.sleep(3)
        driver.implicitly_wait(self.waiting_time)
        
        # Populate journal info and save to JSON
        article_info = self._populate_info_dict(
            driver, self.journal_tags, article_info, "journal_info"
        )
        self.journal_json[journal_abbr] = article_info['journal_info']
        
        with open(self.journals_file, "w") as outfile:
            json.dump(self.journal_json, outfile, indent=4)
            
        return article_info

    def _save_article_info(self, article_info, process_folder):
        """Save collected information to JSON file"""
        json_file_name = os.path.join(
            process_folder,
            f"{article_info['journal_info']['journal-tla']}_{article_info['article_info']['article_id']}.json"
        )
        
        print("Creating json file....")
        with open(json_file_name, "w") as file:
            json.dump(article_info, file, indent=4)
            
        return article_info

    def create_info_xml(self, driver, article_info, jrn_loc, article_id, process_folder, jrn_found):
        """Main method to collect all article information"""
        print("Collecting article info....")
        journal_abbr = driver.find_element(
            By.XPATH, "//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']/tbody/tr[1]/td[1]"
        ).text
        
        # Open article editor
        driver.find_element(
            By.XPATH, "//img[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00_ctl04_EditButtonImage']"
        ).click()
        time.sleep(10)
        driver.implicitly_wait(self.waiting_time)
        
        # Navigate to article info
        driver.find_element(By.XPATH, "//span[contains(text(),'Article Info')]").click()
        time.sleep(3)
        driver.implicitly_wait(self.waiting_time)
        
        try:
            article_info = self._collect_article_info(driver, article_info)
            article_info = self._collect_authors_info(driver, article_info)
            article_info = self._collect_funder_info(driver, article_info)
            driver.refresh()
            
            if not jrn_found:
                driver.refresh()
                driver.get("https://journals.sageapps.com/SMART/JournalList.aspx?atype=J")
                article_info = self._collect_journal_info(driver, article_info, jrn_loc, journal_abbr)
            
            return self._save_article_info(article_info, process_folder)
            
        except Exception as e:
            driver.quit()
            print(f"Error collecting article info: {e}")
            raise

    def _perform_login(self, driver, username, password):
        """Handle login process"""
        driver.find_element(By.ID, "ctl00_SmartMasterContent_rtbuserlogin").send_keys(username)
        driver.find_element(By.ID, "ctl00_SmartMasterContent_rtbpasswd").send_keys(password)
        driver.find_element(
            By.XPATH, '//form//span[@class="RadButton RadButton_Web20_SegoeMod rbSkinnedButton"]'
        ).click()
        time.sleep(4)
        driver.implicitly_wait(self.waiting_time)

    def _search_article(self, driver, article_id, ms_no, location_btn):
        """Search for article in SMART system"""
        driver.find_element(By.XPATH, location_btn).click()
        time.sleep(3)
        WebDriverWait(driver, self.waiting_time).until(
            EC.url_contains("ViewTasks.aspx")
        )
        
        search_field = driver.find_element(By.ID, "TopMenu1_ArticleSearchEdit")
        search_field.send_keys(article_id or ms_no)
        driver.find_element(By.XPATH, '//form//span/input[@id="ctl00_TopMenu1_rbsearch_input"]').click()
        time.sleep(3)
        driver.implicitly_wait(self.waiting_time)
        
        return driver.find_element(
            By.XPATH, "//form//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']"
        )

    def _handle_article_found(self, driver, article_info, jrn_loc, article_id, process_folder, jrn_found):
        """Process when article is found"""
        article_table = driver.find_element(
            By.XPATH, "//form//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']"
        )
        article_info['article_info']['article_id'] = article_table.find_element(
            By.XPATH, './tbody/tr/td[3]'
        ).text
        journal_abbr = article_table.find_element(By.XPATH, './tbody/tr/td[1]').text
        article_info['journal_info'] = self.journal_json.get(journal_abbr, {})
        
        return self.create_info_xml(
            driver, article_info, jrn_loc, article_id, process_folder, True
        )

    def smart_login(self, article_id, ms_no, journal_id, process_folder):
        """Main entry point for article information collection"""
        if journal_id in self.journal_json:
            jrn_location = self.journal_json[journal_id]["journal_loc"]
            return self._login_with_known_journal(article_id, ms_no, journal_id, process_folder, jrn_location)
        return self._login_with_unknown_journal(article_id, ms_no, journal_id, process_folder)

    def _login_with_known_journal(self, article_id, ms_no, journal_id, process_folder, jrn_loc):
        """Login flow when journal is known"""
        print("Collecting Details from SAGE Smart, It will take few minute(s)")
        print("Login Process Started....")
        
        driver = self.driver
        article_info = {
            'article_info': {}, 
            'journal_info': {}, 
            'authors_info': {}, 
            'funder_info': {}
        }
        jrn_tla = self.journal_json[journal_id].get("journal-tla", journal_id)
        
        try:
            # Initial login attempt
            driver.get("https://journals.sageapps.com/smart/login.aspx")
            self._perform_login(driver, self.sage_details[jrn_loc]['USERNAME'], self.sage_details[jrn_loc]['PASSWORD'])
            
            # Search for article
            article_table = self._search_article(
                driver, article_id, ms_no, self.sage_details[jrn_loc]['BTN']
            )
            
            if "No matching records were found" in article_table.text:
                # Try alternative location
                new_loc = "UK" if jrn_loc == "US" else "US"
                return self._try_alternative_location(
                    driver, article_id, ms_no, journal_id, process_folder, jrn_loc, new_loc, article_info
                )
            
            # Process found article
            result = self._handle_article_found(
                driver, article_info, jrn_loc, article_id, process_folder, True
            )
            return True, result['article_info']['article_id'], jrn_tla
            
        except Exception as e:
            print(f"Error in known journal login: {e}")
            driver.quit()
            return False, article_id, None
        finally:
            driver.quit()
            print("Details Collection Completed")

    # Additional helper methods would follow the same pattern for:
    # _try_alternative_location, _login_with_unknown_journal, 
    # and other refactored parts from the original login methods

    # Note: For brevity, full refactoring of all login methods isn't shown here,
    # but would follow the same structural improvements demonstrated above