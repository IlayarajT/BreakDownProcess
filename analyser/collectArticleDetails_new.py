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

# Configure logging
logging.getLogger("urllib3").setLevel(logging.WARNING)
logging.getLogger('selenium').setLevel(logging.CRITICAL)

class ArticleInfoCollector:
    SMART_LOGIN_URL = "https://journals.sageapps.com/smart/login.aspx"
    TASK_PAGE_URL = "https://journals.sageapps.com/smart/ViewTasks.aspx"
    JOURNAL_LIST_URL = "https://journals.sageapps.com/SMART/JournalList.aspx?atype=J"
    
    def __init__(self):
        self.config_folder, _ = getconfig()
        self._load_configuration()
        self._initialize_driver()
        self._load_credentials()
        
    def _load_configuration(self):
        """Load YAML configuration and journal JSON data"""
        config_path = os.path.join(self.config_folder, 'config\\createArticleInfo.yaml')
        with open(config_path, "r") as f:
            self.config = yaml.safe_load(f)
        
        self.wait_time = self.config['waiting_time']
        self.journal_file = os.path.join(self.config_folder, "SupportingFiles\\sageJournalInfo.json")
        
        with open(self.journal_file, "r") as f:
            self.journal_data = json.load(f)
        
        self.journal_tags = self.config['journal']
        self.article_tags = self.config['article']
        self.author_tags = self.config['author']

    def _initialize_driver(self):
        """Set up and configure the WebDriver"""
        options = Options()
        options.add_argument('--no-sandbox')
        options.add_argument('--disable-dev-shm-usage')
        options.headless = self.config['view_web_page']
        
        service = Service(executable_path="firefox.geckodriver")
        self.driver = webdriver.Firefox(options=options, service=service)
        self.driver.set_window_size(1920, 1080)
        self.driver.maximize_window()

    def _load_credentials(self):
        """Load user credentials from configuration"""
        self.credentials = {
            "UK": {
                "username": self.config['UK']['USERNAME'],
                "password": self.config['UK']['PASSWORD'],
                "button": self.config['UK']['BTN']
            },
            "US": {
                "username": self.config['US']['USERNAME'],
                "password": self.config['US']['PASSWORD'],
                "button": self.config['US']['BTN']
            }
        }

    def collect_article_info(self, article_id, ms_number, journal_id, output_folder):
        """Main entry point for article information collection"""
        if journal_id in self.journal_data:
            journal_loc = self.journal_data[journal_id]["journal_loc"]
            return self._collect_with_known_journal(article_id, ms_number, journal_id, output_folder, journal_loc)
        return self._collect_with_unknown_journal(article_id, ms_number, journal_id, output_folder)

    def _collect_with_known_journal(self, article_id, ms_number, journal_id, output_folder, journal_loc):
        """Collect article info when journal is already known"""
        print("Starting SAGE Smart data collection...")
        article_info = self._initialize_article_info_dict()
        journal_tla = self.journal_data[journal_id].get("journal-tla", journal_id)
        
        try:
            # Attempt collection in primary location
            success, found_id = self._attempt_collection(
                journal_loc, article_id, ms_number, article_info, output_folder, True
            )
            if success:
                return True, found_id, journal_tla
            
            # Try alternative location if primary failed
            alt_loc = "UK" if journal_loc == "US" else "US"
            success, found_id = self._attempt_collection(
                alt_loc, article_id, ms_number, article_info, output_folder, True
            )
            return success, found_id, journal_tla if success else None
            
        except Exception as e:
            print(f"Collection failed: {str(e)}")
            return False, article_id, None
        finally:
            self.driver.quit()
            print("Data collection completed")

    def _collect_with_unknown_journal(self, article_id, ms_number, journal_id, output_folder):
        """Collect article info when journal is unknown"""
        print("Starting SAGE Smart data collection...")
        article_info = self._initialize_article_info_dict()
        
        try:
            # Try UK location first
            success, found_id = self._attempt_collection(
                "UK", article_id, ms_number, article_info, output_folder, False
            )
            if success:
                journal_tla = article_info['journal_info'].get('journal-tla', journal_id)
                return True, found_id, journal_tla
            
            # Try US location if UK failed
            success, found_id = self._attempt_collection(
                "US", article_id, ms_number, article_info, output_folder, False
            )
            if success:
                journal_tla = article_info['journal_info'].get('journal-tla', journal_id)
                return True, found_id, journal_tla
            
            return False, article_id, None
            
        except Exception as e:
            print(f"Collection failed: {str(e)}")
            return False, article_id, None
        finally:
            self.driver.quit()
            print("Data collection completed")

    def _initialize_article_info_dict(self):
        """Create the article information dictionary structure"""
        return {
            'article_info': {},
            'journal_info': {},
            'authors_info': {},
            'funder_info': {}
        }

    def _attempt_collection(self, location, article_id, ms_number, article_info, output_folder, journal_known):
        """Perform a single collection attempt in a specific location"""
        self.driver.get(self.SMART_LOGIN_URL)
        self._login(
            self.credentials[location]["username"],
            self.credentials[location]["password"]
        )
        
        # Verify successful login
        if self.driver.current_url != self.TASK_PAGE_URL:
            time.sleep(10)
            self._login(
                self.credentials[location]["username"],
                self.credentials[location]["password"]
            )
        
        # Select location and search for article
        self.driver.find_element(By.XPATH, self.credentials[location]["button"]).click()
        time.sleep(3)
        self._search_article(article_id or ms_number)
        
        # Check if article was found
        if self._is_no_records_found():
            return False, None
        
        # Collect article details
        found_id = self._get_article_id_from_table()
        article_info['article_info']['article_id'] = found_id
        
        if journal_known:
            journal_abbr = self._get_journal_abbr_from_table()
            article_info['journal_info'] = self.journal_data.get(journal_abbr, {})
        else:
            article_info['journal_info']['journal_abbr'] = self._get_journal_abbr_from_table()
            article_info['journal_info']['journal_loc'] = location
        
        self._collect_full_article_info(article_info, location, found_id, output_folder, journal_known)
        return True, found_id

    def _login(self, username, password):
        """Perform login action"""
        self.driver.find_element(By.ID, "ctl00_SmartMasterContent_rtbuserlogin").send_keys(username)
        self.driver.find_element(By.ID, "ctl00_SmartMasterContent_rtbpasswd").send_keys(password)
        self.driver.find_element(
            By.XPATH, '//form//span[@class="RadButton RadButton_Web20_SegoeMod rbSkinnedButton"]'
        ).click()
        time.sleep(4)
        self.driver.implicitly_wait(self.wait_time)

    def _search_article(self, search_term):
        """Search for an article in the system"""
        search_field = self.driver.find_element(By.ID, "TopMenu1_ArticleSearchEdit")
        search_field.clear()
        search_field.send_keys(search_term)
        self.driver.find_element(By.XPATH, '//form//span/input[@id="ctl00_TopMenu1_rbsearch_input"]').click()
        time.sleep(3)
        self.driver.implicitly_wait(self.wait_time)

    def _is_no_records_found(self):
        """Check if no records were found in search results"""
        table = self.driver.find_element(
            By.XPATH, "//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']"
        )
        return "No matching records were found" in table.text

    def _get_article_id_from_table(self):
        """Extract article ID from results table"""
        return self.driver.find_element(
            By.XPATH, '//table[@id="ctl00_SmartMasterContent_ArticleGrid_ctl00"]/tbody/tr/td[3]'
        ).text

    def _get_journal_abbr_from_table(self):
        """Extract journal abbreviation from results table"""
        return self.driver.find_element(
            By.XPATH, '//table[@id="ctl00_SmartMasterContent_ArticleGrid_ctl00"]/tbody/tr/td[1]'
        ).text

    def _collect_full_article_info(self, article_info, location, article_id, output_folder, journal_known):
        """Collect all article details after initial search"""
        print("Collecting article details...")
        journal_abbr = self._get_journal_abbr_from_table()
        
        # Open article editor
        self.driver.find_element(
            By.XPATH, "//img[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00_ctl04_EditButtonImage']"
        ).click()
        time.sleep(10)
        self.driver.implicitly_wait(self.wait_time)
        
        # Navigate to article info
        self.driver.find_element(By.XPATH, "//span[contains(text(),'Article Info')]").click()
        time.sleep(3)
        
        try:
            # Collect article info
            article_info = self._collect_section_info(self.article_tags, article_info, "article_info")
            
            # Collect author info
            self._collect_author_info(article_info)
            
            # Collect funder info
            self._collect_funder_info(article_info)
            
            # Collect journal info if needed
            if not journal_known:
                self.driver.get(self.JOURNAL_LIST_URL)
                time.sleep(3)
                self._select_journal_location(location)
                self.driver.find_element(By.XPATH, f"//a[contains(text(), '{journal_abbr}')]").click()
                time.sleep(3)
                article_info = self._collect_section_info(self.journal_tags, article_info, "journal_info")
                self._update_journal_data(journal_abbr, article_info['journal_info'])
            
            # Save collected data
            self._save_article_info(article_info, output_folder)
            
        except Exception as e:
            print(f"Error during collection: {str(e)}")
            raise

    def _collect_section_info(self, tags, article_info, section, author_id=None):
        """Generic method to collect information for a section"""
        target_dict = article_info[section]
        if author_id is not None:
            if author_id not in target_dict:
                target_dict[author_id] = {}
            target_dict = target_dict[author_id]

        for field, config in tags.items():
            element = self.driver.find_element(By.XPATH, config['tag'])
            
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

    def _collect_author_info(self, article_info):
        """Collect information for all authors"""
        print("Collecting author information...")
        self.driver.find_element(By.XPATH, "//span[contains(text(),'Authors')]").click()
        time.sleep(3)
        
        # Process author table
        authors_table = self.driver.find_element(
            By.XPATH, "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody"
        )
        soup = BeautifulSoup(authors_table.get_attribute('innerHTML'), 'lxml')
        
        # Clean up HTML
        for element in soup.find_all("table", {'summary': 'combobox'}):
            element.decompose()
            
        dom = etree.HTML(str(soup))
        author_count = 0
        
        # Process each author
        for author_id in dom.xpath('/html/body/tr/td/a/img/@id'):
            author_count += 1
            self.driver.find_element(By.XPATH, f"//img[@id='{author_id}']/parent::a").click()
            time.sleep(3)
            
            # Switch to author popup
            self.driver.switch_to.frame(self.driver.find_element(By.CSS_SELECTOR, "iframe[name='rwinauthor']"))
            article_info = self._collect_section_info(self.author_tags, article_info, "authors_info", author_count)
            
            # Close author popup
            self.driver.switch_to.default_content()
            self.driver.find_element(By.XPATH, "//a[@class='rwCloseButton']").click()
            time.sleep(3)
            
        return article_info

    def _collect_funder_info(self, article_info):
        """Collect funder information"""
        print("Collecting funder information...")
        self.driver.find_element(By.XPATH, "//span[@class='rtsTxt'][contains(text(), 'Open Funder')]").click()
        time.sleep(2)
        
        funder_table = self.driver.find_element(
            By.XPATH, "//table[@id='ctl00_ArticleFundRef_uc_fundRefGrid_ctl00']/tbody"
        )
        table_html = funder_table.get_attribute('innerHTML')
        
        if "No matching records were found" in table_html:
            article_info['funder_info'] = False
            return
        
        soup = BeautifulSoup(table_html, 'lxml')
        dom = etree.HTML(str(soup))
        funder_count = 0
        
        for row in dom.xpath('/html/body/tr'):
            funder_count += 1
            cells = row.xpath('td')
            article_info['funder_info'][funder_count] = {
                'id': cells[0].text,
                'funder-name': cells[1].text,
                'funder-id': cells[2].text,
                'grant-id': cells[3].text
            }

    def _select_journal_location(self, location):
        """Select journal location (US/UK)"""
        if location == "UK":
            self.driver.find_element(By.XPATH, "//input[@id='SmartMasterContent_rbllocale_LocaleRBEdit_1']").click()
        elif location == "US":
            self.driver.find_element(By.XPATH, "//input[@id='SmartMasterContent_rbllocale_LocaleRBEdit_0']").click()

    def _update_journal_data(self, journal_abbr, journal_info):
        """Update journal data in JSON file"""
        self.journal_data[journal_abbr] = journal_info
        with open(self.journal_file, "w") as f:
            json.dump(self.journal_data, f, indent=4)

    def _save_article_info(self, article_info, output_folder):
        """Save collected information to JSON file"""
        filename = f"{article_info['journal_info'].get('journal-tla', 'UNKNOWN')}_{article_info['article_info']['article_id']}.json"
        filepath = os.path.join(output_folder, filename)
        
        print("Saving article information...")
        with open(filepath, "w") as f:
            json.dump(article_info, f, indent=4)