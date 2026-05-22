import json
import os
import re
import time
import logging
import warnings
import yaml
from bs4 import BeautifulSoup
from lxml import etree

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.firefox.service import Service as FirefoxService
from webdriver_manager.chrome import ChromeDriverManager
from webdriver_manager.firefox import GeckoDriverManager

from loadconfig import getconfig

warnings.filterwarnings("ignore", message="Exception ignored in.*Popen")

logging.getLogger("urllib3").setLevel(logging.WARNING)
logging.getLogger("selenium").setLevel(logging.CRITICAL)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

log = logging.getLogger(__name__)


class GetArticleId:

    def __init__(self):
        self.configFolder, self.breakDownConfig = getconfig()

        info_yaml = os.path.join(
            self.configFolder, "config", "createArticleInfo.yaml"
        )
        with open(info_yaml, "r", encoding="utf-8") as fh:
            self.sage_details = yaml.safe_load(fh)

        self.waiting_time = int(self.sage_details.get("waiting_time", 30))

        self.journals_file = os.path.join(
            self.configFolder, "SupportingFiles", "sageJournalInfo.json"
        )
        with open(self.journals_file, "r", encoding="utf-8") as fh:
            self.journal_json = json.load(fh)

        # Load BreakDown.json for journal ID validation
        self.breakdown_json_path = os.path.join(
            self.configFolder, "SupportingFiles", "BreakDown.json"
        )
        self.breakdown_journal_data = {}
        try:
            if os.path.exists(self.breakdown_json_path):
                with open(self.breakdown_json_path, "r", encoding="utf-8") as fh:
                    _bd_data = yaml.safe_load(fh)
                    self.breakdown_journal_data = _bd_data.get("journal_details", {})
        except Exception as e:
            print(f"[WARN] Could not load BreakDown.json: {e}")

        self.driver = self._create_driver(headless=True)
        self.driver.set_window_size(1920, 1080)

        self.uk_user = self.sage_details["UK"]["USERNAME"]
        self.uk_pass = self.sage_details["UK"]["PASSWORD"]
        self.us_user = self.sage_details["US"]["USERNAME"]
        self.us_pass = self.sage_details["US"]["PASSWORD"]

        self.journal_tags = self.sage_details["journal"]
        self.article_tags = self.sage_details["article"]
        self.author_tags = self.sage_details["author"]

    # ==============================================================
    # DRIVER CREATION (CHROME → FIREFOX FALLBACK)
    # ==============================================================
    def _create_driver(self, headless=False):
        # ---------- Try Chrome ----------
        try:
            chrome_options = webdriver.ChromeOptions()
            chrome_options.add_experimental_option(
                "excludeSwitches", ["enable-logging"]
            )
            if headless:
                chrome_options.add_argument("--headless=new")
                chrome_options.add_argument("--disable-gpu")
            chrome_options.add_argument("--no-sandbox")
            chrome_options.add_argument("--disable-dev-shm-usage")

            log.info("Launching Chrome WebDriver")
            return webdriver.Chrome(
                service=ChromeService(ChromeDriverManager().install()),
                options=chrome_options,
            )

        except Exception as chrome_error:
            log.warning("Chrome failed, falling back to Firefox: %s", chrome_error)

        # ---------- Fallback Firefox ----------
        firefox_options = webdriver.FirefoxOptions()
        if headless:
            firefox_options.add_argument("-headless")
        firefox_options.add_argument("--no-sandbox")
        firefox_options.add_argument("--disable-dev-shm-usage")

        log.info("Launching Firefox WebDriver")
        return webdriver.Firefox(
            service=FirefoxService(GeckoDriverManager().install()),
            options=firefox_options,
        )

    # ------------------------------------------------------------
    # GENERIC RETRY
    # ------------------------------------------------------------
    def _retry(self, func, label="", retries=3, delay=2):
        for attempt in range(1, retries + 1):
            try:
                return func()
            except Exception as e:
                logging.warning(f"{label} failed ({attempt}/{retries}): {e}")
                if attempt == retries:
                    raise
                time.sleep(delay)

    # ------------------------------------------------------------
    # RADTABSTRIP UTILITIES (GENERIC, SAFE)
    # ------------------------------------------------------------
    def _get_tab_anchor(self, wait, tab_text):
        return wait.until(
            EC.presence_of_element_located(
                (
                    By.XPATH,
                    f"//a[contains(@class,'rtsLink')][.//span[contains(normalize-space(),'{tab_text}')]]"
                )
            )
        )

    def _is_tab_selected(self, tab):
        return "rtsSelected" in (tab.get_attribute("class") or "")

    def _click_tab(self, wait, tab_text, label, content_xpath=None):
        def _click():
            tab = self._get_tab_anchor(wait, tab_text)
            self.driver.execute_script(
                "arguments[0].scrollIntoView({block:'center'});", tab
            )
            self.driver.execute_script("arguments[0].click();", tab)

        self._retry(_click, label)

        def _wait_dom():
            if self.driver.execute_script("return document.readyState") != "complete":
                raise RuntimeError("DOM not ready")
            return True

        self._retry(_wait_dom, f"{label} DOM ready")

        if content_xpath:
            try:
                wait.until(
                    EC.presence_of_element_located((By.XPATH, content_xpath))
                )
                return True
            except Exception:
                logging.info(f"{label}: no content available")
                return False

        return True

    # ------------------------------------------------------------
    # SAFE DRIVER SHUTDOWN
    # ------------------------------------------------------------
    def _safe_quit_driver(self):
        try:
            self.driver.quit()
        except Exception:
            pass

    # ------------------------------------------------------------
    # LOGIN PAGE
    # ------------------------------------------------------------
    def login_page(self, driver, user_name, pass_word):
        print("[INFO]: LogIn Smart.....")
        wait = WebDriverWait(driver, self.waiting_time)
        driver.get("https://journals.sageapps.com/smart/login.aspx")

        wait.until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_SmartMasterContent_rtbuserlogin")
            )
        ).send_keys(user_name)

        wait.until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_SmartMasterContent_rtbpasswd")
            )
        ).send_keys(pass_word)

        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, "//span[contains(@class,'RadButton')]")
            )
        ).click()

        wait.until(EC.url_contains("ViewTasks.aspx"))

    # ------------------------------------------------------------
    # JID MATCHING HELPER
    # ------------------------------------------------------------
    @staticmethod
    def _jid_match(meta_jid, table_jid):
        """Return True if meta_jid and table_jid are prefix-compatible
        (case-insensitive). Used to find the right row in the SMART search
        results when the metadata journal abbreviation may be slightly
        longer or shorter than the SMART acronym (e.g. JVDI vs JVD).

        Examples (True):  JVDI vs JVDI, JVDI vs JVD, SPX vs SP
        Examples (False): JVDI vs AUT, SPX vs SPP
        """
        a = (meta_jid or "").strip().upper()
        b = (table_jid or "").strip().upper()
        if not a or not b:
            return False
        return a.startswith(b) or b.startswith(a)

    def _find_matching_row(self, table, expected_jid):
        """Iterate the rows of the SMART article search results table and
        return the index (1-based, matching XPath tr[N]) of the first row
        whose Journal Acronym (td[1]) is prefix-compatible with the expected
        journal abbreviation. Returns None if no match found.
        """
        if not expected_jid:
            return None
        try:
            rows = table.find_elements(By.XPATH, "./tbody/tr")
            for idx, row in enumerate(rows, start=1):
                try:
                    cells = row.find_elements(By.XPATH, "./td")
                    if not cells:
                        continue
                    table_jid = cells[0].text.strip()
                    if self._jid_match(expected_jid, table_jid):
                        print(f"[INFO] Matching row {idx}: "
                              f"table='{table_jid}' expected='{expected_jid}'")
                        return idx
                except Exception:
                    continue
        except Exception as exc:
            print(f"[WARN] Could not iterate rows: {exc}")
        return None

    # ------------------------------------------------------------
    def create_info_xml(
        self, driver, article_info, jrn_loc, article_id, process_folder, jrn_found,
        matched_row=1
    ):
        print("[INFO]: Collecting Details Smart.....")
        wait = WebDriverWait(driver, self.waiting_time)
        # Pick journal_abbr from the matched row (not blindly the first row).
        # matched_row is 1-based to match XPath tr[N].
        if matched_row is None or matched_row < 1:
            matched_row = 1
        journal_abbr = self._retry(
            lambda: wait.until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        f"//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']"
                        f"/tbody/tr[{matched_row}]/td[1]"
                    )
                )
            ).text,
            "Get journal abbreviation"
        )

        current_url = self.driver.current_url

        def click_edit():
            # Click the edit button in the matched row, not the first one.
            el = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        f"//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']"
                        f"/tbody/tr[{matched_row}]//img[contains(@id,'EditButtonImage')]"
                    )
                )
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", el)
            self.driver.execute_script("arguments[0].click();", el)

        self._retry(click_edit, "Open Article")

        WebDriverWait(self.driver, 20).until(
            lambda d: d.current_url != current_url
                      or d.find_elements(By.XPATH, "//input[@id='articleTitle']")
        )
        article_page_url = driver.current_url
        self._click_tab(wait, "Article Info", "Article Info tab")
        article_info['article_info']['article_id'] = article_id
        article_info = self.create_dic(driver, self.article_tags, article_info, "article_info")

        # Post-process: handle "Other" article type and extract Journal ID
        article_info = self._post_process_article_info(driver, article_info)

        # Use journal ID from portal if available
        portal_jid = article_info.get("article_info", {}).get("journal-id-from-portal")
        if portal_jid:
            journal_abbr = portal_jid

        driver.get(article_page_url)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        # ---------- Authors
        has_authors = self._click_tab(
            wait,
            "Authors",
            "Authors tab",
            "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody"
        )

        if has_authors:
            max_retries = 3
            retry_count = 0
            authors_processed_successfully = False

            while retry_count < max_retries and not authors_processed_successfully:
                try:
                    # Get initial count of authors
                    rows = wait.until(EC.presence_of_all_elements_located((
                        By.XPATH,
                        "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody/tr[contains(@class,'rgRow') or contains(@class,'rgAltRow')]"
                    )))
                    total_authors = len(rows)

                    # Reset authors_info for this attempt
                    if retry_count > 0:
                        article_info["authors_info"] = {}

                    au_count = 0
                    successful_authors = 0

                    # Process each author by index instead of iterating over stale elements
                    for idx in range(total_authors):
                        au_count += 1
                        try:
                            # Re-fetch rows on each iteration to avoid stale element issues
                            current_rows = wait.until(EC.presence_of_all_elements_located((
                                By.XPATH,
                                "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody/tr[contains(@class,'rgRow') or contains(@class,'rgAltRow')]"
                            )))

                            # Get the current row by index
                            row = current_rows[idx]

                            # Find and click edit button
                            edit_btn = row.find_element(By.XPATH, ".//a/img[contains(@id,'Edit')]")

                            driver.execute_script(
                                "arguments[0].scrollIntoView({block:'center'});", edit_btn
                            )
                            driver.execute_script("arguments[0].click();", edit_btn)

                            # Wait for iframe and switch to it
                            wait.until(EC.frame_to_be_available_and_switch_to_it("rwinauthor"))
                            time.sleep(2)

                            # Extract author information
                            article_info = self.create_dic(
                                driver,
                                self.author_tags,
                                article_info,
                                "authors_info",
                                au_count
                            )

                            successful_authors += 1

                        except Exception as e:
                            article_info.setdefault("authors_info", {})[au_count] = {"error": str(e)}
                            print(f"Error processing author {au_count}: {str(e)}")

                        finally:
                            # Always switch back to default content
                            driver.switch_to.default_content()

                            try:
                                # Close the popup
                                close_btn = wait.until(EC.element_to_be_clickable((
                                    By.XPATH, "//a[@class='rwCloseButton']"
                                )))
                                close_btn.click()

                                # Wait for popup to close
                                wait.until(EC.invisibility_of_element_located((By.ID, "rwinauthor")))

                                # Wait for grid to stabilize after popup close
                                time.sleep(1)

                            except Exception as close_error:
                                print(f"Error closing popup for author {au_count}: {close_error}")

                    # Validate: Check if all authors were processed successfully
                    authors_with_errors = sum(
                        1 for author_data in article_info.get("authors_info", {}).values()
                        if isinstance(author_data, dict) and "error" in author_data
                    )

                    if authors_with_errors == 0 and successful_authors == total_authors:
                        authors_processed_successfully = True
                        print(f"Successfully processed all {total_authors} authors")
                    else:
                        print(
                            f"Attempt {retry_count + 1}: Processed {successful_authors}/{total_authors} authors successfully, {authors_with_errors} with errors")
                        retry_count += 1

                        if retry_count < max_retries:
                            print(f"Retrying... (Attempt {retry_count + 1}/{max_retries})")
                            time.sleep(2)  # Wait before retry

                except Exception as e:
                    print(f"Error during author processing attempt {retry_count + 1}: {str(e)}")
                    retry_count += 1

                    if retry_count < max_retries:
                        print(f"Retrying... (Attempt {retry_count + 1}/{max_retries})")
                        time.sleep(2)  # Wait before retry

            # Final validation report
            if not authors_processed_successfully:
                print(f"Warning: Failed to process all authors after {max_retries} attempts")
                total_processed = len(article_info.get("authors_info", {}))
                errors = sum(
                    1 for author_data in article_info.get("authors_info", {}).values()
                    if isinstance(author_data, dict) and "error" in author_data
                )
                print(f"Final result: {total_processed - errors}/{total_processed} authors processed successfully")

        driver.get(article_page_url)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        # ---------- Funders
        has_funders = self._click_tab(
            wait,
            "Open Funder",
            "Funder tab",
            "//table[@id='ctl00_ArticleFundRef_uc_fundRefGrid_ctl00']/tbody"
        )

        article_info["funder_info"] = {}

        if has_funders:
            funder_table = driver.find_element(
                By.XPATH,
                "//table[@id='ctl00_ArticleFundRef_uc_fundRefGrid_ctl00']/tbody"
            )

            funder_text = funder_table.get_attribute("innerHTML")

            if re.search("No matching records were found", funder_text, re.I):
                article_info["funder_info"] = False
            else:
                funder_soup = BeautifulSoup(funder_text, "lxml")
                funder_dom = etree.HTML(str(funder_soup))

                funder_count = 0
                for row in funder_dom.xpath("/html/body/tr"):
                    funder_count += 1
                    article_info["funder_info"][funder_count] = {
                        "id": row.xpath("./td[1]/text()")[0],
                        "funder-name": row.xpath("./td[2]/text()")[0],
                        "funder-id": row.xpath("./td[3]/text()")[0],
                        "grant-id": row.xpath("./td[4]/text()")[0],
                    }
        else:
            article_info["funder_info"] = False

        # ---------- Journal Info (inside create_info_xml)
        if jrn_found is False:
            wait = WebDriverWait(driver, self.waiting_time)

            self._retry(
                lambda: driver.refresh(),
                "Refresh journal page"
            )

            jrn_page = "https://journals.sageapps.com/SMART/JournalList.aspx?atype=J"

            self._retry(
                lambda: driver.get(jrn_page),
                "Open journal list page"
            )

            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            print("Collecting journal info....")

            # ---------- Locale selection
            if jrn_loc == "UK":
                self._retry(
                    lambda: wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                "//input[@id='SmartMasterContent_rbllocale_LocaleRBEdit_1']"
                            )
                        )
                    ).click(),
                    "Select UK locale"
                )

            elif jrn_loc == "US":
                self._retry(
                    lambda: wait.until(
                        EC.element_to_be_clickable(
                            (
                                By.XPATH,
                                "//input[@id='SmartMasterContent_rbllocale_LocaleRBEdit_0']"
                            )
                        )
                    ).click(),
                    "Select US locale"
                )

            # ---------- Open journal
            # IMPORTANT: Use exact match (text starts with "ABBR -" or "ABBR ")
            # to avoid contains() picking up "SPPS - ..." when looking for "SP".
            self._retry(
                lambda: wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            f"//a[starts-with(normalize-space(text()), '{journal_abbr} -') "
                            f"or starts-with(normalize-space(text()), '{journal_abbr} ') "
                            f"or normalize-space(text())='{journal_abbr}']"
                        )
                    )
                ).click(),
                f"Open journal {journal_abbr}"
            )

            wait.until(lambda d: d.execute_script("return document.readyState") == "complete")
            driver.implicitly_wait(self.waiting_time)

            # ---------- Extract journal info
            article_info = self.create_dic(
                driver,
                self.journal_tags,
                article_info,
                "journal_info",
                None
            )

            jrnl_info = article_info.get("journal_info")

            # Sanity check: scraped journal-accronym should match journal_abbr
            # If not, the wrong journal link was clicked (substring match issue)
            scraped_acc = (jrnl_info.get("journal-accronym") or "").strip()
            scraped_tla = (jrnl_info.get("journal-tla") or "").strip()
            if (scraped_acc and scraped_acc != journal_abbr
                    and scraped_tla and scraped_tla != journal_abbr):
                print(f"[WARN] Journal page mismatch: expected '{journal_abbr}', "
                      f"scraped journal-accronym='{scraped_acc}', "
                      f"journal-tla='{scraped_tla}'. "
                      f"The wrong journal link may have been clicked.")

            self.journal_json[journal_abbr] = jrnl_info

            with open(self.journals_file, "w") as outfile:
                outfile.write(json.dumps(self.journal_json, indent=4))

        # Resolve journal ID: prefer journal-accronym over journal-tla
        # if it exists in BreakDown.json
        article_info = self._resolve_journal_id(article_info)

        json_path = os.path.join(
            process_folder,
            f"{article_info['journal_info']['journal-tla']}_{article_id}.json"
        )

        with open(json_path, "w", encoding="utf-8") as fh:
            json.dump(article_info, fh, indent=4)

        return article_info


    def create_info_xml_new(
            self, driver, article_info, jrn_loc, article_id, process_folder, jrn_found,
            matched_row=1
    ):
        print("[INFO]: Collecting Details Smart.....")
        wait = WebDriverWait(driver, self.waiting_time)

        if matched_row is None or matched_row < 1:
            matched_row = 1

        # Get journal abbreviation from the matched row
        journal_abbr = self._retry(
            lambda: wait.until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        f"//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']"
                        f"/tbody/tr[{matched_row}]/td[1]"
                    )
                )
            ).text,
            "Get journal abbreviation"
        )

        # Click Edit button to open article (in the matched row)
        current_url = self.driver.current_url

        def click_edit():
            el = wait.until(
                EC.element_to_be_clickable(
                    (
                        By.XPATH,
                        f"//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']"
                        f"/tbody/tr[{matched_row}]//img[contains(@id,'EditButtonImage')]"
                    )
                )
            )
            self.driver.execute_script("arguments[0].scrollIntoView(true);", el)
            time.sleep(0.5)
            self.driver.execute_script("arguments[0].click();", el)

        self._retry(click_edit, "Open Article")

        # Wait for page to load
        WebDriverWait(self.driver, 20).until(
            lambda d: d.current_url != current_url
                      or d.find_elements(By.XPATH, "//input[@id='articleTitle']")
        )
        article_page_url = driver.current_url

        # ---------- Article Info Tab
        self._click_tab(wait, "Article Info", "Article Info tab")
        article_info['article_info']['article_id'] = article_id
        article_info = self.create_dic(driver, self.article_tags, article_info, "article_info")

        # Post-process: handle "Other" article type and extract Journal ID
        article_info = self._post_process_article_info(driver, article_info)

        # Use journal ID from portal if available
        portal_jid = article_info.get("article_info", {}).get("journal-id-from-portal")
        if portal_jid:
            journal_abbr = portal_jid

        # Return to article page
        driver.get(article_page_url)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        # ---------- Authors Tab - COMPLETELY REWRITTEN FOR RELIABILITY
        has_authors = self._click_tab(
            wait,
            "Authors",
            "Authors tab",
            "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody"
        )

        if has_authors:
            # Wait a moment for the table to fully load
            time.sleep(2)

            # METHOD 1: Try to get all edit images directly first
            all_edit_images = driver.find_elements(By.XPATH, "//img[contains(@id, 'EditButtonImage')]")
            logging.info(f"Found {len(all_edit_images)} edit images directly")

            if len(all_edit_images) > 0:
                # Process by edit images directly (most reliable)
                au_count = 0
                for img in all_edit_images:
                    try:
                        au_count += 1

                        # Find the parent anchor and click
                        edit_link = img.find_element(By.XPATH, "./parent::a")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", edit_link)
                        time.sleep(0.3)
                        self.driver.execute_script("arguments[0].click();", edit_link)

                        # Wait for popup and switch to frame
                        wait.until(EC.frame_to_be_available_and_switch_to_it("rwinauthor"))

                        # Extract author information
                        article_info = self.create_dic(
                            driver,
                            self.author_tags,
                            article_info,
                            "authors_info",
                            au_count
                        )

                        logging.info(f"Successfully processed author {au_count}")

                    except Exception as e:
                        logging.error(f"Error processing author {au_count}: {str(e)}")
                        article_info.setdefault("authors_info", {})[au_count] = {
                            "error": str(e)
                        }

                    finally:
                        # Close popup
                        self._safe_close_popup(driver)

            else:
                # METHOD 2: If no edit images found, try getting by rows
                logging.info("No direct edit images found, trying row-based approach")

                # Get the authors table
                authors_table = driver.find_element(
                    By.XPATH,
                    "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody"
                )

                # Get all rows
                author_rows = authors_table.find_elements(By.XPATH, "./tr")
                logging.info(f"Found {len(author_rows)} author rows")

                au_count = 0

                for idx, row in enumerate(author_rows, 1):
                    try:
                        # Try multiple XPath strategies to find edit button
                        edit_button = None

                        # Strategy 1: Direct img with EditButtonImage in ID
                        try:
                            edit_button = row.find_element(By.XPATH, ".//img[contains(@id, 'EditButtonImage')]")
                        except:
                            pass

                        # Strategy 2: Look for any img that might be an edit button
                        if not edit_button:
                            try:
                                all_imgs = row.find_elements(By.XPATH, ".//img")
                                for img in all_imgs:
                                    if 'edit' in img.get_attribute('src').lower() or 'pencil' in img.get_attribute(
                                            'src').lower():
                                        edit_button = img
                                        break
                            except:
                                pass

                        # Strategy 3: Look for anchor with edit in href or class
                        if not edit_button:
                            try:
                                edit_links = row.find_elements(By.XPATH,
                                                               ".//a[contains(@href, 'Edit') or contains(@class, 'edit')]")
                                if edit_links:
                                    # Find the img within this link
                                    edit_button = edit_links[0].find_element(By.XPATH, ".//img")
                            except:
                                pass

                        if not edit_button:
                            logging.warning(f"Could not find edit button in row {idx} with any strategy")

                            # Debug: Print row HTML for inspection
                            row_html = row.get_attribute("outerHTML")
                            logging.debug(f"Row {idx} HTML: {row_html[:200]}...")
                            continue

                        au_count += 1

                        # Click the edit button
                        edit_link = edit_button.find_element(By.XPATH, "./parent::a")
                        self.driver.execute_script("arguments[0].scrollIntoView(true);", edit_link)
                        time.sleep(0.3)
                        self.driver.execute_script("arguments[0].click();", edit_link)

                        # Wait for popup
                        wait.until(EC.frame_to_be_available_and_switch_to_it("rwinauthor"))

                        # Extract author information
                        article_info = self.create_dic(
                            driver,
                            self.author_tags,
                            article_info,
                            "authors_info",
                            au_count
                        )

                        logging.info(f"Successfully processed author {au_count}")

                    except Exception as e:
                        logging.error(f"Error processing author {au_count}: {str(e)}")
                        article_info.setdefault("authors_info", {})[au_count] = {
                            "error": str(e)
                        }

                    finally:
                        # Always close popup
                        driver.switch_to.default_content()
                        self._safe_close_popup(driver)

                        # Small delay between authors
                        time.sleep(1)

            # Verify we processed all authors
            logging.info(f"Total authors processed: {au_count}")

        # Return to article page
        driver.get(article_page_url)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        # ---------- Funders Tab (unchanged)
        has_funders = self._click_tab(
            wait,
            "Open Funder",
            "Funder tab",
            "//table[@id='ctl00_ArticleFundRef_uc_fundRefGrid_ctl00']/tbody"
        )

        article_info["funder_info"] = {}

        if has_funders:
            funder_table = driver.find_element(
                By.XPATH,
                "//table[@id='ctl00_ArticleFundRef_uc_fundRefGrid_ctl00']/tbody"
            )

            funder_text = funder_table.get_attribute("innerHTML")

            if re.search("No matching records were found", funder_text, re.I):
                article_info["funder_info"] = False
            else:
                funder_soup = BeautifulSoup(funder_text, "lxml")
                funder_dom = etree.HTML(str(funder_soup))

                funder_count = 0
                for row in funder_dom.xpath("/html/body/tr"):
                    funder_count += 1
                    article_info["funder_info"][funder_count] = {
                        "id": row.xpath("./td[1]/text()")[0] if row.xpath("./td[1]/text()") else "",
                        "funder-name": row.xpath("./td[2]/text()")[0] if row.xpath("./td[2]/text()") else "",
                        "funder-id": row.xpath("./td[3]/text()")[0] if row.xpath("./td[3]/text()") else "",
                        "grant-id": row.xpath("./td[4]/text()")[0] if row.xpath("./td[4]/text()") else "",
                    }
        else:
            article_info["funder_info"] = False

        # ---------- Journal Info (unchanged)
        if jrn_found is False:
            # ... (keep your existing journal info code)
            pass

        # Resolve journal ID: prefer journal-accronym over journal-tla
        # if it exists in BreakDown.json
        article_info = self._resolve_journal_id(article_info)

        # Save JSON file
        json_path = os.path.join(
            process_folder,
            f"{article_info['journal_info']['journal-tla']}_{article_id}.json"
        )

        with open(json_path, "w", encoding="utf-8") as fh:
            json.dump(article_info, fh, indent=4)

        return article_info

    def _safe_close_popup(self, driver):
        """Helper method to safely close popups"""
        for attempt in range(3):
            try:
                # Try multiple close button selectors
                close_selectors = [
                    "//a[@class='rwCloseButton']",
                    "//div[@class='rwCloseButton']",
                    "//span[contains(@class, 'rwCloseButton')]",
                    "//button[contains(@class, 'close')]",
                    "//img[contains(@src, 'close')]"
                ]

                for selector in close_selectors:
                    try:
                        close_buttons = driver.find_elements(By.XPATH, selector)
                        for btn in close_buttons:
                            if btn.is_displayed():
                                self.driver.execute_script("arguments[0].click();", btn)
                                time.sleep(0.5)
                                return True
                    except:
                        continue

                # If no button found, try Escape key
                from selenium.webdriver.common.keys import Keys
                body = driver.find_element(By.TAG_NAME, 'body')
                body.send_keys(Keys.ESCAPE)
                time.sleep(0.5)

            except:
                time.sleep(1)

        return False


    def _verify_author_count(self, driver, expected_count=None):
        """Verify that we've processed all authors"""
        try:
            authors_table = driver.find_element(
                By.XPATH,
                "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody"
            )
            rows = authors_table.find_elements(By.XPATH, "./tr")
            return len(rows)
        except:
            return 0


    # ------------------------------------------------------------
    # ENTRY POINTS (UNCHANGED)
    # ------------------------------------------------------------
    def smart_login(self, article_id, ms_no, journal_id, process_folder):
        try:
            if journal_id in self.journal_json:
                jrn_loc = self.journal_json[journal_id]["journal_loc"]
                return self.login_smart(
                    article_id, ms_no, journal_id, process_folder, jrn_loc
                )
            return self.login_regular(
                article_id, ms_no, journal_id, process_folder
            )
        finally:
            self._safe_quit_driver()

    # ------------------------------------------------------------
    # HELPER: LOGIN + LOCALE SELECT + SEARCH
    # ------------------------------------------------------------
    def _login_and_search(self, loc, search_term, expected_jid=None):
        """Login with the given locale credentials, select locale, search for article.

        Returns a tuple (table, matched_row_index):
          - table: the article results table element, or None if no records
          - matched_row_index: 1-based row index whose Journal Acronym matches
            expected_jid via prefix rule, or None if no match found.

        When expected_jid is None, matched_row_index will always be None
        (caller will fall back to row 1).
        """
        driver = self.driver
        wait = WebDriverWait(driver, self.waiting_time)

        self.login_page(
            driver,
            self.sage_details[loc]["USERNAME"],
            self.sage_details[loc]["PASSWORD"],
        )

        # Select locale
        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, self.sage_details[loc]["BTN"])
            )
        ).click()

        # Search for article
        search = wait.until(
            EC.presence_of_element_located((By.ID, "TopMenu1_ArticleSearchEdit"))
        )
        search.send_keys(search_term)

        wait.until(
            EC.element_to_be_clickable(
                (By.ID, "ctl00_TopMenu1_rbsearch_input")
            )
        ).click()

        table = wait.until(
            EC.presence_of_element_located(
                (By.ID, "ctl00_SmartMasterContent_ArticleGrid_ctl00")
            )
        )

        if "No matching records" in table.text:
            return None, None

        matched_row = self._find_matching_row(table, expected_jid)
        return table, matched_row

    def login_smart(self, article_id, ms_no, journal_id, process_folder, jrn_loc):
        driver = self.driver
        search_term = article_id or ms_no

        article_info = {
            "article_info": {},
            "journal_info": self.journal_json[journal_id],
            "authors_info": {},
            "funder_info": {}
        }

        # Determine locale order: try journal's locale first, then the other
        fallback_loc = "US" if jrn_loc == "UK" else "UK"
        locale_order = [jrn_loc, fallback_loc]

        table = None
        matched_row = None
        active_loc = None

        # First pass: find a locale where we get a JID-matching row
        candidate_results = {}
        for loc in locale_order:
            print(f"[INFO]: Trying {loc} locale for article search "
                  f"(expected JID='{journal_id}')...")
            t, mrow = self._login_and_search(loc, search_term, expected_jid=journal_id)
            if t is None:
                print(f"[INFO]: No matching records found in {loc}.")
                continue
            candidate_results[loc] = (t, mrow)
            if mrow is not None:
                table = t
                matched_row = mrow
                active_loc = loc
                print(f"[INFO]: Locale '{loc}' has matching row {mrow} "
                      f"for JID '{journal_id}'.")
                break

        # Fallback: no JID match in any locale, use first available locale
        if table is None and candidate_results:
            for loc in locale_order:
                if loc in candidate_results:
                    table, _ = candidate_results[loc]
                    matched_row = 1
                    active_loc = loc
                    print(f"[WARN] No JID match in any locale. Falling back to "
                          f"{loc} locale, row 1. Expected JID was '{journal_id}'.")
                    break

        if table is None:
            print("[INFO]: No matching records found in any locale.")
            return False, article_id, None

        if matched_row is None:
            matched_row = 1
        try:
            article_id = table.find_element(
                By.XPATH, f"./tbody/tr[{matched_row}]/td[3]"
            ).text
        except Exception:
            article_id = table.find_element(By.XPATH, "./tbody/tr/td[3]").text

        self.create_info_xml(
            driver,
            article_info,
            active_loc,
            article_id,
            process_folder,
            True,
            matched_row=matched_row
        )

        return True, article_id, article_info["journal_info"]["journal-tla"]

    def create_dic(self, driver, tags, article_info, info, au_count=None):
        wait = WebDriverWait(driver, self.waiting_time)

        # Fields that are conditionally present on the page — skip silently
        # if element is not found (no retries, no warnings).
        optional_fields = {"other-article-type"}

        for tag, meta in tags.items():
            xpath = meta["tag"]
            tag_type = meta["type"]

            # For optional fields, do a quick presence check first.
            # If the element doesn't exist, set None and move on silently.
            if tag in optional_fields:
                try:
                    driver.find_element(By.XPATH, xpath)
                except Exception:
                    if au_count is None:
                        article_info.setdefault(info, {})[tag] = None
                    else:
                        article_info.setdefault(info, {}).setdefault(au_count, {})[tag] = None
                    continue

            def elem():
                return self._retry(
                    lambda: wait.until(
                        EC.presence_of_element_located((By.XPATH, xpath))
                    ),
                    f"{info}:{tag}",
                )

            if au_count is None:
                article_info.setdefault(info, {})
                try:
                    if tag_type == "text":
                        article_info[info][tag] = elem().text
                    elif tag_type == "attrib":
                        article_info[info][tag] = elem().get_attribute(
                            meta["attrib"]
                        )
                    elif tag_type == "checked":
                        article_info[info][tag] = elem().is_selected()
                except Exception:
                    article_info[info][tag] = None
            else:
                article_info.setdefault(info, {}).setdefault(au_count, {})
                try:
                    if tag_type == "text":
                        article_info[info][au_count][tag] = elem().text
                    elif tag_type == "attrib":
                        article_info[info][au_count][tag] = elem().get_attribute(
                            meta["attrib"]
                        )
                    elif tag_type == "checked":
                        article_info[info][au_count][tag] = elem().is_selected()
                except Exception:
                    article_info[info][au_count][tag] = None

        return article_info

    # ==============================================================
    # JOURNAL ID RESOLUTION: abbreviation vs TLA via BreakDown.json
    # ==============================================================
    def _resolve_journal_id(self, article_info):
        """
        Resolve the correct journal ID by checking BreakDown.json.

        Candidates (checked in this order against BreakDown.json):
        1. journal-id-from-portal (extracted from Journal dropdown, e.g. "SP")
        2. journal-accronym (from journal details page, e.g. "SPPS")
        3. journal-tla (from journal details page, e.g. "SPP")

        Logic:
        - First candidate found in BreakDown.json wins.
        - If none found, proceed with journal-tla as-is (fallback).
        """
        jrnl_info = article_info.get("journal_info", {})
        art_info = article_info.get("article_info", {})
        portal_jid = (art_info.get("journal-id-from-portal") or "").strip()
        accronym = (jrnl_info.get("journal-accronym") or "").strip()
        tla = (jrnl_info.get("journal-tla") or "").strip()

        if not self.breakdown_journal_data:
            return article_info

        # Build ordered list of candidates
        candidates = []
        if portal_jid:
            candidates.append(("journal-id-from-portal", portal_jid))
        if accronym:
            candidates.append(("journal-accronym", accronym))
        if tla:
            candidates.append(("journal-tla", tla))

        resolved = None
        resolved_source = None
        for source, value in candidates:
            if value in self.breakdown_journal_data:
                resolved = value
                resolved_source = source
                break

        if resolved:
            print(f"[INFO] Journal ID resolved: using {resolved_source} '{resolved}' "
                  f"(found in BreakDown.json)")
            if resolved != tla:
                print(f"[INFO] Overriding journal-tla: '{tla}' -> '{resolved}'")
                article_info["journal_info"]["journal-tla"] = resolved
        else:
            print(f"[WARN] None of portal-jid '{portal_jid}', "
                  f"accronym '{accronym}', tla '{tla}' found in BreakDown.json. "
                  f"Proceeding with journal-tla '{tla}'.")

        return article_info
    def _post_process_article_info(self, driver, article_info):
        """
        Called after create_dic scrapes article_tags on the Article Info tab.

        1. If article-type is "Other" and other-article-type has a valid
           value, replace article-type with other-article-type.
        2. Extract the journal ID (e.g. "CMS") from the jrnl-name field
           which is formatted as "CMS - Journal of Cutaneous Medicine
           and Surgery".
        """
        # ── 1. Handle "Other" article type ────────────────────────────
        art_info = article_info.get("article_info", {})
        art_type = (art_info.get("article-type") or "").strip()
        if art_type.lower() == "other":
            other_type = (art_info.get("other-article-type") or "").strip()
            if other_type and other_type.lower() not in ("other", "none", ""):
                print(f"[INFO] Article type is 'Other', "
                      f"using other-article-type: '{other_type}'")
                article_info["article_info"]["article-type"] = other_type

        # ── 2. Extract Journal ID from jrnl-name ─────────────────────
        jrnl_name = (art_info.get("jrnl-name") or "").strip()
        if jrnl_name and " - " in jrnl_name:
            portal_jid = jrnl_name.split(" - ")[0].strip().upper()
            print(f"[INFO] Journal ID from jrnl-name: '{portal_jid}'")
            article_info["article_info"]["journal-id-from-portal"] = portal_jid

        return article_info

    def login_regular(self, article_id, ms_no, journal_id, process_folder):
        driver = self.driver
        search_term = article_id or ms_no

        article_info = {
            "article_info": {},
            "journal_info": {},
            "authors_info": {},
            "funder_info": {}
        }

        # Try UK first, then US — looking for a row that matches journal_id
        # via prefix rule (e.g. metadata 'JVDI' matches table 'JVD').
        locale_order = ["UK", "US"]

        table = None
        matched_row = None
        active_loc = None

        # First pass: find a locale where we get a JID-matching row
        candidate_results = {}  # loc -> (table, matched_row)
        for loc in locale_order:
            print(f"[INFO]: Trying {loc} locale for article search "
                  f"(expected JID='{journal_id}')...")
            t, mrow = self._login_and_search(loc, search_term, expected_jid=journal_id)
            if t is None:
                print(f"[INFO]: No matching records found in {loc}.")
                continue
            candidate_results[loc] = (t, mrow)
            if mrow is not None:
                table = t
                matched_row = mrow
                active_loc = loc
                print(f"[INFO]: Locale '{loc}' has matching row {mrow} "
                      f"for JID '{journal_id}'.")
                break
            else:
                print(f"[INFO]: Locale '{loc}' has results but no JID match.")

        # Fallback: no locale had a JID-matching row, but some had results.
        # Use the first locale with results and warn the user.
        if table is None and candidate_results:
            for loc in locale_order:
                if loc in candidate_results:
                    table, _ = candidate_results[loc]
                    matched_row = 1  # fall back to first row
                    active_loc = loc
                    print(f"[WARN] No JID match in any locale. Falling back to "
                          f"{loc} locale, row 1. Expected JID was '{journal_id}'.")
                    break

        if table is None:
            print("[INFO]: No matching records found in any locale.")
            return False, article_id, None

        # Pick article_id from the matched row's td[3]
        if matched_row is None:
            matched_row = 1
        try:
            article_id = table.find_element(
                By.XPATH, f"./tbody/tr[{matched_row}]/td[3]"
            ).text
        except Exception:
            article_id = table.find_element(By.XPATH, "./tbody/tr/td[3]").text

        self.create_info_xml(
            driver,
            article_info,
            active_loc,
            article_id,
            process_folder,
            False,
            matched_row=matched_row
        )

        return True, article_id, article_info["journal_info"]["journal-tla"]


# create_info = GetArticleId()
# info_found, article_id, jrn_tla = create_info.smart_login(None, "TAB-25-09-199R1", "TAB", "V:\\FOR_BREAKDOWN\\PROCESS\\Article_Attachments-2026-01-27-21-09-26")
# create_info.create_info_xml()
# print(article_id)
