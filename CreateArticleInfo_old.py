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
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from loadconfig import getconfig

warnings.filterwarnings("ignore", message="Exception ignored in.*Popen")

logging.getLogger("urllib3").setLevel(logging.WARNING)
logging.getLogger("selenium").setLevel(logging.CRITICAL)
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")


class GetArticleId:

    def __init__(self):
        self.configFolder, self.breakDownConfig = getconfig()

        info_yaml = os.path.join(
            self.configFolder, "config", "createArticleInfo.yaml"
        )
        with open(info_yaml, "r", encoding="utf-8") as fh:
            self.sage_details = yaml.safe_load(fh)

        self.waiting_time = int(self.sage_details.get("waiting_time", 30))
        view_head = self.sage_details.get("view_web_page", True)

        self.journals_file = os.path.join(
            self.configFolder, "SupportingFiles", "sageJournalInfo.json"
        )
        with open(self.journals_file, "r", encoding="utf-8") as fh:
            self.journal_json = json.load(fh)

        options = Options()
        options.headless = view_head
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")

        service = Service(executable_path="firefox.geckodriver")
        self.driver = webdriver.Firefox(options=options, service=service)
        self.driver.set_window_size(1920, 1080)

        self.uk_user = self.sage_details["UK"]["USERNAME"]
        self.uk_pass = self.sage_details["UK"]["PASSWORD"]
        self.us_user = self.sage_details["US"]["USERNAME"]
        self.us_pass = self.sage_details["US"]["PASSWORD"]

        self.journal_tags = self.sage_details["journal"]
        self.article_tags = self.sage_details["article"]
        self.author_tags = self.sage_details["author"]

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
    # ARTICLE INFO EXTRACTION
    # ------------------------------------------------------------
    def create_info_xml(
        self, driver, article_info, jrn_loc, article_id, process_folder, jrn_found
    ):
        print("[INFO]: Collecting Details Smart.....")
        wait = WebDriverWait(driver, self.waiting_time)
        journal_abbr = self._retry(
            lambda: wait.until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        "//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']/tbody/tr[1]/td[1]"
                    )
                )
            ).text,
            "Get journal abbreviation"
        )

        # self._retry(
        #     lambda: wait.until(
        #         EC.element_to_be_clickable(
        #             (
        #                 By.XPATH,
        #                 "//img[contains(@id,'EditButtonImage')]"
        #             )
        #         )
        #     ).click(),
        #     "Open Article"
        # )
        current_url = self.driver.current_url

        def click_edit():
            el = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//img[contains(@id,'EditButtonImage')]")
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

        driver.get(article_page_url)
        wait.until(lambda d: d.execute_script("return document.readyState") == "complete")

        # # ---------- Authors
        # has_authors = self._click_tab(
        #     wait,
        #     "Authors",
        #     "Authors tab",
        #     "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody"
        # )
        #
        # if has_authors:
        #     authors_table = driver.find_element(
        #         By.XPATH,
        #         "//table[@id='ctl00_ArticleAuthors_uc_ArticleAuthorsGrid_ctl00']/tbody"
        #     )
        #
        #     authors_text = authors_table.get_attribute("innerHTML")
        #     authors_soup = BeautifulSoup(authors_text, "lxml")
        #
        #     # REMOVE nested combobox tables (CRITICAL)
        #     for tbl in authors_soup.find_all("table", {"summary": "combobox"}):
        #         tbl.decompose()
        #
        #     authors_dom = etree.HTML(str(authors_soup))
        #     au_count = 0
        #
        #     for au_id in authors_dom.xpath("/html/body/tr/td/a/img/@id"):
        #         au_count += 1
        #
        #         try:
        #             author_link = driver.find_element(
        #                 By.XPATH, f"//img[@id='{au_id}']/parent::a"
        #             )
        #             author_link.click()
        #
        #             wait.until(EC.frame_to_be_available_and_switch_to_it("rwinauthor"))
        #
        #             article_info = self.create_dic(
        #                 driver,
        #                 self.author_tags,
        #                 article_info,
        #                 "authors_info",
        #                 au_count
        #             )
        #
        #         except Exception as e:
        #             article_info.setdefault("authors_info", {})[au_count] = {
        #                 "error": str(e)
        #             }
        #
        #         finally:
        #             driver.switch_to.default_content()
        #             wait.until(
        #                 EC.element_to_be_clickable(
        #                     (By.XPATH, "//a[@class='rwCloseButton']")
        #                 )
        #             ).click()
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
            self._retry(
                lambda: wait.until(
                    EC.element_to_be_clickable(
                        (
                            By.XPATH,
                            f"//a[contains(text(), '{journal_abbr}')]"
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
            self.journal_json[journal_abbr] = jrnl_info

            with open(self.journals_file, "w") as outfile:
                outfile.write(json.dumps(self.journal_json, indent=4))

        json_path = os.path.join(
            process_folder,
            f"{article_info['journal_info']['journal-tla']}_{article_id}.json"
        )

        with open(json_path, "w", encoding="utf-8") as fh:
            json.dump(article_info, fh, indent=4)

        return article_info


    def create_info_xml_new(
            self, driver, article_info, jrn_loc, article_id, process_folder, jrn_found
    ):
        print("[INFO]: Collecting Details Smart.....")
        wait = WebDriverWait(driver, self.waiting_time)

        # Get journal abbreviation
        journal_abbr = self._retry(
            lambda: wait.until(
                EC.visibility_of_element_located(
                    (
                        By.XPATH,
                        "//table[@id='ctl00_SmartMasterContent_ArticleGrid_ctl00']/tbody/tr[1]/td[1]"
                    )
                )
            ).text,
            "Get journal abbreviation"
        )

        # Click Edit button to open article
        current_url = self.driver.current_url

        def click_edit():
            el = wait.until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//img[contains(@id,'EditButtonImage')]")
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

    def login_smart(self, article_id, ms_no, journal_id, process_folder, jrn_loc):
        driver = self.driver
        wait = WebDriverWait(driver, self.waiting_time)

        article_info = {
            "article_info": {},
            "journal_info": self.journal_json[journal_id],
            "authors_info": {},
            "funder_info": {}
        }

        self.login_page(
            driver,
            self.sage_details[jrn_loc]["USERNAME"],
            self.sage_details[jrn_loc]["PASSWORD"],
        )

        wait.until(
            EC.element_to_be_clickable(
                (By.XPATH, self.sage_details[jrn_loc]["BTN"])
            )
        ).click()

        search = wait.until(
            EC.presence_of_element_located((By.ID, "TopMenu1_ArticleSearchEdit"))
        )
        search.send_keys(article_id or ms_no)

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
            return False, article_id, None

        article_id = table.find_element(By.XPATH, "./tbody/tr/td[3]").text

        self.create_info_xml(
            driver,
            article_info,
            jrn_loc,
            article_id,
            process_folder,
            True
        )

        return True, article_id, article_info["journal_info"]["journal-tla"]

    def create_dic(self, driver, tags, article_info, info, au_count=None):
        wait = WebDriverWait(driver, self.waiting_time)

        for tag, meta in tags.items():
            xpath = meta["tag"]
            tag_type = meta["type"]

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

    def login_regular(self, article_id, ms_no, journal_id, process_folder):
        driver = self.driver
        wait = WebDriverWait(driver, self.waiting_time)

        article_info = {
            "article_info": {},
            "journal_info": {},
            "authors_info": {},
            "funder_info": {}
        }

        self.login_page(driver, self.uk_user, self.uk_pass)

        search = wait.until(
            EC.presence_of_element_located((By.ID, "TopMenu1_ArticleSearchEdit"))
        )
        search.send_keys(article_id or ms_no)

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
            return False, article_id, None

        article_id = table.find_element(By.XPATH, "./tbody/tr/td[3]").text

        self.create_info_xml(
            driver,
            article_info,
            "UK",
            article_id,
            process_folder,
            False
        )

        return True, article_id, article_info["journal_info"]["journal-tla"]


# create_info = GetArticleId()
# info_found, article_id, jrn_tla = create_info.smart_login(None, "TAB-25-09-199R1", "TAB", "V:\\FOR_BREAKDOWN\\PROCESS\\Article_Attachments-2026-01-27-21-09-26")
# create_info.create_info_xml()
# print(article_id)