import time
from bs4 import BeautifulSoup
from lxml import etree
from selenium.webdriver.common.by import By
from pages.base_page import BasePage
from utils.retry import retry
from utils.metrics import metrics


class JournalPage(BasePage):

    @metrics.track("journal_scan")
    def extract_journals(self, loc):
        table = self.wait_for(
            By.XPATH,
            "//table[@id='SmartMasterContent_dlJournalList']/tbody",
        )
        html = table.get_attribute("innerHTML")

        soup = BeautifulSoup(html, "lxml")
        for t in soup.find_all("table", {"summary": "combobox"}):
            t.decompose()

        dom = etree.HTML(str(soup))
        journals = []

        row = 1
        for _ in dom.xpath("/html/body/tr"):
            for a in dom.xpath(f"/html/body/tr[{row}]/td/a"):
                journals.append((a.text, a.attrib["id"], loc))
            row += 1

        return journals

    def open_journal(self, button_id):
        retry(lambda: self.driver.find_element(By.ID, button_id).click())
        time.sleep(1)
