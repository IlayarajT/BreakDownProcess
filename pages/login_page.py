import time
from selenium.webdriver.common.by import By
from selenium.webdriver import ActionChains
from pages.base_page import BasePage
from utils.metrics import metrics


class LoginPage(BasePage):

    @metrics.track("login")
    def login(self, user, password):
        self.wait_for(By.ID, "ctl00_SmartMasterContent_rtbuserlogin").send_keys(user)
        self.wait_for(By.ID, "ctl00_SmartMasterContent_rtbpasswd").send_keys(password)
        self.wait_for(
            By.XPATH,
            "//form//span[contains(@class,'RadButton')]",
        ).click()

        time.sleep(3)

        content = self.wait_for(
            By.XPATH, "//span[normalize-space()='Content']"
        )
        submenu = self.wait_for(
            By.XPATH,
            "//span[normalize-space()='Maintain Journal/Issue/Article']"
        )

        ActionChains(self.driver).move_to_element(content).move_to_element(submenu).perform()
        submenu.click()
