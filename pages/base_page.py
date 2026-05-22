from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


class BasePage:
    def __init__(self, driver):
        self.driver = driver

    def wait_for(self, by, value, timeout=15):
        return WebDriverWait(self.driver, timeout).until(
            EC.presence_of_element_located((by, value))
        )
