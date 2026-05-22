import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager


def create_driver(view_web_page: bool):
    options = webdriver.ChromeOptions()
    options.add_experimental_option(
        "excludeSwitches", ["enable-logging"]
    )

    env_headless = os.getenv("HEADLESS", "").lower() == "true"
    if env_headless or not view_web_page:
        options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")

    return webdriver.Chrome(
        options=options,
        service=Service(ChromeDriverManager().install()),
    )
