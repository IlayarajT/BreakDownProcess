import time
from selenium.common.exceptions import (
    TimeoutException,
    StaleElementReferenceException,
)

def retry(func, retries=3, delay=1):
    for attempt in range(retries):
        try:
            return func()
        except (TimeoutException, StaleElementReferenceException):
            if attempt == retries - 1:
                raise
            time.sleep(delay)
