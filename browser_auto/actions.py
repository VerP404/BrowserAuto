from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webdriver import WebDriver


def click_xpath(driver: WebDriver, xpath: str) -> None:
    driver.find_element(By.XPATH, xpath).click()


def input_xpath(driver: WebDriver, xpath: str, text: str) -> None:
    element = driver.find_element(By.XPATH, xpath)
    element.clear()
    element.send_keys(text)


def input_enter_xpath(driver: WebDriver, xpath: str, text: str) -> None:
    driver.find_element(By.XPATH, xpath).send_keys(text, Keys.ENTER)


def input_enter_id(driver: WebDriver, element_id: str, text: str) -> None:
    driver.find_element(By.ID, element_id).send_keys(text, Keys.ENTER)
