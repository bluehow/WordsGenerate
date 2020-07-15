import time
import os
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By


def capture(word):
    if os.path.exists('images\\{}.png'.format(word)):
        return
    url = r'https://cn.bing.com/images/'
    locator = (By.ID,'sb_form_q')
    chrome_options = Options()
    chrome_options.add_argument('--headless')
    chrome_options.add_argument('--disable-gpu')
    browser = webdriver.Chrome(r"chromedriver.exe",chrome_options=chrome_options)
    browser.set_window_size(1920, 1042)
    browser.implicitly_wait(5)
    browser.get(url)

    time.sleep(5)
    try:
        search_input = WebDriverWait(browser,10,0.1).until(EC.visibility_of_element_located((locator)))
        search_input.send_keys(word)
        browser.find_element_by_id('sb_form_go').click()
    finally:
        try:
            search_input = WebDriverWait(browser,10,0.1).until(EC.visibility_of_element_located((locator)))
            search_input.send_keys(' ')
        finally:
            time.sleep(2)
            browser.get_screenshot_as_file('images\\{}.png'.format(word))
            browser.close()

