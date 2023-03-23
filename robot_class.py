from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select

from time import sleep




def get_information_class(list_class, country):
    code_country = {'AR': 1, 'CL': 9, 'CO': 11, 'MX': 39, 'PE': 48}
    c = code_country[country]

    driver = webdriver.Chrome(executable_path = '../chromedriver.exe')
    driver.implicitly_wait(30)
    driver.maximize_window()
    driver.get('https://wwclassprod.inc.hpicorp.net/pls/WWCLASS_PROD/SYN_CLS_PART_LOOKUP.part_lookup?p_nl_ctry_cd=000')
    list_class = '\n'.join(list_class)
    default_login = driver.find_element(by=By.XPATH, value='/html/body/table[5]/tbody/tr/td[3]/table/tbody/tr[3]/td/table/tbody/tr/td[1]/table/tbody/tr[5]/td/table/tbody/tr[3]/td[3]/input')
    default_login.click()
    part_lookup = driver.find_element(by=By.LINK_TEXT, value='Part Lookup')
    part_lookup.click()

    select_country = Select(WebDriverWait(driver, 10).until(EC.visibility_of_element_located((By.NAME, 'p_fil_ctry_cd'))))
    select_country.select_by_index(c)

    multiple_part = driver.find_element(by=By.NAME, value='p_multi_part_nrs')
    multiple_part.click()
    multiple_part.send_keys(list_class)
    submit = driver.find_element(by=By.XPATH, value='/html/body/table[5]/tbody/tr/td[3]/table/tbody/tr[3]/td/form/table/tbody/tr[2]/td[2]/table/tbody/tr[14]/td[2]/input')
    submit.click()
    export_results = driver.find_element(by=By.XPATH, value="//input[@value='Export results']")
    export_results.click()
    sleep(10)
    driver.quit()

