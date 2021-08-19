from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager
import time
import pandas as pd
from selenium.common.exceptions import NoSuchElementException

login = input('Please scan your login: ')
password = input('Please scan your password: ')

time.sleep(2)
driver = webdriver.Chrome(ChromeDriverManager().install())
url = "http://kuna/auth/login/?next=/index"
driver.get(url)

#site log in
user_name = driver.find_element_by_xpath("""//*[@id="id_username"]""")
user_name.send_keys(login)
pass_form = driver.find_element_by_xpath("""//*[@id="id_password"]""")
pass_form.send_keys(password)
driver.find_element_by_xpath("""//*[@id="submit-id-submit"]""").click()

# shop order
shop_order = driver.find_element_by_xpath("""//*[@id="bs-example-navbar-collapse-1"]/ul[1]/ul/li[1]/a""").click()
time.sleep(2)

#read excel and make list of serials
get_data = pd.read_excel(r'C:\\Users\\Waldemar.lusiak\\Desktop\\report.xlsx')
serial_numbers_list = get_data['Serial Number'].tolist()

#creating dictionary to get titan test for each serial assigned
try:
    titan_test_dict = {}
    for serial_number in serial_numbers_list:
        serial_number_window = driver.find_element_by_xpath("""//*[@id="serial"]""")
        serial_number_window.send_keys(serial_number)
        get_serial_data = WebDriverWait(driver, 12).until(
        EC.presence_of_element_located((By.XPATH, """/html/body/div[1]/div[2]/div[2]/div/div[1]/div[1]/div/button"""))).click()
        time.sleep(3)
        try:
            titan_test_result = driver.find_element_by_xpath("""//*[@id="grid_so_contract_rec_8"]/td[2]/div""").text
            titan_test_dict.update({serial_number:titan_test_result})
            serial_number_window.clear()
        except NoSuchElementException as e:
            titan_test_dict.update({serial_number:'No ShopOrder available'})
        pass
except NoSuchElementException as e:
    print('No Such Element has occurred or visible')
pass

driver.quit()

# save the titan test in excel for each serial in second column
try:
    into_excel = pd.DataFrame(titan_test_dict.items(),columns=['Serial Number','Titan Test result'])
    into_excel.to_excel(excel_writer='C:\\Users\\Waldemar.lusiak\\Desktop\\report.xlsx',header=True,)

except NoSuchElementException as e:
    print('something went wrong')
pass