import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def search_address(driver, wait, address):

    address_box = wait.until(EC.presence_of_element_located((By.ID, 'FreeText_ADDR')))
    submit_button = driver.find_element(By.ID, 'ext-comp-1010')

    # 紀錄舊資料區塊
    #old_table = driver.find_element(By.XPATH, '//*[@id="ext-gen105"]').text
    address_box.clear()
    address_box.send_keys(address)
    submit_button.click()
    #time.sleep(3)
    
    # 等待新資料出現
    wait.until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="ext-gen107"]/div[1]/table/tbody/tr/td[2]/div'))
    )
    
    # 讀取結果
    result = driver.find_element(By.XPATH, '//*[@id="ext-gen107"]/div[1]/table/tbody/tr/td[2]/div')
    print(result)
    return result.text.strip()

if __name__ == '__main__':

    address = '中壢區中正路76號'

    driver = webdriver.Chrome()
    driver.get('https://addressrs.moi.gov.tw/address/index.cfm?city_id=68000')
    wait = WebDriverWait(driver, 10)

    try:
        full_address = search_address(driver, wait, address)
        print(f"{address} → {full_address}")
    except Exception as e:
        print(f"{address} 查詢失敗：{e}")


    driver.quit()


