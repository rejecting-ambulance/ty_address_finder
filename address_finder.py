import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import unicodedata

def setup_chrome_driver():
    options = Options()
    options.add_argument('--headless=new')  # 使用新版 headless 模式（更穩定）
    options.add_argument('--disable-gpu')   # Windows上有時必須加
    options.add_argument('--no-sandbox')    # 如果你在 Linux 或 Docker 中沒權限時加
    options.add_argument('--window-size=1920,1080')  # 確保元素能正確呈現
    options.add_argument('--disable-logging')  # 關閉日誌
    options.add_argument('--log-level=3')  # 只顯示致命錯誤
    options.add_argument('--disable-dev-shm-usage')  # 關閉開發者工具相關訊息
    options.add_argument('--disable-extensions')  # 關閉擴展相關訊息
    options.add_argument('--disable-web-security')  # 關閉網路安全警告
    options.add_argument('--disable-features=VizDisplayCompositor')  # 關閉顯示相關警告
    options.add_argument('--silent')  # 靜默模式
    options.add_argument('--disable-crash-reporter')  # 關閉崩潰報告
    options.add_argument('--disable-in-process-stack-traces')  # 關閉進程內堆疊追蹤
    options.add_argument('--disable-dev-tools')  # 關閉開發者工具
    options.add_argument('--disable-background-timer-throttling')  # 關閉背景計時器
    options.add_argument('--disable-renderer-backgrounding')  # 關閉渲染器背景化
    options.add_argument('--disable-backgrounding-occluded-windows')  # 關閉被遮蔽視窗的背景化
    options.add_argument('--disable-ipc-flooding-protection')  # 關閉IPC洪水保護
    options.add_experimental_option('excludeSwitches', ['enable-logging'])  # 排除日誌開關
    options.add_experimental_option('useAutomationExtension', False)  # 關閉自動化擴展
    
    # 關閉瀏覽器日誌
    options.add_experimental_option('excludeSwitches', ['enable-automation'])
    options.add_experimental_option('useAutomationExtension', False)
    
    # 設定日誌級別
    import logging
    import os
    
    # 設定環境變數來抑制Chrome的錯誤訊息
    os.environ['CHROME_LOG_FILE'] = 'NUL'  # Windows
    # os.environ['CHROME_LOG_FILE'] = '/dev/null'  # Linux/Mac 請用這行替換上面一行
    
    logging.getLogger('selenium').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)

    driver = webdriver.Chrome(options=options)
    return driver

def wait_class_change(driver, element_id, origin_class, old_class, timeout=10):
    
    WebDriverWait(driver, timeout).until(
        lambda d: d.find_element(By.ID, element_id).get_attribute('class') != origin_class
    )
    WebDriverWait(driver, timeout).until(
        lambda d: d.find_element(By.ID, element_id).get_attribute('class') != old_class
    )


def search_address(driver, wait, address):
    driver.get('https://addressrs.moi.gov.tw/address/index.cfm?city_id=68000')
    address_box = wait.until(EC.presence_of_element_located((By.ID, 'FreeText_ADDR')))
    submit_button = driver.find_element(By.ID, 'ext-comp-1010')

    address_box.clear()
    address_box.send_keys(address)
    submit_button.click()
    
    # 原表單wait
    #wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ext-gen107"]/div[1]/table/tbody/tr/td[2]/div')))
    
    wait_class_change(driver, 'ext-gen97', 'x-panel-bwrap', 'x-panel-bwrap x-masked-relative x-masked')
    
    try:
        result = driver.find_element(By.XPATH, '//*[@id="ext-gen107"]/div/table/tbody/tr/td[2]/div')
        return result.text.strip()
    except:
        return "找不到結果"

def simplify_address(address):
    """
    將地址簡化為：去除里、鄰與號後面的文字。
    回傳：(原地址, 簡化地址, 後綴)
    """
    original_address = address  # 保留原始輸入

    # 移除「里」與「鄰」段
    address = re.sub(r'[^ ]{1,5}里', '', address)
    address = re.sub(r'\d{1,3}鄰', '', address)

    # 處理號後的尾端文字
    split_chars = ['號', '及', '、', '.']
    split_indices = [(address.find(c), c) for c in split_chars if address.find(c) != -1]

    if split_indices:
        split_indices.sort()
        index, char = split_indices[0]

        if char == '號':
            simplified = address[:index + 1]
            suffix = address[index + 1:]
        else:
            simplified = address[:index]
            suffix = address[index:]
    else:
        simplified = address
        suffix = ''

    return original_address.strip(), simplified.strip(), suffix.strip()


def fullwidth_to_halfwidth(text):
    half_text = ''
    for char in text:
        code = ord(char)
        if code == 0x3000:
            code = 0x0020
        elif 0xFF01 <= code <= 0xFF5E:
            code -= 0xFEE0
        half_text += chr(code)
    return half_text

def format_simplified_address(addr):
    # 數字轉半形
    addr = fullwidth_to_halfwidth(addr)

    # 去除「0」開頭的鄰編號，如 003鄰 ➜ 3鄰
    addr = re.sub(r'(\D)0*(\d+)鄰', r'\1\2鄰', addr)

    # 阿拉伯數字轉中文段號（1~9段）
    num_to_chinese = {'1': '一', '2': '二', '3': '三', '4': '四', '5': '五',
                      '6': '六', '7': '七', '8': '八', '9': '九'}

    def replace_road_section(match):
        num = match.group(1)
        return num_to_chinese.get(num, num) + '段'

    addr = re.sub(r'(\d)段', replace_road_section, addr)

    return addr.strip()


def remove_ling_with_condition(full_address):
    if "高上里" in full_address:
        # 是高上里 → 不刪除
        return full_address
    else:
        # 不是高上里 → 若有「里」和「鄰」，刪除「里」與「鄰」之間的所有字元（含「鄰」）
        # 範例：AAA里SD鄰11路 → AAA里11路
        return re.sub(r'(里).*?鄰', r'\1', full_address)


def process_no_result_address(original_address):
    """
    處理查無結果的地址：如果原地址有「里」，直接放到「不含鄰的地址」欄
    高上里特殊處理：須有「鄰」才放至「不含鄰的地址」欄
    """
    if "里" in original_address:
        # 如果是高上里，須有「鄰」才放到「不含鄰的地址」欄
        if "高上里" in original_address:
            if re.search(r'\d+鄰', original_address):
                return original_address
            else:
                return "查詢失敗"
        else:
            # 其他有「里」的地址直接放入
            return original_address
    else:
        return "查詢失敗"

def visual_len(text):
    """ 計算文字的實際顯示寬度 """
    width = 0
    for ch in text:
        if unicodedata.east_asian_width(ch) in ('F', 'W'):
            width += 2  # 全形、寬字元
        else:
            width += 1  # 半形
    return width

def pad_text(text, target_width):
    """ 補足空格讓文字達到指定寬度 """
    pad_len = target_width - visual_len(text)
    return text + ' ' * max(pad_len, 0)

def main(file_path):
    df = pd.read_excel(file_path)
    addresses = df['查詢地址'].tolist()
    driver = setup_chrome_driver()
    wait = WebDriverWait(driver, 10)

    full_addresses = []
    simplified_addresses = []

    # 先取最大地址長度，用來對齊箭頭
    max_len = 50

    serial_numbers = []  # A 欄流水號
    for i, address in enumerate(addresses, start=1):
        serial_numbers.append(i)

        if not address or str(address).strip() == 'nan':
            print(f"{i:04d}. 空白資料")
            full_addresses.append("")
            simplified_addresses.append("")
            continue

        try:
            data_address, shorter_address, last_address = simplify_address(address)
            result_address = search_address(driver, wait, shorter_address)

            if result_address == "找不到結果":
                output = f"{i:04d}. {pad_text(address, max_len)} → 查無結果"
                print(output)
                full_addresses.append("查無結果")
                simplified = process_no_result_address(data_address)
            else:
                full_address = f'桃園市{result_address}{last_address}'
                full_address = fullwidth_to_halfwidth(full_address)
                output = f"{i:04d}. {pad_text(address, max_len)} → {full_address}"
                print(output)
                full_addresses.append(full_address)
                simplified = remove_ling_with_condition(full_address)

            formatted_simplified = format_simplified_address(simplified)
            simplified_addresses.append(formatted_simplified)

        except Exception as e:
            print(f"{i:04d}. {pad_text(address, max_len)} → 查詢失敗")
            full_addresses.append("查詢失敗")
            simplified = process_no_result_address(address)
            formatted_simplified = format_simplified_address(simplified)
            simplified_addresses.append(formatted_simplified)

    driver.quit()

    # 寫入 Excel 前，補滿 A 欄序號（即第一欄）
    # 補回序號（根據實際處理過的地址數）
    df.iloc[:, 0] = list(range(1, len(addresses) + 1))

    df['完整地址'] = full_addresses
    df['不含鄰的地址'] = simplified_addresses
    df.to_excel(file_path, index=False)
    print(f"✅ 已完成，請查看：{file_path}")


if __name__ == '__main__':
    file_path = 'address_data.xlsx'
    main(file_path)

