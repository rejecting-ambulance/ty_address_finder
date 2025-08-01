import re
#import os
#import json
import unicodedata

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options


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
    #import os
    
    # 設定環境變數來抑制Chrome的錯誤訊息
    #os.environ['CHROME_LOG_FILE'] = 'NUL'  # Windows
    # os.environ['CHROME_LOG_FILE'] = '/dev/null'  # Linux/Mac 請用這行替換上面一行
    
    logging.getLogger('selenium').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)

    driver = webdriver.Chrome(options=options)
    return driver


'''
def load_exception_rules(json_path='exception_rules.json'):
    if os.path.exists(json_path):
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {"require_ling": []}
'''


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
    將地址簡化為：去除里、鄰與號後的文字，並將 '-' 替換為 '之'。
    回傳：(原地址, 簡化地址, 後綴)
    """
    original_address = address  # 保留原始輸入

    # 移除「區」之後的 XX里（保留前後），避免誤刪區名
    address = re.sub(r'([\u4e00-\u9fff]{1,5}區)[\u4e00-\u9fff]{1,2}里', r'\1', address)
    # 移除 XXX鄰（1~3位數）但保留後面的地址（若有），可加入 lookahead 或結合 word boundary
    address = re.sub(r'(\d{1,3})鄰', '', address)

    # 將簡化地址中的 '-' 取代為 '之'
    address = address.replace('-', '之')

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
    
    #print(f"原地址: {original_address}, 簡化地址: {simplified}, 後綴: {suffix}")
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
    '''
    結果格式化：
    1. 數字轉半形
    2. 去除空格
    3. 將「-」轉回「之」
    4. 去除「0」開頭的鄰編號，如 003鄰 ➜ 3鄰
    5. 阿拉伯數字轉中文段號（1~9段）
    
    '''    
    
    # 數字轉半形
    addr = fullwidth_to_halfwidth(addr)
    addr = addr.replace(' ', '')  # 去除空格
    addr = addr.replace('-', '之')  # 將「-」轉回「之」

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

#EXCEPTION_RULES = load_exception_rules()
def remove_ling_with_condition(full_address):
    # 若地址中有例外名單的里，則不刪除
    '''
    for special_li in EXCEPTION_RULES.get("require_ling", []):
        if special_li in full_address:
            return full_address
    '''
    # 否則執行標準簡化：刪除「里」與「鄰」間文字（含鄰）
    return re.sub(r'(里).*?鄰', r'\1', full_address)


def process_no_result_address(original_address):
    """
    處理查無結果的地址：如果原地址有「里」，直接放到「不含鄰的地址」欄
    高上里特殊處理：須有「鄰」才放至「不含鄰的地址」欄
    """
    if "里" in original_address:
        '''
        for special_li in EXCEPTION_RULES.get("require_ling", []):
            if special_li in original_address:
                # 例外里須包含「鄰」才能保留
                if re.search(r'\d+鄰', original_address):
                    return original_address
                else:
                    return "查詢失敗"
        '''
        # 非例外里，只要有「里」就保留
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

def main():

    print("===============今天想去哪阿?===============")

    driver = setup_chrome_driver()
    wait = WebDriverWait(driver, 10)
    i = 0
    while True:

        address = input('給我地址：\n')    
        i += 1
        max_len = 55  # 用來對齊箭頭

        if not address or str(address).strip() == '':
            print(f"{i}. 空白資料")
            full_address = ""
            simplified = ""
            formatted_simplified = ""
        else:
            try:
                data_address, shorter_address, last_address = simplify_address(address)
                result_address = search_address(driver, wait, shorter_address)

                if result_address == "找不到結果":
                    
                    full_address = "查無結果"
                    
                    simplified = process_no_result_address(data_address)
                    formatted_simplified = format_simplified_address(simplified)

                    
                else:

                    full_address = f'桃園市{result_address}{last_address}'
                    full_address = fullwidth_to_halfwidth(full_address)

                                        
                    simplified = remove_ling_with_condition(full_address)
                    formatted_simplified = format_simplified_address(simplified)
                    
                    formatted_simplified = format_simplified_address(full_address)


            except Exception as e:
                full_address = "查詢失敗"
                simplified = process_no_result_address(data_address)
                formatted_simplified = format_simplified_address(simplified)

        simplified = remove_ling_with_condition(full_address)

        output = f"{i:03d}. {pad_text(address, max_len)}\n   → {pad_text(formatted_simplified, max_len)}\n   → {pad_text(simplified, max_len)}"
        print(f'{output}\n')

    #driver.quit()


if __name__ == '__main__':

    main()

