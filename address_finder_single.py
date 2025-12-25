import re
import os
import logging
import unicodedata

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
import subprocess


def setup_chrome_driver():
    options = Options()

    # --- Headless 模式設定 ---
    options.add_argument('--headless=new')  # 啟用新版無頭模式，模擬真實瀏覽器但不開視窗
    options.add_argument('--disable-gpu')   # 關閉 GPU 加速，避免在部分環境下造成錯誤
    options.add_argument('--no-sandbox')    # 解除沙盒限制（Linux/Docker 無權限環境必加）
    options.add_argument('--disable-dev-shm-usage')  # 避免 /dev/shm 空間不足導致崩潰（Docker 常見）
    options.add_argument('--window-size=1920,1080')  # 指定視窗大小，確保頁面元素完整載入可見

    # --- 日誌與自動化提示設定 ---
    # 排除特定開關，以隱藏「Chrome 正在受自動化控制」提示及多餘的 console log
    options.add_experimental_option(
        'excludeSwitches', 
        ['enable-logging', 'enable-automation']
    )

    # 關閉 Chrome 自動化擴展功能（減少被網站偵測的機率）
    options.add_experimental_option('useAutomationExtension', False)

    # --- 系統級日誌抑制設定 ---
    # 將 Chrome 的內部 log 輸出導向無效位置
    os.environ['CHROME_LOG_FILE'] = os.devnull  # 使用平台相容的 devnull

    # 降低 Selenium 與 urllib3 的日誌輸出層級，只顯示警告以上訊息
    logging.getLogger('selenium').setLevel(logging.WARNING)
    logging.getLogger('urllib3').setLevel(logging.WARNING)

    # 建立並回傳 WebDriver 物件，使用 Service 隱藏 chromedriver 控制台視窗
    try:
        creationflags = subprocess.CREATE_NO_WINDOW
    except Exception:
        creationflags = 0

    service = Service(log_path=os.devnull)
    try:
        service.creationflags = creationflags
    except Exception:
        pass

    # 額外參數以降低 Chrome GPU 與 log 噪音
    options.add_argument('--log-level=3')
    options.add_argument('--disable-software-rasterizer')

    return webdriver.Chrome(service=service, options=options)


def wait_class_change(driver, element_id, origin_class, old_class, timeout=10):
    
    WebDriverWait(driver, timeout).until(
        lambda d: d.find_element(By.ID, element_id).get_attribute('class') != origin_class
    )
    WebDriverWait(driver, timeout).until(
        lambda d: d.find_element(By.ID, element_id).get_attribute('class') != old_class
    )

def wait_mask_cycle(driver, mask_class='ext-el-mask', timeout=20):
    """
    等待遮罩出現再消失，用於等待查詢完成
    """
    try:
        # Step 1. 等待遮罩出現
        WebDriverWait(driver, timeout/2).until(
            EC.presence_of_element_located((By.CLASS_NAME, mask_class))
        )
        #print("[INFO] 遮罩已出現，開始等待消失...")

    except Exception:
        print("[WARN] 查詢遮罩未出現（可能瞬間出現又消失）")

    # Step 2. 等待遮罩消失
    WebDriverWait(driver, timeout).until_not(
        EC.presence_of_element_located((By.CLASS_NAME, mask_class))
    )
    #print("[INFO] 遮罩已消失，查詢完成。")


def search_address(driver, wait, address):
    driver.get('https://addressrs.moi.gov.tw/address/index.cfm?city_id=68000')
    address_box = wait.until(EC.presence_of_element_located((By.ID, 'FreeText_ADDR')))
    #submit_button = driver.find_element(By.ID, 'ext-comp-1010')
    submit_button = driver.find_element(By.ID, 'ext-gen51')

    address_box.clear()
    address_box.send_keys(address)
    submit_button.click()
    
    # 原表單wait
    #wait.until(EC.presence_of_element_located((By.XPATH, '//*[@id="ext-gen107"]/div[1]/table/tbody/tr/td[2]/div')))
    
    #wait_class_change(driver, 'ext-gen97', 'x-panel-bwrap', 'x-panel-bwrap x-masked-relative x-masked')
    wait_mask_cycle(driver)
    try:
        #result = driver.find_element(By.XPATH, '//*[@id="ext-gen107"]/div/table/tbody/tr/td[2]/div')
        result = driver.find_element(By.XPATH, '//*[@id="ext-gen111"]/div/table/tbody/tr/td[2]/div')
        return result.text.strip()
    except:
        return "找不到結果"

def simplify_address(address):
    """
    簡化輸入地址，供後續查詢用。

    進函式時先將全形數字轉為半形，並移除里/鄰；
    處理路/街與段、號之間的中文/阿拉伯數字之轉換（逐字或簡單單位解析），
    並去除號前之前導零。回傳 (original, simplified, suffix)。
    """
    original_address = address

    # 進函式時先將全形數字轉為半形，方便後續處理
    address = fullwidth_to_halfwidth(address)

    # 移除「里」與「鄰」段
    address = re.sub(r'([\u4e00-\u9fff]{1,5}區)[\u4e00-\u9fff]{1,2}里', r'\1', address)
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

    # 1) 街/路 + 阿拉伯數字(1~2位) + 段 -> 將阿拉伯數字轉為中文段號（支援到十位）
    arabic_digits_map = {0: '零', 1: '一', 2: '二', 3: '三', 4: '四', 5: '五', 6: '六', 7: '七', 8: '八', 9: '九'}

    def arabic_to_chinese_section(n: int) -> str:
        if n <= 0:
            return ''
        if n < 10:
            return arabic_digits_map[n]
        tens, ones = divmod(n, 10)
        if tens == 1:
            return '十' + (arabic_digits_map[ones] if ones else '')
        else:
            return arabic_digits_map[tens] + '十' + (arabic_digits_map[ones] if ones else '')

    def _road_digit_to_chinese(m):
        road = m.group(1)
        num_s = m.group(2)
        num_s = num_s.lstrip('0')
        if not num_s:
            return f"{road}0段"
        n = int(num_s)
        if n >= 1 and n <= 99:
            return f"{road}{arabic_to_chinese_section(n)}段"
        else:
            return f"{road}{num_s}段"

    simplified = re.sub(r'([\u4e00-\u9fff]+(?:路|街))0*([1-9]\d?)段', _road_digit_to_chinese, simplified)

    # 2) 街/路 + 中文數字 + 號 -> 將中文數字逐字或簡單單位解析為阿拉伯數字
    char_to_digit = {'零': '0', '〇': '0', '一': '1', '二': '2', '三': '3', '四': '4', '五': '5',
                     '六': '6', '七': '7', '八': '8', '九': '9'}

    def chinese_to_arabic(s: str) -> str:
        unit_chars = set('十百千')
        if any(ch in unit_chars for ch in s):
            digits_map = {'零': 0, '〇': 0, '一': 1, '二': 2, '三': 3, '四': 4, '五': 5,
                          '六': 6, '七': 7, '八': 8, '九': 9}
            unit_map = {'千': 1000, '百': 100, '十': 10}
            total = 0
            num = 0
            for ch in s:
                if ch in digits_map:
                    num = digits_map[ch]
                elif ch in unit_map:
                    unit_val = unit_map[ch]
                    if num == 0:
                        num = 1
                    total += num * unit_val
                    num = 0
                else:
                    num = 0
            total += num
            return str(total)
        else:
            return ''.join(char_to_digit.get(ch, ch) for ch in s)

    def _road_chinese_to_digit(m):
        road = m.group(1)
        chs = m.group(2)
        arabic = chinese_to_arabic(chs)
        return f"{road}{arabic}號"

    simplified = re.sub(r'([\u4e00-\u9fff]+(?:路|街))([零〇一二三四五六七八九十百千]+)號', _road_chinese_to_digit, simplified)

    # 若路/街 和 號 之間有其他文字，也嘗試把緊接在號前的中文數字轉為阿拉伯數字
    def _convert_between_road_and_hao(m):
        prefix = m.group(1)
        chinese_digits = m.group(2)
        return prefix + chinese_to_arabic(chinese_digits) + '號'

    simplified = re.sub(r'([\u4e00-\u9fff]+(?:路|街).*?)([零〇一二三四五六七八九十百千]+)號', _convert_between_road_and_hao, simplified)

    # 出函式前再確保數字為半形並回傳
    simplified = fullwidth_to_halfwidth(simplified)
    suffix = fullwidth_to_halfwidth(suffix)

    # 號前的數字去除前導零，001 -> 1
    simplified = re.sub(r"(\d+)號", lambda m: str(int(m.group(1))) + '號', simplified)
    suffix = re.sub(r"(\d+)號", lambda m: str(int(m.group(1))) + '號', suffix)

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
    addr = addr.replace(',', '，')  # 半形「,」轉回「，」

    # 去除「0」開頭的鄰編號，如 003鄰 ➜ 3鄰
    addr = re.sub(r'(\D)0*(\d+)鄰', r'\1\2鄰', addr)

    # 號前的數字去除前導零，001 -> 1, 016 -> 16, 010 -> 10
    addr = re.sub(r'(\d+)號', lambda m: str(int(m.group(1))) + '號', addr)

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
        if "桃園市" not in original_address:
            original_address = f'桃園市{original_address}'
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
            print(f"{i}. 空白資料，結束查詢。")
            full_address = ""
            simplified = ""
            formatted_simplified = ""
            break
        
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