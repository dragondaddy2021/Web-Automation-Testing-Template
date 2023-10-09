# 網頁測試用例自動化測試腳本開始

# 需要用到的模組導入區
from selenium import webdriver
from selenium.common.exceptions import NoSuchElementException
from time import sleep
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import warnings
from selenium.webdriver import ActionChains 
from selenium.webdriver.common.keys import Keys
import sys
from selenium.webdriver.chrome.options import Options
import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import smtplib
from email.mime.image import MIMEImage
from pathlib import Path
import io
from PIL import Image
from selenium.webdriver.common.action_chains import ActionChains
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2 import service_account
import os
from google.oauth2.service_account import Credentials
from google.auth import credentials
import json
import requests
from git import Repo
import gitlab
import git
from collections import defaultdict
import subprocess
from git import Repo
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import UnexpectedAlertPresentException
from email.message import EmailMessage


'''
重複行為的函式準備區
'''
# 先宣告一個初始序號的全域變數為0(其他非函式但有序號的地方也要加上serial_number += 1及{serial_number}
serial_number = 0

# 以ID定位某元素驗證五秒內是否存在、可見、可點擊的函式
def verify_clickable_ID(browser, element_ID, element_name):
    global serial_number
    serial_number += 1
    try:
        element = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.ID, element_ID)))
        # 找到元素後的操作
        print(f"{serial_number}. 驗證{element_name} ，確實於渲染後五秒內存在、可見、可點擊")
    except:
        # 若不可點擊，則印出錯誤後進入下一步
        print(f"{serial_number}. ▲警告▲：驗證{element_name} ，未於五秒內渲染完成")
        pass

# 以XPATH定位某元素驗證五秒內是否存在、可見、可點擊的函式
def verify_clickable_XPATH(browser, element_XPATH, element_name):
    global serial_number
    serial_number += 1
    try:
        element = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.XPATH, element_XPATH)))
        # 找到元素後的操作
        print(f"{serial_number}. 驗證{element_name} ，確實於渲染後五秒內存在、可見、可點擊")
    except:
        # 若不可點擊，則印出錯誤後進入下一步
        print(f"{serial_number}. ▲警告▲：驗證{element_name} ，未於五秒內渲染完成")
        pass

# 以CLASS_NAME定位某元素驗證五秒內是否存在、可見、可點擊的函式
def verify_clickable_CLASS_NAME(browser, element_CLASS_NAME, element_name):
    global serial_number
    serial_number += 1
    try:
        element = WebDriverWait(browser, 5).until(EC.element_to_be_clickable((By.CLASS_NAME, element_CLASS_NAME)))
        # 找到元素後的操作
        print(f"{serial_number}. 驗證{element_name} ，確實於渲染後五秒內存在、可見、可點擊")
    except:
        # 若不可點擊，則印出錯誤後進入下一步
        print(f"{serial_number}. ▲警告▲：驗證{element_name} ，未於五秒內渲染完成")
        pass

# 以ID定位某元素驗證五秒內是否存在、可見的函式
def verify_visibility_ID(browser, element_ID, element_name):
    global serial_number
    serial_number += 1
    try:
        element = WebDriverWait(browser, 5).until(EC.visibility_of_element_located((By.ID, element_ID)))
        # 找到元素後的操作
        print(f"{serial_number}. 驗證{element_name}，確實於渲染後五秒內存在、可見")
    except:
        # 若不可見，則印出錯誤訊息後進入下一步
        print(f"{serial_number}. ▲警告▲：驗證{element_name}，未於五秒內渲染完成")
        pass

# 以XPATH定位某元素驗證五秒內是否存在、可見的函式
def verify_visibility_XPATH(browser, element_XPATH, element_name):
    global serial_number
    serial_number += 1
    try:
        element = WebDriverWait(browser, 5).until(EC.visibility_of_element_located((By.XPATH, element_XPATH)))
        # 找到元素後的操作
        print(f"{serial_number}. 驗證{element_name}，確實於渲染後五秒內存在、可見")
    except:
        # 若不可見，則印出錯誤訊息後進入下一步
        print(f"{serial_number}. ▲警告▲：驗證{element_name}，未於五秒內渲染完成")
        pass

# 以ID實際點擊某個元素
def click_element_ID(browser, ID):
    element = browser.find_element(By.ID, ID)
    element.click()

# 以XPATH實際點擊某個元素
def click_element_XPATH(browser, XPATH):
    element = browser.find_element(By.XPATH, XPATH)
    element.click()

# 跳轉頁面後等待五秒內是否有新渲染的HTML完成，否則則拋出錯誤
def wait_for_page_load(browser, timeout=5):
    WebDriverWait(browser, timeout).until(EC.staleness_of(browser.find_element(By.TAG_NAME, 'html')))

# 確認網址是否成功跳轉至預期網址
def check_url(browser, element_name):
    expected_url = "預期的URL"
    global serial_number
    serial_number += 1
    try:
        url_match = browser.current_url == expected_url
        if url_match:
            # 當前網址與預期網址相符合後的操作
            print(f"{serial_number}. 驗證{element_name}，可成功引導至XXX網址")
        else:
            # 若網址不相符，則印出錯誤訊息後進入下一步
            print(f"{serial_number}. ▲警告▲：驗證{element_name}，與預期XXX網址不相符，預期網址為：{expected_url}，當前網址為：{browser.current_url}")
    except:
        # 處理其他可能的異常
        pass

# 確認網址是否成功跳轉至預期的網址,url輸入預期網址，element_name輸入驗證項目
def match_url(browser, url, element_name):
    expected_url = url
    global serial_number
    serial_number += 1
    try:
        url_match = browser.current_url == expected_url
        if url_match:
            # 當前網址與預期網址相符合後的操作
            print(f"{serial_number}. 驗證{element_name}，可成功引導至預期網址")
        else:
            # 若網址不相符，則印出錯誤訊息後進入下一步
            print(f"{serial_number}. ▲警告▲：驗證{element_name}，與預期跳轉網址不相符，預期網址為：{expected_url}，當前網址為：{browser.current_url}")
    except:
        # 處理其他可能的異常
        pass

# 操作輸入帳密登入流程的函式
def loging_process(element_name):
    global serial_number
    serial_number += 1
    try:
        # 定位帳密輸入框及登入按鈕
        username_input = browser.find_element(By.ID, 'AccountID')
        password_input = browser.find_element(By.ID, 'PasswordID')
        login_input = browser.find_element(By.ID, 'btn_loginID')
        # 在帳號欄位輸入帳號
        username_input.send_keys('帳號')
        # 輸完帳號後TAB到密碼欄
        username_input.send_keys(Keys.TAB)
        # 在密碼欄位輸入密碼
        password_input.send_keys('密碼')
        # 點擊登入
        login_input.click()
        # 登入流程成功回報
        print(f"{serial_number}. 驗證{element_name}登入流程，本次登入操作成功")
    except:
        # 處理其他可能的異常
        print(f"{serial_number}. ▲警告▲：驗證{element_name}登入流程，本次登入操作失敗")
        pass


# 計算網頁跳轉後渲染耗時的函式，element_name填入驗證項目為何，XPATH替換為您要點擊的按鈕
def calculate_page_visible_render_time(element_name, XPATH, timeout=10):
    global serial_number
    serial_number += 1
    # 以XPATH實際點擊某個元素
    element = browser.find_element(By.XPATH, XPATH)
    element.click()
    # 等待頁面可見
    wait = WebDriverWait(browser, timeout)
    start_time = time.time()
    wait.until(EC.visibility_of(browser.find_element(By.TAG_NAME, 'body')))
    end_time = time.time()

    # 計算可見渲染時間
    visible_render_time = end_time - start_time

    # 印出結果
    print(f"{serial_number}. 驗證{element_name}點擊跳轉後至body渲染完成共耗時 {visible_render_time}秒")

# 獲取視窗寬度的函式
def current_width(element_name):
    global serial_number
    serial_number += 1
    window_width = browser.execute_script('return window.innerWidth;')
    print(f"{serial_number}. 驗證{element_name}，目前的螢幕寬度為{window_width}")


'''
基本設定區
'''

# 執行CMD命令
def run_cmd(command):
    try:
        result = subprocess.run(command, shell=True, check=True, capture_output=True, text=True)
        output = result.stdout.strip()  # 獲取CMD的輸出結果
        return output
    except subprocess.CalledProcessError as e:
        print(f'CMD command execution failed: {e}')
        return None

# 在CMD中創建每日日期為檔名的資料夾
def create_daily_folder():
    now = datetime.datetime.now()
    date_str = now.strftime("%Y%m%d")
    time_str = now.strftime("%H%M%S")
    dirname = f'{date_str}_{time_str}'
    command = f'mkdir {dirname}'
    run_cmd(command)
    return dirname

# 執行函式
dirname = create_daily_folder()

# 建立 Chrome 選項
chrome_options = Options()

'''
# 設定手機版的chrome console模擬行動裝置
mobile_emulation = {"deviceName": "手機型號"}
chrome_options.add_experimental_option("mobileEmulation", mobile_emulation)

# 啟用不開瀏覽器的Headless模式
chrome_options.add_argument('--headless')  
'''

# 設定視窗縮放比例
chrome_options.add_argument("--force-device-scale-factor=0.85")

# 啟用 JavaScript
chrome_options.add_argument("--enable-javascript")

# 模擬當前瀏覽器的User Agent
chrome_options.add_argument("--user-agent=自己在consloe查詢")

# 取得當前時間
current_time = datetime.datetime.now()

# 將時間格式化為特定的字串格式
time_str = current_time.strftime('%H-%M-%S')

# 取得當天日期
today = datetime.date.today()

# 將日期轉換為特定格式的字串
date_str = today.strftime('%Y-%m-%d')

# 將print輸出同步重定向至文件和字符串
class SyncedOutput:
    def __init__(self, *outputs):
        self.outputs = outputs

    def write(self, message):
        for output in self.outputs:
            output.write(message)

    def flush(self):
        for output in self.outputs:
            output.flush()

# 檔名加上日期及時間
filename = f'web_log_{date_str}_{time_str}.txt'
file_path = os.path.join(dirname, filename)
file_output = open(file_path, 'w')
string_output = io.StringIO()
synced_output = SyncedOutput(sys.stdout, file_output, string_output)
sys.stdout = synced_output

# 去除Selenium版本太新因此不建議使用的醜醜警告
warnings.filterwarnings("ignore", category=DeprecationWarning)

# 設定 Chrome 瀏覽器的執行檔路徑， r"..." 來指定 ChromeDriver 路徑，這可以確保程式碼中的反斜線（\）被當作字元而非跳脫字元處理。
chrome_driver_path = r"ChromeDriver的執行擋路徑"

# 啟動 Chrome 瀏覽器(加上路徑和插件的選項)
browser = webdriver.Chrome(options=chrome_options)
browser.service.executable_path = chrome_driver_path


# 網頁載入等待時間統一設置隱式等待時間為10秒
browser.implicitly_wait(10)

# 最大化瀏覽器視窗
browser.maximize_window()

# 滾動至頁面底部
browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")

"""
以下為網頁自動化測試用例
"""

print("【測試流程開始】")

print("\n 以下為網頁自動化測試用例結果")

# 前往網站
browser.get("網站URL")

# 確認網址是否成功跳轉至預期網址
check_url(browser, '是否成功跳轉預期網址')

# 確認登入狀態下的首頁顯示

# 定位登入按鈕並點擊
click_element_XPATH(browser, 'XPATH路徑')

# 呼叫輸入帳密操作登入流程的函式
loging_process('通常狀態下登入的')

# 確認點擊某廣告後是否能正常跳轉

# 確認點擊某廣告後頁面渲染完成耗時多久
calculate_page_visible_render_time('點擊某廣告', 'XPATH路徑', timeout=10)

# 確認是否於五秒內出現某廣告，且存在、可見、可點擊，以第一則某廣告為例子
verify_clickable_XPATH(browser, 'XPATH路徑', '點擊跳轉後的第一則某廣告')

# 回到首頁
browser.back()

# 確認某區塊是否於五秒內存在、可見、可點擊
verify_clickable_XPATH(browser, 'XPATH路徑', '某區塊')

# 確認螢幕寬度是否RWD自適應寬度後小於PC版寬度1268
current_width('RWD縮放')

# 回到首頁
browser.back()

# 點擊某按鈕
click_element_XPATH(browser, '某按鈕的XPATH路徑')

# 切換到新開啟的網頁(要跳到新開的視窗內，不然會停在原頁無法獲取文本)
handles = browser.window_handles
browser.switch_to.window(handles[-1])

# 切換回去首頁的視窗
handles = browser.window_handles
browser.switch_to.window(handles[0])

'''
其餘測試項目按去識別化刪除
'''

# 關閉瀏覽器
browser.quit()

print("\n【測試流程結束】")

"""
測試流程結束
"""






"""
電子郵件寄送專區
"""
# 建立MIMEMultipart物件
content = MIMEMultipart()

# 動態加上當天日期的郵件標題
subject = f"[自動化測試報告] - {date_str}"
content["subject"] = subject

# 其他郵件設定
# 誰寄的
content["from"] = "寄件者郵件地址"
# 寄信給哪些人
content["to"] = "收件者郵件地址1, 收件者郵件地址2"


# 讀取圖片檔案
with open("圖片檔名.jpg", "rb") as f:
    image_data = f.read()

# 建立MIMEImage物件
image = MIMEImage(image_data)

# 設定圖片的內容類型和檔名
image.add_header("Content-Disposition", "attachment", filename="圖片檔名.jpg")
content.attach(image)

# 郵件內文
content.attach(MIMEText("\n★請確認以下測試項目是否出現▲警告▲內容★\n   "))

# 取得print的訊息
message = string_output.getvalue()

# 附加print的訊息
content.attach(MIMEText(message))

with smtplib.SMTP(host="host的IP", port="25") as smtp:  # 設定SMTP伺服器
    try:
        smtp.send_message(content)  # 寄送郵件
        print("測試結果，已成功寄送!")
    except Exception as e:
        print("Error message:▲警告▲：測試結果，未成功寄送 ", e)





'''
將txt檔push到遠端儲存庫上
'''

# 設定專案路徑
project_path = r'C槽的專案資料夾路徑'

# 切換至專案路徑
os.chdir(project_path)

# 創建文件並指定UTF-8編碼
# 取得字串輸出
file_path = os.path.join(dirname, filename)
file_content = string_output.getvalue()
with io.open(file_path, 'w', encoding='utf-8') as file:
    file.write(file_content)

# 確認目前的工作目錄
current_dir = os.getcwd()
print(f"目前的工作目錄: {current_dir}")

# 初始化 Repo 物件
repo = Repo(project_path)

# 確認目前的 Git 狀態
print(f"目前的 Git 狀態: {repo.git.status()}")

# 添加專案資料夾至暫存區
repo.git.add(dirname)

# 提交變更，並加入提交訊息
commit_message = '自動提交報告到遠端儲存庫'
repo.git.commit('-m', commit_message)

# 推送到遠端儲存庫
origin = repo.remote('origin')
origin.push()





# print重定向使用結束後，將print輸出恢復至控制台
sys.stdout = sys.__stdout__

