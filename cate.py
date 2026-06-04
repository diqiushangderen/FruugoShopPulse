import pandas as pd
import time
import os
import random
import threading
from queue import Queue
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, ElementClickInterceptedException
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import math
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import sys
from datetime import datetime

# 读取的xlsx文件结构为
# 店铺id      状态(程序写入)    SKU数(程序写入)   唯一站点(你填写，单站点店铺)
# 28000       空                空                South Korea
# 28000       空                空                Poland

# 定义状态和对应的填充颜色
STATUS_COLORS = {
    "闭店": "FFFF0000",  # 红色
    "假期模式": "FFFFFF00",  # 黄色
    "一切正常": "FF90EE90"  # 浅绿色
}

# ====================== Fruugo 站点域名映射 ======================
FRUUGO_SITE_DOMAINS = {
    "South Korea": "fruugo.kr",
    "Poland": "fruugo.pl",
    "Spain": "fruugo.es",
    "Japan": "fruugo.co.jp",
    "Czech Republic": "fruugo.cz",
    "Finland": "fruugo.fi",
    "United Kingdom": "fruugo.co.uk",
    "Switzerland": "fruugo.ch",
    "Germany": "fruugo.de",
    "France": "fruugo.fr",
    "Italy": "fruugo.it",
    "Australia": "fruugo.com.au",
    "United States": "fruugo.com",
    "Saudi Arabia": "fruugo.sa",
    "Austria": "fruugo.at",
    "Bahrain": "fruugo.bh",
    "Belgium": "fruugo.be",
    "China": "fruugo.cn",
    "Denmark": "fruugo.dk",
    "Egypt": "fruugo.eg",
    "Philippines": "fruugo.ph",
    "Greece": "fruugo.gr",
    "Netherlands": "fruugo.nl",
    "India": "fruugo.in",
    "Ireland": "fruugo.ie",
    "Israel": "fruugo.co.il",
    "Canada": "fruugo.ca",
    "Qatar": "fruugo.qa",
    "Luxembourg": "fruugo.lu",
    "Malaysia": "fruugo.my",
    "Norway": "fruugo.no",
    "New Zealand": "fruugo.co.nz",
    "Portugal": "fruugo.pt",
    "South Africa": "fruugo.co.za",
    "Romania": "fruugo.ro",
    "Singapore": "fruugo.sg",
    "Sweden": "fruugo.se",
    "Slovakia": "fruugo.sk",
    "Turkey": "fruugo.com.tr",
    "Hungary": "fruugo.hu",
    "UAE": "fruugo.ae"
}

# 添加线程锁，用于Excel写入
excel_lock = threading.Lock()

# 创建一个队列用于线程间通信
message_queue = Queue()


class RedirectText:
    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = ""

    def write(self, string):
        self.buffer += string
        if string.endswith('\n'):
            self.text_widget.insert(tk.END, self.buffer)
            self.text_widget.see(tk.END)
            self.buffer = ""
            self.text_widget.update()

    def flush(self):
        if self.buffer:
            self.text_widget.insert(tk.END, self.buffer)
            self.text_widget.see(tk.END)
            self.buffer = ""
            self.text_widget.update()


class CheckWebSituationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("店铺状态+SKU数提取工具")
        self.root.geometry("800x600")

        # 创建主框架
        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 文件选择区域
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.file_path = tk.StringVar()
        self.select_file_btn = ttk.Button(self.file_frame, text="选择Excel文件", command=self.select_file)
        self.select_file_btn.grid(row=0, column=0, padx=5)

        self.file_label = ttk.Label(self.file_frame, textvariable=self.file_path, wraplength=500)
        self.file_label.grid(row=0, column=1, padx=5, sticky=tk.W)

        # 线程数设置区域
        self.thread_frame = ttk.Frame(self.main_frame)
        self.thread_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.thread_label = ttk.Label(self.thread_frame, text="线程数:")
        self.thread_label.grid(row=0, column=0, padx=5)

        self.thread_count = tk.StringVar(value="5")
        self.thread_entry = ttk.Entry(self.thread_frame, textvariable=self.thread_count, width=5)
        self.thread_entry.grid(row=0, column=1, padx=5)

        # 控制按钮区域
        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)

        self.start_btn = ttk.Button(self.control_frame, text="开始运行", command=self.start_check)
        self.start_btn.grid(row=0, column=0, padx=5)

        self.stop_btn = ttk.Button(self.control_frame, text="结束运行", command=self.stop_check, state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=1, padx=5)

        # 状态显示区域
        self.status_frame = ttk.Frame(self.main_frame)
        self.status_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)

        self.status_text = scrolledtext.ScrolledText(self.status_frame, height=20, width=80)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 配置grid权重
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(3, weight=1)
        self.status_frame.columnconfigure(0, weight=1)
        self.status_frame.rowconfigure(0, weight=1)

        # 初始化变量
        self.is_running = False
        self.threads = []

        # 重定向标准输出到文本框
        self.redirect = RedirectText(self.status_text)
        sys.stdout = self.redirect

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            self.file_path.set(filename)

    def start_check(self):
        if not self.file_path.get():
            self.status_text.insert(tk.END, "请先选择Excel文件！\n")
            return

        try:
            thread_count = int(self.thread_count.get())
            if thread_count < 1:
                self.status_text.insert(tk.END, "线程数必须大于0！\n")
                return
        except ValueError:
            self.status_text.insert(tk.END, "请输入有效的线程数！\n")
            return

        self.is_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.select_file_btn.config(state=tk.DISABLED)
        self.thread_entry.config(state=tk.DISABLED)

        check_thread = threading.Thread(target=self.run_check)
        check_thread.daemon = True
        check_thread.start()

    def stop_check(self):
        self.is_running = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.select_file_btn.config(state=tk.NORMAL)
        self.thread_entry.config(state=tk.NORMAL)

        for thread in self.threads:
            thread.join()
        self.threads.clear()

    def run_check(self):
        try:
            thread_count = int(self.thread_count.get())
            check_store_status_multi_thread(self.file_path.get(), thread_count)
        except Exception as e:
            self.status_text.insert(tk.END, f"发生错误: {str(e)}\n")
        finally:
            self.root.after(0, self.stop_check)


def random_sleep(min_seconds=0.8, max_seconds=1.5):
    time.sleep(random.uniform(min_seconds, max_seconds))


def setup_driver():
    options = uc.ChromeOptions()
    options.add_argument('--disable-blink-features=AutomationControlled')
    options.add_argument('--disable-infobars')

    user_agents = [
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/132.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0.0.0 Safari/537.36"
    ]
    options.add_argument(f'user-agent={random.choice(user_agents)}')

    thread_id = threading.current_thread().ident
    profile_dir = f"C:/chrome_profiles/profile_{thread_id}"
    if not os.path.exists(profile_dir):
        os.makedirs(profile_dir)
    options.add_argument(f"--user-data-dir={profile_dir}")

    driver_path = os.path.join(os.path.dirname(__file__), "chromedriver.exe")
    if not os.path.exists(driver_path):
        raise FileNotFoundError(f"ChromeDriver 文件不存在于 {driver_path}")

    driver = uc.Chrome(
        options=options,
        use_subprocess=False,
        driver_executable_path=driver_path,
        version_main=None
    )
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    driver.implicitly_wait(0)
    return driver


def handle_cookie_consent(driver, wait, thread_id):
    try:
        cookie_dialog = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "dialog[data-modal-type='cookie-consent']"))
        )
        try:
            accept_button = wait.until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-test-id='cookie-consent-accept-all']"))
            )
            accept_button.click()
            print(f"线程 {thread_id}: 已接受Cookie设置")
            random_sleep(1, 2)
            return True
        except:
            try:
                driver.execute_script("document.querySelector('dialog[data-modal-type=\"cookie-consent\"]').close();")
                random_sleep(1, 2)
                return True
            except:
                print(f"线程 {thread_id}: 无法处理Cookie对话框")
                return False
    except:
        print(f"线程 {thread_id}: 未检测到Cookie对话框")
        return True


def update_excel_status(excel_path, row_index, status, sale_num):
    """更新Excel：第2列写状态，第3列写SALE后的SKU数"""
    with excel_lock:
        try:
            wb = load_workbook(excel_path)
            ws = wb.active

            status_cell = ws.cell(row=row_index + 2, column=2)
            status_cell.value = status
            status_cell.fill = PatternFill(
                start_color=STATUS_COLORS[status],
                end_color=STATUS_COLORS[status],
                fill_type='solid'
            )

            sku_cell = ws.cell(row=row_index + 2, column=3)
            sku_cell.value = int(sale_num) if sale_num.isdigit() else 0

            wb.save(excel_path)
            print(f"已更新Excel第{row_index + 2}行：状态={status}，SKU数={sale_num}")

        except Exception as e:
            print(f"更新Excel失败: {str(e)}")


def extract_sale_skus(driver, wait, thread_id):
    sale_num = "0"
    try:
        print(f"线程 {thread_id}: 开始提取SKU数...")
        sale_target_xpath = '//*[@id="main"]/div/div/div[2]/div[1]/p'
        sale_element = wait.until(
            EC.visibility_of_element_located((By.XPATH, sale_target_xpath))
        )
        sale_text = sale_element.text.strip()
        print(f"线程 {thread_id}: 目标位置文本：{sale_text}")

        if not sale_text:
            print(f"线程 {thread_id}: 目标位置文本为空")
            return sale_num

        if 'on ' in sale_text:
            after_on = sale_text.split('on ')[1]
            num_segment = ''
            for char in after_on:
                if char.isdigit() or char in (',', '.'):
                    num_segment += char
                else:
                    break
            if num_segment:
                sale_num = num_segment.replace(',', '').replace('.', '')
                sale_num = ''.join(filter(str.isdigit, sale_num)) or "0"

        print(f"线程 {thread_id}: 提取到SKU数：{sale_num}")

    except Exception as e:
        print(f"线程 {thread_id}: 提取SKU数失败: {str(e)}")
        sale_num = "0"
    return sale_num


def check_store_status_after_enter(driver, wait, merchant_id, excel_path, row_index, thread_id, sale_num):
    try:
        print(f"线程 {thread_id}: 开始精细化检查店铺 {merchant_id} 状态...")
        HOLIDAY_XPATH = "//*[text()='Currently unavailable']"
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, HOLIDAY_XPATH)))
            print(f"线程 {thread_id}: 店铺 {merchant_id} → 精准匹配Currently unavailable，判定【假期模式】")
            update_excel_status(excel_path, row_index, "假期模式", sale_num)
            return
        except TimeoutException:
            print(f"线程 {thread_id}: 未检测到Currently unavailable，进入正常状态验证...")
        except Exception as e:
            print(f"线程 {thread_id}: 假期模式判断异常: {str(e)}，进入正常状态验证...")

        status_text = None
        is_normal = False
        css_selector = "#main > div.Product.container > div.row.Product__Top > div.col.col-right.col-auto-md > div:nth-child(2) > div > div.d-none.d-md-block > div > div.mt-16 > div:nth-child(1) > span"
        try:
            status_elem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_selector)))
            status_text = status_elem.text.strip()
            if status_text == "In stock":
                is_normal = True
        except:
            print(f"线程 {thread_id}: CSS选择器未匹配到状态，尝试XPath验证...")
            stock_xpaths = [
                "//strong[text()='In stock']",
                "//span[text()='In stock']",
                "//p//strong[contains(text(), 'In stock')]",
                "//div//span[contains(text(), 'In stock')]"
            ]
            for xpath in stock_xpaths:
                if len(driver.find_elements(By.XPATH, xpath)) > 0:
                    is_normal = True
                    status_text = "In stock"
                    break

        if is_normal and status_text == "In stock":
            print(f"线程 {thread_id}: 店铺 {merchant_id} → 匹配到In stock，判定【一切正常】")
            update_excel_status(excel_path, row_index, "一切正常", sale_num)
        else:
            print(f"线程 {thread_id}: 店铺 {merchant_id} → 无In stock标识，非闭店，判定【假期模式】")
            update_excel_status(excel_path, row_index, "假期模式", sale_num)

    except Exception as e:
        print(f"线程 {thread_id}: 店铺 {merchant_id} 状态检查异常: {str(e)}")
        print(f"线程 {thread_id}: 店铺 {merchant_id} → 检查异常，兜底判定【假期模式】")
        update_excel_status(excel_path, row_index, "假期模式", sale_num)


def check_es_store(driver, wait, merchant_id, thread_id):
    try:
        handle_cookie_consent(driver, wait, thread_id)
        es_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div/footer/div[1]/div[4]/div/div[1]/div/ul/li[10]/a"))
        )
        driver.execute_script("arguments[0].click();", es_link)
        random_sleep(3, 5)

        h1_text = wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div/div/div[2]/div/h1"))
        ).text.strip()
        is_closed = h1_text == "Sorry, we didn't find any products that match your query"
        print(f"线程 {thread_id}: 芬兰站点检查结果 → {'已关闭' if is_closed else '正常开放'}")
        return is_closed
    except Exception as e:
        print(f"线程 {thread_id}: 芬兰站点检查异常: {str(e)} → 判定为开放")
        return False


def check_au_store(driver, wait, merchant_id, thread_id):
    try:
        handle_cookie_consent(driver, wait, thread_id)
        au_link = wait.until(
            EC.element_to_be_clickable((By.XPATH, "/html/body/div/footer/div[1]/div[4]/div/div[1]/div/ul/li[37]/a"))
        )
        driver.execute_script("arguments[0].click();", au_link)
        random_sleep(3, 5)

        h1_text = wait.until(
            EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div/div/div[2]/div/h1"))
        ).text.strip()
        is_closed = h1_text == "Emme valitettavasti löytäneet hakuasi vastaavia tuotteita"
        print(f"线程 {thread_id}: 瑞士站点检查结果 → {'已关闭' if is_closed else '正常开放'}")
        return is_closed
    except Exception as e:
        print(f"线程 {thread_id}: 瑞士站点检查异常: {str(e)} → 判定为开放")
        return False


def check_store_with_retry(driver, merchant_id, excel_path, row_index, thread_id, is_single_site, site_domain=None):
    max_retries = 3
    retry_count = 0

    if is_single_site and site_domain:
        url = f"https://www.{site_domain}/search/?merchantId={merchant_id}&language=en"
        print(f"线程 {thread_id}: 单站点店铺，直接访问 => {url}")
    else:
        url = f"https://www.fruugo.co.uk/search/?merchantId={merchant_id}&language=en"

    while retry_count < max_retries:
        try:
            print(f"\n线程 {thread_id}: 检查商户 {merchant_id} (第{retry_count + 1}次尝试)")
            driver.get(url)
            wait = WebDriverWait(driver, 10)

            handle_cookie_consent(driver, wait, thread_id)
            random_sleep(2, 4)

            sale_num = extract_sale_skus(driver, wait, thread_id)

            try:
                product_link = wait.until(
                    EC.presence_of_element_located((By.XPATH, "/html/body/div/main/div/div/div[2]/div[3]/div[1]/a"))
                )
                driver.execute_script("arguments[0].click();", product_link)
                random_sleep(3, 5)

                check_store_status_after_enter(driver, wait, merchant_id, excel_path, row_index, thread_id, sale_num)
                return

            except (TimeoutException, NoSuchElementException, ElementClickInterceptedException) as e:
                print(f"线程 {thread_id}: 无法进入商品详情页: {str(e)} → 启动多站点交叉验证")

                if is_single_site:
                    try:
                        h1_text = wait.until(EC.presence_of_element_located((By.XPATH, "//main//h1"))).text.strip()
                        is_empty = "Sorry, we didn't find any products that match your query" in h1_text
                    except:
                        is_empty = False

                    if is_empty:
                        print(f"线程 {thread_id}: 单站点无商品 → 闭店")
                        update_excel_status(excel_path, row_index, "闭店", sale_num)
                    else:
                        print(f"线程 {thread_id}: 单站点正常 → 假期模式")
                        update_excel_status(excel_path, row_index, "假期模式", sale_num)
                    return

                es_closed = check_es_store(driver, wait, merchant_id, thread_id)
                if es_closed:
                    au_closed = check_au_store(driver, wait, merchant_id, thread_id)
                    if au_closed:
                        print(f"线程 {thread_id}: 商户 {merchant_id} → 所有站点均关闭，判定【闭店】")
                        update_excel_status(excel_path, row_index, "闭店", sale_num)
                    else:
                        print(f"线程 {thread_id}: 商户 {merchant_id} → 仅芬兰站点关闭，判定【假期模式】")
                        update_excel_status(excel_path, row_index, "假期模式", sale_num)
                else:
                    print(f"线程 {thread_id}: 商户 {merchant_id} → 芬兰站点开放，判定【假期模式】")
                    update_excel_status(excel_path, row_index, "假期模式", sale_num)
                return

        except Exception as e:
            print(f"线程 {thread_id}: 第{retry_count + 1}次检查失败: {str(e)}")
            retry_count += 1
            if retry_count < max_retries:
                print(f"线程 {thread_id}: 3秒后进行下一次重试...")
                time.sleep(3)
            else:
                print(f"线程 {thread_id}: 商户 {merchant_id} → 达到最大重试次数，判定【闭店】")
                update_excel_status(excel_path, row_index, "闭店", "0")


def worker(thread_id, merchant_ids, excel_path, start_index):
    print(f"线程 {thread_id}: 启动成功")
    driver = setup_driver()
    df = pd.read_excel(excel_path)

    try:
        for i, merchant_id in enumerate(merchant_ids):
            actual_index = start_index + i
            print(f"\n===== 线程 {thread_id}: 开始检查第{actual_index + 1}个店铺: {merchant_id} =====")

            is_single_site = False
            site_domain = None
            try:
                if len(df.columns) >= 4:
                    site_val = df.iloc[actual_index, 3]
                    if pd.notna(site_val) and str(site_val).strip() != "":
                        country = str(site_val).strip()
                        site_domain = FRUUGO_SITE_DOMAINS.get(country)
                        if site_domain:
                            is_single_site = True
                            print(f"线程 {thread_id}: 识别为【单站点】=> {country}")
            except:
                pass

            if not is_single_site:
                print(f"线程 {thread_id}: 识别为【多站点店铺】")

            check_store_with_retry(
                driver=driver,
                merchant_id=merchant_id,
                excel_path=excel_path,
                row_index=actual_index,
                thread_id=thread_id,
                is_single_site=is_single_site,
                site_domain=site_domain
            )

            if (i + 1) % 3 == 0:
                wait_time = random.uniform(15, 30)
                print(f"线程 {thread_id}: 已检查3个店铺，反爬休息 {wait_time:.1f} 秒...")
                time.sleep(wait_time)
            else:
                random_sleep(5, 10)
    finally:
        driver.quit()
        print(f"\n线程 {thread_id}: 所有任务处理完成，已退出")


def check_store_status_multi_thread(excel_path, num_threads=5):
    try:
        df = pd.read_excel(excel_path)
        merchant_ids = df.iloc[:, 0].astype(str).tolist()
        total_stores = len(merchant_ids)
        print(f"===== 任务开始 =====")
        print(f"共检测 {total_stores} 个店铺，使用 {num_threads} 个线程并行处理")
        print(f"====================\n")

        base_size = total_stores // num_threads
        remainder = total_stores % num_threads
        threads = []
        start_idx = 0

        for i in range(num_threads):
            chunk_size = base_size + (1 if i < remainder else 0)
            chunk = merchant_ids[start_idx:start_idx + chunk_size]
            if chunk:
                thread = threading.Thread(
                    target=worker,
                    args=(i + 1, chunk, excel_path, start_idx)
                )
                threads.append(thread)
                thread.start()
                print(f"初始化线程 {i + 1} 成功，处理店铺范围：{start_idx + 1}-{start_idx + chunk_size}")
            start_idx += chunk_size

        for thread in threads:
            thread.join()

        print(f"\n===== 所有任务完成 =====")
        print(f"检测结果已写入Excel，格式说明：")
        print(f"- 第2列：店铺状态（闭店/假期模式/一切正常），对应颜色标记")
        print(f"- 第3列：SALE区域提取的纯数字SKU数")
        print(f"========================")

    except Exception as e:
        print(f"\n多线程主函数执行错误: {str(e)}")

# 无敌完美
if __name__ == "__main__":
    root = tk.Tk()
    app = CheckWebSituationGUI(root)
    root.mainloop()