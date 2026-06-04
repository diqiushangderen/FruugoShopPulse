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
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext
import sys

# Excel结构说明：
# A列: 店铺id | B列: 状态(写入) | C列: SKU数(写入) | D列: 唯一站点(若无则留空)

STATUS_COLORS = {
    "闭店": "FFFF0000",      # 红色
    "假期模式": "FFFFFF00",  # 黄色
    "一切正常": "FF90EE90"  # 浅绿色
}

FRUUGO_SITE_DOMAINS = {
    "South Korea": "fruugo.kr", "Poland": "fruugo.pl", "Spain": "fruugo.es",
    "Japan": "fruugo.co.jp", "Czech Republic": "fruugo.cz", "Finland": "fruugo.fi",
    "United Kingdom": "fruugo.co.uk", "Switzerland": "fruugo.ch", "Germany": "fruugo.de",
    "France": "fruugo.fr", "Italy": "fruugo.it", "Australia": "fruugo.com.au",
    "United States": "fruugo.com", "Saudi Arabia": "fruugo.sa", "Austria": "fruugo.at",
    "Bahrain": "fruugo.bh", "Belgium": "fruugo.be", "China": "fruugo.cn",
    "Denmark": "fruugo.dk", "Egypt": "fruugo.eg", "Philippines": "fruugo.ph",
    "Greece": "fruugo.gr", "Netherlands": "fruugo.nl", "India": "fruugo.in",
    "Ireland": "fruugo.ie", "Israel": "fruugo.co.il", "Canada": "fruugo.ca",
    "Qatar": "fruugo.qa", "Luxembourg": "fruugo.lu", "Malaysia": "fruugo.my",
    "Norway": "fruugo.no", "New Zealand": "fruugo.co.nz", "Portugal": "fruugo.pt",
    "South Africa": "fruugo.co.za", "Romania": "fruugo.ro", "Singapore": "fruugo.sg",
    "Sweden": "fruugo.se", "Slovakia": "fruugo.sk", "Turkey": "fruugo.com.tr",
    "Hungary": "fruugo.hu", "UAE": "fruugo.ae"
}

excel_lock = threading.Lock()
message_queue = Queue()


class RedirectText:
    """标准输出重定向至 GUI 文本框组件"""
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
    """GUI 交互界面管理类"""
    def __init__(self, root):
        self.root = root
        self.root.title("店铺状态与 SKU 提取工具")
        self.root.geometry("800x600")

        self.main_frame = ttk.Frame(root, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # 文件选择
        self.file_frame = ttk.Frame(self.main_frame)
        self.file_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.file_path = tk.StringVar()
        self.select_file_btn = ttk.Button(self.file_frame, text="选择Excel文件", command=self.select_file)
        self.select_file_btn.grid(row=0, column=0, padx=5)
        self.file_label = ttk.Label(self.file_frame, textvariable=self.file_path, wraplength=500)
        self.file_label.grid(row=0, column=1, padx=5, sticky=tk.W)

        # 线程设置
        self.thread_frame = ttk.Frame(self.main_frame)
        self.thread_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.thread_label = ttk.Label(self.thread_frame, text="线程数:")
        self.thread_label.grid(row=0, column=0, padx=5)
        self.thread_count = tk.StringVar(value="5")
        self.thread_entry = ttk.Entry(self.thread_frame, textvariable=self.thread_count, width=5)
        self.thread_entry.grid(row=0, column=1, padx=5)

        # 控制台按钮
        self.control_frame = ttk.Frame(self.main_frame)
        self.control_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        self.start_btn = ttk.Button(self.control_frame, text="开始运行", command=self.start_check)
        self.start_btn.grid(row=0, column=0, padx=5)
        self.stop_btn = ttk.Button(self.control_frame, text="结束运行", command=self.stop_check, state=tk.DISABLED)
        self.stop_btn.grid(row=0, column=1, padx=5)

        # 日志流显示
        self.status_frame = ttk.Frame(self.main_frame)
        self.status_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        self.status_text = scrolledtext.ScrolledText(self.status_frame, height=20, width=80)
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(3, weight=1)
        self.status_frame.columnconfigure(0, weight=1)
        self.status_frame.rowconfigure(0, weight=1)

        self.is_running = False
        self.threads = []
        sys.stdout = RedirectText(self.status_text)

    def select_file(self):
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")]
        )
        if filename:
            self.file_path.set(filename)

    def start_check(self):
        if not self.file_path.get():
            print("请先选择Excel文件！")
            return
        try:
            thread_count = int(self.thread_count.get())
            if thread_count < 1:
                print("线程数必须大于0！")
                return
        except ValueError:
            print("请输入有效的线程数！")
            return

        self.is_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.select_file_btn.config(state=tk.DISABLED)
        self.thread_entry.config(state=tk.DISABLED)

        check_thread = threading.Thread(target=self.run_check, daemon=True)
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
            print(f"发生错误: {str(e)}")
        finally:
            self.root.after(0, self.stop_check)


def random_sleep(min_seconds=0.8, max_seconds=1.5):
    time.sleep(random.uniform(min_seconds, max_seconds))


def setup_driver():
    """初始化配置防反爬自动化浏览器环境"""
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
        raise FileNotFoundError(f"ChromeDriver 未能在当前目录下找到: {driver_path}")

    driver = uc.Chrome(options=options, use_subprocess=False, driver_executable_path=driver_path, version_main=None)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def handle_cookie_consent(driver, wait, thread_id):
    """检测并跳过页面的 Cookie 授权弹窗"""
    try:
        wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "dialog[data-modal-type='cookie-consent']")))
        try:
            accept_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-test-id='cookie-consent-accept-all']")))
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
        return True


def update_excel_status(excel_path, row_index, status, sale_num):
    """写入检测状态(B列)与清洗后的SKU数(C列)并应用单元格填充色"""
    with excel_lock:
        try:
            wb = load_workbook(excel_path)
            ws = wb.active

            status_cell = ws.cell(row=row_index + 2, column=2)
            status_cell.value = status
            status_cell.fill = PatternFill(start_color=STATUS_COLORS[status], end_color=STATUS_COLORS[status], fill_type='solid')

            sku_cell = ws.cell(row=row_index + 2, column=3)
            sku_cell.value = int(sale_num) if sale_num.isdigit() else 0

            wb.save(excel_path)
            print(f"已更新Excel第{row_index + 2}行：状态={status}，SKU数={sale_num}")
        except Exception as e:
            print(f"更新Excel失败: {str(e)}")


def extract_sale_skus(driver, wait, thread_id):
    """提取店铺列表页文本中的在售商品总数并做数字化清洗"""
    sale_num = "0"
    try:
        sale_target_xpath = '//*[@id="main"]/div/div/div[2]/div[1]/p'
        sale_element = wait.until(EC.visibility_of_element_located((By.XPATH, sale_target_xpath)))
        sale_text = sale_element.text.strip()

        if not sale_text:
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
    return sale_num


def check_store_status_after_enter(driver, wait, merchant_id, excel_path, row_index, thread_id, sale_num):
    """步入商品详情页，执行精细化库存状态交叉校验"""
    try:
        HOLIDAY_XPATH = "//*[text()='Currently unavailable']"
        try:
            wait.until(EC.presence_of_element_located((By.XPATH, HOLIDAY_XPATH)))
            print(f"线程 {thread_id}: 店铺 {merchant_id} 匹配到不可售，判定【假期模式】")
            update_excel_status(excel_path, row_index, "假期模式", sale_num)
            return
        except TimeoutException:
            pass

        is_normal = False
        css_selector = "#main > div.Product.container > div.row.Product__Top > div.col.col-right.col-auto-md > div:nth-child(2) > div > div.d-none.d-md-block > div > div.mt-16 > div:nth-child(1) > span"
        try:
            status_elem = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, css_selector)))
            if status_elem.text.strip() == "In stock":
                is_normal = True
        except:
            stock_xpaths = ["//strong[text()='In stock']", "//span[text()='In stock']"]
            for xpath in stock_xpaths:
                if len(driver.find_elements(By.XPATH, xpath)) > 0:
                    is_normal = True
                    break

        if is_normal:
            print(f"线程 {thread_id}: 店铺 {merchant_id} 匹配到现货，判定【一切正常】")
            update_excel_status(excel_path, row_index, "一切正常", sale_num)
        else:
            print(f"线程 {thread_id}: 店铺 {merchant_id} 无正常现货标识，判定【假期模式】")
            update_excel_status(excel_path, row_index, "假期模式", sale_num)
    except Exception as e:
        print(f"线程 {thread_id}: 验证状态异常: {str(e)}，兜底判定【假期模式】")
        update_excel_status(excel_path, row_index, "假期模式", sale_num)


def check_es_store(driver, wait, merchant_id, thread_id):
    """多站点验证：跳转芬兰站点测试"""
    try:
        handle_cookie_consent(driver, wait, thread_id)
        es_link = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/footer/div[1]/div[4]/div/div[1]/div/ul/li[10]/a")))
        driver.execute_script("arguments[0].click();", es_link)
        random_sleep(3, 5)

        h1_text = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div/div/div[2]/div/h1"))).text.strip()
        return h1_text == "Sorry, we didn't find any products that match your query"
    except Exception:
        return False


def check_au_store(driver, wait, merchant_id, thread_id):
    """多站点验证：跳转瑞士站点测试"""
    try:
        handle_cookie_consent(driver, wait, thread_id)
        au_link = wait.until(EC.element_to_be_clickable((By.XPATH, "/html/body/div/footer/div[1]/div[4]/div/div[1]/div/ul/li[37]/a")))
        driver.execute_script("arguments[0].click();", au_link)
        random_sleep(3, 5)

        h1_text = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div[1]/main/div/div/div[2]/div/h1"))).text.strip()
        return h1_text == "Emme valitettavasti löytäneet hakuasi vastaavia tuotteita"
    except Exception:
        return False


def check_store_with_retry(driver, merchant_id, excel_path, row_index, thread_id, is_single_site, site_domain=None):
    """带有指数重试的主体检测流逻辑"""
    max_retries = 3
    retry_count = 0
    url = f"https://www.{site_domain}/search/?merchantId={merchant_id}&language=en" if (is_single_site and site_domain) else f"https://www.fruugo.co.uk/search/?merchantId={merchant_id}&language=en"

    while retry_count < max_retries:
        try:
            driver.get(url)
            wait = WebDriverWait(driver, 10)
            handle_cookie_consent(driver, wait, thread_id)
            random_sleep(2, 4)

            sale_num = extract_sale_skus(driver, wait, thread_id)

            try:
                product_link = wait.until(EC.presence_of_element_located((By.XPATH, "/html/body/div/main/div/div/div[2]/div[3]/div[1]/a")))
                driver.execute_script("arguments[0].click();", product_link)
                random_sleep(3, 5)
                check_store_status_after_enter(driver, wait, merchant_id, excel_path, row_index, thread_id, sale_num)
                return
            except (TimeoutException, NoSuchElementException, ElementClickInterceptedException):
                if is_single_site:
                    try:
                        h1_text = wait.until(EC.presence_of_element_located((By.XPATH, "//main//h1"))).text.strip()
                        is_empty = "Sorry, we didn't find any products that match your query" in h1_text
                    except:
                        is_empty = False
                    update_excel_status(excel_path, row_index, "闭店" if is_empty else "假期模式", sale_num)
                    return

                if check_es_store(driver, wait, merchant_id, thread_id):
                    if check_au_store(driver, wait, merchant_id, thread_id):
                        update_excel_status(excel_path, row_index, "闭店", sale_num)
                    else:
                        update_excel_status(excel_path, row_index, "假期模式", sale_num)
                else:
                    update_excel_status(excel_path, row_index, "假期模式", sale_num)
                return
        except Exception as e:
            retry_count += 1
            if retry_count >= max_retries:
                update_excel_status(excel_path, row_index, "闭店", "0")
            else:
                time.sleep(3)


def worker(thread_id, merchant_ids, excel_path, start_index):
    """线程工作节点核心分发循环"""
    driver = setup_driver()
    df = pd.read_excel(excel_path)

    try:
        for i, merchant_id in enumerate(merchant_ids):
            actual_index = start_index + i
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
            except:
                pass

            check_store_with_retry(driver, merchant_id, excel_path, actual_index, thread_id, is_single_site, site_domain)

            # 每跑完3家店铺，触发策略性反爬长休
            if (i + 1) % 3 == 0:
                time.sleep(random.uniform(15, 30))
            else:
                random_sleep(5, 10)
    finally:
        driver.quit()


def check_store_status_multi_thread(excel_path, num_threads=5):
    """多线程总调度中心"""
    try:
        df = pd.read_excel(excel_path)
        merchant_ids = df.iloc[:, 0].astype(str).tolist()
        total_stores = len(merchant_ids)

        base_size = total_stores // num_threads
        remainder = total_stores % num_threads
        threads = []
        start_idx = 0

        for i in range(num_threads):
            chunk_size = base_size + (1 if i < remainder else 0)
            chunk = merchant_ids[start_idx:start_idx + chunk_size]
            if chunk:
                thread = threading.Thread(target=worker, args=(i + 1, chunk, excel_path, start_idx))
                threads.append(thread)
                thread.start()
            start_idx += chunk_size

        for thread in threads:
            thread.join()
        print("\n===== 任务全部执行完毕 =====")
    except Exception as e:
        print(f"\n主线程调度异常: {str(e)}")


if __name__ == "__main__":
    root = tk.Tk()
    app = CheckWebSituationGUI(root)
    root.mainloop()
