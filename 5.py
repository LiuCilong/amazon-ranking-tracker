from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
import openpyxl
import time
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import random
from openpyxl.styles import Font
import os

# ====================== 你的配置（这里先检查！） ======================
# 先确认这里是不是4个关键词！！！
KEYWORDS_LIST = ["modest bathing suit for women",
                 "2 piece bathing suits for women",
                 "bathing suit for women tummy control",
                 "women bathing suits"]
YOUR_ASIN_LIST = ["B0BRWJJDB4", "B0BRWKSJZS"]
MAX_PAGE = 7
ZIP_CODE = "10010"
RESULT_FILE = "亚马逊产品排名记录.xlsx"
COLOR_KEYWORDS = ["black", "white", "red", "blue", "green", "yellow", "pink", "purple", "gray", "brown", "beige",
                  "navy"]
AMAZON_URL = "https://www.amazon.com/"
EDGE_DRIVER_PATH = "./EDGEdriver/msedgedriver.exe"
YOUR_PRODUCT_KEYWORD = "Nleyook"


# ==============================================================

def init_browser():
    print("DEBUG: 正在初始化浏览器...")
    service = Service(executable_path=EDGE_DRIVER_PATH)
    options = webdriver.EdgeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--start-maximized")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    driver = webdriver.Edge(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    print("DEBUG: 浏览器初始化完成！")
    return driver


def set_zip_code(driver, zip_code):
    print(f"正在设置邮编为: {zip_code}")
    try:
        driver.get(AMAZON_URL)
        WebDriverWait(driver, 15).until(EC.presence_of_element_located((By.ID, "twotabsearchtextbox")))
        time.sleep(random.uniform(2, 3))

        zip_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "nav-global-location-popover-link"))
        )
        zip_button.click()
        time.sleep(random.uniform(1, 2))

        zip_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.ID, "GLUXZipUpdateInput"))
        )
        zip_input.clear()
        zip_input.send_keys(zip_code)
        time.sleep(random.uniform(0.5, 1))

        apply_button = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.ID, "GLUXZipUpdate"))
        )
        apply_button.click()
        time.sleep(random.uniform(3, 4))

        try:
            close_button = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable((By.ID, "GLUXConfirmClose"))
            )
            close_button.click()
            time.sleep(random.uniform(1, 2))
        except NoSuchElementException:
            pass
        print(f"邮编{zip_code}设置成功！")
        return True
    except Exception as e:
        print(f"邮编设置失败，错误信息: {e}")
        return False


def get_product_info(driver, page_num):
    print(f"DEBUG: 正在第{page_num}页查找产品...")
    products = driver.find_elements(By.CSS_SELECTOR, "div.s-result-item[data-asin]:not([data-asin=''])")
    print(f"DEBUG: 本页找到{len(products)}个产品")
    total = len(products)
    ad, natural = None, None
    for i, p in enumerate(products, 1):
        try:
            asin = p.get_attribute("data-asin").strip()
            title = p.find_element(By.CSS_SELECTOR, "h2 a span").text.lower()
            if YOUR_PRODUCT_KEYWORD.lower() not in title and asin not in [a.upper() for a in YOUR_ASIN_LIST]:
                continue
            is_ad = "sponsored" in p.text.lower()
            pos = "上" if i <= total / 3 else "中" if i <= total * 2 / 3 else "下"
            color = "无"
            for c in COLOR_KEYWORDS:
                if c in title:
                    color = c.capitalize()
            info = {"页码": page_num, "位置": pos, "颜色": color, "ASIN": asin}
            if is_ad and not ad:
                ad = info
                print(f"DEBUG: 找到广告位产品：{info}")
            if not is_ad and not natural:
                natural = info
                print(f"DEBUG: 找到自然位产品：{info}")
            if ad and natural:
                break
        except Exception as e:
            print(f"DEBUG: 单个产品解析出错：{e}")
            continue
    return ad, natural


def next_page(driver):
    print("DEBUG: 尝试翻页...")
    try:
        next_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "a.s-pagination-next"))
        )
        driver.execute_script("arguments[0].click();", next_btn)
        time.sleep(random.uniform(3, 5))
        print("DEBUG: 翻页成功！")
        return True
    except Exception as e:
        print(f"翻页失败: {e}")
        return False


def search_keyword(driver, kw):
    print(f"DEBUG: 正在搜索关键词：{kw}")
    try:
        driver.get(AMAZON_URL)
        print("DEBUG: 回到亚马逊首页")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "twotabsearchtextbox")))
        time.sleep(random.uniform(1, 2))

        search_box = driver.find_element(By.ID, "twotabsearchtextbox")
        search_box.clear()
        search_box.send_keys(kw)
        print(f"DEBUG: 已输入关键词：{kw}")
        time.sleep(random.uniform(0.5, 1))

        search_btn = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "nav-search-submit-button"))
        )
        search_btn.click()
        print("DEBUG: 点击搜索按钮")
        time.sleep(random.uniform(4, 6))

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div.s-result-item[data-asin]"))
        )
        print("DEBUG: 搜索结果加载完成！")
        return True
    except Exception as e:
        print(f"搜索{kw}失败: {e}")
        return False


def main():
    print("DEBUG: 程序开始运行！")
    print(f"DEBUG: 关键词列表：{KEYWORDS_LIST}")
    print(f"DEBUG: 关键词总数：{len(KEYWORDS_LIST)}")

    if os.path.exists(RESULT_FILE):
        try:
            with open(RESULT_FILE, 'a'):
                pass
        except PermissionError:
            print("❌ 请关闭Excel文件再运行！")
            return

    driver = init_browser()
    set_zip_code(driver, ZIP_CODE)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["关键词", "广告页", "广告位", "广告色", "广告ASIN", "自然页", "自然位", "自然色", "自然ASIN", "备注"])
    for c in ws[1]:
        c.font = Font(bold=True)

    for idx, kw in enumerate(KEYWORDS_LIST, 1):
        print(f"\n===== 第{idx}/{len(KEYWORDS_LIST)}个关键词：{kw} =====")
        ad_res = {"页码": "无", "位置": "无", "颜色": "无", "ASIN": "无"}
        nat_res = {"页码": "无", "位置": "无", "颜色": "无", "ASIN": "无"}
        remark = ""

        try:
            if not search_keyword(driver, kw):
                remark = "搜索页面加载失败"
                print(f"DEBUG: 关键词{kw}搜索失败，跳过")
                continue

            cur_page = 1
            find_ad, find_nat = False, False

            while cur_page <= MAX_PAGE:
                a, n = get_product_info(driver, cur_page)
                if a and not find_ad:
                    ad_res, find_ad = a, True
                if n and not find_nat:
                    nat_res, find_nat = n, True
                if find_ad and find_nat:
                    print("DEBUG: 已找到广告+自然位，提前结束翻页")
                    break
                if cur_page < MAX_PAGE and not next_page(driver):
                    print("DEBUG: 翻页失败，停止查找")
                    break
                cur_page += 1

            remark = "7页未找到" if not find_ad and not find_nat else ""
            print(f"✅ {kw} 完成 | 广告:{ad_res} | 自然:{nat_res}")

        except Exception as e:
            remark = f"错误：{str(e)[:50]}"
            print(f"❌ {kw} 异常：{remark}")

        ws.append([kw, ad_res["页码"], ad_res["位置"], ad_res["颜色"], ad_res["ASIN"],
                   nat_res["页码"], nat_res["位置"], nat_res["颜色"], nat_res["ASIN"], remark])
        wb.save(RESULT_FILE)
        print(f"DEBUG: 已保存{kw}的结果到Excel")
        time.sleep(random.uniform(3, 5))

    print("\n🎉 全部关键词执行完成！")
    driver.quit()
    wb.save(RESULT_FILE)
    print("DEBUG: 程序正常退出！")


if __name__ == "__main__":
    main()