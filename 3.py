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

# ====================== 你的配置（不用动） ======================
KEYWORDS_LIST = ["modest bathing suit for women"]
YOUR_ASIN_LIST = ["B0BRWJJDB4"]
MAX_PAGE = 7
ZIP_CODE = "10010"
RESULT_FILE = "亚马逊产品排名记录.xlsx"
COLOR_KEYWORDS = ["black","white","red","blue","green","yellow","pink","purple","gray","brown","beige","navy"]
AMAZON_URL = "https://www.amazon.com/"
EDGE_DRIVER_PATH = "./EDGEdriver/msedgedriver.exe"
YOUR_PRODUCT_KEYWORD = "Nleyook"
# ==============================================================

def init_browser():
    service = Service(executable_path=EDGE_DRIVER_PATH)
    options = webdriver.EdgeOptions()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_argument("--start-maximized")
    driver = webdriver.Edge(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver


def set_zip_code(driver, zip_code):
    """自动设置亚马逊配送邮编为10010，匹配美国站搜索环境"""
    print(f"正在设置邮编为: {zip_code}")
    try:
        driver.get(AMAZON_URL)
        time.sleep(random.uniform(2, 3))

        # 点击左上角邮编设置按钮
        zip_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "nav-global-location-popover-link"))
        )
        zip_button.click()
        time.sleep(random.uniform(1, 2))

        # 输入邮编
        zip_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "GLUXZipUpdateInput"))
        )
        zip_input.clear()
        zip_input.send_keys(zip_code)
        time.sleep(random.uniform(0.5, 1))

        # 点击应用按钮
        apply_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "GLUXZipUpdate"))
        )
        apply_button.click()
        time.sleep(random.uniform(2, 3))

        # 关闭弹窗（如有）
        try:
            close_button = driver.find_element(By.ID, "GLUXConfirmClose")
            if close_button.is_displayed():
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
    products = driver.find_elements(By.CSS_SELECTOR, "div[data-asin]:not([data-asin=''])")
    total = len(products)
    ad, natural = None, None

    for i, p in enumerate(products, 1):
        try:
            asin = p.get_attribute("data-asin").strip()
            title = p.find_element(By.CSS_SELECTOR, "h2 a span").text.lower()
            if YOUR_PRODUCT_KEYWORD.lower() not in title and asin not in [a.upper() for a in YOUR_ASIN_LIST]:
                continue

            # 判断广告
            is_ad = "sponsored" in p.text.lower()
            pos = "上" if i<=total/3 else "中" if i<=total*2/3 else "下"
            color = "无"
            for c in COLOR_KEYWORDS:
                if c in title: color = c.capitalize()

            info = {"页码":page_num,"位置":pos,"颜色":color,"ASIN":asin}
            if is_ad and not ad: ad = info
            if not is_ad and not natural: natural = info
            if ad and natural: break
        except:
            continue
    return ad, natural

def next_page(driver):
    try:
        driver.execute_script("document.querySelector('a[aria-label*=\"next\"]').click();")
        time.sleep(random.uniform(4,6))
        return True
    except:
        return False

def main():
    # 关闭Excel文件检查
    if os.path.exists(RESULT_FILE):
        try:
            with open(RESULT_FILE, 'a'): pass
        except PermissionError:
            print("❌ 请关闭Excel文件再运行！")
            return

    driver = init_browser()
    set_zip_code(driver,ZIP_CODE)

    # Excel初始化
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["关键词","广告页","广告位","广告色","广告ASIN","自然页","自然位","自然色","自然ASIN","备注"])
    for c in ws[1]: c.font = Font(bold=True)

    for kw in KEYWORDS_LIST:
        print(f"\n搜索：{kw}")
        ad_res = {"页码":"无","位置":"无","颜色":"无","ASIN":"无"}
        nat_res = {"页码":"无","位置":"无","颜色":"无","ASIN":"无"}
        remark = ""

        try:
            # ✅ JS强制点击搜索框（无视遮挡！）
            driver.execute_script(f"document.getElementById('twotabsearchtextbox').value='{kw}';")
            driver.execute_script("document.getElementById('nav-search-submit-button').click();")
            time.sleep(4)

            cur_page = 1
            find_ad, find_nat = False, False
            while cur_page <= MAX_PAGE:
                a, n = get_product_info(driver, cur_page)
                if a and not find_ad: ad_res, find_ad = a, True
                if n and not find_nat: nat_res, find_nat = n, True
                if find_ad and find_nat: break
                if cur_page < MAX_PAGE and not next_page(driver): break
                cur_page += 1

            remark = "7页未找到" if not find_ad and not find_nat else ""
        except Exception as e:
            remark = f"错误：{str(e)[:30]}"

        ws.append([kw, ad_res["页码"],ad_res["位置"],ad_res["颜色"],ad_res["ASIN"],
                   nat_res["页码"],nat_res["位置"],nat_res["颜色"],nat_res["ASIN"], remark])
        wb.save(RESULT_FILE)
        time.sleep(3)

    print("\n🎉 全部完成！")
    driver.quit()
    wb.save(RESULT_FILE)

if __name__ == "__main__":
    main()