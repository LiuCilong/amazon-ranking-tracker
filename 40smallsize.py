from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
import openpyxl
import time
from selenium.common.exceptions import NoSuchElementException
import random
from openpyxl.styles import Font
import os

# ====================== 你的配置（不用动） ======================
# 可以替换为你图中的长尾精准词
# KEYWORDS_LIST = ["plus size tankini for women", "plus size swimdress with shorts"]
KEYWORDS_LIST = [
    "swim romper for women",
    "romper swimsuit",
    "romper bathing suit for women",
    "womens romper swimsuit",
    "one piece swimsuit women with shorts",
    "one piece swimsuit shorts",
    "bathing suit romper for women",
    "romper bathing suit",
    "romper swimsuits for women",
    "swimsuit romper",
    "womens swim romper bathing suit",
    "one piece shorts swimsuit women",
    "one piece shorts romper",
    "black one piece swimsuits for women",
    "shorts jumpsuit for women"
]



YOUR_ASIN_LIST = [
    "B0GJSQR5XF", "B0GJT91RX3", "B0GJSYD5FF", "B0GJSTHG41", "B0GJT39VC8",
    "B0GJT2QK5Q", "B0GJSYJZKK", "B0GJT62KWP", "B0GJSX8Y5P", "B0GJSPYB26",
    "B0GJT533PQ", "B0GJT55JN7", "B0GJSX7C7L", "B0GJSXK2VN", "B0GJSVYB8N",
    "B0GJSPS2FD", "B0GJSXFKTN", "B0GJSS23F8", "B0GJSVX22D", "B0GJSWRBPR",
    "B0GJSK6G4K", "B0GJT1BT8S", "B0GJT4LPGY", "B0GJT4K3J1", "B0GJSM62KF"
]

MAX_PAGE = 7
ZIP_CODE = "10010"
RESULT_FILE = "亚马逊产品排名记录40smallszie0505.xlsx"
AMAZON_URL = "https://www.amazon.com/"
EDGE_DRIVER_PATH = "./EDGEdriver/msedgedriver.exe"
# ==============================================================

TARGET_ASINS = [a.upper() for a in YOUR_ASIN_LIST]


def init_browser():
    service = Service(executable_path=EDGE_DRIVER_PATH)
    options = webdriver.EdgeOptions()
    # 👇 加这一行，开启Edge无痕（InPrivate）模式
    options.add_argument("--inprivate")
    # 你原来的反爬参数（保留不变）
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option('useAutomationExtension', False)
    options.add_argument("--start-maximized")
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-popup-blocking")
    driver = webdriver.Edge(service=service, options=options)
    driver.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    return driver

def set_zip_code(driver, zip_code):
    """自动设置亚马逊配送邮编"""
    print(f"正在设置邮编为: {zip_code}")
    try:
        driver.get(AMAZON_URL)
        time.sleep(random.uniform(2, 3))

        zip_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "nav-global-location-popover-link"))
        )
        zip_button.click()
        time.sleep(random.uniform(1, 2))

        zip_input = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "GLUXZipUpdateInput"))
        )
        zip_input.clear()
        zip_input.send_keys(zip_code)
        time.sleep(random.uniform(0.5, 1))

        apply_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "GLUXZipUpdate"))
        )
        apply_button.click()
        time.sleep(random.uniform(2, 3))

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
        print(f"邮编设置失败: {e}")
        return False


def get_product_info(driver, page_num):
    # 限定抓取范围为标准搜索网格
    # products = driver.find_elements(By.CSS_SELECTOR, "div.s-main-slot div[data-component-type='s-search-result']")
    products = driver.find_elements(By.CSS_SELECTOR, "div[data-asin]:not([data-asin=''])")
    total = len(products)
    ad, natural = None, None

    for i, p in enumerate(products, 1):
        try:
            asin = p.get_attribute("data-asin").strip().upper()

            if asin not in TARGET_ASINS:
                continue

            # ==========================================
            # 【全新广告判定逻辑：多维特征探测】
            # ==========================================
            is_ad = False

            # 策略 1：检查外层容器类名 (针对顶部大横幅或特殊排版)
            container_classes = p.get_attribute("class") or ""
            if "AdHolder" in container_classes or "s-sponsored-search-result" in container_classes:
                is_ad = True

            # 策略 2：向下寻找核心的 Sponsored 标签元素 (结合你的实测发现)
            if not is_ad:
                try:
                    # 将你发现的两种形态全部写入 CSS 选择器进行联合查找
                    # 只要命中其中任意一个，即判定为广告位
                    ad_badges = p.find_elements(By.CSS_SELECTOR,
                                                ".puis-sponsored-label-info-icon, "  # 你发现的形态 2 徽章
                                                ".puis-sponsored-label-text, "  # 你发现的形态 2 文本
                                                "span[aria-label*='Sponsored information'], "  # 你发现的形态 1 (模糊匹配增强稳定性)
                                                ".s-sponsored-label-info-icon"  # 兜底的老版本类名
                                                )
                    if len(ad_badges) > 0:
                        is_ad = True
                except Exception:
                    pass
            # ==========================================

            pos = f"{i}/{total}"

            info = {"页码": page_num, "位置": pos, "ASIN": asin}

            if is_ad and not ad:
                ad = info
            if not is_ad and not natural:
                natural = info
            if ad and natural:
                break
        except Exception as e:  # 把错误信息存到 e 里
            print(f"⚠️ 商品加载失败，已自动跳过！错误原因：{str(e)[:50]}")  # 加这行提示
            continue
    return ad, natural


def next_page(driver):
    try:
        next_btn = driver.find_elements(By.CSS_SELECTOR, ".s-pagination-next")
        if next_btn:
            if "s-pagination-disabled" in next_btn[0].get_attribute("class"):
                return False
            driver.execute_script("arguments[0].click();", next_btn[0])
            time.sleep(random.uniform(4, 6))
            return True
        return False
    except Exception:
        return False


def main():
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
    ws.append(["关键词", "广告页", "广告位", "广告ASIN", "自然页", "自然位", "自然ASIN", "备注"])
    for c in ws[1]: c.font = Font(bold=True)

    for kw in KEYWORDS_LIST:
        print(f"\n搜索：{kw}")
        ad_res = {"页码": "无", "位置": "无", "ASIN": "无"}
        nat_res = {"页码": "无", "位置": "无", "ASIN": "无"}
        remark = ""

        try:
            driver.get(AMAZON_URL)
            time.sleep(random.uniform(2, 3))

            search_box = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.ID, "twotabsearchtextbox"))
            )
            search_box.clear()
            search_box.send_keys(kw)

            driver.execute_script("document.getElementById('nav-search-submit-button').click();")
            time.sleep(4)

            cur_page = 1
            find_ad, find_nat = False, False

            while cur_page <= MAX_PAGE:
                print(f"正在扫描第 {cur_page} 页...")
                a, n = get_product_info(driver, cur_page)

                if a and not find_ad:
                    ad_res, find_ad = a, True
                if n and not find_nat:
                    nat_res, find_nat = n, True
                if find_ad and find_nat:
                    break
                if cur_page < MAX_PAGE and not next_page(driver):
                    break

                cur_page += 1

            remark = "7页未找全" if not (find_ad and find_nat) else ""

        except Exception as e:
            remark = f"错误：{str(e)[:30]}"

        ws.append([kw, ad_res["页码"], ad_res["位置"], ad_res["ASIN"],
                   nat_res["页码"], nat_res["位置"], nat_res["ASIN"], remark])
        wb.save(RESULT_FILE)
        time.sleep(3)

    print("\n🎉 全部完成！")
    driver.quit()


if __name__ == "__main__":
    main()