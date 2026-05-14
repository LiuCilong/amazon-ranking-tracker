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

# ====================== 配置 ======================
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
RESULT_FILE = "亚马逊产品排名记录60smallszie0512.xlsx"
AMAZON_URL = "https://www.amazon.com/"
EDGE_DRIVER_PATH = "./EDGEdriver/msedgedriver.exe"
# ==============================================================

TARGET_ASINS = [a.upper() for a in YOUR_ASIN_LIST]


def init_browser():
    service = Service(executable_path=EDGE_DRIVER_PATH)
    options = webdriver.EdgeOptions()
    options.add_argument("--inprivate")
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


# 搜索结果容器选择器（自然 + 广告），多处复用
RESULT_CONTAINER_SEL = (
    "div.s-main-slot div[data-component-type='s-search-result'], "
    "div.s-main-slot div[data-component-type='sp-sponsored-result']"
)


def get_product_info(driver, page_num):
    # 【新增】同时抓取自然结果和广告容器，避免广告容器被排除在扫描范围之外
    products = driver.find_elements(By.CSS_SELECTOR, RESULT_CONTAINER_SEL)
    total = len(products)
    ad, natural = None, None

    for i, p in enumerate(products, 1):
        try:
            asin = p.get_attribute("data-asin").strip().upper()

            if asin not in TARGET_ASINS:
                continue

            is_ad = False

            # ============================================================
            # 【原有判断1】检查外层容器 class（精确匹配，保留）
            # ============================================================
            container_classes = p.get_attribute("class") or ""
            if "AdHolder" in container_classes or "s-sponsored-search-result" in container_classes:
                is_ad = True

            # ============================================================
            # 【原有判断2】精确类名 + aria-label（保留）
            # ============================================================
            if not is_ad:
                try:
                    ad_badges = p.find_elements(By.CSS_SELECTOR,
                                                ".puis-sponsored-label-info-icon, "
                                                ".puis-sponsored-label-text, "
                                                "span[aria-label*='Sponsored information'], "
                                                ".s-sponsored-label-info-icon"
                                                )
                    if len(ad_badges) > 0:
                        is_ad = True
                except Exception:
                    pass

            # ============================================================
            # 【新增判断3】data-component-type 属性（最稳定）
            # ============================================================
            if not is_ad:
                comp_type = p.get_attribute("data-component-type") or ""
                if "sp-sponsored" in comp_type or comp_type == "ad":
                    is_ad = True

            # ============================================================
            # 【新增判断4】class 通配符匹配（不怕 Amazon 改名）
            # ============================================================
            if not is_ad:
                if any(kw in container_classes for kw in
                       ["AdHolder", "sponsored-search", "sponsored-result"]):
                    is_ad = True

            # ============================================================
            # 【新增判断5】内部元素通配符匹配
            # ============================================================
            if not is_ad:
                try:
                    ad_badges = p.find_elements(By.CSS_SELECTOR,
                                                "[class*='sponsored-label'], "
                                                "[class*='puis-sponsored'], "
                                                "[aria-label*='Sponsored'], "
                                                "[data-component-type*='sponsored']"
                                                )
                    if len(ad_badges) > 0:
                        is_ad = True
                except Exception:
                    pass

            # ============================================================
            # 【新增判断6】终极兜底——文字内容包含 "Sponsored"
            # 只看前500字符，避免匹配到评论/描述中的 sponsored 字样
            # ============================================================
            if not is_ad:
                try:
                    text = p.text.lower()
                    if "sponsored" in text[:500]:
                        is_ad = True
                except Exception:
                    pass

            pos = f"{i}/{total}"

            info = {"页码": page_num, "位置": pos, "ASIN": asin}

            if is_ad and not ad:
                ad = info
            if not is_ad and not natural:
                natural = info
            if ad and natural:
                break
        except Exception as e:
            print(f"商品加载失败，已自动跳过！错误原因：{str(e)[:50]}")
            continue
    return ad, natural


# 【修复】翻页后用 WebDriverWait 等待新页面加载，而非固定 sleep
def next_page(driver):
    try:
        next_btn = driver.find_elements(By.CSS_SELECTOR, ".s-pagination-next")
        if not next_btn:
            return False
        cls = next_btn[0].get_attribute("class") or ""
        if "s-pagination-disabled" in cls:
            return False
        driver.execute_script("arguments[0].click();", next_btn[0])
        # 等待搜索结果容器出现，最多15秒
        try:
            WebDriverWait(driver, 15).until(
                EC.presence_of_element_located(
                    (By.CSS_SELECTOR, RESULT_CONTAINER_SEL)
                )
            )
        except Exception:
            time.sleep(3)
        return True
    except Exception:
        return False


def main():
    if os.path.exists(RESULT_FILE):
        try:
            with open(RESULT_FILE, 'a'):
                pass
        except PermissionError:
            print("请关闭Excel文件再运行！")
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

            # 【修复】等待搜索结果加载再开始扫描
            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_element_located(
                        (By.CSS_SELECTOR, RESULT_CONTAINER_SEL)
                    )
                )
            except Exception:
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

    print("\n全部完成！")
    driver.quit()


if __name__ == "__main__":
    main()
