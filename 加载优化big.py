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
KEYWORDS_LIST = ["plus size bathing suit","plus size bathing suits for women","bathing suits for women","women bathing suits",
                 "womens bathing suits tankinis","tummy control bathing suits for women","2 piece bathing suits for women",
                 "two piece bathing suits for women","womens bathing suits tummy control","bathing suit for women tummy control",
                 "modest bathing suit for women","black bathing suits for women","becca bathing suit","bathing suit bottoms for women"
                ,"plus size swimsuit for women","plus size bathing suit for women","plus size tankini for women","plus size swimdress with shorts"]



YOUR_ASIN_LIST = [
    "B0GFN8MLZB", "B0CSFYXM1C", "B0BRWJS2GR", "B0CTPZ2W5Q", "B0GFNFBXWV",
    "B0DX1RF8JS", "B0DX1G55FK", "B0CTQ5Y6P5", "B0C27V1ZTX", "B0DX1JN6D6",
    "B0BRWJF3GK", "B0CSFZ1TQ2", "B0BRWH6THR", "B0DX1BQCZS", "B0C4YSMBXV",
    "B0DNKLCRL1", "B0BW5ZJ6LT", "B0DYN9QX4K", "B0DWZV1NT1", "B0BZVXH3WQ",
    "B0F7QXK712", "B0C4YS63X8", "B0F7QYD1VS", "B0CTLSF83V", "B0DWZZ4BLK",
    "B0GFNFTR3W", "B0BRWJXBFN", "B0C4YRPBGY", "B0BRWL6N2Q", "B0D7GXBN2B",
    "B0GFN127C6", "B0DWZZT29C", "B0BZVZ3HBJ", "B0GFNFQNVC", "B0DT6Q9M2D",
    "B0DRFPPLZ1", "B0GFNBVNYC", "B0GFN2KKDW", "B0DX13BDKR", "B0CSFZ7CRK",
    "B0BRWJPQ74", "B0DT6YWF6H", "B0CSG1R35L", "B0BRWL12KL", "B0C4YNR3L5",
    "B0C4YRC4JB", "B0CSG1VF8Z", "B0BZW1KCWK", "B0DWZYBQRK", "B0CSFYJZD3",
    "B0DNMGSJ78", "B0C4YQ42VN", "B0BRWKSJZS", "B0CSG18VB9", "B0DX1LLKV5",
    "B0GFNHLNV8", "B0BRWK6TLR", "B0CTQCCHZ4", "B0BRWK8QHS", "B0DNKKLDD4",
    "B0DX1DVP4H", "B0F7QRVCQ5", "B0C4YRQS47", "B0F7QXPY7H", "B0BRWJJDB4",
    "B0BRWJMM27", "B0DX1DQ6LF", "B0F7QTXMNW", "B0C4YN3B9L", "B0C4YNR42R",
    "B0GFN29H6F", "B0DT6QP38H", "B0DT6R6NHL", "B0C4YNFW3C", "B0BRWJD7J5",
    "B0DT73NCCJ", "B0DRFMGFYQ", "B0GFN14H8C", "B0DT6QJP4S", "B0DNKK47GV",
    "B0GFNBWL54", "B0DX1646BZ", "B0GFN7KGCW", "B0DNKK47H1", "B0BRWJLJNV",
    "B0CTQ9YQCS", "B0D7GQVQZQ", "B0DX13NJ2K", "B0CTQ4QXHJ", "B0DT6QNJDZ",
    "B0GFMXM5W2", "B0DNKHLTWR", "B0F7QSFQ5J", "B0D7H5STC8", "B0DT6R88VY",
    "B0DWZWZ61R", "B0BRWJFL6Y", "B0FG7VJ5YM", "B0GFN72NC2", "B0DT6S2BVR",
    "B0F7QY8ZWX", "B0DT6T2NZR", "B0BRWJFL6T", "B0DWZQN9YH", "B0GFN8DWKL",
    "B0DT6JZ3GS", "B0DRFS4BYJ", "B0GFN5SCK5", "B0GFNB1MCL", "B0BRWL38V5",
    "B0DNKJN59K", "B0C4YQCMLB", "B0DT6SWP2Z", "B0BRWJSYXJ", "B0DX13NRXP",
    "B0CSG172RK", "B0BZVX5M48", "B0GFN7GVTN", "B0DWZNHPT8", "B0DX166VSF",
    "B0GFN27W34", "B0DT6T6YBP", "B0GFN7NW8C", "B0BRWK4JSM", "B0CTQ1RBV9",
    "B0GFN85LZ5", "B0DWZY6HQ7"
]

MAX_PAGE = 7
ZIP_CODE = "10010"
RESULT_FILE = "亚马逊产品排名精确记录bigsize.xlsx"
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
    # ==========================================
    # 【新增修复：模拟真人分段滑动页面，强制触发懒加载】
    # ==========================================
    for scroll in range(1, 5):
        # 每次往下拉 1/4 的屏幕高度
        driver.execute_script(f"window.scrollTo(0, document.body.scrollHeight * ({scroll}/4));")
        time.sleep(random.uniform(0.8, 1.2))  # 停顿一下，等待底部图片和 DOM 渲染
    # ==========================================

    # 限定抓取范围为标准搜索网格
    products = driver.find_elements(By.CSS_SELECTOR, "div.s-main-slot div[data-component-type='s-search-result']")
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

            # 策略 1：检查外层容器类名
            container_classes = p.get_attribute("class") or ""
            if "AdHolder" in container_classes or "s-sponsored-search-result" in container_classes:
                is_ad = True

            # 策略 2：向下寻找核心的 Sponsored 标签元素
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
            # ==========================================

            pos = f"{i}/{total}"

            info = {"页码": page_num, "位置": pos, "ASIN": asin}

            if is_ad and not ad:
                ad = info
            if not is_ad and not natural:
                natural = info
            if ad and natural:
                break
        except Exception as e:
            print(f"⚠️ 商品加载失败，已自动跳过！错误原因：{str(e)[:50]}")
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