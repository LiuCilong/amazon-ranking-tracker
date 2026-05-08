from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.edge.service import Service
from selenium.common.exceptions import NoSuchElementException, TimeoutException
import openpyxl
import time
import random
from openpyxl.styles import Font

# ====================== 【唯一需要你修改的区域，直接在这里增删改即可】 ======================
# 1. 要搜索的关键词列表，每行一个关键词，用英文双引号括起来，英文逗号分隔
KEYWORDS_LIST = ["modest bathing suit for women","modest bathing suit for women"]
# 2. 你的产品子ASIN列表，把要找的所有子ASIN都放这里（子ASIN唯一对应一个颜色，精准匹配）
YOUR_ASIN_LIST = [
   "B0BRWJJDB4"
]

# 3. 最多搜索页数，固定为7页，不用改
MAX_PAGE = 7

# 4. 美国邮编，固定10010，不用改
ZIP_CODE = "10010"

# 5. 结果保存的Excel文件名，可自定义
RESULT_FILE = "亚马逊产品排名记录1.xlsx"

# 6. 产品颜色关键词，补充你家产品的所有颜色，用于精准提取颜色
COLOR_KEYWORDS = [
    "black", "white", "red", "blue", "green", "yellow", "pink", "purple",
    "gray", "brown", "beige", "navy", "khaki", "orange", "burgundy", "olive"
]
# 7. 亚马逊美国站地址，不用改
AMAZON_URL = "https://www.amazon.com/"

# 8. Edge驱动路径，和你的文件夹结构完全匹配，不用改！除非你改了文件夹名字
EDGE_DRIVER_PATH = "./EDGEdriver/msedgedriver.exe"

YOUR_PRODUCT_KEYWORD = "Nleyook"
# ========================================================================================

def init_browser():
    """初始化Edge浏览器，使用你项目目录下的本地驱动，自带反爬屏蔽"""
    # 加载你项目里的Edge驱动
    service = Service(executable_path=EDGE_DRIVER_PATH)
    options = webdriver.EdgeOptions()
    # 屏蔽亚马逊自动化检测，避免触发人机验证
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    # 窗口最大化，避免元素加载不全
    options.add_argument("--start-maximized")
    # 初始化浏览器
    driver = webdriver.Edge(service=service, options=options)
    # 移除自动化标识
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
    """提取当前页的产品信息，返回(第一个你的广告位, 第一个你的自然位)，无则返回None"""
    # 等待产品加载完成
    try:
        WebDriverWait(driver, 10).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div[data-asin]"))
        )
    except TimeoutException:
        print(f"第{page_num}页未加载到产品")
        return None, None

    # 获取所有有效产品（排除空ASIN）
    all_products = driver.find_elements(By.CSS_SELECTOR, "div[data-asin]:not([data-asin=''])")
    total_products = len(all_products)
    if total_products == 0:
        print(f"第{page_num}页无有效产品")
        return None, None

    print(f"第{page_num}页共找到{total_products}个产品")
    first_ad, first_natural = None, None

    # 遍历当前页所有产品
    for index, product in enumerate(all_products, start=1):
        try:
            # 获取产品ASIN和标题
            asin = product.get_attribute("data-asin")
            title = product.find_element(By.CSS_SELECTOR, "h2 a span").text.strip()
            # 判断是否为你的产品（标题包含专属词，不区分大小写）
            is_your_product = YOUR_PRODUCT_KEYWORD.lower() in title.lower()
            if not is_your_product:
                continue

            # 判断是广告位还是自然位
            is_sponsored = False
            try:
                # 匹配亚马逊2种广告标识格式
                product.find_element(By.CSS_SELECTOR, "span[data-component-type='sp-sponsored-label-text']")
                is_sponsored = True
            except NoSuchElementException:
                try:
                    product.find_element(By.XPATH, ".//span[text()='Sponsored']")
                    is_sponsored = True
                except NoSuchElementException:
                    pass

            # 计算位置：上/中/下（按当前页产品总数均分3段）
            if index <= total_products / 3:
                position = "上"
            elif index <= total_products * 2 / 3:
                position = "中"
            else:
                position = "下"

            # 提取产品颜色（匹配你设置的颜色关键词）
            color = "无"
            title_lower = title.lower()
            for c in COLOR_KEYWORDS:
                if c in title_lower:
                    color = c.capitalize()
                    break

            # 组装产品信息
            product_info = {
                "页码": page_num,
                "位置": position,
                "颜色": color,
                "ASIN": asin,
                "标题": title
            }

            # 只记录第一个广告位和第一个自然位
            if is_sponsored and first_ad is None:
                first_ad = product_info
                print(f"找到第一个广告位：第{page_num}页{position}部，颜色{color}")
            elif not is_sponsored and first_natural is None:
                first_natural = product_info
                print(f"找到第一个自然位：第{page_num}页{position}部，颜色{color}")

            # 两个都找到就提前结束遍历，提升效率
            if first_ad is not None and first_natural is not None:
                break

        except Exception:
            # 单个产品提取失败不中断程序，直接跳过
            continue

    return first_ad, first_natural

def go_to_next_page(driver, current_page):
    """自动翻到下一页，成功返回True，无下一页返回False"""
    try:
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[text()='Next']"))
        )
        next_button.click()
        print(f"成功跳转到第{current_page + 1}页")
        time.sleep(random.uniform(3, 5))  # 翻页后增加等待，规避反爬
        return True
    except (NoSuchElementException, TimeoutException):
        print(f"无下一页，当前为第{current_page}页")
        return False

def main():
    # 1. 加载关键词列表（直接用代码里的列表，无需外部文件）
    keywords = KEYWORDS_LIST
    if not keywords:
        print("关键词列表为空，请在代码顶部的KEYWORDS_LIST里添加关键词！")
        return
    print(f"成功读取到{len(keywords)}个关键词")

    # 2. 初始化Excel结果文件
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "排名记录"
    # 设置Excel表头
    headers = [
        "关键词", "第一个广告位页码", "第一个广告位位置", "第一个广告位颜色", "广告位ASIN",
        "第一个自然位页码", "第一个自然位位置", "第一个自然位颜色", "自然位ASIN", "备注"
    ]
    ws.append(headers)
    # 表头加粗
    for cell in ws[1]:
        cell.font = Font(bold=True)

    # 3. 初始化浏览器，设置邮编
    driver = init_browser()
    zip_success = set_zip_code(driver, ZIP_CODE)
    if not zip_success:
        print("邮编设置失败，程序退出")
        driver.quit()
        return

    # 4. 逐个处理关键词
    for keyword in keywords:
        print(f"\n====================== 正在处理关键词: {keyword} ======================")
        # 初始化结果为无
        ad_result = {"页码": "无", "位置": "无", "颜色": "无", "ASIN": "无"}
        natural_result = {"页码": "无", "位置": "无", "颜色": "无", "ASIN": "无"}
        remark = ""

        try:
            # 搜索关键词
            search_box = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "twotabsearchtextbox"))
            )
            search_box.clear()
            search_box.send_keys(keyword)
            time.sleep(random.uniform(0.5, 1))
            # 点击搜索按钮
            search_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "nav-search-submit-button"))
            )
            search_button.click()
            time.sleep(random.uniform(3, 5))

            # 遍历最多7页搜索结果
            current_page = 1
            found_ad, found_natural = False, False
            while current_page <= MAX_PAGE:
                first_ad, first_natural = get_product_info(driver, current_page)

                # 记录第一个匹配到的广告位/自然位
                if first_ad is not None and not found_ad:
                    ad_result = first_ad
                    found_ad = True
                if first_natural is not None and not found_natural:
                    natural_result = first_natural
                    found_natural = True

                # 两个都找到就提前结束翻页，节省时间
                if found_ad and found_natural:
                    print("该关键词的广告位和自然位均已找到，提前结束翻页")
                    break

                # 翻页操作
                if current_page < MAX_PAGE:
                    has_next = go_to_next_page(driver, current_page)
                    if not has_next:
                        remark = f"仅找到{current_page}页，不足7页"
                        break
                current_page += 1

            # 7页未找到的备注
            if not found_ad and not found_natural:
                remark = "7页内未找到该关键词对应的产品"

        except Exception as e:
            print(f"处理关键词【{keyword}】出错，错误信息: {e}")
            remark = f"处理出错: {str(e)}"

        # 写入Excel一行数据
        row_data = [
            keyword,
            ad_result["页码"], ad_result["位置"], ad_result["颜色"], ad_result["ASIN"],
            natural_result["页码"], natural_result["位置"], natural_result["颜色"], natural_result["ASIN"],
            remark
        ]
        ws.append(row_data)
        # 每处理一个关键词就保存一次，避免程序崩溃丢失数据
        wb.save(RESULT_FILE)
        # 关键词之间加随机等待，规避亚马逊反爬
        time.sleep(random.uniform(2, 4))

    # 程序结束
    # 程序全部执行完成
    print(f"\n====================== 全部关键词处理完成！结果已保存到【{RESULT_FILE}】 ======================")
    driver.quit()
    wb.save(RESULT_FILE)
    wb.close()


if __name__ == "__main__":
    main()







