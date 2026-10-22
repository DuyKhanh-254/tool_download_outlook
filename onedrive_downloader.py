#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
OneDrive & SharePoint PDF Downloader (CLI)
=========================================
Công cụ tự động tải file PDF từ SharePoint và OneDrive trong trường hợp bị khóa
tính năng tải xuống (download disabled / restricted view).

Hỗ trợ trích xuất direct download URL từ API, trích xuất raw canvas qua JavaScript,
hoặc chụp screenshot từng trang ở độ phân giải cao và ghép thành file PDF hoàn chỉnh.
"""

import argparse
import base64
import io
import os
import re
import shutil
import sys
import tempfile
import time
from typing import List, Optional, Tuple
from urllib.parse import parse_qs, unquote, urlparse

# Đảm bảo console Windows in ký tự tiếng Việt UTF-8 không bị lỗi charmap
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

import requests
from PIL import Image
from PyPDF2 import PdfReader
from selenium import webdriver
from selenium.common.exceptions import (
    NoSuchElementException,
    SessionNotCreatedException,
    TimeoutException,
    WebDriverException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait


def validate_url(url: str) -> bool:
    """
    Kiểm tra tính hợp lệ của URL SharePoint hoặc OneDrive.

    Args:
        url: Chuỗi URL cần kiểm tra.

    Returns:
        True nếu URL khớp với định dạng SharePoint/OneDrive, ngược lại False.
    """
    if not url:
        return False
    pattern = r"^https?://([a-zA-Z0-9_-]+)\.(sharepoint\.com|onedrive\.live\.com)/.*"
    return bool(re.match(pattern, url.strip(), re.IGNORECASE))


def extract_filename_from_url(url: str, driver: Optional[webdriver.Chrome] = None) -> str:
    """
    Trích xuất tên file từ tham số URL hoặc tiêu đề trang web. Luôn đảm bảo đuôi là .pdf.

    Args:
        url: URL của tài liệu.
        driver: Selenium WebDriver (nếu có để đọc document.title).

    Returns:
        Tên file PDF chuẩn hóa.
    """
    filename = "document.pdf"
    try:
        parsed = urlparse(url)
        params = parse_qs(parsed.query)

        # Kiểm tra tham số 'id' thường dùng trong SharePoint
        if "id" in params and params["id"]:
            raw_path = unquote(params["id"][0])
            name = os.path.basename(raw_path)
            if name:
                filename = name

        # Nếu chưa tìm được, đọc title từ driver
        if filename == "document.pdf" and driver:
            title = driver.title
            if title:
                # Loại bỏ hậu tố SharePoint/OneDrive
                title = re.sub(r"\s*-\s*(SharePoint|OneDrive|Word|PowerPoint).*$", "", title, flags=re.I).strip()
                if title:
                    filename = title
    except Exception:
        pass

    # Chuẩn hóa đuôi file luôn là .pdf
    if not filename.lower().endswith(".pdf"):
        # Đổi các đuôi khác (docx, pptx) thành pdf
        filename = re.sub(r"\.[a-zA-Z0-9]+$", "", filename) + ".pdf"

    # Xóa các ký tự không hợp lệ cho tên file trên Windows
    filename = re.sub(r'[\\/*?:"<>|]', "_", filename)
    return filename


def init_driver(
    browser_type: str = "chrome",
    profile_dir: Optional[str] = None,
    headless: bool = False,
    verbose: bool = False,
) -> webdriver.Remote:
    """
    Khởi tạo Selenium WebDriver (Chrome hoặc Firefox) với cấu hình tối ưu.

    Args:
        browser_type: Loại trình duyệt ('chrome' hoặc 'firefox').
        profile_dir: Đường dẫn đến thư mục Profile để giữ phiên đăng nhập.
        headless: Chạy ẩn không hiển thị cửa sổ.
        verbose: In log chi tiết.

    Returns:
        Đối tượng WebDriver.
    """
    browser_type = browser_type.lower()

    if browser_type == "chrome":
        options = webdriver.ChromeOptions()

        # Cấu hình độ phân giải và High DPI (2x) để hình ảnh tài liệu sắc nét
        options.add_argument("--window-size=1920,1080")
        options.add_argument("--force-device-scale-factor=2")
        options.add_argument("--high-dpi-support=1")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--disable-blink-features=AutomationControlled")
        options.add_experimental_option("excludeSwitches", ["enable-automation"])
        options.add_experimental_option("useAutomationExtension", False)

        if profile_dir:
            if os.path.exists(profile_dir):
                if verbose:
                    print(f"[+] Sử dụng Chrome Profile: {profile_dir}")
                options.add_argument(f"--user-data-dir={profile_dir}")
            else:
                print(f"[!] Cảnh báo: Thư mục profile '{profile_dir}' không tồn tại. Tiếp tục với profile mặc định.")

        if headless:
            options.add_argument("--headless=new")

        try:
            driver = webdriver.Chrome(options=options)
            return driver
        except SessionNotCreatedException as e:
            if "user data directory is already in use" in str(e).lower():
                print("\n[!] LỖI XUNG ĐỘT PROFILE:")
                print("    Thư mục Chrome User Data đang bị khóa do bạn đang mở một cửa sổ Chrome khác.")
                print("    Giải pháp: Vui lòng đóng tất cả cửa sổ Chrome hiện tại rồi chạy lại lệnh,")
                print("    hoặc không dùng tùy chọn --profile-dir và đăng nhập trực tiếp trên cửa sổ tool mở ra.\n")
            raise

    elif browser_type == "firefox":
        options = webdriver.FirefoxOptions()
        options.add_argument("--width=1920")
        options.add_argument("--height=1080")
        if headless:
            options.add_argument("-headless")
        if profile_dir and os.path.exists(profile_dir):
            if verbose:
                print(f"[+] Sử dụng Firefox Profile: {profile_dir}")
            options.profile = profile_dir

        return webdriver.Firefox(options=options)

    else:
        raise ValueError(f"Trình duyệt '{browser_type}' không được hỗ trợ. Chỉ hỗ trợ 'chrome' hoặc 'firefox'.")


def extract_direct_url(session: requests.Session, url: str) -> Optional[str]:
    """
    Phương pháp 2: Thử trích xuất hoặc kiểm tra Direct Download URL từ SharePoint API.

    Args:
        session: Requests session có chứa cookie xác thực.
        url: URL SharePoint xem trước.

    Returns:
        URL tải trực tiếp nếu tìm thấy và hoạt động, ngược lại None.
    """
    try:
        # Thử thêm download=1 vào URL
        test_url = url
        if "download=1" not in test_url:
            separator = "&" if "?" in test_url else "?"
            test_url = f"{test_url}{separator}download=1"

        resp = session.head(test_url, allow_redirects=True, timeout=10)
        content_type = resp.headers.get("content-type", "").lower()
        if resp.status_code == 200 and ("pdf" in content_type or "octet-stream" in content_type):
            return test_url
    except Exception:
        pass
    return None


def merge_images_to_pdf(image_paths: List[str], output_path: str) -> str:
    """
    Phương pháp ghép các ảnh chụp từng trang thành một file PDF duy nhất bằng Pillow.

    Args:
        image_paths: Danh sách đường dẫn ảnh các trang đã chụp.
        output_path: Đường dẫn file PDF đầu ra.

    Returns:
        Đường dẫn file PDF đã lưu.
    """
    if not image_paths:
        raise ValueError("Danh sách ảnh rỗng, không thể tạo PDF.")

    pil_images = []
    for p in image_paths:
        try:
            img = Image.open(p)
            # Chuyển đổi về không gian màu RGB (tránh lỗi khi xuất PDF từ RGBA)
            if img.mode != "RGB":
                img = img.convert("RGB")
            pil_images.append(img)
        except Exception as e:
            print(f"[!] Bỏ qua ảnh lỗi {p}: {e}")

    if not pil_images:
        raise RuntimeError("Không có ảnh hợp lệ để ghép thành PDF.")

    # Đảm bảo thư mục đích tồn tại
    out_dir = os.path.dirname(output_path)
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)

    # Lưu file PDF đầu ra
    first_image = pil_images[0]
    first_image.save(
        output_path,
        save_all=True,
        append_images=pil_images[1:],
        resolution=150.0,
        quality=95,
        format="PDF",
    )

    # Đóng các ảnh
    for img in pil_images:
        img.close()

    return output_path


def validate_pdf(output_path: str) -> Tuple[bool, int, int]:
    """
    Kiểm tra file PDF đầu ra: số trang, kích thước file và tính toàn vẹn bằng PyPDF2.

    Args:
        output_path: Đường dẫn file PDF cần kiểm tra.

    Returns:
        Tuple (is_valid, page_count, file_size_bytes).
    """
    if not os.path.exists(output_path):
        return False, 0, 0

    size = os.path.getsize(output_path)
    if size == 0:
        return False, 0, 0

    try:
        reader = PdfReader(output_path)
        page_count = len(reader.pages)
        if page_count > 0:
            return True, page_count, size
    except Exception as e:
        print(f"[!] Lỗi khi đọc kiểm tra file PDF: {e}")

    return False, 0, size


def capture_pdf_via_screenshots(
    driver: webdriver.Remote,
    url: str,
    temp_dir: str,
    page_delay: float = 2.0,
    start_page: int = 1,
    total_pages_override: Optional[int] = None,
    verbose: bool = False,
) -> List[str]:
    """
    Phương pháp 1 & 3: Mở file PDF trong preview viewer, cuộn qua từng trang,
    trích xuất dữ liệu từ thẻ <canvas> (Method 3) hoặc chụp screenshot element trang (Method 1).

    Args:
        driver: Selenium WebDriver đã xác thực.
        url: URL của file cần tải.
        temp_dir: Thư mục lưu ảnh tạm thời.
        page_delay: Thời gian chờ giữa các trang (giây).
        start_page: Trang bắt đầu quét (mặc định 1).
        total_pages_override: Số trang người dùng tự chỉ định (nếu có).
        verbose: In log chi tiết.

    Returns:
        Danh sách đường dẫn các file ảnh đã chụp.
    """
    print(f"[*] Đang điều hướng đến URL: {url}")
    driver.get(url)

    # Chờ trang nạp và xác nhận xem trước
    time.sleep(3)

    # Kiểm tra nếu trang bị chuyển hướng đến trang đăng nhập Microsoft
    current_url = driver.current_url.lower()
    if "login.microsoftonline.com" in current_url:
        print("\n" + "=" * 60)
        print("[!] PHÁT HIỆN YÊU CẦU ĐĂNG NHẬP MICROSOFT OFFICE 365:")
        print("    Vui lòng hoàn tất đăng nhập trực tiếp trên cửa sổ trình duyệt vừa mở.")
        print("    Tool sẽ tự động chờ bạn đăng nhập xong để tiếp tục...")
        print("=" * 60)

        login_success = False
        for _ in range(180):
            time.sleep(1)
            if "login.microsoftonline.com" not in driver.current_url.lower():
                login_success = True
                break
        if not login_success:
            raise TimeoutException("Quá thời gian chờ đăng nhập (180 giây).")
        print("[+] Đăng nhập thành công! Đang tải tài liệu...")
        time.sleep(3)

    # Chờ viewer xuất hiện
    wait = WebDriverWait(driver, 25)
    try:
        wait.until(
            EC.presence_of_element_located(
                (By.CSS_SELECTOR, "canvas, div.page, [data-page-number], #viewerContainer, #WACViewPanel_ScrollViewer")
            )
        )
    except TimeoutException:
        print("[!] Cảnh báo: Không phát hiện được viewer tiêu chuẩn trong 25 giây đầu, đang thử quét lại DOM...")

    # Phát hiện tổng số trang
    total_pages = 1
    if total_pages_override and total_pages_override > 0:
        total_pages = total_pages_override
        print(f"[+] Sử dụng số trang do người dùng chỉ định: {total_pages} trang.")
    else:
        # 1. Tìm thuộc tính aria-setsize (Office Online / Fluent UI)
        try:
            set_size_elems = driver.find_elements(By.CSS_SELECTOR, "[aria-setsize]")
            for el in set_size_elems:
                val = el.get_attribute("aria-setsize")
                if val and val.isdigit() and int(val) > 1:
                    total_pages = max(total_pages, int(val))
        except Exception:
            pass

        # 2. Tìm input số trang / spinbutton
        try:
            inputs = driver.find_elements(By.CSS_SELECTOR, "input[type='number'], input[role='spinbutton'], input.pageNumber")
            for inp in inputs:
                max_attr = inp.get_attribute("aria-valuemax") or inp.get_attribute("max")
                if max_attr and max_attr.isdigit() and int(max_attr) > 1:
                    total_pages = max(total_pages, int(max_attr))
        except Exception:
            pass

        # 3. Quét text trong body tìm mẫu 'X of Y' hoặc 'X / Y'
        try:
            body_text = driver.find_element(By.TAG_NAME, "body").text
            match = re.search(r"(?:Trang|Page)?\s*(\d+)\s*(?:/|of|của|trên)\s*(\d+)", body_text, flags=re.I)
            if match:
                total_pages = max(total_pages, int(match.group(2)))
        except Exception:
            pass

        # 4. Quét div trang DOM
        page_elements = driver.find_elements(By.CSS_SELECTOR, "div.page[data-page-number], div.page, .pdf-page, .WacPage")
        if page_elements:
            total_pages = max(total_pages, len(page_elements))

        print(f"[+] Phát hiện ước tính: {total_pages} trang.")

    # Cảnh báo nếu file quá lớn (> 500 trang)
    if total_pages > 500:
        print(f"[!] CẢNH BÁO: File có tới {total_pages} trang (> 500 trang).")
        confirm = input("    Bạn có chắc chắn muốn tiếp tục không? (y/n): ")
        if confirm.strip().lower() != "y":
            print("[*] Đã dừng theo yêu cầu của người dùng.")
            return []

    captured_images: List[str] = []

    # Đưa tài liệu về trang bắt đầu (start_page)
    if start_page > 1:
        print(f"[*] Đang chuyển đến trang bắt đầu: {start_page}...")
    else:
        driver.execute_script("window.scrollTo(0, 0);")

    # Thuật toán cuộn & chụp từng trang
    for page_num in range(start_page, total_pages + 1):
        print(f"[*] Đang chụp trang {page_num}/{total_pages}...", end="\r", flush=True)

        # Cách 1: Thử điều hướng qua input số trang
        driver.execute_script("""
            const pageNum = arguments[0];
            const inp = document.querySelector('input[type="number"], input[role="spinbutton"], input.pageNumber');
            if (inp) {
                inp.value = pageNum;
                inp.dispatchEvent(new Event('input', {bubbles: true}));
                inp.dispatchEvent(new Event('change', {bubbles: true}));
                inp.dispatchEvent(new KeyboardEvent('keydown', {key: 'Enter', keyCode: 13, bubbles: true}));
            }
        """, page_num)

        # Cách 2: Tìm phần tử trang trong DOM nếu có
        page_elements = driver.find_elements(
            By.CSS_SELECTOR, f"div.page[data-page-number='{page_num}'], [data-page-index='{page_num - 1}'], div.page"
        )
        target_el = page_elements[0] if page_elements else None

        if target_el:
            driver.execute_script("arguments[0].scrollIntoView({behavior: 'instant', block: 'center'});", target_el)
        else:
            # Cuộn cửa sổ / phát phím PageDown
            driver.execute_script("""
                const scrollContainer = document.querySelector('#viewerContainer, .pdfViewer, #main-content') || document.documentElement;
                scrollContainer.scrollBy(0, scrollContainer.clientHeight * 0.9);
                document.dispatchEvent(new KeyboardEvent('keydown', {key: 'PageDown', code: 'PageDown', keyCode: 34, bubbles: true}));
            """)

        time.sleep(page_delay)

        img_path = os.path.join(temp_dir, f"page_{page_num:04d}.png")

        # THỬ METHOD 3: Canvas capture qua JavaScript toDataURL (Chất lượng nguyên bản cực nét)
        canvas_captured = False
        try:
            canvas = None
            if target_el:
                c_list = target_el.find_elements(By.TAG_NAME, "canvas")
                if c_list:
                    canvas = c_list[0]
            if not canvas:
                # Tìm canvas lớn nhất đang hiển thị
                visible_canvases = driver.find_elements(By.TAG_NAME, "canvas")
                for c in visible_canvases:
                    if c.is_displayed() and c.size["width"] > 150:
                        canvas = c
                        break

            if canvas:
                data_url = driver.execute_script("return arguments[0].toDataURL('image/png');", canvas)
                if data_url and data_url.startswith("data:image/png;base64,"):
                    b64_data = data_url.split(",")[1]
                    img_bytes = base64.b64decode(b64_data)
                    with open(img_path, "wb") as f:
                        f.write(img_bytes)
                    canvas_captured = True
                    if verbose:
                        print(f"\n[+] Trang {page_num}: Đã trích xuất trực tiếp từ thẻ <canvas>.")
        except Exception as e:
            if verbose:
                print(f"\n[!] Canvas capture lỗi ở trang {page_num}, chuyển sang screenshot: {e}")

        # THỬ METHOD 1: Screenshot Element Fallback
        if not canvas_captured:
            try:
                if target_el and target_el.size["width"] > 50 and target_el.size["height"] > 50:
                    target_el.screenshot(img_path)
                else:
                    driver.save_screenshot(img_path)
                if verbose:
                    print(f"\n[+] Trang {page_num}: Đã chụp qua screenshot element.")
            except Exception as e:
                print(f"\n[!] Không thể chụp trang {page_num}: {e}")
                continue

        if os.path.exists(img_path) and os.path.getsize(img_path) > 0:
            captured_images.append(img_path)

    print(f"\n[+] Đã chụp thành công {len(captured_images)}/{total_pages} trang.")
    return captured_images


def main(
    url: str,
    output_path: Optional[str] = None,
    browser_type: str = "chrome",
    profile_dir: Optional[str] = None,
    headless: bool = False,
    page_delay: float = 2.0,
    start_page: int = 1,
    total_pages: Optional[int] = None,
    keep_temp: bool = False,
    verbose: bool = False,
) -> int:
    """
    Hàm chính điều phối toàn bộ quy trình tải tài liệu PDF.

    Args:
        url: URL SharePoint hoặc OneDrive của file.
        output_path: Đường dẫn lưu file PDF đầu ra.
        browser_type: 'chrome' hoặc 'firefox'.
        profile_dir: Thư mục profile trình duyệt.
        headless: Chạy ẩn không mở giao diện.
        page_delay: Thời gian chờ giữa các trang (giây).
        start_page: Trang bắt đầu quét (mặc định 1).
        total_pages: Tổng số trang tùy chỉnh (nếu muốn ghi đè).
        keep_temp: Giữ lại ảnh tạm sau khi hoàn thành.
        verbose: In log chi tiết.

    Returns:
        0 nếu thành công, 1 nếu có lỗi.
    """
    start_time = time.time()
    print("=" * 65)
    print("   OneDrive & SharePoint PDF Downloader (Restricted View)")
    print("=" * 65)

    # Bước 1: Kiểm tra tính hợp lệ của URL
    if not validate_url(url):
        print(f"[!] LỖI: URL không hợp lệ. Vui lòng cung cấp URL từ SharePoint hoặc OneDrive.")
        print(f"    URL đã nhập: {url}")
        return 1

    temp_dir = tempfile.mkdtemp(prefix="sp_pdf_")
    driver = None

    try:
        # Bước 2: Khởi tạo trình duyệt
        print(f"[*] Khởi tạo trình duyệt {browser_type.upper()}...")
        driver = init_driver(
            browser_type=browser_type,
            profile_dir=profile_dir,
            headless=headless,
            verbose=verbose,
        )

        # Trích xuất tên file
        filename = extract_filename_from_url(url, driver)
        if not output_path:
            output_path = filename
        elif not output_path.lower().endswith(".pdf"):
            output_path += ".pdf"

        # Bước 3: Thử phương pháp Direct URL trước (Method 2)
        print("[*] Đang kiểm tra phương thức Direct URL qua cookie phiên...")
        session = requests.Session()
        # Chuyển cookie từ Selenium sang requests
        driver.get(url)
        time.sleep(2)
        for cookie in driver.get_cookies():
            session.cookies.set(cookie["name"], cookie["value"], domain=cookie.get("domain", ""))

        direct_url = extract_direct_url(session, url)
        if direct_url:
            print(f"[+] Tìm thấy Direct Download URL hợp lệ! Đang tải trực tiếp...")
            resp = session.get(direct_url, stream=True, timeout=30)
            if resp.status_code == 200:
                with open(output_path, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=8192):
                        f.write(chunk)
                is_valid, pages, size = validate_pdf(output_path)
                if is_valid:
                    elapsed = time.time() - start_time
                    print("\n" + "=" * 65)
                    print(f"🎉 TẢI THÀNH CÔNG (Phương pháp: Direct Download)")
                    print(f"   - File đầu ra: {os.path.abspath(output_path)}")
                    print(f"   - Số trang:    {pages} trang")
                    print(f"   - Dung lượng:  {size / (1024 * 1024):.2f} MB")
                    print(f"   - Thời gian:   {elapsed:.2f} giây")
                    print("=" * 65)
                    return 0

        # Bước 4 & 5: Phương pháp Selenium Canvas / Screenshot Capture (Method 1 & 3)
        print("[*] Chuyển sang phương thức trích xuất Canvas & Screenshot từ Viewer...")
        image_paths = capture_pdf_via_screenshots(
            driver=driver,
            url=url,
            temp_dir=temp_dir,
            page_delay=page_delay,
            start_page=start_page,
            total_pages_override=total_pages,
            verbose=verbose,
        )

        if not image_paths:
            print("[!] LỖI: Không thu thập được trang nào từ tài liệu.")
            return 1

        # Bước 6: Ghép ảnh thành PDF
        print(f"[*] Đang ghép {len(image_paths)} trang thành file PDF...")
        merge_images_to_pdf(image_paths, output_path)

        # Bước 7: Kiểm định file PDF
        is_valid, pages, size = validate_pdf(output_path)
        if not is_valid:
            print("[!] CẢNH BÁO: File PDF tạo ra có thể bị lỗi cấu trúc.")
            return 1

        elapsed = time.time() - start_time
        print("\n" + "=" * 65)
        print(f"🎉 TẢI VÀ ĐÓNG GÓI PDF THÀNH CÔNG!")
        print(f"   - File đầu ra: {os.path.abspath(output_path)}")
        print(f"   - Số trang:    {pages} trang")
        print(f"   - Dung lượng:  {size / (1024 * 1024):.2f} MB")
        print(f"   - Thời gian:   {elapsed:.2f} giây (Trung bình: {elapsed / pages:.2f}s/trang)")
        print("=" * 65)

        return 0

    except KeyboardInterrupt:
        print("\n[!] Người dùng hủy tiến trình.")
        return 1
    except Exception as e:
        print(f"\n[!] Đã xảy ra lỗi trong quá trình xử lý: {e}")
        if verbose:
            import traceback

            traceback.print_exc()
        return 1
    finally:
        # Đóng tài nguyên
        if driver:
            try:
                driver.quit()
            except Exception:
                pass

        # Xóa thư mục tạm nếu không có cờ --keep-temp
        if not keep_temp and os.path.exists(temp_dir):
            try:
                shutil.rmtree(temp_dir)
            except Exception:
                pass
        elif keep_temp:
            print(f"[i] Giữ lại thư mục ảnh tạm: {temp_dir}")


def parse_arguments() -> argparse.Namespace:
    """Xử lý các tham số dòng lệnh."""
    parser = argparse.ArgumentParser(
        description="Tải tài liệu PDF từ SharePoint/OneDrive bị chặn download (Restricted View).",
        formatter_class=argparse.ArgumentDefaultsHelpFormatter,
    )

    parser.add_argument(
        "--url",
        "-u",
        required=True,
        type=str,
        help="URL của file SharePoint/OneDrive cần tải (Bắt buộc).",
    )
    parser.add_argument(
        "--output",
        "-o",
        type=str,
        default=None,
        help="Đường dẫn file PDF đầu ra (Mặc định: tên gốc của file).",
    )
    parser.add_argument(
        "--browser",
        "-b",
        type=str,
        choices=["chrome", "firefox"],
        default="chrome",
        help="Loại trình duyệt sử dụng để xem tài liệu.",
    )
    parser.add_argument(
        "--profile-dir",
        "-p",
        type=str,
        default=None,
        help="Đường dẫn thư mục Profile của trình duyệt (giữ phiên đăng nhập sẵn).",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        default=False,
        help="Chạy ẩn trình duyệt không hiển thị cửa sổ.",
    )
    parser.add_argument(
        "--page-delay",
        "-d",
        type=float,
        default=2.0,
        help="Thời gian chờ nạp giữa các trang (giây).",
    )
    parser.add_argument(
        "--start-page",
        type=int,
        default=1,
        help="Trang bắt đầu quét (mặc định: 1).",
    )
    parser.add_argument(
        "--total-pages",
        type=int,
        default=None,
        help="Tổng số trang cần tải (tự động phát hiện nếu để trống).",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        default=False,
        help="Giữ lại các file ảnh tạm thời trong thư mục temp.",
    )
    parser.add_argument(
        "--verbose",
        "-v",
        action="store_true",
        default=False,
        help="In chi tiết các bước xử lý (debug).",
    )

    return parser.parse_args()


if __name__ == "__main__":
    args = parse_arguments()
    exit_code = main(
        url=args.url,
        output_path=args.output,
        browser_type=args.browser,
        profile_dir=args.profile_dir,
        headless=args.headless,
        page_delay=args.page_delay,
        start_page=args.start_page,
        total_pages=args.total_pages,
        keep_temp=args.keep_temp,
        verbose=args.verbose,
    )
    sys.exit(exit_code)
