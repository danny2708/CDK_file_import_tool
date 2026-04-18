#!/usr/bin/env python3
"""
Tool crawl văn bản QC-TC từ vr.org.vn và nhập vào hệ thống quantri-vr.e-office.vn.

Workflow:
1) Crawl metadata + file từ vr.org.vn.
2) Upload sẵn toàn bộ file vào Kho dữ liệu.
3) Nhập văn bản từ cuối lên, bỏ qua 6 văn bản cuối đã nhập.
4) Đồng bộ lĩnh vực theo trạng thái hiện tại của form draft.
5) File đính kèm chỉ chọn lại từ Kho dữ liệu, không upload tại bước nhập.
"""

from __future__ import annotations

import argparse
import json
import logging
import re
import sys
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Any, Iterable, Optional
from urllib.parse import parse_qs, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from playwright.sync_api import BrowserContext, Page, TimeoutError as PlaywrightTimeoutError, sync_playwright


# =========================
# CẤU HÌNH
# =========================

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
RAW_DIR = DATA_DIR / "raw"
DOWNLOAD_DIR = DATA_DIR / "downloads"
SCREENSHOT_DIR = DATA_DIR / "screenshots"
STATE_DIR = DATA_DIR / "state"
OUTPUT_JSON = DATA_DIR / "documents.json"
LOG_FILE = DATA_DIR / "tool.log"
PROFILE_DIR = DATA_DIR / "playwright_profile"

SOURCE_LIST_URL = "https://www.vr.org.vn/quy-chuan-tieu-chuan/Pages/default.aspx"
SOURCE_DETAIL_URL = "https://www.vr.org.vn/quy-chuan-tieu-chuan/Pages/default.aspx?ItemID={item_id}"
TARGET_BASE_URL = "https://quantri-vr.e-office.vn"
TARGET_CREATE_URL = f"{TARGET_BASE_URL}/van-ban/them-moi"
TARGET_LIST_URL = f"{TARGET_BASE_URL}/van-ban/danh-sach"

FIELD_MAP_PATH = DATA_DIR / "field_mapping.json"

DEFAULT_FIELD_MAP: dict[str, str] = {
    "lĩnh vực hoạt động": "Lĩnh vực hoạt động",
    "tàu biển": "Tàu biển",
    "phương tiện thủy nội địa": "Phương tiện thủy nội địa",
    "công trình biển": "Công trình biển",
    "phương tiện đường sắt": "Phương tiện đường sắt",
    "sản phẩm công nghiệp": "Sản phẩm công nghiệp",
    "phương tiện cơ giới đường bộ": "Phương tiện cơ giới đường bộ",
    "thẩm định thiết kế tàu biển": "Thẩm định thiết kế tàu biển",
    "thẩm định thiết kế phương tiện thủy nội địa": "Thẩm định thiết kế phương tiện thủy nội địa",
    "kiểm tra, giám sát đóng mới hoán cải phương tiện thủy nội": "Kiểm tra, giám sát đóng mới hoán cải phương tiện thủy nội",
    "kiểm tra, giám sát đóng mới/hoán cải phương tiện thủy nội địa": "Kiểm tra, giám sát đóng mới/hoán cải phương tiện thủy nội địa",
    "giám sát đóng mới phương tiện thủy nội địa": "Giám sát đóng mới phương tiện thủy nội địa",
    "giám sát hoán cải phương tiện thủy nội địa": "Giám sát hoán cải phương tiện thủy nội địa",
    "kiểm tra phương tiện thủy nội địa trong khai thác": "Kiểm tra phương tiện thủy nội địa trong khai thác",
    "kiểm tra, xác nhận năng lực cơ sở đóng mới, hoán cải, sửa chữa và phục hồi phương tiện thủy nội địa": "Kiểm tra, xác nhận năng lực cơ sở đóng mới, hoán cải, sửa chữa và phục hồi phương tiện thủy nội địa",
    "thẩm định thiết kế sản phẩm công nghiệp sử dụng cho phương tiện thủy nội địa": "Thẩm định thiết kế sản phẩm công nghiệp sử dụng cho phương tiện thủy nội địa",
    "kiểm tra tàu biển trong khai thác": "Kiểm tra tàu biển trong khai thác",
    "giám sát đóng mới tàu biển": "Giám sát đóng mới tàu biển",
    "đánh giá chứng nhận an toàn, an ninh, loa động hang hải": "Đánh giá chứng nhận an toàn, an ninh, loa động hang hải",
    "đánh giá chứng nhận an toàn, an ninh, lao động hàng hải": "Đánh giá chứng nhận an toàn, an ninh, loa động hang hải",
    "đánh giá, công nhận cơ sở": "Đánh giá, công nhận cơ sở",
    "kiểm tra công trình biển trong khai thác": "Kiểm tra công trình biển trong khai thác",
    "kiểm tra, giám sát đóng mới hoán cải và nhập khẩu công trình biển": "Kiểm tra, giám sát đóng mới hoán cải và nhập khẩu công trình biển",
    "thẩm định thiết kế": "Thẩm định thiết kế",
    "khác": "Khác",
    "chứng nhận sản phẩm công nghiệp": "Chứng nhận sản phẩm công nghiệp",
    "kiểm tra chứng nhận spcn nhập khẩu": "Kiểm tra chứng nhận SPCN nhập khẩu",
    "chứng nhận spcn nhập khẩu": "Kiểm tra chứng nhận SPCN nhập khẩu",
    "kiểm tra, giám sát chế tạo spcn": "Kiểm tra, giám sát chế tạo SPCN",
    "công nhận sản phẩm công nghiệp": "Công nhận sản phẩm công nghiệp",
    "đánh giá năng lực cơ sở chế tạo spcn": "Đánh giá năng lực cơ sở chế tạo SPCN",
    "kiểm tra theo luật thiết bị nâng, bình chịu áp lực, nồi hơi": "Kiểm tra theo luật thiết bị nâng, bình chịu áp lực, nồi hơi",
    "kiểm tra chứng nhận thợ - hàn": "Kiểm tra chứng nhận thợ - hàn",
    "thẩm định thiết kế ptcgđb": "Thẩm định thiết kế PTCGĐB",
    "thẩm định thiết kế ptcgdb": "Thẩm định thiết kế PTCGĐB",
    "thẩm định thiết kế ptcgdđ": "Thẩm định thiết kế PTCGĐB",
    "chứng nhận chất lượng kiểu loại xe sản xuát, lắp ráp": "Chứng nhận chất lượng kiểu loại xe sản xuát, lắp ráp",
    "chứng nhận chất lượng kiểu loại xe sản xuất, lắp ráp": "Chứng nhận chất lượng kiểu loại xe sản xuát, lắp ráp",
    "kiểm tra chứng nhận chất lượng kiểu loại xe cơ giới": "Kiểm tra chứng nhận chất lượng kiểu loại xe cơ giới",
    "kiểm tra chứng nhận chất lượng xe cơ giới nhập khẩu": "Kiểm tra chứng nhận chất lượng xe cơ giới nhập khẩu",
    "thử nghiệm ô tô, rơ moóc, sơ mi rơ moóc": "Thử nghiệm ô tô, rơ moóc, sơ mi rơ moóc",
    "thử nghiệm xe cơ giới, xe máy chuyên dùng và phụ tùng xe cơ giới": "Thử nghiệm xe cơ giới, xe máy chuyên dùng và phụ tùng xe cơ giới",
    "thử nghiệm chất lượng ptcgđb và linh kiên": "Thử nghiệm chất lượng PTCGĐB và linh kiên",
    "thử nghiệm chất lượng ptcgdb và linh kiện": "Thử nghiệm chất lượng PTCGĐB và linh kiên",
    "kiểm định xe cơ giới và xe bốn bánh có gắn động cơ đang lưu hành": "Kiểm định xe cơ giới và xe bốn bánh có gắn động cơ đang lưu hành",
    "kiểm tra xe máy chuyên dùng trong khai thác sử dụng": "Kiểm tra xe máy chuyên dùng trong khai thác sử dụng",
    "kiểm tra chứng nhận điều kiện kinh doanh hoạt động kiểm định xe cơ giới": "Kiểm tra chứng nhận điều kiện kinh doanh hoạt động kiểm định xe cơ giới",
    "kiểm tra chất lượng atkt và bvmt xe nhập khẩu": "Kiểm tra chất lượng ATKT và BVMT xe nhập khẩu",
    "thẩm định thiết kế xe cơ giới lắp ráp từ xe cơ sở": "Thẩm định thiết kế xe cơ giới lắp ráp từ xe cơ sở",
    "thẩm định an toàn hệ thống đường sắt đô thị": "Thẩm định an toàn hệ thống đường sắt đô thị",
    "thẩm định thiết kế phương tiện, thiết bị giao thông đường sắt": "Thẩm định thiết kế phương tiện, thiết bị giao thông đường sắt",
    "kiểm tra chứng nhận chất lượng phương tiện, thiết bị giao thông đường sắt sxlr": "Kiểm tra chứng nhận chất lượng phương tiện, thiết bị giao thông đường sắt SXLR",
    "kiểm tra chứng nhận chất lượng phương tiện, thiết bị giao thông đường sắt nhập khẩu": "Kiểm tra chứng nhận chất lượng phương tiện, thiết bị giao thông đường sắt nhập khẩu",
    "kiểm tra phương tiện đường sắt trong khai thác": "Kiểm tra phương tiện đường sắt trong khai thác",
}

FIXED_FORM_VALUES = {
    "loai_van_ban": "Tiêu chuẩn",
    "co_quan_ban_hanh": "Bộ giao thông vận tải",
    "nguoi_ky": "Bộ giao thông vận tải",
    "chuc_danh": "b",
}

CRAWL_CONFIG = {
    "page_start": 1,
    "page_end": 10,
    "known_total": 190,
    "skip_last_already_imported": 6,
}

PLAYWRIGHT_DEFAULT_TIMEOUT_MS = 30000


# =========================
# MODEL
# =========================

@dataclass
class Attachment:
    name: str
    url: str
    local_path: str = ""


@dataclass
class SourceDocument:
    item_id: int
    source_url: str
    title: str
    so_ky_hieu: str
    loai_qc_tc: str
    linh_vuc_raw: list[str]
    linh_vuc_mapped: list[str]
    trich_yeu: str
    noi_dung_text: str
    noi_dung_html: str
    attachments: list[Attachment] = field(default_factory=list)


# =========================
# HỖ TRỢ CHUNG
# =========================

def ensure_dirs() -> None:
    for path in [DATA_DIR, RAW_DIR, DOWNLOAD_DIR, SCREENSHOT_DIR, STATE_DIR]:
        path.mkdir(parents=True, exist_ok=True)


def setup_logging(verbose: bool = True) -> None:
    ensure_dirs()
    handlers: list[logging.Handler] = [logging.FileHandler(LOG_FILE, encoding="utf-8")]
    if verbose:
        handlers.append(logging.StreamHandler(sys.stdout))

    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=handlers,
        force=True,
    )


def slugify_filename(name: str) -> str:
    name = name.strip().replace("/", "_").replace("\\", "_")
    name = re.sub(r"\s+", " ", name)
    name = re.sub(r"[^\w\-. ()À-ỹ]", "_", name, flags=re.UNICODE)
    return name[:240].strip() or "file"


def clean_text(text: str) -> str:
    text = re.sub(r"\s+", " ", text or "")
    return text.strip()


def normalize_key(text: str) -> str:
    return clean_text(text).lower()


def read_json(path: Path, default: Any) -> Any:
    if not path.exists():
        return default
    try:
        content = path.read_text(encoding="utf-8").strip()
        if not content:
            return default
        return json.loads(content)
    except Exception:
        return default


def get_field_map() -> dict[str, str]:
    file_map = read_json(FIELD_MAP_PATH, {})
    normalized_file_map = {normalize_key(k): v for k, v in file_map.items()}
    normalized_default_map = {normalize_key(k): v for k, v in DEFAULT_FIELD_MAP.items()}
    normalized_default_map.update(normalized_file_map)
    return normalized_default_map


def export_default_field_mapping() -> None:
    if FIELD_MAP_PATH.exists():
        return
    write_json(FIELD_MAP_PATH, DEFAULT_FIELD_MAP)


def write_json(path: Path, data: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def save_state(name: str, data: Any) -> None:
    write_json(STATE_DIR / f"{name}.json", data)


def load_state(name: str, default: Any) -> Any:
    return read_json(STATE_DIR / f"{name}.json", default)


# =========================
# CRAWLER
# =========================

class VrCrawler:
    def __init__(self, session: Optional[requests.Session] = None):
        self.session = session or requests.Session()
        self.session.headers.update(
            {
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/147.0.0.0 Safari/537.36"
                )
            }
        )

    def fetch(self, url: str) -> str:
        logging.info("GET %s", url)
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        resp.encoding = resp.apparent_encoding or resp.encoding or "utf-8"
        return resp.text

    def crawl_list_pages(self, page_start: int, page_end: int) -> list[int]:
        item_ids: list[int] = []
        for page_num in range(page_start, page_end + 1):
            url = f"{SOURCE_LIST_URL}?Page={page_num}"
            html = self.fetch(url)
            (RAW_DIR / f"list_page_{page_num}.html").write_text(html, encoding="utf-8")
            page_ids = self.extract_item_ids_from_list(html)
            logging.info("Page %s -> %s item IDs", page_num, len(page_ids))
            item_ids.extend(page_ids)
        return list(dict.fromkeys(item_ids))

    def extract_item_ids_from_list(self, html: str) -> list[int]:
        soup = BeautifulSoup(html, "lxml")
        item_ids: list[int] = []
        for a in soup.select("a[href*='ItemID=']"):
            href = a.get("href", "")
            if "ItemID=" not in href:
                continue
            full_url = urljoin(SOURCE_LIST_URL, href)
            qs = parse_qs(urlparse(full_url).query)
            item_id_raw = (qs.get("ItemID") or [None])[0]
            if item_id_raw and str(item_id_raw).isdigit():
                item_ids.append(int(item_id_raw))
        return item_ids

    def crawl_detail(self, item_id: int) -> SourceDocument:
        url = SOURCE_DETAIL_URL.format(item_id=item_id)
        html = self.fetch(url)
        (RAW_DIR / f"detail_{item_id}.html").write_text(html, encoding="utf-8")
        return self.parse_detail(item_id=item_id, url=url, html=html)

    def parse_detail(self, item_id: int, url: str, html: str) -> SourceDocument:
        soup = BeautifulSoup(html, "lxml")
        root = soup.select_one("div.qc-qp-tc") or soup

        title = clean_text((root.select_one("h3") or soup.select_one("h3")).get_text(" ", strip=True))
        so_ky_hieu = self.extract_value_by_label(html, "Số ký hiệu")
        if not so_ky_hieu:
            text_all = clean_text(root.get_text(" ", strip=True))
            m = re.search(
                r"Số ký hiệu\s*:?\s*([A-ZÀ-Ỹ0-9][^:]{3,100}?)\s*(?:Loại QC|Loại QC-TC|TÀI LIỆU ĐÍNH KÈM|$)",
                text_all,
                flags=re.IGNORECASE,
            )
            if m:
                so_ky_hieu = clean_text(m.group(1))

        loai_qc_tc = self.extract_value_by_label(html, "Loại QC - TC") or self.extract_value_by_label(html, "Loại QC-TC")
        if not loai_qc_tc:
            loai_qc_tc = self.extract_loai_qc_tc_from_text(root.get_text(" ", strip=True))

        linh_vuc_value = self.extract_value_by_label(html, "Lĩnh vực / Loại hình công việc")
        linh_vuc_raw = self.parse_linh_vuc(linh_vuc_value)
        linh_vuc_mapped = self.map_linh_vuc(linh_vuc_raw)

        des_div = root.select_one("div.des")
        content_div = root.select_one("div.content")
        trich_yeu = clean_text(title)
        noi_dung_text = clean_text(content_div.get_text(" ", strip=True) if content_div else "")
        noi_dung_html = str(content_div) if content_div else ""

        if not noi_dung_text and des_div:
            noi_dung_text = clean_text(des_div.get_text(" ", strip=True))
        if not noi_dung_html and des_div:
            noi_dung_html = str(des_div)

        attachments = self.extract_attachments(root, url)

        return SourceDocument(
            item_id=item_id,
            source_url=url,
            title=title,
            so_ky_hieu=clean_text(so_ky_hieu),
            loai_qc_tc=clean_text(loai_qc_tc),
            linh_vuc_raw=linh_vuc_raw,
            linh_vuc_mapped=linh_vuc_mapped,
            trich_yeu=trich_yeu,
            noi_dung_text=noi_dung_text,
            noi_dung_html=noi_dung_html,
            attachments=attachments,
        )

    def extract_value_by_label(self, html: str, label: str) -> str:
        label_re = re.escape(label)
        patterns = [
            rf"{label_re}\s*:\s*(.+?)(?:<|\n|$)",
            rf"{label_re}</div>\s*<div[^>]*>\s*(.+?)(?:</div>|<)",
        ]
        for pattern in patterns:
            match = re.search(pattern, html, flags=re.IGNORECASE | re.DOTALL)
            if match:
                value = BeautifulSoup(match.group(1), "lxml").get_text(" ", strip=True)
                return clean_text(value)
        return ""

    def extract_loai_qc_tc_from_text(self, text: str) -> str:
        text = clean_text(text)
        match = re.search(r"Loại\s*QC\s*-?\s*TC\s*:\s*(Quy chuẩn|Tiêu chuẩn)", text, flags=re.IGNORECASE)
        return clean_text(match.group(1)) if match else ""

    def parse_linh_vuc(self, raw_value: str) -> list[str]:
        if not raw_value:
            return []
        parts = [clean_text(x) for x in raw_value.split(",")]
        return [x for x in parts if x]

    def map_linh_vuc(self, values: Iterable[str]) -> list[str]:
        mapped: list[str] = []
        field_map = get_field_map()
        unmatched: list[str] = []

        for value in values:
            key = normalize_key(value)
            if key in field_map:
                target_value = field_map[key]
                if target_value not in mapped:
                    mapped.append(target_value)
                continue

            matched = False
            for source_key, target_value in field_map.items():
                if source_key in key or key in source_key:
                    if target_value not in mapped:
                        mapped.append(target_value)
                    matched = True
                    break

            if not matched:
                candidate = clean_text(value)
                if candidate and candidate not in mapped:
                    mapped.append(candidate)
                    unmatched.append(candidate)

        if unmatched:
            missing = load_state("unmatched_fields", [])
            for item in unmatched:
                if item not in missing:
                    missing.append(item)
            save_state("unmatched_fields", missing)
            logging.warning("Có lĩnh vực chưa map cứng, đang dùng raw value tạm thời: %s", unmatched)

        return mapped

    def extract_attachments(self, root: BeautifulSoup, base_url: str) -> list[Attachment]:
        results: list[Attachment] = []
        for a in root.select("p.attach ~ ul a, ul a[href*='Attachments/'], a[href*='Attachments/']"):
            href = a.get("href", "").strip()
            if not href:
                continue
            file_url = urljoin(base_url, href)
            file_name = clean_text(a.get_text(" ", strip=True)) or Path(urlparse(file_url).path).name
            att = Attachment(name=file_name, url=file_url)
            if att.url not in {x.url for x in results}:
                results.append(att)
        return results

    def download_attachments(self, doc: SourceDocument) -> SourceDocument:
        doc_dir = DOWNLOAD_DIR / f"{doc.item_id:04d}_{slugify_filename(doc.so_ky_hieu or doc.title)}"
        doc_dir.mkdir(parents=True, exist_ok=True)
        for att in doc.attachments:
            ext = Path(urlparse(att.url).path).suffix or Path(att.name).suffix
            filename = slugify_filename(Path(att.name).stem) + ext
            out_path = doc_dir / filename
            if not out_path.exists():
                logging.info("Download file: %s", att.url)
                resp = self.session.get(att.url, timeout=60)
                resp.raise_for_status()
                out_path.write_bytes(resp.content)
            att.local_path = str(out_path)
        return doc


def crawl_all_documents() -> list[SourceDocument]:
    crawler = VrCrawler()
    item_ids = crawler.crawl_list_pages(
        page_start=CRAWL_CONFIG["page_start"],
        page_end=CRAWL_CONFIG["page_end"],
    )

    logging.info("Tổng item IDs lấy được: %s", len(item_ids))
    documents: list[SourceDocument] = []
    for idx, item_id in enumerate(item_ids, start=1):
        logging.info("[%s/%s] Crawl item_id=%s", idx, len(item_ids), item_id)
        try:
            doc = crawler.crawl_detail(item_id)
            doc = crawler.download_attachments(doc)
            documents.append(doc)
            write_json(OUTPUT_JSON, [asdict(x) for x in documents])
        except Exception as exc:
            logging.exception("Lỗi crawl item_id=%s: %s", item_id, exc)
            errors = load_state("crawl_errors", [])
            errors.append({"item_id": item_id, "error": str(exc)})
            save_state("crawl_errors", errors)

    logging.info("Đã crawl xong %s văn bản", len(documents))
    return documents


# =========================
# PLAYWRIGHT: TIỆN ÍCH
# =========================

class EOfficeBot:
    def __init__(self, headless: bool = False, slow_mo: int = 100, interactive_login: bool = False):
        self.headless = headless
        self.slow_mo = slow_mo
        self.interactive_login = interactive_login
        self.playwright = None
        self.context: Optional[BrowserContext] = None
        self.page: Optional[Page] = None

    def __enter__(self) -> "EOfficeBot":
        ensure_dirs()
        self.playwright = sync_playwright().start()
        self.context = self.playwright.chromium.launch_persistent_context(
            user_data_dir=str(PROFILE_DIR),
            headless=self.headless,
            slow_mo=self.slow_mo,
            accept_downloads=True,
            viewport={"width": 1600, "height": 1000},
        )
        self.context.set_default_timeout(PLAYWRIGHT_DEFAULT_TIMEOUT_MS)
        pages = self.context.pages
        self.page = pages[0] if pages else self.context.new_page()
        if self.interactive_login:
            self.ensure_login()
        return self

    def __exit__(self, exc_type, exc_val, exc_tb) -> None:
        try:
            if self.context:
                self.context.close()
        finally:
            if self.playwright:
                self.playwright.stop()

    def ensure_login(self) -> None:
        assert self.page is not None

        print("\n=== DEBUG LOGIN MODE ===", flush=True)
        print(">>> Chrome đã mở.", flush=True)
        print(">>> Hãy tự mở URL nội bộ trong cửa sổ này:", flush=True)
        print(f">>> {TARGET_LIST_URL}", flush=True)
        print(">>> Đăng nhập xong thì quay lại terminal nhấn Enter.\n", flush=True)

        try:
            self.page.set_content(
                "<html><body style='font-family:Arial;padding:24px'>"
                "<h2>Playwright đã mở</h2>"
                "<p>Đăng nhập hệ thống rồi quay lại terminal nhấn Enter.</p>"
                "</body></html>"
            )
        except Exception:
            pass

        input("Nhấn Enter để tiếp tục... ")

        print(">>> Đang chuyển sang trang thêm mới...", flush=True)
        self.page.goto(TARGET_CREATE_URL, wait_until="load", timeout=60000)
        self.wait_for_form_ready()
        print(">>> Đã vào form thêm mới.", flush=True)

    def goto_create_form(self) -> None:
        assert self.page is not None
        self.page.goto(TARGET_CREATE_URL, wait_until="load", timeout=60000)
        self.wait_for_form_ready()

    def wait_for_form_ready(self) -> None:
        assert self.page is not None
        self.page.wait_for_url("**/van-ban/them-moi", timeout=60000)
        print(f"DEBUG: current url = {self.page.url}", flush=True)

        self.page.wait_for_selector("input[formcontrolname='officialNumber']", timeout=30000)
        self.page.wait_for_selector("textarea[formcontrolname='source']", timeout=30000)
        self.page.wait_for_selector("app-file-attachment a.btn", timeout=30000)

        print("DEBUG: form ready", flush=True)

    def screenshot(self, name: str) -> None:
        assert self.page is not None
        self.page.screenshot(path=str(SCREENSHOT_DIR / name), full_page=True)

    @staticmethod
    def _xpath_literal(value: str) -> str:
        if "'" not in value:
            return f"'{value}'"
        if '"' not in value:
            return f'"{value}"'
        parts = value.split("'")
        return "concat(" + ", \"'\", ".join(f"'{part}'" for part in parts) + ")"

    def _first_visible(self, locator):
        limit = min(locator.count(), 50)
        for i in range(limit):
            try:
                candidate = locator.nth(i)
                if candidate.is_visible():
                    return candidate
            except Exception:
                pass
        return None

    def _matching_text_locators(self, root, text: str) -> list[Any]:
        literal = self._xpath_literal(text)
        return [
            root.get_by_text(text, exact=True),
            root.locator(f"xpath=.//*[normalize-space(text())={literal}]"),
            root.locator(f"xpath=.//*[normalize-space(.)={literal}]"),
            root.locator(f"xpath=.//*[contains(normalize-space(.), {literal})]"),
        ]

    def _click_text_in_root(self, root, text: str) -> bool:
        click_target_selectors = [
            "xpath=ancestor-or-self::*[@role='treeitem'][1]",
            "xpath=ancestor-or-self::a[1]",
            "xpath=ancestor-or-self::button[1]",
            "xpath=ancestor-or-self::li[1]",
            "xpath=ancestor-or-self::span[1]",
            "xpath=ancestor-or-self::div[1]",
        ]

        for matches in self._matching_text_locators(root, text):
            for i in range(min(matches.count(), 30)):
                try:
                    item = matches.nth(i)
                    if not item.is_visible():
                        continue

                    click_targets = []
                    for selector in click_target_selectors:
                        click_targets.append(item.locator(selector))
                    click_targets.append(item)

                    for target in click_targets:
                        visible_target = self._first_visible(target)
                        if visible_target is None:
                            continue
                        try:
                            visible_target.click(force=True, timeout=1500)
                            time.sleep(0.5)
                            return True
                        except Exception:
                            pass
                except Exception:
                    pass
        return False

    def _locator_has_selected_files(self, locator) -> bool:
        try:
            count = min(locator.count(), 5)
        except Exception:
            return False

        for i in range(count):
            try:
                selected_count = locator.nth(i).evaluate(
                    "(el) => el instanceof HTMLInputElement && el.files ? el.files.length : 0"
                )
                if selected_count and int(selected_count) > 0:
                    return True
            except Exception:
                pass
        return False

    def _is_effectively_enabled(self, locator) -> bool:
        try:
            return bool(
                locator.evaluate(
                    """
                    (el) => {
                        const disabledAttr = el.getAttribute && (
                            el.getAttribute('disabled') !== null ||
                            el.getAttribute('aria-disabled') === 'true'
                        );
                        const disabledProp = 'disabled' in el ? !!el.disabled : false;
                        const className = (el.className || '').toString().toLowerCase();
                        return !(disabledAttr || disabledProp || className.includes('disabled'));
                    }
                    """
                )
            )
        except Exception:
            return True

    def _exact_action_candidates(self, root, text: str):
        literal = self._xpath_literal(text)
        return [
            root.get_by_role("button", name=text),
            root.locator(f"xpath=.//button[normalize-space(.)={literal}]"),
            root.locator(f"xpath=.//a[normalize-space(.)={literal}]"),
            root.locator(f"xpath=.//*[self::button or self::a or @role='button'][normalize-space(.)={literal}]"),
        ]

    def _modal_footer_action_candidates(self, modal_root, text: str):
        literal = self._xpath_literal(text)
        footer = modal_root.locator(".ant-modal-footer, [nz-modal-footer]").first
        return [
            footer.get_by_role("button", name=text),
            footer.locator(f"xpath=.//button[normalize-space(.)={literal}]"),
            footer.locator(f"xpath=.//a[normalize-space(.)={literal}]"),
            footer.locator(
                f"xpath=.//*[self::button or self::a or @role='button'][normalize-space(.)={literal}]"
            ),
        ]

    def _upload_popup_root(self):
        assert self.page is not None
        popup_titles = ["Upload file", "Chọn file để upload"]
        container_selectors = [
            "div.ant-modal-content",
            "div.ant-modal",
            "div.ant-modal-wrap",
            "nz-modal-container",
            "div[role='dialog']",
        ]

        for title in popup_titles:
            for selector in container_selectors:
                try:
                    popup = self._first_visible(self.page.locator(selector).filter(has_text=title))
                    if popup is not None:
                        return popup
                except Exception:
                    pass

        title_locators = [
            self.page.get_by_text("Upload file", exact=True),
            self.page.get_by_text("Chọn file để upload", exact=True),
            self.page.locator("h1, h2, h3, h4, .ant-modal-title").filter(has_text="Upload file"),
        ]
        ancestor_selectors = [
            "xpath=ancestor::div[contains(@class,'ant-modal-content')][1]",
            "xpath=ancestor::div[contains(@class,'ant-modal')][1]",
            "xpath=ancestor::div[contains(@class,'ant-modal-wrap')][1]",
            "xpath=ancestor::nz-modal-container[1]",
            "xpath=ancestor::*[@role='dialog'][1]",
        ]

        for title in title_locators:
            visible_title = self._first_visible(title)
            if visible_title is None:
                continue
            for selector in ancestor_selectors:
                try:
                    popup = self._first_visible(visible_title.locator(selector))
                    if popup is not None:
                        return popup
                except Exception:
                    pass
        return None

    def _is_kho_modal_open(self) -> bool:
        assert self.page is not None
        try:
            return (
                self._first_visible(
                    self.page.locator(
                        "div.ant-modal-wrap, div.ant-modal, nz-modal-container, div[role='dialog']"
                    ).filter(has_text="Kho dữ liệu")
                )
                is not None
            )
        except Exception:
            return False

    def _kho_file_search_input(self, modal=None):
        modal = modal or self._kho_root()
        selectors = [
            "input[placeholder*='Tìm kiếm file']",
            "input[placeholder*='Tim kiem file']",
            "input[placeholder='Tìm kiếm file']",
            "input[placeholder='Tim kiem file']",
        ]
        for selector in selectors:
            try:
                candidate = self._first_visible(modal.locator(selector))
                if candidate is not None:
                    return candidate
            except Exception:
                pass
        return modal.locator("input").first

    def _trigger_kho_search(self, modal=None) -> None:
        modal = modal or self._kho_root()
        search_input = self._kho_file_search_input(modal)
        try:
            search_input.press("Enter")
            time.sleep(0.8)
            return
        except Exception:
            pass

        search_icons = [
            modal.locator("i.bi-search").locator("xpath=ancestor::*[@role='button'][1]"),
            modal.locator("i.bi-search").locator("xpath=ancestor::a[1]"),
            modal.locator("i.bi-search"),
        ]
        for candidate in search_icons:
            try:
                btn = self._first_visible(candidate)
                if btn is None:
                    continue
                btn.click(force=True, timeout=2000)
                time.sleep(0.8)
                return
            except Exception:
                pass

    def _kho_refresh_candidates(self, modal=None) -> list[Any]:
        modal = modal or self._kho_root()
        return [
            modal.locator("[nztype='reload'], [data-icon='reload'], [data-icon='redo'], [class*='reload']"),
            modal.locator("i.bi-arrow-repeat").locator("xpath=ancestor::a[1]"),
            modal.locator("[nztooltiptitle*='Làm mới']").locator("xpath=.//a | ancestor::a[1]"),
            modal.locator("span").filter(has_text="Làm mới").locator("xpath=ancestor::a[1]"),
        ]

    def _first_visible_kho_result_item(self, modal=None):
        modal = modal or self._kho_root()
        result_locators = [
            modal.locator(".right-fckEditor .card[id] .hoverDiv"),
            modal.locator(".right-fckEditor tbody tr.odd a.text-gray-800"),
            modal.locator(".right-fckEditor .card[id]"),
            modal.locator(".right-fckEditor tbody tr.odd"),
            modal.locator(".card[id] .hoverDiv"),
            modal.locator("tbody tr.odd a.text-gray-800"),
            modal.locator(".card[id]"),
            modal.locator("tbody tr.odd"),
        ]
        for locator in result_locators:
            try:
                candidate = self._first_visible(locator)
                if candidate is not None:
                    return candidate
            except Exception:
                pass
        return None

    def _file_appears_in_kho(self, file_name: str, attempts: int = 15, refresh_every: int = 5) -> bool:
        modal = self._kho_root()
        search_terms = [file_name, Path(file_name).stem]
        search_input = self._kho_file_search_input(modal)

        for attempt in range(attempts):
            try:
                search_input.fill("")
                search_input.fill(file_name)
                self._trigger_kho_search(modal)
            except Exception:
                pass

            for term in search_terms:
                try:
                    if modal.get_by_text(term, exact=False).count() > 0:
                        return True
                except Exception:
                    pass

            for selector in [
                f"[nztooltiptitle='{file_name}']",
                f"[ng-reflect-nz-tooltip-title='{file_name}']",
                f"[title='{file_name}']",
            ]:
                try:
                    if modal.locator(selector).count() > 0:
                        return True
                except Exception:
                    pass

            if self._first_visible_kho_result_item(modal) is not None:
                return True

            if refresh_every > 0 and attempt % refresh_every == refresh_every - 1:
                for candidate in self._kho_refresh_candidates(modal):
                    try:
                        refresh_btn = self._first_visible(candidate)
                        if refresh_btn is None:
                            continue
                        refresh_btn.click(force=True, timeout=2000)
                        time.sleep(1)
                        self._trigger_kho_search(modal)
                        break
                    except Exception:
                        pass

            time.sleep(1.5)
        return False

    def _kho_root(self):
        assert self.page is not None
        modal_title = "Kho d\u1eef li\u1ec7u"
        container_selectors = [
            "div.ant-modal-wrap",
            "div.ant-modal",
            "nz-modal-container",
            "div[role='dialog']",
            "div.cdk-overlay-pane",
            "div.modal.show",
            "div.modal",
        ]

        for selector in container_selectors:
            candidate = self._first_visible(self.page.locator(selector).filter(has_text=modal_title))
            if candidate is not None:
                return candidate

        title_candidates = [
            self.page.get_by_text(modal_title, exact=True),
            self.page.locator("h1, h2, h3, h4, .modal-title, .ant-modal-title").filter(has_text=modal_title),
        ]
        ancestor_selectors = [
            "xpath=ancestor::div[contains(@class,'ant-modal-wrap')][1]",
            "xpath=ancestor::div[contains(@class,'ant-modal')][1]",
            "xpath=ancestor::nz-modal-container[1]",
            "xpath=ancestor::*[@role='dialog'][1]",
            "xpath=ancestor::div[contains(@class,'cdk-overlay-pane')][1]",
            "xpath=ancestor::div[contains(@class,'modal')][1]",
        ]

        for title in title_candidates:
            visible_title = self._first_visible(title)
            if visible_title is None:
                continue
            for selector in ancestor_selectors:
                candidate = self._first_visible(visible_title.locator(selector))
                if candidate is not None:
                    return candidate
        modal = self.page.locator("text=Kho dữ liệu").locator(
            "xpath=ancestor::div[contains(@class,'modal') or contains(@class,'ant-modal')][1]"
        )
        if modal.count() > 0:
            return modal
        return self.page.locator("body")

    def debug_dump_kho(self) -> None:
        root = self._kho_root()
        try:
            self.screenshot("debug_kho_root.png")
        except Exception:
            pass

        try:
            html = root.first.inner_html()
            (SCREENSHOT_DIR / "debug_kho_root.html").write_text(html, encoding="utf-8")
        except Exception:
            pass

        try:
            txt = root.first.inner_text()
            (SCREENSHOT_DIR / "debug_kho_root.txt").write_text(txt, encoding="utf-8")
        except Exception:
            pass

    def debug_list_visible_texts(self, keyword: str = "") -> None:
        root = self._kho_root()
        lines = []
        all_nodes = root.locator("*")
        count = min(all_nodes.count(), 500)
        for i in range(count):
            try:
                node = all_nodes.nth(i)
                if not node.is_visible():
                    continue
                txt = clean_text(node.inner_text())
                if not txt:
                    continue
                if keyword and keyword.lower() not in txt.lower():
                    continue
                lines.append(txt)
            except Exception:
                pass

        (SCREENSHOT_DIR / "debug_visible_texts.txt").write_text(
            "\n".join(dict.fromkeys(lines)),
            encoding="utf-8",
        )

    def debug_dump_upload_popup(self, popup) -> None:
        try:
            html = popup.inner_html()
            (SCREENSHOT_DIR / "debug_upload_popup.html").write_text(html, encoding="utf-8")
        except Exception:
            pass
        try:
            txt = popup.inner_text()
            (SCREENSHOT_DIR / "debug_upload_popup.txt").write_text(txt, encoding="utf-8")
        except Exception:
            pass

    # -------------------------
    # Form helpers
    # -------------------------

    def fill_so_ky_hieu(self, value: str) -> None:
        assert self.page is not None
        input_box = self.page.locator("input[formcontrolname='officialNumber']").first
        input_box.fill("")
        input_box.fill(value)

    def set_fixed_fields(self, loai_qc_tc: str) -> None:
        assert self.page is not None

        self.page.locator("nz-select[formcontrolname='idTypeOfDocument']").first.click(force=True)
        self.page.locator("nz-option-item, .ant-select-item-option").filter(
            has_text=FIXED_FORM_VALUES["loai_van_ban"]
        ).first.click(force=True)

        self.page.locator("nz-select[formcontrolname='vanBanQCTC']").first.click(force=True)
        target_qctc = "Văn bản quy chuẩn" if "quy chuẩn" in (loai_qc_tc or "").lower() else "Văn bản tiêu chuẩn"
        self.page.locator("nz-option-item, .ant-select-item-option").filter(
            has_text=target_qctc
        ).first.click(force=True)

        selects = self.page.locator("nz-table nz-select")
        if selects.count() >= 2:
            selects.nth(0).click(force=True)
            self.page.locator("nz-option-item, .ant-select-item-option").filter(
                has_text=FIXED_FORM_VALUES["co_quan_ban_hanh"]
            ).first.click(force=True)

            selects.nth(1).click(force=True)
            self.page.locator("nz-option-item, .ant-select-item-option").filter(
                has_text=FIXED_FORM_VALUES["nguoi_ky"]
            ).first.click(force=True)

        row_inputs = self.page.locator("nz-table input")
        if row_inputs.count() >= 1:
            row_inputs.nth(0).fill("")
            row_inputs.nth(0).fill(FIXED_FORM_VALUES["chuc_danh"])

    def fill_trich_yeu(self, value: str) -> None:
        assert self.page is not None
        textarea = self.page.locator("textarea[formcontrolname='source']").first
        textarea.fill("")
        textarea.fill(value)

    def fill_noi_dung(self, html: str, plain_text: str) -> None:
        assert self.page is not None
        editor = self.page.locator(
            ".ck-content[contenteditable='true'], .ck-editor__editable[contenteditable='true'], .ql-editor, [contenteditable='true']"
        ).first
        try:
            editor.wait_for(state="visible", timeout=5000)
            editor.click(force=True)
            editor.evaluate(
                """
                (el, value) => {
                    el.innerHTML = value;
                    el.dispatchEvent(new Event('input', { bubbles: true }));
                    el.dispatchEvent(new Event('change', { bubbles: true }));
                    el.dispatchEvent(new KeyboardEvent('keyup', { bubbles: true }));
                }
                """,
                html or f"<p>{plain_text}</p>",
            )
        except PlaywrightTimeoutError:
            logging.warning("Không tìm thấy contenteditable editor rõ ràng, thử dán plain text bằng keyboard.")
            self.page.keyboard.press("Control+A")
            self.page.keyboard.type(plain_text[:20000], delay=5)

    # -------------------------
    # LĨNH VỰC
    # -------------------------

    def _fields_select(self):
        assert self.page is not None
        return self.page.locator("nz-select[formcontrolname='fields']").first

    def get_selected_fields(self) -> list[str]:
        root = self._fields_select()
        tags = root.locator(".ant-select-selection-item")
        results: list[str] = []
        for i in range(tags.count()):
            text = clean_text(tags.nth(i).inner_text()).replace("×", "").strip()
            if text:
                results.append(text)
        return results

    def remove_field(self, field_name: str) -> None:
        root = self._fields_select()
        item = root.locator(".ant-select-selection-item", has_text=field_name).first
        close_icon = item.locator(".ant-select-selection-item-remove, .anticon-close, [aria-label='close']").first
        if close_icon.count() > 0:
            close_icon.click(force=True)
        else:
            item.click(force=True)
            self.page.keyboard.press("Backspace")

    def add_field(self, field_name: str) -> None:
        root = self._fields_select()
        root.click(force=True)

        search_input = root.locator("input.ant-select-selection-search-input").first
        try:
            search_input.fill(field_name)
        except Exception:
            pass

        option = self.page.locator(".ant-select-item-option, nz-option-item").filter(
            has_text=field_name
        ).first
        option.wait_for(state="visible")
        option.click(force=True)
        time.sleep(0.2)

    def sync_fields(self, target_fields: list[str]) -> None:
        current_fields = self.get_selected_fields()
        current_set = set(current_fields)
        target_set = set(target_fields)

        for field in sorted(current_set - target_set):
            logging.info("Bỏ lĩnh vực thừa: %s", field)
            self.remove_field(field)

        for field in sorted(target_set - current_set):
            logging.info("Thêm lĩnh vực mới: %s", field)
            self.add_field(field)

    # -------------------------
    # FILE ĐÍNH KÈM
    # -------------------------

    def get_selected_file_names(self) -> list[str]:
        assert self.page is not None
        names: list[str] = []
        rows = self.page.locator("app-file-attachment .scroll-y .m-0")
        for i in range(rows.count()):
            text = clean_text(rows.nth(i).inner_text())
            if text:
                names.append(text)
        return names

    def remove_all_selected_files(self) -> None:
        assert self.page is not None
        while True:
            trash_buttons = self.page.locator("app-file-attachment .ki-trash").locator("xpath=ancestor::a[1]")
            if trash_buttons.count() == 0:
                break
            trash_buttons.first.click(force=True)
            time.sleep(0.3)

    def open_file_modal(self) -> None:
        assert self.page is not None
        if self._is_kho_modal_open():
            return
        btn = self.page.locator("app-file-attachment a.btn").filter(has_text="Chọn file").first
        btn.wait_for(state="visible", timeout=30000)
        btn.click(force=True)
        self.page.wait_for_timeout(1500)
        self.debug_dump_kho()
        self.debug_list_visible_texts()

    def close_file_modal(self) -> None:
        assert self.page is not None
        if not self._is_kho_modal_open():
            return
        modal = self._kho_root()
        close_candidates = self._modal_footer_action_candidates(modal, "Đóng") + [
            modal.locator("button[aria-label='Close']"),
            modal.locator(".ant-modal-close"),
        ]
        for candidate in close_candidates:
            try:
                close_btn = self._first_visible(candidate)
                if close_btn is None:
                    continue
                close_btn.click(force=True, timeout=2000)
                time.sleep(0.8)
                if not self._is_kho_modal_open():
                    return
            except Exception:
                pass

    def select_folder_in_kho(self, folder_name: str = "files") -> None:
        modal = self._kho_root()
        if self._click_text_in_root(modal, folder_name):
            return

        candidates = modal.get_by_text(folder_name, exact=True)
        count = candidates.count()

        for i in range(count):
            try:
                item = candidates.nth(i)
                if item.is_visible():
                    item.click(force=True)
                    time.sleep(0.5)
                    return
            except Exception:
                pass

        fuzzy = modal.locator(f":text('{folder_name}')")
        for i in range(fuzzy.count()):
            try:
                item = fuzzy.nth(i)
                if item.is_visible():
                    item.click(force=True)
                    time.sleep(0.5)
                    return
            except Exception:
                pass

        self.screenshot("debug_kho_modal_no_files.png")
        self.debug_dump_kho()
        self.debug_list_visible_texts(folder_name)
        raise RuntimeError(f"Không tìm thấy folder '{folder_name}' trong Kho dữ liệu.")

    def open_upload_menu_for_current_folder(self) -> None:
        modal = self._kho_root()
        self.select_folder_in_kho("files")
        upload_menu_item_texts = ["T\u1ea3i file l\u00ean", "Upload file"]
        upload_icon_selectors = [
            "[nztype='upload']",
            "[data-icon='upload']",
            "[class*='upload']",
            "[title*='T\u1ea3i file']",
            "[aria-label*='T\u1ea3i file']",
            "svg[data-icon='upload']",
            "i[class*='upload']",
        ]
        folder_rows = []

        for matches in self._matching_text_locators(modal, "files"):
            for i in range(min(matches.count(), 10)):
                try:
                    item = matches.nth(i)
                    if not item.is_visible():
                        continue
                    for selector in [
                        "xpath=ancestor-or-self::*[@role='treeitem'][1]",
                        "xpath=ancestor-or-self::li[1]",
                        "xpath=ancestor-or-self::tr[1]",
                        "xpath=ancestor-or-self::div[1]",
                    ]:
                        candidate = self._first_visible(item.locator(selector))
                        if candidate is not None:
                            folder_rows.append(candidate)
                    folder_rows.append(item)
                except Exception:
                    pass

        for row in folder_rows:
            for selector in upload_icon_selectors:
                loc = row.locator(selector)
                for i in range(min(loc.count(), 10)):
                    try:
                        el = loc.nth(i)
                        if not el.is_visible():
                            continue
                        el.click(force=True, timeout=1000)
                        time.sleep(0.5)
                        for text in upload_menu_item_texts:
                            try:
                                menu_item = self.page.get_by_text(text, exact=True).first
                                if menu_item.is_visible():
                                    menu_item.click(force=True)
                                    time.sleep(0.8)
                                    return
                            except Exception:
                                pass
                    except Exception:
                        pass

        for text in upload_menu_item_texts:
            try:
                menu_item = self.page.get_by_text(text, exact=True).first
                if menu_item.is_visible():
                    menu_item.click(force=True)
                    time.sleep(0.8)
                    return
            except Exception:
                pass
        menu_text = "T\u1ea3i file l\u00ean"
        search_roots = []

        for matches in self._matching_text_locators(modal, "files"):
            for i in range(min(matches.count(), 10)):
                try:
                    item = matches.nth(i)
                    if not item.is_visible():
                        continue
                    search_roots.append(item)
                    for selector in [
                        "xpath=ancestor-or-self::*[@role='treeitem'][1]",
                        "xpath=ancestor-or-self::li[1]",
                        "xpath=ancestor-or-self::tr[1]",
                        "xpath=ancestor-or-self::div[contains(@class,'tree')][1]",
                        "xpath=ancestor-or-self::div[1]",
                    ]:
                        candidate = self._first_visible(item.locator(selector))
                        if candidate is not None:
                            search_roots.append(candidate)
                except Exception:
                    pass

        search_roots.append(modal)
        selectors = [
            "[nztype='upload']",
            "[data-icon='upload']",
            "[class*='upload']",
            "[title*='T\u1ea3i file']",
            "[aria-label*='T\u1ea3i file']",
            "i",
            "a",
            "button",
            "span[role='button']",
        ]

        for root in search_roots:
            for selector in selectors:
                loc = root.locator(selector)
                for i in range(min(loc.count(), 30)):
                    try:
                        el = loc.nth(i)
                        if not el.is_visible():
                            continue
                        el.click(force=True, timeout=1000)
                        time.sleep(0.5)
                        if self.page.get_by_text(menu_text, exact=True).count() > 0:
                            menu_item = self.page.get_by_text(menu_text, exact=True).first
                            menu_item.wait_for(state="visible", timeout=10000)
                            menu_item.click(force=True)
                            time.sleep(0.8)
                            return
                    except Exception:
                        pass

        for matches in self._matching_text_locators(modal, "files"):
            for i in range(min(matches.count(), 10)):
                try:
                    item = matches.nth(i)
                    if not item.is_visible():
                        continue
                    item.click(button="right", force=True, timeout=1000)
                    time.sleep(0.5)
                    if self.page.get_by_text(menu_text, exact=True).count() > 0:
                        menu_item = self.page.get_by_text(menu_text, exact=True).first
                        menu_item.wait_for(state="visible", timeout=10000)
                        menu_item.click(force=True)
                        time.sleep(0.8)
                        return
                except Exception:
                    pass

        left_panel = modal.locator("div").nth(0)

        selectors = [
            "[nztype='upload']",
            "[data-icon='upload']",
            "[class*='upload']",
            "[title*='T\u1ea3i file']",
            "[aria-label*='T\u1ea3i file']",
            "i",
            "a",
            "button",
        ]

        clicked = False
        for selector in selectors:
            loc = left_panel.locator(selector)
            for i in range(min(loc.count(), 30)):
                try:
                    el = loc.nth(i)
                    if not el.is_visible():
                        continue
                    el.click(force=True, timeout=1000)
                    time.sleep(0.5)
                    if self.page.get_by_text("Tải file lên", exact=True).count() > 0:
                        clicked = True
                        break
                    if self.page.get_by_text("Upload file", exact=True).count() > 0:
                        clicked = True
                        break
                    if modal.get_by_text("Tải file cho thư mục này", exact=True).count() > 0:
                        clicked = True
                        break
                except Exception:
                    pass
            if clicked:
                break

        if not clicked:
            self.screenshot("debug_kho_modal_menu_not_found.png")
            self.debug_dump_kho()
            raise RuntimeError("Không mở được menu thao tác của folder 'files' trong Kho dữ liệu.")

        for text in ["Tải file lên", "Upload file", "Tải file cho thư mục này"]:
            try:
                menu_item = self.page.get_by_text(text, exact=True).first
                menu_item.wait_for(state="visible", timeout=3000)
                menu_item.click(force=True)
                time.sleep(0.8)
                return
            except Exception:
                pass
        raise RuntimeError("Mở được menu nhưng không thấy mục upload file.")

    def search_file_in_modal(self, file_name: str) -> None:
        modal = self._kho_root()
        search_input = self._kho_file_search_input(modal)
        search_input.fill("")
        search_input.fill(file_name)
        self._trigger_kho_search(modal)
        time.sleep(1)

    def select_file_card_in_modal(self, file_name: str) -> None:
        modal = self._kho_root()
        for selector in [
            f"[nztooltiptitle='{file_name}']",
            f"[ng-reflect-nz-tooltip-title='{file_name}']",
            f"[title='{file_name}']",
        ]:
            try:
                candidate = self._first_visible(modal.locator(selector))
                if candidate is not None:
                    candidate.click(force=True)
                    return
            except Exception:
                pass

        result_item = self._first_visible_kho_result_item(modal)
        if result_item is not None:
            result_item.click(force=True)
            return

        card = modal.locator("text=" + file_name).first
        card.wait_for(state="visible")
        card.click(force=True)

    def confirm_selected_files_in_modal(self) -> None:
        assert self.page is not None
        self.page.locator("button, a").filter(has_text="Chọn").last.click(force=True)
        time.sleep(1)

    def attach_existing_files(self, file_names: list[str]) -> None:
        self.remove_all_selected_files()
        self.open_file_modal()
        for file_name in file_names:
            logging.info("Chọn file từ kho: %s", file_name)
            self.search_file_in_modal(file_name)
            self.select_file_card_in_modal(file_name)
            time.sleep(0.2)
        self.confirm_selected_files_in_modal()

    def upload_file_to_kho(self, file_path: Path) -> None:
        assert self.page is not None

        self.open_file_modal()
        if self._file_appears_in_kho(file_path.name, attempts=2, refresh_every=0):
            logging.info("File đã có sẵn trong kho, bỏ qua upload: %s", file_path.name)
            return

        self.open_upload_menu_for_current_folder()
        modal = self._kho_root()

        upload_popup = self._upload_popup_root()
        if upload_popup is None:
            self.screenshot("debug_upload_popup_not_found.png")
            self.debug_dump_kho()
            raise RuntimeError(f"Kh\u00f4ng t\u00ecm th\u1ea5y popup upload cho: {file_path.name}")
        self.debug_dump_upload_popup(upload_popup)

        clicked_file_chooser = False
        file_input_candidates = [
            upload_popup.locator(".ant-upload input[type='file']"),
            upload_popup.locator("input[type='file'][accept]"),
            upload_popup.locator("input[type='file']"),
        ]
        for file_input in file_input_candidates:
            try:
                if file_input.count() == 0:
                    continue
                for i in range(min(file_input.count(), 5)):
                    try:
                        file_input.nth(i).set_input_files(str(file_path.resolve()))
                        time.sleep(0.5)
                        clicked_file_chooser = True
                        try:
                            if upload_popup.get_by_text(file_path.name, exact=False).count() > 0:
                                break
                        except Exception:
                            pass
                        break
                    except Exception:
                        pass
                if clicked_file_chooser:
                    break
            except Exception:
                pass

        if not clicked_file_chooser:
            chooser_targets = [
                upload_popup.locator(".ant-upload.ant-upload-drag").first,
                upload_popup.locator(".ant-upload").first,
                upload_popup.get_by_text("Chọn file để upload", exact=True).first,
            ]
            for target in chooser_targets:
                try:
                    if target.count() == 0:
                        continue
                    with self.page.expect_file_chooser(timeout=3000) as fc_info:
                        target.click(force=True)
                    fc_info.value.set_files(str(file_path.resolve()))
                    clicked_file_chooser = True
                    time.sleep(0.5)
                    break
                except Exception:
                    pass

        if not clicked_file_chooser:
            self.screenshot("debug_upload_popup_no_filechooser.png")
            raise RuntimeError(f"Không gán được file vào input upload cho: {file_path.name}")

        file_name_visible = False
        for _ in range(10):
            try:
                if upload_popup.get_by_text(file_path.name, exact=False).count() > 0:
                    file_name_visible = True
                    break
            except Exception:
                pass
            time.sleep(0.5)

        if not file_name_visible:
            logging.warning("Popup upload chua hien thi ten file, van thu upload: %s", file_path.name)
            self.screenshot("debug_upload_popup_file_not_visible.png")

        upload_enabled = False
        for _ in range(10):
            for button_text in ["Upload", "T\u1ea3i l\u00ean", "T\u1ea3i file l\u00ean"]:
                try:
                    for candidate in self._modal_footer_action_candidates(upload_popup, button_text):
                        visible_btn = self._first_visible(candidate)
                        if visible_btn is None:
                            continue
                        if self._is_effectively_enabled(visible_btn):
                            upload_enabled = True
                            break
                    if upload_enabled:
                        break
                except Exception:
                    pass
            if upload_enabled:
                break
            time.sleep(0.5)

        if not upload_enabled:
            self.screenshot("debug_upload_button_disabled.png")
            raise RuntimeError(f"Nút upload vẫn đang bị disable sau khi chọn file: {file_path.name}")

        time.sleep(0.5)

        uploaded = False
        for button_text in ["Upload", "T\u1ea3i l\u00ean", "T\u1ea3i file l\u00ean"]:
            try:
                button_candidates = self._modal_footer_action_candidates(upload_popup, button_text)
                for candidate in button_candidates:
                    visible_btn = self._first_visible(candidate)
                    if visible_btn is None:
                        continue
                    if not self._is_effectively_enabled(visible_btn):
                        continue
                    visible_btn.click(force=True)
                    time.sleep(2)
                    uploaded = True
                    break
                if uploaded:
                    break
            except Exception:
                pass
        if not uploaded:
            self.screenshot("debug_upload_submit_not_found.png")
            raise RuntimeError(f"Kh\u00f4ng t\u00ecm th\u1ea5y n\u00fat upload cho: {file_path.name}")

        popup_closed = False
        for _ in range(20):
            try:
                if upload_popup.count() == 0 or not upload_popup.first.is_visible():
                    popup_closed = True
                    break
            except Exception:
                popup_closed = True
                break
            time.sleep(1)

        if not popup_closed:
            logging.warning("Popup upload không tự đóng sau khi submit, thử đóng thủ công: %s", file_path.name)
            close_candidates = self._modal_footer_action_candidates(upload_popup, "Thoát") + [
                upload_popup.locator("button[aria-label='Close']"),
                upload_popup.locator(".ant-modal-close"),
            ]
            for candidate in close_candidates:
                try:
                    close_btn = self._first_visible(candidate)
                    if close_btn is None:
                        continue
                    close_btn.click(force=True, timeout=2000)
                    popup_closed = True
                    break
                except Exception:
                    pass
            time.sleep(1)

        if not self._file_appears_in_kho(file_path.name, attempts=10, refresh_every=3):
            logging.info("Không thấy file ngay trong modal hiện tại, thử mở lại Kho dữ liệu: %s", file_path.name)
            self.close_file_modal()
            self.open_file_modal()
            self.select_folder_in_kho("files")

        if not self._file_appears_in_kho(file_path.name, attempts=15, refresh_every=3):
            self.screenshot("debug_upload_finish_timeout.png")
            raise RuntimeError(f"Upload không hoàn tất hoặc không thấy file sau khi upload: {file_path.name}")
        logging.info("Đã upload xong file vào kho: %s", file_path.name)

    def ensure_files_uploaded(self, docs: list[SourceDocument]) -> None:
        uploaded_state = load_state("uploaded_files", {})
        self.goto_create_form()
        self.open_file_modal()
        try:
            for doc in docs:
                for att in doc.attachments:
                    if not att.local_path:
                        continue
                    local_path = Path(att.local_path)
                    if not local_path.exists():
                        logging.warning("Thiếu file local: %s", local_path)
                        continue
                    if uploaded_state.get(local_path.name):
                        continue
                    logging.info("Upload sẵn file vào kho: %s", local_path.name)
                    self.upload_file_to_kho(local_path)
                    uploaded_state[local_path.name] = True
                    save_state("uploaded_files", uploaded_state)
        finally:
            self.close_file_modal()

    # -------------------------
    # LƯU FORM
    # -------------------------

    def save_document(self) -> None:
        assert self.page is not None
        self.page.locator("button, a").filter(has_text="Lưu văn bản").first.click(force=True)
        time.sleep(2)

    def fill_document_form(self, doc: SourceDocument) -> None:
        logging.info("Điền form cho: %s | %s", doc.so_ky_hieu, doc.title)
        self.goto_create_form()
        self.set_fixed_fields(doc.loai_qc_tc or "Tiêu chuẩn")
        self.fill_so_ky_hieu(doc.so_ky_hieu)
        self.fill_trich_yeu(doc.trich_yeu)
        self.fill_noi_dung(doc.noi_dung_html, doc.noi_dung_text)
        self.sync_fields(doc.linh_vuc_mapped)
        self.attach_existing_files([Path(a.local_path).name for a in doc.attachments if a.local_path])

    def import_documents(self, docs: list[SourceDocument], dry_run: bool = False) -> None:
        imported_state = load_state("imported_documents", {})

        ordered_docs = list(reversed(docs))
        ordered_docs = ordered_docs[CRAWL_CONFIG["skip_last_already_imported"] :]

        logging.info(
            "Sẽ xử lý %s văn bản sau khi đảo ngược và bỏ qua %s văn bản cuối đã nhập.",
            len(ordered_docs),
            CRAWL_CONFIG["skip_last_already_imported"],
        )

        for idx, doc in enumerate(ordered_docs, start=1):
            key = str(doc.item_id)
            if imported_state.get(key) == "success":
                logging.info("Bỏ qua item_id=%s vì đã import thành công trước đó.", doc.item_id)
                continue

            try:
                logging.info("[%s/%s] Import item_id=%s", idx, len(ordered_docs), doc.item_id)
                self.fill_document_form(doc)
                self.screenshot(f"before_save_{doc.item_id}.png")
                if not dry_run:
                    self.save_document()
                imported_state[key] = "success"
                save_state("imported_documents", imported_state)
                time.sleep(1)
            except Exception as exc:
                logging.exception("Import lỗi item_id=%s: %s", doc.item_id, exc)
                self.screenshot(f"error_{doc.item_id}.png")
                imported_state[key] = {"status": "error", "message": str(exc)}
                save_state("imported_documents", imported_state)
                raise


# =========================
# NẠP / CHUYỂN ĐỔI DỮ LIỆU
# =========================

def load_documents_from_json() -> list[SourceDocument]:
    raw = read_json(OUTPUT_JSON, [])
    docs: list[SourceDocument] = []
    for item in raw:
        attachments = [Attachment(**att) for att in item.get("attachments", [])]
        docs.append(
            SourceDocument(
                item_id=item["item_id"],
                source_url=item["source_url"],
                title=item.get("title", ""),
                so_ky_hieu=item.get("so_ky_hieu", ""),
                loai_qc_tc=item.get("loai_qc_tc", ""),
                linh_vuc_raw=item.get("linh_vuc_raw", []),
                linh_vuc_mapped=item.get("linh_vuc_mapped", []),
                trich_yeu=item.get("trich_yeu", ""),
                noi_dung_text=item.get("noi_dung_text", ""),
                noi_dung_html=item.get("noi_dung_html", ""),
                attachments=attachments,
            )
        )
    return docs


def export_summary_csv(docs: list[SourceDocument]) -> Path:
    csv_path = DATA_DIR / "documents_summary.csv"

    normal_docs: list[SourceDocument] = []
    missing_docs: list[SourceDocument] = []
    for doc in docs:
        if doc.so_ky_hieu:
            normal_docs.append(doc)
        else:
            missing_docs.append(doc)

    ordered_docs = normal_docs + missing_docs
    lines = ["item_id,so_ky_hieu,loai_qc_tc,linh_vuc_mapped,file_count,source_url"]
    for doc in ordered_docs:
        display_so_ky_hieu = doc.so_ky_hieu or f"[THIẾU] {doc.title}"
        line = '"{}","{}","{}","{}",{},"{}"'.format(
            doc.item_id,
            display_so_ky_hieu.replace('"', '""'),
            (doc.loai_qc_tc or "").replace('"', '""'),
            "; ".join(doc.linh_vuc_mapped).replace('"', '""'),
            len(doc.attachments),
            doc.source_url.replace('"', '""'),
        )
        lines.append(line)

    csv_path.write_text("\n".join(lines), encoding="utf-8-sig")
    logging.info("CSV: %s bản ghi OK, %s bản ghi thiếu số ký hiệu", len(normal_docs), len(missing_docs))
    return csv_path


def export_missing_so_ky_hieu(docs: list[SourceDocument]) -> None:
    missing = [doc for doc in docs if not doc.so_ky_hieu]
    if not missing:
        logging.info("Không có bản ghi thiếu số ký hiệu")
        return

    path = STATE_DIR / "missing_so_ky_hieu.json"
    data = [{"item_id": doc.item_id, "title": doc.title, "url": doc.source_url} for doc in missing]
    write_json(path, data)
    logging.warning("Có %s văn bản thiếu số ký hiệu → %s", len(missing), path)


# =========================
# CLI
# =========================

def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Crawl + upload + import văn bản QC-TC")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("crawl", help="Crawl dữ liệu và tải file từ vr.org.vn")
    sub.add_parser("init-field-map", help="Tạo file field_mapping.json mẫu để bổ sung đủ 48 lĩnh vực")

    p_upload = sub.add_parser("upload-files", help="Upload sẵn toàn bộ file vào Kho dữ liệu")
    p_upload.add_argument("--interactive-login", action="store_true")
    p_upload.add_argument("--headless", action="store_true")
    p_upload.add_argument("--slow-mo", type=int, default=100)

    p_import = sub.add_parser("import", help="Nhập văn bản vào hệ thống từ cuối lên")
    p_import.add_argument("--interactive-login", action="store_true")
    p_import.add_argument("--headless", action="store_true")
    p_import.add_argument("--slow-mo", type=int, default=100)
    p_import.add_argument("--dry-run", action="store_true")

    sub.add_parser("summary", help="Xuất file CSV tóm tắt để rà soát trước khi import")
    return parser


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()
    setup_logging(verbose=True)
    export_default_field_mapping()

    if args.command == "init-field-map":
        export_default_field_mapping()
        logging.info("Đã tạo/giữ nguyên file mapping tại: %s", FIELD_MAP_PATH)
        return

    if args.command == "crawl":
        docs = crawl_all_documents()
        csv_path = export_summary_csv(docs)
        export_missing_so_ky_hieu(docs)
        logging.info("Đã xuất summary: %s", csv_path)
        return

    docs = load_documents_from_json()
    if not docs:
        raise SystemExit("Chưa có documents.json. Hãy chạy lệnh crawl trước.")

    if args.command == "summary":
        csv_path = export_summary_csv(docs)
        logging.info("Đã xuất summary: %s", csv_path)
        return

    with EOfficeBot(
        headless=args.headless,
        slow_mo=args.slow_mo,
        interactive_login=args.interactive_login,
    ) as bot:
        if args.command == "upload-files":
            bot.ensure_files_uploaded(docs)
            logging.info("Đã upload xong file vào Kho dữ liệu.")
            return

        if args.command == "import":
            bot.import_documents(docs, dry_run=args.dry_run)
            logging.info("Đã import xong batch văn bản.")
            return


if __name__ == "__main__":
    main()
