
import os
import re
import threading
import html as html_lib
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
from urllib.parse import urlparse

import requests
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

DEFAULT_URL = "https://www.dm5.cn/m1337660/"
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)


class DM5CrawlerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("DM5 特化逐頁下載")
        self.root.geometry("1120x860")

        self.urls: list[str] = []
        self.page_images: dict[int, list[str]] = {}
        self.downloaded_urls: set[str] = set()
        self.is_running = False

        self.url_var = tk.StringVar(value=DEFAULT_URL)
        self.status_var = tk.StringVar(value="就緒")
        self.scroll_times_var = tk.StringVar(value="8")
        self.scroll_wait_var = tk.StringVar(value="1200")
        self.timeout_var = tk.StringVar(value="30000")
        self.save_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "downloads"))
        self.min_width_var = tk.StringVar(value="300")
        self.min_height_var = tk.StringVar(value="300")
        self.start_page_var = tk.StringVar(value="1")
        self.end_page_var = tk.StringVar(value="5")
        self.use_title_chapter_dir_var = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        ttk.Label(top, text="DM5 網址").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.url_var, width=110).grid(
            row=0, column=1, columnspan=7, sticky="ew", padx=6
        )

        ttk.Label(top, text="起始頁").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.start_page_var, width=8).grid(
            row=1, column=1, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="結束頁").grid(row=1, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.end_page_var, width=8).grid(
            row=1, column=3, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="最小寬").grid(row=1, column=4, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.min_width_var, width=8).grid(
            row=1, column=5, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="最小高").grid(row=1, column=6, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.min_height_var, width=8).grid(
            row=1, column=7, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="捲動次數").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.scroll_times_var, width=8).grid(
            row=2, column=1, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="每次等待(ms)").grid(row=2, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.scroll_wait_var, width=10).grid(
            row=2, column=3, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="頁面 timeout(ms)").grid(row=2, column=4, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.timeout_var, width=12).grid(
            row=2, column=5, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="下載資料夾").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.save_dir_var, width=90).grid(
            row=3, column=1, columnspan=5, sticky="ew", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="選擇", command=self.choose_dir).grid(
            row=3, column=6, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(
            top,
            text="自動建立資料夾（標題/篇章）",
            variable=self.use_title_chapter_dir_var,
        ).grid(row=3, column=7, sticky="w", padx=6, pady=(8, 0))

        ttk.Button(top, text="顯示自動轉換範本", command=self.preview_template).grid(
            row=4, column=3, sticky="e", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="自動抓最大頁數", command=self.fetch_max_page).grid(
            row=4, column=4, sticky="e", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="開始找當前頁", command=self.start_find).grid(
            row=4, column=5, sticky="e", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="逐頁下載範圍", command=self.start_download_range).grid(
            row=4, column=6, sticky="e", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="下載目前列表", command=self.start_download_current).grid(
            row=4, column=7, sticky="e", padx=6, pady=(8, 0)
        )

        top.columnconfigure(1, weight=1)

        middle = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        middle.pack(fill="both", expand=True)

        ttk.Label(middle, textvariable=self.status_var).pack(anchor="w", pady=(0, 8))
        self.result_text = ScrolledText(middle, wrap="word")
        self.result_text.pack(fill="both", expand=True)

    def choose_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.save_dir_var.get() or os.getcwd())
        if selected:
            self.save_dir_var.set(selected)

    def append_text(self, text: str) -> None:
        self.result_text.insert("end", text)
        self.result_text.see("end")

    def set_status(self, text: str) -> None:
        self.status_var.set(text)

    def clear_output(self) -> None:
        self.result_text.delete("1.0", "end")

    def get_int(self, value: str, field_name: str) -> int:
        try:
            return int(value.strip())
        except ValueError as exc:
            raise ValueError(f"{field_name} 必須是整數") from exc

    def normalize_dm5_template(self, raw_url: str) -> str:
        url = raw_url.strip()
        if not url:
            raise ValueError("請輸入 DM5 網址")

        url = url.replace("（#）", "(#)")

        if "(#)" in url:
            return url

        m = re.match(r"^(https?://www\.dm5\.cn/m\d+?)/?$", url)
        if m:
            return f"{m.group(1)}-p(#)"

        converted, count = re.subn(r"-p\d+(?=(?:[#/?]|$))", "-p(#)", url)
        if count > 0:
            return converted

        converted, count = re.subn(r"#ipg\d+", "#ipg(#)", url)
        if count > 0:
            return converted

        raise ValueError("只支援 DM5 的作品首頁、-p頁碼網址，或已含 (#) 的範本")

    def build_url_for_page(self, page_num: int) -> str:
        template = self.normalize_dm5_template(self.url_var.get())
        return template.replace("(#)", str(page_num))

    def build_chapter_home_url(self) -> str:
        raw_url = self.url_var.get().strip()
        m = re.search(r"(https?://www\.dm5\.cn/m\d+)/?", raw_url)
        if not m:
            raise ValueError("無法從目前網址判斷 DM5 章節首頁")
        return m.group(1) + "/"

    def preview_template(self) -> None:
        try:
            template = self.normalize_dm5_template(self.url_var.get())
        except ValueError as exc:
            messagebox.showerror("錯誤", str(exc))
            return
        self.append_text(f"自動轉換範本：{template}\n")
        messagebox.showinfo("範本", template)

    def fetch_max_page(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        try:
            template = self.normalize_dm5_template(self.url_var.get())
        except ValueError as exc:
            messagebox.showerror("錯誤", str(exc))
            return
        self.is_running = True
        self.set_status("正在抓取最大頁數...")
        threading.Thread(target=self.fetch_max_page_worker, args=(template,), daemon=True).start()

    def extract_max_page_from_html(self, html: str) -> int:
        m = re.search(r"DM5_IMAGE_COUNT\s*=\s*(\d+)", html, re.IGNORECASE)
        if m:
            return int(m.group(1))

        pager_matches = re.findall(r"/m\d+-p(\d+)/", html, re.IGNORECASE)
        if pager_matches:
            return max(int(x) for x in pager_matches)

        ipg_matches = re.findall(r"#ipg(\d+)", html, re.IGNORECASE)
        if ipg_matches:
            return max(int(x) for x in ipg_matches)

        raise ValueError("找不到 DM5_IMAGE_COUNT 或分頁資訊")

    def fetch_max_page_worker(self, template: str) -> None:
        try:
            page1_url = template.replace("(#)", "1")
            headers = {"User-Agent": UA, "Referer": page1_url}
            response = requests.get(page1_url, headers=headers, timeout=30)
            response.raise_for_status()
            max_page = self.extract_max_page_from_html(response.text)
            self.root.after(0, lambda m=max_page, u=page1_url: self.on_fetch_max_page_success(m, u))
        except Exception as exc:
            err = exc
            self.root.after(0, lambda err=err: self.on_fetch_max_page_error(err))

    def on_fetch_max_page_success(self, max_page: int, page1_url: str) -> None:
        self.is_running = False
        self.end_page_var.set(str(max_page))
        self.set_status(f"已抓到最大頁數：{max_page}")
        self.append_text(f"自動抓最大頁數成功：{max_page}（來源：{page1_url}）\n")

    def on_fetch_max_page_error(self, exc: Exception) -> None:
        self.is_running = False
        self.set_status("抓取最大頁數失敗")
        self.append_text(f"抓取最大頁數失敗：{exc}\n")
        messagebox.showerror("錯誤", str(exc))

    def is_probably_comic_image(self, url: str) -> bool:
        lowered = url.lower()
        if not re.search(r"\.(jpg|jpeg)(\?|$)", lowered):
            return False
        blacklist = ["logo", "avatar", "cover", "banner", "ads", "ad.", "icon", "emoji", "sprite"]
        return not any(word in lowered for word in blacklist)

    def collect_dom_urls(self, page, min_width: int, min_height: int) -> list[str]:
        return page.evaluate(
            """
            ({minWidth, minHeight}) => {
                return [...document.images]
                    .map(img => {
                        const url = img.currentSrc || img.src || "";
                        return {
                            url,
                            w: img.naturalWidth || 0,
                            h: img.naturalHeight || 0,
                            complete: !!img.complete
                        };
                    })
                    .filter(x => x.complete && x.url && x.w >= minWidth && x.h >= minHeight)
                    .map(x => x.url);
            }
            """,
            {"minWidth": min_width, "minHeight": min_height},
        )

    def clean_html_text(self, text: str) -> str:
        text = re.sub(r"<[^>]+>", "", text)
        text = html_lib.unescape(text)
        text = re.sub(r"\s+", " ", text)
        return text.strip()

    def sanitize_path_part(self, text: str) -> str:
        text = self.clean_html_text(text)
        text = re.sub(r'[\\/:*?"<>|]+', "_", text)
        text = text.strip(" .")
        return text[:120] if text else ""

    def extract_title_and_chapter_from_html(self, html: str) -> tuple[str, str]:
        title_text = ""
        chapter_text = ""

        block_match = re.search(
            r'<div class="title">(.*?)</div>\s*<div class="right-bar">',
            html,
            re.IGNORECASE | re.DOTALL,
        )
        block = block_match.group(1) if block_match else html

        pair_match = re.search(
            r'<span class="right-arrow">\s*<a [^>]*>(.*?)</a>\s*</span>\s*'
            r'<span class="active right-arrow">\s*(.*?)\s*</span>',
            block,
            re.IGNORECASE | re.DOTALL,
        )
        if pair_match:
            title_text = self.clean_html_text(pair_match.group(1))
            chapter_text = self.clean_html_text(pair_match.group(2))

        if not title_text:
            m = re.search(
                r'<span class="right-arrow">\s*<a [^>]*title="([^"]+)"[^>]*>',
                block,
                re.IGNORECASE | re.DOTALL,
            )
            if m:
                title_text = self.clean_html_text(m.group(1))

        if not chapter_text:
            m = re.search(
                r'<span class="active right-arrow">\s*(.*?)\s*</span>',
                block,
                re.IGNORECASE | re.DOTALL,
            )
            if m:
                chapter_text = self.clean_html_text(m.group(1))

        if (not title_text or not chapter_text):
            m = re.search(r'DM5_CTITLE\s*=\s*"([^"]+)"', html, re.IGNORECASE)
            if m:
                ctitle = self.clean_html_text(m.group(1))
                m2 = re.search(r"(第\s*\d+\s*话.*)$", ctitle)
                if m2:
                    if not chapter_text:
                        chapter_text = self.clean_html_text(m2.group(1))
                    if not title_text:
                        title_text = self.clean_html_text(ctitle[: m2.start()].strip())
                elif not title_text:
                    title_text = ctitle

        return title_text, chapter_text

    def resolve_save_dir(self, base_dir: str, page_url: str, html_text: str | None = None) -> tuple[str, str, str]:
        final_dir = base_dir
        title_text = ""
        chapter_text = ""

        if self.use_title_chapter_dir_var.get():
            if html_text is None:
                headers = {"User-Agent": UA, "Referer": page_url}
                response = requests.get(page_url, headers=headers, timeout=30)
                response.raise_for_status()
                html_text = response.text

            title_text, chapter_text = self.extract_title_and_chapter_from_html(html_text)
            safe_title = self.sanitize_path_part(title_text)
            safe_chapter = self.sanitize_path_part(chapter_text)

            if safe_title and safe_chapter:
                final_dir = os.path.join(base_dir, safe_title, safe_chapter)
            elif safe_title:
                final_dir = os.path.join(base_dir, safe_title)

        os.makedirs(final_dir, exist_ok=True)
        return final_dir, title_text, chapter_text

    def prepare_single_find(self) -> tuple[str, int, int, int, int, int]:
        template = self.normalize_dm5_template(self.url_var.get())
        scroll_times = self.get_int(self.scroll_times_var.get(), "捲動次數")
        scroll_wait = self.get_int(self.scroll_wait_var.get(), "每次等待")
        timeout = self.get_int(self.timeout_var.get(), "頁面 timeout")
        min_width = self.get_int(self.min_width_var.get(), "最小寬")
        min_height = self.get_int(self.min_height_var.get(), "最小高")
        return template, scroll_times, scroll_wait, timeout, min_width, min_height

    def start_find(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        try:
            self.prepare_single_find()
        except ValueError as exc:
            messagebox.showerror("錯誤", str(exc))
            return
        self.is_running = True
        self.clear_output()
        self.set_status("正在找當前頁圖片...")
        threading.Thread(target=self.find_images_worker, daemon=True).start()

    def find_images_worker(self) -> None:
        try:
            template, scroll_times, scroll_wait, timeout, min_width, min_height = self.prepare_single_find()
            current_url = self.build_url_for_page(self.get_int(self.start_page_var.get(), "起始頁"))
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False)
                context = browser.new_context(viewport={"width": 1400, "height": 900}, user_agent=UA)
                page = context.new_page()
                page.set_default_timeout(timeout)
                page.goto(current_url, wait_until="domcontentloaded")
                page.wait_for_timeout(3000)
                for _ in range(scroll_times):
                    page.mouse.wheel(0, 2000)
                    page.wait_for_timeout(scroll_wait)
                dom_urls = self.collect_dom_urls(page, min_width, min_height)
                browser.close()
            seen = set()
            self.urls = [u for u in dom_urls if self.is_probably_comic_image(u) and not (u in seen or seen.add(u))]
            self.root.after(0, self.on_find_success)
        except Exception as exc:
            err = exc
            self.root.after(0, lambda err=err: self.on_find_error(err))

    def on_find_success(self) -> None:
        self.is_running = False
        self.set_status(f"找圖完成，共 {len(self.urls)} 張")
        self.append_text(f"找到 {len(self.urls)} 個 jpg 候選網址\n")
        self.append_text("-" * 80 + "\n")
        for i, url in enumerate(self.urls, 1):
            self.append_text(f"{i}. {url}\n")
        if not self.urls:
            messagebox.showwarning("結果", "沒有找到符合條件的 jpg 圖片")

    def on_find_error(self, exc: Exception) -> None:
        self.is_running = False
        self.set_status("作業失敗")
        self.append_text(f"發生錯誤：{exc}\n")
        messagebox.showerror("錯誤", str(exc))

    def start_download_current(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        if not self.urls:
            messagebox.showwarning("提示", "目前沒有可下載的圖片，請先按『開始找當前頁』")
            return
        save_dir = self.save_dir_var.get().strip()
        if not save_dir:
            messagebox.showerror("錯誤", "請先指定下載資料夾")
            return
        os.makedirs(save_dir, exist_ok=True)
        self.is_running = True
        self.set_status("正在下載目前列表...")
        threading.Thread(target=self.download_current_worker, daemon=True).start()

    def start_download_range(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        save_dir = self.save_dir_var.get().strip()
        if not save_dir:
            messagebox.showerror("錯誤", "請先指定下載資料夾")
            return
        try:
            start_page = self.get_int(self.start_page_var.get(), "起始頁")
            end_page = self.get_int(self.end_page_var.get(), "結束頁")
            if start_page <= 0 or end_page <= 0:
                raise ValueError("頁碼必須大於 0")
            if start_page > end_page:
                raise ValueError("起始頁不可大於結束頁")
            self.prepare_single_find()
        except ValueError as exc:
            messagebox.showerror("錯誤", str(exc))
            return
        os.makedirs(save_dir, exist_ok=True)
        self.is_running = True
        self.clear_output()
        self.page_images.clear()
        self.downloaded_urls.clear()
        self.set_status("正在逐頁下載 jpg...")
        threading.Thread(target=self.download_range_worker, daemon=True).start()

    def build_filename(self, index: int, page_num: int, url: str) -> str:
        parsed = urlparse(url)
        original_name = os.path.basename(parsed.path)
        base, ext = os.path.splitext(original_name)
        safe_base = re.sub(r"[^a-zA-Z0-9._-]+", "_", base).strip("_") or f"image_{index:03d}"
        ext = ".jpg" if ext.lower() not in {".jpg", ".jpeg"} else ".jpg"
        return f"p{page_num:03d}_{index:03d}_{safe_base}{ext}"

    def load_cookies_into_requests(self, context, session: requests.Session) -> None:
        for cookie in context.cookies():
            name = cookie.get("name", "")
            if not name:
                continue
            session.cookies.set(
                name,
                cookie.get("value", ""),
                domain=cookie.get("domain"),
                path=cookie.get("path", "/"),
            )

    def download_urls(
        self,
        session: requests.Session,
        headers: dict,
        urls: list[str],
        page_num: int,
        save_dir: str,
    ) -> tuple[int, int]:
        success_count = 0
        fail_count = 0
        for idx, url in enumerate(urls, 1):
            if url in self.downloaded_urls:
                continue
            try:
                response = session.get(url, headers=headers, timeout=30)
                response.raise_for_status()
                filename = self.build_filename(idx, page_num, url)
                path = os.path.join(save_dir, filename)
                with open(path, "wb") as f:
                    f.write(response.content)
                self.downloaded_urls.add(url)
                success_count += 1
                self.root.after(0, lambda p=path: self.append_text(f"已下載：{p}\n"))
            except Exception as exc:
                fail_count += 1
                self.root.after(0, lambda u=url, e=exc: self.append_text(f"下載失敗：{u}\n原因：{e}\n"))
        return success_count, fail_count

    def download_current_worker(self) -> None:
        base_dir = self.save_dir_var.get().strip()
        page_url = self.build_url_for_page(self.get_int(self.start_page_var.get(), "起始頁"))
        headers = {"User-Agent": UA, "Referer": self.url_var.get().strip()}
        success_count = 0
        fail_count = 0
        session = requests.Session()
        try:
            final_dir, title_text, chapter_text = self.resolve_save_dir(base_dir, page_url)
            if self.use_title_chapter_dir_var.get():
                self.root.after(
                    0,
                    lambda d=final_dir, t=title_text, c=chapter_text:
                        self.append_text(f"下載資料夾：{d}\n標題：{t or '(未抓到)'}\n篇章：{c or '(未抓到)'}\n"),
                )
            for i, url in enumerate(self.urls, 1):
                response = session.get(url, headers=headers, timeout=30)
                response.raise_for_status()
                filename = self.build_filename(i, 1, url)
                path = os.path.join(final_dir, filename)
                with open(path, "wb") as f:
                    f.write(response.content)
                success_count += 1
                self.root.after(0, lambda p=path: self.append_text(f"已下載：{p}\n"))
        except Exception as exc:
            fail_count += 1
            self.root.after(0, lambda e=exc: self.append_text(f"下載發生錯誤：{e}\n"))
        self.root.after(0, lambda: self.on_download_done(success_count, fail_count, final_dir if 'final_dir' in locals() else base_dir))

    def download_range_worker(self) -> None:
        start_page = self.get_int(self.start_page_var.get(), "起始頁")
        end_page = self.get_int(self.end_page_var.get(), "結束頁")
        template, scroll_times, scroll_wait, timeout, min_width, min_height = self.prepare_single_find()

        page1_html = None
        page1_url = template.replace("(#)", "1")

        try:
            headers = {"User-Agent": UA, "Referer": page1_url}
            response = requests.get(page1_url, headers=headers, timeout=30)
            response.raise_for_status()
            page1_html = response.text
            max_page = self.extract_max_page_from_html(page1_html)
            if end_page != max_page:
                end_page = max_page
                self.root.after(0, lambda m=max_page: self.end_page_var.set(str(m)))
                self.root.after(0, lambda m=max_page: self.append_text(f"已自動校正結束頁為 {m}\n"))
        except Exception as exc:
            self.root.after(0, lambda e=exc: self.append_text(f"自動抓最大頁數失敗，沿用手動結束頁：{e}\n"))

        base_dir = self.save_dir_var.get().strip()
        try:
            final_dir, title_text, chapter_text = self.resolve_save_dir(base_dir, page1_url, page1_html)
            if self.use_title_chapter_dir_var.get():
                self.root.after(
                    0,
                    lambda d=final_dir, t=title_text, c=chapter_text:
                        self.append_text(f"下載資料夾：{d}\n標題：{t or '(未抓到)'}\n篇章：{c or '(未抓到)'}\n"),
                )
        except Exception as exc:
            self.root.after(0, lambda e=exc: self.append_text(f"建立標題/篇章資料夾失敗，改用原始資料夾：{e}\n"))
            final_dir = base_dir
            os.makedirs(final_dir, exist_ok=True)

        total_success = 0
        total_fail = 0
        all_urls: list[str] = []

        try:
            with sync_playwright() as p:
                browser = p.chromium.launch(headless=False)
                context = browser.new_context(viewport={"width": 1400, "height": 900}, user_agent=UA)
                page = context.new_page()
                page.set_default_timeout(timeout)

                session = requests.Session()
                headers = {"User-Agent": UA}

                for page_num in range(start_page, end_page + 1):
                    current_url = self.build_url_for_page(page_num)
                    self.root.after(0, lambda n=page_num, u=current_url: self.append_text(f"\n[頁 {n}] 開啟：{u}\n"))
                    page.goto(current_url, wait_until="domcontentloaded")
                    page.wait_for_timeout(2500)

                    for _ in range(scroll_times):
                        page.mouse.wheel(0, 2000)
                        page.wait_for_timeout(scroll_wait)

                    dom_urls = self.collect_dom_urls(page, min_width, min_height)
                    current_urls: list[str] = []
                    seen = set()
                    for url in dom_urls:
                        if self.is_probably_comic_image(url) and url not in seen:
                            seen.add(url)
                            current_urls.append(url)

                    self.page_images[page_num] = current_urls
                    for url in current_urls:
                        if url not in all_urls:
                            all_urls.append(url)

                    self.root.after(0, lambda n=page_num, c=len(current_urls): self.append_text(f"[頁 {n}] 找到 {c} 張 jpg\n"))

                    self.load_cookies_into_requests(context, session)
                    headers["Referer"] = current_url
                    success_count, fail_count = self.download_urls(session, headers, current_urls, page_num, final_dir)
                    total_success += success_count
                    total_fail += fail_count
                    self.root.after(0, lambda n=page_num, s=success_count, f=fail_count: self.append_text(f"[頁 {n}] 已下載 {s}，失敗 {f}\n"))

                browser.close()

            self.urls = all_urls
            self.root.after(0, lambda: self.on_download_done(total_success, total_fail, final_dir))
        except PlaywrightTimeoutError as exc:
            err = exc
            self.root.after(0, lambda err=err: self.on_find_error(err))
        except Exception as exc:
            err = exc
            self.root.after(0, lambda err=err: self.on_find_error(err))

    def on_download_done(self, success_count: int, fail_count: int, save_dir: str) -> None:
        self.is_running = False
        self.set_status(f"下載完成，成功 {success_count}，失敗 {fail_count}")
        messagebox.showinfo(
            "完成",
            f"下載完成\n成功：{success_count}\n失敗：{fail_count}\n資料夾：{save_dir}",
        )


def main() -> None:
    root = tk.Tk()
    app = DM5CrawlerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
