import os
import re
import threading
import asyncio
import html as html_lib
import tkinter as tk
import xml.etree.ElementTree as ET
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText
from urllib.parse import urlparse

import requests
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeoutError

# ==========================================
# 預設設定
# ==========================================
DEFAULT_URL = ""
UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36"
)

class DM5CrawlerApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("DM5 特化逐頁下載 (非同步並行版)")
        self.root.geometry("1150x920")

        self.urls: list[str] = []
        self.page_images: dict[int, list[str]] = {}
        self.downloaded_urls: set[str] = set()
        
        # 狀態控制變數
        self.is_running = False
        self.is_paused = False
        self.cancel_event = threading.Event()
        self.pause_event_async = None
        self._loop = None

        # UI 變數
        self.rss_url_var = tk.StringVar(value="")
        self.url_var = tk.StringVar(value=DEFAULT_URL)
        self.status_var = tk.StringVar(value="就緒")
        self.scroll_times_var = tk.StringVar(value="0")
        self.scroll_wait_var = tk.StringVar(value="1200")
        self.timeout_var = tk.StringVar(value="100")
        self.save_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "downloads"))
        self.min_width_var = tk.StringVar(value="300")
        self.min_height_var = tk.StringVar(value="300")
        self.start_page_var = tk.StringVar(value="1")
        self.end_page_var = tk.StringVar(value="5")
        self.use_title_chapter_dir_var = tk.BooleanVar(value=False)
        self.max_concurrent_var = tk.StringVar(value="5")
        self.image_format_var = tk.StringVar(value="JPG & PNG")
        self.manual_verify_var = tk.BooleanVar(value=False)

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        # 第 0 列: RSS 批次
        ttk.Label(top, text="RSS 網址 (批次)").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.rss_url_var, width=110).grid(row=0, column=1, columnspan=7, sticky="ew", padx=6)
        ttk.Button(top, text="批次下載 RSS 列表", command=self.start_batch_rss).grid(row=0, column=8, columnspan=2, sticky="e", padx=6)

        # 第 1 列: 一般網址
        ttk.Label(top, text="DM5 網址").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.url_var, width=110).grid(row=1, column=1, columnspan=7, sticky="ew", padx=6, pady=(8, 0))

        # 第 2 列: 頁面與尺寸與格式
        ttk.Label(top, text="起始頁").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.start_page_var, width=8).grid(row=2, column=1, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(top, text="結束頁").grid(row=2, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.end_page_var, width=8).grid(row=2, column=3, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(top, text="最小寬").grid(row=2, column=4, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.min_width_var, width=8).grid(row=2, column=5, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(top, text="最小高").grid(row=2, column=6, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.min_height_var, width=8).grid(row=2, column=7, sticky="w", padx=6, pady=(8, 0))
        
        ttk.Label(top, text="圖片格式").grid(row=2, column=8, sticky="w", pady=(8, 0), padx=6)
        format_cb = ttk.Combobox(top, textvariable=self.image_format_var, values=["僅 JPG", "僅 PNG", "JPG & PNG"], width=10, state="readonly")
        format_cb.grid(row=2, column=9, sticky="w", padx=6, pady=(8, 0))

        # 第 3 列: 捲動與時間與並發
        ttk.Label(top, text="捲動次數").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.scroll_times_var, width=8).grid(row=3, column=1, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(top, text="每次等待(ms)").grid(row=3, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.scroll_wait_var, width=10).grid(row=3, column=3, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(top, text="頁面 timeout(ms)").grid(row=3, column=4, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.timeout_var, width=12).grid(row=3, column=5, sticky="w", padx=6, pady=(8, 0))
        
        ttk.Label(top, text="同時最大頁數").grid(row=3, column=6, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.max_concurrent_var, width=8).grid(row=3, column=7, sticky="w", padx=6, pady=(8, 0))

        # 第 4 列: 資料夾與驗證
        ttk.Label(top, text="下載資料夾").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.save_dir_var, width=90).grid(row=4, column=1, columnspan=5, sticky="ew", padx=6, pady=(8, 0))
        ttk.Button(top, text="選擇", command=self.choose_dir).grid(row=4, column=6, sticky="w", pady=(8, 0))
        ttk.Checkbutton(top, text="自動建立資料夾（標題/篇章）", variable=self.use_title_chapter_dir_var).grid(row=4, column=7, sticky="w", padx=6, pady=(8, 0))
        ttk.Checkbutton(top, text="啟用手動通關 (首頁暫停)", variable=self.manual_verify_var).grid(row=4, column=8, columnspan=2, sticky="w", padx=6, pady=(8, 0))

        # 第 5 列: 功能按鈕
        ttk.Button(top, text="自動抓最大頁數", command=self.fetch_max_page).grid(row=5, column=4, sticky="e", padx=6, pady=(8, 0))
        
        self.pause_btn = ttk.Button(top, text="暫停", command=self.toggle_pause, state="disabled")
        self.pause_btn.grid(row=5, column=5, sticky="e", padx=6, pady=(8, 0))
        
        self.cancel_btn = ttk.Button(top, text="終止", command=self.cancel_task, state="disabled")
        self.cancel_btn.grid(row=5, column=6, sticky="e", padx=6, pady=(8, 0))
        
        ttk.Button(top, text="並行逐頁下載範圍", command=self.start_download_range).grid(row=5, column=7, sticky="e", padx=6, pady=(8, 0))
        ttk.Button(top, text="下載目前列表", command=self.start_download_current).grid(row=5, column=8, columnspan=2, sticky="e", padx=6, pady=(8, 0))

        top.columnconfigure(1, weight=1)

        middle = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        middle.pack(fill="both", expand=True)
        ttk.Label(middle, textvariable=self.status_var).pack(anchor="w", pady=(0, 8))
        self.result_text = ScrolledText(middle, wrap="word")
        self.result_text.pack(fill="both", expand=True)

    # ==========================================
    # 流程控制 (暫停 / 繼續 / 終止)
    # ==========================================
    def toggle_pause(self) -> None:
        if not self.is_running or not self.pause_event_async:
            return
            
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_btn.config(text="繼續")
            self.set_status("已暫停 (等待手動繼續...)")
            self.append_text(">>> 程式已暫停。若需手動驗證，請於瀏覽器操作後點擊「繼續」。\n")
            if hasattr(self, '_loop') and self._loop and self._loop.is_running():
                self._loop.call_soon_threadsafe(self.pause_event_async.clear)
        else:
            self.pause_btn.config(text="暫停")
            self.set_status("正在執行中...")
            self.append_text(">>> 程式繼續執行。\n")
            if hasattr(self, '_loop') and self._loop and self._loop.is_running():
                self._loop.call_soon_threadsafe(self.pause_event_async.set)

    def cancel_task(self) -> None:
        if self.is_running:
            self.cancel_event.set()
            self.set_status("正在終止中...")
            self.append_text(">>> 收到終止要求，正在安全結束任務...\n")
            if self.is_paused and hasattr(self, '_loop') and self._loop and self._loop.is_running():
                self._loop.call_soon_threadsafe(self.pause_event_async.set)

    def reset_control_state(self) -> None:
        self.cancel_event.clear()
        self.is_paused = False
        self.pause_btn.config(text="暫停", state="normal")
        self.cancel_btn.config(state="normal")

    async def wait_if_paused(self):
        if self.cancel_event.is_set():
            raise asyncio.CancelledError("User cancelled the task.")
        if self.pause_event_async:
            await self.pause_event_async.wait()
        if self.cancel_event.is_set():
            raise asyncio.CancelledError("User cancelled the task.")

    # ==========================================
    # 輔助函式
    # ==========================================
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
        if "(#)" in url: return url
        m = re.match(r"^(https?://www\.dm5\.cn/m\d+?)/?$", url)
        if m: return f"{m.group(1)}-p(#)"
        converted, count = re.subn(r"-p\d+(?=(?:[#/?]|$))", "-p(#)", url)
        if count > 0: return converted
        converted, count = re.subn(r"#ipg\d+", "#ipg(#)", url)
        if count > 0: return converted
        raise ValueError("只支援 DM5 的作品首頁、-p頁碼網址，或已含 (#) 的範本")

    def build_url_for_page(self, page_num: int) -> str:
        template = self.normalize_dm5_template(self.url_var.get())
        return template.replace("(#)", str(page_num))

    def preview_template(self) -> None:
        try:
            template = self.normalize_dm5_template(self.url_var.get())
        except ValueError as exc:
            messagebox.showerror("錯誤", str(exc))
            return
        self.append_text(f"自動轉換範本：{template}\n")
        messagebox.showinfo("範本", template)

    def is_probably_comic_image(self, url: str) -> bool:
        lowered = url.lower()
        fmt = self.image_format_var.get()
        if fmt == "僅 JPG":
            pattern = r"\.(jpg|jpeg)(\?|$)"
        elif fmt == "僅 PNG":
            pattern = r"\.png(\?|$)"
        else:
            pattern = r"\.(jpg|jpeg|png)(\?|$)"
            
        if not re.search(pattern, lowered):
            return False
            
        blacklist = ["logo", "avatar", "cover", "banner", "ads", "ad.", "icon", "emoji", "sprite"]
        return not any(word in lowered for word in blacklist)

    async def collect_dom_urls_async(self, page, min_width: int, min_height: int) -> list[str]:
        return await page.evaluate(
            """({minWidth, minHeight}) => {
                return [...document.images].map(img => {
                    const url = img.currentSrc || img.src || "";
                    return { url, w: img.naturalWidth || 0, h: img.naturalHeight || 0, complete: !!img.complete };
                }).filter(x => x.complete && x.url && x.w >= minWidth && x.h >= minHeight).map(x => x.url);
            }""", {"minWidth": min_width, "minHeight": min_height})

    def clean_html_text(self, text: str) -> str:
        text = re.sub(r"<[^>]+>", "", text)
        text = html_lib.unescape(text)
        return re.sub(r"\s+", " ", text).strip()

    def sanitize_path_part(self, text: str) -> str:
        text = self.clean_html_text(text)
        text = re.sub(r'[\\/:*?"<>|]+', "_", text).strip(" .")
        return text[:120] if text else ""

    def extract_title_and_chapter_from_html(self, html: str) -> tuple[str, str]:
        title_text, chapter_text = "", ""
        block_match = re.search(r'<div class="title">(.*?)</div>\s*<div class="right-bar">', html, re.IGNORECASE | re.DOTALL)
        block = block_match.group(1) if block_match else html

        pair_match = re.search(r'<span class="right-arrow">\s*<a [^>]*>(.*?)</a>\s*</span>\s*<span class="active right-arrow">\s*(.*?)\s*</span>', block, re.IGNORECASE | re.DOTALL)
        if pair_match:
            title_text, chapter_text = self.clean_html_text(pair_match.group(1)), self.clean_html_text(pair_match.group(2))

        if not title_text:
            m = re.search(r'<span class="right-arrow">\s*<a [^>]*title="([^"]+)"[^>]*>', block, re.IGNORECASE | re.DOTALL)
            if m: title_text = self.clean_html_text(m.group(1))

        if not chapter_text:
            m = re.search(r'<span class="active right-arrow">\s*(.*?)\s*</span>', block, re.IGNORECASE | re.DOTALL)
            if m: chapter_text = self.clean_html_text(m.group(1))

        if not title_text or not chapter_text:
            m = re.search(r'DM5_CTITLE\s*=\s*"([^"]+)"', html, re.IGNORECASE)
            if m:
                ctitle = self.clean_html_text(m.group(1))
                m2 = re.search(r"(第\s*\d+\s*话.*)$", ctitle)
                if m2:
                    if not chapter_text: chapter_text = self.clean_html_text(m2.group(1))
                    if not title_text: title_text = self.clean_html_text(ctitle[: m2.start()].strip())
                elif not title_text:
                    title_text = ctitle

        return title_text, chapter_text

    def resolve_save_dir(self, base_dir: str, page_url: str, html_text: str | None = None) -> tuple[str, str, str]:
        final_dir, title_text, chapter_text = base_dir, "", ""
        if self.use_title_chapter_dir_var.get():
            if html_text is None:
                response = requests.get(page_url, headers={"User-Agent": UA, "Referer": page_url}, timeout=30)
                response.raise_for_status()
                html_text = response.text
            title_text, chapter_text = self.extract_title_and_chapter_from_html(html_text)
            safe_title, safe_chapter = self.sanitize_path_part(title_text), self.sanitize_path_part(chapter_text)
            if safe_title and safe_chapter: final_dir = os.path.join(base_dir, safe_title, safe_chapter)
            elif safe_title: final_dir = os.path.join(base_dir, safe_title)

        os.makedirs(final_dir, exist_ok=True)
        return final_dir, title_text, chapter_text

    def build_filename(self, index: int, page_num: int, url: str) -> str:
        parsed = urlparse(url)
        base, ext = os.path.splitext(os.path.basename(parsed.path))
        safe_base = re.sub(r"[^a-zA-Z0-9._-]+", "_", base).strip("_") or f"image_{index:03d}"
        ext = ext.lower()
        if ext == ".jpeg": ext = ".jpg"
        elif ext not in {".jpg", ".png"}: ext = ".jpg"
        return f"p{page_num:03d}_{index:03d}_{safe_base}{ext}"

    def prepare_single_find(self) -> tuple[str, int, int, int, int, int]:
        template = self.normalize_dm5_template(self.url_var.get())
        scroll_times = self.get_int(self.scroll_times_var.get(), "捲動次數")
        scroll_wait = self.get_int(self.scroll_wait_var.get(), "每次等待")
        timeout = self.get_int(self.timeout_var.get(), "頁面 timeout")
        min_width = self.get_int(self.min_width_var.get(), "最小寬")
        min_height = self.get_int(self.min_height_var.get(), "最小高")
        return template, scroll_times, scroll_wait, timeout, min_width, min_height

    # ==========================================
    # 最大頁數與 RSS 解析
    # ==========================================
    def extract_max_page_from_html(self, html: str) -> int:
        m = re.search(r"DM5_IMAGE_COUNT\s*=\s*(\d+)", html, re.IGNORECASE)
        if m: return int(m.group(1))
        pager_matches = re.findall(r"/m\d+-p(\d+)/", html, re.IGNORECASE)
        if pager_matches: return max(int(x) for x in pager_matches)
        ipg_matches = re.findall(r"#ipg(\d+)", html, re.IGNORECASE)
        if ipg_matches: return max(int(x) for x in ipg_matches)
        raise ValueError("找不到 DM5_IMAGE_COUNT 或分頁資訊")

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

    def fetch_max_page_worker(self, template: str) -> None:
        try:
            page1_url = template.replace("(#)", "1")
            response = requests.get(page1_url, headers={"User-Agent": UA, "Referer": page1_url}, timeout=30)
            response.raise_for_status()
            max_page = self.extract_max_page_from_html(response.text)
            self.root.after(0, lambda m=max_page, u=page1_url: self.on_fetch_max_page_success(m, u))
        except Exception as exc:
            self.root.after(0, lambda err=exc: self.on_fetch_max_page_error(err))

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

    def _sync_fetch_max_page(self, raw_url: str) -> int:
        try:
            template = self.normalize_dm5_template(raw_url)
            page1_url = template.replace("(#)", "1")
            response = requests.get(page1_url, headers={"User-Agent": UA, "Referer": page1_url}, timeout=30)
            response.raise_for_status()
            return self.extract_max_page_from_html(response.text)
        except Exception as e:
            self.root.after(0, lambda err=e: self.append_text(f"嘗試抓取最大頁數時發生錯誤: {err}\n"))
            return 0

    def start_batch_rss(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        rss_url = self.rss_url_var.get().strip()
        save_dir = self.save_dir_var.get().strip()
        if not rss_url or not save_dir:
            messagebox.showerror("錯誤", "請確認已輸入 RSS 網址並指定下載資料夾")
            return
        self.is_running = True
        self.clear_output()
        self.set_status("正在解析 RSS 列表...")
        threading.Thread(target=self.batch_rss_worker, args=(rss_url,), daemon=True).start()

    def batch_rss_worker(self, rss_url: str) -> None:
        try:
            response = requests.get(rss_url, headers={"User-Agent": UA}, timeout=30)
            response.raise_for_status()
            root = ET.fromstring(response.text)
            
            target_urls = []
            for item in root.findall('.//item'):
                link_elem = item.find('link')
                title_elem = item.find('title')
                if link_elem is not None and link_elem.text:
                    target_urls.append((title_elem.text.strip() if title_elem is not None else "未知", link_elem.text.strip()))
                    
            if not target_urls: raise ValueError("在 RSS 中找不到任何項目")
            target_urls.reverse()
            
            self.root.after(0, lambda c=len(target_urls): self.append_text(f"成功解析 RSS，共 {c} 話準備下載。\n"))
            self.root.after(0, lambda: self.append_text("-" * 40 + "\n"))
            
            for idx, (chapter_title, chapter_url) in enumerate(target_urls, 1):
                if self.cancel_event.is_set(): break
                self.root.after(0, lambda t=chapter_title: self.append_text(f"\n>>> 開始處理：{t}\n"))
                self.root.after(0, lambda u=chapter_url: self.url_var.set(u))
                
                max_page = self._sync_fetch_max_page(chapter_url)
                if max_page == 0:
                    self.root.after(0, lambda t=chapter_title: self.append_text(f"跳過 {t}：無法取得最大頁數\n"))
                    continue
                    
                self.root.after(0, lambda: self.start_page_var.set("1"))
                self.root.after(0, lambda m=max_page: self.end_page_var.set(str(m)))
                import time; time.sleep(1)
                
                self._run_download_range_async(is_batch=True)
                self.root.after(0, lambda t=chapter_title: self.append_text(f"<<< {t} 處理完成\n"))
                
            if not self.cancel_event.is_set():
                self.root.after(0, lambda: self.set_status("批次 RSS 下載完成！"))
                self.root.after(0, lambda: messagebox.showinfo("完成", "批次 RSS 下載完成！"))
            self.is_running = False

        except Exception as exc:
            self.is_running = False
            self.root.after(0, lambda e=exc: self.append_text(f"批次解析失敗：{e}\n"))
            self.root.after(0, lambda: self.set_status("批次作業失敗"))

    # ==========================================
    # 下載邏輯 (單一 / 目前列表)
    # ==========================================
    async def get_playwright_cookies(self, context) -> dict:
        return {c['name']: c['value'] for c in await context.cookies()}

    def download_urls_requests(self, session: requests.Session, headers: dict, urls: list[str], page_num: int, save_dir: str) -> tuple[int, int]:
        success_count, fail_count = 0, 0
        for idx, url in enumerate(urls, 1):
            if url in self.downloaded_urls: continue
            if self.cancel_event.is_set(): break
            try:
                response = session.get(url, headers=headers, timeout=30)
                response.raise_for_status()
                path = os.path.join(save_dir, self.build_filename(idx, page_num, url))
                with open(path, "wb") as f: f.write(response.content)
                self.downloaded_urls.add(url)
                success_count += 1
                self.root.after(0, lambda p=path: self.append_text(f"已下載：{p}\n"))
            except Exception as exc:
                fail_count += 1
                self.root.after(0, lambda u=url, e=exc: self.append_text(f"下載失敗：{u}\n原因：{e}\n"))
        return success_count, fail_count

    def start_download_current(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        if not self.urls:
            messagebox.showwarning("提示", "目前沒有可下載的圖片")
            return
        save_dir = self.save_dir_var.get().strip()
        if not save_dir:
            messagebox.showerror("錯誤", "請先指定下載資料夾")
            return
        os.makedirs(save_dir, exist_ok=True)
        self.is_running = True
        self.reset_control_state()
        self.set_status("正在下載目前列表...")
        threading.Thread(target=self.download_current_worker, daemon=True).start()

    def download_current_worker(self) -> None:
        base_dir = self.save_dir_var.get().strip()
        page_url = self.build_url_for_page(self.get_int(self.start_page_var.get(), "起始頁"))
        session = requests.Session()
        success_count, fail_count = 0, 0
        try:
            final_dir, t_text, c_text = self.resolve_save_dir(base_dir, page_url)
            if self.use_title_chapter_dir_var.get():
                self.root.after(0, lambda d=final_dir, t=t_text, c=c_text: self.append_text(f"下載資料夾：{d}\n標題：{t}\n篇章：{c}\n"))
            for i, url in enumerate(self.urls, 1):
                if self.cancel_event.is_set(): break
                response = session.get(url, headers={"User-Agent": UA, "Referer": self.url_var.get().strip()}, timeout=30)
                response.raise_for_status()
                path = os.path.join(final_dir, self.build_filename(i, 1, url))
                with open(path, "wb") as f: f.write(response.content)
                success_count += 1
                self.root.after(0, lambda p=path: self.append_text(f"已下載：{p}\n"))
        except Exception as exc:
            fail_count += 1
            self.root.after(0, lambda e=exc: self.append_text(f"下載發生錯誤：{e}\n"))
        self.root.after(0, lambda: self.on_download_done(success_count, fail_count, final_dir if 'final_dir' in locals() else base_dir))

    # ==========================================
    # 非同步並行範圍下載核心
    # ==========================================
    def start_download_range(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        if not self.save_dir_var.get().strip():
            messagebox.showerror("錯誤", "請先指定下載資料夾")
            return
        try:
            if self.get_int(self.start_page_var.get(), "起始頁") > self.get_int(self.end_page_var.get(), "結束頁"):
                raise ValueError("起始頁不可大於結束頁")
            if self.get_int(self.max_concurrent_var.get(), "同時最大頁數") <= 0:
                raise ValueError("同時最大頁數必須大於 0")
            self.prepare_single_find()
        except ValueError as exc:
            messagebox.showerror("錯誤", str(exc))
            return
            
        os.makedirs(self.save_dir_var.get().strip(), exist_ok=True)
        self.is_running = True
        self.clear_output()
        self.page_images.clear()
        self.downloaded_urls.clear()
        
        threading.Thread(target=self._run_download_range_async, daemon=True).start()

    def _run_download_range_async(self, is_batch=False):
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        self._loop = loop
        try:
            loop.run_until_complete(self.download_range_worker_async(is_batch))
        finally:
            loop.close()
            self._loop = None

    async def download_single_page_async(self, context, page_num: int, semaphore: asyncio.Semaphore, final_dir: str):
        template, scroll_times, scroll_wait, timeout, min_width, min_height = self.prepare_single_find()
        current_url = self.build_url_for_page(page_num)
        
        async with semaphore:
            await self.wait_if_paused()
            page = await context.new_page()
            page.set_default_timeout(timeout)
            self.root.after(0, lambda n=page_num, u=current_url: self.append_text(f"\n[頁 {n}] 開始處理：{u}\n"))
            
            try:
                await page.goto(current_url, wait_until="domcontentloaded")
            except PlaywrightTimeoutError:
                pass
            except Exception as e:
                if "Target closed" not in str(e):
                    self.root.after(0, lambda n=page_num, err=e: self.append_text(f"[頁 {n}] 開啟失敗：{err}\n"))
                await page.close()
                return 0, 0
                
            # 手動人類辨識
            start_page = self.get_int(self.start_page_var.get(), "起始頁")
            if page_num == start_page and self.manual_verify_var.get():
                self.root.after(0, lambda: self.append_text(f"[頁 {page_num}] 觸發手動通關，已自動暫停。\n"))
                self.root.after(0, self.toggle_pause) 
            
            await self.wait_if_paused()
            await page.wait_for_timeout(2500)

            for _ in range(scroll_times):
                await self.wait_if_paused()
                await page.mouse.wheel(0, 2000)
                await page.wait_for_timeout(scroll_wait)

            # 空白重試機制
            max_retries = 3
            current_urls: list[str] = []
            
            for retry_count in range(max_retries + 1):
                await self.wait_if_paused()
                dom_urls = await self.collect_dom_urls_async(page, min_width, min_height)
                seen = set()
                current_urls.clear()
                for url in dom_urls:
                    if self.is_probably_comic_image(url) and url not in seen:
                        seen.add(url)
                        current_urls.append(url)
                
                if current_urls:
                    break 
                else:
                    if retry_count < max_retries:
                        self.root.after(0, lambda n=page_num, r=retry_count+1: self.append_text(f"[頁 {n}] 未抓到圖片，500ms 後第 {r} 次重試...\n"))
                        await page.wait_for_timeout(500)
                    else:
                        self.root.after(0, lambda n=page_num: self.append_text(f"[頁 {n}] 重試 {max_retries} 次後仍未抓到圖片。\n"))

            self.page_images[page_num] = current_urls
            self.root.after(0, lambda n=page_num, c=len(current_urls): self.append_text(f"[頁 {n}] 找到 {c} 張圖片\n"))

            if not current_urls:
                await page.close()
                return 0, 0

            await self.wait_if_paused()
            playwright_cookies = await self.get_playwright_cookies(context)
            session = requests.Session()
            session.cookies.update(playwright_cookies)
            
            if self.cancel_event.is_set():
                await page.close()
                raise asyncio.CancelledError()

            success, fail = self.download_urls_requests(session, {"User-Agent": UA, "Referer": current_url}, current_urls, page_num, final_dir)
            self.root.after(0, lambda n=page_num, s=success, f=fail: self.append_text(f"[頁 {n}] 已下載 {s}，失敗 {f}\n"))
            
            await page.close()
            return success, fail

    async def download_range_worker_async(self, is_batch: bool = False) -> None:
        self.root.after(0, self.reset_control_state)
        self.pause_event_async = asyncio.Event()
        self.pause_event_async.set()

        start_page, end_page = self.get_int(self.start_page_var.get(), "起始頁"), self.get_int(self.end_page_var.get(), "結束頁")
        max_concurrent = self.get_int(self.max_concurrent_var.get(), "同時最大頁數")
        template, _, _, _, _, _ = self.prepare_single_find()
        
        self.root.after(0, lambda: self.set_status(f"正在並行下載 ({max_concurrent} 頁同時進行)..."))

        page1_url = template.replace("(#)", "1")
        page1_html = None
        try:
            response = requests.get(page1_url, headers={"User-Agent": UA, "Referer": page1_url}, timeout=30)
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
                self.root.after(0, lambda d=final_dir, t=title_text, c=chapter_text: self.append_text(f"下載資料夾：{d}\n標題：{t}\n篇章：{c}\n"))
        except Exception as exc:
            self.root.after(0, lambda e=exc: self.append_text(f"建立標題/篇章資料夾失敗，改用原始資料夾：{e}\n"))
            final_dir, os.makedirs(base_dir, exist_ok=True) = base_dir, None

        total_success, total_fail = 0, 0
        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=False)
                context = await browser.new_context(viewport={"width": 1400, "height": 900}, user_agent=UA)
                semaphore = asyncio.Semaphore(max_concurrent)
                
                tasks = [self.download_single_page_async(context, pn, semaphore, final_dir) for pn in range(start_page, end_page + 1)]
                results = await asyncio.gather(*tasks, return_exceptions=True)
                
                for res in results:
                    if isinstance(res, tuple):
                        total_success += res[0]; total_fail += res[1]

                await browser.close()

            all_urls = []
            for urls_in_page in self.page_images.values():
                for url in urls_in_page:
                    if url not in all_urls: all_urls.append(url)
            self.urls = all_urls

            self.root.after(0, lambda: self.on_download_done(total_success, total_fail, final_dir, is_batch))
            
        except Exception as exc:
            self.root.after(0, lambda err=exc: self.append_text(f"發生錯誤：{err}\n"))
            if not is_batch: self.is_running = False

    def on_download_done(self, success_count: int, fail_count: int, save_dir: str, is_batch: bool = False) -> None:
        self.pause_btn.config(text="暫停", state="disabled")
        self.cancel_btn.config(state="disabled")
        
        if self.cancel_event.is_set():
             if not is_batch: self.is_running = False
             self.set_status("任務已終止")
             self.append_text("\n=== 使用者手動終止了任務 ===\n")
             return

        if not is_batch:
            self.is_running = False
            self.set_status(f"下載完成，成功 {success_count}，失敗 {fail_count}")
            messagebox.showinfo("完成", f"下載完成\n成功：{success_count}\n失敗：{fail_count}\n資料夾：{save_dir}")
        else:
            self.append_text(f"單話下載完成，成功 {success_count}，失敗 {fail_count}\n")

def main() -> None:
    root = tk.Tk()
    app = DM5CrawlerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
