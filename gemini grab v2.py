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
        self.root.title("DM5 特化逐頁下載 (共用 Session 版)")
        self.root.geometry("1280x920") 

        self.urls: list[str] = []
        self.page_images: dict[int, list[str]] = {}
        self.downloaded_urls: set[str] = set()
        
        # 佇列排程 (存放 Tuple: (標題, 網址))
        self.target_queue: list[tuple[str, str]] = []
        
        # 狀態控制變數
        self.is_running = False
        self.is_paused = False
        self.cancel_event = threading.Event()
        self.pause_event_async = None
        self._loop = None
        self.chapter_end_flag = False # 動態終止旗標

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

        # 第 0 列: RSS 批次與手動
        ttk.Label(top, text="RSS 網址 (批次)").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.rss_url_var, width=80).grid(row=0, column=1, columnspan=5, sticky="ew", padx=6)
        ttk.Button(top, text="解析 RSS 列表", command=self.parse_rss).grid(row=0, column=6, columnspan=2, sticky="e", padx=6)
        ttk.Button(top, text="手動貼上 RSS", command=self.open_manual_rss_dialog).grid(row=0, column=8, columnspan=2, sticky="e", padx=6)

        # 第 1 列: 一般網址 (單話)
        ttk.Label(top, text="DM5 網址 (單話)").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.url_var, width=80).grid(row=1, column=1, columnspan=5, sticky="ew", padx=6, pady=(8, 0))
        ttk.Button(top, text="解析單話網址", command=self.parse_single).grid(row=1, column=6, columnspan=2, sticky="e", padx=6, pady=(8, 0))

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

        # 第 5 列: 下載控制區
        action_frame = ttk.Frame(top)
        action_frame.grid(row=5, column=0, columnspan=10, sticky="e", pady=(12, 0))
        
        ttk.Button(action_frame, text="▶ 開始佇列下載", command=self.start_queue_download).pack(side="left", padx=4)
        self.pause_btn = ttk.Button(action_frame, text="⏸ 暫停", command=self.toggle_pause, state="disabled")
        self.pause_btn.pack(side="left", padx=4)
        self.cancel_btn = ttk.Button(action_frame, text="⏹ 終止", command=self.cancel_task, state="disabled")
        self.cancel_btn.pack(side="left", padx=4)

        top.columnconfigure(1, weight=1)

        # ==========================================
        # 分割畫面 (左右兩欄)
        # ==========================================
        middle = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        middle.pack(fill="both", expand=True)
        ttk.Label(middle, textvariable=self.status_var).pack(anchor="w", pady=(0, 4))
        
        paned = ttk.PanedWindow(middle, orient=tk.HORIZONTAL)
        paned.pack(fill="both", expand=True)

        left_frame = ttk.Frame(paned)
        right_frame = ttk.Frame(paned)
        
        paned.add(left_frame, weight=1)
        paned.add(right_frame, weight=2)

        # 左邊：佇列
        ttk.Label(left_frame, text="待下載佇列 (Queue)").pack(anchor="w", pady=(0, 4))
        self.queue_text = ScrolledText(left_frame, wrap="word", width=40, bg="#f4f4f4")
        self.queue_text.pack(fill="both", expand=True)
        self.queue_text.insert("end", "佇列為空。\n請先解析 RSS 或是單話網址。")
        self.queue_text.config(state="disabled")

        # 右邊：日誌
        ttk.Label(right_frame, text="執行進度與日誌 (Logs)").pack(anchor="w", pady=(0, 4))
        self.log_text = ScrolledText(right_frame, wrap="word")
        self.log_text.pack(fill="both", expand=True)

    # ==========================================
    # 流程控制 (暫停 / 繼續 / 終止)
    # ==========================================
    def toggle_pause(self) -> None:
        if not self.is_running or not self.pause_event_async:
            return
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_btn.config(text="▶ 繼續")
            self.set_status("已暫停 (等待手動繼續...)")
            self.append_log(">>> 程式已暫停。若需手動驗證，請於瀏覽器操作後點擊「繼續」。\n")
            if hasattr(self, '_loop') and self._loop and self._loop.is_running():
                self._loop.call_soon_threadsafe(self.pause_event_async.clear)
        else:
            self.pause_btn.config(text="⏸ 暫停")
            self.set_status("正在執行中...")
            self.append_log(">>> 程式繼續執行。\n")
            if hasattr(self, '_loop') and self._loop and self._loop.is_running():
                self._loop.call_soon_threadsafe(self.pause_event_async.set)

    def cancel_task(self) -> None:
        if self.is_running:
            self.cancel_event.set()
            self.set_status("正在終止中...")
            self.append_log(">>> 收到終止要求，正在安全結束任務...\n")
            if self.is_paused and hasattr(self, '_loop') and self._loop and self._loop.is_running():
                self._loop.call_soon_threadsafe(self.pause_event_async.set)

    def reset_control_state(self) -> None:
        self.cancel_event.clear()
        self.is_paused = False
        self.chapter_end_flag = False
        self.pause_btn.config(text="⏸ 暫停", state="normal")
        self.cancel_btn.config(state="normal")

    async def wait_if_paused(self):
        if self.cancel_event.is_set(): raise asyncio.CancelledError("User cancelled the task.")
        if self.pause_event_async: await self.pause_event_async.wait()
        if self.cancel_event.is_set(): raise asyncio.CancelledError("User cancelled the task.")

    # ==========================================
    # 輔助與 UI 更新函式
    # ==========================================
    def choose_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.save_dir_var.get() or os.getcwd())
        if selected: self.save_dir_var.set(selected)

    def append_log(self, text: str) -> None:
        self.log_text.insert("end", text)
        self.log_text.see("end")

    def clear_log(self) -> None:
        self.log_text.delete("1.0", "end")

    def set_status(self, text: str) -> None:
        self.status_var.set(text)

    def update_queue_display(self) -> None:
        self.queue_text.config(state="normal")
        self.queue_text.delete("1.0", "end")
        if not self.target_queue:
            self.queue_text.insert("end", "佇列為空。\n請先解析 RSS 或是單話網址。")
        else:
            for i, (title, url) in enumerate(self.target_queue, 1):
                self.queue_text.insert("end", f"[{i}] {title}\n    {url}\n\n")
        self.queue_text.config(state="disabled")

    def get_int(self, value: str, field_name: str) -> int:
        try: return int(value.strip())
        except ValueError as exc: raise ValueError(f"{field_name} 必須是整數") from exc

    def normalize_dm5_template(self, raw_url: str) -> str:
        url = raw_url.strip()
        if not url: raise ValueError("請輸入 DM5 網址")
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

    def build_filename(self, index: int, page_num: int, url: str) -> str:
        parsed = urlparse(url)
        base, ext = os.path.splitext(os.path.basename(parsed.path))
        safe_base = re.sub(r"[^a-zA-Z0-9._-]+", "_", base).strip("_") or f"image_{index:03d}"
        ext = ext.lower()
        if ext == ".jpeg": ext = ".jpg"
        elif ext not in {".jpg", ".png"}: ext = ".jpg"
        return f"p{page_num:03d}_{index:03d}_{safe_base}{ext}"

    def is_probably_comic_image(self, url: str) -> bool:
        lowered = url.lower()
        if "404.png" in lowered or "error" in lowered:
            return True
            
        fmt = self.image_format_var.get()
        if fmt == "僅 JPG": pattern = r"\.(jpg|jpeg)(\?|$)"
        elif fmt == "僅 PNG": pattern = r"\.png(\?|$)"
        else: pattern = r"\.(jpg|jpeg|png)(\?|$)"
        if not re.search(pattern, lowered): return False
        blacklist = ["logo", "avatar", "cover", "banner", "ads", "ad.", "icon", "emoji", "sprite"]
        return not any(word in lowered for word in blacklist)

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

    def extract_max_page_from_html(self, html: str) -> int:
        m = re.search(r"DM5_IMAGE_COUNT\s*=\s*(\d+)", html, re.IGNORECASE)
        if m: return int(m.group(1))
        pager_matches = re.findall(r"/m\d+-p(\d+)/", html, re.IGNORECASE)
        if pager_matches: return max(int(x) for x in pager_matches)
        ipg_matches = re.findall(r"#ipg(\d+)", html, re.IGNORECASE)
        if ipg_matches: return max(int(x) for x in ipg_matches)
        raise ValueError("找不到 DM5_IMAGE_COUNT 或分頁資訊")

    def _sync_fetch_max_page(self, raw_url: str) -> int:
        try:
            template = self.normalize_dm5_template(raw_url)
            page1_url = template.replace("(#)", "1")
            cookies = {"isAdult": "1", "fastshow": "true"}
            response = requests.get(page1_url, headers={"User-Agent": UA, "Referer": page1_url}, cookies=cookies, timeout=30)
            response.raise_for_status()
            return self.extract_max_page_from_html(response.text)
        except Exception as e:
            self.root.after(0, lambda err=e: self.append_log(f"嘗試事先抓取最大頁數失敗 (交給動態偵測): {err}\n"))
            return 0

    def prepare_single_find(self) -> tuple[str, int, int, int, int, int]:
        template = self.normalize_dm5_template(self.url_var.get())
        scroll_times = self.get_int(self.scroll_times_var.get(), "捲動次數")
        scroll_wait = self.get_int(self.scroll_wait_var.get(), "每次等待")
        timeout = self.get_int(self.timeout_var.get(), "頁面 timeout")
        min_width = self.get_int(self.min_width_var.get(), "最小寬")
        min_height = self.get_int(self.min_height_var.get(), "最小高")
        return template, scroll_times, scroll_wait, timeout, min_width, min_height

    # ==========================================
    # 階段一：解析至佇列 (Parse Single / RSS)
    # ==========================================
    def parse_single(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        url = self.url_var.get().strip()
        if not url:
            messagebox.showerror("錯誤", "請先輸入 DM5 網址")
            return
        self.set_status("正在解析單話網址...")
        threading.Thread(target=self._parse_single_worker, args=(url,), daemon=True).start()

    def _parse_single_worker(self, url: str) -> None:
        try:
            template = self.normalize_dm5_template(url)
            page1_url = template.replace("(#)", "1")
            
            headers = {"User-Agent": UA, "Referer": page1_url}
            cookies = {"isAdult": "1", "fastshow": "true"}
            response = requests.get(page1_url, headers=headers, cookies=cookies, timeout=30)
            response.raise_for_status()
            
            title_text, chapter_text = self.extract_title_and_chapter_from_html(response.text)
            full_title = f"{title_text} {chapter_text}".strip() or "未知標題"
            
            self.target_queue = [(full_title, page1_url)]
            self.root.after(0, self.update_queue_display)
            self.root.after(0, lambda: self.set_status("解析完成！請點擊「開始佇列下載」"))
            self.root.after(0, lambda: self.append_log(f"解析單話成功：{full_title}\n"))
        except Exception as exc:
            self.root.after(0, lambda e=exc: self.append_log(f"解析單話失敗：{e}\n"))
            self.root.after(0, lambda: self.set_status("解析失敗"))

    def parse_rss(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        rss_url = self.rss_url_var.get().strip()
        if not rss_url:
            messagebox.showerror("錯誤", "請先輸入 RSS 網址")
            return
        self.set_status("正在從網路抓取並解析 RSS...")
        threading.Thread(target=self._parse_rss_worker, args=(rss_url,), daemon=True).start()

    def _parse_rss_worker(self, rss_url: str) -> None:
        try:
            headers = {"User-Agent": UA}
            cookies = {"isAdult": "1", "fastshow": "true"}
            response = requests.get(rss_url, headers=headers, cookies=cookies, timeout=30)
            response.raise_for_status()
            
            raw_xml = response.text
            if not raw_xml.strip():
                raise ValueError("伺服器回傳了完全空白的內容！")
            
            self._process_rss_xml(raw_xml)
        except Exception as exc:
            self.root.after(0, lambda e=exc: self.append_log(f"網路抓取 RSS 失敗：{e}\n"))
            self.root.after(0, lambda: self.set_status("解析失敗"))

    def open_manual_rss_dialog(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return

        dialog = tk.Toplevel(self.root)
        dialog.title("手動貼上 RSS XML 內容")
        dialog.geometry("800x600")

        ttk.Label(dialog, text="請在下方貼上從瀏覽器複製的 RSS XML 原始碼：").pack(anchor="w", padx=10, pady=5)
        text_area = ScrolledText(dialog, wrap="word")
        text_area.pack(fill="both", expand=True, padx=10, pady=5)

        def on_confirm():
            raw_xml = text_area.get("1.0", "end").strip()
            if not raw_xml:
                messagebox.showwarning("警告", "內容不能為空")
                return
            dialog.destroy()
            self.set_status("正在解析手動輸入的 RSS...")
            threading.Thread(target=self._process_rss_xml, args=(raw_xml,), daemon=True).start()

        ttk.Button(dialog, text="確定解析", command=on_confirm).pack(pady=10)

    def _process_rss_xml(self, raw_xml: str) -> None:
        try:
            start_idx = raw_xml.find('<rss')
            if start_idx != -1: raw_xml = raw_xml[start_idx:]
            
            raw_xml = re.sub(r'&(?!(amp|lt|gt|quot|apos|#\d+);)', '&amp;', raw_xml)
            
            root = ET.fromstring(raw_xml)
            new_queue = []
            
            for item in root.findall('.//item'):
                link_elem = item.find('link')
                title_elem = item.find('title')
                if link_elem is not None and link_elem.text:
                    title_text = title_elem.text.strip() if title_elem is not None and title_elem.text else "未知"
                    new_queue.append((title_text, link_elem.text.strip()))
                    
            if not new_queue: 
                raise ValueError("在 RSS 中找不到任何有效項目")
            
            new_queue.reverse()
            self.target_queue = new_queue
            
            self.root.after(0, self.update_queue_display)
            self.root.after(0, lambda c=len(new_queue): self.append_log(f"成功解析 XML，已將 {c} 話加入佇列。\n"))
            self.root.after(0, lambda: self.set_status("解析完成！請點擊「開始佇列下載」"))

        except Exception as exc:
            self.root.after(0, lambda e=exc: self.append_log(f"XML 解析失敗：{e}\n"))
            self.root.after(0, lambda: self.set_status("解析失敗"))

    # ==========================================
    # 階段二：整併的下載排程器
    # ==========================================
    def start_queue_download(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
        if not self.target_queue:
            messagebox.showwarning("警告", "佇列為空，請先解析 RSS 或是單話網址。")
            return
        if not self.save_dir_var.get().strip():
            messagebox.showerror("錯誤", "請先指定下載資料夾")
            return

        self.is_running = True
        self.reset_control_state()
        self.clear_log()
        self.set_status(f"開始下載佇列中的 {len(self.target_queue)} 個項目...")
        
        # 啟動唯一一個 Thread 來執行整個佇列
        threading.Thread(target=self._queue_download_worker, daemon=True).start()

    def _queue_download_worker(self) -> None:
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        self._loop = loop
        try:
            loop.run_until_complete(self._process_queue_async())
        except Exception as exc:
            self.root.after(0, lambda e=exc: self.append_log(f"佇列處理發生嚴重錯誤: {e}\n"))
            self.root.after(0, lambda: self.set_status("下載異常中斷"))
        finally:
            loop.close()
            self._loop = None
            self.is_running = False
            self.root.after(0, self.reset_control_state)

    async def _process_queue_async(self):
        total_items = len(self.target_queue)
        
        self.pause_event_async = asyncio.Event()
        self.pause_event_async.set()

        # ============================================================
        # 核心改動：在這裡啟動 Playwright，並貫穿整個 Queue 的下載生命週期
        # ============================================================
        async with async_playwright() as p:
            browser = await p.chromium.launch(headless=False)
            context = await browser.new_context(viewport={"width": 1400, "height": 900}, user_agent=UA)
            
            for idx, (chapter_title, chapter_url) in enumerate(self.target_queue, 1):
                if self.cancel_event.is_set(): break
                
                self.root.after(0, lambda i=idx, tot=total_items, t=chapter_title: self.append_log(f"\n{'='*40}\n>>> 開始處理 ({i}/{tot})：{t}\n"))
                self.root.after(0, lambda u=chapter_url: self.url_var.set(u))
                
                max_page = self._sync_fetch_max_page(chapter_url)
                if max_page == 0:
                    self.root.after(0, lambda: self.append_log("⚠️ 無法事先取得最大頁數，啟用動態偵測 (上限 999 頁)\n"))
                    max_page = 999
                    
                self.root.after(0, lambda: self.start_page_var.set("1"))
                self.root.after(0, lambda m=max_page: self.end_page_var.set(str(m)))
                
                await asyncio.sleep(1) # 讓 UI 有時間更新變數
                
                # 傳入 shared context 進行單話下載
                await self._download_chapter_async(context, chapter_title, 1, max_page)
                
                self.root.after(0, lambda t=chapter_title: self.append_log(f"<<< {t} 處理完成\n{'='*40}\n"))
            
            # 整個佇列跑完後才關閉瀏覽器
            await browser.close()
            
        if not self.cancel_event.is_set():
            self.root.after(0, lambda: self.set_status("佇列下載完成！"))
            self.root.after(0, lambda: messagebox.showinfo("完成", "佇列中所有項目已下載完成！"))
        else:
            self.root.after(0, lambda: self.set_status("任務已終止"))
            self.root.after(0, lambda: self.append_log("\n=== 使用者手動終止了任務 ===\n"))

    # ==========================================
    # 核心下載邏輯 (Async Playwright)
    # ==========================================
    async def get_playwright_cookies(self, context) -> dict:
        return {c['name']: c['value'] for c in await context.cookies()}

    def download_urls_requests(self, session: requests.Session, headers: dict, urls: list[str], page_num: int, save_dir: str) -> tuple[int, int]:
        success_count, fail_count = 0, 0
        for idx, url in enumerate(urls, 1):
            if url in self.downloaded_urls: continue
            if self.cancel_event.is_set(): break
            
            path = os.path.join(save_dir, self.build_filename(idx, page_num, url))
            if os.path.exists(path):
                self.downloaded_urls.add(url)
                success_count += 1
                self.root.after(0, lambda p=path: self.append_log(f"檔案已存在，跳過下載：{p}\n"))
                continue
                
            try:
                response = session.get(url, headers=headers, timeout=30)
                response.raise_for_status()
                with open(path, "wb") as f: f.write(response.content)
                self.downloaded_urls.add(url)
                success_count += 1
                self.root.after(0, lambda p=path: self.append_log(f"已下載：{p}\n"))
            except Exception as exc:
                fail_count += 1
                self.root.after(0, lambda u=url, e=exc: self.append_log(f"下載失敗：{u}\n原因：{e}\n"))
        return success_count, fail_count

    async def collect_dom_urls_async(self, page, min_width: int, min_height: int) -> list[str]:
        return await page.evaluate(
            """({minWidth, minHeight}) => {
                return [...document.images].map(img => {
                    const url = img.currentSrc || img.src || "";
                    return { url, w: img.naturalWidth || 0, h: img.naturalHeight || 0, complete: !!img.complete };
                }).filter(x => {
                    if (!x.complete || !x.url) return false;
                    const lowerUrl = x.url.toLowerCase();
                    if (lowerUrl.includes("404.png") || lowerUrl.includes("error")) return true;
                    return x.w >= minWidth && x.h >= minHeight;
                }).map(x => x.url);
            }""", {"minWidth": min_width, "minHeight": min_height})

    async def download_single_page_async(self, context, page_num: int, semaphore: asyncio.Semaphore, final_dir: str):
        if getattr(self, 'chapter_end_flag', False):
            return 0, 0

        template, scroll_times, scroll_wait, timeout, min_width, min_height = self.prepare_single_find()
        current_url = self.build_url_for_page(page_num)
        
        async with semaphore:
            if getattr(self, 'chapter_end_flag', False):
                return 0, 0

            await self.wait_if_paused()
            page = await context.new_page()
            page.set_default_timeout(timeout)
            self.root.after(0, lambda n=page_num, u=current_url: self.append_log(f"\n[頁 {n}] 開始處理：{u}\n"))
            
            try:
                await page.goto(current_url, wait_until="domcontentloaded")
            except PlaywrightTimeoutError:
                pass
            except Exception as e:
                if "Target closed" not in str(e):
                    self.root.after(0, lambda n=page_num, err=e: self.append_log(f"[頁 {n}] 開啟失敗：{err}\n"))
                await page.close()
                return 0, 0
                
            start_page = self.get_int(self.start_page_var.get(), "起始頁")
            if page_num == start_page and self.manual_verify_var.get():
                self.root.after(0, lambda n=page_num: self.append_log(f"[頁 {n}] 觸發手動通關，已自動暫停。\n"))
                self.pause_event_async.clear()
                self.is_paused = True
                self.root.after(0, lambda: self.pause_btn.config(text="▶ 繼續"))
                self.root.after(0, lambda: self.set_status("已暫停 (等待手動繼續...)"))
                self.root.after(0, lambda: self.manual_verify_var.set(False))
            
            await self.wait_if_paused()
            await page.wait_for_timeout(2500)

            for _ in range(scroll_times):
                await self.wait_if_paused()
                await page.mouse.wheel(0, 2000)
                await page.wait_for_timeout(scroll_wait)

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
                
                if any("404.png" in u.lower() or "error" in u.lower() for u in current_urls):
                    self.chapter_end_flag = True
                    self.root.after(0, lambda n=page_num: self.append_log(f"[頁 {n}] 偵測到 404 錯誤圖片，標記為本話結束。\n"))
                    await page.close()
                    return 0, 0

                if current_urls:
                    break 
                else:
                    if retry_count < max_retries:
                        self.root.after(0, lambda n=page_num, r=retry_count+1: self.append_log(f"[頁 {n}] 未抓到圖片，500ms 後第 {r} 次重試...\n"))
                        await page.wait_for_timeout(500)
                    else:
                        self.chapter_end_flag = True
                        self.root.after(0, lambda n=page_num: self.append_log(f"[頁 {n}] 持續無畫面，標記為本話結尾。\n"))
                        await page.close()
                        return 0, 0

            if current_urls:
                current_set = frozenset(current_urls)
                for prev_page, prev_urls in self.page_images.items():
                    if prev_page != page_num and current_set == frozenset(prev_urls):
                        self.chapter_end_flag = True
                        self.root.after(0, lambda n=page_num, p=prev_page: self.append_log(f"[頁 {n}] 圖片內容與第 {p} 頁完全相同，觸發備案機制，標記為本話結束。\n"))
                        await page.close()
                        return 0, 0

            self.page_images[page_num] = current_urls
            self.root.after(0, lambda n=page_num, c=len(current_urls): self.append_log(f"[頁 {n}] 找到 {c} 張圖片\n"))

            await self.wait_if_paused()
            playwright_cookies = await self.get_playwright_cookies(context)
            session = requests.Session()
            session.cookies.update(playwright_cookies)
            
            if self.cancel_event.is_set():
                await page.close()
                raise asyncio.CancelledError()

            success, fail = self.download_urls_requests(session, {"User-Agent": UA, "Referer": current_url}, current_urls, page_num, final_dir)
            self.root.after(0, lambda n=page_num, s=success, f=fail: self.append_log(f"[頁 {n}] 已下載 {s}，失敗 {f}\n"))
            
            await page.close()
            return success, fail

    async def _download_chapter_async(self, context, batch_title: str, start_page: int, end_page: int) -> None:
        self.chapter_end_flag = False 
        self.page_images.clear()

        max_concurrent = self.get_int(self.max_concurrent_var.get(), "同時最大頁數")
        self.root.after(0, lambda: self.set_status(f"正在並行下載 ({max_concurrent} 頁同時進行)..."))

        base_dir = self.save_dir_var.get().strip()
        final_dir = base_dir
        
        if self.use_title_chapter_dir_var.get():
            if batch_title:
                safe_title = self.sanitize_path_part(batch_title)
                final_dir = os.path.join(base_dir, safe_title)
                os.makedirs(final_dir, exist_ok=True)
                self.root.after(0, lambda d=final_dir: self.append_log(f"下載資料夾：{d}\n"))
            else:
                os.makedirs(final_dir, exist_ok=True)

        total_success, total_fail = 0, 0
        
        semaphore = asyncio.Semaphore(max_concurrent)
        
        tasks = [self.download_single_page_async(context, pn, semaphore, final_dir) for pn in range(start_page, end_page + 1)]
        results = await asyncio.gather(*tasks, return_exceptions=True)
        
        for res in results:
            if isinstance(res, tuple):
                total_success += res[0]
                total_fail += res[1]

        all_urls = []
        for urls_in_page in self.page_images.values():
            for url in urls_in_page:
                if url not in all_urls: all_urls.append(url)
        self.urls = all_urls

        self.root.after(0, lambda s=total_success, f=total_fail: self.append_log(f"單話下載完成，成功 {s}，失敗 {f}\n"))


def main() -> None:
    root = tk.Tk()
    app = DM5CrawlerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()