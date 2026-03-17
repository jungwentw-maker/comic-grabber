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
# 2. 不要預設網址 (留空)
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
        self.root.geometry("1120x920") # 稍微拉高一點容納新 UI

        self.urls: list[str] = []
        self.page_images: dict[int, list[str]] = {}
        self.downloaded_urls: set[str] = set()
        self.is_running = False
        # 新增圖片格式選擇變數
        self.image_format_var = tk.StringVar(value="JPG & PNG")

        # ==========================================
        # 4. 預留 RSS 輸入空格
        # ==========================================
        self.rss_url_var = tk.StringVar(value="")
        
        self.url_var = tk.StringVar(value=DEFAULT_URL)
        self.status_var = tk.StringVar(value="就緒")
        
        # ==========================================
        # 3. 預設捲動次數為 0
        # ==========================================
        self.scroll_times_var = tk.StringVar(value="0")
        self.scroll_wait_var = tk.StringVar(value="1200")
        
        # ==========================================
        # 1. 預設頁面 timeout 改成 100ms
        # ==========================================
        self.timeout_var = tk.StringVar(value="100")
        self.save_dir_var = tk.StringVar(value=os.path.join(os.getcwd(), "downloads"))
        self.min_width_var = tk.StringVar(value="300")
        self.min_height_var = tk.StringVar(value="300")
        self.start_page_var = tk.StringVar(value="1")
        self.end_page_var = tk.StringVar(value="5")
        self.use_title_chapter_dir_var = tk.BooleanVar(value=False)
        
        # ==========================================
        # 5. 自行設定同時下載的最大頁數
        # ==========================================
        self.max_concurrent_var = tk.StringVar(value="5")

        self._build_ui()

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=10)
        top.pack(fill="x")

        # --- 新增 RSS 輸入區塊 ---
        ttk.Label(top, text="RSS 網址 (批次)").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.rss_url_var, width=110).grid(
            row=0, column=1, columnspan=6, sticky="ew", padx=6
        )
        ttk.Button(top, text="批次下載 RSS 列表", command=self.start_batch_rss).grid(
            row=0, column=7, sticky="e", padx=6
        )

        ttk.Label(top, text="DM5 網址").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.url_var, width=110).grid(
            row=1, column=1, columnspan=7, sticky="ew", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="起始頁").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.start_page_var, width=8).grid(
            row=2, column=1, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="結束頁").grid(row=2, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.end_page_var, width=8).grid(
            row=2, column=3, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="最小寬").grid(row=2, column=4, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.min_width_var, width=8).grid(
            row=2, column=5, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="最小高").grid(row=2, column=6, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.min_height_var, width=8).grid(
            row=2, column=7, sticky="w", padx=6, pady=(8, 0)
        )
        
        ttk.Label(top, text="圖片格式").grid(row=2, column=8, sticky="w", pady=(8, 0), padx=(10, 0))
        format_cb = ttk.Combobox(
            top, textvariable=self.image_format_var, values=["僅 JPG", "僅 PNG", "JPG & PNG"], width=10, state="readonly"
        )
        format_cb.grid(row=2, column=9, sticky="w", padx=6, pady=(8, 0))

        ttk.Label(top, text="捲動次數").grid(row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.scroll_times_var, width=8).grid(
            row=3, column=1, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="每次等待(ms)").grid(row=3, column=2, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.scroll_wait_var, width=10).grid(
            row=3, column=3, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="頁面 timeout(ms)").grid(row=3, column=4, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.timeout_var, width=12).grid(
            row=3, column=5, sticky="w", padx=6, pady=(8, 0)
        )
        
        # --- 新增同時最大頁數設定 ---
        ttk.Label(top, text="同時最大頁數").grid(row=3, column=6, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.max_concurrent_var, width=8).grid(
            row=3, column=7, sticky="w", padx=6, pady=(8, 0)
        )

        ttk.Label(top, text="下載資料夾").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(top, textvariable=self.save_dir_var, width=90).grid(
            row=4, column=1, columnspan=5, sticky="ew", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="選擇", command=self.choose_dir).grid(
            row=4, column=6, sticky="w", pady=(8, 0)
        )
        ttk.Checkbutton(
            top,
            text="自動建立資料夾（標題/篇章）",
            variable=self.use_title_chapter_dir_var,
        ).grid(row=4, column=7, sticky="w", padx=6, pady=(8, 0))

        ttk.Button(top, text="顯示自動轉換範本", command=self.preview_template).grid(
            row=5, column=3, sticky="e", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="自動抓最大頁數", command=self.fetch_max_page).grid(
            row=5, column=4, sticky="e", padx=6, pady=(8, 0)
        )
        # 註：因為全面改為 async，若保留 start_find 也需改寫。這裡保留結構，但如果主要用逐頁下載可先專注於 start_download_range
        ttk.Button(top, text="開始找當前頁", command=self.start_find).grid(
            row=5, column=5, sticky="e", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="並行逐頁下載範圍", command=self.start_download_range).grid(
            row=5, column=6, sticky="e", padx=6, pady=(8, 0)
        )
        ttk.Button(top, text="下載目前列表", command=self.start_download_current).grid(
            row=5, column=7, sticky="e", padx=6, pady=(8, 0)
        )

        top.columnconfigure(1, weight=1)

        middle = ttk.Frame(self.root, padding=(10, 0, 10, 10))
        middle.pack(fill="both", expand=True)

        ttk.Label(middle, textvariable=self.status_var).pack(anchor="w", pady=(0, 8))
        self.result_text = ScrolledText(middle, wrap="word")
        self.result_text.pack(fill="both", expand=True)

    def start_batch_rss(self) -> None:
        if self.is_running:
            messagebox.showinfo("提示", "目前有工作進行中")
            return
            
        rss_url = self.rss_url_var.get().strip()
        if not rss_url:
            messagebox.showerror("錯誤", "請先輸入 RSS 網址")
            return
            
        save_dir = self.save_dir_var.get().strip()
        if not save_dir:
            messagebox.showerror("錯誤", "請先指定下載資料夾")
            return
            
        self.is_running = True
        self.clear_output()
        self.set_status("正在解析 RSS 列表...")
        threading.Thread(target=self.batch_rss_worker, args=(rss_url,), daemon=True).start()

    def batch_rss_worker(self, rss_url: str) -> None:
        try:
            # 1. 抓取並解析 RSS XML
            headers = {"User-Agent": UA}
            response = requests.get(rss_url, headers=headers, timeout=30)
            response.raise_for_status()
            
            root = ET.fromstring(response.text)
            
            # 萃取所有 <link>
            target_urls = []
            for item in root.findall('.//item'):
                link_elem = item.find('link')
                title_elem = item.find('title')
                if link_elem is not None and link_elem.text:
                    url = link_elem.text.strip()
                    title = title_elem.text.strip() if title_elem is not None else "未知章節"
                    target_urls.append((title, url))
                    
            if not target_urls:
                raise ValueError("在 RSS 中找不到任何項目")
                
            # 反轉列表：讓它從第 1 話開始下載
            target_urls.reverse()
            
            self.root.after(0, lambda c=len(target_urls): self.append_text(f"成功解析 RSS，共 {c} 話準備下載。\n"))
            self.root.after(0, lambda: self.append_text("-" * 40 + "\n"))
            
            # 2. 依序執行每一話的下載任務
            for idx, (chapter_title, chapter_url) in enumerate(target_urls, 1):
                self.root.after(0, lambda t=chapter_title: self.append_text(f"\n>>> 開始處理：{t}\n"))
                
                # 更新目前的 URL (讓原本的邏輯能吃到)
                self.root.after(0, lambda u=chapter_url: self.url_var.set(u))
                
                # 同步呼叫自動抓最大頁數 (不用 Thread，因為這裡已經是在背景 Thread)
                max_page = self._sync_fetch_max_page(chapter_url)
                if max_page == 0:
                    self.root.after(0, lambda t=chapter_title: self.append_text(f"跳過 {t}：無法取得最大頁數\n"))
                    continue
                    
                self.root.after(0, lambda: self.start_page_var.set("1"))
                self.root.after(0, lambda m=max_page: self.end_page_var.set(str(m)))
                
                # 等待一點時間讓 UI 更新變數
                import time
                time.sleep(1)
                
                # 直接呼叫非同步的下載工作，傳入 is_batch=True 避免被 on_download_done 中斷狀態
                asyncio.run(self.download_range_worker_async(is_batch=True))
                
                self.root.after(0, lambda t=chapter_title: self.append_text(f"<<< {t} 處理完成\n"))
                
            self.is_running = False
            self.root.after(0, lambda: self.set_status("批次 RSS 下載完成！"))
            self.root.after(0, lambda: messagebox.showinfo("完成", "批次 RSS 下載完成！"))

        except Exception as exc:
            self.is_running = False
            self.root.after(0, lambda e=exc: self.append_text(f"批次解析失敗：{e}\n"))
            self.root.after(0, lambda: self.set_status("批次作業失敗"))

    def _sync_fetch_max_page(self, raw_url: str) -> int:
        """同步版本的抓取最大頁數，專供批次排程使用"""
        try:
            template = self.normalize_dm5_template(raw_url)
            page1_url = template.replace("(#)", "1")
            headers = {"User-Agent": UA, "Referer": page1_url}
            response = requests.get(page1_url, headers=headers, timeout=30)
            response.raise_for_status()
            return self.extract_max_page_from_html(response.text)
        except Exception as e:
            self.root.after(0, lambda err=e: self.append_text(f"嘗試抓取最大頁數時發生錯誤: {err}\n"))
            return 0

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
        if m: return int(m.group(1))
        pager_matches = re.findall(r"/m\d+-p(\d+)/", html, re.IGNORECASE)
        if pager_matches: return max(int(x) for x in pager_matches)
        ipg_matches = re.findall(r"#ipg(\d+)", html, re.IGNORECASE)
        if ipg_matches: return max(int(x) for x in ipg_matches)
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
        
        # 根據 UI 選項決定正則表達式
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

    # 由於 Playwright API 轉換為 async，這裡改寫成 async 方法
    async def collect_dom_urls_async(self, page, min_width: int, min_height: int) -> list[str]:
        return await page.evaluate(
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
            html, re.IGNORECASE | re.DOTALL,
        )
        block = block_match.group(1) if block_match else html

        pair_match = re.search(
            r'<span class="right-arrow">\s*<a [^>]*>(.*?)</a>\s*</span>\s*'
            r'<span class="active right-arrow">\s*(.*?)\s*</span>',
            block, re.IGNORECASE | re.DOTALL,
        )
        if pair_match:
            title_text = self.clean_html_text(pair_match.group(1))
            chapter_text = self.clean_html_text(pair_match.group(2))

        if not title_text:
            m = re.search(r'<span class="right-arrow">\s*<a [^>]*title="([^"]+)"[^>]*>', block, re.IGNORECASE | re.DOTALL)
            if m: title_text = self.clean_html_text(m.group(1))

        if not chapter_text:
            m = re.search(r'<span class="active right-arrow">\s*(.*?)\s*</span>', block, re.IGNORECASE | re.DOTALL)
            if m: chapter_text = self.clean_html_text(m.group(1))

        if (not title_text or not chapter_text):
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

    # ---------------------------------------------------------
    # 因為 Playwright 已改為 Async，單頁尋找圖的功能也需使用 Async (為了不讓程式碼太長，這裡用 asyncio.run 包裝原本的邏輯)
    # ---------------------------------------------------------
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
        threading.Thread(target=self._run_find_async, daemon=True).start()

    def _run_find_async(self):
        asyncio.run(self.find_images_worker_async())

    async def find_images_worker_async(self) -> None:
        try:
            template, scroll_times, scroll_wait, timeout, min_width, min_height = self.prepare_single_find()
            current_url = self.build_url_for_page(self.get_int(self.start_page_var.get(), "起始頁"))
            
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=False)
                context = await browser.new_context(viewport={"width": 1400, "height": 900}, user_agent=UA)
                page = await context.new_page()
                page.set_default_timeout(timeout)
                
                try:
                    await page.goto(current_url, wait_until="domcontentloaded")
                except PlaywrightTimeoutError:
                    pass # 忽略 timeout，繼續執行後續動作

                await page.wait_for_timeout(3000)
                for _ in range(scroll_times):
                    await page.mouse.wheel(0, 2000)
                    await page.wait_for_timeout(scroll_wait)
                
                dom_urls = await self.collect_dom_urls_async(page, min_width, min_height)
                await browser.close()
                
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

    def build_filename(self, index: int, page_num: int, url: str) -> str:
        parsed = urlparse(url)
        original_name = os.path.basename(parsed.path)
        base, ext = os.path.splitext(original_name)
        safe_base = re.sub(r"[^a-zA-Z0-9._-]+", "_", base).strip("_") or f"image_{index:03d}"
        
        # 保留原本的副檔名，如果是 .jpeg 就轉 .jpg
        ext = ext.lower()
        if ext == ".jpeg":
            ext = ".jpg"
        elif ext not in {".jpg", ".png"}:
            ext = ".jpg" # 預防萬一的 fallback
            
        return f"p{page_num:03d}_{index:03d}_{safe_base}{ext}"

    async def get_playwright_cookies(self, context) -> dict:
        cookies = await context.cookies()
        cookie_dict = {}
        for cookie in cookies:
             cookie_dict[cookie['name']] = cookie['value']
        return cookie_dict

    def download_urls_requests(
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

    # ==========================================
    # 核心改寫：並行逐頁下載範圍 (Async)
    # ==========================================
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
            max_concurrent = self.get_int(self.max_concurrent_var.get(), "同時最大頁數")
            if start_page <= 0 or end_page <= 0:
                raise ValueError("頁碼必須大於 0")
            if start_page > end_page:
                raise ValueError("起始頁不可大於結束頁")
            if max_concurrent <= 0:
                raise ValueError("同時最大頁數必須大於 0")
            self.prepare_single_find()
        except ValueError as exc:
            messagebox.showerror("錯誤", str(exc))
            return
            
        os.makedirs(save_dir, exist_ok=True)
        self.is_running = True
        self.clear_output()
        self.page_images.clear()
        self.downloaded_urls.clear()
        self.set_status(f"正在並行下載 ({max_concurrent} 頁同時進行)...")
        
        # 使用 Thread 啟動 asyncio 的事件迴圈
        threading.Thread(target=self._run_download_range_async, daemon=True).start()

    def _run_download_range_async(self):
        asyncio.run(self.download_range_worker_async())

    async def download_single_page_async(self, context, page_num: int, semaphore: asyncio.Semaphore, final_dir: str):
        """處理單一頁面的解析與下載"""
        template, scroll_times, scroll_wait, timeout, min_width, min_height = self.prepare_single_find()
        current_url = self.build_url_for_page(page_num)
        
        async with semaphore:
            # 建立新分頁 (共用 context)
            page = await context.new_page()
            page.set_default_timeout(timeout)
            
            self.root.after(0, lambda n=page_num, u=current_url: self.append_text(f"\n[頁 {n}] 開始處理：{u}\n"))
            
            try:
                # 這裡設定 timeout 為 UI 輸入的 100ms，通常會 Timeout
                await page.goto(current_url, wait_until="domcontentloaded")
            except PlaywrightTimeoutError:
                # 攔截 TimeoutError 繼續執行，因為 100ms 幾乎一定會觸發
                pass
            except Exception as e:
                self.root.after(0, lambda n=page_num, err=e: self.append_text(f"[頁 {n}] 開啟失敗：{err}\n"))
                await page.close()
                return 0, 0
                
            # 給一點緩衝時間讓圖片載入
            await page.wait_for_timeout(2500)

            for _ in range(scroll_times):
                await page.mouse.wheel(0, 2000)
                await page.wait_for_timeout(scroll_wait)

            dom_urls = await self.collect_dom_urls_async(page, min_width, min_height)
            
            # 過濾網址
            current_urls: list[str] = []
            seen = set()
            for url in dom_urls:
                if self.is_probably_comic_image(url) and url not in seen:
                    seen.add(url)
                    current_urls.append(url)

            self.page_images[page_num] = current_urls
            self.root.after(0, lambda n=page_num, c=len(current_urls): self.append_text(f"[頁 {n}] 找到 {c} 張 jpg\n"))

            # 取得該 Context 的 Cookie 提供給 requests 使用
            playwright_cookies = await self.get_playwright_cookies(context)
            
            # 下載圖片 (使用 requests 於同步執行緒中，因量不大暫以迴圈處理)
            session = requests.Session()
            session.cookies.update(playwright_cookies)
            headers = {"User-Agent": UA, "Referer": current_url}
            
            success, fail = self.download_urls_requests(session, headers, current_urls, page_num, final_dir)
            self.root.after(0, lambda n=page_num, s=success, f=fail: self.append_text(f"[頁 {n}] 已下載 {s}，失敗 {f}\n"))
            
            await page.close()
            return success, fail

    async def download_range_worker_async(self, is_batch: bool = False) -> None:
        start_page = self.get_int(self.start_page_var.get(), "起始頁")
        end_page = self.get_int(self.end_page_var.get(), "結束頁")
        max_concurrent = self.get_int(self.max_concurrent_var.get(), "同時最大頁數")
        template, _, _, _, _, _ = self.prepare_single_find()

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

        # 啟動 Playwright 並行處理
        try:
            async with async_playwright() as p:
                browser = await p.chromium.launch(headless=False) # 測試期可開著看，穩定了可改 True
                context = await browser.new_context(viewport={"width": 1400, "height": 900}, user_agent=UA)
                
                # 建立 Semaphore 控制並行數量
                semaphore = asyncio.Semaphore(max_concurrent)
                
                tasks = []
                for page_num in range(start_page, end_page + 1):
                    # 建立任務
                    task = self.download_single_page_async(context, page_num, semaphore, final_dir)
                    tasks.append(task)
                
                # 等待所有頁面處理完成
                results = await asyncio.gather(*tasks, return_exceptions=True)
                
                # 統計結果
                for res in results:
                    if isinstance(res, tuple):
                        total_success += res[0]
                        total_fail += res[1]

                await browser.close()

            # 更新所有抓到過的 URL 給 UI
            all_urls = []
            for urls_in_page in self.page_images.values():
                for url in urls_in_page:
                    if url not in all_urls:
                        all_urls.append(url)
            self.urls = all_urls

            self.root.after(0, lambda: self.on_download_done(total_success, total_fail, final_dir, is_batch))
            
        except Exception as exc:
            err = exc
            self.root.after(0, lambda err=err: self.on_find_error(err))

    def on_download_done(self, success_count: int, fail_count: int, save_dir: str, is_batch: bool = False) -> None:
        if not is_batch:
            self.is_running = False
            self.set_status(f"下載完成，成功 {success_count}，失敗 {fail_count}")
            messagebox.showinfo(
                "完成",
                f"下載完成\n成功：{success_count}\n失敗：{fail_count}\n資料夾：{save_dir}",
            )
        else:
            self.append_text(f"單話下載完成，成功 {success_count}，失敗 {fail_count}\n")


def main() -> None:
    root = tk.Tk()
    app = DM5CrawlerApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()