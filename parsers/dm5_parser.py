import re

from utils.text_utils import clean_html_text


def normalize_dm5_template(raw_url: str) -> str:
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


def extract_title_and_chapter_from_html(html: str) -> tuple[str, str]:
    title_text, chapter_text = "", ""
    block_match = re.search(
        r'<div class="title">(.*?)</div>\s*<div class="right-bar">',
        html,
        re.IGNORECASE | re.DOTALL,
    )
    block = block_match.group(1) if block_match else html

    pair_match = re.search(
        r'<span class="right-arrow">\s*<a [^>]*>(.*?)</a>\s*</span>\s*<span class="active right-arrow">\s*(.*?)\s*</span>',
        block,
        re.IGNORECASE | re.DOTALL,
    )
    if pair_match:
        title_text = clean_html_text(pair_match.group(1))
        chapter_text = clean_html_text(pair_match.group(2))

    if not title_text:
        m = re.search(
            r'<span class="right-arrow">\s*<a [^>]*title="([^"]+)"[^>]*>',
            block,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            title_text = clean_html_text(m.group(1))

    if not chapter_text:
        m = re.search(
            r'<span class="active right-arrow">\s*(.*?)\s*</span>',
            block,
            re.IGNORECASE | re.DOTALL,
        )
        if m:
            chapter_text = clean_html_text(m.group(1))

    if not title_text or not chapter_text:
        m = re.search(r'DM5_CTITLE\s*=\s*"([^"]+)"', html, re.IGNORECASE)
        if m:
            ctitle = clean_html_text(m.group(1))
            m2 = re.search(r"(第\s*\d+\s*话.*)$", ctitle)
            if m2:
                if not chapter_text:
                    chapter_text = clean_html_text(m2.group(1))
                if not title_text:
                    title_text = clean_html_text(ctitle[: m2.start()].strip())
            elif not title_text:
                title_text = ctitle

    return title_text, chapter_text


def extract_max_page_from_html(html: str) -> int:
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
