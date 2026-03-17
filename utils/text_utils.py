import html as html_lib
import re


def clean_html_text(text: str) -> str:
    text = re.sub(r"<[^>]+>", "", text)
    text = html_lib.unescape(text)
    return re.sub(r"\s+", " ", text).strip()


def sanitize_path_part(text: str) -> str:
    text = clean_html_text(text)
    text = re.sub(r'[\\/:*?"<>|]+', "_", text).strip(" .")
    return text[:120] if text else ""
