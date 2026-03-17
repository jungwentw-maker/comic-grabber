import re
import xml.etree.ElementTree as ET


def parse_rss_xml_to_queue(raw_xml: str) -> list[tuple[str, str]]:
    start_idx = raw_xml.find('<rss')
    if start_idx != -1:
        raw_xml = raw_xml[start_idx:]

    raw_xml = re.sub(r'&(?!(amp|lt|gt|quot|apos|#\d+);)', '&amp;', raw_xml)

    root = ET.fromstring(raw_xml)
    new_queue: list[tuple[str, str]] = []

    for item in root.findall('.//item'):
        link_elem = item.find('link')
        title_elem = item.find('title')
        if link_elem is not None and link_elem.text:
            title_text = title_elem.text.strip() if title_elem is not None and title_elem.text else "未知"
            new_queue.append((title_text, link_elem.text.strip()))

    if not new_queue:
        raise ValueError("在 RSS 中找不到任何有效項目")

    new_queue.reverse()
    return new_queue
