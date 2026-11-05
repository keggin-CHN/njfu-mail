################宏定义开始#################
import os
USERNAME = os.getenv("NJFU_USERNAME", "your_username")
PASSWORD = os.getenv("NJFU_PASSWORD", "your_password")
DOMAIN = os.getenv("NJFU_DOMAIN", "njfu.edu.cn")
EXPORT_START_DATE = os.getenv("NJFU_START_DATE", "2026-10-30")
EXPORT_END_DATE = os.getenv("NJFU_END_DATE", "2026-11-06")
OUTPUT_DOCX_PATH = "收件箱邮件导出.docx"
ATTACHMENTS_DIR = "收件箱邮件附件"
#####################宏定义结束###################################

import base64
import hashlib
import html
import os
import re
import sys
import time
from datetime import datetime, timedelta
from email.header import decode_header
from html.parser import HTMLParser
from urllib.parse import parse_qs, unquote, unquote_to_bytes, urljoin, urlparse

from curl_cffi import requests
from Crypto.Cipher import PKCS1_v1_5
from Crypto.PublicKey import RSA
from docx import Document

LOGIN_URL = "https://mail.njfu.edu.cn/"
POST_LOGIN_URL = "https://mail.njfu.edu.cn/cgi-bin/login"
MAIL_LIST_URL = "https://mail.njfu.edu.cn/cgi-bin/mail_list"
READ_MAIL_URL = "https://mail.njfu.edu.cn/cgi-bin/readmail"
ATTACH_URL = "https://mail.njfu.edu.cn/cgi-bin/download"
RSA_N_HEX = "CF87D7B4C864F4842F1D337491A48FFF54B73A17300E8E42FA365420393AC0346AE55D8AFAD975DFA175FAF0106CBA81AF1DDE4ACEC284DAC6ED9A0D8FEB1CC070733C58213EFFED46529C54CEA06D774E3CC7E073346AEBD6C66FC973F299EB74738E400B22B1E7CDC54E71AED059D228DFEB5B29C530FF341502AE56DDCFE9"
RSA_E_HEX = "10001"


class SimpleHTMLTextExtractor(HTMLParser):
    BLOCK_TAGS = {"p", "div", "br", "tr", "table", "li", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6"}

    def __init__(self):
        super().__init__()
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag == "br":
            self.parts.append("\n")

    def handle_endtag(self, tag):
        if tag in self.BLOCK_TAGS:
            self.parts.append("\n")

    def handle_data(self, data):
        if data:
            self.parts.append(data)

    def get_text(self):
        text = "".join(self.parts)
        return normalize_text(text)


def safe_print(message):
    encoding = sys.stdout.encoding or "utf-8"
    print(str(message).encode(encoding, errors="replace").decode(encoding, errors="replace"))


def extract_first(pattern, text, default=""):
    match = re.search(pattern, text, re.S)
    return match.group(1).strip() if match else default


def parse_export_range():
    start_date = datetime.strptime(EXPORT_START_DATE, "%Y-%m-%d")
    end_date = datetime.strptime(EXPORT_END_DATE, "%Y-%m-%d") + timedelta(days=1) - timedelta(milliseconds=1)
    if end_date < start_date:
        raise RuntimeError("导出结束时间不能早于开始时间")
    return int(start_date.timestamp() * 1000), int(end_date.timestamp() * 1000)


def normalize_text(text):
    text = html.unescape(text)
    text = text.replace("\xa0", " ")
    text = text.replace("\u2003", " ")
    text = text.replace("\u2002", " ")
    text = text.replace("\u2009", " ")
    text = text.replace("\ufeff", "")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"[ \t]{2,}", " ", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def decode_filename(name):
    if not name:
        return name

    value = html.unescape(str(name)).strip()
    value = value.replace("+", " ")

    if re.match(r"=\?.+\?=", value):
        try:
            decoded_parts = []
            for chunk, encoding in decode_header(value):
                if isinstance(chunk, bytes):
                    decoded_parts.append(chunk.decode(encoding or "utf-8", errors="replace"))
                else:
                    decoded_parts.append(chunk)
            value = "".join(decoded_parts)
        except Exception:
            pass

    for _ in range(2):
        try:
            raw_bytes = unquote_to_bytes(value)
            if b"%" in raw_bytes and raw_bytes == value.encode("utf-8", errors="ignore"):
                break
            try:
                new_value = raw_bytes.decode("utf-8")
            except UnicodeDecodeError:
                new_value = raw_bytes.decode("gb18030", errors="replace")
        except Exception:
            new_value = unquote(value)

        if new_value == value:
            break
        value = new_value

    mojibake_hints = ("Ã", "Â", "æ", "ç", "é", "ä", "å", "ö", "ñ")
    if any(hint in value for hint in mojibake_hints):
        for encoding in ("utf-8", "gb18030"):
            try:
                repaired = value.encode("latin1").decode(encoding)
                if repaired and repaired != value:
                    value = repaired
                    break
            except Exception:
                continue

    return value


def sanitize_filename(name, default="unnamed"):
    name = decode_filename(name)
    name = normalize_text(name) or default
    name = re.sub(r'[\\/:*?"<>|]', "_", name)
    name = name.strip(" .")
    return name or default


def ensure_unique_path(path):
    if not os.path.exists(path):
        return path
    base, ext = os.path.splitext(path)
    index = 1
    while True:
        candidate = f"{base}_{index}{ext}"
        if not os.path.exists(candidate):
            return candidate
        index += 1


def extract_filename_from_response(response):
    content_disposition = response.headers.get("Content-Disposition", "")
    if not content_disposition:
        return ""

    filename_star_match = re.search(r"filename\*\s*=\s*([^;]+)", content_disposition, re.I)
    if filename_star_match:
        value = filename_star_match.group(1).strip().strip('"')
        if "''" in value:
            value = value.split("''", 1)[1]
        return decode_filename(value)

    filename_match = re.search(r'filename\s*=\s*(?:"([^"]+)"|([^;]+))', content_disposition, re.I)
    if filename_match:
        value = filename_match.group(1) or filename_match.group(2) or ""
        return decode_filename(value.strip())

    return ""


def calculate_file_sha256(path):
    digest = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            digest.update(chunk)
    return digest.hexdigest()


def find_existing_file_by_hash(directory, content_hash):
    if not os.path.isdir(directory):
        return ""
    for entry in os.scandir(directory):
        if entry.is_file() and calculate_file_sha256(entry.path) == content_hash:
            return entry.path
    return ""


def extract_raw_query_param(query, key):
    match = re.search(r"(?:^|&)" + re.escape(key) + r"=([^&]*)", query)
    return match.group(1) if match else ""


def parse_attachment_links(detail_html, sid, mailid):
    attachments = []
    seen = set()

    patterns = [
        re.compile(r'<a[^>]+href="([^"]*?/cgi-bin/download[^"#]*)"[^>]*>(.*?)</a>', re.I | re.S),
        re.compile(r'<a[^>]+href="([^"]*?download[^"#]*mailid=' + re.escape(str(mailid)) + r'[^"#]*)"[^>]*>(.*?)</a>', re.I | re.S),
    ]

    for pattern in patterns:
        for href, label_html in pattern.findall(detail_html):
            label = normalize_text(re.sub(r"<[^>]+>", "", label_html))
            full_url = urljoin(LOGIN_URL, html.unescape(href))
            parsed = urlparse(full_url)
            query = parse_qs(parsed.query)
            attachment_id = query.get("id", [""])[0] or query.get("fileid", [""])[0]
            raw_filename = extract_raw_query_param(parsed.query, "filename")
            filename = decode_filename(raw_filename) or label or f"attachment_{len(attachments) + 1}"
            dedupe_key = ("attachment", attachment_id or sanitize_filename(filename).lower() or full_url)
            if dedupe_key in seen:
                continue
            seen.add(dedupe_key)
            attachments.append({
                "url": full_url,
                "filename": sanitize_filename(filename),
                "label": label,
            })

    js_pattern = re.compile(r"download(?:File|Attach)?\((.*?)\)", re.I | re.S)
    for args_blob in js_pattern.findall(detail_html):
        args = re.findall(r'"([^"]*)"|\'([^\']*)\'', args_blob)
        flat_args = [a or b for a, b in args]
        if not flat_args:
            continue
        attach_id = flat_args[0]
        filename = flat_args[1] if len(flat_args) > 1 else f"attachment_{len(attachments) + 1}"
        params = {"sid": sid, "mailid": mailid, "id": attach_id}
        full_url = ATTACH_URL + "?" + "&".join(f"{k}={requests.utils.quote(str(v), safe='')}" for k, v in params.items())
        dedupe_key = ("attachment", attach_id or sanitize_filename(filename).lower() or full_url)
        if dedupe_key in seen:
            continue
        seen.add(dedupe_key)
        attachments.append({
            "url": full_url,
            "filename": sanitize_filename(filename),
            "label": filename,
        })

    return attachments


def download_attachments(session, sid, mailid, detail_html, base_dir, sender_addr=""):
    attachments = parse_attachment_links(detail_html, sid, mailid)
    if not attachments:
        return []

    sender_dir_name = sanitize_filename(sender_addr or mailid)
    mail_dir = os.path.join(base_dir, sender_dir_name)
    os.makedirs(mail_dir, exist_ok=True)
    downloaded = []

    saved_hashes = set()

    for index, attachment in enumerate(attachments, start=1):
        url = attachment["url"]
        parsed = urlparse(url)
        query = parse_qs(parsed.query)
        if "sid" not in query:
            sep = "&" if parsed.query else "?"
            url = f"{url}{sep}sid={sid}"
        response = session.get(url)
        response.raise_for_status()

        content = response.content
        content_hash = hashlib.sha256(content).hexdigest()
        if content_hash in saved_hashes:
            continue

        header_filename = extract_filename_from_response(response)
        filename = sanitize_filename(header_filename or attachment["filename"] or f"attachment_{index}")
        preferred_path = os.path.join(mail_dir, filename)

        existing_same_content_path = find_existing_file_by_hash(mail_dir, content_hash)
        if existing_same_content_path:
            saved_hashes.add(content_hash)
            continue

        if os.path.exists(preferred_path):
            if calculate_file_sha256(preferred_path) == content_hash:
                saved_hashes.add(content_hash)
                continue
            save_path = ensure_unique_path(preferred_path)
        else:
            save_path = preferred_path

        with open(save_path, "wb") as f:
            f.write(content)

        saved_hashes.add(content_hash)
        downloaded.append({
            "filename": os.path.basename(save_path),
            "path": save_path,
            "size": len(content),
        })

    return downloaded


def format_size(num_bytes):
    units = ["B", "KB", "MB", "GB"]
    size = float(num_bytes)
    for unit in units:
        if size < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}"
        size /= 1024


def login_session():
    session = requests.Session(impersonate="chrome116")
    response = session.get(LOGIN_URL)
    response.raise_for_status()
    timestamp_match = re.search(r'nTimeStamp\s*:\s*"(\d+)"', response.text)
    if not timestamp_match:
        raise RuntimeError("未找到 nTimeStamp，无法登录")
    timestamp = timestamp_match.group(1)
    message = f"{PASSWORD}\n{timestamp}\n"
    pub_key = RSA.construct((int(RSA_N_HEX, 16), int(RSA_E_HEX, 16)))
    cipher = PKCS1_v1_5.new(pub_key)
    encrypted = cipher.encrypt(message.encode("utf-8"))
    p_b64 = base64.b64encode(encrypted).decode("utf-8")
    payload = {
        "uin": USERNAME,
        "domain": DOMAIN,
        "pp": p_b64,
        "verifycode": "",
        "f": "biz",
        "aliastype": "other",
        "p": p_b64,
        "ppp": "",
        "fun": "",
        "errtemplate": "new_index/index",
    }
    post_response = session.post(POST_LOGIN_URL, data=payload, allow_redirects=False)
    post_response.raise_for_status()
    sid_match = re.search(r'targetUrl\s*=\s*urlHead\s*\+\s*"frame_html\?sid=([^"&]+)', post_response.text)
    if not sid_match:
        raise RuntimeError("登录成功后未找到 sid")
    sid = sid_match.group(1)
    return session, sid


def fetch_inbox_page(session, sid, page):
    response = session.get(
        MAIL_LIST_URL,
        params={
            "sid": sid,
            "folderid": "1",
            "page": str(page),
            "s": "getmail",
            "today_tips": "400",
            "loc": "frame_html,,,22",
        },
    )
    response.raise_for_status()
    return response.text


def parse_total_pages(list_html):
    match = re.search(r'(\d+)\s*/\s*(\d+)\s*页', list_html)
    if not match:
        raise RuntimeError("未找到分页信息")
    return int(match.group(2))


def parse_mail_entries(list_html):
    pattern = re.compile(r'totime="(\d+)"[^>]*name="mailid"\s+value="([^"]+)"', re.S)
    return [(mailid, int(totime)) for totime, mailid in pattern.findall(list_html)]


def fetch_mail_detail(session, sid, mailid):
    response = session.get(READ_MAIL_URL, params={"sid": sid, "mailid": mailid})
    response.raise_for_status()
    return response.text


def parse_mail_detail(detail_html, mailid, mail_timestamp_ms):
    subject = normalize_text(extract_first(r'<span id="subject"[^>]*>(.*?)</span>', detail_html, default="(无主题)"))

    sender_name = normalize_text(extract_first(r'from:\s*\{\s*name\s*:\s*"([^"]*)"', detail_html, default=""))
    if not sender_name:
        sender_name = normalize_text(extract_first(r'mailnick="([^"]*)"', detail_html, default=""))
    if not sender_name:
        sender_name = normalize_text(extract_first(r'fromaddr="[^"]*"\s+title="([^"]*)"', detail_html, default=""))
    if not sender_name:
        sender_name = normalize_text(extract_first(r'replyTo:\s*\[\s*\{\s*name\s*:\s*"([^"]*)"', detail_html, default=""))

    sender_addr = normalize_text(extract_first(r'from:\s*\{.*?addr\s*:\s*"([^"]*)"', detail_html, default=""))
    if not sender_addr:
        sender_addr = normalize_text(extract_first(r'\bfrom="([^"]+@[^"]+)"', detail_html, default=""))
    if not sender_addr:
        sender_addr = normalize_text(extract_first(r'\bfromaddr="([^"]+@[^"]+)"', detail_html, default=""))
    if not sender_addr:
        sender_addr = normalize_text(extract_first(r'\baddr="([^"]+@[^"]+)"', detail_html, default=""))
    if not sender_addr:
        sender_addr = normalize_text(extract_first(r'mimefrom="([^"]+@[^"]+)"', detail_html, default=""))
    if not sender_addr:
        sender_addr = normalize_text(extract_first(r'replyTo:\s*\[\s*\{.*?addr\s*:\s*"([^"]+@[^"]+)"', detail_html, default=""))

    date_text = normalize_text(extract_first(r'date:\s*"([^"]*)"', detail_html, default=""))
    content_html = extract_first(r'<div id="mailContentContainer"[^>]*>(.*?)</div></div></div>', detail_html, default="")
    if not content_html:
        content_html = extract_first(r'<div id="mailContentContainer"[^>]*>(.*?)</div>', detail_html, default="")
    extractor = SimpleHTMLTextExtractor()
    extractor.feed(content_html)
    body_text = extractor.get_text()
    if not body_text:
        body_text = f"[邮件 {mailid} 未提取到正文]"
    return {
        "mailid": mailid,
        "subject": subject,
        "sender_name": sender_name,
        "sender_addr": sender_addr,
        "date": date_text,
        "timestamp_ms": mail_timestamp_ms,
        "body": body_text,
        "attachments": [],
    }


def export_to_docx(mails, output_path):
    document = Document()
    document.add_heading("收件箱邮件导出", level=0)
    document.add_paragraph(f"导出时间：{time.strftime('%Y-%m-%d %H:%M:%S')}")
    document.add_paragraph(f"筛选日期：{EXPORT_START_DATE} 到 {EXPORT_END_DATE}")
    document.add_paragraph(f"邮件总数：{len(mails)}")
    for index, mail in enumerate(mails, start=1):
        document.add_page_break()
        document.add_heading(f"{index}. {mail['subject']}", level=1)
        document.add_paragraph(f"发件人：{mail['sender_name']} <{mail['sender_addr']}>")
        document.add_paragraph(f"时间：{mail['date']}")
        document.add_paragraph(f"MailID：{mail['mailid']}")
        document.add_paragraph("正文：")
        document.add_paragraph(mail["body"])
        document.add_paragraph("附件：")
        if mail.get("attachments"):
            for attachment in mail["attachments"]:
                document.add_paragraph(
                    f"- {attachment['filename']}（{format_size(attachment['size'])}）\n  保存路径：{attachment['path']}"
                )
        else:
            document.add_paragraph("- 无")
    document.save(output_path)


def main():
    start_ms, end_ms = parse_export_range()
    session, sid = login_session()
    os.makedirs(ATTACHMENTS_DIR, exist_ok=True)
    first_page_html = fetch_inbox_page(session, sid, 0)
    total_pages = parse_total_pages(first_page_html)
    selected_entries = []
    seen = set()
    for page in range(total_pages):
        html_text = first_page_html if page == 0 else fetch_inbox_page(session, sid, page)
        entries = parse_mail_entries(html_text)
        for mailid, totime in entries:
            if start_ms <= totime <= end_ms and mailid not in seen:
                seen.add(mailid)
                selected_entries.append((mailid, totime))
        safe_print(f"已扫描列表页 {page + 1}/{total_pages}，符合日期范围的邮件 {len(selected_entries)} 封")

    mails = []
    for index, (mailid, totime) in enumerate(selected_entries, start=1):
        safe_print(f"[{index}/{len(selected_entries)}] 抓取邮件 {mailid}")
        detail_html = fetch_mail_detail(session, sid, mailid)
        mail = parse_mail_detail(detail_html, mailid, totime)
        try:
            attachments = download_attachments(session, sid, mailid, detail_html, ATTACHMENTS_DIR, mail["sender_addr"])
        except Exception as e:
            safe_print(f"  附件下载失败 {mailid}: {e}")
            attachments = []
        mail["attachments"] = attachments
        mails.append(mail)
        safe_print(f"  已抓取主题：{mail['subject']}，附件 {len(attachments)} 个")

    mails.sort(key=lambda item: item["timestamp_ms"])
    export_to_docx(mails, OUTPUT_DOCX_PATH)
    safe_print(f"导出完成：{OUTPUT_DOCX_PATH}，共 {len(mails)} 封邮件；附件目录：{ATTACHMENTS_DIR}")


if __name__ == "__main__":
    main()
