"""
南京林业大学邮箱自动化处理模块 (NJFUMailClient)
基于 curl_cffi 与 PyCryptodome，支持作业自动识别、正文贴图抓取、PDF栅格化
"""

import base64
import html
import os
import re
import sys
import time
from datetime import datetime
from email.header import decode_header
from html.parser import HTMLParser
from urllib.parse import parse_qs, unquote, unquote_to_bytes, urlparse

from curl_cffi import requests
from Crypto.Cipher import PKCS1_v1_5
from Crypto.PublicKey import RSA
import pymupdf


class NJFUMailClient:
    LOGIN_URL = "https://mail.njfu.edu.cn/"
    POST_LOGIN_URL = "https://mail.njfu.edu.cn/cgi-bin/login"
    MAIL_LIST_URL = "https://mail.njfu.edu.cn/cgi-bin/mail_list"
    READ_MAIL_URL = "https://mail.njfu.edu.cn/cgi-bin/readmail"
    DOWNLOAD_URL = "https://mail.njfu.edu.cn/cgi-bin/download"
    VIEWFILE_URL = "https://mail.njfu.edu.cn/cgi-bin/viewfile"

    RSA_N_HEX = "CF87D7B4C864F4842F1D337491A48FFF54B73A17300E8E42FA365420393AC0346AE55D8AFAD975DFA175FAF0106CBA81AF1DDE4ACEC284DAC6ED9A0D8FEB1CC070733C58213EFFED46529C54CEA06D774E3CC7E073346AEBD6C66FC973F299EB74738E400B22B1E7CDC54E71AED059D228DFEB5B29C530FF341502AE56DDCFE9"
    RSA_E_HEX = "10001"

    def __init__(self, username, password, domain="njfu.edu.cn"):
        self.username = username
        self.password = password
        self.domain = domain
        self.sid = ""
        # 强制直连，模拟 Chrome 116 TLS
        self.session = requests.Session(
            impersonate="chrome116",
            proxies={"http": None, "https": None}
        )

    def login(self):
        """登录并获取全局 sid"""
        resp = self.session.get(self.LOGIN_URL)
        resp.raise_for_status()
        ts_match = re.search(r'nTimeStamp\s*:\s*"(\d+)"', resp.text)
        if not ts_match:
            raise RuntimeError("未能从首页获取到 nTimeStamp")
        timestamp = ts_match.group(1)

        # RSA PKCS#1 v1.5 加密
        message = f"{self.password}\n{timestamp}\n".encode("utf-8")
        pub_key = RSA.construct((int(self.RSA_N_HEX, 16), int(self.RSA_E_HEX, 16)))
        cipher = PKCS1_v1_5.new(pub_key)
        p_b64 = base64.b64encode(cipher.encrypt(message)).decode("utf-8")

        payload = {
            "uin": self.username,
            "domain": self.domain,
            "pp": p_b64,
            "verifycode": "",
            "f": "biz",
            "aliastype": "other",
            "p": p_b64,
            "ppp": "",
            "fun": "",
            "errtemplate": "new_index/index",
        }
        post_resp = self.session.post(self.POST_LOGIN_URL, data=payload, allow_redirects=False)
        post_resp.raise_for_status()

        sid_match = re.search(r'targetUrl\s*=\s*urlHead\s*\+\s*"frame_html\?sid=([^"&]+)', post_resp.text)
        if not sid_match:
            raise RuntimeError("登录失败，未找到 sid。请检查账号密码。")
        self.sid = sid_match.group(1)
        return self.sid

    def get_mail_entries(self, folder_id=1, max_pages=3):
        """获取邮件条目列表 [(mailid, totime_ms), ...]"""
        entries = []
        seen = set()
        for page in range(max_pages):
            time.sleep(0.4)
            resp = self.session.get(
                self.MAIL_LIST_URL,
                params={
                    "sid": self.sid,
                    "folderid": str(folder_id),
                    "page": str(page),
                    "s": "getmail",
                }
            )
            resp.raise_for_status()
            matches = re.findall(r'totime="(\d+)"[^>]*name="mailid"\s+value="([^"]+)"', resp.text)
            if not matches:
                break
            for totime, mailid in matches:
                if mailid not in seen:
                    seen.add(mailid)
                    entries.append((mailid, int(totime)))
        return entries

    def get_mail_detail(self, mailid):
        """获取邮件详细信息与正文"""
        time.sleep(0.3)
        resp = self.session.get(self.READ_MAIL_URL, params={"sid": self.sid, "mailid": mailid})
        resp.raise_for_status()
        html_text = resp.text

        # 主题
        subject = self._extract(r'<span id="subject"[^>]*>(.*?)</span>', html_text, "(无主题)")
        # 发件人
        sender_name = self._extract(r'from:\s*\{\s*name\s*:\s*"([^"]*)"', html_text)
        sender_addr = self._extract(r'from:\s*\{.*?addr\s*:\s*"([^"]*)"', html_text)
        # 时间
        date_str = self._extract(r'date:\s*"([^"]*)"', html_text)

        # 正文容器
        content_m = re.search(r'<div id="mailContentContainer"[^>]*>(.*?)</div></div></div>', html_text, re.S)
        if not content_m:
            content_m = re.search(r'<div id="mailContentContainer"[^>]*>(.*?)</div>', html_text, re.S)
        content_html = content_m.group(1) if content_m else ""

        # 普通附件
        attachments = []
        att_matches = re.findall(r'<a[^>]+href="([^"]*?/cgi-bin/download[^"#]*)"[^>]*>(.*?)</a>', html_text)
        for href, label in att_matches:
            full_url = href if href.startswith("http") else f"https://mail.njfu.edu.cn{href}"
            parsed = urlparse(full_url)
            query = parse_qs(parsed.query)
            fn = query.get("filename", [""])[0] or label
            attachments.append({
                "url": full_url,
                "filename": self._clean_filename(fn),
            })

        # 内联图片
        inline_imgs = []
        for img_src in re.findall(r'<img[^>]+src=["\']([^"\']+)["\']', content_html):
            if "viewfile" in img_src or "readmail" in img_src:
                full_src = img_src if img_src.startswith("http") else f"https://mail.njfu.edu.cn{img_src}"
                inline_imgs.append(full_src)

        return {
            "mailid": mailid,
            "subject": html.unescape(subject).strip(),
            "sender_name": html.unescape(sender_name).strip(),
            "sender_addr": sender_addr.strip(),
            "date": date_str.strip(),
            "attachments": attachments,
            "inline_images": inline_imgs,
            "content_html": content_html,
        }

    def download_resource(self, url):
        """下载附件或正文图片二进制数据"""
        if "sid=" not in url:
            sep = "&" if "?" in url else "?"
            url = f"{url}{sep}sid={self.sid}"
        resp = self.session.get(url)
        resp.raise_for_status()
        return resp.content

    def reply_mail(self, mailid, to_addr, subject, html_body):
        """回复指定邮件 (带 CSRF Token 解析)"""
        template_url = f"https://mail.njfu.edu.cn/cgi-bin/readtemplate?action=reply&mailid={mailid}&t=compose&sid={self.sid}"
        resp = self.session.get(template_url)
        resp.raise_for_status()

        form_m = re.search(r'<form[^>]+name="frm"[^>]*>(.*?)</form>', resp.text, re.S)
        if not form_m:
            raise RuntimeError("未在模板页面中解析到 frm 表单")

        payload = {}
        for inp in re.finditer(r'<input[^>]+>', form_m.group(1)):
            nm = re.search(r'name=["\']([^"\']+)["\']', inp.group(0))
            vm = re.search(r'value=["\']([^"\']*)["\']', inp.group(0))
            tm = re.search(r'type=["\']([^"\']+)["\']', inp.group(0))
            if not nm:
                continue
            t = tm.group(1).lower() if tm else "text"
            if t in ("button", "submit", "reset"):
                continue
            payload[nm.group(1)] = vm.group(1) if vm else ""

        payload["to"] = to_addr
        payload["subject"] = subject
        payload["content__html"] = html_body
        payload["actiontype"] = "send"
        payload["savesendbox"] = "1"
        payload["sid"] = self.sid

        send_url = f"https://mail.njfu.edu.cn/cgi-bin/compose_send?sid={self.sid}"
        send_resp = self.session.post(send_url, data=payload)
        send_resp.raise_for_status()
        return "成功" in send_resp.text or "mail_sent" in send_resp.text

    def export_pdf_as_images(self, pdf_bytes, output_jpg_path, dpi=200):
        """将 PDF 渲染为高清 JPG 图片"""
        doc = pymupdf.open(stream=pdf_bytes, filetype="pdf")
        for i in range(len(doc)):
            pix = doc[i].get_pixmap(dpi=dpi)
            path = output_jpg_path if len(doc) == 1 else f"{os.path.splitext(output_jpg_path)[0]}_{i+1}.jpg"
            pix.save(path)

    @staticmethod
    def _extract(pattern, text, default=""):
        m = re.search(pattern, text, re.S)
        return m.group(1).strip() if m else default

    @staticmethod
    def _clean_filename(name):
        name = html.unescape(name).replace("+", " ")
        name = unquote(name)
        return re.sub(r'[\\/:*?"<>|]', "_", name).strip()

