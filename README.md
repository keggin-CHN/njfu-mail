# 南京林业大学校园邮箱自动化工具集 (njfu-mail)

针对南京林业大学校园邮箱系统（基于腾讯企业邮箱 Tencent Exmail 深度定制架构）的 Python 自动化工具库与协议接口全集。支持自动化认证登录、邮件检索、作业批量收集、正文贴图抓取、PDF 无损渲染以及邮件快捷回复。

---

## 核心特性

- **完整 RSA 逆向鉴权**：内置 PKCS#1 v1.5 与动态时间戳加密，无需人工干预登录。
- **全格式作业收集**：
  - 支持标准文件附件（图片、压缩包、文档）。
  - **支持正文内嵌贴图抓取**：解决手机端 QQ 邮箱直接贴图导致附件数显示为 0 的行业漏收痛点。
  - **PDF 自动高清转图**：利用 PyMuPDF 将提交的 PDF 作业直接转换为高分辨率 JPG，方便统一阅卷。
- **反爬与安全规避**：
  - 基于 `curl_cffi` 完整模拟 Chrome TLS (JA3) 握手特征，绕过腾讯网关的静默重置。
  - 支持显式旁路本地系统代理（避免 Clash/v2rayN 节点干扰校园网）。
  - 请求频率平滑控制与重试退避机制。
- **邮件发送与回复**：支持动态提取 CSRF Token 并一键向学生发送答复或提醒邮件。

---

## 目录结构

```text
njfu-mail/
├── 接口文档/                             # 详尽的逆向接口协议文档
│   ├── README.md                        # 文档导航目录
│   ├── 01_认证与登录接口.md              # RSA加密、nTimeStamp、POST登录获取sid
│   ├── 02_邮件列表与检索接口.md          # 分页控制、时间戳检索与早停策略
│   ├── 03_邮件详情与正文查看接口.md      # 元数据提取与正文HTML清洗
│   ├── 04_附件与正文内嵌图片下载接口.md  # 附件/贴图下载、防乱码、PDF栅格化
│   ├── 05_网络环境与反爬风控避坑指南.md  # TLS指纹拦截、本地代理冲突与频控
│   ├── 06_开箱即用Python封装类(SDK).md   # 面向对象 SDK 架构与使用示例
│   └── 07_邮件发送与快捷回复接口.md      # 邮件回复/发送、动态 CSRF Token
├── njfu_client.py                       # 开箱即用的面向对象核心 SDK 模块
├── test.py                              # 批量邮件/作业导出脚本示例
├── .env.example                         # 环境变量配置示例
├── .gitignore                           # 严格的隐私与学生作业防泄漏过滤配置
└── README.md                            # 项目主说明文档
```

---

## 安装依赖

运行前请安装必需的第三方依赖包：

```bash
pip install curl_cffi pycryptodome pymupdf python-docx
```

---

## 快速配置

复制配置模板为 `.env`：

```bash
cp .env.example .env
```

配置您的账号信息：
```env
NJFU_USERNAME=2410403132
NJFU_PASSWORD=your_password
NJFU_DOMAIN=njfu.edu.cn
```

---

## 快速使用

### 1. 使用核心客户端 SDK (`njfu_client.py`)

```python
import os
from njfu_client import NJFUMailClient

# 实例化客户端（已内置直连防代理拦截与 Chrome TLS 伪装）
client = NJFUMailClient(
    username=os.getenv("NJFU_USERNAME"),
    password=os.getenv("NJFU_PASSWORD")
)

# 登录获取会话令牌
sid = client.login()
print(f"登录成功！SID: {sid}")

# 获取收件箱最新邮件
entries = client.get_mail_entries(folder_id=1, max_pages=1)
for mailid, totime in entries[:5]:
    mail = client.get_mail_detail(mailid)
    print(f"[{mail['date']}] {mail['sender_name']}: {mail['subject']}")
```

### 2. 回复邮件通知学生

```python
client.reply_mail(
    mailid="TARGET_MAIL_ID",
    to_addr="student@qq.com",
    subject="回复：作业提交提醒",
    html_body="<div>同学你好：作业已收到！</div>"
)
```

---

## 协议与文档

详细的抓包分析、请求参数表、响应字段解析与安全风控说明，请查阅 [**`接口文档/`**](./接口文档/README.md)。

---

## 免责声明与安全规范

本项目仅供学术研究、教学管理与自动化收作业等正当用途使用。请妥善保管个人统一身份认证账号与密码，严禁将包含真实学号、密码或同学作业个人信息的配置文件提交至公共代码仓库。
