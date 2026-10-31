#!/usr/bin/env python3
"""ipa.go.jp の SCS評価制度のページと文書を取得し、最新版と過去の版を分けて保存する。

このスクリプトは非公開リポジトリ（SCS-AP-Document-Archiver）に置き、
取得結果は環境変数 ARCHIVE_DIR で指定したフォルダ（公開リポジトリ SCS-AP-Document のチェックアウト先）に書き込む。
ARCHIVE_DIR を指定しないときは、このリポジトリの site/ に書き込む（手元での試験用。site/ は git の管理外）。

保存先のフォルダ構成（掲載ページを単位に分類する。ホスト名や rcu1hd…-att のような不透明な ID は使わない）
  README.md                                SCS評価制度の概要
  documents/<分類>/<ページ名>.md            最新版のページ本文
  documents/<分類>/<ページ名>/<名前>        最新版の原本（PDF・Word・Excel など）
  documents/<分類>/<ページ名>/<名前>.md     最新版のテキスト版
  archive/<分類>/<検知日>_<名前>            過去の版の原本
  archive/<分類>/<検知日>_<ページ名>.md     過去の版のページ本文
  history/CHANGELOG.md                     更新履歴（差分つき）
  history/documents.md                     取得対象のページ・文書と過去の版の一覧
  history/diffs/<版>/<名前>.diff            前の版との差分
  history/catalog.json                     全ページ・全文書・全版の台帳

  分類は掲載ページの URL から作る。例：https://www.ipa.go.jp/security/scs/requirements-criteria.html
    -> ページ本文 documents/scs/requirements-criteria.md
       掲載資料   documents/scs/requirements-criteria/<名前>

取得対象
  - トップページ（SCOPES の start）と、その配下（prefix）のページ
  - トップページの本文からリンクされている、ALLOWED_HOSTS 上のページ（リンクが増えれば自動で追加）
  - 上記のページの本文からリンクされている文書（PDF・Word・Excel・PowerPoint・CSV・ZIP）
"""
import datetime as dt
import difflib
import email.utils
import hashlib
import json
import os
import pathlib
import re
import shutil
import subprocess
import sys
import tempfile
import time
import unicodedata
import zipfile
from urllib import robotparser
from urllib.parse import quote, urljoin, urlparse, urlunparse

import openpyxl
import requests
from bs4 import BeautifulSoup
from markdownify import markdownify
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

# ---------- 設定 ----------
HOST = "www.ipa.go.jp"
ALLOWED_HOSTS = {HOST}   # 取得してよいホスト。他のサイト（例: www.meti.go.jp）も取得するならここに追加
SCOPES = [  # 取得範囲。start がトップページ、prefix 配下のページはすべてたどる
    {"name": "SCS評価制度", "start": f"https://{HOST}/security/scs/index.html", "prefix": "/security/scs/"},
]
DOC_EXT = (".pdf", ".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".ppt", ".pptx", ".csv", ".zip")
MAIN_SELECTORS = (                   # 本文領域の候補（上から順に試す）。サイトのデザインが変わったらここに追加
    "main.container--col2__main",    # IPA の通常のページ
    "main",                          # 調達情報など、レイアウトの違うページ
    "#mainContent",
    "article",
)
MAX_BYTES = 90 * 1024 * 1024        # GitHub の 100MB 制限の手前で止める
# 連絡先は GitHub の Settings → Secrets and variables → Actions → Variables の CONTACT で設定
UA = f"ipa-doc-archiver/1.0 (contact: {os.environ.get('CONTACT') or 'you@example.com'})"
WAIT = 1.5                           # リクエスト間隔（秒）

SCRIPT_DIR = pathlib.Path(__file__).resolve().parent
TOOL_ROOT = SCRIPT_DIR.parent                    # 非公開リポジトリのルート（.commit_msg などの作業ファイルを置く）
ROOT = pathlib.Path(os.environ.get("ARCHIVE_DIR") or TOOL_ROOT / "site").resolve()   # 取得結果の保存先（公開リポジトリ）
DOCS = ROOT / "documents"            # 最新版
ARCHIVE = ROOT / "archive"           # 過去の版
HISTORY = ROOT / "history"           # 更新履歴
DIFFS = HISTORY / "diffs"
CATALOG = HISTORY / "catalog.json"
JST = dt.timezone(dt.timedelta(hours=9))
TODAY = dt.datetime.now(JST).strftime("%Y-%m-%d")

CHANGELOG_DIFF_EVENTS = 30          # CHANGELOG に差分を埋め込む件数（新しい順）。古いものは差分ファイルへのリンクのみ
CHANGELOG_DIFF_LINES = 150          # CHANGELOG に埋め込む差分の行数の上限（1件あたり）
MSG_DIFF_LINES = 200                # コミットメッセージに載せる差分の行数の上限（1件あたり）
MSG_MAX_BYTES = 60_000              # コミットメッセージ全体の上限

ERRORS = []                          # 取得に失敗したもの（最後に .errors に書き出し、Actions で通知する）


def warn(msg):
    print(f"::warning::{msg}")
    ERRORS.append(msg)


# ---------- HTTP ----------
session = requests.Session()
session.headers["User-Agent"] = UA
_retry = Retry(total=3, backoff_factor=2, status_forcelist=(429, 500, 502, 503, 504),
               allowed_methods=frozenset(["GET", "HEAD"]), respect_retry_after_header=True)
session.mount("https://", HTTPAdapter(max_retries=_retry))
session.mount("http://", HTTPAdapter(max_retries=_retry))
_robots = {}


class Gone(Exception):
    """404 / 410：ページや文書がサイトから削除された"""


def robots_for(host):
    if host not in _robots:
        rp = robotparser.RobotFileParser()
        r = session.get(f"https://{host}/robots.txt", timeout=30)
        if r.status_code in (401, 403):
            rp.disallow_all = True
        elif 400 <= r.status_code < 500:
            rp.allow_all = True
        else:
            r.raise_for_status()
            rp.parse(r.text.splitlines())
        _robots[host] = rp
    return _robots[host]


def get(url, headers=None):
    if not robots_for(urlparse(url).netloc).can_fetch(UA, url):
        raise PermissionError(f"robots.txt で禁止されています: {url}")
    time.sleep(WAIT)
    r = session.get(url, headers=headers or {}, timeout=60)
    if r.status_code in (404, 410):
        raise Gone(f"{r.status_code}: {url}")
    if r.status_code != 304:
        r.raise_for_status()
    return r


def probe(url):
    """取得対象から外れたページがまだ存在するかを確かめる"""
    try:
        get(url)
        return "alive"
    except Gone:
        return "gone"
    except Exception:
        return "error"


# ---------- ユーティリティ ----------
def normalize(url):
    p = urlparse(url)
    path = p.path + "index.html" if p.path.endswith("/") else p.path
    return urlunparse(("https", p.netloc, path, "", "", ""))

def is_doc_path(path):
    return path.lower().endswith(DOC_EXT)

def is_page_path(path):
    last = path.rsplit("/", 1)[-1].lower()
    return last.endswith((".html", ".htm")) or "." not in last

def url_path(url):
    p = urlparse(url)
    return p.netloc + p.path

def leaf(url):
    return urlparse(url).path.rsplit("/", 1)[-1] or "index.html"

SECURITY_PREFIX = "security"          # IPA の /security/ はどのページにも共通なので分類から外す
OPAQUE_DIR = re.compile(r"^r[a-z0-9]{6,}-att$")          # 例: rcu1hd0000007a2i-att（IPAの添付ファイル置き場）

def page_parts(url):
    """掲載ページの URL から保存先の分類とページ名を作る。
    例: https://www.ipa.go.jp/security/scs/requirements-criteria.html -> ["scs", "requirements-criteria"]
        -> ページ本文 documents/scs/requirements-criteria.md、掲載資料 documents/scs/requirements-criteria/"""
    parts = [p for p in urlparse(url).path.split("/") if p]
    if parts and parts[0] == SECURITY_PREFIX:
        parts = parts[1:]
    if parts:
        parts[-1] = re.sub(r"\.html?$", "", parts[-1])
    return [p for p in parts if p and not OPAQUE_DIR.match(p)] or ["other"]

def page_folder(url):
    """(分類のフォルダ, ページ名)"""
    parts = page_parts(url)
    return parts[:-1], parts[-1]

def folder_path(base, parts):
    return base.joinpath(*parts) if parts else base

def rel(path):
    return pathlib.Path(path).relative_to(ROOT).as_posix()

def link_from(path, base):
    """base フォルダにある Markdown から path へのリンク"""
    return quote(os.path.relpath(ROOT / path, base).replace(os.sep, "/"))

def text_of(el):
    return el.get_text(" ", strip=True) if el else ""

def tidy(md):
    return re.sub(r"\n{3,}", "\n\n", re.sub(r"[ \t\u3000]+\n", "\n", md)).strip()

def clean_title(t):  # 「…(PDF:165 KB)」のサイズ表記を除く
    return re.sub(r"\s*[（(][^()（）]*?\d[\d.,]*\s*[KMG]?B[)）]\s*$", "", t, flags=re.I).strip()

def esc(s):
    return (s or "").replace("|", "\\|")

def norm(t):
    """比較用の正規化。全角・半角の違いや空白の違いを吸収する（例：「★3・★4　要求事項」と「★3・★4 要求事項」）"""
    return re.sub(r"\s+", " ", unicodedata.normalize("NFKC", t or "")).strip()

def same_doc(a, b):
    """ファイル名が変わっても同じ文書かを判定する。
    両方に文書番号（SCS-400 など）があれば番号で、なければ正規化したタイトルで比べる"""
    if a.get("doc_no") and b.get("doc_no"):
        return a["doc_no"] == b["doc_no"]
    return norm(a.get("title")) == norm(b.get("title"))

def latest(entry):
    vs = entry.get("versions") or []
    return vs[-1] if vs else None

VER_IN_NAME = re.compile(r"-v(\d+(?:\.\d+)*)(?=[-_.]|$)", re.I)

def rfc1123_date(value):
    """Last-Modified（例: Fri, 17 Jul 2026 04:55:12 GMT）を日本時間の yyyy-mm-dd にする"""
    if not value:
        return None
    try:
        return email.utils.parsedate_to_datetime(value).astimezone(JST).strftime("%Y-%m-%d")
    except (TypeError, ValueError):
        return None

def jp_date_to_iso(value):
    m = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", value or "")
    return f"{m.group(1)}-{int(m.group(2)):02d}-{int(m.group(3)):02d}" if m else None

def file_version(v, entry=None):
    """資料のバージョン表記。
    ファイル名に vX.X（例: SCS-410-self-assess-guide-v1.0）があればそれを、なければファイルの最終更新日
    （yyyy-mm-dd）を返す。更新日時が取れない資料だけ、公式の更新日・初回取得日を使う。"""
    name = leaf(v.get("url") or v.get("original") or "")
    m = VER_IN_NAME.search(pathlib.PurePosixPath(name).stem)
    if m:
        return f"v{m.group(1)}"
    return (rfc1123_date(v.get("last_modified"))
            or jp_date_to_iso((entry or {}).get("official_updated"))
            or v.get("date") or "")

def latest_version_label(entry):
    v = latest(entry)
    return file_version(v, entry) if v else ""


# ---------- 保存先 ----------
def current_paths(url, is_doc, page=None):
    """最新版の保存先 (原本, テキスト版)。資料は掲載ページのフォルダに置く"""
    parts, name = page_folder(page or url)
    if is_doc:
        folder = folder_path(DOCS, parts + [name])
        fname = leaf(url)
        return folder / fname, folder / f"{fname}.md"
    return None, folder_path(DOCS, parts) / f"{name}.md"


def archive_paths(url, is_doc, version, page=None):
    """過去の版の保存先 (原本, テキスト版)。版の名前は検知日（例: 2026-10-28）"""
    parts, name = page_folder(page or url)
    prefix = f"{version}_"
    if is_doc:
        folder = folder_path(ARCHIVE, parts + [name])
        fname = leaf(url)
        return folder / f"{prefix}{fname}", folder / f"{prefix}{fname}.md"
    return None, folder_path(ARCHIVE, parts) / f"{prefix}{name}.md"


# ---------- ページ ----------
DOC_NO = re.compile(r"^SCS-[1-9][0-9]{2}[A-Z]*$")          # 文書番号（SCS-0911 のような日付は除く）
DOC_NO_IN_FILE = re.compile(r"(?:^|/)(SCS-[1-9][0-9]{2}[A-Z]*)(?=[-_.])")

def soup_new(el, name):
    root = el
    while root.parent is not None:
        root = root.parent
    return root.new_tag(name)

def find_main(soup):
    for sel in MAIN_SELECTORS:
        el = soup.select_one(sel)
        if el is not None and len(text_of(el)) > 20:
            return el, sel
    return None, None

def clean_main(main):
    """本文を差分の読みやすい Markdown になるよう整える"""
    # ナビゲーション・目次のジャンプボタン・アコーディオンの「閉じる」など、本文でない要素を除く
    for el in main.select("nav, aside, header, footer, script, style, noscript, button, "
                          ".breadcrumb, .anchorlink-list, .anchor"):
        el.decompose()
    for h in main.find_all("h2"):
        nxt = h.find_next_sibling(True)
        if text_of(h) == "目次" and (nxt is None or nxt.name == "h2"):   # 中身を除いて空になった目次
            h.decompose()
    # よくある質問（acc-box）：「### Q1-1 質問」「**A1-1** 回答」の形にする
    for box in main.select("ol.acc-box, ul.acc-box"):
        wrap = soup_new(box, "div")
        for item in box.select(".acc-box__item"):
            q = f"{text_of(item.select_one('.acc-box__trigger__label'))} {text_of(item.select_one('.acc-box__trigger__txt'))}".strip()
            h = soup_new(box, "h3"); h.string = q or text_of(item.select_one(".acc-box__trigger")); wrap.append(h)
            label = text_of(item.select_one(".acc-box__child__label"))
            if label:
                p_ = soup_new(box, "p"); b = soup_new(box, "strong"); b.string = label; p_.append(b); wrap.append(p_)
            body = item.select_one(".acc-box__child__txt-wrap") or item.select_one(".acc-box__child")
            if body:
                for child in list(body.children):
                    wrap.append(child.extract())
        box.replace_with(wrap)
    # 定義リスト（文書名・概要の表など）：「**見出し**」の段落 + 中身にする（GitHub は定義リストを表示できないため）
    for dl in main.select("dl"):
        wrap = soup_new(dl, "div")
        for dt_ in dl.find_all("dt"):
            p_ = soup_new(dl, "p"); b = soup_new(dl, "strong"); b.string = text_of(dt_); p_.append(b); wrap.append(p_)
            dd = dt_.find_next_sibling("dd")
            if dd:
                for child in list(dd.children):
                    wrap.append(child.extract())
        dl.replace_with(wrap)
    return main

def doc_context(a, url):
    """文書リンクの周辺から、見出し・文書番号・公式の更新日を取り出す"""
    h = a.find_previous(["h1", "h2", "h3"])
    section = text_of(h) if h is not None and h.name != "h1" else ""
    doc_no = updated = None
    dd = a.find_parent("dd")
    dt_ = dd.find_previous_sibling("dt") if dd else None
    if dd and dt_ and text_of(dt_) == "文書名":           # 「文書名」欄にあるリンクだけ文書番号を採る
        texts = [text_of(x) for x in dd.select(".list__item__txt, p")]
        doc_no = next((t for t in texts if DOC_NO.match(t)), None)
        m = re.search(r"更新日[：:]\s*([0-9０-９]+年[0-9０-９]+月[0-9０-９]+日)", text_of(dd))
        updated = m.group(1) if m else None
    if not doc_no:                                        # 別紙などはファイル名から（例：SCS-410AP-...）
        m = DOC_NO_IN_FILE.search(urlparse(url).path)
        doc_no = m.group(1) if m else None
    return {"section": section, "doc_no": doc_no, "official_updated": updated}

def parse_page(url, html):
    soup = BeautifulSoup(html, "html.parser")
    main, selector = find_main(soup)
    if main is None:
        raise RuntimeError(f"本文領域が見つかりません（ページ構造が変わった可能性。MAIN_SELECTORS を確認）: {url}")
    for tag, attr in (("a", "href"), ("img", "src")):
        for el in soup.select(f"{tag}[{attr}]"):
            el[attr] = urljoin(url, el[attr].strip())

    def links_in(el):
        out = []
        for a in el.select("a[href]"):
            if a["href"].startswith("http"):
                out.append((normalize(a["href"]), re.sub(r"\s+", " ", a.get_text(" ", strip=True))))
        return out

    page_links = [h for h, _ in links_in(soup)]          # ページ全体（サイドメニュー含む）のリンク
    body_links = links_in(main)                           # 本文のリンク

    # 文書は本文のリンクだけを対象にし、見出し・文書番号・公式の更新日を付ける
    docs, seen = [], set()
    for a in main.select("a[href]"):
        if not a["href"].startswith("http"):
            continue
        href = normalize(a["href"])
        if is_doc_path(urlparse(href).path) and href not in seen:
            seen.add(href)
            title = clean_title(re.sub(r"\s+", " ", a.get_text(" ", strip=True))) or leaf(href)
            docs.append((href, {"title": title, **doc_context(a, href)}))
    # 「配付資料」「議事要旨」のように同じページで名前が重なるものは見出しを付けて区別する
    counts = {}
    for _, d in docs:
        counts[norm(d["title"])] = counts.get(norm(d["title"]), 0) + 1
    for _, d in docs:
        if counts[norm(d["title"])] > 1 and d["section"]:
            d["title"] = f"{d['section']} {d['title']}"

    info = {
        "title": text_of(main.find("h1")) or text_of(soup.find("title")) or url,
        "dates": [text_of(p) for p in main.select(".news-detail__update-info p")],
        "changelog": [{"date": text_of(li.select_one(".date")), "text": text_of(li.select_one(".txt"))}
                      for li in main.select(".change-log-box__list__item")],
        "selector": selector,
    }
    top = BeautifulSoup(str(main), "html.parser")         # README 用に加工前の本文を残す
    md = tidy(markdownify(str(clean_main(main)), heading_style="ATX", bullets="-"))
    info["sha256"] = hashlib.sha256(md.encode()).hexdigest()
    return info, page_links, body_links, docs, md, top


def crawl(scope):
    """取得対象のページをたどる。
    - prefix 配下のページ：本文とサイドメニューのリンクから再帰的にたどる
    - トップページの本文にあるリンク：ALLOWED_HOSTS 上のページなら prefix の外でも取得する（その先はたどらない）
    """
    start, prefix = scope["start"], scope["prefix"]
    queue, seen = [(start, "トップページ")], set()
    pages, docs, failed, gone, external, top = {}, {}, set(), set(), {}, None
    while queue:
        url, via = queue.pop(0)
        if url in seen:
            continue
        seen.add(url)
        try:
            r = get(url)
        except Gone:
            if url == start:
                sys.exit(f"トップページが見つかりません（404）: {url}")
            gone.add(url)
            continue
        except Exception as e:
            if url == start:
                sys.exit(f"トップページを取得できません: {url}: {e}")
            warn(f"ページの取得に失敗: {url}: {e}")
            failed.add(url)
            continue
        ctype = r.headers.get("Content-Type", "")
        if "html" not in ctype:
            continue                                      # 拡張子のないリンクが HTML 以外だった
        r.encoding = "utf-8"
        try:
            info, page_links, body_links, page_docs, md, main = parse_page(url, r.text)
        except Exception as e:
            if url == start:
                sys.exit(str(e))
            warn(str(e))
            failed.add(url)
            continue
        info.update(scope=scope["name"], order=len(pages), via=via, body=md)
        pages[url] = info
        if url == start:
            top = main

        for href, meta in page_docs:                      # 文書：許可したホスト上ならパスを問わず保存
            if urlparse(href).netloc in ALLOWED_HOSTS:
                d = docs.setdefault(href, {**meta, "pages": []})
                if url not in d["pages"]:
                    d["pages"].append(url)

        if urlparse(url).path.startswith(prefix):         # 取得範囲内のページからは範囲内のページをたどる
            for href in page_links:
                p = urlparse(href)
                if p.netloc == HOST and p.path.startswith(prefix) and is_page_path(p.path) \
                        and not is_doc_path(p.path) and href not in seen:
                    queue.append((href, via if url != start else "トップページ配下"))
        if url == start:                                  # トップページの本文のリンクは範囲外でも取得する
            for href, text in body_links:
                p = urlparse(href)
                if is_doc_path(p.path) or not is_page_path(p.path):
                    continue
                if p.netloc in ALLOWED_HOSTS:
                    if not p.path.startswith(prefix) and href not in seen:
                        queue.append((href, "トップページのリンク"))
                else:
                    external.setdefault(href, text)       # 許可していないホストは一覧にだけ載せる
    return pages, docs, top, failed, gone, external


# ---------- 文書のテキスト化 ----------
def run_text(cmd):
    return subprocess.run(cmd, check=True, capture_output=True, timeout=300).stdout.decode("utf-8", "replace")

def rows_to_md(title, rows):
    """表（行のリスト）を「1行 = 1ブロック」の Markdown にする。セル内改行が多い表でも差分が読みやすい"""
    out = [f"# シート: {title}\n"]
    rows = [["" if c is None else str(c).strip() for c in r] for r in rows]
    rows = [r for r in rows if any(r)]
    if not rows:
        return out
    width = max(sum(1 for c in r if c) for r in rows)
    hi = next(i for i, r in enumerate(rows) if sum(1 for c in r if c) >= max(1, width // 2))
    out += [" ".join(c for c in r if c) + "\n" for r in rows[:hi]]   # 表の上のタイトル・注記
    header = rows[hi]
    for n, r in enumerate(rows[hi + 1:], 1):
        out.append(f"## {n}")
        for j, c in enumerate(r):
            if c:
                h = header[j] if j < len(header) and header[j] else f"列{j + 1}"
                out.append(f"- **{h}**: " + c.replace("\n", "\n  "))
        out.append("")
    return out

def xlsx_to_md(path):
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    out = []
    for ws in wb.worksheets:
        out += rows_to_md(ws.title, ws.iter_rows(values_only=True))
    return "\n".join(out)

def xls_to_md(path):
    import xlrd
    wb = xlrd.open_workbook(path)
    out = []
    for ws in wb.sheets():
        out += rows_to_md(ws.name, (ws.row_values(i) for i in range(ws.nrows)))
    return "\n".join(out)

def docx_to_md(path):
    """本文の段落と表を文書の順番どおりに出す"""
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(path)
    out = []
    for block in d.element.body.iterchildren():
        tag = block.tag.rsplit("}", 1)[-1]
        if tag == "p":
            p = Paragraph(block, d)
            t = p.text.strip()
            if not t:
                continue
            style = (p.style.name or "") if p.style is not None else ""
            m = re.match(r"(?:Heading|見出し)\s*(\d)", style)
            out.append(f"{'#' * min(int(m.group(1)) + 1, 6)} {t}" if m else t)
            out.append("")
        elif tag == "tbl":
            rows = []
            for row in Table(block, d).rows:
                cells = []
                for c in row.cells:
                    txt = c.text.strip().replace("\n", "<br>").replace("|", "\\|")
                    if not cells or cells[-1] != txt:     # 結合セルの重複を除く
                        cells.append(txt)
                rows.append(cells)
            if rows:
                n = max(len(r) for r in rows)
                rows = [r + [""] * (n - len(r)) for r in rows]
                out += ["| " + " | ".join(rows[0]) + " |", "|" + "---|" * n]
                out += ["| " + " | ".join(r) + " |" for r in rows[1:]]
                out.append("")
    return "\n".join(out)

def pptx_to_md(path):
    from pptx import Presentation
    out = []
    for i, slide in enumerate(Presentation(path).slides, 1):
        out.append(f"## スライド {i}\n")
        for shape in slide.shapes:
            if shape.has_text_frame and shape.text_frame.text.strip():
                out.append(shape.text_frame.text.strip() + "\n")
            if getattr(shape, "has_table", False) and shape.has_table:
                for row in shape.table.rows:
                    out.append("| " + " | ".join(c.text.strip().replace("\n", " ") for c in row.cells) + " |")
                out.append("")
    return "\n".join(out)

def via_libreoffice(path, fmt):
    """古い形式は LibreOffice で新しい形式に変換してから読む（catdoc などがないときの予備）"""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("変換ツールがありません（catdoc / antiword / libreoffice）")
    with tempfile.TemporaryDirectory() as tmp:
        subprocess.run([soffice, "--headless", "--convert-to", fmt, "--outdir", tmp, str(path)],
                       check=True, capture_output=True, timeout=300)
        out = next(pathlib.Path(tmp).glob(f"*.{fmt}"))
        return {"docx": docx_to_md, "xlsx": xlsx_to_md, "pptx": pptx_to_md}[fmt](out)

def zip_to_md(path):
    """ZIP は中身の一覧（名前・サイズ・CRC）をテキスト版にする。中身が変われば差分に出る"""
    out = ["| ファイル | サイズ | CRC32 |", "|---|---|---|"]
    with zipfile.ZipFile(path) as z:
        for i in z.infolist():
            name = i.filename
            if not (i.flag_bits & 0x800):                # UTF-8 フラグがない日本語名は CP932 として読む
                try:
                    name = name.encode("cp437").decode("cp932")
                except (UnicodeEncodeError, UnicodeDecodeError):
                    pass
            out.append(f"| {esc(name)} | {i.file_size:,} | {i.CRC:08x} |")
    return "\n".join(out)

def to_markdown(src, url):
    """原本からテキスト版（Markdown）を作る。対応していない形式は None"""
    ext = src.suffix.lower()
    head = f"<!-- source: {url} -->\n\n"
    if ext == ".pdf":
        return head + "```text\n" + run_text(["pdftotext", "-layout", "-enc", "UTF-8", str(src), "-"]).rstrip() + "\n```\n"
    if ext in (".xlsx", ".xlsm"):
        return head + xlsx_to_md(src)
    if ext == ".xls":
        try:
            return head + xls_to_md(src)
        except Exception:
            return head + via_libreoffice(src, "xlsx")
    if ext == ".docx":
        return head + docx_to_md(src)
    if ext == ".doc":
        for cmd in (["antiword", "-w", "0", str(src)], ["catdoc", "-d", "utf-8", "-w", str(src)]):
            if shutil.which(cmd[0]):
                try:
                    text = run_text(cmd).strip()
                    if text:                              # 空なら次の方法を試す
                        return head + "```text\n" + text + "\n```\n"
                except Exception:
                    pass
        return head + via_libreoffice(src, "docx")
    if ext == ".pptx":
        return head + pptx_to_md(src)
    if ext == ".ppt":
        if shutil.which("catppt"):
            try:
                text = run_text(["catppt", "-d", "utf-8", str(src)]).strip()
                if text:                                  # 空なら LibreOffice で変換する
                    return head + "```text\n" + text + "\n```\n"
            except Exception:
                pass
        return head + via_libreoffice(src, "pptx")
    if ext == ".csv":
        raw = src.read_bytes()
        for enc in ("utf-8-sig", "cp932"):
            try:
                return head + "```text\n" + raw.decode(enc).rstrip() + "\n```\n"
            except UnicodeDecodeError:
                pass
    if ext == ".zip":
        return head + zip_to_md(src)
    return None


# ---------- 版の保存（最新版は documents/、過去の版は archive/） ----------
def read(path):
    p = ROOT / path if path else None
    return p.read_text("utf-8") if p and p.is_file() else ""

def make_diff(old_text, new_text, old_name, new_name):
    lines = difflib.unified_diff(old_text.splitlines(), new_text.splitlines(),
                                 fromfile=old_name, tofile=new_name, lineterm="", n=1)
    return "\n".join(lines)

def diff_stat(diff):
    add = sum(1 for l in diff.splitlines() if l.startswith("+") and not l.startswith("+++"))
    rem = sum(1 for l in diff.splitlines() if l.startswith("-") and not l.startswith("---"))
    return add, rem

def version_name(versions, url, is_doc, page=None):
    """版の名前は検知日。同じ日に2回目の変更があれば 2026-10-28-2 のように番号を付ける"""
    used = {v["version"] for v in versions}
    for p in archive_paths(url, is_doc, "x", page):
        if p is not None and p.parent.is_dir():           # 台帳にないファイルが残っていても上書きしない
            used |= {f.name.split("_", 1)[0] for f in p.parent.iterdir() if "_" in f.name}
    name, n = TODAY, 1
    while name in used:
        n += 1
        name = f"{TODAY}-{n}"
    return name

def move(src_rel, dst):
    """documents/ にあるファイルを archive/ に移し、新しいパスを返す"""
    if not src_rel or not (ROOT / src_rel).is_file():
        return None
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.move(str(ROOT / src_rel), str(dst))
    return rel(dst)

def archive_version(v, url, is_doc, page=None):
    """最新版だった v を archive/ に移す（版の名前はその版を検知した日）"""
    if v.get("archived"):
        return
    page = page or v.get("page") or url
    v["page"] = page
    orig_dst, md_dst = archive_paths(v.get("url", url), is_doc, v["version"], page)
    if v.get("original"):
        v["original"] = move(v["original"], orig_dst) or v["original"]
    if v.get("markdown"):
        v["markdown"] = move(v["markdown"], md_dst) or v["markdown"]
    v["archived"] = True

def save_diff(url, old_md_text, old_name, new_md_rel, ver):
    if not old_md_text:
        return None, 0, 0
    diff = make_diff(old_md_text, read(new_md_rel), old_name, new_md_rel)
    if not diff:
        return None, 0, 0
    out = DIFFS / ver / (url_path(url).replace("/", "__") + ".diff")
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(diff + "\n", "utf-8")
    add, rem = diff_stat(diff)
    return rel(out), add, rem

def add_version(entry, url, is_doc, info, md_text=None, data=None):
    """新しい版を documents/ に保存し、それまでの最新版を archive/ に移す"""
    versions = entry.setdefault("versions", [])
    page = entry.get("page") or url
    prev = latest(entry)
    old_md_text, old_name = (read(prev.get("markdown")), None) if prev else ("", None)
    if prev:
        archive_version(prev, url, is_doc, page)
        old_name = prev.get("markdown")
    ver = version_name(versions, url, is_doc, page)
    orig_path, md_path = current_paths(url, is_doc, page)
    v = {"version": ver, "date": TODAY, "url": url, "page": page, **info}
    if data is not None:
        orig_path.parent.mkdir(parents=True, exist_ok=True)
        orig_path.write_bytes(data)
        v["original"] = rel(orig_path)
        if md_text is None:
            try:
                md_text = to_markdown(orig_path, url)
            except Exception as e:
                warn(f"テキスト化に失敗: {url}: {e}")
    if md_text is not None:
        md_path.parent.mkdir(parents=True, exist_ok=True)
        md_path.write_text(md_text, "utf-8")
        v["markdown"] = rel(md_path)
        v["diff"], v["added"], v["removed"] = save_diff(url, old_md_text, old_name, v["markdown"], ver)
    versions.append(v)
    return v

def retire(entry, url, is_doc):
    """掲載が終わったページ・文書の最新版も archive/ に移す（documents/ は掲載中のものだけにする）"""
    v = latest(entry)
    if v:
        archive_version(v, url, is_doc, entry.get("page") or url)


def migrate(cat):
    """旧構成（documents/originals/www.ipa.go.jp/... の深い階層）で保存したファイルを、
    掲載ページ単位の新しい分類へ移す。台帳の保存先もあわせて更新する。"""
    moved = 0
    for kind, is_doc in (("pages", False), ("documents", True)):
        for url, e in cat[kind].items():
            page = url if not is_doc else (e.get("pages") or [url])[0]
            e["page"] = page
            for v in e.get("versions", []):
                if v.get("archived"):
                    orig, md = archive_paths(v.get("url", url), is_doc, v["version"], page)
                else:
                    orig, md = current_paths(v.get("url", url), is_doc, page)
                for key, dst in (("original", orig), ("markdown", md)):
                    cur = v.get(key)
                    if not cur or dst is None or cur == rel(dst):
                        continue
                    src = ROOT / cur
                    if not src.is_file() or dst.exists():
                        continue
                    dst.parent.mkdir(parents=True, exist_ok=True)
                    shutil.move(str(src), str(dst))
                    v[key] = rel(dst)
                    moved += 1
    return moved


# ---------- 文書の取得 ----------
def full_refresh():
    """週1回（月曜）と FULL_REFRESH=1 のときは Last-Modified を使わずに全文書を取り直す。
    サーバーの更新日時が変わらないまま中身だけ差し替えられた場合も検知するため"""
    return os.environ.get("FULL_REFRESH") == "1" or dt.datetime.now(dt.timezone(dt.timedelta(hours=9))).weekday() == 0

def fetch_doc(url, prev):
    last = latest(prev) or {}
    headers = {}
    if last.get("last_modified") and not last.get("archived") and not full_refresh():
        headers["If-Modified-Since"] = last["last_modified"]   # 変わっていなければ再ダウンロードしない
    r = get(url, headers)
    if r.status_code == 304:
        return None
    data = r.content
    if len(data) > MAX_BYTES:
        raise ValueError(f"サイズ超過 {len(data)} bytes")
    return data, r.headers.get("Last-Modified")


# ---------- 台帳と変更検知 ----------
def update_catalog(old, pages, docs, failed, gone):
    events, cat = [], {"pages": {}, "documents": {}}

    def ev(kind, title, url, v=None, note=None):
        e = {"date": TODAY, "type": kind, "title": title, "url": url}
        if note:
            e["note"] = note
        if v:
            e.update({k: v.get(k) for k in ("version", "original", "markdown", "diff", "added", "removed")})
        events.append(e)

    # --- ページ ---
    for url, info in pages.items():
        body = info.pop("body")
        prev = old["pages"].get(url, {})
        entry = {**info, "status": "active", "first_seen": prev.get("first_seen", TODAY),
                 "page": url, "versions": prev.get("versions", [])}
        last = latest(entry)
        if last and last["sha256"] == info["sha256"] and not last.get("archived"):
            cat["pages"][url] = entry                     # 本文に変化なし
            continue
        head = ["---", f"source: {url}", f"title: {json.dumps(info['title'], ensure_ascii=False)}",
                f"dates: {json.dumps(info['dates'], ensure_ascii=False)}", "---", ""]
        v = add_version(entry, url, False, {"sha256": info["sha256"]}, "\n".join(head) + body + "\n")
        cat["pages"][url] = entry
        if not prev or prev.get("status") != "active":
            note = None if info["via"] == "トップページ" else f"{info['via']}から追加"
            ev("ページ追加" if not prev else "ページ再掲載", info["title"], url, v, note)
            continue
        new_log = [e for e in info["changelog"] if e not in prev.get("changelog", [])]
        if new_log:
            for e in new_log:
                ev("ページ更新", f'{info["title"]} [{e["date"]}] {e["text"]}', url, v)
        elif prev.get("dates") != info["dates"]:
            ev("ページ更新", f'{info["title"]}（{" / ".join(info["dates"])}）', url, v)
        else:
            ev("本文変更（更新履歴に記載なし）", info["title"], url, v)

    for url, prev in old["pages"].items():
        if url in pages:
            continue
        if prev["status"] != "active":
            cat["pages"][url] = prev
            continue
        # 今回たどれなかったページ：取得失敗なら前回のまま、404 なら削除、まだあればリンク切れ（取得対象外）
        state = "error" if url in failed else "gone" if url in gone else probe(url)
        if state == "error":
            cat["pages"][url] = prev
            continue
        entry = {**prev, "status": "removed" if state == "gone" else "unlinked", "removed": TODAY}
        retire(entry, url, False)
        cat["pages"][url] = entry
        ev("ページ削除" if state == "gone" else "取得対象外（リンクがなくなった）", prev["title"], url)

    # --- 文書 ---
    gone_docs = {u: d for u, d in old["documents"].items() if u not in docs and d["status"] == "active"}
    replaced = set()
    for url, d in docs.items():
        prev = old["documents"].get(url, {})
        entry = {**prev, **d, "status": "active", "first_seen": prev.get("first_seen", TODAY),
                 "versions": prev.get("versions", [])}
        entry.pop("removed", None)
        entry["page"] = prev.get("page") or (d.get("pages") or [url])[0]
        try:
            got = fetch_doc(url, prev)
        except Gone:
            warn(f"リンク先の文書がありません（404）: {url}")
            if prev:
                cat["documents"][url] = prev
            continue
        except Exception as e:
            warn(f"文書の取得に失敗: {url}: {e}")
            if prev:
                cat["documents"][url] = {**prev, **d, "versions": prev.get("versions", [])}
            continue
        cat["documents"][url] = entry
        if got is None:
            continue                                      # 304: 変化なし
        data, last_modified = got
        sha = hashlib.sha256(data).hexdigest()
        last = latest(entry)
        if last and last["sha256"] == sha and not last.get("archived"):
            last["last_modified"] = last_modified         # 中身は同じ（日時だけ更新）
            continue

        info = {"sha256": sha, "size": len(data), "last_modified": last_modified}
        if not prev or prev.get("status") != "active":
            before = next((u for u, g in gone_docs.items() if same_doc(g, d)), None)
            if before:   # 同じ文書がファイル名を変えて掲載された＝改訂。旧ファイル名の版も引き継ぐ
                replaced.add(before)
                entry["replaces"] = before
                entry["versions"] = old["documents"][before].get("versions", [])
                v = add_version(entry, url, True, info, data=data)
                ev("文書差し替え", f'{d["title"]}（{leaf(before)} → {leaf(url)}）', url, v)
            else:
                v = add_version(entry, url, True, info, data=data)
                ev("文書追加" if not prev else "文書再掲載", d["title"], url, v)
        else:
            v = add_version(entry, url, True, info, data=data)
            ev("文書更新", d["title"], url, v)

    for url, prev in old["documents"].items():
        if url in cat["documents"]:
            continue
        if url in replaced:
            cat["documents"][url] = {**prev, "status": "replaced", "removed": TODAY, "versions": []}
            continue
        if prev["status"] != "active":
            cat["documents"][url] = prev
            continue
        # 掲載していたページを今回取得できなかったときは判断を保留する
        if any(p in failed for p in prev.get("pages", [])):
            cat["documents"][url] = prev
            continue
        entry = {**prev, "status": "removed", "removed": TODAY}
        retire(entry, url, True)
        cat["documents"][url] = entry
        ev("文書掲載終了", prev["title"], url)
    return cat, events


# ---------- README（概要のみ） ----------
def section_md(main, heading):
    h = next((h for h in main.find_all("h2") if heading in h.get_text()), None)
    if not h:
        return None
    html = ""
    for sib in h.find_next_siblings():
        if sib.name == "h2":
            break
        html += str(sib)
    return tidy(markdownify(html, heading_style="ATX", bullets="-"))

def overview_md(main):
    """トップページの「本制度の概要」。見出しが変わったときは最初の段落群を使う"""
    md = section_md(main, "概要")
    if md:
        return md
    first = main.find("h2")
    html = "".join(str(s) for s in (first.find_previous_siblings()[::-1] if first else main.children)
                   if getattr(s, "name", None) not in ("h1",) and "update-info" not in " ".join(s.get("class", []) if hasattr(s, "get") else []))
    return tidy(markdownify(html, heading_style="ATX", bullets="-")) or "_（取得できませんでした）_"

def write_readme(top, cat):
    tpl = (SCRIPT_DIR / "README.template.md").read_text("utf-8")
    active_pages = sum(1 for p in cat["pages"].values() if p["status"] == "active")
    active_docs = sum(1 for d in cat["documents"].values() if d["status"] == "active")
    last = next((e["date"] for e in cat["events"]), "—")
    parts = {"overview": overview_md(top), "pages": str(active_pages), "documents": str(active_docs),
             "last_change": last, "checked": TODAY}
    for k, v in parts.items():
        tpl = tpl.replace("{{" + k + "}}", v)
    (ROOT / "README.md").write_text(tpl, "utf-8")


# ---------- history/（更新履歴・文書一覧） ----------
def news_md(main, limit=10):
    h = next((h for h in main.find_all("h2") if "お知らせ" in h.get_text()), None)
    ul = h.find_next_sibling("ul") if h else None
    if not ul:
        return "_（取得できませんでした）_"
    out = []
    for li in ul.find_all("li", recursive=False)[:limit]:
        out.append(f"**{text_of(li.find('span'))}**\n")
        for sub in li.select("ul li"):
            a, t = sub.find("a"), text_of(sub)
            out.append(f"- [{t}]({a['href']})" if a else f"- {t}")
        out.append("")
    return "\n".join(out).strip()

def saved_links(v, base):
    out = []
    if v and v.get("original"):
        out.append(f"[原本]({link_from(v['original'], base)})")
    if v and v.get("markdown"):
        out.append(f"[md]({link_from(v['markdown'], base)})")
    return " / ".join(out)

def write_changelog(cat, top):
    base = HISTORY
    out = ["# 更新履歴", "",
           "アーカイブが検知した変更です（新しい順）。「差分を表示」を開くと前の版からの変更点（`-` 削除 / `+` 追加）が見られます。",
           "取得対象のページ・文書と過去の版は [documents.md](documents.md)、差分ファイルは [diffs/](diffs/) にあります。", "",
           "## 公式サイトのお知らせ（最新）", "", news_md(top), "", "## 検知した変更"]
    embedded, current = 0, None
    for e in cat["events"]:
        if e["date"] != current:
            current = e["date"]
            out += ["", f"### {current}", ""]
        stat = f"（+{e['added']} / -{e['removed']}）" if e.get("diff") else ""
        note = f"（{e['note']}）" if e.get("note") else ""
        links = " / ".join(x for x in (saved_links(e, base), f"[diff]({link_from(e['diff'], base)})" if e.get("diff") else "") if x)
        out.append(f"- **{e['type']}** [{esc(e['title'])}]({e['url']}){note}{stat}" + (f" — {links}" if links else ""))
        if e.get("diff") and embedded < CHANGELOG_DIFF_EVENTS:
            embedded += 1
            lines = read(e["diff"]).splitlines()
            more = len(lines) - CHANGELOG_DIFF_LINES
            out += ["", "  <details><summary>差分を表示</summary>", "", "  ````diff",
                    *("  " + l for l in lines[:CHANGELOG_DIFF_LINES]), "  ````", ""]
            if more > 0:
                out.append(f"  （残り {more} 行は [差分ファイル]({link_from(e['diff'], base)}) を参照）\n")
            out += ["  </details>", ""]
    if not cat["events"]:
        out += ["", "_（まだありません）_"]
    (HISTORY / "CHANGELOG.md").write_text("\n".join(out).rstrip() + "\n", "utf-8")
    # 過去の版へのリンクは archive/ に移ると変わるため、イベントのリンク先は台帳の最新パスに合わせて書き直している

def write_documents(cat, external=None):
    """取得対象のページ・文書と過去の版の一覧を history/documents.md に書く。"""
    base = HISTORY
    STATUS = {"active": "掲載中", "removed": "掲載終了", "unlinked": "取得対象外", "replaced": "差し替え済み"}
    out = ["# 取得対象のページと文書", "", f"最終確認：{TODAY}", "", "## ページ", "",
           "| ページ | 経路 | 最終更新（公式） | 状態 | 最新バージョン | 保存 |", "|---|---|---|---|---|---|"]
    for url, p in sorted(cat["pages"].items(), key=lambda x: (x[1]["status"] != "active", x[1].get("order", 999))):
        upd = next((re.sub(r"^.*?[：:]", "", d) for d in p.get("dates", []) if "更新" in d),
                   p["changelog"][0]["date"] if p.get("changelog") else "")
        out.append(f"| [{esc(p['title'])}]({url}) | {p.get('via', '')} | {upd} | {STATUS.get(p['status'], p['status'])} | "
                   f"{latest_version_label(p)} | {saved_links(latest(p), base)} |")

    out += ["", "## 文書", ""]
    by_page = {}
    for url, d in cat["documents"].items():
        if d["status"] == "replaced":
            continue
        key = d["pages"][0] if d["status"] == "active" and d.get("pages") else "_removed"
        by_page.setdefault(key, []).append((url, d))
    order = sorted((k for k in by_page if k != "_removed"), key=lambda u: cat["pages"].get(u, {}).get("order", 999))
    for purl in order + (["_removed"] if "_removed" in by_page else []):
        title = "掲載が終わった文書" if purl == "_removed" else cat["pages"].get(purl, {}).get("title", purl)
        out += [f"### {title}", "", "| 文書番号 | 文書 | 形式 | サイズ | 公式の更新日 | 初回取得 | 最終変更 | 最新バージョン | 保存 |",
                "|---|---|---|---|---|---|---|---|---|"]
        for url, d in sorted(by_page[purl], key=lambda x: (x[1].get("doc_no") or "~", x[1]["title"])):
            v = latest(d) or {}
            ext = pathlib.Path(urlparse(url).path).suffix[1:].upper()
            size = f"{v['size'] / 1024:,.0f} KB" if v.get("size") else ""
            out.append(f"| {d.get('doc_no') or ''} | [{esc(d['title'])}]({url}) | {ext} | {size} | "
                       f"{d.get('official_updated') or ''} | {d['first_seen']} | {v.get('date', '')} | "
                       f"{latest_version_label(d)} | {saved_links(v, base)} |")
        out.append("")

    out += ["## 過去の版", ""]
    groups = (
        ("ページの過去の版", list(cat["pages"].items())),
        ("文書の過去の版", [(u, e) for u, e in cat["documents"].items() if e["status"] != "replaced"]),
    )
    for heading, entries in groups:
        out += [f"### {heading}", ""]
        items = [(e["title"], u, e) for u, e in entries
                 if len(e.get("versions", [])) >= 2 or (e.get("versions") and e["status"] != "active")]
        if not items:
            out += ["_（過去の版はまだありません）_", ""]
            continue
        for title, url, e in sorted(items, key=lambda x: (latest(x[2]) or {}).get("date", ""), reverse=True):
            out += [f"<details><summary>{esc(title)}（{STATUS.get(e['status'], e['status'])}）</summary>", "",
                    "| バージョン | 保存先 | サイズ | 変更行 | 保存 | 差分 |", "|---|---|---|---|---|---|"]
            for v in reversed(e["versions"]):
                size = f"{v['size'] / 1024:,.0f} KB" if v.get("size") else ""
                chg = f"+{v.get('added', 0)} / -{v.get('removed', 0)}" if v.get("diff") else "—"
                diff = f"[diff]({link_from(v['diff'], base)})" if v.get("diff") else "—"
                where = "archive" if v.get("archived") else "documents（最新）"
                out.append(f"| {file_version(v, e)} | {where} | {size} | {chg} | {saved_links(v, base)} | {diff} |")
            out += ["", "</details>", ""]

    if external:
        out += ["## トップページからリンクされている外部サイト（取得対象外）", "",
                "取得の対象にする場合は `tools/archive.py` の `ALLOWED_HOSTS` にホスト名を追加してください。", ""]
        out += [f"- [{esc(t) or u}]({u})" for u, t in sorted(external.items())]
        out.append("")
    (HISTORY / "documents.md").write_text("\n".join(out).rstrip() + "\n", "utf-8")


def relink_events(cat):
    """イベントに記録した保存先を、台帳の現在のパス（archive/ に移った後のパス）に合わせる"""
    where = {}
    for kind in ("pages", "documents"):
        for e in cat[kind].values():
            for v in e.get("versions", []):
                where[(v.get("url"), v["version"])] = v
    for e in cat["events"]:
        v = where.get((e["url"], e.get("version")))
        if v:
            e["original"], e["markdown"] = v.get("original"), v.get("markdown")


# ---------- コミットメッセージ ----------
def write_commit_message(events):
    subject = f"[IPA] {len(events)}件の変更を検知 ({TODAY})" if events else f"[IPA] 台帳のみ更新 ({TODAY})"
    out = [subject, ""] + [f"- {e['type']}: {e['title']}" + (f"（{e['note']}）" if e.get("note") else "") for e in events]
    if ERRORS:
        out += ["", "取得できなかったもの（前回の状態のまま）:"] + [f"- {m}" for m in ERRORS]
    with_diff = [e for e in events if e.get("diff")]
    if with_diff:
        out += ["", "=" * 60, "文書の差分", "=" * 60]
        for e in with_diff:
            lines = read(e["diff"]).splitlines()
            block = ["", f"■ {e['title']}（+{e['added']} / -{e['removed']}）"] + lines[:MSG_DIFF_LINES]
            if len(lines) > MSG_DIFF_LINES:
                block.append(f"... (残り {len(lines) - MSG_DIFF_LINES} 行は {e['diff']} を参照)")
            if len("\n".join(out + block).encode()) > MSG_MAX_BYTES:
                out += ["", "... 以降の差分は省略しました。history/CHANGELOG.md または history/diffs/ を参照してください。"]
                break
            out += block
    (TOOL_ROOT / ".commit_msg").write_text("\n".join(out) + "\n", "utf-8")


# ---------- メイン ----------
def main():
    if not ROOT.is_dir():
        sys.exit(f"保存先のフォルダがありません（ARCHIVE_DIR を確認）: {ROOT}")
    old = json.loads(CATALOG.read_text("utf-8")) if CATALOG.exists() else {"pages": {}, "documents": {}, "events": []}
    pages, docs, failed, gone, external, top = {}, {}, set(), set(), {}, None
    for s in SCOPES:
        p, d, m, f, g, x = crawl(s)
        pages.update(p); docs.update(d); failed |= f; gone |= g; external.update(x)
        top = top or m

    active = sum(1 for p in old["pages"].values() if p["status"] == "active")
    if active and len(pages) < active * 0.5:   # サイト障害で「全部削除」と誤判定するのを防ぐ
        sys.exit(f"取得ページ数が急減しました（{active} → {len(pages)}）。中断します")

    moved = migrate(old)          # 旧構成（www.ipa.go.jp/... の深い階層）を掲載ページ単位の分類へ移す
    if moved:
        print(f"{moved} 件のファイルを新しいフォルダ構成に移しました")
    cat, events = update_catalog(old, pages, docs, failed, gone)
    cat["events"] = events + old.get("events", [])
    relink_events(cat)
    for d in (DOCS, ARCHIVE, HISTORY, DIFFS):   # 空でも git add できるようにフォルダを必ず作る
        d.mkdir(parents=True, exist_ok=True)
        (d / ".gitkeep").touch()
    CATALOG.write_text(json.dumps(cat, ensure_ascii=False, indent=2, sort_keys=True) + "\n", "utf-8")
    write_readme(top, cat)
    write_changelog(cat, top)
    write_documents(cat, external)
    write_commit_message(events)
    (TOOL_ROOT / ".errors").write_text("\n".join(ERRORS) + ("\n" if ERRORS else ""), "utf-8")

if __name__ == "__main__":
    main()
