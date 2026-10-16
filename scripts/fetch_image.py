"""从网上取背景介绍、原理讲解用的图片，并记录来源和许可。

用法：
    python fetch_image.py search "关键词" [--kind photo|diagram|any] [--limit 8]
    python fetch_image.py get "File:Xxx.svg" <输出目录> [--name 文件名] [--width 1600]
    python fetch_image.py add <图片 URL> <输出目录> --name 文件名 --title 题名 --source 来源 \\
        [--author 作者] [--license 许可] [--license-url 许可链接] [--date 发布日期] [--page 来源页面]

search：在 Wikimedia Commons 检索，列出文件名、像素尺寸、许可、作者和描述。常用词可用中文，专业术语用英文更准；
        要图上是中文标注的示意图，在关键词后加 intitle:zh。
        --kind diagram 只要矢量图（SVG，适合原理示意），photo 只要照片和位图。
get：   下载 Commons 文件。SVG 和大图取 --width 宽的 PNG / JPG 渲染（python-pptx 不能嵌 SVG）。
add：   下载并登记 Commons 以外任意网站的图（百度图片、百科、公众号、论文、厂商页面等）。
        用浏览器 UA 请求，给了 --page 时把它作为 Referer（公众号、百科的图床会查防盗链）；WebP 转成 PNG 保存。
        --page 必须是图片真正所在的网页：公众号图片配别的网站作 Referer 时，会返回「未经允许不可引用」的 140×140 占位图。

依赖 pillow：uv run --no-project --with pillow python -X utf8 fetch_image.py ...

每张图旁边写一个同名 .json，记下来源和许可（页面上不写出处，web_figure() 只检查 .json 存在）：
    reference  GB/T 7714-2015 电子资源著录：作者. 题名[EB/OL]. (发布日期)[引用日期]. URL.，需要参考文献页时用
    以及 title、author、source、license、license_url、page、image_url、date、accessed。
"""
import argparse
import datetime
import html
import io
import json
import re
import sys
import time
import urllib.parse
import urllib.request
from pathlib import Path

from PIL import Image

API = "https://commons.wikimedia.org/w/api.php"
# Wikimedia 要求自报来源的 User-Agent，否则可能被拒
UA = "hust-academic-ppt/1.0 (academic slides; python-urllib)"
# 其他网站（百度、公众号图床等）会拒绝非浏览器 UA
BROWSER_UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"
KIND_FILTER = {"photo": " filetype:bitmap", "diagram": " filetype:drawing", "any": ""}


def http_get(url, params=None, headers=None):
    if params:
        url += "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers=headers or {"User-Agent": UA})
    # Commons 连续请求十几次后会返回 429，按 Retry-After 等一会儿就能继续
    for attempt in range(4):
        try:
            with urllib.request.urlopen(req, timeout=60) as resp:
                return resp.read(), resp.headers.get("Content-Type", "")
        except urllib.error.HTTPError as e:
            if e.code != 429 or attempt == 3:
                raise RuntimeError(f"请求失败（HTTP {e.code}）：{url}") from e
            wait = int(e.headers.get("Retry-After") or 0) or 10 * 2 ** attempt
            print(f"限流（HTTP 429），{wait} 秒后重试", file=sys.stderr)
            time.sleep(wait)


def strip_html(s):
    """extmetadata 的值多为 HTML 字符串，也有数字（如 Assessments），统一转成纯文本。"""
    return re.sub(r"\s+", " ", html.unescape(re.sub(r"<[^>]+>", " ", str(s)))).strip()


def commons_query(params):
    body, _ = http_get(API, {"action": "query", "prop": "imageinfo", "iiprop": "url|size|mime|extmetadata",
                             "format": "json", "formatversion": 2, **params})
    return json.loads(body)["query"].get("pages", [])


def meta_of(page):
    """Commons 页面 → 来源记录。"""
    ii = page["imageinfo"][0]
    m = {k: strip_html(v["value"]) for k, v in ii["extmetadata"].items()}
    author = m.get("Artist") or "佚名"
    if re.match(r"(?i)unknown", author):  # Commons 的「Unknown author」模板
        author = "佚名"
    date = re.search(r"\d{4}(-\d{2}(-\d{2})?)?", m.get("DateTimeOriginal", "") or m.get("DateTime", ""))
    return {
        "title": m.get("ObjectName") or Path(page["title"].removeprefix("File:")).stem,
        "author": author,
        "source": "Wikimedia Commons",
        "license": m.get("LicenseShortName", ""),
        "license_url": m.get("LicenseUrl", ""),
        "page": ii["descriptionurl"],
        "image_url": (ii.get("thumburl") or ii["url"]).split("?")[0],  # 去掉 utm 跟踪参数
        "date": date.group(0) if date else "",
        "size": (ii["width"], ii["height"]),
        "mime": ii["mime"],
        "description": m.get("ImageDescription", ""),
    }


def search(query, kind, limit):
    pages = commons_query({"generator": "search", "gsrsearch": query + KIND_FILTER[kind], "gsrnamespace": 6,
                           "gsrlimit": limit})
    if not pages:
        print("没有结果。换英文关键词，或把 --kind 改成 any。")
    for p in sorted(pages, key=lambda p: p["index"]):
        r = meta_of(p)
        print(f"{p['title']}\n    {r['size'][0]}×{r['size'][1]}  {r['mime']}  {r['license']}  作者：{r['author'][:40]}"
              f"\n    {r['description'][:80]}\n    {r['page']}")


def reference(r, accessed):
    """GB/T 7714-2015 电子资源著录。"""
    date = f"({r['date']})" if r["date"] else ""
    return f"{r['author']}. {r['title']}[EB/OL]. {date}[{accessed}]. {r['page'] or r['image_url']}."


def save(r, image_url, out, name, headers=None):
    data, ctype = http_get(image_url, headers=headers)
    ctype = ctype.split(";")[0].strip()
    if ctype == "image/webp":  # 公众号、百度的图常是 WebP，PPT 不支持
        buf = io.BytesIO()
        Image.open(io.BytesIO(data)).save(buf, "PNG")
        data, ctype = buf.getvalue(), "image/png"
    if not ctype.startswith("image/"):
        raise ValueError(f"{image_url} 返回的不是图片（Content-Type: {ctype}），可能是防盗链或验证页，"
                         "用 --page 给出图片所在网页，或换图片直链")
    ext = {"image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif"}.get(ctype)
    if ext is None:
        raise ValueError(f"{image_url} 是 {ctype}，PPT 只用 PNG / JPG / GIF，换一张或用 Commons 的渲染图")
    out.mkdir(parents=True, exist_ok=True)
    img = out / f"{name}{ext}"
    img.write_bytes(data)
    accessed = datetime.date.today().isoformat()
    record = {k: r.get(k, "") for k in ("title", "author", "source", "license", "license_url", "page", "date")}
    record.update(image_url=image_url, accessed=accessed, reference=reference(r, accessed))
    img.with_suffix(".json").write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding="utf-8")
    # 描述里常有图上编号的含义（如「1) 阴极 2) 阳极」），写图注和讲解时要和图上的标号对应
    w, h = Image.open(img).size  # 防盗链占位图、图标这类小图在这里就能看出来
    print(f"written {img}  {w}×{h}\n  {r['author']} / {r['source']} / {r['license']}\n  {record['reference']}"
          f"\n  描述：{r.get('description', '')[:300]}")


def get(file, out, name, width):
    title = file if file.startswith("File:") else "File:" + file
    pages = commons_query({"titles": title, "iiurlwidth": width})
    if not pages or "imageinfo" not in pages[0]:
        raise ValueError(f"Commons 上没有 {title}")
    r = meta_of(pages[0])
    if not r["license"]:
        raise ValueError(f"{title} 没有许可信息，不能使用")
    save(r, r["image_url"], out, name or re.sub(r"[^\w\-]+", "_", r["title"]))


def main():
    ap = argparse.ArgumentParser()
    sub = ap.add_subparsers(dest="cmd", required=True)
    s = sub.add_parser("search")
    s.add_argument("query")
    s.add_argument("--kind", choices=list(KIND_FILTER), default="any")
    s.add_argument("--limit", type=int, default=8)
    g = sub.add_parser("get")
    g.add_argument("file")
    g.add_argument("out")
    g.add_argument("--name")
    g.add_argument("--width", type=int, default=1600)
    a = sub.add_parser("add")
    a.add_argument("url")
    a.add_argument("out")
    for k in ("name", "title", "source"):
        a.add_argument("--" + k, required=True)
    a.add_argument("--author", default="佚名")  # 网图常查不到作者，GB/T 7714 著录为「佚名」
    for k in ("license", "license-url", "date", "page"):
        a.add_argument("--" + k, default="")
    args = ap.parse_args()

    if args.cmd == "search":
        search(args.query, args.kind, args.limit)
    elif args.cmd == "get":
        get(args.file, Path(args.out), args.name, args.width)
    else:
        r = {"title": args.title, "author": args.author, "source": args.source, "license": args.license,
             "license_url": args.license_url, "date": args.date, "page": args.page, "image_url": args.url}
        headers = {"User-Agent": BROWSER_UA}
        if args.page:
            headers["Referer"] = args.page
        save(r, args.url, Path(args.out), args.name, headers)


if __name__ == "__main__":
    main()
