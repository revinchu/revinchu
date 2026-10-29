"""매거진: content/articles/<slug>.<lang>.md (앞머리 + 간단한 마크다운) → HTML.

앞머리(--- 사이)에 title, description, category, published, updated, hero, hero_alt, hero_credit, answer, tags.
본문 문법: '## ' 소제목(목차), '### ' 하위 제목(FAQ 에서는 질문), '- '/'1. ' 목록, '> ' 팁 상자,
'![설명](이미지 "출처")' 사진, '| 표 |', '**굵게**', '[링크](/{lang}/hospitals)' ({lang} 은 언어 코드로 바뀜).
'## FAQ'(언어별 표기) 구역은 FAQPage 구조화 데이터, '## Sources' 구역은 출처 목록이 된다.
글은 import 시점에 한 번만 변환한다 (Cloudflare 메모리 스냅샷에 포함 → 요청마다 CPU 를 쓰지 않음).
"""
import html
import os
import re

try:  # Cloudflare Workers 빌드는 글을 _embedded.py 로 묶는다 (cloudflare/build.py)
    from . import _embedded
except ImportError:
    _embedded = None

CATEGORIES = ["k-medical", "k-beauty", "k-pop", "k-food", "travel"]
FAQ_HEADINGS = {"faq", "자주 묻는 질문", "常见问题", "よくある質問", "câu hỏi thường gặp"}
SOURCE_HEADINGS = {"sources", "출처", "参考资料", "出典", "nguồn tham khảo"}


def _raw_files():
    if _embedded is not None and getattr(_embedded, "ARTICLES", None):
        return _embedded.ARTICLES
    folder = os.path.join(os.path.dirname(__file__), "content", "articles")
    if not os.path.isdir(folder):
        return {}
    return {f: open(os.path.join(folder, f), encoding="utf-8").read() for f in sorted(os.listdir(folder)) if f.endswith(".md")}


def _inline(text, lang):
    text = html.escape(text, quote=False)
    text = re.sub(r"\*\*(.+?)\*\*", r"<strong>\1</strong>", text)

    def link(m):
        label, url = m.group(1), m.group(2).replace("{lang}", lang)
        external = url.startswith("http")
        attrs = ' target="_blank" rel="noopener"' if external else ""
        return f'<a href="{html.escape(url)}"{attrs}>{label}</a>'
    return re.sub(r"\[([^\]]+)\]\(([^)\s]+)\)", link, text)


def _plain(text):
    text = re.sub(r"\*\*(.+?)\*\*", r"\1", text)
    return re.sub(r"\[([^\]]+)\]\([^)]+\)", r"\1", text)


def _anchor(text, n):
    slug = re.sub(r"[^\w\s-]", "", text.lower()).strip()
    slug = re.sub(r"[\s_]+", "-", slug)
    return slug[:60] if re.fullmatch(r"[a-z0-9-]+", slug or "") else f"s{n}"


def _render(body, lang):
    out, toc, faq, sources = [], [], [], []
    section, question = "", None
    for n, block in enumerate(re.split(r"\n\s*\n", body.strip())):
        lines = [l.rstrip() for l in block.strip().split("\n")]
        first = lines[0]
        if first.startswith("## "):
            title = first[3:].strip()
            low = title.lower()
            section = "faq" if low in FAQ_HEADINGS else "sources" if low in SOURCE_HEADINGS else ""
            anchor = "faq" if section == "faq" else "sources" if section == "sources" else _anchor(title, n)
            if section != "sources":
                toc.append((anchor, title))
            out.append(f'<h2 id="{anchor}">{_inline(title, lang)}</h2>')
            lines = lines[1:]
            if not lines:
                continue
            first = lines[0]
        if first.startswith("### "):
            title = first[4:].strip()
            if section == "faq":
                question = {"q": _plain(title), "a": []}
                faq.append(question)
                out.append(f'<h3 class="faq-q">{_inline(title, lang)}</h3>')
            else:
                out.append(f"<h3>{_inline(title, lang)}</h3>")
            lines = lines[1:]
            if not lines:
                continue
            first = lines[0]
        if all(l.startswith("- ") for l in lines):
            items = [l[2:] for l in lines]
            if section == "sources":
                sources += items
            out.append("<ul>" + "".join(f"<li>{_inline(i, lang)}</li>" for i in items) + "</ul>")
        elif all(re.match(r"\d+\. ", l) for l in lines):
            start = int(re.match(r"\d+", first).group())  # 소제목으로 나뉜 번호 목록은 이어진 번호로 시작
            items = [re.sub(r"^\d+\. ", "", l) for l in lines]
            out.append((f'<ol start="{start}">' if start != 1 else "<ol>")
                       + "".join(f"<li>{_inline(i, lang)}</li>" for i in items) + "</ol>")
        elif first.startswith("> "):
            text = " ".join(l[2:] if l.startswith("> ") else l for l in lines)
            out.append(f'<aside class="tip">{_inline(text, lang)}</aside>')
        elif first.startswith("!["):
            m = re.match(r'!\[([^\]]*)\]\((\S+)(?:\s+"([^"]*)")?\)', first)
            if m:
                alt, url, credit = m.group(1), m.group(2), m.group(3) or ""
                cap = html.escape(alt) + (f' <span class="credit">© {html.escape(credit)}</span>' if credit else "")
                out.append(f'<figure><img src="{html.escape(url)}" alt="{html.escape(alt)}" loading="lazy" '
                           f'referrerpolicy="no-referrer"><figcaption>{cap}</figcaption></figure>')
        elif first.startswith("|"):
            rows = [[c.strip() for c in l.strip("|").split("|")] for l in lines if not re.match(r"^\|[\s:|-]+\|$", l)]
            head, body_rows = rows[0], rows[1:]
            out.append('<div class="table-wrap"><table class="data-table"><thead><tr>'
                       + "".join(f"<th>{_inline(c, lang)}</th>" for c in head) + "</tr></thead><tbody>"
                       + "".join("<tr>" + "".join(f"<td>{_inline(c, lang)}</td>" for c in r) + "</tr>" for r in body_rows)
                       + "</tbody></table></div>")
        else:
            text = " ".join(lines)
            if section == "faq" and question is not None:
                question["a"].append(_plain(text))
            out.append(f"<p>{_inline(text, lang)}</p>")
    return "\n".join(out), toc, [{"q": f["q"], "a": " ".join(f["a"])} for f in faq], sources


def _reading_minutes(body, lang):
    text = _plain(re.sub(r"!\[[^\]]*\]\([^)]*\)", "", body))
    if lang in ("ko", "zh", "ja"):
        per_min = 500 if lang == "ko" else 400
        return max(1, round(len(re.sub(r"\s", "", text)) / per_min))
    return max(1, round(len(text.split()) / 230))


def _parse(name, raw):
    slug, lang = name[:-3].rsplit(".", 1)
    meta, body = {}, raw
    m = re.match(r"^---\n(.*?)\n---\n(.*)$", raw.replace("\r", ""), re.S)
    if m:
        for line in m.group(1).split("\n"):
            if ":" in line:
                k, v = line.split(":", 1)
                meta[k.strip()] = v.strip()
        body = m.group(2)
    content, toc, faq, sources = _render(body, lang)
    return {
        "slug": slug, "lang": lang, "title": meta.get("title", slug), "description": meta.get("description", ""),
        "category": meta.get("category", "travel"), "published": meta.get("published", ""),
        "updated": meta.get("updated") or meta.get("published", ""), "hero": meta.get("hero", ""),
        "hero_alt": meta.get("hero_alt", ""), "hero_credit": meta.get("hero_credit", ""),
        "answer": meta.get("answer", ""), "tags": [t.strip() for t in meta.get("tags", "").split(",") if t.strip()],
        "html": content, "toc": toc, "faq": faq, "sources": sources, "minutes": _reading_minutes(body, lang),
    }


ARTICLES = {}  # slug → {lang: article}
for _name, _raw in _raw_files().items():
    _a = _parse(_name, _raw)
    ARTICLES.setdefault(_a["slug"], {})[_a["lang"]] = _a


def languages(slug):
    return sorted(ARTICLES.get(slug, {}))


def get(slug, lang):
    """요청 언어 글. 없으면 영어(→ 한국어) 판을 fallback=True 로 돌려준다 (canonical 은 원문으로)."""
    versions = ARTICLES.get(slug)
    if not versions:
        return None
    if lang in versions:
        return {**versions[lang], "fallback": False}
    base = versions.get("en") or versions.get("ko") or next(iter(versions.values()))
    return {**base, "fallback": True}


def listing(lang, category=None):
    items = [get(slug, lang) for slug in ARTICLES]
    items = [a for a in items if a and (category is None or a["category"] == category)]
    return sorted(items, key=lambda a: (a["published"], a["slug"]), reverse=True)
