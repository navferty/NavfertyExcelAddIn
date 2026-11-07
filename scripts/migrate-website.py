#!/usr/bin/env python3
"""Convert the legacy website into initial Hugo content.

Requires Beautiful Soup and pandoc. Review generated content before publishing.
"""
import json
import re
import subprocess
from pathlib import Path
from urllib.parse import urljoin, urlsplit, urlunsplit

from bs4 import BeautifulSoup

ROOT = Path(__file__).resolve().parents[1]
SITE = ROOT / "Website"
ORIGIN = "https://navferty.net"


def rewrite_url(value, original_url):
    if not value or value.startswith(("#", "mailto:", "tel:", "data:")):
        return value
    absolute = urlsplit(urljoin(original_url, value))
    if absolute.netloc != "navferty.net":
        return value
    path = absolute.path
    if path.endswith("/index.html"):
        path = path[:-10]
    elif path.endswith(".html"):
        path = path[:-5] + "/"
    return urlunsplit(("", "", path, absolute.query, absolute.fragment))


def convert(path, lang):
    relative = path.relative_to(SITE)
    original_url = f"{ORIGIN}/{relative.as_posix()}"
    soup = BeautifulSoup(path.read_text(encoding="utf-8"), "html.parser")
    main = soup.select_one("main#content")
    if main is None:
        raise ValueError(f"No main content in {path}")
    commands = main.select_one("ul.other-commands")
    if commands:
        heading = commands.find_previous_sibling("h2")
        if heading:
            heading.decompose()
        commands.decompose()
    for heading in main.select("h1[id], h2[id], h3[id], h4[id], h5[id], h6[id]"):
        anchor = soup.new_tag("a", id=heading["id"])
        heading.insert_before(anchor)
    for element in main.select("[href], [src]"):
        for attribute in ("href", "src"):
            if element.has_attr(attribute):
                element[attribute] = rewrite_url(element[attribute], original_url)
    description = soup.select_one('meta[name="description"]')
    heading = main.find("h1")
    title = heading.get_text(" ", strip=True) if heading else soup.title.get_text(strip=True)
    image = soup.select_one('meta[property="og:image"]')
    meta = {"title": title, "description": description.get("content", "") if description else ""}
    if image:
        meta["image"] = image["content"]
    content_path = path.relative_to(SITE / lang)
    if content_path.name == "index.html":
        output = SITE / "content" / lang / content_path.parent / "_index.md"
    else:
        output = SITE / "content" / lang / content_path.with_suffix(".md")
    output.parent.mkdir(parents=True, exist_ok=True)
    body = main.decode_contents()
    if content_path == Path("index.html"):
        markdown = body
    else:
        markdown = subprocess.run(
            ["pandoc", "-f", "html", "-t", "gfm", "--wrap=none"],
            input=body, text=True, capture_output=True, check=True,
        ).stdout
        markdown = re.sub(
            r"<img\b[^>]*>", lambda match: match.group(0) if re.search(r"\balt=", match.group(0))
            else match.group(0).replace("<img ", '<img alt="" ', 1), markdown,
        )
    output.write_text(json.dumps(meta, ensure_ascii=False, indent=2) + "\n\n" + markdown, encoding="utf-8")


if __name__ == "__main__":
    for language in ("en", "ru"):
        for source in sorted((SITE / language).rglob("*.html")):
            convert(source, language)
