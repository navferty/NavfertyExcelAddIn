#!/usr/bin/env python3
"""Check local links and images in a built Hugo site."""
import sys
from html.parser import HTMLParser
from pathlib import Path
from urllib.parse import unquote, urljoin, urlsplit


class Document(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.ids = set()
        self.references = []

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        if attrs.get("id"):
            self.ids.add(attrs["id"])
        for attr in ("href", "src"):
            if attrs.get(attr):
                self.references.append((tag, attrs[attr]))


def main(root):
    root = root.resolve()
    documents = {}
    for path in root.rglob("*.html"):
        document = Document()
        document.feed(path.read_text(encoding="utf-8"))
        documents[path] = document
    errors = []
    for path, document in documents.items():
        relative = path.relative_to(root)
        page_url = "/" + relative.as_posix()
        if page_url.endswith("/index.html"):
            page_url = page_url[:-10]
        for tag, reference in document.references:
            target = urlsplit(urljoin("https://navferty.net" + page_url, reference))
            if target.netloc != "navferty.net" or target.scheme not in ("http", "https"):
                continue
            destination = (root / unquote(target.path).lstrip("/")).resolve()
            if not destination.is_relative_to(root):
                errors.append(f"{relative}: invalid path {reference}")
                continue
            if destination.is_dir():
                destination /= "index.html"
            if not destination.is_file():
                errors.append(f"{relative}: missing {tag} target {reference}")
            elif target.fragment and destination in documents and unquote(target.fragment) not in documents[destination].ids:
                errors.append(f"{relative}: missing fragment {reference}")
    if errors:
        print("\n".join(errors), file=sys.stderr)
        return 1
    print(f"Checked {len(documents)} HTML files: local links and fragments resolve.")
    return 0


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: verify-hugo.py <built-site-directory>")
    raise SystemExit(main(Path(sys.argv[1])))
