#!/usr/bin/env python3
"""Regression checks for the built bilingual Hugo preview."""
import sys
import unittest
from html.parser import HTMLParser
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BUILD = Path(sys.argv.pop(1)).resolve()


class Elements(HTMLParser):
    def __init__(self, path):
        super().__init__(convert_charrefs=True)
        self.tags = []
        self.feed(path.read_text(encoding="utf-8"))

    def handle_starttag(self, tag, attrs):
        self.tags.append((tag, dict(attrs)))

    def find(self, tag, **attrs):
        return [
            element for name, element in self.tags
            if name == tag and all(element.get(key) == value for key, value in attrs.items())
        ]


class HugoOutputTests(unittest.TestCase):
    def test_keyboard_markup_is_preserved(self):
        for language in ("en", "ru"):
            for source in (ROOT / "Website" / language).rglob("*.html"):
                original = Elements(source)
                count = len(original.find("kbd"))
                if not count:
                    continue
                relative = source.relative_to(ROOT / "Website" / language)
                target = BUILD / language / relative.with_suffix("") / "index.html"
                if relative.name == "index.html":
                    target = BUILD / language / relative
                with self.subTest(page=str(relative), language=language):
                    self.assertEqual(len(Elements(target).find("kbd")), count)

    def test_header_search_and_translations(self):
        for language, other in (("en", "ru"), ("ru", "en")):
            page = Elements(BUILD / language / "install" / "index.html")
            with self.subTest(language=language):
                form = page.find("form", id="header-search-form")[0]
                self.assertEqual(form["action"], f"/{language}/search/")
                self.assertIn("hidden", form)
                self.assertEqual(page.find("input", id="header-search-query")[0]["name"], "q")
                self.assertEqual(page.find("a", **{"data-lang": other})[0]["href"], f"/{other}/install/")
                search = Elements(BUILD / language / "search" / "index.html")
                self.assertEqual(search.find("input", id="search-query")[0]["name"], "q")


if __name__ == "__main__":
    unittest.main()
