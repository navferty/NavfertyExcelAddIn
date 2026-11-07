#!/usr/bin/env python3
"""Regression checks for the built bilingual Hugo preview."""
import sys
import json
import tomllib
import unittest
from html.parser import HTMLParser
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
BUILD = Path(sys.argv.pop(1)).resolve()


class Elements(HTMLParser):
    def __init__(self, path):
        super().__init__(convert_charrefs=True)
        self.tags = []
        self.scripts = {}
        self.current_script = None
        self.feed(path.read_text(encoding="utf-8"))

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        self.tags.append((tag, attrs))
        if tag == "script" and attrs.get("id"):
            self.current_script = attrs["id"]
            self.scripts[self.current_script] = ""

    def handle_data(self, data):
        if self.current_script:
            self.scripts[self.current_script] += data

    def handle_endtag(self, tag):
        if tag == "script":
            self.current_script = None

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
                    self.assertGreaterEqual(len(Elements(target).find("kbd")), count)

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

    def test_localization_catalogs_cover_configured_languages(self):
        config = tomllib.loads((ROOT / "Website" / "hugo.toml").read_text(encoding="utf-8"))
        language_codes = set(config["languages"])
        self.assertEqual(config["params"]["defaultLanguage"], config["defaultContentLanguage"])
        catalogs = {
            path.stem: tomllib.loads(path.read_text(encoding="utf-8"))
            for path in (ROOT / "Website" / "i18n").glob("*.toml")
        }
        self.assertEqual(set(catalogs), language_codes)
        keys = set(next(iter(catalogs.values())))
        for language, catalog in catalogs.items():
            with self.subTest(language=language):
                self.assertEqual(set(catalog), keys)
                page = Elements(BUILD / language / "install" / "index.html")
                messages = json.loads(page.scripts["site-messages"])
                self.assertEqual(messages["searchFailed"], catalog["searchFailed"])
                self.assertFalse(page.find("script", src="/site.js"))
        client_script = (ROOT / "Website" / "hugo-assets" / "site.js").read_text(encoding="utf-8")
        self.assertNotIn("=== 'ru'", client_script)
        self.assertNotIn("=== 'en'", client_script)


if __name__ == "__main__":
    unittest.main()
