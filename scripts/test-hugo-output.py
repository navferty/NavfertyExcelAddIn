#!/usr/bin/env python3
"""Regression checks for the built Hugo site."""
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
        self.open_tags = []
        self.nested_anchors = []
        self.unnamed_links = []
        self.unnamed_buttons = []
        self.open_links = []
        self.open_buttons = []
        self.labels = set()
        self.controls = []
        self.feed(path.read_text(encoding="utf-8"))

    def handle_starttag(self, tag, attrs):
        attrs = dict(attrs)
        self.tags.append((tag, attrs))
        if tag == "a" and "a" in self.open_tags:
            self.nested_anchors.append(attrs.get("href"))
        if tag == "a":
            self.open_links.append({"href": attrs.get("href"), "name": attrs.get("aria-label", "")})
        if tag == "button":
            self.open_buttons.append({"name": attrs.get("aria-label", "")})
        if tag == "label" and attrs.get("for"):
            self.labels.add(attrs["for"])
        if tag in ("input", "select", "textarea"):
            self.controls.append((attrs, "label" in self.open_tags))
        if tag == "img":
            for link in self.open_links:
                link["name"] += attrs.get("alt") or ""
            for button in self.open_buttons:
                button["name"] += attrs.get("alt") or ""
        if tag == "script" and attrs.get("id"):
            self.current_script = attrs["id"]
            self.scripts[self.current_script] = ""
        if tag not in {
            "area", "base", "br", "col", "embed", "hr", "img", "input",
            "link", "meta", "param", "source", "track", "wbr",
        }:
            self.open_tags.append(tag)

    def handle_data(self, data):
        if self.current_script:
            self.scripts[self.current_script] += data
        for link in self.open_links:
            link["name"] += data
        for button in self.open_buttons:
            button["name"] += data

    def handle_endtag(self, tag):
        if tag == "script":
            self.current_script = None
        if tag == "a" and self.open_links:
            link = self.open_links.pop()
            if not link["name"].strip():
                self.unnamed_links.append(link["href"])
        if tag == "button" and self.open_buttons:
            button = self.open_buttons.pop()
            if not button["name"].strip():
                self.unnamed_buttons.append(button)
        while self.open_tags:
            if self.open_tags.pop() == tag:
                break

    def find(self, tag, **attrs):
        return [
            element for name, element in self.tags
            if name == tag and all(element.get(key) == value for key, value in attrs.items())
        ]


class HugoOutputTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.config = tomllib.loads((ROOT / "Website" / "hugo.toml").read_text(encoding="utf-8"))
        cls.languages = set(cls.config["languages"])

    def test_keyboard_markup_is_preserved(self):
        for language in self.languages:
            content_root = ROOT / "Website" / "content" / language
            for source in content_root.rglob("*.md"):
                original = Elements(source)
                count = len(original.find("kbd"))
                if not count:
                    continue
                relative = source.relative_to(content_root).with_suffix("")
                parts = list(relative.parts)
                if parts[-1] == "_index":
                    parts.pop()
                target = BUILD / language / Path(*parts) / "index.html"
                with self.subTest(page=str(relative), language=language):
                    self.assertGreaterEqual(len(Elements(target).find("kbd")), count)

    def test_header_search_and_translations(self):
        for language in self.languages:
            page = Elements(BUILD / language / "install" / "index.html")
            with self.subTest(language=language):
                form = page.find("form", id="header-search-form")[0]
                self.assertEqual(form["action"], f"/{language}/search/")
                self.assertIn("hidden", form)
                self.assertEqual(page.find("input", id="header-search-query")[0]["name"], "q")
                language_links = page.find("a", hreflang=language)
                self.assertEqual(language_links[0]["href"], f"/{language}/install/")
                self.assertEqual(language_links[0]["aria-current"], "page")
                search = Elements(BUILD / language / "search" / "index.html")
                self.assertEqual(search.find("input", id="search-query")[0]["name"], "q")

    def test_all_images_have_alternatives_or_are_decorative(self):
        for page_path in BUILD.rglob("*.html"):
            page = Elements(page_path)
            for image in page.find("img"):
                with self.subTest(page=str(page_path.relative_to(BUILD)), src=image.get("src")):
                    self.assertIn("alt", image)
                    if not (image["alt"] or "").strip():
                        self.assertIn("/icons/", image.get("src", ""))

    def test_links_and_form_controls_are_accessible(self):
        for page_path in BUILD.rglob("*.html"):
            page = Elements(page_path)
            relative = str(page_path.relative_to(BUILD))
            with self.subTest(page=relative):
                self.assertEqual(page.nested_anchors, [])
                self.assertEqual(page.unnamed_links, [])
                self.assertEqual(page.unnamed_buttons, [])
                for attrs, wrapped in page.controls:
                    if attrs.get("type") == "hidden":
                        continue
                    labelled = (
                        attrs.get("aria-label")
                        or attrs.get("aria-labelledby")
                        or attrs.get("id") in page.labels
                        or wrapped
                    )
                    self.assertTrue(labelled, f"Unlabelled form control: {attrs}")

    def test_localization_catalogs_cover_configured_languages(self):
        config = self.config
        language_codes = self.languages
        self.assertEqual(config["params"]["defaultLanguage"], config["defaultContentLanguage"])
        mounts = {
            mount["lang"]: mount["source"]
            for mount in config["module"]["mounts"]
            if mount.get("target") == "content" and "lang" in mount
        }
        self.assertEqual(set(mounts), language_codes)
        for language in language_codes:
            self.assertEqual(mounts[language], config["languages"][language]["contentDir"])
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

    def test_feature_index_and_unused_taxonomies(self):
        for language in self.languages:
            page = Elements(BUILD / language / "features" / "index.html")
            feature_files = list((ROOT / "Website" / "content" / language / "features").glob("*.md"))
            for source in feature_files:
                if source.stem != "_index":
                    with self.subTest(language=language, feature=source.stem):
                        self.assertTrue(page.find("a", href=f"/{language}/features/{source.stem}/"))
            self.assertFalse((BUILD / language / "categories" / "index.html").exists())
            self.assertFalse((BUILD / language / "tags" / "index.html").exists())


if __name__ == "__main__":
    unittest.main()
