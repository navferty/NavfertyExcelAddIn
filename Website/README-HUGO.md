# Website

Hugo is the source and build system for the production website. The generated
site is deployed by `.github/workflows/website.yml` on pushes to `master` and by
`.github/workflows/release.yml` when a release is published. Both pipelines
build and validate the same Hugo output before deployment.

Install Hugo Extended 0.123.7, then run:

```sh
bash scripts/build-website.sh /tmp/navferty-website
hugo server --source Website
```

For Visual Studio on Windows, install Hugo Extended 0.123.7 and ensure `hugo`
is on `PATH` (restart Visual Studio after changing `PATH`). Open
`NavfertyExcelAddIn.slnx`, right-click the `Website` project and select
**Set as Startup Project**, then press **F5** or **Ctrl+F5**. The launcher
starts Hugo on localhost, opens the browser when the site is ready and watches
for edits. Stop debugging, close the launcher console, or press Ctrl+C to stop
the server. F5 debugs the launcher, not Hugo templates; use browser developer
tools to debug site JavaScript. An unavailable Hugo executable is reported in
the launcher's console; versions other than Extended 0.123.7 produce a warning.
To start without opening a browser, pass `--no-browser` to the launcher.

The build script checks generated local links, content/navigation markup,
language selection (including a temporary region-specific language), and search
ranking. Use a fresh destination directory for each build so removed pages do
not remain in the output. The deployment pipelines synchronize only the generated site
to the web root and remove obsolete files from earlier site versions.
The old `.html` URLs are intentionally not kept or redirected; the published
pages use Hugo's clean URLs.

## Adding a language

1. Add `[languages.<code>]` to `hugo.toml`, with `languageName`, `languageCode`,
   `weight`, and `contentDir = "content/<code>"`. Add a matching
   `[[module.mounts]]` entry with `source = "content/<code>"`,
   `target = "content"`, and `lang = "<code>"`. Use the same code in the
   language key, directory, mount and catalog; lowercase region-specific codes
   such as `pt-br` work too.
2. Add `i18n/<code>.toml` with all keys from an existing catalog, translated
   for that language. Add translated home, articles, features, search, install,
   and support content under `content/<code>/`. Match relative filenames across
   languages so Hugo associates translations. Translate descriptions, image
   alternatives and any embedded links; don't copy `/en/` or `/ru/` links
   unchanged. A page missing a translation will not appear in that page's
   language menu.
3. Build into a fresh directory using the command above. Check the generated
   language home, feature index, search index, translation links, sitemap and
   root redirect. Follow [the manual checklist](TESTING.md), including keyboard,
   screen-reader and narrow-screen checks for the new language.

The shared templates and JavaScript enumerate configured languages and use the
catalogs rather than branching on specific codes. A localized URL always opens
that language and saves it for later root visits. At the root, a supported saved
language wins, then browser languages are matched in order (including a fallback
from regional to less specific codes), then the default language is used.
`websiteLanguage` is the current storage key; the old site's preference is not
read. If storage is unavailable, the browser preference and default still work.
Keep `params.defaultLanguage` aligned with `defaultContentLanguage` for SEO
`x-default` links.

Give informative images concise, contextual `alt` text. Use an empty `alt` only
for decorative icons whose meaning is already conveyed by nearby text; the
generated-site tests enforce this distinction.

The header search opens from its link or the `/` keyboard shortcut. Submitting
it opens the current language's search results; Escape closes the input.
Matches in titles rank first, followed by matches in descriptions and page text.
