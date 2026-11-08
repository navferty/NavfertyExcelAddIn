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

The build script checks generated local links, content/navigation markup, and
search ranking. The deployment pipelines synchronize only the generated site
to the web root and remove obsolete files from earlier site versions.
The old `.html` URLs are intentionally not kept or redirected; the published
pages use Hugo's clean URLs.

To add a language, add its Hugo language configuration, translated content, and
an `i18n/<language-code>.toml` catalog with the same keys as the existing
catalogs. The shared templates and JavaScript read translated UI strings from
those catalogs and enumerate configured languages; they do not branch on
specific language codes. Keep `params.defaultLanguage` aligned with
`defaultContentLanguage` for SEO `x-default` links.

Give informative images concise, contextual `alt` text. Use an empty `alt` only
for decorative icons whose meaning is already conveyed by nearby text; the
generated-site tests enforce this distinction.

The header search opens from its link or the `/` keyboard shortcut. Submitting
it opens the current language's search results; Escape closes the input.
Matches in titles rank first, followed by matches in descriptions and page text.
