# Hugo migration preview

The existing static site remains the production source. The Hugo site is an
in-progress migration and must not be deployed until its content, accessibility,
SEO, and deployment cutover have been reviewed.

Install Hugo Extended 0.123.7, then run:

```sh
hugo server --source Website
```

To build and check local links:

```sh
hugo --source Website --destination /tmp/navferty-hugo-site --gc
python scripts/verify-hugo.py /tmp/navferty-hugo-site
python scripts/test-hugo-output.py /tmp/navferty-hugo-site
node scripts/test-hugo-search.js
```

The original HTML is retained while converted Markdown is reviewed. The
one-time conversion script requires `pandoc` and Python Beautiful Soup;
running it again overwrites the migrated content. Do not run it after editing
the Markdown unless those edits are backed up.

The header search opens from its link or the `/` keyboard shortcut. Submitting
it opens the current language's search results; Escape closes the input.
Matches in titles rank first, followed by matches in descriptions and page text.
