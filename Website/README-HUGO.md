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
```

The original HTML is retained while converted Markdown is reviewed. The
one-time conversion script requires `pandoc` and Python Beautiful Soup;
running it again overwrites the migrated content. Do not run it after editing
the Markdown unless those edits are backed up.
