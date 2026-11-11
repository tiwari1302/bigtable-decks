# bigtable-decks

Each PowerPoint is one chapter of the site. Read a chapter in the browser, or download the `.pptx` from that chapter.

`chapters/` and `media/` are generated. Do not hand-edit them, and do not read them unless the task is about one specific slide.

To rebuild after a deck or the viewer changes:

```bash
python3 scripts/build_site.py
```

Edit `assets/site.css`, `assets/site.js`, or `scripts/build_site.py` for the site itself.
