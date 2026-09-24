# Meow Toolkit · Website Navigation (miaonav)

> 🌐 Language: [中文](./README.md)

A static, bookmark-generated "curated web tools directory" featuring category tiling, linked tag filtering, dark mode and a bilingual Chinese/English UI, plus a desktop visual editor.

- 🌐 Live demo: **https://www.meowtool.com/miaonav**
- 🍴 Forked from: [Pintree, `pintree-old-pages` branch](https://github.com/Pintree-io/pintree/tree/pintree-old-pages)
- ✍️ Author: Cheng

---

## 1. Features

**Navigation pages (`index.html` / `en.html`)**

- Bilingual: two same-source pages with a one-click language toggle, plus canonical / hreflang tags
- Homepage tiled into 9 top-level categories, each with an emoji and a site count; sidebar scroll-spy highlights the active category
- Two/three-level linked tag filtering; each section previews 18 sites and offers a "View more" detail page (breadcrumbs + recursive grouping)
- Search matches both site titles and descriptions; a sticky mobile search bar and a two-column mobile card layout are included
- Dark-mode preference is persisted in `localStorage` and survives language switches without flashing
- All site icons are served locally from `assets/logo/`, falling back to a default placeholder on error

**Desktop editor (`Website navigation tool/`)**

- Import / export Excel (`.xlsx`) and JSON; table-based editing, searching and sorting
- One-click export of Pintree-compatible `pintree.json`; multi-source favicon downloader converts icons to PNG
- Visual category emoji management (right-click "Set emoji…"; written into JSON on export)
- Built-in Baidu / Tencent translation incrementally produces `pintree.en.json` (batched disk writes; interrupted jobs resume)
- Ships in both a Chinese-UI build (`Website navigation tool.py`) and an English-UI build (`Website navigation tool EN.py`); the two are feature-equivalent

**Also included**

- Full SEO tags (canonical / Open Graph / Twitter Card)
- Umami, Google Analytics and Microsoft Clarity analytics

---

## 2. Project structure

```
miaonav/
├── index.html                     # Chinese navigation page
├── en.html                        # English navigation page (same source, kept in sync manually)
├── css/
│   ├── styles.css                 # Custom styles (.mn-* series + mobile adaptations)
│   └── tailwind.css               # Tailwind build output
├── json/
│   ├── pintree.json               # Chinese directory data (9 top-level categories, with emoji)
│   ├── pintree.en.json            # English directory data (produced by the translator)
│   └── _translation_config.json   # Translation API credentials (see notes below; never publish)
├── assets/
│   ├── logo.svg                   # Site logo
│   ├── og.webp                    # Social share image
│   ├── favicon/                   # Site favicons
│   ├── default-icon.svg           # Placeholder when a site icon fails to load
│   └── logo/                      # Per-site icons (local PNGs)
└── Website navigation tool/
    ├── Website navigation tool.py     # Desktop visual editor (Tkinter, Chinese UI)
    ├── Website navigation tool EN.py  # Desktop visual editor (Tkinter, English UI)
    ├── translation.py                 # Baidu / Tencent translation API wrapper
    └── _translation_cache.json        # Incremental translation cache (auto-generated at runtime)
```

---

## 3. Quick start

### 1. Local preview

The pages load JSON via `fetch`, so serve them over HTTP (don't open `file://` directly):

```bash
python -m http.server 8000
# Chinese: http://localhost:8000/index.html
# English: http://localhost:8000/en.html
```

### 2. Desktop editor (recommended)

```bash
pip install openpyxl requests pillow
python "Website navigation tool/Website navigation tool.py"
# For the English UI, use instead:
python "Website navigation tool/Website navigation tool EN.py"
```

1. Import the existing `json/pintree.json` or an Excel file, then edit, sort and assign emoji
2. Click "Export JSON" to generate `json/pintree.json`; refresh the page to see changes
3. The translation toolbar has 4 buttons:
   - **Translation settings…**: choose a provider (Baidu / Tencent), enter APPID/secret, set QPS rate, test the connection
   - **Translate now**: translates only new/missing entries against the Chinese JSON (incremental)
   - **Re-translate all**: clears the cache and fully re-translates, overwriting the English JSON
   - **Export English JSON**: translates the current in-memory data and outputs only `pintree.en.json`

> 🔑 Credentials live in `json/_translation_config.json` — replace them with your own Baidu / Tencent keys. The file contains secrets; in a public repo, untrack it from Git.

### 3. Manual JSON editing

The data follows the Pintree bookmark structure:

```json
{
  "type": "folder",
  "title": "Search Tools",
  "emoji": "🔍",
  "children": [
    {
      "type": "link",
      "title": "Felo",
      "icon": "assets/logo/felo.ai.png",
      "url": "https://felo.ai/search",
      "description": "AI-powered intelligent search platform"
    }
  ]
}
```

### 4. Deployment

Purely static — upload the whole directory to any static host. Every local resource (CSS / JS / images / JSON / inter-page links) uses relative paths, so the site works at the domain root or any subpath. Third-party analytics scripts and SEO canonical links are absolute URLs, as required.

---

## 4. Changelog

### V2.1.0 (2026-09-24)

- Added an English-UI desktop editor `Website navigation tool EN.py`: feature-equivalent to the Chinese build (`Website navigation tool.py`) with a fully English interface, for non-Chinese users

### V2.0.0 (2026-09-23)

- Translation upgrade: Baidu / Tencent providers and 4 toolbar buttons; the incremental cache is written atomically per batch, so interrupted/stopped jobs can resume without producing a dirty file; the progress window gained a progress bar
- New English page `en.html` with a language-toggle pill and full I18N of runtime strings
- Category icons now come from the JSON `emoji` field (configured for all 9 top-level categories)
- Dark-mode preference is persisted, so full-page language switches no longer flash back to light mode
- Search now matches site descriptions as well as titles
- All site icons and footer images are local relative references; the hardcoded `<base href="/miaonav/">` was removed

### V1.0.1 (2026-09-03)

- Added a sticky mobile search bar with content synced to the desktop box
- Two-card-per-row layout on phones, with tighter icons and spacing
- SEO fix: the sidebar brand name is now the page's single H1

### V1.0.0 (2026-08-27) | [Tag](https://github.com/cheng01315/miaonav/tree/v1.01_miaonav)

First fork release: desktop visual editor, tiled homepage with tag filtering, scroll-spy sidebar, localized icons, the description field, custom branding and analytics.

---

## 5. Acknowledgements

The page skeleton and data format are derived from **Pintree** ([pintree-old-pages](https://github.com/Pintree-io/pintree/tree/pintree-old-pages)). Thanks to the original authors.

---

## 6. License

MIT License — modified and redistributed from Pintree. Please retain the original author and project attribution.
