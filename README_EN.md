# MeowTool Nav (miaonav)

> 🌐 Language: [中文](./README.md)

> A curated "tool website directory" generated from bookmarks, with category tiling, tag-linked filtering, dark mode and a local visual editor.

- 🌐 Live demo: **https://www.meowtool.com/miaonav**
- 🍴 Forked from: **https://github.com/Pintree-io/pintree/tree/pintree-old-pages** (Pintree legacy `pintree-old-pages` branch)
- ✍️ Author: Cheng

---

## 1. Introduction

This project is a derivative work built on the **Pintree** `pintree-old-pages` branch. The original Pintree turns browser bookmarks into a navigation website. On top of that, this fork ships a **productized, Chinese-user-oriented** redesign:

- A redesigned homepage (category tiling + two/three-level tag-linked filtering + scroll-spy highlighting);
- A brand-new **desktop visual editor** so you can manage the directory without installing a browser extension or hand-writing JSON;
- Polished branding, SEO and analytics, ready to publish.

---

## 2. Differences from the original Pintree (what's new)

### V1.0.1

Release Date: 2026.9.3

**1. Mobile Search Box (New)**

- Added a **sticky search bar** (`lg:hidden`, visible only on small screens) below the top bar, featuring a search icon + a "Clear" button, with search triggered by pressing Enter.
- Logic reuses the desktop `searchBookmarks` / `clearSearchResults` functions, with **input content synced across both ends**; clearing one also clears the other.
- Detail tweaks: The magnifying glass icon first received `my-auto` to fix vertical centering, and was then **completely removed** as per your request, with left padding reduced from `pl-9` to `pl-4` for a cleaner look.
- Also addresses a historical gap: the 640–1024px range previously had no search box at all, now covered by this mobile search bar.

**2. Mobile Card Layout (New Adaptation)**

- Within `@media (max-width:640px)`, changed the grid from `auto-fill minmax(210px,1fr)` to a fixed **`repeat(2, 1fr)`**, i.e., 2 cards per row on phones.
- Supporting adjustments: reduced card padding, icon size (40→32px), and description truncated to 1 line; category jump `scroll-margin-top` adjusted to 8.5rem to prevent titles from being obscured by the sticky search bar.

**3. SEO (Basic Completion)**

- The homepage previously had **no H1 at all**; the desktop sidebar brand name "Meow Tools Collection" has been upgraded to the **page's single `<h1>`**.
- The same text in the mobile top bar / drawer menu remains as `<a>` to avoid duplicate H1s diluting SEO weight.

**Files Involved**

- `index.html`: mobile search box DOM, H1 heading
- `css/styles.css`: mobile 2-column grid + card adaptations



### V1.0.1

Release Date: 2026.8.27

| Area | Original Pintree (pintree-old-pages) | This fork (miaonav) |
| --- | --- | --- |
| Editing data | Must install the "Pintree Bookmarks Exporter" Chrome extension, export a JSON, then manually replace `json/pintree.json` | Ships a **desktop GUI editor** (`Website navigation tool/Website navigation tool.py`) — import Excel/JSON, edit visually, sort, one-click export |
| Icon fetching | Relies on remote favicon URLs | Built-in **multi-source favicon downloader** (Google / faviconkit / Yandex / favicon.im fallback), downloads locally to `assets/logo/` and unifies to PNG |
| Homepage layout | Folder-card grid / bookmark list | **Category-tile homepage**: each category has an emoji icon and a site count |
| Filtering | Minimal | Per-category **secondary tags + expandable third-level linked tags**, 18-item preview + "View more" detail page |
| Navigation UX | Plain sidebar | Sidebar **scroll-spy highlighting** of the active category; detail pages with breadcrumbs and recursive grouping |
| Data fields | title / url / icon | Adds a **description field** (shown as a two-line card summary) |
| Branding & visuals | Generic Pintree skin | New "喵喵工具集 / meowtool" brand, custom logo, favicon, OG image, full `.mn-*` custom styles |
| Analytics | None | Umami + Google Analytics + Microsoft Clarity |
| SEO / compliance | Basic meta | Full canonical / Open Graph / Twitter Card |

### Highlights of the new capabilities

1. **Desktop visual editor (`Website navigation tool/Website navigation tool.py`)**
   - Import / export Excel (`.xlsx`) and JSON;
   - Table-based add / delete / edit of entries;
   - Search and filter;
   - Category sorting and link sorting (move up / down / top / bottom);
   - One-click export to Pintree-compatible `pintree.json`;
   - Auto-detects `assets/logo` under the project root and manages icons.
   - Dependencies: `openpyxl` (Excel), `requests` (icon download), `Pillow` (icon to PNG). Run with:
     ```bash
     pip install openpyxl requests pillow
     python "Website navigation tool/Website navigation tool.py"
     ```

2. **Category tiling + tag-linked filtering**
   - The homepage lays out top-level categories as tiled sections, each with an emoji icon and a total site count;
   - Each section shows a secondary-tag (sub-category) row that expands into third-level tags, forming a two/three-level linkage;
   - Each section previews the first 18 sites; "View more" opens the full detail page (breadcrumbs + recursive grouping).

3. **Scroll-spy highlighting**
   - An `IntersectionObserver` auto-highlights the corresponding sidebar category while scrolling the homepage, improving orientation in long lists.

4. **Local & unified icons**
   - The editor auto-downloads site favicons into `assets/logo/`, unifies them to PNG, removing the dependency on third-party realtime services for faster, more stable loads.

---

## 3. Project structure

```
miaonav/
├── index.html                 # Main navigation page (tiled home + tag filter + detail)
├── css/
│   ├── styles.css            # Custom styles (includes the new .mn-* series)
│   └── tailwind.css          # Tailwind build output
├── json/
│   └── pintree.json          # Directory data (categories / links / icons / descriptions)
├── assets/
│   ├── logo.svg              # Site logo
│   ├── og.webp               # Social share image
│   ├── favicon/              # Site favicon assets
│   ├── logo/                 # Per-site icons (downloaded by the editor)
│   └── default-icon.svg      # Fallback placeholder when an icon is missing
└── Website navigation tool/
    └── Website navigation tool.py   # Desktop visual editor (new)
```

---

## 4. Usage

### Option A: Desktop editor (recommended)
1. Install deps: `pip install openpyxl requests pillow`
2. Run `Website navigation tool/Website navigation tool.py`
3. Import existing `json/pintree.json` or an Excel file, edit and sort
4. Click "导出 json" to produce `json/pintree.json`; refresh the page

### Option B: Edit JSON manually (compatible with original)
Edit `json/pintree.json` following the Pintree structure:
```json
{
  "type": "folder",
  "title": "Search Tools",
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

### Local preview
For browser security reasons, serve over a local HTTP server (don't open `index.html` via `file://`):
```bash
python -m http.server 8000
# open http://localhost:8000
```

### Deploy
Host the whole directory (`index.html`, `css/`, `json/`, `assets/`) on any static host. The live version runs at **https://www.meowtool.com/miaonav**.

---

## 5. Acknowledgements

- The page skeleton and data format are derived from **Pintree** ([pintree-old-pages](https://github.com/Pintree-io/pintree/tree/pintree-old-pages)). Thanks to the original authors.
- The original project is released under the **MIT License**; this fork keeps the same license.

---

## 6. License

MIT License — modified from Pintree. Please retain the original author and project attribution.
