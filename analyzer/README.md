# SP Page Analyzer

Brand-agnostic SharePoint page analysis tool. Extracts structure, design tokens, images, and screenshots from any SharePoint page — Classic or Modern.

## What It Does

1. **Detects page type** — Classic (master pages, `#s4-workspace`) vs Modern (React, `CanvasZone`)
2. **Extracts design tokens** — background colors, text colors, font families, font sizes (by frequency)
3. **Maps page structure** — sections, dimensions, background styles, child counts, webpart IDs
4. **Inventories images** — all `<img>` sources + CSS background images with natural/display dimensions
5. **Captures screenshots** — viewport, full-page, and scroll segments (handles inner scroll containers)
6. **Lists custom assets** — CSS/JS files loaded from Style Library or SiteAssets
7. **Extracts navigation** — nav links with text and URLs

## Usage

```bash
cd sp-page-analyzer
npm install

# Basic analysis
node analyze.js https://contoso.sharepoint.com/sites/Intranet/Pages/home.aspx

# With interactive login
node analyze.js https://contoso.sharepoint.com/sites/HR --auth

# Custom output dir and wait time
node analyze.js https://contoso.sharepoint.com/sites/Marketing --output ./marketing --wait 8

# Skip screenshots (faster, JSON only)
node analyze.js https://contoso.sharepoint.com/sites/IT --no-screenshots
```

## Output

```
output/
├── analysis.json          # Full structured analysis
├── screenshot-viewport.png  # Above-the-fold
├── screenshot-full.png      # Full page (when possible)
└── screenshot-scroll-N.png  # Scroll segments (Classic SP)
```

### analysis.json structure

```json
{
  "meta": {
    "pageType": "Classic|Modern|Unknown",
    "hasCustomJS": true,
    "scrollContainer": "#s4-workspace"
  },
  "designTokens": {
    "colors": { "backgrounds": {}, "text": {} },
    "typography": { "fontFamilies": [], "fontSizes": [] }
  },
  "structure": [
    { "index": 0, "name": "Hero Section", "dimensions": {...}, "styles": {...} }
  ],
  "images": [
    { "src": "https://...", "naturalWidth": 1920, "displayWidth": 1108 }
  ],
  "navigation": [
    { "text": "Home", "href": "https://..." }
  ],
  "customAssets": {
    "stylesheets": ["https://.../App.css"],
    "scripts": ["https://.../App.js"]
  }
}
```

## Authentication

The tool uses a persistent browser profile (`~/.sp-analyzer-browser`). On first run with `--auth`, sign in interactively. Subsequent runs reuse the session cookies.

## Next Steps

This tool is the foundation for:
- **SP Page Analyzer MCP Server** — expose as a Claude Code MCP tool
- **SharePoint Skill for Claude Code** — automated page analysis + SPFx scaffolding
