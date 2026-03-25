---
name: spfx-delivery-ops
description: >
  SPFx project delivery workflow — resolution hierarchy, canvas overrides,
  CI/CD automation, image deployment, ES5 gotchas, and full delivery checklist.
  Complements @sudharsank/spfx-enterprise-skills (coding standards).
  Use for new builds, redesigns, Classic-to-Modern migrations.
metadata:
  filePattern:
    - "**/src/webparts/**/*.ts"
    - "**/config/config.json"
    - "**/gulpfile.js"
    - "**/*.sppkg"
  bashPattern:
    - "gulp (bundle|serve|package-solution)"
    - "m365 spo"
    - "nvm use 18"
  priority: 80
---

# SPFx Delivery & Operations Skill

> This skill covers HOW TO DELIVER SPFx projects — the deployment pipeline, practical
> gotchas, and workflow patterns that come from hands-on SharePoint delivery.
> For coding standards, UX design, theming, accessibility, performance, CSS governance,
> and build toolchain, use `@sudharsank/spfx-enterprise-skills` (14 skills, 1058 lines).

---

## 1. Three-Tier Resolution Hierarchy

**Every SharePoint customisation decision must follow this order. No exceptions.**

| Tier | Approach | Characteristics |
|------|----------|----------------|
| **1. Out-of-box** | Native SP settings, page types, section layouts | No code. Supported. Survives updates. Always try first. |
| **2. Custom SPFx** | Webparts and extensions via the SPFx framework | Supported by Microsoft. Maintainable. Deployable via App Catalog. |
| **3. DOM hacks** | CSS injection, MutationObserver, `!important` overrides | **Unsupported.** Can break on any SP update without notice. Warn stakeholders explicitly. |

### Tier 1 — Native Settings to Exhaust First

Before writing any code, configure these via Site Settings or PowerShell:

```powershell
# Header layout: Minimal or Compact (shrinks native chrome)
# Site Settings → Change the look → Header

# Hide site title from header bar
# Change the look → Header → toggle title off

# Disable footer entirely
Set-PnPWeb -FooterEnabled $false

# Hide left/top navigation
Set-PnPWeb -QuickLaunchEnabled $false

# Set page as Home page (hides title area, author/date chrome)
# Site Pages → right-click page → Make homepage

# Use full-width section layout (removes column width constraints)
# Page editor → Add section → Full-width
```

### Tier 2 — Custom SPFx (the main build)

When OOB cannot achieve the design, build custom webparts. The modular architecture
decision framework:

| Page Complexity | Architecture | Trade-offs |
|----------------|-------------|------------|
| Standard layouts (images, text, hero, news) | OOB webparts only | Fast, limited visual fidelity |
| Quick demo, visual fidelity needed | Single monolithic SPFx webpart | Pixel-perfect but hardcoded |
| Production, site owners edit content | Modular SPFx webparts (3-6 per page) | Best balance of fidelity + editability |

**Rule of thumb:** Group visually contiguous sections into one webpart. Separate sections
with distinct interactive behaviour (carousels, tab panels). Never exceed 5-6 webparts
per page — too many makes editing fragile.

### Tier 3 — DOM Hacks (Last Resort with Warning)

If native settings and SPFx webparts leave gaps (e.g., the suite bar "waffle" menu
still visible), DOM manipulation is the only option. **Always document:**

1. What element is being overridden
2. Why Tier 1 and Tier 2 cannot solve it
3. The exact CSS selectors or DOM paths targeted
4. A warning that this may break on any SP update

---

## 2. Canvas Zone Override Patterns

SharePoint wraps every webpart in constrained `CanvasZone` / `CanvasSection` containers
that cap width at ~1200px. For full-width designs, apply this SCSS pattern:

### The Override (SCSS Module)

```scss
// In each webpart's .module.scss or in a shared _canvas-override.scss mixin
.myWebPart {
  width: 100%;

  // CRITICAL: Break out of SP canvas constraints
  :global {
    .CanvasZone {
      max-width: none !important;
      padding: 0 !important;
    }
    .CanvasSection {
      max-width: none !important;
      padding: 0 !important;
    }
  }
}
```

### Shared Mixin Pattern (Recommended)

Create `src/common/_canvas-override.scss` and import across all webparts:

```scss
// src/common/_canvas-override.scss
@mixin canvas-override {
  :global {
    .CanvasZone {
      max-width: none !important;
      padding: 0 !important;
    }
    .CanvasSection {
      max-width: none !important;
      padding: 0 !important;
    }
  }
}

@mixin webpart-root {
  width: 100%;
  @include canvas-override;
}
```

```scss
// In each webpart .module.scss
@import '../../common/canvas-override';

.heroCarousel {
  @include webpart-root;
  // ... webpart-specific styles
}
```

### Prerequisites

1. Set `supportedHosts: ["SharePointWebPart", "SharePointFullPage"]` in the manifest
2. Use **full-width section layout** in the page editor (Tier 1 first)
3. Apply the `:global` SCSS override only if Tier 1 full-width sections are still constrained

> **Warning:** `.CanvasZone` and `.CanvasSection` are internal SP class names. Microsoft
> can rename them without notice. SPFx 1.21 Flexible Sections may reduce or eliminate
> the need for this hack.

---

## 3. Background Persistence for SPA Re-renders

SharePoint Modern is a React SPA that continuously re-renders. Inline styles,
injected DOM elements, and attribute changes get wiped on navigation and re-render
cycles. This is the pattern to make styles stick.

### MutationObserver + requestAnimationFrame Pattern

Use this in Application Customizer extensions that need persistent visual changes:

```typescript
// In your ApplicationCustomizer's onInit() or after rendering placeholders

private _applyPersistentStyles(): void {
  // Set a data attribute for CSS specificity boost
  document.documentElement.setAttribute('data-custom-header', 'true');

  // Inject a stylesheet (survives longer than inline styles)
  const styleId = 'custom-persistent-styles';
  if (!document.getElementById(styleId)) {
    const style = document.createElement('style');
    style.id = styleId;
    style.textContent = `
      html[data-custom-header] .SPPageChrome {
        /* your overrides here */
      }
    `;
    document.head.appendChild(style);
  }
}

private _startObserver(): void {
  let frameRequested = false;

  const observer = new MutationObserver(() => {
    if (!frameRequested) {
      frameRequested = true;
      requestAnimationFrame(() => {
        this._applyPersistentStyles();
        frameRequested = false;
      });
    }
  });

  observer.observe(document.documentElement, {
    attributes: true,
    childList: true,
    subtree: true
  });
}
```

### Why Each Piece Matters

| Technique | Purpose |
|-----------|---------|
| `data-attribute` on `<html>` | Specificity boost without inline styles. `html[data-custom-header] .target` beats almost any SP selector. |
| `<style>` element injection | Survives longer than `element.style.x = y` which SP re-renders wipe. |
| `MutationObserver` | Detects when SP re-renders and reapplies your changes. |
| `requestAnimationFrame` | Batches rapid-fire mutation callbacks into one paint cycle. Without this, you get hundreds of reapplications per second. |

---

## 4. CI/CD Automation with GitHub Actions

### Certificate Auth Setup (PEM Format)

The `m365` CLI (CLI for Microsoft 365) authenticates via Entra ID app registration
with certificate auth. **Use PEM format, not PFX** — PFX requires a password parameter
that complicates CI secrets management.

```bash
# Generate certificate locally
openssl req -x509 -newkey rsa:2048 -keyout key.pem -out cert.pem -days 365 -nodes \
  -subj "/CN=SPFxDeploy"

# Combine into single PEM file for m365 CLI
cat key.pem cert.pem > combined.pem

# Base64 encode for GitHub secret storage
base64 -i combined.pem | tr -d '\n' > cert-base64.txt
```

Upload `cert.pem` to your Entra ID app registration (Certificates & secrets).
Store the base64-encoded combined PEM as a GitHub Actions secret.

### GitHub Actions Workflow

```yaml
name: Deploy SPFx to SharePoint

on:
  push:
    branches: [main]
    paths:
      - 'src/**'
      - 'config/**'
      - 'package.json'

jobs:
  build-deploy:
    runs-on: ubuntu-latest
    steps:
      - uses: actions/checkout@v4

      - uses: actions/setup-node@v4
        with:
          node-version: 18

      - run: npm ci

      - name: Bundle for production
        run: npx gulp bundle --ship

      - name: Package solution
        run: npx gulp package-solution --ship

      - name: Deploy to App Catalog
        run: |
          npx m365 login \
            --authType certificate \
            --certificateBase64Encoded "${{ secrets.M365_CERT_BASE64 }}" \
            --appId "${{ secrets.M365_APP_ID }}" \
            --tenant "${{ secrets.M365_TENANT_ID }}"

          npx m365 spo app add \
            --filePath sharepoint/solution/*.sppkg \
            --appCatalogUrl "${{ secrets.SP_APP_CATALOG_URL }}" \
            --overwrite

          npx m365 spo app deploy \
            --name "$(ls sharepoint/solution/*.sppkg | xargs basename)" \
            --appCatalogUrl "${{ secrets.SP_APP_CATALOG_URL }}"
```

### Required GitHub Secrets

| Secret | Value |
|--------|-------|
| `M365_CERT_BASE64` | Base64-encoded combined PEM (key + cert) |
| `M365_APP_ID` | Entra ID app registration Application (client) ID |
| `M365_TENANT_ID` | Azure AD tenant ID |
| `SP_APP_CATALOG_URL` | `https://contoso.sharepoint.com/sites/appcatalog` |

### Tenant-Wide Deployment

Set in `config/package-solution.json` to skip per-site activation:

```json
{
  "solution": {
    "skipFeatureDeployment": true
  }
}
```

---

## 5. Image Deployment to SiteAssets

### Upload Workflow

1. Prepare images locally (optimise, name descriptively)
2. Upload to the target site's document library (typically `SiteAssets/images/`)
3. After upload, SharePoint assigns each image a GUID — **ignore the GUID**
4. Reference images by filename, constructing the URL dynamically

### The GUID Trick

SharePoint stores images with internal GUIDs after upload, but they remain accessible
by their original filename path. Always reference by filename:

```typescript
// CORRECT: reference by filename — portable across sites
const imageUrl = `${baseUrl}/SiteAssets/images/hero-banner.jpg`;

// WRONG: hardcoded GUID URL — breaks when image is re-uploaded or site is copied
const imageUrl = `https://contoso.sharepoint.com/sites/Demo/SiteAssets/Forms/AllItems.aspx?id=%2Fsites%2FDemo%2FSiteAssets%2Fimages%2F{guid}`;
```

### Dynamic baseUrl Construction

Every webpart should build image paths from SP context + a configurable property:

```typescript
// In the webpart class render() method
const baseUrl = this.context.pageContext.web.absoluteUrl
  + '/' + (this.properties.siteAssetsPath || 'SiteAssets/images');

// Pass to React component
const element = React.createElement(MyComponent, {
  baseUrl: baseUrl,
  heroImage: this.properties.heroImageFilename || 'hero-default.jpg'
});
```

```tsx
// In the React component
<img src={`${this.props.baseUrl}/${this.props.heroImage}`} alt="Hero banner" />
```

### Pre-Deployment Image Checklist

- [ ] All referenced image filenames exist in the target document library
- [ ] No hardcoded URLs or tenant-specific paths in code
- [ ] Missing images cause silent 404s — check browser console after deployment
- [ ] The `siteAssetsPath` property is exposed in PropertyPane for site owners

---

## 6. ES5 Target Gotchas

SPFx `tsconfig.json` targets ES5. TypeScript transpiles `const/let` and arrow functions
correctly, but **runtime APIs that do not exist in ES5 are NOT polyfilled**. These will
fail silently or throw errors in older browsers.

### Forbidden APIs and Safe Alternatives

```typescript
// --- string.includes() ---
// BROKEN in ES5:
if (title.includes('Draft')) { ... }
// SAFE:
if (title.indexOf('Draft') > -1) { ... }

// --- Array.from() ---
// BROKEN in ES5:
const items = Array.from(nodeList);
// SAFE:
const items = Array.prototype.slice.call(nodeList);

// --- Object.assign() ---
// BROKEN in ES5:
const merged = Object.assign({}, defaults, overrides);
// SAFE (spread is transpiled by TS):
const merged = { ...defaults, ...overrides };
// OR manual merge:
const merged: Record<string, unknown> = {};
for (const key in defaults) {
  if (Object.prototype.hasOwnProperty.call(defaults, key)) {
    merged[key] = defaults[key];
  }
}
for (const key in overrides) {
  if (Object.prototype.hasOwnProperty.call(overrides, key)) {
    merged[key] = overrides[key];
  }
}

// --- for...of on non-arrays ---
// BROKEN in ES5 (requires Symbol.iterator):
for (const item of nodeList) { ... }
// SAFE:
for (let i = 0; i < nodeList.length; i++) {
  const item = nodeList[i];
  // ...
}

// --- template literals in innerHTML ---
// TypeScript transpiles these correctly in .ts files.
// But if you're injecting HTML strings in Application Customizers
// via innerHTML, the transpiled output uses string concatenation.
// This is SAFE — just be aware of the output:
topPlaceholder.domElement.innerHTML = `
  <div class="header">Welcome, ${userName}</div>
`;
// Transpiles to: '<div class="header">Welcome, ' + userName + '</div>'

// --- Array.prototype.find() ---
// BROKEN in ES5:
const match = items.find(item => item.id === targetId);
// SAFE:
let match: ItemType | undefined;
for (let i = 0; i < items.length; i++) {
  if (items[i].id === targetId) {
    match = items[i];
    break;
  }
}

// --- Array.prototype.findIndex() ---
// BROKEN in ES5:
const idx = items.findIndex(item => item.active);
// SAFE:
let idx = -1;
for (let i = 0; i < items.length; i++) {
  if (items[i].active) { idx = i; break; }
}

// --- Promise (if not polyfilled) ---
// SPFx includes a Promise polyfill, so Promise is SAFE.
// But async/await requires the tslib __awaiter helper —
// ensure tslib is in your dependencies (SPFx scaffolding includes it).
```

### Quick Reference Table

| ES5-Unsafe API | Safe Alternative |
|---------------|-----------------|
| `string.includes()` | `string.indexOf() > -1` |
| `Array.from()` | `Array.prototype.slice.call()` |
| `Object.assign()` | Spread syntax (TS transpiles) or manual merge |
| `for...of` on NodeList/Map/Set | Traditional `for` loop |
| `Array.find()` | `for` loop with `break` |
| `Array.findIndex()` | `for` loop returning index |
| `Object.entries()` | `Object.keys()` + index access |
| `Object.values()` | `Object.keys().map(k => obj[k])` |
| `String.startsWith()` | `string.indexOf(prefix) === 0` |
| `String.endsWith()` | `string.indexOf(suffix, string.length - suffix.length) !== -1` |
| `Number.isNaN()` | `isNaN()` (global function) |

---

## 7. config.json Registration

**This is the gotcha that kills builds silently.** Every webpart and extension MUST be
registered in `config/config.json`. If you add a new webpart and forget this step,
the build succeeds but the webpart does not appear in the toolbox.

### Webpart Registration

```json
{
  "bundles": {
    "hero-carousel-web-part": {
      "components": [
        {
          "entrypoint": "./lib/webparts/heroCarousel/HeroCarouselWebPart.js",
          "manifest": "./src/webparts/heroCarousel/HeroCarouselWebPart.manifest.json"
        }
      ]
    },
    "content-cards-web-part": {
      "components": [
        {
          "entrypoint": "./lib/webparts/contentCards/ContentCardsWebPart.js",
          "manifest": "./src/webparts/contentCards/ContentCardsWebPart.manifest.json"
        }
      ]
    }
  },
  "localizedResources": {
    "HeroCarouselWebPartStrings": "lib/webparts/heroCarousel/loc/{locale}.js",
    "ContentCardsWebPartStrings": "lib/webparts/contentCards/loc/{locale}.js"
  }
}
```

### Extension Registration

```json
{
  "bundles": {
    "custom-header-application-customizer": {
      "components": [
        {
          "entrypoint": "./lib/extensions/customHeader/CustomHeaderApplicationCustomizer.js",
          "manifest": "./src/extensions/customHeader/CustomHeaderApplicationCustomizer.manifest.json"
        }
      ]
    }
  }
}
```

### Diagnosis Checklist

If a webpart does not appear in the toolbox after deployment:

1. Check `config/config.json` — is the bundle registered?
2. Check the manifest file — does the `id` (GUID) match, is `componentType` correct?
3. Check `package-solution.json` — is `skipFeatureDeployment: true` set?
4. Check the App Catalog — was the solution trusted/deployed?
5. Clear browser cache and re-add the webpart from the toolbox

---

## 8. Node Version Constraints

SPFx 1.20 requires **Node.js 18.x** (>=18.17.1 <19.0.0). This is non-negotiable —
builds fail with cryptic errors on other versions.

### Project Setup

```bash
# .nvmrc (commit to repo root)
18
```

```bash
# Before any gulp or npm command
nvm use 18

# Verify
node --version  # Must show v18.x.x
```

### CI/CD Node Setup

```yaml
- uses: actions/setup-node@v4
  with:
    node-version: 18
```

### Common Symptoms of Wrong Node Version

| Symptom | Cause |
|---------|-------|
| `gulp bundle` hangs or crashes with `ERR_OSSL_EVP_UNSUPPORTED` | Node 17+ OpenSSL 3.0 breaking change |
| `npm install` fails on `node-gyp` native modules | Node version mismatch with prebuilt binaries |
| `Cannot find module 'node:fs'` | SPFx toolchain does not support Node.js `node:` prefix imports |
| Build completes but `.sppkg` is empty/corrupt | Node version silently produces bad output |

### Version Matrix

| SPFx Version | Required Node | Build System |
|-------------|--------------|-------------|
| 1.18-1.20 | 18.x | Gulp |
| 1.21 | 18.x | Gulp |
| 1.22+ | 18.x | Heft (Rush Stack) |

---

## 9. Multi-Page PropertyPane Pattern

For webparts with many configurable fields (10+), split across multiple PropertyPane
pages. Each page gets a header description and grouped fields.

```typescript
protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  return {
    pages: [
      // PAGE 1: Main Content
      {
        header: { description: 'Main Content Settings' },
        groups: [
          {
            groupName: 'Banner',
            groupFields: [
              PropertyPaneTextField('bannerTitle', { label: 'Banner Title' }),
              PropertyPaneTextField('bannerSubtitle', { label: 'Subtitle' }),
              PropertyPaneTextField('bannerImage', { label: 'Banner Image Filename' })
            ]
          },
          {
            groupName: 'Call to Action',
            groupFields: [
              PropertyPaneTextField('ctaText', { label: 'Button Text' }),
              PropertyPaneTextField('ctaUrl', { label: 'Button URL' }),
              PropertyPaneToggle('ctaEnabled', { label: 'Show CTA Button' })
            ]
          }
        ]
      },
      // PAGE 2: Carousel Slides (Numbered Slot Pattern)
      {
        header: { description: 'Carousel Slides' },
        groups: [
          {
            groupName: 'Slide 1',
            groupFields: [
              PropertyPaneTextField('slide1Title', { label: 'Title' }),
              PropertyPaneTextField('slide1Image', { label: 'Image Filename' }),
              PropertyPaneTextField('slide1Link', { label: 'Link URL' })
            ]
          },
          {
            groupName: 'Slide 2',
            groupFields: [
              PropertyPaneTextField('slide2Title', { label: 'Title' }),
              PropertyPaneTextField('slide2Image', { label: 'Image Filename' }),
              PropertyPaneTextField('slide2Link', { label: 'Link URL' })
            ]
          }
          // ... up to 6 slides — skip empty slots at render time
        ]
      },
      // PAGE 3: Display Settings
      {
        header: { description: 'Display & Behaviour' },
        groups: [
          {
            groupName: 'Layout',
            groupFields: [
              PropertyPaneDropdown('autoRotateInterval', {
                label: 'Auto-rotate Interval',
                options: [
                  { key: 0, text: 'Off' },
                  { key: 3000, text: '3 seconds' },
                  { key: 5000, text: '5 seconds' },
                  { key: 8000, text: '8 seconds' }
                ]
              }),
              PropertyPaneToggle('showNavDots', { label: 'Show Navigation Dots' }),
              PropertyPaneTextField('siteAssetsPath', {
                label: 'Image Library Path',
                description: 'Relative path from site root (e.g., SiteAssets/images)'
              })
            ]
          }
        ]
      }
    ]
  };
}
```

### Numbered Slot Pattern for Repeating Items

For configurable arrays (carousel slides, resource links, team members), use numbered
properties instead of complex collection editors:

```typescript
// At render time, build the array from numbered slots, skipping empty ones
private _buildSlides(): ISlide[] {
  const slides: ISlide[] = [];
  for (let i = 1; i <= 6; i++) {
    const title = (this.properties as Record<string, string>)[`slide${i}Title`];
    const image = (this.properties as Record<string, string>)[`slide${i}Image`];
    if (title && title.trim().length > 0) {
      slides.push({
        title: title,
        image: image || '',
        link: (this.properties as Record<string, string>)[`slide${i}Link`] || ''
      });
    }
  }
  return slides;
}
```

---

## 10. SP Page Analyzer / MCP Tooling

### sp-page-analyzer

A brand-agnostic Playwright extraction pipeline that reverse-engineers any live
SharePoint page. Run as a CLI tool or MCP server.

```bash
# CLI usage
node analyze.js https://contoso.sharepoint.com/sites/Intranet/Pages/home.aspx --auth
```

**6-step pipeline:**

| Step | Output |
|------|--------|
| 1. Load & Auth | Page fully rendered in browser (handles login redirects) |
| 2. Detect Type | Classic vs Modern, scroll container identification |
| 3. Design Tokens | Color palette, typography stack, spacing (ranked by frequency) |
| 4. Page Structure | Section tree with dimensions, styles, webpart IDs |
| 5. Image Inventory | All `<img>` + CSS `background-image` sources with dimensions |
| 6. Screenshots | Viewport, full-page, and scroll-segment captures |

**MCP server tools** (4 tools for Claude Code integration):

| Tool | Purpose |
|------|---------|
| `analyze_page` | Full 6-step pipeline on a URL |
| `extract_design_tokens` | Design tokens only (colors, fonts, sizes) |
| `capture_screenshots` | Visual reference screenshots |
| `list_images` | Complete image asset inventory |

### Graph API via MS365 MCP

Claude Code can read SP page structure directly via the ms365 MCP server:

```
# Read page canvas layout (Modern pages only)
read_resource(uri: "page:///sites/{siteId}/pages/{pageId}")
# Returns: canvasLayout with horizontalSections, columns, webparts

# Search SharePoint content
sharepoint_search(query: "quarterly report", limit: 20)

# Read files from document libraries
read_resource(uri: "file:///{driveId}/{itemId}")

# Search folders
sharepoint_folder_search(query: "SiteAssets", siteUrl: "https://contoso.sharepoint.com/sites/Demo")
```

**Key limitation:** Graph API has no equivalent for Classic page content. Classic pages
store content in wiki fields, Content Editor webparts, or Script Editor webparts —
none exposed via Graph. Use Playwright/sp-page-analyzer for Classic pages.

---

## 11. Classic-to-Modern Migration Workflow

### Two-Phase Strategy

**Phase 1: Rebuild Classic as Modern (fidelity first)**
- Analyse the Classic page using sp-page-analyzer (Playwright only — Graph API cannot read Classic)
- Extract design tokens, section structure, image inventory
- Map every Classic section to the resolution hierarchy (OOB, SPFx, or hack)
- Rebuild using Modern page types, native features first, then custom webparts
- Goal: visual parity with the Classic page on a Modern site

**Phase 2: Enhance with Modern capabilities**
- Add features impossible in Classic: responsive layouts, theme variants, PropertyPane editing
- Replace static content with dynamic sources (lists, Graph API, external APIs)
- Add AI-powered features (search, recommendations, content freshness)

### Classic Element Migration Map

| Classic Element | Modern Equivalent |
|----------------|-------------------|
| Master Page / Page Layout | Site design + SPFx Application Customizer |
| Content Editor Webpart | SPFx webpart with rich text property |
| Script Editor Webpart | SPFx webpart (no direct equivalent) |
| jQuery/Angular widgets | React components inside SPFx |
| `#s4-workspace` scroll | `[data-automation-id="contentScrollRegion"]` |
| Wiki page content zones | Modern page sections + Text webpart |
| SharePoint Designer workflows | Power Automate |
| InfoPath forms | Power Apps |

### Page Type Selection (Critical First Decision)

Choose the right Modern page type **before writing code** — wrong choice creates
unnecessary chrome-hiding hacks:

| Page Type | Native Chrome | Best For |
|-----------|--------------|----------|
| **Home page** | Minimal — no title area, no author/date | Landing pages, dashboards, portals |
| **Article page** | Title area, author, date, banner image | News articles, blog content |
| **News post** | Title, author, date, publishing metadata | Announcements, updates |

---

## 12. Full Delivery Checklist

### Phase 1: Discovery

- [ ] Define scope: new build / redesign / Classic-to-Modern migration
- [ ] Gather design requirements (mockups, reference pages, brand guidelines)
- [ ] Run `sp-page-analyzer` on source page (if replicating)
- [ ] Extract design tokens into `_tokens.scss` (colors, fonts, sizes, spacing)
- [ ] Map every page section to the resolution hierarchy (OOB → SPFx → hack)
- [ ] Choose correct page type (home, article, news post)
- [ ] Identify all image assets and prepare for SiteAssets upload
- [ ] Determine webpart count and boundaries (max 5-6 per page)

### Phase 2: Build

- [ ] `nvm use 18` — verify Node version
- [ ] Scaffold SPFx project (Yeoman or SPFx CLI for 1.23+)
- [ ] Create shared `_tokens.scss` and `_canvas-override.scss` mixins
- [ ] Build each webpart with full PropertyPane editability
- [ ] Use multi-page PropertyPane for 10+ fields
- [ ] Use numbered slot pattern for repeating items (slides, links, resources)
- [ ] Build Application Customizer for site-wide chrome if needed
- [ ] Register ALL bundles in `config/config.json` (the silent-fail gotcha)
- [ ] Avoid ES5-unsafe APIs (see Section 6 table)
- [ ] Run `gulp bundle --ship` — fix ALL TS/lint errors
- [ ] Run `gulp package-solution --ship` — verify bundle size (<200KB per webpart)

### Phase 3: Deploy

- [ ] Configure native site settings (Tier 1): header, footer, navigation, page type
- [ ] Upload images to SiteAssets document library
- [ ] Verify all referenced image filenames exist (no silent 404s)
- [ ] Deploy `.sppkg` to App Catalog (CI/CD or manual)
- [ ] Trust/deploy the solution in App Catalog
- [ ] Assemble page: add webparts in correct order, configure each via PropertyPane

### Phase 4: QA

- [ ] **Console errors**: No 404 images, no CSP violations, no unhandled exceptions
- [ ] **Visual fidelity**: Side-by-side comparison with design reference or source page
- [ ] **PropertyPane**: Every editable field updates the webpart live; no fields silently fail
- [ ] **Responsive**: Test at 375px, 768px, 1024px — no horizontal scroll, images scale
- [ ] **Accessibility**: Keyboard navigation works, interactive elements focusable, alt text present
- [ ] **Cross-browser**: Edge + Chrome (primary corporate SP browsers)
- [ ] **Live data**: API-driven webparts succeed; fallback/error states render gracefully
- [ ] **Page structure**: Verify via Graph API `canvasLayout` that sections match expectations

### QA Tools

| Tool | Purpose |
|------|---------|
| `sp-page-analyzer` | Before/after comparison (design tokens, section counts, image inventory) |
| SP Page Diagnostics Tool | Microsoft browser extension — slow webparts, large images, excessive API calls |
| Browser DevTools Console | 404s, CSP violations, JS errors |
| DevTools Responsive Mode | Layout at mobile/tablet/desktop breakpoints |

---

## 13. SPFx Version Roadmap (March 2026)

| Version | Key Changes | Build System |
|---------|-------------|-------------|
| **1.20** | Current stable. Node 18.x required. | Gulp |
| **1.21** | Flexible Sections — free drag/resize, may reduce CanvasZone hacks | Gulp |
| **1.22** | Gulp replaced by Heft (Rush Stack). Webpack remains. npm audit clean. | Heft |
| **1.23** | Open-sourced templates. New SPFx CLI (preview, replaces Yeoman). | Heft |

### Key Retirements

| Technology | Status | Migration |
|-----------|--------|-----------|
| Azure ACS (legacy add-in auth) | **Retired April 2026** | Entra ID |
| SharePoint Add-ins | Retired 2025 | SPFx |
| SharePoint 2013 Workflows | Retired 2025 | Power Automate |
| Field Customizers | Postponed retirement | Column Formatting |
| Gulp build system | Superseded in 1.22 | Heft |

---

## Quick Reference: Common Build Commands

```bash
# Environment setup
nvm use 18
npm ci                                    # Clean install from lockfile

# Development
gulp serve                                # Local workbench + hot reload

# Production build (SPFx <=1.20)
gulp bundle --ship                        # Minified production bundle
gulp package-solution --ship              # Generate .sppkg

# Production build (SPFx >=1.22)
npx heft build --production               # Replaces gulp bundle --ship

# Deploy via CLI for Microsoft 365
m365 login --authType certificate \
  --certificateBase64Encoded $CERT \
  --appId $APP_ID --tenant $TENANT_ID
m365 spo app add --filePath sharepoint/solution/*.sppkg \
  --appCatalogUrl $CATALOG_URL --overwrite
m365 spo app deploy --name solution.sppkg \
  --appCatalogUrl $CATALOG_URL

# Upgrade between SPFx versions
m365 spfx project upgrade --output md     # Generates migration steps
```
