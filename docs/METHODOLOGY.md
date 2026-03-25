# SharePoint Development Learnings — Page Replication Methodology
## Code & Canvas | March 2026

---

## 1. Analysing a SharePoint Page (Classic or Modern)

### Techniques Used

| Technique | What It Reveals | Limitations | How We Overcame |
|-----------|----------------|-------------|-----------------|
| **MS365 MCP (Graph API)** | Page metadata, canvas layout, webpart types, section structure, image URLs | Cannot see custom JS-rendered content; only sees OOB webpart configurations | Combine with Playwright — Graph API for structure, Playwright for rendered output |
| **Playwright browser automation** | Live DOM structure, computed CSS, rendered visual output, JS-generated content | Requires authentication; Classic SP snapshots may be empty initially | Use persistent browser profile for auth reuse; wait 5s+ for JS rendering |
| **JavaScript evaluation in browser** | Exact colors, font stacks, element dimensions, scroll containers, image sizes | CSP restrictions may block some scripts | Run `page.evaluate()` which executes in page context, bypassing CSP for our own scripts |
| **Saved HTML analysis** | Offline analysis of page structure, CSS classes, image references | **Static snapshot; misses ALL dynamic content** — JS-rendered sections, authenticated resources, carousels, interactive state | **Replaced by live Playwright extraction** — captures the rendered DOM including all JS output. Saved HTML should only be used as a last-resort fallback when browser access is impossible |
| **Full-page screenshots** | Visual reference for pixel-perfect matching | Classic SP uses inner scroll containers (`#s4-workspace`), `window.scrollTo()` doesn't work | **Auto-detect scroll container** — the `sp-page-analyzer` tool checks for `#s4-workspace` (Classic) or `[data-automation-id="contentScrollRegion"]` (Modern) and scrolls the correct element |

### Recommended Approach: SP Page Analyzer Tool

We built `sp-page-analyzer/` — a brand-agnostic Playwright extraction pipeline that automates the entire analysis in one command:

```bash
node analyze.js https://contoso.sharepoint.com/sites/Intranet/Pages/home.aspx --auth
```

**6-step pipeline:**

| Step | What It Does | Output |
|------|-------------|--------|
| 1. Load & Auth | Navigate to page, handle login redirects, wait for JS rendering | Page fully rendered in browser |
| 2. Detect Type | Identify Classic vs Modern, find scroll container | `pageType`, `scrollContainer` |
| 3. Design Tokens | Sample 1000 elements for background colors, text colors, fonts, sizes (ranked by frequency) | Color palette, typography stack, page dimensions |
| 4. Page Structure | Map sections with dimensions, styles, child counts, webpart IDs | Section tree with coordinates |
| 5. Image Inventory | All `<img>` elements + CSS `background-image` values with natural/display dimensions | Complete asset list for replication |
| 6. Screenshots | Viewport, full-page, and scroll-segment captures (handles inner scroll containers) | Visual reference files |

**Why this replaces saved HTML:**
- Saved HTML files lose all JS-rendered content (carousels, dynamic widgets, Angular/React output)
- Saved HTML cannot capture computed styles (only source CSS)
- Saved HTML breaks authenticated image URLs
- The Playwright pipeline captures the **live rendered state** — exactly what a user sees

### Key Discoveries

1. **Classic vs Modern page detection**: Classic SP pages use `#DeltaPlaceHolderMain`, `#s4-workspace`, master pages, and AngularJS/jQuery. Modern SP pages use React, `CanvasZone`, `data-automation-id` attributes. The `sp-page-analyzer` auto-detects this.

2. **Classic SP scroll container**: The page body has `overflow: hidden`. The actual scroll container is `#s4-workspace`. Standard `window.scrollTo()` does nothing. Must target:
   ```js
   document.getElementById('s4-workspace').scrollTop = 2000;
   ```
   Modern SP uses `[data-automation-id="contentScrollRegion"]` as its scroll container.

3. **Design token extraction**: Use `getComputedStyle()` on sampled elements and rank by frequency. The most-used background color is likely the page background; the most-used text color is the body text. Cross-reference with custom CSS files for named values.

4. **Graph API for Modern pages**: The `canvasLayout.horizontalSections` array reveals every section, column, and webpart. Each webpart includes `webPartType` (GUID), `properties`, and `serverProcessedContent` (images, links, text). Use this for structural analysis of Modern pages, then Playwright for visual verification.

5. **Graph API limitation for Classic pages**: Graph API has no equivalent for Classic page content. Classic pages store content in wiki fields, Content Editor webparts, or Script Editor webparts — none of which are exposed via Graph API. **Playwright is the only reliable analysis tool for Classic pages.**

6. **Image asset discovery**: Intranet sites store images across multiple libraries (not just SiteAssets). The `sp-page-analyzer` finds all image sources automatically — both `<img>` elements and CSS `background-image` values — giving a complete inventory for replication.

7. **Authentication strategy**: Use Playwright's persistent browser context (`launchPersistentContext`) with a user data directory. First run: interactive login (`--auth`). Subsequent runs: reuses session cookies automatically. This avoids the complexity of certificate auth or token management for analysis-only workflows.

---

## 2. Rebuilding/Replicating the Page and Webparts

### Step 1: Assess the Source Page

Before writing any code, use the `sp-page-analyzer` output to categorize every section on the source page:

| Section Type | SP Modern Approach | Example |
|-------------|-------------------|---------|
| **Static banner** (image link, no interaction) | OOB Image webpart or single SPFx webpart | Announcement banners, footer images |
| **Rotating content** (carousel, auto-rotate) | Custom SPFx webpart with timer + state | Hero carousels, banner slideshows |
| **Interactive panel** (click-to-change, tabs) | Custom SPFx webpart with component state | News panels with sidebar selection |
| **Icon/link grid** (repeating items) | Custom SPFx webpart with configurable slots | Resource bars, quick links |
| **Custom header/footer** (site-wide chrome) | SPFx Application Customizer extension | Suite bar, navigation, welcome message |
| **Simple text/heading** | OOB Text webpart | Section dividers, headings |

### Step 2: Choose the Right Architecture

The number of custom SPFx webparts depends on the page complexity. Use this decision framework:

| Approach | When to Use | Trade-offs |
|----------|-------------|------------|
| **OOB webparts only** | Page has standard layouts (images, text, hero, news) | Fast to build, limited visual fidelity, no custom interaction |
| **Single monolithic SPFx webpart** | Quick demo or prototype — visual fidelity needed, editability not | Pixel-perfect, but hardcoded content; can't hand off to site owners |
| **Modular SPFx webparts** | Production use — site owners must edit content | Best balance of fidelity + editability; more files to maintain |

**Rule of thumb for modular architecture:**
- Group visually contiguous sections that share no distinct editing needs into one webpart
- Separate sections that have distinct interactive behavior (carousels, clickable panels)
- Don't create more than 5-6 webparts per page — too many makes page editing fragile
- Always use an Application Customizer extension for site-wide chrome (headers, footers, backgrounds)

### Step 3: Build with Editability

Every piece of content that a site owner might want to change must be exposed via the PropertyPane:

| Content Type | PropertyPane Control | Pattern |
|-------------|---------------------|---------|
| Text (titles, labels) | `PropertyPaneTextField` | Direct binding to prop |
| Image paths | `PropertyPaneTextField` | Relative filename; webpart prepends `baseUrl` |
| Enum choices (categories, intervals) | `PropertyPaneDropdown` | Predefined options array |
| Show/hide toggles | `PropertyPaneToggle` | Controls conditional rendering |
| Repeating items (slides, stories, resources) | Numbered slots (`slide1Title`, `slide2Title`...) | Build array at render time, skip empty slots |

**Multi-page PropertyPane**: When a webpart has more than ~10 configurable fields, split across multiple property pane pages (e.g., Page 1: Main Content, Page 2: Side Content, Page 3: Settings).

### Key Implementation Patterns

1. **baseUrl construction from SP context** — Never hardcode image paths. Build dynamically from the site URL:
   ```typescript
   const baseUrl = this.context.pageContext.web.absoluteUrl + '/' + this.properties.siteAssetsPath;
   ```
   This makes the same `.sppkg` work on any site — just upload images to the configured path.

2. **Canvas zone override mixin** — SharePoint Modern wraps webparts in constrained `CanvasZone` containers. For full-width layouts, override in a shared SCSS mixin:
   ```scss
   :global {
     .CanvasZone { max-width: none !important; padding: 0 !important; }
     .CanvasSection { max-width: none !important; padding: 0 !important; }
   }
   ```

3. **Shared SCSS design tokens** — Create a `_tokens.scss` partial with the client's extracted colors, typography, shadows, and layout constants. Share across all webparts via `@import`. Include reusable mixins for webpart root setup, content wrapper centering, and canvas zone overrides.

4. **Interactive state in class components** — SPFx uses React 17 class components (not hooks). Use `this.state` and `this.setState` for interactive features. Timer-based features (carousels) use `setInterval` in `componentDidMount` and clear in `componentWillUnmount`.

5. **User personalization from SP context** — Access the current user without additional API calls:
   ```typescript
   this.context.pageContext.user.displayName
   this.context.pageContext.user.email
   ```

6. **Data-attribute specificity boost** — For Application Customizers that inject CSS overriding SP's own styles, set a custom attribute on `<html>` (e.g., `html[data-custom-header]`) and prefix all CSS selectors with it. This gives high specificity without inline styles, plus clean activation/deactivation by adding/removing the attribute.

---

## 3. QA/QC the Page

### Build-Time Checks

| Check | Tool | What to Look For |
|-------|------|-----------------|
| TypeScript compilation | `gulp bundle --ship` | Unused variables (TS6133), const reassignment (TS2588), type errors |
| Linting | SPFx built-in (ESLint or tslint for older projects) | `no-var`, `no-explicit-any`, unused imports |
| SCSS compilation | SPFx sass task | Missing imports, undefined variables, invalid selectors |
| Package generation | `gulp package-solution --ship` | Manifest validation, bundle size (aim for <200KB per webpart) |

**Note:** Older SPFx projects (v1.15 and below) use `tslint`; newer versions use `eslint`. Check `config/tslint.json` vs `eslint.config.*` to know which applies.

### Common TS/Lint Issues & Fixes

1. **TS6133 — unused variables**: Methods referenced via `.bind(this)` in JSX are not detected as "used" by TypeScript. **Preferred fix:** use arrow functions as class properties (e.g., `private _onClick = (): void => { ... }`), which preserve `this` binding and are correctly detected as used. **Quick fix:** make them `public`, but this leaks internal API.

2. **no-var lint rule**: SPFx ESLint config enforces `const/let` over `var`. AI agents often generate `var` for ES5 compatibility — use `const/let` instead; TypeScript handles the ES5 downcompilation.

3. **const reassignment**: Bulk replacing `var` → `const` breaks variables that are reassigned (in switch statements, if/else blocks). Variables that are reassigned after declaration must use `let`.

4. **Unused imports**: `import * as strings from '...'` — if localization strings aren't used in the file, remove the import entirely. ESLint disable comments don't suppress TypeScript compiler errors (TS6133).

### Runtime Checks

| Check | How | What to Verify |
|-------|-----|---------------|
| **Console errors** | Browser DevTools → Console | No 404 images, no CSP violations, no unhandled exceptions |
| **Visual fidelity** | Playwright screenshots or manual comparison | Demo page matches source page side-by-side |
| **Property pane** | Edit mode → click each webpart → open property pane | Every editable field updates the webpart live; no fields silently fail |
| **Responsive/mobile** | DevTools responsive mode (375px, 768px, 1024px) | Layout adapts; no horizontal scroll; images scale; text readable |
| **Accessibility** | Keyboard-only navigation + screen reader spot check | Tab order logical; interactive elements focusable; alt text on images |
| **Cross-browser** | Test in Edge + Chrome (corporate SP users) | No layout breaks or missing features between browsers |
| **Page structure** | Graph API `canvasLayout` read post-deployment | Webpart types and section structure match expectations |
| **Live data** | Check any webparts that fetch external data | API calls succeed; fallback/error states render gracefully |

### Automated QA (Optional)

The `sp-page-analyzer` tool can be extended for before/after comparison:
1. Run against the **source page** → save baseline output
2. Run against the **replica page** → save comparison output
3. Diff design tokens, section counts, and image inventories

Microsoft's **SP Page Diagnostics Tool** (browser extension) can also identify slow webparts, oversized images, and excessive API calls on the deployed page.

---

## 4. Challenges/Blockers and How They Were Resolved

### Challenge 1: Node.js Version Incompatibility
- **Problem**: System Node v23.11.0, SPFx requires >=18.17.1 <19.0.0
- **Resolution**: Use `nvm use 18` before every build command
- **Prevention**: Add `.nvmrc` file with `18` to the project root

### Challenge 2: Classic SP Page Analysis via Browser
- **Problem**: Playwright snapshot was empty on first load (JS-rendered content)
- **Current fix**: Hardcoded `waitForTimeout(5000)` after page load
- **Best practice**: Replace with `waitForSelector()` targeting a known rendered element — Classic: `#DeltaPlaceHolderMain .ms-webpart-zone`, Modern: `[data-automation-id="CanvasZone"]`. This waits exactly as long as needed (1s on fast networks, 15s on slow) instead of a fixed delay. Similarly, replace `waitForTimeout` between scroll segments with `waitForFunction()` that confirms scroll position changed. **Planned:** apply these smart waits when wrapping the analyzer as an MCP server.
- **Problem**: Page didn't scroll with `window.scrollTo()`
- **Resolution**: Auto-detect scroll container — Classic: `#s4-workspace`, Modern: `[data-automation-id="contentScrollRegion"]`. The `sp-page-analyzer` handles this automatically.

### Challenge 3: SharePoint Modern Chrome Hiding

> **⚠️ RESOLUTION HIERARCHY — always follow this order:**
> 1. **Out-of-box** — native SP settings (no code, supported, survives updates)
> 2. **Custom webparts** — SPFx components (supported by Microsoft, maintainable)
> 3. **Hacks** — DOM manipulation, CSS injection, MutationObserver (unsupported — **must warn stakeholders this is not a true build and may break on any SP update**)

**Step 1 — Configure native site settings (REQUIRED FIRST):**

| Setting | What It Controls | How to Apply |
|---------|-----------------|-------------|
| Header layout → **Minimal** or **Compact** | Shrinks/simplifies the site header | Site Settings → Change the look → Header |
| Hide site title | Removes title text from header bar | Change the look → Header → toggle title off |
| Disable footer | Removes the site footer entirely | PowerShell: `Set-PnPWeb -FooterEnabled $false` or Site Settings → Change the look → Footer |
| Hide navigation | Removes left/top navigation panel | PowerShell: `Set-PnPWeb -QuickLaunchEnabled $false` |
| Set page as **Home page** | Hides article title area and page header chrome | Site Pages → right-click page → Make homepage |
| Full-width section layout | Removes column width constraints per section | Page editor → Add section → Full-width |

**Step 2 — Code-based overrides (ONLY for remaining gaps):**

After native settings are applied, some elements may still need hiding (e.g., the suite bar "waffle" menu, "Edit" button, specific SP-injected banners). For these:
- **Problem**: SP Modern pages continuously re-render (React), resetting inline styles
- **Resolution**: MutationObserver watches for attribute/childList changes and reapplies styles using `requestAnimationFrame` for batching
- **⚠️ Warning**: DOM manipulation targeting native SP elements is **unsupported by Microsoft**. Element IDs and class names can change in any update without notice. This technique is appropriate for **demos and controlled environments only** — not production sites managed by end users.

**Document what was configured:** Always record which native settings were changed and which elements required code overrides, so site owners can replicate the setup on new sites without the developer.

### Challenge 4: Canvas Zone Width Constraints

Following the resolution hierarchy:

1. **Out-of-box**: Use **full-width section layout** in the page editor (Page → Add section → Full-width). This is the native way to allow content to span the full page width.
2. **Custom webparts**: Set `supportedHosts: ["SharePointWebPart", "SharePointFullPage"]` in the webpart manifest. Design the webpart's internal layout to fill available width using standard CSS (flexbox/grid).
3. **Hack (if still constrained)**: Even with full-width sections, SP wraps webpart content in max-width `CanvasZone`/`CanvasSection` containers. Override with `:global` SCSS and `!important`:
   ```scss
   :global {
     .CanvasZone { max-width: none !important; padding: 0 !important; }
     .CanvasSection { max-width: none !important; padding: 0 !important; }
   }
   ```
   **⚠️ This is a hack.** These class names are internal SP implementation details. Microsoft can rename or restructure them at any time. Warn stakeholders that this is not a true build — it's a visual override that may break on future SP updates.

### Challenge 5: Wrong Page Type Creates Unnecessary Work

- **Problem**: Different SharePoint page types (home, article, news post, wiki) display different amounts of native chrome (title area, author/date, navigation, metadata). Choosing the wrong page type means writing hack code to hide chrome that the correct type would never show.
- **Resolution**: **Assess the page requirement before creating anything.** Determine:
  1. What is the page's purpose? (landing page, content article, news, dashboard)
  2. What native chrome should be visible vs hidden?
  3. Which SP page type gives the closest native starting point?

| Page Type | Native Chrome Shown | Best For |
|-----------|-------------------|----------|
| **Home page** | Minimal — no title area, no author/date | Landing pages, dashboards, portals |
| **Article page** | Title area, author, date, banner image | News articles, blog-style content |
| **News post** | Title, author, date, publishing metadata | Announcements, company updates |
| **Wiki page** (Classic only) | Editable content zones, sidebar | Legacy knowledge bases |

**Why this matters:** Picking the right page type is a Tier 1 (out-of-box) decision that can eliminate entire categories of Tier 3 hacks. Always make this decision during requirements gathering, not after code is written.

### Challenge 6: Image Path Resolution

**Custom webpart resolution (Tier 2):** Images must be uploaded to a SharePoint document library (typically SiteAssets) — local filesystem paths don't work in SP.

- Every webpart exposes a configurable `siteAssetsPath` property via PropertyPane
- `baseUrl` is constructed dynamically: `context.pageContext.web.absoluteUrl + '/' + siteAssetsPath`
- This makes the same `.sppkg` portable across sites — just upload images and set the path
- **Pre-deployment check:** Verify all referenced images exist in the target library before deploying. Missing images (e.g., `Vision2030-logo.png`) cause silent 404s that only show in the browser console

### Challenge 7: Parallel Agent Code Quality
- **Problem**: Background agents generated code with `var` declarations, unused variables, and inconsistent patterns
- **Resolution**: Two-part approach:
  1. **Upfront:** Include SPFx linting rules in agent prompts — specify `const/let` (no `var`), arrow functions for class methods, and `import` cleanup requirements
  2. **Post-generation:** Run `gulp bundle --ship` as a validation gate before packaging. Fix any remaining lint/TS errors before proceeding to `gulp package-solution --ship`
- **Prevention for future projects:** Create an SPFx code generation checklist (or Claude Code skill) that agents must follow, covering the common pitfalls documented in Section 3

---

## 5. Validated Learnings & Tips/Tricks

### From Online Research (Validated March 2026)

**SharePoint Pages Graph API (now GA)**
- `GET /sites/{site-id}/pages/{page-id}/microsoft.graph.sitePage/canvasLayout` returns full section/column/webpart structure
- `GET /sites/{site-id}/pages/{page-id}/microsoft.graph.sitePage/webparts` lists all webparts on a page
- Site creation via Graph was added November 2025
- Source: [Microsoft Learn - canvasLayout Resource](https://learn.microsoft.com/en-us/graph/api/resources/canvaslayout)

**PropertyPane Best Practices**
- Use PropertyPane for **configuration**, not content editing. Content should be edited inline in the webpart canvas (the "directness" principle)
- Use `@pnp/spfx-property-controls` for ready-made controls: `PropertyFieldCollectionData` (array editing), `PropertyFieldListPicker`, `PropertyFieldPeoplePicker`, `PropertyFieldColorPicker`
- Use `loadPropertyPaneResources()` to lazy-load heavy dependencies only when the pane opens
- Populate dropdowns dynamically by fetching data during `onInit()`
- Source: [PnP Property Controls](https://pnp.github.io/sp-dev-fx-property-controls/)

**Hiding Native SP Chrome (Warning)**
- DOM manipulation targeting native SharePoint elements is **unsupported and fragile** per Microsoft
- Microsoft can change element IDs/classes in any update without notice
- Source: [Microsoft Learn - Page Placeholders](https://learn.microsoft.com/en-us/sharepoint/dev/spfx/extensions/get-started/using-page-placeholder-with-extensions)

**Classic → Modern Migration**
- Master pages, page layouts, Script Editor, Content Editor webparts have **no modern equivalent**
- Must be rebuilt as SPFx webparts from scratch
- SharePoint Designer workflows → Power Automate
- PnP Page Transformation Framework (`ConvertTo-PnPPage`) can auto-convert some classic webparts
- "Classic Publishing Sites Must Be Modernized, Not Migrated" — don't attempt lift-and-shift
- Source: [Microsoft Learn - Modernize Classic Sites](https://learn.microsoft.com/en-us/sharepoint/dev/transform/modernize-classic-sites)

**QA Tools**
- **Page Diagnostics Tool**: Microsoft browser extension that identifies slow webparts, large images, excessive API calls
- **Percy/Applitools/Chromatic**: Visual regression testing tools usable with Playwright + SharePoint
- Source: [Microsoft Learn - Page Diagnostics](https://learn.microsoft.com/en-us/microsoft-365/enterprise/page-diagnostics-for-spo)

**SPFx 1.21 — Flexible Sections (March 2025)**
- Webparts can be dragged freely and resized without strict column constraints
- May reduce or eliminate the need for CanvasZone CSS overrides (see Challenge 4)
- Source: Microsoft 365 Developer Blog, March 2025

**SPFx 1.22 — Gulp Replaced by Heft (December 2025)**
- Gulp is no longer the build task runner — replaced by Heft (from Rush Stack)
- Webpack remains the bundler, orchestrated by Heft instead of Gulp
- All npm audit vulnerabilities resolved in scaffolded projects
- Existing projects on Gulp still work temporarily but should plan migration
- ⚠️ Projects on SPFx ≤1.20 still use `gulp bundle --ship` / `gulp package-solution --ship`
- Source: Microsoft Learn - SPFx v1.22 Release Notes

**SPFx 1.23 — Open-Source Templates & CLI Preview (March 2026)**
- Open-sourced SPFx solution templates
- New SPFx CLI (preview) to eventually replace the Yeoman generator
- Source: Microsoft 365 Developer Blog, February 2026

**Field Customizers Retirement**
- Retirement announced June 2025, then postponed after community feedback
- Still functional but future uncertain — consider Column Formatting as alternative
- Source: Voitanos blog

**Azure ACS Retirement — April 2, 2026**
- Azure ACS (legacy add-in auth) retired **April 2, 2026**
- All authentication must use Microsoft Entra ID
- Any existing SharePoint Add-ins must be migrated to SPFx
- Source: Microsoft Learn

### Tips for SharePoint SPFx Development

1. **Use `supportedHosts: ["SharePointWebPart", "SharePointFullPage"]`** for full-page webparts that need to break out of column constraints.

2. **Use `supportsThemeVariants: true`** in manifest to receive theme change events and adapt to site themes.

3. **Use `@pnp/spfx-property-controls`** instead of building custom property pane controls from scratch — saves significant development time.

4. **Use leaf-level imports from Fluent UI**: `import { PrimaryButton } from "@fluentui/react/lib/Button"` instead of importing the whole package — reduces bundle size.

5. **Use SPFx Library Components** for shared code across webparts — "write once, use everywhere" pattern.

6. **Navigation loading from SP list**: Create a `HeaderNavigation` list with Title, LinkURL, SortOrder, IsActive columns. Load via REST API in `onInit()` with fallback defaults.

7. **Use `data-attribute` specificity boost**: Set a custom attribute on `<html>` (e.g., `html[data-custom-bg]`) and prefix all CSS selectors with it — gives extremely high specificity without inline styles, plus clean activation/deactivation by adding/removing the attribute.

8. **Externalize large libraries to CDN**: Use the `externals` config in `config/config.json` to exclude heavy dependencies from your bundle and load them from CDN instead. This can reduce bundle size dramatically.

9. **Use dynamic imports for lazy loading**: Use `import()` and `React.lazy` with `Suspense` to lazy-load components not needed at first render (charts, editors, heavy UI). Split code into multiple bundles loaded on demand.

10. **Run webpack Bundle Analyzer**: Add `--stats` to your build command and use `webpack-bundle-analyzer` to visualize dependency sizes and identify bloat. Aim for <200KB per webpart bundle.

11. **Lock package-lock.json in git**: Always commit `package-lock.json`. Without it, `npm install` pulls latest compatible versions that may break your build silently.

12. **Use tenant-wide deployment**: Set `skipFeatureDeployment: true` in `package-solution.json` to make webparts available across all sites immediately upon App Catalog deployment — no per-site activation needed.

13. **Use CLI for Microsoft 365 for CI/CD**: The `m365` CLI (Node.js, cross-platform) is ideal for automating deployments in GitHub Actions or Azure DevOps. Commands: `m365 spo app add`, `m365 spo app deploy`. Replaces manual App Catalog uploads.

14. **Use PnPjs for API calls**: `@pnp/sp` and `@pnp/graph` provide a fluent API for SharePoint REST and Microsoft Graph calls with built-in caching, batching, and error handling. Significantly cleaner than raw `fetch()` calls.

15. **Use PnP React Controls**: `@pnp/spfx-controls-react` provides SharePoint-aware React components — `PeoplePicker`, `ListView`, `Carousel`, `DateTimePicker`, `ChartControl` — that save significant development time over building from scratch.

16. **Use React Testing Library, not Enzyme**: Enzyme is deprecated. Use Jest (or Vitest for faster execution) with React Testing Library for component tests. Mock SPFx-specific modules (`@microsoft/sp-core-library`, etc.).

17. **Separate dev/staging/production with site collection App Catalogs**: Use site collection app catalogs for testing before promoting to the tenant-wide app catalog. This prevents untested webparts from appearing across the organization.

18. **Use the SPFx upgrade tool**: Run `m365 spfx project upgrade` (from CLI for Microsoft 365) to automatically generate upgrade steps when moving between SPFx versions. Saves hours of manual migration work.

---

## 6. Overall Learnings

### What Worked Well

1. **Multi-tool analysis approach**: Combining Graph API (structure), Playwright (visual), and JS evaluation (CSS extraction) gave a complete picture of the real intranet page that no single tool could provide alone.

2. **Modular webpart architecture**: Breaking into 4 webparts hit the sweet spot — enough granularity for editability, not so many that page editing becomes fragile.

3. **Shared design tokens**: A single `_tokens.scss` with reusable mixins (`webpart-root`, `content-wrapper`, `canvas-override`) eliminated duplication across all webpart SCSS files.

4. **Parallel agent construction**: Building all 4 webparts simultaneously via background agents reduced total build time from ~30min to ~10min.

5. **Existing pattern reuse**: Following the exact patterns from `HeroCarouselWebPart.ts` (lifecycle, property pane, theme handling) ensured consistency and reduced errors.

### What Could Be Improved

1. **Agent code quality**: AI agents generating SPFx code must be given the project's linting rules upfront (`const/let` not `var`, arrow functions for class methods, no unused imports). Post-generation, always run the build as a validation gate before packaging.

2. **Live data verification**: Any webpart that fetches external data (APIs, RSS, live feeds) should be tested with real endpoints on the deployed page — not just with hardcoded demo data. Verify error states and fallbacks.

3. **Interactive state testing**: Webparts with user interaction (carousels, tab panels, click-to-change) must be tested on the actual deployed SharePoint page, not just in local workbench. SP's React rendering can interfere with component state.

4. **Page type selection**: Always determine the correct page type (home, article, news post) during requirements gathering — before writing code. Wrong page type creates unnecessary chrome-hiding work (see Challenge 5).

5. **Image asset pre-check**: Verify all referenced images exist in the target document library before deploying. Missing images cause silent 404s visible only in the browser console — easily missed during QA.

### Key Architectural Insight

> **Classic SP pages are essentially custom web applications** hosted inside SharePoint. They use their own JS frameworks (Angular, jQuery), CSS, and render everything client-side. Replicating them in Modern SP requires understanding that Modern SP is also a React SPA — but one that constrains you to the SPFx webpart model. The art is in making SPFx webparts that break free of the constraints (CanvasZone width, section gaps, theme overrides) while still playing nice with the SP editing experience.

---

## 7. SPFx Platform Landscape (March 2026)

This section is a self-contained reference for anyone starting or maintaining SPFx development. It covers the current framework state, version roadmap, key libraries, common pitfalls, performance techniques, AI tooling, and a generic delivery checklist.

### 7.1 SPFx Version Roadmap

| Version | Release | Key Changes |
|---------|---------|-------------|
| **1.20** | Mid-2025 | Current stable for most production sites. Gulp-based build (`gulp bundle --ship` / `gulp package-solution --ship`). Node.js 18.x required. |
| **1.21** | March 2025 | **Flexible sections** — webparts can be dragged freely and resized without strict column constraints. May reduce or eliminate CanvasZone CSS override hacks (see Section 4, Challenge 4). |
| **1.22** | December 2025 | **Build toolchain overhaul** — Gulp replaced by Heft (Rush Stack). Webpack remains the bundler but is now orchestrated by Heft. All npm audit vulnerabilities resolved in scaffolded projects. Existing Gulp projects still work temporarily but should plan migration. |
| **1.23** | March 2026 (preview) | **Open-sourced solution templates** on GitHub. New **SPFx CLI** (preview) to eventually replace the Yeoman generator. |

**Release cadence**: Microsoft publishes quarterly SPFx releases with monthly roadmap updates on the Microsoft 365 Developer Blog.

### 7.2 Retirements & Breaking Changes

| Technology | Status | Migration Path |
|-----------|--------|---------------|
| **Azure ACS** (legacy add-in auth) | **Retired April 2, 2026** | All auth must use Microsoft Entra ID. Migrate existing SharePoint Add-ins to SPFx. |
| **SharePoint 2013 Workflows** | Retired 2025 | Migrate to Power Automate |
| **SharePoint Add-ins** | Retired 2025 | Rebuild as SPFx solutions |
| **SPFx Isolated Web Parts** | Retired 2025 | Use standard SPFx webparts with Entra ID auth |
| **SharePoint Mail API** | Retired 2025 | Use Microsoft Graph Mail API |
| **Field Customizers** | Retirement announced June 2025, **postponed** after community feedback | Column Formatting recommended as alternative. Still functional but future uncertain. |
| **Gulp build system** | Superseded in SPFx 1.22 | Migrate to Heft. SPFx ≤1.20 projects still use Gulp. |

### 7.3 Key Libraries & Accelerators

| Library | Purpose | When to Use |
|---------|---------|-------------|
| **`@pnp/spfx-controls-react`** | SharePoint-aware React components — PeoplePicker, ListView, Carousel, DateTimePicker, ChartControl | Anytime you'd build a SP-specific UI control from scratch. Saves significant dev time. |
| **`@pnp/spfx-property-controls`** | Ready-made PropertyPane controls — CollectionData (array editing), ListPicker, PeoplePicker, ColorPicker | Complex configuration UIs. Replaces hundreds of lines of custom PropertyPane code. |
| **`@pnp/sp` / `@pnp/graph`** (PnPjs) | Fluent API for SharePoint REST and Microsoft Graph with built-in caching, batching, error handling | Recommended over raw `fetch()` for all SP/Graph API calls. |
| **Fluent UI** (`@fluentui/react`) | Microsoft's design system (buttons, dialogs, icons, etc.) | Already bundled with SPFx. Use leaf-level imports to control bundle size. |
| **CLI for Microsoft 365** (`m365`) | Cross-platform CLI for SP/Graph operations, CI/CD automation, project upgrades | Deployment automation, `m365 spfx project upgrade`, tenant management. |
| **SPFx Library Components** | Shared code across webparts (write once, use everywhere) | When 2+ webparts share logic, utilities, or services. |

**PnP controls vs Fluent UI**: PnP controls are accelerators layered on Fluent UI that solve data-bound scenarios specific to SharePoint (e.g., PeoplePicker wraps Fluent UI but adds tenant-specific search, validation, and user retrieval). Use Fluent UI directly for generic UI; use PnP controls when the component needs SharePoint context.

### 7.4 Common Pitfalls

1. **SPFx/Node version mismatch** — Accounts for ~90% of setup failures. SPFx version must match SharePoint version. Node.js must be 18.x. Add `.nvmrc` to every project.

2. **Unlocked dependencies** — Without `package-lock.json` committed, `npm install` pulls latest compatible versions that silently break builds. Always commit the lock file.

3. **Overly broad API permissions** — Graph permissions granted hastily during development (e.g., `Sites.ReadWrite.All`) often ship to production unrestricted. Audit and scope down before deployment.

4. **Code duplication across webparts** — Copy-pasting between webparts causes inconsistent bug fixes, bloated bundles, and merge conflicts. Use SPFx Library Components for shared code.

5. **`var` declarations** — SPFx ESLint config enforces `const/let`. AI code generators often produce `var` for ES5 compatibility — TypeScript handles the downcompilation, so always use `const/let`.

6. **Staying on outdated SPFx versions** — Loses security patches for Node/Gulp dependencies and risks build breaks from OS or npm updates. Use `m365 spfx project upgrade` to generate migration steps.

7. **Hardcoded image paths** — Images must be in a SharePoint document library. Construct paths from `context.pageContext.web.absoluteUrl + '/' + siteAssetsPath`. Never hardcode URLs.

### 7.5 Performance Optimization

| Technique | Impact | How |
|-----------|--------|-----|
| **Release builds** | ~85% size reduction (e.g., 1255KB debug → 177KB release) | Always use `--ship` flag: `gulp bundle --ship` |
| **CDN externals** | Removes large libraries from bundle entirely | Configure `externals` in `config/config.json` to load from CDN |
| **Dynamic imports** | Defers heavy components to load on demand | `import()` + `React.lazy` + `Suspense` for charts, editors, rich UI |
| **Leaf-level Fluent imports** | Avoids importing the entire Fluent UI package | `import { PrimaryButton } from "@fluentui/react/lib/Button"` |
| **Tree shaking** | Eliminates dead code paths automatically | Enabled by default in Webpack production mode |
| **Bundle Analyzer** | Visual treemap of what's consuming space | Add `--stats` flag, use `webpack-bundle-analyzer` |
| **Cache-first with service workers** | Accelerates CSS/asset loading for repeat visits | Cache frequently accessed resources first |

**Target**: <200KB per webpart bundle in production.

### 7.6 AI Tooling for SharePoint Development

| Category | Exists? | Best Option | Notes |
|----------|---------|-------------|-------|
| **Claude Code skill for SPFx** | 1 community package | `@sudharsank/spfx-enterprise-skills` | npm-installable skill pack; distributes SKILL.md files with SPFx patterns. Install via `npx @sudharsank/spfx-enterprise-skills --list-skills` |
| **MCP server for SP page analysis** | **Yes (ours)** | `sp-page-analyzer` MCP server | Playwright-based extraction of design tokens, structure, images, screenshots. 4 tools: `analyze_page`, `extract_design_tokens`, `capture_screenshots`, `list_images` |
| **MCP server for SP data access** | 8-10+ | sekops-ch, Composio, built-in ms365 MCP | Read/write documents, lists, sites via Graph API |
| **MCP server for SPFx code gen** | **No** | None | All SharePoint MCP servers focus on data access, not code generation |
| **Copilot extension for SPFx** | Mature | SPFx Toolkit (`pnp/vscode-viva`) | `@spfx` chat participant, `/new` scaffolding, agent mode for site operations. Most capable AI tool for SPFx, but tied to VS Code + GitHub Copilot |
| **AI tool that generates full SPFx webparts** | **No** | Does not exist | No dedicated tool auto-generates complete webparts from prompts |

Sources: [spknowledge.com](https://spknowledge.com/2026/03/18/from-skills-repo-to-npx-installer/), [Composio](https://composio.dev/toolkits/share_point/framework/claude-code), [SPFx Toolkit](https://pnp.github.io/vscode-viva/), [PnP Blog](https://pnp.github.io/blog/post/spfx-toolkit-vscode-chat-pre-release/)

### 7.7 Build & Deploy Reference

**Build pipeline (SPFx ≤1.20, Gulp):**
```bash
nvm use 18                           # Ensure correct Node version
npm install                          # Install dependencies (lock file must exist)
gulp bundle --ship                   # Production bundle (minified, tree-shaken)
gulp package-solution --ship         # Generate .sppkg package
```

**Build pipeline (SPFx ≥1.22, Heft):**
```bash
nvm use 18
npm install
npx heft build --production          # Replaces gulp bundle --ship
# Package generation — refer to SPFx 1.22 docs for updated commands
```

**CI/CD (GitHub Actions with CLI for Microsoft 365):**
```yaml
- uses: actions/setup-node@v4
  with: { node-version: 18 }
- run: npm ci
- run: npx gulp bundle --ship
- run: npx gulp package-solution --ship
- run: |
    npx m365 login --authType certificate --certificateBase64Encoded ${{ secrets.M365_CERT }} --appId ${{ secrets.M365_APP_ID }} --tenant ${{ secrets.M365_TENANT_ID }}
    npx m365 spo app add --filePath sharepoint/solution/*.sppkg --appCatalogUrl ${{ secrets.SP_APP_CATALOG_URL }} --overwrite
    npx m365 spo app deploy --name *.sppkg --appCatalogUrl ${{ secrets.SP_APP_CATALOG_URL }}
```

**Tenant-wide deployment**: Set `"skipFeatureDeployment": true` in `config/package-solution.json` to make webparts available on all sites immediately — no per-site activation needed.

### 7.8 Generic Delivery Checklist

This checklist applies to any SharePoint page build or replication project, regardless of client or brand.

**Planning:**
- [ ] Gather requirements: page purpose, page type, content sections, editing needs
- [ ] Analyse source (if replicating): run `sp-page-analyzer` or manual inspection
- [ ] Extract design tokens: colors, fonts, sizes, spacing
- [ ] Map sections to architecture: OOB webparts vs custom SPFx vs extensions
- [ ] Apply resolution hierarchy: (1) out-of-box → (2) custom webparts → (3) hacks with warning
- [ ] Choose correct page type (home, article, news post) — see Section 4, Challenge 5

**Build:**
- [ ] Build with editability: expose all changeable content via PropertyPane
- [ ] Use shared design tokens (`_tokens.scss`) across all webparts
- [ ] Verify all webpart/extension manifests registered in `config/config.json`
- [ ] Run `gulp bundle --ship` — fix all TS/lint errors before packaging
- [ ] Run `gulp package-solution --ship` — verify bundle size (<200KB per webpart)

**Deploy:**
- [ ] Configure site settings: header layout, footer, navigation, page type
- [ ] Upload images to target document library
- [ ] Verify all referenced images exist (no silent 404s)
- [ ] Deploy `.sppkg` to App Catalog (automated via CI/CD or manual upload)
- [ ] Assemble page with webparts in correct order

**QA:**
- [ ] Console errors: no 404 images, no CSP violations, no unhandled exceptions
- [ ] Visual fidelity: demo page matches source/design side-by-side
- [ ] Property pane: every editable field updates the webpart live
- [ ] Responsive: 375px, 768px, 1024px — no horizontal scroll, images scale, text readable
- [ ] Accessibility: keyboard navigation, focusable interactives, alt text on images
- [ ] Cross-browser: Edge + Chrome (corporate SP users)
- [ ] Live data: API calls succeed, fallback/error states render gracefully
