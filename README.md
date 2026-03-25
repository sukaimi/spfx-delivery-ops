# SPFx Delivery Ops

**Three tools that give AI agents everything they need to build and ship SharePoint sites.**

One writes the code. One ships the project. One reads the page.

---

## Why This Exists

AI-assisted SPFx development needs three things to work end-to-end:

1. **Good code** -- [`@sudharsank/spfx-enterprise-skills`](https://github.com/nickhsm/spfx-enterprise-skills) teaches Claude how to write clean SPFx code: architecture, theming, accessibility, performance.
2. **A delivery workflow** -- This skill teaches Claude how to ship it: which sections need custom code vs native webparts, how to handle SharePoint's quirks, how to deploy via CI/CD, and how to QA the result.
3. **A starting point** -- The MCP analyzer reads a live SharePoint page and extracts what's already there (colors, fonts, images, layout) so both skills have something to work from.

Without all three, you get correct code that never ships, or shipped code that doesn't match the page. This repo provides pieces 2 and 3, designed to work alongside piece 1.

---

## What's Included

### 1. The Skill -- `spfx-delivery-ops`

A Claude Code skill (`SKILL.md`) that covers the full delivery lifecycle: from analyzing a page, through build decisions, to deployment and QA.

What it teaches Claude:

- **OOB vs custom vs DOM hack** -- A decision framework for every page section. Use native webparts when you can, custom SPFx when you need interactivity, and Application Customizer overrides only as a last resort.
- **Canvas zone overrides** -- The `:global` SCSS patterns for full-width layouts
- **Multi-page PropertyPane** -- Patterns for webparts with 20+ configurable fields
- **CI/CD deployment** -- Certificate auth with CLI for Microsoft 365, App Catalog automation
- **Image deployment** -- Dynamic `baseUrl` from SP context, no hardcoded paths
- **ES5 gotchas** -- What breaks when AI generates modern JS for SPFx's ES5 target
- **Full delivery checklist** -- Discovery, build, deploy, QA steps

**Install:**

```bash
# Copy SKILL.md into your SPFx project's Claude Code skills directory
mkdir -p .claude/skills
cp spfx-delivery-ops/SKILL.md .claude/skills/spfx-delivery-ops.md
```

Or reference it in your `.claude/settings.json`:

```json
{
  "skills": ["./spfx-delivery-ops/SKILL.md"]
}
```

### 2. The Analyzer -- `sp-page-analyzer/`

A Playwright-based tool that reads any live SharePoint page (Classic or Modern) and extracts structured data. Also runs as an MCP server so Claude can call it directly.

This is how Claude sees the starting point before either skill does its work.

**What it extracts:**

| Output | Description |
|--------|-------------|
| Design tokens | Background colors, text colors, font families, font sizes (ranked by frequency) |
| Page structure | Sections with dimensions, styles, child counts, webpart IDs |
| Image inventory | All `<img>` sources + CSS background images with natural/display dimensions |
| Screenshots | Viewport, full-page, and scroll segments (handles SP's inner scroll containers) |
| Navigation | Nav links with text and URLs |
| Custom assets | CSS/JS files loaded from Style Library or SiteAssets |
| Page type | Classic (master pages, `#s4-workspace`) vs Modern (React, `CanvasZone`) |

**Quick start:**

```bash
cd sp-page-analyzer
npm install

# Analyze any SharePoint page
node analyze.js https://contoso.sharepoint.com/sites/Intranet --auth

# First run opens a browser for login. Subsequent runs reuse the session.
```

**Output:**

```
output/
├── analysis.json            # Full structured analysis
├── screenshot-viewport.png  # Above the fold
├── screenshot-full.png      # Full page
└── screenshot-scroll-N.png  # Scroll segments (Classic SP)
```

**As an MCP server** (for Claude Code / Claude Desktop):

```json
{
  "mcpServers": {
    "sp-page-analyzer": {
      "command": "node",
      "args": ["./sp-page-analyzer/mcp-server.js"]
    }
  }
}
```

This gives Claude Code direct access to `analyze_page`, `extract_design_tokens`, `capture_screenshots`, and `list_images` tools.

### 3. The Field Manual -- `METHODOLOGY.md`

Everything learned from building SharePoint Modern sites with AI assistance, in 7 sections:

1. **Analysing a SharePoint page** -- Techniques, tool comparison, key discoveries
2. **Rebuilding/replicating pages** -- Architecture decisions, PropertyPane patterns, implementation
3. **QA/QC** -- Build-time checks, visual comparison, common TS/lint fixes
4. **CI/CD deployment** -- CLI for M365, certificate auth, GitHub Actions pipeline
5. **Canvas zone and layout control** -- Full-width overrides, section mapping
6. **Application Customizer patterns** -- Header/footer injection, background persistence, MutationObserver
7. **Lessons learned** -- What worked, what didn't, what to do differently

Not a tutorial -- a field manual from an actual delivery.

---

## The Trio

These three pieces are designed to work together:

```
.claude/skills/
├── spfx-delivery-ops.md              # Ships the project (this repo)
└── spfx-enterprise-skills/           # Writes the code (@sudharsank)
    ├── spfx-enterprise-code-and-performance/SKILL.md
    ├── spfx-theme-and-brand-integration/SKILL.md
    └── ...

.mcp.json
└── sp-page-analyzer                  # Reads the page (this repo)
```

| Piece | Role | Source |
|-------|------|--------|
| `@sudharsank/spfx-enterprise-skills` | Writes clean SPFx code | [github.com/sudharsank](https://github.com/nickhsm/spfx-enterprise-skills) |
| `spfx-delivery-ops` skill | Ships the project end-to-end | This repo |
| `sp-page-analyzer` MCP | Reads the live page as a starting point | This repo |

Coding standards + delivery workflow + page analysis = full coverage.

---

## Screenshots

Visual QA output from the delivery workflow:

| Screenshot | Description |
|------------|-------------|
| `qa-screenshots/desktop-fullpage.png` | Full-page desktop render after deployment |
| `qa-screenshots/desktop-above-fold.png` | Above-the-fold viewport capture |
| `qa-screenshots/mobile-375x812.png` | Mobile responsive check |
| `qa-screenshots/tablet-768x1024.png` | Tablet breakpoint check |
| `qa-screenshots/banner-carousel-propertypane-p1.png` | PropertyPane editability verification |
| `qa-screenshots/hero-stories-propertypane-p1.png` | Multi-page PropertyPane in action |

---

## Quick Start

### Prerequisites

- Node.js 18.x (SPFx 1.20 requirement)
- An SPFx project (existing or new)
- Claude Code, Cursor, or any AI coding assistant that supports skills/instructions

### Option A: Use the skill only

Copy `spfx-delivery-ops/SKILL.md` into your project's AI skill directory. Your AI assistant now understands SPFx delivery workflow.

### Option B: Use the page analyzer

```bash
git clone https://github.com/cac-io/spfx-delivery-ops.git
cd spfx-delivery-ops/sp-page-analyzer
npm install
node analyze.js https://your-sharepoint-site.sharepoint.com/sites/YourSite --auth
```

Use the `analysis.json` output to inform your SPFx architecture decisions.

### Option C: Full toolkit

1. Clone the repo
2. Copy the skill into your SPFx project
3. Configure the MCP server for Claude Code
4. Run the analyzer against your target page
5. Let your AI assistant use both the analysis output and the delivery skill to build and ship

---

## Project Structure

```
spfx-delivery-ops/
├── spfx-delivery-ops/
│   └── SKILL.md                  # Claude Code delivery workflow skill
├── sp-page-analyzer/
│   ├── analyze.js                # CLI analysis tool
│   ├── mcp-server.js             # MCP server for Claude Code
│   └── README.md                 # Analyzer documentation
├── SHAREPOINT-LEARNINGS.md       # Methodology document
├── qa-screenshots/               # Example QA output
│   ├── desktop-fullpage.png
│   ├── mobile-375x812.png
│   └── ...
└── README.md                     # This file
```

---

## Multi-Editor Support

| Tool | Skill | MCP Analyzer |
|------|-------|--------------|
| Claude Code | Works natively (SKILL.md) | Works natively (.mcp.json) |
| Cursor | Convert to .cursorrules | Works (MCP supported) |
| GitHub Copilot | Convert to copilot-instructions.md | Not yet supported |
| Windsurf | Convert to .windsurfrules | Works (MCP supported) |
| VS Code + Continue | Convert to .continue/ config | Works (MCP supported) |

The skill file is currently in Claude Code's SKILL.md format. The knowledge inside is editor-agnostic -- PRs welcome for native format conversions. The MCP analyzer works with any editor that supports the Model Context Protocol.

---

## Contributing

Contributions are welcome. This toolkit grew out of a real delivery -- if you've found patterns that work (or don't), we'd like to hear about them.

**Areas where contributions would be especially valuable:**

- Additional page analyzer extractors (e.g., list view analysis, search configuration)
- CI/CD patterns for Azure DevOps (we currently cover GitHub Actions)
- Multi-language / multi-geo deployment patterns
- SPFx 1.21+ compatibility notes
- Classic-to-Modern migration case studies

**To contribute:**

1. Fork the repo
2. Create a feature branch (`git checkout -b feature/your-feature`)
3. Commit your changes
4. Open a pull request with a clear description of what you added and why

---

## License

MIT License. See [LICENSE](LICENSE) for details.

---

## Credits

Built by [Code&Canvas](https://codeandcanvas.io) while delivering a SharePoint Modern intranet for a Fortune 500 client.

**Acknowledgments:**

- [Sudharsan Kesavanarayanan](https://github.com/nickhsm) -- for `spfx-enterprise-skills`, the coding standards skill pack that this delivery skill complements
- [PnP Community](https://pnp.github.io/) -- for CLI for Microsoft 365 and the broader SharePoint dev ecosystem
- [Anthropic](https://anthropic.com) -- Claude Code made AI-assisted SPFx delivery practical

---

*Built with AI assistance. Shipped by humans.*
