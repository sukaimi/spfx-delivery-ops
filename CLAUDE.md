# CLAUDE.md — SPFx Delivery Ops Skill

## What This Is
A Claude Code skill package for SharePoint Framework project delivery. Covers the practical delivery side — resolution hierarchy, canvas overrides, CI/CD, image deployment, ES5 gotchas, and full QA workflow.

**Complements** `@sudharsank/spfx-enterprise-skills` (coding standards, 14 skills). This skill covers what that pack doesn't: how to actually ship.

## Folder Structure
```
spfx-delivery-skill/
├── CLAUDE.md              ← You are here
├── README.md              ← Public repo README (for GitHub publication)
├── skill/
│   └── SKILL.md           ← The Claude Code skill file (install this)
├── analyzer/
│   ├── mcp-server.js      ← MCP server for Claude Code integration
│   ├── lib/analyzer.js    ← Page analysis library (Playwright-based)
│   ├── package.json       ← Dependencies
│   └── README.md          ← Analyzer documentation
├── docs/
│   └── METHODOLOGY.md     ← Full methodology document (7 sections)
├── screenshots/           ← QA evidence screenshots
└── examples/              ← Example webpart patterns (future)
```

## How to Use

### Install the skill into a project
```bash
cp skill/SKILL.md /path/to/your-spfx-project/.claude/skills/spfx-delivery-ops/SKILL.md
```

### Install the MCP server
```bash
cd analyzer && npm install
```
Add to your `.mcp.json`:
```json
{
  "mcpServers": {
    "sp-page-analyzer": {
      "command": "node",
      "args": ["/path/to/analyzer/mcp-server.js"]
    }
  }
}
```

### Use alongside enterprise skills
```bash
npx @sudharsank/spfx-enterprise-skills install --host claude --mode project --project-path . --skills all
```

## Key Principles
1. **Three-tier resolution**: OOB features first → Custom SPFx → DOM hacks (last resort)
2. **Brand-agnostic**: No client names, no hardcoded colors, discovers what's on the page
3. **Delivery-focused**: Discovery → Build → Deploy → QA workflow
4. **Battle-tested**: Built delivering a real enterprise intranet with 4 modular webparts

## Related Resources
- Enterprise skills: https://github.com/sudharsank/spfx-enterprise-skills
- SPFx Toolkit: https://pnp.github.io/vscode-viva/
- PnP CLI M365 MCP: https://github.com/pnp/cli-microsoft365-mcp-server
