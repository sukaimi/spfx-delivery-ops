#!/usr/bin/env node

/**
 * SP Page Analyzer — MCP Server
 *
 * Wraps the SharePoint page analysis tool as an MCP (Model Context Protocol)
 * server using stdio transport. Exposes four tools:
 *
 *   - analyze_page          Full 6-step pipeline
 *   - extract_design_tokens Design token extraction only
 *   - capture_screenshots   Screenshot capture only
 *   - list_images           Image inventory only
 *
 * Usage:
 *   node mcp-server.js          (stdio transport — for Claude Code / MCP clients)
 */

const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const { z } = require('zod');
const {
  analyzePage,
  analyzeDesignTokens,
  analyzeScreenshots,
  analyzeImages,
} = require('./lib/analyzer.js');

const server = new McpServer({
  name: 'sp-page-analyzer',
  version: '1.0.0',
});

// ---------------------------------------------------------------------------
// Tool: analyze_page
// ---------------------------------------------------------------------------
server.tool(
  'analyze_page',
  'Run the full 6-step SharePoint page analysis pipeline. Extracts page type, design tokens, structure, images, navigation, custom assets, and captures screenshots. Returns structured JSON.',
  {
    url: z.string().url().describe('The SharePoint page URL to analyze'),
    auth: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        'If true, opens a visible browser for interactive login. Default false (headless).'
      ),
  },
  async ({ url, auth }) => {
    try {
      const result = await analyzePage(url, { auth });
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text',
            text: `Analysis failed: ${error.message}`,
          },
        ],
      };
    }
  }
);

// ---------------------------------------------------------------------------
// Tool: extract_design_tokens
// ---------------------------------------------------------------------------
server.tool(
  'extract_design_tokens',
  'Extract design tokens (colors, typography, spacing, dimensions) from a SharePoint page. Useful for understanding the visual design system in use.',
  {
    url: z.string().url().describe('The SharePoint page URL to analyze'),
    auth: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        'If true, opens a visible browser for interactive login. Default false (headless).'
      ),
  },
  async ({ url, auth }) => {
    try {
      const result = await analyzeDesignTokens(url, { auth });
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text',
            text: `Design token extraction failed: ${error.message}`,
          },
        ],
      };
    }
  }
);

// ---------------------------------------------------------------------------
// Tool: capture_screenshots
// ---------------------------------------------------------------------------
server.tool(
  'capture_screenshots',
  'Capture screenshots of a SharePoint page (viewport, full-page, and scroll segments). Returns base64-encoded PNG data keyed by filename.',
  {
    url: z.string().url().describe('The SharePoint page URL to screenshot'),
    auth: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        'If true, opens a visible browser for interactive login. Default false (headless).'
      ),
  },
  async ({ url, auth }) => {
    try {
      const result = await analyzeScreenshots(url, { auth });
      // Return metadata as text, screenshots are base64 in the JSON
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text',
            text: `Screenshot capture failed: ${error.message}`,
          },
        ],
      };
    }
  }
);

// ---------------------------------------------------------------------------
// Tool: list_images
// ---------------------------------------------------------------------------
server.tool(
  'list_images',
  'Extract an inventory of all images on a SharePoint page, including inline images (with dimensions, alt text, position) and CSS background images.',
  {
    url: z.string().url().describe('The SharePoint page URL to scan for images'),
    auth: z
      .boolean()
      .optional()
      .default(false)
      .describe(
        'If true, opens a visible browser for interactive login. Default false (headless).'
      ),
  },
  async ({ url, auth }) => {
    try {
      const result = await analyzeImages(url, { auth });
      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(result, null, 2),
          },
        ],
      };
    } catch (error) {
      return {
        isError: true,
        content: [
          {
            type: 'text',
            text: `Image extraction failed: ${error.message}`,
          },
        ],
      };
    }
  }
);

// ---------------------------------------------------------------------------
// Start the server
// ---------------------------------------------------------------------------
async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((error) => {
  console.error('MCP server failed to start:', error);
  process.exit(1);
});
