/**
 * SP Page Analyzer — Modular library version
 *
 * Exposes individual pipeline steps so they can be called independently
 * from the MCP server or CLI.
 */

const { chromium } = require('playwright');
const path = require('path');
const fs = require('fs');
const os = require('os');

const USER_DATA_DIR = path.join(os.homedir(), '.sp-analyzer-browser');

/**
 * Launch a persistent browser context and navigate to the target URL.
 * Returns { context, page, pageType }.
 */
async function openPage(url, { auth = false, waitTime = 5 } = {}) {
  const context = await chromium.launchPersistentContext(USER_DATA_DIR, {
    headless: !auth,
    viewport: { width: 1440, height: 900 },
    args: ['--disable-blink-features=AutomationControlled'],
  });

  const page = await context.newPage();
  await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });

  // Check for login redirect
  const currentUrl = page.url();
  if (
    currentUrl.includes('login.microsoftonline.com') ||
    currentUrl.includes('login.live.com')
  ) {
    if (auth) {
      await page.waitForURL('**/sharepoint.com/**', { timeout: 120000 });
    } else {
      await context.close();
      throw new Error(
        'Login required. Set auth=true for interactive login, or authenticate in the browser first.'
      );
    }
  }

  // Wait for JS rendering
  await page.waitForTimeout(waitTime * 1000);

  // Detect page type
  const pageType = await page.evaluate(() => {
    return {
      isModern: !!(
        document.querySelector('[data-automation-id="contentScrollRegion"]') ||
        document.querySelector('.CanvasZone') ||
        document.querySelector('[data-sp-feature-tag]')
      ),
      isClassic: !!(
        document.querySelector('#s4-workspace') ||
        document.querySelector('#DeltaPlaceHolderMain') ||
        document.querySelector('.ms-rtestate-field')
      ),
      hasCustomJS: !!(
        document.querySelector('script[src*="Style Library"]') ||
        document.querySelector('script[src*="SiteAssets"]')
      ),
      scrollContainer: document.querySelector('#s4-workspace')
        ? '#s4-workspace'
        : document.querySelector('[data-automation-id="contentScrollRegion"]')
          ? '[data-automation-id="contentScrollRegion"]'
          : 'window',
      pageTitle: document.title,
      url: window.location.href,
    };
  });

  return { context, page, pageType };
}

/**
 * Step 3: Extract design tokens (colors, typography, spacing, dimensions).
 */
async function extractDesignTokens(page) {
  return page.evaluate(() => {
    const tokens = {
      colors: { backgrounds: {}, text: {}, borders: {} },
      typography: { fontFamilies: [], fontSizes: [], fontWeights: [] },
      spacing: {},
      dimensions: {},
    };

    const bgColorMap = new Map();
    const textColorMap = new Map();
    const fontFamilySet = new Set();
    const fontSizeSet = new Set();

    const allElements = document.querySelectorAll('*');
    const sampleSize = Math.min(allElements.length, 1000);

    for (let i = 0; i < sampleSize; i++) {
      const el = allElements[i];
      const cs = getComputedStyle(el);
      const rect = el.getBoundingClientRect();

      if (rect.height < 5 || rect.width < 5) continue;

      const bg = cs.backgroundColor;
      if (bg && bg !== 'rgba(0, 0, 0, 0)' && bg !== 'transparent') {
        bgColorMap.set(bg, (bgColorMap.get(bg) || 0) + 1);
      }

      if (
        el.textContent &&
        el.textContent.trim().length > 0 &&
        el.children.length === 0
      ) {
        const color = cs.color;
        textColorMap.set(color, (textColorMap.get(color) || 0) + 1);
      }

      if (cs.fontFamily) fontFamilySet.add(cs.fontFamily);
      if (cs.fontSize) fontSizeSet.add(cs.fontSize);
    }

    tokens.colors.backgrounds = Object.fromEntries(
      [...bgColorMap.entries()].sort((a, b) => b[1] - a[1]).slice(0, 20)
    );
    tokens.colors.text = Object.fromEntries(
      [...textColorMap.entries()].sort((a, b) => b[1] - a[1]).slice(0, 15)
    );
    tokens.typography.fontFamilies = [...fontFamilySet].slice(0, 10);
    tokens.typography.fontSizes = [...fontSizeSet].sort().slice(0, 20);

    const body = document.body;
    const html = document.documentElement;
    tokens.dimensions = {
      viewportWidth: window.innerWidth,
      viewportHeight: window.innerHeight,
      pageWidth: Math.max(body.scrollWidth, html.scrollWidth),
      pageHeight: Math.max(body.scrollHeight, html.scrollHeight),
    };

    return tokens;
  });
}

/**
 * Step 4: Extract page structure (sections, web parts, content areas).
 */
async function extractStructure(page) {
  return page.evaluate(() => {
    const sections = [];

    const canvasSections = document.querySelectorAll(
      '[data-automation-id="CanvasSection"], .CanvasZone'
    );
    const classicZones = document.querySelectorAll(
      '#DeltaPlaceHolderMain > *, .ms-rte-layoutszone-inner > *'
    );
    const mainContent = document.querySelector(
      'main, [role="main"], #contentBox, #DeltaPlaceHolderMain, article'
    );

    const sourceElements =
      canvasSections.length > 0
        ? canvasSections
        : classicZones.length > 0
          ? classicZones
          : mainContent
            ? mainContent.children
            : [];

    for (let i = 0; i < sourceElements.length; i++) {
      const el = sourceElements[i];
      const rect = el.getBoundingClientRect();
      if (rect.height < 10) continue;

      const cs = getComputedStyle(el);
      const section = {
        index: i,
        tag: el.tagName.toLowerCase(),
        id: el.id || null,
        className:
          el.className && typeof el.className === 'string'
            ? el.className.substring(0, 200)
            : null,
        name:
          el.getAttribute('data-automation-id') ||
          el.getAttribute('aria-label') ||
          el.id ||
          (el.querySelector('h1,h2,h3,h4') || {}).textContent ||
          null,
        dimensions: {
          x: Math.round(rect.x),
          y: Math.round(rect.y),
          width: Math.round(rect.width),
          height: Math.round(rect.height),
        },
        styles: {
          backgroundColor:
            cs.backgroundColor !== 'rgba(0, 0, 0, 0)'
              ? cs.backgroundColor
              : null,
          backgroundImage:
            cs.backgroundImage !== 'none'
              ? cs.backgroundImage.substring(0, 200)
              : null,
          color: cs.color,
          padding: cs.padding,
          margin: cs.margin,
        },
        childCount: el.children.length,
        textContent: (el.textContent || '').trim().substring(0, 100),
      };

      const wpType = el.querySelector('[data-sp-web-part-id]');
      if (wpType) {
        section.webPartId = wpType.getAttribute('data-sp-web-part-id');
      }

      sections.push(section);
    }

    return sections;
  });
}

/**
 * Step 5: Extract all images (inline and background).
 */
async function extractImages(page) {
  return page.evaluate(() => {
    const imgs = [];
    document.querySelectorAll('img[src]').forEach((img) => {
      const rect = img.getBoundingClientRect();
      if (rect.width < 20 && rect.height < 20) return;
      if (img.src.startsWith('data:')) return;
      if (img.src.includes('_layouts/15/images')) return;

      imgs.push({
        src: img.src,
        alt: img.alt || '',
        naturalWidth: img.naturalWidth,
        naturalHeight: img.naturalHeight,
        displayWidth: Math.round(rect.width),
        displayHeight: Math.round(rect.height),
        x: Math.round(rect.x),
        y: Math.round(rect.y),
      });
    });

    const bgImages = [];
    document.querySelectorAll('*').forEach((el) => {
      const cs = getComputedStyle(el);
      if (cs.backgroundImage && cs.backgroundImage !== 'none') {
        const urlMatch = cs.backgroundImage.match(
          /url\(["']?([^"')]+)["']?\)/
        );
        if (
          urlMatch &&
          !urlMatch[1].startsWith('data:') &&
          !urlMatch[1].includes('_layouts')
        ) {
          bgImages.push({
            src: urlMatch[1],
            type: 'background-image',
            element:
              el.tagName.toLowerCase() +
              (el.id ? '#' + el.id : '') +
              (el.className && typeof el.className === 'string'
                ? '.' + el.className.split(' ')[0]
                : ''),
          });
        }
      }
    });

    return { images: imgs, backgroundImages: bgImages };
  });
}

/**
 * Extract custom CSS and JS references.
 */
async function extractCustomAssets(page) {
  return page.evaluate(() => {
    const assets = { stylesheets: [], scripts: [] };

    document.querySelectorAll('link[rel="stylesheet"]').forEach((l) => {
      if (
        l.href &&
        !l.href.includes('_layouts') &&
        !l.href.includes('microsoftonline') &&
        !l.href.includes('cdn.office.net') &&
        !l.href.includes('res-')
      ) {
        assets.stylesheets.push(l.href);
      }
    });

    document.querySelectorAll('script[src]').forEach((s) => {
      if (
        s.src &&
        !s.src.includes('_layouts') &&
        !s.src.includes('microsoftonline') &&
        !s.src.includes('cdn.office.net') &&
        !s.src.includes('res-') &&
        !s.src.includes('microsoft.com') &&
        !s.src.includes('googletagmanager') &&
        !s.src.includes('google-analytics')
      ) {
        assets.scripts.push(s.src);
      }
    });

    return assets;
  });
}

/**
 * Extract navigation structure.
 */
async function extractNavigation(page) {
  return page.evaluate(() => {
    const navItems = [];
    const navSelectors = [
      'nav a',
      '[role="navigation"] a',
      '#DeltaTopNavigation a',
      '.ms-core-listMenu-item',
      '[data-automation-id="HorizontalNav"] a',
    ];

    for (const selector of navSelectors) {
      const links = document.querySelectorAll(selector);
      if (links.length > 0) {
        links.forEach((a) => {
          const text = (a.textContent || '').trim();
          if (text && text.length > 0 && text.length < 100) {
            navItems.push({ text, href: a.href, selector });
          }
        });
        break;
      }
    }

    return navItems;
  });
}

/**
 * Step 6: Capture screenshots (viewport, full-page, scroll segments).
 * Returns an object mapping screenshot names to base64-encoded PNG data.
 */
async function captureScreenshots(page, pageType) {
  const screenshots = {};

  // Viewport screenshot
  const viewportBuf = await page.screenshot({ type: 'png' });
  screenshots['screenshot-viewport.png'] = viewportBuf.toString('base64');

  // Scroll to bottom to trigger lazy loading, then back to top
  const scrollContainer = pageType.scrollContainer;
  if (scrollContainer !== 'window') {
    await page.evaluate((sel) => {
      const el = document.querySelector(sel);
      if (el) el.scrollTop = el.scrollHeight;
    }, scrollContainer);
    await page.waitForTimeout(2000);
    await page.evaluate((sel) => {
      const el = document.querySelector(sel);
      if (el) el.scrollTop = 0;
    }, scrollContainer);
    await page.waitForTimeout(1000);
  }

  // Full page screenshot
  try {
    const fullBuf = await page.screenshot({ fullPage: true, type: 'png' });
    screenshots['screenshot-full.png'] = fullBuf.toString('base64');
  } catch (_e) {
    // Full-page screenshot may fail on pages with inner scroll containers
  }

  // Scrolled screenshots for pages with inner scroll containers
  if (scrollContainer !== 'window') {
    const scrollHeight = await page.evaluate((sel) => {
      const el = document.querySelector(sel);
      return el ? el.scrollHeight : 0;
    }, scrollContainer);

    const viewportHeight = 900;
    const numScreenshots = Math.min(Math.ceil(scrollHeight / viewportHeight), 5);

    for (let i = 0; i < numScreenshots; i++) {
      await page.evaluate(
        (sel, scrollTo) => {
          const el = document.querySelector(sel);
          if (el) el.scrollTop = scrollTo;
        },
        scrollContainer,
        i * viewportHeight
      );
      await page.waitForTimeout(500);
      const buf = await page.screenshot({ type: 'png' });
      screenshots[`screenshot-scroll-${i}.png`] = buf.toString('base64');
    }
  }

  return screenshots;
}

/**
 * Run the full 6-step analysis pipeline.
 * Returns the complete analysis object (same shape as analysis.json from the CLI).
 */
async function analyzePage(url, { auth = false, waitTime = 5 } = {}) {
  const { context, page, pageType } = await openPage(url, { auth, waitTime });

  try {
    const [designTokens, structure, imageData, customAssets, navigation] =
      await Promise.all([
        extractDesignTokens(page),
        extractStructure(page),
        extractImages(page),
        extractCustomAssets(page),
        extractNavigation(page),
      ]);

    const screenshots = await captureScreenshots(page, pageType);

    return {
      meta: {
        analyzedUrl: url,
        finalUrl: page.url(),
        pageTitle: pageType.pageTitle,
        pageType: pageType.isModern
          ? 'Modern'
          : pageType.isClassic
            ? 'Classic'
            : 'Unknown',
        hasCustomJS: pageType.hasCustomJS,
        scrollContainer: pageType.scrollContainer,
        analyzedAt: new Date().toISOString(),
        toolVersion: '1.0.0',
      },
      designTokens,
      structure,
      images: imageData.images,
      backgroundImages: imageData.backgroundImages,
      navigation,
      customAssets,
      screenshots: Object.keys(screenshots),
    };
  } finally {
    await context.close();
  }
}

/**
 * Run only the design token extraction step.
 */
async function analyzeDesignTokens(url, { auth = false, waitTime = 5 } = {}) {
  const { context, page, pageType } = await openPage(url, { auth, waitTime });
  try {
    const designTokens = await extractDesignTokens(page);
    return {
      meta: {
        analyzedUrl: url,
        finalUrl: page.url(),
        pageTitle: pageType.pageTitle,
        pageType: pageType.isModern
          ? 'Modern'
          : pageType.isClassic
            ? 'Classic'
            : 'Unknown',
        analyzedAt: new Date().toISOString(),
      },
      designTokens,
    };
  } finally {
    await context.close();
  }
}

/**
 * Run only the screenshot capture step.
 */
async function analyzeScreenshots(url, { auth = false, waitTime = 5 } = {}) {
  const { context, page, pageType } = await openPage(url, { auth, waitTime });
  try {
    const screenshots = await captureScreenshots(page, pageType);
    return {
      meta: {
        analyzedUrl: url,
        finalUrl: page.url(),
        pageTitle: pageType.pageTitle,
        analyzedAt: new Date().toISOString(),
      },
      screenshots,
    };
  } finally {
    await context.close();
  }
}

/**
 * Run only the image inventory step.
 */
async function analyzeImages(url, { auth = false, waitTime = 5 } = {}) {
  const { context, page, pageType } = await openPage(url, { auth, waitTime });
  try {
    const imageData = await extractImages(page);
    return {
      meta: {
        analyzedUrl: url,
        finalUrl: page.url(),
        pageTitle: pageType.pageTitle,
        analyzedAt: new Date().toISOString(),
      },
      images: imageData.images,
      backgroundImages: imageData.backgroundImages,
    };
  } finally {
    await context.close();
  }
}

module.exports = {
  analyzePage,
  analyzeDesignTokens,
  analyzeScreenshots,
  analyzeImages,
};
