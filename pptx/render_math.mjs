#!/usr/bin/env node
/**
 * LaTeX → PNG renderer using KaTeX + Playwright.
 *
 * Usage:
 *   node pptx/render_math.mjs '<latex>' output.png [--display] [--fontsize=28]
 *
 * Renders a single LaTeX expression to a transparent PNG image.
 * Uses KaTeX for rendering and Playwright for screenshot.
 */

import { chromium } from 'playwright';
import katex from 'katex';
import { writeFileSync, readFileSync } from 'fs';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

const __dirname = dirname(fileURLToPath(import.meta.url));
const katexCssPath = join(__dirname, '..', 'node_modules', 'katex', 'dist', 'katex.min.css');
const katexCss = readFileSync(katexCssPath, 'utf-8');
// Inline font faces from KaTeX
const fontsDir = join(__dirname, '..', 'node_modules', 'katex', 'dist', 'fonts');

const args = process.argv.slice(2);
const displayMode = args.includes('--display');
const fontSizeArg = args.find(a => a.startsWith('--fontsize='));
const fontSize = fontSizeArg ? parseInt(fontSizeArg.split('=')[1]) : 28;
const filtered = args.filter(a => !a.startsWith('--'));

const latex = filtered[0];
const outPath = filtered[1];

if (!latex || !outPath) {
    console.error('Usage: node render_math.mjs <latex> <output.png> [--display] [--fontsize=N]');
    process.exit(1);
}

const html = katex.renderToString(latex, {
    displayMode,
    output: 'html',
    throwOnError: false,
});

// Fix font URLs to absolute paths
const fixedCss = katexCss.replace(/url\(fonts\//g, `url(file://${fontsDir}/`);

const fullHtml = `<!DOCTYPE html>
<html>
<head>
<style>
${fixedCss}
* { margin: 0; padding: 0; }
body {
    background: transparent;
    display: inline-block;
    padding: 8px 12px;
}
.katex {
    font-size: ${fontSize}px;
    color: #1a1a2e;
}
</style>
</head>
<body>${html}</body>
</html>`;

const browser = await chromium.launch();
const page = await browser.newPage();
await page.setContent(fullHtml, { waitUntil: 'networkidle' });

// Get bounding box of the body
const bbox = await page.evaluate(() => {
    const body = document.body;
    const rect = body.getBoundingClientRect();
    return { width: Math.ceil(rect.width), height: Math.ceil(rect.height) };
});

await page.setViewportSize({ width: bbox.width + 4, height: bbox.height + 4 });

await page.screenshot({
    path: outPath,
    omitBackground: true,
    clip: { x: 0, y: 0, width: bbox.width + 4, height: bbox.height + 4 },
});

await browser.close();
