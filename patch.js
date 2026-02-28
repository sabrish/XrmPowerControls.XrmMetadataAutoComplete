#!/usr/bin/env node
/**
 * patch.js — applies development-only patches to node_modules/pcf-start.
 * Runs automatically after every `npm install` via the "postinstall" script.
 * Safe to run multiple times (idempotent).
 *
 * Patches applied:
 *   1. pcf-start/index.html — loads React 18 platform lib and sets window.React
 *      to Reactv18 so the harness's dynamic Fluent load uses React 18, not 16.
 *   2. pcf-start/lib/mock-api.js — fetch interceptor that returns realistic mock
 *      Dataverse metadata for all 5 supported types so the control can be tested
 *      in the harness without a live Dataverse connection.
 */
'use strict';

const fs   = require('fs');
const path = require('path');

const pcfStartDir   = path.join(__dirname, 'node_modules', 'pcf-start');
const indexHtmlPath = path.join(pcfStartDir, 'index.html');
const libDir        = path.join(pcfStartDir, 'lib');
const mockApiPath   = path.join(libDir, 'mock-api.js');

// ─────────────────────────────────────────────────────────────────────────────
// Guard: pcf-start is a devDependency and won't be present in production
//        installs (npm ci --omit=dev).  Skip silently in that case.
// ─────────────────────────────────────────────────────────────────────────────
if (!fs.existsSync(pcfStartDir)) {
  console.log('patch.js: pcf-start not installed — skipping harness patches.');
  process.exit(0);
}

if (!fs.existsSync(indexHtmlPath)) {
  console.error('patch.js ERROR: pcf-start is installed but index.html was not found at:');
  console.error(' ', indexHtmlPath);
  console.error('The pcf-start package layout may have changed. Check the package contents.');
  process.exit(1);
}

// ─────────────────────────────────────────────────────────────────────────────
// 1.  Patch pcf-start/index.html
// ─────────────────────────────────────────────────────────────────────────────
let html = fs.readFileSync(indexHtmlPath, 'utf8');

// a) React 18 platform lib + make window.React point to Reactv18 permanently
//    so the harness's dynamic Fluent UMD load uses React 18 (not the CDN React 16).
if (!html.includes('react_18_3_1.js')) {
  const TARGET_A = 'window["Reactv16"]=window.React;window["ReactDOMv16"]=window.ReactDOM;';
  const before = html;
  html = html.replace(
    TARGET_A,
    TARGET_A +
    '</script><script src="/lib/react_18_3_1.js"></script>' +
    '<script>window.React=window.Reactv18;window.ReactDOM=window.ReactDOMv18;'
  );
  if (html === before) {
    console.error('patch.js ERROR: React 18 patch failed — the expected string was not found in index.html.');
    console.error('The pcf-start template may have changed. Update TARGET_A in patch.js to match.');
    process.exit(1);
  }
  console.log('  Applied: React 18 script tags');
}

// b) Mock API interceptor — must load BEFORE harness.js so fetch is patched
//    before any PCF control init code runs.
if (!html.includes('mock-api.js')) {
  const TARGET_B = '<script defer="defer" src="harness.js">';
  const before = html;
  html = html.replace(
    TARGET_B,
    '<script src="/lib/mock-api.js"></script>' + TARGET_B
  );
  if (html === before) {
    console.error('patch.js ERROR: mock-api patch failed — the expected string was not found in index.html.');
    console.error('The pcf-start template may have changed. Update TARGET_B in patch.js to match.');
    process.exit(1);
  }
  console.log('  Applied: mock-api.js script tag');
}

fs.writeFileSync(indexHtmlPath, html);
console.log('Patched pcf-start/index.html');

// ─────────────────────────────────────────────────────────────────────────────
// 2.  Copy mock-api.src.js → pcf-start/lib/mock-api.js
//     Keeping the source as a real .js file avoids any template-literal
//     escape issues when embedding regex patterns.
// ─────────────────────────────────────────────────────────────────────────────
const mockApiSrc = path.join(__dirname, 'mock-api.src.js');
fs.mkdirSync(libDir, { recursive: true });
fs.copyFileSync(mockApiSrc, mockApiPath);
console.log('Copied mock-api.src.js → pcf-start/lib/mock-api.js');
