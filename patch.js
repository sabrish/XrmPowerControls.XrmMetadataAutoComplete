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

const pcfStartDir  = path.join(__dirname, 'node_modules', 'pcf-start');
const indexHtmlPath = path.join(pcfStartDir, 'index.html');
const libDir        = path.join(pcfStartDir, 'lib');
const mockApiPath   = path.join(libDir, 'mock-api.js');

// ─────────────────────────────────────────────────────────────────────────────
// 1.  Patch pcf-start/index.html
// ─────────────────────────────────────────────────────────────────────────────
let html = fs.readFileSync(indexHtmlPath, 'utf8');

// a) React 18 platform lib + make window.React point to Reactv18 permanently
//    so the harness's dynamic Fluent UMD load uses React 18 (not the CDN React 16).
if (!html.includes('react_18_3_1.js')) {
  html = html.replace(
    'window["Reactv16"]=window.React;window["ReactDOMv16"]=window.ReactDOM;',
    'window["Reactv16"]=window.React;window["ReactDOMv16"]=window.ReactDOM;' +
    '</script><script src="/lib/react_18_3_1.js"></script>' +
    '<script>window.React=window.Reactv18;window.ReactDOM=window.ReactDOMv18;'
  );
  console.log('  Applied: React 18 script tags');
}

// b) Mock API interceptor — must load BEFORE harness.js so fetch is patched
//    before any PCF control init code runs.
if (!html.includes('mock-api.js')) {
  html = html.replace(
    '<script defer="defer" src="harness.js">',
    '<script src="/lib/mock-api.js"></script><script defer="defer" src="harness.js">'
  );
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

