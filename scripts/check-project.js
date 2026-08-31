'use strict';

const fs = require('node:fs');
const vm = require('node:vm');

const requiredFiles = [
  'shortcuts_sidebar.gs',
  'Sidebar.html',
  'appsscript.json',
];

for (const file of requiredFiles) {
  if (!fs.existsSync(file)) throw new Error(`Missing required file: ${file}`);
}

const serverSource = fs.readFileSync('shortcuts_sidebar.gs', 'utf8');
new vm.Script(serverSource, { filename: 'shortcuts_sidebar.gs' });

const html = fs.readFileSync('Sidebar.html', 'utf8');
if (!/^<!doctype html>/i.test(html)) throw new Error('Sidebar.html needs a doctype.');
if (!html.includes('role="status"')) throw new Error('Sidebar status must be accessible.');
if (html.includes('.innerHTML')) throw new Error('Do not render server text with innerHTML.');

const scripts = [...html.matchAll(/<script>([\s\S]*?)<\/script>/gi)];
if (scripts.length !== 1) throw new Error('Expected exactly one inline sidebar script.');
new vm.Script(scripts[0][1], { filename: 'Sidebar.html:inline-script' });

const manifest = JSON.parse(fs.readFileSync('appsscript.json', 'utf8'));
const scopes = manifest.oauthScopes || [];
const expectedScopes = [
  'https://www.googleapis.com/auth/spreadsheets.currentonly',
  'https://www.googleapis.com/auth/script.container.ui',
];
for (const scope of expectedScopes) {
  if (!scopes.includes(scope)) throw new Error(`Missing least-privilege scope: ${scope}`);
}
if (scopes.some(scope => scope.includes('/auth/drive'))) {
  throw new Error('A Drive scope was added; review and document why it is necessary.');
}

console.log('Project files, JavaScript syntax, HTML safeguards, and manifest are valid.');
