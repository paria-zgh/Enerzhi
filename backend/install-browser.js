// backend/install-browser.js
const { execSync } = require('child_process');
const path = require('path');

const cliPath = path.resolve(__dirname, 'node_modules/playwright-core/cli.js');
console.log('Downloading Chromium browser directly via node...');
execSync(`node "${cliPath}" install chromium`, { stdio: 'inherit' });
