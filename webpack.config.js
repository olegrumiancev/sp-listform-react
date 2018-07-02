// Rewrite webpack config if needed, aka "eject"
const path = require('path');
const config = require('sp-build-tasks/dist/webpack/config');

config.resolve = { alias: { 'react': path.resolve(__dirname, 'node_modules', 'react') } }

module.exports = config;