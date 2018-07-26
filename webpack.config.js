const path = require('path');
const { DefinePlugin } = require('webpack');
const configs = require('sp-build-tasks/dist/webpack/config');

configs.forEach(conf => {
  // conf.plugins = conf.plugins || [];
  // conf.plugins.push(new DefinePlugin(defineOptions));
  conf.output.chunkFilename = '[name].js';
});

module.exports = configs;
