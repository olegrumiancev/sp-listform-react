// Rewrite webpack config if needed, aka "eject"
const path = require('path');
const webpack = require("webpack");
const config = require('sp-build-tasks/dist/webpack/config');
config.resolve = { alias: { 'react': path.resolve(__dirname, 'node_modules', 'react') } }

if (config[0]) {
  if (config[0].optimization) {
    if (config[0].optimization.minimizer && config[0].optimization.minimizer.length > 0) {
      let uglify = config[0].optimization.minimizer[0];
      uglify.options.sourceMap = false;
      uglify.options.uglifyOptions.ie8 = false;
      uglify.options.uglifyOptions.safari10 = false;
      uglify.options.uglifyOptions.keep_fnames = false;
    }
    //config[0].optimization.splitChunks =
    //   { chunks: 'all' };
      // {
      //   chunks: 'async',
      //   minSize: 40000,
      //   maxSize: 0,
      //   minChunks: 1,
      //   maxAsyncRequests: 5,
      //   maxInitialRequests: 3,
      //   automaticNameDelimiter: '~',
      //   name: true,
      //   cacheGroups: {
      //     vendors: {
      //       test: /[\\/]node_modules[\\/]/,
      //       priority: -10
      //     },
      //     default: {
      //       minChunks: 2,
      //       priority: -20,
      //       reuseExistingChunk: true
      //     }
      //   }
      // };
  }
  if (!config[0].plugins) {
    config[0].plugins = [];
  }
  // config[0].plugins.push(new webpack.IgnorePlugin(/^\.\/locale$/, [/moment$/]));
}

module.exports = config;