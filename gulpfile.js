'use strict';

// eslint-disable-next-line @typescript-eslint/no-require-imports, no-undef
const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not camelCase and will not be type-safe.`);

var getTasks = build.rig.getTasks;
build.rig.getTasks = function () {
  var result = getTasks.call(build.rig);

  result.set('serve', result.get('serve-deprecated'));

  return result;
};

// Add font loader configuration for handling TTF, WOFF, etc.
const fontLoaderConfig = {
  test: /\.(woff(2)?|ttf|eot|svg)(\?v=\d+\.\d+\.\d+)?$/,
  type: 'asset/resource',
  generator: {
    filename: 'fonts/[name][ext]'
  }
};

// Merge the font loader into webpack configuration
build.configureWebpack.mergeConfig({
  additionalConfiguration: (generatedConfiguration) => {
    generatedConfiguration.module.rules.push(fontLoaderConfig);
    return generatedConfiguration;
  }
});

// eslint-disable-next-line @typescript-eslint/no-require-imports, no-undef
build.initialize(require('gulp'));
