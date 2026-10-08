// No signing credentials or release feed are used for automatic PR artifacts.
const { createPreviewConfig } = require('../../tools/release/preview-metadata.cjs')
const base = require('./electron-builder.cjs')

module.exports = createPreviewConfig(base, {
  pr: process.env.WISWORK_PREVIEW_PR,
  commit: process.env.WISWORK_PREVIEW_COMMIT,
  builtAt: process.env.WISWORK_PREVIEW_BUILT_AT,
  version: require('./package.json').version,
})
