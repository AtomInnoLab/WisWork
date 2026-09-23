const { app, nativeImage } = require('electron')
const assert = require('node:assert/strict')
app
  .whenReady()
  .then(async () => {
    const { normalizePresentationImage } = require(process.argv.at(-1))
    const image = nativeImage.createFromBitmap(
      Buffer.from([255, 0, 0, 255, 0, 255, 0, 255, 0, 0, 255, 255, 255, 255, 255, 255]),
      { width: 2, height: 2 },
    )
    for (const bytes of [image.toPNG(), image.toJPEG(90)]) {
      const normalized = await normalizePresentationImage(bytes)
      assert.equal(normalized.width, 2)
      assert.equal(normalized.height, 2)
      assert.equal(nativeImage.createFromBuffer(Buffer.from(normalized.bytes)).isEmpty(), false)
      assert.equal(Buffer.from(normalized.bytes).subarray(1, 4).toString(), 'PNG')
    }
    await assert.rejects(normalizePresentationImage(Buffer.from('not an image')))
    console.log(
      'Native Electron PNG/JPEG decode, PNG normalization and invalid-input rejection passed',
    )
    app.exit(0)
  })
  .catch((error) => {
    console.error(error)
    app.exit(1)
  })
