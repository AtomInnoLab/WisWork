import { writeFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { createCanvas, loadImage } from '@napi-rs/canvas'

const sourcePath = process.argv[2]
if (!sourcePath) throw new Error('usage: node tools/build-wiswork-icons.mjs <source.png>')

const root = resolve(import.meta.dirname, '..')
const source = await loadImage(resolve(sourcePath))

function renderPng(size) {
  const canvas = createCanvas(size, size)
  const ctx = canvas.getContext('2d')
  ctx.clearRect(0, 0, size, size)
  ctx.imageSmoothingEnabled = true
  ctx.imageSmoothingQuality = 'high'
  const scale = Math.min(size / source.width, size / source.height)
  const width = source.width * scale
  const height = source.height * scale
  ctx.drawImage(source, (size - width) / 2, (size - height) / 2, width, height)
  return canvas.toBuffer('image/png')
}

const png1024 = renderPng(1024)
const pngTargets = [
  'apps/shell/build/icon.png',
  'apps/shell/build/icon-mac.png',
  'apps/shell/src/renderer/src/assets/app-icon.png',
  'apps/docs/src/renderer/assets/app-icon.png',
  'apps/sheets/src/renderer/assets/app-icon.png',
  'apps/slides/src/renderer/assets/app-icon.png',
]
for (const target of pngTargets) writeFileSync(resolve(root, target), png1024)

function buildIco() {
  const sizes = [16, 24, 32, 48, 64, 128, 256]
  const images = sizes.map(renderPng)
  const headerSize = 6 + sizes.length * 16
  const header = Buffer.alloc(headerSize)
  header.writeUInt16LE(0, 0)
  header.writeUInt16LE(1, 2)
  header.writeUInt16LE(sizes.length, 4)
  let offset = headerSize
  sizes.forEach((size, index) => {
    const entry = 6 + index * 16
    header.writeUInt8(size === 256 ? 0 : size, entry)
    header.writeUInt8(size === 256 ? 0 : size, entry + 1)
    header.writeUInt8(0, entry + 2)
    header.writeUInt8(0, entry + 3)
    header.writeUInt16LE(1, entry + 4)
    header.writeUInt16LE(32, entry + 6)
    header.writeUInt32LE(images[index].length, entry + 8)
    header.writeUInt32LE(offset, entry + 12)
    offset += images[index].length
  })
  return Buffer.concat([header, ...images])
}

function buildIcns() {
  const entries = [
    ['ic07', renderPng(128)],
    ['ic08', renderPng(256)],
    ['ic09', renderPng(512)],
    ['ic10', png1024],
  ]
  const chunks = entries.map(([type, image]) => {
    const header = Buffer.alloc(8)
    header.write(type, 0, 4, 'ascii')
    header.writeUInt32BE(image.length + 8, 4)
    return Buffer.concat([header, image])
  })
  const body = Buffer.concat(chunks)
  const header = Buffer.alloc(8)
  header.write('icns', 0, 4, 'ascii')
  header.writeUInt32BE(body.length + 8, 4)
  return Buffer.concat([header, body])
}

writeFileSync(resolve(root, 'apps/shell/build/icon.ico'), buildIco())
writeFileSync(resolve(root, 'apps/shell/build/icon.icns'), buildIcns())

console.log(`WisWork icons generated from ${resolve(sourcePath)}`)
