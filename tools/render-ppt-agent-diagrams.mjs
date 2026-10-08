import fs from 'node:fs/promises'
import path from 'node:path'
import { chromium } from 'playwright'

const root = process.cwd()
const assetDir = path.join(root, 'docs/product/assets')
const tempDir = path.join('/tmp', 'wiswork-ppt-agent-diagrams')
const files = (await fs.readdir(assetDir)).filter((file) => file.endsWith('.svg'))

await fs.mkdir(tempDir, { recursive: true })

const browser = await chromium.launch({
  executablePath: '/usr/bin/google-chrome',
  headless: true,
  args: ['--no-sandbox', '--disable-gpu'],
})

try {
  const page = await browser.newPage({ viewport: { width: 1600, height: 1200 } })
  for (const file of files) {
    const svgPath = path.join(assetDir, file)
    const rawPath = path.join(tempDir, file.replace(/\.svg$/, '-raw.png'))
    await page.goto(`file://${svgPath}`)
    await page.locator('svg').screenshot({ path: rawPath })
    console.log(`${rawPath}\t${path.join(assetDir, file.replace(/\.svg$/, '.png'))}`)
  }
} finally {
  await browser.close()
}
