// JS 描画が必要な IR ページ用：headless Chromium で描画後の HTML を標準出力へ
// usage: node render.js <url>
const { chromium } = require('playwright');
(async () => {
  const browser = await chromium.launch();
  try {
    const page = await browser.newPage({ userAgent: 'Mozilla/5.0 (research crawler; financial-institution disclosure survey)' });
    await page.goto(process.argv[2], { waitUntil: 'networkidle', timeout: 45000 });
    await page.waitForTimeout(1500);
    process.stdout.write(await page.content());
  } catch (e) {
    process.exitCode = 1;
  } finally {
    await browser.close();
  }
})();
