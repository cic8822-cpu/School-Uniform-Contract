const { chromium } = require('C:/Users/최익창/AppData/Roaming/npm/node_modules/playwright');
const path = require('path');

(async () => {
  const browser = await chromium.launch();
  const page = await browser.newPage({ viewport: { width: 1160, height: 1000 }, deviceScaleFactor: 2 });
  const fileUrl = 'file:///' + path.resolve(__dirname, 'uniform-contract-cardnews.html').replace(/\\/g, '/');
  await page.goto(fileUrl, { waitUntil: 'networkidle' });
  await page.waitForTimeout(300);
  await page.evaluate(() => {
    const nav = document.querySelector('.quicknav');
    if (nav) nav.style.display = 'none';
  });

  const targets = [
    { sel: '#m-web .poster', out: 'method-1-web.png' },
    { sel: '#m-exe .poster', out: 'method-2-exe.png' },
    { sel: '#m-xl .poster', out: 'method-3-excel.png' },
  ];
  for (const t of targets) {
    const el = await page.$(t.sel);
    await el.screenshot({ path: path.resolve(__dirname, 'images', t.out) });
    console.log('saved', t.out);
  }

  // also a full-page hero shot for the README top
  await page.setViewportSize({ width: 1160, height: 900 });
  const hero = await page.$('.hero');
  await hero.screenshot({ path: path.resolve(__dirname, 'images', 'hero.png') });
  console.log('saved hero.png');

  await browser.close();
})();
