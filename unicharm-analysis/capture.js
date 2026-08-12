/* Headless screenshot capture for Unicharm AEM migration analysis.
 * Captures an above-the-fold (_top) and a full-page (_full) PNG per site.
 * Saves to ./screenshots and writes a results log to ./capture-results.json.
 * Screenshots stay on disk — never loaded into the agent context. */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const OUT = path.join(__dirname, 'screenshots');
const sites = require('./sites.json').sites;
const CONCURRENCY = 4;
const NAV_TIMEOUT = 45000;

fs.mkdirSync(OUT, { recursive: true });

async function capture(browser, site) {
  const ctx = await browser.newContext({
    viewport: { width: 1440, height: 1000 },
    userAgent:
      'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36',
    ignoreHTTPSErrors: true,
    locale: 'en-US',
  });
  const page = await ctx.newPage();
  const res = { id: site.id, url: site.url, brand: site.brand, top: false, full: false, error: null, status: null, finalUrl: null };
  try {
    const resp = await page.goto(site.url, { waitUntil: 'domcontentloaded', timeout: NAV_TIMEOUT });
    res.status = resp ? resp.status() : null;
    // let lazy content / carousels settle; dismiss obvious cookie/consent overlays
    await page.waitForTimeout(3500);
    try {
      await page.evaluate(() => {
        const sel = ['[id*="cookie" i] button', '[class*="cookie" i] button', '[id*="consent" i] button', '.gdpr button', '#onetrust-accept-btn-handler'];
        for (const s of sel) { const el = document.querySelector(s); if (el) { el.click(); break; } }
      });
    } catch (e) { /* ignore */ }
    await page.waitForTimeout(800);
    res.finalUrl = page.url();
    // above the fold
    await page.screenshot({ path: path.join(OUT, `${site.id}_top.png`), fullPage: false });
    res.top = true;
    // full page (cap height by scrolling to trigger lazy images first)
    await page.evaluate(async () => {
      await new Promise((resolve) => {
        let y = 0; const step = 800;
        const t = setInterval(() => {
          window.scrollBy(0, step); y += step;
          if (y >= document.body.scrollHeight || y > 12000) { clearInterval(t); resolve(); }
        }, 120);
      });
    });
    await page.waitForTimeout(1000);
    await page.evaluate(() => window.scrollTo(0, 0));
    await page.waitForTimeout(500);
    await page.screenshot({ path: path.join(OUT, `${site.id}_full.png`), fullPage: true });
    res.full = true;
  } catch (e) {
    res.error = String(e).split('\n')[0].slice(0, 200);
  } finally {
    await ctx.close();
  }
  console.log(`${res.top ? 'OK ' : '   '}${res.full ? 'F' : ' '} ${site.id} ${site.brand.padEnd(20)} ${res.status || ''} ${res.error || ''}`);
  return res;
}

(async () => {
  const browser = await chromium.launch({ args: ['--no-sandbox', '--disable-dev-shm-usage'] });
  const results = [];
  const queue = [...sites];
  async function worker() {
    while (queue.length) {
      const site = queue.shift();
      results.push(await capture(browser, site));
    }
  }
  await Promise.all(Array.from({ length: CONCURRENCY }, worker));
  await browser.close();
  fs.writeFileSync(path.join(__dirname, 'capture-results.json'), JSON.stringify(results, null, 2));
  const ok = results.filter((r) => r.top).length;
  const full = results.filter((r) => r.full).length;
  const failed = results.filter((r) => !r.top);
  console.log(`\nDONE: ${ok}/${sites.length} top, ${full}/${sites.length} full.`);
  if (failed.length) console.log('FAILED:', failed.map((f) => `${f.id}(${f.error || f.status})`).join(', '));
})();
