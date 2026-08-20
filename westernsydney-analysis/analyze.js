/* Headless capture + AEM block/DOM extraction for Western Sydney University pages.
 * Screenshots -> ./screenshots ; structural JSON -> ./page-structures.json */
const { chromium } = require('playwright');
const fs = require('fs');
const path = require('path');

const OUT = path.join(__dirname, 'screenshots');
fs.mkdirSync(OUT, { recursive: true });

const PAGES = [
  ['home',        'https://www.westernsydney.edu.au/'],
  ['news-list',   'https://www.westernsydney.edu.au/news-centre/stories/2025'],
  ['news-story',  'https://www.westernsydney.edu.au/news-centre/stories/2025/chancellors-address-women-of-western-sydney-awards-2025'],
  ['expert-op',   'https://www.westernsydney.edu.au/news-centre/expert-opinion/2013-whitlam-oration-delivered-by-noel-pearson'],
  ['institute',   'https://www.westernsydney.edu.au/ics'],
  ['inst-about',  'https://www.westernsydney.edu.au/ics/about-us'],
  ['school',      'https://www.westernsydney.edu.au/schools/grs'],
  ['events-list', 'https://www.westernsydney.edu.au/babylab/events/previous-events'],
  ['people',      'https://www.westernsydney.edu.au/challenging-racism-project/our-people'],
  ['research',    'https://www.westernsydney.edu.au/babylab/research'],
  ['content-deep','https://www.westernsydney.edu.au/inherent-requirements/bachelor-of-applied-leadership-and-critical-thinking'],
  ['news-centre', 'https://www.westernsydney.edu.au/news-centre'],
];

const UA = 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

(async () => {
  const browser = await chromium.launch({ args: ['--no-sandbox', '--disable-dev-shm-usage'] });
  const ctx = await browser.newContext({
    viewport: { width: 1440, height: 1000 }, userAgent: UA, ignoreHTTPSErrors: true, locale: 'en-AU',
  });
  const results = [];
  for (const [key, url] of PAGES) {
    const page = await ctx.newPage();
    const rec = { key, url, ok: false, error: null };
    try {
      const resp = await page.goto(url, { waitUntil: 'domcontentloaded', timeout: 45000 });
      rec.status = resp ? resp.status() : null;
      await page.waitForTimeout(3000);
      // dismiss cookie banners
      try { await page.evaluate(() => {
        for (const s of ['#onetrust-accept-btn-handler','[aria-label*="accept" i] button','button[title*="Accept" i]','.cookie button']) {
          const el = document.querySelector(s); if (el) { el.click(); break; }
        }
      }); } catch (e) {}
      await page.waitForTimeout(500);
      // extract structure
      rec.data = await page.evaluate(() => {
        const uniq = (a) => Array.from(new Set(a));
        // AEM Core Component types via data-cmp-data-layer or class cmp-*
        const cmpClasses = uniq(Array.from(document.querySelectorAll('[class*="cmp-"]'))
          .flatMap(el => Array.from(el.classList)).filter(c => c.startsWith('cmp-'))
          .map(c => c.split('__')[0])).sort();
        const dataLayer = uniq(Array.from(document.querySelectorAll('[data-cmp-data-layer]'))
          .map(el => { try { return Object.values(JSON.parse(el.getAttribute('data-cmp-data-layer')))[0]?.['@type']; } catch (e) { return null; } })
          .filter(Boolean)).sort();
        // generic block hints
        const sectionClasses = uniq(Array.from(document.querySelectorAll('main *[class]'))
          .flatMap(el => Array.from(el.classList))
          .filter(c => /(hero|banner|card|teaser|carousel|accordion|tab|list|cta|nav|footer|header|search|breadcrumb|grid|feature|quote|gallery|video|form|promo|tile|slider)/i.test(c)))
          .slice(0, 60).sort();
        const iframes = uniq(Array.from(document.querySelectorAll('iframe')).map(f => { try { return new URL(f.src).host; } catch(e){ return f.src; } }).filter(Boolean));
        const scripts = uniq(Array.from(document.querySelectorAll('script[src]')).map(s => { try { return new URL(s.src).host; } catch(e){ return null; } }).filter(Boolean));
        const forms = document.querySelectorAll('form').length;
        const h1 = document.querySelector('h1') ? document.querySelector('h1').innerText.slice(0, 90) : null;
        const title = document.title;
        return { title, h1, cmpClasses, dataLayer, sectionClasses, iframes, scripts, forms,
                 headerNav: !!document.querySelector('header nav, nav[aria-label]'),
                 breadcrumb: !!document.querySelector('[class*="breadcrumb" i], nav[aria-label*="readcrumb"]') };
      });
      await page.screenshot({ path: path.join(OUT, `${key}_top.png`), fullPage: false });
      // full page (bounded)
      await page.evaluate(async () => { await new Promise(r => { let y=0; const t=setInterval(()=>{window.scrollBy(0,900);y+=900;if(y>=document.body.scrollHeight||y>10000){clearInterval(t);r();}},110);}); });
      await page.waitForTimeout(700); await page.evaluate(() => window.scrollTo(0,0)); await page.waitForTimeout(400);
      await page.screenshot({ path: path.join(OUT, `${key}_full.png`), fullPage: true });
      rec.ok = true;
    } catch (e) { rec.error = String(e).split('\n')[0].slice(0,160); }
    await page.close();
    results.push(rec);
    console.log(`${rec.ok?'OK':'  '} ${key.padEnd(13)} ${rec.status||''} cmp=${rec.data?rec.data.cmpClasses.length:'-'} sect=${rec.data?rec.data.sectionClasses.length:'-'} ${rec.error||''}`);
  }
  await browser.close();
  fs.writeFileSync(path.join(__dirname, 'page-structures.json'), JSON.stringify(results, null, 2));
  console.log('\nDONE. structures -> page-structures.json');
})();
