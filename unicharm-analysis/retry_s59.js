const { chromium } = require('playwright');
const path = require('path');
(async () => {
  const browser = await chromium.launch({
    args: ['--no-sandbox','--disable-dev-shm-usage','--ignore-certificate-errors','--ignore-ssl-errors'],
  });
  const ctx = await browser.newContext({ viewport:{width:1440,height:1000}, ignoreHTTPSErrors:true,
    userAgent:'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36' });
  const page = await ctx.newPage();
  const targets = ['https://www.babyjoy.com.eg/ar/home.html','http://www.babyjoy.com.eg/','https://babyjoy.com.eg/ar/home.html'];
  for (const t of targets) {
    try {
      const r = await page.goto(t, { waitUntil:'domcontentloaded', timeout:40000 });
      console.log('OK', t, r && r.status(), '->', page.url());
      await page.waitForTimeout(3500);
      await page.screenshot({ path: path.join(__dirname,'screenshots','s59_top.png'), fullPage:false });
      await page.evaluate(async()=>{await new Promise(res=>{let y=0;const i=setInterval(()=>{window.scrollBy(0,800);y+=800;if(y>=document.body.scrollHeight||y>12000){clearInterval(i);res();}},120);});});
      await page.waitForTimeout(1000); await page.evaluate(()=>window.scrollTo(0,0)); await page.waitForTimeout(500);
      await page.screenshot({ path: path.join(__dirname,'screenshots','s59_full.png'), fullPage:true });
      console.log('CAPTURED via', t);
      break;
    } catch(e){ console.log('FAIL', t, String(e).split('\n')[0].slice(0,120)); }
  }
  await browser.close();
})();
