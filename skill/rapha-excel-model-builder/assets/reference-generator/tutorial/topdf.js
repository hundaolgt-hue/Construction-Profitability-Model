// Render tutorial HTML to A4 PDF with Chromium (Playwright), with running header and page numbers.
// usage: node topdf.js in.html out.pdf
const { chromium } = require('/opt/node22/lib/node_modules/playwright');
(async () => {
  const [inp, out] = process.argv.slice(2);
  const browser = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium-1194/chrome-linux/chrome' }).catch(async () => chromium.launch());
  const page = await browser.newPage();
  await page.goto('file://' + require('path').resolve(inp), { waitUntil: 'networkidle' });
  await page.evaluate(() => document.fonts.ready);
  const hdr = `<div style="width:100%;font-family:'Noto Sans Ethiopic';font-size:7.5px;color:#5B6675;padding:0 16mm;display:flex;justify-content:space-between;">
      <span>የራፋ የግንባታ ፕሮጀክት ትርፋማነት ሞዴል — የአጠቃቀም መመሪያ</span><span style="color:#D98E2B">Rapha Development Profitability Model</span></div>`;
  const ftr = `<div style="width:100%;font-family:'Noto Sans Ethiopic';font-size:7.5px;color:#5B6675;padding:0 16mm;display:flex;justify-content:space-between;">
      <span>የዕቅድ መሣሪያ ብቻ — የኢንቨስትመንት፣ የግብር ወይም የሕግ ምክር አይደለም</span><span>ገጽ <span class="pageNumber"></span> / <span class="totalPages"></span></span></div>`;
  await page.pdf({ path: out, format: 'A4', printBackground: true, displayHeaderFooter: true,
    headerTemplate: hdr, footerTemplate: ftr, margin: { top: '16mm', bottom: '16mm', left: '0mm', right: '0mm' } });
  await browser.close();
})();
