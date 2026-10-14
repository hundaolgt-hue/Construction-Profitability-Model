const { chromium } = require('/opt/node22/lib/node_modules/playwright');
(async()=>{const b=await chromium.launch({executablePath:'/opt/pw-browsers/chromium-1194/chrome-linux/chrome'});const p=await b.newPage();
await p.goto('file://'+require('path').resolve('ty.html'));await p.evaluate(()=>document.fonts.ready);
await p.pdf({path:'../output/dist/Rapha_Thank_You.pdf',format:'A4',printBackground:true,margin:{top:0,bottom:0,left:0,right:0},pageRanges:'1'});await b.close();})();
