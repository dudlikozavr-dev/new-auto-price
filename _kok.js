// Product photo fetcher for kokete.ru / miamia.ru / braff.ru
// Usage: node _kok.js <outDir> <urls.txt | url1 url2 ...>
// For each product URL: title -> article + color, download unique full-size gallery photos
// to <outDir>/<site>_<article>_<n>.<ext>
const https=require('https'), fs=require('fs'), crypto=require('crypto'), path=require('path');
const outDir=process.argv[2];
let urls=process.argv.slice(3);
if(urls.length===1 && urls[0].endsWith('.txt')) urls=fs.readFileSync(urls[0],'utf8').split(/\r?\n/).map(s=>s.trim()).filter(Boolean);
fs.mkdirSync(outDir,{recursive:true});
const sleep=ms=>new Promise(x=>setTimeout(x,ms));

const SITES={
  'kokete.ru':  {key:'kok',    keep:/\/upload\/iblock\//i, byArticle:false},
  'miamia.ru':  {key:'miamia', keep:/\/upload\/iblock\//i, byArticle:true},
  'braff.ru':   {key:'braff',  keep:/\/images\/detailed\//i, byArticle:true},
};
const siteOf=h=>SITES[h.replace(/^www\./,'')]||{key:h.replace(/^www\./,'').split('.')[0], keep:/\/(?:upload|images|media)\//i, byArticle:false};

function fetch(u){return new Promise(res=>{
  const url=new URL(u);
  const req=https.get({host:url.host,path:url.pathname+url.search,rejectUnauthorized:false,
    headers:{'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64)','Referer':url.origin+'/','Accept':'*/*'},timeout:25000},r=>{
    if(r.statusCode>=300&&r.statusCode<400&&r.headers.location){r.resume();return res(fetch(new URL(r.headers.location,u).href));}
    const chunks=[]; r.on('data',c=>chunks.push(c)); r.on('end',()=>res({code:r.statusCode, buf:Buffer.concat(chunks)}));
  });
  req.on('timeout',()=>{req.destroy();res({code:0,buf:Buffer.alloc(0)});});
  req.on('error',()=>res({code:0,buf:Buffer.alloc(0)}));
});}
async function fetchR(u){for(let i=0;i<3;i++){const r=await fetch(u);if(r&&r.code===200&&r.buf.length>0)return r;await sleep(1500);}return{code:0,buf:Buffer.alloc(0)};}

const COLORS=['графит','молочный','кремовый','оранжевый','изумрудный','мультицвет','мультцвет','бежевый','розовый','чёрный','черный','белый','голубой','синий','пудра','бирюзовый','зелёный','зеленый','лиловый','коралловый','персиковый','серый','красный','бордовый','шоколадный','фиолетовый'];

(async()=>{
  const summary=[];
  for(const raw of urls){
    const u=raw.split('?')[0];
    const site=siteOf(new URL(u).host);
    const r=await fetchR(u);
    if(r.code!==200){ summary.push({url:u, err:'page fetch failed'}); continue; }
    const html=r.buf.toString('utf8');
    const title=(html.match(/<title>([^<]+)<\/title>/i)||[])[1]||'';
    const seg=(u.match(/\/(\d+)\/?$/)||[])[1]||'';
    const art=(/^\d{4}$/.test(seg)?seg:null)||(title.match(/\b(\d{4})\b/)||[])[1]||'????';
    const color=COLORS.find(c=>title.toLowerCase().includes(c))||'?';
    const descRaw=(html.match(/<meta itemprop="description" content="([\s\S]*?)"\s*\/?>/i)||html.match(/<meta name="description" content="([\s\S]*?)"\s*\/?>/i)||[])[1]||'';
    const desc=descRaw.replace(/&#(\d+);/g,(_,c)=>String.fromCharCode(+c)).replace(/&nbsp;/g,' ').replace(/&quot;/g,'"').replace(/&amp;/g,'&')
      .split(/\n\s*\n\s*Коллекци[яю]/)[0].replace(/^Только у нас\s*/,'').trim();
    const cands=new Set();
    for(const m of html.matchAll(/(?:src|data-src|data-large|href|content)=["']([^"']*\.(?:png|jpe?g|webp))["']/gi)) cands.add(m[1]);
    for(const m of html.matchAll(/\/(?:upload\/iblock|images\/detailed)\/[^"'\s<>)]+?\.(?:png|jpe?g|webp)/gi)) cands.add(m[0]);
    const imgs=[...cands].map(p=>new URL(p,u).href)
      .filter(p=>site.keep.test(p))
      .filter(p=>!/thumbnail|logo|icon|sprite|banner|pagespeed|no_photo|watermark|resize_cache/i.test(p))
      .filter(p=>!site.byArticle || art==='????' || decodeURIComponent(p).includes(art));
    // one product photo can be served in several renditions under the same file name — keep the biggest
    const best=new Map();
    for(const ip of imgs){
      const ir=await fetchR(ip);
      if(ir.code!==200||ir.buf.length<3000) continue;
      const key=decodeURIComponent(path.basename(new URL(ip).pathname)).toLowerCase();
      const prev=best.get(key);
      if(!prev||ir.buf.length>prev.buf.length) best.set(key,{buf:ir.buf, ext:ip.split('.').pop().toLowerCase()});
    }
    const saved=[]; const hashes=new Set(); let n=0;
    for(const {buf,ext} of best.values()){
      const h=crypto.createHash('md5').update(buf).digest('hex');
      if(hashes.has(h)) continue; hashes.add(h);
      n++; const fn=site.key+'_'+art+'_'+n+'.'+ext;
      fs.writeFileSync(path.join(outDir,fn), buf);
      saved.push({fn, bytes:buf.length});
    }
    summary.push({url:u, site:site.key, article:art, color, title:title.replace(/ купить.*$/,''), desc, photos:saved.length, files:saved.map(s=>s.fn)});
    process.stderr.write(site.key+':'+art+'('+saved.length+') ');
    await sleep(400);
  }
  console.log(JSON.stringify(summary,null,1));
})();
