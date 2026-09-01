// Новые карточки на сайте: CSV для импорта + дозаливка фото/описания/коллекций через API.
//
// Почему два шага: POST /admin/products.json молча игнорирует option_values в любом виде
// (проверены option_values_attributes / option_values / option_value_ids, с option_names_attributes
// и без, id и title, а также PUT после создания) — товар создаётся, но варианты остаются без
// размера и цвета, и второй вариант падает с 422 «Варианты должны быть уникальными».
// Размеры и цвета умеет проставить только штатный импорт CSV. Всё остальное API отдаёт нормально.
//
//   node _create_cards.js --csv <арт> [<арт> ...]      -> import_new_cards.csv
//   node _create_cards.js --finish <арт> [<арт> ...]   -> фото + описание + коллекции на уже созданные
const https=require('https'), fs=require('fs'), path=require('path');
const AUTH='Basic '+Buffer.from('0be53397a378bea9f795b3525c71831e:928b5f36f68936151c69c7ae6854ca5d').toString('base64');
const HOST='elenason.myinsales.ru';
const PHOTOS='photos', BOM='\ufeff';
const MODE=process.argv.includes('--finish')?'finish':(process.argv.includes('--props')?'props':'csv');
const arts=process.argv.slice(2).filter(a=>/^\d+$/.test(a));
const sup=JSON.parse(fs.readFileSync('_supplier.json','utf8'));
const donor=['_kok_out3.json','_kok_out4.json'].filter(f=>fs.existsSync(f))
  .reduce((a,f)=>a.concat(JSON.parse(fs.readFileSync(f,'utf8'))),[]);
const sleep=ms=>new Promise(x=>setTimeout(x,ms));
const sizeMap={'XS':'XS(42)','S':'S(44)','M':'M(46)','L':'L(48)','XL':'XL(50)','2XL':'2XL(52)','3XL':'3XL(54)','4XL':'4XL(56)','5XL':'5XL(58)',
 'S/M':'S/M(44-46)','L/XL':'L/XL(48-50)','2XL/3XL':'2XL/3XL(52-54)','3XL/4XL':'3XL/4XL(54-56)','4XL/5XL':'4XL/5XL(56-58)','Free':'Free'};

function req(method,p,body){return new Promise(res=>{
  const data=body?Buffer.from(JSON.stringify(body),'utf8'):null;
  const r=https.request({host:HOST,path:p,method,timeout:120000,
    headers:Object.assign({Authorization:AUTH},data?{'Content-Type':'application/json','Content-Length':data.length}:{})},x=>{
    let d=''; x.on('data',c=>d+=c); x.on('end',()=>{let j=null;try{j=JSON.parse(d)}catch(e){} res({code:x.statusCode,json:j,body:d});});
  });
  r.on('timeout',()=>{r.destroy();res({code:0,body:'timeout'})});
  r.on('error',e=>res({code:0,body:String(e)}));
  if(data)r.write(data); r.end();
});}
async function reqR(method,p,body,ok=[200,201]){
  for(let i=0;i<4;i++){const r=await req(method,p,body); if(ok.includes(r.code)||r.code===422)return r; await sleep(2500*(i+1));}
  return {code:0,body:'failed after retries'};
}

// прайс "Пижама Mia-Amore 5862 \"Kanti\"" -> "Пижама Mia-Amore Kanti 5862"
function mkTitle(name){
  const type=(name.match(/^(\S+)/)||[])[1]||'';
  const num=(name.match(/\b(\d{3,5})\b/)||[])[1]||'';
  const model=(name.match(/"([^"]+)"/)||name.match(/[«"]([^»"]+)[»"]/)||[])[1]||'';
  return `${type} Mia-Amore ${model} ${num}`.replace(/\s+/g,' ').trim();
}
const typeOf=t=>(String(t).match(/^(\S+)/)||[])[1]||'';
const esc=s=>String(s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');

function rows(art){
  const base='ми'+art, s=sup[base];
  if(!s) return null;
  const title=mkTitle(s.name), out=[], skipped=[];
  for(const v of s.variants){
    const sz=sizeMap[v.size];
    if(!sz){ skipped.push(v.size); continue; }
    const price=Math.round(v.price*1.8);
    out.push({base, sku:base+' '+v.size.toLowerCase()+' '+v.color, title, color:v.color, size:sz,
      price, old:price*2, cost:v.price, stock:v.stock, barcode:String(v.barcode)});
  }
  return {base, title, rows:out, skipped, sizes:[...new Set(out.map(r=>r.size))]};
}

function buildDescription(art, sizes){
  const d=donor.find(x=>x.article===String(art));
  let raw=((d||{}).desc||'').trim();
  // донор иногда отдаёт вместо описания товара общий текст про бренд — такой не берём
  if(/Итальянская марка|Бренд miamia|по выгодной цене|интернет-магазин/i.test(raw)) raw='';
  const sizeLine='Размеры: '+sizes.join(', ');
  const paras=raw?raw.split(/\n+/).map(x=>x.trim()).filter(Boolean).map(x=>'<p>'+esc(x)+'</p>'):[];
  return {html:paras.join('\n')+(paras.length?'\n':'')+'<p>'+esc(sizeLine)+'</p>', fromDonor:!!raw};
}

function photosFor(art){
  if(!fs.existsSync(PHOTOS)) return [];
  return fs.readdirSync(PHOTOS).filter(f=>new RegExp('_'+art+'_\\d+\\.(png|jpe?g|webp)$','i').test(f))
    .sort((a,b)=>(+((a.match(/_(\d+)\./)||[])[1]||0))-(+((b.match(/_(\d+)\./)||[])[1]||0)));
}

// карточка того же типа/линейки, у которой копируем категорию и коллекции
function sibling(base, name, art, cards){
  const model=(name.match(/"([^"]+)"/)||[])[1]||'';
  const type=typeOf(name);
  // у пляжной линейки коллекции задаёт модельный ряд (Antigua целиком в «пляжной одежде»),
  // у домашней — тип изделия (сорочка и сарафан одной линейки лежат в разных разделах)
  const d=donor.find(x=>x.article===String(art));
  const beach=/\/beach\/|plyazh/i.test((d||{}).url||'');
  let best=null, bestScore=0;
  for(const c of cards){
    if((c.combos||[]).some(v=>String(v.sku).toLowerCase().startsWith(base))) continue;
    let s=0;
    if(type && typeOf(c.title).toLowerCase()===type.toLowerCase()) s+=beach?9:12;
    if(model && new RegExp('\\b'+model.replace(/[.*+?^${}()|[\]\\]/g,'\\$&')+'\\b','i').test(c.title)) s+=beach?12:9;
    if(s>bestScore){bestScore=s;best=c;}
  }
  return best;
}

async function doCsv(){
  // «Артикул варианта» импорт игнорирует — в sku варианта он пишет колонку «Артикул»,
  // поэтому артикул варианта кладём именно туда, иначе все строки одного товара
  // выглядят одним и тем же вариантом и импорт заводит только первую
  const header=['Название','Артикул','Цвет','Размер','Цена','Старая цена','Себестоимость','Остаток','Штрихкод'];
  const lines=[header.join(';')]; const summary=[];
  for(const art of arts){
    const r=rows(art);
    if(!r){ console.log('ми'+art+': нет в прайсе — пропуск'); continue; }
    for(const x of r.rows) lines.push([x.title,x.sku,x.color,x.size,x.price,x.old,x.cost,x.stock,x.barcode].join(';'));
    summary.push(r);
  }
  fs.writeFileSync('import_new_cards.csv', BOM+lines.join('\r\n')+'\r\n','utf8');
  console.log('import_new_cards.csv — строк (вариантов):', lines.length-1);
  for(const r of summary){
    console.log('  '+r.base+' | "'+r.title+'" | '+r.rows.length+' вар | цена '+r.rows[0].price+' | размеры '+r.sizes.join(', ')
      +(r.skipped.length?' | ПРОПУЩЕНЫ '+r.skipped.join(','):''));
  }
}

async function doFinish(){
  process.stderr.write('ищу карточки в каталоге: ');
  const want=new Set(arts.map(a=>'ми'+a));
  const found={}; const all=[];
  for(let page=1;page<=80 && Object.keys(found).length<want.size;page++){
    const r=await reqR('GET','/admin/products.json?per_page=250&page='+page);
    if(r.code!==200||!Array.isArray(r.json)||!r.json.length) break;
    for(const p of r.json){
      const mi=(p.variants||[]).filter(v=>String(v.sku||'').trim().toLowerCase().startsWith('ми'));
      if(!mi.length) continue;
      const base=String(mi[0].sku).trim().toLowerCase().split(/\s+/)[0];
      all.push({pid:p.id,title:p.title,combos:mi.map(v=>({sku:v.sku}))});
      if(want.has(base)&&!found[base]) found[base]=p;
    }
    process.stderr.write(page+' ');
    await sleep(300);
  }
  process.stderr.write('\n');
  const report=[];
  for(const art of arts){
    const base='ми'+art, p=found[base];
    if(!p){ console.log(base+': карточки на сайте нет — сначала импортируй CSV'); continue; }
    const r=rows(art);
    const desc=buildDescription(art, r.sizes);
    const sib=sibling(base, sup[base].name, art, all);
    let colls=[], cat=null;
    if(sib){ const t=await reqR('GET','/admin/products/'+sib.pid+'.json'); if(t.code===200){ colls=t.json.collections_ids||[]; cat=t.json.category_id; } }

    const upd={description:desc.html};
    if(cat && p.category_id!==cat) upd.category_id=cat;
    const ur=await reqR('PUT','/admin/products/'+p.id+'.json',{product:upd});

    let added=0;
    for(const cid of colls){
      if((p.collections_ids||[]).includes(cid)) continue;
      const cr=await reqR('POST','/admin/collects.json',{collect:{collection_id:cid,product_id:p.id}});
      if(cr.code===201) added++;
      await sleep(300);
    }

    const have=(p.images||[]).length;
    const pics=have?[]:photosFor(art);
    let up=0;
    for(const f of pics){
      const b=fs.readFileSync(path.join(PHOTOS,f));
      const ir=await reqR('POST','/admin/products/'+p.id+'/images.json',{image:{attachment:b.toString('base64'),filename:f}});
      if(ir.code===201||ir.code===200) up++; else console.log('   фото '+f+': '+ir.code+' '+String(ir.body).slice(0,120));
      await sleep(500);
    }
    console.log(base+' pid '+p.id+' | '+p.title+' | описание '+(ur.code===200?(desc.fromDonor?'с сайта-донора':'только размеры'):'ОШИБКА '+ur.code)
      +' | коллекций +'+added+(sib?' (по образцу «'+sib.title+'»)':'')+' | фото '+(have?'уже было '+have:up+'/'+pics.length));
    report.push({base,pid:p.id,title:p.title,desc:desc.fromDonor,collections:added,photos:up});
    await sleep(600);
  }
  fs.writeFileSync('_create_cards_report.json', JSON.stringify(report,null,1),'utf8');
}

// Характеристики (свойства) API не отдаёт: characteristics_attributes / characteristics /
// characteristic_ids молча игнорируются, POST /admin/products/{id}/characteristics.json — 404.
// Ставятся только импортом, колонками, и с включённой галкой «Добавлять/Удалять/Обновлять параметры».
async function doProps(){
  // «Цена» обязательна: без неё импорт не стартует («Не указано поле с ценами продажи»).
  // Значения те же, что уже стоят на сайте, поэтому перезапись ничего не меняет.
  const header=['Название','Артикул','Цена','Бренд','Состав','Цвет','SKU_HIDDEN'];
  const lines=[header.join(';')]; const summary=[];
  for(const art of arts){
    const r=rows(art);
    if(!r){ console.log('ми'+art+': нет в прайсе — пропуск'); continue; }
    const d=donor.find(x=>x.article===String(art));
    const sostav=((d||{}).sostav||'').trim();
    for(const x of r.rows) lines.push([x.title,x.sku,x.price,'Mia-Amore',sostav,x.color,x.base].join(';'));
    summary.push({base:r.base,title:r.title,n:r.rows.length,sostav,colors:[...new Set(r.rows.map(x=>x.color))]});
  }
  fs.writeFileSync('import_props.csv', BOM+lines.join('\r\n')+'\r\n','utf8');
  console.log('import_props.csv — строк:', lines.length-1);
  for(const s of summary) console.log('  '+s.base+' | состав: '+(s.sostav||'НЕТ — заполни руками')+' | цвет: '+s.colors.join('/'));
}

(MODE==='finish'?doFinish():MODE==='props'?doProps():doCsv()).catch(e=>{console.error(e);process.exit(1);});
