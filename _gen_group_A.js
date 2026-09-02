// group_A_dozavesti.csv из _split.json — сводка «что дозавести в существующие карточки».
//
// ЗАЧЕМ ОТДЕЛЬНЫЙ СКРИПТ. Файл от 27.07 собирали разово и генератор не сохранили,
// а пересобирать приходится регулярно: импорт группы A ищет товар ПО НАЗВАНИЮ,
// а названия карточек в каталоге меняются. Проверка 02.09.2026: из 102 строк файла
// от 27.07 разошлись с живым каталогом 76 (75%). Такой файл импорт молча не найдёт.
//
// Порядок: node _scan_catalog.js -> node _split.js -> node _gen_group_A.js
const fs = require('fs');
const { groupA } = JSON.parse(fs.readFileSync('_split.json', 'utf8'));

// размеры в понятном порядке, а не как пришли от поставщика
const PORYADOK = ['XS','S','M','L','XL','2XL','3XL','4XL','5XL',
                  'S/M','L/XL','2XL/3XL','3XL/4XL','4XL/5XL'];
const poRazmeru = (a, b) => {
  const ia = PORYADOK.indexOf(a), ib = PORYADOK.indexOf(b);
  return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || String(a).localeCompare(b);
};

const BOM = '﻿';
const shapka = ['Артикул','Название','pid','Цвета на сайте','Добавить цвета',
                'Размеры добавить','Вариантов добавить'];
const stroki = [shapka.join(';')];

for (const g of groupA) {
  const razmery = [...new Set(g.missing.map(v => v.size))].sort(poRazmeru);
  stroki.push([
    g.base,
    // название берём с САЙТА, а не из прайса: по нему импорт и матчит карточку
    String(g.title || '').replace(/;/g, ','),
    (g.pids || []).join(' '),
    (g.siteColors || []).join('/'),
    (g.missingColors || []).join('/'),
    razmery.join('/'),
    g.missingCount,
  ].join(';'));
}

fs.writeFileSync('group_A_dozavesti.csv', BOM + stroki.join('\r\n') + '\r\n', 'utf8');
const vsego = groupA.reduce((s, x) => s + x.missingCount, 0);
console.log(`group_A_dozavesti.csv: ${groupA.length} карточек, ${vsego} вариантов дозавести`);

// карточек без названия на сайте быть не должно — если есть, матчить будет нечем
const bezImeni = groupA.filter(x => !x.title);
if (bezImeni.length) console.log('WARN без названия:', bezImeni.map(x => x.base).join(', '));

// ⚠ у части баз карточек НЕСКОЛЬКО (раскладка по цветам). Название берётся из первой,
// и импорт положит новый цвет именно в неё — а это может быть карточка другого цвета
// (ми4592 «…Lilirose черный» + светло-бежевый). Такие строки глазами, до импорта.
const mnogoKart = groupA.filter(x => (x.pids || []).length > 1);
if (mnogoKart.length) {
  console.log(`\n⚠ карточек на базу больше одной — ${mnogoKart.length}, проверить вручную:`);
  mnogoKart.forEach(x => console.log(
    `  ${x.base} | ${x.pids.length} карточек | берём «${x.title}» | добавить: ${(x.missingColors||[]).join('/')}`));
}
