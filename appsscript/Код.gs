var API_KEY = 'c044b7f6f791cb35307817cbc4cb5cd6';
var API_PASS = '9b14cd1c4d5aadccffc30edf211d5203';
var DOMAIN = 'elenason.myinsales.ru';

function getAuth() {
  return Utilities.base64Encode(API_KEY + ':' + API_PASS);
}

function apiGet(path) {
  var response = UrlFetchApp.fetch('https://' + DOMAIN + path, {
    method: 'get',
    headers: { 'Authorization': 'Basic ' + getAuth() },
    muteHttpExceptions: true
  });
  return { code: response.getResponseCode(), body: response.getContentText() };
}

function apiPut(path, payload) {
  var response = UrlFetchApp.fetch('https://' + DOMAIN + path, {
    method: 'put',
    contentType: 'application/json',
    headers: { 'Authorization': 'Basic ' + getAuth() },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  return { code: response.getResponseCode(), body: response.getContentText() };
}

// Автоматическое обновление прайса поставщика по ссылке
function обновитьПрайсПоставщика() {
  var url = 'http://miamia.ru/1c/ostatki_Platina.xlsx';

  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (response.getResponseCode() !== 200) {
    SpreadsheetApp.getUi().alert('Ошибка загрузки файла: ' + response.getResponseCode());
    return;
  }

  var blob = response.getBlob().setName('ostatki-Platina.xlsx');
  var xlsxFile = DriveApp.createFile(blob);
  var converted = Drive.Files.copy(
    { name: '_temp_platina_auto', mimeType: 'application/vnd.google-apps.spreadsheet' },
    xlsxFile.getId()
  );
  xlsxFile.setTrashed(true);

  var tempSS = SpreadsheetApp.openById(converted.id);
  var sheet = tempSS.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(8, 1, lastRow - 7, 18).getValues();

  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var articul = row[0];
    var cena = row[16];
    if (!articul || parseFloat(String(cena).replace(',', '.')) == 0) continue;
    result.push([articul, row[3], row[13], row[14], cena, row[17]]);
  }

  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Поставщик');
  if (destSheet.getLastRow() > 1) {
    destSheet.getRange(2, 1, destSheet.getLastRow() - 1, 6).clearContent();
  }
  if (result.length > 0) {
    destSheet.getRange(2, 1, result.length, 6).setValues(result);
  }

  Drive.Files.remove(converted.id);

  обновитьСводную();
}

// ШАГ 1: Конвертировать прайс поставщика
function шаг1_Конвертировать() {
  var files = DriveApp.getFilesByName('ostatki-Platina.xlsx');
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert('Файл не найден на Drive');
    return;
  }
  var fileId = files.next().getId();
  var tempFile = Drive.Files.copy(
    { name: '_temp_platina', mimeType: 'application/vnd.google-apps.spreadsheet' },
    fileId
  );
  PropertiesService.getScriptProperties().setProperty('tempFileId', tempFile.id);
  SpreadsheetApp.getUi().alert('Конвертация готова. Теперь запусти шаг2_Импорт');
}

// ШАГ 2: Импорт прайса поставщика
function шаг2_Импорт() {
  var tempId = PropertiesService.getScriptProperties().getProperty('tempFileId');
  if (!tempId) {
    SpreadsheetApp.getUi().alert('Сначала запусти шаг1_Конвертировать');
    return;
  }

  var tempSS = SpreadsheetApp.openById(tempId);
  var sheet = tempSS.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(8, 1, lastRow - 7, 18).getValues();

  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var articul = row[0];
    var cena = row[16];
    if (!articul || parseFloat(String(cena).replace(',', '.')) == 0) continue;
    result.push([articul, row[3], row[13], row[14], cena, row[17]]);
  }

  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Поставщик');
  if (destSheet.getLastRow() > 1) {
    destSheet.getRange(2, 1, destSheet.getLastRow() - 1, 6).clearContent();
  }
  if (result.length > 0) {
    destSheet.getRange(2, 1, result.length, 6).setValues(result);
  }

  Drive.Files.remove(tempId);
  PropertiesService.getScriptProperties().deleteProperty('tempFileId');
  SpreadsheetApp.getUi().alert('Загружено строк: ' + result.length);
}

// ШАГ 1: Проверить наличие файла на Drive (конвертация не нужна для CSV)
function шаг1_КонвертироватьМойПрайс() {
  var files = DriveApp.searchFiles('title contains "shop_products_prices_and_stocks" and trashed = false');
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert('Файл shop_products_prices_and_stocks*.csv не найден на Drive.\nЗагрузи файл на Google Drive и повтори.');
    return;
  }
  SpreadsheetApp.getUi().alert('Файл найден. Запусти шаг2_ИмпортМойПрайс');
}

// ШАГ 2: Импорт моего прайса из CSV (InSales выгрузка)
function шаг2_ИмпортМойПрайс() {
  var files = DriveApp.searchFiles('title contains "shop_products_prices_and_stocks" and trashed = false');
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert('Файл shop_products_prices_and_stocks*.csv не найден на Drive');
    return;
  }

  var file = files.next();
  while (files.hasNext()) {
    var next = files.next();
    if (next.getDateCreated() > file.getDateCreated()) file = next;
  }

  var content = file.getBlob().getDataAsString('UTF-16LE');
  if (content.charCodeAt(0) === 0xFEFF) content = content.slice(1);

  var lines = content.split('\n');
  var result = [];

  for (var i = 1; i < lines.length; i++) {
    var cols = lines[i].split('\t');
    if (cols.length < 9) continue;
    var articul = cols[2].trim();
    if (!articul) continue;
    result.push([
      '',              // A: ID_товара (нет в CSV)
      cols[0].trim(),  // B: ID_варианта
      articul,         // C: Артикул
      '',              // D: Штрихкод (нет в CSV)
      cols[3].trim(),  // E: Цена_продажи
      cols[8].trim()   // F: Остаток
    ]);
  }

  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Мой прайс');
  if (destSheet.getLastRow() > 1) {
    destSheet.getRange(2, 1, destSheet.getLastRow() - 1, 6).clearContent();
  }
  if (result.length > 0) {
    destSheet.getRange(2, 1, result.length, 6).setValues(result);
  }
  SpreadsheetApp.getUi().alert('Загружено строк: ' + result.length);
}

// Обновить Сводную и Новые товары
function обновитьСводную() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var suppSheet = ss.getSheetByName('Поставщик');
  var mySheet = ss.getSheetByName('Мой прайс');
  var svodSheet = ss.getSheetByName('Сводная');
  var novSheet = ss.getSheetByName('Новые товары');

  // Поставщик: A=артикул, B=название, C=размер+цвет, D=штрихкод, E=цена, F=остаток
  var suppLastRow = suppSheet.getLastRow();
  if (suppLastRow < 2) { SpreadsheetApp.getUi().alert('Лист Поставщик пуст'); return; }
  var suppData = suppSheet.getRange(2, 1, suppLastRow - 1, 6).getValues();

  // Мой прайс: A=ID_товара, B=ID_варианта, C=артикул, D=штрихкод, E=цена_продажи, F=остаток
  var myLastRow = mySheet.getLastRow();
  if (myLastRow < 2) { SpreadsheetApp.getUi().alert('Лист Мой прайс пуст'); return; }
  var myData = mySheet.getRange(2, 1, myLastRow - 1, 6).getValues();

  // Строим карты поставщика: базовый артикул → массив индексов; штрихкод → индекс
  var suppByBase = {};
  var suppByBarcode = {};
  for (var i = 0; i < suppData.length; i++) {
    var sArt = String(suppData[i][0]).trim().toLowerCase();
    var sBase = sArt.split(' ')[0];
    var sBarcode = String(suppData[i][3]).trim();
    if (!suppByBase[sBase]) suppByBase[sBase] = [];
    suppByBase[sBase].push(i);
    if (sBarcode && sBarcode !== '0' && sBarcode !== '') suppByBarcode[sBarcode] = i;
  }

  var suppMatched = {};
  var svodRows = [];

  for (var j = 0; j < myData.length; j++) {
    var myRow = myData[j];
    var myArt = String(myRow[2]).trim();
    var myBase = myArt.toLowerCase().split(' ')[0];
    var myBarcode = String(myRow[3]).trim();

    var matchIdx = null;

    // Сначала по базовому артикулу
    if (suppByBase[myBase]) {
      // Уточнение по штрихкоду если есть
      if (myBarcode && myBarcode !== '0' && myBarcode !== '') {
        for (var k = 0; k < suppByBase[myBase].length; k++) {
          var candIdx = suppByBase[myBase][k];
          if (String(suppData[candIdx][3]).trim() === myBarcode) {
            matchIdx = candIdx;
            break;
          }
        }
      }
      // Если по штрихкоду не нашли — берём первый по базовому артикулу
      if (matchIdx === null) matchIdx = suppByBase[myBase][0];
    }

    // Если по артикулу не нашли — ищем по штрихкоду
    if (matchIdx === null && myBarcode && myBarcode !== '0' && myBarcode !== '') {
      if (suppByBarcode[myBarcode] !== undefined) matchIdx = suppByBarcode[myBarcode];
    }

    if (matchIdx !== null) {
      suppMatched[matchIdx] = true;
      var sr = suppData[matchIdx];
      svodRows.push([
        myArt,       // A: Артикул мой
        myBase,      // B: Базовый артикул
        sr[1],       // C: Название
        sr[2],       // D: Размер+цвет
        sr[3],       // E: Штрихкод поставщика
        sr[4],       // F: Цена поставщика
        'Совпал',    // G: Статус
        myRow[4],    // H: Старая цена
        myRow[4],    // I: Цена продажи
        sr[5],       // J: Остаток поставщика
        myBarcode,   // K: Штрихкод мой
        myRow[1],    // L: ID_варианта
        myRow[0]     // M: ID_товара
      ]);
    } else {
      svodRows.push([
        myArt,              // A
        myBase,             // B
        '', '', '', '',     // C-F
        'Нет у поставщика', // G
        myRow[4],           // H: Старая цена
        myRow[4],           // I: Цена продажи
        0,                  // J: Остаток = 0
        myBarcode,          // K
        myRow[1],           // L
        myRow[0]            // M
      ]);
    }
  }

  // Записываем Сводную
  svodSheet.clearContents();
  svodSheet.getRange(1, 1, 1, 13).setValues([[
    'Артикул', 'Базовый арт.', 'Название', 'Размер+цвет',
    'Штрихкод пост.', 'Цена пост.', 'Статус', 'Старая цена',
    'Цена продажи', 'Остаток', 'Штрихкод мой', 'ID_варианта', 'ID_товара'
  ]]);
  if (svodRows.length > 0) {
    svodSheet.getRange(2, 1, svodRows.length, 13).setValues(svodRows);
  }

  // Записываем Новые товары
  var novRows = [];
  for (var s = 0; s < suppData.length; s++) {
    if (!suppMatched[s]) {
      var nr = suppData[s];
      novRows.push([nr[0], nr[1], nr[2], nr[3], nr[4], nr[5]]);
    }
  }
  novSheet.clearContents();
  novSheet.getRange(1, 1, 1, 6).setValues([['Артикул', 'Название', 'Размер+цвет', 'Штрихкод', 'Цена пост.', 'Остаток']]);
  if (novRows.length > 0) {
    novSheet.getRange(2, 1, novRows.length, 6).setValues(novRows);
    novSheet.getRange(2, 1, novRows.length, 6).setBackground('#FFCCCC');
  }

  var совпало = svodRows.filter(function(r) { return r[6] === 'Совпал'; }).length;
  var нетУПост = svodRows.filter(function(r) { return r[6] === 'Нет у поставщика'; }).length;
  SpreadsheetApp.getUi().alert(
    'Готово!\n' +
    'Совпало: ' + совпало + '\n' +
    'Нет у поставщика (остаток обнулён): ' + нетУПост + '\n' +
    'Новые товары поставщика: ' + novRows.length
  );
}

// ТЕСТ: Выгрузка 5 строк в InSales по ID варианта
function тестВыгрузки5строк() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Сводная');
  var data = sheet.getDataRange().getValues();

  var toUpdate = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][6] === 'Совпал' && toUpdate.length < 5) {
      toUpdate.push({
        variantId: data[i][11],  // L: ID_варианта
        sku: data[i][0],          // A: Артикул
        price: data[i][8],        // I: Цена_продажи
        stock: data[i][9]         // J: Остаток
      });
    }
  }

  var results = [];
  for (var j = 0; j < toUpdate.length; j++) {
    var item = toUpdate[j];
    if (!item.variantId) {
      results.push(item.sku + ': нет ID варианта');
      continue;
    }

    var getResp = apiGet('/admin/products/variants/' + item.variantId + '.json');
    if (getResp.code !== 200) {
      results.push(item.sku + ': вариант не найден (' + getResp.code + ')');
      Utilities.sleep(300);
      continue;
    }
    var productId = JSON.parse(getResp.body).product_id;

    var putResp = apiPut(
      '/admin/products/' + productId + '/variants/' + item.variantId + '.json',
      { variant: { price: item.price, quantity: item.stock } }
    );

    var code = putResp.code;
    results.push(item.sku + ': ' + code + (code === 200 ? ' OK' : ' ОШИБКА'));
    Utilities.sleep(300);
  }

  SpreadsheetApp.getUi().alert(results.join('\n'));
}
