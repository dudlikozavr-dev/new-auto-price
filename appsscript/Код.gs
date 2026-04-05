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

// ШАГ 1: Конвертировать mi_products CSV в Google Sheets на Drive
function шаг1_КонвертироватьМойПрайс() {
  var files = DriveApp.searchFiles('title contains "mi_products" and trashed = false');
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert('Файл mi_products*.csv не найден на Drive.\nЗагрузи файл на Google Drive и повтори.');
    return;
  }
  var file = files.next();
  while (files.hasNext()) {
    var next = files.next();
    if (next.getDateCreated() > file.getDateCreated()) file = next;
  }

  var tempFile = Drive.Files.copy(
    { name: '_temp_mystore', mimeType: 'application/vnd.google-apps.spreadsheet' },
    file.getId()
  );
  PropertiesService.getScriptProperties().setProperty('tempMyStoreId', tempFile.id);
  SpreadsheetApp.getUi().alert('Конвертация готова (' + file.getName() + ').\nТеперь запусти шаг2_ИмпортМойПрайс');
}

// ШАГ 2: Импорт моего прайса из сконвертированного файла
function шаг2_ИмпортМойПрайс() {
  var tempId = PropertiesService.getScriptProperties().getProperty('tempMyStoreId');
  if (!tempId) {
    SpreadsheetApp.getUi().alert('Сначала запусти шаг1_КонвертироватьМойПрайс');
    return;
  }

  var tempSS = SpreadsheetApp.openById(tempId);
  var sheet = tempSS.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var articul = String(row[2]).trim();
    if (!articul) continue;
    result.push([
      '',                     // A: ID_товара (нет в CSV)
      String(row[0]).trim(),  // B: ID_варианта
      articul,                // C: Артикул
      '',                     // D: Штрихкод (нет в CSV)
      row[3],                 // E: Цена_продажи
      row[8]                  // F: Остаток
    ]);
  }

  Drive.Files.remove(tempId);
  PropertiesService.getScriptProperties().deleteProperty('tempMyStoreId');

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
  var suppByFullArt = {};
  for (var i = 0; i < suppData.length; i++) {
    var sArt = String(suppData[i][0]).trim().toLowerCase();
    var sBase = sArt.split(' ')[0];
    var sSizeColor = String(suppData[i][2]).trim().toLowerCase();
    var sBarcode = String(suppData[i][3]).trim();
    if (!suppByBase[sBase]) suppByBase[sBase] = [];
    suppByBase[sBase].push(i);
    if (sBarcode && sBarcode !== '0' && sBarcode !== '') suppByBarcode[sBarcode] = i;
    suppByFullArt[sArt] = i;
    if (sSizeColor) suppByFullArt[sBase + ' ' + sSizeColor] = i;
  }

  var suppMatched = {};
  var svodRows = [];

  for (var j = 0; j < myData.length; j++) {
    var myRow = myData[j];
    var myArt = String(myRow[2]).trim();
    var myArtNorm = myArt.toLowerCase().trim();
    var myBase = myArtNorm.split(' ')[0];
    var myBarcode = String(myRow[3]).trim();

    var matchIdx = null;

    // 1. По полному артикулу (точное совпадение)
    if (suppByFullArt[myArtNorm] !== undefined) {
      matchIdx = suppByFullArt[myArtNorm];
    }

    // 2. По штрихкоду
    if (matchIdx === null && myBarcode && myBarcode !== '0' && myBarcode !== '') {
      if (suppByBarcode[myBarcode] !== undefined) matchIdx = suppByBarcode[myBarcode];
    }

    // 3. По базовому артикулу (уточнение по штрихкоду поставщика)
    if (matchIdx === null && suppByBase[myBase]) {
      if (myBarcode && myBarcode !== '0' && myBarcode !== '') {
        for (var k = 0; k < suppByBase[myBase].length; k++) {
          var candIdx = suppByBase[myBase][k];
          if (String(suppData[candIdx][3]).trim() === myBarcode) {
            matchIdx = candIdx;
            break;
          }
        }
      }
      if (matchIdx === null) matchIdx = suppByBase[myBase][0];
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
    var colors = svodRows.map(function(r) {
      var color = r[6] === 'Совпал' ? '#C6EFCE' : '#FFCCCC';
      return new Array(13).fill(color);
    });
    svodSheet.getRange(2, 1, svodRows.length, 13).setBackgrounds(colors);
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

// ТЕСТ: Выгрузка 5 случайных строк в InSales по ID варианта
function тестВыгрузки5строк() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Сводная');
  var data = sheet.getDataRange().getValues();

  var matched = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][6] === 'Совпал' && data[i][11]) matched.push(data[i]);
  }

  // Перемешать и взять 5 случайных
  for (var i = matched.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var tmp = matched[i]; matched[i] = matched[j]; matched[j] = tmp;
  }

  var toUpdate = matched.slice(0, 5).map(function(r) {
    return {
      variantId: r[11],  // L: ID_варианта
      sku: r[0],          // A: Артикул
      price: r[8],        // I: Цена_продажи
      stock: r[9]         // J: Остаток
    };
  });

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
