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

// ШАГ 1: Конвертировать мой прайс
function шаг1_КонвертироватьМойПрайс() {
  var files = DriveApp.getFilesByName('price-carstvo-sna.xls');
  if (!files.hasNext()) {
    SpreadsheetApp.getUi().alert('Файл price-carstvo-sna.xls не найден на Drive');
    return;
  }
  var fileId = files.next().getId();
  var tempFile = Drive.Files.copy(
    { name: '_temp_mystore', mimeType: 'application/vnd.google-apps.spreadsheet' },
    fileId
  );
  PropertiesService.getScriptProperties().setProperty('tempMyStoreId', tempFile.id);
  SpreadsheetApp.getUi().alert('Конвертация готова. Теперь запусти шаг2_ИмпортМойПрайс');
}

// ШАГ 2: Импорт моего прайса
function шаг2_ИмпортМойПрайс() {
  var tempId = PropertiesService.getScriptProperties().getProperty('tempMyStoreId');

  var ss;
  if (tempId) {
    ss = SpreadsheetApp.openById(tempId);
  } else {
    var files = DriveApp.getFilesByName('price-carstvo-sna');
    if (!files.hasNext()) {
      SpreadsheetApp.getUi().alert('Файл price-carstvo-sna не найден на Drive');
      return;
    }
    ss = SpreadsheetApp.openById(files.next().getId());
    tempId = null;
  }

  var sheet = ss.getSheets()[0];
  var lastRow = sheet.getLastRow();
  var data = sheet.getRange(2, 1, lastRow - 1, 30).getValues();

  var result = [];
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    var articul = row[23];
    var cena = row[26];
    if (!articul || parseFloat(String(cena).replace(',', '.')) == 0) continue;
    result.push([
      row[0],   // A: ID_товара
      row[22],  // B: ID_варианта (из XLS, не совпадает с API ID)
      articul,  // C: Артикул
      row[24],  // D: Штрихкод
      row[26],  // E: Цена_продажи
      row[29]   // F: Остаток
    ]);
  }

  var destSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Мой прайс');
  if (destSheet.getLastRow() > 1) {
    destSheet.getRange(2, 1, destSheet.getLastRow() - 1, 6).clearContent();
  }
  if (result.length > 0) {
    destSheet.getRange(2, 1, result.length, 6).setValues(result);
  }

  if (tempId) {
    Drive.Files.remove(tempId);
    PropertiesService.getScriptProperties().deleteProperty('tempMyStoreId');
  }
  SpreadsheetApp.getUi().alert('Загружено строк: ' + result.length);
}

// ТЕСТ: Выгрузка 5 строк в InSales по SKU
function тестВыгрузки5строк() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Сводная');
  var data = sheet.getDataRange().getValues();

  var toUpdate = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][6] === 'Совпал' && toUpdate.length < 5) {
      toUpdate.push({
        productId: data[i][12],  // M: ID_товара
        sku: data[i][0],         // A: Артикул
        price: data[i][8],       // I: Цена_продажи
        stock: data[i][9]        // J: Остаток_поставщика
      });
    }
  }

  var productIds = [];
  toUpdate.forEach(function(x) {
    if (productIds.indexOf(x.productId) === -1) productIds.push(x.productId);
  });

  var skuToVariant = {};
  for (var p = 0; p < productIds.length; p++) {
    var resp = apiGet('/admin/products/' + productIds[p] + '/variants.json');
    if (resp.code === 200) {
      var variants = JSON.parse(resp.body);
      variants.forEach(function(v) {
        skuToVariant[v.sku.toLowerCase()] = { id: v.id, productId: productIds[p] };
      });
    }
    Utilities.sleep(300);
  }

  var results = [];
  for (var j = 0; j < toUpdate.length; j++) {
    var item = toUpdate[j];
    var variantInfo = skuToVariant[item.sku.toLowerCase()];

    if (!variantInfo) {
      results.push(item.sku + ': SKU не найден в API');
      continue;
    }

    var resp2 = apiPut(
      '/admin/products/' + variantInfo.productId + '/variants/' + variantInfo.id + '.json',
      { variant: { price: item.price, quantity: item.stock } }
    );

    var code = resp2.code;
    results.push(item.sku + ': ' + code + (code === 200 ? ' OK' : ' ОШИБКА'));
    Utilities.sleep(300);
  }

  SpreadsheetApp.getUi().alert(results.join('\n'));
}
