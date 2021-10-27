// создание меню команд
function onOpen() {
  SpreadsheetApp.getUi()
  // добавление пункта меню в оболочку
  .createMenu('Команды')
  // добавление подпунктов
  .addItem('Информация о товаре', 'skuInfo')
  .addItem('Продажи за период', 'skuSales')
  .addItem('Похожие товары', 'skuSimilar')
  .addToUi();
}

const apiUrl = 'https://mpstats.io/api/';

function skuInfo () {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Товарная позиция');
  var sku = ss.getRange(2,1).getValue();
  var token = ss.getRange(2,4).getValue();
  var servUrl = apiUrl + 'wb/get/item/' + sku;

  var hdr = {
      'X-Mpstats-TOKEN': token,
      'Content-Type': 'application/json',
      'maxRedirects' : 20
  };

  var options = {
      'method': 'GET',
      'headers': hdr,
  };

  let response = UrlFetchApp.fetch(servUrl, options);
  let data = JSON.parse(response);

  ss.getRange(5,2,19,3).clearContent();

  ss.getRange(5,2).setValue(data.item.id);
  ss.getRange(6,2).setValue(data.item.name);
  ss.getRange(7,2).setValue(data.item.full_name);
  ss.getRange(8,2).setValue(data.item.link);
  ss.getRange(9,2).setValue(data.item.brand);
  ss.getRange(10,2).setValue(data.item.seller);
  ss.getRange(11,2).setValue(data.item.rating);
  ss.getRange(12,2).setValue(data.item.comments);
  ss.getRange(13,2).setValue(data.item.price);
  ss.getRange(14,2).setValue(data.item.final_price);
  ss.getRange(15,2).setValue(data.item.discount);
  ss.getRange(16,2).setValue(data.item.updated);
  ss.getRange(17,2).setValue(data.item.first_date);
  ss.getRange(18,2).setValue(data.item.is_new);
  ss.getRange(19,2).setValue(data.item.sizeandstores);

  for (var i=0; i<data.photos.length; i++) {
    ss.getRange(20+i*2,2).setValue(data.photos[i].f); 
    ss.getRange(21+i*2,2).setValue(data.photos[i].t); 
  }
}

function skuSales () {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Товарная позиция');
  var sku = ss.getRange(2,1).getValue();
  var token = ss.getRange(2,4).getValue();
  var d1 = Utilities.formatDate(ss.getRange(2,2).getValue(),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var d2 = Utilities.formatDate(ss.getRange(2,3).getValue(),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var servUrl = apiUrl + 'wb/get/item/' + sku + '/sales?d1=' + d1 + '&d2=' + d2;

  var hdr = {
      'X-Mpstats-TOKEN': token,
      'Content-Type': 'application/json',
      'maxRedirects' : 20
  };

  var options = {
      'method': 'GET',
      'headers': hdr,
  };

  let response = UrlFetchApp.fetch(servUrl, options);
  let data = JSON.parse(response);

  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('продажи');

  for (var i=0; i<data.length; i++) {
    ss.appendRow([data[i].no_data,data[i].data,data[i].balance,data[i].sales,data[i].rating,data[i].price,data[i].final_price,data[i].is_new,data[i].comments,data[i].discount,data[i].basic_sale,data[i].basic_price,data[i].promo_sale,data[i].client_sale,data[i].client_price]);
  }
}

function skuSimilar () {
  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Товарная позиция');
  var sku = ss.getRange(2,1).getValue();
  var token = ss.getRange(2,4).getValue();
  var d1 = Utilities.formatDate(ss.getRange(2,2).getValue(),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var d2 = Utilities.formatDate(ss.getRange(2,3).getValue(),SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), 'yyyy-MM-dd');
  var servUrl = apiUrl + 'wb/get/item/' + sku + '/similar?d1=' + d1 + '&d2=' + d2;

  var hdr = {
      'X-Mpstats-TOKEN': token,
      'Content-Type': 'application/json',
      'maxRedirects' : 20
  };

  var options = {
      'method': 'GET',
      'headers': hdr,
  };

  let response = UrlFetchApp.fetch(servUrl, options);
  let data = JSON.parse(response);

  ss = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('похожие');

  for (var i=0; i<data.length; i++) {
    ss.appendRow([data[i].id,data[i].name,data[i].final_price,data[i].sales,data[i].brand,data[i].seller,data[i].revenue,data[i].rating,data[i].comments,data[i].balance,data[i].is_fbs,data[i].thumb,data[i].thumb_middle]);
  }
}
