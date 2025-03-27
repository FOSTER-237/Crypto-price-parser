var cachedData = {}; // Кэш для хранения данных
var cacheTimestamp = null; // Время последнего обновления кэша
var CACHE_EXPIRATION_MINUTES = 60; // Время жизни кэша (в минутах)

function fetchCryptoData(symbols, convert) {
  var url = `https://pro-api.coinmarketcap.com/v1/cryptocurrency/quotes/latest?symbol=${symbols.join(
    ","
  )}&convert=${convert}`;
  var options = {
    method: "get",
    headers: {
      "X-CMC_PRO_API_KEY": "СЮДА ВСТАВЬТЕ ВАШ API КЛЮЧ", // Укажите ваш API-ключ
    },
  };

  try {
    var response = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(response.getContentText());
    return json.data;
  } catch (e) {
    throw new Error("Ошибка загрузки данных: " + e.message);
  }
}

function fetchSymbolsFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "СЮДА ВСТАВЬТЕ НАЗВАНИЕ ЛИСТА"
  ); // Укажите ваш лист
  if (!sheet) {
    throw new Error("Указанный лист 'СЮДА ВСТАВЬТЕ НАЗВАНИЕ ЛИСТА' не найден!");
  }

  var ranges = ["C6:C38"]; // Ваши диапазоны
  var data = [];

  for (var i = 0; i < ranges.length; i++) {
    try {
      var range = sheet.getRange(ranges[i]);
      Logger.log("Обрабатываем диапазон: " + ranges[i]);
      data.push(...range.getValues().flat().filter(Boolean)); // Собираем тикеры
    } catch (e) {
      Logger.log(
        "Ошибка с диапазоном: " + ranges[i] + ". Сообщение: " + e.message
      );
    }
  }

  if (data.length === 0) {
    throw new Error("Список тикеров пуст!");
  }

  return data;
}

function getCryptoPriceOptimized(symbol, convert) {
  // Приводим тикер к верхнему регистру для унификации
  var normalizedSymbol = symbol.toUpperCase();

  // Проверяем, актуален ли кэш
  var cache = CacheService.getScriptCache();
  var cacheData = cache.get("cryptoCache");
  var cacheTimestamp = cache.get("cryptoCacheTimestamp");

  // Если кэш отсутствует или устарел, обновляем его
  if (
    !cacheData ||
    !cacheTimestamp ||
    (new Date() - new Date(cacheTimestamp)) / 60000 > CACHE_EXPIRATION_MINUTES
  ) {
    updateCache(convert);
    cacheData = cache.get("cryptoCache"); // Перезагружаем кэш
  }

  // Разбираем кэш из строки JSON в объект
  var cachedData = cacheData ? JSON.parse(cacheData) : {};

  // Возвращаем значение из кэша, если оно есть
  if (cachedData[normalizedSymbol]) {
    return cachedData[normalizedSymbol];
  }

  // Если данные не найдены в кэше
  return "Not Found";
}

function updateCache(convert) {
  var symbolsToFetch = fetchSymbolsFromSheet(); // Загружаем тикеры из таблицы
  if (symbolsToFetch.length === 0) {
    throw new Error("Список криптовалют пуст!");
  }

  var data = fetchCryptoData(symbolsToFetch, convert);

  // Создаём объект для кэша
  var cacheObject = {};
  for (var key in data) {
    cacheObject[key] = data[key].quote[convert].price;
  }

  // Сохраняем данные в кэш
  var cache = CacheService.getScriptCache();
  cache.put(
    "cryptoCache",
    JSON.stringify(cacheObject),
    CACHE_EXPIRATION_MINUTES * 60
  ); // Сохраняем на указанный срок
  cache.put(
    "cryptoCacheTimestamp",
    new Date().toISOString(),
    CACHE_EXPIRATION_MINUTES * 60
  );
}

function fetchSymbolsFromSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
    "СЮДА ВСТАВЬТЕ НАЗВАНИЕ ЛИСТА"
  ); // Укажите название листа
  var ranges = ["C6:C38"]; // Укажите диапазоны
  var allValues = [];

  // Собираем значения из указанных диапазонов
  ranges.forEach(function (range) {
    var values = sheet.getRange(range).getValues().flat();
    allValues = allValues.concat(values);
  });

  // Убираем пустые ячейки и возвращаем результат
  return allValues.filter(Boolean).map(function (value) {
    return value.toString().toUpperCase().trim();
  });
}
