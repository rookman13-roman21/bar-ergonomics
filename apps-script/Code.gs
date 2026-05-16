/**
 * Барная эргономика — автозаполнение оборудования
 *
 * УСТАНОВКА (один раз):
 * 1. Откройте таблицу → Расширения → Apps Script
 * 2. Скопируйте этот файл в Code.gs
 * 3. Создайте файл Dialog.html → скопируйте содержимое Dialog.html
 * 4. Сохраните (Ctrl+S)
 * 5. Запустите функцию setupTriggers() → разрешите права доступа
 * 6. Обновите таблицу — появится меню «🔧 Оборудование»
 */

// ──────────────────────────────────────────
// Константы
// ──────────────────────────────────────────

const EQUIPMENT_SHEET  = 'Оборудование';
const LIBRARY_SHEET    = 'Бибилиотека оборудования';
const LOG_SHEET        = 'Лог изменений';
const LIB_HEADER_ROW   = 2;
const LIB_DATA_START   = 3;

// Индексы столбцов библиотеки (0-based)
// Структура: A=Название, B=Производитель, C=Размеры, D=кВт, E=Вода, F=Цена, G=Категория
const LIB = {
  NAME:       0,  // A: Название
  PRODUCER:   1,  // B: Производитель / модель
  DIMENSIONS: 2,  // C: Размеры Ш-Г-В
  KW:         3,  // D: кВт
  WATER:      4,  // E: Подключение к воде
  PRICE:      5,  // F: Цена
  CATEGORY:   6,  // G: Категория
};

// Номера столбцов «Оборудование» (1-based для Sheets API)
const EQ = {
  NUMBER:     1,  // A: Номер на плане
  CATEGORY:   2,  // B: Категория (выпадающий список категорий)
  NAME:       3,  // C: Производитель / модель (зависит от B)
  QTY:        4,  // D: Количество (не трогаем)
  DIMENSIONS: 5,  // E: Размеры ← авто
  KW:         6,  // F: кВт ← авто
  WATER:      7,  // G: Подключение к воде ← авто
  PRICE:      8,  // H: Цена ← авто
  TOTAL:      9,  // I: Сумма (не трогаем — там формула)
};

// Каноничный список категорий — единый источник истины для GPT-prompt и валидаций
const KNOWN_CATEGORIES = [
  'Кофемашины', 'Кофемолки', 'Аксессуары эспрессо', 'Фильтр кофе',
  'Лёд', 'Холодильное оборудование', 'Блендеры', 'Соковыжималки',
  'Горячая вода', 'Водоподготовка', 'Спешиалти напитки', 'Посудомойка',
  'Аксессуары бара', 'Сантехника', 'Касса', 'Кухонное оборудование',
];

const MAX_EQ_ROWS = 200;  // макс. число строк оборудования

// ──────────────────────────────────────────
// Меню
// ──────────────────────────────────────────

function onOpen() {
  try {
    sanitizeEquipmentSheet_(); // чистим строки с пустой категорией но непустыми авто-полями
    resetGreenHighlights_();   // сбрасываем зелёный фон с прошлой сессии
  } catch(initErr) { /* не блокируем открытие таблицы */ }
  try {
    SpreadsheetApp.getUi()
      .createMenu('🔧 Оборудование')
      .addItem('➕ Добавить из библиотеки', 'showAddEquipmentDialog')
      .addSeparator()
      .addItem('🔄 Обновить выпадающие списки', 'refreshDropdowns')
      .addItem('🔢 Обновить итоги', 'updateTotalsRow')
      .addItem('🗑 Очистить строку', 'clearSelectedRow')
      .addItem('🧹 Санировать таблицу', 'sanitizeEquipmentSheet_')
      .addSeparator()
      .addItem('🔒 Защитить библиотеку', 'protectLibrary')
      .addItem('🔑 Настроить OpenAI ключ', 'setOpenAIKey')
      .addItem('🔗 Импортировать из URL', 'importFromUrlMenu')
      .addItem('💰 Обновить цены из URL', 'refreshPricesFromUrls')
      .addSeparator()
      .addSubMenu(
        SpreadsheetApp.getUi().createMenu('📊 Сортировка библиотеки')
          .addItem('По категории (А→Я)', 'sortLibraryByCategory')
          .addItem('По названию (А→Я)', 'sortLibraryByName')
          .addItem('По цене (↑ дешевле)', 'sortLibraryByPriceAsc')
          .addItem('По цене (↓ дороже)', 'sortLibraryByPriceDesc')
          .addItem('По подключению к воде', 'sortLibraryByWater')
          .addItem('По потреблению кВт (↑)', 'sortLibraryByKw')
      )
      .addItem('❓ Справка', 'showHelp')
      .addToUi();
  } catch(e) {
    // getUi() недоступен вне UI-контекста (time-based trigger) — пропускаем меню
  }
}

// ──────────────────────────────────────────
// Автозаполнение при редактировании (onEdit)
// ──────────────────────────────────────────

function onEdit(e) {
  const sheet = e.source.getActiveSheet();

  // ── Импорт из URL в библиотеке ──
  if (sheet.getName() === LIBRARY_SHEET) {
    handleLibraryUrlInput_(sheet, e);
    return;
  }

  if (sheet.getName() !== EQUIPMENT_SHEET) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (row < 2) return;

  // ── B изменён: выбрана категория ──
  if (col === EQ.CATEGORY) {
    // Сбрасываем C и авто-поля
    sheet.getRange(row, EQ.NAME).clearContent().clearDataValidations();
    clearAutoFilledCells(sheet, row);
    const cat = e.value;
    if (cat && cat.trim()) setModelDropdown_(sheet, row, cat.trim());
    return;
  }

  // ── C изменён: выбрана модель → автозаполнение ──
  if (col !== EQ.NAME) return;

  const selectedName = e.value;
  if (!selectedName || selectedName.trim() === '') {
    clearAutoFilledCells(sheet, row);
    return;
  }

  const libData = getLibraryData_();
  const match   = findByName_(libData, selectedName);
  if (!match) return;

  // Извлекаем чистое название модели (убираем префикс «Название • »)
  const sepIdx_    = selectedName.indexOf(' • ');
  const cleanName_ = sepIdx_ !== -1 ? selectedName.slice(sepIdx_ + 3).trim() : selectedName;

  // #11: предупреждение о дубликате
  const dupRow = findDuplicate_(sheet, row, cleanName_);
  if (dupRow) {
    SpreadsheetApp.getUi().alert(`⚠️ «${cleanName_}» уже есть в строке ${dupRow}`);
  }

  fillRow_(sheet, row, match);

  // Переносим гиперссылку из библиотеки в ячейку C — только если её там ещё нет.
  // Проверка обрывает каскад: 1-й вызов onEdit → ссылки нет → устанавливаем → onEdit снова.
  // 2-й вызов onEdit → ссылка уже есть → пропускаем → цикл остановлен.
  const existingRt_ = sheet.getRange(row, EQ.NAME).getRichTextValue();
  if (!existingRt_ || !existingRt_.getLinkUrl()) {
    const libRt_ = getLibraryRichText_(selectedName);
    if (libRt_ && libRt_.getLinkUrl()) {
      const linked_ = SpreadsheetApp.newRichTextValue()
        .setText(selectedName)
        .setLinkUrl(libRt_.getLinkUrl())
        .build();
      sheet.getRange(row, EQ.NAME).setRichTextValue(linked_);
    }
  }

  // Восстанавливаем dropdown с полным списком категории —
  // выбранное значение есть в списке, поэтому ошибки нет, стрелка остаётся.
  // setDataValidation меняет только метаданные ячейки, не значение → не триггерит onEdit.
  const cat_ = String(sheet.getRange(row, EQ.CATEGORY).getValue() || '').trim();
  if (cat_) setModelDropdown_(sheet, row, cat_);

  flashGreen_(sheet, row);
  logChange_(cleanName_, row);  // #9
}

// ──────────────────────────────────────────
// Импорт оборудования из URL (вставка в столбец B библиотеки)
// ──────────────────────────────────────────

/**
 * Срабатывает при вставке значения в столбец B библиотеки.
 * Если значение — URL, ставит его в очередь и создаёт отложенный триггер (1 сек).
 * Сам UrlFetchApp вызывать нельзя из onEdit — нет прав.
 */
function handleLibraryUrlInput_(sheet, e) {
  const col = e.range.getColumn();
  const row = e.range.getRow();
  if (col !== 2) return;
  if (row < LIB_DATA_START) return;

  const val = String(e.value || '').trim();
  if (!/^https?:\/\//i.test(val)) return;

  // Добавляем задание в очередь (поддерживает вставку нескольких URL подряд)
  const props_  = PropertiesService.getScriptProperties();
  const rawQ    = props_.getProperty('URL_QUEUE');
  const queue_  = rawQ ? JSON.parse(rawQ) : [];
  queue_.push({ url: val, row: row, sheetName: sheet.getName() });
  props_.setProperty('URL_QUEUE', JSON.stringify(queue_));

  // Создаём триггер только если его ещё нет — избегаем дублей
  const alreadyScheduled = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'processQueuedUrl_');
  if (!alreadyScheduled) {
    ScriptApp.newTrigger('processQueuedUrl_')
      .timeBased()
      .after(2000)
      .create();
  }

  SpreadsheetApp.getActiveSpreadsheet()
    .toast('⏳ URL получен, запускаю распознавание…', 'Импорт URL', 30);
}

/**
 * Запускается отложенным триггером — имеет полные права, может вызывать UrlFetchApp.
 * Читает задание из очереди, загружает страницу, заполняет строку библиотеки.
 */
function processQueuedUrl_() {
  // Удаляем себя
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processQueuedUrl_')
    .forEach(t => ScriptApp.deleteTrigger(t));

  const props__ = PropertiesService.getScriptProperties();
  const rawQ__  = props__.getProperty('URL_QUEUE');
  if (!rawQ__) return;

  let queue__;
  try { queue__ = JSON.parse(rawQ__); } catch(e) { props__.deleteProperty('URL_QUEUE'); return; }
  if (!queue__.length) { props__.deleteProperty('URL_QUEUE'); return; }

  // Извлекаем первую задачу
  const task = queue__.shift();

  if (queue__.length > 0) {
    // Остались ещё URL — сохраняем и планируем следующий триггер
    props__.setProperty('URL_QUEUE', JSON.stringify(queue__));
    ScriptApp.newTrigger('processQueuedUrl_').timeBased().after(2000).create();
  } else {
    props__.deleteProperty('URL_QUEUE');
  }

  const { url, row, sheetName } = task;
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;

  ss.toast('⏳ Загружаю страницу и распознаю оборудование…', 'Импорт URL', 60);

  try {
    const data = fetchEquipmentFromUrl_(url);
    if (!data || !data.producer) {
      ss.toast('❌ Не удалось распознать оборудование на странице', 'Ошибка', 8);
      return;
    }

    sheet.getRange(row, 1, 1, 7).setValues([[
      data.name       || '',   // A
      data.producer   || '',   // B
      data.dimensions || '',   // C
      data.kw         || '',   // D
      data.water      || '',   // E
      data.price      || '',   // F
      data.category   || '',   // G
    ]]);

    const richText = SpreadsheetApp.newRichTextValue()
      .setText(data.producer)
      .setLinkUrl(url)
      .build();
    sheet.getRange(row, 2).setRichTextValue(richText);

    CacheService.getScriptCache().remove('LIBRARY_DATA');

    ss.toast(`✅ Добавлено: ${data.name} — ${data.producer}`, 'Готово', 6);
  } catch (err) {
    ss.toast('❌ ' + err.message, 'Ошибка импорта', 10);
    logChange_('❌ Ошибка импорта URL: ' + url + ' — ' + err.message, 0);
  }
}

/** Загружает страницу и через OpenAI извлекает характеристики оборудования */
function fetchEquipmentFromUrl_(url) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('API ключ не настроен — запустите меню «🔑 Настроить OpenAI ключ»');

  // 1. Загружаем страницу
  let pageText = '';
  try {
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'Accept-Encoding': 'gzip',
        'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'ru-RU,ru;q=0.9',
      },
    });
    const code = resp.getResponseCode();
    if (code !== 200) throw new Error(`HTTP ${code} при загрузке страницы`);
    const html = resp.getContentText('UTF-8');
    pageText = html
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/\s+/g, ' ')
      .slice(0, 10000);
    if (pageText.trim().length < 50) throw new Error('Страница пустая или заблокирована');
  } catch (e) {
    if (e.message.indexOf('HTTP') === 0 || e.message.indexOf('Страница') === 0) throw e;
    throw new Error('Не удалось загрузить страницу: ' + e.message);
  }

  // 2. GPT-4o-mini: извлекаем структурированные данные
  const prompt =
    'Ты эксперт по барному оборудованию. Извлеки из текста страницы данные об оборудовании ' +
    'и верни ONLY JSON без markdown:\n' +
    '{\n' +
    '  "name": "общее название на русском (например: Кофемашина)",\n' +
    '  "producer": "Бренд и ПОЛНОЕ название модели без слова ‘Кофемашина’/юридического типа. ' +
    'Сохраняй все цифры, версии, опции полностью — не обрезай. ' +
    'Например: Sanremo D8 PRO 2 высокие группы, подсветка групп+Cold Touch",\n' +
    '  "dimensions": "габариты самого прибора без упаковки строго в порядке: Ширина x Глубина x Высота в мм (например: 560x580x430). Если на сайте порядок другой — переставь правильно. Если несколько вариантов — бери рабочие (меньшие) размеры. Или пусто",\n' +
    '  "kw": число_кВт_или_пустая_строка,\n' +
    '  "water": "да / нет / опционально или пусто",\n' +
    '  "price": число_в_рублях_или_пустая_строка,\n' +
    '  "category": "одна из: ' + KNOWN_CATEGORIES.join(' | ') + '"\n' +
    '}\n\nТекст страницы:\n' + pageText;

  let apiResp;
  try {
    apiResp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method:      'post',
      contentType: 'application/json',
      headers:     { 'Authorization': 'Bearer ' + apiKey },
      payload:     JSON.stringify({
        model:       'gpt-4o-mini',
        messages:    [{ role: 'user', content: prompt }],
        temperature: 0,
        max_tokens:  400,
      }),
      muteHttpExceptions: true,
    });
  } catch (e) {
    throw new Error('Нет доступа к OpenAI API: ' + e.message);
  }

  const json = JSON.parse(apiResp.getContentText());
  if (json.error) throw new Error('OpenAI: ' + json.error.message);

  const raw   = json.choices[0].message.content.trim();
  const clean = raw.replace(/^```json\s*/i, '').replace(/```\s*$/, '').trim();

  let result;
  try {
    result = JSON.parse(clean);
  } catch (e) {
    throw new Error('GPT вернул некорректный JSON: ' + clean.slice(0, 200));
  }
  return result;
}

/**
 * Импорт из URL через меню — резервный способ если onEdit не срабатывает.
 * Спрашивает URL, затем спрашивает строку, заполняет библиотеку.
 */
function importFromUrlMenu() {
  const ui = SpreadsheetApp.getUi();

  const urlResult = ui.prompt(
    '🔗 Импорт оборудования по URL',
    'Вставьте ссылку на страницу товара:',
    ui.ButtonSet.OK_CANCEL
  );
  if (urlResult.getSelectedButton() !== ui.Button.OK) return;
  const url = urlResult.getResponseText().trim();
  if (!/^https?:\/\//i.test(url)) {
    ui.alert('❌ Это не ссылка. Должна начинаться с http:// или https://');
    return;
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LIBRARY_SHEET);
  if (!sheet) { ui.alert('❌ Вкладка библиотеки не найдена'); return; }

  // Находим первую пустую строку в библиотеке
  const lastRow  = sheet.getLastRow();
  const insertAt = Math.max(lastRow + 1, LIB_DATA_START);

  ss.toast('⏳ Загружаю страницу и распознаю оборудование…', 'Импорт URL', 60);

  try {
    const data = fetchEquipmentFromUrl_(url);
    if (!data || !data.producer) {
      ui.alert('❌ Не удалось распознать оборудование на странице');
      return;
    }

    sheet.getRange(insertAt, 1, 1, 7).setValues([[
      data.name       || '',   // A
      data.producer   || '',   // B
      data.dimensions || '',   // C
      data.kw         || '',   // D
      data.water      || '',   // E
      data.price      || '',   // F
      data.category   || '',   // G
    ]]);

    const richText = SpreadsheetApp.newRichTextValue()
      .setText(data.producer)
      .setLinkUrl(url)
      .build();
    sheet.getRange(insertAt, 2).setRichTextValue(richText);

    CacheService.getScriptCache().remove('LIBRARY_DATA');

    ui.alert(`✅ Добавлено в строку ${insertAt}:\n${data.name} — ${data.producer}\nКатегория: ${data.category}`);
  } catch (err) {
    ui.alert('❌ Ошибка: ' + err.message);
  }
}

/** Сохраняет OpenAI API ключ в Script Properties */
function setOpenAIKey() {
  const ui     = SpreadsheetApp.getUi();
  const result = ui.prompt(
    '🔑 OpenAI API ключ',
    'Вставьте ваш ключ (начинается с sk-):',
    ui.ButtonSet.OK_CANCEL
  );
  if (result.getSelectedButton() !== ui.Button.OK) return;
  const key = result.getResponseText().trim();
  if (!key.startsWith('sk-')) {
    ui.alert('❌ Неверный формат — ключ должен начинаться с sk-');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', key);
  ui.alert('✅ Ключ сохранён!\n\nТеперь вставьте любую ссылку на оборудование в столбец B библиотеки.');
}

// ──────────────────────────────────────────
// Работа с библиотекой
// ──────────────────────────────────────────

/** Возвращает все строки данных библиотеки (начиная с LIB_DATA_START) */
function getLibraryData_() {
  // #5: кэшируем на 10 минут — избегаем лишних запросов к Sheets
  const cache  = CacheService.getScriptCache();
  const cached = cache.get('LIBRARY_DATA');
  if (cached) {
    try { return JSON.parse(cached); } catch(e) {}
  }
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib) return [];
  const lastRow = lib.getLastRow();
  if (lastRow < LIB_DATA_START) return [];
  const data = lib
    .getRange(LIB_DATA_START, 1, lastRow - LIB_DATA_START + 1, 7)
    .getValues();
  try { cache.put('LIBRARY_DATA', JSON.stringify(data), 600); } catch(e) {}
  return data;
}

/**
 * Ищет строку в библиотеке по имени производителя и возвращает RichTextValue ячейки B.
 * Используется для переноса гиперссылки в столбец C «Оборудование».
 */
function getLibraryRichText_(producerName) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib) return null;
  const lastRow = lib.getLastRow();
  if (lastRow < LIB_DATA_START) return null;
  const norm   = s => String(s).trim().toLowerCase().replace(/\s+/g, ' ');
  // Поддержка составной строки «Название • Производитель»
  const sepIdx = producerName.indexOf(' • ');
  const searchName = sepIdx !== -1 ? producerName.slice(sepIdx + 3).trim() : producerName;
  const target = norm(searchName);
  const values = lib.getRange(LIB_DATA_START, LIB.PRODUCER + 1, lastRow - LIB_DATA_START + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (norm(values[i][0]) === target) {
      return lib.getRange(LIB_DATA_START + i, LIB.PRODUCER + 1).getRichTextValue();
    }
  }
  return null;
}

/** Поиск строки по модели/производителю (столбец B библиотеки), с fallback на название */
function findByName_(libData, name) {
  const norm   = s => String(s).trim().toLowerCase().replace(/\s+/g, ' ');
  const target = norm(name);

  // Если строка в формате «Название • Производитель» — извлекаем часть после « • »
  const sepIdx = target.indexOf(' • ');
  const producerTarget = sepIdx !== -1 ? target.slice(sepIdx + 3).trim() : target;

  // Точное совпадение по столбцу B (Производитель / модель)
  for (const row of libData) {
    if (norm(row[LIB.PRODUCER]) === producerTarget) return row;
  }
  // Частичное по столбцу B
  for (const row of libData) {
    const n = norm(row[LIB.PRODUCER]);
    if (n && (n.includes(producerTarget) || producerTarget.includes(n))) return row;
  }
  // Fallback: точное по столбцу A (Название)
  for (const row of libData) {
    if (norm(row[LIB.NAME]) === target) return row;
  }
  return null;
}

/**
 * Устанавливает зависимый дропдаун в C(row) — список моделей выбранной категории.
 * Вызывается из onEdit при изменении B.
 */
function setModelDropdown_(sheet, row, category) {
  const libData = getLibraryData_();
  const models  = [];
  for (const r of libData) {
    const cat      = String(r[LIB.CATEGORY] || '').trim();
    const name     = String(r[LIB.NAME]     || '').trim();
    const producer = String(r[LIB.PRODUCER] || '').trim();
    if (cat === category && producer) {
      // Показываем «Название • Производитель/Модель» для наглядности
      const label = name ? `${name} • ${producer}` : producer;
      models.push(label);
    }
  }
  if (models.length === 0) return;
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(models, true)
    .setAllowInvalid(true)
    .setHelpText('Выберите модель из библиотеки')
    .build();
  sheet.getRange(row, EQ.NAME).setDataValidation(rule);
}

/** Очищает строку цены "р.110 000" → число 110000 (или оставляет как есть) */
function parsePrice_(raw) {
  if (!raw) return '';
  const s = String(raw)
    .replace(/р\.\s*/i, '')   // убираем "р."
    .replace(/\u00a0/g, '')   // убираем неразрывные пробелы
    .replace(/\s/g, '')       // убираем обычные пробелы
    .replace(',', '.');       // запятая → точка
  const n = parseFloat(s);
  return isNaN(n) ? raw : n;  // если число — отдаём число, иначе оригинал
}

/** Заполняет авто-поля строки в «Оборудование» */
function fillRow_(sheet, row, libRow) {
  const kw = String(libRow[LIB.KW] || '').trim();

  // #4: batch — EQ.DIMENSIONS(5)→KW(6)→WATER(7)→PRICE(8), C(EQ.NAME) заполняет пользователь
  sheet.getRange(row, EQ.DIMENSIONS, 1, 4).setValues([[
    libRow[LIB.DIMENSIONS] || '',
    kw,
    libRow[LIB.WATER]      || '',
    parsePrice_(libRow[LIB.PRICE]),
  ]]);

  // #6: автоформула суммы в I, если ячейка пустая
  const totalCell = sheet.getRange(row, EQ.TOTAL);
  if (!totalCell.getFormula() && !totalCell.getValue()) {
    totalCell.setFormula(`=D${row}*H${row}`);
  }
}

/** Очищает авто-заполненные поля C и E..I при сбросе строки. D (количество) не трогаем. */
function clearAutoFilledCells(sheet, row) {
  sheet.getRange(row, EQ.NAME, 1, 1).clearContent();                                    // C
  sheet.getRange(row, EQ.DIMENSIONS, 1, EQ.TOTAL - EQ.DIMENSIONS + 1).clearContent();  // E,F,G,H,I
  sheet.getRange(row, EQ.NAME, 1, 6).setBackground(null);                               // сброс фона C-H
}

/** Подсветка зелёным — визуальная обратная связь. Сбрасывается при следующем onOpen. */
function flashGreen_(sheet, row) {
  sheet.getRange(row, EQ.NAME, 1, 6).setBackground('#c8e6c9');
  // #1: запоминаем строку — сбросим фон при следующем открытии таблицы
  const props    = PropertiesService.getScriptProperties();
  const existing = props.getProperty('GREEN_ROWS') || '';
  const set      = new Set(existing ? existing.split(',') : []);
  set.add(sheet.getName() + ':' + row);
  props.setProperty('GREEN_ROWS', Array.from(set).join(','));
}

/** Очищает строки где B (категория) пустая, но C/E-I содержат остаточные данные.
 * Вызывается при каждом открытии таблицы (onOpen). */
function sanitizeEquipmentSheet_() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET);
  if (!sheet) return;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Читаем B..I за один запрос
  const numCols = EQ.TOTAL - EQ.CATEGORY + 1;  // 8 столбцов: B,C,D,E,F,G,H,I
  const vals    = sheet.getRange(2, EQ.CATEGORY, lastRow - 1, numCols).getValues();

  const rowsToClear = [];
  for (let i = 0; i < vals.length; i++) {
    const cat = String(vals[i][0] || '').trim();  // B
    if (cat) continue;                             // категория есть — строка ок
    // B пустая — проверяем C..I (индексы 1..7 в строке)
    const hasResidue = vals[i].slice(1).some(v => String(v || '').trim() !== '');
    if (hasResidue) rowsToClear.push(i + 2);
  }

  if (rowsToClear.length === 0) return;

  for (const row of rowsToClear) {
    sheet.getRange(row, EQ.NAME, 1, 1).clearContent();                                    // C
    sheet.getRange(row, EQ.DIMENSIONS, 1, EQ.TOTAL - EQ.DIMENSIONS + 1).clearContent();  // E,F,G,H,I
    sheet.getRange(row, EQ.NAME, 1, 6).setBackground(null);                               // сброс фона C-H
  }
}

/** Сбрасывает зелёный фон со всех запомненных строк (#1) — batch по листу */
function resetGreenHighlights_() {
  const props  = PropertiesService.getScriptProperties();
  const stored = props.getProperty('GREEN_ROWS');
  if (!stored) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Группируем строки по имени листа → один setBackground на диапазон
  const bySheet = {};
  for (const key of stored.split(',')) {
    if (!key) continue;
    const [sheetName, rowStr] = key.split(':');
    const r = parseInt(rowStr);
    if (isNaN(r)) continue;
    if (!bySheet[sheetName]) bySheet[sheetName] = [];
    bySheet[sheetName].push(r);
  }
  for (const [sheetName, rows] of Object.entries(bySheet)) {
    const s = ss.getSheetByName(sheetName);
    if (!s) continue;
    // Сбрасываем все строки одним batch-запросом через RangeList
    const ranges = rows.map(r => s.getRange(r, EQ.NAME, 1, 6));
    s.getRangeList(ranges.map(r => r.getA1Notation())).setBackground(null);
  }
  props.deleteProperty('GREEN_ROWS');
}

// ──────────────────────────────────────────
// Диалог выбора по категориям (Sidebar)
// ──────────────────────────────────────────

/** Открывает боковую панель «Добавить из библиотеки» */
function showAddEquipmentDialog() {
  const html = HtmlService.createHtmlOutputFromFile('Dialog')
    .setTitle('Добавить оборудование')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/** Возвращает уникальные категории для диалога */
function getCategories() {
  const libData = getLibraryData_();
  const cats    = new Set();
  for (const row of libData) {
    const cat = String(row[LIB.CATEGORY] || '').trim();
    if (cat) cats.add(cat);
  }
  const sorted = Array.from(cats).sort();
  sorted.unshift('— Все —');
  return sorted;
}

/** Возвращает оборудование (с деталями) по категории */
function getEquipmentByCategory(category) {
  const libData = getLibraryData_();
  const result  = [];
  for (const row of libData) {
    const name = String(row[LIB.NAME] || '').trim();
    const cat  = String(row[LIB.CATEGORY] || '').trim();
    if (!name) continue;
    if (category !== '— Все —' && cat !== category) continue;

    const kw = String(row[LIB.KW] || '').trim();

    result.push({
      name:       name,
      producer:   String(row[LIB.PRODUCER]   || '').trim(),
      dimensions: String(row[LIB.DIMENSIONS] || '').trim(),
      kw:         kw,
      water:      String(row[LIB.WATER]      || '').trim(),
      price:      parsePrice_(row[LIB.PRICE]),  // #10: число вместо строки
      category:   cat,
    });
  }
  return result;
}

/** Возвращает строки «Оборудование» для выбора в диалоге: заполненные + первая пустая (#5) */
function getEquipmentRows() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET);
  const last  = Math.max(sheet.getLastRow(), 25);
  const values = sheet.getRange(2, EQ.NUMBER, last - 1, 3).getValues();  // A, B, C
  const rows = [];
  let firstEmptyAdded = false;
  for (let i = 0; i < values.length; i++) {
    const num  = String(values[i][0] || '').trim();
    const name = String(values[i][2] || '').trim();  // C = модель
    if (name && name !== 'ИТОГО') {
      // Заполненная строка — всегда показываем
      rows.push({ row: i + 2, num: num || '—', name: name });
    } else if (!name && !firstEmptyAdded) {
      // Первая пустая строка — показываем одну (#5)
      rows.push({ row: i + 2, num: num || '—', name: '' });
      firstEmptyAdded = true;
    }
  }
  return rows;
}

/** Записывает оборудование в строку (вызывается из диалога) */
function addEquipmentToRow(rowNumber, equipmentName) {
  if (rowNumber < 2 || rowNumber > MAX_EQ_ROWS) {
    return { success: false, error: `Неверный номер строки (макс. ${MAX_EQ_ROWS})` };
  }
  const libData = getLibraryData_();
  const match   = findByName_(libData, equipmentName);
  if (!match) {
    return { success: false, error: 'Оборудование не найдено в библиотеке' };
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET);

  const eqName = match[LIB.PRODUCER] || match[LIB.NAME];
  sheet.getRange(rowNumber, EQ.CATEGORY).setValue(match[LIB.CATEGORY] || '');
  sheet.getRange(rowNumber, EQ.NAME).setValue(eqName);

  // Переносим гиперссылку из библиотеки (столбец B) в столбец C «Оборудование»
  const rt = getLibraryRichText_(eqName);
  if (rt && rt.getLinkUrl()) {
    const linked = SpreadsheetApp.newRichTextValue()
      .setText(eqName)
      .setLinkUrl(rt.getLinkUrl())
      .build();
    sheet.getRange(rowNumber, EQ.NAME).setRichTextValue(linked);
  }

  fillRow_(sheet, rowNumber, match);
  flashGreen_(sheet, rowNumber);
  logChange_(match[LIB.PRODUCER] || match[LIB.NAME], rowNumber);  // #9

  return { success: true, row: rowNumber, name: match[LIB.NAME] };
}

// ──────────────────────────────────────────
// Обновление выпадающих списков
// ──────────────────────────────────────────

function refreshDropdowns() {
  const libData = getLibraryData_();
  const cats    = new Set();

  for (const row of libData) {
    const cat = String(row[LIB.CATEGORY] || '').trim();
    if (cat) cats.add(cat);
  }

  if (cats.size === 0) {
    SpreadsheetApp.getUi().alert('Библиотека пуста — список не обновлён');
    return;
  }

  const catList = Array.from(cats).sort();  // #3: алфавитный порядок
  CacheService.getScriptCache().remove('LIBRARY_DATA');  // #5: сброс кэша

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET);

  // Диапазон: до lastRow + 10 запасных строк, минимум 50
  const lastEqRow = Math.max(sheet.getLastRow() + 10, 50);

  // B — список категорий
  const catRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(catList, true)
    .setAllowInvalid(true)
    .setHelpText('Выберите категорию оборудования')
    .build();
  sheet.getRange(2, EQ.CATEGORY, lastEqRow - 1, 1).setDataValidation(catRule);

  // C — очищаем старые валидации (будут проставляться динамически onEdit)
  sheet.getRange(2, EQ.NAME, lastEqRow - 1, 1).clearDataValidations();

  // Восстанавливаем выпадающий список C для строк, где уже выбрана категория в B
  const existingLastRow = sheet.getLastRow();
  if (existingLastRow >= 2) {
    const catValues = sheet.getRange(2, EQ.CATEGORY, existingLastRow - 1, 1).getValues();
    for (let i = 0; i < catValues.length; i++) {
      const cat = String(catValues[i][0] || '').trim();
      if (cat) {
        setModelDropdown_(sheet, i + 2, cat);
      }
    }
  }

  SpreadsheetApp.getUi().alert(`✅ Обновлено: ${catList.length} категорий в столбце B (строки 2–${lastEqRow})\nВыпадающие списки в столбце C восстановлены для уже заполненных строк`);
}

// ──────────────────────────────────────────
// Очистка строки
// ──────────────────────────────────────────

function clearSelectedRow() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (sheet.getName() !== EQUIPMENT_SHEET) {
    SpreadsheetApp.getUi().alert('Перейдите на вкладку «Оборудование»');
    return;
  }
  const row = sheet.getActiveRange().getRow();
  if (row < 2) return;

  // Очищаем B (категория), C (модель + валидация), E-I (авто) — не трогаем A и D
  sheet.getRange(row, EQ.CATEGORY).clearContent();
  sheet.getRange(row, EQ.NAME).clearContent().clearDataValidations();
  clearAutoFilledCells(sheet, row);
}

// ──────────────────────────────────────────
// Установка триггеров (запустить один раз)
// ──────────────────────────────────────────

function setupTriggers() {
  // Удаляем дублирующиеся onEdit-триггеры
  for (const t of ScriptApp.getProjectTriggers()) {
    if (t.getHandlerFunction() === 'onEdit') ScriptApp.deleteTrigger(t);
  }
  // Создаём installable-триггер (надёжнее simple trigger)
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  SpreadsheetApp.getUi().alert(
    '✅ Триггер установлен!\n\nТеперь при выборе категории в столбце B ' +
    'появится список моделей в столбце C, а характеристики заполнятся автоматически.'
  );
}

// ──────────────────────────────────────────
// Валидация дубликатов (#11)
// ──────────────────────────────────────────

/** Возвращает номер строки где уже выбрано это оборудование, или null */
function findDuplicate_(sheet, currentRow, name) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, EQ.NAME, lastRow - 1, 1).getValues();
  // Убираем префикс «Тип • » при сравнении — корректно работает и со старыми строками
  const stripPrefix = s => { const sep = s.indexOf(' • '); return sep !== -1 ? s.slice(sep + 3).trim() : s; };
  const norm = s => stripPrefix(String(s).trim().toLowerCase());
  const normName = norm(name);
  for (let i = 0; i < values.length; i++) {
    const r = i + 2;
    if (r === currentRow) continue;
    if (norm(values[i][0]) === normName) return r;
  }
  return null;
}

// ──────────────────────────────────────────
// Лог изменений (#9)
// ──────────────────────────────────────────

/** Записывает событие в вкладку «Лог изменений» */
function logChange_(equipmentName, rowNumber) {
  try {
    const ss  = SpreadsheetApp.getActiveSpreadsheet();
    let   log = ss.getSheetByName(LOG_SHEET);
    if (!log) {
      log = ss.insertSheet(LOG_SHEET);
      log.getRange(1, 1, 1, 4)
         .setValues([['Дата', 'Пользователь', 'Строка', 'Оборудование']])
         .setFontWeight('bold');
      log.setFrozenRows(1);
    }
    const user = Session.getActiveUser().getEmail() || '—';
    log.appendRow([new Date(), user, rowNumber, equipmentName]);
  } catch(e) { /* не блокируем основной поток */ }
}

// ──────────────────────────────────────────
// Итоги (#7)
// ──────────────────────────────────────────

/** Добавляет/обновляет строку ИТОГО с суммой столбца I */
function updateTotalsRow() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(EQUIPMENT_SHEET);
  if (!sheet) return;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  // Находим последнюю непустую строку в B
  const bVals = sheet.getRange(2, EQ.NAME, lastRow - 1, 1).getValues();
  let dataEnd = 1;
  for (let i = 0; i < bVals.length; i++) {
    if (String(bVals[i][0]).trim() && String(bVals[i][0]).trim() !== 'ИТОГО') {
      dataEnd = i + 2;
    }
  }
  if (dataEnd < 2) return;

  const totRow = dataEnd + 1;
  // Очищаем все существующие строки ИТОГО во всём листе
  const allB = sheet.getRange(2, EQ.NAME, lastRow - 1, 1).getValues();
  for (let i = 0; i < allB.length; i++) {
    if (String(allB[i][0]).trim() === 'ИТОГО') {
      sheet.getRange(i + 2, 1, 1, 10).clearContent();
    }
  }

  sheet.getRange(totRow, EQ.NAME).setValue('ИТОГО').setFontWeight('bold');
  sheet.getRange(totRow, EQ.TOTAL)
       .setFormula(`=SUM(I2:I${dataEnd})`)
       .setFontWeight('bold');

  SpreadsheetApp.getUi().alert(`✅ Итого обновлено (строки 2–${dataEnd})`);
}

// ──────────────────────────────────────────
// Защита библиотеки (#8)
// ──────────────────────────────────────────

/** Включает защиту вкладки «Библиотека оборудования» (режим предупреждения) */
function protectLibrary() {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  const lib = ss.getSheetByName(LIBRARY_SHEET);
  if (!lib) {
    SpreadsheetApp.getUi().alert('Вкладка библиотеки не найдена');
    return;
  }
  // Снимаем старые защиты
  lib.getProtections(SpreadsheetApp.ProtectionType.SHEET)
     .forEach(p => p.remove());
  lib.protect()
     .setDescription('Библиотека оборудования — не редактировать вручную')
     .setWarningOnly(true);
  SpreadsheetApp.getUi().alert(
    '🔒 Библиотека защищена!\n\n' +
    'При попытке редактировать появится предупреждение.\n' +
    'Для снятия: Данные → Защищённые листы и диапазоны.'
  );
}

// ──────────────────────────────────────────
// Обновление цен из URL
// ──────────────────────────────────────────

/**
 * Запускает обновление цен для всех позиций библиотеки с сохранёнными URL.
 * Ссылки берутся из гиперссылок в столбце B (Производитель / модель).
 */
function refreshPricesFromUrls() {
  const ui    = SpreadsheetApp.getUi();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LIBRARY_SHEET);
  if (!sheet) { ui.alert('Вкладка библиотеки не найдена'); return; }

  const lastRow = sheet.getLastRow();
  if (lastRow < LIB_DATA_START) { ui.alert('Библиотека пуста'); return; }

  // Собираем строки с гиперссылками в столбце B
  const queue = [];
  for (let row = LIB_DATA_START; row <= lastRow; row++) {
    const rt = sheet.getRange(row, LIB.PRODUCER + 1).getRichTextValue();
    if (rt) {
      const url = rt.getLinkUrl();
      if (url && /^https?:\/\//i.test(url)) {
        const name = String(sheet.getRange(row, LIB.NAME + 1).getValue() || '').trim();
        queue.push({ row, url, name });
      }
    }
  }

  if (queue.length === 0) {
    ui.alert('❌ В библиотеке нет позиций с сохранёнными ссылками.\n\nСсылки сохраняются автоматически при импорте через URL.');
    return;
  }

  const confirm = ui.alert(
    '💰 Обновление цен',
    `Найдено позиций со ссылками: ${queue.length}\n\nЗапустить обновление цен? Это может занять несколько минут.`,
    ui.ButtonSet.OK_CANCEL
  );
  if (confirm !== ui.Button.OK) return;

  // Сохраняем очередь в ScriptProperties
  const props = PropertiesService.getScriptProperties();
  props.setProperty('PRICE_REFRESH_QUEUE', JSON.stringify(queue));
  props.setProperty('PRICE_REFRESH_TOTAL', String(queue.length));
  props.setProperty('PRICE_REFRESH_DONE',  '0');
  props.setProperty('PRICE_REFRESH_UPDATED', '0');

  // Создаём триггер если ещё нет
  const already = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'processPriceRefreshQueue_');
  if (!already) {
    ScriptApp.newTrigger('processPriceRefreshQueue_').timeBased().after(1000).create();
  }

  ss.toast(`⏳ Запускаю обновление цен для ${queue.length} позиций…`, 'Обновление цен', 15);
}

/** Обрабатывает одну позицию из очереди обновления цен (запускается триггером) */
function processPriceRefreshQueue_() {
  // Удаляем себя
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'processPriceRefreshQueue_')
    .forEach(t => ScriptApp.deleteTrigger(t));

  const props = PropertiesService.getScriptProperties();
  const rawQ  = props.getProperty('PRICE_REFRESH_QUEUE');
  if (!rawQ) return;

  let queue;
  try { queue = JSON.parse(rawQ); } catch(e) { props.deleteProperty('PRICE_REFRESH_QUEUE'); return; }
  if (!queue.length) { props.deleteProperty('PRICE_REFRESH_QUEUE'); return; }

  const task    = queue.shift();
  const total   = parseInt(props.getProperty('PRICE_REFRESH_TOTAL')   || '0');
  const done    = parseInt(props.getProperty('PRICE_REFRESH_DONE')    || '0') + 1;
  const updated = parseInt(props.getProperty('PRICE_REFRESH_UPDATED') || '0');

  if (queue.length > 0) {
    props.setProperty('PRICE_REFRESH_QUEUE', JSON.stringify(queue));
    props.setProperty('PRICE_REFRESH_DONE', String(done));
    ScriptApp.newTrigger('processPriceRefreshQueue_').timeBased().after(2000).create();
  } else {
    // Очередь закончилась — чистим после обработки последнего элемента
    props.deleteProperty('PRICE_REFRESH_QUEUE');
  }

  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LIBRARY_SHEET);
  if (!sheet) return;

  ss.toast(`⏳ [${done}/${total}] Проверяю цену: ${task.name || '…'}`, 'Обновление цен', 30);

  let updatedCount = updated;
  try {
    const newPrice = fetchPriceFromUrl_(task.url);
    if (newPrice !== null) {
      const oldRaw   = sheet.getRange(task.row, LIB.PRICE + 1).getValue();
      const oldPrice = parseFloat(String(oldRaw).replace(/[\s\u00a0]/g, '').replace(',', '.')) || 0;
      sheet.getRange(task.row, LIB.PRICE + 1).setValue(newPrice);
      CacheService.getScriptCache().remove('LIBRARY_DATA');

      // Логируем если цена изменилась
      if (oldPrice && Math.abs(newPrice - oldPrice) > 0.01) {
        logChange_(`💰 Цена обновлена: ${task.name} — ${oldPrice.toLocaleString('ru-RU')} → ${newPrice.toLocaleString('ru-RU')} ₽`, task.row);
        updatedCount++;
        props.setProperty('PRICE_REFRESH_UPDATED', String(updatedCount));
      }
    }
  } catch(err) {
    // Ошибка на одной позиции — не останавливаем всю очередь
    logChange_(`⚠️ Не удалось обновить цену: ${task.name} — ${err.message}`, task.row);
  }

  // Финальный тост при завершении очереди
  if (queue.length === 0) {
    const finalUpdated = parseInt(props.getProperty('PRICE_REFRESH_UPDATED') || '0');
    props.deleteProperty('PRICE_REFRESH_TOTAL');
    props.deleteProperty('PRICE_REFRESH_DONE');
    props.deleteProperty('PRICE_REFRESH_UPDATED');
    ss.toast(
      finalUpdated > 0
        ? `✅ Готово: проверено ${total}, обновлено ${finalUpdated} цен`
        : `✅ Готово: проверено ${total}, цены актуальны`,
      'Обновление цен', 8
    );
  }
}

/**
 * Загружает страницу и через GPT-4o-mini извлекает только текущую цену.
 * Лёгкий запрос (max_tokens: 20) — дешевле полного парсинга.
 */
function fetchPriceFromUrl_(url) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!apiKey) throw new Error('API ключ не настроен');

  // Загружаем страницу
  let pageText = '';
  try {
    const resp = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      followRedirects: true,
      headers: {
        'Accept-Encoding': 'gzip',
        'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1',
        'Accept': 'text/html,application/xhtml+xml',
        'Accept-Language': 'ru-RU,ru;q=0.9',
      },
    });
    if (resp.getResponseCode() !== 200) return null;
    const html = resp.getContentText('UTF-8');
    pageText = html
      .replace(/<script[\s\S]*?<\/script>/gi, ' ')
      .replace(/<style[\s\S]*?<\/style>/gi, ' ')
      .replace(/<[^>]+>/g, ' ')
      .replace(/&nbsp;/g, ' ')
      .replace(/\s+/g, ' ')
      .slice(0, 6000);  // меньше чем при полном парсинге — экономим токены
    if (pageText.trim().length < 50) return null;
  } catch(e) { return null; }

  // GPT: только цена
  const prompt =
    'Найди текущую цену товара на странице. ' +
    'Верни ТОЛЬКО число в рублях без пробелов, знаков и слов. Например: 125000\n' +
    'Если цена не найдена или товар недоступен — верни null.\n\n' +
    'Текст страницы:\n' + pageText;

  let apiResp;
  try {
    apiResp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post', contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + apiKey },
      payload: JSON.stringify({
        model:       'gpt-4o-mini',
        messages:    [{ role: 'user', content: prompt }],
        temperature: 0,
        max_tokens:  20,
      }),
      muteHttpExceptions: true,
    });
  } catch(e) { throw new Error('OpenAI недоступен: ' + e.message); }

  const json = JSON.parse(apiResp.getContentText());
  if (json.error) throw new Error('OpenAI: ' + json.error.message);

  const raw = json.choices[0].message.content.trim();
  if (raw === 'null' || raw === '') return null;
  const n = parseFloat(raw.replace(/[\s\u00a0]/g, '').replace(',', '.'));
  return isNaN(n) ? null : n;
}

// ──────────────────────────────────────────
// Сортировка библиотеки
// ──────────────────────────────────────────

/**
 * Сортирует библиотеку оборудования по указанному столбцу.
 * @param {number} col       - 1-based номер столбца для сортировки
 * @param {boolean} ascending - направление сортировки
 * @param {string} label     - человекочитаемое название для сообщения
 */
function sortLibrary_(col, ascending, label) {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(LIBRARY_SHEET);
  if (!sheet) {
    SpreadsheetApp.getUi().alert('Вкладка библиотеки не найдена');
    return;
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < LIB_DATA_START) {
    SpreadsheetApp.getUi().alert('Библиотека пуста');
    return;
  }
  // Сортируем только строки с данными, не трогая заголовок
  sheet
    .getRange(LIB_DATA_START, 1, lastRow - LIB_DATA_START + 1, 7)
    .sort({ column: col, ascending: ascending });

  // Сбрасываем кэш — данные изменились
  CacheService.getScriptCache().remove('LIBRARY_DATA');

  SpreadsheetApp.getActiveSpreadsheet()
    .toast(`✅ Библиотека отсортирована: ${label}`, 'Сортировка', 4);
}

function sortLibraryByCategory()  { sortLibrary_(LIB.CATEGORY  + 1, true,  'по категории А→Я'); }
function sortLibraryByName()      { sortLibrary_(LIB.NAME      + 1, true,  'по названию А→Я'); }
function sortLibraryByPriceAsc()  { sortLibrary_(LIB.PRICE     + 1, true,  'по цене ↑ дешевле'); }
function sortLibraryByPriceDesc() { sortLibrary_(LIB.PRICE     + 1, false, 'по цене ↓ дороже'); }
function sortLibraryByWater()     { sortLibrary_(LIB.WATER     + 1, true,  'по подключению к воде'); }
function sortLibraryByKw()        { sortLibrary_(LIB.KW        + 1, true,  'по потреблению кВт ↑'); }

// ──────────────────────────────────────────
// Справка
// ──────────────────────────────────────────

function showHelp() {
  SpreadsheetApp.getUi().alert(
    '🔧 Барная эргономика — справка\n\n' +

    '─── ДОБАВЛЕНИЕ ОБОРУДОВАНИЯ ───\n\n' +

    '1. Вручную через таблицу:\n' +
    '   • В столбце B выберите категорию →\n' +
    '   • В столбце C появится список моделей →\n' +
    '   • Выберите модель — поля E,F,G,H заполнятся сами\n\n' +

    '2. Через боковую панель:\n' +
    '   • Меню → «➕ Добавить из библиотеки»\n' +
    '   • Выберите категорию, модель, строку → «Добавить»\n\n' +

    '3. Импорт по URL (автоматически через GPT):\n' +
    '   • Вставьте ссылку на товар в столбец B библиотеки\n' +
    '   • Или: Меню → «🔗 Импортировать из URL»\n' +
    '   • Характеристики и цена заполнятся автоматически\n\n' +

    '─── ОБСЛУЖИВАНИЕ ───\n\n' +

    '🔄 Обновить выпадающие списки\n' +
    '   После добавления новых позиций в Библиотеку\n\n' +

    '💰 Обновить цены из URL\n' +
    '   Перезагружает страницы всех позиций с ссылками\n' +
    '   и обновляет изменившиеся цены. Изменения\n' +
    '   фиксируются в «Лог изменений»\n\n' +

    '🔢 Обновить итоги\n' +
    '   Добавляет/пересчитывает строку ИТОГО (столбец I)\n\n' +

    '🗑 Очистить строку\n' +
    '   Удаляет данные выделенной строки (A и D не трогает)\n\n' +

    '🧹 Санировать таблицу\n' +
    '   Убирает «мусорные» данные в строках без категории\n\n' +

    '📊 Сортировка библиотеки\n' +
    '   По категории, названию, цене, воде, кВт\n\n' +

    '─── НАСТРОЙКА ───\n\n' +

    '🔑 Настроить OpenAI ключ — необходим для импорта\n' +
    '   по URL и обновления цен (ключ начинается с sk-)\n\n' +

    '🔒 Защитить библиотеку — включает предупреждение\n' +
    '   при попытке редактировать библиотеку вручную\n\n' +

    '─── СТРУКТУРА ТАБЛИЦЫ ───\n\n' +
    'A — Номер на плане   B — Категория\n' +
    'C — Модель           D — Количество (вручную)\n' +
    'E — Размеры          F — кВт\n' +
    'G — Вода             H — Цена   I — Сумма (D×H)'
  );
}
