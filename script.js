const STORAGE_KEY = 'ozon-normalized-dashboard-v1';

const TEMPLATE_SPECS = {
  'Начисления': {
    required: true,
    label: 'Начисления',
    headers: [
      'ID начисления','Дата начисления','Группа услуг','Тип начисления','Артикул','SKU','Название товара','Количество',
      'Цена продавца','Дата принятия заказа в обработку или оказания услуги','Платформа продажи','Схема работы',
      'Вознаграждение Ozon, %','Индекс локализации, %','Среднее время доставки, часы','Сумма итого, руб.'
    ],
    dateFields: ['Дата начисления']
  },
  'Заказы': {
    required: true,
    label: 'Заказы',
    headers: [
      'Номер заказа','Номер отправления','Принят в обработку','Дата отгрузки','Статус','Дата доставки',
      'Фактическая дата передачи в доставку','Сумма отправления','Код валюты отправления','Название товара','SKU','Артикул',
      'Ваша цена','Код валюты товара','Оплачено покупателем','Код валюты покупателя','Количество','Стоимость доставки',
      'Связанные отправления','Выкуп товара','Цена товара до скидок','Скидка %','Скидка руб','Акции','Объемный вес товаров, кг',
      'Кластер отгрузки','Кластер доставки','Склад отгрузки','Регион доставки','Город доставки','Способ доставки',
      'Сегмент клиента','Юридическое лицо','Способ оплаты','Адрес покупателя','Штрихкод ювелирного изделия'
    ],
    dateFields: ['Принят в обработку']
  },
  'Продвижение': {
    required: true,
    label: 'Продвижение',
    headers: ['SKU','Название товара','Инструмент','Место размещения','ID кампании','Расход, ₽','ДРР, %','Продажи, ₽','Заказы, шт','CTR, %','Показы','Клики'],
    dateFields: []
  },
  'Возвраты': {
    required: true,
    label: 'Возвраты',
    headers: [
      'Схема','Наименование товара','Номер отправления','Артикул товара','OZON SKU ID','Дата оформления заказа','Дата возврата',
      'Статус возврата','Дата статуса','Статус компенсации','Дата статуса компенсации','Обязательность маркировки','Причина возврата',
      'Комментарий покупателя','Тип покупателя','Количество возвращаемых товаров','Отправление вскрыто','Целевое место назначения',
      'Адрес размещения','Переход в "В пункте выдачи"','Местоположение','Дата возврата продавцу','Кол-во дней хранения',
      'Последний день бесплатного размещения','Штрихкод возврата','Стоимость хранения','Стоимость утилизации','Стоимость товара',
      'Процент комиссии','Комиссия','Цена без комиссии'
    ],
    dateFields: ['Дата статуса']
  },
  'Реализация': {
    required: true,
    label: 'Реализация',
    headers: [
      '№ п/п','Название товара','Артикул','SKU','Штрих-код товара','Реализовано на сумму, руб.',
      'Выплаты по механикам лояльности партнёров, руб.','Баллы за скидки','Кол-во реализовано','Цена реализации, руб.',
      'Вознаграждение за продажу по категории (справочно)','Цена до скидок по поручению продавца, руб., (справочно)',
      'Базовое вознаграждение Ozon, руб.','Итого к начислению, руб.','Возвращено на сумму, руб.',
      'Выплаты по механикам лояльности партнёров при возврате, руб.','Баллы за скидки при возврате','Кол-во возвращено',
      'Цена реализации при возврате, руб.','Базовое вознаграждение Ozon при возврате, руб.','Итого возвращено, руб.',
      'Номер отправления','Дата отправления','Номер счет-фактуры','Дата счет-фактуры'
    ],
    dateFields: []
  }
};

const STATUS_GRID = document.getElementById('sheetStatusGrid');
const CSV_INPUTS = document.getElementById('csvInputs');
const CSV_MODE = document.getElementById('csvMode');
const RESULTS_SECTION = document.getElementById('resultsSection');
const ANALYZE_BTN = document.getElementById('analyzeBtn');
const RESET_BTN = document.getElementById('resetBtn');
const TEMPLATE_FILE_INPUT = document.getElementById('templateFile');
const MONTH_INPUT = document.getElementById('reportMonth');
const PRODUCT_SEARCH = document.getElementById('productSearch');
const COPY_SUMMARY_BTN = document.getElementById('copySummaryBtn');
const EXPORT_CSV_BTN = document.getElementById('exportCsvBtn');
const TOGGLE_CSV_BTN = document.getElementById('toggleCsvBtn');
const BUILD_FROM_CSV_BTN = document.getElementById('buildFromCsvBtn');

const DEFAULT_SETTINGS = {
  month: '',
  defaults: { costPrice: 0, packaging: 0, inbound: 0, otherVar: 0 },
  allocationMode: 'none',
  skuCosts: {}
};

const ACCRUAL_IGNORE_TYPES = new Set([
  'Выручка','Программы партнёров','Баллы за скидки','Вознаграждение за продажу',
  'Возврат выручки','Возврат вознаграждения'
]);

const state = {
  settings: loadSettings(),
  sheetData: null,
  sourceMode: null,
  analysis: null,
  csvFiles: {}
};

init();

function init() {
  if (!state.settings.month) {
    MONTH_INPUT.value = monthToInputValue(new Date());
  } else {
    MONTH_INPUT.value = state.settings.month;
  }

  document.getElementById('defaultCostPrice').value = state.settings.defaults.costPrice || '';
  document.getElementById('defaultPackaging').value = state.settings.defaults.packaging || '';
  document.getElementById('defaultInbound').value = state.settings.defaults.inbound || '';
  document.getElementById('defaultOtherVar').value = state.settings.defaults.otherVar || '';
  document.getElementById('allocationMode').value = state.settings.allocationMode || 'none';

  renderStatusCards({});
  renderCsvInputs();

  TEMPLATE_FILE_INPUT.addEventListener('change', handleWorkbookUpload);
  ANALYZE_BTN.addEventListener('click', rebuildAnalysis);
  RESET_BTN.addEventListener('click', resetAll);
  PRODUCT_SEARCH.addEventListener('input', renderProductTable);
  COPY_SUMMARY_BTN.addEventListener('click', copySummary);
  EXPORT_CSV_BTN.addEventListener('click', exportProductCsv);
  TOGGLE_CSV_BTN.addEventListener('click', () => {
    CSV_MODE.classList.toggle('hidden');
    TOGGLE_CSV_BTN.textContent = CSV_MODE.classList.contains('hidden') ? 'Показать' : 'Скрыть';
  });
  BUILD_FROM_CSV_BTN.addEventListener('click', handleCsvBuild);
  MONTH_INPUT.addEventListener('change', () => {
    state.settings.month = MONTH_INPUT.value;
    persistSettings();
    renderStatusCards(state.sheetData ? validateDataset(state.sheetData).status : {});
    if (state.analysis) rebuildAnalysis();
  });

  ['defaultCostPrice','defaultPackaging','defaultInbound','defaultOtherVar','allocationMode'].forEach((id) => {
    document.getElementById(id).addEventListener('input', handleSettingChange);
    document.getElementById(id).addEventListener('change', handleSettingChange);
  });
}

function handleSettingChange() {
  state.settings.defaults.costPrice = toNumber(document.getElementById('defaultCostPrice').value);
  state.settings.defaults.packaging = toNumber(document.getElementById('defaultPackaging').value);
  state.settings.defaults.inbound = toNumber(document.getElementById('defaultInbound').value);
  state.settings.defaults.otherVar = toNumber(document.getElementById('defaultOtherVar').value);
  state.settings.allocationMode = document.getElementById('allocationMode').value;
  persistSettings();
  if (state.analysis) rebuildAnalysis();
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return structuredClone(DEFAULT_SETTINGS);
    const parsed = JSON.parse(raw);
    return {
      ...structuredClone(DEFAULT_SETTINGS),
      ...parsed,
      defaults: { ...DEFAULT_SETTINGS.defaults, ...(parsed.defaults || {}) },
      skuCosts: parsed.skuCosts || {}
    };
  } catch (error) {
    return structuredClone(DEFAULT_SETTINGS);
  }
}

function persistSettings() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.settings));
}

function renderStatusCards(statusMap) {
  STATUS_GRID.innerHTML = '';
  Object.keys(TEMPLATE_SPECS).forEach((sheetName) => {
    const spec = TEMPLATE_SPECS[sheetName];
    const info = statusMap[sheetName] || { state: 'missing', rows: 0, message: 'Файл ещё не загружен.' };

    const card = document.createElement('div');
    card.className = 'status-card';

    const title = document.createElement('h3');
    title.textContent = spec.label;

    const meta = document.createElement('div');
    meta.className = 'status-meta';

    const requiredPill = document.createElement('span');
    requiredPill.className = `status-pill ${spec.required ? 'status-ok' : 'status-warn'}`;
    requiredPill.textContent = spec.required ? 'Обязательный лист' : 'Опционально';
    meta.appendChild(requiredPill);

    const statePill = document.createElement('span');
    statePill.className = `status-pill ${info.state === 'ok' ? 'status-ok' : info.state === 'warn' ? 'status-warn' : 'status-bad'}`;
    statePill.textContent = info.state === 'ok' ? 'Готово' : info.state === 'warn' ? 'Проверьте' : 'Нет данных';
    meta.appendChild(statePill);

    const rowText = document.createElement('p');
    rowText.textContent = `Строк: ${formatInteger(info.rows || 0)}`;

    const message = document.createElement('p');
    message.textContent = info.message || '';

    card.appendChild(title);
    card.appendChild(meta);
    card.appendChild(rowText);
    card.appendChild(message);
    STATUS_GRID.appendChild(card);
  });
}

function renderCsvInputs() {
  CSV_INPUTS.innerHTML = '';
  Object.keys(TEMPLATE_SPECS).forEach((sheetName) => {
    const spec = TEMPLATE_SPECS[sheetName];
    const card = document.createElement('div');
    card.className = 'csv-input-card';

    const title = document.createElement('h3');
    title.textContent = spec.label;

    const help = document.createElement('p');
    help.className = 'muted';
    help.textContent = `Загрузи CSV для вкладки «${spec.label}». Заголовки должны совпадать с шаблоном.`;

    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.csv,.txt';
    input.dataset.sheet = sheetName;
    input.addEventListener('change', (event) => {
      const file = event.target.files?.[0];
      state.csvFiles[sheetName] = file || null;
      card.dataset.ready = file ? '1' : '0';
    });

    card.appendChild(title);
    card.appendChild(help);
    card.appendChild(input);
    CSV_INPUTS.appendChild(card);
  });
}

async function handleWorkbookUpload(event) {
  const file = event.target.files?.[0];
  if (!file) return;

  if (!window.XLSX) {
    renderStatusCards({
      'Начисления': { state: 'warn', rows: 0, message: 'Библиотека чтения Excel недоступна. Используй резервный режим CSV ниже.' }
    });
    return;
  }

  try {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = window.XLSX.read(arrayBuffer, { type: 'array', cellDates: true });
    const data = extractWorkbookData(workbook);
    state.sheetData = data;
    state.sourceMode = 'xlsx';
    state.analysis = null;

    const validation = validateDataset(data);
    renderStatusCards(validation.status);

    if (validation.suggestedMonth) {
      MONTH_INPUT.value = validation.suggestedMonth;
      state.settings.month = validation.suggestedMonth;
      persistSettings();
    }

    ANALYZE_BTN.disabled = !validation.ready;
    RESULTS_SECTION.classList.add('hidden');
  } catch (error) {
    console.error(error);
    renderStatusCards({
      'Начисления': { state: 'warn', rows: 0, message: `Не удалось прочитать книгу: ${error.message}` }
    });
    ANALYZE_BTN.disabled = true;
    RESULTS_SECTION.classList.add('hidden');
  }
}

async function handleCsvBuild() {
  try {
    const dataset = {};
    for (const sheetName of Object.keys(TEMPLATE_SPECS)) {
      const file = state.csvFiles[sheetName];
      if (!file) throw new Error(`Не загружен CSV для листа «${sheetName}».`);
      const text = await file.text();
      dataset[sheetName] = parseCsvSheet(text, TEMPLATE_SPECS[sheetName]);
    }
    state.sheetData = dataset;
    state.sourceMode = 'csv';
    state.analysis = null;
    const validation = validateDataset(dataset);
    if (validation.suggestedMonth) {
      MONTH_INPUT.value = validation.suggestedMonth;
      state.settings.month = validation.suggestedMonth;
      persistSettings();
    }
    renderStatusCards(validation.status);
    ANALYZE_BTN.disabled = !validation.ready;
    RESULTS_SECTION.classList.add('hidden');
  } catch (error) {
    console.error(error);
    alert(error.message);
  }
}

function extractWorkbookData(workbook) {
  const result = {};
  for (const [sheetName, spec] of Object.entries(TEMPLATE_SPECS)) {
    const sheet = workbook.Sheets[sheetName];
    if (!sheet) {
      result[sheetName] = null;
      continue;
    }
    const aoa = window.XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: null, blankrows: false });
    result[sheetName] = parseArraySheet(aoa, spec);
  }
  return result;
}

function parseArraySheet(aoa, spec) {
  const rows = (aoa || []).map((row) => trimTrailingBlanks(row));
  const nonEmptyRows = rows.filter((row) => row.some((cell) => !isBlank(cell)));
  if (!nonEmptyRows.length) throw new Error(`Лист «${spec.label}» пуст.`);
  const header = trimTrailingBlanks(nonEmptyRows[0]);
  const expected = spec.headers;
  const actualNormalized = header.map(normalizeHeader);
  const expectedNormalized = expected.map(normalizeHeader);
  const prefix = actualNormalized.slice(0, expected.length);

  if (JSON.stringify(prefix) !== JSON.stringify(expectedNormalized)) {
    throw new Error(`Лист «${spec.label}» не похож на нормализованный шаблон. Проверь строку 1 и порядок столбцов.`);
  }

  const records = [];
  for (let i = 1; i < nonEmptyRows.length; i += 1) {
    const row = nonEmptyRows[i];
    const values = row.slice(0, expected.length);
    if (!values.some((cell) => !isBlank(cell))) continue;
    const record = {};
    expected.forEach((headerName, index) => {
      record[headerName] = normalizeCellValue(values[index]);
    });
    records.push(record);
  }
  return records;
}

function parseCsvSheet(text, spec) {
  const delimiter = detectDelimiter(text);
  const rows = parseDelimitedText(text, delimiter).map((row) => trimTrailingBlanks(row));
  return parseArraySheet(rows, spec);
}

function detectDelimiter(text) {
  const firstLine = (text || '').split(/\r?\n/).find((line) => line.trim().length > 0) || '';
  const delimiters = ['\t', ';', ','];
  let best = ';';
  let bestCount = -1;
  delimiters.forEach((delimiter) => {
    const count = firstLine.split(delimiter).length;
    if (count > bestCount) {
      best = delimiter;
      bestCount = count;
    }
  });
  return best;
}

function parseDelimitedText(text, delimiter) {
  const rows = [];
  let current = '';
  let row = [];
  let inQuotes = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];
    const next = text[i + 1];

    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (!inQuotes && char === delimiter) {
      row.push(current);
      current = '';
      continue;
    }

    if (!inQuotes && (char === '\n' || char === '\r')) {
      if (char === '\r' && next === '\n') i += 1;
      row.push(current);
      rows.push(row);
      row = [];
      current = '';
      continue;
    }

    current += char;
  }

  if (current.length > 0 || row.length > 0) {
    row.push(current);
    rows.push(row);
  }

  return rows;
}

function validateDataset(dataset) {
  const status = {};
  let ready = true;
  const months = [];

  for (const [sheetName, spec] of Object.entries(TEMPLATE_SPECS)) {
    const rows = dataset?.[sheetName];
    if (!rows || !rows.length) {
      status[sheetName] = {
        state: 'bad',
        rows: 0,
        message: 'Нет данных на этом листе.'
      };
      if (spec.required) ready = false;
      continue;
    }

    let stateLevel = 'ok';
    let message = 'Структура соответствует шаблону.';
    const monthCheck = checkMonthConsistency(rows, spec.dateFields, MONTH_INPUT.value);
    if (monthCheck.externalRows > 0) {
      stateLevel = 'warn';
      message = `Есть строки вне выбранного месяца: ${formatInteger(monthCheck.externalRows)}. Проверь диапазон копирования.`;
    }
    if (monthCheck.primaryMonth) months.push(monthCheck.primaryMonth);

    status[sheetName] = {
      state: stateLevel,
      rows: rows.length,
      message
    };
  }

  const suggestedMonth = months.length ? mostFrequent(months) : '';
  return { ready, status, suggestedMonth };
}

function checkMonthConsistency(rows, dateFields, selectedMonth) {
  if (!dateFields || !dateFields.length) return { externalRows: 0, primaryMonth: '' };
  const monthCounts = {};
  let externalRows = 0;

  rows.forEach((row) => {
    const monthsInRow = dateFields
      .map((field) => parseDate(row[field]))
      .filter(Boolean)
      .map((date) => monthToInputValue(date));
    if (!monthsInRow.length) return;
    const primary = monthsInRow[0];
    monthCounts[primary] = (monthCounts[primary] || 0) + 1;
    if (selectedMonth && primary !== selectedMonth) externalRows += 1;
  });

  const primaryMonth = Object.keys(monthCounts).sort((a, b) => monthCounts[b] - monthCounts[a])[0] || '';
  return { externalRows, primaryMonth };
}

function rebuildAnalysis() {
  if (!state.sheetData) return;
  const validation = validateDataset(state.sheetData);
  renderStatusCards(validation.status);
  if (!validation.ready) {
    alert('Сначала загрузи все обязательные листы из нормализованного шаблона.');
    return;
  }

  const analysis = buildAnalysis(state.sheetData, state.settings);
  state.analysis = analysis;
  RESULTS_SECTION.classList.remove('hidden');

  renderKpis(analysis);
  renderDailyChart(analysis.dailySales);
  renderExpenseChart(analysis);
  renderTopProductsChart(analysis.products);
  renderReturnReasons(analysis.returnReasons);
  renderSharedExpenses(analysis.sharedExpenses);
  renderInsights(analysis);
  renderProductTable();
}

function buildAnalysis(dataset, settings) {
  const accruals = dataset['Начисления'] || [];
  const orders = dataset['Заказы'] || [];
  const promo = dataset['Продвижение'] || [];
  const returns = dataset['Возвраты'] || [];
  const realization = dataset['Реализация'] || [];

  const hasPromoSheet = promo.length > 0;
  const products = new Map();
  const dailySales = new Map();
  const returnReasons = new Map();
  const sharedExpenses = new Map();
  const deliveredOrderSet = new Set();
  const allOrderSet = new Set();

  const totals = {
    grossSales: 0,
    financialReturns: 0,
    netSales: 0,
    soldQty: 0,
    returnQtyFinancial: 0,
    operationalReturnsQty: 0,
    orders: 0,
    deliveredQty: 0,
    cancelledQty: 0,
    openQty: 0,
    payoutBeforeOps: 0,
    baseCommissionNet: 0,
    acquiring: 0,
    logistics: 0,
    storage: 0,
    promotion: 0,
    penalties: 0,
    otherMp: 0,
    compensation: 0,
    manualCosts: 0,
    sharedNet: 0,
    directContribution: 0,
    finalContribution: 0
  };

  function ensureProduct(sku, article = '', name = '') {
    const key = normalizeSku(sku);
    if (!key) return null;
    if (!products.has(key)) {
      products.set(key, {
        sku: key,
        article: article || '',
        name: name || '',
        grossSales: 0,
        financialReturns: 0,
        netSales: 0,
        soldQty: 0,
        returnQtyFinancial: 0,
        operationalReturnsQty: 0,
        ordersQty: 0,
        deliveredQty: 0,
        cancelledQty: 0,
        ordersSet: new Set(),
        payoutBeforeOps: 0,
        baseCommission: 0,
        returnCommissionRefund: 0,
        baseCommissionNet: 0,
        acquiring: 0,
        logistics: 0,
        storage: 0,
        promotion: 0,
        promoExtra: 0,
        penalties: 0,
        otherMp: 0,
        compensation: 0,
        adSales: 0,
        adOrders: 0,
        impressions: 0,
        clicks: 0,
        reasons: new Map(),
        manualCostPerUnit: 0,
        manualCosts: 0,
        directContribution: 0,
        sharedAllocated: 0,
        finalContribution: 0
      });
    }
    const item = products.get(key);
    if (!item.article && article) item.article = article;
    if (!item.name && name) item.name = name;
    return item;
  }

  realization.forEach((row) => {
    const item = ensureProduct(row['SKU'], row['Артикул'], row['Название товара']);
    if (!item) return;

    item.grossSales += toNumber(row['Реализовано на сумму, руб.']);
    item.financialReturns += toNumber(row['Возвращено на сумму, руб.']);
    item.soldQty += toNumber(row['Кол-во реализовано']);
    item.returnQtyFinancial += toNumber(row['Кол-во возвращено']);
    item.payoutBeforeOps += toNumber(row['Итого к начислению, руб.']) - toNumber(row['Итого возвращено, руб.']);
    item.baseCommission += toNumber(row['Базовое вознаграждение Ozon, руб.']);
    item.returnCommissionRefund += toNumber(row['Базовое вознаграждение Ozon при возврате, руб.']);

    const date = parseDate(row['Дата отправления']) || parseDate(row['Дата счет-фактуры']);
    if (date) {
      const key = formatDateKey(date);
      const prev = dailySales.get(key) || 0;
      dailySales.set(key, prev + toNumber(row['Реализовано на сумму, руб.']) - toNumber(row['Возвращено на сумму, руб.']));
    }
  });

  orders.forEach((row) => {
    const item = ensureProduct(row['SKU'], row['Артикул'], row['Название товара']);
    if (!item) return;

    const qty = toNumber(row['Количество']) || 1;
    const orderNo = String(row['Номер заказа'] || '').trim();
    const status = String(row['Статус'] || '').trim();

    item.ordersQty += qty;
    if (orderNo) {
      item.ordersSet.add(orderNo);
      allOrderSet.add(orderNo);
    }

    if (status.toLowerCase().includes('доставлен')) {
      item.deliveredQty += qty;
      if (orderNo) deliveredOrderSet.add(orderNo);
    } else if (status.toLowerCase().includes('отмен')) {
      item.cancelledQty += qty;
    } else {
      totals.openQty += qty;
    }
  });

  returns.forEach((row) => {
    const item = ensureProduct(row['OZON SKU ID'], row['Артикул товара'], row['Наименование товара']);
    if (!item) return;

    const qty = toNumber(row['Количество возвращаемых товаров']) || 1;
    item.operationalReturnsQty += qty;

    const reason = String(row['Причина возврата'] || 'Причина не указана').trim() || 'Причина не указана';
    item.reasons.set(reason, (item.reasons.get(reason) || 0) + qty);
    returnReasons.set(reason, (returnReasons.get(reason) || 0) + qty);
  });

  promo.forEach((row) => {
    const item = ensureProduct(row['SKU'], '', row['Название товара']);
    if (!item) return;

    item.promotion += toNumber(row['Расход, ₽']);
    item.adSales += toNumber(row['Продажи, ₽']);
    item.adOrders += toNumber(row['Заказы, шт']);
    item.impressions += toNumber(row['Показы']);
    item.clicks += toNumber(row['Клики']);
  });

  accruals.forEach((row) => {
    const amount = toNumber(row['Сумма итого, руб.']);
    const type = String(row['Тип начисления'] || '').trim();
    const group = String(row['Группа услуг'] || '').trim();
    const sku = normalizeSku(row['SKU']);
    const article = row['Артикул'];
    const name = row['Название товара'];

    if (ACCRUAL_IGNORE_TYPES.has(type)) return;
    if (hasPromoSheet && type === 'Оплата за клик') return;

    const bucket = classifyAccrualBucket(type, group);
    const target = sku ? ensureProduct(sku, article, name) : null;

    if (!bucket) return;

    if (target) {
      applyAccrualToTarget(target, bucket, amount);
    } else {
      applySharedAccrual(sharedExpenses, bucket, amount);
    }
  });

  let totalNetUnits = 0;
  let totalNetSales = 0;
  let sharedNet = 0;

  products.forEach((item) => {
    item.baseCommissionNet = item.baseCommission - item.returnCommissionRefund;
    item.netSales = item.grossSales - item.financialReturns;
    item.netUnits = Math.max(0, item.soldQty - item.returnQtyFinancial);

    const skuCostSettings = state.settings.skuCosts[item.sku] || {};
    const unitCostPrice = chooseValue(skuCostSettings.costPrice, settings.defaults.costPrice);
    const unitPackaging = chooseValue(skuCostSettings.packaging, settings.defaults.packaging);
    const unitInbound = chooseValue(skuCostSettings.inbound, settings.defaults.inbound);
    const unitOther = chooseValue(skuCostSettings.otherVar, settings.defaults.otherVar);

    item.manualCostPerUnit = unitCostPrice + unitPackaging + unitInbound + unitOther;
    item.manualCosts = item.manualCostPerUnit * item.netUnits;
    item.directMarketplaceCosts = item.acquiring + item.logistics + item.storage + item.promotion + item.promoExtra + item.penalties + item.otherMp - item.compensation;
    item.directContribution = item.payoutBeforeOps - item.acquiring - item.logistics - item.storage - item.promotion - item.promoExtra - item.penalties - item.otherMp + item.compensation - item.manualCosts;
    item.drr = item.netSales > 0 ? item.promotion / item.netSales : 0;
    item.directMargin = item.netSales > 0 ? item.directContribution / item.netSales : 0;

    totalNetUnits += item.netUnits;
    totalNetSales += item.netSales;
  });

  sharedExpenses.forEach((value) => {
    sharedNet += value.net;
  });

  products.forEach((item) => {
    const weight = calculateAllocationWeight(item, settings.allocationMode, totalNetSales, totalNetUnits);
    item.sharedAllocated = sharedNet * weight;
    item.finalContribution = item.directContribution - item.sharedAllocated;
    item.finalMargin = item.netSales > 0 ? item.finalContribution / item.netSales : 0;
  });

  const productList = Array.from(products.values()).map((item) => ({
    ...item,
    ordersCount: item.ordersSet.size
  }));

  productList.sort((a, b) => b.finalContribution - a.finalContribution);

  productList.forEach((item) => {
    totals.grossSales += item.grossSales;
    totals.financialReturns += item.financialReturns;
    totals.netSales += item.netSales;
    totals.soldQty += item.soldQty;
    totals.returnQtyFinancial += item.returnQtyFinancial;
    totals.operationalReturnsQty += item.operationalReturnsQty;
    totals.deliveredQty += item.deliveredQty;
    totals.cancelledQty += item.cancelledQty;
    totals.payoutBeforeOps += item.payoutBeforeOps;
    totals.baseCommissionNet += item.baseCommissionNet;
    totals.acquiring += item.acquiring;
    totals.logistics += item.logistics;
    totals.storage += item.storage;
    totals.promotion += item.promotion + item.promoExtra;
    totals.penalties += item.penalties;
    totals.otherMp += item.otherMp;
    totals.compensation += item.compensation;
    totals.manualCosts += item.manualCosts;
    totals.directContribution += item.directContribution;
    totals.finalContribution += item.finalContribution;
  });

  sharedExpenses.forEach((value) => {
    totals.sharedNet += value.net;
  });

  totals.orders = allOrderSet.size;
  totals.averageCheck = deliveredOrderSet.size ? totals.netSales / deliveredOrderSet.size : 0;
  totals.openQty += 0;

  return {
    monthLabel: formatMonthLabel(MONTH_INPUT.value || guessMonthFromTotals(dataset) || monthToInputValue(new Date())),
    products: productList,
    totals,
    dailySales: Array.from(dailySales.entries()).sort(([a], [b]) => a.localeCompare(b)).map(([date, amount]) => ({ date, amount })),
    returnReasons: Array.from(returnReasons.entries()).sort((a, b) => b[1] - a[1]).map(([reason, qty]) => ({ reason, qty })),
    sharedExpenses: Array.from(sharedExpenses.entries()).sort((a, b) => Math.abs(b[1].net) - Math.abs(a[1].net)).map(([label, data]) => ({ label, ...data })),
    allocationMode: settings.allocationMode
  };
}

function classifyAccrualBucket(type, group) {
  if (['Эквайринг'].includes(type)) return 'Эквайринг';
  if (['Логистика','Доставка до места выдачи','Обратная логистика','Обработка отправления Drop-off (ПВЗ)','Обработка отправления Drop-off партнёрами (ПВЗ)','Обработка возвратов, отмен и невыкупов партнёрами','Обработка отменённых и невостребованных товаров','Обработка возвратов Ozon'].includes(type)) {
    return 'Логистика и обработка';
  }
  if (['Кросс-докинг','Размещение товаров на складах Ozon','Упаковка товара партнёрами','Обеспечение материалами для упаковки товара'].includes(type)) {
    return 'FBO / хранение / упаковка';
  }
  if (['Звёздные товары','Продвижение с оплатой за заказ','Подписка Premium','Оплата за клик'].includes(type)) {
    return 'Продвижение';
  }
  if (type.toLowerCase().includes('слот') || type.toLowerCase().includes('просроч') || group.toLowerCase().includes('штраф')) {
    return 'Штрафы';
  }
  if (type === 'Взаимозачет требований между Договорами') return 'Прочие удержания';
  if (group.toLowerCase().includes('компенсац') || type.toLowerCase().includes('потеря')) return 'Компенсации';
  return 'Прочие удержания';
}

function applyAccrualToTarget(target, bucket, amount) {
  if (bucket === 'Компенсации') {
    target.compensation += amount;
    return;
  }
  const expense = amount < 0 ? -amount : 0;
  const income = amount > 0 ? amount : 0;
  if (bucket === 'Эквайринг') target.acquiring += expense - income;
  else if (bucket === 'Логистика и обработка') target.logistics += expense - income;
  else if (bucket === 'FBO / хранение / упаковка') target.storage += expense - income;
  else if (bucket === 'Продвижение') target.promoExtra += expense - income;
  else if (bucket === 'Штрафы') target.penalties += expense - income;
  else target.otherMp += expense - income;
}

function applySharedAccrual(sharedExpenses, bucket, amount) {
  if (!sharedExpenses.has(bucket)) {
    sharedExpenses.set(bucket, { expense: 0, income: 0, net: 0 });
  }
  const item = sharedExpenses.get(bucket);
  if (bucket === 'Компенсации') {
    if (amount >= 0) item.income += amount;
    else item.expense += -amount;
  } else {
    if (amount < 0) item.expense += -amount;
    else item.income += amount;
  }
  item.net = item.expense - item.income;
}

function calculateAllocationWeight(item, mode, totalNetSales, totalNetUnits) {
  if (mode === 'revenue') {
    return totalNetSales > 0 ? item.netSales / totalNetSales : 0;
  }
  if (mode === 'units') {
    return totalNetUnits > 0 ? item.netUnits / totalNetUnits : 0;
  }
  return 0;
}

function renderKpis(analysis) {
  const grid = document.getElementById('kpiGrid');
  const totalExtraMp = analysis.totals.acquiring + analysis.totals.logistics + analysis.totals.storage + analysis.totals.promotion + analysis.totals.penalties + analysis.totals.otherMp - analysis.totals.compensation + analysis.totals.sharedNet;
  const items = [
    { label: 'Выручка нетто', value: formatMoney(analysis.totals.netSales), foot: `За вычетом финансовых возвратов. ${analysis.monthLabel}` },
    { label: 'Заказы', value: formatInteger(analysis.totals.orders), foot: `Средний чек: ${formatMoney(analysis.totals.averageCheck)}` },
    { label: 'Выкуплено, шт', value: formatInteger(analysis.totals.deliveredQty), foot: `По статусам из листа «Заказы».` },
    { label: 'Отмены, шт', value: formatInteger(analysis.totals.cancelledQty), foot: `Операционные отмены до выкупа.` },
    { label: 'Операционные возвраты, шт', value: formatInteger(analysis.totals.operationalReturnsQty), foot: `По листу «Возвраты».` },
    { label: 'К начислению по реализации', value: formatMoney(analysis.totals.payoutBeforeOps), foot: `Уже после базовой комиссии Ozon.` },
    { label: 'Доп. расходы Ozon + ручные', value: formatMoney(totalExtraMp + analysis.totals.manualCosts), foot: `Логистика, эквайринг, реклама, FBO, штрафы и ручные затраты.` },
    { label: analysis.allocationMode === 'none' ? 'Вклад до общих расходов' : 'Вклад после распределения', value: formatMoney(analysis.allocationMode === 'none' ? analysis.totals.directContribution : analysis.totals.finalContribution), foot: analysis.allocationMode === 'none' ? 'Общие расходы без SKU не распределены по товарам.' : 'С учётом выбранного правила распределения.' }
  ];

  grid.innerHTML = '';
  items.forEach((item) => {
    const card = document.createElement('div');
    card.className = 'kpi-card';
    card.innerHTML = `
      <div class="kpi-label">${escapeHtml(item.label)}</div>
      <div class="kpi-value">${escapeHtml(item.value)}</div>
      <div class="kpi-foot">${escapeHtml(item.foot)}</div>
    `;
    grid.appendChild(card);
  });
}

function renderDailyChart(points) {
  const box = document.getElementById('dailyChart');
  if (!points.length) {
    box.innerHTML = '<div class="muted">Недостаточно дат для графика.</div>';
    return;
  }

  const width = 920;
  const height = 320;
  const padding = { top: 20, right: 20, bottom: 42, left: 50 };
  const innerWidth = width - padding.left - padding.right;
  const innerHeight = height - padding.top - padding.bottom;
  const maxValue = Math.max(...points.map((point) => point.amount), 1);
  const barWidth = innerWidth / Math.max(points.length, 1) - 4;

  const bars = points.map((point, index) => {
    const x = padding.left + index * (barWidth + 4);
    const h = (Math.max(point.amount, 0) / maxValue) * innerHeight;
    const y = padding.top + innerHeight - h;
    const label = point.date.slice(8);
    return `
      <rect x="${x}" y="${y}" width="${Math.max(barWidth, 4)}" height="${h}" rx="5" fill="rgba(47,109,246,0.9)"></rect>
      <text class="axis-label" x="${x + Math.max(barWidth, 4) / 2}" y="${height - 14}" text-anchor="middle">${label}</text>
    `;
  }).join('');

  const gridLines = [0, 0.25, 0.5, 0.75, 1].map((ratio) => {
    const y = padding.top + innerHeight - ratio * innerHeight;
    const label = formatCompactMoney(maxValue * ratio);
    return `
      <line x1="${padding.left}" y1="${y}" x2="${width - padding.right}" y2="${y}" stroke="#e6eef8" stroke-width="1"></line>
      <text class="axis-label" x="${padding.left - 10}" y="${y + 4}" text-anchor="end">${label}</text>
    `;
  }).join('');

  box.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" role="img" aria-label="Дневная выручка">
      ${gridLines}
      ${bars}
    </svg>
  `;
}

function renderExpenseChart(analysis) {
  const rows = [
    { label: 'Логистика и обработка', value: analysis.totals.logistics },
    { label: 'Эквайринг', value: analysis.totals.acquiring },
    { label: 'FBO / хранение / упаковка', value: analysis.totals.storage },
    { label: 'Реклама и продвижение', value: analysis.totals.promotion },
    { label: 'Штрафы', value: analysis.totals.penalties },
    { label: 'Прочие удержания', value: analysis.totals.otherMp },
    { label: 'Ручные затраты', value: analysis.totals.manualCosts },
    { label: 'Общие расходы без SKU', value: analysis.totals.sharedNet }
  ].filter((row) => Math.abs(row.value) > 0.005);

  const colors = ['#2f6df6','#4e8cff','#6aa0ff','#8ab4ff','#a4c6ff','#bfd8ff','#89c29b','#f0b95f'];
  const box = document.getElementById('expenseChart');
  const legend = document.getElementById('expenseLegend');

  if (!rows.length) {
    box.innerHTML = '<div class="muted">Нет расходов для построения структуры.</div>';
    legend.innerHTML = '';
    return;
  }

  const maxValue = Math.max(...rows.map((row) => Math.abs(row.value)), 1);
  const width = 920;
  const height = 320;
  const padding = { top: 20, right: 20, bottom: 20, left: 220 };
  const innerWidth = width - padding.left - padding.right;
  const barHeight = 26;
  const gap = 14;

  const bars = rows.map((row, index) => {
    const y = padding.top + index * (barHeight + gap);
    const w = (Math.abs(row.value) / maxValue) * innerWidth;
    return `
      <text class="bar-label" x="14" y="${y + 18}">${escapeHtml(row.label)}</text>
      <rect x="${padding.left}" y="${y}" width="${w}" height="${barHeight}" rx="10" fill="${colors[index % colors.length]}"></rect>
      <text class="value-label" x="${padding.left + w + 8}" y="${y + 18}">${escapeHtml(formatMoney(row.value))}</text>
    `;
  }).join('');

  box.innerHTML = `<svg viewBox="0 0 ${width} ${height}">${bars}</svg>`;

  legend.innerHTML = '';
  rows.forEach((row, index) => {
    const item = document.createElement('div');
    item.className = 'legend-item';
    item.innerHTML = `
      <div class="legend-item-main">
        <span class="legend-color" style="background:${colors[index % colors.length]}"></span>
        <span class="legend-name">${escapeHtml(row.label)}</span>
      </div>
      <span class="legend-value">${escapeHtml(formatMoney(row.value))}</span>
    `;
    legend.appendChild(item);
  });
}

function renderTopProductsChart(products) {
  const box = document.getElementById('topProductsChart');
  const top = [...products].sort((a, b) => b.finalContribution - a.finalContribution).slice(0, 8);
  if (!top.length) {
    box.innerHTML = '<div class="muted">Нет данных по товарам.</div>';
    return;
  }

  const maxValue = Math.max(...top.map((item) => Math.max(item.finalContribution, 0.01)));
  const width = 920;
  const height = 320;
  const padding = { top: 20, right: 20, bottom: 20, left: 260 };
  const innerWidth = width - padding.left - padding.right;
  const barHeight = 24;
  const gap = 12;

  const bars = top.map((item, index) => {
    const y = padding.top + index * (barHeight + gap);
    const w = Math.max(0, (item.finalContribution / maxValue) * innerWidth);
    return `
      <text class="bar-label" x="14" y="${y + 17}">${escapeHtml(shorten(item.article || item.name, 30))}</text>
      <rect x="${padding.left}" y="${y}" width="${w}" height="${barHeight}" rx="10" fill="rgba(29,156,109,0.82)"></rect>
      <text class="value-label" x="${padding.left + w + 8}" y="${y + 17}">${escapeHtml(formatMoney(item.finalContribution))}</text>
    `;
  }).join('');

  box.innerHTML = `<svg viewBox="0 0 ${width} ${height}">${bars}</svg>`;
}

function renderReturnReasons(reasons) {
  const list = document.getElementById('returnReasonsList');
  list.innerHTML = '';
  if (!reasons.length) {
    list.innerHTML = '<div class="muted">На листе возвратов нет строк.</div>';
    return;
  }
  reasons.slice(0, 8).forEach((item) => {
    const row = document.createElement('div');
    row.className = 'reason-item';
    row.innerHTML = `
      <div class="reason-item-main">
        <span class="reason-name">${escapeHtml(item.reason)}</span>
      </div>
      <span class="reason-value">${escapeHtml(formatInteger(item.qty))} шт</span>
    `;
    list.appendChild(row);
  });
}

function renderSharedExpenses(sharedExpenses) {
  const list = document.getElementById('sharedExpenseList');
  list.innerHTML = '';
  if (!sharedExpenses.length) {
    list.innerHTML = '<div class="muted">Все дополнительные расходы удалось разложить по SKU.</div>';
    return;
  }

  sharedExpenses.forEach((item) => {
    const row = document.createElement('div');
    row.className = 'expense-item';
    row.innerHTML = `
      <div class="expense-item-main">
        <span class="expense-name">${escapeHtml(item.label)}</span>
      </div>
      <span class="expense-value">${escapeHtml(formatMoney(item.net))}</span>
    `;
    list.appendChild(row);
  });
}

function renderInsights(analysis) {
  const list = document.getElementById('insightsList');
  list.innerHTML = '';

  const topByRevenue = [...analysis.products].sort((a, b) => b.netSales - a.netSales)[0];
  const topByContribution = [...analysis.products].sort((a, b) => b.finalContribution - a.finalContribution)[0];
  const weakByContribution = [...analysis.products].sort((a, b) => a.finalContribution - b.finalContribution)[0];
  const heavyReturns = [...analysis.products].sort((a, b) => b.operationalReturnsQty - a.operationalReturnsQty)[0];

  const messages = [
    {
      title: 'Главный драйвер выручки',
      text: topByRevenue ? `${topByRevenue.article || topByRevenue.sku} даёт ${formatMoney(topByRevenue.netSales)} нетто-выручки.` : 'Нет данных.'
    },
    {
      title: 'Лучший вклад в экономику',
      text: topByContribution ? `${topByContribution.article || topByContribution.sku} приносит ${formatMoney(topByContribution.finalContribution)} после выбранных правил распределения.` : 'Нет данных.'
    },
    {
      title: 'Слабое место',
      text: weakByContribution ? `${weakByContribution.article || weakByContribution.sku} показывает ${formatMoney(weakByContribution.finalContribution)}. Его стоит проверить по цене, рекламе и возвратам.` : 'Нет данных.'
    },
    {
      title: 'Риски по возвратам',
      text: heavyReturns && heavyReturns.operationalReturnsQty > 0 ? `${heavyReturns.article || heavyReturns.sku} имеет ${formatInteger(heavyReturns.operationalReturnsQty)} операционных возвратов.` : 'Критичных возвратов по товарам не видно.'
    },
    {
      title: 'Что ещё не разложено по товарам',
      text: analysis.totals.sharedNet > 0 ? `${formatMoney(analysis.totals.sharedNet)} остаётся в общих расходах без SKU.` : 'Нераспределённых расходов без SKU почти нет.'
    }
  ];

  messages.forEach((item) => {
    const row = document.createElement('div');
    row.className = 'insight-item';
    row.innerHTML = `<div><strong>${escapeHtml(item.title)}</strong><p>${escapeHtml(item.text)}</p></div>`;
    list.appendChild(row);
  });
}

const TABLE_COLUMNS = [
  { key: 'sku', label: 'SKU', type: 'text' },
  { key: 'article', label: 'Артикул', type: 'text' },
  { key: 'name', label: 'Товар', type: 'text' },
  { key: 'netUnits', label: 'Чистые ед.', type: 'number' },
  { key: 'ordersCount', label: 'Заказы', type: 'number' },
  { key: 'cancelledQty', label: 'Отмены', type: 'number' },
  { key: 'operationalReturnsQty', label: 'Возвраты', type: 'number' },
  { key: 'netSales', label: 'Выручка нетто', type: 'money' },
  { key: 'payoutBeforeOps', label: 'К начислению', type: 'money' },
  { key: 'baseCommissionNet', label: 'Базовая комиссия', type: 'money' },
  { key: 'acquiring', label: 'Эквайринг', type: 'money' },
  { key: 'logistics', label: 'Логистика', type: 'money' },
  { key: 'storage', label: 'FBO / хранение', type: 'money' },
  { key: 'promotion', label: 'Реклама по SKU', type: 'money' },
  { key: 'promoExtra', label: 'Прочее продвижение', type: 'money' },
  { key: 'otherMpTotal', label: 'Прочие расходы МП', type: 'money' },
  { key: 'manualCosts', label: 'Ручные затраты', type: 'money' },
  { key: 'sharedAllocated', label: 'Распределено общих', type: 'money' },
  { key: 'finalContribution', label: 'Итоговый вклад', type: 'money' },
  { key: 'finalMargin', label: 'Маржа', type: 'percent' },
  { key: 'drr', label: 'ДРР', type: 'percent' },
  { key: 'manualInputs', label: 'Ручной ввод', type: 'manual' }
];

let tableSort = { key: 'finalContribution', desc: true };

function renderProductTable() {
  if (!state.analysis) return;
  const thead = document.querySelector('#productTable thead');
  const tbody = document.querySelector('#productTable tbody');

  const headers = TABLE_COLUMNS.map((column) => `<th data-key="${column.key}">${escapeHtml(column.label)}</th>`).join('');
  thead.innerHTML = `<tr>${headers}</tr>`;

  thead.querySelectorAll('th').forEach((th) => {
    th.addEventListener('click', () => {
      const key = th.dataset.key;
      if (tableSort.key === key) {
        tableSort.desc = !tableSort.desc;
      } else {
        tableSort = { key, desc: key !== 'sku' && key !== 'article' && key !== 'name' };
      }
      renderProductTable();
    });
  });

  const search = PRODUCT_SEARCH.value.trim().toLowerCase();
  const rows = state.analysis.products
    .map((item) => ({
      ...item,
      otherMpTotal: item.penalties + item.otherMp
    }))
    .filter((item) => {
      if (!search) return true;
      return [item.sku, item.article, item.name].some((value) => String(value || '').toLowerCase().includes(search));
    })
    .sort((a, b) => sortRows(a, b, tableSort));

  tbody.innerHTML = rows.map((item) => renderProductRow(item)).join('');

  tbody.querySelectorAll('input[data-sku]').forEach((input) => {
    input.addEventListener('change', handleManualSkuInput);
  });
}

function renderProductRow(item) {
  const cells = TABLE_COLUMNS.map((column) => {
    if (column.type === 'manual') {
      const skuSettings = state.settings.skuCosts[item.sku] || {};
      return `
        <td>
          <div style="display:grid;grid-template-columns:repeat(2,minmax(90px,1fr));gap:8px">
            <input data-sku="${escapeHtml(item.sku)}" data-field="costPrice" type="number" step="0.01" value="${safeInputValue(skuSettings.costPrice)}" placeholder="Себестоим." />
            <input data-sku="${escapeHtml(item.sku)}" data-field="packaging" type="number" step="0.01" value="${safeInputValue(skuSettings.packaging)}" placeholder="Упаковка" />
            <input data-sku="${escapeHtml(item.sku)}" data-field="inbound" type="number" step="0.01" value="${safeInputValue(skuSettings.inbound)}" placeholder="До склада" />
            <input data-sku="${escapeHtml(item.sku)}" data-field="otherVar" type="number" step="0.01" value="${safeInputValue(skuSettings.otherVar)}" placeholder="Прочее" />
          </div>
        </td>
      `;
    }

    const rawValue = item[column.key];
    let content = '';
    let className = '';
    if (column.type === 'money') {
      content = formatMoney(rawValue);
      if (rawValue < 0) className = 'cell-negative';
      if (rawValue > 0 && ['finalContribution'].includes(column.key)) className = 'cell-positive';
    } else if (column.type === 'number') {
      content = formatInteger(rawValue || 0);
    } else if (column.type === 'percent') {
      content = formatPercent(rawValue || 0);
      if (column.key === 'finalMargin') {
        className = rawValue >= 0 ? 'cell-positive' : 'cell-negative';
      }
    } else {
      content = escapeHtml(shorten(rawValue || '', column.key === 'name' ? 52 : 18));
      className = column.key === 'name' ? '' : 'cell-strong';
    }
    return `<td class="${className}">${content}</td>`;
  }).join('');
  return `<tr>${cells}</tr>`;
}

function handleManualSkuInput(event) {
  const { sku, field } = event.target.dataset;
  const value = event.target.value;
  if (!state.settings.skuCosts[sku]) state.settings.skuCosts[sku] = {};
  if (value === '') {
    delete state.settings.skuCosts[sku][field];
    if (!Object.keys(state.settings.skuCosts[sku]).length) delete state.settings.skuCosts[sku];
  } else {
    state.settings.skuCosts[sku][field] = toNumber(value);
  }
  persistSettings();
  rebuildAnalysis();
}

function copySummary() {
  if (!state.analysis) return;
  const top = state.analysis.products.slice(0, 5);
  const lines = [
    `Ozon · ${state.analysis.monthLabel}`,
    `Выручка нетто: ${formatMoney(state.analysis.totals.netSales)}`,
    `Заказы: ${formatInteger(state.analysis.totals.orders)}`,
    `Выкуплено, шт: ${formatInteger(state.analysis.totals.deliveredQty)}`,
    `Отмены, шт: ${formatInteger(state.analysis.totals.cancelledQty)}`,
    `Операционные возвраты, шт: ${formatInteger(state.analysis.totals.operationalReturnsQty)}`,
    `К начислению по реализации: ${formatMoney(state.analysis.totals.payoutBeforeOps)}`,
    `Прямой вклад: ${formatMoney(state.analysis.totals.directContribution)}`,
    state.analysis.allocationMode === 'none'
      ? `Общие расходы без SKU не распределялись: ${formatMoney(state.analysis.totals.sharedNet)}`
      : `Вклад после распределения общих расходов: ${formatMoney(state.analysis.totals.finalContribution)}`,
    '',
    'Топ товаров по итоговому вкладу:'
  ];
  top.forEach((item, index) => {
    lines.push(`${index + 1}. ${item.article || item.sku} — ${formatMoney(item.finalContribution)} при выручке ${formatMoney(item.netSales)}`);
  });
  navigator.clipboard.writeText(lines.join('\n'))
    .then(() => alert('Сводка скопирована.'))
    .catch(() => alert('Не удалось скопировать сводку.'));
}

function exportProductCsv() {
  if (!state.analysis) return;
  const headers = [
    'SKU','Артикул','Товар','Чистые ед.','Заказы','Отмены','Операционные возвраты','Выручка нетто','К начислению',
    'Базовая комиссия','Эквайринг','Логистика','FBO / хранение','Реклама по SKU','Прочее продвижение','Прочие расходы МП',
    'Ручные затраты','Распределено общих','Итоговый вклад','Маржа','ДРР'
  ];
  const rows = state.analysis.products.map((item) => [
    item.sku, item.article, item.name, item.netUnits, item.ordersCount, item.cancelledQty, item.operationalReturnsQty, item.netSales,
    item.payoutBeforeOps, item.baseCommissionNet, item.acquiring, item.logistics, item.storage, item.promotion, item.promoExtra,
    item.penalties + item.otherMp, item.manualCosts, item.sharedAllocated, item.finalContribution, item.finalMargin, item.drr
  ]);

  const csv = [headers, ...rows].map((row) => row.map(csvEscape).join(';')).join('\n');
  const blob = new Blob(['\uFEFF' + csv], { type: 'text/csv;charset=utf-8;' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  const month = MONTH_INPUT.value || 'report';
  link.download = `ozon-products-${month}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function resetAll() {
  TEMPLATE_FILE_INPUT.value = '';
  PRODUCT_SEARCH.value = '';
  state.sheetData = null;
  state.analysis = null;
  state.sourceMode = null;
  state.csvFiles = {};
  ANALYZE_BTN.disabled = true;
  RESULTS_SECTION.classList.add('hidden');
  renderStatusCards({});
  renderCsvInputs();
}

function sortRows(a, b, sort) {
  const av = a[sort.key];
  const bv = b[sort.key];
  if (typeof av === 'string' || typeof bv === 'string') {
    return sort.desc ? String(bv || '').localeCompare(String(av || ''), 'ru') : String(av || '').localeCompare(String(bv || ''), 'ru');
  }
  return sort.desc ? (bv || 0) - (av || 0) : (av || 0) - (bv || 0);
}

function parseDate(value) {
  if (!value && value !== 0) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === 'number') {
    const epoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(epoch.getTime() + value * 86400000);
    return Number.isNaN(date.getTime()) ? null : date;
  }
  const trimmed = String(value).trim();
  if (!trimmed) return null;
  const iso = new Date(trimmed);
  if (!Number.isNaN(iso.getTime())) return iso;

  const match = trimmed.match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
  if (match) {
    const [, dd, mm, yyyy, hh = '0', min = '0', sec = '0'] = match;
    return new Date(Number(yyyy), Number(mm) - 1, Number(dd), Number(hh), Number(min), Number(sec));
  }
  return null;
}

function normalizeCellValue(value) {
  if (value === undefined || value === null) return '';
  if (typeof value === 'string') return value.trim();
  return value;
}

function trimTrailingBlanks(row) {
  const copy = [...row];
  while (copy.length && isBlank(copy[copy.length - 1])) copy.pop();
  return copy;
}

function isBlank(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function normalizeHeader(text) {
  return String(text || '').replace(/\s+/g, ' ').trim().toLowerCase();
}

function normalizeSku(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}

function toNumber(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return Number.isFinite(value) ? value : 0;
  const normalized = String(value)
    .replace(/\u00A0/g, '')
    .replace(/\s+/g, '')
    .replace(',', '.')
    .replace(/(?!^-)[^0-9.\-]/g, '');
  const num = Number(normalized);
  return Number.isFinite(num) ? num : 0;
}

function chooseValue(specific, fallback) {
  return specific === undefined || specific === null || specific === '' ? (fallback || 0) : toNumber(specific);
}

function monthToInputValue(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  return `${year}-${month}`;
}

function formatMonthLabel(monthValue) {
  if (!monthValue) return 'Отчёт';
  const [year, month] = monthValue.split('-').map(Number);
  const date = new Date(year, month - 1, 1);
  return new Intl.DateTimeFormat('ru-RU', { month: 'long', year: 'numeric' }).format(date);
}

function formatDateKey(date) {
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function formatMoney(value) {
  return new Intl.NumberFormat('ru-RU', { style: 'currency', currency: 'RUB', maximumFractionDigits: 0 }).format(value || 0);
}

function formatCompactMoney(value) {
  const abs = Math.abs(value || 0);
  if (abs >= 1000000) return `${(value / 1000000).toFixed(1)}м`;
  if (abs >= 1000) return `${(value / 1000).toFixed(0)}к`;
  return `${Math.round(value || 0)}`;
}

function formatInteger(value) {
  return new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 }).format(value || 0);
}

function formatPercent(value) {
  return `${(Number(value || 0) * 100).toFixed(1)}%`;
}

function csvEscape(value) {
  const str = String(value ?? '');
  return /[;"\n]/.test(str) ? `"${str.replace(/"/g, '""')}"` : str;
}

function safeInputValue(value) {
  return value === undefined || value === null ? '' : String(value);
}

function mostFrequent(values) {
  const counter = {};
  values.forEach((value) => { counter[value] = (counter[value] || 0) + 1; });
  return Object.keys(counter).sort((a, b) => counter[b] - counter[a])[0] || '';
}

function guessMonthFromTotals(dataset) {
  const month = checkMonthConsistency(dataset['Начисления'] || [], TEMPLATE_SPECS['Начисления'].dateFields, '').primaryMonth;
  return month;
}

function shorten(value, maxLength) {
  const str = String(value || '');
  return str.length > maxLength ? `${str.slice(0, maxLength - 1)}…` : str;
}

function escapeHtml(value) {
  return String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#039;');
}
