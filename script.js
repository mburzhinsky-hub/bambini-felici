const STORAGE_KEY = 'mp-analytics-v4-ozon-fix';

const state = {
  marketplace: 'ozon',
  month: '',
  files: {
    ozon: {
      accruals: [],
      orders: [],
      promotion: [],
      returns: [],
      realization: []
    },
    wb: {
      weeklyDetail: [],
      orderFeed: []
    }
  },
  manualCosts: {
    defaults: {
      cogs: '',
      packaging: '',
      inbound: '',
      other: ''
    },
    items: {}
  },
  lastAnalysis: null,
  filterText: '',
  resultsFilterText: ''
};

const nfMoney = new Intl.NumberFormat('ru-RU', {
  style: 'currency',
  currency: 'RUB',
  maximumFractionDigits: 0
});

const nfNumber = new Intl.NumberFormat('ru-RU', {
  maximumFractionDigits: 0
});

const nfDecimal = new Intl.NumberFormat('ru-RU', {
  minimumFractionDigits: 1,
  maximumFractionDigits: 1
});

const MARKETPLACE_CONFIG = {
  ozon: {
    label: 'Ozon',
    reports: {
      accruals: {
        key: 'accruals',
        title: 'Отчёт по начислениям',
        required: true,
        multiple: false,
        formatsLabel: 'XLSX, XLS, CSV, XML',
        accept: '.xlsx,.xls,.csv,.xml',
        where: 'Финансы / Аналитика Ozon → выгрузка начислений за месяц.',
        why: 'Даёт все расходы маркетплейса по операциям: логистику, эквайринг, платное продвижение, штрафы, хранение и прочие сервисы.',
        fields: ['ID начисления', 'Дата начисления', 'Группа услуг', 'Тип начисления', 'Артикул', 'SKU', 'Название товара', 'Количество', 'Сумма итого, руб.'],
        hints: ['id начисления', 'дата начисления', 'группа услуг', 'тип начисления', 'сумма итого'],
        parser: parseOzonAccruals
      },
      orders: {
        key: 'orders',
        title: 'Отчёт по заказам',
        required: true,
        multiple: false,
        formatsLabel: 'CSV, XML, XLSX, XLS',
        accept: '.csv,.xml,.xlsx,.xls',
        where: 'Аналитика → Отчёты → Заказы.',
        why: 'Нужен для числа заказов, текущих статусов, отмен и связки заказа с товаром и отправлением.',
        fields: ['Номер заказа', 'Номер отправления', 'Принят в обработку', 'Статус', 'Название товара', 'SKU', 'Артикул', 'Количество'],
        hints: ['номер заказа', 'номер отправления', 'принят в обработку', 'статус', 'название товара', 'sku'],
        parser: parseOzonOrders
      },
      promotion: {
        key: 'promotion',
        title: 'Аналитика продвижения',
        required: true,
        multiple: false,
        formatsLabel: 'XLSX, XLS',
        accept: '.xlsx,.xls',
        where: 'Продвижение → Аналитика продвижения → скачайте выгрузку за месяц.',
        why: 'Даёт рекламные расходы и метрики по SKU: расход, продажи, заказы, клики, показы. Нужен для полной картинки по товарам.',
        fields: ['SKU', 'Название товара', 'Расход, ₽', 'Продажи, ₽', 'Заказы, шт', 'Показы', 'Клики', 'В корзину'],
        hints: ['sku', 'расход', 'продажи', 'заказы', 'клики', 'показы'],
        parser: parseOzonPromotion
      },
      returns: {
        key: 'returns',
        title: 'Отчёт о возвратах',
        required: true,
        multiple: false,
        formatsLabel: 'XLSX, XLS',
        accept: '.xlsx,.xls',
        where: 'FBO / Возвраты → отчёт по возвратам.',
        why: 'Показывает операционные возвраты, причины возврата и статус возвратных отправлений. Это отдельный слой аналитики поверх финансовых возвратов.',
        fields: ['Наименование товара', 'Номер отправления', 'Артикул товара', 'OZON SKU ID', 'Дата возврата', 'Статус возврата', 'Причина возврата', 'Количество возвращаемых товаров'],
        hints: ['дата возврата', 'статус возврата', 'причина возврата', 'ozon sku id', 'количество возвращаемых товаров'],
        parser: parseOzonReturns
      },
      realization: {
        key: 'realization',
        title: 'Позаказный отчёт о реализации товара',
        required: true,
        multiple: false,
        formatsLabel: 'XLSX, XLS',
        accept: '.xlsx,.xls',
        where: 'Финансы → Документы → Отчёты о реализации → позаказный отчёт.',
        why: 'Главный источник по товарной выручке по цене продавца, базовой комиссии Ozon, сумме к начислению и финансовым возвратам по SKU.',
        fields: ['Название товара', 'Артикул', 'SKU', 'Реализовано на сумму', 'Баллы за скидки', 'Итого к начислению', 'Возвращено клиентом', 'Отправление'],
        hints: ['реализовано на сумму', 'итого к начислению', 'возвращено клиентом', 'отправление', 'sku'],
        parser: parseOzonRealization
      }
    }
  },
  wb: {
    label: 'Wildberries',
    reports: {
      weeklyDetail: {
        key: 'weeklyDetail',
        title: 'Детализация еженедельного отчёта реализации',
        required: true,
        multiple: true,
        formatsLabel: 'XLSX, XLS',
        accept: '.xlsx,.xls',
        where: 'Финансовые отчёты → Еженедельные → откройте нужный отчёт по номеру → скачайте детализацию в XLSX. Для месяца загрузите все недели, которые его пересекают.',
        why: 'Главный источник финансов WB: продажи, возвраты, комиссии, эквайринг, ПВЗ, логистика, хранение, удержания, штрафы, приёмка.',
        fields: ['тип документа', 'обоснование для оплаты', 'дата заказа', 'дата продажи', 'код номенклатуры / артикул поставщика', 'товар', 'к перечислению продавцу', 'реализовано', 'комиссия', 'логистика', 'удержания'],
        hints: ['обоснован', 'датапродажи', 'кодноменклатуры', 'артикулпоставщика', 'кперечислению', 'логист'],
        parser: parseWbWeeklyDetail
      },
      orderFeed: {
        key: 'orderFeed',
        title: 'Отчёт «Лента заказов»',
        required: false,
        multiple: false,
        formatsLabel: 'XLSX, XLS',
        accept: '.xlsx,.xls',
        where: 'Аналитика продавца → Лента заказов → выгрузка Excel.',
        why: 'Опциональный файл для точного количества заказов, отказов и возвратов. Важное ограничение: отчёт покрывает только последние 30 дней.',
        fields: ['дата оформления заказа', 'дата текущего статуса', 'статус', 'ID заказа', 'товар', 'цена со скидкой продавца'],
        hints: ['статус', 'idзаказа', 'датаоформлениязаказа', 'товар', 'цена'],
        parser: parseWbOrderFeed
      }
    }
  }
};

const REPORT_SEQUENCE = {
  ozon: ['accruals', 'orders', 'promotion', 'returns', 'realization'],
  wb: ['weeklyDetail', 'orderFeed']
};

const CHART_COLORS = ['#2554ff', '#08b2c2', '#6e7cff', '#a979ff', '#f39c44', '#ff6f8f', '#7cb342', '#607d8b'];

const dom = {};

document.addEventListener('DOMContentLoaded', init);

function init() {
  cacheDom();
  loadState();
  initMonth();
  bindEvents();
  renderAll();

  if (typeof window.XLSX === 'undefined') {
    showToast('Не удалось загрузить модуль для чтения Excel. Проверьте интернет-соединение или сохраните библиотеку локально рядом с проектом.', 'error', 7000);
  }
}

function cacheDom() {
  dom.analysisMonth = document.getElementById('analysisMonth');
  dom.marketSwitch = document.getElementById('marketSwitch');
  dom.instructionCards = document.getElementById('instructionCards');
  dom.validationList = document.getElementById('validationList');
  dom.buildStatus = document.getElementById('buildStatus');
  dom.uploadCards = document.getElementById('uploadCards');
  dom.manualDefaults = document.getElementById('manualDefaults');
  dom.manualCostSearch = document.getElementById('manualCostSearch');
  dom.manualHint = document.getElementById('manualHint');
  dom.manualCostTable = document.getElementById('manualCostTable');
  dom.buildAnalysisBtn = document.getElementById('buildAnalysisBtn');
  dom.resetBtn = document.getElementById('resetBtn');
  dom.resultsSection = document.getElementById('resultsSection');
  dom.resultsSubtitle = document.getElementById('resultsSubtitle');
  dom.availabilityGrid = document.getElementById('availabilityGrid');
  dom.kpiGrid = document.getElementById('kpiGrid');
  dom.lineChart = document.getElementById('lineChart');
  dom.donutChart = document.getElementById('donutChart');
  dom.barChart = document.getElementById('barChart');
  dom.insightsList = document.getElementById('insightsList');
  dom.expenseBreakdown = document.getElementById('expenseBreakdown');
  dom.weakProducts = document.getElementById('weakProducts');
  dom.tableNote = document.getElementById('tableNote');
  dom.productTable = document.getElementById('productTable');
  dom.copySummaryBtn = document.getElementById('copySummaryBtn');
  dom.exportCsvBtn = document.getElementById('exportCsvBtn');
  dom.limitationsBlock = document.getElementById('limitationsBlock');
  dom.toastContainer = document.getElementById('toastContainer');
  dom.resultsProductSearch = document.getElementById('resultsProductSearch');
}

function bindEvents() {
  dom.analysisMonth.addEventListener('change', () => {
    state.month = dom.analysisMonth.value;
    saveState();
    state.lastAnalysis = null;
    renderAll();
  });

  dom.marketSwitch.addEventListener('click', (event) => {
    const button = event.target.closest('.market-btn');
    if (!button) return;
    state.marketplace = button.dataset.marketplace;
    state.lastAnalysis = null;
    saveState();
    renderAll();
  });

  dom.uploadCards.addEventListener('change', async (event) => {
    const input = event.target.closest('input[type="file"]');
    if (!input) return;
    const reportKey = input.dataset.reportKey;
    const files = Array.from(input.files || []);
    if (!files.length) return;
    await handleFiles(reportKey, files);
    input.value = '';
  });

  dom.uploadCards.addEventListener('click', async (event) => {
    const clearBtn = event.target.closest('[data-clear-report]');
    if (clearBtn) {
      const reportKey = clearBtn.dataset.clearReport;
      state.files[state.marketplace][reportKey] = [];
      state.lastAnalysis = null;
      renderAll();
      return;
    }

    const removeBtn = event.target.closest('[data-remove-file]');
    if (removeBtn) {
      const reportKey = removeBtn.dataset.reportKey;
      const fileId = removeBtn.dataset.removeFile;
      state.files[state.marketplace][reportKey] = state.files[state.marketplace][reportKey].filter((item) => item.id !== fileId);
      state.lastAnalysis = null;
      renderAll();
    }
  });

  dom.uploadCards.addEventListener('dragover', (event) => {
    const dropzone = event.target.closest('.dropzone');
    if (!dropzone) return;
    event.preventDefault();
    dropzone.classList.add('dragover');
  });

  dom.uploadCards.addEventListener('dragleave', (event) => {
    const dropzone = event.target.closest('.dropzone');
    if (!dropzone) return;
    dropzone.classList.remove('dragover');
  });

  dom.uploadCards.addEventListener('drop', async (event) => {
    const dropzone = event.target.closest('.dropzone');
    if (!dropzone) return;
    event.preventDefault();
    dropzone.classList.remove('dragover');
    const reportKey = dropzone.dataset.reportKey;
    const files = Array.from(event.dataTransfer?.files || []);
    if (!files.length) return;
    await handleFiles(reportKey, files);
  });

  dom.manualDefaults.addEventListener('input', handleManualDefaultChange);
  dom.manualCostTable.addEventListener('input', handleManualItemChange);
  dom.manualCostSearch.addEventListener('input', () => {
    state.filterText = dom.manualCostSearch.value.trim().toLowerCase();
    renderManualCostTable();
  });

  dom.resultsProductSearch.addEventListener('input', () => {
    state.resultsFilterText = dom.resultsProductSearch.value.trim().toLowerCase();
    renderResultsTable();
  });

  dom.buildAnalysisBtn.addEventListener('click', () => {
    const validation = getValidationState();
    if (!validation.canBuild) {
      showToast('Сначала загрузите все обязательные отчёты по текущей площадке.', 'warn');
      return;
    }
    buildAnalysis();
  });

  dom.resetBtn.addEventListener('click', () => {
    const confirmed = window.confirm('Сбросить все загруженные файлы и результаты для обеих площадок?');
    if (!confirmed) return;
    resetAllData();
  });

  dom.copySummaryBtn.addEventListener('click', async () => {
    if (!state.lastAnalysis) return;
    const text = createSummaryText(state.lastAnalysis);
    try {
      await navigator.clipboard.writeText(text);
      showToast('Сводка скопирована в буфер обмена.', 'success');
    } catch (error) {
      showToast('Не удалось скопировать сводку. Браузер заблокировал доступ к буферу обмена.', 'error');
    }
  });

  dom.exportCsvBtn.addEventListener('click', () => {
    if (!state.lastAnalysis) return;
    exportProductsCsv(state.lastAnalysis);
  });
}

function loadState() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return;
    const parsed = JSON.parse(raw);
    if (parsed.marketplace) state.marketplace = parsed.marketplace;
    if (parsed.month) state.month = parsed.month;
    if (parsed.manualCosts) state.manualCosts = parsed.manualCosts;
  } catch (error) {
    console.warn('Не удалось восстановить сохранённое состояние.', error);
  }
}

function saveState() {
  const payload = {
    marketplace: state.marketplace,
    month: state.month,
    manualCosts: state.manualCosts
  };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(payload));
}

function initMonth() {
  if (!state.month) {
    const today = new Date();
    const month = String(today.getMonth() + 1).padStart(2, '0');
    state.month = `${today.getFullYear()}-${month}`;
  }
  dom.analysisMonth.value = state.month;
}

function resetAllData() {
  state.files = {
    ozon: { accruals: [], orders: [], promotion: [], returns: [], realization: [] },
    wb: { weeklyDetail: [], orderFeed: [] }
  };
  state.lastAnalysis = null;
  state.filterText = '';
  state.resultsFilterText = '';
  dom.manualCostSearch.value = '';
  dom.resultsProductSearch.value = '';
  renderAll();
  showToast('Файлы и результаты сброшены. Ручные затраты сохранены.', 'success');
}

function renderAll() {
  renderMarketSwitch();
  renderInstructionCards();
  renderUploadCards();
  renderValidation();
  renderManualDefaults();
  renderManualCostTable();
  renderLimitations();
  renderResults();
  updateBuildButtonState();
}

function renderMarketSwitch() {
  dom.marketSwitch.querySelectorAll('.market-btn').forEach((button) => {
    button.classList.toggle('active', button.dataset.marketplace === state.marketplace);
  });
}

function renderInstructionCards() {
  const config = MARKETPLACE_CONFIG[state.marketplace];
  const html = REPORT_SEQUENCE[state.marketplace]
    .map((reportKey) => {
      const report = config.reports[reportKey];
      return `
        <article class="instruction-card">
          <div class="badge-row">
            <span class="badge ${report.required ? 'required' : 'optional'}">${report.required ? 'Обязательный' : 'Опциональный'}</span>
            <span class="badge">${escapeHtml(report.formatsLabel)}</span>
            ${report.multiple ? '<span class="badge">Можно несколько файлов</span>' : '<span class="badge">Один файл</span>'}
          </div>
          <h3 class="instruction-title">${escapeHtml(report.title)}</h3>
          <p class="section-note">${escapeHtml(report.why)}</p>
          <ul class="info-list">
            <li>
              <span class="info-label">Где искать</span>
              <span>${escapeHtml(report.where)}</span>
            </li>
            <li>
              <span class="info-label">Формат</span>
              <span>${escapeHtml(report.formatsLabel)}</span>
            </li>
            <li>
              <span class="info-label">Используемые поля</span>
              <span>${escapeHtml(report.fields.join(' • '))}</span>
            </li>
          </ul>
        </article>
      `;
    })
    .join('');

  dom.instructionCards.innerHTML = html;
}

function renderUploadCards() {
  const config = MARKETPLACE_CONFIG[state.marketplace];
  const html = REPORT_SEQUENCE[state.marketplace]
    .map((reportKey) => {
      const report = config.reports[reportKey];
      const files = state.files[state.marketplace][reportKey] || [];
      const fileListHtml = files.length
        ? `<div class="file-list">${files.map((item) => renderUploadedFileItem(reportKey, item)).join('')}</div>`
        : `<div class="small-note">Файлы пока не загружены.</div>`;

      return `
        <article class="upload-card">
          <div>
            <div class="badge-row">
              <span class="badge ${report.required ? 'required' : 'optional'}">${report.required ? 'Обязательный' : 'Опциональный'}</span>
              <span class="badge">${report.multiple ? 'Можно добавлять частями' : 'Перезаписывает прошлый файл'}</span>
            </div>
            <h3 class="upload-title">${escapeHtml(report.title)}</h3>
            <div class="small-note">${escapeHtml(report.why)}</div>
          </div>

          <label class="dropzone" data-report-key="${reportKey}">
            <input type="file" data-report-key="${reportKey}" accept="${report.accept}" ${report.multiple ? 'multiple' : ''} />
            <div>
              <strong>Перетащите файл сюда или нажмите для выбора</strong>
              <p>
                ${escapeHtml(report.formatsLabel)}
                ${report.multiple ? ' • можно загрузить несколько файлов подряд' : ''}
              </p>
            </div>
          </label>

          ${fileListHtml}

          <div class="file-actions">
            <button class="btn btn-secondary" data-clear-report="${reportKey}" type="button">Очистить блок</button>
          </div>
        </article>
      `;
    })
    .join('');

  dom.uploadCards.innerHTML = html;
}

function renderUploadedFileItem(reportKey, item) {
  const statusClass = item.status === 'error' ? 'error' : item.status === 'warn' ? 'partial' : 'exact';
  const statusLabel = item.status === 'error' ? 'Ошибка' : item.status === 'warn' ? 'Частично' : 'Готово';
  const sub = item.error
    ? item.error
    : `${nfNumber.format(item.rowCount || 0)} строк • за выбранный месяц: ${nfNumber.format(item.monthRows || 0)}`;

  return `
    <div class="file-item">
      <div class="file-item-meta">
        <div class="file-name">${escapeHtml(item.name)}</div>
        <div class="file-sub">${escapeHtml(sub)}</div>
        ${item.warnings?.length ? `<div class="small-note">${escapeHtml(item.warnings.join(' • '))}</div>` : ''}
      </div>
      <div class="file-actions">
        <span class="badge ${statusClass}">${statusLabel}</span>
        <button class="icon-btn" type="button" title="Удалить файл" data-remove-file="${item.id}" data-report-key="${reportKey}">×</button>
      </div>
    </div>
  `;
}

function renderValidation() {
  const validation = getValidationState();
  const config = MARKETPLACE_CONFIG[state.marketplace];

  const html = REPORT_SEQUENCE[state.marketplace]
    .map((reportKey) => {
      const report = config.reports[reportKey];
      const item = validation.items.find((entry) => entry.reportKey === reportKey);
      const badgeClass = item.status === 'ready' ? 'exact' : item.status === 'optional' ? 'optional' : item.status === 'error' ? 'error' : 'partial';
      const badgeLabel = item.status === 'ready' ? 'Готово' : item.status === 'optional' ? 'Опционально' : item.status === 'error' ? 'Ошибка' : 'Не загружено';
      return `
        <div class="validation-item">
          <div class="validation-copy">
            <strong>${escapeHtml(report.title)}</strong>
            <span>${escapeHtml(item.message)}</span>
          </div>
          <span class="badge ${badgeClass}">${badgeLabel}</span>
        </div>
      `;
    })
    .join('');

  dom.validationList.innerHTML = html;
  dom.buildStatus.classList.remove('hidden', 'ok', 'warn', 'error');

  if (validation.canBuild) {
    dom.buildStatus.classList.add('ok');
    dom.buildStatus.textContent = validation.statusMessage;
  } else if (validation.hasBlockingErrors) {
    dom.buildStatus.classList.add('error');
    dom.buildStatus.textContent = validation.statusMessage;
  } else {
    dom.buildStatus.classList.add('warn');
    dom.buildStatus.textContent = validation.statusMessage;
  }
}

function getValidationState() {
  const config = MARKETPLACE_CONFIG[state.marketplace];
  const items = [];
  let canBuild = true;
  let hasBlockingErrors = false;

  for (const reportKey of REPORT_SEQUENCE[state.marketplace]) {
    const report = config.reports[reportKey];
    const files = state.files[state.marketplace][reportKey] || [];
    const failed = files.some((item) => item.status === 'error');
    const okFiles = files.filter((item) => item.status !== 'error');

    if (failed) {
      items.push({
        reportKey,
        status: 'error',
        message: 'Есть файл с ошибкой структуры или формата. Очистите блок и загрузите корректный файл.'
      });
      if (report.required) {
        canBuild = false;
        hasBlockingErrors = true;
      }
      continue;
    }

    if (!okFiles.length) {
      items.push({
        reportKey,
        status: report.required ? 'missing' : 'optional',
        message: report.required ? 'Файл ещё не загружен.' : 'Можно не загружать — анализ всё равно построится, но часть метрик будет недоступна.'
      });
      if (report.required) canBuild = false;
      continue;
    }

    const monthRows = okFiles.reduce((sum, item) => sum + (item.monthRows || 0), 0);
    const reportMessage = report.multiple
      ? `Загружено ${okFiles.length} файлов. Строк за выбранный месяц: ${nfNumber.format(monthRows)}.`
      : `Файл загружен. Строк за выбранный месяц: ${nfNumber.format(monthRows)}.`;

    items.push({
      reportKey,
      status: 'ready',
      message: reportMessage
    });
  }

  let statusMessage = '';
  if (canBuild) {
    statusMessage = 'Все обязательные отчёты загружены. Анализ можно строить.';
  } else if (hasBlockingErrors) {
    statusMessage = 'Есть ошибки формата. Исправьте проблемные блоки загрузки.';
  } else {
    statusMessage = 'Не хватает обязательных файлов для текущей площадки.';
  }

  return { canBuild, hasBlockingErrors, statusMessage, items };
}

function updateBuildButtonState() {
  const validation = getValidationState();
  dom.buildAnalysisBtn.disabled = !validation.canBuild;
}

function renderManualDefaults() {
  const defaults = state.manualCosts.defaults;
  dom.manualDefaults.innerHTML = `
    <div class="default-card">
      <label>
        Себестоимость по умолчанию, ₽
        <input type="number" step="0.01" min="0" data-default-cost="cogs" value="${escapeHtml(defaults.cogs ?? '')}" />
      </label>
    </div>
    <div class="default-card">
      <label>
        Упаковка по умолчанию, ₽
        <input type="number" step="0.01" min="0" data-default-cost="packaging" value="${escapeHtml(defaults.packaging ?? '')}" />
      </label>
    </div>
    <div class="default-card">
      <label>
        Логистика до склада, ₽
        <input type="number" step="0.01" min="0" data-default-cost="inbound" value="${escapeHtml(defaults.inbound ?? '')}" />
      </label>
    </div>
    <div class="default-card">
      <label>
        Прочие затраты на единицу, ₽
        <input type="number" step="0.01" min="0" data-default-cost="other" value="${escapeHtml(defaults.other ?? '')}" />
      </label>
    </div>
  `;
}

function renderManualCostTable() {
  const products = collectDetectedProducts(state.marketplace);
  const filtered = products.filter((item) => {
    if (!state.filterText) return true;
    const haystack = `${item.sku} ${item.article || ''} ${item.name}`.toLowerCase();
    return haystack.includes(state.filterText);
  });

  const thead = dom.manualCostTable.querySelector('thead');
  const tbody = dom.manualCostTable.querySelector('tbody');

  if (!products.length) {
    dom.manualHint.textContent = 'Загрузите и распарсьте файлы — здесь появится список товаров для ручного ввода.';
    thead.innerHTML = '';
    tbody.innerHTML = '<tr><td colspan="8" class="small-note">Пока нет данных по товарам.</td></tr>';
    return;
  }

  dom.manualHint.textContent = 'Изменения сохраняются в localStorage. Для расчёта на уровне SKU используйте эту таблицу, для массового сценария — значения по умолчанию выше.';

  thead.innerHTML = `
    <tr>
      <th>SKU / артикул</th>
      <th>Товар</th>
      <th>Выкупы, шт.</th>
      <th>Себестоимость</th>
      <th>Упаковка</th>
      <th>До склада</th>
      <th>Прочее</th>
      <th>Итого на единицу</th>
    </tr>
  `;

  tbody.innerHTML = filtered
    .map((item) => {
      const key = getManualCostKey(state.marketplace, item.sku);
      const saved = state.manualCosts.items[key] || {};
      const effective = getManualUnitCost(state.marketplace, item.sku);
      return `
        <tr>
          <td class="mono">${escapeHtml(item.sku || '—')}${item.article ? `<div class="small-note">${escapeHtml(item.article)}</div>` : ''}</td>
          <td>${escapeHtml(item.name || 'Без названия')}</td>
          <td class="mono">${nfNumber.format(item.unitsSold || 0)}</td>
          <td><input class="table-input" type="number" step="0.01" min="0" data-item-cost="cogs" data-sku="${escapeHtml(item.sku)}" value="${escapeHtml(saved.cogs ?? '')}" /></td>
          <td><input class="table-input" type="number" step="0.01" min="0" data-item-cost="packaging" data-sku="${escapeHtml(item.sku)}" value="${escapeHtml(saved.packaging ?? '')}" /></td>
          <td><input class="table-input" type="number" step="0.01" min="0" data-item-cost="inbound" data-sku="${escapeHtml(item.sku)}" value="${escapeHtml(saved.inbound ?? '')}" /></td>
          <td><input class="table-input" type="number" step="0.01" min="0" data-item-cost="other" data-sku="${escapeHtml(item.sku)}" value="${escapeHtml(saved.other ?? '')}" /></td>
          <td class="mono">${fmtMoney(effective.total)}</td>
        </tr>
      `;
    })
    .join('');
}


function collectDetectedProducts(marketplace) {
  if (state.lastAnalysis && state.lastAnalysis.marketplace === marketplace && state.lastAnalysis.month === state.month) {
    return state.lastAnalysis.productSummary.map((item) => ({
      sku: item.sku,
      article: item.article,
      name: item.name,
      unitsSold: item.unitsSold || 0,
      revenue: item.revenue || 0
    }));
  }

  const map = new Map();
  const month = state.month;

  function ensureCandidate(row) {
    const sku = normalizeSku(row?.sku);
    const article = safeString(row?.article);
    const name = safeString(row?.name) || 'Без названия';
    const key = sku || article || name;
    if (!map.has(key)) {
      map.set(key, { sku: sku || article || name, article, name, unitsSold: 0, revenue: 0 });
    }
    const item = map.get(key);
    if (!item.article && article) item.article = article;
    if ((!item.name || item.name === 'Без названия') && name) item.name = name;
    return item;
  }

  if (marketplace === 'ozon') {
    for (const file of state.files.ozon.realization) {
      if (file.status === 'error') continue;
      for (const row of file.parsed.rows || []) {
        if (!isDateInSelectedMonth(row.eventDate || row.reportDate, month)) continue;
        const item = ensureCandidate(row);
        item.unitsSold += row.saleQty || 0;
        item.revenue += row.revenue || 0;
      }
    }

    for (const file of state.files.ozon.promotion) {
      if (file.status === 'error') continue;
      for (const row of file.parsed.rows || []) {
        if (!isDateInSelectedMonth(row.eventDate || row.reportDate, month)) continue;
        ensureCandidate(row);
      }
      for (const row of file.parsed.unionRows || []) {
        if (!isDateInSelectedMonth(row.eventDate, month)) continue;
        ensureCandidate({ sku: row.promotedSku, name: row.promotedName });
      }
    }

    for (const key of ['returns', 'orders', 'accruals']) {
      for (const file of state.files.ozon[key]) {
        if (file.status === 'error') continue;
        for (const row of file.parsed.rows || []) {
          if (!isDateInSelectedMonth(row.eventDate || row.orderDate || row.returnDate, month)) continue;
          ensureCandidate(row);
        }
      }
    }
  }

  if (marketplace === 'wb') {
    for (const file of state.files.wb.weeklyDetail) {
      if (file.status === 'error') continue;
      for (const row of file.parsed.rows || []) {
        if (!isDateInSelectedMonth(row.eventDate, month)) continue;
        const item = ensureCandidate(row);
        if (row.isSale) item.unitsSold += row.quantity || 0;
        item.revenue += row.revenueSigned || 0;
      }
    }
  }

  return Array.from(map.values()).sort((a, b) => b.revenue - a.revenue || b.unitsSold - a.unitsSold || a.name.localeCompare(b.name, 'ru'));
}


function handleManualDefaultChange(event) {
  const input = event.target.closest('[data-default-cost]');
  if (!input) return;
  state.manualCosts.defaults[input.dataset.defaultCost] = input.value;
  saveState();
  renderManualCostTable();
  rebuildIfNeeded();
}

function handleManualItemChange(event) {
  const input = event.target.closest('[data-item-cost]');
  if (!input) return;
  const sku = input.dataset.sku;
  const field = input.dataset.itemCost;
  const key = getManualCostKey(state.marketplace, sku);
  state.manualCosts.items[key] = state.manualCosts.items[key] || {};
  state.manualCosts.items[key][field] = input.value;
  saveState();
  renderManualCostTable();
  rebuildIfNeeded();
}

function rebuildIfNeeded() {
  if (!state.lastAnalysis) return;
  const currentMarketplace = state.lastAnalysis.marketplace;
  if (currentMarketplace !== state.marketplace) return;
  buildAnalysis(true);
}

async function handleFiles(reportKey, files) {
  const reportConfig = MARKETPLACE_CONFIG[state.marketplace].reports[reportKey];
  if (!reportConfig) return;

  if (!reportConfig.multiple) {
    state.files[state.marketplace][reportKey] = [];
  }

  for (const file of files) {
    const fileId = getFileId(file);
    const alreadyExists = state.files[state.marketplace][reportKey].some((item) => item.id === fileId);
    if (alreadyExists) continue;

    try {
      const parsed = await parseReportFile(file, reportConfig);
      state.files[state.marketplace][reportKey].push(parsed);
      showToast(`Файл «${file.name}» обработан.`, parsed.status === 'warn' ? 'warn' : 'success');
    } catch (error) {
      console.error(error);
      state.files[state.marketplace][reportKey].push({
        id: fileId,
        name: file.name,
        status: 'error',
        error: error.message || 'Не удалось прочитать файл.'
      });
      showToast(`Ошибка чтения файла «${file.name}».`, 'error');
    }
  }

  state.lastAnalysis = null;
  renderAll();
}

async function parseReportFile(file, reportConfig) {
  const rawTable = await parseSpreadsheetLikeFile(file, reportConfig.hints);
  const parsed = await reportConfig.parser(rawTable, file, {
    selectedMonth: state.month,
    marketplace: state.marketplace,
    reportKey: reportConfig.key
  });
  const monthRows = (parsed.rows || []).filter((row) => isDateInSelectedMonth(row.eventDate || row.activityDate || row.orderDate || row.returnDate || row.reportDate, state.month)).length;

  return {
    id: getFileId(file),
    name: file.name,
    status: parsed.warnings?.length ? 'warn' : 'ok',
    warnings: parsed.warnings || [],
    rowCount: parsed.rows?.length || 0,
    monthRows,
    parsed
  };
}

async function parseSpreadsheetLikeFile(file, hints = []) {
  const extension = getFileExtension(file.name);

  if (extension === 'csv' || extension === 'txt') {
    const text = await file.text();
    return parseCsvText(text, hints, file.name);
  }

  if (extension === 'xml') {
    const text = await file.text();
    return parseXmlOrWorkbook(text, hints, file.name);
  }

  if (typeof window.XLSX === 'undefined') {
    throw new Error('Библиотека XLSX не загрузилась.');
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array', cellDates: false, raw: false });
  return pickBestSheetFromWorkbook(workbook, hints, file.name);
}

function parseCsvText(text, hints, fileName) {
  const rows = simpleCsvParse(text);
  return buildSheetCandidate(rows, hints, fileName, 'CSV');
}

function parseXmlOrWorkbook(text, hints, fileName) {
  if (typeof window.XLSX !== 'undefined') {
    try {
      const workbook = XLSX.read(text, { type: 'string', dense: true, raw: false });
      return pickBestSheetFromWorkbook(workbook, hints, fileName);
    } catch (error) {
      // fallback below
    }
  }

  const rows = parseGenericXmlRows(text);
  return buildObjectsCandidate(rows, hints, fileName, 'XML');
}

function pickBestSheetFromWorkbook(workbook, hints, fileName) {
  const candidates = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    const aoa = sheetToAoaRobust(sheet);
    return buildSheetCandidate(aoa, hints, fileName, sheetName);
  }).filter(Boolean);

  if (!candidates.length) {
    throw new Error('В файле не найден лист с табличными данными.');
  }

  candidates.sort((a, b) => b.score - a.score || b.rows.length - a.rows.length);
  return candidates[0];
}

function sheetToAoaRobust(sheet) {
  if (!sheet || typeof XLSX === 'undefined') return [];

  if (Array.isArray(sheet)) {
    const denseRows = [];
    for (let r = 0; r < sheet.length; r += 1) {
      const sourceRow = sheet[r];
      if (!sourceRow) continue;
      const resultRow = [];
      for (let c = 0; c < sourceRow.length; c += 1) {
        const cell = sourceRow[c];
        if (cell == null) {
          resultRow[c] = '';
        } else if (typeof cell === 'object' && cell) {
          resultRow[c] = cell.w ?? cell.v ?? '';
        } else {
          resultRow[c] = cell;
        }
      }
      denseRows[r] = resultRow;
    }
    if (denseRows.length > 1) return denseRows;
  }

  const cellKeys = Object.keys(sheet).filter((key) => /^[A-Z]+\d+$/i.test(key));
  if (!cellKeys.length) {
    return XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', raw: false, blankrows: false });
  }

  let minRow = Infinity;
  let minCol = Infinity;
  let maxRow = 0;
  let maxCol = 0;

  cellKeys.forEach((key) => {
    const address = XLSX.utils.decode_cell(key);
    minRow = Math.min(minRow, address.r);
    minCol = Math.min(minCol, address.c);
    maxRow = Math.max(maxRow, address.r);
    maxCol = Math.max(maxCol, address.c);
  });

  const range = {
    s: { r: Number.isFinite(minRow) ? minRow : 0, c: Number.isFinite(minCol) ? minCol : 0 },
    e: { r: maxRow, c: maxCol }
  };

  return XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: '',
    raw: false,
    blankrows: false,
    range
  });
}

function buildSheetCandidate(aoa, hints, fileName, sourceName) {
  if (!Array.isArray(aoa) || !aoa.length) return null;
  const headerIndex = detectHeaderRow(aoa, hints);
  const headerRow = aoa[headerIndex] || [];
  const headers = makeUniqueHeaders(headerRow);
  const rows = [];

  for (let i = headerIndex + 1; i < aoa.length; i += 1) {
    const row = aoa[i] || [];
    if (isRowEmpty(row)) continue;
    const obj = {};
    let nonEmpty = 0;
    headers.forEach((header, index) => {
      const value = row[index] ?? '';
      if (!isBlank(value)) nonEmpty += 1;
      obj[header] = value;
    });
    if (nonEmpty < 2) continue;
    rows.push(obj);
  }

  const normalizedHeaders = headers.map(normalizeHeader);
  const hintScore = hints.reduce((score, hint) => score + Number(normalizedHeaders.some((header) => header.includes(normalizeHeader(hint)))), 0);
  return {
    fileName,
    sourceName,
    headers,
    normalizedHeaders,
    rows,
    score: hintScore * 20 + rows.length * 0.1 + headers.length
  };
}

function buildObjectsCandidate(objects, hints, fileName, sourceName) {
  if (!Array.isArray(objects) || !objects.length) {
    throw new Error('XML не содержит табличных записей.');
  }
  const headers = makeUniqueHeaders(Object.keys(objects[0]));
  const normalizedHeaders = headers.map(normalizeHeader);
  const rows = objects.map((row) => {
    const obj = {};
    headers.forEach((header) => {
      obj[header] = row[header] ?? '';
    });
    return obj;
  });

  const hintScore = hints.reduce((score, hint) => score + Number(normalizedHeaders.some((header) => header.includes(normalizeHeader(hint)))), 0);
  return {
    fileName,
    sourceName,
    headers,
    normalizedHeaders,
    rows,
    score: hintScore * 20 + rows.length * 0.1 + headers.length
  };
}

function detectHeaderRow(aoa, hints) {
  const maxRows = Math.min(25, aoa.length);
  let bestIndex = 0;
  let bestScore = -Infinity;

  for (let i = 0; i < maxRows; i += 1) {
    const row = aoa[i] || [];
    const values = row.map((value) => String(value ?? '').trim()).filter(Boolean);
    if (values.length < 2) continue;

    const joined = normalizeHeader(values.join(' '));
    const textCells = values.filter((value) => /[A-Za-zА-Яа-я]/.test(value)).length;
    const numericCells = values.filter((value) => /^[-+]?\d+[\d\s,.-]*$/.test(value)).length;
    const hintHits = hints.filter((hint) => joined.includes(normalizeHeader(hint))).length;
    const score = values.length * 4 + textCells * 3 + hintHits * 12 - numericCells * 1.5;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex;
}


function parseOzonAccruals(rawTable) {
  const resolver = createResolver(rawTable.headers);
  const columns = {
    accrualId: resolver.find(['id начисления', 'ид начисления']),
    eventDate: resolver.find(['дата начисления', 'дата операции', 'дата']),
    group: resolver.find(['группа услуг']),
    type: resolver.find(['тип начисления', 'тип операции']),
    article: resolver.find(['артикул']),
    sku: resolver.find(['sku']),
    name: resolver.find(['название товара', 'наименование товара']),
    quantity: resolver.find(['количество']),
    sellerPrice: resolver.find(['цена продавца']),
    orderDate: resolver.find(['дата принятия заказа в обработку или оказания услуги', 'дата принятия заказа']),
    platform: resolver.find(['платформа продажи']),
    scheme: resolver.find(['схема работы']),
    amount: resolver.find(['сумма итого руб', 'сумма итого, руб.', 'сумма итого'])
  };

  if (!columns.eventDate || !columns.group || !columns.type || !columns.amount) {
    throw new Error('Не удалось распознать Ozon «Отчёт по начислениям». Нужны столбцы с датой начисления, группой услуг, типом начисления и суммой.');
  }

  const rows = rawTable.rows
    .map((row) => {
      const groupText = safeString(getCell(row, columns.group));
      const typeText = safeString(getCell(row, columns.type));
      const category = detectOzonAccrualCategory(groupText, typeText);

      return {
        eventDate: parseDateValue(getCell(row, columns.eventDate)),
        orderDate: parseDateValue(getCell(row, columns.orderDate)),
        shipmentId: safeString(getCell(row, columns.accrualId)),
        sku: normalizeSku(getCell(row, columns.sku)),
        article: safeString(getCell(row, columns.article)),
        name: safeString(getCell(row, columns.name)),
        quantity: Math.max(0, parseNumber(getCell(row, columns.quantity)) || 0),
        sellerPrice: parseNumber(getCell(row, columns.sellerPrice)) || 0,
        platform: safeString(getCell(row, columns.platform)),
        scheme: safeString(getCell(row, columns.scheme)),
        groupText,
        typeText,
        category,
        amount: parseNumber(getCell(row, columns.amount)) || 0,
        hasProduct: Boolean(normalizeSku(getCell(row, columns.sku)) || safeString(getCell(row, columns.article)) || safeString(getCell(row, columns.name)))
      };
    })
    .filter((row) => row.eventDate || row.shipmentId || row.sku || row.article || row.name || Math.abs(row.amount) > 0.01);

  return { rows, warnings: [] };
}

function parseOzonOrders(rawTable) {
  const resolver = createResolver(rawTable.headers);
  const columns = {
    orderId: resolver.find(['номер заказа']),
    shipmentId: resolver.find(['номер отправления']),
    orderDate: resolver.find(['принят в обработку', 'дата заказа', 'дата оформления заказа']),
    shipmentDate: resolver.find(['дата отгрузки']),
    status: resolver.find(['статус']),
    deliveryDate: resolver.find(['дата доставки']),
    deliveryToCarrierDate: resolver.find(['фактическая дата передачи в доставку']),
    amount: resolver.find(['сумма отправления']),
    name: resolver.find(['название товара']),
    sku: resolver.find(['sku']),
    article: resolver.find(['артикул']),
    sellerPrice: resolver.find(['ваша цена']),
    customerPaid: resolver.find(['оплачено покупателем']),
    quantity: resolver.find(['количество'])
  };

  if (!columns.orderId || !columns.shipmentId || !columns.orderDate || !columns.status) {
    throw new Error('Не удалось распознать Ozon «Отчёт по заказам». Нужны столбцы «Номер заказа», «Номер отправления», «Принят в обработку» и «Статус».');
  }

  const rows = rawTable.rows
    .map((row) => {
      const statusText = safeString(getCell(row, columns.status));
      const normalizedStatus = normalizeHeader(statusText);

      return {
        orderDate: parseDateValue(getCell(row, columns.orderDate)),
        shipmentDate: parseDateValue(getCell(row, columns.shipmentDate)),
        statusDate: parseDateValue(getCell(row, columns.deliveryDate)) || parseDateValue(getCell(row, columns.shipmentDate)) || parseDateValue(getCell(row, columns.orderDate)),
        activityDate: parseDateValue(getCell(row, columns.deliveryDate)) || parseDateValue(getCell(row, columns.shipmentDate)) || parseDateValue(getCell(row, columns.orderDate)),
        eventDate: parseDateValue(getCell(row, columns.orderDate)),
        orderId: safeString(getCell(row, columns.orderId)),
        shipmentId: safeString(getCell(row, columns.shipmentId)),
        sku: normalizeSku(getCell(row, columns.sku)),
        article: safeString(getCell(row, columns.article)),
        name: safeString(getCell(row, columns.name)),
        status: statusText,
        quantity: Math.max(0, parseNumber(getCell(row, columns.quantity)) || 0),
        sellerPrice: parseNumber(getCell(row, columns.sellerPrice)) || 0,
        amount: parseNumber(getCell(row, columns.amount)) || 0,
        customerPaid: parseNumber(getCell(row, columns.customerPaid)) || 0,
        isCancel: /отмен/.test(normalizedStatus),
        isReturn: /возврат/.test(normalizedStatus),
        isDelivered: /доставлен|выдан|получ/.test(normalizedStatus),
        isInProgress: /доставля|принят|собран|ожида/.test(normalizedStatus)
      };
    })
    .filter((row) => row.orderId || row.shipmentId);

  return { rows, warnings: [] };
}

async function parseOzonPromotion(rawTable, file, context = {}) {
  let statisticsTable = rawTable;
  let unionRows = [];
  let reportMonth = extractMonthKeyFromText(file?.name || '') || context.selectedMonth || '';

  if (typeof XLSX !== 'undefined' && file) {
    try {
      const workbook = await readWorkbookFromFile(file);
      const statSheetName = workbook.SheetNames.find((name) => normalizeHeader(name).includes('statistics')) || workbook.SheetNames[0];

      if (statSheetName) {
        const statAoa = sheetToAoaRobust(workbook.Sheets[statSheetName]);
        reportMonth = extractMonthKeyFromText(safeString(statAoa[0]?.[0])) || reportMonth;
        statisticsTable = buildSheetCandidate(statAoa, ['sku', 'расход', 'продажи', 'заказы', 'клики', 'показы'], file.name, statSheetName) || rawTable;
      }

      const unionSheetName = workbook.SheetNames.find((name) => normalizeHeader(name).includes('union'));
      if (unionSheetName) {
        const unionAoa = sheetToAoaRobust(workbook.Sheets[unionSheetName]);
        const unionTable = buildSheetCandidate(unionAoa, ['sku в продвижении', 'sku из объединенной карточки', 'продажи', 'заказы'], file.name, unionSheetName);

        if (unionTable) {
          const unionResolver = createResolver(unionTable.headers);
          const unionColumns = {
            promotedSku: unionResolver.find(['sku в продвижении']),
            promotedName: unionResolver.find(['название товара в продвижении']),
            unitedSku: unionResolver.find(['sku из объединенной карточки']),
            unitedName: unionResolver.find(['название товара из объединенной карточки']),
            sales: unionResolver.find(['продажи']),
            orders: unionResolver.find(['заказы'])
          };

          unionRows = unionTable.rows
            .map((row) => ({
              eventDate: monthAnchorDate(reportMonth || context.selectedMonth),
              promotedSku: normalizeSku(getCell(row, unionColumns.promotedSku)),
              promotedName: safeString(getCell(row, unionColumns.promotedName)),
              unitedSku: normalizeSku(getCell(row, unionColumns.unitedSku)),
              unitedName: safeString(getCell(row, unionColumns.unitedName)),
              sales: parseNumber(getCell(row, unionColumns.sales)) || 0,
              orders: parseNumber(getCell(row, unionColumns.orders)) || 0
            }))
            .filter((row) => row.promotedSku || row.unitedSku || row.sales || row.orders);
        }
      }
    } catch (error) {
      console.warn('Не удалось отдельно прочитать листы отчёта продвижения Ozon, используем базовый парсер.', error);
    }
  }

  const resolver = createResolver(statisticsTable.headers);
  const columns = {
    sku: resolver.find(['sku']),
    name: resolver.find(['название товара']),
    instrument: resolver.find(['инструмент']),
    placement: resolver.find(['место размещения']),
    campaignId: resolver.find(['id кампании']),
    spend: resolver.find(['расход']),
    attributedSales: resolver.find(['продажи']),
    attributedOrders: resolver.find(['заказы']),
    ctr: resolver.find(['ctr']),
    impressions: resolver.find(['показы']),
    clicks: resolver.find(['клики']),
    addToCart: resolver.find(['в корзину']),
    addToCartRate: resolver.find(['конверсия в корзину']),
    costPerOrder: resolver.find(['затраты на заказ']),
    cpc: resolver.find(['стоимость клика']),
    drr: resolver.find(['дрр'])
  };

  if (!columns.sku || !columns.spend) {
    throw new Error('Не удалось распознать Ozon «Аналитику продвижения». Нужны столбцы SKU и Расход, ₽.');
  }

  const warnings = [];
  if (reportMonth && context.selectedMonth && reportMonth !== context.selectedMonth) {
    warnings.push(`Файл продвижения относится к ${formatMonthRu(reportMonth)}, а выбран ${formatMonthRu(context.selectedMonth)}.`);
  }

  const anchorDate = monthAnchorDate(reportMonth || context.selectedMonth);

  const rows = statisticsTable.rows
    .map((row) => ({
      eventDate: anchorDate,
      reportDate: anchorDate,
      reportMonth: reportMonth || context.selectedMonth || '',
      sku: normalizeSku(getCell(row, columns.sku)),
      name: safeString(getCell(row, columns.name)),
      instrument: safeString(getCell(row, columns.instrument)),
      placement: safeString(getCell(row, columns.placement)),
      campaignId: safeString(getCell(row, columns.campaignId)),
      spend: Math.abs(parseNumber(getCell(row, columns.spend)) || 0),
      attributedSales: parseNumber(getCell(row, columns.attributedSales)) || 0,
      attributedOrders: parseNumber(getCell(row, columns.attributedOrders)) || 0,
      ctr: parseNumber(getCell(row, columns.ctr)) || 0,
      impressions: parseNumber(getCell(row, columns.impressions)) || 0,
      clicks: parseNumber(getCell(row, columns.clicks)) || 0,
      addToCart: parseNumber(getCell(row, columns.addToCart)) || 0,
      addToCartRate: parseNumber(getCell(row, columns.addToCartRate)) || 0,
      costPerOrder: parseNumber(getCell(row, columns.costPerOrder)) || 0,
      cpc: parseNumber(getCell(row, columns.cpc)) || 0,
      drr: parseNumber(getCell(row, columns.drr)) || 0,
      hasProduct: Boolean(normalizeSku(getCell(row, columns.sku)) || safeString(getCell(row, columns.name)))
    }))
    .filter((row) => row.sku || row.name || row.spend || row.attributedSales || row.attributedOrders);

  return { rows, unionRows, warnings };
}

function parseOzonReturns(rawTable) {
  const resolver = createResolver(rawTable.headers);
  const columns = {
    scheme: resolver.find(['схема']),
    name: resolver.find(['наименование товара']),
    shipmentId: resolver.find(['номер отправления']),
    article: resolver.find(['артикул товара']),
    sku: resolver.find(['ozon sku id']),
    orderDate: resolver.find(['дата оформления заказа']),
    returnDate: resolver.find(['дата возврата']),
    status: resolver.find(['статус возврата']),
    statusDate: resolver.find(['дата статуса']),
    compensationStatus: resolver.find(['статус компенсации']),
    reason: resolver.find(['причина возврата']),
    buyerComment: resolver.find(['комментарий покупателя']),
    quantity: resolver.find(['количество возвращаемых товаров']),
    returnedToSellerDate: resolver.find(['дата возврата продавцу']),
    storageCost: resolver.find(['стоимость хранения']),
    disposalCost: resolver.find(['стоимость утилизации']),
    commission: resolver.find(['комиссия'])
  };

  if (!columns.shipmentId || !columns.returnDate || !columns.name) {
    throw new Error('Не удалось распознать Ozon «Отчёт о возвратах». Нужны столбцы с отправлением, датой возврата и названием товара.');
  }

  const rows = rawTable.rows
    .map((row) => ({
      eventDate: parseDateValue(getCell(row, columns.returnDate)) || parseDateValue(getCell(row, columns.statusDate)) || parseDateValue(getCell(row, columns.returnedToSellerDate)),
      returnDate: parseDateValue(getCell(row, columns.returnDate)),
      orderDate: parseDateValue(getCell(row, columns.orderDate)),
      statusDate: parseDateValue(getCell(row, columns.statusDate)),
      shipmentId: safeString(getCell(row, columns.shipmentId)),
      sku: normalizeSku(getCell(row, columns.sku)),
      article: safeString(getCell(row, columns.article)),
      name: safeString(getCell(row, columns.name)),
      scheme: safeString(getCell(row, columns.scheme)),
      status: safeString(getCell(row, columns.status)),
      compensationStatus: safeString(getCell(row, columns.compensationStatus)),
      reason: safeString(getCell(row, columns.reason)),
      buyerComment: safeString(getCell(row, columns.buyerComment)),
      quantity: Math.max(0, parseNumber(getCell(row, columns.quantity)) || 0),
      storageCost: parseNumber(getCell(row, columns.storageCost)) || 0,
      disposalCost: parseNumber(getCell(row, columns.disposalCost)) || 0,
      commission: parseNumber(getCell(row, columns.commission)) || 0,
      hasProduct: Boolean(normalizeSku(getCell(row, columns.sku)) || safeString(getCell(row, columns.article)) || safeString(getCell(row, columns.name)))
    }))
    .filter((row) => row.eventDate || row.shipmentId || row.sku || row.article || row.name);

  return { rows, warnings: [] };
}

async function parseOzonRealization(rawTable, file, context = {}) {
  if (typeof XLSX === 'undefined' || !file) {
    throw new Error('Для позаказного отчёта о реализации нужен Excel-парсер XLSX.');
  }

  const workbook = await readWorkbookFromFile(file);
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const aoa = sheetToAoaRobust(sheet);

  const headerRowIndex = aoa.findIndex((row) => {
    const normalized = row.map((cell) => normalizeHeader(cell));
    return normalized[0] === 'п п' && normalized.includes('название товара') && normalized.includes('артикул') && normalized.includes('sku');
  });

  if (headerRowIndex < 0 || !aoa[headerRowIndex + 1]) {
    throw new Error('Не удалось распознать структуру позаказного отчёта о реализации Ozon.');
  }

  const periodText = aoa
    .slice(0, headerRowIndex)
    .flatMap((row) => row || [])
    .map((value) => safeString(value))
    .find((value) => /период/i.test(value)) || '';

  const reportMonth = extractMonthKeyFromText(periodText) || extractMonthKeyFromText(file.name) || context.selectedMonth || '';
  const anchorDate = monthAnchorDate(reportMonth || context.selectedMonth);

  const warnings = [];
  if (reportMonth && context.selectedMonth && reportMonth !== context.selectedMonth) {
    warnings.push(`Позаказный отчёт о реализации относится к ${formatMonthRu(reportMonth)}, а выбран ${formatMonthRu(context.selectedMonth)}.`);
  }

  const rows = [];
  for (let index = headerRowIndex + 3; index < aoa.length; index += 1) {
    const row = aoa[index] || [];
    const serial = safeString(row[0]);
    const name = safeString(row[1]);
    const article = safeString(row[2]);
    const sku = normalizeSku(row[3]);
    const shipmentId = safeString(row[21]);

    if (!serial && !name && !article && !sku && !shipmentId) continue;
    if (!shipmentId && /^итого/.test(normalizeHeader(name))) break;
    if (!shipmentId && !sku && !article && !name) break;

    const saleAmount = parseNumber(row[5]) || 0;
    const salePartnerPayout = parseNumber(row[6]) || 0;
    const saleDiscountPoints = parseNumber(row[7]) || 0;
    const saleQty = Math.max(0, parseNumber(row[8]) || 0);
    const salePrice = parseNumber(row[9]) || 0;
    const baseCommissionSale = parseNumber(row[12]) || 0;
    const payoutSale = parseNumber(row[13]) || 0;

    const returnAmount = parseNumber(row[14]) || 0;
    const returnPartnerPayout = parseNumber(row[15]) || 0;
    const returnDiscountPoints = parseNumber(row[16]) || 0;
    const returnQty = Math.max(0, parseNumber(row[17]) || 0);
    const returnPrice = parseNumber(row[18]) || 0;
    const baseCommissionReturn = parseNumber(row[19]) || 0;
    const payoutReturn = parseNumber(row[20]) || 0;

    rows.push({
      eventDate: anchorDate,
      reportDate: anchorDate,
      reportMonth,
      shipmentDate: parseDateValue(row[22]),
      shipmentId,
      sku,
      article,
      name,
      saleAmount,
      salePartnerPayout,
      saleDiscountPoints,
      saleQty,
      salePrice,
      baseCommissionSale,
      payoutSale,
      returnAmount,
      returnPartnerPayout,
      returnDiscountPoints,
      returnQty,
      returnPrice,
      baseCommissionReturn,
      payoutReturn,
      revenue: saleAmount + salePartnerPayout + saleDiscountPoints - returnAmount - returnPartnerPayout - returnDiscountPoints,
      realizedRevenue: saleAmount - returnAmount,
      commission: baseCommissionSale - baseCommissionReturn,
      payout: payoutSale - payoutReturn,
      hasProduct: Boolean(sku || article || name)
    });
  }

  return { rows, warnings };
}


function detectOzonAccrualCategory(groupText, typeText) {
  const groupNorm = normalizeHeader(groupText);
  const typeNorm = normalizeHeader(typeText);
  const normalized = `${groupNorm} ${typeNorm}`.trim();

  if (/выручка|программы партнеров|баллы за скидки|возврат выручки/.test(typeNorm) || /продажи/.test(groupNorm)) {
    return 'base-revenue';
  }

  if (/вознаграждение ozon/.test(groupNorm) || /вознаграждение за продажу|возврат вознаграждения/.test(typeNorm)) {
    return 'base-commission';
  }

  if (/эквайринг/.test(normalized)) {
    return 'payment';
  }

  if (/логистик|обратная логистика|доставка до места выдачи|кросс докинг|кросс-докинг|обработка отправления drop off|обработка отправления drop off партнерами|обработка возвратов|обработка отмененных|обработка отменённых/.test(normalized)) {
    return 'logistics';
  }

  if (/оплата за клик/.test(typeNorm)) {
    return 'promotion-cpc';
  }

  if (/продвижение и реклама/.test(groupNorm) || /продвижение с оплатой за заказ|звездные товары|звёздные товары|premium/.test(normalized)) {
    return 'promotion-other';
  }

  if (/размещение товаров на складах|хранен/.test(normalized)) {
    return 'storage';
  }

  if (/другие услуги и штрафы/.test(groupNorm) || /штраф|нерекомендованный слот|просроченная отгрузка|превышение индекса ошибок/.test(normalized)) {
    return 'penalties';
  }

  if (/компенсации и декомпенсации/.test(groupNorm) || /компенсац|потеря по вине ozon|возмещ/.test(normalized)) {
    return 'compensation';
  }

  return 'other';
}

async function readWorkbookFromFile(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: 'array', raw: false, cellDates: false });
}

function monthAnchorDate(month) {
  if (!month) return null;
  const [year, mon] = String(month).split('-').map(Number);
  if (!year || !mon) return null;
  return new Date(year, mon - 1, 1);
}

function extractMonthKeyFromText(text) {
  const value = safeString(text);

  let match = value.match(/(20\d{2})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (match) {
    return `${match[1]}-${String(Number(match[2])).padStart(2, '0')}`;
  }

  match = value.match(/(\d{1,2})[./-](\d{1,2})[./-](20\d{2})/);
  if (match) {
    return `${match[3]}-${String(Number(match[2])).padStart(2, '0')}`;
  }

  match = value.match(/(20\d{2})(\d{2})(\d{2})/);
  if (match) {
    return `${match[1]}-${match[2]}`;
  }

  match = value.match(/(20\d{2})[-_](\d{2})/);
  if (match) {
    return `${match[1]}-${match[2]}`;
  }

  return '';
}

function normalizeSku(value) {
  const text = safeString(value);
  if (!text) return '';
  return text.replace(/\.0+$/, '');
}

function parseOzonSettlements(rawTable) {
  const resolver = createResolver(rawTable.headers);
  const columns = {
    date: resolver.find(['дата', 'date']),
    type: resolver.find(['тип операции', 'операция', 'вид операции']),
    description: resolver.find(['описание', 'комментарий', 'назначение']),
    amount: resolver.find(['сумма', 'итого', 'amount'])
  };

  if (!columns.amount || (!columns.type && !columns.description)) {
    throw new Error('Не удалось распознать отчёт о взаиморасчётах. Нужны столбцы с типом операции или описанием и суммой.');
  }

  const rows = rawTable.rows.map((row) => {
    const typeText = `${getCell(row, columns.type)} ${getCell(row, columns.description)}`.trim();
    const normalized = normalizeHeader(typeText);
    return {
      eventDate: parseDateValue(getCell(row, columns.date)),
      typeText,
      amount: parseNumber(getCell(row, columns.amount)) || 0,
      category: classifySettlementCategory(normalized)
    };
  }).filter((row) => row.typeText || row.amount);

  return {
    rows,
    warnings: ['Отчёт о взаиморасчётах используется только для общих месячных расходов и сверки, чтобы не задвоить построчные начисления по заказам.']
  };
}

function classifySettlementCategory(normalizedText) {
  if (/продвиж|реклам/.test(normalizedText)) return 'promotion';
  if (/хранен/.test(normalizedText)) return 'storage';
  if (/штраф|неустойк|удерж/.test(normalizedText)) return 'penalties';
  if (/компенсац|возмещ/.test(normalizedText)) return 'compensation';
  if (/логист|достав/.test(normalizedText)) return 'ignore-logistics';
  if (/комис|эквайр|платеж/.test(normalizedText)) return 'ignore-direct';
  return 'otherMarketplace';
}

function parseWbWeeklyDetail(rawTable) {
  const resolver = createResolver(rawTable.headers);
  const columns = {
    subject: resolver.find(['предмет']),
    wbNmId: resolver.find(['код номенклатуры', 'nm id', 'nmid']),
    supplierArticle: resolver.find(['артикул поставщика', 'артикул продавца']),
    name: resolver.find(['название', 'наименование']),
    documentType: resolver.find(['тип документа']),
    paymentReason: resolver.find(['обоснование для оплаты']),
    orderDate: resolver.find(['дата заказа покупателем', 'дата заказа']),
    saleDate: resolver.find(['дата продажи']),
    quantity: resolver.find(['кол-во', 'количество']),
    retailAmount: resolver.find(['вайлдберриз реализовал товар', 'реализовал товар', 'пр']),
    sellerPayout: resolver.find(['к перечислению продавцу за реализованный товар', 'к перечислению продавцу']),
    commissionBase: resolver.find(['вознаграждение вайлдберриз', 'вознаграждение вб', 'вв']),
    commissionVat: resolver.find(['ндс с вознаграждения']),
    paymentFee: resolver.find(['компенсация платёжных услуг', 'эквайринг', 'комиссия за интеграцию платёжных сервисов']),
    pvzFee: resolver.find(['возмещение за выдачу и возврат товаров на пвз']),
    deliveryCost: resolver.find(['услуги по доставке товара покупателю']),
    fineAmount: resolver.find(['общая сумма штрафов']),
    correction: resolver.find(['корректировка вознаграждения вайлдберриз', 'корректировка вознаграждения']),
    operationKinds: resolver.find(['виды логистики, штрафов и корректировок вв', 'виды логистики']),
    storage: resolver.find(['хранение']),
    retentions: resolver.find(['удержания']),
    acceptance: resolver.find(['операции при приёмке', 'операции при приемке']),
    transportReimbursement: resolver.find(['возмещение издержек по перевозке/по складским операциям с товаром', 'возмещение издержек']),
    orderId: resolver.find(['id корзины заказа', 'id заказа']),
    srid: resolver.find(['srid']),
    warehouse: resolver.find(['склад'])
  };

  if (!columns.paymentReason || !columns.saleDate || (!columns.supplierArticle && !columns.wbNmId) || !columns.sellerPayout) {
    throw new Error('Не удалось распознать детализацию еженедельного отчёта WB. Проверьте, что загружен именно XLSX детализации, а не печатная форма отчёта.');
  }

  const warnings = [];
  if (!columns.retailAmount) warnings.push('Не найден столбец фактической цены продажи — выручка будет считаться по сумме к перечислению и может быть занижена.');

  const rows = rawTable.rows.map((row) => {
    const reasonText = safeString(getCell(row, columns.paymentReason));
    const docTypeText = safeString(getCell(row, columns.documentType));
    const operationText = `${reasonText} ${getCell(row, columns.operationKinds)}`.trim();
    const normalizedReason = normalizeHeader(reasonText);
    const normalizedDocType = normalizeHeader(docTypeText);
    const isSale = normalizedReason.includes('продажа') || normalizedDocType.includes('продажа');
    const isReturn = normalizedReason.includes('возврат') || normalizedDocType.includes('возврат');
    const quantity = Math.max(1, parseNumber(getCell(row, columns.quantity)) || 1);
    const revenueRaw = parseNumber(getCell(row, columns.retailAmount));
    const payoutRaw = parseNumber(getCell(row, columns.sellerPayout));
    const revenue = revenueRaw ?? payoutRaw ?? 0;

    return {
      eventDate: parseDateValue(getCell(row, columns.saleDate)),
      orderDate: parseDateValue(getCell(row, columns.orderDate)),
      orderId: safeString(getCell(row, columns.orderId)) || safeString(getCell(row, columns.srid)),
      srid: safeString(getCell(row, columns.srid)),
      sku: safeString(getCell(row, columns.supplierArticle)),
      article: safeString(getCell(row, columns.wbNmId)),
      subject: safeString(getCell(row, columns.subject)),
      name: safeString(getCell(row, columns.name)),
      warehouse: safeString(getCell(row, columns.warehouse)),
      reasonText,
      operationText,
      isSale,
      isReturn,
      quantity,
      revenue,
      revenueSigned: isReturn ? -Math.abs(revenue) : isSale ? Math.abs(revenue) : 0,
      payout: payoutRaw || 0,
      payoutSigned: isReturn ? -Math.abs(payoutRaw || 0) : isSale ? Math.abs(payoutRaw || 0) : 0,
      commission: (parseNumber(getCell(row, columns.commissionBase)) || 0) + (parseNumber(getCell(row, columns.commissionVat)) || 0),
      paymentFee: parseNumber(getCell(row, columns.paymentFee)) || 0,
      pvzFee: parseNumber(getCell(row, columns.pvzFee)) || 0,
      logistics: parseNumber(getCell(row, columns.deliveryCost)) || 0,
      fine: parseNumber(getCell(row, columns.fineAmount)) || 0,
      correction: parseNumber(getCell(row, columns.correction)) || 0,
      storage: parseNumber(getCell(row, columns.storage)) || 0,
      retention: parseNumber(getCell(row, columns.retentions)) || 0,
      acceptance: parseNumber(getCell(row, columns.acceptance)) || 0,
      transportReimbursement: parseNumber(getCell(row, columns.transportReimbursement)) || 0,
      hasProduct: Boolean(safeString(getCell(row, columns.supplierArticle)) || safeString(getCell(row, columns.wbNmId)) || safeString(getCell(row, columns.name)))
    };
  }).filter((row) => row.eventDate || row.orderId || row.sku || row.article);

  return { rows, warnings };
}

function parseWbOrderFeed(rawTable) {
  const resolver = createResolver(rawTable.headers);
  const columns = {
    orderDate: resolver.find(['дата оформления заказа', 'дата заказа']),
    statusDate: resolver.find(['дата текущего статуса', 'дата статуса']),
    status: resolver.find(['статус заказа', 'статус']),
    orderId: resolver.find(['id заказа', 'номер заказа', 'заказ']),
    name: resolver.find(['товар', 'название']),
    price: resolver.find(['цена со скидкой продавца', 'цена'])
  };

  if (!columns.orderDate || !columns.status || !columns.orderId) {
    throw new Error('Не удалось распознать выгрузку «Ленты заказов». Нужны столбцы даты заказа, статуса и ID заказа.');
  }

  const rows = rawTable.rows.map((row) => {
    const status = safeString(getCell(row, columns.status));
    const normalized = normalizeHeader(status);
    return {
      orderDate: parseDateValue(getCell(row, columns.orderDate)),
      statusDate: parseDateValue(getCell(row, columns.statusDate)),
      eventDate: parseDateValue(getCell(row, columns.statusDate)) || parseDateValue(getCell(row, columns.orderDate)),
      orderId: safeString(getCell(row, columns.orderId)),
      name: safeString(getCell(row, columns.name)),
      status,
      isBuyout: /выкуп/.test(normalized),
      isCancel: /отказ|техническая отмен/.test(normalized),
      isReturn: /возврат/.test(normalized),
      price: parseNumber(getCell(row, columns.price)) || 0
    };
  }).filter((row) => row.orderId);

  return { rows, warnings: [] };
}

function buildAnalysis(silent = false) {
  const validation = getValidationState();
  if (!validation.canBuild) return;

  const analysis = state.marketplace === 'ozon' ? analyzeOzon() : analyzeWb();
  state.lastAnalysis = analysis;
  renderResults();
  renderLimitations();
  if (!silent) showToast('Анализ построен.', 'success');
}


function analyzeOzon() {
  const month = state.month;
  const accrualRows = state.files.ozon.accruals.flatMap((file) => file.parsed.rows || []).filter((row) => isDateInSelectedMonth(row.eventDate, month));
  const orderRows = state.files.ozon.orders.flatMap((file) => file.parsed.rows || []);
  const monthOrderRows = orderRows.filter((row) => isDateInSelectedMonth(row.orderDate, month));

  const promotionFile = state.files.ozon.promotion[0]?.parsed || {};
  const promotionRows = (promotionFile.rows || []).filter((row) => isDateInSelectedMonth(row.eventDate || row.reportDate, month));
  const promotionUnionRows = (promotionFile.unionRows || []).filter((row) => isDateInSelectedMonth(row.eventDate, month));

  const returnRows = state.files.ozon.returns.flatMap((file) => file.parsed.rows || []).filter((row) => isDateInSelectedMonth(row.eventDate || row.returnDate, month));
  const realizationRows = state.files.ozon.realization.flatMap((file) => file.parsed.rows || []).filter((row) => isDateInSelectedMonth(row.eventDate || row.reportDate, month));

  const ordersPlaced = uniqueBy(monthOrderRows, (row) => row.shipmentId || row.orderId).length;
  const canceledOrders = uniqueBy(monthOrderRows.filter((row) => row.isCancel), (row) => row.shipmentId || row.orderId).length;
  const buyoutOrders = uniqueBy(realizationRows.filter((row) => (row.saleQty || 0) > 0), (row) => row.shipmentId || `${row.sku}-${row.article}-${row.name}`).length;

  const revenue = sum(realizationRows.map((row) => row.revenue || 0));
  const payoutBeforeServices = sum(realizationRows.map((row) => row.payout || 0));
  const buyoutUnits = sum(realizationRows.map((row) => row.saleQty || 0));
  const financialReturnUnits = sum(realizationRows.map((row) => row.returnQty || 0));
  const operationalReturnUnits = sum(returnRows.map((row) => row.quantity || 0));
  const averageCheck = buyoutOrders ? revenue / buyoutOrders : null;

  const productAnalytics = buildOzonProductAnalytics({
    realizationRows,
    accrualRows,
    promotionRows,
    promotionUnionRows,
    returnRows,
    orderRows: monthOrderRows,
    usePromotionFile: Boolean(state.files.ozon.promotion.length)
  });

  const productSummary = productAnalytics.products;
  applyManualCostsToProducts(productSummary, 'ozon');

  const expenses = mergeExpenseGroups(productAnalytics.expenses, productAnalytics.commonExpenses);
  const commonCosts = sum(Object.values(productAnalytics.commonExpenses));
  const linkedMarketCosts = sum(productSummary.map((item) => item.marketplaceCosts));
  const totalManualCosts = sum(productSummary.map((item) => item.manualCosts));
  const netResult = revenue - linkedMarketCosts - commonCosts - totalManualCosts;

  const dailySeries = buildDailySeries(accrualRows.filter((row) => row.category === 'base-revenue').map((row) => ({ date: row.eventDate, value: row.amount || 0 })), month);

  const availability = {
    exact: ['Выручка по цене продавца', 'К начислению по товарам', 'Заказы', 'Продажи по SKU', 'Базовая комиссия Ozon', 'Логистика', 'Эквайринг', 'Реклама по SKU', 'Финансовые возвраты', 'Операционные возвраты и причины'],
    partial: ['Отмены считаются по текущему статусу заказов в выгрузке Ozon: в файле нет отдельной даты отмены.'],
    manual: ['Себестоимость', 'Упаковка', 'Логистика до склада', 'Прочие внутренние расходы']
  };

  const limitations = [
    'Позаказный отчёт о реализации задаёт товарную выручку по цене продавца, сумму к начислению и финансовые возвраты.',
    'Отчёт по начислениям используется как слой расходов Ozon поверх товарной выручки: логистика, эквайринг, продвижение, штрафы, хранение и прочие сервисы.',
    'Если рекламный файл загружен, расходы за «Оплату за клик» берутся из него, чтобы точнее разложить рекламу по SKU и не задвоить начисления.',
    'Отчёт о возвратах показывает операционные возвраты, которые могут быть шире финансовых возвратов в позаказной реализации.',
    'Ручные затраты применяются к чистому количеству проданных единиц в выбранном месяце.'
  ];

  const insights = buildInsights({
    marketplace: 'ozon',
    revenue,
    ordersPlaced,
    canceledOrders,
    returnUnits: operationalReturnUnits,
    buyoutUnits,
    totalMarketplaceCosts: linkedMarketCosts + commonCosts,
    commonCosts,
    productSummary
  });

  const returnReasonMap = new Map();
  returnRows.forEach((row) => {
    const key = row.reason || 'Без причины';
    returnReasonMap.set(key, (returnReasonMap.get(key) || 0) + (row.quantity || 0));
  });
  const topReturnReason = Array.from(returnReasonMap.entries()).sort((a, b) => b[1] - a[1])[0];
  const adSpend = sum(productSummary.map((item) => item.promotion || 0));
  if (topReturnReason && topReturnReason[1] > 0) {
    insights.push(`Самая частая причина возврата: «${topReturnReason[0]}» — ${nfNumber.format(topReturnReason[1])} шт.`);
  }
  if (adSpend && revenue && adSpend / revenue > 0.12) {
    insights.push(`Расходы на продвижение занимают ${fmtPct(adSpend / revenue)} выручки по цене продавца. Проверьте товары с высоким ДРР и низким вкладом в итог.`);
  }

  return {
    marketplace: 'ozon',
    month,
    title: 'Ozon',
    subtitle: 'Картина собрана из пяти слоёв Ozon: начисления, заказы, продвижение, возвраты и позаказная реализация.',
    availability,
    limitations,
    kpis: [
      makeKpi('Выручка', revenue, 'точно', 'По цене продавца, уже с учётом возвратов', 'money'),
      makeKpi('К начислению по товарам', payoutBeforeServices, 'точно', 'После базовой комиссии Ozon, до логистики и прочих сервисов', 'money'),
      makeKpi('Заказы', ordersPlaced, 'точно', 'По заказам, принятым в обработку в выбранном месяце', 'number'),
      makeKpi('Выкуплено товаров', buyoutUnits, 'точно', 'По позаказной реализации', 'number'),
      makeKpi('Отмены', canceledOrders, 'частично', 'По текущему статусу заказа в выгрузке заказов', 'number'),
      makeKpi('Финансовые возвраты', financialReturnUnits, 'точно', 'По строкам возвратов в позаказной реализации', 'number'),
      makeKpi('Операционные возвраты', operationalReturnUnits, 'точно', 'По отдельному отчёту возвратов Ozon', 'number'),
      makeKpi('Средний чек', averageCheck, 'точно', 'Выручка / число выкупленных отправлений', 'money'),
      makeKpi('Расходы маркетплейса', linkedMarketCosts + commonCosts, 'точно', 'Комиссия, логистика, эквайринг, продвижение, хранение, штрафы и прочие сервисы', 'money'),
      makeKpi('Итог после ручных затрат', netResult, totalManualCosts ? 'точно' : 'частично', totalManualCosts ? 'Включая ручной блок затрат' : 'Себестоимость не заполнена или заполнена частично', 'money')
    ],
    dailySeries,
    expenses,
    productSummary,
    commonCosts,
    totalManualCosts,
    netResult,
    insights: insights.slice(0, 6),
    tableNote: commonCosts
      ? `Общие расходы Ozon без привязки к SKU: ${fmtMoney(commonCosts)}. Они показаны отдельно и не размазаны по товарам. Финансовых возвратов: ${nfNumber.format(financialReturnUnits)} шт., операционных возвратов: ${nfNumber.format(operationalReturnUnits)} шт.`
      : `Все найденные расходы удалось привязать к товарам. Финансовых возвратов: ${nfNumber.format(financialReturnUnits)} шт., операционных возвратов: ${nfNumber.format(operationalReturnUnits)} шт.`
  };
}

function analyzeWb() {
  const month = state.month;
  const detailRows = state.files.wb.weeklyDetail.flatMap((file) => file.parsed.rows || []).filter((row) => isDateInSelectedMonth(row.eventDate, month));
  const orderFeedRows = (state.files.wb.orderFeed[0]?.parsed.rows || []);

  const revenue = sum(detailRows.map((row) => row.revenueSigned || 0));
  const buyoutUnits = sum(detailRows.filter((row) => row.isSale).map((row) => row.quantity || 0));
  const returnUnits = sum(detailRows.filter((row) => row.isReturn).map((row) => row.quantity || 0));
  const buyoutOrders = uniqueBy(detailRows.filter((row) => row.isSale), (row) => row.orderId || row.srid).length;
  const averageCheck = buyoutOrders ? revenue / buyoutOrders : null;

  const exactOrders = state.files.wb.orderFeed.length
    ? uniqueBy(orderFeedRows.filter((row) => isDateInSelectedMonth(row.orderDate, month)), (row) => row.orderId).length
    : null;
  const exactCancels = state.files.wb.orderFeed.length
    ? uniqueBy(orderFeedRows.filter((row) => isDateInSelectedMonth(row.statusDate || row.eventDate, month) && row.isCancel), (row) => row.orderId).length
    : null;

  const expenses = {
    commission: sum(detailRows.map((row) => row.commission || 0)),
    logistics: sum(detailRows.map((row) => row.logistics || 0)),
    payment: sum(detailRows.map((row) => row.paymentFee || 0)) + sum(detailRows.map((row) => row.pvzFee || 0)),
    storage: sum(detailRows.map((row) => row.storage || 0)),
    promotion: 0,
    penalties: sum(detailRows.map((row) => row.fine || 0)) + positiveOnly(sum(detailRows.map((row) => row.correction || 0))),
    otherMarketplace: sum(detailRows.map((row) => row.acceptance || 0)) + positiveOnly(sum(detailRows.map((row) => row.retention || 0))),
    compensation: -sum(detailRows.map((row) => row.transportReimbursement || 0))
  };

  const productSummary = buildWbProducts(detailRows);
  applyManualCostsToProducts(productSummary, 'wb');

  const linkedMarketCosts = sum(productSummary.map((item) => item.marketplaceCosts));
  const commonCosts = calcWbCommonCosts(detailRows);
  const totalManualCosts = sum(productSummary.map((item) => item.manualCosts));
  const netResult = revenue - linkedMarketCosts - commonCosts - totalManualCosts;

  const dailySeries = buildDailySeries(detailRows.map((row) => ({ date: row.eventDate, value: row.revenueSigned || 0 })), month);

  const availability = {
    exact: ['Выручка', 'Выкупленные товары', 'Возвраты', 'Комиссия WB', 'Эквайринг / платёжные услуги', 'Логистика', 'Хранение', 'Приёмка', 'Штрафы / удержания'],
    partial: state.files.wb.orderFeed.length
      ? ['Количество заказов и отмен посчитано точно по «Ленте заказов».']
      : ['Без «Ленты заказов» точное количество заказов и отмен недоступно: финансовая детализация WB не покрывает весь поток заказов как отдельный архив за любой месяц.'],
    manual: ['Себестоимость', 'Упаковка', 'Логистика до склада', 'Прочие внутренние расходы']
  };

  const limitations = [
    'Для точного месячного анализа по WB нужно загрузить все недельные детализации, которые пересекают выбранный месяц.',
    'Если «Лента заказов» не загружена, приложение честно не считает точные заказы и отмены.',
    'Общие удержания и часть расходов WB могут быть не привязаны к SKU — такие суммы показываются отдельно.',
    'Ручные затраты применяются к чистому количеству проданных единиц в выбранном месяце.'
  ];

  const insights = buildInsights({
    marketplace: 'wb',
    revenue,
    ordersPlaced: exactOrders,
    canceledOrders: exactCancels,
    returnUnits,
    buyoutUnits,
    totalMarketplaceCosts: linkedMarketCosts + commonCosts,
    commonCosts,
    productSummary
  });

  return {
    marketplace: 'wb',
    month,
    title: 'Wildberries',
    subtitle: state.files.wb.orderFeed.length
      ? 'Экономика построена по детализации еженедельных отчётов реализации и дополнена «Лентой заказов». '
      : 'Экономика построена по детализации еженедельных отчётов реализации. Заказы и отмены без «Ленты заказов» остаются недоступны.',
    availability,
    limitations,
    kpis: [
      makeKpi('Выручка', revenue, 'точно', 'Чистая выручка по операциям месяца', 'money'),
      makeKpi('Заказы', exactOrders, exactOrders == null ? 'недоступно' : 'точно', exactOrders == null ? 'Нужна «Лента заказов» за период' : 'По дате заказа из «Ленты заказов»', 'number'),
      makeKpi('Выкуплено товаров', buyoutUnits, 'точно', 'Количество единиц по строкам продажи', 'number'),
      makeKpi('Отмены', exactCancels, exactCancels == null ? 'недоступно' : 'точно', exactCancels == null ? 'Нужна «Лента заказов» за период' : 'По статусам отказа / техотмены', 'number'),
      makeKpi('Возвраты', returnUnits, 'точно', 'Количество возвращённых единиц', 'number'),
      makeKpi('Средний чек', averageCheck, 'точно', 'Выручка / количество выкупленных заказов', 'money'),
      makeKpi('Расходы маркетплейса', linkedMarketCosts + commonCosts, 'точно', 'Комиссия, логистика, эквайринг, хранение, удержания, приёмка', 'money'),
      makeKpi('Итог после ручных затрат', netResult, totalManualCosts ? 'точно' : 'частично', totalManualCosts ? 'Включая ручной блок затрат' : 'Себестоимость не заполнена или заполнена частично', 'money')
    ],
    dailySeries,
    expenses,
    productSummary,
    commonCosts,
    totalManualCosts,
    netResult,
    insights,
    tableNote: commonCosts
      ? `Общие расходы WB, не привязанные к SKU: ${fmtMoney(commonCosts)}. Они показаны отдельно и не распределены по товарам.`
      : 'Все найденные расходы удалось привязать к операциям или товарам.'
  };
}


function buildOzonProductAnalytics({ realizationRows, accrualRows, promotionRows, promotionUnionRows, returnRows, orderRows, usePromotionFile }) {
  const map = new Map();
  const commonExpenses = {
    commission: 0,
    payment: 0,
    logistics: 0,
    storage: 0,
    promotion: 0,
    penalties: 0,
    otherMarketplace: 0,
    compensation: 0
  };

  function ensureProduct(row) {
    const sku = normalizeSku(row?.sku);
    const article = safeString(row?.article);
    const name = safeString(row?.name) || 'Без названия';
    const key = sku || article || name;
    if (!map.has(key)) {
      const item = makeProductAccumulator(sku || article || name, name);
      item.article = article;
      item.name = name;
      map.set(key, item);
    }
    const item = map.get(key);
    if (!item.article && article) item.article = article;
    if ((!item.name || item.name === 'Без названия') && name) item.name = name;
    return item;
  }

  realizationRows.forEach((row) => {
    const item = ensureProduct(row);
    const orderKey = row.shipmentId || `${item.sku}-${item.article}-${item.name}-${item.rows + 1}`;
    item.rows += 1;
    item.revenue += row.revenue || 0;
    item.realizedRevenue += row.realizedRevenue || 0;
    item.payout += row.payout || 0;
    item.commission += row.commission || 0;
    item.unitsSold += row.saleQty || 0;
    item.unitsReturned += row.returnQty || 0;
    if ((row.saleQty || 0) > 0) item.buyoutOrders.add(orderKey);
  });

  accrualRows.forEach((row) => {
    const shouldSkipCpc = usePromotionFile && row.category === 'promotion-cpc';
    if (row.category === 'base-revenue' || row.category === 'base-commission' || shouldSkipCpc) {
      return;
    }

    const amount = Math.abs(row.amount || 0);
    const target = row.hasProduct ? ensureProduct(row) : null;

    function applyToTarget(field, signedAmount = amount) {
      if (target) {
        target[field] += signedAmount;
      } else if (field in commonExpenses) {
        commonExpenses[field] += signedAmount;
      }
    }

    switch (row.category) {
      case 'payment':
        applyToTarget('payment');
        break;
      case 'logistics':
        applyToTarget('logistics');
        break;
      case 'promotion-cpc':
      case 'promotion-other':
        applyToTarget('promotion');
        break;
      case 'storage':
        applyToTarget('storage');
        break;
      case 'penalties':
        applyToTarget('penalties');
        break;
      case 'compensation':
        applyToTarget('compensation', -amount);
        break;
      default:
        if (target) {
          target.otherServices += amount;
        } else {
          commonExpenses.otherMarketplace += amount;
        }
        break;
    }
  });

  promotionRows.forEach((row) => {
    const item = ensureProduct(row);
    item.promotion += row.spend || 0;
    item.adSpend += row.spend || 0;
    item.adAttributedSales += row.attributedSales || 0;
    item.adAttributedOrders += row.attributedOrders || 0;
    item.adClicks += row.clicks || 0;
    item.adImpressions += row.impressions || 0;
    item.adAddToCart += row.addToCart || 0;
    item.rows += 1;
  });

  promotionUnionRows.forEach((row) => {
    const item = ensureProduct({ sku: row.promotedSku, name: row.promotedName });
    item.adUnionSales += row.sales || 0;
    item.adUnionOrders += row.orders || 0;
    item.adAttributedSales += row.sales || 0;
    item.adAttributedOrders += row.orders || 0;
  });

  returnRows.forEach((row) => {
    const item = ensureProduct(row);
    item.operationalReturns += row.quantity || 0;
    const reasonKey = row.reason || 'Без причины';
    item.returnReasons.set(reasonKey, (item.returnReasons.get(reasonKey) || 0) + (row.quantity || 0));
    item.rows += 1;
  });

  orderRows.forEach((row) => {
    const item = ensureProduct(row);
    const orderKey = row.shipmentId || row.orderId || `${item.sku}-${item.article}-${item.name}-${item.rows + 1}`;
    item.ordersPlacedSet.add(orderKey);
    if (row.isCancel) item.canceledOrdersSet.add(orderKey);
  });

  const products = finalizeProducts(
    Array.from(map.values()).map((item) => {
      item.otherMarketplace = (item.storage || 0) + (item.penalties || 0) + (item.otherServices || 0) + (item.compensation || 0);
      item.marketplaceCosts = (item.commission || 0) + (item.payment || 0) + (item.logistics || 0) + (item.promotion || 0) + item.otherMarketplace;
      item.ordersPlaced = item.ordersPlacedSet.size;
      item.canceledOrders = item.canceledOrdersSet.size;
      return item;
    })
  );

  const expenses = {
    commission: sum(products.map((item) => item.commission || 0)),
    payment: sum(products.map((item) => item.payment || 0)),
    logistics: sum(products.map((item) => item.logistics || 0)),
    storage: sum(products.map((item) => item.storage || 0)),
    promotion: sum(products.map((item) => item.promotion || 0)),
    penalties: sum(products.map((item) => item.penalties || 0)),
    otherMarketplace: sum(products.map((item) => item.otherServices || 0)),
    compensation: sum(products.map((item) => item.compensation || 0))
  };


  return {
    products,
    expenses,
    commonExpenses: {
      commission: commonExpenses.commission,
      payment: commonExpenses.payment,
      logistics: commonExpenses.logistics,
      storage: commonExpenses.storage,
      promotion: commonExpenses.promotion,
      penalties: commonExpenses.penalties,
      otherMarketplace: commonExpenses.otherMarketplace,
      compensation: commonExpenses.compensation
    }
  };
}

function buildWbProducts(rows) {
  const map = new Map();
  let fallbackCounter = 0;

  rows.forEach((row) => {
    const rawKey = row.sku || row.article || row.name || `Без SKU ${fallbackCounter += 1}`;
    if (!map.has(rawKey)) {
      map.set(rawKey, makeProductAccumulator(rawKey, row.name || row.article || row.sku || 'Без названия'));
    }
    const item = map.get(rawKey);
    const orderKey = row.orderId || row.srid || `${rawKey}-${item.rows + 1}`;
    item.rows += 1;
    item.name = item.name || row.name || rawKey;
    if (row.isSale) {
      item.unitsSold += row.quantity || 0;
      item.buyoutOrders.add(orderKey);
    }
    if (row.isReturn) item.unitsReturned += row.quantity || 0;

    item.revenue += row.revenueSigned || 0;
    item.payout += row.payoutSigned || 0;

    if (row.hasProduct) {
      const costs = (row.commission || 0) + (row.paymentFee || 0) + (row.pvzFee || 0) + (row.logistics || 0) + (row.fine || 0) + positiveOnly(row.correction || 0) + (row.acceptance || 0) + positiveOnly(row.retention || 0) + (row.storage || 0) - (row.transportReimbursement || 0);
      item.commission += (row.commission || 0) + (row.paymentFee || 0);
      item.logistics += row.logistics || 0;
      item.otherMarketplace += (row.pvzFee || 0) + (row.fine || 0) + positiveOnly(row.correction || 0) + (row.acceptance || 0) + positiveOnly(row.retention || 0) + (row.storage || 0) - (row.transportReimbursement || 0);
      item.marketplaceCosts += costs;
    }
  });

  return finalizeProducts(Array.from(map.values()));
}

function makeProductAccumulator(sku, name) {
  return {
    sku,
    article: '',
    name,
    rows: 0,
    unitsSold: 0,
    unitsReturned: 0,
    operationalReturns: 0,
    revenue: 0,
    realizedRevenue: 0,
    payout: 0,
    commission: 0,
    payment: 0,
    logistics: 0,
    promotion: 0,
    storage: 0,
    penalties: 0,
    otherServices: 0,
    otherMarketplace: 0,
    compensation: 0,
    marketplaceCosts: 0,
    manualCosts: 0,
    netResult: 0,
    marginPct: null,
    adSpend: 0,
    adAttributedSales: 0,
    adAttributedOrders: 0,
    adClicks: 0,
    adImpressions: 0,
    adAddToCart: 0,
    adUnionSales: 0,
    adUnionOrders: 0,
    ordersPlaced: 0,
    canceledOrders: 0,
    buyoutOrders: new Set(),
    ordersPlacedSet: new Set(),
    canceledOrdersSet: new Set(),
    returnReasons: new Map()
  };
}

function finalizeProducts(items) {
  return items
    .map((item) => ({
      ...item,
      buyoutOrdersCount: item.buyoutOrders.size,
      buyoutOrders: undefined,
      ordersPlacedSet: undefined,
      canceledOrdersSet: undefined,
      topReturnReason: Array.from(item.returnReasons.entries()).sort((a, b) => b[1] - a[1])[0]?.[0] || '',
      returnReasons: undefined
    }))
    .sort((a, b) => b.revenue - a.revenue || b.netResult - a.netResult || a.name.localeCompare(b.name, 'ru'));
}

function applyManualCostsToProducts(products, marketplace) {
  products.forEach((item) => {
    const unitCosts = getManualUnitCost(marketplace, item.sku);
    const unitsForManual = Math.max(0, (item.unitsSold || 0) - (item.unitsReturned || 0));
    item.manualCosts = unitsForManual * unitCosts.total;
    item.netResult = item.revenue - item.marketplaceCosts - item.manualCosts;
    item.marginPct = item.revenue ? item.netResult / item.revenue : null;
  });
}

function getManualUnitCost(marketplace, sku) {
  const defaults = state.manualCosts.defaults || {};
  const item = state.manualCosts.items[getManualCostKey(marketplace, sku)] || {};

  const cogs = parseNumber(item.cogs) ?? parseNumber(defaults.cogs) ?? 0;
  const packaging = parseNumber(item.packaging) ?? parseNumber(defaults.packaging) ?? 0;
  const inbound = parseNumber(item.inbound) ?? parseNumber(defaults.inbound) ?? 0;
  const other = parseNumber(item.other) ?? parseNumber(defaults.other) ?? 0;

  return {
    cogs,
    packaging,
    inbound,
    other,
    total: cogs + packaging + inbound + other
  };
}

function getManualCostKey(marketplace, sku) {
  return `${marketplace}::${sku}`;
}

function summarizeSettlementRows(rows) {
  const totals = {
    commission: 0,
    logistics: 0,
    payment: 0,
    storage: 0,
    promotion: 0,
    penalties: 0,
    otherMarketplace: 0,
    compensation: 0
  };

  rows.forEach((row) => {
    switch (row.category) {
      case 'promotion':
        totals.promotion += Math.abs(row.amount);
        break;
      case 'storage':
        totals.storage += Math.abs(row.amount);
        break;
      case 'penalties':
        totals.penalties += Math.abs(row.amount);
        break;
      case 'compensation':
        totals.compensation -= Math.abs(row.amount);
        break;
      case 'otherMarketplace':
        totals.otherMarketplace += Math.abs(row.amount);
        break;
      default:
        break;
    }
  });

  return totals;
}

function mergeExpenseGroups(base, extra) {
  const result = { ...base };
  Object.keys(extra).forEach((key) => {
    result[key] = (result[key] || 0) + (extra[key] || 0);
  });
  return result;
}

function calcOzonCommonCosts(accrualRows, settlementRows) {
  const fromAccruals = accrualRows
    .filter((row) => !row.hasProduct)
    .reduce((sum, row) => sum + (row.commission || 0) + (row.logistics || 0) + (row.paymentFee || 0) + (row.storage || 0) + (row.promotion || 0) + (row.penalty || 0) + (row.service || 0) - (row.compensation || 0), 0);

  const fromSettlements = settlementRows.reduce((sum, row) => {
    if (row.category === 'ignore-logistics' || row.category === 'ignore-direct') return sum;
    if (row.category === 'compensation') return sum - Math.abs(row.amount);
    return sum + Math.abs(row.amount);
  }, 0);

  return fromAccruals + fromSettlements;
}

function calcWbCommonCosts(detailRows) {
  return detailRows
    .filter((row) => !row.hasProduct)
    .reduce((sum, row) => sum + (row.storage || 0) + positiveOnly(row.retention || 0) + (row.acceptance || 0) + (row.fine || 0) + positiveOnly(row.correction || 0) - (row.transportReimbursement || 0), 0);
}

function buildDailySeries(rows, month) {
  const map = new Map();
  const dates = getMonthDateList(month);
  dates.forEach((date) => map.set(date, 0));

  rows.forEach((entry) => {
    const key = toDateKey(entry.date);
    if (!key || !map.has(key)) return;
    map.set(key, (map.get(key) || 0) + (entry.value || 0));
  });

  return Array.from(map.entries()).map(([date, value]) => ({ date, value }));
}

function buildInsights({ marketplace, revenue, ordersPlaced, canceledOrders, returnUnits, buyoutUnits, totalMarketplaceCosts, commonCosts, productSummary }) {
  const insights = [];
  const top1 = productSummary[0];
  const top3Revenue = sum(productSummary.slice(0, 3).map((item) => item.revenue));
  const negativeItems = productSummary.filter((item) => item.netResult < 0);
  const returnRate = buyoutUnits ? returnUnits / buyoutUnits : 0;
  const expenseRate = revenue ? totalMarketplaceCosts / revenue : 0;

  if (revenue && top1 && top1.revenue / revenue > 0.35) {
    insights.push(`Один товар даёт ${fmtPct(top1.revenue / revenue)} выручки. Есть заметная зависимость от лидера — стоит контролировать остатки и риски по этой позиции.`);
  }

  if (revenue && top3Revenue / revenue > 0.7) {
    insights.push(`Три лидера формируют ${fmtPct(top3Revenue / revenue)} выручки. Ассортимент зависит от узкой группы SKU.`);
  }

  if (returnRate > 0.12) {
    insights.push(`Доля возвратов высокая: ${fmtPct(returnRate)} от выкупленных единиц. Проверьте карточки, размерные сетки, комплектацию и ожидания покупателей.`);
  } else if (returnRate > 0.05) {
    insights.push(`Возвраты заметны: ${fmtPct(returnRate)} от выкупленных единиц. Имеет смысл отдельно посмотреть товары с наибольшим возвратом.`);
  }

  if (expenseRate > 0.3) {
    insights.push(`Расходы маркетплейса съедают ${fmtPct(expenseRate)} выручки. В первую очередь проверьте комиссию, логистику и удержания.`);
  }

  if (commonCosts > 0 && revenue) {
    insights.push(`Есть ${fmtMoney(commonCosts)} общих расходов без привязки к SKU. Их важно контролировать отдельно, чтобы не искажать экономику отдельных товаров.`);
  }

  if (negativeItems.length) {
    insights.push(`Убыточных SKU: ${nfNumber.format(negativeItems.length)}. Они собраны в блоке «Слабые товары».`);
  }

  if (marketplace === 'wb' && ordersPlaced == null) {
    insights.push('Для WB точное количество заказов и отмен появится только после загрузки «Ленты заказов» за нужный период.');
  }

  if (ordersPlaced != null && canceledOrders != null && ordersPlaced > 0) {
    const cancelRate = canceledOrders / ordersPlaced;
    if (cancelRate > 0.1) {
      insights.push(`Отмены занимают ${fmtPct(cancelRate)} от заказов. Проверьте сроки доставки, остатки и причины отказов.`);
    }
  }

  if (!insights.length) {
    insights.push('Явных аномалий в выбранном месяце не обнаружено. Основные метрики выглядят ровно.');
  }

  return insights.slice(0, 6);
}

function makeKpi(label, value, stateLabel, sub, kind) {
  return { label, value, stateLabel, sub, kind };
}

function renderResults() {
  if (!state.lastAnalysis || state.lastAnalysis.marketplace !== state.marketplace) {
    dom.resultsSection.classList.add('hidden');
    return;
  }

  const analysis = state.lastAnalysis;
  dom.resultsSection.classList.remove('hidden');
  dom.resultsSubtitle.textContent = analysis.subtitle;
  renderAvailability(analysis);
  renderKpis(analysis.kpis);
  renderLineChart(analysis.dailySeries);
  renderDonutChart(analysis.expenses);
  renderBarChart(analysis.productSummary);
  renderInsights(analysis.insights);
  renderExpenseBreakdown(analysis.expenses, analysis.commonCosts, analysis.totalManualCosts, analysis.netResult);
  renderWeakProducts(analysis.productSummary);
  dom.tableNote.textContent = analysis.tableNote;
  renderResultsTable();
}

function renderAvailability(analysis) {
  dom.availabilityGrid.innerHTML = `
    <section class="availability-card">
      <h3>Точно считается</h3>
      <ul>${analysis.availability.exact.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>
    </section>
    <section class="availability-card">
      <h3>Считается частично / зависит от доп. файла</h3>
      <ul>${analysis.availability.partial.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>
    </section>
    <section class="availability-card">
      <h3>Требует ручного ввода</h3>
      <ul>${analysis.availability.manual.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>
    </section>
  `;
}

function renderKpis(kpis) {
  dom.kpiGrid.innerHTML = kpis
    .map((kpi) => {
      const badgeClass = kpi.stateLabel === 'точно' ? 'exact' : kpi.stateLabel === 'частично' ? 'partial' : 'unavailable';
      const badgeLabel = kpi.stateLabel === 'точно' ? 'Точно' : kpi.stateLabel === 'частично' ? 'Частично' : 'Недоступно';
      return `
        <article class="kpi-card">
          <div class="kpi-top">
            <span class="kpi-label">${escapeHtml(kpi.label)}</span>
            <span class="badge ${badgeClass}">${badgeLabel}</span>
          </div>
          <div class="kpi-value ${typeof kpi.value === 'number' && kpi.value < 0 ? 'metric-negative' : ''}">${formatKpiValue(kpi.value, kpi.kind)}</div>
          <div class="kpi-sub">${escapeHtml(kpi.sub)}</div>
        </article>
      `;
    })
    .join('');
}

function renderInsights(insights) {
  dom.insightsList.innerHTML = insights.map((item) => `<li>${escapeHtml(item)}</li>`).join('');
}

function renderExpenseBreakdown(expenses, commonCosts, totalManualCosts, netResult) {
  const items = [
    ['Комиссия', expenses.commission || 0],
    ['Логистика', expenses.logistics || 0],
    ['Эквайринг / платёжные услуги', expenses.payment || 0],
    ['Хранение', expenses.storage || 0],
    ['Продвижение', expenses.promotion || 0],
    ['Штрафы', expenses.penalties || 0],
    ['Прочие расходы MP', expenses.otherMarketplace || 0],
    ['Компенсации / возмещения', expenses.compensation || 0],
    ['Общие расходы без SKU', commonCosts || 0],
    ['Ручные затраты', totalManualCosts || 0],
    ['Итог после затрат', netResult || 0]
  ].filter(([, value], index) => index >= 8 || Math.abs(value) > 0.01);

  dom.expenseBreakdown.innerHTML = items
    .map(([label, value]) => `
      <div class="expense-row">
        <div>
          <strong>${escapeHtml(label)}</strong>
          <span>${value < 0 ? 'Снижает расходы / добавляет доход' : 'Сумма за выбранный месяц'}</span>
        </div>
        <div class="mono ${value < 0 ? 'metric-positive' : value === netResult ? (value < 0 ? 'metric-negative' : 'metric-positive') : ''}">${fmtMoney(value)}</div>
      </div>
    `)
    .join('');
}

function renderWeakProducts(products) {
  const weak = products
    .filter((item) => item.netResult < 0 || (item.revenue > 0 && ((item.operationalReturns || item.unitsReturned) / Math.max(item.unitsSold, 1)) > 0.2))
    .slice(0, 6);

  if (!weak.length) {
    dom.weakProducts.innerHTML = '<div class="small-note">Явно слабых позиций в выбранном месяце не найдено.</div>';
    return;
  }

  dom.weakProducts.innerHTML = weak
    .map((item) => `
      <div class="weak-row">
        <div>
          <strong>${escapeHtml(item.name)}</strong>
          <span>${escapeHtml(item.sku || 'Без SKU')} • выручка ${fmtMoney(item.revenue)}${item.operationalReturns ? ` • опер. возвраты ${nfNumber.format(item.operationalReturns)}` : ''}</span>
        </div>
        <div class="mono metric-negative">${fmtMoney(item.netResult)}</div>
      </div>
    `)
    .join('');
}


function renderResultsTable() {
  if (!state.lastAnalysis || state.lastAnalysis.marketplace !== state.marketplace) return;
  const products = state.lastAnalysis.productSummary.filter((item) => {
    if (!state.resultsFilterText) return true;
    const haystack = `${item.sku} ${item.article || ''} ${item.name}`.toLowerCase();
    return haystack.includes(state.resultsFilterText);
  });

  const thead = dom.productTable.querySelector('thead');
  const tbody = dom.productTable.querySelector('tbody');

  if (state.marketplace === 'ozon') {
    thead.innerHTML = `
      <tr>
        <th>SKU / артикул</th>
        <th>Товар</th>
        <th>Выручка</th>
        <th>К начислению</th>
        <th>Выкупы, шт.</th>
        <th>Фин. возвраты, шт.</th>
        <th>Опер. возвраты, шт.</th>
        <th>Комиссия / эквайринг</th>
        <th>Логистика</th>
        <th>Продвижение</th>
        <th>Прочие MP</th>
        <th>Ручные затраты</th>
        <th>Итог</th>
        <th>Маржа</th>
      </tr>
    `;

    tbody.innerHTML = products
      .map((item) => `
        <tr>
          <td class="mono">${escapeHtml(item.sku || '—')}${item.article ? `<div class="small-note">${escapeHtml(item.article)}</div>` : ''}</td>
          <td>${escapeHtml(item.name || 'Без названия')}</td>
          <td class="mono">${fmtMoney(item.revenue)}</td>
          <td class="mono">${fmtMoney(item.payout)}</td>
          <td class="mono">${nfNumber.format(item.unitsSold || 0)}</td>
          <td class="mono">${nfNumber.format(item.unitsReturned || 0)}</td>
          <td class="mono">${nfNumber.format(item.operationalReturns || 0)}</td>
          <td class="mono">${fmtMoney((item.commission || 0) + (item.payment || 0))}</td>
          <td class="mono">${fmtMoney(item.logistics)}</td>
          <td class="mono">${fmtMoney(item.promotion)}</td>
          <td class="mono">${fmtMoney(item.otherMarketplace)}</td>
          <td class="mono">${fmtMoney(item.manualCosts)}</td>
          <td class="mono ${item.netResult < 0 ? 'metric-negative' : 'metric-positive'}">${fmtMoney(item.netResult)}</td>
          <td class="mono">${item.marginPct == null ? '—' : fmtPct(item.marginPct)}</td>
        </tr>
      `)
      .join('');

    return;
  }

  thead.innerHTML = `
    <tr>
      <th>SKU / артикул</th>
      <th>Товар</th>
      <th>Выручка</th>
      <th>Выкупы, шт.</th>
      <th>Возвраты, шт.</th>
      <th>Комиссия и платёжные услуги</th>
      <th>Логистика</th>
      <th>Прочие MP</th>
      <th>Ручные затраты</th>
      <th>Итог</th>
      <th>Маржа</th>
    </tr>
  `;

  tbody.innerHTML = products
    .map((item) => `
      <tr>
        <td class="mono">${escapeHtml(item.sku || '—')}</td>
        <td>${escapeHtml(item.name || 'Без названия')}</td>
        <td class="mono">${fmtMoney(item.revenue)}</td>
        <td class="mono">${nfNumber.format(item.unitsSold || 0)}</td>
        <td class="mono">${nfNumber.format(item.unitsReturned || 0)}</td>
        <td class="mono">${fmtMoney((item.commission || 0) + (item.payment || 0))}</td>
        <td class="mono">${fmtMoney(item.logistics)}</td>
        <td class="mono">${fmtMoney(item.otherMarketplace)}</td>
        <td class="mono">${fmtMoney(item.manualCosts)}</td>
        <td class="mono ${item.netResult < 0 ? 'metric-negative' : 'metric-positive'}">${fmtMoney(item.netResult)}</td>
        <td class="mono">${item.marginPct == null ? '—' : fmtPct(item.marginPct)}</td>
      </tr>
    `)
    .join('');
}


function renderLineChart(series) {
  if (!series.length || series.every((point) => Math.abs(point.value) < 0.001)) {
    dom.lineChart.innerHTML = '<div class="chart-empty">За выбранный месяц нет данных для построения дневной динамики.</div>';
    return;
  }

  const width = 920;
  const height = 320;
  const padL = 56;
  const padR = 18;
  const padT = 24;
  const padB = 42;
  const values = series.map((item) => item.value);
  const min = Math.min(0, ...values);
  const max = Math.max(...values, 0);
  const range = max - min || 1;
  const innerW = width - padL - padR;
  const innerH = height - padT - padB;

  const points = series.map((item, index) => {
    const x = padL + (innerW * index) / Math.max(series.length - 1, 1);
    const y = padT + innerH - ((item.value - min) / range) * innerH;
    return { x, y, label: item.date, value: item.value };
  });

  const path = points.map((point, index) => `${index === 0 ? 'M' : 'L'} ${point.x} ${point.y}`).join(' ');
  const baseline = padT + innerH - ((0 - min) / range) * innerH;
  const xLabels = pickSparseLabels(series.map((item) => item.date));
  const yTicks = Array.from({ length: 5 }, (_, index) => min + ((max - min) * index) / 4);

  const svg = `
    <svg class="chart-svg" viewBox="0 0 ${width} ${height}" role="img" aria-label="График дневной выручки">
      <defs>
        <linearGradient id="lineArea" x1="0" y1="0" x2="0" y2="1">
          <stop offset="0%" stop-color="rgba(37,84,255,0.25)" />
          <stop offset="100%" stop-color="rgba(37,84,255,0.02)" />
        </linearGradient>
      </defs>
      ${yTicks.map((tick) => {
        const y = padT + innerH - ((tick - min) / range) * innerH;
        return `<g><line x1="${padL}" y1="${y}" x2="${width - padR}" y2="${y}" stroke="rgba(18,32,51,0.08)" /><text x="${padL - 10}" y="${y + 4}" text-anchor="end" fill="#66758c" font-size="12">${escapeHtml(shortMoneyLabel(tick))}</text></g>`;
      }).join('')}
      <line x1="${padL}" y1="${baseline}" x2="${width - padR}" y2="${baseline}" stroke="rgba(18,32,51,0.14)" />
      <path d="${path} L ${points[points.length - 1].x} ${height - padB} L ${points[0].x} ${height - padB} Z" fill="url(#lineArea)" stroke="none"></path>
      <path d="${path}" fill="none" stroke="#2554ff" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"></path>
      ${points.map((point) => `<circle cx="${point.x}" cy="${point.y}" r="3.5" fill="#2554ff"></circle>`).join('')}
      ${xLabels.map((entry) => `<text x="${entry.x}" y="${height - 14}" text-anchor="middle" fill="#66758c" font-size="12">${escapeHtml(entry.label.slice(8))}</text>`).join('')}
    </svg>
  `;

  dom.lineChart.innerHTML = svg;
}

function renderDonutChart(expenses) {
  const items = [
    ['Комиссия', expenses.commission || 0],
    ['Логистика', expenses.logistics || 0],
    ['Эквайринг / платёжные услуги', expenses.payment || 0],
    ['Хранение', expenses.storage || 0],
    ['Продвижение', expenses.promotion || 0],
    ['Штрафы', expenses.penalties || 0],
    ['Прочее', expenses.otherMarketplace || 0]
  ].filter(([, value]) => value > 0);

  if (!items.length) {
    dom.donutChart.innerHTML = '<div class="chart-empty">В выбранных данных не найдено расходных категорий для круговой диаграммы.</div>';
    return;
  }

  const total = sum(items.map(([, value]) => value));
  const radius = 72;
  const center = 110;
  const circumference = 2 * Math.PI * radius;
  let offset = 0;

  const circles = items.map(([label, value], index) => {
    const segment = (value / total) * circumference;
    const circle = `
      <circle
        cx="${center}"
        cy="${center}"
        r="${radius}"
        fill="none"
        stroke="${CHART_COLORS[index % CHART_COLORS.length]}"
        stroke-width="26"
        stroke-dasharray="${segment} ${circumference - segment}"
        stroke-dashoffset="${-offset}"
        stroke-linecap="butt"
        transform="rotate(-90 ${center} ${center})"
      />
    `;
    offset += segment;
    return circle;
  }).join('');

  const legend = items.map(([label, value], index) => `
    <div class="expense-row">
      <div>
        <strong>${escapeHtml(label)}</strong>
        <span>${fmtPct(value / total)} от расходов</span>
      </div>
      <div class="mono" style="color:${CHART_COLORS[index % CHART_COLORS.length]}">${fmtMoney(value)}</div>
    </div>
  `).join('');

  dom.donutChart.innerHTML = `
    <div style="display:grid; grid-template-columns: 220px 1fr; gap: 16px; align-items:center;">
      <svg class="chart-svg" viewBox="0 0 220 220" role="img" aria-label="Структура расходов">
        <circle cx="${center}" cy="${center}" r="${radius}" fill="none" stroke="rgba(18,32,51,0.08)" stroke-width="26"></circle>
        ${circles}
        <text x="${center}" y="${center - 4}" text-anchor="middle" fill="#66758c" font-size="14">Итого</text>
        <text x="${center}" y="${center + 24}" text-anchor="middle" fill="#122033" font-size="20" font-weight="700">${escapeHtml(shortMoneyLabel(total))}</text>
      </svg>
      <div class="expense-breakdown">${legend}</div>
    </div>
  `;
}

function renderBarChart(products) {
  const top = products.filter((item) => item.revenue > 0).slice(0, 8);
  if (!top.length) {
    dom.barChart.innerHTML = '<div class="chart-empty">Недостаточно данных для топа товаров.</div>';
    return;
  }

  const max = Math.max(...top.map((item) => item.revenue), 1);
  dom.barChart.innerHTML = `
    <div class="expense-breakdown">
      ${top.map((item, index) => {
        const width = (item.revenue / max) * 100;
        return `
          <div>
            <div style="display:flex; justify-content:space-between; gap:12px; margin-bottom:6px; align-items:center;">
              <strong>${escapeHtml(trimText(item.name, 46))}</strong>
              <span class="mono">${fmtMoney(item.revenue)}</span>
            </div>
            <div style="height:12px; border-radius:999px; background:rgba(18,32,51,0.08); overflow:hidden;">
              <div style="height:100%; width:${width}%; background:${CHART_COLORS[index % CHART_COLORS.length]}; border-radius:999px;"></div>
            </div>
          </div>
        `;
      }).join('')}
    </div>
  `;
}

function renderLimitations() {
  const defaults = [
    'Приложение не придумывает метрики: если нужных полей в загруженных файлах нет, показатель помечается как недоступный или частичный.',
    'Файлы не отправляются на сервер. Обработка идёт в браузере, а настройки ручных затрат сохраняются только в localStorage.',
    'При изменении структуры официальных отчётов может потребоваться обновить словари колонок в script.js.'
  ];

  const current = state.lastAnalysis?.marketplace === state.marketplace ? state.lastAnalysis.limitations : [];
  const items = [...defaults, ...current].slice(0, 7);
  dom.limitationsBlock.innerHTML = `<ul>${items.map((item) => `<li>${escapeHtml(item)}</li>`).join('')}</ul>`;
}

function createSummaryText(analysis) {
  const monthLabel = formatMonthRu(analysis.month);
  const top = analysis.productSummary.slice(0, 5);
  const weak = analysis.productSummary.filter((item) => item.netResult < 0).slice(0, 5);
  return [
    `${analysis.title} • ${monthLabel}`,
    '',
    'Ключевые показатели:',
    ...analysis.kpis.map((item) => `- ${item.label}: ${formatKpiValue(item.value, item.kind)} (${item.stateLabel})`),
    '',
    'Топ товаров:',
    ...(top.length ? top.map((item) => `- ${item.sku || '—'} • ${item.name} • выручка ${fmtMoney(item.revenue)} • итог ${fmtMoney(item.netResult)}`) : ['- Нет данных']),
    '',
    'Слабые товары:',
    ...(weak.length ? weak.map((item) => `- ${item.sku || '—'} • ${item.name} • итог ${fmtMoney(item.netResult)}`) : ['- Явно убыточных позиций не найдено']),
    '',
    'Выводы:',
    ...analysis.insights.map((item) => `- ${item}`)
  ].join('\n');
}


function exportProductsCsv(analysis) {
  const rows = analysis.productSummary.map((item) => {
    if (analysis.marketplace === 'ozon') {
      return {
        SKU: item.sku || '',
        Артикул: item.article || '',
        Товар: item.name || '',
        Выручка: round2(item.revenue),
        'К начислению': round2(item.payout),
        'Выкупы, шт': item.unitsSold || 0,
        'Финансовые возвраты, шт': item.unitsReturned || 0,
        'Операционные возвраты, шт': item.operationalReturns || 0,
        'Комиссия Ozon': round2(item.commission || 0),
        Эквайринг: round2(item.payment || 0),
        Логистика: round2(item.logistics || 0),
        Продвижение: round2(item.promotion || 0),
        'Прочие расходы MP': round2(item.otherMarketplace || 0),
        'Ручные затраты': round2(item.manualCosts || 0),
        Итог: round2(item.netResult || 0),
        'Маржа, %': item.marginPct == null ? '' : round2(item.marginPct * 100)
      };
    }

    return {
      SKU: item.sku || '',
      Товар: item.name || '',
      Выручка: round2(item.revenue),
      'Выкупы, шт': item.unitsSold || 0,
      'Возвраты, шт': item.unitsReturned || 0,
      'Комиссия и платежи': round2((item.commission || 0) + (item.payment || 0)),
      Логистика: round2(item.logistics),
      'Прочие расходы MP': round2(item.otherMarketplace),
      'Ручные затраты': round2(item.manualCosts),
      Итог: round2(item.netResult),
      'Маржа, %': item.marginPct == null ? '' : round2(item.marginPct * 100)
    };
  });

  const csv = objectsToCsv(rows);
  const monthSafe = analysis.month.replace('-', '_');
  downloadBlob(`${analysis.marketplace}_${monthSafe}_товары.csv`, csv, 'text/csv;charset=utf-8;');
}


function createResolver(headers) {
  const normalized = headers.map((header) => ({ raw: header, normalized: normalizeHeader(header) }));

  function find(aliases) {
    const list = Array.isArray(aliases) ? aliases : [aliases];
    for (const alias of list) {
      const aliasNorm = normalizeHeader(alias);
      const exact = normalized.find((entry) => entry.normalized === aliasNorm);
      if (exact) return exact.raw;
      const partial = normalized.find((entry) => entry.normalized.includes(aliasNorm));
      if (partial) return partial.raw;
      const tokenMatch = normalized.find((entry) => {
        const tokens = aliasNorm.split(' ').filter(Boolean);
        return tokens.length > 1 && tokens.every((token) => entry.normalized.includes(token));
      });
      if (tokenMatch) return tokenMatch.raw;
    }
    return null;
  }

  return { find };
}

function getCell(row, columnName) {
  if (!columnName) return '';
  return row[columnName] ?? '';
}

function simpleCsvParse(text) {
  const cleanText = String(text ?? '').replace(/^\uFEFF/, '');
  const separator = detectCsvSeparator(cleanText);
  const lines = cleanText.replace(/\r\n/g, '\n').replace(/\r/g, '\n').split('\n');
  return lines.filter((line, index, arr) => line !== '' || index < arr.length - 1).map((line) => parseCsvLine(line, separator));
}

function parseCsvLine(line, separator) {
  const result = [];
  let current = '';
  let inQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];
    if (char === '"') {
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }
    if (char === separator && !inQuotes) {
      result.push(current);
      current = '';
      continue;
    }
    current += char;
  }
  result.push(current);
  return result;
}

function detectCsvSeparator(text) {
  const sample = String(text ?? '').replace(/^\uFEFF/, '').split(/\r?\n/).slice(0, 6).join('\n');
  const candidates = [',', ';', '\t'];
  let best = ';';
  let max = -1;
  candidates.forEach((candidate) => {
    const count = (sample.match(new RegExp(escapeRegExp(candidate), 'g')) || []).length;
    if (count > max) {
      max = count;
      best = candidate;
    }
  });
  return best;
}

function parseGenericXmlRows(text) {
  const parser = new DOMParser();
  const doc = parser.parseFromString(text, 'application/xml');
  if (doc.querySelector('parsererror')) {
    throw new Error('XML-файл повреждён или не читается браузером.');
  }

  const nodes = Array.from(doc.getElementsByTagName('*')).filter((node) => {
    const elementChildren = Array.from(node.children || []);
    return elementChildren.length > 1 && elementChildren.every((child) => child.children.length === 0);
  });

  if (!nodes.length) {
    throw new Error('В XML не удалось найти повторяющиеся табличные записи.');
  }

  const groups = new Map();
  nodes.forEach((node) => {
    const key = node.tagName;
    groups.set(key, (groups.get(key) || 0) + 1);
  });

  const bestTag = Array.from(groups.entries()).sort((a, b) => b[1] - a[1])[0][0];
  const rows = Array.from(doc.getElementsByTagName(bestTag)).map((node) => {
    const obj = {};
    Array.from(node.children).forEach((child) => {
      obj[child.tagName] = child.textContent || '';
    });
    return obj;
  });

  return rows;
}

function isDateInSelectedMonth(value, month) {
  const date = parseDateValue(value);
  if (!date || !month) return false;
  return toMonthKey(date) === month;
}

function getMonthDateList(month) {
  const [year, mon] = month.split('-').map(Number);
  const date = new Date(year, mon - 1, 1);
  const list = [];
  while (date.getMonth() === mon - 1) {
    list.push(toDateKey(date));
    date.setDate(date.getDate() + 1);
  }
  return list;
}

function toMonthKey(date) {
  const d = parseDateValue(date);
  if (!d) return '';
  const month = String(d.getMonth() + 1).padStart(2, '0');
  return `${d.getFullYear()}-${month}`;
}

function toDateKey(date) {
  const d = parseDateValue(date);
  if (!d) return '';
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
}

function parseDateValue(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (value == null || value === '') return null;

  if (typeof value === 'number') {
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const date = new Date(excelEpoch.getTime() + value * 86400000);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const text = String(value).trim();
  if (!text) return null;

  const iso = text.match(/^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})/);
  if (iso) {
    const date = new Date(Number(iso[1]), Number(iso[2]) - 1, Number(iso[3]));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const ru = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})/);
  if (ru) {
    const year = ru[3].length === 2 ? Number(`20${ru[3]}`) : Number(ru[3]);
    const date = new Date(year, Number(ru[2]) - 1, Number(ru[1]));
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const native = new Date(text);
  return Number.isNaN(native.getTime()) ? null : native;
}

function parseNumber(value) {
  if (value == null || value === '') return null;
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  const text = String(value)
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, '')
    .replace(/−/g, '-')
    .replace(/,/g, '.');
  if (!text || text === '-') return null;
  const number = Number(text.replace(/[^0-9+-.]/g, ''));
  return Number.isFinite(number) ? number : null;
}

function formatKpiValue(value, kind) {
  if (value == null || value === '' || Number.isNaN(value)) return '—';
  if (kind === 'money') return fmtMoney(value);
  if (kind === 'number') return nfNumber.format(value);
  return String(value);
}

function fmtMoney(value) {
  if (value == null || Number.isNaN(value)) return '—';
  return nfMoney.format(value);
}

function fmtPct(value) {
  if (value == null || Number.isNaN(value)) return '—';
  return `${nfDecimal.format(value * 100)}%`;
}

function shortMoneyLabel(value) {
  const abs = Math.abs(value);
  if (abs >= 1000000) return `${nfDecimal.format(value / 1000000)} млн`;
  if (abs >= 1000) return `${nfDecimal.format(value / 1000)} тыс`;
  return nfNumber.format(round2(value));
}

function formatMonthRu(month) {
  const [year, mon] = month.split('-').map(Number);
  const date = new Date(year, mon - 1, 1);
  return new Intl.DateTimeFormat('ru-RU', { month: 'long', year: 'numeric' }).format(date);
}

function getFileExtension(name) {
  return String(name).split('.').pop().toLowerCase();
}

function getFileId(file) {
  return `${file.name}::${file.size}::${file.lastModified}`;
}

function positiveOnly(value) {
  return value > 0 ? value : 0;
}

function sum(values) {
  return values.reduce((acc, value) => acc + (Number(value) || 0), 0);
}

function uniqueBy(items, getKey) {
  const map = new Map();
  items.forEach((item) => {
    const key = getKey(item);
    if (key == null || key === '') return;
    if (!map.has(key)) map.set(key, item);
  });
  return Array.from(map.values());
}

function makeUniqueHeaders(headers) {
  const used = new Map();
  return headers.map((header, index) => {
    const base = String(header ?? '').trim() || `Колонка ${index + 1}`;
    const count = used.get(base) || 0;
    used.set(base, count + 1);
    return count ? `${base} (${count + 1})` : base;
  });
}

function normalizeHeader(value) {
  return String(value ?? '')
    .toLowerCase()
    .replace(/ё/g, 'е')
    .replace(/\u00a0/g, ' ')
    .replace(/[«»"'`]/g, '')
    .replace(/[()\[\]{}:;,.!?%№/\\]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function safeString(value) {
  return String(value ?? '').trim();
}

function isBlank(value) {
  return value == null || String(value).trim() === '';
}

function isRowEmpty(row) {
  return !row || row.every((value) => isBlank(value));
}

function objectsToCsv(rows) {
  if (!rows.length) return '';
  const headers = Object.keys(rows[0]);
  const escape = (value) => {
    const text = value == null ? '' : String(value);
    if (/[";,\n]/.test(text)) return `"${text.replace(/"/g, '""')}"`;
    return text;
  };
  return [headers.join(';'), ...rows.map((row) => headers.map((header) => escape(row[header])).join(';'))].join('\n');
}

function downloadBlob(fileName, content, mimeType) {
  const blob = new Blob([content], { type: mimeType });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

function escapeHtml(value) {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeRegExp(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function pickSparseLabels(labels) {
  const result = [];
  const max = 6;
  const step = Math.max(1, Math.floor(labels.length / max));
  labels.forEach((label, index) => {
    if (index === 0 || index === labels.length - 1 || index % step === 0) {
      result.push({ label, index });
    }
  });
  return result.map((item) => ({ ...item, x: 56 + ((920 - 56 - 18) * item.index) / Math.max(labels.length - 1, 1) }));
}

function round2(value) {
  return Math.round((value || 0) * 100) / 100;
}

function trimText(value, length) {
  const text = String(value ?? '');
  return text.length > length ? `${text.slice(0, length - 1)}…` : text;
}

function showToast(message, type = 'success', timeout = 3600) {
  const toast = document.createElement('div');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  dom.toastContainer.appendChild(toast);
  window.setTimeout(() => {
    toast.remove();
  }, timeout);
}
