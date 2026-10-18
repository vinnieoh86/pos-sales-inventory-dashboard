(function () {
  const VERSION = "20261018-sales-studio-v15";
  const DAY_MS = 86400000;
  const GKPOS_DEPARTMENT_ORDER = ["GENERAL", "HAIR", "BEAUTY PRODUCTS", "COSMETIC", "ELECTRIC", "JEWELRY", "GROCERY"];
  const GKPOS_CATEGORY_DEPARTMENTS = new Map(Object.entries({
    "BODY OIL": "BEAUTY PRODUCTS", "CHEMICAL": "BEAUTY PRODUCTS", "SKIN CARE": "BEAUTY PRODUCTS",
    "BRONZER": "COSMETIC", "CONCEALER": "COSMETIC", "COSMETIC ETC": "COSMETIC", "DEODORANT": "COSMETIC",
    "EYELASHES": "COSMETIC", "EYELINER": "COSMETIC", "EYESHADOW": "COSMETIC", "LIPGLOSS": "COSMETIC",
    "LIPLINER": "COSMETIC", "LIPSTICK": "COSMETIC", "LIQUID MAKE-UP": "COSMETIC",
    "MAKE UP SPONGE / BRUSH": "COSMETIC", "MASCARA": "COSMETIC", "NAIL CARE": "COSMETIC",
    "NAIL POLISH": "COSMETIC", "POWDER MAKE UP": "COSMETIC",
    "CLIPPER": "ELECTRIC", "CURLING IRON": "ELECTRIC", "DRYER": "ELECTRIC", "ELECTRIC ETC": "ELECTRIC",
    "ELECTRIC HAIR ROLLER": "ELECTRIC", "FLAT IRON": "ELECTRIC", "HEADPHONES": "ELECTRIC",
    "STOVE": "ELECTRIC", "TRIMMER": "ELECTRIC",
    "BAGS": "GENERAL", "BARBER ITEMS": "GENERAL", "CLOTHES": "GENERAL", "COMBS / BRUSHES": "GENERAL",
    "GENERAL": "GENERAL", "HAIR ACCESSORIES": "GENERAL", "SCARF & BAND": "GENERAL", "SHOES": "GENERAL",
    "SUNGLASS/GLASSES": "GENERAL", "WIG CAPS": "GENERAL",
    "CANDY": "GROCERY", "DRINKS": "GROCERY", "FOOD": "GROCERY", "FROZEN": "GROCERY",
    "HOUSE HOLD": "GROCERY", "MEDICINE": "GROCERY", "TOBACCO": "GROCERY",
    "BRAIDING": "HAIR", "BUNDLES": "HAIR", "CLOSURE": "HAIR", "CROCHET": "HAIR", "EXTENSION": "HAIR",
    "HAIR BUN": "HAIR", "HALF BUNDLE": "HAIR", "HUMAN HAIR": "HAIR", "HUMAN WIG": "HAIR",
    "PONYTAIL": "HAIR", "SYNTH. HAIR": "HAIR", "WIGS": "HAIR",
    "BODY JEWELRY": "JEWELRY", "BRACELET / ANKLET": "JEWELRY", "EARRING": "JEWELRY",
    "JEWELRY ETC": "JEWELRY", "NECKLACE": "JEWELRY", "NOSE RINGS": "JEWELRY", "RINGS / TOE RINGS": "JEWELRY"
  }));

  state._remoteSalesLoadedStart = state._remoteSalesLoadedStart || "";
  state._remoteSalesLoadedEnd = state._remoteSalesLoadedEnd || "";
  state._remoteSalesAvailableStart = state._remoteSalesAvailableStart || "";
  state._remoteSalesAvailableEnd = state._remoteSalesAvailableEnd || "";
  state._remoteSalesLoadPromise = null;
  state.activePresetDays = 30;

  const style = document.createElement("link");
  style.rel = "stylesheet";
  style.href = `sales-dashboard-patch.css?v=${VERSION}`;
  document.head.append(style);

  const dashboard = document.querySelector('[data-tab-view="dashboard"]');
  if (dashboard && !document.querySelector("#dashboardOverview")) {
    dashboard.insertAdjacentHTML("afterbegin", `
      <section class="dashboard-overview" id="dashboardOverview" aria-label="Sales dashboard status">
        <div>
          <p class="eyebrow">Sales intelligence</p>
          <h1>Performance at a glance</h1>
          <p class="dashboard-overview__copy">Compare revenue, units, profit, departments, categories, vendors, and products across any period.</p>
        </div>
        <div class="dashboard-overview__status">
          <span class="dashboard-sync-badge" id="dashboardDataStatus">Loading sales history...</span>
          <small id="dashboardDataRange">Waiting for shared data.</small>
        </div>
      </section>
      <nav class="sales-mode-tabs" id="salesModeTabs" aria-label="Sales dashboard views">
        <button type="button" data-sales-mode="overview">Overview</button>
        <button type="button" class="active" data-sales-mode="calendar">Calendar</button>
        <button type="button" data-sales-mode="items">Items</button>
      </nav>
      <section class="sales-health-panel" id="salesDataHealth" aria-live="polite">
        <div class="sales-health-heading">
          <div><p class="eyebrow">Data health</p><h2>Sync &amp; recovery</h2></div>
          <button type="button" class="sales-secondary-button" id="salesHealthRefresh">Repair &amp; refresh</button>
        </div>
        <div class="sales-health-grid" id="salesHealthGrid"><div class="sales-loading">Checking sales and inventory sync...</div></div>
        <div class="sales-health-alerts" id="salesHealthAlerts"></div>
      </section>
      <section class="sales-workspace" id="salesWorkspace">
        <section class="sales-mode-view" data-sales-mode-view="overview" hidden>
          <div class="sales-section-heading">
            <div><p class="eyebrow">Executive analysis</p><h2>Business pulse</h2></div>
            <small id="salesOverviewFocus">Current period compared with the same period last year.</small>
          </div>
          <div class="trend-explorer-controls sales-overview-controls">
            <label>Period<select id="salesTrendPeriod"><option value="day">Day</option><option value="month" selected>Month</option><option value="quarter">Quarter</option><option value="half">Half-year</option><option value="year">Year</option></select></label>
            <label>Period date<input type="date" id="salesTrendDate" /></label>
            <label>Analyze<select id="salesTrendGroup"><option value="" selected>Total sales</option><option value="department">Department</option><option value="category">Category</option><option value="vendor">Vendor</option><option value="item">Item</option></select></label>
            <label>Selection<input id="salesTrendEntity" list="salesTrendEntities" placeholder="All sales" disabled /><datalist id="salesTrendEntities"></datalist></label>
            <label>Metric<select id="salesTrendMetric"><option value="sales">Net sales $</option><option value="units">Units</option><option value="profit">Profit</option><option value="margin">Margin %</option></select></label>
            <div class="sales-segmented" id="salesChartType"><button type="button" class="active" data-chart-type="line">Line</button><button type="button" data-chart-type="bar">Bar</button></div>
            <label class="sales-check"><input type="checkbox" id="salesTrendCompare" checked /> Compare prior year</label>
            <button type="button" class="sales-secondary-button sales-clear-button" id="salesTrendClear">Clear all</button>
          </div>
          <article class="sales-card sales-trend-card" id="salesTrendCard">
            <div class="sales-card__heading"><div><p class="eyebrow" id="salesTrendKicker">Total sales</p><h3 id="salesTrendTitle">Monthly performance</h3></div><div class="sales-chart-actions"><div class="sales-chart-legend" id="salesTrendLegend"></div><button type="button" class="sales-secondary-button" id="salesTrendExpand">Expand</button></div></div>
            <canvas id="salesTrendExplorerChart" height="320"></canvas>
            <div class="sales-chart-summary" id="salesTrendSummary"></div>
          </article>
          <div class="yoy-card-grid" id="salesYoyCards"></div>
          <div class="sales-executive-grid" id="salesExecutiveSummary"></div>
          <div class="sales-overview-grid sales-overview-grid--single">
            <article class="sales-card">
              <div class="sales-card__heading"><div><p class="eyebrow">Sales mix</p><h3>Departments &amp; categories</h3></div></div>
              <div class="sales-department-tree" id="salesOverviewCategories"></div>
            </article>
          </div>
        </section>

        <section class="sales-mode-view active" data-sales-mode-view="calendar">
          <div class="calendar-toolbar">
            <button type="button" class="sales-icon-button" id="salesCalendarPrev" aria-label="Previous month">&#8249;</button>
            <div><p class="eyebrow">Monthly sales</p><h2 id="salesCalendarTitle">Calendar</h2><div class="calendar-month-summary" id="salesCalendarMonthSummary"></div></div>
            <button type="button" class="sales-icon-button" id="salesCalendarNext" aria-label="Next month">&#8250;</button>
            <label class="calendar-date-jump">Jump to date<input type="date" id="salesCalendarDatePicker" /></label>
            <button type="button" class="sales-secondary-button" id="salesCalendarToday">Latest month</button>
          </div>
          <div class="sales-calendar" id="salesCalendar"></div>
        </section>

        <section class="sales-mode-view" data-sales-mode-view="items" hidden>
          <div class="sales-section-heading">
            <div><p class="eyebrow">Itemized sales</p><h2>Products for the selected dates</h2></div>
            <small id="salesItemsRange">Uses the current dashboard date range.</small>
          </div>
          <div class="sales-item-toolbar">
            <label>From<input type="date" id="salesItemsStart" /></label>
            <label>To<input type="date" id="salesItemsEnd" /></label>
            <label>Department<select id="salesItemDepartment"><option value="">All departments</option></select></label>
            <label>Category<select id="salesItemCategory"><option value="">All categories</option></select></label>
            <label>Vendor<select id="salesItemVendor"><option value="">All vendors</option></select></label>
            <input type="search" id="salesItemSearch" placeholder="Search product or UPC" />
            <select id="salesItemSort" aria-label="Sort items"><option value="sales">Highest sales</option><option value="units">Most units</option><option value="profit">Highest profit</option><option value="cost">Highest total cost</option><option value="product">Product A-Z</option></select>
            <button type="button" class="sales-secondary-button sales-clear-button" id="salesItemsClear">Clear all</button>
          </div>
          <div class="sales-item-totals" id="salesItemsTotals"></div>
          <div class="sales-table-wrap"><table class="sales-item-table"><thead><tr>
            <th><button type="button" data-item-sort="product">Product</button></th>
            <th><button type="button" data-item-sort="department">Department</button></th>
            <th><button type="button" data-item-sort="category">Category</button></th>
            <th><button type="button" data-item-sort="vendor">Vendor</button></th>
            <th><button type="button" data-item-sort="units">Units</button></th>
            <th><button type="button" data-item-sort="price">Avg price</button></th>
            <th><button type="button" data-item-sort="sales">Total sales</button></th>
            <th><button type="button" data-item-sort="unitCost">Unit cost</button></th>
            <th><button type="button" data-item-sort="cost">Total cost</button></th>
            <th><button type="button" data-item-sort="profit">Total profit</button></th>
          </tr></thead><tbody id="salesItemRows"></tbody></table></div>
        </section>
      </section>`);
  }

  function shiftIsoDate(iso, days) {
    const date = new Date(`${iso}T12:00:00Z`);
    if (!Number.isFinite(date.getTime())) return "";
    date.setUTCDate(date.getUTCDate() + Number(days || 0));
    return date.toISOString().slice(0, 10);
  }

  function salesDisplayDate(iso) {
    if (!iso || !/^\d{4}-\d{2}-\d{2}$/.test(String(iso).slice(0, 10))) return cleanCell(iso);
    const [year, month, day] = String(iso).slice(0, 10).split("-");
    return `${month}/${day}/${year}`;
  }

  function dashboardDataStatus() {
    const status = document.querySelector("#dashboardDataStatus");
    const range = document.querySelector("#dashboardDataRange");
    if (!status || !range || state._remoteSalesLoadPromise) return;
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    const earliest = state._remoteSalesAvailableStart || state.dates[0] || "";
    if (!latest) {
      status.textContent = "No sales data loaded";
      range.textContent = "The nightly sync will appear here after its first successful upload.";
      return;
    }
    status.textContent = `Synced through ${salesDisplayDate(latest)}`;
    range.textContent = earliest
      ? `History available ${salesDisplayDate(earliest)} through ${salesDisplayDate(latest)}. Selected periods load on demand.`
      : `Latest completed sales date: ${salesDisplayDate(latest)}.`;
  }

  function localIsoDate(date = new Date()) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, "0");
    const day = String(date.getDate()).padStart(2, "0");
    return `${year}-${month}-${day}`;
  }

  function expectedLatestSalesDate() {
    const now = new Date();
    const closeReached = now.getHours() > 20 || (now.getHours() === 20 && now.getMinutes() >= 15);
    let date = new Date(now.getFullYear(), now.getMonth(), now.getDate() - (closeReached ? 0 : 1), 12);
    while (date.getDay() === 0) date = new Date(date.getFullYear(), date.getMonth(), date.getDate() - 1, 12);
    return localIsoDate(date);
  }

  function formatHealthStamp(value) {
    if (!value) return "Not reported";
    const date = new Date(value);
    if (!Number.isFinite(date.getTime())) return String(value);
    return date.toLocaleString("en-US", { month: "short", day: "numeric", hour: "numeric", minute: "2-digit" });
  }

  async function renderSalesDataHealth(force = false) {
    const target = document.querySelector("#salesHealthGrid");
    const alerts = document.querySelector("#salesHealthAlerts");
    if (!target || !alerts) return;
    if (!force && salesUi.healthLoadedAt && Date.now() - salesUi.healthLoadedAt < 60000) return;
    if (salesUi.healthPromise) return salesUi.healthPromise;
    target.innerHTML = `<div class="sales-loading">Checking sales and inventory sync...</div>`;
    salesUi.healthPromise = (async () => {
      const expectedEnd = expectedLatestSalesDate();
      const scanStart = shiftIsoDate(expectedEnd, -29);
      const [syncRows, newestProducts, salesDays] = await Promise.all([
        directSelectRows("sync_state", { select: "id,last_sync_kind,updated_at,latest_inventory_date,latest_sales_date", id: "eq.pos-dashboard-main", limit: "1" }).catch(() => []),
        directSelectRows("products", { select: "updated_at", order: "updated_at.desc", limit: "1" }).catch(() => []),
        salesFetchRollups("DAY", scanStart, expectedEnd),
      ]);
      const sync = syncRows[0] || {};
      const availableDates = new Set(salesDays.map((row) => row.date));
      const missing = [];
      for (let iso = scanStart; iso <= expectedEnd; iso = shiftIsoDate(iso, 1)) {
        if (salesDate(iso).getUTCDay() !== 0 && !availableDates.has(iso)) missing.push(iso);
      }
      const latestSales = [cleanCell(sync.latest_sales_date), ...availableDates].filter(Boolean).sort().at(-1) || "";
      const inventoryDate = cleanCell(sync.latest_inventory_date);
      const productUpdatedAt = newestProducts[0]?.updated_at || "";
      const salesState = missing.length ? "warning" : "ok";
      const inventoryAge = inventoryDate ? Math.floor((salesDate(localIsoDate()) - salesDate(inventoryDate.slice(0, 10))) / DAY_MS) : 999;
      const inventoryState = inventoryAge > 1 ? "warning" : "ok";
      const productAgeHours = productUpdatedAt ? (Date.now() - Date.parse(productUpdatedAt)) / 3600000 : 999;
      const deltaState = productAgeHours > 24 ? "warning" : "ok";
      target.innerHTML = `
        <article class="sales-health-card ${salesState}"><small>Sales</small><b>${latestSales ? `Through ${escapeHtml(salesDisplayDate(latestSales))}` : "No successful upload"}</b><span>${missing.length ? `${missing.length} missing business day${missing.length === 1 ? "" : "s"}` : "Last 30 days complete"}</span></article>
        <article class="sales-health-card ${inventoryState}" title="A full inventory snapshot is the complete product and quantity file. The recovery watchdog repairs an overdue snapshot during business hours."><small>Full inventory</small><b>${inventoryDate ? escapeHtml(salesDisplayDate(inventoryDate)) : "Not reported"}</b><span>${inventoryAge > 1 ? `Full snapshot overdue by ${number.format(inventoryAge)} days` : "Current full snapshot available"}</span></article>
        <article class="sales-health-card ${deltaState}"><small>Inventory updates</small><b>${escapeHtml(formatHealthStamp(productUpdatedAt))}</b><span>${deltaState === "ok" ? "Recent product sync" : "No recent delta upload"}</span></article>`;
      if (missing.length) {
        const preview = missing.slice(-8).map(salesDisplayDate).join(", ");
        alerts.className = "sales-health-alerts warning";
        alerts.textContent = sync.last_sync_kind === "recovery-requested"
          ? `Recovery queued for ${missing.length} sales date${missing.length === 1 ? "" : "s"}: ${preview}${missing.length > 8 ? " …" : ""}`
          : `Recovery needed for ${missing.length} sales date${missing.length === 1 ? "" : "s"}: ${preview}${missing.length > 8 ? " …" : ""}`;
      } else {
        alerts.className = "sales-health-alerts";
        alerts.textContent = "No missing sales dates detected in the last 30 business days.";
      }
      salesUi.healthLoadedAt = Date.now();
    })().catch((error) => {
      target.innerHTML = `<article class="sales-health-card error"><small>Data health</small><b>Status check failed</b><span>${escapeHtml(error?.message || "Unable to read sync status")}</span></article>`;
      alerts.textContent = "PowerShell recovery will continue independently.";
    }).finally(() => { salesUi.healthPromise = null; });
    return salesUi.healthPromise;
  }

  async function requestSalesRecovery() {
    const button = document.querySelector("#salesHealthRefresh");
    const alerts = document.querySelector("#salesHealthAlerts");
    if (button) { button.disabled = true; button.textContent = "Queuing repair…"; }
    try {
      await supabaseUpsertRows("sync_state", [{
        id: SYNC_STATE_ROW_ID,
        last_sync_kind: "recovery-requested",
        updated_at: new Date().toISOString(),
      }], "id");
      if (alerts) {
        alerts.className = "sales-health-alerts warning";
        alerts.textContent = "Recovery queued. The store sync computer will repair missing dates on its next 30-minute watchdog check.";
      }
      showToast("Recovery requested. Checking cloud data again…", 3200, "success");
    } catch (error) {
      if (alerts) {
        alerts.className = "sales-health-alerts error";
        alerts.textContent = "The recovery request could not be queued. The automatic watchdog will still retry.";
      }
      showToast("Recovery request could not be queued.", 3200, "warning");
    } finally {
      salesUi.healthLoadedAt = 0;
      salesUi.rollups.clear();
      await renderSalesDataHealth(true);
      if (button) { button.disabled = false; button.textContent = "Repair & refresh"; }
    }
  }

  const originalRender = render;
  render = function patchedRender() {
    const result = originalRender.apply(this, arguments);
    dashboardDataStatus();
    renderSalesWorkspace();
    return result;
  };

  const originalNormalizeSalesRow = normalizeSalesRow;
  normalizeSalesRow = function patchedNormalizeSalesRow(row, date) {
    const patched = {
      ...row,
      PRODUCT: row.PRODUCT || row.Product || row.item_name,
      VENDOR: row.VENDOR || row.Vendor || row.vendor_name,
      QTY: row.QTY || row.Qty || row.Quantity,
      SBTLS: row.SBTLS || row.Subtotal || row.Sales,
      COST: row.COST || row.Cost,
      PRF: row.PRF || row.Profit,
      CATEGORY: row.CATEGORY || row.Category,
      DEPARTMENT: row.DEPARTMENT || row.Department,
      CODE: row.CODE || row.Code,
    };
    return originalNormalizeSalesRow(patched, date);
  };

  const originalDateFromRows = dateFromRows;
  dateFromRows = function patchedDateFromRows(rows) {
    const sample = (rows || []).find((row) => row && typeof row === "object");
    if (sample) {
      const salesDate = cleanCell(sample.SaleDate || sample.saleDate || sample.sales_date || sample["Sale Date"]);
      const iso = normalizeItemDate(salesDate);
      if (iso) {
        const [yyyy, mm, dd] = iso.split("-");
        return { iso, compact: `${mm}${dd}${yyyy}` };
      }
    }
    return originalDateFromRows(rows);
  };

  const directSelectRows = supabaseSelectRows;

  async function remoteSalesBounds() {
    const [firstRows, lastRows] = await Promise.all([
      directSelectRows("daily_sales", { select: "sales_date", order: "sales_date.asc", limit: 1 }).catch(() => []),
      directSelectRows("daily_sales", { select: "sales_date", order: "sales_date.desc", limit: 1 }).catch(() => []),
    ]);
    const start = cleanCell(firstRows[0]?.sales_date);
    const end = cleanCell(lastRows[0]?.sales_date);
    if (start) state._remoteSalesAvailableStart = start;
    if (end) state._remoteSalesAvailableEnd = end;
    return { start, end };
  }

  function remoteSalesRange(start, end) {
    if (!start || !end || start > end) return Promise.resolve([]);
    return directSelectRows("daily_sales", {
      select: "*",
      and: `(sales_date.gte.${start},sales_date.lte.${end})`,
      order: "sales_date.asc",
    }).then((rows) => rows.filter((row) => !String(row.code || "").startsWith("__ROLLUP__")));
  }

  const originalRestoreShared = restoreSharedDataFromSupabase;
  restoreSharedDataFromSupabase = async function patchedRestoreShared(options = {}) {
    const bounds = await remoteSalesBounds();
    const end = bounds.end || shiftIsoDate(new Date().toISOString().slice(0, 10), -1);
    const start = shiftIsoDate(end, -64);
    const priorSelect = supabaseSelectRows;
    supabaseSelectRows = function rangeAwareSelect(tableName, query = {}) {
      if (tableName === "daily_sales") return remoteSalesRange(start, end);
      return priorSelect(tableName, query);
    };
    try {
      const result = await originalRestoreShared(options);
      if (result && result !== "kept-local") {
        state._remoteSalesLoadedStart = start;
        state._remoteSalesLoadedEnd = end;
      }
      return result;
    } finally {
      supabaseSelectRows = priorSelect;
      dashboardDataStatus();
    }
  };

  async function ensureRemoteSalesRange(start, end, options = {}) {
    if (!start || !end || start > end) return false;
    const loadedStart = state._remoteSalesLoadedStart;
    const loadedEnd = state._remoteSalesLoadedEnd;
    if (!options.force && loadedStart && loadedEnd && start >= loadedStart && end <= loadedEnd) return false;
    if (state._remoteSalesLoadPromise) return state._remoteSalesLoadPromise;

    state._remoteSalesLoadPromise = (async () => {
      const status = document.querySelector("#dashboardDataStatus");
      if (status) status.textContent = `Loading sales ${start} through ${end}...`;
      try {
        const rows = await remoteSalesRange(start, end);
        const hydrated = rows.map(hydrateSalesFromSupabase).filter((row) => row.code || row.product);
        state.rawSales = state.rawSales.filter((row) => row.date < start || row.date > end).concat(hydrated);
        state.rawSales.sort((a, b) => a.date.localeCompare(b.date));
        state.dates = [...new Set(state.rawSales.map((row) => row.date))].sort();
        state._remoteSalesLoadedStart = loadedStart && loadedStart < start ? loadedStart : start;
        state._remoteSalesLoadedEnd = loadedEnd && loadedEnd > end ? loadedEnd : end;
        bumpDataStamp();
        updateFilterOptions();
        render();
        return true;
      } catch (error) {
        console.warn("Requested sales range could not be loaded", error);
        if (!options.silent) showToast("Sales history could not be loaded for that period.", 3200, "warning");
        return false;
      } finally {
        state._remoteSalesLoadPromise = null;
        dashboardDataStatus();
      }
    })();
    return state._remoteSalesLoadPromise;
  }

  function loadSelectedSalesRange() {
    const start = els.startDate.value;
    const end = els.endDate.value;
    if (!start || !end) return Promise.resolve(false);
    const dayCount = Math.max(1, Math.round((new Date(`${end}T00:00:00`) - new Date(`${start}T00:00:00`)) / DAY_MS) + 1);
    return ensureRemoteSalesRange(shiftIsoDate(start, -dayCount) || start, end);
  }

  const originalApplyDatePreset = applyDatePreset;
  let firstPreset = true;
  applyDatePreset = function patchedApplyDatePreset(days) {
    const selected = firstPreset && Number(days) === 90 ? 30 : days;
    firstPreset = false;
    state.activePresetDays = selected;
    const result = originalApplyDatePreset(selected);
    if (selected === "all") {
      els.startDate.value = state._remoteSalesAvailableStart || els.startDate.value;
      els.endDate.value = state._remoteSalesAvailableEnd || els.endDate.value;
      render();
      void ensureRemoteSalesRange(els.startDate.value, els.endDate.value);
    } else {
      void loadSelectedSalesRange();
    }
    return result;
  };

  let importContext = null;
  const originalLoadFiles = loadFiles;
  loadFiles = async function patchedLoadFiles(fileList) {
    const files = [...fileList].filter((file) => /\.csv$/i.test(file.name));
    const salesDates = new Set();
    for (const file of files) {
      const rows = parseCsv(await file.text());
      const isInventory = /(^|[^a-z])(current[_ -]?inventory|all[_ -]?inventory|inventory[_ -]?snapshot)/i.test(file.name)
        || looksLikeInventoryRows(rows);
      if (isInventory) continue;
      const date = dateFromFileName(file.name) || dateFromRows(rows);
      if (date?.iso) salesDates.add(date.iso);
    }
    importContext = { salesDates };
    try {
      return await originalLoadFiles(fileList);
    } finally {
      importContext = null;
    }
  };

  const originalSharedSync = syncSharedDataToSupabase;
  syncSharedDataToSupabase = async function patchedSharedSync(options = {}) {
    if (!importContext) return originalSharedSync(options);
    const productSyncOk = await originalSharedSync({ ...options, productsOnly: true });
    const dates = [...importContext.salesDates];
    if (!dates.length) return productSyncOk;
    try {
      await supabaseDeleteRowsByValues("daily_sales", "sales_date", dates);
      const dateSet = new Set(dates);
      const rows = state.rawSales
        .filter((row) => dateSet.has(row.date) && (row.code || row.product))
        .map(salesRowForSupabase);
      if (rows.length) await supabaseInsertRows("daily_sales", rows);
      await updateSharedSyncState("daily-sales");
      return true;
    } catch (error) {
      console.warn("Date-scoped sales sync failed", error);
      return false;
    }
  };

  [els.startDate, els.endDate].filter(Boolean).forEach((input) => {
    input.addEventListener("change", () => { void loadSelectedSalesRange(); });
  });
  els.compareToggle?.addEventListener("change", () => {
    if (els.compareToggle.checked) void loadSelectedSalesRange();
  });

  const salesUi = {
    mode: "calendar",
    calendarMonth: "",
    selectedDate: "",
    chartType: "line",
    itemSort: "sales",
    itemSortDirection: "desc",
    trendRequest: 0,
    rollups: new Map(),
    healthLoadedAt: 0,
    healthPromise: null,
  };

  function salesIso(date) {
    return date.toISOString().slice(0, 10);
  }

  function salesDate(iso) {
    return new Date(`${iso}T12:00:00Z`);
  }

  function salesShiftYear(iso, years) {
    const date = salesDate(iso);
    date.setUTCFullYear(date.getUTCFullYear() + years);
    return salesIso(date);
  }

  function salesMonthBounds(monthKey) {
    const [year, month] = monthKey.split("-").map(Number);
    const start = new Date(Date.UTC(year, month - 1, 1, 12));
    const end = new Date(Date.UTC(year, month, 0, 12));
    return { start: salesIso(start), end: salesIso(end), year, month };
  }

  let salesHierarchyStamp = "";
  let salesHierarchyByCode = new Map();

  function salesHierarchyIndex() {
    const inventoryRows = state.inventoryRows?.length
      ? state.inventoryRows
      : [...(state.latestInventory?.values?.() || [])];
    const stamp = `${state._dataCacheStamp || 0}|${inventoryRows.length}`;
    if (stamp === salesHierarchyStamp) return salesHierarchyByCode;
    salesHierarchyStamp = stamp;
    salesHierarchyByCode = new Map();
    inventoryRows.forEach((row) => {
      const key = codeKey(row.code);
      if (key && !salesHierarchyByCode.has(key)) salesHierarchyByCode.set(key, row);
    });
    return salesHierarchyByCode;
  }

  function salesHierarchyRow(row) {
    const inventory = salesHierarchyIndex().get(codeKey(row.code)) || {};
    const category = cleanCell(row.category) || cleanCell(inventory.category) || "Unassigned";
    const mappedDepartment = GKPOS_CATEGORY_DEPARTMENTS.get(category.toUpperCase()) || "";
    const rowDepartment = cleanCell(row.department);
    const inventoryDepartment = cleanCell(inventory.department);
    const department = mappedDepartment || inventoryDepartment || (rowDepartment && rowDepartment !== category ? rowDepartment : "Unassigned");
    return { ...row, department, category };
  }

  function salesRangeRows(start, end, predicate = null) {
    return state.rawSales
      .filter((row) => row.date >= start && row.date <= end)
      .map(salesHierarchyRow)
      .filter((row) => !predicate || predicate(row));
  }

  function salesTotals(rows) {
    return rows.reduce((sum, row) => {
      sum.sales += Number(row.sales || 0);
      sum.units += Number(row.units || row.qty || 0);
      sum.cost += Number(row.cost || row.cost_sold || 0);
      sum.profit += Number(row.profit || 0);
      return sum;
    }, { sales: 0, units: 0, cost: 0, profit: 0 });
  }

  function salesDelta(current, prior) {
    if (!prior) return null;
    return ((current - prior) / Math.abs(prior)) * 100;
  }

  function salesDeltaText(value) {
    if (value == null || !Number.isFinite(value)) return "No prior-year data";
    return `${value >= 0 ? "+" : ""}${value.toFixed(1)}%`;
  }

  function salesRollupCode(kind) {
    return kind === "DAY" ? "eq.__ROLLUP__DAY" : `like.__ROLLUP__${kind}*`;
  }

  async function salesFetchRollups(kind, start, end, options = {}) {
    const key = `${kind}|${start}|${end}`;
    const cached = salesUi.rollups.get(key);
    if (!options.force && cached && Date.now() - cached.at < 60000) return cached.promise;
    const promise = directSelectRows("daily_sales", {
      select: "sales_date,code,product,vendor,category,department,qty,sales,cost_sold,profit",
      and: `(sales_date.gte.${start},sales_date.lte.${end})`,
      code: salesRollupCode(kind),
      order: "sales_date.asc",
    }).then((rows) => {
      const mapped = rows.map((row) => ({
        date: cleanCell(row.sales_date),
        code: cleanCell(row.code),
        product: cleanCell(row.product),
        vendor: cleanCell(row.vendor) || "Unassigned",
        category: cleanCell(row.category) || "Unassigned",
        department: cleanCell(row.department) || cleanCell(row.category) || "Unassigned",
        units: Number(row.qty || 0),
        sales: Number(row.sales || 0),
        cost: Number(row.cost_sold || 0),
        profit: Number(row.profit || 0),
      }));
      if (kind === "DAY" && mapped.length) {
        const dates = mapped.map((row) => row.date).filter(Boolean).sort();
        const first = dates[0], last = dates.at(-1);
        if (first && (!state._remoteSalesAvailableStart || first < state._remoteSalesAvailableStart)) state._remoteSalesAvailableStart = first;
        if (last && (!state._remoteSalesAvailableEnd || last > state._remoteSalesAvailableEnd)) state._remoteSalesAvailableEnd = last;
        dashboardDataStatus();
      }
      return mapped;
    }).catch((error) => {
      salesUi.rollups.delete(key);
      console.warn("Sales rollup request failed", { kind, start, end, error });
      return [];
    });
    salesUi.rollups.set(key, { at: Date.now(), promise });
    return promise;
  }

  function salesAggregateDaily(rows) {
    const byDate = new Map();
    rows.forEach((row) => {
      const current = byDate.get(row.date) || { date: row.date, sales: 0, units: 0, cost: 0, profit: 0 };
      current.sales += Number(row.sales || 0);
      current.units += Number(row.units || 0);
      current.cost += Number(row.cost || 0);
      current.profit += Number(row.profit || 0);
      byDate.set(row.date, current);
    });
    return [...byDate.values()].sort((a, b) => a.date.localeCompare(b.date));
  }

  function salesAggregateGroup(rows, field) {
    const groups = new Map();
    rows.forEach((row) => {
      const name = cleanCell(row[field]) || "Unassigned";
      const current = groups.get(name) || { name, sales: 0, units: 0, cost: 0, profit: 0 };
      current.sales += Number(row.sales || 0);
      current.units += Number(row.units || 0);
      current.cost += Number(row.cost || 0);
      current.profit += Number(row.profit || 0);
      groups.set(name, current);
    });
    return [...groups.values()];
  }

  function salesSetCanvas(canvas, height) {
    if (!canvas) return null;
    const width = Math.max(canvas.clientWidth || 720, 320);
    const ratio = window.devicePixelRatio || 1;
    canvas.width = width * ratio;
    canvas.height = height * ratio;
    canvas.style.height = `${height}px`;
    const ctx = canvas.getContext("2d");
    ctx.setTransform(ratio, 0, 0, ratio, 0, 0);
    ctx.clearRect(0, 0, width, height);
    return { ctx, width, height };
  }

  function bindSalesChartHover(canvas, points, options = {}) {
    if (!canvas) return;
    let tooltip = document.querySelector("#salesChartTooltip");
    if (!tooltip) {
      tooltip = document.createElement("div");
      tooltip.id = "salesChartTooltip";
      tooltip.className = "sales-chart-tooltip";
      tooltip.hidden = true;
      document.body.append(tooltip);
    }
    let hoverFrame = 0;
    canvas.onmousemove = (event) => {
      cancelAnimationFrame(hoverFrame);
      const clientX = event.clientX;
      const clientY = event.clientY;
      hoverFrame = requestAnimationFrame(() => {
      if (!points.length) { tooltip.hidden = true; return; }
      const rect = canvas.getBoundingClientRect();
      const pointerX = clientX - rect.left;
      const closest = points.reduce((best, point) => Math.abs(point.x - pointerX) < Math.abs(best.x - pointerX) ? point : best, points[0]);
      const threshold = Math.max(12, rect.width / Math.max(points.length, 1) / 2);
      if (Math.abs(closest.x - pointerX) > threshold) { tooltip.hidden = true; return; }
      const formatter = options.percent ? { format: (value) => `${Number(value || 0).toFixed(1)}%` } : (options.currency === false ? number : currency);
      const current = closest.current;
      const prior = closest.prior;
      const label = current?.tooltipLabel || current?.label || prior?.tooltipLabel || prior?.label || "";
      const currentText = current ? `<span class="current-dot"></span>${escapeHtml(options.currentLabel || "Current")}: <b>${formatter.format(current.value)}</b>${current.units == null ? "" : ` · ${number.format(current.units)} units`}` : "";
      const priorText = prior ? `<span class="prior-dot"></span>${escapeHtml(options.priorLabel || "Prior year")}: <b>${formatter.format(prior.value)}</b>${prior.units == null ? "" : ` · ${number.format(prior.units)} units`}` : "";
      const delta = current && prior && Number(prior.value) ? salesDelta(Number(current.value), Number(prior.value)) : null;
      tooltip.innerHTML = `<strong>${escapeHtml(label)}</strong>${currentText ? `<br>${currentText}` : ""}${priorText ? `<br>${priorText}` : ""}${delta == null ? "" : `<br><em class="${delta < 0 ? "down" : "up"}">${salesDeltaText(delta)} vs prior year</em>`}`;
      tooltip.hidden = false;
      tooltip.style.left = `${Math.min(clientX + 14, window.innerWidth - 270)}px`;
      tooltip.style.top = `${Math.max(8, clientY - 100)}px`;
      });
    };
    canvas.onmouseleave = () => { cancelAnimationFrame(hoverFrame); tooltip.hidden = true; };
  }

  function salesDrawChart(canvas, current, prior = [], options = {}) {
    const setup = salesSetCanvas(canvas, Number(options.height || canvas?.getAttribute("height") || 260));
    if (!setup) return;
    const { ctx, width, height } = setup;
    const left = 48, right = width - 18, top = 18, bottom = height - 34;
    const values = [...current, ...prior].map((row) => Number(row.value || 0));
    const max = Math.max(...values, 1);
    ctx.strokeStyle = "#dfe7e3";
    ctx.fillStyle = "#6d7773";
    ctx.font = "11px Arial";
    for (let i = 0; i < 5; i += 1) {
      const y = top + ((bottom - top) * i) / 4;
      ctx.beginPath(); ctx.moveTo(left, y); ctx.lineTo(right, y); ctx.stroke();
      const label = max * (1 - i / 4);
      ctx.fillText(options.percent ? `${label.toFixed(0)}%` : (options.currency === false ? number.format(label) : compactCurrency(label)), 4, y + 4);
    }
    if (!current.length) {
      bindSalesChartHover(canvas, [], options);
      ctx.textAlign = "center";
      ctx.font = "15px Arial";
      ctx.fillText("No data for this selection yet.", width / 2, height / 2);
      ctx.textAlign = "left";
      return;
    }
    const count = Math.max(current.length, prior.length, 1);
    const xFor = (index) => count === 1 ? (left + right) / 2 : left + ((right - left) * index) / (count - 1);
    const yFor = (value) => bottom - ((bottom - top) * Number(value || 0)) / max;
    const drawLine = (rows, color, dashed) => {
      if (!rows.length) return;
      ctx.strokeStyle = color; ctx.lineWidth = dashed ? 1.8 : 2.8; ctx.setLineDash(dashed ? [6, 5] : []);
      ctx.beginPath();
      rows.forEach((row, index) => { const x = xFor(index), y = yFor(row.value); if (!index) ctx.moveTo(x, y); else ctx.lineTo(x, y); });
      ctx.stroke(); ctx.setLineDash([]);
    };
    if (options.type === "bar") {
      const slot = Math.max(4, (right - left) / count);
      const bar = Math.max(2, Math.min(18, slot * 0.34));
      current.forEach((row, index) => { const x = xFor(index) - bar; const y = yFor(row.value); ctx.fillStyle = "#16835b"; ctx.fillRect(x, y, bar, bottom - y); });
      prior.forEach((row, index) => { const x = xFor(index) + 1; const y = yFor(row.value); ctx.fillStyle = "#9cb7db"; ctx.fillRect(x, y, bar, bottom - y); });
    } else {
      drawLine(prior, "#8da9d2", true);
      drawLine(current, "#16835b", false);
    }
    ctx.fillStyle = "#6d7773"; ctx.font = "11px Arial"; ctx.textAlign = "center";
    const labelStep = count <= 31 && width >= 760 ? 1 : Math.max(1, Math.ceil(count / 12));
    current.forEach((row, index) => {
      if (index % labelStep !== 0 && index !== current.length - 1) return;
      ctx.fillText(row.label || "", xFor(index), height - 10);
    });
    ctx.textAlign = "left";
    const hoverPoints = Array.from({ length: count }, (_, index) => ({
      x: xFor(index),
      current: current[index] || null,
      prior: prior[index] || null,
    }));
    bindSalesChartHover(canvas, hoverPoints, options);
  }

  function compactCurrency(value) {
    const numberValue = Number(value || 0);
    if (Math.abs(numberValue) >= 1000000) return `$${(numberValue / 1000000).toFixed(1)}m`;
    if (Math.abs(numberValue) >= 1000) return `$${(numberValue / 1000).toFixed(0)}k`;
    return `$${Math.round(numberValue)}`;
  }

  function renderSalesYoyCards(rows, latest) {
    const target = document.querySelector("#salesYoyCards");
    if (!target || !latest) return;
    const latestDate = salesDate(latest);
    const previousLatest = salesShiftYear(latest, -1);
    const monthStart = `${latest.slice(0, 7)}-01`;
    const previousMonthStart = salesShiftYear(monthStart, -1);
    const yearStart = `${latest.slice(0, 4)}-01-01`;
    const previousYearStart = salesShiftYear(yearStart, -1);
    const total = (start, end) => salesTotals(rows.filter((row) => row.date >= start && row.date <= end));
    const cards = [
      { label: "Selected day", current: total(latest, latest), prior: total(previousLatest, previousLatest), detail: salesDisplayDate(latest) },
      { label: "Month to date", current: total(monthStart, latest), prior: total(previousMonthStart, previousLatest), detail: latestDate.toLocaleDateString("en-US", { month: "long", year: "numeric", timeZone: "UTC" }) },
      { label: "Year to date", current: total(yearStart, latest), prior: total(previousYearStart, previousLatest), detail: latest.slice(0, 4) },
    ];
    target.innerHTML = cards.map((card) => {
      const delta = salesDelta(card.current.sales, card.prior.sales);
      const hasPrior = card.prior.sales || card.prior.units;
      const difference = card.current.sales - card.prior.sales;
      const margin = card.current.sales ? (card.current.profit / card.current.sales) * 100 : 0;
      return `<article class="yoy-card">
        <div class="yoy-card__top"><span>${escapeHtml(card.label)}</span><small>${escapeHtml(card.detail)}</small></div>
        <strong>${currency.format(card.current.sales)}</strong>
        <div class="yoy-card__compare ${delta != null && delta < 0 ? "down" : "up"}">${hasPrior ? `${salesDeltaText(delta)} · ${difference >= 0 ? "+" : "−"}${currency.format(Math.abs(difference))}` : "No prior-year match"}</div>
        <div class="yoy-card__foot"><span><small>Prior year</small><b>${hasPrior ? currency.format(card.prior.sales) : "—"}</b></span><span><small>Units</small><b>${number.format(card.current.units)}</b></span><span><small>Margin</small><b>${margin.toFixed(1)}%</b></span></div>
      </article>`;
    }).join("");
    renderSalesExecutiveSummary(rows, latest);
  }

  function salesOperatingDays(start, end) {
    let count = 0;
    for (let iso = start; iso <= end; iso = shiftIsoDate(iso, 1)) {
      if (salesDate(iso).getUTCDay() !== 0) count += 1;
    }
    return count;
  }

  function renderSalesExecutiveSummary(rows, latest) {
    const target = document.querySelector("#salesExecutiveSummary");
    if (!target || !latest) return;
    const yearStart = `${latest.slice(0, 4)}-01-01`;
    const monthStart = `${latest.slice(0, 7)}-01`;
    const monthEnd = salesMonthBounds(latest.slice(0, 7)).end;
    const yearRows = rows.filter((row) => row.date >= yearStart && row.date <= latest);
    const monthRows = rows.filter((row) => row.date >= monthStart && row.date <= latest);
    const ytd = salesTotals(yearRows);
    const mtd = salesTotals(monthRows);
    const sellingDays = Math.max(1, monthRows.filter((row) => Number(row.sales || 0) !== 0).length);
    const totalOperatingDays = salesOperatingDays(monthStart, monthEnd);
    const projection = (mtd.sales / sellingDays) * totalOperatingDays;
    const avgDay = ytd.sales / Math.max(1, yearRows.filter((row) => Number(row.sales || 0) !== 0).length);
    const grossMargin = ytd.sales ? (ytd.profit / ytd.sales) * 100 : 0;
    target.innerHTML = `
      <article><small>YTD gross profit</small><b>${currency.format(ytd.profit)}</b><span>${grossMargin.toFixed(1)}% gross margin</span></article>
      <article><small>Average selling day</small><b>${currency.format(avgDay)}</b><span>${number.format(yearRows.filter((row) => Number(row.sales || 0) !== 0).length)} completed days</span></article>
      <article><small>Month-end projection</small><b>${currency.format(projection)}</b><span>Based on ${number.format(sellingDays)} completed selling days</span></article>
      <article><small>YTD total cost</small><b>${currency.format(ytd.cost)}</b><span>Finalized cost of goods sold</span></article>`;
  }

  function renderSalesDepartmentTree(rows) {
    const target = document.querySelector("#salesOverviewCategories");
    if (!target) return;
    const departmentRank = (name) => {
      const normalized = cleanCell(name).toUpperCase().replace(/^BEAUTY PRODUCT$/, "BEAUTY PRODUCTS");
      const index = GKPOS_DEPARTMENT_ORDER.indexOf(normalized);
      return index === -1 ? GKPOS_DEPARTMENT_ORDER.length : index;
    };
    const departments = salesAggregateGroup(rows, "department").sort((a, b) => {
      const rankDifference = departmentRank(a.name) - departmentRank(b.name);
      return rankDifference || a.name.localeCompare(b.name);
    });
    target.innerHTML = departments.length ? departments.map((department) => {
      const departmentRows = rows.filter((row) => (cleanCell(row.department) || "Unassigned") === department.name);
      const categories = salesAggregateGroup(departmentRows, "category").sort((a, b) => b.sales - a.sales);
      const tone = Math.min(departmentRank(department.name) + 1, GKPOS_DEPARTMENT_ORDER.length + 1);
      return `<details class="sales-department department-tone-${tone}">
        <summary><span><b>${escapeHtml(department.name)}</b><small>${number.format(department.units)} units</small></span><span class="sales-mix-actions"><strong>${currency.format(department.sales)}</strong><button type="button" data-open-department="${escapeHtml(department.name)}">Chart</button></span></summary>
        <div class="sales-category-branches">${categories.map((category) => `<button type="button" class="sales-category-chart" data-open-category="${escapeHtml(category.name)}"><span><b>${escapeHtml(category.name)}</b><small>${number.format(category.units)} units</small></span><strong>${currency.format(category.sales)}</strong></button>`).join("")}</div>
      </details>`;
    }).join("") : `<div class="sales-empty">Department totals will appear after sales load.</div>`;
  }

  async function renderSalesOverview() {
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    if (!latest) return;
    const trendDate = document.querySelector("#salesTrendDate");
    if (trendDate && !trendDate.value) trendDate.value = salesUi.selectedDate || latest;
    const focus = trendDate?.value || salesUi.selectedDate || latest;
    const focusLabel = document.querySelector("#salesOverviewFocus");
    if (focusLabel) focusLabel.textContent = `${salesDisplayDate(focus)} compared with the same period last year.`;
    const previousYearStart = salesShiftYear(`${focus.slice(0, 4)}-01-01`, -1);
    const rollups = await salesFetchRollups("DAY", previousYearStart, focus);
    const fallback = salesAggregateDaily(salesRangeRows(previousYearStart, focus));
    const daily = rollups.length ? rollups : fallback;
    const coverageStart = daily[0]?.date || "";
    const serverShouldCoverComparison = state._remoteSalesAvailableStart && state._remoteSalesAvailableStart <= previousYearStart;
    if (serverShouldCoverComparison && (!rollups.length || !coverageStart || coverageStart > shiftIsoDate(previousYearStart, 7))) {
      const cards = document.querySelector("#salesYoyCards");
      const executive = document.querySelector("#salesExecutiveSummary");
      if (cards) cards.innerHTML = `<div class="sales-loading sales-history-warning">Loading complete year-over-year history… Partial totals are intentionally hidden.</div>`;
      if (executive) executive.innerHTML = "";
      return;
    }
    renderSalesYoyCards(daily, focus);
    await renderSalesTrend();
  }

  function renderSalesCalendarGrid(rows) {
    const target = document.querySelector("#salesCalendar");
    const title = document.querySelector("#salesCalendarTitle");
    if (!target || !salesUi.calendarMonth) return;
    const bounds = salesMonthBounds(salesUi.calendarMonth);
    const date = new Date(Date.UTC(bounds.year, bounds.month - 1, 1, 12));
    if (title) title.textContent = date.toLocaleDateString("en-US", { month: "long", year: "numeric", timeZone: "UTC" });
    const datePicker = document.querySelector("#salesCalendarDatePicker");
    if (datePicker) datePicker.value = salesUi.selectedDate?.slice(0, 7) === salesUi.calendarMonth ? salesUi.selectedDate : bounds.start;
    const daily = new Map(salesAggregateDaily(rows).map((row) => [row.date, row]));
    const firstDay = date.getUTCDay();
    const daysInMonth = Number(bounds.end.slice(8));
    const cells = [];
    ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].forEach((day) => cells.push(`<div class="sales-calendar__weekday">${day}</div>`));
    for (let index = 0; index < firstDay; index += 1) cells.push(`<div class="sales-calendar__day muted"></div>`);
    for (let day = 1; day <= daysInMonth; day += 1) {
      const iso = `${salesUi.calendarMonth}-${String(day).padStart(2, "0")}`;
      const row = daily.get(iso);
      const classes = ["sales-calendar__day", row ? "has-sales" : "", iso === salesUi.selectedDate ? "selected" : ""].filter(Boolean).join(" ");
      cells.push(`<button type="button" class="${classes}" data-sales-date="${iso}"><span>${day}</span>${row ? `<strong>${currency.format(row.sales)}</strong><small>${number.format(row.units)} units</small>` : `<small>No sales</small>`}</button>`);
    }
    target.innerHTML = cells.join("");
  }

  async function renderSalesCalendar() {
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || salesIso(new Date());
    if (!salesUi.calendarMonth) salesUi.calendarMonth = latest.slice(0, 7);
    const requestedMonth = salesUi.calendarMonth;
    const bounds = salesMonthBounds(salesUi.calendarMonth);
    const priorBounds = { start: salesShiftYear(bounds.start, -1), end: salesShiftYear(bounds.end, -1) };
    const [rollups, priorRollups] = await Promise.all([
      salesFetchRollups("DAY", bounds.start, bounds.end),
      salesFetchRollups("DAY", priorBounds.start, priorBounds.end),
    ]);
    let rows = rollups;
    if (!rows.length) {
      await ensureRemoteSalesRange(bounds.start, bounds.end, { silent: true });
      if (requestedMonth !== salesUi.calendarMonth) return;
      rows = salesRangeRows(bounds.start, bounds.end);
    }
    renderSalesCalendarGrid(rows);
    const latestRollupDate = rows.reduce((max, row) => row.date > max ? row.date : max, "");
    const currentMonth = salesIso(new Date()).slice(0, 7);
    const knownEnd = state._remoteSalesAvailableEnd || latestRollupDate;
    const isOpenMonth = salesUi.calendarMonth === currentMonth || (knownEnd && knownEnd.slice(0, 7) === salesUi.calendarMonth);
    const cutoff = isOpenMonth && latestRollupDate ? latestRollupDate : bounds.end;
    const priorCutoff = salesShiftYear(cutoff, -1);
    const currentTotal = salesTotals(rows.filter((row) => row.date <= cutoff));
    const priorTotal = salesTotals(priorRollups.filter((row) => row.date <= priorCutoff));
    const delta = salesDelta(currentTotal.sales, priorTotal.sales);
    const summary = document.querySelector("#salesCalendarMonthSummary");
    if (summary) summary.innerHTML = `<b>${currency.format(currentTotal.sales)}</b><span>${cutoff === bounds.end ? "month total" : `through ${salesDisplayDate(cutoff)}`}</span><em class="${delta != null && delta < 0 ? "down" : "up"}">${priorTotal.sales ? `${salesDeltaText(delta)} vs same period ${Number(bounds.year) - 1} · ${currency.format(priorTotal.sales)}` : "No prior-year match"}</em>`;
  }

  async function renderSalesDayDetail(date) {
    const target = document.querySelector("#salesDayDetail");
    if (!target || !date) return;
    target.innerHTML = `<div class="sales-loading">Loading ${escapeHtml(date)}...</div>`;
    await ensureRemoteSalesRange(date, date);
    const rows = salesRangeRows(date, date);
    const totals = salesTotals(rows);
    const categories = salesAggregateGroup(rows, "category").sort((a, b) => b.sales - a.sales);
    const items = salesAggregateGroup(rows.map((row) => ({ ...row, item: row.product || row.code })), "item").sort((a, b) => b.sales - a.sales).slice(0, 30);
    target.innerHTML = `<div class="sales-day-detail__heading">
      <div><p class="eyebrow">Daily detail</p><h2>${salesDate(date).toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric", timeZone: "UTC" })}</h2></div>
      <div class="sales-day-kpis"><span><small>Sales</small><b>${currency.format(totals.sales)}</b></span><span><small>Units</small><b>${number.format(totals.units)}</b></span><span><small>Profit</small><b>${currency.format(totals.profit)}</b></span></div>
    </div>
    <div class="sales-day-detail__grid">
      <div><h3>Category / department breakdown</h3><div class="sales-ranked-list">${categories.slice(0, 20).map((row) => `<button type="button" data-open-trend="${escapeHtml(row.name)}"><span><b>${escapeHtml(row.name)}</b><small>${number.format(row.units)} units</small></span><strong>${currency.format(row.sales)}</strong></button>`).join("") || `<div class="sales-empty">No category sales.</div>`}</div></div>
      <div><h3>Top items</h3><div class="sales-ranked-list sales-ranked-list--items">${items.map((row) => `<button type="button" data-open-item="${escapeHtml(row.name)}"><span><b>${escapeHtml(row.name)}</b><small>${number.format(row.units)} units</small></span><strong>${currency.format(row.sales)}</strong></button>`).join("") || `<div class="sales-empty">No item sales.</div>`}</div></div>
    </div>`;
  }

  function salesTrendDates(period, anchor, latest) {
    const date = salesDate(anchor || latest);
    const year = date.getUTCFullYear();
    const month = date.getUTCMonth();
    let startDate = new Date(Date.UTC(year, month, date.getUTCDate(), 12));
    let endDate = new Date(startDate);
    if (period === "month") {
      startDate = new Date(Date.UTC(year, month, 1, 12));
      endDate = new Date(Date.UTC(year, month + 1, 0, 12));
    } else if (period === "quarter") {
      const firstMonth = Math.floor(month / 3) * 3;
      startDate = new Date(Date.UTC(year, firstMonth, 1, 12));
      endDate = new Date(Date.UTC(year, firstMonth + 3, 0, 12));
    } else if (period === "half") {
      const firstMonth = month < 6 ? 0 : 6;
      startDate = new Date(Date.UTC(year, firstMonth, 1, 12));
      endDate = new Date(Date.UTC(year, firstMonth + 6, 0, 12));
    } else if (period === "year") {
      startDate = new Date(Date.UTC(year, 0, 1, 12));
      endDate = new Date(Date.UTC(year, 11, 31, 12));
    }
    const start = salesIso(startDate);
    let end = salesIso(endDate);
    if (latest && start <= latest && end > latest) end = latest;
    return { start, end };
  }

  function salesTrendOptions(group) {
    if (!group) return [];
    const field = group === "item" ? "product" : group;
    const inventoryRows = state.inventoryRows?.length ? state.inventoryRows : [...(state.latestInventory?.values?.() || [])];
    const inventoryField = group === "item" ? "name" : field;
    return [...new Set([
      ...state.rawSales.map((row) => cleanCell(row[field])),
      ...inventoryRows.map((row) => cleanCell(row[inventoryField]) || (group === "item" ? cleanCell(row.item_name) : "")),
    ].filter(Boolean))].sort(compareDisplayValue);
  }

  function updateSalesTrendOptions() {
    const group = document.querySelector("#salesTrendGroup")?.value || "";
    const list = document.querySelector("#salesTrendEntities");
    if (!list) return;
    const values = salesTrendOptions(group);
    list.innerHTML = values.slice(0, 5000).map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
    const input = document.querySelector("#salesTrendEntity");
    if (input) {
      input.disabled = !group;
      input.placeholder = group ? `Search ${group}` : "All sales";
      if (!group) input.value = "";
    }
  }

  function salesBucketSeries(rows, metric, monthly) {
    const grouped = new Map();
    rows.forEach((row) => {
      const key = monthly ? row.date.slice(0, 7) : row.date;
      const current = grouped.get(key) || { key, label: monthly ? key.slice(5) : row.date.slice(5), tooltipLabel: monthly ? salesDisplayDate(`${key}-01`) : salesDisplayDate(row.date), value: 0, sales: 0, profit: 0, units: 0 };
      current.sales += Number(row.sales || 0);
      current.profit += Number(row.profit || 0);
      current.units += Number(row.units || row.qty || 0);
      current.value = metric === "margin" ? (current.sales ? (current.profit / current.sales) * 100 : 0) : Number(current[metric] || 0);
      grouped.set(key, current);
    });
    return [...grouped.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([, row]) => row);
  }

  function salesAlignedTrendSeries(currentRows, priorRows, metric, monthly, dates) {
    const currentBuckets = salesBucketSeries(currentRows, metric, monthly);
    const priorBuckets = salesBucketSeries(priorRows, metric, monthly);
    const currentMap = new Map(currentBuckets.map((row) => [row.key, row.value]));
    const priorMap = new Map(priorBuckets.map((row) => [salesShiftYear(`${row.key}${monthly ? "-01" : ""}`, 1).slice(0, monthly ? 7 : 10), row.value]));
    const current = [];
    const prior = [];
    if (monthly) {
      let cursor = `${dates.start.slice(0, 7)}-01`;
      const last = `${dates.end.slice(0, 7)}-01`;
      while (cursor <= last) {
        const key = cursor.slice(0, 7);
        current.push({ label: key.slice(5), tooltipLabel: salesDisplayDate(`${key}-01`), value: Number(currentMap.get(key) || 0) });
        const priorKey = salesShiftYear(cursor, -1).slice(0, 7);
        prior.push({ label: key.slice(5), tooltipLabel: salesDisplayDate(`${priorKey}-01`), value: Number(priorMap.get(key) || 0) });
        const date = salesDate(cursor); date.setUTCMonth(date.getUTCMonth() + 1); cursor = salesIso(date).slice(0, 7) + "-01";
      }
    } else {
      for (let iso = dates.start; iso <= dates.end; iso = shiftIsoDate(iso, 1)) {
        current.push({ label: iso.slice(5), tooltipLabel: salesDisplayDate(iso), value: Number(currentMap.get(iso) || 0) });
        const priorIso = salesShiftYear(iso, -1);
        prior.push({ label: iso.slice(5), tooltipLabel: salesDisplayDate(priorIso), value: Number(priorMap.get(iso) || 0) });
      }
    }
    return { current, prior };
  }

  async function renderSalesTrend() {
    const request = ++salesUi.trendRequest;
    const group = document.querySelector("#salesTrendGroup")?.value || "";
    const entity = cleanCell(document.querySelector("#salesTrendEntity")?.value);
    const metric = document.querySelector("#salesTrendMetric")?.value || "sales";
    const period = document.querySelector("#salesTrendPeriod")?.value || "month";
    const compare = !!document.querySelector("#salesTrendCompare")?.checked;
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    const anchor = document.querySelector("#salesTrendDate")?.value || salesUi.selectedDate || latest;
    const title = document.querySelector("#salesTrendTitle");
    const kicker = document.querySelector("#salesTrendKicker");
    if (kicker) kicker.textContent = group ? `${group[0].toUpperCase()}${group.slice(1)} analysis` : "Total sales";
    if (title) title.textContent = group ? (entity || `Choose a ${group}`) : "All-store performance";
    if (!latest || (group && !entity)) {
      const summary = document.querySelector("#salesTrendSummary");
      const legend = document.querySelector("#salesTrendLegend");
      if (legend) legend.innerHTML = "";
      if (summary) summary.innerHTML = `<span>Choose a ${escapeHtml(group)} to view its performance.</span>`;
      return salesDrawChart(document.querySelector("#salesTrendExplorerChart"), [], [], { height: 300 });
    }
    const dates = salesTrendDates(period, anchor, latest);
    const priorDates = { start: salesShiftYear(dates.start, -1), end: salesShiftYear(dates.end, -1) };
    let rows = [];
    if (!group) {
      rows = await salesFetchRollups("DAY", compare ? priorDates.start : dates.start, dates.end);
    } else if (group === "item") {
      const match = state.rawSales.find((row) => cleanCell(row.product).toLowerCase() === entity.toLowerCase());
      if (match?.code) {
        rows = await directSelectRows("daily_sales", {
          select: "sales_date,code,product,vendor,category,department,qty,sales,cost_sold,profit",
          and: `(sales_date.gte.${compare ? priorDates.start : dates.start},sales_date.lte.${dates.end})`,
          code: `eq.${match.code}`,
          order: "sales_date.asc",
        }).then((result) => result.map(hydrateSalesFromSupabase)).catch(() => []);
      }
    } else {
      const rollupKind = group === "vendor" ? "VENDOR__" : "CATEGORY__";
      rows = await salesFetchRollups(rollupKind, compare ? priorDates.start : dates.start, dates.end);
      rows = rows.filter((row) => cleanCell(row[group]).toLowerCase() === entity.toLowerCase());
      if (!rows.length) rows = salesRangeRows(compare ? priorDates.start : dates.start, dates.end, (row) => cleanCell(row[group]).toLowerCase() === entity.toLowerCase());
    }
    if (request !== salesUi.trendRequest) return;
    const monthly = period === "quarter" || period === "half" || period === "year";
    const currentRows = rows.filter((row) => row.date >= dates.start && row.date <= dates.end);
    const priorRows = compare ? rows.filter((row) => row.date >= priorDates.start && row.date <= priorDates.end) : [];
    const aligned = salesAlignedTrendSeries(currentRows, priorRows, metric, monthly, dates);
    const current = aligned.current;
    const prior = compare ? aligned.prior : [];
    const currentYear = dates.end.slice(0, 4);
    const priorYear = String(Number(currentYear) - 1);
    const legend = document.querySelector("#salesTrendLegend");
    if (legend) legend.innerHTML = `<span><i class="current"></i>${currentYear}</span>${compare ? `<span><i class="prior"></i>${priorYear}</span>` : ""}`;
    salesDrawChart(document.querySelector("#salesTrendExplorerChart"), current, prior, { type: salesUi.chartType, height: 320, currency: metric !== "units" && metric !== "margin", percent: metric === "margin", currentLabel: currentYear, priorLabel: priorYear });
    const currentTotals = salesTotals(currentRows);
    const priorTotals = salesTotals(priorRows);
    const currentTotal = metric === "margin" ? (currentTotals.sales ? (currentTotals.profit / currentTotals.sales) * 100 : 0) : current.reduce((sum, row) => sum + row.value, 0);
    const priorTotal = metric === "margin" ? (priorTotals.sales ? (priorTotals.profit / priorTotals.sales) * 100 : 0) : prior.reduce((sum, row) => sum + row.value, 0);
    const formatter = metric === "units" ? number : metric === "margin" ? { format: (value) => `${Number(value || 0).toFixed(1)}%` } : currency;
    const summary = document.querySelector("#salesTrendSummary");
    if (summary) summary.innerHTML = `<b>${formatter.format(currentTotal)}</b><span>${salesDisplayDate(dates.start)}–${salesDisplayDate(dates.end)}</span>${compare ? `<b class="${currentTotal >= priorTotal ? "up" : "down"}">${salesDeltaText(salesDelta(currentTotal, priorTotal))}</b><span>vs ${priorYear} (${formatter.format(priorTotal)})</span>` : ""}`;
    const mixRows = await salesFetchRollups("CATEGORY__", dates.start, dates.end);
    if (request === salesUi.trendRequest) renderSalesDepartmentTree(mixRows);
  }

  async function renderSalesItems() {
    const body = document.querySelector("#salesItemRows");
    if (!body) return;
    const start = document.querySelector("#salesItemsStart")?.value || els.startDate.value || state.dates[0] || "";
    const end = document.querySelector("#salesItemsEnd")?.value || els.endDate.value || state.dates.at(-1) || "";
    const range = document.querySelector("#salesItemsRange");
    if (range) range.textContent = start && end ? `${salesDisplayDate(start)} through ${salesDisplayDate(end)}` : "Choose a date range.";
    const query = cleanCell(document.querySelector("#salesItemSearch")?.value).toLowerCase();
    const departmentSelect = document.querySelector("#salesItemDepartment");
    const categorySelect = document.querySelector("#salesItemCategory");
    const vendorSelect = document.querySelector("#salesItemVendor");
    const sortSelect = document.querySelector("#salesItemSort");
    const sort = salesUi.itemSort || sortSelect?.value || "sales";
    if (sortSelect && sortSelect.value !== sort && [...sortSelect.options].some((option) => option.value === sort)) sortSelect.value = sort;
    const map = new Map();
    salesRangeRows(start, end).forEach((row) => {
      const key = codeKey(row.code) || cleanCell(row.product).toLowerCase();
      const current = map.get(key) || { code: row.code, product: row.product || row.code, department: row.department, category: row.category, vendor: row.vendor, units: 0, sales: 0, cost: 0, profit: 0 };
      current.units += Number(row.units || 0); current.sales += Number(row.sales || 0); current.cost += Number(row.cost || 0); current.profit += Number(row.profit || 0);
      map.set(key, current);
    });
    const setFilterOptions = (select, values, allLabel) => {
      if (!select) return "";
      const current = select.value;
      const options = [...new Set(values.map((value) => cleanCell(value) || "Unassigned"))].sort((a, b) => a.localeCompare(b));
      select.innerHTML = `<option value="">${allLabel}</option>${options.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}`;
      select.value = options.includes(current) ? current : "";
      return select.value;
    };
    const allRows = [...map.values()].map((row) => ({ ...row, price: row.units ? row.sales / row.units : 0, unitCost: row.units ? row.cost / row.units : 0 }));
    const departmentOrder = (name) => {
      const index = GKPOS_DEPARTMENT_ORDER.indexOf(cleanCell(name).toUpperCase());
      return index === -1 ? GKPOS_DEPARTMENT_ORDER.length : index;
    };
    const departments = [...new Set(allRows.map((row) => cleanCell(row.department) || "Unassigned"))]
      .sort((a, b) => departmentOrder(a) - departmentOrder(b) || a.localeCompare(b));
    const department = setFilterOptions(departmentSelect, departments, "All departments");
    const departmentRows = allRows.filter((row) => !department || (cleanCell(row.department) || "Unassigned") === department);
    const category = setFilterOptions(categorySelect, departmentRows.map((row) => row.category), "All categories");
    const categoryRows = departmentRows.filter((row) => !category || (cleanCell(row.category) || "Unassigned") === category);
    const vendor = setFilterOptions(vendorSelect, categoryRows.map((row) => row.vendor), "All vendors");
    let rows = categoryRows
      .filter((row) => !vendor || (cleanCell(row.vendor) || "Unassigned") === vendor)
      .filter((row) => !query || [row.code, row.product].join(" ").toLowerCase().includes(query));
    const direction = salesUi.itemSortDirection === "asc" ? 1 : -1;
    rows.sort((a, b) => ["product", "department", "category", "vendor"].includes(sort)
      ? direction * cleanCell(a[sort]).localeCompare(cleanCell(b[sort]))
      : direction * (Number(a[sort] || 0) - Number(b[sort] || 0)));
    const totals = salesTotals(rows);
    const finalizedDays = await salesFetchRollups("DAY", start, end);
    const hasItemFilters = Boolean(department || category || vendor || query);
    const finalizedSubtotal = !hasItemFilters && finalizedDays.length ? salesTotals(finalizedDays).sales : totals.sales;
    const totalTarget = document.querySelector("#salesItemsTotals");
    if (totalTarget) totalTarget.innerHTML = `<span><small>Products</small><b>${number.format(rows.length)}</b></span><span><small>Units</small><b>${number.format(totals.units)}</b></span><span><small>Item sales</small><b>${currency.format(totals.sales)}</b></span><span><small>Total cost</small><b>${currency.format(totals.cost)}</b></span><span><small>Final subtotal</small><b>${currency.format(finalizedSubtotal)}</b></span><span><small>Total profit</small><b>${currency.format(totals.profit)}</b></span>`;
    document.querySelectorAll("[data-item-sort]").forEach((button) => {
      const active = button.dataset.itemSort === sort;
      button.classList.toggle("active", active);
      button.dataset.direction = active ? salesUi.itemSortDirection : "";
    });
    body.innerHTML = rows.slice(0, 500).map((row) => `<tr data-open-item="${escapeHtml(row.product)}"><td><b>${escapeHtml(row.product)}</b><small>${escapeHtml(row.code)}</small></td><td>${escapeHtml(row.department || "Unassigned")}</td><td>${escapeHtml(row.category || "Unassigned")}</td><td>${escapeHtml(row.vendor || "Unassigned")}</td><td>${number.format(row.units)}</td><td>${currency.format(row.price)}</td><td>${currency.format(row.sales)}</td><td>${currency.format(row.unitCost)}</td><td>${currency.format(row.cost)}</td><td>${currency.format(row.profit)}</td></tr>`).join("") || `<tr><td colspan="10" class="sales-empty">No matching sales items.</td></tr>`;
  }

  function setSalesMode(mode) {
    salesUi.mode = mode;
    document.querySelectorAll("[data-sales-mode]").forEach((button) => button.classList.toggle("active", button.dataset.salesMode === mode));
    document.querySelectorAll("[data-sales-mode-view]").forEach((view) => {
      const active = view.dataset.salesModeView === mode;
      view.hidden = !active; view.classList.toggle("active", active);
    });
    if (mode === "overview") void renderSalesOverview();
    if (mode === "calendar") void renderSalesCalendar();
    if (mode === "items") {
      const startInput = document.querySelector("#salesItemsStart"), endInput = document.querySelector("#salesItemsEnd");
      const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
      if (startInput && !startInput.value) startInput.value = shiftIsoDate(latest, -29);
      if (endInput && !endInput.value) endInput.value = latest;
      const start = startInput?.value || els.startDate.value, end = endInput?.value || els.endDate.value;
      void ensureRemoteSalesRange(start, end).then(renderSalesItems);
    }
  }

  function renderSalesWorkspace() {
    if (!document.querySelector("#salesWorkspace")) return;
    void renderSalesDataHealth();
    if (salesUi.mode === "overview") void renderSalesOverview();
    else if (salesUi.mode === "calendar") void renderSalesCalendar();
    else if (salesUi.mode === "items") void renderSalesItems();
  }

  document.querySelector("#salesModeTabs")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-sales-mode]");
    if (button) setSalesMode(button.dataset.salesMode);
  });
  document.querySelector("#salesHealthRefresh")?.addEventListener("click", () => { void requestSalesRecovery(); });
  document.querySelector("#salesCalendar")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-sales-date]");
    if (!button) return;
    salesUi.selectedDate = button.dataset.salesDate;
    const trendDate = document.querySelector("#salesTrendDate"), period = document.querySelector("#salesTrendPeriod");
    if (trendDate) trendDate.value = salesUi.selectedDate;
    if (period) period.value = "day";
    const startInput = document.querySelector("#salesItemsStart"), endInput = document.querySelector("#salesItemsEnd");
    if (startInput) startInput.value = salesUi.selectedDate;
    if (endInput) endInput.value = salesUi.selectedDate;
    setSalesMode("overview");
  });
  document.querySelector("#salesCalendarPrev")?.addEventListener("click", () => {
    const bounds = salesMonthBounds(salesUi.calendarMonth);
    const date = new Date(Date.UTC(bounds.year, bounds.month - 2, 1, 12));
    salesUi.calendarMonth = salesIso(date).slice(0, 7); salesUi.selectedDate = ""; void renderSalesCalendar();
  });
  document.querySelector("#salesCalendarNext")?.addEventListener("click", () => {
    const bounds = salesMonthBounds(salesUi.calendarMonth);
    const date = new Date(Date.UTC(bounds.year, bounds.month, 1, 12));
    salesUi.calendarMonth = salesIso(date).slice(0, 7); salesUi.selectedDate = ""; void renderSalesCalendar();
  });
  document.querySelector("#salesCalendarToday")?.addEventListener("click", () => {
    salesUi.calendarMonth = (state._remoteSalesAvailableEnd || state.dates.at(-1) || salesIso(new Date())).slice(0, 7); salesUi.selectedDate = ""; void renderSalesCalendar();
  });
  document.querySelector("#salesCalendarDatePicker")?.addEventListener("change", (event) => {
    if (!event.target.value) return;
    salesUi.calendarMonth = event.target.value.slice(0, 7);
    salesUi.selectedDate = event.target.value;
    void renderSalesCalendar();
  });
  ["#salesTrendGroup", "#salesTrendMetric", "#salesTrendPeriod", "#salesTrendCompare", "#salesTrendDate"].forEach((selector) => document.querySelector(selector)?.addEventListener("change", () => {
    if (selector === "#salesTrendGroup") { const input = document.querySelector("#salesTrendEntity"); if (input) input.value = ""; updateSalesTrendOptions(); }
    void renderSalesTrend();
  }));
  document.querySelector("#salesTrendEntity")?.addEventListener("change", () => { void renderSalesTrend(); });
  document.querySelector("#salesTrendEntity")?.addEventListener("keydown", (event) => { if (event.key === "Enter") void renderSalesTrend(); });
  document.querySelector("#salesTrendClear")?.addEventListener("click", () => {
    const group = document.querySelector("#salesTrendGroup"), entity = document.querySelector("#salesTrendEntity");
    const metric = document.querySelector("#salesTrendMetric"), period = document.querySelector("#salesTrendPeriod"), compare = document.querySelector("#salesTrendCompare");
    if (group) group.value = "";
    if (entity) entity.value = "";
    if (metric) metric.value = "sales";
    const date = document.querySelector("#salesTrendDate");
    if (period) period.value = "month";
    if (date) date.value = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    if (compare) compare.checked = true;
    salesUi.chartType = "line";
    document.querySelectorAll("[data-chart-type]").forEach((button) => button.classList.toggle("active", button.dataset.chartType === "line"));
    updateSalesTrendOptions();
    void renderSalesTrend();
  });
  document.querySelector("#salesChartType")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-chart-type]"); if (!button) return;
    salesUi.chartType = button.dataset.chartType;
    document.querySelectorAll("[data-chart-type]").forEach((item) => item.classList.toggle("active", item === button));
    void renderSalesTrend();
  });
  document.querySelector("#salesTrendExpand")?.addEventListener("click", () => {
    const card = document.querySelector("#salesTrendCard");
    card?.classList.toggle("expanded");
    const button = document.querySelector("#salesTrendExpand");
    if (button) button.textContent = card?.classList.contains("expanded") ? "Close" : "Expand";
    void renderSalesTrend();
  });
  document.querySelector("#salesItemSearch")?.addEventListener("input", renderSalesItems);
  document.querySelector("#salesItemDepartment")?.addEventListener("change", () => {
    const category = document.querySelector("#salesItemCategory"), vendor = document.querySelector("#salesItemVendor");
    if (category) category.value = "";
    if (vendor) vendor.value = "";
    void renderSalesItems();
  });
  document.querySelector("#salesItemCategory")?.addEventListener("change", () => {
    const vendor = document.querySelector("#salesItemVendor");
    if (vendor) vendor.value = "";
    void renderSalesItems();
  });
  document.querySelector("#salesItemVendor")?.addEventListener("change", () => { void renderSalesItems(); });
  document.querySelector("#salesItemsClear")?.addEventListener("click", () => {
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    const values = {
      salesItemsStart: latest ? shiftIsoDate(latest, -29) : "",
      salesItemsEnd: latest,
      salesItemDepartment: "",
      salesItemCategory: "",
      salesItemVendor: "",
      salesItemSearch: "",
      salesItemSort: "sales",
    };
    Object.entries(values).forEach(([id, value]) => { const element = document.querySelector(`#${id}`); if (element) element.value = value; });
    salesUi.itemSort = "sales";
    salesUi.itemSortDirection = "desc";
    if (values.salesItemsStart && values.salesItemsEnd) void ensureRemoteSalesRange(values.salesItemsStart, values.salesItemsEnd).then(renderSalesItems);
    else void renderSalesItems();
  });
  document.querySelector("#salesItemSort")?.addEventListener("change", (event) => {
    salesUi.itemSort = event.target.value;
    salesUi.itemSortDirection = salesUi.itemSort === "product" ? "asc" : "desc";
    renderSalesItems();
  });
  document.querySelector(".sales-item-table thead")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-item-sort]");
    if (!button) return;
    const next = button.dataset.itemSort;
    if (salesUi.itemSort === next) salesUi.itemSortDirection = salesUi.itemSortDirection === "asc" ? "desc" : "asc";
    else { salesUi.itemSort = next; salesUi.itemSortDirection = ["product", "department", "category", "vendor"].includes(next) ? "asc" : "desc"; }
    renderSalesItems();
  });
  ["#salesItemsStart", "#salesItemsEnd"].forEach((selector) => document.querySelector(selector)?.addEventListener("change", () => {
    const start = document.querySelector("#salesItemsStart")?.value;
    const end = document.querySelector("#salesItemsEnd")?.value;
    if (start && end && start <= end) void ensureRemoteSalesRange(start, end).then(renderSalesItems);
  }));
  document.querySelector("#salesWorkspace")?.addEventListener("click", (event) => {
    const trend = event.target.closest("[data-open-trend], [data-open-category]");
    if (trend) {
      setSalesMode("overview");
      const group = document.querySelector("#salesTrendGroup"), input = document.querySelector("#salesTrendEntity");
      if (group) group.value = "category"; updateSalesTrendOptions(); if (input) input.value = trend.dataset.openTrend || trend.dataset.openCategory; void renderSalesTrend();
    }
    const department = event.target.closest("[data-open-department]");
    if (department) {
      setSalesMode("overview");
      const group = document.querySelector("#salesTrendGroup"), input = document.querySelector("#salesTrendEntity");
      if (group) group.value = "department"; updateSalesTrendOptions(); if (input) input.value = department.dataset.openDepartment; void renderSalesTrend();
    }
    const item = event.target.closest("[data-open-item]");
    if (item) {
      setSalesMode("overview");
      const group = document.querySelector("#salesTrendGroup"), input = document.querySelector("#salesTrendEntity");
      if (group) group.value = "item"; updateSalesTrendOptions(); if (input) input.value = item.dataset.openItem; void renderSalesTrend();
    }
  });
  window.addEventListener("resize", () => { if (activeTabName() === "dashboard") renderSalesWorkspace(); });

  dashboardDataStatus();
  setSalesMode("calendar");
})();

