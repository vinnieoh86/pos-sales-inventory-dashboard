(function () {
  const VERSION = "20261017-sales-studio-v9";
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
        <button type="button" data-sales-mode="trends">Trends</button>
        <button type="button" data-sales-mode="items">Items</button>
      </nav>
      <section class="sales-workspace" id="salesWorkspace">
        <section class="sales-mode-view" data-sales-mode-view="overview" hidden>
          <div class="sales-section-heading">
            <div><p class="eyebrow">Year-over-year</p><h2>Business pulse</h2></div>
            <small id="salesOverviewFocus">Selected day, month, and year compared with last year.</small>
          </div>
          <div class="yoy-card-grid" id="salesYoyCards"></div>
          <div class="sales-overview-grid">
            <article class="sales-card">
              <div class="sales-card__heading"><div><p class="eyebrow">Recent performance</p><h3>Last 14 completed days</h3></div></div>
              <canvas id="salesOverviewChart" height="210"></canvas>
            </article>
            <article class="sales-card">
              <div class="sales-card__heading"><div><p class="eyebrow">Sales mix</p><h3>Departments &amp; categories</h3></div></div>
              <div class="sales-department-tree" id="salesOverviewCategories"></div>
            </article>
          </div>
        </section>

        <section class="sales-mode-view active" data-sales-mode-view="calendar">
          <div class="calendar-toolbar">
            <button type="button" class="sales-icon-button" id="salesCalendarPrev" aria-label="Previous month">&#8249;</button>
            <div><p class="eyebrow">Monthly sales</p><h2 id="salesCalendarTitle">Calendar</h2></div>
            <button type="button" class="sales-icon-button" id="salesCalendarNext" aria-label="Next month">&#8250;</button>
            <button type="button" class="sales-secondary-button" id="salesCalendarToday">Latest month</button>
          </div>
          <div class="sales-calendar" id="salesCalendar"></div>
        </section>

        <section class="sales-mode-view" data-sales-mode-view="trends" hidden>
          <div class="sales-section-heading">
            <div><p class="eyebrow">Trend explorer</p><h2>Compare any category, vendor, or item</h2></div>
            <small>Use line view for direction and bar view for period-to-period magnitude.</small>
          </div>
          <div class="trend-explorer-controls">
            <label>Analyze<select id="salesTrendGroup"><option value="category">Category</option><option value="vendor">Vendor</option><option value="item">Item</option></select></label>
            <label>Selection<input id="salesTrendEntity" list="salesTrendEntities" placeholder="Choose or type a name" /><datalist id="salesTrendEntities"></datalist></label>
            <label>Metric<select id="salesTrendMetric"><option value="sales">Sales $</option><option value="units">Units</option><option value="profit">Profit</option></select></label>
            <label>Period<select id="salesTrendPeriod"><option value="30">30 days</option><option value="90">90 days</option><option value="365">12 months</option><option value="ytd">Year to date</option></select></label>
            <div class="sales-segmented" id="salesChartType"><button type="button" class="active" data-chart-type="line">Line</button><button type="button" data-chart-type="bar">Bar</button></div>
            <label class="sales-check"><input type="checkbox" id="salesTrendCompare" checked /> Compare last year</label>
          </div>
          <article class="sales-card sales-trend-card" id="salesTrendCard">
            <div class="sales-card__heading"><div><p class="eyebrow" id="salesTrendKicker">Category trend</p><h3 id="salesTrendTitle">Choose a category</h3></div><button type="button" class="sales-secondary-button" id="salesTrendExpand">Expand</button></div>
            <canvas id="salesTrendExplorerChart" height="300"></canvas>
            <div class="sales-chart-summary" id="salesTrendSummary"></div>
          </article>
        </section>

        <section class="sales-mode-view" data-sales-mode-view="items" hidden>
          <div class="sales-section-heading">
            <div><p class="eyebrow">Itemized sales</p><h2>Products for the selected dates</h2></div>
            <small id="salesItemsRange">Uses the current dashboard date range.</small>
          </div>
          <div class="sales-item-toolbar">
            <label>From<input type="date" id="salesItemsStart" /></label>
            <label>To<input type="date" id="salesItemsEnd" /></label>
            <input type="search" id="salesItemSearch" placeholder="Search product, UPC, vendor, or category" />
            <select id="salesItemSort"><option value="sales">Highest sales</option><option value="units">Most units</option><option value="profit">Highest profit</option><option value="cost">Highest total cost</option><option value="product">Product A-Z</option></select>
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
    status.textContent = `Synced through ${latest}`;
    range.textContent = earliest
      ? `History available ${earliest} through ${latest}. Selected periods load on demand.`
      : `Latest completed sales date: ${latest}.`;
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
    return `like.__ROLLUP__${kind}*`;
  }

  async function salesFetchRollups(kind, start, end) {
    const key = `${kind}|${start}|${end}`;
    if (salesUi.rollups.has(key)) return salesUi.rollups.get(key);
    const promise = directSelectRows("daily_sales", {
      select: "sales_date,code,product,vendor,category,department,qty,sales,cost_sold,profit",
      and: `(sales_date.gte.${start},sales_date.lte.${end})`,
      code: salesRollupCode(kind),
      order: "sales_date.asc",
    }).then((rows) => rows.map((row) => ({
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
    }))).catch(() => []);
    salesUi.rollups.set(key, promise);
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
      ctx.fillText(options.currency === false ? number.format(label) : compactCurrency(label), 4, y + 4);
    }
    if (!current.length) {
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
    const first = current[0]?.label || "";
    const last = current.at(-1)?.label || "";
    ctx.fillStyle = "#6d7773"; ctx.font = "11px Arial";
    ctx.fillText(first, left, height - 10);
    if (last && last !== first) ctx.fillText(last, Math.max(left, right - ctx.measureText(last).width), height - 10);
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
      { label: "Selected day", current: total(latest, latest), prior: total(previousLatest, previousLatest), detail: latest },
      { label: "Month to date", current: total(monthStart, latest), prior: total(previousMonthStart, previousLatest), detail: latestDate.toLocaleDateString("en-US", { month: "long", year: "numeric", timeZone: "UTC" }) },
      { label: "Year to date", current: total(yearStart, latest), prior: total(previousYearStart, previousLatest), detail: latest.slice(0, 4) },
    ];
    target.innerHTML = cards.map((card) => {
      const delta = salesDelta(card.current.sales, card.prior.sales);
      const hasPrior = card.prior.sales || card.prior.units;
      return `<article class="yoy-card">
        <div class="yoy-card__top"><span>${escapeHtml(card.label)}</span><small>${escapeHtml(card.detail)}</small></div>
        <strong>${currency.format(card.current.sales)}</strong>
        <div class="yoy-card__compare ${delta != null && delta < 0 ? "down" : "up"}">${hasPrior ? salesDeltaText(delta) : "Awaiting prior-year history"}</div>
        <small>${hasPrior ? `${currency.format(card.prior.sales)} last year` : `${number.format(card.current.units)} units`}</small>
      </article>`;
    }).join("");
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
      return `<details class="sales-department">
        <summary><span><b>${escapeHtml(department.name)}</b><small>${number.format(department.units)} units</small></span><strong>${currency.format(department.sales)}</strong></summary>
        <div class="sales-category-branches">${categories.map((category) => {
          const categoryRows = departmentRows.filter((row) => (cleanCell(row.category) || "Unassigned") === category.name);
          const items = salesAggregateGroup(categoryRows.map((row) => ({ ...row, item: row.product || row.code })), "item").sort((a, b) => b.sales - a.sales);
          return `<details class="sales-category-branch">
            <summary><span><b>${escapeHtml(category.name)}</b><small>${number.format(category.units)} units</small></span><strong>${currency.format(category.sales)}</strong></summary>
            <div class="sales-category-items">${items.slice(0, 100).map((item) => `<button type="button" data-open-item="${escapeHtml(item.name)}"><span><b>${escapeHtml(item.name)}</b><small>${number.format(item.units)} units</small></span><strong>${currency.format(item.sales)}</strong></button>`).join("")}</div>
          </details>`;
        }).join("")}</div>
      </details>`;
    }).join("") : `<div class="sales-empty">Department totals will appear after sales load.</div>`;
  }

  async function renderSalesOverview() {
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    if (!latest) return;
    const focus = salesUi.selectedDate || latest;
    const focusLabel = document.querySelector("#salesOverviewFocus");
    if (focusLabel) focusLabel.textContent = `${salesDate(focus).toLocaleDateString("en-US", { month: "long", day: "numeric", year: "numeric", timeZone: "UTC" })} with month/year comparisons.`;
    const previousYearStart = salesShiftYear(`${focus.slice(0, 4)}-01-01`, -1);
    const rollups = await salesFetchRollups("DAY", previousYearStart, focus);
    const fallback = salesAggregateDaily(salesRangeRows(previousYearStart, focus));
    const daily = rollups.length ? rollups : fallback;
    renderSalesYoyCards(daily, focus);
    const recentStart = shiftIsoDate(focus, -13);
    const recent = daily.filter((row) => row.date >= recentStart && row.date <= focus).map((row) => ({ label: row.date.slice(5), value: row.sales }));
    salesDrawChart(document.querySelector("#salesOverviewChart"), recent, [], { type: "bar", height: 210 });
    await ensureRemoteSalesRange(focus, focus, { silent: true });
    renderSalesDepartmentTree(salesRangeRows(focus, focus));
  }

  function renderSalesCalendarGrid(rows) {
    const target = document.querySelector("#salesCalendar");
    const title = document.querySelector("#salesCalendarTitle");
    if (!target || !salesUi.calendarMonth) return;
    const bounds = salesMonthBounds(salesUi.calendarMonth);
    const date = new Date(Date.UTC(bounds.year, bounds.month - 1, 1, 12));
    if (title) title.textContent = date.toLocaleDateString("en-US", { month: "long", year: "numeric", timeZone: "UTC" });
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
    const rollups = await salesFetchRollups("DAY", bounds.start, bounds.end);
    let rows = rollups;
    if (!rows.length) {
      await ensureRemoteSalesRange(bounds.start, bounds.end, { silent: true });
      if (requestedMonth !== salesUi.calendarMonth) return;
      rows = salesRangeRows(bounds.start, bounds.end);
    }
    renderSalesCalendarGrid(rows);
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

  function salesTrendDates(period, end) {
    if (period === "ytd") return { start: `${end.slice(0, 4)}-01-01`, end };
    return { start: shiftIsoDate(end, -(Math.max(1, Number(period)) - 1)), end };
  }

  function salesTrendOptions(group) {
    const field = group === "item" ? "product" : group;
    return [...new Set(state.rawSales.map((row) => cleanCell(row[field])).filter(Boolean))].sort(compareDisplayValue);
  }

  function updateSalesTrendOptions() {
    const group = document.querySelector("#salesTrendGroup")?.value || "category";
    const list = document.querySelector("#salesTrendEntities");
    if (!list) return;
    const values = salesTrendOptions(group);
    list.innerHTML = values.slice(0, 5000).map((value) => `<option value="${escapeHtml(value)}"></option>`).join("");
    const input = document.querySelector("#salesTrendEntity");
    if (input && !input.value && values.length) input.value = values[0];
  }

  function salesBucketSeries(rows, metric, monthly) {
    const grouped = new Map();
    rows.forEach((row) => {
      const key = monthly ? row.date.slice(0, 7) : row.date;
      const current = grouped.get(key) || { label: monthly ? key : row.date.slice(5), value: 0 };
      current.value += Number(row[metric] || 0);
      grouped.set(key, current);
    });
    return [...grouped.entries()].sort(([a], [b]) => a.localeCompare(b)).map(([, row]) => row);
  }

  async function renderSalesTrend() {
    const request = ++salesUi.trendRequest;
    const group = document.querySelector("#salesTrendGroup")?.value || "category";
    const entity = cleanCell(document.querySelector("#salesTrendEntity")?.value);
    const metric = document.querySelector("#salesTrendMetric")?.value || "sales";
    const period = document.querySelector("#salesTrendPeriod")?.value || "30";
    const compare = !!document.querySelector("#salesTrendCompare")?.checked;
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    const title = document.querySelector("#salesTrendTitle");
    const kicker = document.querySelector("#salesTrendKicker");
    if (kicker) kicker.textContent = `${group[0].toUpperCase()}${group.slice(1)} trend`;
    if (title) title.textContent = entity || `Choose a ${group}`;
    if (!latest || !entity) return salesDrawChart(document.querySelector("#salesTrendExplorerChart"), [], [], { height: 300 });
    const dates = salesTrendDates(period, latest);
    const priorDates = { start: salesShiftYear(dates.start, -1), end: salesShiftYear(dates.end, -1) };
    let rows = [];
    if (group === "item") {
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
    const monthly = period === "365" || period === "ytd";
    const currentRows = rows.filter((row) => row.date >= dates.start && row.date <= dates.end);
    const priorRows = compare ? rows.filter((row) => row.date >= priorDates.start && row.date <= priorDates.end) : [];
    const current = salesBucketSeries(currentRows, metric, monthly);
    const prior = salesBucketSeries(priorRows, metric, monthly);
    salesDrawChart(document.querySelector("#salesTrendExplorerChart"), current, prior, { type: salesUi.chartType, height: 300, currency: metric !== "units" });
    const currentTotal = current.reduce((sum, row) => sum + row.value, 0);
    const priorTotal = prior.reduce((sum, row) => sum + row.value, 0);
    const formatter = metric === "units" ? number : currency;
    const summary = document.querySelector("#salesTrendSummary");
    if (summary) summary.innerHTML = `<b>${formatter.format(currentTotal)}</b><span>selected period</span>${compare ? `<b class="${currentTotal >= priorTotal ? "up" : "down"}">${salesDeltaText(salesDelta(currentTotal, priorTotal))}</b><span>vs last year (${formatter.format(priorTotal)})</span>` : ""}`;
  }

  async function renderSalesItems() {
    const body = document.querySelector("#salesItemRows");
    if (!body) return;
    const start = document.querySelector("#salesItemsStart")?.value || els.startDate.value || state.dates[0] || "";
    const end = document.querySelector("#salesItemsEnd")?.value || els.endDate.value || state.dates.at(-1) || "";
    const range = document.querySelector("#salesItemsRange");
    if (range) range.textContent = start && end ? `${start} through ${end}` : "Choose a date range.";
    const query = cleanCell(document.querySelector("#salesItemSearch")?.value).toLowerCase();
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
    let rows = [...map.values()].map((row) => ({ ...row, price: row.units ? row.sales / row.units : 0, unitCost: row.units ? row.cost / row.units : 0 }))
      .filter((row) => !query || [row.code, row.product, row.department, row.category, row.vendor].join(" ").toLowerCase().includes(query));
    const direction = salesUi.itemSortDirection === "asc" ? 1 : -1;
    rows.sort((a, b) => ["product", "department", "category", "vendor"].includes(sort)
      ? direction * cleanCell(a[sort]).localeCompare(cleanCell(b[sort]))
      : direction * (Number(a[sort] || 0) - Number(b[sort] || 0)));
    const totals = salesTotals(rows);
    const finalizedDays = await salesFetchRollups("DAY", start, end);
    const finalizedSubtotal = finalizedDays.length ? salesTotals(finalizedDays).sales : totals.sales;
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
    if (mode === "trends") { updateSalesTrendOptions(); void renderSalesTrend(); }
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
    if (salesUi.mode === "overview") void renderSalesOverview();
    else if (salesUi.mode === "calendar") void renderSalesCalendar();
    else if (salesUi.mode === "trends") { updateSalesTrendOptions(); void renderSalesTrend(); }
    else if (salesUi.mode === "items") void renderSalesItems();
  }

  document.querySelector("#salesModeTabs")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-sales-mode]");
    if (button) setSalesMode(button.dataset.salesMode);
  });
  document.querySelector("#salesCalendar")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-sales-date]");
    if (!button) return;
    salesUi.selectedDate = button.dataset.salesDate;
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
  ["#salesTrendGroup", "#salesTrendMetric", "#salesTrendPeriod", "#salesTrendCompare"].forEach((selector) => document.querySelector(selector)?.addEventListener("change", () => {
    if (selector === "#salesTrendGroup") { const input = document.querySelector("#salesTrendEntity"); if (input) input.value = ""; updateSalesTrendOptions(); }
    void renderSalesTrend();
  }));
  document.querySelector("#salesTrendEntity")?.addEventListener("change", () => { void renderSalesTrend(); });
  document.querySelector("#salesTrendEntity")?.addEventListener("keydown", (event) => { if (event.key === "Enter") void renderSalesTrend(); });
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
    const trend = event.target.closest("[data-open-trend]");
    if (trend) {
      setSalesMode("trends");
      const group = document.querySelector("#salesTrendGroup"), input = document.querySelector("#salesTrendEntity");
      if (group) group.value = "category"; updateSalesTrendOptions(); if (input) input.value = trend.dataset.openTrend; void renderSalesTrend();
    }
    const item = event.target.closest("[data-open-item]");
    if (item) {
      setSalesMode("trends");
      const group = document.querySelector("#salesTrendGroup"), input = document.querySelector("#salesTrendEntity");
      if (group) group.value = "item"; updateSalesTrendOptions(); if (input) input.value = item.dataset.openItem; void renderSalesTrend();
    }
  });
  window.addEventListener("resize", () => { if (activeTabName() === "dashboard") renderSalesWorkspace(); });

  dashboardDataStatus();
  setSalesMode("calendar");
})();

