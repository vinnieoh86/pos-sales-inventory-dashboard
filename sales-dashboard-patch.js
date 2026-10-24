(function () {
  const VERSION = "20261024-sales-studio-v31";
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

  function upgradeSalesStudioLayout() {
    const overview = document.querySelector('[data-sales-mode-view="overview"]');
    const calendar = document.querySelector('[data-sales-mode-view="calendar"]');
    const itemsView = document.querySelector('[data-sales-mode-view="items"]');
    const hero = document.querySelector("#dashboardOverview");
    const tabs = document.querySelector("#salesModeTabs");
    const health = document.querySelector("#salesDataHealth");
    const controls = document.querySelector(".sales-overview-controls");
    if (!overview || !calendar || !controls || overview.dataset.masterLayout === "true") return;
    overview.dataset.masterLayout = "true";

    hero?.remove();
    tabs?.querySelector('[data-sales-mode="items"]')?.remove();
    const calendarTab = tabs?.querySelector('[data-sales-mode="calendar"]');
    const overviewTab = tabs?.querySelector('[data-sales-mode="overview"]');
    if (tabs && calendarTab && overviewTab) {
      calendarTab.classList.add("active");
      overviewTab.classList.remove("active");
      overviewTab.textContent = "Overview / Charts";
      tabs.append(calendarTab, overviewTab);
    }

    const workspace = document.querySelector("#salesWorkspace");
    if (workspace && tabs && health) {
      const dock = document.createElement("section");
      dock.className = "sales-top-dock";
      dock.id = "salesTopDock";
      workspace.parentNode.insertBefore(dock, workspace);
      dock.append(tabs, health);
      health.querySelector(".sales-health-heading h2")?.replaceChildren(document.createTextNode("Data health"));
      const healthEyebrow = health.querySelector(".sales-health-heading .eyebrow");
      if (healthEyebrow) healthEyebrow.textContent = "Sync & recovery";
    }

    controls.classList.add("sales-master-controls");
    controls.innerHTML = `
      <label>Period<select id="salesPeriodUnit"><option value="day">Day</option><option value="week">Week</option><option value="month" selected>Month</option><option value="quarter">Quarter</option><option value="year">Year</option></select></label>
      <div class="sales-period-nav"><button type="button" class="sales-icon-button" id="salesTrendPrev" aria-label="Previous period">&#8249;</button><label>Start<input type="date" id="salesTrendStart" /></label><label>End<input type="date" id="salesTrendDate" /></label><button type="button" class="sales-icon-button" id="salesTrendNext" aria-label="Next period">&#8250;</button></div>
      <label>Chart view<select id="salesTrendBucket"><option value="day" selected>Daily</option><option value="week">Weekly</option><option value="month">Monthly</option></select></label>
      <details class="sales-vendor-picker" id="salesVendorPicker"><summary>Vendors <span id="salesVendorCount">All</span></summary><div><input type="search" id="salesVendorSearch" placeholder="Find vendor" autocomplete="off" data-lpignore="true"/><div id="salesVendorOptions"></div></div></details>
      <button type="button" class="sales-secondary-button" id="salesVendorCompare">Compare vendors</button>
      <label>Metric<select id="salesTrendMetric"><option value="sales">Net sales $</option><option value="units">Units</option><option value="profit">Profit</option><option value="margin">Margin %</option></select></label>
      <div class="sales-segmented" id="salesChartType"><button type="button" class="active" data-chart-type="bar">Bar</button><button type="button" data-chart-type="line">Line</button></div>
      <label class="sales-check"><input type="checkbox" id="salesTrendCompare" checked /> Compare prior year</label>
      <button type="button" class="sales-secondary-button sales-clear-button" id="salesTrendClear">Clear all</button>`;

    const heading = overview.querySelector(".sales-section-heading");
    if (heading) {
      heading.innerHTML = `<div><p class="eyebrow">Executive analysis</p><h2>Business pulse</h2></div><div class="sales-drill-path" id="salesDrillPath"><button type="button" data-drill-level="all">All departments</button></div>`;
    }

    const trendCard = document.querySelector("#salesTrendCard");
    const mixGrid = document.querySelector(".sales-overview-grid");
    if (trendCard && mixGrid) {
      const analysisGrid = document.createElement("div");
      analysisGrid.className = "sales-analysis-grid";
      trendCard.parentNode.insertBefore(analysisGrid, trendCard);
      analysisGrid.append(trendCard, mixGrid);
      const mixTitle = mixGrid.querySelector("h3");
      if (mixTitle) mixTitle.textContent = "Browse departments, categories & items";
    }

    document.querySelector("#salesExecutiveSummary")?.remove();
    const cards = document.querySelector("#salesYoyCards");
    if (cards) {
      cards.className = "sales-period-kpis sales-global-kpis";
      const dock = document.querySelector("#salesTopDock");
      dock?.insertAdjacentElement("afterend", cards);
    }

    const dayNav = document.createElement("section");
    dayNav.className = "sales-day-nav";
    dayNav.id = "salesDayNav";
    dayNav.hidden = true;
    dayNav.innerHTML = `<button type="button" class="sales-secondary-button" id="salesDayBack">&#8249; Back to calendar</button><div><p class="eyebrow">Daily sales detail</p><h2 id="salesCalendarDayTitle">Choose a calendar day</h2><p id="salesDayComparison"></p></div>`;
    document.querySelector("#salesTopDock")?.insertAdjacentElement("afterend", dayNav);

    const dayPage = document.createElement("section");
    dayPage.className = "sales-mode-view sales-day-page";
    dayPage.dataset.salesModeView = "day";
    dayPage.id = "salesDayPage";
    dayPage.hidden = true;
    dayPage.innerHTML = `<section class="sales-calendar-selection" id="salesCalendarSelection"><div class="sales-section-heading"><div><p class="eyebrow">Department breakdown</p><h2>Sales mix for this day</h2></div><button type="button" class="sales-secondary-button" id="salesCalendarClearFilters">Clear department filter</button></div><div class="sales-calendar-mix" id="salesCalendarMix"></div></section>`;
    workspace?.append(dayPage);

    if (itemsView) {
      itemsView.hidden = false;
      itemsView.removeAttribute("data-sales-mode-view");
      itemsView.classList.add("sales-items-inline");
      const heading = itemsView.querySelector("h2");
      if (heading) heading.textContent = "Items sold in this selection";
      const itemToolbar = itemsView.querySelector(".sales-item-toolbar");
      if (itemToolbar) itemToolbar.innerHTML = `<label class="sales-item-search">Search sold items<input type="search" id="salesItemSearch" list="salesSoldItems" placeholder="Product, UPC or item number" autocomplete="off" name="salesSoldItemLookup" data-lpignore="true" data-1p-ignore="true" data-form-type="other"/><datalist id="salesSoldItems"></datalist></label><label class="sales-item-history-range" hidden>History<select id="salesItemHistoryRange"><option value="12">Last 12 months</option><option value="24" selected>Last 24 months</option></select></label><label>Sort<select id="salesItemSort" aria-label="Sort items"><option value="sales">Highest sales</option><option value="units">Most units</option><option value="profit">Highest profit</option><option value="cost">Highest total cost</option><option value="product">Product A-Z</option></select></label><button type="button" class="sales-secondary-button" id="salesItemsClear">Clear search</button>`;
      const headRow = itemsView.querySelector(".sales-item-table thead tr");
      if (headRow) headRow.id = "salesItemHeadRow";
      overview.append(itemsView);
    }
  }

  upgradeSalesStudioLayout();

  function syncSalesStickyOffsets() {
    const command = document.querySelector(".command-bar");
    const dock = document.querySelector("#salesTopDock");
    document.documentElement.style.setProperty("--sales-command-height", `${Math.ceil(command?.getBoundingClientRect().height || 0)}px`);
    document.documentElement.style.setProperty("--sales-dock-height", `${Math.ceil(dock?.getBoundingClientRect().height || 0)}px`);
  }

  function syncSalesModeLayout(mode) {
    const dock = document.querySelector("#salesTopDock");
    const kpis = document.querySelector("#salesYoyCards");
    const dayPage = document.querySelector("#salesDayPage");
    const daySelection = document.querySelector("#salesCalendarSelection");
    const items = document.querySelector(".sales-items-inline");
    const overview = document.querySelector('[data-sales-mode-view="overview"]');
    if (!dock || !kpis || !dayPage || !daySelection || !items || !overview) return;
    if (mode === "day") {
      let summary = dayPage.querySelector(".sales-day-summary-grid");
      if (!summary) {
        summary = document.createElement("section");
        summary.className = "sales-day-summary-grid";
        dayPage.prepend(summary);
      }
      summary.append(kpis, daySelection);
      dayPage.append(items);
    } else if (mode === "overview") {
      const controls = overview.querySelector(".sales-master-controls");
      if (controls) controls.insertAdjacentElement("afterend", kpis);
      else overview.prepend(kpis);
      dayPage.prepend(daySelection);
      overview.append(items);
    } else {
      dock.insertAdjacentElement("afterend", kpis);
      dayPage.prepend(daySelection);
      overview.append(items);
    }
    requestAnimationFrame(syncSalesStickyOffsets);
  }

  const salesStickyObserver = typeof ResizeObserver === "function" ? new ResizeObserver(syncSalesStickyOffsets) : null;
  salesStickyObserver?.observe(document.querySelector(".command-bar"));
  salesStickyObserver?.observe(document.querySelector("#salesTopDock"));
  window.addEventListener("load", syncSalesStickyOffsets, { once: true });
  requestAnimationFrame(syncSalesStickyOffsets);

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

  function salesDisplayTime(value) {
    const raw = cleanCell(value);
    if (!raw) return "Daily total";
    const date = new Date(raw);
    if (Number.isFinite(date.getTime())) return date.toLocaleTimeString("en-US", { hour: "numeric", minute: "2-digit" });
    const match = raw.match(/(?:^|\s)(\d{1,2}:\d{2}(?::\d{2})?\s*(?:AM|PM)?)/i);
    return match ? match[1].toUpperCase() : raw;
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
    const normalized = originalNormalizeSalesRow(patched, date);
    const timestamp = cleanCell(row.Timestamp || row.TIMESTAMP || row.DateTime || row.DATETIME || row.SaleTime || row.SALE_TIME || row.Time || row.TIME);
    return timestamp ? { ...normalized, timestamp } : normalized;
  };

  const originalHydrateSalesFromSupabase = hydrateSalesFromSupabase;
  hydrateSalesFromSupabase = function patchedHydrateSalesFromSupabase(row) {
    const hydrated = originalHydrateSalesFromSupabase(row);
    const timestamp = cleanCell(row.sale_timestamp || row.sales_timestamp || row.sold_at || row.sale_time || row.timestamp);
    return timestamp ? { ...hydrated, timestamp } : hydrated;
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
    mode: "",
    calendarMonth: "",
    selectedDate: "",
    calendarDepartment: "",
    calendarCategory: "",
    chartType: "bar",
    itemSort: "sales",
    itemSortDirection: "desc",
    drill: { department: "", category: "", item: "" },
    vendors: new Set(),
    vendorCompare: false,
    periodUnit: "month",
    trendRangeMonths: 1,
    trendBucket: "day",
    itemHistoryMonths: 24,
    trendRequest: 0,
    rollups: new Map(),
    itemDetails: new Map(),
    itemRequest: 0,
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

  function salesPeriodBounds(unit, anchor) {
    const date = salesDate(anchor);
    const year = date.getUTCFullYear();
    const month = date.getUTCMonth();
    if (unit === "day") return { start: anchor, end: anchor };
    if (unit === "week") {
      const start = new Date(date);
      const weekday = start.getUTCDay();
      start.setUTCDate(start.getUTCDate() + (weekday === 0 ? -6 : 1 - weekday));
      return { start: salesIso(start), end: shiftIsoDate(salesIso(start), 6) };
    }
    if (unit === "quarter") {
      const quarterMonth = Math.floor(month / 3) * 3;
      return {
        start: salesIso(new Date(Date.UTC(year, quarterMonth, 1, 12))),
        end: salesIso(new Date(Date.UTC(year, quarterMonth + 3, 0, 12))),
      };
    }
    if (unit === "year") return { start: `${year}-01-01`, end: `${year}-12-31` };
    return salesMonthBounds(`${year}-${String(month + 1).padStart(2, "0")}`);
  }

  function setSalesPeriod(anchor, unit = salesUi.periodUnit) {
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || salesIso(new Date());
    const safeAnchor = anchor && anchor <= latest ? anchor : latest;
    const bounds = salesPeriodBounds(unit, safeAnchor);
    const start = document.querySelector("#salesTrendStart");
    const end = document.querySelector("#salesTrendDate");
    if (start) start.value = bounds.start;
    if (end) end.value = bounds.end;
    salesUi.periodUnit = unit;
    return bounds;
  }

  function selectedSalesPeriod(latest) {
    const start = document.querySelector("#salesTrendStart")?.value;
    const selectedEnd = document.querySelector("#salesTrendDate")?.value;
    if (!start || !selectedEnd || start > selectedEnd) return salesPeriodBounds(salesUi.periodUnit, latest);
    return { start, end: selectedEnd > latest ? latest : selectedEnd, selectedEnd };
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
      select: "*",
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

  async function salesFetchAdjustments(start, end) {
    const [grossRows, discountRows, returnRows] = await Promise.all([
      salesFetchRollups("GROSS", start, end),
      salesFetchRollups("DISCOUNT", start, end),
      salesFetchRollups("RETURNS", start, end),
    ]);
    return {
      gross: salesTotals(grossRows).sales,
      discount: salesTotals(discountRows).sales,
      returns: salesTotals(returnRows).sales,
      available: grossRows.length > 0 || discountRows.length > 0 || returnRows.length > 0,
    };
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
      const priorText = prior ? `<span class="prior-dot"></span>${escapeHtml(options.priorLabel || "Prior year")}: <b>${formatter.format(prior.value)}</b>${prior.units == null ? "" : ` · ${number.format(prior.units)} units`}` : "";
      const currentText = current ? `<span class="current-dot"></span>${escapeHtml(options.currentLabel || "Current")}: <b>${formatter.format(current.value)}</b>${current.units == null ? "" : ` · ${number.format(current.units)} units`}` : "";
      const delta = current && prior && Number(prior.value) ? salesDelta(Number(current.value), Number(prior.value)) : null;
      tooltip.innerHTML = `<strong>${escapeHtml(label)}</strong>${priorText ? `<br>${priorText}` : ""}${currentText ? `<br>${currentText}` : ""}${delta == null ? "" : `<br><em class="${delta < 0 ? "down" : "up"}">${salesDeltaText(delta)} vs prior year</em>`}`;
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
      prior.forEach((row, index) => { const x = xFor(index) - bar; const y = yFor(row.value); ctx.fillStyle = "#9cb7db"; ctx.fillRect(x, y, bar, bottom - y); });
      current.forEach((row, index) => { const x = xFor(index) + 1; const y = yFor(row.value); ctx.fillStyle = "#16835b"; ctx.fillRect(x, y, bar, bottom - y); });
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

  const SALES_TREND_COLORS = ["#16835b", "#4f6f91", "#8b5fb5", "#d16b43", "#c39a24", "#168b91", "#b54f73", "#637a36", "#805e49", "#667085"];

  function salesLightenHex(hex, amount = .68) {
    const value = String(hex || "#16835b").replace("#", "");
    const normalized = value.length === 3 ? value.split("").map((part) => `${part}${part}`).join("") : value;
    const numberValue = Number.parseInt(normalized, 16);
    if (!Number.isFinite(numberValue)) return "#b7d8ca";
    const channel = (shift) => Math.round(((numberValue >> shift) & 255) + (255 - ((numberValue >> shift) & 255)) * amount);
    return `rgb(${channel(16)}, ${channel(8)}, ${channel(0)})`;
  }

  function salesTrendBucketKey(iso, bucket) {
    if (bucket === "day") return iso.slice(0, 10);
    if (bucket === "month") return iso.slice(0, 7);
    const date = salesDate(iso);
    const day = date.getUTCDay();
    date.setUTCDate(date.getUTCDate() + (day === 0 ? -6 : 1 - day));
    return salesIso(date);
  }

  function salesTrendBucketLabel(key, bucket) {
    if (bucket === "day") return salesDate(key).toLocaleDateString("en-US", { month: "numeric", day: "numeric", timeZone: "UTC" });
    if (bucket === "month") return salesDate(`${key}-01`).toLocaleDateString("en-US", { month: "short", year: "2-digit", timeZone: "UTC" });
    return salesDate(key).toLocaleDateString("en-US", { month: "short", day: "numeric", timeZone: "UTC" });
  }

  function salesTrendBucketKeys(dates, bucket) {
    const keys = [];
    if (bucket === "day") {
      for (let cursor = dates.start; cursor <= dates.end; cursor = shiftIsoDate(cursor, 1)) keys.push(cursor);
      return keys;
    }
    if (bucket === "month") {
      let cursor = `${dates.start.slice(0, 7)}-01`;
      const last = `${dates.end.slice(0, 7)}-01`;
      while (cursor <= last) {
        keys.push(cursor.slice(0, 7));
        const date = salesDate(cursor); date.setUTCMonth(date.getUTCMonth() + 1); cursor = `${salesIso(date).slice(0, 7)}-01`;
      }
      return keys;
    }
    let cursor = salesTrendBucketKey(dates.start, "week");
    const last = salesTrendBucketKey(dates.end, "week");
    while (cursor <= last) { keys.push(cursor); cursor = shiftIsoDate(cursor, 7); }
    return keys;
  }

  function salesTrendSeries(rows, priorRows, metric, bucket, dimension, fallbackName, dates, compare) {
    const keys = salesTrendBucketKeys(dates, bucket);
    const groupName = (row) => dimension ? (cleanCell(row[dimension]) || "Unassigned") : fallbackName;
    const totals = new Map();
    rows.forEach((row) => {
      const name = groupName(row);
      const current = totals.get(name) || 0;
      totals.set(name, current + (metric === "margin" ? Number(row.sales || 0) : salesMetricValue(row, metric)));
    });
    let names = [...totals.keys()];
    if (dimension === "department") names.sort((a, b) => {
      const ai = GKPOS_DEPARTMENT_ORDER.indexOf(a.toUpperCase()), bi = GKPOS_DEPARTMENT_ORDER.indexOf(b.toUpperCase());
      return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi) || a.localeCompare(b);
    });
    else names.sort((a, b) => Number(totals.get(b) || 0) - Number(totals.get(a) || 0));
    names = names.slice(0, dimension === "category" || dimension === "vendor" ? 10 : 12);

    const build = (source, prior = false) => names.map((name, index) => {
      const buckets = new Map();
      source.filter((row) => groupName(row) === name).forEach((row) => {
        let key = salesTrendBucketKey(row.date, bucket);
        if (prior) key = salesTrendBucketKey(salesShiftYear(row.date, 1), bucket);
        const current = buckets.get(key) || { sales: 0, profit: 0, units: 0 };
        current.sales += Number(row.sales || 0); current.profit += Number(row.profit || 0); current.units += Number(row.units || row.qty || 0);
        buckets.set(key, current);
      });
      return {
        name,
        color: SALES_TREND_COLORS[index % SALES_TREND_COLORS.length],
        prior,
        values: keys.map((key) => {
          const value = buckets.get(key) || { sales: 0, profit: 0, units: 0 };
          return metric === "margin" ? (value.sales ? (value.profit / value.sales) * 100 : 0) : Number(value[metric] || 0);
        }),
      };
    });
    return { labels: keys.map((key) => salesTrendBucketLabel(key, bucket)), series: [...(compare ? build(priorRows, true) : []), ...build(rows, false)] };
  }

  function salesDepartmentComparisonSeries(rows, priorRows, metric, compare) {
    const currentGroups = new Map(salesAggregateGroup(rows, "department").map((row) => [row.name, row]));
    const priorGroups = new Map(salesAggregateGroup(priorRows, "department").map((row) => [row.name, row]));
    const labels = [...new Set([...currentGroups.keys(), ...priorGroups.keys()])]
      .filter((name) => name && name !== "Unassigned")
      .sort((a, b) => {
        const ai = GKPOS_DEPARTMENT_ORDER.indexOf(a.toUpperCase());
        const bi = GKPOS_DEPARTMENT_ORDER.indexOf(b.toUpperCase());
        return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi) || a.localeCompare(b);
      });
    const valueFor = (map, name) => {
      const row = map.get(name) || { sales: 0, profit: 0, units: 0 };
      return metric === "margin" ? (row.sales ? (row.profit / row.sales) * 100 : 0) : Number(row[metric] || 0);
    };
    const base = { name: "Department total", color: "#16835b" };
    return {
      labels,
      series: [
        ...(compare ? [{ ...base, prior: true, values: labels.map((name) => valueFor(priorGroups, name)) }] : []),
        { ...base, prior: false, values: labels.map((name) => valueFor(currentGroups, name)) },
      ],
    };
  }

  function bindSalesMultiChartHover(canvas, labels, series, xPositions, options = {}, hitAreas = []) {
    if (!canvas) return;
    let tooltip = document.querySelector("#salesChartTooltip");
    if (!tooltip) { tooltip = document.createElement("div"); tooltip.id = "salesChartTooltip"; tooltip.className = "sales-chart-tooltip"; tooltip.hidden = true; document.body.append(tooltip); }
    let marker = document.querySelector("#salesChartMarker");
    if (!marker) { marker = document.createElement("div"); marker.id = "salesChartMarker"; marker.className = "sales-chart-marker"; marker.hidden = true; document.body.append(marker); }
    let hoverFrame = 0;
    const formatter = options.percent ? { format: (value) => `${Number(value || 0).toFixed(1)}%` } : (options.currency === false ? number : currency);
    const hitAt = (event) => {
      const rect = canvas.getBoundingClientRect();
      const x = event.clientX - rect.left;
      const y = event.clientY - rect.top;
      return [...hitAreas].reverse().find((hit) => hit.kind === "bar" && x >= hit.x - 3 && x <= hit.x + hit.width + 3 && y >= hit.y - 3 && y <= hit.y + hit.height + 3)
        || hitAreas.reduce((best, hit) => {
          if (hit.kind !== "point") return best;
          const distance = Math.hypot(x - hit.x, y - hit.y);
          return distance <= 13 && (!best || distance < best.distance) ? { ...hit, distance } : best;
        }, null);
    };
    const showHit = (hit, locked = false) => {
      if (!hit) return;
      const rect = canvas.getBoundingClientRect();
      const markerWidth = hit.kind === "bar" ? Math.max(4, hit.width) : 14;
      const markerHeight = hit.kind === "bar" ? Math.max(5, hit.height) : 14;
      marker.hidden = false;
      marker.className = `sales-chart-marker ${hit.kind}${locked ? " locked" : ""}`;
      marker.style.left = `${rect.left + hit.x - (hit.kind === "point" ? 7 : 0)}px`;
      marker.style.top = `${rect.top + hit.y - (hit.kind === "point" ? 7 : 0)}px`;
      marker.style.width = `${markerWidth}px`;
      marker.style.height = `${markerHeight}px`;
      if (hit.comparison) {
        const delta = Number(hit.priorValue) ? salesDelta(Number(hit.currentValue), Number(hit.priorValue)) : null;
        tooltip.innerHTML = `<strong>${escapeHtml(hit.series.name)} · ${escapeHtml(hit.label)}</strong><br><span style="display:inline-block;width:.58rem;height:.58rem;margin-right:.35rem;border-radius:2px;background:${salesLightenHex(hit.series.color)}"></span>${escapeHtml(options.priorLabel || "Prior")}: <b>${formatter.format(hit.priorValue)}</b><br><span style="display:inline-block;width:.58rem;height:.58rem;margin-right:.35rem;border-radius:2px;background:${hit.series.color}"></span>${escapeHtml(options.currentLabel || "Current")}: <b>${formatter.format(hit.currentValue)}</b>${delta == null ? "" : `<br><em class="${delta < 0 ? "down" : "up"}">${salesDeltaText(delta)} · ${formatter.format(Number(hit.currentValue) - Number(hit.priorValue))}</em>`}${locked ? `<br><em>Locked · click another bar to move</em>` : `<br><em>Click to lock this comparison</em>`}`;
      } else {
        tooltip.innerHTML = `<strong>${escapeHtml(hit.label)}</strong><br><span style="display:inline-block;width:.58rem;height:.58rem;margin-right:.35rem;border-radius:50%;background:${hit.series.color};opacity:${hit.series.prior ? .5 : 1}"></span>${escapeHtml(hit.series.name)}${hit.series.prior ? " (prior year)" : ""}: <b>${formatter.format(hit.value)}</b>${locked ? `<br><em>Locked · click another bar to move</em>` : `<br><em>Click to lock this bar</em>`}`;
      }
      tooltip.hidden = false;
      tooltip.classList.toggle("locked", locked);
      tooltip.style.left = `${Math.min(rect.left + hit.x + markerWidth + 12, window.innerWidth - 300)}px`;
      tooltip.style.top = `${Math.max(8, rect.top + hit.y - 78)}px`;
    };
    canvas.onmousemove = (event) => {
      if (canvas._salesLockedHit) { showHit(canvas._salesLockedHit, true); return; }
      cancelAnimationFrame(hoverFrame);
      hoverFrame = requestAnimationFrame(() => {
        const hit = hitAt(event);
        if (hit) { canvas.style.cursor = "pointer"; showHit(hit, false); }
        else { canvas.style.cursor = "default"; tooltip.hidden = true; marker.hidden = true; }
      });
    };
    canvas.onclick = (event) => {
      const hit = hitAt(event);
      if (!hit) return;
      const same = canvas._salesLockedHit?.key === hit.key;
      canvas._salesLockedHit = same ? null : hit;
      if (same) { tooltip.hidden = true; marker.hidden = true; }
      else showHit(hit, true);
    };
    canvas.onmouseleave = () => {
      cancelAnimationFrame(hoverFrame);
      if (canvas._salesLockedHit) showHit(canvas._salesLockedHit, true);
      else { tooltip.hidden = true; marker.hidden = true; }
    };
  }

  function salesDrawMultiSeriesChart(canvas, labels, series, options = {}) {
    const setup = salesSetCanvas(canvas, Number(options.height || 350));
    if (!setup) return;
    const { ctx, width, height } = setup;
    const left = 58, right = width - 18, top = 20, bottom = height - 42;
    const values = series.flatMap((item) => item.values).map(Number);
    const max = Math.max(...values, 1);
    ctx.strokeStyle = "#dfe7e3"; ctx.fillStyle = "#6d7773"; ctx.font = "11px Arial";
    for (let i = 0; i < 5; i += 1) {
      const y = top + ((bottom - top) * i) / 4;
      ctx.beginPath(); ctx.moveTo(left, y); ctx.lineTo(right, y); ctx.stroke();
      const value = max * (1 - i / 4);
      ctx.fillText(options.percent ? `${value.toFixed(0)}%` : (options.currency === false ? number.format(value) : compactCurrency(value)), 4, y + 4);
    }
    if (!labels.length || !series.length) { ctx.textAlign = "center"; ctx.fillText("No data for this selection yet.", width / 2, height / 2); ctx.textAlign = "left"; return; }
    const count = labels.length;
    const slot = (right - left) / Math.max(count, 1);
    const xPositions = labels.map((_, index) => left + slot * index + slot / 2);
    const yFor = (value) => bottom - ((bottom - top) * Number(value || 0)) / max;
    const hitAreas = [];
    if (options.type === "bar" || options.type === "combo") {
      const logicalNames = [...new Set(series.map((item) => item.name))];
      const bar = Math.max(3, Math.min(24, slot * .82 / Math.max(logicalNames.length, 1)));
      logicalNames.forEach((name, logicalIndex) => {
        const currentSeries = series.find((item) => item.name === name && !item.prior);
        const priorSeries = series.find((item) => item.name === name && item.prior);
        const color = currentSeries?.color || priorSeries?.color || SALES_TREND_COLORS[logicalIndex % SALES_TREND_COLORS.length];
        labels.forEach((label, index) => {
          const currentValue = Number(currentSeries?.values[index] || 0);
          const priorValue = Number(priorSeries?.values[index] || 0);
          const baseX = xPositions[index] - (bar * logicalNames.length) / 2 + logicalIndex * bar;
          const fullWidth = Math.max(2, bar - 1);
          const overlayWidth = Math.max(2, fullWidth * .58);
          const currentY = yFor(currentValue);
          const priorY = yFor(priorValue);
          if (currentValue) {
            ctx.globalAlpha = .95;
            ctx.fillStyle = color;
            ctx.fillRect(baseX, currentY, fullWidth, bottom - currentY);
          }
          if (priorValue && options.type === "bar") {
            ctx.globalAlpha = .96;
            ctx.fillStyle = salesLightenHex(color);
            ctx.fillRect(baseX + (fullWidth - overlayWidth) / 2, priorY, overlayWidth, bottom - priorY);
          }
          if (currentValue || priorValue) {
            const topY = Math.min(currentValue ? currentY : bottom, priorValue ? priorY : bottom);
            hitAreas.push({ kind: "bar", comparison: true, x: baseX, y: topY, width: fullWidth, height: bottom - topY, label, value: currentValue, currentValue, priorValue, series: currentSeries || priorSeries || { name, color }, key: `${index}|${name}` });
          }
        });
      });
      ctx.globalAlpha = 1;
      if (options.type === "combo") {
        const priorSeries = series.find((item) => item.prior);
        if (priorSeries) {
          ctx.strokeStyle = "#d66b6b";
          ctx.fillStyle = "#fff";
          ctx.lineWidth = 2.5;
          ctx.setLineDash([]);
          ctx.beginPath();
          priorSeries.values.forEach((value, index) => {
            const x = xPositions[index];
            const y = yFor(value);
            if (!index) ctx.moveTo(x, y); else ctx.lineTo(x, y);
          });
          ctx.stroke();
          priorSeries.values.forEach((value, index) => {
            if (!Number(value || 0)) return;
            const x = xPositions[index], y = yFor(value);
            ctx.beginPath(); ctx.arc(x, y, 3.5, 0, Math.PI * 2); ctx.fill(); ctx.stroke();
          });
        }
      }
    } else {
      series.forEach((item, seriesIndex) => {
        ctx.strokeStyle = item.color; ctx.globalAlpha = item.prior ? .38 : 1; ctx.lineWidth = item.prior ? 1.6 : 2.5; ctx.setLineDash(item.prior ? [6, 5] : []);
        ctx.beginPath(); item.values.forEach((value, index) => { const x = xPositions[index], y = yFor(value); if (!index) ctx.moveTo(x, y); else ctx.lineTo(x, y); if (Number(value || 0) !== 0) hitAreas.push({ kind: "point", x, y, label: labels[index], value: Number(value || 0), series: item, key: `${index}|${seriesIndex}` }); }); ctx.stroke();
      });
      ctx.globalAlpha = 1; ctx.setLineDash([]);
    }
    ctx.fillStyle = "#6d7773"; ctx.font = "11px Arial"; ctx.textAlign = "center";
    const step = Math.max(1, Math.ceil(labels.length / 12));
    labels.forEach((label, index) => { if (index % step === 0 || index === labels.length - 1) ctx.fillText(label, xPositions[index], height - 12); });
    ctx.textAlign = "left";
    if (canvas._salesLockedHit && !hitAreas.some((hit) => hit.key === canvas._salesLockedHit.key)) canvas._salesLockedHit = null;
    bindSalesMultiChartHover(canvas, labels, series, xPositions, options, hitAreas);
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
    const filters = salesMasterFilters();
    const level = filters.category ? "product" : filters.department ? "category" : "department";
    const title = target.closest(".sales-card")?.querySelector("h3");
    if (title) title.textContent = level === "department" ? "Department mix" : level === "category" ? `${filters.department} categories` : `${filters.category} items`;
    let groups = salesAggregateGroup(rows.filter((row) => !String(row.code || "").startsWith("__ROLLUP__") || level !== "product"), level)
      .filter((row) => row.name && row.name !== "Unassigned")
      .sort((a, b) => level === "department"
        ? departmentRank(a.name) - departmentRank(b.name) || a.name.localeCompare(b.name)
        : b.sales - a.sales);
    if (level === "product") groups = groups.slice(0, 60);
    if (!groups.length) {
      target.innerHTML = `<div class="sales-empty">No ${level === "product" ? "items" : level === "category" ? "categories" : "departments"} were found for this period.</div>`;
      return;
    }
    const total = groups.reduce((sum, row) => sum + Number(row.sales || 0), 0);
    const max = Math.max(...groups.map((row) => Number(row.sales || 0)), 1);
    let running = 0;
    const donutSegments = groups.slice(0, 12).map((row, index) => {
      const share = total ? (Number(row.sales || 0) / total) * 100 : 0;
      const offset = -running;
      running += share;
      return `<circle class="sales-donut-segment" cx="21" cy="21" r="15.9155" pathLength="100" fill="none" stroke="${SALES_TREND_COLORS[index % SALES_TREND_COLORS.length]}" stroke-dasharray="${share.toFixed(3)} ${(100 - share).toFixed(3)}" stroke-dashoffset="${offset.toFixed(3)}" data-open-department="${escapeHtml(row.name)}" data-donut-name="${escapeHtml(row.name)}" data-donut-sales="${Number(row.sales || 0)}" data-donut-units="${Number(row.units || 0)}" data-donut-share="${share.toFixed(1)}"><title>${escapeHtml(row.name)} · ${currency.format(row.sales)} · ${number.format(row.units)} units</title></circle>`;
    }).join("");
    const actionAttribute = level === "department" ? "data-open-department" : level === "category" ? "data-open-category" : "data-open-item";
    const list = groups.map((row, index) => {
      const share = total ? (Number(row.sales || 0) / total) * 100 : 0;
      return `<button type="button" class="sales-mix-row" ${actionAttribute}="${escapeHtml(row.name)}">
        <i style="--mix-color:${SALES_TREND_COLORS[index % SALES_TREND_COLORS.length]};--mix-width:${Math.max(2, (Number(row.sales || 0) / max) * 100).toFixed(1)}%"></i>
        <span><b>${escapeHtml(row.name)}</b><small>${number.format(row.units)} units · ${share.toFixed(1)}%</small></span>
        <strong>${currency.format(row.sales)}</strong><em>›</em>
      </button>`;
    }).join("");
    target.innerHTML = `<div class="sales-mix-browser ${level === "department" ? "has-donut" : "ranked-only"}">
      ${level === "department" ? `<div class="sales-donut-wrap"><div class="sales-donut"><svg viewBox="0 0 42 42" role="img" aria-label="Department sales mix">${donutSegments}</svg><span><b>${currency.format(total)}</b><small>Total sales</small></span></div><p>Hover for totals · click a slice to open categories.</p></div>` : ""}
      <div class="sales-mix-list">${list}</div>
    </div>`;
    target.querySelectorAll(".sales-donut-segment").forEach((segment) => {
      segment.addEventListener("pointerenter", (event) => {
        let tooltip = document.querySelector("#salesChartTooltip");
        if (!tooltip) { tooltip = document.createElement("div"); tooltip.id = "salesChartTooltip"; tooltip.className = "sales-chart-tooltip"; document.body.append(tooltip); }
        tooltip.innerHTML = `<strong>${escapeHtml(segment.dataset.donutName)}</strong><br><b>${currency.format(Number(segment.dataset.donutSales || 0))}</b> · ${number.format(Number(segment.dataset.donutUnits || 0))} units<br><em>${segment.dataset.donutShare}% of department sales · click to open</em>`;
        tooltip.hidden = false;
        tooltip.style.left = `${Math.min(event.clientX + 14, window.innerWidth - 300)}px`;
        tooltip.style.top = `${Math.max(8, event.clientY - 72)}px`;
      });
      segment.addEventListener("pointermove", (event) => {
        const tooltip = document.querySelector("#salesChartTooltip");
        if (!tooltip) return;
        tooltip.style.left = `${Math.min(event.clientX + 14, window.innerWidth - 300)}px`;
        tooltip.style.top = `${Math.max(8, event.clientY - 72)}px`;
      });
      segment.addEventListener("pointerleave", () => { const tooltip = document.querySelector("#salesChartTooltip"); if (tooltip && !tooltip.classList.contains("locked")) tooltip.hidden = true; });
    });
  }

  async function renderSalesOverview() {
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    if (!latest) return;
    updateSalesMasterOptions();
    const trendDate = document.querySelector("#salesTrendDate");
    if (trendDate && !trendDate.value) setSalesPeriod(salesUi.selectedDate || latest, salesUi.periodUnit);
    const focus = trendDate?.value || salesUi.selectedDate || latest;
    const focusLabel = document.querySelector("#salesOverviewFocus");
    if (focusLabel) focusLabel.textContent = `Every chart, KPI and item below follows the master filters.`;
    await renderSalesTrend();
  }

  function renderSalesCalendarGrid(rows, priorRows = []) {
    const target = document.querySelector("#salesCalendar");
    const title = document.querySelector("#salesCalendarTitle");
    if (!target || !salesUi.calendarMonth) return;
    const bounds = salesMonthBounds(salesUi.calendarMonth);
    const date = new Date(Date.UTC(bounds.year, bounds.month - 1, 1, 12));
    if (title) title.textContent = date.toLocaleDateString("en-US", { month: "long", year: "numeric", timeZone: "UTC" });
    const datePicker = document.querySelector("#salesCalendarDatePicker");
    if (datePicker) datePicker.value = salesUi.selectedDate?.slice(0, 7) === salesUi.calendarMonth ? salesUi.selectedDate : bounds.start;
    const daily = new Map(salesAggregateDaily(rows).map((row) => [row.date, row]));
    const priorDaily = new Map(salesAggregateDaily(priorRows).map((row) => [row.date, row]));
    const firstDay = date.getUTCDay();
    const daysInMonth = Number(bounds.end.slice(8));
    const cells = [];
    ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].forEach((day) => cells.push(`<div class="sales-calendar__weekday">${day}</div>`));
    for (let index = 0; index < firstDay; index += 1) cells.push(`<div class="sales-calendar__day muted"></div>`);
    for (let day = 1; day <= daysInMonth; day += 1) {
      const iso = `${salesUi.calendarMonth}-${String(day).padStart(2, "0")}`;
      const row = daily.get(iso);
      const priorIso = salesShiftYear(iso, -1);
      const priorRow = priorDaily.get(priorIso);
      const delta = row && priorRow ? salesDelta(row.sales, priorRow.sales) : null;
      const classes = ["sales-calendar__day", row || priorRow ? "has-sales" : "", iso === salesUi.selectedDate ? "selected" : ""].filter(Boolean).join(" ");
      cells.push(`<button type="button" class="${classes}" data-sales-date="${iso}"><span class="sales-calendar__date">${day}</span><div class="sales-calendar__compare"><span class="prior"><small>${Number(bounds.year) - 1}</small><strong>${priorRow ? currency.format(priorRow.sales) : "—"}</strong></span><span class="current"><small>${bounds.year}</small><strong>${row ? currency.format(row.sales) : "—"}</strong></span></div>${delta == null ? "" : `<em class="${delta < 0 ? "down" : "up"}">${salesDeltaText(delta)}</em>`}</button>`);
    }
    target.innerHTML = cells.join("");
  }

  function renderSalesCalendarMix(rows) {
    const target = document.querySelector("#salesCalendarMix");
    const title = document.querySelector("#salesCalendarDayTitle");
    if (title) title.textContent = salesUi.selectedDate ? salesDisplayDate(salesUi.selectedDate) : "Choose a calendar day";
    if (!target) return;
    const scoped = rows.filter((row) => !salesUi.calendarDepartment || row.department === salesUi.calendarDepartment);
    const dimension = salesUi.calendarDepartment ? "category" : "department";
    const groups = salesAggregateGroup(scoped, dimension).sort((a, b) => b.sales - a.sales);
    const allTotals = salesTotals(rows);
    target.innerHTML = `<button type="button" class="${!salesUi.calendarDepartment && !salesUi.calendarCategory ? "active" : ""}" data-calendar-all><b>All ${salesUi.calendarDepartment ? "categories" : "departments"}</b><span>${currency.format(allTotals.sales)} · ${number.format(allTotals.units)} units</span></button>${groups.map((group) => `<button type="button" class="${(dimension === "department" ? salesUi.calendarDepartment : salesUi.calendarCategory) === group.name ? "active" : ""}" data-calendar-${dimension}="${escapeHtml(group.name)}"><b>${escapeHtml(group.name)}</b><span>${currency.format(group.sales)} · ${number.format(group.units)} units</span></button>`).join("")}`;
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
    const latestRollupDate = rows.reduce((max, row) => row.date > max ? row.date : max, "");
    if (salesUi.selectedDate && salesUi.selectedDate.slice(0, 7) !== salesUi.calendarMonth) salesUi.selectedDate = "";
    renderSalesCalendarGrid(rows, priorRollups);
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

    const yearStart = `${cutoff.slice(0, 4)}-01-01`;
    const priorYearStart = salesShiftYear(yearStart, -1);
    const [ytdRollups, priorYtdRollups, adjustments] = await Promise.all([
      salesFetchRollups("DAY", yearStart, cutoff),
      salesFetchRollups("DAY", priorYearStart, priorCutoff),
      salesFetchAdjustments(bounds.start, cutoff),
    ]);
    renderSalesPeriodKpis(rows.filter((row) => row.date <= cutoff), priorRollups.filter((row) => row.date <= priorCutoff), ytdRollups, priorYtdRollups, { start: bounds.start, end: cutoff }, { start: priorBounds.start, end: priorCutoff }, "month", adjustments);
  }

  async function renderSalesDayPage() {
    const selected = salesUi.selectedDate;
    if (!selected) { setSalesMode("calendar"); return; }
    const priorSelected = salesShiftYear(selected, -1);
    const yearStart = `${selected.slice(0, 4)}-01-01`;
    const priorYearStart = salesShiftYear(yearStart, -1);
    const title = document.querySelector("#salesCalendarDayTitle");
    if (title) title.textContent = salesDate(selected).toLocaleDateString("en-US", { weekday: "long", month: "long", day: "numeric", year: "numeric", timeZone: "UTC" });
    await ensureRemoteSalesRange(selected, selected, { silent: true });
    const dayRows = salesRangeRows(selected, selected);
    renderSalesCalendarMix(dayRows);
    const filters = {
      department: salesUi.calendarDepartment,
      category: salesUi.calendarCategory,
      vendor: salesUi.vendors.size === 1 ? [...salesUi.vendors][0] : "",
      vendors: [...salesUi.vendors],
      item: "",
    };
    const isFiltered = Boolean(filters.department || filters.category || filters.vendors.length);
    const [dayRollup, priorDayRollup, ytdRollups, priorYtdRollups, adjustments] = await Promise.all([
      salesFetchRollups("DAY", selected, selected),
      salesFetchRollups("DAY", priorSelected, priorSelected),
      salesFetchRollups("DAY", yearStart, selected),
      salesFetchRollups("DAY", priorYearStart, priorSelected),
      salesFetchAdjustments(selected, selected),
    ]);
    let currentRows = dayRollup;
    let priorRows = priorDayRollup;
    let currentYtdRows = ytdRollups;
    let priorYtdRows = priorYtdRollups;
    let scopedAdjustments = adjustments;
    if (isFiltered) {
      [currentRows, priorRows, currentYtdRows, priorYtdRows] = await Promise.all([
        salesFetchFilteredRows(selected, selected, filters),
        salesFetchFilteredRows(priorSelected, priorSelected, filters),
        salesFetchFilteredRows(yearStart, selected, filters),
        salesFetchFilteredRows(priorYearStart, priorSelected, filters),
      ]);
      const allStoreSales = salesTotals(dayRollup).sales;
      const ratio = allStoreSales ? salesTotals(currentRows).sales / allStoreSales : 0;
      scopedAdjustments = { ...adjustments, gross: adjustments.gross * ratio, discount: adjustments.discount * ratio, returns: adjustments.returns * ratio, allocated: true };
    }
    const current = salesTotals(currentRows);
    const prior = salesTotals(priorRows);
    const delta = salesDelta(current.sales, prior.sales);
    const comparison = document.querySelector("#salesDayComparison");
    if (comparison) comparison.innerHTML = `<span>${salesDisplayDate(priorSelected)}: ${currency.format(prior.sales)}</span><b>${salesDisplayDate(selected)}: ${currency.format(current.sales)}</b>${delta == null ? "" : `<em class="${delta < 0 ? "down" : "up"}">${salesDeltaText(delta)} year over year</em>`}`;
    renderSalesPeriodKpis(currentRows, priorRows, currentYtdRows, priorYtdRows, { start: selected, end: selected }, { start: priorSelected, end: priorSelected }, "day", scopedAdjustments);
    await renderSalesItems({ start: selected, end: selected }, null, dayRows);
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

  function salesTrendDates(monthCount, anchor, latest) {
    const endDate = salesDate(anchor || latest);
    const months = Math.max(1, Number(monthCount || 6));
    const startDate = new Date(Date.UTC(endDate.getUTCFullYear(), endDate.getUTCMonth() - months + 1, 1, 12));
    let end = salesIso(endDate);
    if (latest && end > latest) end = latest;
    return { start: salesIso(startDate), end };
  }

  let salesMasterCatalogStamp = "";
  let salesMasterCatalogRows = [];

  function salesMasterCatalog() {
    const inventoryRows = state.inventoryRows?.length ? state.inventoryRows : [...(state.latestInventory?.values?.() || [])];
    const stamp = `${state._dataCacheStamp || 0}|${inventoryRows.length}`;
    if (stamp === salesMasterCatalogStamp) return salesMasterCatalogRows;
    salesMasterCatalogStamp = stamp;
    const seen = new Set();
    salesMasterCatalogRows = [];
    [...inventoryRows, ...state.rawSales].forEach((source) => {
      const row = salesHierarchyRow({
        code: cleanCell(source.code || source.Code),
        product: cleanCell(source.product || source.name || source.item_name || source.PRODUCT),
        department: cleanCell(source.department || source.Department || source.DEPARTMENT),
        category: cleanCell(source.category || source.Category || source.CATEGORY),
        vendor: cleanCell(source.vendor || source.vendor_name || source.VENDOR),
      });
      const key = codeKey(row.code) || row.product.toLowerCase();
      if (!key || seen.has(key)) return;
      seen.add(key);
      salesMasterCatalogRows.push({ ...row, vendor: row.vendor || "Unassigned" });
    });
    return salesMasterCatalogRows;
  }

  function salesMasterFilters() {
    return {
      department: cleanCell(salesUi.drill.department),
      category: cleanCell(salesUi.drill.category),
      vendor: salesUi.vendors.size === 1 ? [...salesUi.vendors][0] : "",
      vendors: [...salesUi.vendors],
      item: cleanCell(salesUi.drill.item),
    };
  }

  function salesSetSelectOptions(select, values, label, preferredOrder = null) {
    if (!select) return "";
    const current = select.value;
    let options = [...new Set(values.map((value) => cleanCell(value) || "Unassigned"))];
    options.sort((a, b) => preferredOrder
      ? (preferredOrder.indexOf(a.toUpperCase()) === -1 ? 999 : preferredOrder.indexOf(a.toUpperCase())) - (preferredOrder.indexOf(b.toUpperCase()) === -1 ? 999 : preferredOrder.indexOf(b.toUpperCase())) || a.localeCompare(b)
      : a.localeCompare(b));
    select.innerHTML = `<option value="">${label}</option>${options.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}`;
    select.value = options.includes(current) ? current : "";
    return select.value;
  }

  function updateSalesMasterOptions() {
    const catalog = salesMasterCatalog();
    const vendors = [...new Set(catalog.map((row) => cleanCell(row.vendor) || "Unassigned"))].sort(compareDisplayValue);
    const needle = cleanCell(document.querySelector("#salesVendorSearch")?.value).toLowerCase();
    const options = document.querySelector("#salesVendorOptions");
    if (options) options.innerHTML = vendors.filter((vendor) => !needle || vendor.toLowerCase().includes(needle)).map((vendor) => `<label><input type="checkbox" value="${escapeHtml(vendor)}" ${salesUi.vendors.has(vendor) ? "checked" : ""}/> <span>${escapeHtml(vendor)}</span></label>`).join("");
    const count = document.querySelector("#salesVendorCount");
    if (count) count.textContent = salesUi.vendors.size ? `${salesUi.vendors.size} selected` : "All";
    const itemRows = catalog.filter((row) => (!salesUi.drill.department || row.department === salesUi.drill.department) && (!salesUi.drill.category || row.category === salesUi.drill.category) && (!salesUi.vendors.size || salesUi.vendors.has(row.vendor)));
    const list = document.querySelector("#salesSoldItems");
    if (list) list.innerHTML = itemRows.slice(0, 7000).map((row) => `<option value="${escapeHtml(row.product)}">${escapeHtml(row.code)}</option>`).join("");
    renderSalesDrillPath();
  }

  function renderSalesDrillPath() {
    const target = document.querySelector("#salesDrillPath");
    if (!target) return;
    const filters = salesMasterFilters();
    const pieces = [`<button type="button" data-drill-level="all">All departments</button>`];
    if (filters.department) pieces.push(`<span>›</span><button type="button" data-drill-level="department">${escapeHtml(filters.department)}</button>`);
    if (filters.category) pieces.push(`<span>›</span><button type="button" data-drill-level="category">${escapeHtml(filters.category)}</button>`);
    if (filters.item) pieces.push(`<span>›</span><button type="button" data-drill-level="item">${escapeHtml(filters.item)}</button>`);
    if (filters.vendors.length) pieces.push(`<em>${escapeHtml(filters.vendors.join(", "))}</em>`);
    target.innerHTML = pieces.join("");
  }

  function salesFilterTitle(filters) {
    return filters.item || filters.category || filters.department || filters.vendor || "All-store performance";
  }

  function salesRowsMatchFilters(row, filters) {
    const itemText = `${cleanCell(row.product)} ${cleanCell(row.code)}`.toLowerCase();
    return (!filters.department || cleanCell(row.department).toLowerCase() === filters.department.toLowerCase())
      && (!filters.category || cleanCell(row.category).toLowerCase() === filters.category.toLowerCase())
      && (!(filters.vendors || []).length || filters.vendors.some((vendor) => cleanCell(row.vendor).toLowerCase() === vendor.toLowerCase()))
      && (!filters.item || itemText.includes(filters.item.toLowerCase()));
  }

  async function salesFetchFilteredRows(start, end, filters) {
    if (salesUi.vendorCompare && !filters.department && !filters.category && !filters.item) {
      const rows = await salesFetchRollups("VENDOR__", start, end);
      return rows.filter((row) => salesRowsMatchFilters(row, filters));
    }
    if (!filters.department && !filters.category && !(filters.vendors || []).length && !filters.item) return salesFetchRollups("CATEGORY__", start, end);
    if (!(filters.vendors || []).length && !filters.item) {
      const rows = await salesFetchRollups("CATEGORY__", start, end);
      return rows.filter((row) => salesRowsMatchFilters(row, filters));
    }
    const catalog = salesMasterCatalog();
    const itemNeedle = filters.item.toLowerCase();
    const exactItem = filters.item ? catalog.find((row) => row.code.toLowerCase() === itemNeedle || row.product.toLowerCase() === itemNeedle) : null;
    const query = {
      select: "sales_date,code,product,vendor,category,department,qty,sales,cost_sold,profit",
      and: `(sales_date.gte.${start},sales_date.lte.${end})`,
      order: "sales_date.asc",
    };
    if (exactItem?.code) query.code = `eq.${exactItem.code}`;
    else {
      query.code = "not.like.__ROLLUP__*";
      if (filters.item) query.product = `ilike.*${filters.item.replace(/[*,()]/g, "")}*`;
      if (filters.department) query.department = `eq.${filters.department}`;
      if (filters.category) query.category = `eq.${filters.category}`;
      if ((filters.vendors || []).length === 1) query.vendor = `eq.${filters.vendors[0]}`;
    }
    const fetchQuery = (requestQuery) => directSelectRows("daily_sales", requestQuery).catch(() => []);
    const rawRows = (filters.vendors || []).length > 1 && !exactItem?.code
      ? (await Promise.all(filters.vendors.map((vendor) => fetchQuery({ ...query, vendor: `eq.${vendor}` })))).flat()
      : await fetchQuery(query);
    const rows = rawRows.map(hydrateSalesFromSupabase);
    return rows.map(salesHierarchyRow).filter((row) => salesRowsMatchFilters(row, filters));
  }

  async function salesFetchDetailRows(start, end, filters) {
    if (!filters.item && !filters.category && !(filters.vendors || []).length) return [];
    const query = {
      select: "*",
      and: `(sales_date.gte.${start},sales_date.lte.${end})`,
      code: "not.like.__ROLLUP__*",
      order: "sales_date.asc",
    };
    if (filters.department) query.department = `eq.${filters.department}`;
    if (filters.category) query.category = `eq.${filters.category}`;
    if ((filters.vendors || []).length === 1) query.vendor = `eq.${filters.vendors[0]}`;
    const fetchQuery = (requestQuery) => directSelectRows("daily_sales", requestQuery).catch(() => []);
    const rawRows = (filters.vendors || []).length > 1
      ? (await Promise.all(filters.vendors.map((vendor) => fetchQuery({ ...query, vendor: `eq.${vendor}` })))).flat()
      : await fetchQuery(query);
    return rawRows.map(hydrateSalesFromSupabase).map(salesHierarchyRow).filter((row) => salesRowsMatchFilters(row, filters));
  }

  async function salesFetchPeriodItemRows(start, end, filters) {
    const key = `${start}|${end}|${filters.department}|${filters.category}|${(filters.vendors || []).join(",")}`;
    const cached = salesUi.itemDetails.get(key);
    if (cached && Date.now() - cached.at < 300000) return cached.promise;
    const promise = (async () => {
      const query = {
        select: "sales_date,code,product,vendor,category,department,qty,sales,cost_sold,profit",
        and: `(sales_date.gte.${start},sales_date.lte.${end})`,
        code: "not.like.__ROLLUP__*",
        order: "sales_date.asc",
      };
      if (filters.department) query.department = `eq.${filters.department}`;
      if (filters.category) query.category = `eq.${filters.category}`;
      if ((filters.vendors || []).length === 1) query.vendor = `eq.${filters.vendors[0]}`;
      const fetchQuery = (requestQuery) => directSelectRows("daily_sales", requestQuery).catch(() => []);
      const rawRows = (filters.vendors || []).length > 1
        ? (await Promise.all(filters.vendors.map((vendor) => fetchQuery({ ...query, vendor: `eq.${vendor}` })))).flat()
        : await fetchQuery(query);
      return rawRows.map(hydrateSalesFromSupabase).map(salesHierarchyRow).filter((row) => salesRowsMatchFilters(row, filters));
    })().catch((error) => {
      salesUi.itemDetails.delete(key);
      console.warn("Period item details failed", error);
      return [];
    });
    salesUi.itemDetails.set(key, { at: Date.now(), promise });
    return promise;
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

  function salesChartDimension(filters) {
    if (salesUi.vendorCompare) return "vendor";
    if (filters.item || filters.category) return "";
    if (filters.department) return "category";
    return "department";
  }

  function salesMetricValue(row, metric) {
    if (metric === "margin") return Number(row.sales || 0) ? (Number(row.profit || 0) / Number(row.sales || 0)) * 100 : 0;
    return Number(row[metric] || 0);
  }

  function salesDimensionSeries(currentRows, priorRows, metric, dimension) {
    if (dimension === "date") {
      const current = salesBucketSeries(currentRows, metric, false).map((row) => ({ ...row, label: row.key.slice(5) }));
      const prior = salesBucketSeries(priorRows, metric, false).map((row) => ({ ...row, label: row.key.slice(5) }));
      return { current, prior };
    }
    const aggregate = (rows) => salesAggregateGroup(rows, dimension).map((row) => ({
      name: row.name,
      label: row.name.length > 18 ? `${row.name.slice(0, 17)}…` : row.name,
      tooltipLabel: row.name,
      value: salesMetricValue(row, metric),
      units: row.units,
      sales: row.sales,
      profit: row.profit,
    }));
    const currentGroups = aggregate(currentRows);
    const priorGroups = aggregate(priorRows);
    const names = [...new Set([...currentGroups, ...priorGroups].map((row) => row.name))];
    const currentMap = new Map(currentGroups.map((row) => [row.name, row]));
    const priorMap = new Map(priorGroups.map((row) => [row.name, row]));
    names.sort((a, b) => {
      if (dimension === "department") {
        const ar = GKPOS_DEPARTMENT_ORDER.indexOf(a.toUpperCase());
        const br = GKPOS_DEPARTMENT_ORDER.indexOf(b.toUpperCase());
        return (ar < 0 ? 999 : ar) - (br < 0 ? 999 : br) || a.localeCompare(b);
      }
      return Number(currentMap.get(b)?.value || priorMap.get(b)?.value || 0) - Number(currentMap.get(a)?.value || priorMap.get(a)?.value || 0);
    });
    const limited = dimension === "product" ? names.slice(0, 35) : names;
    const blank = (name) => ({ name, label: name.length > 18 ? `${name.slice(0, 17)}…` : name, tooltipLabel: name, value: 0, units: 0 });
    return {
      current: limited.map((name) => currentMap.get(name) || blank(name)),
      prior: limited.map((name) => priorMap.get(name) || blank(name)),
    };
  }

  function renderSalesPeriodKpis(currentRows, priorRows, ytdCurrentRows, ytdPriorRows, dates, priorDates, period, adjustments = null) {
    const target = document.querySelector("#salesYoyCards");
    if (!target) return;
    const current = salesTotals(currentRows);
    const prior = salesTotals(priorRows);
    const ytd = salesTotals(ytdCurrentRows);
    const priorYtd = salesTotals(ytdPriorRows);
    const margin = current.sales ? (current.profit / current.sales) * 100 : 0;
    const priorMargin = prior.sales ? (prior.profit / prior.sales) * 100 : 0;
    const delta = salesDelta(current.sales, prior.sales);
    const periodName = { day: "Selected day", week: "Selected week", month: "Selected month", quarter: "Selected quarter", half: "Selected half-year", year: "Selected year", range: "Selected range" }[period] || "Selected period";
    const range = dates.start === dates.end ? salesDisplayDate(dates.start) : `${salesDisplayDate(dates.start)}–${salesDisplayDate(dates.end)}`;
    const priorYear = priorDates.end.slice(0, 4);
    const hasAdjustments = adjustments?.available;
    const currentYear = dates.end.slice(0, 4);
    const row = (label, priorValue, currentValue, formatter = currency) => {
      const change = salesDelta(currentValue, priorValue);
      return `<div class="sales-compare-row"><b>${label}</b><span><small>${priorYear}</small><strong>${formatter.format(priorValue)}</strong></span><span><small>${currentYear}</small><strong>${formatter.format(currentValue)}</strong></span><em class="${change != null && change < 0 ? "down" : "up"}">${salesDeltaText(change)}</em></div>`;
    };
    target.innerHTML = `<section class="sales-summary-panel"><div class="sales-compare-card">
      <header><div><small>${periodName}</small><b>${escapeHtml(range)}</b></div><span>Prior</span><span>Current</span><span>Change</span></header>
      ${row("Net sales", prior.sales, current.sales)}
      ${row("Units", prior.units, current.units, number)}
      ${row("Total cost", prior.cost, current.cost)}
      ${row("Gross profit", prior.profit, current.profit)}
      ${row("Margin", priorMargin, margin, { format: (value) => `${Number(value || 0).toFixed(1)}%` })}
    </div>
    <div class="sales-current-facts">
      <article><small>Gross sales</small><b>${hasAdjustments ? currency.format(adjustments.gross) : "—"}</b></article>
      <article><small>Discounts</small><b>${hasAdjustments ? currency.format(adjustments.discount) : "—"}</b></article>
      <article><small>Returns</small><b>${hasAdjustments ? currency.format(adjustments.returns) : "—"}</b></article>
      <article><small>YTD through ${salesDisplayDate(dates.end)}</small><b>${currency.format(ytd.sales)}</b><span class="${salesDelta(ytd.sales, priorYtd.sales) != null && salesDelta(ytd.sales, priorYtd.sales) < 0 ? "down" : "up"}">${salesDeltaText(salesDelta(ytd.sales, priorYtd.sales))}</span></article>
    </div></section>`;
  }

  async function renderSalesTrend() {
    const request = ++salesUi.trendRequest;
    const filters = salesMasterFilters();
    const metric = document.querySelector("#salesTrendMetric")?.value || "sales";
    const bucket = document.querySelector("#salesTrendBucket")?.value || salesUi.trendBucket || "day";
    salesUi.trendBucket = bucket;
    const compare = !!document.querySelector("#salesTrendCompare")?.checked;
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    const anchor = document.querySelector("#salesTrendDate")?.value || salesUi.selectedDate || latest;
    if (!latest || !anchor) return;
    const requestedDates = selectedSalesPeriod(latest);
    const historyMonths = Number(document.querySelector("#salesItemHistoryRange")?.value || salesUi.itemHistoryMonths || 24);
    salesUi.itemHistoryMonths = historyMonths;
    const dates = filters.item ? salesTrendDates(historyMonths, anchor, latest) : requestedDates;
    const priorDates = { start: salesShiftYear(dates.start, -1), end: salesShiftYear(dates.end, -1) };
    const currentYtdStart = `${dates.end.slice(0, 4)}-01-01`;
    const priorYtdStart = salesShiftYear(currentYtdStart, -1);
    const fetchStart = [dates.start, currentYtdStart, ...(compare ? [priorDates.start, priorYtdStart] : [])].sort()[0];
    const title = document.querySelector("#salesTrendTitle");
    const kicker = document.querySelector("#salesTrendKicker");
    const dimension = salesChartDimension(filters);
    if (kicker) kicker.textContent = filters.item ? `${historyMonths}-month item history` : `${bucket === "day" ? "Daily" : bucket === "week" ? "Weekly" : "Monthly"} view · ${salesDisplayDate(dates.start)}–${salesDisplayDate(dates.end)}`;
    if (title) title.textContent = filters.item ? `${filters.item} sales history` : filters.category ? `${filters.category} sales trend` : filters.department ? `${filters.department} sales trend` : salesUi.vendorCompare ? "Vendor trends" : "Department totals";
    document.body.classList.toggle("sales-item-history-active", !!filters.item);
    const summary = document.querySelector("#salesTrendSummary");
    if (summary) summary.innerHTML = `<span>Loading ${escapeHtml(salesDisplayDate(dates.start))} through ${escapeHtml(salesDisplayDate(dates.end))}…</span>`;

    const rows = await salesFetchFilteredRows(fetchStart, dates.end, filters);
    if (request !== salesUi.trendRequest) return;
    const currentRows = rows.filter((row) => row.date >= dates.start && row.date <= dates.end);
    const priorRows = compare ? rows.filter((row) => row.date >= priorDates.start && row.date <= priorDates.end) : [];
    const ytdCurrentRows = rows.filter((row) => row.date >= currentYtdStart && row.date <= dates.end);
    const ytdPriorRows = compare ? rows.filter((row) => row.date >= priorYtdStart && row.date <= priorDates.end) : [];
    const currentYear = dates.end.slice(0, 4);
    const priorYear = String(Number(currentYear) - 1);
    const fallbackName = filters.category || filters.department || salesFilterTitle(filters);
    const snapshotDepartments = !filters.department && !filters.category && !filters.item && !filters.vendors.length && !salesUi.vendorCompare;
    const trend = snapshotDepartments
      ? salesDepartmentComparisonSeries(currentRows, priorRows, metric, compare)
      : salesTrendSeries(currentRows, priorRows, metric, bucket, dimension, fallbackName, dates, compare);
    const currentSeriesCount = trend.series.filter((item) => !item.prior).length;
    const adaptiveChartType = salesUi.chartType === "line" ? "line" : (snapshotDepartments || currentSeriesCount > 1 ? "bar" : "combo");
    const legend = document.querySelector("#salesTrendLegend");
    if (legend) legend.innerHTML = filters.item ? "" : snapshotDepartments
      ? `<span><i style="background:#16835b"></i>${currentYear} department totals</span>${compare ? `<span class="sales-prior-key">Light overlay = ${priorYear}</span>` : ""}`
      : `<span><i style="background:#16835b"></i>${currentYear} ${adaptiveChartType === "combo" ? "bars" : "sales"}</span>${compare ? `<span class="sales-prior-key">${adaptiveChartType === "combo" ? "Red line" : "Light overlay"} = ${priorYear}</span>` : ""}`;
    if (!filters.item) salesDrawMultiSeriesChart(document.querySelector("#salesTrendExplorerChart"), trend.labels, trend.series, { type: adaptiveChartType, height: 350, currency: metric !== "units" && metric !== "margin", percent: metric === "margin", priorLabel: priorYear, currentLabel: currentYear });
    const currentTotals = salesTotals(currentRows);
    const priorTotals = salesTotals(priorRows);
    const currentTotal = metric === "margin" ? (currentTotals.sales ? (currentTotals.profit / currentTotals.sales) * 100 : 0) : salesMetricValue(currentTotals, metric);
    const priorTotal = metric === "margin" ? (priorTotals.sales ? (priorTotals.profit / priorTotals.sales) * 100 : 0) : salesMetricValue(priorTotals, metric);
    const formatter = metric === "units" ? number : metric === "margin" ? { format: (value) => `${Number(value || 0).toFixed(1)}%` } : currency;
    if (summary) summary.innerHTML = `<b>${formatter.format(currentTotal)}</b><span>${salesDisplayDate(dates.start)}–${salesDisplayDate(dates.end)}</span>${compare ? `<b class="${currentTotal >= priorTotal ? "up" : "down"}">${salesDeltaText(salesDelta(currentTotal, priorTotal))}</b><span>vs ${priorYear} (${formatter.format(priorTotal)})</span>` : ""}`;
    let adjustments = await salesFetchAdjustments(dates.start, dates.end);
    const scoped = Boolean(filters.department || filters.category || filters.item || filters.vendors.length);
    if (scoped && adjustments.available) {
      const allStoreRows = await salesFetchRollups("DAY", dates.start, dates.end);
      const allStoreSales = salesTotals(allStoreRows).sales;
      const ratio = allStoreSales ? currentTotals.sales / allStoreSales : 0;
      adjustments = { ...adjustments, gross: adjustments.gross * ratio, discount: adjustments.discount * ratio, returns: adjustments.returns * ratio, allocated: true };
    }
    if (request !== salesUi.trendRequest) return;
    renderSalesPeriodKpis(currentRows, priorRows, ytdCurrentRows, ytdPriorRows, dates, priorDates, "range", adjustments);

    const detailRows = filters.item ? currentRows : await salesFetchDetailRows(dates.start, dates.end, filters);
    if (request !== salesUi.trendRequest) return;
    renderSalesDepartmentTree(detailRows.length ? detailRows : currentRows);
    void renderSalesItems(dates, filters, detailRows);
  }

  async function renderSalesItems(selectedDates = null, selectedFilters = null, sourceRows = null) {
    const body = document.querySelector("#salesItemRows");
    if (!body) return;
    const request = ++salesUi.itemRequest;
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || "";
    const anchor = document.querySelector("#salesTrendDate")?.value || latest;
    const calendarMode = salesUi.mode === "day";
    const calendarDay = salesUi.selectedDate || latest;
    const dates = calendarMode ? { start: calendarDay, end: calendarDay } : (selectedDates || selectedSalesPeriod(latest));
    const start = dates.start;
    const end = dates.end;
    const filters = calendarMode ? {
      department: salesUi.calendarDepartment,
      category: salesUi.calendarCategory,
      vendor: salesUi.vendors.size === 1 ? [...salesUi.vendors][0] : "",
      vendors: [...salesUi.vendors],
      item: "",
    } : (selectedFilters || salesMasterFilters());
    const range = document.querySelector("#salesItemsRange");
    if (range) range.textContent = calendarMode ? `${salesDisplayDate(start)} only` : (start && end ? `${salesDisplayDate(start)} through ${salesDisplayDate(end)}` : "Choose a date range.");
    const searchQuery = cleanCell(document.querySelector("#salesItemSearch")?.value).toLowerCase();
    const exactHistory = calendarMode ? "" : cleanCell(filters.item);
    document.querySelector(".sales-items-inline")?.classList.toggle("item-history-view", !!exactHistory);
    const historyRange = document.querySelector(".sales-item-history-range");
    if (historyRange) historyRange.hidden = !exactHistory;
    const query = exactHistory.toLowerCase() || searchQuery;
    const sortSelect = document.querySelector("#salesItemSort");
    const sort = exactHistory ? (salesUi.itemSort === "product" ? "date" : salesUi.itemSort) : (salesUi.itemSort || sortSelect?.value || "sales");
    if (sortSelect && sortSelect.value !== sort && [...sortSelect.options].some((option) => option.value === sort)) sortSelect.value = sort;
    const localCoverageComplete = state._remoteSalesLoadedStart && state._remoteSalesLoadedEnd && start >= state._remoteSalesLoadedStart && end <= state._remoteSalesLoadedEnd;
    let baseSource = Array.isArray(sourceRows) && sourceRows.length
      ? sourceRows.filter((row) => salesRowsMatchFilters(row, filters))
      : (localCoverageComplete ? salesRangeRows(start, end).filter((row) => salesRowsMatchFilters(row, filters)) : []);
    if (!baseSource.length && start && end) {
      const totalTarget = document.querySelector("#salesItemsTotals");
      if (totalTarget) totalTarget.innerHTML = `<span><small>Products</small><b>Loading…</b><em>Building the sold-item list for this period.</em></span>`;
      body.innerHTML = `<tr><td colspan="10" class="sales-empty">Loading products sold from ${salesDisplayDate(start)} through ${salesDisplayDate(end)}…</td></tr>`;
      baseSource = await salesFetchPeriodItemRows(start, end, filters);
      if (request !== salesUi.itemRequest) return;
    }
    const matchedSource = baseSource.filter((row) => !searchQuery || `${cleanCell(row.product)} ${cleanCell(row.code)}`.toLowerCase().includes(searchQuery));
    const map = new Map();
    matchedSource.forEach((row) => {
      const timestamp = cleanCell(row.timestamp || row.saleTime || row.time);
      const key = exactHistory ? `${row.date}|${timestamp || "daily"}` : (codeKey(row.code) || cleanCell(row.product).toLowerCase());
      const current = map.get(key) || { code: row.code, product: row.product || row.code, department: row.department, category: row.category, vendor: row.vendor, timestamp, units: 0, sales: 0, cost: 0, profit: 0 };
      current.date = exactHistory ? row.date : "";
      current.units += Number(row.units || 0); current.sales += Number(row.sales || 0); current.cost += Number(row.cost || 0); current.profit += Number(row.profit || 0);
      map.set(key, current);
    });
    const allRows = [...map.values()].map((row) => ({ ...row, price: row.units ? row.sales / row.units : 0, unitCost: row.units ? row.cost / row.units : 0 }));
    let rows = allRows
      .filter((row) => !query || [row.code, row.product].join(" ").toLowerCase().includes(query));
    const direction = salesUi.itemSortDirection === "asc" ? 1 : -1;
    rows.sort((a, b) => ["date", "product", "department", "category", "vendor"].includes(sort)
      ? direction * cleanCell(a[sort]).localeCompare(cleanCell(b[sort]))
      : direction * (Number(a[sort] || 0) - Number(b[sort] || 0)));
    const totals = salesTotals(rows);
    const finalizedDays = await salesFetchRollups("DAY", start, end);
    const hasItemFilters = Boolean(filters.department || filters.category || (filters.vendors || []).length || query);
    const finalizedSubtotal = !hasItemFilters && finalizedDays.length ? salesTotals(finalizedDays).sales : totals.sales;
    const zeroCostCount = rows.filter((row) => row.sales && !row.cost).length;
    const totalTarget = document.querySelector("#salesItemsTotals");
    if (totalTarget) totalTarget.innerHTML = `<span><small>Products</small><b>${number.format(rows.length)}</b></span><span><small>Units</small><b>${number.format(totals.units)}</b></span><span><small>Item sales</small><b>${currency.format(totals.sales)}</b></span><span class="${zeroCostCount ? "warning" : ""}"><small>Total cost</small><b>${currency.format(totals.cost)}</b>${zeroCostCount ? `<em>${number.format(zeroCostCount)} missing-cost item${zeroCostCount === 1 ? "" : "s"}</em>` : ""}</span><span><small>Final subtotal</small><b>${currency.format(finalizedSubtotal)}</b></span><span><small>Total profit</small><b>${currency.format(totals.profit)}</b></span>`;
    const head = document.querySelector("#salesItemHeadRow");
    if (head) head.innerHTML = exactHistory
      ? `<th><button type="button" data-item-sort="date">Sold date</button></th><th>Time</th><th>Product</th><th>Vendor</th><th><button type="button" data-item-sort="units">Qty</button></th><th><button type="button" data-item-sort="price">Avg net price</button></th><th><button type="button" data-item-sort="sales">Net sales</button></th><th><button type="button" data-item-sort="unitCost">Unit cost</button></th><th><button type="button" data-item-sort="cost">Total cost</button></th><th><button type="button" data-item-sort="profit">Profit</button></th><th>Margin</th>`
      : `<th><button type="button" data-item-sort="product">Product</button></th><th><button type="button" data-item-sort="department">Department</button></th><th><button type="button" data-item-sort="category">Category</button></th><th><button type="button" data-item-sort="vendor">Vendor</button></th><th><button type="button" data-item-sort="units">Units</button></th><th><button type="button" data-item-sort="price">Avg net price</button></th><th><button type="button" data-item-sort="sales">Net sales</button></th><th><button type="button" data-item-sort="unitCost">Unit cost</button></th><th><button type="button" data-item-sort="cost">Total cost</button></th><th><button type="button" data-item-sort="profit">Profit</button></th>`;
    document.querySelectorAll("[data-item-sort]").forEach((button) => {
      const active = button.dataset.itemSort === sort;
      button.classList.toggle("active", active);
      button.dataset.direction = active ? salesUi.itemSortDirection : "";
    });
    body.innerHTML = rows.slice(0, 500).map((row) => exactHistory
      ? `<tr><td><b>${salesDisplayDate(row.date)}</b></td><td>${escapeHtml(salesDisplayTime(row.timestamp))}</td><td><b>${escapeHtml(row.product)}</b><small>${escapeHtml(row.code)}</small></td><td>${escapeHtml(row.vendor || "Unassigned")}</td><td>${number.format(row.units)}</td><td>${currency.format(row.price)}</td><td>${currency.format(row.sales)}</td><td>${currency.format(row.unitCost)}</td><td>${currency.format(row.cost)}</td><td>${currency.format(row.profit)}</td><td>${row.sales ? `${((row.profit / row.sales) * 100).toFixed(1)}%` : "—"}</td></tr>`
      : `<tr data-open-item="${escapeHtml(row.product)}"><td><b>${escapeHtml(row.product)}</b><small>${escapeHtml(row.code)}</small></td><td>${escapeHtml(row.department || "Unassigned")}</td><td>${escapeHtml(row.category || "Unassigned")}</td><td>${escapeHtml(row.vendor || "Unassigned")}</td><td>${number.format(row.units)}</td><td>${currency.format(row.price)}</td><td>${currency.format(row.sales)}</td><td>${currency.format(row.unitCost)}</td><td>${currency.format(row.cost)}</td><td>${currency.format(row.profit)}</td></tr>`).join("") || `<tr><td colspan="10" class="sales-empty">No matching sales items.</td></tr>`;
  }

  function setSalesMode(mode) {
    if (mode === "items") mode = "overview";
    const changed = salesUi.mode !== mode;
    salesUi.mode = mode;
    document.body.classList.toggle("sales-day-open", mode === "day");
    const dayNav = document.querySelector("#salesDayNav");
    if (dayNav) dayNav.hidden = mode !== "day";
    document.querySelectorAll("[data-sales-mode]").forEach((button) => button.classList.toggle("active", button.dataset.salesMode === mode || (mode === "day" && button.dataset.salesMode === "calendar")));
    document.querySelectorAll("[data-sales-mode-view]").forEach((view) => {
      const active = view.dataset.salesModeView === mode;
      view.hidden = !active; view.classList.toggle("active", active);
    });
    syncSalesModeLayout(mode);
    if (changed && mode === "overview") void renderSalesOverview();
    if (changed && mode === "calendar") void renderSalesCalendar();
    if (changed && mode === "day") void renderSalesDayPage();
    if (changed) window.scrollTo({ top: 0, left: 0, behavior: "auto" });
  }

  function renderSalesWorkspace() {
    if (!document.querySelector("#salesWorkspace")) return;
    void renderSalesDataHealth();
    if (salesUi.mode === "overview") void renderSalesOverview();
    else if (salesUi.mode === "calendar") void renderSalesCalendar();
    else if (salesUi.mode === "day") void renderSalesDayPage();
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
    salesUi.calendarDepartment = "";
    salesUi.calendarCategory = "";
    setSalesMode("day");
  });
  document.querySelector("#salesDayBack")?.addEventListener("click", () => setSalesMode("calendar"));
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
    salesUi.calendarDepartment = "";
    salesUi.calendarCategory = "";
    void renderSalesCalendar();
  });
  ["#salesTrendMetric", "#salesTrendBucket", "#salesTrendCompare", "#salesTrendStart", "#salesTrendDate"].forEach((selector) => document.querySelector(selector)?.addEventListener("change", () => { void renderSalesTrend(); }));
  document.querySelector("#salesPeriodUnit")?.addEventListener("change", (event) => {
    const anchor = document.querySelector("#salesTrendDate")?.value || state._remoteSalesAvailableEnd || state.dates.at(-1) || salesIso(new Date());
    setSalesPeriod(anchor, event.target.value);
    void renderSalesTrend();
  });
  document.querySelector("#salesTrendClear")?.addEventListener("click", () => {
    const metric = document.querySelector("#salesTrendMetric"), unit = document.querySelector("#salesPeriodUnit"), bucket = document.querySelector("#salesTrendBucket"), compare = document.querySelector("#salesTrendCompare");
    salesUi.drill = { department: "", category: "", item: "" };
    salesUi.vendors.clear();
    salesUi.vendorCompare = false;
    document.querySelector("#salesVendorCompare")?.classList.remove("active");
    const itemSearch = document.querySelector("#salesItemSearch"); if (itemSearch) itemSearch.value = "";
    if (metric) metric.value = "sales";
    if (unit) unit.value = "month";
    if (bucket) bucket.value = "day";
    setSalesPeriod(state._remoteSalesAvailableEnd || state.dates.at(-1) || salesIso(new Date()), "month");
    if (compare) compare.checked = true;
    salesUi.chartType = "bar";
    document.querySelectorAll("[data-chart-type]").forEach((button) => button.classList.toggle("active", button.dataset.chartType === "bar"));
    updateSalesMasterOptions();
    void renderSalesTrend();
  });
  function shiftSalesAnchor(direction) {
    const input = document.querySelector("#salesTrendStart");
    if (!input?.value) return;
    const date = salesDate(input.value);
    const unit = document.querySelector("#salesPeriodUnit")?.value || salesUi.periodUnit || "month";
    if (unit === "day") date.setUTCDate(date.getUTCDate() + direction);
    else if (unit === "week") date.setUTCDate(date.getUTCDate() + (direction * 7));
    else if (unit === "quarter") date.setUTCMonth(date.getUTCMonth() + (direction * 3));
    else if (unit === "year") date.setUTCFullYear(date.getUTCFullYear() + direction);
    else date.setUTCMonth(date.getUTCMonth() + direction);
    const latest = state._remoteSalesAvailableEnd || state.dates.at(-1) || salesIso(new Date());
    setSalesPeriod(salesIso(date) > latest ? latest : salesIso(date), unit);
    void renderSalesTrend();
  }
  document.querySelector("#salesTrendPrev")?.addEventListener("click", () => shiftSalesAnchor(-1));
  document.querySelector("#salesTrendNext")?.addEventListener("click", () => shiftSalesAnchor(1));
  document.querySelector("#salesVendorSearch")?.addEventListener("input", updateSalesMasterOptions);
  document.querySelector("#salesVendorOptions")?.addEventListener("change", (event) => {
    const checkbox = event.target.closest('input[type="checkbox"]');
    if (!checkbox) return;
    if (checkbox.checked) salesUi.vendors.add(checkbox.value); else salesUi.vendors.delete(checkbox.value);
    updateSalesMasterOptions();
    if (salesUi.mode === "day") void renderSalesDayPage(); else void renderSalesTrend();
  });
  document.querySelector("#salesVendorCompare")?.addEventListener("click", (event) => {
    salesUi.vendorCompare = !salesUi.vendorCompare;
    event.currentTarget.classList.toggle("active", salesUi.vendorCompare);
    event.currentTarget.textContent = salesUi.vendorCompare ? "Comparing vendors" : "Compare vendors";
    setSalesMode("overview");
  });
  document.querySelector("#salesDrillPath")?.addEventListener("click", (event) => {
    const button = event.target.closest("[data-drill-level]");
    if (!button) return;
    const level = button.dataset.drillLevel;
    if (level === "all") salesUi.drill = { department: "", category: "", item: "" };
    if (level === "department") salesUi.drill = { department: salesUi.drill.department, category: "", item: "" };
    if (level === "category") salesUi.drill.item = "";
    const search = document.querySelector("#salesItemSearch"); if (search && !salesUi.drill.item) search.value = "";
    updateSalesMasterOptions(); void renderSalesTrend();
  });
  document.querySelector("#salesCalendarMix")?.addEventListener("click", (event) => {
    const all = event.target.closest("[data-calendar-all]");
    const department = event.target.closest("[data-calendar-department]");
    const category = event.target.closest("[data-calendar-category]");
    if (!all && !department && !category) return;
    if (all) { salesUi.calendarDepartment = ""; salesUi.calendarCategory = ""; }
    if (department) { salesUi.calendarDepartment = department.dataset.calendarDepartment; salesUi.calendarCategory = ""; }
    if (category) salesUi.calendarCategory = category.dataset.calendarCategory;
    void renderSalesDayPage();
  });
  document.querySelector("#salesCalendarClearFilters")?.addEventListener("click", () => {
    salesUi.calendarDepartment = ""; salesUi.calendarCategory = "";
    const search = document.querySelector("#salesItemSearch"); if (search) search.value = "";
    void renderSalesDayPage();
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
  let salesItemSearchTimer = 0;
  document.querySelector("#salesItemSearch")?.addEventListener("input", () => {
    window.clearTimeout(salesItemSearchTimer);
    salesItemSearchTimer = window.setTimeout(() => { void renderSalesItems(); }, 180);
  });
  document.querySelector("#salesItemSearch")?.addEventListener("change", (event) => {
    if (salesUi.mode !== "overview") return;
    const value = cleanCell(event.target.value);
    const match = salesMasterCatalog().find((row) => row.product.toLowerCase() === value.toLowerCase() || row.code.toLowerCase() === value.toLowerCase());
    if (match) {
      salesUi.drill = { department: match.department, category: match.category, item: match.product };
      salesUi.itemSort = "date";
      salesUi.itemSortDirection = "desc";
      updateSalesMasterOptions();
      void renderSalesTrend();
    }
  });
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
    const search = document.querySelector("#salesItemSearch"); if (search) search.value = "";
    if (salesUi.mode === "overview" && salesUi.drill.item) {
      salesUi.drill.item = "";
      updateSalesMasterOptions();
      void renderSalesTrend();
      return;
    }
    salesUi.itemSort = "sales";
    salesUi.itemSortDirection = "desc";
    void renderSalesItems();
  });
  document.querySelector("#salesItemSort")?.addEventListener("change", (event) => {
    salesUi.itemSort = event.target.value;
    salesUi.itemSortDirection = salesUi.itemSort === "product" ? "asc" : "desc";
    renderSalesItems();
  });
  document.querySelector("#salesItemHistoryRange")?.addEventListener("change", (event) => {
    salesUi.itemHistoryMonths = Number(event.target.value || 24);
    void renderSalesTrend();
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
      if (salesUi.mode !== "overview") setSalesMode("overview");
      const categoryName = trend.dataset.openTrend || trend.dataset.openCategory;
      const mapped = GKPOS_CATEGORY_DEPARTMENTS.get(cleanCell(categoryName).toUpperCase());
      salesUi.drill.department = mapped || salesUi.drill.department;
      salesUi.drill.category = categoryName;
      salesUi.drill.item = "";
      updateSalesMasterOptions(); void renderSalesTrend();
    }
    const department = event.target.closest("[data-open-department]");
    if (department) {
      if (salesUi.mode !== "overview") setSalesMode("overview");
      salesUi.drill = { department: department.dataset.openDepartment, category: "", item: "" };
      updateSalesMasterOptions(); void renderSalesTrend();
    }
    const item = event.target.closest("[data-open-item]");
    if (item) {
      if (salesUi.mode !== "overview") setSalesMode("overview");
      const catalogItem = salesMasterCatalog().find((row) => row.product === item.dataset.openItem);
      if (catalogItem) salesUi.drill = { department: catalogItem.department, category: catalogItem.category, item: catalogItem.product };
      else salesUi.drill.item = item.dataset.openItem;
      salesUi.itemSort = "date"; salesUi.itemSortDirection = "desc";
      const input = document.querySelector("#salesItemSearch"); if (input) input.value = item.dataset.openItem;
      updateSalesMasterOptions(); void renderSalesTrend();
    }
  });
  let salesResizeTimer = 0;
  window.addEventListener("resize", () => {
    window.clearTimeout(salesResizeTimer);
    salesResizeTimer = window.setTimeout(() => { if (activeTabName() === "dashboard") renderSalesWorkspace(); }, 180);
  });

  dashboardDataStatus();
  setSalesMode("calendar");
})();

