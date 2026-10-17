(function () {
  const VERSION = "20261017-sales-history-dashboard";
  const DAY_MS = 86400000;

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
    });
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

  dashboardDataStatus();
})();

