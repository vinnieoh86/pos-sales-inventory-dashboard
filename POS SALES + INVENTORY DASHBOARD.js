const state = {
  rawSales: [],
  inventories: new Map(),
  latestInventory: new Map(),
  excelItems: new Map(),
  excelByPlu: new Map(),
  excelByItemNumber: new Map(),
  reorderOverrides: JSON.parse(localStorage.getItem("posDashboardReorderOverrides:v1") || "{}"),
  visibleColumns: JSON.parse(localStorage.getItem("posDashboardVisibleColumns:v3") || "null"),
  columnOrder: JSON.parse(localStorage.getItem("posDashboardColumnOrder:v3") || "null"),
  arrangeColumns: false,
  detailOrder: JSON.parse(localStorage.getItem("posDashboardDetailOrder:v1") || "null"),
  detailFilters: JSON.parse(localStorage.getItem("posDashboardDetailFilters:v1") || "null"),
  parentRules: JSON.parse(localStorage.getItem("posDashboardParentRules:v1") || "[]"),
  attributeRules: JSON.parse(localStorage.getItem("posDashboardAttributeRules:v1") || "[]"),
  inventorySort: { key: "product", dir: "asc" },
  dates: [],
  filteredSkus: [],
  inventoryRows: [],
  parentRows: [],
  activePresetDays: 90,
  // Performance cache â€” invalidated on data load, valid across filter/search changes
  _skuCache: null,      // full unfiltered SKU map keyed by codeKey
  _skuCacheStamp: 0,    // increments when raw data changes
  _dataCacheStamp: 0,   // stamp written on load; cache valid when stamps match
  _salesIndex: null,
  _salesIndexStamp: 0,
  _salesWindowsCache: new Map(),
  _dailyTotals: new Map(),
  _inventoryRowIndex: new Map(),
  _filteredSkuIndex: new Map(),
  _loadedFileSignatures: new Set(),
  _activeDetailCode: "",
  _activeTabRenderToken: 0,
  _activeTabRenderHandle: 0,
  _parentPartsCache: new Map(),
  _datePresetsReady: false,
  tabSearches: JSON.parse(localStorage.getItem("posDashboardTabSearches:v1") || "{\"dashboard\":\"\",\"inventory\":\"\",\"ordering\":\"\"}"),
  countSessions: JSON.parse(localStorage.getItem("posDashboardCountSessions:v1") || "[]"),
  adjustmentLog: JSON.parse(localStorage.getItem("posDashboardAdjustLog:v1") || "[]"),
  selectedInventoryCodes: new Set(),
  selectedSkuCodes: new Set(),
  _pinnedAdjustCode: null,
  vendorRules: JSON.parse(localStorage.getItem("posDashboardVendorRules:v1") || "[]"),
  vendorRuleEditId: null,
  vendorRuleSelectedDays: new Set(),
  orderSort: { key: "recommendedOrder", dir: "desc" },
  orderVendorFilter: "Active",
  orderVisibleColumns: JSON.parse(localStorage.getItem("posOrderColumns:v1") || "null"),
  activeCountSession: null,
  countQtyBuffer: "0",
  selectedCountItemCode: "",
  countStage: "search",
  pendingDuplicateCount: null,
  countReportMode: "input",
  // Stock adjust modal state
  stockAdjustItem: null,
  stockAdjustAction: null,   // "add" | "remove" | "set"
  stockAdjustQtyBuffer: "0",
  pendingDeleteSessionId: null,
  pendingSubmitSessionId: null,
};

const ENABLE_CUSTOM_PARENT_RULES = true;
const ENABLE_CUSTOM_ATTRIBUTE_RULES = false;

localStorage.removeItem("posDashboardActiveCountSession:v1");

// Increment this whenever raw data changes to bust the SKU cache
function bumpDataStamp() {
  // After data loads, apply the 90D default if it hasn't been overridden
  clearTimeout(state._applyPresetTimer);
  state._applyPresetTimer = setTimeout(() => {
    // Only auto-apply 90D if it hasn't been manually changed by the user
    if (state.activePresetDays === 90 && state.dates.length) {
      applyDatePreset(90);
    }
  }, 200);
  state._dataCacheStamp += 1;
  state._skuCache = null;
  state._inventoryCache = null;
  state._salesIndex = null;
  state._salesIndexStamp = 0;
  state._salesWindowsCache = new Map();
  state._dailyTotals = new Map();
  state._inventoryRowIndex = new Map();
  state._countSearchIndex = null; // invalidate count search index too
  // Pre-warm inventory cache in background after a short delay so the
  // UI stays responsive but the next render is instant
  clearTimeout(state._prewarmTimer);
  state._prewarmTimer = setTimeout(() => {
    if (activeTabName() !== "inventory") {
      // Build but don't display — just warms the cache
      buildInventoryRows({ ignoreQuery: true, ignoreFilters: true, ignoreStateFilter: true });
    }
  }, 400);
}

// Debounce helper â€” delays fast-typing renders
function debounce(fn, ms) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

const DB_NAME = "posDashboardHistory_launch421";
const DB_VERSION = 1;
const DB_STORE = "app";
const DB_KEY = "state_v2";

const els = {
  fileInput: document.querySelector("#fileInput"),
  folderInput: document.querySelector("#folderInput"),
  excelInput: document.querySelector("#excelInput"),
  dropZone: document.querySelector("#dropZone"),
  fileCount: document.querySelector("#fileCount"),
  dateCoverage: document.querySelector("#dateCoverage"),
  excelStatus: document.querySelector("#excelStatus"),
  searchInput: document.querySelector("#searchInput"),
  startDate: document.querySelector("#startDate"),
  endDate: document.querySelector("#endDate"),
  departmentFilter: document.querySelector("#departmentFilter"),
  categoryFilter: document.querySelector("#categoryFilter"),
  vendorFilter: document.querySelector("#vendorFilter"),
  colorFilter: document.querySelector("#colorFilter"),
  leadDays: document.querySelector("#leadDays"),
  safetyDays: document.querySelector("#safetyDays"),
  daysOfInventory: document.querySelector("#daysOfInventory"),
  clearFiltersButton: document.querySelector("#clearFiltersButton"),
  clearFilterButtons: document.querySelectorAll("[data-clear-filter]"),
  inventoryQuickTools: document.querySelector("#inventoryQuickTools"),
  createPoShortcut: document.querySelector("#createPoShortcut"),
  chooseSalesButton: document.querySelector("#chooseSalesButton"),
  chooseFolderButton: document.querySelector("#chooseFolderButton"),
  chooseExcelButton: document.querySelector("#chooseExcelButton"),
  totalSales: document.querySelector("#totalSales"),
  salesDelta: document.querySelector("#salesDelta"),
  unitsSold: document.querySelector("#unitsSold"),
  avgDailyUnits: document.querySelector("#avgDailyUnits"),
  grossProfit: document.querySelector("#grossProfit"),
  costSold: document.querySelector("#costSold"),
  costTotal: document.querySelector("#costTotal"),
  marginRate: document.querySelector("#marginRate"),
  riskCount: document.querySelector("#riskCount"),
  trendChart: document.querySelector("#trendChart"),
  segmentMetric: document.querySelector("#segmentMetric"),
  segmentGroup: document.querySelector("#segmentGroup"),
  segmentTitle: document.querySelector("#segmentTitle"),
  categoryPanelTitle: document.querySelector("#categoryPanelTitle"),
  categoryPanelHint: document.querySelector("#categoryPanelHint"),
  compareToggle: document.querySelector("#compareToggle"),
  compareLegend: document.querySelector("#compareLegend"),
  comparisonSummary: document.querySelector("#comparisonSummary"),
  comparePeriod: document.querySelector("#comparePeriod"),
  compareGroup: document.querySelector("#compareGroup"),
  compareCards: document.querySelector("#compareCards"),
  departmentBars: document.querySelector("#departmentBars"),
  categoryBars: document.querySelector("#categoryBars"),
  orderCards: document.querySelector("#orderCards"),
  skuBody: document.querySelector("#skuBody"),
  inventoryBody: document.querySelector("#inventoryBody"),
  inventorySummary: document.querySelector("#inventorySummary"),
  parentGrid: document.querySelector("#parentGrid"),
  parentsSearch: document.querySelector("#parentsSearch"),
  detailDrawer: document.querySelector("#detailDrawer"),
  dropInbox: document.querySelector("#dropInbox"),
  columnPickerPanel: document.querySelector("#columnPickerPanel"),
  sortMode: document.querySelector("#sortMode"),
  inventoryStateFilter: document.querySelector("#inventoryStateFilter"),
  arrangeColumnsButton: document.querySelector("#arrangeColumnsButton"),
  downloadOrder: document.querySelector("#downloadOrder"),
  downloadSku: document.querySelector("#downloadSku"),
  downloadInventory: document.querySelector("#downloadInventory"),
  downloadParents: document.querySelector("#downloadParents"),
  parentRuleName: document.querySelector("#parentRuleName"),
  parentRuleAliases: document.querySelector("#parentRuleAliases"),
  addParentRuleButton: document.querySelector("#addParentRuleButton"),
  parentRuleList: document.querySelector("#parentRuleList"),
  attributeRuleType: document.querySelector("#attributeRuleType"),
  attributeRuleValue: document.querySelector("#attributeRuleValue"),
  attributeRuleAliases: document.querySelector("#attributeRuleAliases"),
  addAttributeRuleButton: document.querySelector("#addAttributeRuleButton"),
  attributeRuleList: document.querySelector("#attributeRuleList"),
  attributeRuleCount: document.querySelector("#attributeRuleCount"),
  countWorkspace: document.querySelector("#countWorkspace"),
  countWorkspaceEmpty: document.querySelector("#countWorkspaceEmpty"),
  countLaunchCard: document.querySelector("#countLaunchCard"),
  countSummaryStrip: document.querySelector("#countSummaryStrip"),
  activeCountTitle: document.querySelector("#activeCountTitle"),
  activeCountMeta: document.querySelector("#activeCountMeta"),
  countReviewButton: document.querySelector("#countReviewButton"),
  closeCountSessionButton: document.querySelector("#closeCountSessionButton"),
  countSessionModal: document.querySelector("#countSessionModal"),
  countSaveSessionButton: document.querySelector("#countSaveSessionButton"),
  countDeleteSessionButton: document.querySelector("#countDeleteSessionButton"),
  countInputReportButton: document.querySelector("#countInputReportButton"),
  countComparisonReportButton: document.querySelector("#countComparisonReportButton"),
  countSessionBody: document.querySelector("#countSessionBody"),
  countEntryBody: document.querySelector("#countEntryBody"),
  countSearchInput: document.querySelector("#countSearchInput"),
  countSearchButton: document.querySelector("#countSearchButton"),
  countClearSearchButton: document.querySelector("#countClearSearchButton"),
  countSelectedItem: document.querySelector("#countSelectedItem"),
  countQuantityDisplay: document.querySelector("#countQuantityDisplay"),
  countKeyButtons: document.querySelectorAll("[data-count-key]"),
  countSetupModal: document.querySelector("#countSetupModal"),
  countDateInput: document.querySelector("#countDateInput"),
  countVendorInput: document.querySelector("#countVendorInput"),
  countCategoryInput: document.querySelector("#countCategoryInput"),
  countStartButton: document.querySelector("#countStartButton"),
  countCancelButton: document.querySelector("#countCancelButton"),
  countDuplicateModal: document.querySelector("#countDuplicateModal"),
  countDuplicateMessage: document.querySelector("#countDuplicateMessage"),
  countDuplicateAddButton: document.querySelector("#countDuplicateAddButton"),
  countDuplicateResetButton: document.querySelector("#countDuplicateResetButton"),
  countDuplicateCancelButton: document.querySelector("#countDuplicateCancelButton"),
  countReportModal: document.querySelector("#countReportModal"),
  countReportTitle: document.querySelector("#countReportTitle"),
  countReportMeta: document.querySelector("#countReportMeta"),
  countReportHead: document.querySelector("#countReportHead"),
  countReportBody: document.querySelector("#countReportBody"),
  countPdfReportButton: document.querySelector("#countPdfReportButton"),
  countExcelReportButton: document.querySelector("#countExcelReportButton"),
  countContinueButton: document.querySelector("#countContinueButton"),
  countSubmitButton: document.querySelector("#countSubmitButton"),
  zeroNegativeStockButton: document.querySelector("#zeroNegativeStockButton"),
  stockAdjustModal: document.querySelector("#stockAdjustModal"),
  stockAdjustStage1: document.querySelector("#stockAdjustStage1"),
  stockAdjustStage2: document.querySelector("#stockAdjustStage2"),
  stockAdjustStage3: document.querySelector("#stockAdjustStage3"),
  stockAdjustTitle: document.querySelector("#stockAdjustTitle"),
  stockAdjustEyebrow: document.querySelector("#stockAdjustEyebrow"),
  stockAdjustMeta: document.querySelector("#stockAdjustMeta"),
  stockAdjustActionLabel: document.querySelector("#stockAdjustActionLabel"),
  stockAdjustQtyDisplay: document.querySelector("#stockAdjustQtyDisplay"),
  stockAdjustCancelButton: document.querySelector("#stockAdjustCancelButton"),
  confirmDeleteSessionModal: document.querySelector("#confirmDeleteSessionModal"),
  confirmDeleteSessionMessage: document.querySelector("#confirmDeleteSessionMessage"),
  confirmDeleteSessionYes: document.querySelector("#confirmDeleteSessionYes"),
  confirmDeleteSessionNo: document.querySelector("#confirmDeleteSessionNo"),
  confirmSubmitCountModal: document.querySelector("#confirmSubmitCountModal"),
  confirmSubmitCountMessage: document.querySelector("#confirmSubmitCountMessage"),
  confirmSubmitCountYes: document.querySelector("#confirmSubmitCountYes"),
  confirmSubmitCountNo: document.querySelector("#confirmSubmitCountNo"),
  confirmZeroNegModal: document.querySelector("#confirmZeroNegModal"),
  confirmZeroNegMessage: document.querySelector("#confirmZeroNegMessage"),
  confirmZeroNegYes: document.querySelector("#confirmZeroNegYes"),
  confirmZeroNegNo: document.querySelector("#confirmZeroNegNo"),
  adjustLogBody: document.querySelector("#adjustLogBody"),
  exportAdjustPdfButton: document.querySelector("#exportAdjustPdfButton"),
  exportAdjustExcelButton: document.querySelector("#exportAdjustExcelButton"),
  clearAdjustLogButton: document.querySelector("#clearAdjustLogButton"),
  sessionHistoryModal: document.querySelector("#sessionHistoryModal"),
  sessionHistoryVendorFilter: document.querySelector("#sessionHistoryVendorFilter"),
  sessionHistoryPeriodFilter: document.querySelector("#sessionHistoryPeriodFilter"),
  sessionHistoryCloseButton: document.querySelector("#sessionHistoryCloseButton"),
  finalCountReportModal: document.querySelector("#finalCountReportModal"),
  finalReportTitle: document.querySelector("#finalReportTitle"),
  finalReportMeta: document.querySelector("#finalReportMeta"),
  finalReportBody: document.querySelector("#finalReportBody"),
  finalReportPdfButton: document.querySelector("#finalReportPdfButton"),
  finalReportExcelButton: document.querySelector("#finalReportExcelButton"),
  finalReportCloseButton: document.querySelector("#finalReportCloseButton"),
  openSessionHistoryButton: document.querySelector("#openSessionHistoryButton"),
  selectAllInventory: document.querySelector("#selectAllInventory"),
  selectAllSku: document.querySelector("#selectAllSku"),
  countCloseReportButton: document.querySelector("#countCloseReportButton"),
  countInputViewButton: document.querySelector("#countInputViewButton"),
  countComparisonViewButton: document.querySelector("#countComparisonViewButton"),
};

const inventoryColumns = [
  ["pending", "PO"],
  ["code", "Code"],
  ["product", "Item"],
  ["plu", "PLU"],
  ["itemNumber", "Item #"],
  ["sizeAttr", "Sub"],
  ["subType", "Type"],
  ["containerAttr", "Tag"],
  ["category", "Category"],
  ["vendor", "Vendor"],
  ["state", "State"],
  ["addDate", "Add Date"],
  ["stock", "Stock"],
  ["units", "Sold"],
  ["velocity", "SV/day"],
  ["unitCost", "Cost"],
  ["price", "Price"],
  ["inventoryCost", "Cost Total"],
  ["caseSize", "Case"],
  ["reorderMin", "Min"],
  ["reorderMax", "Max"],
  ["needs", "Needs"],
];

const orderColumns = [
  { key: "status",           label: "Status",      defaultOn: true  },
  { key: "pending",          label: "Pending",     defaultOn: true  },
  { key: "code",             label: "Code",        defaultOn: true  },
  { key: "product",          label: "Item",        defaultOn: true  },
  { key: "vendor",           label: "Vendor",      defaultOn: true  },
  { key: "plu",              label: "PLU",         defaultOn: false },
  { key: "velocity",         label: "SV/day",      defaultOn: true  },
  { key: "units",            label: "Sold",        defaultOn: true  },
  { key: "stock",            label: "Stock",       defaultOn: true  },
  { key: "reorderMin",       label: "Min",         defaultOn: true  },
  { key: "reorderMax",       label: "Max",         defaultOn: true  },
  { key: "recommendedOrder", label: "Rec. Order",  defaultOn: true  },
  { key: "caseOrder",        label: "Case Order",  defaultOn: true  },
  { key: "caseSize",         label: "Case Size",   defaultOn: true  },
  { key: "unitCost",         label: "Unit Cost",   defaultOn: false },
  { key: "totalCost",        label: "Total Cost",  defaultOn: true  },
];
if (!state.orderVisibleColumns) {
  state.orderVisibleColumns = Object.fromEntries(orderColumns.map((c) => [c.key, c.defaultOn]));
}

const hoverTooltip = document.createElement("div");
hoverTooltip.className = "row-hover-tooltip";
hoverTooltip.hidden = true;
document.body.append(hoverTooltip);

if (!state.visibleColumns) {
  // Sensible defaults: hide rarely-needed columns to keep table in viewport
  const defaultOff = new Set(["plu","itemNumber","sizeAttr","subType","containerAttr","addDate","inventoryCost"]);
  state.visibleColumns = Object.fromEntries(inventoryColumns.map(([key]) => [key, !defaultOff.has(key)]));
}
if (!state.columnOrder) {
  state.columnOrder = inventoryColumns.map(([key]) => key);
}
const validColumnKeys = inventoryColumns.map(([key]) => key);
state.columnOrder = state.columnOrder.filter((key) => validColumnKeys.includes(key));
validColumnKeys.forEach((key) => {
  if (!state.columnOrder.includes(key)) state.columnOrder.push(key);
  if (!(key in state.visibleColumns)) state.visibleColumns[key] = true;
});

const currency = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const number = new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 });
const svNumber = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function activeTabName() {
  return document.querySelector(".tab-button.active")?.dataset.tab || "dashboard";
}

function saveActiveTabSearch() {
  const tab = activeTabName();
  if (!["dashboard", "inventory", "ordering"].includes(tab)) return;
  state.tabSearches[tab] = els.searchInput?.value || "";
  localStorage.setItem("posDashboardTabSearches:v1", JSON.stringify(state.tabSearches));
}

els.fileInput.addEventListener("change", (event) => loadFiles(event.target.files));
els.folderInput.addEventListener("change", (event) => loadFiles(event.target.files));
els.excelInput.addEventListener("change", (event) => loadExcelFile(event.target.files[0]));
els.dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  els.dropZone.classList.add("dragging");
});
els.dropZone.addEventListener("dragleave", () => els.dropZone.classList.remove("dragging"));
els.dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  els.dropZone.classList.remove("dragging");
  loadDroppedFiles(event.dataTransfer.files);
});
els.dropInbox.addEventListener("dragover", (event) => {
  event.preventDefault();
  els.dropInbox.classList.add("dragging");
});
els.dropInbox.addEventListener("dragenter", (event) => {
  event.preventDefault();
  els.dropInbox.classList.add("dragging");
});
els.dropInbox.addEventListener("dragleave", () => els.dropInbox.classList.remove("dragging"));
els.dropInbox.addEventListener("drop", (event) => {
  event.preventDefault();
  els.dropInbox.classList.remove("dragging");
  loadDroppedFiles(event.dataTransfer.files);
});
els.dropInbox.addEventListener("click", (event) => {
  if (event.target.closest("button, input, select, label, a")) return;
  els.fileInput.click();
});
["dragenter", "dragover", "drop"].forEach((type) => {
  document.addEventListener(type, (event) => {
    const hasFiles = [...(event.dataTransfer?.types || [])].includes("Files");
    if (!hasFiles) return;
    event.preventDefault();
  });
});
document.addEventListener("dragleave", (event) => {
  if (event.target === document.documentElement || event.target === document.body) {
    els.dropInbox.classList.remove("dragging");
  }
});

const renderDebounced = debounce(render, 120);

// Date navigation arrows + period mode
document.querySelector("#dateNavPrev")?.addEventListener("click", () => shiftDateRange(-1));
document.querySelector("#dateNavNext")?.addEventListener("click", () => shiftDateRange(1));
document.querySelector("#datePeriodMode")?.addEventListener("change", (e) => {
  state.datePeriodMode = e.target.value;
});
const renderParentsDebounced = debounce(renderParents, 120);
const syncStickyHeightsDebounced = debounce(syncStickyHeights, 60);
window.addEventListener("resize", syncStickyHeightsDebounced);
// Text/number inputs debounce so fast typing doesn't rebuild 1000+ rows on every keystroke
[els.searchInput, els.leadDays, els.safetyDays, els.daysOfInventory].filter(Boolean).forEach((input) => input.addEventListener("input", renderDebounced));
els.parentsSearch?.addEventListener("input", renderParentsDebounced);
els.searchInput?.addEventListener("input", () => {
  const upper = els.searchInput.value.toUpperCase();
  if (els.searchInput.value !== upper) els.searchInput.value = upper;
  saveActiveTabSearch();
});
// Enter selects all text so next scan/type replaces it immediately
els.searchInput?.addEventListener("keydown", (e) => {
  if (e.key === "Enter") { e.preventDefault(); els.searchInput.select(); }
});
els.searchInput?.addEventListener("focus", () => {
  setTimeout(() => els.searchInput?.select(), 0);
});
els.parentsSearch?.addEventListener("input", () => {
  const upper = els.parentsSearch.value.toUpperCase();
  if (els.parentsSearch.value !== upper) els.parentsSearch.value = upper;
});
// Selects and date pickers render immediately (single discrete change)
[els.startDate, els.endDate, els.departmentFilter, els.categoryFilter, els.vendorFilter,
 els.colorFilter, els.segmentMetric, els.segmentGroup, els.sortMode, els.inventoryStateFilter,
 els.comparePeriod, els.compareGroup].filter(Boolean)
  .forEach((input) => input.addEventListener("input", render));
[els.startDate, els.endDate].filter(Boolean).forEach((input) => {
  input.addEventListener("input", () => {
    state.activePresetDays = null;
  });
});

els.compareToggle?.addEventListener("change", () => {
  if (els.compareLegend) els.compareLegend.hidden = !els.compareToggle.checked;
  renderTrend();
});

[els.searchInput, els.startDate, els.endDate, els.leadDays, els.safetyDays].forEach((input) => {
  input.addEventListener("focus", () => input.select?.());
  input.addEventListener("dblclick", () => input.select?.());
  input.addEventListener("keydown", (event) => {
    if (event.key === "Delete") {
      input.value = "";
      input.dispatchEvent(new Event("input", { bubbles: true }));
    }
  });
});

[els.departmentFilter, els.categoryFilter, els.vendorFilter, els.colorFilter, els.inventoryStateFilter].forEach((select) => {
  select.addEventListener("keydown", (event) => {
    if (event.key === "Delete") {
      select.value = "";
      select.dispatchEvent(new Event("input", { bubbles: true }));
    }
  });
});

els.clearFiltersButton.addEventListener("click", () => clearFilters());
els.clearFilterButtons.forEach((button) => {
  button.addEventListener("click", (event) => {
    event.preventDefault();
    event.stopPropagation();
    clearSingleFilter(button.dataset.clearFilter);
  });
});
els.chooseSalesButton.addEventListener("click", () => els.fileInput.click());
els.chooseFolderButton.addEventListener("click", () => els.folderInput.click());
els.chooseExcelButton.addEventListener("click", () => els.excelInput.click());
els.createPoShortcut?.addEventListener("click", () => {
  // Copy current search/filter context to ordering tab
  state.tabSearches["ordering"] = els.searchInput?.value || "";
  switchTab("ordering");
});
els.countLaunchCard?.addEventListener("click", openCountSetupModal);
els.closeCountSessionButton?.addEventListener("click", closeActiveCountSession);
els.countReviewButton?.addEventListener("click", () => openCountReport(state.activeCountSession?.id, "input"));
els.countSaveSessionButton?.addEventListener("click", saveCountSession);
els.countDeleteSessionButton?.addEventListener("click", deleteCountSession);
els.countInputReportButton?.addEventListener("click", () => openCountReport(state.activeCountSession?.id, "input"));
els.countComparisonReportButton?.addEventListener("click", () => openCountReport(state.activeCountSession?.id, "comparison"));
els.countStartButton?.addEventListener("click", startCountSessionFromModal);
els.countCancelButton?.addEventListener("click", closeCountSetupModal);
els.countSetupModal?.addEventListener("click", (event) => {
  if (event.target === els.countSetupModal) closeCountSetupModal();
});
els.countSearchButton?.addEventListener("click", handleCountLookup);
els.countClearSearchButton?.addEventListener("click", clearCountLookup);
els.countSearchInput?.addEventListener("keydown", (event) => {
  const dropdown = document.querySelector("#countSearchDropdown");
  if (dropdown && !dropdown.hidden) {
    const items = [...dropdown.querySelectorAll(".count-dd-item:not(.count-dd-out)")];
    const active = dropdown.querySelector(".count-dd-item--active");
    if (event.key === "ArrowDown") {
      event.preventDefault();
      const next = active ? (items.indexOf(active) + 1) % items.length : 0;
      dropdown.querySelectorAll(".count-dd-item").forEach((el) => el.classList.remove("count-dd-item--active"));
      items[next]?.classList.add("count-dd-item--active");
      return;
    }
    if (event.key === "ArrowUp") {
      event.preventDefault();
      const prev = active ? (items.indexOf(active) - 1 + items.length) % items.length : items.length - 1;
      dropdown.querySelectorAll(".count-dd-item").forEach((el) => el.classList.remove("count-dd-item--active"));
      items[prev]?.classList.add("count-dd-item--active");
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
      const chosen = dropdown.querySelector(".count-dd-item--active") || items[0];
      if (chosen?.dataset.code) { selectCountDropdownItem(chosen.dataset.code); return; }
    }
    if (event.key === "Escape") { hideCountDropdown(); return; }
  }
  if (event.key === "Enter") {
    event.preventDefault();
    event.stopPropagation();
    hideCountDropdown();
    handleCountLookup();
    return;
  }
  if (event.key === "Delete") {
    event.preventDefault();
    els.countSearchInput.value = "";
    hideCountDropdown();
  }
});
// Debounced count search — avoids scanning 25k items on every keystroke
const renderCountDropdownDebounced = debounce((val) => renderCountDropdown(val), 120);
els.countSearchInput?.addEventListener("input", () => {
  renderCountDropdownDebounced(els.countSearchInput.value);
});
els.countSearchInput?.addEventListener("focus", () => els.countSearchInput.select?.());
els.countSearchInput?.addEventListener("click", () => els.countSearchInput.select?.());
els.countKeyButtons?.forEach((button) => button.addEventListener("click", () => handleCountKey(button.dataset.countKey)));
els.countDuplicateModal?.addEventListener("click", (event) => {
  if (event.target === els.countDuplicateModal) closeDuplicateCountModal();
});

// When the scan/count modal is open, clicking anywhere that isn't an interactive element
// forces focus back to the search/scan bar so scanning always works immediately
els.countSessionModal?.addEventListener("click", (event) => {
  if (els.countDuplicateModal && !els.countDuplicateModal.hidden) return;
  if (els.countReportModal && !els.countReportModal.hidden) return;
  const interactive = event.target.closest("button, input, select, textarea, a, label, [data-count-key], .count-keypad");
  if (!interactive) {
    setTimeout(() => focusCountSearch(), 0);
  }
});
els.countDuplicateAddButton?.addEventListener("click", () => resolveDuplicateCount("add"));
els.countDuplicateResetButton?.addEventListener("click", () => resolveDuplicateCount("reset"));
els.countDuplicateCancelButton?.addEventListener("click", closeDuplicateCountModal);
els.countReportModal?.addEventListener("click", (event) => {
  if (event.target === els.countReportModal) closeCountReport();
});
els.countCloseReportButton?.addEventListener("click", closeCountReport);
document.querySelector("#countPdfReportButton")?.addEventListener("click", () => exportCountReportPdf());
document.querySelector("#countExcelReportButton")?.addEventListener("click", () => exportCountReportExcel());
document.querySelector("#countContinueButton")?.addEventListener("click", () => continueCountFromReport());
document.querySelector("#countSubmitButton")?.addEventListener("click", () => openConfirmSubmitCount());
document.querySelector("#zeroNegativeStockButton")?.addEventListener("click", () => openConfirmZeroNeg());
document.querySelector("#confirmDeleteSessionYes")?.addEventListener("click", () => confirmDeleteSavedSession());
document.querySelector("#confirmDeleteSessionNo")?.addEventListener("click", () => { document.querySelector("#confirmDeleteSessionModal").hidden = true; });
document.querySelector("#confirmSubmitCountYes")?.addEventListener("click", () => submitAndApplyCount());
document.querySelector("#confirmSubmitCountNo")?.addEventListener("click", () => { document.querySelector("#confirmSubmitCountModal").hidden = true; });
document.querySelector("#confirmZeroNegYes")?.addEventListener("click", () => applyZeroNegatives());
document.querySelector("#confirmZeroNegNo")?.addEventListener("click", () => { document.querySelector("#confirmZeroNegModal").hidden = true; });
document.querySelector("#stockAdjustCancelButton")?.addEventListener("click", () => closeStockAdjustModal());
document.querySelector("#exportAdjustPdfButton")?.addEventListener("click", () => exportAdjustLogPdf());
document.querySelector("#exportAdjustExcelButton")?.addEventListener("click", () => exportAdjustLogExcel());
document.querySelector("#clearAdjustLogButton")?.addEventListener("click", () => {
  if (!confirm("Clear the entire stock adjustment log?")) return;
  state.adjustmentLog = [];
  localStorage.setItem("posDashboardAdjustLog:v1", "[]");
  renderAdjustLog();
  showToast("Adjustment log cleared.", 2400, "warning");
});

// Session history modal
document.querySelector("#openSessionHistoryButton")?.addEventListener("click", () => openSessionHistoryModal());
document.querySelector("#sessionHistoryCloseButton")?.addEventListener("click", () => { document.querySelector("#sessionHistoryModal").hidden = true; });
document.querySelector("#sessionHistoryModal")?.addEventListener("click", (e) => { if (e.target === document.querySelector("#sessionHistoryModal")) document.querySelector("#sessionHistoryModal").hidden = true; });
document.querySelector("#sessionHistoryVendorFilter")?.addEventListener("change", () => renderCountSessionRows());
document.querySelector("#sessionHistoryPeriodFilter")?.addEventListener("change", () => renderCountSessionRows());

// Final count report modal
document.querySelector("#finalReportCloseButton")?.addEventListener("click", () => { document.querySelector("#finalCountReportModal").hidden = true; });
document.querySelector("#finalReportPdfButton")?.addEventListener("click", () => exportFinalCountReportPdf());
document.querySelector("#finalReportExcelButton")?.addEventListener("click", () => exportFinalCountReportExcel());

// Select-all for ordering and sku tabs (inventory select-all is wired in renderInventoryHeader)
document.querySelector("#selectAllSku")?.addEventListener("change", (e) => {
  const checked = e.target.checked;
  document.querySelectorAll("#skuBody .row-checkbox").forEach((cb) => {
    cb.checked = checked;
    const code = cb.dataset.code;
    if (checked) state.selectedSkuCodes.add(code);
    else state.selectedSkuCodes.delete(code);
  });
});

// Stock action buttons (stage 1 of adjust modal)
document.querySelector("#stockActionAdd")?.addEventListener("click", () => beginStockAdjustQty("add"));
document.querySelector("#stockActionRemove")?.addEventListener("click", () => beginStockAdjustQty("remove"));
document.querySelector("#stockActionSet")?.addEventListener("click", () => beginStockAdjustQty("set"));

// Stock keypad
document.querySelectorAll("[data-stock-key]").forEach((btn) => {
  btn.addEventListener("click", () => handleStockKey(btn.dataset.stockKey));
});

// Reason buttons (stage 3)
document.querySelectorAll(".stock-reason-btn").forEach((btn) => {
  btn.addEventListener("click", () => finalizeStockAdjust(btn.dataset.reason));
});
els.countInputViewButton?.addEventListener("click", () => {
  state.countReportMode = "input";
  openCountReport(state.activeCountSession?.id || state.countSessions[0]?.id, "input");
});
els.countComparisonViewButton?.addEventListener("click", () => {
  state.countReportMode = "comparison";
  openCountReport(state.activeCountSession?.id || state.countSessions[0]?.id, "comparison");
});
els.arrangeColumnsButton.addEventListener("click", () => {
  state.arrangeColumns = !state.arrangeColumns;
  els.arrangeColumnsButton.classList.toggle("active-edit", state.arrangeColumns);
  els.arrangeColumnsButton.textContent = state.arrangeColumns ? "Done arranging" : "Arrange columns";
  renderInventory();
});
els.addParentRuleButton.addEventListener("click", () => addParentRule());
els.addAttributeRuleButton.addEventListener("click", () => addAttributeRule());

document.querySelectorAll("[data-tab]").forEach((button) => {
  button.addEventListener("click", () => switchTab(button.dataset.tab));
});

document.querySelectorAll("[data-sort]").forEach((header) => {
  header.addEventListener("click", () => {
    const key = header.dataset.sort;
    state.inventorySort = {
      key,
      dir: state.inventorySort.key === key && state.inventorySort.dir === "asc" ? "desc" : "asc",
    };
    renderInventory();
  });
});

document.addEventListener("keydown", (event) => {
  // Stock adjust modal: Esc closes at any stage; digits/Enter/Back work on stage 2
  if (els.stockAdjustModal && !els.stockAdjustModal.hidden) {
    if (event.key === "Escape") { event.preventDefault(); closeStockAdjustModal(); return; }
    if (!els.stockAdjustStage2.hidden) {
      if (/^\d$/.test(event.key)) { event.preventDefault(); handleStockKey(event.key); return; }
      if (event.key === "Backspace") { event.preventDefault(); handleStockKey("back"); return; }
      if (event.key === ".") { event.preventDefault(); handleStockKey("."); return; }
      if (event.key === "Enter") { event.preventDefault(); handleStockKey("enter"); return; }
    }
    return; // swallow all other keys while modal is open
  }
  if (document.querySelector("#vendorPoModal") && !document.querySelector("#vendorPoModal").hidden) {
    if (event.key === "Escape") { document.querySelector("#vendorPoModal").hidden = true; return; }
  }
  // Report modals
  for (const id of ["reportAdjustModal","reportPoModal","reportCountModal"]) {
    const el = document.querySelector("#" + id);
    if (el && !el.hidden && event.key === "Escape") { el.hidden = true; return; }
  }
  if (document.querySelector("#vendorRuleModal") && !document.querySelector("#vendorRuleModal").hidden) {
    if (event.key === "Escape") { document.querySelector("#vendorRuleModal").hidden = true; return; }
    if (event.key === "Enter") { event.preventDefault(); saveVendorRule(); return; }
  }
  if (document.querySelector("#sessionHistoryModal") && !document.querySelector("#sessionHistoryModal").hidden) {
    if (event.key === "Escape") { return; }
  }
  if (document.querySelector("#finalCountReportModal") && !document.querySelector("#finalCountReportModal").hidden) {
    if (event.key === "Escape") { document.querySelector("#finalCountReportModal").hidden = true; return; }
  }
  // Confirm modals: Esc dismisses
  if (document.querySelector("#confirmDeleteSessionModal") && !document.querySelector("#confirmDeleteSessionModal").hidden) {
    if (event.key === "Escape") { document.querySelector("#confirmDeleteSessionModal").hidden = true; return; }
  }
  if (document.querySelector("#confirmSubmitCountModal") && !document.querySelector("#confirmSubmitCountModal").hidden) {
    if (event.key === "Escape") { document.querySelector("#confirmSubmitCountModal").hidden = true; return; }
  }
  if (document.querySelector("#confirmZeroNegModal") && !document.querySelector("#confirmZeroNegModal").hidden) {
    if (event.key === "Escape") { document.querySelector("#confirmZeroNegModal").hidden = true; return; }
  }
  if (!els.countSetupModal.hidden && event.key === "Escape") {
    closeCountSetupModal();
    return;
  }
  // Enter in setup modal navigates: date → vendor → category → status → start
  if (!els.countSetupModal.hidden && event.key === "Enter") {
    event.preventDefault();
    const focused = document.activeElement;
    if (focused === els.countDateInput) {
      els.countVendorInput.focus();
    } else if (focused === els.countVendorInput) {
      els.countCategoryInput.focus();
    } else if (focused === els.countCategoryInput) {
      const statusEl = document.querySelector("#countStatusInput");
      if (statusEl) statusEl.focus(); else startCountSessionFromModal();
    } else {
      startCountSessionFromModal();
    }
    return;
  }
  if (!els.countDuplicateModal.hidden && event.key === "Escape") {
    closeDuplicateCountModal();
    return;
  }
  if (!els.countSessionModal.hidden) {
    if (event.key === "Escape") {
      event.preventDefault();
      return;
    }
  }
  if (!els.countReportModal.hidden && event.key === "Escape") {
    closeCountReport();
    return;
  }
  if (!els.countDuplicateModal.hidden) return;
  if (!els.countSessionModal.hidden && event.target === els.countSearchInput && event.key === "Enter") return;
  if (!els.countReportModal.hidden) return;
  if (document.querySelector(".tab-button.active")?.dataset.tab === "counts" && state.activeCountSession && state.countStage === "qty") {
    if (/^\d$/.test(event.key)) {
      event.preventDefault();
      handleCountKey(event.key);
      return;
    }
    if (event.key === "Backspace") {
      event.preventDefault();
      handleCountKey("back");
      return;
    }
    if (event.key === ".") {
      event.preventDefault();
      handleCountKey(".");
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      applyCountEntry();
      return;
    }
  }
  if (event.key === "Escape") {
    els.detailDrawer.hidden = true;
    state._activeDetailCode = "";
    document.querySelectorAll("details[open]").forEach((detail) => detail.removeAttribute("open"));
  }
});

document.addEventListener("click", (event) => {
  const interactiveTarget = event.target.closest("button, input, select, summary, details, a, label, td, th, tr, .detail-drawer, .row-hover-tooltip");
  if (
    !els.detailDrawer.hidden &&
    !event.target.closest(".detail-drawer") &&
    !event.target.closest("tbody tr") &&
    !event.target.closest("[data-detail-code]") &&
    !event.target.closest("#datePresets")
  ) {
    els.detailDrawer.hidden = true;
    state._activeDetailCode = "";
  }
  if (!interactiveTarget && event.target.closest(".app, .panel, .metrics, .controls, .tab-view, .sticky-pills")) {
    els.searchInput?.focus();
  }
  document.querySelectorAll(".column-picker[open]").forEach((detail) => {
    if (!detail.contains(event.target)) detail.removeAttribute("open");
  });
  document.querySelectorAll(".detail-picker[open]").forEach((detail) => {
    if (!detail.contains(event.target)) detail.removeAttribute("open");
  });
});

els.downloadOrder.addEventListener("click", () => {
  downloadCsv("recommended-order.csv", currentOrderRows());
});
document.querySelector("#exportPoExcel")?.addEventListener("click", () => exportPoExcel());
document.querySelector("#exportPoPdf")?.addEventListener("click", () => exportPoPdf());
els.downloadSku.addEventListener("click", () => downloadCsv("sku-performance.csv", state.filteredSkus));
els.downloadInventory.addEventListener("click", () => downloadCsv("all-inventory.csv", state.inventoryRows));
els.downloadParents.addEventListener("click", () => downloadCsv("parent-styles.csv", state.parentRows));
els.inventoryBody?.addEventListener("click", (event) => {
  // Checkbox toggle
  const cb = event.target.closest(".row-checkbox");
  if (cb) {
    event.stopPropagation();
    if (cb.checked) state.selectedInventoryCodes.add(cb.dataset.code);
    else state.selectedInventoryCodes.delete(cb.dataset.code);
    return;
  }
  const copyButton = event.target.closest(".copy-code");
  if (copyButton) {
    event.stopPropagation();
    copyText(copyButton.dataset.copyCode || copyButton.textContent.trim(), copyButton);
    return;
  }
  if (event.target.closest(".mini-input, .reset-override")) return;
  // Stock cell click → open stock adjust modal
  const stockCell = event.target.closest("td[data-col='stock']");
  if (stockCell) {
    const row = stockCell.closest("tr[data-item-code]");
    if (row) {
      const item = state._inventoryRowIndex.get(codeKey(row.dataset.itemCode));
      if (item) { openStockAdjustModal(item); return; }
    }
  }
  const row = event.target.closest("tr[data-item-code]");
  if (!row) return;
  const item = state._inventoryRowIndex.get(codeKey(row.dataset.itemCode));
  if (item) showDetail(item);
});
els.inventoryBody?.addEventListener("mousedown", (event) => {
  if (event.target.closest(".mini-input, .copy-code, .reset-override")) {
    event.stopPropagation();
  }
});
els.inventoryBody?.addEventListener("mousemove", (event) => {
  const row = event.target.closest("tr[data-item-code]");
  if (!row) { hideHoverTooltip(); return; }
  // Build tooltip lazily on first hover (deferred flag set during render)
  if (row.dataset.tooltipDeferred) {
    delete row.dataset.tooltipDeferred;
    const item = state._inventoryRowIndex?.get(codeKey(row.dataset.itemCode));
    if (item) {
      const windows = salesWindowsFor(item.code).filter((e) => ["7D","30D","60D","90D","6M","365D"].includes(e.label));
      row.dataset.tooltipHtml = `
        <div class="row-hover-tooltip__title">${escapeHtml(item.product)}</div>
        <div class="row-hover-tooltip__line">Vendor: ${escapeHtml(item.vendor||"-")}</div>
        <div class="row-hover-tooltip__line">Sold: ${number.format(item.units)}</div>
        <div class="row-hover-tooltip__line">Stock: ${number.format(item.stock)}</div>
        <div class="row-hover-tooltip__line">Cost: ${currency.format(item.unitCost)}</div>
        <div class="row-hover-tooltip__line">Price: ${currency.format(item.price)}</div>
        <div class="row-hover-tooltip__windows">${windows.map((e)=>`<span>${e.label}: ${number.format(e.units)}</span>`).join('<span class="row-hover-tooltip__sep">|</span>')}</div>`;
    }
  }
  if (row.dataset.tooltipHtml) showHoverTooltip(row.dataset.tooltipHtml, event);
  else hideHoverTooltip();
});
els.inventoryBody?.addEventListener("mouseleave", () => hideHoverTooltip());
els.orderCards?.addEventListener("click", (event) => {
  const copyButton = event.target.closest(".copy-code");
  if (copyButton) {
    event.stopPropagation();
    copyText(copyButton.dataset.copyCode || copyButton.textContent.trim(), copyButton);
    return;
  }
  const card = event.target.closest("[data-detail-code]");
  if (!card) return;
  const item = findCurrentItemByCode(card.dataset.detailCode);
  if (item) showDetail(item);
});
els.skuBody?.addEventListener("click", (event) => {
  const row = event.target.closest("tr[data-detail-code]");
  if (!row) return;
  const item = findCurrentItemByCode(row.dataset.detailCode);
  if (item) showDetail(item);
});
renderColumnPicker();
renderAttributeRules();
document.querySelectorAll(".rules-box, .parent-rule-editor").forEach((node) => {
  node.hidden = !ENABLE_CUSTOM_PARENT_RULES && !ENABLE_CUSTOM_ATTRIBUTE_RULES;
});

document.addEventListener("input", (event) => {
  const input = event.target.closest("[data-reorder-field]");
  if (!input) return;
  const code = input.dataset.code;
  const field = input.dataset.reorderField;
  const val = input.value.trim();
  if (val === "") {
    // Empty field = clear override, revert to auto
    if (state.reorderOverrides[code]) {
      delete state.reorderOverrides[code][field];
      if (!Object.keys(state.reorderOverrides[code]).length) delete state.reorderOverrides[code];
    }
    showToast(`${field === "min" ? "Min" : "Max"} reset to auto for ${code}`);
  } else {
    state.reorderOverrides[code] = state.reorderOverrides[code] || {};
    state.reorderOverrides[code][field] = toNumber(val);
    showToast(`ðŸ”’ Manual ${field === "min" ? "Min" : "Max"} set â€” clear field to restore auto`);
  }
  localStorage.setItem("posDashboardReorderOverrides:v1", JSON.stringify(state.reorderOverrides));
  bumpDataStamp();
  renderDebounced();
});

// Reset-override button: click the ðŸ”’ icon to clear that field's override
document.addEventListener("click", (event) => {
  const resetBtn = event.target.closest(".reset-override");
  if (!resetBtn) return;
  event.stopPropagation();
  const { code, field } = resetBtn.dataset;
  if (state.reorderOverrides[code]) {
    delete state.reorderOverrides[code][field];
    if (!Object.keys(state.reorderOverrides[code]).length) delete state.reorderOverrides[code];
    localStorage.setItem("posDashboardReorderOverrides:v1", JSON.stringify(state.reorderOverrides));
  }
  showToast(`${field === "min" ? "Min" : "Max"} restored to auto for ${code}`);
  bumpDataStamp();
  render();
});

async function loadDroppedFiles(fileList) {
  const files = [...fileList];
  const excelFile = files.find((file) => /\.(xlsx|xls)$/i.test(file.name));
  const csvFiles = files.filter((file) => /\.csv$/i.test(file.name));
  if (csvFiles.length) await loadFiles(csvFiles);
  if (excelFile) await loadExcelFile(excelFile);
}

async function loadFiles(fileList) {
  const files = [...fileList].filter((file) => file.name.toLowerCase().endsWith(".csv"));
  if (!files.length) return;
  els.fileCount.textContent = `Reading ${files.length} files...`;
  try {
    let selectedInventory = null;
    let skippedDuplicates = 0;
    let processedFiles = 0;
    for (const file of files) {
      const isCurrentInventoryFile = /(^|[^a-z])current[_ ]inventory/i.test(file.name);
      const signature = fileSignature(file);
      if (!isCurrentInventoryFile && state._loadedFileSignatures.has(signature)) {
        skippedDuplicates += 1;
        continue;
      }
      const date = dateFromFileName(file.name);
      if (!date) continue;
      const rows = parseCsv(await file.text());
      if (isCurrentInventoryFile) {
        const normalized = rows.map((row) => normalizeInventoryRow(row, date)).filter((row) => row.code || row.product);
        if (normalized.length && (!selectedInventory || date.iso >= selectedInventory.date)) {
          selectedInventory = { date: date.iso, rows: normalized };
        }
      } else {
        state.rawSales = state.rawSales.filter((row) => row.date !== date.iso);
        state.rawSales.push(...rows.map((row) => normalizeSalesRow(row, date)).filter(Boolean));
      }
      state._loadedFileSignatures.add(signature);
      processedFiles += 1;
      els.fileCount.textContent = `Read ${file.name}`;
    }

    if (!processedFiles) {
      if (skippedDuplicates && (state.rawSales.length || state.latestInventory.size || state.excelItems.size)) {
        const synced = await syncSharedDataToSupabase({ silent: true });
        els.fileCount.textContent = synced
          ? `Synced existing local data to shared cloud storage.`
          : `Duplicate files skipped. Shared sync still needs attention.`;
        return;
      }
      els.fileCount.textContent = skippedDuplicates
        ? `Skipped ${skippedDuplicates} duplicate file${skippedDuplicates === 1 ? "" : "s"}`
        : "No new CSV files were imported.";
      return;
    }

    if (selectedInventory) {
      const currentLatest = latestInventoryDate();
      if (!currentLatest || selectedInventory.date >= currentLatest) {
        state.inventories = new Map([[selectedInventory.date, selectedInventory.rows]]);
      }
    }
    state.dates = [...new Set(state.rawSales.map((row) => row.date))].sort();
    buildLatestInventory();
    bumpDataStamp();
    updateFilterOptions();
    setDefaultDates();
    await savePersistedState();
    await syncSharedDataToSupabase({ silent: true });
    render();
    if (skippedDuplicates) {
      els.fileCount.textContent = `${fileSummary()} â€¢ skipped ${skippedDuplicates} duplicate file${skippedDuplicates === 1 ? "" : "s"}`;
    }
  } catch (error) {
    console.error("CSV import failed", error);
    els.fileCount.textContent = `CSV import failed: ${error.message || error}`;
  }
}

async function loadExcelFile(file) {
  if (!file) return;
  try {
    const signature = fileSignature(file);
    if (state._loadedFileSignatures.has(signature)) {
      const synced = await syncSharedDataToSupabase({ productsOnly: true, silent: true });
      els.excelStatus.textContent = synced
        ? `Skipped duplicate Excel file: ${file.name}. Synced existing product data to cloud.`
        : `Skipped duplicate Excel file: ${file.name}`;
      return;
    }
    els.excelStatus.textContent = `Reading ${file.name}...`;
    const xlsx = await ensureXlsxReader();
    if (!xlsx) {
      els.excelStatus.textContent = "Excel reader could not load. Check internet access or save the workbook as CSV for now.";
      return;
    }
    const workbook = xlsx.read(await file.arrayBuffer(), { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { defval: "" });
    state.excelItems = new Map();
    state.excelByPlu = new Map();
    state.excelByItemNumber = new Map();
    rows.forEach((row) => {
      const item = normalizeExcelRow(row);
      addExcelIndex(item);
    });
    state._loadedFileSignatures.add(signature);
    updateFilterOptions();
    updateInventoryStateFilter();
    bumpDataStamp();
    els.excelStatus.textContent = `${number.format(state.excelItems.size)} Excel items imported for ordering fields.`;
    await savePersistedState();
    await syncSharedDataToSupabase({ productsOnly: true, silent: true });
    render();
  } catch (error) {
    console.error("Excel import failed", error);
    els.excelStatus.textContent = `Excel import failed: ${error.message || error}`;
  }
}

function ensureXlsxReader() {
  if (window.XLSX) return Promise.resolve(window.XLSX);
  return new Promise((resolve) => {
    const existing = document.querySelector("script[data-xlsx-fallback]");
    const script = existing || document.createElement("script");
    const done = () => resolve(window.XLSX || null);
    script.onload = done;
    script.onerror = () => resolve(null);
    if (!existing) {
      script.dataset.xlsxFallback = "true";
      script.src = "https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js";
      document.head.append(script);
    }
    setTimeout(done, 10000);
  });
}

function addExcelIndex(item) {
  if (item.code) {
    state.excelItems.set(item.code, item);
    state.excelItems.set(codeKey(item.code), item);
  }
  if (item.plu) {
    state.excelByPlu.set(item.plu, item);
    state.excelByPlu.set(codeKey(item.plu), item);
  }
  if (item.itemNumber) {
    state.excelByItemNumber.set(item.itemNumber, item);
    state.excelByItemNumber.set(codeKey(item.itemNumber), item);
  }
}

function rebuildExcelIndexes() {
  state.excelByPlu = new Map();
  state.excelByItemNumber = new Map();
  [...state.excelItems.values()].forEach((item) => addExcelIndex(item));
}

function latestInventoryDate() {
  return [...state.inventories.keys()].sort().at(-1) || "";
}

function findExcelFor(item = {}) {
  return (
    state.excelItems.get(item.code) ||
    state.excelItems.get(codeKey(item.code)) ||
    state.excelByPlu.get(item.plu) ||
    state.excelByPlu.get(codeKey(item.plu)) ||
    state.excelByItemNumber.get(item.itemNumber) ||
    state.excelByItemNumber.get(codeKey(item.itemNumber)) ||
    state.excelItems.get(item.plu) ||
    state.excelItems.get(item.itemNumber) ||
    {}
  );
}

function normalizeExcelRow(row) {
  const code = normalizeCode(row.code ?? row.CODE);
  return {
    code,
    product: cleanCell(row.item_name),
    vendor: cleanCell(row.vendor_name),
    processingTime: toNumber(row.processing_time),
    leadTime: toNumber(row.lead_time),
    safetyStock: toNumber(row.safety_stock),
    daysOfInventory: toNumber(row.days_of_inventory),
    saleWindowSum: toNumber(row.sale_window_sum),
    saleVelocity: toNumber(row.sale_velocity),
    stock: toNumber(row.stock),
    reorderMin: toNumber(row.reorder_qty_min),
    reorderMax: toNumber(row.reorder_qty_max),
    daysBeforeRestock: toNumber(row.days_before_restock),
    state: cleanCell(row.state),
    plu: cleanCell(row.PLU ?? row.plu),
    itemNumber: cleanCell(row.item_number),
    category: cleanCell(row.category),
    addDate: cleanCell(row.add_date ?? row["ADD DATE"] ?? row["Add Date"] ?? row.addDate),
    cost: toNumber(row.cost),
    price: toNumber(row.price),
    caseSize: toNumber(row.case_size) || 1,
    maxOrderQty: toNumber(row.max_order_qty),
    qtyNeeded: toNumber(row.qty_needed),
    orderPendingId: cleanCell(row.order_pending_id),
    orderPendingStale: cleanCell(row.order_pending_stale),
    poPendingClearedTimes: toNumber(row.PO_pending_cleared_times),
    overrideValues: cleanCell(row.override_values),
  };
}

function parseCsv(text) {
  const lines = String(text || "").replace(/^\uFEFF/, "").split(/\r\n|\n|\r/).filter((line) => line.trim() !== "");
  if (!lines.length) return [];
  const headers = parseCsvLine(lines.shift()).map((header, index) => cleanHeader(header) || `H${index + 1}`);
  return lines.map((line) => {
    const values = alignCsvValues(headers, parseCsvLine(line));
    return Object.fromEntries(headers.map((header, index) => [header, cleanCell(values[index])]));
  });
}

function parseCsvLine(line) {
  let row = [];
  let cell = "";
  let insideQuotes = false;
  for (let i = 0; i < line.length; i += 1) {
    const char = line[i];
    const next = line[i + 1];
    if (char === '"' && insideQuotes && next === '"') {
      cell += '"';
      i += 1;
    } else if (char === '"' && insideQuotes && (next === "," || next === undefined)) {
      insideQuotes = false;
    } else if (char === '"' && !insideQuotes && cell === "") {
      insideQuotes = true;
    } else if (char === "," && !insideQuotes) {
      row.push(cell);
      cell = "";
    } else {
      cell += char;
    }
  }
  row.push(cell);
  return row;
}

function alignCsvValues(headers, values) {
  if (values.length <= headers.length) return values;
  const mergeIndex = ["NAME", "PRODUCT", "item_name"].map((name) => headers.indexOf(name)).find((index) => index > -1);
  if (mergeIndex === undefined || mergeIndex < 0) return values;
  const extra = values.length - headers.length;
  return values
    .slice(0, mergeIndex)
    .concat(values.slice(mergeIndex, mergeIndex + extra + 1).join(","))
    .concat(values.slice(mergeIndex + extra + 1));
}

function normalizeSalesRow(row, date) {
  const code = normalizeCode(row.CODE);
  const product = cleanCell(row.PRODUCT);
  const department = cleanCell(row.DEPARTMENT);
  if (!code && !product && !department) return null;
  if (!code && product && !department) return null;
  return {
    date: date.iso,
    code,
    product,
    department: department || "Unassigned",
    category: cleanCell(row["::CAT::"]) || "Unassigned",
    vendor: cleanCell(row["::VENDOR::"]) || "Unassigned",
    units: toNumber(row.QTY),
    sales: toNumber(row.SBTLS),
    cost: toNumber(row.COST),
    profit: toNumber(row.PRF),
  };
}

function normalizeInventoryRow(row, date) {
  return {
    date: date.iso,
    code: normalizeCode(row.CODE),
    category: cleanCell(row.CATEGORY),
    product: cleanCell(row.NAME),
    plu: cleanCell(row.PLU),
    itemNumber: cleanCell(row["ITEM NUMBER"]),
    price: toNumber(row.PRICE),
    cost: toNumber(row.COST),
    stock: toNumber(row.STOCK),
    vendor: cleanCell(row.VENDOR),
    vendorCode: cleanCell(row["VENDOR CODE"]),
    color: cleanCell(row.COLOR),
    size: cleanCell(row.SIZE),
    length: cleanCell(row.LENGTH),
    manufacture: cleanCell(row.MANUFACTURE),
    memo: cleanCell(row.MEMO),
  };
}

function buildLatestInventory() {
  state.latestInventory = new Map();
  [...state.inventories.entries()].sort(([a], [b]) => a.localeCompare(b)).forEach(([, rows]) => {
    rows.forEach((row) => {
      if (row.code) state.latestInventory.set(codeKey(row.code), row);
    });
  });
}

function inventoryIndexForDate(asOfDate) {
  const index = new Map();
  [...state.inventories.entries()]
    .filter(([date]) => date <= asOfDate)
    .sort(([a], [b]) => a.localeCompare(b))
    .forEach(([, rows]) => {
      rows.forEach((row) => {
        if (row.code) index.set(codeKey(row.code), row);
      });
  });
  return index.size ? index : state.latestInventory;
}

function rangeDayCount(start, end) {
  if (!start || !end) return Math.max(filteredSalesDates().length, 1);
  const startMs = new Date(`${start}T00:00:00`).getTime();
  const endMs = new Date(`${end}T00:00:00`).getTime();
  if (!Number.isFinite(startMs) || !Number.isFinite(endMs) || endMs < startMs) return 1;
  return Math.max(Math.round((endMs - startMs) / 86400000) + 1, 1);
}

function formatVelocity(value) {
  return svNumber.format(Number(value || 0));
}

function findCurrentItemByCode(code) {
  const key = codeKey(code);
  return (
    state._inventoryRowIndex.get(key) ||
    state._filteredSkuIndex.get(key) ||
    state.filteredSkus.find((item) => codeKey(item.code) === key) ||
    state.inventoryRows.find((item) => codeKey(item.code) === key) ||
    null
  );
}

function showHoverTooltip(html, event) {
  if (!html) {
    hideHoverTooltip();
    return;
  }
  hoverTooltip.innerHTML = html;
  hoverTooltip.hidden = false;
  const offset = 18;
  const { clientX, clientY } = event;
  const maxX = window.innerWidth - hoverTooltip.offsetWidth - 12;
  const maxY = window.innerHeight - hoverTooltip.offsetHeight - 12;
  hoverTooltip.style.left = `${Math.max(12, Math.min(clientX + offset, maxX))}px`;
  hoverTooltip.style.top = `${Math.max(12, Math.min(clientY + offset, maxY))}px`;
}

function hideHoverTooltip() {
  hoverTooltip.hidden = true;
}

function syncStickyHeights() {
  const root = document.documentElement;
  const commandBar = document.querySelector(".command-bar");
  const metrics = document.querySelector(".metrics");
  const filters = document.querySelector(".sticky-filters");
  if (commandBar) root.style.setProperty("--command-bar-height", `${commandBar.offsetHeight}px`);
  if (metrics) root.style.setProperty("--metrics-height", `${metrics.offsetHeight}px`);
  if (filters) root.style.setProperty("--filters-height", `${filters.offsetHeight}px`);
}

function mountInventoryQuickTools() {
  if (!els.inventoryQuickTools || els.inventoryQuickTools.dataset.mounted === "true") return;
  const source = document.querySelector("[data-inventory-tools]");
  if (!source) return;
  const row = document.createElement("div");
  row.className = "inventory-quick-tools__row";
  [els.inventoryStateFilter, els.arrangeColumnsButton, document.querySelector(".column-picker"), els.downloadInventory, els.createPoShortcut]
    .filter(Boolean)
    .forEach((node) => row.append(node));
  els.inventoryQuickTools.append(row);
  els.inventoryQuickTools.dataset.mounted = "true";
  source.hidden = true;
}

function currentInventoryRows() {
  const rows = buildInventoryRows();
  // Pin the most recently adjusted item to the top so it stays visible after a stock change
  if (state._pinnedAdjustCode) {
    const pinnedKey = codeKey(state._pinnedAdjustCode);
    const idx = rows.findIndex((r) => codeKey(r.code) === pinnedKey);
    if (idx > 0) {
      const [pinned] = rows.splice(idx, 1);
      rows.unshift(pinned);
    }
  }
  state.inventoryRows = rows;
  state._inventoryRowIndex = new Map(rows.map((item) => [codeKey(item.code), item]));
  return rows;
}

function latestExcelAddDate() {
  return [...state.excelItems.values()]
    .map((item) => cleanCell(item.addDate))
    .filter(Boolean)
    .sort(compareDateValue)
    .at(-1) || "";
}

function currentOrderRows() {
  const today = new Date().toLocaleDateString("en-US", { weekday: "long" }).toLowerCase();
  const vendorFilter = getOrderVendorFilter ? getOrderVendorFilter() : "Active";
  return currentInventoryRows()
    .filter((item) => {
      if (item.isOrderable === false) return false;
      if (item.recommendedOrder <= 0 && !item.qtyNeeded) return false;
      if (state.vendorRules.length && vendorFilter !== "") {
        const vendorName = (item.vendor || "").toUpperCase();
        const rule = state.vendorRules.find((r) => r.vendor?.toUpperCase() === vendorName);
        if (vendorFilter === "Active" && rule && rule.status !== "Active") return false;
        if (vendorFilter === "Disabled" && (!rule || rule.status !== "Disabled")) return false;
      }
      return true;
    })
    .sort((a, b) => (b.recommendedOrder || b.qtyNeeded || 0) - (a.recommendedOrder || a.qtyNeeded || 0));
}

function persistCountSessions() {
  localStorage.setItem("posDashboardCountSessions:v1", JSON.stringify(state.countSessions));
  localStorage.setItem("posDashboardActiveCountSession:v1", JSON.stringify(state.activeCountSession));
}

let _persistCountTimer = null;
function persistActiveCountSession() {
  // Defer localStorage write off the critical path — batches rapid scans
  clearTimeout(_persistCountTimer);
  _persistCountTimer = setTimeout(() => {
    localStorage.setItem("posDashboardActiveCountSession:v1", JSON.stringify(state.activeCountSession));
  }, 300);
}

function allCountCandidateRows() {
  if (state._countCandidateCache && state._countCandidateStamp === state._dataCacheStamp) {
    return state._countCandidateCache;
  }
  state._countCandidateCache = buildInventoryRows({ ignoreQuery: true, ignoreFilters: true, ignoreStateFilter: true });
  state._countCandidateStamp = state._dataCacheStamp;
  return state._countCandidateCache;
}

function filteredCountCandidateRows(session = state.activeCountSession) {
  if (!session) return allCountCandidateRows();
  const cacheKey = (session.id || "") + "|" + state._dataCacheStamp;
  if (state._filteredCountCache && state._filteredCountCacheKey === cacheKey) {
    return state._filteredCountCache;
  }
  const rows = allCountCandidateRows().filter((item) => {
    const vendorFilter = (session.vendor || session.department || "").trim().toUpperCase();
    if (vendorFilter && (item.vendor || "").trim().toUpperCase() !== vendorFilter) return false;
    const catFilter = (session.category || "").trim().toUpperCase();
    if (catFilter && (item.category || "").trim().toUpperCase() !== catFilter) return false;
    const statusFilter = (session.status || "").trim().toLowerCase();
    if (statusFilter && (item.state || "").trim().toLowerCase() !== statusFilter) return false;
    return true;
  });
  state._filteredCountCache = rows;
  state._filteredCountCacheKey = cacheKey;
  return rows;
}

function countSessionLabel(session) {
  if (!session) return "Physical count";
  const vendorLabel = session.vendor || session.department || "All vendors";
  const parts = [session.date || "No date", vendorLabel, session.category || "All categories"];
  return parts.join(" · ");
}

function openCountSetupModal() {
  els.countDateInput.value = new Date().toISOString().slice(0, 10);
  populateCountSetupOptions();
  // Always reset to "All" for a fresh count
  els.countVendorInput.value = "";
  els.countCategoryInput.value = "";
  const statusEl = document.querySelector("#countStatusInput");
  if (statusEl) statusEl.value = "";
  els.countSetupModal.hidden = false;
  els.countDateInput.focus();
}

function closeCountSetupModal() {
  els.countSetupModal.hidden = true;
}

function populateCountSetupOptions() {
  // Pull from ALL inventory rows — no filters applied
  const allRows = [...state.latestInventory.values()];
  const excelRows = [...state.excelItems.values()];
  const combined = allRows.length ? allRows : excelRows;
  fillSelect(els.countVendorInput, unique(combined.map((r) => r.vendor).filter(Boolean)));
  fillSelect(els.countCategoryInput, unique(combined.map((r) => r.category).filter(Boolean)));
  const statusEl = document.querySelector("#countStatusInput");
  if (statusEl) {
    // Collect states from data, but always include the known set
    const knownStates = ["Active", "Disabled", "Discontinued", "Force Order"];
    const dataStates = unique(combined.map((r) => r.state).filter(Boolean));
    const allStates = unique([...knownStates, ...dataStates]);
    fillSelect(statusEl, allStates);
  }
}

function startCountSessionFromModal() {
  const statusEl = document.querySelector("#countStatusInput");
  const session = {
    id: `count-${Date.now()}`,
    date: els.countDateInput.value || new Date().toISOString().slice(0, 10),
    vendor: els.countVendorInput.value || "",
    category: els.countCategoryInput.value || "",
    status: statusEl ? (statusEl.value || "") : "",
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    entries: [],
  };
  state.activeCountSession = session;
  state.countQtyBuffer = "0";
  state.selectedCountItemCode = "";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  persistActiveCountSession();
  closeCountSetupModal();
  buildCountSearchIndex();
  renderCountsWorkspace();
  showToast(`Started count: ${countSessionLabel(session)}`);
}

function closeActiveCountSession() {
  if (!state.activeCountSession) return;
  state.activeCountSession = null;
  state.selectedCountItemCode = "";
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  persistCountSessions();
  renderCountsWorkspace();
}

function saveCountSession() {
  if (!state.activeCountSession) return;
  const session = {
    ...state.activeCountSession,
    savedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
  };
  state.countSessions = [session, ...state.countSessions.filter((item) => item.id !== session.id)];

  // Apply physical counts: update stock values in latestInventory to reflect counted quantities
  const latestByCode = new Map();
  (session.entries || []).forEach((entry) => {
    latestByCode.set(codeKey(entry.code), entry);
  });
  let updatedCount = 0;
  latestByCode.forEach((entry, key) => {
    const existing = state.latestInventory.get(key);
    if (existing) {
      existing.stock = Number(entry.countedQty || 0);
      state.latestInventory.set(key, existing);
      updatedCount++;
    } else {
      // Item may be keyed differently — try to find it
      state.latestInventory.forEach((item, k) => {
        if (codeKey(item.code) === key) {
          item.stock = Number(entry.countedQty || 0);
        }
      });
    }
  });

  state.activeCountSession = null;
  state.selectedCountItemCode = "";
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  bumpDataStamp();
  persistCountSessions();
  renderCountsWorkspace();
  // Defer the expensive full render until after modal closes
  setTimeout(() => render(), 50);
  showToast(`Count saved · ${updatedCount} item stock quantities updated`, 3200, "success");
}

function deleteCountSession() {
  if (!state.activeCountSession) return;
  const label = countSessionLabel(state.activeCountSession);
  state.activeCountSession = null;
  state.selectedCountItemCode = "";
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  persistCountSessions();
  renderCountsWorkspace();
  showToast(`Deleted unsaved count: ${label}`, 3200, "warning");
}

function handleCountKey(key) {
  if (key === "clear") {
    state.countQtyBuffer = "0";
  } else if (key === "back") {
    state.countQtyBuffer = state.countQtyBuffer.length > 1 ? state.countQtyBuffer.slice(0, -1) : "0";
  } else if (key === "done") {
    applyCountEntry();
    return;
  } else if (key === ".") {
    if (!state.countQtyBuffer.includes(".")) state.countQtyBuffer += ".";
  } else {
    state.countQtyBuffer = state.countQtyBuffer === "0" ? key : `${state.countQtyBuffer}${key}`;
  }
  renderCountQuantity();
}

function renderCountQuantity() {
  if (els.countQuantityDisplay) els.countQuantityDisplay.textContent = state.countQtyBuffer || "0";
}

function findCountMatch(query) {
  const raw = cleanCell(query).trim();
  if (!raw) return null;
  const needle = raw.toLowerCase();
  const normalizedNeedle = codeKey(raw); // strips leading zeros for numeric matching

  // First: try exact code match (barcode scan) — check across ALL inventory rows, not just filtered
  // This ensures scans work even with strict vendor/category/status filters
  const allRows = allCountCandidateRows();
  const exactMatch = allRows.find((item) =>
    codeKey(item.code) === normalizedNeedle ||
    codeKey(item.plu) === normalizedNeedle ||
    codeKey(item.itemNumber) === normalizedNeedle ||
    item.code.toLowerCase() === needle ||
    (item.plu && item.plu.toLowerCase() === needle) ||
    (item.itemNumber && item.itemNumber.toLowerCase() === needle)
  );
  if (exactMatch) return exactMatch;

  // Second: try filtered pool with partial text match
  const filtered = filteredCountCandidateRows();
  return filtered.find((item) =>
    [item.code, item.product, item.plu, item.itemNumber, item.vendor, item.category]
      .some((value) => String(value || "").toLowerCase().includes(needle))
  );
}

function currentSelectedCountItem() {
  if (!state.selectedCountItemCode) return null;
  return filteredCountCandidateRows().find((item) => codeKey(item.code) === codeKey(state.selectedCountItemCode)) || null;
}

function renderSelectedCountItem() {
  const item = currentSelectedCountItem();
  if (!els.countSelectedItem) return;
  if (!item) {
    els.countSelectedItem.classList.remove("is-ready");
    els.countSelectedItem.innerHTML = `<p class="muted">Scan or search an item to begin counting.</p>`;
    return;
  }
  els.countSelectedItem.classList.add("is-ready");
  const previousEntry = state.activeCountSession?.entries?.filter((entry) => codeKey(entry.code) === codeKey(item.code)).at(-1);
  els.countSelectedItem.innerHTML = `
    <div class="count-item-card">
      <strong>${escapeHtml(item.product)}</strong>
      <div class="count-item-card__meta">
        <span><b>Code</b> ${escapeHtml(item.code || "-")}</span>
        <span><b>PLU</b> ${escapeHtml(item.plu || "-")}</span>
        <span><b>Vendor</b> ${escapeHtml(item.vendor || "-")}</span>
        <span><b>Category</b> ${escapeHtml(item.category || "-")}</span>
      </div>
      <small>${previousEntry ? `Last counted: ${number.format(previousEntry.countedQty || 0)} (${escapeHtml(previousEntry.mode)})` : "First entry for this item will set the counted quantity."}</small>
      <small><b>Next step:</b> enter the physical qty on the keypad, then press Enter or Done.</small>
    </div>`;
}

function handleCountLookup() {
  if (!state.activeCountSession) {
    showToast("Start a physical count first.", 3000, "warning");
    return;
  }
  const query = (els.countSearchInput.value || "").trim();
  if (!query) return;
  const match = findCountMatch(query);
  if (!match) {
    state.selectedCountItemCode = "";
    renderSelectedCountItem();
    // Immediately flash red — don't wait for qty entry
    if (els.countSearchInput) {
      els.countSearchInput.classList.add("count-search-error");
      setTimeout(() => els.countSearchInput && els.countSearchInput.classList.remove("count-search-error"), 1200);
    }
    showToast("Item not found in this count scope.", 3400, "warning");
    return;
  }
  const inScope = filteredCountCandidateRows().some((item) => codeKey(item.code) === codeKey(match.code));
  if (!inScope) {
    state.selectedCountItemCode = "";
    state.countStage = "search";
    state.countQtyBuffer = "0";
    renderSelectedCountItem();
    renderCountQuantity();
    if (els.countSearchInput) {
      els.countSearchInput.classList.add("count-search-error");
      setTimeout(() => els.countSearchInput && els.countSearchInput.classList.remove("count-search-error"), 1200);
    }
    focusCountSearch();
    showToast("Item is outside the selected vendor/category scope.", 3400, "warning");
    return;
  }
  state.selectedCountItemCode = match.code;
  state.countStage = "qty";
  state.countQtyBuffer = "0";
  hideCountDropdown();
  els.countSearchInput?.blur();
  renderSelectedCountItem();
  renderCountQuantity();
}

function clearCountLookup() {
  if (els.countSearchInput) els.countSearchInput.value = "";
  hideCountDropdown();
  state.selectedCountItemCode = "";
  state.countQtyBuffer = "0";
  state.countStage = "search";
  renderCountQuantity();
  renderSelectedCountItem();
  focusCountSearch();
}

function openDuplicateCountModal(item, qty, existing) {
  state.pendingDuplicateCount = { item, qty, existing };
  if (els.countDuplicateMessage) {
    els.countDuplicateMessage.textContent = `${item.product} was already counted as ${number.format(existing?.countedQty || 0)}. Add this new quantity or reset it?`;
  }
  els.countDuplicateModal.hidden = false;
}

function closeDuplicateCountModal() {
  state.pendingDuplicateCount = null;
  els.countDuplicateModal.hidden = true;
  state.countStage = "qty";
}

function commitCountEntry(item, qty, mode) {
  const session = {
    ...state.activeCountSession,
    updatedAt: new Date().toISOString(),
    entries: [...(state.activeCountSession.entries || [])],
  };
  const existing = session.entries.filter((entry) => codeKey(entry.code) === codeKey(item.code)).at(-1);
  const countedQty = mode === "add"
    ? Math.max(0, Number(existing?.countedQty || 0) + qty)
    : qty;
  session.entries.push({
    code: item.code,
    product: item.product,
    vendor: item.vendor || "",
    category: item.category || "",
    originalQty: item.stock || 0,
    inputQty: qty,
    countedQty,
    mode,
    unitCost: item.unitCost || 0,
    recordedAt: new Date().toISOString(),
  });
  state.activeCountSession = session;
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.selectedCountItemCode = "";
  persistActiveCountSession();
  // Fast path: prepend new row, update summary, reset UI — no full workspace rebuild
  renderCountEntryRows(true);
  renderSelectedCountItem();
  renderCountQuantity();
  updateCountSummaryStrip();
  if (els.countSearchInput) els.countSearchInput.value = "";
  hideCountDropdown();
  focusCountSearch();
  // Toast deferred so it doesn't block the next scan
  requestAnimationFrame(() => showToast(
    mode === "add"
      ? `Added ${number.format(qty)} to ${item.code}`
      : mode === "reset"
        ? `Reset ${item.code} to ${number.format(qty)}`
        : `Set ${item.code} to ${number.format(qty)}`,
    2000,
    "success",
  ));
}

function resolveDuplicateCount(mode) {
  const pending = state.pendingDuplicateCount;
  if (!pending) return;
  els.countDuplicateModal.hidden = true;
  state.pendingDuplicateCount = null;
  commitCountEntry(pending.item, pending.qty, mode === "add" ? "add" : "reset");
}

function applyCountEntry() {
  if (!state.activeCountSession) {
    showToast("Start a physical count first.", 3000, "warning");
    return;
  }
  const item = currentSelectedCountItem();
  if (!item) {
    showToast("Search and select an item first.", 3000, "warning");
    return;
  }
  const qty = Math.max(0, Number(state.countQtyBuffer || "0"));
  const existing = state.activeCountSession.entries?.filter((entry) => codeKey(entry.code) === codeKey(item.code)).at(-1);
  if (existing) {
    openDuplicateCountModal(item, qty, existing);
    return;
  }
  commitCountEntry(item, qty, "set");
}

function focusCountSearch() {
  if (!els.countSearchInput) return;
  setTimeout(() => {
    if (!els.countSearchInput) return;
    els.countSearchInput.focus();
    els.countSearchInput.select?.();
  }, 0);
}

function hideCountDropdown() {
  const dd = document.querySelector("#countSearchDropdown");
  if (dd) dd.hidden = true;
}

function buildCountSearchIndex() {
  // Force-refresh the candidate cache on session start
  state._countCandidateCache = null;
  state._filteredCountCache = null;
  const allRows = allCountCandidateRows();
  state._countSearchIndex = allRows.map((item) => ({
    item,
    haystack: [item.code, item.product, item.plu, item.itemNumber, item.vendor, item.category]
      .map((v) => String(v || "").toLowerCase()).join("|"),
    codeKey: codeKey(item.code),
  }));
  // Pre-compute filtered scope codes
  state._countFilteredCodes = new Set(
    filteredCountCandidateRows().map((r) => codeKey(r.code))
  );
  state._countIndexStamp = state._dataCacheStamp;
}

function renderCountDropdown(query) {
  const dd = document.querySelector("#countSearchDropdown");
  if (!dd) return;
  // Only show dropdown in search stage
  if (state.countStage && state.countStage !== "search") { dd.hidden = true; return; }
  const raw = cleanCell(query).trim();
  if (!raw || raw.length < 2) { dd.hidden = true; return; }
  const needle = raw.toLowerCase();

  // Use pre-built index for speed
  const index = state._countSearchIndex || [];
  const inScopeCodes = state._countFilteredCodes || new Set();
  const matches = [];
  for (const entry of index) {
    if (entry.haystack.includes(needle)) {
      matches.push(entry);
      if (matches.length >= 18) break;
    }
  }

  if (!matches.length) { dd.hidden = true; return; }

  const inScopeMatches = matches.filter((e) => inScopeCodes.has(e.codeKey));
  const outScopeMatches = matches.filter((e) => !inScopeCodes.has(e.codeKey));
  const session = state.activeCountSession;
  const scopeLabel = [session?.vendor, session?.category, session?.status].filter(Boolean).join(" · ") || "All items";
  const counted = new Set((session?.entries || []).map((e) => codeKey(e.code)));

  let html = "";
  if (inScopeMatches.length) {
    html += `<div class="count-dd-group-label">✓ In scope — ${escapeHtml(scopeLabel)}</div>`;
    html += inScopeMatches.map(({ item }) => {
      const alreadyCounted = counted.has(codeKey(item.code));
      return `<div class="count-dd-item${alreadyCounted ? " count-dd-counted" : ""}" data-code="${escapeHtml(item.code)}">
        <span class="count-dd-name">${escapeHtml(item.product)}</span>
        <span class="count-dd-meta">${escapeHtml(item.code)}${item.vendor ? ` · ${escapeHtml(item.vendor)}` : ""}${alreadyCounted ? " · <b>Counted</b>" : ""}</span>
      </div>`;
    }).join("");
  }
  if (outScopeMatches.length) {
    html += `<div class="count-dd-group-label count-dd-out-label">✗ Outside scope</div>`;
    html += outScopeMatches.map(({ item }) => `
      <div class="count-dd-item count-dd-out" title="Not in this session's scope">
        <span class="count-dd-name">${escapeHtml(item.product)}</span>
        <span class="count-dd-meta">${escapeHtml(item.code)}${item.vendor ? ` · ${escapeHtml(item.vendor)}` : ""}</span>
      </div>`).join("");
  }

  dd.innerHTML = html;
  dd.hidden = false;

  dd.querySelectorAll(".count-dd-item:not(.count-dd-out)").forEach((el) => {
    el.addEventListener("mousedown", (e) => {
      e.preventDefault();
      selectCountDropdownItem(el.dataset.code);
    });
  });
}

function selectCountDropdownItem(code) {
  hideCountDropdown();
  if (!state.activeCountSession) return;
  // Find item in filtered pool
  const item = filteredCountCandidateRows().find((r) => codeKey(r.code) === codeKey(code));
  if (!item) { showToast("Item not in session scope.", 2800, "warning"); return; }
  if (els.countSearchInput) els.countSearchInput.value = item.product;
  state.selectedCountItemCode = item.code;
  state.countStage = "qty";
  state.countQtyBuffer = "0";
  renderSelectedCountItem();
  renderCountQuantity();
}

function findCountSessionById(sessionId) {
  if (!sessionId) return null;
  if (state.activeCountSession?.id === sessionId) return state.activeCountSession;
  return state.countSessions.find((session) => session.id === sessionId) || null;
}

function currentCountSessionCandidates(session) {
  return filteredCountCandidateRows(session);
}

function openCountReport(sessionId = state.activeCountSession?.id, mode = state.countReportMode || "input") {
  const session = findCountSessionById(sessionId);
  if (!session) {
    showToast("No saved count report found yet.", 3000, "warning");
    return;
  }
  state.countReportMode = mode;
  state.countReportOpenId = sessionId;
  if (els.countReportTitle) {
    els.countReportTitle.textContent = mode === "comparison"
      ? `${countSessionLabel(session)} · Comparison report`
      : `${countSessionLabel(session)} · Input log`;
  }
  if (els.countReportHead) {
    els.countReportHead.innerHTML = mode === "comparison"
      ? `<tr><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th>Qty before</th><th>Qty after</th><th>Qty diff</th><th>Cost diff</th><th>Status</th></tr>`
      : `<tr><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th>Qty before</th><th>Qty after</th><th>Variance</th><th>Mode</th><th>Date/Time</th></tr>`;
  }
  if (els.countReportMeta) {
    const vendorLabel = session.vendor || session.department || "All vendors";
    const totalCandidates = currentCountSessionCandidates(session).length;
    els.countReportMeta.innerHTML = `
      <span><b>Date</b> ${escapeHtml(session.date || "-")}</span>
      <span><b>Vendor</b> ${escapeHtml(vendorLabel)}</span>
      <span><b>Category</b> ${escapeHtml(session.category || "All")}</span>
      <span><b>Entries</b> ${number.format((session.entries || []).length)}</span>
      <span><b>Items in scope</b> ${number.format(totalCandidates)}</span>
      <span><b>Started</b> ${escapeHtml(new Date(session.startedAt).toLocaleString())}</span>`;
  }
  renderCountReportRows(session, mode);
  els.countReportModal.hidden = false;
}

function closeCountReport() {
  els.countReportModal.hidden = true;
}

function exportCountReportPdf() {
  const sessionId = state.activeCountSession?.id || state.countSessions[0]?.id;
  const session = findCountSessionById(sessionId);
  if (!session) { showToast("No count report to export.", 3000, "warning"); return; }
  const mode = state.countReportMode || "input";
  const dateStr = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const vendorLabel = session.vendor || session.department || "All vendors";
  const entries = session.entries || [];

  let tableHtml = "";
  if (mode === "comparison") {
    const allItems = currentCountSessionCandidates(session);
    const latestByCode = new Map();
    entries.forEach((entry) => latestByCode.set(codeKey(entry.code), entry));
    tableHtml = `<table>
      <thead><tr><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th class="num">Qty Before</th><th class="num">Qty After</th><th class="num">Qty Diff</th><th class="num">Cost Diff</th><th>Status</th></tr></thead>
      <tbody>${allItems.map((item) => {
        const entry = latestByCode.get(codeKey(item.code));
        const orig = Number(item.stock || 0);
        const final = entry ? Number(entry.countedQty || 0) : null;
        const diff = entry ? final - orig : null;
        const costDiff = entry ? diff * Number(item.unitCost || 0) : null;
        const cls = diff == null ? "" : diff > 0 ? "var-up" : diff < 0 ? "var-down" : "";
        return `<tr class="${cls}"><td>${escapeHtml(item.code)}</td><td>${escapeHtml(item.product)}</td><td>${escapeHtml(item.vendor || "-")}</td><td>${escapeHtml(item.category || "-")}</td>
          <td class="num">${number.format(orig)}</td>
          <td class="num">${entry ? number.format(final) : "NULL"}</td>
          <td class="num">${entry ? (diff > 0 ? `+${number.format(diff)}` : number.format(diff)) : "NULL"}</td>
          <td class="num">${entry ? currency.format(costDiff) : "NULL"}</td>
          <td>${entry ? "Scanned" : "Not scanned"}</td></tr>`;
      }).join("")}</tbody></table>`;
  } else {
    tableHtml = `<table>
      <thead><tr><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th class="num">Qty Before</th><th class="num">Counted</th><th class="num">Variance</th><th>Mode</th><th>Time</th></tr></thead>
      <tbody>${[...entries].reverse().map((entry) => {
        const variance = Number(entry.countedQty || 0) - Number(entry.originalQty || 0);
        const cls = variance > 0 ? "var-up" : variance < 0 ? "var-down" : "";
        return `<tr class="${cls}"><td>${escapeHtml(entry.code)}</td><td>${escapeHtml(entry.product)}</td><td>${escapeHtml(entry.vendor || "-")}</td><td>${escapeHtml(entry.category || "-")}</td>
          <td class="num">${number.format(entry.originalQty || 0)}</td>
          <td class="num">${number.format(entry.countedQty || 0)}</td>
          <td class="num">${variance > 0 ? `+${number.format(variance)}` : number.format(variance)}</td>
          <td>${escapeHtml(entry.mode || "set")}</td>
          <td>${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</td></tr>`;
      }).join("")}</tbody></table>`;
  }

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <title>Physical Count Report – ${escapeHtml(session.date || "-")}</title>
  <style>
    body { font-family: Arial, sans-serif; font-size: 11px; color: #1c2320; margin: 0; padding: 24px; }
    h1 { font-size: 18px; margin: 0 0 4px; }
    .meta { color: #66716d; margin-bottom: 20px; font-size: 11px; display: flex; gap: 20px; flex-wrap: wrap; }
    .meta span { background: #f0f4f2; padding: 3px 8px; border-radius: 4px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 8px; }
    th { background: #eef7f0; text-align: left; padding: 5px 6px; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #dce3df; }
    td { padding: 4px 6px; border-bottom: 1px solid #eee; }
    .num { text-align: right; }
    .var-up td { color: #16835b; } .var-down td { color: #c0392b; }
    @media print { body { padding: 8px; } }
  </style></head><body>
  <h1>Physical Count Report</h1>
  <div class="meta">
    <span><b>Date</b> ${escapeHtml(session.date || "-")}</span>
    <span><b>Vendor</b> ${escapeHtml(vendorLabel)}</span>
    <span><b>Category</b> ${escapeHtml(session.category || "All")}</span>
    <span><b>Entries</b> ${number.format(entries.length)}</span>
    <span><b>Mode</b> ${mode === "comparison" ? "Comparison" : "Input Log"}</span>
    <span><b>Generated</b> ${dateStr}</span>
  </div>
  ${tableHtml}
  </body></html>`;

  const win = window.open("", "_blank", "width=1000,height=750");
  if (!win) { showToast("Pop-up blocked — allow pop-ups to export PDF.", 3500, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 500);
  showToast("Print dialog opened — use 'Save as PDF'.", 3200, "success");
}

async function exportCountReportExcel() {
  const sessionId = state.activeCountSession?.id || state.countSessions[0]?.id;
  const session = findCountSessionById(sessionId);
  if (!session) { showToast("No count report to export.", 3000, "warning"); return; }
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available.", 3000, "warning"); return; }

  const vendorLabel = session.vendor || session.department || "All vendors";
  const entries = session.entries || [];
  const wb = xlsx.utils.book_new();

  // Input log sheet
  const inputData = [
    ["Physical Count — Input Log", "", "", `Date: ${session.date || "-"}`, `Vendor: ${vendorLabel}`, `Category: ${session.category || "All"}`],
    [],
    ["Code", "Item", "Vendor", "Category", "Qty Before", "Counted", "Variance", "Mode", "Time"],
    ...[...entries].reverse().map((entry) => {
      const variance = Number(entry.countedQty || 0) - Number(entry.originalQty || 0);
      return [entry.code, entry.product, entry.vendor || "", entry.category || "", entry.originalQty || 0, entry.countedQty || 0, variance, entry.mode || "set", new Date(entry.recordedAt).toLocaleString()];
    }),
  ];
  const wsInput = xlsx.utils.aoa_to_sheet(inputData);
  wsInput["!cols"] = [12, 32, 14, 14, 10, 10, 10, 8, 20].map((w) => ({ wch: w }));
  xlsx.utils.book_append_sheet(wb, wsInput, "Input Log");

  // Comparison sheet
  const allItems = currentCountSessionCandidates(session);
  const latestByCode = new Map();
  entries.forEach((entry) => latestByCode.set(codeKey(entry.code), entry));
  const compData = [
    ["Physical Count — Comparison", "", "", `Date: ${session.date || "-"}`, `Vendor: ${vendorLabel}`, `Category: ${session.category || "All"}`],
    [],
    ["Code", "Item", "Vendor", "Category", "Qty Before", "Qty After", "Qty Diff", "Cost Diff", "Status"],
    ...allItems.map((item) => {
      const entry = latestByCode.get(codeKey(item.code));
      const orig = Number(item.stock || 0);
      const final = entry ? Number(entry.countedQty || 0) : null;
      const diff = entry != null ? final - orig : null;
      const costDiff = entry != null ? diff * Number(item.unitCost || 0) : null;
      return [item.code, item.product, item.vendor || "", item.category || "", orig, final ?? "NULL", diff ?? "NULL", costDiff ?? "NULL", entry ? "Scanned" : "Not scanned"];
    }),
  ];
  const wsComp = xlsx.utils.aoa_to_sheet(compData);
  wsComp["!cols"] = [12, 32, 14, 14, 10, 10, 10, 10, 12].map((w) => ({ wch: w }));
  xlsx.utils.book_append_sheet(wb, wsComp, "Comparison");

  xlsx.writeFile(wb, `PhysicalCount_${session.date || "report"}.xlsx`);
  showToast(`Count report exported — ${entries.length} entries, ${allItems.length} items in scope`, 3200, "success");
}

function renderCountReportRows(session, mode = state.countReportMode || "input") {
  if (!els.countReportBody) return;
  const entries = session?.entries || [];
  if (mode === "comparison") {
    const allItems = currentCountSessionCandidates(session);
    if (!allItems.length) {
      els.countReportBody.innerHTML = `<tr><td colspan="9" class="empty-cell">No items matched this count criteria.</td></tr>`;
      return;
    }
    const latestByCode = new Map();
    entries.forEach((entry) => latestByCode.set(codeKey(entry.code), entry));
    els.countReportBody.innerHTML = allItems
      .map((item) => {
        const entry = latestByCode.get(codeKey(item.code));
        const originalQty = Number(item.stock || 0);
        const finalQty = entry ? Number(entry.countedQty || 0) : null;
        const qtyDiff = entry ? finalQty - originalQty : null;
        const costDiff = entry ? qtyDiff * Number(item.unitCost || 0) : null;
        const qtyClass = qtyDiff == null ? "" : qtyDiff > 0 ? "variance-up" : qtyDiff < 0 ? "variance-down" : "variance-flat";
        const costClass = costDiff == null ? "" : costDiff > 0 ? "variance-up" : costDiff < 0 ? "variance-down" : "variance-flat";
        return `
          <tr>
            <td>${escapeHtml(item.code || "-")}</td>
            <td>${escapeHtml(item.product || "-")}</td>
            <td>${escapeHtml(item.vendor || "-")}</td>
            <td>${escapeHtml(item.category || "-")}</td>
            <td class="num">${number.format(originalQty)}</td>
            <td class="num">${entry ? number.format(finalQty) : `<span class="muted">NULL</span>`}</td>
            <td class="num ${qtyClass}">${entry ? (qtyDiff > 0 ? `+${number.format(qtyDiff)}` : number.format(qtyDiff)) : `<span class="muted">NULL</span>`}</td>
            <td class="${costClass}">${entry ? currency.format(costDiff) : `<span class="muted">NULL</span>`}</td>
            <td>${entry ? escapeHtml(`Scanned ${new Date(entry.recordedAt).toLocaleString()}`) : `<span class="muted">Not scanned</span>`}</td>
          </tr>`;
      })
      .join("");
    return;
  }
  if (!entries.length) {
    els.countReportBody.innerHTML = `<tr><td colspan="9" class="empty-cell">No items counted yet.</td></tr>`;
    return;
  }
  els.countReportBody.innerHTML = entries
    .slice()
    .reverse()
    .map((entry) => {
      const variance = Number(entry.countedQty || 0) - Number(entry.originalQty || 0);
      const varianceClass = variance > 0 ? "variance-up" : variance < 0 ? "variance-down" : "variance-flat";
      const varianceLabel = variance > 0 ? `+${number.format(variance)}` : number.format(variance);
      return `
        <tr>
          <td>${escapeHtml(entry.code || "-")}</td>
          <td>${escapeHtml(entry.product || "-")}</td>
          <td>${escapeHtml(entry.vendor || "-")}</td>
          <td>${escapeHtml(entry.category || "-")}</td>
          <td class="num">${number.format(entry.originalQty || 0)}</td>
          <td class="num">${number.format(entry.countedQty || 0)}</td>
          <td class="num ${varianceClass}">${varianceLabel}</td>
          <td>${escapeHtml(entry.mode || "set")}</td>
          <td>${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</td>
        </tr>`;
    })
    .join("");
}

function openSessionHistoryModal() {
  // Populate vendor filter
  const vendors = [...new Set(state.countSessions.map((s) => s.vendor || s.department || "").filter(Boolean))];
  if (els.sessionHistoryVendorFilter) {
    const cur = els.sessionHistoryVendorFilter.value;
    els.sessionHistoryVendorFilter.innerHTML = `<option value="">All vendors</option>` +
      vendors.map((v) => `<option value="${escapeHtml(v)}" ${v === cur ? "selected" : ""}>${escapeHtml(v)}</option>`).join("");
  }
  document.querySelector("#sessionHistoryModal").hidden = false;
  renderCountSessionRows();
}

function sessionMatchesPeriodFilter(session, period) {
  if (!period) return true;
  const now = new Date();
  const sessionDate = new Date(session.updatedAt || session.startedAt);
  const y = now.getFullYear(), m = now.getMonth(), d = now.getDate();
  if (period === "day") {
    return sessionDate.getFullYear() === y && sessionDate.getMonth() === m && sessionDate.getDate() === d;
  }
  if (period === "week") {
    const startOfWeek = new Date(y, m, d - now.getDay());
    return sessionDate >= startOfWeek;
  }
  if (period === "month") return sessionDate.getFullYear() === y && sessionDate.getMonth() === m;
  if (period === "quarter") {
    const q = Math.floor(m / 3);
    const startOfQ = new Date(y, q * 3, 1);
    return sessionDate >= startOfQ;
  }
  if (period === "year") return sessionDate.getFullYear() === y;
  return true;
}

function renderCountSessionRows() {
  if (!els.countSessionBody) return;
  const vendorFilter = els.sessionHistoryVendorFilter?.value || "";
  const periodFilter = els.sessionHistoryPeriodFilter?.value || "";

  const filtered = state.countSessions
    .slice()
    .sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")))
    .filter((s) => {
      if (vendorFilter && (s.vendor || s.department || "") !== vendorFilter) return false;
      if (!sessionMatchesPeriodFilter(s, periodFilter)) return false;
      return true;
    });

  if (!filtered.length) {
    els.countSessionBody.innerHTML = `<tr><td colspan="11" class="empty-cell">${state.countSessions.length ? "No sessions match filters." : "No physical count sessions saved yet."}</td></tr>`;
    return;
  }
  els.countSessionBody.innerHTML = filtered
    .map((session) => `
      <tr>
        <td>${escapeHtml(session.date || "-")}</td>
        <td>${escapeHtml(session.vendor || session.department || "All")}</td>
        <td>${escapeHtml(session.category || "All")}</td>
        <td>${escapeHtml(session.status || "All")}</td>
        <td>${escapeHtml(new Date(session.startedAt).toLocaleString())}</td>
        <td class="num">${number.format((session.entries || []).length)}</td>
        <td>${escapeHtml(new Date(session.updatedAt || session.startedAt).toLocaleString())}</td>
        <td><button type="button" class="secondary-button count-inline-report-button" data-count-report="${escapeHtml(session.id)}">Input/Compare</button></td>
        <td>${session.submittedAt ? `<button type="button" class="secondary-button final-report-btn" data-final-report="${escapeHtml(session.id)}">Final Report</button>` : `<span class="muted">Not submitted</span>`}</td>
        <td>${session.preCountSnapshot ? (session.restoredAt ? `<span class="muted">Restored</span>` : `<button type="button" class="restore-count-btn" data-restore-session="${escapeHtml(session.id)}">Restore</button>`) : ""}</td>
        <td><button type="button" class="delete-session-btn" data-delete-session="${escapeHtml(session.id)}">Delete</button></td>
      </tr>`)
    .join("");

  els.countSessionBody.querySelectorAll("[data-count-report]").forEach((btn) => {
    btn.addEventListener("click", (e) => { e.stopPropagation(); openCountReport(btn.dataset.countReport); });
  });
  els.countSessionBody.querySelectorAll("[data-final-report]").forEach((btn) => {
    btn.addEventListener("click", (e) => { e.stopPropagation(); openFinalCountReport(btn.dataset.finalReport); });
  });
  els.countSessionBody.querySelectorAll("[data-delete-session]").forEach((btn) => {
    btn.addEventListener("click", (e) => { e.stopPropagation(); openConfirmDeleteSession(btn.dataset.deleteSession); });
  });
  els.countSessionBody.querySelectorAll("[data-restore-session]").forEach((btn) => {
    btn.addEventListener("click", (e) => { e.stopPropagation(); restorePreviousCount(btn.dataset.restoreSession); });
  });
}

function renderCountEntryRows(prependOnly = false) {
  if (!els.countEntryBody) return;
  const entries = state.activeCountSession?.entries || [];
  if (!entries.length) {
    els.countEntryBody.innerHTML = `<tr><td colspan="8" class="empty-cell">No items counted yet.</td></tr>`;
    return;
  }
  // Fast path: just prepend the newest entry row
  if (prependOnly && entries.length > 0) {
    const entry = entries[entries.length - 1];
    const variance = Number(entry.countedQty || 0) - Number(entry.originalQty || 0);
    const varClass = variance > 0 ? "entry-positive" : variance < 0 ? "entry-negative" : "entry-exact";
    const varLabel = variance > 0 ? `+${number.format(variance)}` : number.format(variance);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(entry.code || "-")}</td>
      <td>${escapeHtml(entry.product || "-")}</td>
      <td>${escapeHtml(entry.vendor || "-")}</td>
      <td>${escapeHtml(entry.category || "-")}</td>
      <td class="num">${entry.mode === "add" ? `+${number.format(entry.inputQty || 0)}` : number.format(entry.inputQty || 0)}</td>
      <td class="num">${number.format(entry.countedQty || 0)}</td>
      <td>${escapeHtml(entry.mode || "set")}</td>
      <td class="num ${varClass}">${varLabel}</td>`;
    // Remove "no items" placeholder if present
    const placeholder = els.countEntryBody.querySelector("td[colspan]");
    if (placeholder) placeholder.closest("tr").remove();
    els.countEntryBody.prepend(tr);
    return;
  }
  // Full rebuild
  els.countEntryBody.innerHTML = entries
    .slice()
    .reverse()
    .map((entry) => {
      const variance = Number(entry.countedQty || 0) - Number(entry.originalQty || 0);
      const varClass = variance > 0 ? "entry-positive" : variance < 0 ? "entry-negative" : "entry-exact";
      const varLabel = variance > 0 ? `+${number.format(variance)}` : number.format(variance);
      return `
      <tr>
        <td>${escapeHtml(entry.code || "-")}</td>
        <td>${escapeHtml(entry.product || "-")}</td>
        <td>${escapeHtml(entry.vendor || "-")}</td>
        <td>${escapeHtml(entry.category || "-")}</td>
        <td class="num">${entry.mode === "add" ? `+${number.format(entry.inputQty || 0)}` : number.format(entry.inputQty || 0)}</td>
        <td class="num">${number.format(entry.countedQty || 0)}</td>
        <td>${escapeHtml(entry.mode || "set")}</td>
        <td class="num ${varClass}">${varLabel}</td>
      </tr>`;
    })
    .join("");
}

function updateCountSummaryStrip() {
  if (!state.activeCountSession || !els.countSummaryStrip) return;
  const entries = state.activeCountSession.entries || [];
  const date = state.activeCountSession.date || "-";
  const vendor = state.activeCountSession.vendor || "All vendors";
  const category = state.activeCountSession.category || "All categories";
  els.countSummaryStrip.innerHTML = `
    <span><b>${escapeHtml(date)}</b></span>
    <span><b>${escapeHtml(vendor)}</b></span>
    <span><b>${escapeHtml(category)}</b></span>
    <span><b>${number.format(entries.length)}</b> items counted</span>`;
}

function renderCountsWorkspace() {
  populateCountSetupOptions();
  const active = state.activeCountSession;
  if (els.countSessionModal) els.countSessionModal.hidden = !active;
  els.countWorkspaceEmpty.hidden = false;
  if (els.closeCountSessionButton) els.closeCountSessionButton.hidden = true;
  if (els.countReviewButton) els.countReviewButton.hidden = true;
  if (active) {
    els.activeCountTitle.textContent = countSessionLabel(active);
    const vendorLabel = active.vendor || active.department || "All vendors";
    els.activeCountMeta.innerHTML = `
      <span>${escapeHtml(active.date || "-")}</span>
      <span>${escapeHtml(vendorLabel)}</span>
      <span>${escapeHtml(active.category || "All categories")}</span>
      <span>${number.format((active.entries || []).length)} entries</span>`;
  } else if (els.activeCountMeta) {
    els.activeCountMeta.innerHTML = "";
  }
  renderCountQuantity();
  renderSelectedCountItem();
  renderCountEntryRows();
  renderCountSessionRows();
  if (active && state.countStage === "search") focusCountSearch();
}

function renderActiveTab() {
  hideHoverTooltip();
  const activeTab = document.querySelector(".tab-button.active")?.dataset.tab || "dashboard";
  if (els.inventoryQuickTools) {
    els.inventoryQuickTools.hidden = activeTab !== "inventory";
  }
  if (activeTab === "dashboard") {
    renderTrend();
    renderBars();
  } else if (activeTab === "inventory") {
    renderInventory();
  } else if (activeTab === "counts") {
    renderCountsWorkspace();
  } else if (activeTab === "reports") {
    renderAdjustLog();
  } else if (activeTab === "vendors") {
    renderVendorRules();
  } else if (activeTab === "settings") {
    renderSettings();
  } else if (activeTab === "ordering") {
    renderOrders();
    renderTable();
  } else if (activeTab === "parents") {
    renderParents();
  }
  refreshDetailDrawer();
}

function applyRoleRestrictions() {
  const userMode = isUserRole();
  // Hide admin-only elements for basic users
  const adminOnly = [
    "#downloadInventoryCsvBtn", ".download-inventory-csv", "[data-admin-only]",
    "#exportPoExcel", "#exportPoPdf", "#downloadOrder",
    "#exportAdjustPdfButton", "#exportAdjustExcelButton",
    "#clearAdjustLogButton", ".arrange-columns-btn", "#arrangeColumnsButton",
    ".column-picker", "#downloadInventory", "#createPoShortcut", "#openSessionHistoryButton"
  ];
  adminOnly.forEach((sel) => {
    document.querySelectorAll(sel).forEach((el) => { el.style.display = userMode ? "none" : ""; });
  });
  // Hide cost columns for user role
  if (userMode) {
    document.querySelectorAll("#inventory th[data-col='unitCost'], #inventoryBody td[data-col='unitCost'], #inventory th[data-col='inventoryCost'], #inventoryBody td[data-col='inventoryCost']").forEach((el) => { el.style.display = "none"; });
    // Hide metrics strip cost info
    document.querySelectorAll("#costSold, #grossProfit").forEach((el) => {
      const article = el?.closest("article");
      if (article) article.style.display = "none";
    });
    const inventorySummary = document.querySelector("#inventorySummary");
    if (inventorySummary) inventorySummary.style.display = "none";
    // Hide summary strip cost items
    document.querySelectorAll(".inventory-summary .cost-item, [data-cost-summary]").forEach((el) => { el.style.display = "none"; });
    // Only show products and inventory tabs
    document.querySelectorAll(".tab-button").forEach((btn) => {
      const tab = btn.dataset.tab;
      if (!["inventory","counts"].includes(tab)) btn.style.display = "none";
    });
  } else {
    document.querySelectorAll(".tab-button").forEach((btn) => { btn.style.display = ""; });
    const metricsZone = document.querySelector("#metricsHoverZone");
    if (metricsZone) metricsZone.style.display = "";
  }
}

function queueActiveTabRender() {
  const token = ++state._activeTabRenderToken;
  if (state._activeTabRenderHandle) {
    clearTimeout(state._activeTabRenderHandle);
  }
  state._activeTabRenderHandle = setTimeout(() => {
    if (token !== state._activeTabRenderToken) return;
    renderActiveTab();
    state._activeTabRenderHandle = 0;
  }, 0);
}

function render() {
  if (els.countSessionModal && !els.countSessionModal.hidden) return;
  const skuRows = buildSkuRows();
  state.filteredSkus = sortSkuRows(skuRows);
  state._filteredSkuIndex = new Map(state.filteredSkus.map((item) => [codeKey(item.code), item]));
  const metricRows = buildSkuRows({ ignoreStateFilter: true });
  const dates = filteredSalesDates();
  const selectedDayCount = rangeDayCount(els.startDate.value || state.dates[0], els.endDate.value || state.dates[state.dates.length - 1]);
  const totals = metricRows.reduce(
    (sum, sku) => ({
      sales: sum.sales + sku.sales,
      units: sum.units + sku.units,
      costSold: sum.costSold + sku.costSold,
      inventoryCost: sum.inventoryCost + sku.inventoryCost,
      profit: sum.profit + sku.profit,
      atRisk: sum.atRisk + (sku.status === "stockout" || sku.status === "watch" ? 1 : 0),
    }),
    { sales: 0, units: 0, costSold: 0, inventoryCost: 0, profit: 0, atRisk: 0 },
  );

  els.totalSales.textContent = currency.format(totals.sales);
  els.unitsSold.textContent = number.format(totals.units);
  els.grossProfit.textContent = currency.format(totals.profit);
  els.costSold.textContent = currency.format(totals.costSold);
  els.costTotal.textContent = `Inventory cost ${currency.format(totals.inventoryCost)}`;
  els.marginRate.textContent = `${percent(totals.profit, totals.sales)} margin`;
  els.riskCount.textContent = number.format(totals.atRisk);
  els.avgDailyUnits.textContent = `${formatVelocity(totals.units / Math.max(selectedDayCount, 1))} per day`;
  els.salesDelta.textContent = dates.length ? `${selectedDayCount} days in view` : "Load CSVs to begin.";
  els.fileCount.textContent = fileSummary();
  els.dateCoverage.textContent = coverageSummary();
  renderDatePresets();
  syncStickyHeights();
  queueActiveTabRender();
  applyRoleRestrictions();
}

function buildSkuRows(options = {}) {
  const query = options.ignoreQuery ? "" : els.searchInput.value.trim().toLowerCase();
  const start = els.startDate.value || "0000-00-00";
  const end = els.endDate.value || "9999-99-99";
  const department = options.ignoreFilters ? "" : els.departmentFilter.value;
  const category = options.ignoreFilters ? "" : els.categoryFilter.value;
  const vendor = options.ignoreFilters ? "" : els.vendorFilter.value;
  const color = options.ignoreFilters ? "" : els.colorFilter.value;
  const stateFilter = options.ignoreFilters || options.ignoreStateFilter ? "" : els.inventoryStateFilter.value;
  const leadDays = 0;
  const safetyDays = Math.max(0, toNumber(els.safetyDays.value) || 7);
  const daysOfInventory = Math.max(0, toNumber(els.daysOfInventory?.value) || 0);

  // Cache busts on date range OR any ordering parameter change
  const cacheKey = `${start}|${end}|${leadDays}|${safetyDays}|${daysOfInventory}`;
  if (!state._skuCache || state._skuCacheStamp !== state._dataCacheStamp || state._skuCacheKey !== cacheKey) {
    state._skuCache = _buildRawSkuMap(start, end, leadDays, safetyDays, daysOfInventory);
    state._skuCacheStamp = state._dataCacheStamp;
    state._skuCacheKey = cacheKey;
  }
  const allRows = state._skuCache;

  // Fast filter pass on already-aggregated rows
  return [...allRows.values()].filter((sku) => {
    if (department && sku.department !== department) return false;
    if (category && sku.category !== category) return false;
    if (vendor && sku.vendor !== vendor) return false;
    if (color && sku.color !== color) return false;
    if (stateFilter) {
      const s = (sku.state || "").toLowerCase().trim();
      if (stateFilter === "Active") {
        if (s !== "active" && s !== "force order") return false;
      } else if ((sku.state || "") !== stateFilter) return false;
    }
    if (query && !sku._haystack.includes(query.toLowerCase())) return false;
    return true;
  });
}

function _buildRawSkuMap(start, end, leadDays, safetyDays, daysOfInventory) {
  const dayCount = rangeDayCount(start, end);
  const inventoryIndex = inventoryIndexForDate(end);
  const inventoryByKey = indexRowsByCodeKey([...inventoryIndex.values()]);
  const grouped = new Map();

  state.rawSales.filter((row) => row.date >= start && row.date <= end).forEach((row) => {
    const inventory = inventoryIndex.get(row.code) || inventoryByKey.get(codeKey(row.code)) || {};
    const excel = findExcelFor(inventory.code ? inventory : row);
    // Skip discontinued/disabled items from ordering calculations
    const itemState = (excel.state || "").toLowerCase();
    const merged = {
      code: row.code || `${row.product}-${row.category}`,
      product: row.product || inventory.product || excel.product || "Unnamed item",
      department: row.department || "Unassigned",
      category: row.category || inventory.category || excel.category || "Unassigned",
      vendor: row.vendor !== "Unassigned" ? row.vendor : inventory.vendor || excel.vendor || "Unassigned",
      color: inventory.color || "",
      size: inventory.size || "",
      plu: inventory.plu || excel.plu || "",
      itemNumber: inventory.itemNumber || excel.itemNumber || "",
      addDate: excel.addDate || "",
      state: excel.state || "",
      itemState,
      caseSize: excel.caseSize || 1,
      reorderMin: 0,
      reorderMax: 0,
      stock: hasValue(inventory.stock) ? inventory.stock : excel.stock || 0,
      unitCost: pickNumber(inventory.cost, excel.cost),
      price: pickNumber(inventory.price, excel.price),
      sales: row.sales,
      units: row.units,
      costSold: row.cost || (row.units * pickNumber(inventory.cost, excel.cost)),
      profit: row.profit,
    };
    Object.assign(merged, parentPartsFor(merged));
    merged._haystack = [merged.code, merged.product, merged.department, merged.category, merged.vendor,
      merged.color, merged.size, merged.parent, merged.subType,
      merged.plu, String(merged.plu||""), merged.itemNumber, String(merged.itemNumber||""),
      merged.state, merged.addDate, merged.caseSize]
      .filter(v => v != null && v !== "").join(" ").toLowerCase().trim();

    const groupKey = codeKey(merged.code);
    const existing = grouped.get(groupKey) || { ...merged, sales: 0, units: 0, costSold: 0, profit: 0 };
    existing.product = bestLabel(existing.product, merged.product);
    existing.vendor = bestLabel(existing.vendor, merged.vendor);
    existing.category = bestLabel(existing.category, merged.category);
    existing.color = bestLabel(existing.color, merged.color);
    existing.sales += merged.sales;
    existing.units += merged.units;
    existing.costSold += merged.costSold;
    existing.profit += merged.profit;
    existing.stock = merged.stock;
    grouped.set(groupKey, existing);
  });

  return new Map([...grouped.entries()].map(([key, sku]) => {
    const velocity = sku.units / dayCount;
    const override = state.reorderOverrides[sku.code] || {};
    const isOverridden = override.min != null || override.max != null;
    const dynamic = orderingTargets({ velocity, safetyDays, daysOfInventory });
    const dynamicMaxWithDoi = dynamic.max;
    const manualMin = override.min ?? dynamic.min;
    const manualMax = override.max ?? dynamicMaxWithDoi;
    const recommendedOrder = recommendedOrderQty({
      stock: sku.stock,
      min: manualMin,
      max: manualMax,
      caseSize: sku.caseSize,
    });
    const caseOrder = roundOrderToNearestCase(recommendedOrder, sku.caseSize);
    const daysSupply = velocity > 0 ? sku.stock / velocity : Infinity;
    const margin = sku.sales > 0 ? sku.profit / sku.sales : 0;
    const inventoryCost = sku.stock * sku.unitCost;
    const full = {
      ...sku, velocity, daysSupply, recommendedOrder, caseOrder, margin,
      reorderMin: manualMin, reorderMax: manualMax, inventoryCost,
      isOverridden, dynamicMin: dynamic.min, dynamicMax: dynamicMaxWithDoi,
      ...parentPartsFor(sku),
      status: classifySku({ velocity, daysSupply, margin, recommendedOrder, stock: sku.stock, itemState: sku.itemState }),
    };
    return [key, full];
  }));
}

function classifySku({ velocity, daysSupply, margin, recommendedOrder, stock, itemState }) {
  const state = (itemState || "").toLowerCase();
  if (state === "discontinued") return "discontinued";
  if (state === "disabled") return "disabled";
  if (state === "force order") return "forceorder";
  if (velocity > 0 && daysSupply <= 3) return "stockout";
  if (recommendedOrder > 0 && velocity >= 0.15) return "watch";
  if (velocity >= 1 && margin >= 0.25 && daysSupply <= 30) return "grow";
  if (velocity < 0.05 && stock > 10) return "cut";
  return "steady";
}

function sortSkuRows(rows) {
  const mode = els.sortMode.value;
  return [...rows].sort((a, b) => {
    if (mode === "daysSupply") return finite(a.daysSupply) - finite(b.daysSupply);
    return (b[mode] || 0) - (a[mode] || 0);
  });
}

function renderTrend() {
  const ctx = els.trendChart.getContext("2d");
  const width = els.trendChart.clientWidth || 900;
  const height = Number(els.trendChart.getAttribute("height"));
  els.trendChart.width = width * window.devicePixelRatio;
  els.trendChart.height = height * window.devicePixelRatio;
  ctx.scale(window.devicePixelRatio, window.devicePixelRatio);
  ctx.clearRect(0, 0, width, height);
  drawGrid(ctx, width, height);

  const days = filteredSalesDates().map((date) => ({ date, ...dailyTotalsFor(date) }));
  if (!days.length) return drawEmpty(ctx, width, height, "Load CSVs to see daily sales history.");

  // Prior period comparison
  const comparing = els.compareToggle?.checked && days.length > 0;
  let priorDays = [];
  if (comparing) {
    const spanMs = (days.length - 1) * 86400000;
    const currentStartMs = new Date(`${days[0].date}T00:00:00`).getTime();
    const priorEndMs = currentStartMs - 86400000;
    const priorStartMs = priorEndMs - spanMs;
    priorDays = state.dates
      .filter((d) => {
        const t = new Date(`${d}T00:00:00`).getTime();
        return t >= priorStartMs && t <= priorEndMs;
      })
      .map((date) => ({ date, ...dailyTotalsFor(date) }));
  }

  const allSales = [...days, ...priorDays].map((d) => d.sales);
  const allUnits = [...days, ...priorDays].map((d) => d.units);
  const maxSales = Math.max(...allSales, 1);
  const maxUnits = Math.max(...allUnits, 1);

  if (comparing && priorDays.length) {
    drawLineScaled(ctx, priorDays, "sales", "#a8d5bf", maxSales, width, height, true);
    drawLineScaled(ctx, priorDays, "units", "#aac2e8", maxUnits, width, height, true);
  }
  drawLineScaled(ctx, days, "sales", "#16835b", maxSales, width, height, false);
  drawLineScaled(ctx, days, "units", "#3d75c9", maxUnits, width, height, false);

  ctx.fillStyle = "#66716d";
  ctx.font = "11px Arial";
  ctx.fillText(days[0].date.slice(5), 12, height - 10);
  ctx.fillText(days[days.length - 1].date.slice(5), width - 54, height - 10);

  // Comparison summary strip
  if (comparing && priorDays.length && els.comparisonSummary) {
    const curSales = days.reduce((s, d) => s + d.sales, 0);
    const priorSales = priorDays.reduce((s, d) => s + d.sales, 0);
    const curUnits = days.reduce((s, d) => s + d.units, 0);
    const priorUnits = priorDays.reduce((s, d) => s + d.units, 0);
    const salesChg = priorSales ? ((curSales - priorSales) / priorSales) * 100 : 0;
    const unitsChg = priorUnits ? ((curUnits - priorUnits) / priorUnits) * 100 : 0;
    const fmt = (v) => (v >= 0 ? `+${v.toFixed(1)}%` : `${v.toFixed(1)}%`);
    const cls = (v) => v >= 0 ? "compare-up" : "compare-down";
    els.comparisonSummary.hidden = false;
    els.comparisonSummary.innerHTML = `
      <span>vs prior ${days.length}-day period:</span>
      <b class="${cls(salesChg)}">Sales ${fmt(salesChg)}</b>
      <span>(${currency.format(curSales)} vs ${currency.format(priorSales)})</span>
      <b class="${cls(unitsChg)}">Units ${fmt(unitsChg)}</b>
      <span>(${number.format(curUnits)} vs ${number.format(priorUnits)})</span>`;
  } else if (els.comparisonSummary) {
    els.comparisonSummary.hidden = true;
  }
}

function drawLineScaled(ctx, days, key, color, maxVal, width, height, dashed) {
  if (!days.length) return;
  const left = 12, right = width - 12, top = 16, bottom = height - 28;
  ctx.strokeStyle = color;
  ctx.lineWidth = dashed ? 1.5 : 2.5;
  if (dashed) ctx.setLineDash([4, 4]); else ctx.setLineDash([]);
  ctx.beginPath();
  days.forEach((day, i) => {
    const x = days.length === 1 ? width / 2 : left + ((right - left) * i) / (days.length - 1);
    const y = bottom - ((bottom - top) * day[key]) / maxVal;
    if (i === 0) ctx.moveTo(x, y); else ctx.lineTo(x, y);
  });
  ctx.stroke();
  ctx.setLineDash([]);
}

function drawGrid(ctx, width, height) {
  ctx.strokeStyle = "#dce3df";
  ctx.lineWidth = 1;
  for (let i = 0; i < 5; i++) {
    const y = 18 + ((height - 46) / 4) * i;
    ctx.beginPath(); ctx.moveTo(10, y); ctx.lineTo(width - 10, y); ctx.stroke();
  }
}

function drawEmpty(ctx, width, height, message) {
  ctx.fillStyle = "#66716d";
  ctx.font = "16px Arial";
  ctx.textAlign = "center";
  ctx.fillText(message, width / 2, height / 2);
  ctx.textAlign = "left";
}

function renderBars() {
  const metric = els.segmentMetric.value;
  const group = els.segmentGroup?.value || "department";
  const sourceRows = currentInventoryRows();
  if (els.segmentTitle) els.segmentTitle.textContent = { department: "Departments", category: "Categories", vendor: "Vendors" }[group] || group;
  if (els.categoryPanelTitle) {
    els.categoryPanelTitle.textContent = {
      department: "Top categories by department",
      category: "Top categories overall",
      vendor: "Top categories by vendor",
    }[group] || "Top categories";
  }
  if (els.categoryPanelHint) {
    els.categoryPanelHint.textContent = group === "category"
      ? "Click a category bar to show its top 10 and bottom 10 items."
      : "Click a bar to show top 10 and bottom 10 items.";
  }

  const primaryRows = groupBy(sourceRows, group, metric);
  renderBarGroup(els.departmentBars, primaryRows, metric, group, false, "", "", sourceRows);

  const catRows = groupBy(sourceRows, "category", metric).slice(0, 10);
  renderCategoryBreakdown("", group, metric, catRows, sourceRows);
  renderDrillDown("", group, sourceRows);
  renderCompareCards();
}

function renderCategoryBreakdown(groupValue = "", groupKey = "department", metric = "sales", fallbackRows = [], sourceRows = currentInventoryRows()) {
  if (!els.categoryBars) return;
  let rows = fallbackRows;
  let drillKey = "";
  if (els.categoryPanelTitle) {
    if (groupValue && groupKey === "department") {
      els.categoryPanelTitle.textContent = `${groupValue} categories`;
    } else if (groupValue && groupKey !== "department") {
      els.categoryPanelTitle.textContent = `${groupValue} items`;
    }
  }
  if (groupValue && groupKey === "department") {
    rows = groupBy(sourceRows.filter((sku) => sku.department === groupValue), "category", metric);
    drillKey = "category";
  } else if (groupValue && groupKey !== "department") {
    rows = groupBy(sourceRows.filter((sku) => sku[groupKey] === groupValue), "product", metric);
    drillKey = "product";
  }
  renderBarGroup(els.categoryBars, rows.slice(0, 12), metric, drillKey, true, groupValue, groupKey, sourceRows);
}

function renderBarGroup(target, rows, metric, drillKey = "", expandable = false, groupValue = "", groupKey = "", sourceRows = currentInventoryRows()) {
  target.innerHTML = "";
  if (!rows.length) { target.innerHTML = `<p class="muted">Load CSVs or loosen filters.</p>`; return; }
  const max = Math.max(...rows.map((r) => r.value), 1);
  rows.slice(0, 12).forEach((row) => {
    const item = document.createElement("div");
    item.className = "bar-row";
    if (drillKey) item.dataset.drillValue = row.name;
    const sellers = expandable ? filteredBarRowsFor(row.name, drillKey, groupValue, groupKey, sourceRows) : [];
    const topSellers = sellers
      .slice()
      .sort((a, b) => (b[metric] || 0) - (a[metric] || 0))
      .slice(0, 10);
    const bottomSellers = sellers
      .slice()
      .sort((a, b) => (a[metric] || 0) - (b[metric] || 0))
      .slice(0, 10);
    item.innerHTML = expandable ? `
      <details class="bar-detail">
        <summary class="bar-summary">
          <span class="bar-label" title="${escapeHtml(row.name)}">${escapeHtml(row.name)}</span>
          <span class="bar-track"><i class="bar-fill" style="width:${Math.max((row.value / max) * 100, 2)}%"></i></span>
          <span class="bar-value">${formatMetric(row.value, metric)}</span>
        </summary>
        <div class="bar-sellers">
          <div class="bar-seller-group">
            <span class="bar-seller-heading">Top 10</span>
            ${topSellers.map((sku) => `<button type="button" data-detail-code="${escapeHtml(sku.code)}"><span class="ranked-item-name">${escapeHtml(sku.product)}</span><b>${formatMetric(sku[metric] || 0, metric)}</b></button>`).join("") || `<span class="muted">No matching items.</span>`}
          </div>
          <div class="bar-seller-group">
            <span class="bar-seller-heading">Bottom 10</span>
            ${bottomSellers.map((sku) => `<button type="button" data-detail-code="${escapeHtml(sku.code)}"><span class="ranked-item-name">${escapeHtml(sku.product)}</span><b>${formatMetric(sku[metric] || 0, metric)}</b></button>`).join("") || `<span class="muted">No matching items.</span>`}
          </div>
        </div>
      </details>`
      : `
      <span class="bar-label" title="${escapeHtml(row.name)}">${escapeHtml(row.name)}</span>
      <span class="bar-track"><i class="bar-fill" style="width:${Math.max((row.value / max) * 100, 2)}%"></i></span>
      <span class="bar-value">${formatMetric(row.value, metric)}</span>`;
    if (drillKey) {
      item.addEventListener("click", () => {
        target.querySelectorAll(".bar-row").forEach((r) => r.classList.remove("selected"));
        item.classList.add("selected");
        if (target === els.departmentBars) {
          renderCategoryBreakdown(row.name, drillKey, metric, [], sourceRows);
        }
        renderDrillDown(row.name, drillKey, sourceRows);
      });
    }
    target.append(item);
  });
  target.querySelectorAll("[data-detail-code]").forEach((btn) => {
    btn.addEventListener("click", (event) => {
      event.stopPropagation();
      const item = findCurrentItemByCode(btn.dataset.detailCode);
      if (item) showDetail(item);
    });
  });
}

function filteredBarRowsFor(rowName, drillKey, groupValue = "", groupKey = "", sourceRows = currentInventoryRows()) {
  if (drillKey === "category" && groupKey === "department" && groupValue) {
    return sourceRows.filter((sku) => sku.department === groupValue && sku.category === rowName);
  }
  if (drillKey === "product" && groupValue && groupKey) {
    return sourceRows.filter((sku) => sku[groupKey] === groupValue && sku.product === rowName);
  }
  if (drillKey === "category") return sourceRows.filter((sku) => sku.category === rowName);
  if (drillKey === "product") return sourceRows.filter((sku) => sku.product === rowName);
  return sourceRows.filter((sku) => sku[drillKey] === rowName);
}

function renderDrillDown(groupValue = "", groupKey = "department", sourceRows = currentInventoryRows()) {
  const target = document.querySelector("#drillDownPanel");
  if (!target) return;
  const metric = els.segmentMetric.value;
  const rows = groupValue ? sourceRows.filter((sku) => sku[groupKey] === groupValue) : sourceRows;
  // Drill one level deeper: if grouped by deptâ†’show categories; by categoryâ†’show items; by vendorâ†’show items
  const subKey = groupKey === "department" ? "category" : "product";
  const subGroups = groupKey === "department"
    ? groupBy(rows, "category", metric).slice(0, 8)
    : groupBy(rows, "product", metric).slice(0, 10);

  target.innerHTML = `
    <div class="drill-heading">
      <strong>${groupValue ? escapeHtml(groupValue) : `Click a ${groupKey}`}</strong>
      <span>${groupKey === "department" ? "â†’ categories â†’ items" : "â†’ top items"}</span>
    </div>
    ${subGroups.map((sub) => {
      // Top 8 items within that sub-group, sorted by metric descending (not random)
      const rankedItems = rows
        .filter((sku) => (groupKey === "department" ? sku.category : sku.product) === sub.name)
        .sort((a, b) => (b[metric] || 0) - (a[metric] || 0));
      const topItems = rankedItems.slice(0, 10);
      const bottomItems = rankedItems.slice().reverse().slice(0, 10);
      const label = groupKey === "department" ? `Top ${topItems.length} by ${metric}` : `Top and bottom ${Math.min(rankedItems.length, 10)} by ${metric}`;
      return `<details class="drill-category">
        <summary>
          <b>${escapeHtml(sub.name)}</b>
          <span>${formatMetric(sub.value, metric)}</span>
          <small class="drill-label">${label}</small>
        </summary>
        <div class="drill-split">
          <div class="drill-group top">
            <div class="drill-divider">Top 10</div>
            <div class="drill-list">
              ${topItems.map((sku) => `
                <button type="button" data-detail-code="${escapeHtml(sku.code)}">
                  <span class="ranked-item-name">${escapeHtml(sku.product)}</span>
                  <span>${formatMetric(sku[metric] || 0, metric)}</span>
                </button>`).join("")}
            </div>
          </div>
          ${groupKey === "department" ? "" : `
          <div class="drill-group bottom">
            <div class="drill-divider">Bottom 10</div>
            <div class="drill-list">
              ${bottomItems.map((sku) => `
                <button type="button" data-detail-code="${escapeHtml(sku.code)}">
                  <span class="ranked-item-name">${escapeHtml(sku.product)}</span>
                  <span>${formatMetric(sku[metric] || 0, metric)}</span>
                </button>`).join("")}
            </div>
          </div>`}
        </div>
      </details>`;
    }).join("")}`;
  target.querySelectorAll("[data-detail-code]").forEach((btn) => {
    btn.addEventListener("click", () => showDetail(state.filteredSkus.find((sku) => sku.code === btn.dataset.detailCode)));
  });
}

function renderCompareCards() {
  const target = els.compareCards;
  if (!target) return;
  const periodVal = els.comparePeriod?.value || "30";
  const groupKey = els.compareGroup?.value || "department";
  if (!state.dates.length) { target.innerHTML = `<p class="muted">Load sales data to see comparisons.</p>`; return; }

  // Determine current window
  const lastDate = state.dates.at(-1);
  const lastMs = new Date(`${lastDate}T00:00:00`).getTime();
  let curDays, priorDays;

  if (periodVal === "custom") {
    const start = els.startDate.value || state.dates[0];
    const end = els.endDate.value || lastDate;
    curDays = state.dates.filter((d) => d >= start && d <= end);
    const spanMs = (curDays.length - 1) * 86400000 || 86400000;
    const curStartMs = new Date(`${curDays[0] || start}T00:00:00`).getTime();
    const priorEnd = new Date(curStartMs - 86400000).toISOString().slice(0, 10);
    const priorStart = new Date(curStartMs - 86400000 - spanMs).toISOString().slice(0, 10);
    priorDays = state.dates.filter((d) => d >= priorStart && d <= priorEnd);
  } else if (periodVal === "ytd") {
    const year = lastDate.slice(0, 4);
    curDays = state.dates.filter((d) => d.startsWith(year));
    priorDays = state.dates.filter((d) => d.startsWith(String(Number(year) - 1)));
  } else {
    const n = Number(periodVal);
    const curStartMs = lastMs - (n - 1) * 86400000;
    const priorEndMs = curStartMs - 86400000;
    const priorStartMs = priorEndMs - (n - 1) * 86400000;
    curDays = state.dates.filter((d) => new Date(`${d}T00:00:00`).getTime() >= curStartMs);
    priorDays = state.dates.filter((d) => {
      const t = new Date(`${d}T00:00:00`).getTime();
      return t >= priorStartMs && t <= priorEndMs;
    });
  }

  const curDaySet = new Set(curDays);
  const priorDaySet = new Set(priorDays);

  // Aggregate by groupKey for both windows
  const aggregate = (daySet) => {
    const map = new Map();
    state.rawSales.filter((r) => daySet.has(r.date)).forEach((r) => {
      const key = r[groupKey] || "Unassigned";
      const inv = state.latestInventory.get(codeKey(r.code)) || {};
      const resolved = groupKey === "vendor"
        ? (r.vendor !== "Unassigned" ? r.vendor : inv.vendor || "Unassigned")
        : r[groupKey] || inv[groupKey] || "Unassigned";
      const row = map.get(resolved) || { name: resolved, sales: 0, units: 0, profit: 0 };
      row.sales += r.sales; row.units += r.units; row.profit += r.profit;
      map.set(resolved, row);
    });
    return map;
  };

  const curMap = aggregate(curDaySet);
  const priorMap = aggregate(priorDaySet);
  const allKeys = [...new Set([...curMap.keys(), ...priorMap.keys()])];
  const rows = allKeys.map((k) => {
    const cur = curMap.get(k) || { sales: 0, units: 0, profit: 0 };
    const pri = priorMap.get(k) || { sales: 0, units: 0, profit: 0 };
    const salesChg = pri.sales ? ((cur.sales - pri.sales) / pri.sales) * 100 : null;
    const unitsChg = pri.units ? ((cur.units - pri.units) / pri.units) * 100 : null;
    return { name: k, cur, pri, salesChg, unitsChg };
  }).sort((a, b) => b.cur.sales - a.cur.sales);

  if (!rows.length) { target.innerHTML = `<p class="muted">No sales data for the selected period.</p>`; return; }

  const periodLabel = { "30": "30 days", "60": "60 days", "90": "90 days", "182": "6 months", "365": "12 months", ytd: "YTD", custom: "current range" }[periodVal];
  const fmtChg = (v) => v == null ? "â€”" : (v >= 0 ? `â–² +${v.toFixed(1)}%` : `â–¼ ${v.toFixed(1)}%`);
  const chgCls = (v) => v == null ? "" : v >= 0 ? "compare-up" : "compare-down";

  target.innerHTML = `
    <p class="compare-period-label">Current ${periodLabel} vs prior same-length period Â· grouped by ${groupKey}</p>
    <div class="compare-card-grid">
      ${rows.slice(0, 16).map((r) => `
        <div class="compare-card">
          <strong>${escapeHtml(r.name)}</strong>
          <div class="compare-row">
            <span>Sales</span>
            <b>${currency.format(r.cur.sales)}</b>
            <small class="${chgCls(r.salesChg)}">${fmtChg(r.salesChg)}</small>
          </div>
          <div class="compare-row">
            <span>Units</span>
            <b>${number.format(r.cur.units)}</b>
            <small class="${chgCls(r.unitsChg)}">${fmtChg(r.unitsChg)}</small>
          </div>
          <div class="compare-row prior">
            <span>Prior sales</span>
            <b>${currency.format(r.pri.sales)}</b>
          </div>
        </div>`).join("")}
    </div>`;
}

function renderOrderColumnPicker() {
  const panel = document.querySelector("#orderColumnPickerPanel");
  if (!panel || panel.dataset.built) return;
  panel.dataset.built = "1";
  panel.innerHTML = `<div class="column-picker-grid">${orderColumns.map((c) => `
    <label class="column-choice">
      <input type="checkbox" data-order-col="${c.key}" ${state.orderVisibleColumns[c.key] ? "checked" : ""} />
      <span>${c.label}</span>
    </label>`).join("")}</div>`;
  panel.querySelectorAll("[data-order-col]").forEach((cb) => {
    cb.addEventListener("change", () => {
      state.orderVisibleColumns[cb.dataset.orderCol] = cb.checked;
      localStorage.setItem("posOrderColumns:v1", JSON.stringify(state.orderVisibleColumns));
      // Rebuild without re-opening picker
      delete document.querySelector("#orderColumnPickerPanel")?.dataset.built;
      renderOrders();
    });
  });
}

function renderOrders() {
  // Show live formula in header
  const safety = toNumber(els.safetyDays.value) || 7;
  const doi = toNumber(els.daysOfInventory?.value) || 0;
  const note = document.getElementById("formulaNote");
  if (note) note.textContent = `Min = ceil(SV x ${safety}d)  |  Max = ceil(Min${doi ? ` + (SV x ${doi}d DOI)` : ""})  |  Case rounds order only`;

  // Today's ordering alert
  const today = new Date().toLocaleDateString("en-US", { weekday: "long" }).toLowerCase();
  const todayVendors = state.vendorRules.filter((r) => r.status === "Active" && (r.orderDays || []).includes(today));
  const disabledCount = state.vendorRules.filter((r) => r.status !== "Active").length;

  // Write alert to the dedicated banner, not into orderCards (so it persists)
  const banner = document.querySelector("#orderAlertBanner");
  if (banner) {
    let alertHtml = "";
    if (todayVendors.length) {
      alertHtml += `<div class="order-day-alert">
        📅 <b>Order today:</b>
        ${todayVendors.map((r) => {
          const minAlerts2 = (() => { if (!r.minOrder) return false; const vo = currentOrderRows().filter((item) => (item.vendor||"").toUpperCase()===r.vendor?.toUpperCase()); const total = vo.reduce((s,i)=>s+(i.caseOrder||0)*(i.unitCost||0),0); return total < r.minOrder; })();
          const isPending = state.pendingOrders?.some((po) => po.vendor === r.vendor && !po.cleared);
          return `<button type="button" class="order-vendor-chip-btn${minAlerts2?" order-min-warn":""}${isPending?" order-pending-chip":""}" data-vendor-order="${escapeHtml(r.vendor)}">${escapeHtml(r.vendor)}${minAlerts2?` ⚠ min`:""}${isPending?" 🕐":""}</button>`;
        }).join("")}
        <button type="button" id="submitAllPoButton" class="count-submit-btn" style="margin-left:auto">Submit All PO</button>
      </div>`;
    }
    banner.innerHTML = alertHtml;
    // Wire chips
    banner.querySelectorAll(".order-vendor-chip-btn").forEach((btn) => {
      btn.addEventListener("click", () => openVendorAnalysisPanel(btn.dataset.vendorOrder));
    });
    document.querySelector("#submitAllPoButton")?.addEventListener("click", () => submitAllPo());
  }

  const orders = currentOrderRows().slice(0, 80);
  // Render column picker
  renderOrderColumnPicker();

  const visibleCols = orderColumns.filter((c) => state.orderVisibleColumns[c.key]);
  const sortKey = state.orderSort?.key || "recommendedOrder";
  const sortDir = state.orderSort?.dir || "desc";

  // Sort orders
  const sorted = currentOrderRows().sort((a, b) => {
    const av = a[sortKey] ?? 0, bv = b[sortKey] ?? 0;
    if (typeof av === "string") return sortDir === "asc" ? av.localeCompare(bv) : bv.localeCompare(av);
    return sortDir === "asc" ? av - bv : bv - av;
  }).slice(0, 120);

  els.orderCards.innerHTML = "";
  if (!sorted.length) {
    els.orderCards.innerHTML = `<p class="muted">No reorder recommendations. Discontinued, disabled, and vendor-disabled items are excluded.</p>`;
    return;
  }

  const headerHtml = visibleCols.map((c) => {
    const isActive = c.key === sortKey;
    const arrow = isActive ? (sortDir === "asc" ? " ↑" : " ↓") : "";
    return `<th class="order-sortable-th${isActive ? " sort-active" : ""}" data-order-sort="${c.key}">${c.label}${arrow}</th>`;
  }).join("");

  els.orderCards.innerHTML = `
    <div class="table-wrap order-table-wrap" style="overflow-x:auto;max-height:none">
      <table class="order-table" style="table-layout:auto;min-width:max-content">
        <thead>
          <tr>
            <th class="checkbox-col"><input type="checkbox" id="selectAllOrdering" title="Select all"></th>
            ${headerHtml}
          </tr>
        </thead>
        <tbody>
          ${sorted.map((sku) => {
            const isPend = isPendingOrder(sku.code);
            const cellMap = {
              status:           `<td>${stateBadgeHtml(sku)}</td>`,
              pending:          `<td style="text-align:center;width:2rem">${isPend ? `<button type="button" class="pending-clear-btn" data-code="${escapeHtml(sku.code)}" title="Click to clear pending">&#x1F550;</button>` : ""}</td>`,
              code:             `<td class="copy-code" data-copy-code="${escapeHtml(sku.code)}" style="color:#2470c4;font-weight:700;cursor:pointer">${escapeHtml(sku.code)}</td>`,
              product:          `<td class="sku-name">${escapeHtml(sku.product)}</td>`,
              vendor:           `<td>${escapeHtml(sku.vendor || "-")}</td>`,
              plu:              `<td>${escapeHtml(sku.plu || "-")}</td>`,
              velocity:         `<td class="num">${formatVelocity(sku.velocity || 0)}</td>`,
              units:            `<td class="num">${number.format(sku.units || 0)}</td>`,
              stock:            `<td class="num ${(sku.stock||0) < 0 ? "entry-negative" : ""}">${number.format(sku.stock || 0)}</td>`,
              reorderMin:       `<td class="num">${number.format(sku.reorderMin || 0)}</td>`,
              reorderMax:       `<td class="num">${number.format(sku.reorderMax || 0)}</td>`,
              recommendedOrder: `<td class="num order-highlight"><input type="number" class="order-rec-input mini-input" data-code="${escapeHtml(sku.code)}" value="${sku.recommendedOrder || sku.qtyNeeded || 0}" min="0" style="width:3.8rem;text-align:center;font-weight:700" /></td>`,
              caseOrder:        `<td class="num order-highlight"><b>${number.format(sku.caseOrder || 0)}</b></td>`,
              caseSize:         `<td class="num">${number.format(sku.caseSize || 1)}</td>`,
              unitCost:         `<td class="num">${currency.format(sku.unitCost || 0)}</td>`,
              totalCost:        `<td class="num">${currency.format((sku.caseOrder || 0) * (sku.unitCost || 0))}</td>`,
            };
            return `<tr data-detail-code="${escapeHtml(sku.code)}">
              <td class="checkbox-col"><input type="checkbox" class="row-checkbox order-checkbox" data-code="${escapeHtml(sku.code)}" ${state.selectedSkuCodes.has(sku.code) ? "checked" : ""}></td>
              ${visibleCols.map((c) => cellMap[c.key] || "<td></td>").join("")}
            </tr>`;
          }).join("")}
        </tbody>
      </table>
    </div>`;

  // Wire sortable headers
  els.orderCards.querySelectorAll(".order-sortable-th").forEach((th) => {
    th.style.cursor = "pointer";
    th.addEventListener("click", () => {
      const key = th.dataset.orderSort;
      state.orderSort = { key, dir: state.orderSort?.key === key && state.orderSort?.dir === "asc" ? "desc" : "asc" };
      renderOrders();
    });
  });
  // Wire checkboxes
  els.orderCards.querySelectorAll(".order-checkbox").forEach((cb) => {
    cb.addEventListener("change", () => {
      if (cb.checked) state.selectedSkuCodes.add(cb.dataset.code);
      else state.selectedSkuCodes.delete(cb.dataset.code);
    });
  });
  const selectAllOrdering = document.querySelector("#selectAllOrdering");
  if (selectAllOrdering) {
    selectAllOrdering.addEventListener("change", (e) => {
      els.orderCards.querySelectorAll(".order-checkbox").forEach((cb) => {
        cb.checked = e.target.checked;
        if (e.target.checked) state.selectedSkuCodes.add(cb.dataset.code);
        else state.selectedSkuCodes.delete(cb.dataset.code);
      });
    });
  }
  // Wire editable Rec Order inputs — stopPropagation prevents opening detail drawer
  els.orderCards.querySelectorAll(".order-rec-input").forEach((input) => {
    input.addEventListener("click", (e) => e.stopPropagation());
    input.addEventListener("change", () => {
      const code = input.dataset.code;
      const val = Math.max(0, toNumber(input.value) || 0);
      // Override in local state
      if (!state._orderRecOverrides) state._orderRecOverrides = new Map();
      state._orderRecOverrides.set(codeKey(code), val);
      input.value = val;
    });
  });
  // Wire per-item clear-pending buttons
  els.orderCards.querySelectorAll(".pending-clear-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const code = btn.dataset.code;
      state.pendingOrders = (state.pendingOrders || []).map((po) => {
        if (po.cleared) return po;
        const newCodes = (po.codes || []).filter((c) => c !== codeKey(code));
        return { ...po, codes: newCodes };
      });
      savePendingOrders();
      renderOrders();
      showToast("Pending cleared for " + code, 2000);
    });
  });
  // Vendor analysis panel wiring now handled in the banner block above
}

function renderTable() {
  const rows = state.filteredSkus.slice(0, 500);
  els.skuBody.innerHTML = "";
  if (!rows.length) {
    els.skuBody.innerHTML = `<tr><td colspan="20" class="empty-cell">No matching SKUs yet.</td></tr>`;
    return;
  }
  const fragment = document.createDocumentFragment();
  rows.forEach((sku) => {
    const tr = document.createElement("tr");
    tr.dataset.detailCode = sku.code;
    const isChecked = state.selectedSkuCodes.has(sku.code);
    tr.innerHTML = `
      <td class="checkbox-col"><input type="checkbox" class="row-checkbox" data-code="${escapeHtml(sku.code)}" ${isChecked ? "checked" : ""}></td>
      <td><span class="badge ${sku.status}">${labelStatus(sku.status)}</span></td>
      <td>${escapeHtml(sku.code)}</td>
      <td class="sku-name">${escapeHtml(sku.product)}</td>
      <td>${escapeHtml(sku.plu || "-")}</td>
      <td>${escapeHtml(sku.itemNumber || "-")}</td>
      <td>${escapeHtml(sku.department)}</td>
      <td>${escapeHtml(sku.category)}</td>
      <td>${escapeHtml(sku.vendor)}</td>
      <td>${escapeHtml(sku.color || "-")}</td>
      <td class="num">${formatVelocity(sku.velocity)}</td>
      <td class="num">${number.format(sku.units)}</td>
      <td class="num">${currency.format(sku.sales)}</td>
      <td class="num">${currency.format(sku.costSold)}</td>
      <td class="num">${currency.format(sku.profit)}</td>
      <td class="num">${number.format(sku.stock)}</td>
      <td class="num">${currency.format(sku.unitCost)}</td>
      <td class="num">${currency.format(sku.price)}</td>
      <td class="num">${formatDays(sku.daysSupply)}</td>
      <td class="num">${number.format(sku.recommendedOrder)}</td>`;
    fragment.append(tr);
  });
  els.skuBody.append(fragment);
  // delegate checkbox clicks
  els.skuBody.querySelectorAll(".row-checkbox").forEach((cb) => {
    cb.addEventListener("change", () => {
      if (cb.checked) state.selectedSkuCodes.add(cb.dataset.code);
      else state.selectedSkuCodes.delete(cb.dataset.code);
    });
  });
}

function renderInventory() {
  const rows = currentInventoryRows();
  els.inventoryBody.innerHTML = "";
  renderInventorySummary(rows);
  renderInventoryHeader();
  if (!rows.length) {
    els.inventoryBody.innerHTML = `<tr><td colspan="20" class="empty-cell">Load CSV inventory or the Excel import to view all items.</td></tr>`;
    return;
  }

  // Chunked rendering — build DOM in batches via requestAnimationFrame so the
  // browser stays responsive. Tooltip HTML is deferred until mouseover.
  const CHUNK = 80;
  const visible = rows.slice(0, 1200);
  let offset = 0;

  function renderChunk() {
    const fragment = document.createDocumentFragment();
    const end = Math.min(offset + CHUNK, visible.length);
    for (let i = offset; i < end; i++) {
      const item = visible[i];
      const tr = document.createElement("tr");
      tr.dataset.itemCode = item.code;
      tr.dataset.tooltipDeferred = "1"; // tooltip built on first hover, not upfront
      const isChecked = state.selectedInventoryCodes.has(item.code);
      const cbTd = document.createElement("td");
      cbTd.className = "checkbox-col";
      const cbInput = document.createElement("input");
      cbInput.type = "checkbox";
      cbInput.className = "row-checkbox";
      cbInput.dataset.code = item.code;
      cbInput.checked = isChecked;
      cbInput.addEventListener("change", () => {
        if (cbInput.checked) state.selectedInventoryCodes.add(item.code);
        else state.selectedInventoryCodes.delete(item.code);
      });
      cbTd.append(cbInput);
      tr.append(cbTd);
      tr.insertAdjacentHTML("beforeend", state.columnOrder.map((key) => inventoryCellHtml(key, item)).join(""));
      fragment.append(tr);
    }
    els.inventoryBody.append(fragment);
    offset = end;
    if (offset < visible.length) {
      requestAnimationFrame(renderChunk);
    } else {
      els.inventoryBody.querySelectorAll(".pending-clear-btn").forEach((btn) => {
        btn.addEventListener("click", (event) => {
          event.stopPropagation();
          const code = btn.dataset.code;
          state.pendingOrders = (state.pendingOrders || []).map((po) => {
            if (po.cleared) return po;
            const newCodes = (po.codes || []).filter((c) => c !== codeKey(code));
            return { ...po, codes: newCodes };
          });
          savePendingOrders();
          renderInventory();
          renderOrders();
          showToast("Pending cleared for " + code, 2000, "success");
        });
      });
      applyColumnVisibility();
      syncInventoryHeaderOffset();
    }
  }
  requestAnimationFrame(renderChunk);
}

function syncInventoryHeaderOffset() {
  const root = document.documentElement;
  root.style.setProperty("--inventory-summary-height", "0px");
}

function renderInventoryHeader() {
  const row = document.querySelector("#inventory thead tr");
  if (!row) return;
  const labels = Object.fromEntries(inventoryColumns);
  // Always start with the checkbox th, then data columns
  row.innerHTML = `<th class="checkbox-col" style="width:28px;min-width:28px;max-width:28px"><input type="checkbox" id="selectAllInventory" title="Select / deselect all" /></th>` +
    state.columnOrder.map((key) => `
    <th data-col="${key}" data-sort="${key}" ${state.arrangeColumns ? 'draggable="true"' : ""}>
      ${escapeHtml(labels[key] || key)}
    </th>`).join("");
  // Re-wire select-all after rebuild
  const cb = row.querySelector("#selectAllInventory");
  cb?.addEventListener("change", (e) => {
    document.querySelectorAll("#inventoryBody .row-checkbox").forEach((c) => {
      c.checked = e.target.checked;
      if (e.target.checked) state.selectedInventoryCodes.add(c.dataset.code);
      else state.selectedInventoryCodes.delete(c.dataset.code);
    });
  });
  row.querySelectorAll("[data-sort]").forEach((header) => {
    header.addEventListener("click", () => {
      if (state.arrangeColumns) return;
      state._pinnedAdjustCode = null;
      const key = header.dataset.sort;
      state.inventorySort = {
        key,
        dir: state.inventorySort.key === key && state.inventorySort.dir === "asc" ? "desc" : "asc",
      };
      renderInventory();
    });
    header.addEventListener("dragstart", (event) => {
      if (!state.arrangeColumns) return;
      event.dataTransfer.setData("text/plain", header.dataset.col);
    });
    header.addEventListener("dragover", (event) => {
      if (state.arrangeColumns) event.preventDefault();
    });
    header.addEventListener("drop", (event) => {
      if (!state.arrangeColumns) return;
      event.preventDefault();
      moveColumn(event.dataTransfer.getData("text/plain"), header.dataset.col);
    });
  });
  applyColumnVisibility();
  updateSortHeaders();
}

function stateBadgeHtml(item) {
  const s = (item.state || "").toLowerCase();
  const cls = s === "active" ? "state-active"
    : s === "discontinued" ? "state-discontinued"
    : s === "disabled" ? "state-disabled"
    : s === "force order" ? "state-forceorder"
    : "state-unknown";
  return `<span class="state-badge ${cls}">${escapeHtml(item.state || "-")}</span>`;
}

function inventoryCellHtml(key, item) {
  const override = state.reorderOverrides[item.code] || {};
  const minOverridden = override.min != null;
  const maxOverridden = override.max != null;
  const values = {
    pending: `<td data-col="pending" class="num pending-col">${isPendingOrder(item.code) ? `<button type="button" class="pending-clear-btn pending-inline-btn" data-code="${escapeHtml(item.code)}" title="PO pending - click to clear">&#x1F550;</button>` : ""}</td>`,
    code: `<td data-col="code" class="copy-code" data-copy-code="${escapeHtml(item.code)}" title="Click code to copy">${escapeHtml(item.code)}</td>`,
    product: `<td data-col="product" class="sku-name">${escapeHtml(item.product)}</td>`,
    plu: `<td data-col="plu">${escapeHtml(item.plu || "-")}</td>`,
    itemNumber: `<td data-col="itemNumber">${escapeHtml(item.itemNumber || "-")}</td>`,
    subType: `<td data-col="subType">${escapeHtml(item.subType || "-")}</td>`,
    sizeAttr: `<td data-col="sizeAttr">${escapeHtml(item.sizeAttr || "-")}</td>`,
    containerAttr: `<td data-col="containerAttr">${escapeHtml(item.containerAttr || "-")}</td>`,
    category: `<td data-col="category">${escapeHtml(item.category || "-")}</td>`,
    vendor: `<td data-col="vendor">${escapeHtml(item.vendor || "-")}</td>`,
    state: `<td data-col="state">${stateBadgeHtml(item)}</td>`,
    addDate: `<td data-col="addDate">${escapeHtml(item.addDate || "-")}</td>`,
    stock: `<td data-col="stock" class="num stock-col stock-clickable" title="Click to adjust stock">${number.format(item.stock)}</td>`,
    units: `<td data-col="units" class="num sold-col">${number.format(item.units)}</td>`,
    velocity: `<td data-col="velocity" class="num">${formatVelocity(item.velocity)}</td>`,
    unitCost: `<td data-col="unitCost" class="num">${currency.format(item.unitCost)}</td>`,
    price: `<td data-col="price" class="num">${currency.format(item.price)}</td>`,
    inventoryCost: `<td data-col="inventoryCost" class="num">${currency.format(item.inventoryCost)}</td>`,
    caseSize: `<td data-col="caseSize" class="num">${number.format(item.caseSize || 1)}</td>`,
    reorderMin: `<td data-col="reorderMin" class="num order-col ${minOverridden ? "override-cell" : ""}">
      ${minOverridden ? `<button class="reset-override" data-code="${escapeHtml(item.code)}" data-field="min" title="Reset to auto (${number.format(item.dynamicMin)})">ðŸ”’</button>` : ""}
      <input class="mini-input ${minOverridden ? "overridden" : ""}" data-code="${escapeHtml(item.code)}" data-reorder-field="min"
        value="${escapeHtml(item.reorderMin)}" placeholder="${number.format(item.dynamicMin)}" title="${minOverridden ? `Auto: ${number.format(item.dynamicMin)} â€” manually set` : `Auto: SVÃ—Safety = ${number.format(item.dynamicMin)}`}" />
    </td>`,
    reorderMax: `<td data-col="reorderMax" class="num order-col ${maxOverridden ? "override-cell" : ""}">
      ${maxOverridden ? `<button class="reset-override" data-code="${escapeHtml(item.code)}" data-field="max" title="Reset to auto (${number.format(item.dynamicMax)})">ðŸ”’</button>` : ""}
      <input class="mini-input ${maxOverridden ? "overridden" : ""}" data-code="${escapeHtml(item.code)}" data-reorder-field="max"
        value="${escapeHtml(item.reorderMax)}" placeholder="${number.format(item.dynamicMax)}" title="${maxOverridden ? `Auto: ${number.format(item.dynamicMax)} â€” manually set` : `Auto: Min+(SVÃ—DOI) = ${number.format(item.dynamicMax)}`}" />
    </td>`,
    needs: (() => {
      const needed = item.recommendedOrder || 0;
      if (!needed || item.isOrderable === false) return `<td data-col="needs" class="num"></td>`;
      return `<td data-col="needs" class="num needs-col needs-alert">${number.format(needed)}</td>`;
    })(),
  };
  return values[key] || "";
}

function moveColumn(from, to) {
  if (!from || !to || from === to) return;
  const order = state.columnOrder.filter((key) => key !== from);
  order.splice(order.indexOf(to), 0, from);
  state.columnOrder = order;
  localStorage.setItem("posDashboardColumnOrder:v3", JSON.stringify(state.columnOrder));
  renderInventory();
}

function moveDetailField(from, to) {
  if (!from || !to || from === to) return;
  const defaultOrder = [
    "code", "plu", "itemNumber", "parent", "subType", "sizeAttr", "containerAttr", "otherAttrs", "qty", "windows", "sales", "costSold", "velocity",
    "vendor", "category", "color", "state", "addDate", "stock", "unitCost", "price", "inventoryCost",
    "caseSize", "minMax", "recommendedOrder",
  ];
  const order = (state.detailOrder || defaultOrder).filter((key) => key !== from);
  order.splice(order.indexOf(to), 0, from);
  state.detailOrder = order;
  localStorage.setItem("posDashboardDetailOrder:v1", JSON.stringify(order));
}

function renderInventorySummary(rows) {
  const totals = rows.reduce(
    (sum, item) => ({
      items: sum.items + 1,
      stock: sum.stock + item.stock,
      sold: sum.sold + item.units,
      velocity: sum.velocity + item.velocity,
      unitCost: sum.unitCost + item.unitCost,
      inventoryCost: sum.inventoryCost + item.inventoryCost,
    }),
    { items: 0, stock: 0, sold: 0, velocity: 0, unitCost: 0, inventoryCost: 0 },
  );
  const addDates = rows.map((item) => cleanCell(item.addDate)).filter(Boolean).sort(compareDateValue);
  const earliestAddDate = addDates[0] || "";
  const latestAdd = addDates.at(-1) || latestExcelAddDate();
  els.inventorySummary.innerHTML = `
    <span><b>${number.format(totals.items)}</b> items showing</span>
    <span><b>${number.format(totals.stock)}</b> stock</span>
    <span><b>${number.format(totals.sold)}</b> sold</span>
    <span><b>${formatVelocity(totals.velocity)}</b> avg/day</span>
    <span><b>${currency.format(totals.unitCost)}</b> cost sum</span>
    <span><b>${currency.format(totals.inventoryCost)}</b> stock cost</span>
    <span><b>${escapeHtml(earliestAddDate || "-")}</b> oldest add</span>
    <span><b>${escapeHtml(latestAdd || "-")}</b> newest add</span>`;
}

function buildInventoryRows(options = {}) {
  const query = options.ignoreQuery ? "" : els.searchInput.value.trim().toLowerCase();
  const stateFilter = (options.ignoreFilters || options.ignoreStateFilter) ? "" : els.inventoryStateFilter.value;
  const start = els.startDate.value || "0000-00-00";
  const end = els.endDate.value || "9999-99-99";
  const leadDays = 0;
  const safetyDays = Math.max(0, toNumber(els.safetyDays.value) || 7);
  const daysOfInventory = Math.max(0, toNumber(els.daysOfInventory?.value) || 0);
  const cacheKey = `${start}|${end}|${leadDays}|${safetyDays}|${daysOfInventory}`;
  if (!state._inventoryCache || state._inventoryCacheStamp !== state._dataCacheStamp || state._inventoryCacheKey !== cacheKey) {
    const inventoryIndex = state.latestInventory;
    const salesByCode = new Map(buildSkuRows({ ignoreQuery: true, ignoreFilters: true }).map((sku) => [codeKey(sku.code), sku]));
    const hasCurrentInventory = inventoryIndex.size > 0;
    const codes = hasCurrentInventory
      ? new Set([...inventoryIndex.keys()])
      : new Set([...state.excelItems.values()].map((item) => item.code).filter(Boolean).concat([...salesByCode.values()].map((sku) => sku.code)));

    state._inventoryCache = [...codes].map((code) => {
      const inventory = inventoryIndex.get(codeKey(code)) || {};
      const excel = findExcelFor(inventory.code ? inventory : { code });
      const sales = salesByCode.get(codeKey(code)) || {};
      const itemCode = inventory.code || code;
      const override = state.reorderOverrides[itemCode] || state.reorderOverrides[code] || {};
      const isOverridden = override.min != null || override.max != null;
      const velocity = sales.velocity || excel.saleVelocity || 0;
      const caseSize = excel.caseSize || sales.caseSize || 1;
      const itemState = (excel.state || "").toLowerCase();
      const isOrderable = !["discontinued", "disabled"].includes(itemState);
      // Use vendor-specific safety/doi if a rule exists for this vendor
      const vendorName = (inventory.vendor || excel.vendor || sales.vendor || "").toUpperCase();
      const vendorRule = state.vendorRules.find((r) => r.vendor?.toUpperCase() === vendorName && r.status === "Active");
      const effectiveSafety = vendorRule ? vendorRule.safetyDays : safetyDays;
      const effectiveDoi = vendorRule ? vendorRule.daysOfInventory : daysOfInventory;
      const dynamic = orderingTargets({ velocity, safetyDays: effectiveSafety, daysOfInventory: effectiveDoi });
      const dynamicMaxWithDoi = dynamic.max;
      const stock = hasValue(inventory.stock) ? inventory.stock : excel.stock ?? sales.stock ?? 0;
      const item = {
        code: itemCode,
        product: bestItemName(inventory.product, excel.product, sales.product, inventory.plu, code),
        department: sales.department || "",
        category: inventory.category || excel.category || sales.category || "",
        vendor: inventory.vendor || excel.vendor || sales.vendor || "",
        plu: inventory.plu || excel.plu || sales.plu || "",
        itemNumber: inventory.itemNumber || excel.itemNumber || sales.itemNumber || "",
        color: inventory.color || sales.color || "",
        state: excel.state || "",
        itemState,
        isOrderable,
        isOverridden,
        addDate: excel.addDate || inventory.date || "",
        snapshotDate: inventory.date || "",
        stock,
        units: sales.units || 0,
        sales: sales.sales || 0,
        costSold: sales.costSold || 0,
        profit: sales.profit || 0,
        velocity,
        unitCost: pickNumber(inventory.cost, excel.cost, sales.unitCost),
        price: pickNumber(inventory.price, excel.price, sales.price),
        caseSize,
        dynamicMin: isOrderable ? dynamic.min : 0,
        dynamicMax: isOrderable ? dynamicMaxWithDoi : 0,
        reorderMin: isOrderable ? (override.min ?? dynamic.min) : 0,
        reorderMax: isOrderable ? (override.max ?? dynamicMaxWithDoi) : 0,
        qtyNeeded: sales.recommendedOrder || 0,
        saleWindowSum: excel.saleWindowSum || sales.units || 0,
      };
      item.inventoryCost = item.stock * item.unitCost;
      Object.assign(item, parentPartsFor(item));
      item.subGroup = item.sizeAttr || item.containerAttr || item.otherAttrs || "";
      item.typeGroup = item.subType || item.color || item.itemNumber || item.code;
      item.recommendedOrder = isOrderable ? recommendedOrderQty({
        stock: item.stock,
        min: item.reorderMin,
        max: item.reorderMax,
        caseSize: item.caseSize,
      }) : 0;
      item.caseOrder = item.recommendedOrder > 0 ? roundOrderToNearestCase(item.recommendedOrder, item.caseSize) : 0;
      item.daysSupply = velocity > 0 ? stock / velocity : Infinity;
      item._haystack = [item.code, item.product, item.vendor, item.category, item.color, item.plu, item.itemNumber, item.state, item.parent, item.subType, item.sizeAttr, item.containerAttr, item.otherAttrs, item.subGroup, item.typeGroup, String(item.plu||''), String(item.itemNumber||''), item.department].filter(Boolean).join(" ").toLowerCase().trim();
      return item;
    });
    state._inventoryCacheStamp = state._dataCacheStamp;
    state._inventoryCacheKey = cacheKey;
  }

  return state._inventoryCache.filter((item) => {
    // Force Order is treated as Active — always include it when filtering by Active
    if (stateFilter) {
      const itemState = item.state || "";
      if (stateFilter === "Active") {
        if (itemState !== "Active" && itemState.toLowerCase() !== "force order") return false;
      } else if (itemState !== stateFilter) {
        return false;
      }
    }
    if (query && !item._haystack.includes(query.toLowerCase())) return false;
    if (!options.ignoreFilters) {
      if (els.departmentFilter.value && item.department && item.department !== els.departmentFilter.value) return false;
      if (els.categoryFilter.value && item.category !== els.categoryFilter.value) return false;
      if (els.vendorFilter.value && item.vendor !== els.vendorFilter.value) return false;
      if (els.colorFilter.value && item.color !== els.colorFilter.value && item.subType !== els.colorFilter.value) return false;
    }
    return true;
  }).sort(compareInventoryRows);
}

function renderParents() {
  currentInventoryRows();
  if (ENABLE_CUSTOM_PARENT_RULES) renderParentRules();
  if (ENABLE_CUSTOM_ATTRIBUTE_RULES) renderAttributeRules();
  const query = (els.parentsSearch?.value || "").trim().toLowerCase();
  const grouped = new Map();
  state.inventoryRows.forEach((item) => {
    const parent = item.parent || item.product;
    if (query && !`${parent} ${item.subGroup || ""} ${item.typeGroup || ""} ${item.product}`.toLowerCase().includes(query)) return;
    const row = grouped.get(parent) || { parent, units: 0, stock: 0, sales: 0, inventoryCost: 0, children: [] };
    row.units += item.units || 0;
    row.stock += item.stock || 0;
    row.sales += item.sales || 0;
    row.inventoryCost += item.inventoryCost || 0;
    row.children.push({ ...item, sales: item.sales || 0 });
    grouped.set(parent, row);
  });
  state.parentRows = [...grouped.values()].map((group) => ({
    ...group,
    subtypeCount: new Set(group.children.map((child) => child.subType || child.color || child.itemNumber || child.code)).size,
    children: group.children.sort(compareSubTypeOrder),
  })).sort((a, b) => b.units - a.units);
  els.parentGrid.innerHTML = "";
  if (!state.parentRows.length) {
    els.parentGrid.innerHTML = `<p class="muted">Load the current inventory file to build parent style groups.</p>`;
    return;
  }
  state.parentRows.slice(0, 80).forEach((group) => {
    const details = document.createElement("details");
    details.className = "parent-card";

    // Build 3-level tree: Parent -> Sub -> Type
    const sizeGroups = buildSizeGroups(group.children);
    const hasSize = sizeGroups.some((sg) => sg.size !== "");

    details.innerHTML = `
      <summary>
        <strong>${escapeHtml(group.parent)}</strong>
        <span>${number.format(group.subtypeCount)} sub/types</span>
        <span>${number.format(group.units)} sold</span>
        <span>${number.format(group.stock)} stock</span>
        <span>${currency.format(group.sales)}</span>
      </summary>
      <div class="parent-children ${hasSize ? "has-size-groups" : ""}">
        ${hasSize ? sizeGroups.map((sg) => `
          <details class="size-group">
            <summary class="size-group-summary">
              <b>${escapeHtml(sg.size || "Other")}</b>
              <span>${number.format(sg.units)} sold</span>
              <span>${number.format(sg.stock)} stock</span>
              <span>${number.format(sg.children.length)} types</span>
            </summary>
            <div class="size-group-children">
              ${sg.children.map((child) => skuButtonHtml(child)).join("")}
            </div>
          </details>`).join("")
        : group.children.slice(0, 60).map((child) => skuButtonHtml(child)).join("")}
      </div>`;

    details.querySelectorAll("[data-detail-code]").forEach((button) => {
      button.addEventListener("click", () => showDetail(state.inventoryRows.find((item) => item.code === button.dataset.detailCode)));
    });
    els.parentGrid.append(details);
  });
}

function skuButtonHtml(child) {
  return `<button type="button" data-detail-code="${escapeHtml(child.code)}">
    <b>${escapeHtml(childLabel(child, { omitSize: true }))}</b>
    <span>${number.format(child.units)} sold</span>
    <span>${number.format(child.stock)} stock</span>
    <span>${currency.format(child.sales)}</span>
  </button>`;
}

// Groups children by their sizeAttr, aggregating units/stock per size bucket
function buildSizeGroups(children) {
  const map = new Map();
  children.forEach((child) => {
    const size = cleanCell(child.subGroup || child.sizeAttr || "");
    const sg = map.get(size) || { size, units: 0, stock: 0, children: [] };
    sg.units += child.units || 0;
    sg.stock += child.stock || 0;
    sg.children.push(child);
    map.set(size, sg);
  });
  return [...map.values()].sort((a, b) => {
    if (a.size && !b.size) return -1;
    if (!a.size && b.size) return 1;
    const aNum = parseInt(a.size, 10);
    const bNum = parseInt(b.size, 10);
    if (Number.isFinite(aNum) && Number.isFinite(bNum) && aNum !== bNum) return aNum - bNum;
    return a.size.localeCompare(b.size);
  });
}

function updateFilterOptions() {
  const salesWithInventory = state.rawSales.map((row) => ({ ...row, inventory: state.latestInventory.get(codeKey(row.code)) || {} }));
  const inventoryRows = [...state.latestInventory.values()];
  fillSelect(els.departmentFilter, unique(salesWithInventory.map((row) => row.department)));
  fillSelect(els.categoryFilter, unique(salesWithInventory.map((row) => row.category || row.inventory.category).concat(inventoryRows.map((row) => row.category))));
  fillSelect(els.vendorFilter, unique(salesWithInventory.map((row) => (row.vendor !== "Unassigned" ? row.vendor : row.inventory.vendor)).concat(inventoryRows.map((row) => row.vendor))));
  fillSelect(els.colorFilter, unique(salesWithInventory.map((row) => row.inventory.color).concat(inventoryRows.map((row) => row.color))));
  updateInventoryStateFilter();
}

function updateInventoryStateFilter() {
  fillSelect(els.inventoryStateFilter, unique([...state.excelItems.values()].map((item) => item.state)));
  if (![...els.inventoryStateFilter.options].some((option) => option.value === els.inventoryStateFilter.value)) {
    els.inventoryStateFilter.value = "";
  }
}

function fillSelect(select, values) {
  const current = select.value;
  select.innerHTML = `<option value="">All</option>`;
  values.filter(Boolean).sort(compareDisplayValue).forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = value;
    select.append(option);
  });
  select.value = values.includes(current) ? current : "";
}

function setDefaultDates() {
  if (!state.dates.length) return;
  els.startDate.value = state.dates[0];
  els.endDate.value = state.dates[state.dates.length - 1];
  state.activePresetDays = "all";
}

function filteredSalesDates() {
  const start = els.startDate.value || "0000-00-00";
  const end = els.endDate.value || "9999-99-99";
  return state.dates.filter((date) => date >= start && date <= end);
}

function groupBy(rows, key, metric) {
  const grouped = new Map();
  rows.forEach((row) => {
    const name = row[key] || "Unassigned";
    const existing = grouped.get(name) || { name, sales: 0, units: 0, profit: 0, velocity: 0 };
    existing.sales += row.sales;
    existing.units += row.units;
    existing.profit += row.profit;
    existing.velocity += row.velocity;
    grouped.set(name, existing);
  });
  return [...grouped.values()].map((row) => ({ name: row.name, value: row[metric] || 0 })).sort((a, b) => b.value - a.value);
}

function downloadCsv(fileName, rows) {
  if (!rows.length) return;
  const preferred = ["status", "parent", "subType", "sizeAttr", "containerAttr", "otherAttrs", "code", "product", "plu", "itemNumber", "department", "category", "vendor", "color", "state", "addDate", "units", "sales", "costSold", "profit", "velocity", "stock", "unitCost", "price", "inventoryCost", "caseSize", "reorderMin", "reorderMax", "daysSupply", "recommendedOrder", "qtyNeeded"];
  const available = new Set(rows.flatMap((row) => Object.keys(row)).filter((key) => key !== "children"));
  const headers = preferred.filter((key) => available.has(key)).concat([...available].filter((key) => !preferred.includes(key)));
  const csv = [headers.join(",")].concat(rows.map((row) => headers.map((header) => csvCell(row[header], header)).join(","))).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

// â”€â”€ PO Export helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€

function buildPoRows() {
  return currentOrderRows()
    .sort((a, b) => (a.vendor || "").localeCompare(b.vendor || "") || (b.recommendedOrder - a.recommendedOrder))
    .map((item) => ({
      vendor:           item.vendor || "-",
      code:             item.code,
      product:          item.product,
      plu:              item.plu || "-",
      itemNumber:       item.itemNumber || "-",
      caseSize:         item.caseSize || 1,
      orderQty:         item.recommendedOrder || item.qtyNeeded || 0,
      caseOrderQty:     item.caseOrder || 0,
      unitCost:         item.unitCost || 0,
      totalCost:        (item.caseOrder || 0) * (item.unitCost || 0),
      currentStock:     item.stock || 0,
      svPerDay:         Number(item.velocity || 0).toFixed(2),
      daysSupply:       Number.isFinite(item.daysSupply) ? Math.round(item.daysSupply) : "∞",
      reorderMin:       item.reorderMin || 0,
      reorderMax:       item.reorderMax || 0,
      isOverridden:     item.isOverridden ? "Yes" : "No",
    }));
}

// Helper: make xlsx cell a forced-text string (preserves leading zeros)
function xlsxTextCell(value) {
  return { v: String(value ?? ""), t: "s" };
}

function applyXlsxTextToCodeColumns(ws, data, codeColIndexes) {
  // data is array-of-arrays; codeColIndexes are 0-based column indices to force as text
  if (!ws || !data) return;
  const XLSX = window.XLSX;
  if (!XLSX) return;
  data.forEach((row, r) => {
    codeColIndexes.forEach((c) => {
      const cell_ref = XLSX.utils.encode_cell({ r, c });
      if (ws[cell_ref]) {
        ws[cell_ref].t = "s";
        ws[cell_ref].z = "@";
      }
    });
  });
}

async function exportPoExcel() {
  const rows = buildPoRows();
  if (!rows.length) { showToast("No items to order."); return; }
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available."); return; }

  // Group by vendor for separate sheets
  const vendors = [...new Set(rows.map((r) => r.vendor))];
  const wb = xlsx.utils.book_new();

  // Summary sheet â€” all items
  const summaryData = [
    ["PO Summary", "", "", "", "", "", `Generated: ${new Date().toLocaleDateString()}`],
    [],
    ["Vendor", "Code", "Product", "PLU", "Item #", "Case", "Order Qty", "Case Order", "Unit Cost", "Total Cost", "Stock", "SV/Day", "Days Supply", "Min", "Max", "Override"],
    ...rows.map((r) => [r.vendor, r.code, r.product, r.plu, r.itemNumber, r.caseSize, r.orderQty, r.caseOrderQty, r.unitCost, r.totalCost, r.currentStock, r.svPerDay, r.daysSupply, r.reorderMin, r.reorderMax, r.isOverridden]),
    [],
    ["", "", "", "", "", "", `TOTAL`, "", rows.reduce((s, r) => s + r.totalCost, 0)],
  ];
  const wsSummary = xlsx.utils.aoa_to_sheet(summaryData);
  wsSummary["!cols"] = [14,12,36,10,12,6,9,10,10,8,8,10,6,6,8].map((w) => ({ wch: w }));
  xlsx.utils.book_append_sheet(wb, wsSummary, "All Items");

  // One sheet per vendor
  vendors.forEach((vendor) => {
    const vRows = rows.filter((r) => r.vendor === vendor);
    const vData = [
      [`Purchase Order â€” ${vendor}`, "", `Date: ${new Date().toLocaleDateString()}`],
      [],
    ["Code", "Product", "PLU", "Item #", "Case", "Order Qty", "Case Order", "Unit Cost", "Total Cost", "Stock", "SV/Day", "Override"],
    ...vRows.map((r) => [r.code, r.product, r.plu, r.itemNumber, r.caseSize, r.orderQty, r.caseOrderQty, r.unitCost, r.totalCost, r.currentStock, r.svPerDay, r.isOverridden]),
      [],
      ["", "", "", "", "", "TOTAL", "", vRows.reduce((s, r) => s + r.totalCost, 0)],
    ];
    const ws = xlsx.utils.aoa_to_sheet(vData);
    ws["!cols"] = [12,34,10,12,6,9,10,10,8,8,8].map((w) => ({ wch: w }));
    // Safe sheet name â€” Excel limits to 31 chars, no special chars
    const sheetName = vendor.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Vendor";
    xlsx.utils.book_append_sheet(wb, ws, sheetName);
  });

  xlsx.writeFile(wb, `PO_${new Date().toISOString().slice(0, 10)}.xlsx`);
  showToast(`PO exported â€” ${rows.length} items across ${vendors.length} vendor(s)`);
}

function exportPoPdf() {
  const rows = buildPoRows();
  if (!rows.length) { showToast("No items to order."); return; }

  const vendors = [...new Set(rows.map((r) => r.vendor))];
  const dateStr = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const grandTotal = rows.reduce((s, r) => s + r.totalCost, 0);

  let html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <title>Purchase Order ${dateStr}</title>
  <style>
    body { font-family: Arial, sans-serif; font-size: 11px; color: #1c2320; margin: 0; padding: 24px; }
    h1 { font-size: 18px; margin: 0 0 4px; } 
    .meta { color: #66716d; margin-bottom: 20px; font-size: 11px; }
    h2 { font-size: 13px; margin: 24px 0 6px; padding: 4px 8px; background: #1c2320; color: #fff; border-radius: 4px; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 8px; }
    th { background: #eef7f0; text-align: left; padding: 5px 6px; font-size: 10px; text-transform: uppercase; border-bottom: 2px solid #dce3df; }
    td { padding: 4px 6px; border-bottom: 1px solid #dce3df; vertical-align: top; }
    tr:nth-child(even) td { background: #fafafa; }
    .num { text-align: right; }
    .total-row td { font-weight: 900; background: #eef7f0; border-top: 2px solid #16835b; }
    .override { color: #d79b25; font-weight: 700; }
    .grand { font-size: 13px; font-weight: 900; color: #16835b; text-align: right; margin-top: 16px; }
    .lock { font-size: 10px; }
    @media print { body { padding: 8px; } h2 { break-before: auto; } }
  </style></head><body>
  <h1>Purchase Order</h1>
  <div class="meta">Generated ${dateStr} &nbsp;Â·&nbsp; ${rows.length} items &nbsp;Â·&nbsp; ${vendors.length} vendor(s)</div>`;

  vendors.forEach((vendor) => {
    const vRows = rows.filter((r) => r.vendor === vendor);
    const vTotal = vRows.reduce((s, r) => s + r.totalCost, 0);
    html += `<h2>${escapeHtml(vendor)}</h2>
    <table>
      <thead><tr>
        <th>Code</th><th>Product</th><th>PLU</th><th>Case</th>
        <th class="num">Stock</th><th class="num">SV/Day</th>
        <th class="num">Min</th><th class="num">Max</th>
        <th class="num">Order Qty</th><th class="num">Case Order</th><th class="num">Unit Cost</th><th class="num">Total</th>
      </tr></thead><tbody>
      ${vRows.map((r) => `<tr>
        <td>${escapeHtml(r.code)}</td>
        <td>${escapeHtml(r.product)}${r.isOverridden === "Yes" ? ' <span class="lock override" title="Manual override">ðŸ”’</span>' : ""}</td>
        <td>${escapeHtml(r.plu)}</td>
        <td class="num">${r.caseSize}</td>
        <td class="num">${r.currentStock}</td>
        <td class="num">${r.svPerDay}</td>
        <td class="num">${r.reorderMin}</td>
        <td class="num">${r.reorderMax}</td>
        <td class="num"><b>${r.orderQty}</b></td>
        <td class="num"><b>${r.caseOrderQty}</b></td>
        <td class="num">${currency.format(r.unitCost)}</td>
        <td class="num">${currency.format(r.totalCost)}</td>
      </tr>`).join("")}
      <tr class="total-row"><td colspan="10">Vendor Total</td><td></td><td class="num">${currency.format(vTotal)}</td></tr>
      </tbody></table>`;
  });

  html += `<div class="grand">Grand Total: ${currency.format(grandTotal)}</div>
  </body></html>`;

  const win = window.open("", "_blank", "width=900,height=700");
  if (!win) { showToast("Pop-up blocked — allow pop-ups first.", 3000, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 400);
  // After PDF review, offer to submit POs
  setTimeout(() => {
    if (confirm("PO PDF opened for review.\n\nSubmit POs now? Items will be marked Pending for 10 days.")) {
      const allVendors = [...new Set(currentOrderRows().map((r) => r.vendor).filter(Boolean))];
      allVendors.forEach((v) => submitVendorPo(v));
      showToast(`PO submitted for ${allVendors.length} vendor${allVendors.length > 1 ? "s" : ""}`, 3000, "success");
      renderOrders();
    }
  }, 1800);
}

function dateFromFileName(fileName) {
  const match = fileName.match(/(\d{2})(\d{2})(\d{4})/);
  if (!match) return null;
  const [, mm, dd, yyyy] = match;
  return { iso: `${yyyy}-${mm}-${dd}`, compact: `${mm}${dd}${yyyy}` };
}

function fileSignature(file) {
  return [file.name || "", file.size || 0, file.lastModified || 0].join("|");
}

function switchTab(tab) {
  saveActiveTabSearch();
  if (["dashboard", "inventory", "ordering"].includes(tab) && els.searchInput) {
    els.searchInput.value = state.tabSearches[tab] || "";
  }
  document.querySelectorAll("[data-tab]").forEach((button) => button.classList.toggle("active", button.dataset.tab === tab));
  document.querySelectorAll("[data-tab-view]").forEach((view) => view.classList.toggle("active", view.dataset.tabView === tab));
  queueActiveTabRender();
}

function clearFilters() {
  state._pinnedAdjustCode = null;
  els.searchInput.value = "";
  saveActiveTabSearch();
  if (els.parentsSearch) els.parentsSearch.value = "";
  els.departmentFilter.value = "";
  els.categoryFilter.value = "";
  els.vendorFilter.value = "";
  els.colorFilter.value = "";
  els.startDate.value = state.dates[0] || "";
  els.endDate.value = state.dates[state.dates.length - 1] || "";
  els.inventoryStateFilter.value = "";
  state.activePresetDays = "all";
  render();
}

function clearSingleFilter(filterId) {
  const target = els[filterId];
  if (!target) return;
  target.value = "";
  target.dispatchEvent(new Event("input", { bubbles: true }));
}

function refreshDetailDrawer() {
  if (els.detailDrawer.hidden || !state._activeDetailCode) return;
  const item = findCurrentItemByCode(state._activeDetailCode);
  if (item) showDetail(item);
}

function showDetail(item) {
  if (!item) return;
  state._activeDetailCode = item.code;
  const sales = buildSkuRows({ ignoreQuery: true, ignoreFilters: true }).find((sku) => codeKey(sku.code) === codeKey(item.code)) || item;
  const windows = salesWindowsFor(item.code);
  const fields = [
    ["code", "Code", "ids", `<b class="copy-code" title="Click to copy">${escapeHtml(item.code)}</b>`],
    ["plu", "PLU", "ids", `<b>${escapeHtml(item.plu || sales.plu || "-")}</b>`],
    ["itemNumber", "Item #", "ids", `<b>${escapeHtml(item.itemNumber || sales.itemNumber || "-")}</b>`],
    ["parent", "Parent", "ids", `<b>${escapeHtml(item.parent || sales.parent || "-")}</b>`],
    ["sizeAttr", "Sub", "ids", `<b>${escapeHtml(item.subGroup || item.sizeAttr || sales.subGroup || sales.sizeAttr || "-")}</b>`],
    ["subType", "Type", "ids", `<b>${escapeHtml(item.typeGroup || item.subType || sales.typeGroup || sales.subType || "-")}</b>`],
    ["containerAttr", "Tag", "ids", `<b>${escapeHtml(item.containerAttr || sales.containerAttr || item.otherAttrs || sales.otherAttrs || "-")}</b>`],
    ["qty", "Qty sold in view", "sales", `<b>${number.format(sales.units || item.units || 0)}</b>`],
    ["windows", "Qty windows", "sales", `<b>${windows.map((entry) => `${entry.label}: ${number.format(entry.units)}`).join(" | ")}</b>`],
    ["sales", "Sold total price", "sales", `<b>${currency.format(sales.sales || 0)}</b>`],
    ["costSold", "Sold total cost", "sales", `<b>${currency.format(sales.costSold || ((sales.units || item.units || 0) * (item.unitCost || sales.unitCost || 0)))}</b>`],
    ["velocity", "Sales velocity", "sales", `<b>${formatVelocity(sales.velocity || item.velocity || 0)} / day</b>`],
    ["vendor", "Vendor", "inventory", `<b>${escapeHtml(item.vendor || sales.vendor || "-")}</b>`],
    ["category", "Category", "inventory", `<b>${escapeHtml(item.category || sales.category || "-")}</b>`],
    ["color", "Color", "inventory", `<b>${escapeHtml(item.color || sales.color || "-")}</b>`],
    ["state", "State", "inventory", `<b>${escapeHtml(item.state || sales.state || "-")}</b>`],
    ["addDate", "Add date", "inventory", `<b>${escapeHtml(item.addDate || sales.addDate || "-")}</b>`],
    ["stock", "Stock", "inventory", `<b>${number.format(item.stock ?? sales.stock ?? 0)}</b>`],
    ["unitCost", "Cost", "inventory", `<b>${currency.format(item.unitCost || sales.unitCost || 0)}</b>`],
    ["price", "Price", "inventory", `<b>${currency.format(item.price || sales.price || 0)}</b>`],
    ["inventoryCost", "Stock cost total", "inventory", `<b>${currency.format(item.inventoryCost || sales.inventoryCost || 0)}</b>`],
    ["caseSize", "Case size", "ordering", `<b>${number.format(item.caseSize || sales.caseSize || 1)}</b>`],
    ["minMax", "Min / Max", "ordering", `<b>${number.format(item.reorderMin || sales.reorderMin || 0)} / ${number.format(item.reorderMax || sales.reorderMax || 0)}</b>`],
    ["recommendedOrder", "Recommended order", "ordering", `<b>${number.format(sales.recommendedOrder || item.recommendedOrder || item.qtyNeeded || 0)}</b>`],
    ["caseOrder", "Case order", "ordering", `<b>${number.format(item.caseOrder || 0)}</b>`],
  ];
  const fieldKeys = fields.map(([key]) => key);
  const detailOrder = (state.detailOrder || fieldKeys).filter((key) => fieldKeys.includes(key));
  fieldKeys.forEach((key) => {
    if (!detailOrder.includes(key)) detailOrder.push(key);
  });
  const visibleFields = new Set((state.detailFilters || fieldKeys).filter((key) => fieldKeys.includes(key)));
  fieldKeys.forEach((key) => {
    if (!visibleFields.size) visibleFields.add(key);
  });
  const sortedFields = [...fields].sort((a, b) => detailOrder.indexOf(a[0]) - detailOrder.indexOf(b[0]));
  els.detailDrawer.hidden = false;
  els.detailDrawer.innerHTML = `
    <button class="drawer-close" type="button">Close</button>
    <button class="sort-details-button" type="button">Sort details</button>
    <p class="eyebrow">Item detail</p>
    <h2>${escapeHtml(item.product || sales.product)}</h2>
    <details class="detail-picker">
      <summary>Detail fields</summary>
      <div class="detail-filter">${fields.map(([key, label]) => `
        <label><input type="checkbox" ${visibleFields.has(key) ? "checked" : ""} data-detail-filter="${key}" />${label}</label>`).join("")}</div>
    </details>
    <div class="detail-grid">
      ${sortedFields.map(([key, label, section, value]) => `<span draggable="false" data-detail-key="${key}" data-detail-section="${section}" style="${visibleFields.has(key) ? "" : "display:none;"}">${label} ${value}</span>`).join("")}
    </div>`;
  els.detailDrawer.querySelector(".drawer-close").addEventListener("click", () => {
    els.detailDrawer.hidden = true;
    state._activeDetailCode = "";
  });
  els.detailDrawer.querySelector(".copy-code").addEventListener("click", (event) => copyText(item.code, event.currentTarget));
  els.detailDrawer.querySelectorAll("[data-detail-filter]").forEach((input) => {
    input.addEventListener("change", () => {
      const visible = new Set([...els.detailDrawer.querySelectorAll("[data-detail-filter]:checked")].map((item) => item.dataset.detailFilter));
      state.detailFilters = [...visible];
      localStorage.setItem("posDashboardDetailFilters:v1", JSON.stringify(state.detailFilters));
      els.detailDrawer.querySelectorAll("[data-detail-key]").forEach((node) => {
        const show = visible.has(node.dataset.detailKey);
        node.hidden = !show;
        node.style.display = show ? "" : "none";
      });
    });
  });
  els.detailDrawer.querySelector(".sort-details-button").addEventListener("click", () => {
    els.detailDrawer.classList.toggle("detail-sort-mode");
    const active = els.detailDrawer.classList.contains("detail-sort-mode");
    els.detailDrawer.querySelector(".sort-details-button").textContent = active ? "Done sorting" : "Sort details";
    els.detailDrawer.querySelectorAll("[data-detail-key]").forEach((tile) => {
      tile.draggable = active;
      tile.addEventListener("dragstart", (event) => event.dataTransfer.setData("text/plain", tile.dataset.detailKey));
      tile.addEventListener("dragover", (event) => active && event.preventDefault());
      tile.addEventListener("drop", (event) => {
        if (!active) return;
        event.preventDefault();
        moveDetailField(event.dataTransfer.getData("text/plain"), tile.dataset.detailKey);
        showDetail(item);
        setTimeout(() => els.detailDrawer.querySelector(".sort-details-button")?.click(), 0);
      });
    });
  });
}

function parentNameFor(item) {
  return parentPartsFor(item).parent;
}

function parentheticalBrandMatch(base) {
  // Matches patterns like "HH YAKI WVG 10 (EMPIRE) :: 1B" or "SOME PRODUCT (FAYGO) FLAVOR"
  // Checks if any text inside parentheses matches a known parent rule alias or brand
  const parenMatch = base.match(/\(([^)]+)\)/);
  if (!parenMatch) return null;
  const brandCandidate = parenMatch[1].trim();
  // Check against user-defined parent rules first
  const ruleResult = parentRuleParts(brandCandidate);
  if (ruleResult) {
    const withoutParen = normalizeProductName(base.replace(/\s*\([^)]+\)\s*/g, " "));
    return { parent: ruleResult.parent, subType: withoutParen };
  }
  // Check against built-in known brands (uppercase match)
  const knownBrands = ["EMPIRE", "FAYGO", "LITTLE DEBBIE", "HERSHEYS", "HERSHEY", "REESES", "SNICKERS"];
  const upperCandidate = brandCandidate.toUpperCase();
  const matched = knownBrands.find((b) => upperCandidate === b || upperCandidate.startsWith(b));
  if (!matched) return null;
  const withoutParen = normalizeProductName(base.replace(/\s*\([^)]+\)\s*/g, " "));
  return { parent: matched, subType: withoutParen };
}

function parentPartsFor(item) {
  const cacheKey = [
    cleanCell(item.code || ""),
    cleanCell(item.product || item.itemName || ""),
    cleanCell(item.color || ""),
    cleanCell(item.category || ""),
    cleanCell(item.department || ""),
  ].join("|");
  if (state._parentPartsCache.has(cacheKey)) return state._parentPartsCache.get(cacheKey);
  const base = normalizeProductName(cleanCell(item.product || item.itemName || ""));
  const color = cleanCell(item.color || "");
  const _isHair = isHairCategory(item);
  let parent = base;
  let subType = color;
  if (parent.includes("::")) {
    const index = parent.lastIndexOf("::");
    const leftPart = normalizeProductName(parent.slice(0, index));
    const rightPart = normalizeProductName(parent.slice(index + 2)) || subType;
    const aliasParent = hairParentAlias(leftPart);
    const parsed = applyAttributeRules({
      parent: aliasParent || leftPart || item.code,
      subType: rightPart,
      sizeAttr: _isHair ? extractHairLength(leftPart)?.length || "" : "",
      _isHair,
    });
    state._parentPartsCache.set(cacheKey, parsed);
    return parsed;
  }
  const rule = parentRuleParts(base);
  if (rule) {
    const parsed = applyAttributeRules({ ...rule, _isHair });
    state._parentPartsCache.set(cacheKey, parsed);
    return parsed;
  }
  // Extract parenthetical brand names before the :: split so (EMPIRE) isn't buried in the parent
  const parenBrand = parentheticalBrandMatch(base);
  if (parenBrand) {
    const parsed = applyAttributeRules({ ...parenBrand, _isHair });
    state._parentPartsCache.set(cacheKey, parsed);
    return parsed;
  }
  const hyphenIndex = parent.lastIndexOf(" - ");
  if (hyphenIndex > -1) {
    const right = normalizeProductName(parent.slice(hyphenIndex + 3));
    if (right && (_isHair || looksLikeColorToken(right) || looksLikeVariantText(right))) {
      const parsed = applyAttributeRules({
        parent: normalizeProductName(parent.slice(0, hyphenIndex)),
        subType: right,
        _isHair,
      });
      state._parentPartsCache.set(cacheKey, parsed);
      return parsed;
    }
  }
  const grocery = groceryParentParts(base, item);
  if (grocery) {
    const parsed = applyAttributeRules({ ...grocery, _isHair });
    state._parentPartsCache.set(cacheKey, parsed);
    return parsed;
  }
  if (color && parent.toLowerCase().endsWith(color.toLowerCase())) {
    parent = parent.slice(0, parent.length - color.length);
    subType = subType || color;
  }
  parent = parent.replace(/\s*[-_/]\s*(TT\/)?[A-Z0-9#.\/]{1,14}$/i, (match) => {
      if (looksLikeColorToken(match)) {
        subType = subType || cleanCell(match).replace(/^[-_/\s]+/, "");
        return "";
      }
      return match;
    })
    .replace(/\s+(1|1B|1\/B|2|4|27|30|33|99J|130|350|530|613|BUG|BG|BL|RED|BLUE|GREEN|PURPLE|PINK|GREY|GRAY|WHITE|BLACK|BROWN|BALAYAGE|SUNSET|H1B\/F1B|F1B\/30|F1B\/27)$/i, (match) => {
      subType = subType || cleanCell(match);
      return "";
    })
    .replace(/\s{2,}/g, " ")
    .trim();
  const parsed = applyAttributeRules({ parent: parent || base || item.code, subType: subType || "", _isHair });
  state._parentPartsCache.set(cacheKey, parsed);
  return parsed;
}

function normalizeProductName(value) {
  return value.replace(/\s{2,}/g, " ").replace(/\s+-\s*$/, "").trim();
}

function looksLikeColorToken(value) {
  const token = cleanCell(value).replace(/^[-_/\s]+/, "").toUpperCase();
  return /^(TT\/)?([0-9]{1,3}|1B|1\/B|99J|BUG|BG|BL|F1B|H1B|H1B\/F1B|F1B\/[0-9]{1,3}|OM[A-Z0-9]+|BALAYAGE|SUNSET|LEMON|COPPER|RED|BLUE|GREEN|PURPLE|PINK|GREY|GRAY|WHITE|BLACK|BROWN|BLONDE|AUBURN|BURGUNDY)([\/#-][A-Z0-9]+)?$/.test(token);
}

function isHairCategory(item) {
  return /\b(BRAID|BRAIDING|HAIR|WIG|CROCHET|WEAVE|HUMAN HAIR|SYNTHETIC)\b/i.test(`${item.category || ""} ${item.department || ""}`);
}

// Hair-specific: extract length token (10,12,14,16,18,20,22,24,26,28,30) from subType
// Returns { length: "14", remainder: "1B" } or null
function extractHairLength(subType) {
  // Match standalone length numbers 8â€“30 (even), possibly prefixed with space or parenthesis
  const match = subType.match(/(?:^|\s|\()(\b(?:8|10|12|14|16|18|20|22|24|26|28|30)\b)(?:["']?\s*(?:INCH|IN\b|")?)/i);
  if (!match) return null;
  const len = match[1];
  const remainder = subType.replace(match[0], " ").replace(/\s{2,}/g, " ").trim();
  return { length: `${len}"`, remainder };
}

function hairParentAlias(value) {
  const cleanValue = normalizeProductName(value);
  const rule = parentRuleParts(cleanValue);
  if (rule?.parent) return rule.parent;
  const parenMatch = cleanValue.match(/\(([^)]+)\)/);
  if (parenMatch) {
    const aliasRule = parentRuleParts(parenMatch[1]);
    if (aliasRule?.parent) return aliasRule.parent;
    return normalizeProductName(parenMatch[1]);
  }
  return "";
}

function applyAttributeRules(parts) {
  let subType = normalizeProductName(parts.subType || "");
  const found = { size: cleanCell(parts.sizeAttr || ""), container: "", other: [] };

  // Hair-department auto-length extraction: promote numeric length to sizeAttr
  if (parts._isHair && !found.size) {
    const hair = extractHairLength(subType);
    if (hair) {
      found.size = hair.length;
      subType = hair.remainder;
    }
  }

  if (ENABLE_CUSTOM_ATTRIBUTE_RULES) {
    state.attributeRules.forEach((rule) => {
      if (!rule.value || !(rule.aliases || []).length) return;
      const matchedAlias = rule.aliases.find((alias) => aliasMatches(subType, alias));
      if (!matchedAlias) return;
      subType = removeAliasText(subType, matchedAlias);
      if (rule.type === "size") found.size = rule.value;
      else if (rule.type === "container") found.container = rule.value;
      else found.other.push(rule.value);
    });
  }
  return {
    ...parts,
    subType: normalizeProductName(subType),
    sizeAttr: found.size,
    containerAttr: found.container,
    otherAttrs: found.other.join(", "),
  };
}

function looksLikeVariantText(value) {
  return /\b(CHOCOLATE|ALMOND|ALMONDS|COOKIE|CREAM|CREME|CARAMEL|PEANUT|ROLL|ROLLS|BAR|BARS|KING|REGULAR|MINI|MILK|DARK|WHITE|ORIGINAL|HOT|BBQ|RANCH|CHEDDAR|SOUR)\b/i.test(value);
}

function groceryParentParts(base, item) {
  const category = cleanCell(item.category || "").toUpperCase();
  if (!/\b(CANDY|CANDIES|GROCERY|SNACK|SNACKS|DRINK|DRINKS|FOOD|BEVERAGE|BEVERAGES|SODA|SODAS)\b/.test(category)) return null;
  const cleaned = normalizeProductName(base);
  const brands = ["LITTLE DEBBIE", "KIT KAT", "M&M", "HERSHEYS", "HERSHEY", "REESES", "SNICKERS", "TWIX", "SKITTLES", "HARIBO", "HOSTESS", "FAYGO"];
  const upper = cleaned.toUpperCase();
  const brand = brands.find((name) => upper === name || upper.startsWith(`${name} `) || upper.startsWith(`${name}-`));
  if (!brand) return null;
  const remainder = normalizeProductName(cleaned.slice(brand.length));
  return {
    parent: brand,
    subType: remainder || "",
  };
}

function applyAttributeRulesLegacy_unused(parts) {
  let subType = normalizeProductName(parts.subType || "");
  const found = { size: "", container: "", other: [] };
  state.attributeRules.forEach((rule) => {
    if (!rule.value || !(rule.aliases || []).length) return;
    const matchedAlias = rule.aliases.find((alias) => aliasMatches(subType, alias));
    if (!matchedAlias) return;
    subType = removeAliasText(subType, matchedAlias);
    if (rule.type === "size") found.size = rule.value;
    else if (rule.type === "container") found.container = rule.value;
    else found.other.push(rule.value);
  });
  return {
    ...parts,
    subType: normalizeProductName(subType),
    sizeAttr: found.size,
    containerAttr: found.container,
    otherAttrs: found.other.join(", "),
  };
}

function aliasMatches(value, alias) {
  const normalizedValue = compactMatch(value);
  const normalizedAlias = compactMatch(alias);
  return normalizedAlias && normalizedValue.includes(normalizedAlias);
}

function compactMatch(value) {
  return normalizeForMatch(value).replace(/[^A-Z0-9]/g, "");
}

function removeAliasText(value, alias) {
  const candidates = unique([alias, alias.replace(/^#/, ""), alias.replace(/^#/, "").replace(/([0-9])([A-Z])/gi, "$1 $2")]);
  let result = value;
  candidates.forEach((candidate) => {
    const clean = cleanCell(candidate);
    if (!clean) return;
    result = result.replace(new RegExp(`(^|\\s|-)#?${escapeRegex(clean)}(?=$|\\s|-)`, "ig"), " ");
  });
  return normalizeProductName(result.replace(/\s{2,}/g, " "));
}

function childLabel(child, options = {}) {
  return [child.typeGroup || child.subType || child.color || child.itemNumber || child.code, options.omitSize ? "" : (child.subGroup || child.sizeAttr), child.containerAttr, child.otherAttrs]
    .filter(Boolean)
    .join(" - ");
}

function parentAttributeHtml(group) {
  const chips = []
    .concat(attributeCountChips(group.children, "sizeAttr", "sizeAttr", "Size"))
    .concat(attributeCountChips(group.children, "containerAttr", "containerAttr", "Container"))
    .concat(attributeCountChips(group.children, "otherAttrs", "otherAttrs", "Tag"));
  if (!chips.length) return "";
  return `
    <div class="parent-attribute-row">
      <button type="button" data-parent-attr="all">All sub/types</button>
      ${chips.join("")}
    </div>`;
}

function attributeCountChips(children, field, dataKey, label) {
  const counts = new Map();
  children.forEach((child) => {
    const value = cleanCell(child[field]);
    if (value) counts.set(value, (counts.get(value) || 0) + 1);
  });
  return [...counts.entries()].sort((a, b) => a[0].localeCompare(b[0])).map(([value, count]) => `
    <button type="button" data-parent-attr="${escapeHtml(dataKey)}" data-attr-value="${escapeHtml(value)}">
      ${escapeHtml(label)}: ${escapeHtml(value)} <b>${number.format(count)}</b>
    </button>`);
}

function parentRuleParts(base) {
  if (!ENABLE_CUSTOM_PARENT_RULES || !state.parentRules.length) return null;
  const normalized = normalizeForMatch(base);
  for (const rule of state.parentRules) {
    const alias = (rule.aliases || []).find((entry) => {
      const cleanAlias = normalizeForMatch(entry.replace(/^#/, ""));
      return cleanAlias && (normalized === cleanAlias || normalized.startsWith(`${cleanAlias} `));
    });
    if (!alias) continue;
    const cleanAlias = normalizeForMatch(alias.replace(/^#/, ""));
    const subType = base.slice(base.toUpperCase().indexOf(cleanAlias.toUpperCase()) + cleanAlias.length).replace(/^['S\s-]+/i, "").trim();
    return { parent: rule.parent, subType: normalizeProductName(subType) };
  }
  return null;
}

function addAttributeRule() {
  const type = els.attributeRuleType.value || "other";
  const value = cleanCell(els.attributeRuleValue.value).toUpperCase();
  const aliases = els.attributeRuleAliases.value.split(",").map((item) => cleanCell(item)).filter(Boolean);
  if (!value || !aliases.length) return;
  const id = `${type}:${value}`;
  state.attributeRules = state.attributeRules.filter((rule) => rule.id !== id).concat({ id, type, value, aliases });
  localStorage.setItem("posDashboardAttributeRules:v1", JSON.stringify(state.attributeRules));
  els.attributeRuleValue.value = "";
  els.attributeRuleAliases.value = "";
  render();
}

function removeAttributeRule(id) {
  state.attributeRules = state.attributeRules.filter((rule) => rule.id !== id);
  localStorage.setItem("posDashboardAttributeRules:v1", JSON.stringify(state.attributeRules));
  render();
}

function renderAttributeRules() {
  if (!els.attributeRuleList) return;
  els.attributeRuleCount.textContent = number.format(state.attributeRules.length);
  if (!state.attributeRules.length) {
    els.attributeRuleList.innerHTML = `<span class="muted">No sub/type attribute rules yet.</span>`;
    return;
  }
  els.attributeRuleList.innerHTML = state.attributeRules.map((rule) => `
    <button type="button" data-remove-attribute-rule="${escapeHtml(rule.id)}">
      <b>${escapeHtml(rule.value)}</b>
      <span>${escapeHtml((rule.aliases || []).join(", "))}</span>
      <i>Remove</i>
    </button>`).join("");
  els.attributeRuleList.querySelectorAll("[data-remove-attribute-rule]").forEach((button) => {
    button.addEventListener("click", () => removeAttributeRule(button.dataset.removeAttributeRule));
  });
}

function addParentRule() {
  const parent = cleanCell(els.parentRuleName.value).toUpperCase();
  const aliases = els.parentRuleAliases.value.split(",").map((item) => cleanCell(item)).filter(Boolean);
  if (!parent || !aliases.length) return;
  state.parentRules = state.parentRules.filter((rule) => rule.parent !== parent).concat({ parent, aliases });
  localStorage.setItem("posDashboardParentRules:v1", JSON.stringify(state.parentRules));
  els.parentRuleName.value = "";
  els.parentRuleAliases.value = "";
  render();
}

function removeParentRule(parent) {
  state.parentRules = state.parentRules.filter((rule) => rule.parent !== parent);
  localStorage.setItem("posDashboardParentRules:v1", JSON.stringify(state.parentRules));
  render();
}

function renderParentRules() {
  if (!els.parentRuleList) return;
  if (!state.parentRules.length) {
    els.parentRuleList.innerHTML = `<span class="muted">No app-only parent rules yet.</span>`;
    return;
  }
  els.parentRuleList.innerHTML = `
    <details class="compact-rule-list">
      <summary><span>${number.format(state.parentRules.length)}</span> parent alias rules</summary>
      <div class="rule-list-inner">
        ${state.parentRules.map((rule) => `
          <button type="button" data-remove-parent-rule="${escapeHtml(rule.parent)}">
            <b>${escapeHtml(rule.parent)}</b>
            <span>${escapeHtml((rule.aliases || []).join(", "))}</span>
            <i>Remove</i>
          </button>`).join("")}
      </div>
    </details>`;
  els.parentRuleList.querySelectorAll("[data-remove-parent-rule]").forEach((button) => {
    button.addEventListener("click", () => removeParentRule(button.dataset.removeParentRule));
  });
}

function normalizeForMatch(value) {
  return cleanCell(value).replace(/#/g, "").replace(/'/g, "").toUpperCase();
}

function roundToCase(quantity, caseSize) {
  const size = Math.max(1, toNumber(caseSize) || 1);
  return Math.ceil(quantity / size) * size;
}

function roundOrderToNearestCase(quantity, caseSize) {
  const size = Math.max(1, toNumber(caseSize) || 1);
  const qty = Math.max(0, Math.ceil(toNumber(quantity) || 0));
  if (!qty) return 0;
  return Math.max(size, Math.round(qty / size) * size);
}

function roundInventoryTargetToCase(stock, targetQty, caseSize) {
  const size = Math.max(1, toNumber(caseSize) || 1);
  const desiredTotal = Math.max(Math.ceil(toNumber(stock) || 0), Math.ceil(toNumber(targetQty) || 0));
  return Math.ceil(desiredTotal / size) * size;
}

function recommendedOrderQty({ stock, min, max }) {
  const currentStock = Math.max(0, toNumber(stock) || 0);
  const minQty = Math.max(0, Math.ceil(toNumber(min) || 0));
  const maxQty = Math.max(minQty, Math.ceil(toNumber(max) || 0));
  if (currentStock >= minQty) return 0;
  return Math.max(0, Math.ceil(maxQty - currentStock));
}

function renderColumnPicker() {
  els.columnPickerPanel.innerHTML = `<div class="column-picker-grid order-cp-grid">${inventoryColumns.map(([key, label]) => `
    <label class="column-choice">
      <input type="checkbox" data-column-toggle="${key}" ${state.visibleColumns[key] ? "checked" : ""} />
      <span>${label}</span>
    </label>`).join("")}</div>`;
  els.columnPickerPanel.querySelectorAll("[data-column-toggle]").forEach((input) => {
    input.addEventListener("change", () => {
      state.visibleColumns[input.dataset.columnToggle] = input.checked;
      localStorage.setItem("posDashboardVisibleColumns:v3", JSON.stringify(state.visibleColumns));
      applyColumnVisibility();
    });
  });
}

function applyColumnVisibility() {
  document.querySelectorAll("[data-col]").forEach((cell) => {
    cell.classList.toggle("hidden-column", state.visibleColumns[cell.dataset.col] === false);
  });
}

function updateSortHeaders() {
  document.querySelectorAll("[data-sort]").forEach((header) => {
    const active = header.dataset.sort === state.inventorySort.key;
    header.dataset.sortDir = active ? state.inventorySort.dir : "";
  });
}

function compareInventoryRows(a, b) {
  const { key, dir } = state.inventorySort;
  const av = a[key];
  const bv = b[key];
  const direction = dir === "asc" ? 1 : -1;
  if (typeof av === "number" || typeof bv === "number") return ((av || 0) - (bv || 0)) * direction;
  if (key === "addDate") return compareDateValue(String(av || ""), String(bv || "")) * direction;
  if (key === "product" || key === "subType" || key === "plu" || key === "itemNumber") {
    return compareVariantAwareRows(a, b) * direction;
  }
  return compareAlphaValue(String(av || ""), String(bv || "")) * direction;
}

function compareVariantAwareRows(a, b) {
  const aFamily = variantFamilyLabel(a);
  const bFamily = variantFamilyLabel(b);
  const family = compareAlphaValue(aFamily, bFamily);
  if (family !== 0) return family;
  const size = compareAlphaValue(String(a.subGroup || a.sizeAttr || ""), String(b.subGroup || b.sizeAttr || ""));
  if (size !== 0) return size;
  const colorRank = variantToneRank(a) - variantToneRank(b);
  if (colorRank !== 0) return colorRank;
  const type = compareAlphaValue(variantToneLabel(a), variantToneLabel(b));
  if (type !== 0) return type;
  return compareAlphaValue(String(a.product || a.code || ""), String(b.product || b.code || ""));
}

function variantSourceLabel(item) {
  return String(item.parent || item.product || item.typeGroup || item.subType || item.color || item.plu || item.code || "");
}

function variantTailToken(value) {
  return cleanCell(String(value || "")).split(/\s+/).at(-1) || "";
}

function variantFamilyLabel(item) {
  const source = cleanCell(variantSourceLabel(item));
  const tail = variantTailToken(source);
  if (!tail) return source;
  return colorSortRank(tail) < 100 ? cleanCell(source.slice(0, -tail.length)) : source;
}

function variantToneLabel(item) {
  const source = variantSourceLabel(item);
  const tail = variantTailToken(source);
  if (colorSortRank(tail) < 100) return tail;
  return String(item.typeGroup || item.subType || item.color || item.plu || tail || "");
}

function variantToneRank(item) {
  return colorSortRank(variantToneLabel(item));
}

function compareAlphaValue(a, b) {
  return String(a || "").localeCompare(String(b || ""), undefined, { numeric: true, sensitivity: "base" });
}

function compareDateValue(a, b) {
  return parseLooseDate(a) - parseLooseDate(b);
}

function parseLooseDate(value) {
  const clean = cleanCell(value);
  if (!clean) return 0;
  if (/^\d{4}-\d{2}-\d{2}$/.test(clean)) return Date.parse(`${clean}T00:00:00`) || 0;
  const mdY = clean.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mdY) {
    const [, mm, dd, yyyy] = mdY;
    return Date.parse(`${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}T00:00:00`) || 0;
  }
  return Date.parse(clean) || 0;
}

function compareSubTypeOrder(a, b) {
  return compareDisplayValue(
    String(a.subType || a.color || a.itemNumber || a.code || ""),
    String(b.subType || b.color || b.itemNumber || b.code || ""),
  );
}

function compareDisplayValue(a, b) {
  const ap = colorSortRank(a);
  const bp = colorSortRank(b);
  if (ap !== bp) return ap - bp;
  return String(a || "").localeCompare(String(b || ""), undefined, { numeric: true, sensitivity: "base" });
}

function colorSortRank(value) {
  const token = cleanCell(value).toUpperCase().replace(/\s+/g, "");
  if (token === "1") return 1;
  if (token === "1B" || token === "1/B") return 2;
  if (token === "2") return 3;
  if (token === "4") return 4;
  if (/(^|[^0-9])27([^0-9]|$)/.test(token) || token.includes("/27") || token.includes("OM27") || token.includes("T27")) return 5;
  if (/(^|[^0-9])30([^0-9]|$)/.test(token) || token.includes("/30") || token.includes("OM30") || token.includes("T30")) return 6;
  return 100;
}

function indexRowsByCodeKey(rows) {
  const index = new Map();
  rows.forEach((row) => {
    const key = codeKey(row.code);
    if (key && !index.has(key)) index.set(key, row);
  });
  return index;
}

function copyText(value, target) {
  if (!value) return;
  navigator.clipboard?.writeText(value);
  const previous = target.textContent;
  target.textContent = "Copied";
  setTimeout(() => { target.textContent = previous; }, 800);
}

function showToast(message, duration = 2800, tone = "info") {
  let toast = document.getElementById("posToast");
  if (!toast) {
    toast = document.createElement("div");
    toast.id = "posToast";
    toast.className = "pos-toast";
    document.body.append(toast);
  }
  toast.textContent = message;
  toast.dataset.tone = tone;
  toast.classList.add("visible");
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => toast.classList.remove("visible"), duration);
}

// Date preset helpers â€” sets start/end inputs and re-renders
function applyDatePreset(days) {
  if (!state.dates.length) return;
  // Data is always 1 day delayed (prior day import), so anchor end to yesterday
  const yesterday = new Date();
  yesterday.setDate(yesterday.getDate() - 1);
  const yesterdayIso = yesterday.toISOString().slice(0, 10);
  const lastDataDate = state.dates.at(-1) || yesterdayIso;
  // Use whichever is more recent: yesterday or last data date
  const end = lastDataDate > yesterdayIso ? lastDataDate : yesterdayIso;
  if (days === "ytd") {
    els.startDate.value = `${end.slice(0, 4)}-01-01`;
    els.endDate.value = end;
    render();
    return;
  }
  if (days === "all") {
    els.startDate.value = state.dates[0] || "";
    els.endDate.value = end;
    render();
    return;
  }
  const endMs = new Date(`${end}T00:00:00`).getTime();
  // days=1 means just yesterday; days=7 means yesterday back 6 more days
  const startMs = endMs - (days - 1) * 86400000;
  const startIso = new Date(startMs).toISOString().slice(0, 10);
  els.startDate.value = startIso;
  els.endDate.value = end;
  render();
}

// Navigate date range by period — arrows follow the active preset button
function shiftDateRange(direction) {
  const preset = state.activePresetDays;
  // Determine shift size from active preset
  let shiftMs = 86400000; // default: 1 day
  if (preset === 7)    shiftMs = 7 * 86400000;
  else if (preset === 30)   shiftMs = 30 * 86400000;
  else if (preset === 60)   shiftMs = 60 * 86400000;
  else if (preset === 90)   shiftMs = 90 * 86400000;
  else if (preset === 183)  shiftMs = 183 * 86400000;
  else if (preset === 365)  shiftMs = 365 * 86400000;
  else if (preset === 1)    shiftMs = 86400000;
  const start = new Date(`${els.startDate.value || state.dates[0]}T00:00:00`);
  const end   = new Date(`${els.endDate.value   || state.dates.at(-1)}T00:00:00`);
  const delta = direction * shiftMs;
  els.startDate.value = new Date(start.getTime() + delta).toISOString().slice(0, 10);
  els.endDate.value   = new Date(end.getTime()   + delta).toISOString().slice(0, 10);
  render();
}

function renderDatePresets() {
  const container = document.getElementById("datePresets");
  if (!container) return;
  const presets = [
    { label: "1D", days: 1 },
    { label: "7D", days: 7 },
    { label: "30D", days: 30 },
    { label: "60D", days: 60 },
    { label: "90D", days: 90 },
    { label: "6M", days: 183 },
    { label: "1Y", days: 365 },
    { label: "YTD", days: "ytd" },
    { label: "ALL", days: "all" },
  ];
  if (!state._datePresetsReady) {
    container.innerHTML = presets.map((p) =>
      `<button type="button" class="preset-chip" data-preset-days="${p.days}">${p.label}</button>`
    ).join("");
    container.querySelectorAll("[data-preset-days]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const presetValue = btn.dataset.presetDays === "all"
          ? "all"
          : btn.dataset.presetDays === "ytd"
            ? "ytd"
            : Number(btn.dataset.presetDays);
        state.activePresetDays = presetValue;
        if (btn.dataset.presetDays === "all") {
          els.startDate.value = state.dates[0] || "";
          els.endDate.value = state.dates[state.dates.length - 1] || "";
          render();
        } else {
          applyDatePreset(btn.dataset.presetDays === "ytd" ? "ytd" : Number(btn.dataset.presetDays));
        }
      });
    });
    state._datePresetsReady = true;
  }
  container.querySelectorAll(".preset-chip").forEach((chip) => {
    const value = chip.dataset.presetDays === "all"
      ? "all"
      : chip.dataset.presetDays === "ytd"
        ? "ytd"
        : Number(chip.dataset.presetDays);
    chip.classList.toggle("active", value === state.activePresetDays);
  });
}

function ytdDays() {
  const now = new Date();
  const jan1 = new Date(now.getFullYear(), 0, 1);
  return Math.ceil((now - jan1) / 86400000) + 1;
}

function bestItemName(...values) {
  const found = values.find((value) => cleanCell(value));
  return found ? cleanCell(found) : "Unnamed item";
}

function hasValue(value) {
  return value !== undefined && value !== null && value !== "";
}

function pickNumber(...values) {
  const found = values.find((value) => Number.isFinite(value) && value !== 0);
  return found || 0;
}

function orderingTargets({ velocity, safetyDays, daysOfInventory }) {
  // Targets represent true unit need. Case size only affects the final order qty.
  const rawMin = Math.max(0, velocity * Math.max(0, safetyDays || 0));
  const min = Math.max(0, Math.ceil(rawMin));
  const max = Math.max(min, Math.ceil(rawMin + (velocity * Math.max(0, daysOfInventory || 0))));
  return { min, max };
}

function ensureSalesIndexes() {
  if (state._salesIndex && state._salesIndexStamp === state._dataCacheStamp) return;
  const salesIndex = new Map();
  const dailyTotals = new Map();
  state.rawSales.forEach((row) => {
    const dateKey = row.date;
    const dayTotal = dailyTotals.get(dateKey) || { sales: 0, units: 0 };
    dayTotal.sales += row.sales || 0;
    dayTotal.units += row.units || 0;
    dailyTotals.set(dateKey, dayTotal);

    const itemKey = codeKey(row.code);
    if (!itemKey) return;
    const byDate = salesIndex.get(itemKey) || new Map();
    const totals = byDate.get(dateKey) || { sales: 0, units: 0 };
    totals.sales += row.sales || 0;
    totals.units += row.units || 0;
    byDate.set(dateKey, totals);
    salesIndex.set(itemKey, byDate);
  });
  state._salesIndex = salesIndex;
  state._dailyTotals = dailyTotals;
  state._salesWindowsCache = new Map();
  state._salesIndexStamp = state._dataCacheStamp;
}

function dailyTotalsFor(date) {
  ensureSalesIndexes();
  return state._dailyTotals.get(date) || { sales: 0, units: 0 };
}

function lowerBound(values, target) {
  let low = 0;
  let high = values.length;
  while (low < high) {
    const mid = Math.floor((low + high) / 2);
    if (values[mid] < target) low = mid + 1;
    else high = mid;
  }
  return low;
}

function salesWindowsFor(code) {
  const end = els.endDate.value || state.dates.at(-1);
  if (!end) return [];
  ensureSalesIndexes();
  const targetCode = codeKey(code);
  const cacheKey = `${targetCode}|${end}`;
  if (state._salesWindowsCache.has(cacheKey)) return state._salesWindowsCache.get(cacheKey);
  const endTime = new Date(`${end}T00:00:00`).getTime();
  const byDate = state._salesIndex.get(targetCode);
  const windows = [
    ["1D", 1],
    ["7D", 7],
    ["30D", 30],
    ["60D", 60],
    ["90D", 90],
    ["6M", 183],
    ["365D", 365],
  ];
  if (!byDate?.size) {
    const empty = windows.map(([label]) => ({ label, units: 0 }));
    state._salesWindowsCache.set(cacheKey, empty);
    return empty;
  }
  const series = [...byDate.entries()]
    .map(([date, totals]) => ({ time: new Date(`${date}T00:00:00`).getTime(), units: totals.units || 0 }))
    .sort((a, b) => a.time - b.time);
  const times = series.map((entry) => entry.time);
  const prefix = [];
  let runningUnits = 0;
  series.forEach((entry) => {
    runningUnits += entry.units;
    prefix.push(runningUnits);
  });
  const sumBetween = (startTime, finishTime) => {
    const startIndex = lowerBound(times, startTime);
    const endIndex = lowerBound(times, finishTime + 1) - 1;
    if (startIndex >= times.length || endIndex < startIndex) return 0;
    return prefix[endIndex] - (startIndex > 0 ? prefix[startIndex - 1] : 0);
  };
  const results = windows.map(([label, days]) => {
    const startTime = endTime - (days - 1) * 86400000;
    const units = sumBetween(startTime, endTime);
    return { label, units };
  });
  state._salesWindowsCache.set(cacheKey, results);
  return results;
}

function toNumber(value) {
  const cleaned = String(value ?? "").replace(/="/g, "").replace(/"/g, "").replace(/,/g, "").trim();
  const parsed = Number(cleaned);
  return Number.isFinite(parsed) ? parsed : 0;
}

function cleanHeader(value) {
  return String(value ?? "").trim();
}

function cleanCell(value) {
  return String(value ?? "").replace(/^="/, "").replace(/"$/, "").trim();
}

function normalizeCode(value) {
  return cleanCell(value).replace(/^=/, "").replace(/^"|"$/g, "");
}

function codeKey(value) {
  const code = normalizeCode(value);
  if (!code) return "";
  if (/^\d+$/.test(code)) return code.replace(/^0+/, "") || "0";
  return code.toUpperCase();
}

function bestLabel(current, next) {
  if (!current || current === "Unassigned" || current === "-") return next || current;
  return current;
}

function unique(values) {
  return [...new Set(values.map((value) => cleanCell(value)).filter(Boolean))];
}

function percent(part, whole) {
  if (!whole) return "0%";
  return `${Math.round((part / whole) * 100)}%`;
}

function finite(value) {
  return Number.isFinite(value) ? value : Number.MAX_SAFE_INTEGER;
}

function formatDays(value) {
  if (!Number.isFinite(value)) return "-";
  if (value < 0) return "0";
  return number.format(value);
}

function formatMetric(value, metric) {
  if (metric === "sales" || metric === "profit") return currency.format(value);
  return number.format(value);
}

function labelStatus(status) {
  return {
    grow: "Grow", watch: "Order", cut: "Cut", stockout: "Critical", steady: "Steady",
    discontinued: "Discontinued", disabled: "Disabled", forceorder: "Force Order",
  }[status] || "Steady";
}

function fileSummary() {
  const salesDays = state.dates.length;
  const inventoryDays = state.inventories.size;
  const excelItems = state.excelItems.size;
  if (!salesDays && !inventoryDays && !excelItems) return "No files loaded";
  return `${salesDays} sales days Â· ${inventoryDays} inventory snapshots Â· ${number.format(excelItems)} Excel items`;
}

function coverageSummary() {
  const inventoryDates = [...state.inventories.keys()].sort();
  if (!state.dates.length && !inventoryDates.length) return "Choose POS sales and inventory CSV exports.";
  const salesText = state.dates.length ? `Sales ${state.dates[0]} through ${state.dates[state.dates.length - 1]}` : "No sales days loaded";
  const inventoryText = inventoryDates.length ? `current inventory ${inventoryDates[inventoryDates.length - 1]}` : "no inventory snapshot";
  return `${salesText}; ${inventoryText}`;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function csvCell(value, key) {
  // Preserve leading zeros for code/UPC columns by formatting as Excel text
  if (key === "code" || key === "plu" || key === "itemNumber") {
    const str = String(value ?? "");
    if (/^\d+$/.test(str)) return `"=""${str}"""`;
    return `"${str.replace(/"/g, '""')}"`;
  }
  const clean = Number.isFinite(value) ? value : String(value ?? "");
  return `"${String(clean).replace(/"/g, '""')}"`;
}

function openDb() {
  return new Promise((resolve, reject) => {
    const request = indexedDB.open(DB_NAME, DB_VERSION);
    request.onupgradeneeded = () => request.result.createObjectStore(DB_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

async function readPersistedState() {
  const db = await openDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readonly");
    const request = tx.objectStore(DB_STORE).get(DB_KEY);
    request.onsuccess = () => resolve(request.result || null);
    request.onerror = () => reject(request.error);
  });
}

async function savePersistedState() {
  const db = await openDb();
  const payload = {
    rawSales: state.rawSales,
    inventoryDate: latestInventoryDate(),
    inventoryRows: [...state.latestInventory.values()],
    excelRows: [...state.excelItems.values()],
    loadedFileSignatures: [...state._loadedFileSignatures],
  };
  return new Promise((resolve, reject) => {
    const tx = db.transaction(DB_STORE, "readwrite");
    tx.objectStore(DB_STORE).put(payload, DB_KEY);
    tx.oncomplete = () => resolve();
    tx.onerror = () => reject(tx.error);
  });
}

async function restorePersistedState() {
  try {
    const saved = await readPersistedState();
    if (!saved) return;
    state.rawSales = saved.rawSales || [];
    state.dates = [...new Set(state.rawSales.map((row) => row.date))].sort();
    state.inventories = new Map();
    if (saved.inventoryDate && saved.inventoryRows?.length) {
      state.inventories.set(saved.inventoryDate, saved.inventoryRows);
    }
    buildLatestInventory();
    state.excelItems = new Map();
    (saved.excelRows || []).forEach((item) => addExcelIndex(item));
    rebuildExcelIndexes();
    state._loadedFileSignatures = new Set(saved.loadedFileSignatures || []);
    bumpDataStamp();
    updateFilterOptions();
    updateInventoryStateFilter();
    setDefaultDates();
  } catch (error) {
    console.warn("Could not restore POS dashboard history", error);
  }
}

function supabaseRestUrl(tableName) {
  return `${SUPABASE_URL}/rest/v1/${tableName}`;
}

async function supabaseSelectRows(tableName, query = {}) {
  const pageSize = 1000;
  const rows = [];
  let offset = 0;
  while (true) {
    const url = new URL(supabaseRestUrl(tableName));
    Object.entries(query).forEach(([key, value]) => {
      if (value !== undefined && value !== null && value !== "") url.searchParams.set(key, value);
    });
    url.searchParams.set("limit", String(pageSize));
    url.searchParams.set("offset", String(offset));
    const response = await fetch(url.toString(), { headers: supabaseHeaders() });
    if (!response.ok) throw new Error(`${tableName} read failed (${response.status})`);
    const page = await response.json();
    rows.push(...page);
    if (page.length < pageSize) break;
    offset += pageSize;
  }
  return rows;
}

async function supabaseDeleteAllRows(tableName) {
  const url = new URL(supabaseRestUrl(tableName));
  url.searchParams.set("id", "not.is.null");
  const response = await fetch(url.toString(), {
    method: "DELETE",
    headers: supabaseHeaders(),
  });
  if (!response.ok) throw new Error(`${tableName} delete failed (${response.status})`);
}

async function supabaseInsertRows(tableName, rows) {
  if (!rows.length) return;
  const chunkSize = 500;
  for (let i = 0; i < rows.length; i += chunkSize) {
    const response = await fetch(supabaseRestUrl(tableName), {
      method: "POST",
      headers: supabaseHeaders(true),
      body: JSON.stringify(rows.slice(i, i + chunkSize)),
    });
    if (!response.ok) throw new Error(`${tableName} insert failed (${response.status})`);
  }
}

function productRowForSupabase(item) {
  const excel = findExcelFor(item);
  return {
    code: item.code || excel.code || "",
    product: item.product || excel.product || "",
    plu: item.plu || excel.plu || "",
    item_number: item.itemNumber || excel.itemNumber || "",
    vendor: item.vendor || excel.vendor || "",
    category: item.category || excel.category || "",
    department: item.department || excel.department || "",
    color: item.color || excel.color || "",
    state: excel.state || item.state || "",
    stock: Number(item.stock || excel.stock || 0),
    price: Number(item.price || excel.price || 0),
    unit_cost: Number(item.cost || excel.cost || 0),
    case_size: Number(excel.caseSize || 1),
    add_date: excel.addDate || null,
    updated_at: new Date().toISOString(),
  };
}

function salesRowForSupabase(row) {
  return {
    sales_date: row.date,
    code: row.code || "",
    product: row.product || "",
    vendor: row.vendor || "",
    category: row.category || "",
    department: row.department || "",
    qty: Number(row.units || 0),
    sales: Number(row.sales || 0),
    cost_sold: Number(row.cost || 0),
    profit: Number(row.profit || 0),
    source_file: row.date || "",
  };
}

function hydrateInventoryFromSupabase(row) {
  return {
    date: row.updated_at ? String(row.updated_at).slice(0, 10) : new Date().toISOString().slice(0, 10),
    code: normalizeCode(row.code),
    category: cleanCell(row.category),
    product: cleanCell(row.product),
    plu: cleanCell(row.plu),
    itemNumber: cleanCell(row.item_number),
    price: toNumber(row.price),
    cost: toNumber(row.unit_cost),
    stock: toNumber(row.stock),
    vendor: cleanCell(row.vendor),
    vendorCode: "",
    color: cleanCell(row.color),
    size: "",
    length: "",
    manufacture: "",
    memo: "",
    department: cleanCell(row.department),
  };
}

function hydrateExcelFromSupabase(row) {
  return normalizeExcelRow({
    code: row.code,
    item_name: row.product,
    vendor_name: row.vendor,
    PLU: row.plu,
    item_number: row.item_number,
    category: row.category,
    add_date: row.add_date,
    cost: row.unit_cost,
    price: row.price,
    case_size: row.case_size,
    stock: row.stock,
    state: row.state,
  });
}

function hydrateSalesFromSupabase(row) {
  return {
    date: row.sales_date,
    code: normalizeCode(row.code),
    product: cleanCell(row.product),
    department: cleanCell(row.department) || "Unassigned",
    category: cleanCell(row.category) || "Unassigned",
    vendor: cleanCell(row.vendor) || "Unassigned",
    units: toNumber(row.qty),
    sales: toNumber(row.sales),
    cost: toNumber(row.cost_sold),
    profit: toNumber(row.profit),
  };
}

async function restoreSharedDataFromSupabase(options = {}) {
  const { silent = false } = options;
  try {
    const [productRows, salesRows] = await Promise.all([
      supabaseSelectRows("products", { select: "*" }),
      supabaseSelectRows("daily_sales", { select: "*", order: "sales_date.asc" }),
    ]);
    if (!productRows.length && !salesRows.length) return false;

    state.rawSales = salesRows.map(hydrateSalesFromSupabase).filter((row) => row.code || row.product);
    state.dates = [...new Set(state.rawSales.map((row) => row.date))].sort();

    state.excelItems = new Map();
    state.excelByPlu = new Map();
    state.excelByItemNumber = new Map();
    productRows.map(hydrateExcelFromSupabase).forEach((item) => addExcelIndex(item));
    rebuildExcelIndexes();

    const inventoryRows = productRows.map(hydrateInventoryFromSupabase).filter((row) => row.code || row.product);
    state.inventories = new Map();
    if (inventoryRows.length) {
      const inventoryDate = inventoryRows.map((row) => row.date).sort().at(-1) || new Date().toISOString().slice(0, 10);
      state.inventories.set(inventoryDate, inventoryRows);
    }
    buildLatestInventory();
    state._loadedFileSignatures = new Set(["supabase-shared"]);
    bumpDataStamp();
    updateFilterOptions();
    updateInventoryStateFilter();
    setDefaultDates();
    if (!silent) showToast(`Loaded shared data — ${number.format(productRows.length)} products, ${number.format(salesRows.length)} sales rows`, 3200, "success");
    return true;
  } catch (error) {
    if (!silent) showToast("Supabase shared data could not be loaded.", 3200, "warning");
    return false;
  }
}

async function syncSharedDataToSupabase(options = {}) {
  const { productsOnly = false, silent = false } = options;
  try {
    const productRows = [...state.latestInventory.values()]
      .filter((item) => item.code || item.product)
      .map(productRowForSupabase)
      .filter((row) => row.code || row.product);
    await supabaseDeleteAllRows("products");
    await supabaseInsertRows("products", productRows);

    if (!productsOnly) {
      const salesRows = state.rawSales
        .filter((row) => row.code || row.product)
        .map(salesRowForSupabase);
      await supabaseDeleteAllRows("daily_sales");
      await supabaseInsertRows("daily_sales", salesRows);
    }
    if (!silent) showToast("Shared Supabase data updated.", 2800, "success");
    return true;
  } catch (error) {
    if (!silent) showToast("Supabase shared data sync failed. Add write policies first.", 4200, "warning");
    return false;
  }
}

async function initApp() {
  mountInventoryQuickTools();
  await refreshUsersFromSupabase({ silent: true });
  const restoredShared = await restoreSharedDataFromSupabase({ silent: true });
  if (!restoredShared) await restorePersistedState();
  if (els.searchInput) {
    els.searchInput.value = state.tabSearches[activeTabName()] || "";
  }
  // Apply 90D default AFTER all data is loaded — override whatever restorePersistedState set
  if (state.dates.length) {
    state.activePresetDays = 90;
    applyDatePreset(90);
  }
  render();
}

// ── Confirm delete saved session ──────────────────────────────────────────
function openConfirmDeleteSession(sessionId) {
  state.pendingDeleteSessionId = sessionId;
  const session = state.countSessions.find((s) => s.id === sessionId);
  if (els.confirmDeleteSessionMessage) {
    els.confirmDeleteSessionMessage.textContent = session
      ? `Delete "${countSessionLabel(session)}"? This cannot be undone.`
      : "Delete this count? This cannot be undone.";
  }
  document.querySelector("#confirmDeleteSessionModal").hidden = false;
}

function confirmDeleteSavedSession() {
  const id = state.pendingDeleteSessionId;
  state.pendingDeleteSessionId = null;
  document.querySelector("#confirmDeleteSessionModal").hidden = true;
  if (!id) return;
  state.countSessions = state.countSessions.filter((s) => s.id !== id);
  persistCountSessions();
  renderCountSessionRows();
  showToast("Saved count deleted.", 2400, "warning");
}

// ── Continue count from report (re-open the session as active) ─────────────
function continueCountFromReport() {
  const sessionId = state.countReportOpenId;
  const session = findCountSessionById(sessionId);
  if (!session) { showToast("Session not found.", 3000, "warning"); return; }
  // Remove from saved, set as active
  state.countSessions = state.countSessions.filter((s) => s.id !== sessionId);
  state.activeCountSession = { ...session };
  state.countQtyBuffer = "0";
  state.selectedCountItemCode = "";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  persistCountSessions();
  closeCountReport();
  renderCountsWorkspace();
  showToast(`Continuing count: ${countSessionLabel(session)}`, 2800, "success");
}

// ── Submit & apply count (from report modal) ───────────────────────────────
function openConfirmSubmitCount() {
  const sessionId = state.countReportOpenId;
  state.pendingSubmitSessionId = sessionId;
  const session = findCountSessionById(sessionId);
  const entryCount = (session?.entries || []).length;
  const candidates = currentCountSessionCandidates(session || {});
  const nullCount = candidates.length - entryCount;
  if (els.confirmSubmitCountMessage) {
    els.confirmSubmitCountMessage.innerHTML =
      `<b>${entryCount}</b> scanned item${entryCount !== 1 ? "s" : ""} will be updated to physical counts.<br>` +
      (nullCount > 0 ? `<b>${nullCount}</b> unscanned item${nullCount !== 1 ? "s" : ""} in scope will be set to <b>0</b>.<br>` : "") +
      `<br>A pre-count snapshot will be saved so you can restore if needed.<br><br>⚠️ Are you sure you want to apply?`;
  }
  document.querySelector("#confirmSubmitCountModal").hidden = false;
}

function restorePreviousCount(sessionId) {
  const session = state.countSessions.find((s) => s.id === sessionId);
  if (!session?.preCountSnapshot) { showToast("No pre-count snapshot found for this session.", 3000, "warning"); return; }
  if (!confirm("Restore all stock values to their pre-count state? This will undo the submitted counts.")) return;
  let restored = 0;
  Object.entries(session.preCountSnapshot).forEach(([key, oldStock]) => {
    const item = state.latestInventory.get(key);
    if (!item) return;
    if (item.stock === oldStock) return;
    const before = item.stock;
    item.stock = oldStock;
    state.latestInventory.set(key, item);
    restored++;
    state.adjustmentLog.unshift({
      recordedAt: new Date().toISOString(),
      code: item.code,
      product: item.product,
      vendor: item.vendor || "",
      category: item.category || "",
      action: "RESTORE",
      qtyChange: oldStock - before,
      qtyBefore: before,
      qtyAfter: oldStock,
      reason: `Restored from pre-count snapshot: ${countSessionLabel(session)}`,
    });
  });
  session.restoredAt = new Date().toISOString();
  persistCountSessions();
  localStorage.setItem("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  bumpDataStamp();
  renderCountSessionRows();
  renderCountsWorkspace();
  render();
  renderAdjustLog();
  void syncSharedDataToSupabase({ productsOnly: true, silent: true });
  showToast(`Restored — ${restored} stock values reverted`, 3200, "success");
}

function submitAndApplyCount() {
  const sessionId = state.pendingSubmitSessionId;
  state.pendingSubmitSessionId = null;
  document.querySelector("#confirmSubmitCountModal").hidden = true;
  const session = findCountSessionById(sessionId);
  if (!session) { showToast("Session not found.", 3000, "warning"); return; }
  const entries = session.entries || [];
  const latestByCode = new Map();
  entries.forEach((entry) => latestByCode.set(codeKey(entry.code), entry));

  // Save pre-count snapshot for restore
  const snapshot = {};
  state.latestInventory.forEach((item, key) => { snapshot[key] = item.stock; });
  const savedSession = { ...session, preCountSnapshot: snapshot, submittedAt: new Date().toISOString() };
  state.countSessions = state.countSessions.map((s) => s.id === sessionId ? savedSession : s);

  // All items in scope — scanned items get their count, null/unscanned items get 0
  const candidates = currentCountSessionCandidates(session);
  const scopeCodes = new Set(candidates.map((item) => codeKey(item.code)));
  let updated = 0;
  state.latestInventory.forEach((item, key) => {
    if (!scopeCodes.has(key) && !scopeCodes.has(codeKey(item.code))) return;
    const entry = latestByCode.get(key) || latestByCode.get(codeKey(item.code));
    const before = item.stock;
    const after = entry ? Number(entry.countedQty || 0) : 0; // NULL → 0
    if (before === after) return;
    item.stock = after;
    state.latestInventory.set(key, item);
    updated++;
    state.adjustmentLog.unshift({
      recordedAt: new Date().toISOString(),
      code: item.code,
      product: item.product,
      vendor: item.vendor || "",
      category: item.category || "",
      action: "COUNT SUBMIT",
      qtyChange: after - before,
      qtyBefore: before,
      qtyAfter: after,
      reason: entry ? `Physical count: ${countSessionLabel(session)}` : `NULL → 0: ${countSessionLabel(session)}`,
    });
  });
  localStorage.setItem("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  persistCountSessions();
  bumpDataStamp();
  closeCountReport();
  renderCountsWorkspace();
  render();
  renderAdjustLog();
  generateFinalCountExport(session, candidates, latestByCode, snapshot);
  void syncSharedDataToSupabase({ productsOnly: true, silent: true });
  showToast(`Submitted — ${updated} stock values updated`, 3200, "success");
}

// ── Zero negatives ─────────────────────────────────────────────────────────
function openConfirmZeroNeg() {
  const negItems = [...state.latestInventory.values()].filter((item) => (item.stock || 0) < 0);
  if (els.confirmZeroNegMessage) {
    els.confirmZeroNegMessage.textContent = negItems.length
      ? `${negItems.length} item${negItems.length !== 1 ? "s" : ""} have negative stock and will be set to 0. A report entry will be saved for each.`
      : "No items currently have negative stock.";
  }
  document.querySelector("#confirmZeroNegModal").hidden = false;
}

function applyZeroNegatives() {
  document.querySelector("#confirmZeroNegModal").hidden = true;
  const negItems = [...state.latestInventory.values()].filter((item) => (item.stock || 0) < 0);
  if (!negItems.length) { showToast("No negative stock values found.", 2400); return; }
  negItems.forEach((item) => {
    const before = item.stock;
    item.stock = 0;
    state.latestInventory.set(codeKey(item.code), item);
    state.adjustmentLog.unshift({
      recordedAt: new Date().toISOString(),
      code: item.code,
      product: item.product,
      vendor: item.vendor || "",
      category: item.category || "",
      action: "SET",
      qtyChange: -before,
      qtyBefore: before,
      qtyAfter: 0,
      reason: "ZERO NEGATIVES",
    });
  });
  localStorage.setItem("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  bumpDataStamp();
  render();
  renderAdjustLog();
  void syncSharedDataToSupabase({ productsOnly: true, silent: true });
  showToast(`${negItems.length} negative stock values set to 0`, 2800, "success");
}

// ── Stock Adjust Modal (Products tab) ──────────────────────────────────────
function openStockAdjustModal(item) {
  state.stockAdjustItem = item;
  state.stockAdjustAction = null;
  state.stockAdjustQtyBuffer = "0";
  if (els.stockAdjustTitle) els.stockAdjustTitle.textContent = item.product || item.code;
  if (els.stockAdjustEyebrow) els.stockAdjustEyebrow.textContent = "Adjust stock";
  if (els.stockAdjustMeta) {
    els.stockAdjustMeta.innerHTML = `
      <span><b>Code</b> ${escapeHtml(item.code)}</span>
      <span><b>Current QOH</b> ${number.format(item.stock)}</span>
      <span><b>Vendor</b> ${escapeHtml(item.vendor || "-")}</span>
      <span><b>Category</b> ${escapeHtml(item.category || "-")}</span>`;
  }
  showStockAdjustStage(1);
  if (els.stockAdjustModal) els.stockAdjustModal.hidden = false;
}

function closeStockAdjustModal() {
  if (els.stockAdjustModal) els.stockAdjustModal.hidden = true;
  state.stockAdjustItem = null;
  state.stockAdjustAction = null;
  state.stockAdjustQtyBuffer = "0";
}

function showStockAdjustStage(n) {
  [els.stockAdjustStage1, els.stockAdjustStage2, els.stockAdjustStage3].forEach((el, i) => {
    if (el) el.hidden = (i + 1) !== n;
  });
}

function beginStockAdjustQty(action) {
  state.stockAdjustAction = action;
  state.stockAdjustQtyBuffer = "0";
  const labels = { add: "Qty to ADD", remove: "Qty to REMOVE", set: "Set QOH to" };
  if (els.stockAdjustActionLabel) els.stockAdjustActionLabel.textContent = labels[action] || "Quantity";
  if (els.stockAdjustQtyDisplay) els.stockAdjustQtyDisplay.textContent = "0";
  showStockAdjustStage(2);
}

function handleStockKey(key) {
  if (key === "clear") {
    state.stockAdjustQtyBuffer = "0";
  } else if (key === "back") {
    state.stockAdjustQtyBuffer = state.stockAdjustQtyBuffer.length > 1
      ? state.stockAdjustQtyBuffer.slice(0, -1) : "0";
  } else if (key === "enter") {
    if (Number(state.stockAdjustQtyBuffer) >= 0) showStockAdjustStage(3);
    return;
  } else if (key === ".") {
    if (!state.stockAdjustQtyBuffer.includes(".")) state.stockAdjustQtyBuffer += ".";
  } else {
    state.stockAdjustQtyBuffer = state.stockAdjustQtyBuffer === "0" ? key : `${state.stockAdjustQtyBuffer}${key}`;
  }
  if (els.stockAdjustQtyDisplay) els.stockAdjustQtyDisplay.textContent = state.stockAdjustQtyBuffer;
}

function finalizeStockAdjust(reason) {
  const item = state.stockAdjustItem;
  if (!item) return;
  const qty = Math.max(0, Number(state.stockAdjustQtyBuffer || "0"));
  const inventoryItem = state.latestInventory.get(codeKey(item.code));
  if (!inventoryItem) { showToast("Item not found in inventory.", 3000, "warning"); closeStockAdjustModal(); return; }
  const before = inventoryItem.stock;
  let after;
  if (state.stockAdjustAction === "add") after = before + qty;
  else if (state.stockAdjustAction === "remove") after = before - qty;
  else after = qty; // set
  inventoryItem.stock = after;
  state.latestInventory.set(codeKey(item.code), inventoryItem);
  // Also update any duplicate keys
  state.latestInventory.forEach((inv, k) => {
    if (codeKey(inv.code) === codeKey(item.code)) inv.stock = after;
  });
  // Record in log
  state.adjustmentLog.unshift({
    recordedAt: new Date().toISOString(),
    code: item.code,
    product: item.product,
    vendor: item.vendor || "",
    category: item.category || "",
    action: state.stockAdjustAction.toUpperCase(),
    qtyChange: after - before,
    qtyBefore: before,
    qtyAfter: after,
    reason,
  });
  localStorage.setItem("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  state._pinnedAdjustCode = item.code;
  bumpDataStamp();
  closeStockAdjustModal();
  render();
  renderAdjustLog();
  void syncSharedDataToSupabase({ productsOnly: true, silent: true });
  showToast(`${state.stockAdjustAction === "add" ? "Added" : state.stockAdjustAction === "remove" ? "Removed" : "Set"} ${number.format(qty)} — ${item.code} now ${number.format(after)}`, 3000, "success");
}

// ── Adjust log rendering ───────────────────────────────────────────────────
function renderAdjustLog() {
  if (!els.adjustLogBody) return;
  if (!state.adjustmentLog.length) {
    els.adjustLogBody.innerHTML = `<tr><td colspan="10" class="empty-cell">No stock adjustments recorded yet.</td></tr>`;
    return;
  }

  // Separate normal vs. auto-generated (NULL→0 and RESTORE) entries
  const normalEntries = state.adjustmentLog.filter((e) => !["NULL → 0", "RESTORE", "ZERO NEGATIVES"].some((t) => (e.reason || "").toUpperCase().includes(t) || (e.action || "").toUpperCase().includes(t)));
  const autoEntries   = state.adjustmentLog.filter((e) =>  ["NULL → 0", "RESTORE", "ZERO NEGATIVES"].some((t) => (e.reason || "").toUpperCase().includes(t) || (e.action || "").toUpperCase().includes(t)));

  function rowHtml(entry) {
    const change = entry.qtyChange || 0;
    const cls = change > 0 ? "entry-positive" : change < 0 ? "entry-negative" : "entry-exact";
    const changeLabel = change > 0 ? `+${number.format(change)}` : number.format(change);
    const bg = reasonColor_css(entry.reason || entry.action || "");
    return `<tr>
      <td>${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</td>
      <td>${escapeHtml(entry.code || "-")}</td>
      <td>${escapeHtml(entry.product || "-")}</td>
      <td>${escapeHtml(entry.vendor || "-")}</td>
      <td>${escapeHtml(entry.category || "-")}</td>
      <td>${escapeHtml((entry.action || "-").toUpperCase())}</td>
      <td class="num ${cls}">${changeLabel}</td>
      <td class="num">${number.format(entry.qtyBefore ?? 0)}</td>
      <td class="num">${number.format(entry.qtyAfter ?? 0)}</td>
      <td><span class="reason-chip" style="background:${bg}">${escapeHtml((entry.reason || "-").toUpperCase())}</span></td>
    </tr>`;
  }

  const normalHtml = normalEntries.map(rowHtml).join("") ||
    `<tr><td colspan="10" class="empty-cell">No manual adjustments yet.</td></tr>`;

  // Auto entries as a collapsible block spanning all columns
  const autoHtml = autoEntries.length
    ? `<tr class="auto-entries-toggle-row"><td colspan="10">
        <details class="auto-entries-details">
          <summary>${autoEntries.length} auto-generated entr${autoEntries.length === 1 ? "y" : "ies"} (NULL→0 / RESTORE / ZERO NEGATIVES)</summary>
          <table class="count-report-table inner-auto-table">
            <thead><tr><th>Date/Time</th><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th>Action</th><th>Change</th><th>Before</th><th>After</th><th>Reason</th></tr></thead>
            <tbody>${autoEntries.map(rowHtml).join("")}</tbody>
          </table>
        </details>
      </td></tr>`
    : "";

  els.adjustLogBody.innerHTML = normalHtml + autoHtml;
}

// Reason → color map for reports
function reasonColor_css(reason) {
  const r = (reason || "").toUpperCase();
  if (r.includes("DAMAGED"))           return "#e67e22";
  if (r.includes("STOLEN"))            return "#c0392b";
  if (r.includes("LOST"))              return "#8e44ad";
  if (r.includes("MISCOUNT"))          return "#2980b9";
  if (r.includes("SAMPLE"))            return "#16a085";
  if (r.includes("PROMOTION"))         return "#d35400";
  if (r.includes("RETURNED"))          return "#7f8c8d";
  if (r.includes("COUNT SUBMIT"))      return "#1e8bc3";
  if (r.includes("RESTORE"))           return "#27ae60";
  if (r.includes("ZERO NEGATIVES"))    return "#f39c12";
  return "#555";
}

async function exportAdjustLogPdf() {
  if (!state.adjustmentLog.length) { showToast("No adjustment records to export.", 3000, "warning"); return; }
  const dateStr = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const rows = state.adjustmentLog.map((entry) => {
    const change = entry.qtyChange || 0;
    const cls = change > 0 ? "var-up" : change < 0 ? "var-down" : "";
    const changeLabel = change > 0 ? `+${number.format(change)}` : number.format(change);
    const bg = reasonColor_css(entry.reason || entry.action || "");
    return `<tr class="${cls}">
      <td>${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</td>
      <td>${escapeHtml(entry.code || "-")}</td>
      <td>${escapeHtml(entry.product || "-")}</td>
      <td>${escapeHtml(entry.vendor || "-")}</td>
      <td>${escapeHtml(entry.category || "-")}</td>
      <td>${escapeHtml(entry.action || "-")}</td>
      <td class="num">${changeLabel}</td>
      <td class="num">${number.format(entry.qtyBefore ?? 0)}</td>
      <td class="num">${number.format(entry.qtyAfter ?? 0)}</td>
      <td><span style="display:inline-block;padding:2px 7px;border-radius:4px;background:${bg};color:#fff;font-weight:700;font-size:9px">${escapeHtml((entry.reason || "-").toUpperCase())}</span></td>
    </tr>`;
  }).join("");
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Stock Adjustment Log</title>
  <style>
    body { font-family: Arial, sans-serif; font-size: 10px; color: #1c2320; margin: 0; padding: 16px; }
    h1 { font-size: 16px; margin: 0 0 4px; }
    .meta { color: #66716d; margin-bottom: 16px; }
    table { width: 100%; border-collapse: collapse; }
    th { background: #eef7f0; text-align: left; padding: 4px 5px; font-size: 9px; text-transform: uppercase; border-bottom: 2px solid #dce3df; }
    td { padding: 3px 5px; border-bottom: 1px solid #eee; vertical-align: middle; }
    .num { text-align: right; }
    .var-up td { color: #16835b; } .var-down td { color: #c0392b; }
    @media print { body { padding: 4px; } }
  </style></head><body>
  <h1>Stock Adjustment Log</h1>
  <div class="meta">Generated ${dateStr} &nbsp;·&nbsp; ${state.adjustmentLog.length} records</div>
  <table><thead><tr><th>Date/Time</th><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th>Action</th><th>Change</th><th>Qty Before</th><th>Qty After</th><th>Reason</th></tr></thead>
  <tbody>${rows}</tbody></table></body></html>`;
  const win = window.open("", "_blank", "width=1100,height=750");
  if (!win) { showToast("Pop-up blocked — please allow pop-ups.", 3500, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 500);
}

async function exportAdjustLogExcel() {
  if (!state.adjustmentLog.length) { showToast("No adjustment records to export.", 3000, "warning"); return; }
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available.", 3000, "warning"); return; }
  const wb = xlsx.utils.book_new();
  const data = [
    ["Stock Adjustment Log", "", "", "", "", "", "", "", "", `Generated: ${new Date().toLocaleDateString()}`],
    [],
    ["Date/Time", "Code", "Item", "Vendor", "Category", "Action", "Qty Change", "Qty Before", "Qty After", "Reason"],
    ...state.adjustmentLog.map((entry) => [
      new Date(entry.recordedAt).toLocaleString(),
      entry.code || "", entry.product || "", entry.vendor || "", entry.category || "",
      entry.action || "", entry.qtyChange || 0, entry.qtyBefore ?? 0, entry.qtyAfter ?? 0, (entry.reason || "").toUpperCase(),
    ]),
  ];
  const ws = xlsx.utils.aoa_to_sheet(data);
  ws["!cols"] = [20, 12, 32, 14, 14, 12, 10, 10, 10, 26].map((w) => ({ wch: w }));
  xlsx.utils.book_append_sheet(wb, ws, "Adjustment Log");
  xlsx.writeFile(wb, `StockAdjustments_${new Date().toISOString().slice(0, 10)}.xlsx`);
  showToast(`Exported ${state.adjustmentLog.length} adjustment records`, 2800, "success");
}

// ── Final Count Report (submitted sessions only) ────────────────────────────
function openFinalCountReport(sessionId) {
  const session = state.countSessions.find((s) => s.id === sessionId);
  if (!session?.preCountSnapshot) {
    showToast("No submitted snapshot found for this session.", 3000, "warning");
    return;
  }
  state.finalReportSessionId = sessionId;
  if (els.finalReportTitle) els.finalReportTitle.textContent = `Final Count — ${countSessionLabel(session)}`;
  if (els.finalReportMeta) {
    els.finalReportMeta.innerHTML = `
      <span><b>Date</b> ${escapeHtml(session.date || "-")}</span>
      <span><b>Vendor</b> ${escapeHtml(session.vendor || "All")}</span>
      <span><b>Category</b> ${escapeHtml(session.category || "All")}</span>
      <span><b>Status filter</b> ${escapeHtml(session.status || "All")}</span>
      <span><b>Submitted</b> ${escapeHtml(new Date(session.submittedAt).toLocaleString())}</span>
      <span><b>Entries</b> ${number.format((session.entries || []).length)}</span>`;
  }
  renderFinalCountReportRows(session);
  document.querySelector("#finalCountReportModal").hidden = false;
}

function renderFinalCountReportRows(session) {
  if (!els.finalReportBody) return;
  const snapshot = session.preCountSnapshot || {};
  const entries = session.entries || [];
  const latestByCode = new Map();
  entries.forEach((e) => latestByCode.set(codeKey(e.code), e));
  const candidates = currentCountSessionCandidates(session);
  if (!candidates.length) {
    els.finalReportBody.innerHTML = `<tr><td colspan="9" class="empty-cell">No items in scope.</td></tr>`;
    return;
  }
  els.finalReportBody.innerHTML = candidates.map((item) => {
    const key = codeKey(item.code);
    const qtyStart = snapshot[key] ?? Number(item.stock ?? 0);
    const entry = latestByCode.get(key);
    const qtyEnd = entry ? Number(entry.countedQty || 0) : 0;
    const variance = qtyEnd - qtyStart;
    const costVar = variance * Number(item.unitCost || 0);
    const vCls = variance > 0 ? "entry-positive" : variance < 0 ? "entry-negative" : "entry-exact";
    const status = entry ? "Scanned" : "Not scanned (→ 0)";
    return `<tr>
      <td>${escapeHtml(item.code)}</td>
      <td>${escapeHtml(item.product)}</td>
      <td>${escapeHtml(item.vendor || "-")}</td>
      <td>${escapeHtml(item.category || "-")}</td>
      <td class="num">${number.format(qtyStart)}</td>
      <td class="num">${number.format(qtyEnd)}</td>
      <td class="num ${vCls}">${variance > 0 ? "+" : ""}${number.format(variance)}</td>
      <td class="num ${vCls}">${currency.format(costVar)}</td>
      <td>${escapeHtml(status)}</td>
    </tr>`;
  }).join("");
}

function exportFinalCountReportPdf() {
  const session = state.countSessions.find((s) => s.id === state.finalReportSessionId);
  if (!session) return;
  const snapshot = session.preCountSnapshot || {};
  const entries = session.entries || [];
  const latestByCode = new Map();
  entries.forEach((e) => latestByCode.set(codeKey(e.code), e));
  const candidates = currentCountSessionCandidates(session);
  const dateStr = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const rowsHtml = candidates.map((item) => {
    const key = codeKey(item.code);
    const qtyStart = snapshot[key] ?? Number(item.stock ?? 0);
    const entry = latestByCode.get(key);
    const qtyEnd = entry ? Number(entry.countedQty || 0) : 0;
    const variance = qtyEnd - qtyStart;
    const costVar = variance * Number(item.unitCost || 0);
    const cls = variance > 0 ? "var-up" : variance < 0 ? "var-down" : "";
    return `<tr class="${cls}">
      <td>${escapeHtml(item.code)}</td><td>${escapeHtml(item.product)}</td>
      <td>${escapeHtml(item.vendor || "-")}</td><td>${escapeHtml(item.category || "-")}</td>
      <td class="num">${number.format(qtyStart)}</td>
      <td class="num">${number.format(qtyEnd)}</td>
      <td class="num">${variance > 0 ? "+" : ""}${number.format(variance)}</td>
      <td class="num">${currency.format(costVar)}</td>
      <td>${entry ? "Scanned" : "Not scanned (→ 0)"}</td>
    </tr>`;
  }).join("");
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Final Count Report</title>
  <style>
    body{font-family:Arial,sans-serif;font-size:11px;color:#1c2320;margin:0;padding:20px}
    h1{font-size:17px;margin:0 0 4px}
    .meta{display:flex;gap:16px;flex-wrap:wrap;margin-bottom:16px;font-size:10px;color:#555}
    .meta span{background:#f0f4f2;padding:3px 8px;border-radius:4px}
    table{width:100%;border-collapse:collapse}
    th{background:#eef7f0;text-align:left;padding:4px 6px;font-size:9px;text-transform:uppercase;border-bottom:2px solid #dce3df}
    td{padding:3px 6px;border-bottom:1px solid #eee}
    .num{text-align:right}
    .var-up td{color:#16835b}.var-down td{color:#c0392b}
    @media print{body{padding:6px}}
  </style></head><body>
  <h1>Final Physical Count Report</h1>
  <div class="meta">
    <span><b>Count date</b> ${escapeHtml(session.date || "-")}</span>
    <span><b>Vendor</b> ${escapeHtml(session.vendor || "All")}</span>
    <span><b>Category</b> ${escapeHtml(session.category || "All")}</span>
    <span><b>Submitted</b> ${escapeHtml(new Date(session.submittedAt).toLocaleString())}</span>
    <span><b>Generated</b> ${dateStr}</span>
  </div>
  <table><thead><tr><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th>Qty Start</th><th>Qty End</th><th>Variance</th><th>Cost Var</th><th>Status</th></tr></thead>
  <tbody>${rowsHtml}</tbody></table></body></html>`;
  const win = window.open("", "_blank", "width=1100,height=750");
  if (!win) { showToast("Pop-up blocked.", 3000, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 400);
}

async function exportFinalCountReportExcel() {
  const session = state.countSessions.find((s) => s.id === state.finalReportSessionId);
  if (!session) return;
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available.", 3000, "warning"); return; }
  const snapshot = session.preCountSnapshot || {};
  const entries = session.entries || [];
  const latestByCode = new Map();
  entries.forEach((e) => latestByCode.set(codeKey(e.code), e));
  const candidates = currentCountSessionCandidates(session);
  const wb = xlsx.utils.book_new();
  const data = [
    ["Final Physical Count Report", "", `Date: ${session.date || "-"}`, `Vendor: ${session.vendor || "All"}`, `Submitted: ${new Date(session.submittedAt).toLocaleString()}`],
    [],
    ["Code", "Item", "Vendor", "Category", "Qty Start", "Qty End", "Variance", "Cost Variance", "Status"],
    ...candidates.map((item) => {
      const key = codeKey(item.code);
      const qtyStart = snapshot[key] ?? Number(item.stock ?? 0);
      const entry = latestByCode.get(key);
      const qtyEnd = entry ? Number(entry.countedQty || 0) : 0;
      const variance = qtyEnd - qtyStart;
      const costVar = variance * Number(item.unitCost || 0);
      return [item.code, item.product, item.vendor || "", item.category || "", qtyStart, qtyEnd, variance, costVar, entry ? "Scanned" : "Not scanned (→ 0)"];
    }),
  ];
  const ws = xlsx.utils.aoa_to_sheet(data);
  ws["!cols"] = [14, 34, 14, 14, 10, 10, 10, 12, 18].map((w) => ({ wch: w }));
  applyXlsxTextToCodeColumns(ws, data, [0]); // code column = index 0
  xlsx.utils.book_append_sheet(wb, ws, "Final Count");
  xlsx.writeFile(wb, `FinalCount_${session.date || "report"}.xlsx`);
  showToast(`Final count exported — ${candidates.length} items`, 2800, "success");
}

// ── Vendor Rules (Vendors tab) ──────────────────────────────────────────────
function persistVendorRules() {
  localStorage.setItem("posDashboardVendorRules:v1", JSON.stringify(state.vendorRules));
}

const DAYS_OF_WEEK = ["monday","tuesday","wednesday","thursday","friday","saturday","sunday"];
const DAY_SHORT = { monday:"Mon",tuesday:"Tue",wednesday:"Wed",thursday:"Thu",friday:"Fri",saturday:"Sat",sunday:"Sun" };

function renderVendorRules() {
  const body = document.querySelector("#vendorRulesBody");
  if (!body) return;
  const today = new Date().toLocaleDateString("en-US",{weekday:"long"}).toLowerCase();
  const statusFilter = document.querySelector("#vendorStatusFilter")?.value || "";

  // Update summary strip
  const summary = document.querySelector("#vendorRulesSummary");
  if (summary) {
    const active = state.vendorRules.filter((r) => r.status === "Active").length;
    const todayVendors = state.vendorRules.filter((r) => r.status === "Active" && (r.orderDays || []).includes(today));
    summary.innerHTML = `<span><b>${state.vendorRules.length}</b> vendors configured</span>
      <span><b>${active}</b> active for auto-ordering</span>
      ${todayVendors.length ? `<span class="vendor-today-alert">⚠️ <b>${todayVendors.length}</b> vendor${todayVendors.length>1?"s":""} to order TODAY: ${todayVendors.map(v=>v.vendor).join(", ")}</span>` : ""}`;
  }

  const filtered = statusFilter
    ? state.vendorRules.filter((r) => r.status === statusFilter)
    : state.vendorRules;

  if (!filtered.length) {
    body.innerHTML = `<tr><td colspan="10" class="empty-cell">${state.vendorRules.length ? "No vendors match filter." : 'No vendor rules set. Click "+ Add vendor" to create one.'}</td></tr>`;
    return;
  }
  body.innerHTML = filtered
    .slice()
    .sort((a,b) => (a.vendor||"").localeCompare(b.vendor||""))
    .map((rule) => {
      const isActive = rule.status === "Active";
      const isToday = isActive && (rule.orderDays||[]).includes(today);
      const dayChips = DAYS_OF_WEEK.map((d) => {
        const on = (rule.orderDays||[]).includes(d);
        return `<span class="vday-chip vday-inline ${on?"vday-on":"vday-off"} ${d===today&&on?"vday-today":""}" data-rule-id="${escapeHtml(rule.id)}" data-day="${d}" style="cursor:pointer" title="Toggle ${DAY_SHORT[d]}">${DAY_SHORT[d]}</span>`;
      }).join("");
      return `<tr class="${isToday?"vendor-order-today":""}">
        <td class="checkbox-col"><input type="checkbox" class="vendor-row-cb" data-vendor-id="${escapeHtml(rule.id)}" /></td>
        <td>
          <select class="vendor-inline-select vendor-status-select" data-rule-id="${escapeHtml(rule.id)}" data-field="status">
            <option value="Active" ${isActive?"selected":""}>Active</option>
            <option value="Disabled" ${!isActive?"selected":""}>Disabled</option>
          </select>
        </td>
        <td class="vendor-name-cell"><b>${escapeHtml(rule.vendor||"-")}</b></td>
        <td class="num"><input type="number" class="vendor-inline-input" data-rule-id="${escapeHtml(rule.id)}" data-field="safetyDays" value="${rule.safetyDays??21}" min="0" max="365" style="width:4.5rem" /></td>
        <td class="num"><input type="number" class="vendor-inline-input" data-rule-id="${escapeHtml(rule.id)}" data-field="daysOfInventory" value="${rule.daysOfInventory??7}" min="0" max="365" style="width:4.5rem" /></td>
        <td class="num"><input type="number" class="vendor-inline-input" data-rule-id="${escapeHtml(rule.id)}" data-field="minOrder" value="${rule.minOrder||0}" min="0" style="width:5rem" /></td>
        <td><div class="vday-chips">${dayChips}</div></td>
        <td class="vendor-email-cell"><input type="email" class="vendor-inline-input vendor-email-input" data-rule-id="${escapeHtml(rule.id)}" data-field="email" value="${escapeHtml(rule.email||"")}" placeholder="email@vendor.com" /></td>
        <td><input type="text" class="vendor-inline-input" data-rule-id="${escapeHtml(rule.id)}" data-field="notes" value="${escapeHtml(rule.notes||"")}" placeholder="Notes..." /></td>
        <td class="vendor-actions">
          <button type="button" class="delete-session-btn vendor-delete-btn" data-vendor-id="${escapeHtml(rule.id)}">Del</button>
        </td>
      </tr>`;
    }).join("");

  // Inline number/text/email inputs — save on blur or Enter
  body.querySelectorAll(".vendor-inline-input").forEach((input) => {
    const save = () => {
      const rule = state.vendorRules.find((r) => r.id === input.dataset.ruleId);
      if (!rule) return;
      const field = input.dataset.field;
      const val = ["safetyDays","daysOfInventory","minOrder"].includes(field)
        ? Math.max(0, toNumber(input.value)||0) : input.value.trim();
      rule[field] = val;
      rule.updatedAt = new Date().toISOString();
      persistVendorRules();
    };
    input.addEventListener("blur", save);
    input.addEventListener("keydown", (e) => { if (e.key==="Enter") { e.preventDefault(); input.blur(); } });
  });
  // Status select
  body.querySelectorAll(".vendor-status-select").forEach((sel) => {
    sel.addEventListener("change", () => {
      const rule = state.vendorRules.find((r) => r.id === sel.dataset.ruleId);
      if (!rule) return;
      rule.status = sel.value;
      rule.updatedAt = new Date().toISOString();
      persistVendorRules();
      renderVendorRules(); // re-render to update row highlight
    });
  });
  // Day chip toggles
  body.querySelectorAll(".vday-inline").forEach((chip) => {
    chip.addEventListener("click", () => {
      const rule = state.vendorRules.find((r) => r.id === chip.dataset.ruleId);
      if (!rule) return;
      const day = chip.dataset.day;
      rule.orderDays = rule.orderDays || [];
      if (rule.orderDays.includes(day)) rule.orderDays = rule.orderDays.filter((d) => d !== day);
      else rule.orderDays.push(day);
      rule.updatedAt = new Date().toISOString();
      persistVendorRules();
      // Update chip visually without full re-render
      chip.classList.toggle("vday-on", rule.orderDays.includes(day));
      chip.classList.toggle("vday-off", !rule.orderDays.includes(day));
    });
  });
  body.querySelectorAll(".vendor-delete-btn").forEach((btn) => {
    btn.addEventListener("click", () => {
      if (!confirm("Delete this vendor rule?")) return;
      state.vendorRules = state.vendorRules.filter((r) => r.id !== btn.dataset.vendorId);
      persistVendorRules();
      renderVendorRules();
      showToast("Vendor rule deleted.", 2400, "warning");
    });
  });
  body.querySelectorAll(".vendor-row-cb").forEach((cb) => {
    cb.addEventListener("change", () => updateVendorBulkBar());
  });
  // Reset select-all state on re-render
  const selectAll = document.querySelector("#selectAllVendors");
  if (selectAll) selectAll.checked = false;
  updateVendorBulkBar();
}

function openVendorRuleModal(id) {
  const rule = id ? state.vendorRules.find((r) => r.id === id) : null;
  state.vendorRuleEditId = id || null;
  state.vendorRuleSelectedDays = new Set(rule?.orderDays || []);

  document.querySelector("#vendorModalEyebrow").textContent = rule ? "Edit vendor" : "Add vendor";
  document.querySelector("#vendorModalTitle").textContent = rule ? rule.vendor : "New vendor rule";
  document.querySelector("#vrVendor").value = rule?.vendor || "";
  document.querySelector("#vrStatus").value = rule?.status || "Active";
  document.querySelector("#vrSafetyDays").value = rule?.safetyDays ?? 21;
  document.querySelector("#vrDaysOfInventory").value = rule?.daysOfInventory ?? 7;
  document.querySelector("#vrMinOrder").value = rule?.minOrder ?? 0;
  document.querySelector("#vrEmail").value = rule?.email || "";
  document.querySelector("#vrNotes").value = rule?.notes || "";

  // Update day button states
  document.querySelectorAll(".vday-btn").forEach((btn) => {
    btn.classList.toggle("vday-btn-on", state.vendorRuleSelectedDays.has(btn.dataset.day));
  });

  document.querySelector("#vendorRuleModal").hidden = false;
}

function saveVendorRule() {
  const vendor = document.querySelector("#vrVendor").value.trim().toUpperCase();
  if (!vendor) { showToast("Vendor name is required.", 2800, "warning"); return; }

  const rule = {
    id: state.vendorRuleEditId || `vr-${Date.now()}`,
    vendor,
    status: document.querySelector("#vrStatus").value,
    safetyDays: toNumber(document.querySelector("#vrSafetyDays").value) || 21,
    daysOfInventory: toNumber(document.querySelector("#vrDaysOfInventory").value) || 7,
    minOrder: toNumber(document.querySelector("#vrMinOrder").value) || 0,
    email: document.querySelector("#vrEmail").value.trim(),
    notes: document.querySelector("#vrNotes").value.trim(),
    orderDays: [...state.vendorRuleSelectedDays],
    updatedAt: new Date().toISOString(),
  };

  if (state.vendorRuleEditId) {
    state.vendorRules = state.vendorRules.map((r) => r.id === state.vendorRuleEditId ? rule : r);
  } else {
    state.vendorRules.push(rule);
  }
  persistVendorRules();
  document.querySelector("#vendorRuleModal").hidden = true;
  renderVendorRules();
  showToast(`Vendor rule saved: ${rule.vendor}`, 2400, "success");
}

async function exportVendorRules() {
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available.", 3000, "warning"); return; }
  const wb = xlsx.utils.book_new();
  const data = [
    ["Vendor Rules", "", "", "", "", "", `Generated: ${new Date().toLocaleDateString()}`],
    [],
    ["Vendor","Status","Safety Days","Days of Inventory","Min Order $","Order Days","Email","Notes"],
    ...state.vendorRules.map((r) => [
      r.vendor, r.status, r.safetyDays, r.daysOfInventory, r.minOrder,
      (r.orderDays||[]).join(", "), r.email||"", r.notes||""
    ]),
  ];
  const ws = xlsx.utils.aoa_to_sheet(data);
  ws["!cols"] = [20,10,12,14,12,30,28,24].map((w) => ({wch:w}));
  xlsx.utils.book_append_sheet(wb, ws, "Vendor Rules");
  xlsx.writeFile(wb, `VendorRules_${new Date().toISOString().slice(0,10)}.xlsx`);
  showToast("Vendor rules exported.", 2400, "success");
}

// Wire new inline date-period buttons in filter bar
document.querySelectorAll(".date-period-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    const period = btn.dataset.period;
    applyDatePreset(period === "ytd" ? "ytd" : Number(period));
    // Highlight active button
    document.querySelectorAll(".date-period-btn").forEach((b) => b.classList.remove("date-period-btn--active"));
    btn.classList.add("date-period-btn--active");
  });
});
document.querySelector("#addVendorRuleButton")?.addEventListener("click", () => openVendorRuleModal(null));
document.querySelector("#exportVendorRulesButton")?.addEventListener("click", () => exportVendorRules());
document.querySelector("#vendorRuleSaveButton")?.addEventListener("click", () => saveVendorRule());
document.querySelector("#preloadVendorsButton")?.addEventListener("click", () => preloadVendorsFromInventory());
document.querySelector("#vendorStatusFilter")?.addEventListener("change", () => renderVendorRules());

// Select-all vendors
document.querySelector("#selectAllVendors")?.addEventListener("change", (e) => {
  document.querySelectorAll(".vendor-row-cb").forEach((cb) => { cb.checked = e.target.checked; });
  updateVendorBulkBar();
});

// Bulk action buttons
document.querySelector("#vendorBulkActivate")?.addEventListener("click", () => bulkSetVendorStatus("Active"));
document.querySelector("#vendorBulkDisable")?.addEventListener("click", () => bulkSetVendorStatus("Disabled"));
document.querySelector("#vendorBulkDelete")?.addEventListener("click", () => bulkDeleteVendors());
document.querySelector("#vendorBulkClear")?.addEventListener("click", () => {
  document.querySelectorAll(".vendor-row-cb").forEach((cb) => { cb.checked = false; });
  if (document.querySelector("#selectAllVendors")) document.querySelector("#selectAllVendors").checked = false;
  updateVendorBulkBar();
});

function getSelectedVendorIds() {
  return [...document.querySelectorAll(".vendor-row-cb:checked")].map((cb) => cb.dataset.vendorId);
}

function updateVendorBulkBar() {
  const ids = getSelectedVendorIds();
  const bar = document.querySelector("#vendorBulkBar");
  const countEl = document.querySelector("#vendorBulkCount");
  if (bar) bar.hidden = ids.length === 0;
  if (countEl) countEl.textContent = `${ids.length} selected`;
}

function bulkSetVendorStatus(status) {
  const ids = new Set(getSelectedVendorIds());
  if (!ids.size) return;
  state.vendorRules = state.vendorRules.map((r) => ids.has(r.id) ? { ...r, status } : r);
  persistVendorRules();
  renderVendorRules();
  showToast(`${ids.size} vendor${ids.size > 1 ? "s" : ""} set to ${status}`, 2400, "success");
}

function bulkDeleteVendors() {
  const ids = new Set(getSelectedVendorIds());
  if (!ids.size) return;
  if (!confirm(`Delete ${ids.size} vendor rule${ids.size > 1 ? "s" : ""}?`)) return;
  state.vendorRules = state.vendorRules.filter((r) => !ids.has(r.id));
  persistVendorRules();
  renderVendorRules();
  showToast(`${ids.size} vendor rule${ids.size > 1 ? "s" : ""} deleted.`, 2400, "warning");
}

function preloadVendorsFromInventory() {
  // Collect all unique vendor names from inventory + sales data
  const allRows = [...state.latestInventory.values()];
  const excelRows = [...state.excelItems.values()];
  const salesVendors = state.rawSales.map((r) => r.vendor).filter(Boolean);
  const allVendors = unique([
    ...allRows.map((r) => r.vendor),
    ...excelRows.map((r) => r.vendor),
    ...salesVendors,
  ].filter((v) => v && v !== "Unassigned" && v !== "-"));

  if (!allVendors.length) {
    showToast("No inventory or sales data loaded yet — load your CSV files first.", 3500, "warning");
    return;
  }

  const existingNames = new Set(state.vendorRules.map((r) => r.vendor?.toUpperCase()));
  const newVendors = allVendors.filter((v) => !existingNames.has(v.toUpperCase()));

  if (!newVendors.length) {
    showToast("All vendors from your data already have rules.", 2800);
    return;
  }

  newVendors.forEach((v) => {
    state.vendorRules.push({
      id: `vr-${Date.now()}-${Math.random().toString(36).slice(2,6)}`,
      vendor: v.toUpperCase(),
      status: "Disabled",
      safetyDays: 21,
      daysOfInventory: 7,
      minOrder: 0,
      email: "",
      notes: "",
      orderDays: [],
      updatedAt: new Date().toISOString(),
    });
  });

  persistVendorRules();
  renderVendorRules();
  showToast(`Preloaded ${newVendors.length} vendor${newVendors.length > 1 ? "s" : ""} from inventory data.`, 3000, "success");
}
document.querySelector("#vendorRuleCancelButton")?.addEventListener("click", () => { document.querySelector("#vendorRuleModal").hidden = true; });
document.querySelectorAll(".vday-btn").forEach((btn) => {
  btn.addEventListener("click", () => {
    const day = btn.dataset.day;
    if (state.vendorRuleSelectedDays.has(day)) state.vendorRuleSelectedDays.delete(day);
    else state.vendorRuleSelectedDays.add(day);
    btn.classList.toggle("vday-btn-on", state.vendorRuleSelectedDays.has(day));
  });
});

// ── Pending Orders ────────────────────────────────────────────────────────
function loadPendingOrders() {
  return JSON.parse(localStorage.getItem("posPendingOrders:v1") || "[]");
}
function savePendingOrders() {
  localStorage.setItem("posPendingOrders:v1", JSON.stringify(state.pendingOrders));
}
if (!state.pendingOrders) state.pendingOrders = loadPendingOrders();

function isPendingOrder(code) {
  if (!state.pendingOrders?.length) return false;
  return state.pendingOrders.some((po) => !po.cleared && (po.codes||[]).includes(codeKey(code)));
}

function submitVendorPo(vendorName) {
  const items = currentOrderRows().filter((item) => (item.vendor||"").toUpperCase() === vendorName.toUpperCase());
  if (!items.length) { showToast(`No items to order for ${vendorName}.`, 2800, "warning"); return; }
  const clearAt = Date.now() + 10 * 24 * 60 * 60 * 1000; // 10 days
  const po = {
    id: `po-${Date.now()}`,
    vendor: vendorName,
    codes: items.map((item) => codeKey(item.code)),
    submittedAt: new Date().toISOString(),
    clearAt,
    cleared: false,
  };
  state.pendingOrders = [...(state.pendingOrders||[]), po];
  savePendingOrders();
  state.adjustmentLog.unshift({
    recordedAt: new Date().toISOString(),
    code: "-",
    product: "PO submitted: " + vendorName + " (" + items.length + " items)",
    vendor: vendorName,
    category: "-",
    action: "PO SUBMIT",
    qtyChange: 0,
    qtyBefore: 0,
    qtyAfter: 0,
    reason: currency.format(items.reduce((s, i) => s + (i.caseOrder||0)*(i.unitCost||0), 0)) + " total",
  });
  localStorage.setItem("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  showToast("PO submitted for " + vendorName + " - " + items.length + " items pending", 3200, "success");
  renderOrders();
  renderAdjustLog();
}

function submitAllPo() {
  const today = new Date().toLocaleDateString("en-US", { weekday: "long" }).toLowerCase();
  const todayVendors = state.vendorRules.filter((r) => r.status === "Active" && (r.orderDays||[]).includes(today));
  if (!todayVendors.length) { showToast("No vendors scheduled to order today.", 3000, "warning"); return; }
  if (!confirm(`Submit POs for all ${todayVendors.length} vendor${todayVendors.length>1?"s":""} scheduled today?`)) return;
  todayVendors.forEach((r) => submitVendorPo(r.vendor));
}

function clearExpiredPendingOrders() {
  if (!state.pendingOrders?.length) return;
  const now = Date.now();
  let changed = false;
  state.pendingOrders = state.pendingOrders.map((po) => {
    if (!po.cleared && po.clearAt && now > po.clearAt) { changed = true; return { ...po, cleared: true }; }
    return po;
  });
  if (changed) savePendingOrders();
}
// Run on load
clearExpiredPendingOrders();

// ── Vendor Analysis Panel (inline in ordering tab) ────────────────────────
function exportVendorPoPdf(vendorName, items) {
  const today = new Date().toLocaleDateString("en-US", { year:"numeric", month:"long", day:"numeric" });
  const body = document.querySelector("#vpmBody");
  const rows = [...(body?.querySelectorAll("tr[data-vpm-code]") || [])];
  let grandTotal = 0;
  const rowsHtml = rows.map((tr) => {
    const qtyInput = tr.querySelector(".vpm-qty-input");
    const qty = Math.max(0, toNumber(qtyInput?.value||0));
    const uc = parseFloat(qtyInput?.dataset.unitCost||0);
    const lineCost = qty * uc;
    grandTotal += lineCost;
    const cells = [...tr.cells];
    return `<tr>
      <td>${cells[0]?.textContent.trim()||""}</td>
      <td>${cells[1]?.textContent.trim()||""}</td>
      <td class="num">${cells[2]?.textContent.trim()||""}</td>
      <td class="num">${cells[3]?.textContent.trim()||""}</td>
      <td class="num">${cells[4]?.textContent.trim()||""}</td>
      <td class="num">${cells[5]?.textContent.trim()||""}</td>
      <td class="num">${cells[6]?.textContent.trim()||""}</td>
      <td class="num"><b>${qty}</b></td>
      <td class="num">${cells[8]?.textContent.trim()||""}</td>
      <td class="num">${cells[9]?.textContent.trim()||""}</td>
      <td class="num"><b>${currency.format(lineCost)}</b></td>
    </tr>`;
  }).join("");
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>PO - ${escapeHtml(vendorName)}</title>
  <style>
    body { font-family: Arial, sans-serif; font-size: 11px; color: #1c2320; padding: 20px; margin: 0; }
    h1 { font-size: 20px; margin: 0 0 2px; }
    .meta { display:flex; gap:16px; flex-wrap:wrap; margin:8px 0 16px; font-size:10px; color:#555; }
    .meta span { background:#f0f4f2; padding:3px 8px; border-radius:4px; }
    table { width:100%; border-collapse:collapse; }
    th { background:#eef7f0; text-align:left; padding:5px 6px; font-size:9px; text-transform:uppercase; border-bottom:2px solid #dce3df; }
    td { padding:4px 6px; border-bottom:1px solid #eee; }
    .num { text-align:right; }
    .grand { font-size:14px; font-weight:700; text-align:right; padding:10px 6px; border-top:2px solid #1c2320; }
    @media print { body { padding:8px; } }
  </style></head><body>
  <h1>Purchase Order — ${escapeHtml(vendorName)}</h1>
  <div class="meta">
    <span><b>Date</b> ${today}</span>
    <span><b>Items</b> ${rows.length}</span>
    <span><b>Total</b> ${currency.format(grandTotal)}</span>
  </div>
  <table><thead><tr><th>Code</th><th>Item</th><th>Stock</th><th>SV/day</th><th>Min</th><th>Max</th><th>Rec</th><th>Order Qty</th><th>Case</th><th>Unit Cost</th><th>Total</th></tr></thead>
  <tbody>${rowsHtml}</tbody></table>
  <div class="grand">Grand Total: ${currency.format(grandTotal)}</div>
  </body></html>`;
  const win = window.open("", "_blank", "width=900,height=700");
  if (!win) { showToast("Pop-up blocked.", 3000, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 400);
}

document.querySelector("#vpmCloseButton")?.addEventListener("click", () => {
  document.querySelector("#vendorPoModal").hidden = true;
});
// Also wire Esc in the global keydown handler via the existing modals check

function openVendorAnalysisPanel(vendorName) {
  const modal = document.querySelector("#vendorPoModal");
  if (!modal) return;
  const items = currentOrderRows().filter((item) => (item.vendor||"").toUpperCase() === vendorName.toUpperCase());
  const rule = state.vendorRules.find((r) => r.vendor && r.vendor.toUpperCase() === vendorName.toUpperCase());
  const totalCost = items.reduce((s, item) => s + (item.caseOrder||0)*(item.unitCost||0), 0);
  const minOk = !rule || !rule.minOrder || totalCost >= rule.minOrder;
  const isPending = (state.pendingOrders||[]).some((po) => po.vendor === vendorName && !po.cleared);

  const titleEl = document.querySelector("#vpmTitle");
  if (titleEl) titleEl.textContent = vendorName;
  const metaEl = document.querySelector("#vpmMeta");
  if (metaEl) metaEl.innerHTML =
    "<span><b>" + items.length + "</b> items to order</span>" +
    "<span><b>Total</b> " + currency.format(totalCost) + "</span>" +
    (rule && rule.minOrder ? "<span class='" + (minOk ? "" : "order-min-warn-text") + "'><b>Min order</b> " + currency.format(rule.minOrder) + " " + (minOk ? "&#10003;" : "&#9888; below min") + "</span>" : "") +
    (isPending ? "<span class='pending-badge'>&#x1F550; PO pending</span>" : "");

  const clearBtn = document.querySelector("#vpmClearPendingButton");
  if (clearBtn) {
    clearBtn.hidden = !isPending;
    clearBtn.onclick = function() { clearVendorPending(vendorName); modal.hidden = true; };
  }
  const submitBtn = document.querySelector("#vpmSubmitButton");
  if (submitBtn) submitBtn.onclick = function() { submitVendorPo(vendorName); modal.hidden = true; };
  const pdfBtn = document.querySelector("#vpmPdfButton");
  if (pdfBtn) pdfBtn.onclick = function() { exportVendorPoPdf(vendorName, items); };

  const body = document.querySelector("#vpmBody");
  if (body) {
    body.innerHTML = items.map((sku) => {
      const pend = isPendingOrder(sku.code);
      return "<tr data-vpm-code='" + sku.code + "'>" +
        "<td style='color:#2470c4;font-weight:700;white-space:nowrap'>" + escapeHtml(sku.code) + "</td>" +
        "<td class='sku-name'>" + escapeHtml(sku.product) + "</td>" +
        "<td class='num " + ((sku.stock||0)<0?"entry-negative":"") + "'>" + number.format(sku.stock||0) + "</td>" +
        "<td class='num'>" + formatVelocity(sku.velocity||0) + "</td>" +
        "<td class='num'>" + number.format(sku.reorderMin||0) + "</td>" +
        "<td class='num'>" + number.format(sku.reorderMax||0) + "</td>" +
        "<td class='num'>" + number.format(sku.recommendedOrder||0) + "</td>" +
        "<td class='num order-highlight'><input type='number' class='vpm-qty-input mini-input' data-code='" + sku.code + "' data-unit-cost='" + (sku.unitCost||0) + "' value='" + (sku.caseOrder||0) + "' min='0' style='width:4rem;text-align:center;font-weight:700' /></td>" +
        "<td class='num'>" + number.format(sku.caseSize||1) + "</td>" +
        "<td class='num'>" + currency.format(sku.unitCost||0) + "</td>" +
        "<td class='num vpm-line-cost' data-unit-cost='" + (sku.unitCost||0) + "'>" + currency.format((sku.caseOrder||0)*(sku.unitCost||0)) + "</td>" +
        "<td style='text-align:center'>" + (pend ? "&#x1F550;" : "") + "</td>" +
        "</tr>";
    }).join("") || "<tr><td colspan='12' class='empty-cell'>No items to order for this vendor right now.</td></tr>";

    body.querySelectorAll(".vpm-qty-input").forEach(function(input) {
      input.addEventListener("input", function() {
        var row = input.closest("tr");
        var uc = parseFloat(input.dataset.unitCost || 0);
        var qty = Math.max(0, toNumber(input.value)||0);
        var lineCell = row.querySelector(".vpm-line-cost");
        if (lineCell) lineCell.textContent = currency.format(qty * uc);
        var total = 0;
        body.querySelectorAll(".vpm-qty-input").forEach(function(inp) {
          total += Math.max(0, toNumber(inp.value)||0) * parseFloat(inp.dataset.unitCost||0);
        });
        var dispEl = document.querySelector("#vpmTotalDisplay");
        if (dispEl) dispEl.textContent = currency.format(total);
        var gt = document.querySelector("#vpmGrandTotal");
        if (gt) gt.textContent = "Grand Total: " + currency.format(total);
      });
    });
  }

  var gt = document.querySelector("#vpmGrandTotal");
  if (gt) gt.textContent = "Grand Total: " + currency.format(totalCost);
  state._vaVendorName = vendorName;
  modal.hidden = false;
}

function exportVendorPoPdf(vendorName, items) {
  var today = new Date().toLocaleDateString("en-US", { year:"numeric", month:"long", day:"numeric" });
  var body = document.querySelector("#vpmBody");
  var rows = body ? Array.from(body.querySelectorAll("tr[data-vpm-code]")) : [];
  var grandTotal = 0;
  var rowsHtml = rows.map(function(tr) {
    var qtyInput = tr.querySelector(".vpm-qty-input");
    var qty = Math.max(0, toNumber(qtyInput ? qtyInput.value : 0));
    var uc = parseFloat(qtyInput ? qtyInput.dataset.unitCost : 0);
    var lineCost = qty * uc;
    grandTotal += lineCost;
    var cells = Array.from(tr.cells);
    return "<tr><td>" + (cells[0]?cells[0].textContent.trim():"") + "</td>" +
      "<td>" + (cells[1]?cells[1].textContent.trim():"") + "</td>" +
      "<td class='num'>" + (cells[2]?cells[2].textContent.trim():"") + "</td>" +
      "<td class='num'>" + (cells[3]?cells[3].textContent.trim():"") + "</td>" +
      "<td class='num'>" + (cells[4]?cells[4].textContent.trim():"") + "</td>" +
      "<td class='num'>" + (cells[5]?cells[5].textContent.trim():"") + "</td>" +
      "<td class='num'>" + (cells[6]?cells[6].textContent.trim():"") + "</td>" +
      "<td class='num'><b>" + qty + "</b></td>" +
      "<td class='num'>" + (cells[8]?cells[8].textContent.trim():"") + "</td>" +
      "<td class='num'>" + (cells[9]?cells[9].textContent.trim():"") + "</td>" +
      "<td class='num'><b>" + currency.format(lineCost) + "</b></td></tr>";
  }).join("");
  var html = "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>PO - " + escapeHtml(vendorName) + "</title>" +
    "<style>body{font-family:Arial,sans-serif;font-size:11px;color:#1c2320;padding:20px;margin:0}" +
    "h1{font-size:20px;margin:0 0 4px}.meta{display:flex;gap:14px;flex-wrap:wrap;margin:8px 0 14px;font-size:10px;color:#555}" +
    ".meta span{background:#f0f4f2;padding:3px 8px;border-radius:4px}" +
    "table{width:100%;border-collapse:collapse}th{background:#eef7f0;text-align:left;padding:5px 6px;font-size:9px;text-transform:uppercase;border-bottom:2px solid #dce3df}" +
    "td{padding:4px 6px;border-bottom:1px solid #eee}.num{text-align:right}" +
    ".grand{font-size:14px;font-weight:700;text-align:right;padding:10px 6px;border-top:2px solid #1c2320}" +
    "@media print{body{padding:8px}}</style></head><body>" +
    "<h1>Purchase Order &mdash; " + escapeHtml(vendorName) + "</h1>" +
    "<div class='meta'><span><b>Date</b> " + today + "</span><span><b>Items</b> " + rows.length + "</span><span><b>Total</b> " + currency.format(grandTotal) + "</span></div>" +
    "<table><thead><tr><th>Code</th><th>Item</th><th>Stock</th><th>SV/day</th><th>Min</th><th>Max</th><th>Rec</th><th>Order Qty</th><th>Case</th><th>Unit Cost</th><th>Total</th></tr></thead>" +
    "<tbody>" + rowsHtml + "</tbody></table>" +
    "<div class='grand'>Grand Total: " + currency.format(grandTotal) + "</div>" +
    "</body></html>";
  var win = window.open("", "_blank", "width=900,height=700");
  if (!win) { showToast("Pop-up blocked.", 3000, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(function() { win.print(); }, 400);
}

document.querySelector("#vpmCloseButton") && document.querySelector("#vpmCloseButton").addEventListener("click", function() {
  document.querySelector("#vendorPoModal").hidden = true;
});

function clearVendorPending(vendorName) {
  if (!state.pendingOrders?.length) return;
  state.pendingOrders = state.pendingOrders.map((po) =>
    po.vendor === vendorName ? { ...po, cleared: true, clearedAt: new Date().toISOString() } : po
  );
  savePendingOrders();
  renderOrders();
  showToast(`Pending cleared for ${vendorName}`, 2400, "success");
}

// ── Final count export (called after Submit & Apply) ──────────────────────
function generateFinalCountExport(session, candidates, latestByCode, snapshot) {
  const entries = candidates.map((item) => {
    const key = codeKey(item.code);
    const entry = latestByCode.get(key);
    const qtyStart = snapshot ? (snapshot[key] ?? item.stock ?? 0) : (item.stock ?? 0);
    const qtyEnd = entry ? Number(entry.countedQty || 0) : 0;
    return { code: item.code, product: item.product, vendor: item.vendor||"", category: item.category||"", qtyStart, qty: qtyEnd, variance: qtyEnd - qtyStart, scanned: !!entry };
  });
  const dateStr = new Date().toLocaleDateString("en-US",{year:"numeric",month:"2-digit",day:"2-digit"}).replace(/\//g,"-");
  // Store in state for Reports tab
  if (!state.finalCountReports) state.finalCountReports = [];
  state.finalCountReports.unshift({ sessionId: session.id, label: countSessionLabel(session), date: dateStr, submittedAt: new Date().toISOString(), entries });
  localStorage.setItem("posFinalCountReports:v1", JSON.stringify(state.finalCountReports.slice(0,30)));
  // Offer immediate download
  setTimeout(() => {
    if (confirm("Count submitted! Download the final count export for your POS import?\n\nColumns: Code, QTY (ready for POS import)")) {
      exportFinalCountToExcel(state.finalCountReports[0]);
    }
  }, 400);
}

async function exportFinalCountToExcel(report) {
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available.", 3000, "warning"); return; }
  const wb = xlsx.utils.book_new();
  const data = [
    ["Final Inventory Count - " + report.label, "", "", "", "", report.date],
    [],
    ["Code", "Item", "Vendor", "Category", "Qty Before", "QTY", "Variance", "Scanned"],
    ...report.entries.map((e) => [e.code, e.product, e.vendor, e.category, e.qtyStart, e.qty, e.variance, e.scanned ? "Yes" : "No (0)"]),
  ];
  const ws = xlsx.utils.aoa_to_sheet(data);
  ws["!cols"] = [16,32,14,14,10,10,10,10].map((w)=>({wch:w}));
  // Force code column as text
  data.forEach((row, r) => {
    if (r < 3) return;
    const ref = xlsx.utils.encode_cell({r, c:0});
    if (ws[ref]) { ws[ref].t = "s"; ws[ref].z = "@"; }
  });
  xlsx.utils.book_append_sheet(wb, ws, "Final Count");
  xlsx.writeFile(wb, "FinalCount_" + report.date + ".xlsx");
  showToast("Final count exported for POS import.", 2800, "success");
}

// ── Reports tab — load saved final count reports ───────────────────────────
if (!state.finalCountReports) {
  state.finalCountReports = JSON.parse(localStorage.getItem("posFinalCountReports:v1") || "[]");
}

// ── User/PIN auth system ───────────────────────────────────────────────────
const AUTH_KEY = "posAuthUsers:v1";
const SUPABASE_URL = "https://mqkrgieotabpptsbosdh.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_i0n6EMZnW-E40Dx8Od_8mg_eD9GChBy";
if (!state.authUsers) state.authUsers = [];
if (!state.authUsersLoaded) state.authUsersLoaded = false;

function defaultAuthUsers() {
  return [{ name: "Admin", pin: "0000", role: "admin", id: "default-admin", active: true }];
}

function normalizeAuthUser(user) {
  if (!user) return null;
  return {
    id: String(user.id || ""),
    name: String(user.name || "").trim(),
    pin: String(user.pin || "").trim(),
    role: String(user.role || "user").trim().toLowerCase(),
    active: user.active !== false,
    created_at: user.created_at || null,
  };
}

function loadUsers() {
  if (Array.isArray(state.authUsers) && state.authUsers.length) return state.authUsers;
  const cached = JSON.parse(localStorage.getItem(AUTH_KEY) || "[]")
    .map(normalizeAuthUser)
    .filter((user) => user?.name && user?.pin && user.active !== false);
  if (cached.length) {
    state.authUsers = cached;
    return cached;
  }
  const fallback = defaultAuthUsers();
  state.authUsers = fallback;
  return fallback;
}

function saveUsers(users) {
  const normalized = (users || []).map(normalizeAuthUser).filter((user) => user?.name && user?.pin);
  state.authUsers = normalized;
  localStorage.setItem(AUTH_KEY, JSON.stringify(normalized));
}

function supabaseHeaders(preferReturn = false) {
  const headers = {
    apikey: SUPABASE_PUBLISHABLE_KEY,
    Authorization: `Bearer ${SUPABASE_PUBLISHABLE_KEY}`,
    "Content-Type": "application/json",
  };
  if (preferReturn) headers.Prefer = "return=representation";
  return headers;
}

async function supabaseFetchUsers() {
  const url = new URL("/rest/v1/app_users", SUPABASE_URL);
  url.searchParams.set("select", "id,name,pin,role,active,created_at");
  url.searchParams.set("active", "eq.true");
  url.searchParams.set("order", "created_at.asc");
  const response = await fetch(url.toString(), { headers: supabaseHeaders() });
  if (!response.ok) throw new Error(`Supabase read failed (${response.status})`);
  const rows = await response.json();
  return rows.map(normalizeAuthUser).filter((user) => user?.name && user?.pin && user.active !== false);
}

async function refreshUsersFromSupabase(options = {}) {
  const { silent = false } = options;
  try {
    const users = await supabaseFetchUsers();
    if (users.length) saveUsers(users);
    state.authUsersLoaded = true;
    return loadUsers();
  } catch (error) {
    state.authUsersLoaded = true;
    if (!silent) showToast("Using saved login users — Supabase sync unavailable.", 3200, "warning");
    return loadUsers();
  }
}

async function supabaseInsertUser(user) {
  const response = await fetch(`${SUPABASE_URL}/rest/v1/app_users`, {
    method: "POST",
    headers: supabaseHeaders(true),
    body: JSON.stringify([{
      name: user.name,
      pin: user.pin,
      role: user.role,
      active: user.active !== false,
    }]),
  });
  if (!response.ok) throw new Error(`Supabase insert failed (${response.status})`);
  return response.json();
}

async function supabaseUpdateUser(userId, patch) {
  const url = new URL("/rest/v1/app_users", SUPABASE_URL);
  url.searchParams.set("id", `eq.${userId}`);
  const response = await fetch(url.toString(), {
    method: "PATCH",
    headers: supabaseHeaders(true),
    body: JSON.stringify(patch),
  });
  if (!response.ok) throw new Error(`Supabase update failed (${response.status})`);
  return response.json();
}

async function supabaseDeleteUser(userId) {
  const url = new URL("/rest/v1/app_users", SUPABASE_URL);
  url.searchParams.set("id", `eq.${userId}`);
  const response = await fetch(url.toString(), {
    method: "DELETE",
    headers: supabaseHeaders(),
  });
  if (!response.ok) throw new Error(`Supabase delete failed (${response.status})`);
}

function verifyPin(pin) {
  const users = loadUsers();
  return users.find((u) => u.pin === pin && u.active !== false) || null;
}
if (!state.currentUser) state.currentUser = null;
function isAdmin() {
  return !state.currentUser || state.currentUser.role === "admin";
}
function isUserRole() {
  return state.currentUser && state.currentUser.role === "user";
}
if (!state.authRequired) state.authRequired = true;

// ── Render Settings tab ─────────────────────────────────────────────────────
function renderSettings() {
  const panel = document.querySelector("#settingsPanel");
  if (!panel) return;
  const users = loadUsers();
  panel.innerHTML = users.map((u) => `
    <tr>
      <td>${escapeHtml(u.name)}</td>
      <td><span class="badge ${u.role === "admin" ? "state-active" : "state-disabled"}">${u.role.toUpperCase()}</span></td>
      <td>
        <span class="pin-dots" data-user-id="${escapeHtml(u.id)}">&#9679;&#9679;&#9679;&#9679;</span>
        <button type="button" class="secondary-button settings-pin-reveal" data-user-id="${escapeHtml(u.id)}" style="padding:2px 6px;font-size:.7rem">Show</button>
      </td>
      <td>
        <button type="button" class="secondary-button settings-change-pin" data-user-id="${escapeHtml(u.id)}">Change PIN</button>
        ${u.id !== "default-admin" ? `<button type="button" class="delete-session-btn settings-delete-user" data-user-id="${escapeHtml(u.id)}" style="padding:3px 7px;font-size:.75rem">Del</button>` : ""}
      </td>
    </tr>`).join("");

  panel.querySelectorAll(".settings-pin-reveal").forEach((btn) => {
    btn.addEventListener("click", () => {
      const u = loadUsers().find((u) => u.id === btn.dataset.userId);
      if (!u) return;
      const span = panel.querySelector(`.pin-dots[data-user-id="${btn.dataset.userId}"]`);
      if (span) span.textContent = span.textContent.includes("\u25CF") ? u.pin : "\u25CF\u25CF\u25CF\u25CF";
      btn.textContent = btn.textContent === "Show" ? "Hide" : "Show";
    });
  });
  panel.querySelectorAll(".settings-change-pin").forEach((btn) => {
    btn.addEventListener("click", async () => {
      const pin = prompt("Enter new 4-digit PIN:");
      if (!pin || !/^\d{4}$/.test(pin)) { showToast("PIN must be exactly 4 digits.", 3000, "warning"); return; }
      try {
        await supabaseUpdateUser(btn.dataset.userId, { pin });
        await refreshUsersFromSupabase({ silent: true });
        renderSettings();
        showToast("PIN updated.", 2000, "success");
      } catch (error) {
        showToast("Could not update PIN in Supabase.", 3200, "warning");
      }
    });
  });
  panel.querySelectorAll(".settings-delete-user").forEach((btn) => {
    btn.addEventListener("click", async () => {
      if (!confirm("Delete this user?")) return;
      try {
        await supabaseDeleteUser(btn.dataset.userId);
        await refreshUsersFromSupabase({ silent: true });
        renderSettings();
        showToast("User deleted.", 2000, "warning");
      } catch (error) {
        showToast("Could not delete user from Supabase.", 3200, "warning");
      }
    });
  });
}

async function addSettingsUser() {
  const name = document.querySelector("#settingsNewName")?.value.trim();
  const pin = document.querySelector("#settingsNewPin")?.value.trim();
  const role = (document.querySelector("#settingsNewRole")?.value || "user").toLowerCase();
  if (!name) { showToast("Enter a user name.", 3000, "warning"); return; }
  if (!pin || !/^\d{4}$/.test(pin)) { showToast("PIN must be 4 digits.", 3000, "warning"); return; }
  try {
    await supabaseInsertUser({ name, pin, role, active: true });
    if (document.querySelector("#settingsNewName")) document.querySelector("#settingsNewName").value = "";
    if (document.querySelector("#settingsNewPin")) document.querySelector("#settingsNewPin").value = "";
    await refreshUsersFromSupabase({ silent: true });
    renderSettings();
    showToast("User added.", 2000, "success");
  } catch (error) {
    showToast("Could not add user to Supabase.", 3200, "warning");
  }
}

// ── Lock screen ────────────────────────────────────────────────────────────
function showLockScreen() {
  const overlay = document.querySelector("#lockScreen");
  if (overlay) overlay.classList.remove("lock-dismissed");
  state.currentUser = null;
}

function tryUnlock(pin) {
  const user = verifyPin(pin);
  if (!user) {
    const disp = document.querySelector("#lockPinDisplay");
    if (disp) { disp.textContent = "Wrong PIN"; disp.style.color = "#c0392b"; setTimeout(() => { disp.textContent = ""; disp.style.color = ""; }, 1000); }
    return;
  }
  state.currentUser = user;
  const overlay = document.querySelector("#lockScreen");
  if (overlay) overlay.classList.add("lock-dismissed");
  showToast("Welcome, " + user.name + "!", 2000, "success");
  applyRoleRestrictions();
  // Switch to Products tab for basic users
  if (isUserRole()) { const btn = document.querySelector('.tab-button[data-tab="inventory"]'); if (btn) btn.click(); }
}

// ── Render ordering vendor filter ──────────────────────────────────────────
function getOrderVendorFilter() {
  return document.querySelector("#orderVendorFilterSelect")?.value || state.orderVendorFilter || "Active";
}

function filterOrderVendors(vendorName) {
  const filter = getOrderVendorFilter();
  if (filter === "") return true; // All
  const rule = state.vendorRules.find((r) => r.vendor && r.vendor.toUpperCase() === vendorName.toUpperCase());
  if (!rule) return filter !== "Disabled"; // Unruled vendors: show unless filtering to Disabled-only
  return rule.status === filter;
}

// ── PO Modal: case order auto-adjusts when rec order changes ──────────────
function calcCaseOrder(recOrder, caseSize) {
  if (!caseSize || caseSize <= 1) return recOrder;
  return Math.ceil(recOrder / caseSize) * caseSize;
}


// ── Reports tab box buttons ─────────────────────────────────────────────────
document.querySelector("#reportBoxAdjust")?.addEventListener("click", () => {
  document.querySelector("#reportAdjustModal").hidden = false;
  renderAdjustLog();
});
document.querySelector("#reportBoxPO")?.addEventListener("click", () => {
  document.querySelector("#reportPoModal").hidden = false;
  renderPoHistory();
});
document.querySelector("#reportBoxCount")?.addEventListener("click", () => {
  document.querySelector("#reportCountModal").hidden = false;
  renderCountReportHistory();
});

// Generic close buttons for report modals
document.querySelectorAll(".report-modal-close").forEach((btn) => {
  btn.addEventListener("click", () => {
    const id = btn.dataset.closeModal;
    if (id) document.querySelector("#" + id).hidden = true;
  });
});

// ── Report modal Esc + click-outside ────────────────────────────────────────
["reportAdjustModal","reportPoModal","reportCountModal"].forEach((id) => {
  const el = document.querySelector("#" + id);
  if (!el) return;
  el.addEventListener("click", (e) => { if (e.target === el) el.hidden = true; });
});

// ── PO history rendering ─────────────────────────────────────────────────────
function renderPoHistory() {
  const body = document.querySelector("#poHistoryBody");
  if (!body) return;
  const pos = (state.pendingOrders||[]).filter((po) => po.submittedAt);
  if (!pos.length) { body.innerHTML = '<tr><td colspan="5" class="empty-cell">No POs submitted yet.</td></tr>'; return; }
  body.innerHTML = pos.sort((a,b) => b.submittedAt.localeCompare(a.submittedAt)).map((po) => `<tr>
    <td>${new Date(po.submittedAt).toLocaleDateString()}</td>
    <td><b>${escapeHtml(po.vendor)}</b></td>
    <td>${(po.codes||[]).length}</td>
    <td>-</td>
    <td>${po.cleared ? '<span class="state-badge state-disabled">Cleared</span>' : '<span class="pending-badge">Pending</span>'}</td>
  </tr>`).join("");
}

// ── Count report history rendering ──────────────────────────────────────────
function renderCountReportHistory() {
  const body = document.querySelector("#countReportHistoryBody");
  if (!body) return;
  const reports = state.finalCountReports||[];
  if (!reports.length) { body.innerHTML = '<tr><td colspan="5" class="empty-cell">No final counts saved yet.</td></tr>'; return; }
  body.innerHTML = reports.map((r, i) => `<tr>
    <td>${escapeHtml(r.date)}</td>
    <td>${escapeHtml(r.label)}</td>
    <td>${(r.entries||[]).length}</td>
    <td>${new Date(r.submittedAt).toLocaleString()}</td>
    <td><button type="button" class="secondary-button" onclick="exportFinalCountToExcel(state.finalCountReports[${i}])">Excel</button></td>
  </tr>`).join("");
}

// ── Ordering vendor filter ────────────────────────────────────────────────────
document.querySelector("#orderVendorFilterSelect")?.addEventListener("change", () => {
  state.orderVendorFilter = document.querySelector("#orderVendorFilterSelect").value;
  renderOrders();
});

// ── Report Esc key support ────────────────────────────────────────────────────


// ── Lock screen — no lag, visual feedback, keyboard + click ─────────────────
(function() {
  const overlay = document.querySelector("#lockScreen");
  if (!overlay) return;
  if (loadUsers().length > 0) overlay.classList.remove("lock-dismissed");
  else { overlay.classList.add("lock-dismissed"); return; }

  let pin = "";

  function draw() {
    const disp = document.querySelector("#lockPinDisplay");
    if (!disp) return;
    // Show filled green dots as each digit is entered
    const dots = [];
    for (let i = 0; i < 4; i++) {
      dots.push(i < pin.length ? "\u25CF" : "\u25CB");
    }
    disp.textContent = dots.join("  ");
    disp.style.color = pin.length === 4 ? "var(--green,#16835b)" : "";
    disp.style.borderColor = pin.length > 0 ? "var(--green,#16835b)" : "";
  }

  function press(k) {
    if (k === "clear") { pin = ""; draw(); return; }
    if (k === "del" || k === "Backspace") { pin = pin.slice(0, -1); draw(); return; }
    if (/^[0-9]$/.test(k) && pin.length < 4) {
      pin += k; draw();
      if (pin.length === 4) {
        const p = pin; pin = ""; draw();
        tryUnlock(p);
      }
    }
  }

  // Button clicks — mousedown not click to remove any delay
  overlay.addEventListener("mousedown", (e) => {
    const btn = e.target.closest("[data-lock-key]");
    if (!btn) return;
    e.preventDefault(); e.stopPropagation();
    btn.classList.add("lock-key--pressed");
    setTimeout(() => btn.classList.remove("lock-key--pressed"), 150);
    press(btn.dataset.lockKey);
  });

  // Keyboard — capture phase, fires before other handlers
  document.addEventListener("keydown", (e) => {
    if (overlay.classList.contains("lock-dismissed")) return;
    if (e.key !== "Tab" && e.key !== "F5" && e.key !== "F12") {
      e.preventDefault(); e.stopImmediatePropagation();
    }
    press(e.key);
  }, true);

  draw();
})();
initApp();
