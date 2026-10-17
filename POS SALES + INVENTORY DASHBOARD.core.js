const state = {
  rawSales: [],
  inventories: new Map(),
  latestInventory: new Map(),
  excelItems: new Map(),
  excelByPlu: new Map(),
  excelByItemNumber: new Map(),
  reorderOverrides: JSON.parse(localStorage.getItem("posDashboardReorderOverrides:v1") || "{}"),
  itemMeta: {},
  // v4 starts with a compact Products view; users can re-enable detail columns.
  visibleColumns: JSON.parse(localStorage.getItem("posDashboardVisibleColumns:v4") || "null"),
  columnOrder: JSON.parse(localStorage.getItem("posDashboardColumnOrder:v3") || "null"),
  // Reset v1 widths once; those values were captured while CSS tracks conflicted.
  columnWidths: JSON.parse(localStorage.getItem("posDashboardColumnWidths:v2") || "{}"),
  arrangeColumns: false,
  resizeColumns: false,
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
  // Performance cache - invalidated on data load, valid across filter/search changes
  _skuCache: null,      // full unfiltered SKU map keyed by codeKey
  _skuCacheStamp: 0,    // increments when raw data changes
  _dataCacheStamp: 0,   // stamp written on load; cache valid when stamps match
  _salesIndex: null,
  _salesIndexStamp: 0,
  _salesWindowsCache: new Map(),
  _dailyTotals: new Map(),
  _inventoryRowIndex: new Map(),
  _localProductMetaEdits: new Map(),
  _localImportHoldUntil: 0,
  _countSyncLoading: false,
  _countRemoteLoaded: false,
  _countSyncUnavailable: false,
  _countLastSyncAt: "",
  _countSyncInFlight: null,
  _countSessionOpen: false,
  countSoundEnabled: localStorage.getItem("posDashboardCountSoundEnabled:v1") !== "false",
  _countAudioContext: null,
  _countConfirmAudio: null,
  _deletedCountSessionIds: new Set(JSON.parse(localStorage.getItem("posDashboardDeletedCountSessionIds:v1") || "[]")),
  _priceCheckExactIndex: null,
  _priceCheckExactIndexStamp: 0,
  _multiBarcodeIndexStamp: 0,
  scanModeActive: false,
  _deferredSharedSync: false,
  _deferredDashboardOptions: false,
  _filteredSkuIndex: new Map(),
  _loadedFileSignatures: new Set(),
  _activeDetailCode: "",
  _activeTabRenderHandle: 0,
  _activeTabRenderLastAt: 0,
  _activeTabRenderRequested: false,
  _activeTabRenderRunning: false,
  _productsOpened: false,
  _initialRestoreComplete: false,
  _scanIndexBuildQueued: false,
  _dashboardDetailsReady: false,
  _dashboardDetailScheduled: false,
  _dashboardDetailToken: 0,
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
  orderColumnOrder: JSON.parse(localStorage.getItem("posOrderColumnOrder:v1") || "null"),
  activeCountSession: JSON.parse(localStorage.getItem("posDashboardActiveCountSession:v1") || "null"),
  countQtyBuffer: "0",
  selectedCountItemCode: "",
  countStage: "search",
  pendingDuplicateCount: null,
  pendingDuplicateMode: null,
  countReportMode: "input",
  metricsPinned: JSON.parse(localStorage.getItem("posDashboardMetricsPinned:v1") || "false"),
  orderVendorQuickFilter: "",
  orderSubmissionVendors: [],
  orderSubmissionActiveVendor: "",
  orderSubmissionDrafts: {},
  orderArrangeColumns: false,
  orderArrangeSource: "",
  orderDismissedVendors: JSON.parse(localStorage.getItem("posDashboardDismissedOrderVendors:v1") || "[]"),
  productPoReviewVendor: "",
  productPoReviewSort: "rec-desc",
  vendorPoSort: "item-asc",
  uploadLogs: JSON.parse(localStorage.getItem("posDashboardUploadLogs:v1") || "[]"),
  poBackorderCounts: JSON.parse(localStorage.getItem("posDashboardPoBackorderCounts:v1") || "{}"),
  multiBarcodeMap: JSON.parse(localStorage.getItem("posDashboardMultiBarcodeMap:v1") || "{}"),
  multiBarcodeMasters: JSON.parse(localStorage.getItem("posDashboardMultiBarcodeMasters:v1") || "[]"),
  multiBarcodeFileName: localStorage.getItem("posDashboardMultiBarcodeFileName:v1") || "",
  multiBarcodeSearch: "",
  manualMultiBarcodes: JSON.parse(localStorage.getItem("posDashboardManualMultiBarcodes:v1") || "{}"),
  detailDrawerTab: "fields",
  newItemsVendor: "",
  newItemsDepartment: "",
  newItemsCategory: "",
  newItemsYear: "",
  newItemsSearch: "",
  newItemsDormantMonths: "",
  // Stock adjust modal state
  stockAdjustItem: null,
  stockAdjustAction: null,   // "add" | "remove" | "set"
  stockAdjustQtyBuffer: "0",
  pendingDeleteSessionId: null,
  pendingSubmitSessionId: null,
  priceCheckStream: null,
  priceCheckRaf: 0,
  priceCheckDetector: null,
  priceCheckReader: null,
  priceCheckReaderControls: null,
  priceCheckScanner: null,
  priceCheckFullscreen: false,
  priceCheckTorchOn: false,
  priceCheckScanTimer: 0,
  priceCheckLastCode: "",
  priceCheckLastScanAt: 0,
};

const ENABLE_CUSTOM_PARENT_RULES = true;
const ENABLE_CUSTOM_ATTRIBUTE_RULES = false;
const ENABLE_SHARED_SYNC = true;
const DEFAULT_ORDER_SAFETY_DAYS = 21;
const DEFAULT_ORDER_DAYS_OF_INVENTORY = 7;
const COUNT_REMOTE_ACTIVE_HISTORY_TTL_MS = 12 * 60 * 60 * 1000;
const COUNT_SYNC_RETRY_MS = 8000;
const COUNT_DEVICE_ID_KEY = "posDashboardCountDeviceId:v1";
const DEFAULT_PO_COPY_EMAIL = "vinnieoh86@gmail.com";
const PO_EMAIL_FUNCTION_NAME = "send-po-email";
const pendingInventoryRefreshTimers = new Map();
let sharedSyncTimer = 0;
let sharedVendorRulesTimer = 0;
let appInitPromise = null;
let appInitDone = false;
let lastLocalSharedSyncAt = "";
let pendingSharedProductCodes = new Set();
let sharedVendorRulesAvailable = true;
let sharedProductMetaAvailable = true;
let sharedImportLogsAvailable = true;
let sharedCountSessionsAvailable = true;
const SHARED_SYNC_POLL_MS = 2000;
const SHARED_SYNC_INITIAL_DELAY_MS = 1200;
const SHARED_PRODUCT_META_PULL_LIMIT = 5000;

function performanceNow() {
  return typeof performance !== "undefined" && typeof performance.now === "function"
    ? performance.now()
    : Date.now();
}

function logPerformanceEnd(label, startedAt, details = {}) {
  const ms = Number((performanceNow() - startedAt).toFixed(2));
  console.info(`[AppPerf] ${label} end`, { ms, ...details });
}

// Scan mode gets a small capability profile so weaker browsers do less work while scanning.
function createScanPerformanceProfile() {
  const ua = navigator.userAgent || "";
  const cores = Number(navigator.hardwareConcurrency || 0);
  const memory = Number(navigator.deviceMemory || 0);
  const isSilk = /silk|kindle|kf[a-z]{2,}/i.test(ua);
  const isMobile = /android|iphone|ipad|ipod|mobile/i.test(ua);
  const weakHardware = (!!cores && cores <= 2) || (!!memory && memory <= 2);
  const tier = isSilk || weakHardware ? "low" : (isMobile || (!!cores && cores <= 4) || (!!memory && memory <= 4) ? "medium" : "full");
  const profiles = {
    full: {
      camera: { width: 1280, height: 720, fps: 24, qrbox: { width: 260, height: 130 } },
      detectIntervalMs: 42,
      normalPollMs: SHARED_SYNC_POLL_MS,
      scanPollMs: SHARED_SYNC_POLL_MS,
      // PERF-B: countPollMs — sync poll interval while Count Mode is active.
      // Full-tier (desktop): same as normal, counts are interactive and changes should appear quickly.
      countPollMs: 4000,
      deferRestoresWhileScanning: false,
    },
    medium: {
      camera: { width: 960, height: 540, fps: 15, qrbox: { width: 230, height: 115 } },
      detectIntervalMs: 67,
      normalPollMs: SHARED_SYNC_POLL_MS,
      scanPollMs: 12000,
      // Medium (phone/tablet): 10s between count syncs — responsive without hammering the device.
      countPollMs: 10000,
      deferRestoresWhileScanning: true,
    },
    low: {
      camera: { width: 640, height: 480, fps: 10, qrbox: { width: 200, height: 100 } },
      detectIntervalMs: 100,
      normalPollMs: SHARED_SYNC_POLL_MS,
      scanPollMs: 45000,
      // Low (Kindle/Silk): 30s between count syncs — keeps Kindle responsive while counting.
      countPollMs: 30000,
      deferRestoresWhileScanning: true,
    },
  };
  const profile = { tier, isSilk, isMobile, cores, memory, ...profiles[tier] };
  console.info("[ScanPerf] profile", { tier, isSilk, isMobile, cores: cores || "unknown", memory: memory || "unknown" });
  return profile;
}

const scanPerformanceProfile = createScanPerformanceProfile();

function scanModeIsActive() {
  return state.scanModeActive || activeTabName() === "scanmode";
}

function setScanLookupStatus(message, readiness = "") {
  if (!els?.scanModeStatus) return;
  els.scanModeStatus.textContent = message;
  if (readiness) els.scanModeStatus.dataset.readiness = readiness;
  else delete els.scanModeStatus.dataset.readiness;
}

// Scanner lookup is prioritized ahead of Products DOM work on touch devices.
function queueScanExactIndexBuild(options = {}) {
  const isScanActive = scanModeIsActive();
  if (!isScanActive && options.prioritize !== true) return;
  if (!state._initialRestoreComplete && !state.latestInventory.size && !state.excelItems.size) {
    if (isScanActive) setScanLookupStatus("Loading item lookup...", "loading");
    return;
  }
  if (!state.latestInventory.size && !state.excelItems.size) {
    if (isScanActive) setScanLookupStatus("Offline or not ready", "offline");
    return;
  }
  const currentStamp = `${state._dataCacheStamp}:${state._multiBarcodeIndexStamp || 0}`;
  if (state._priceCheckExactIndex && state._priceCheckExactIndexStamp === currentStamp) {
    if (isScanActive) setScanLookupStatus("Scanner ready", "ready");
    return;
  }
  if (state._scanIndexBuildQueued) return;
  if (isScanActive) setScanLookupStatus("Loading item lookup...", "loading");
  state._scanIndexBuildQueued = true;
  const build = () => {
    state._scanIndexBuildQueued = false;
    if (!scanModeIsActive() && options.prioritize !== true) return;
    ensurePriceCheckExactIndex();
    if (scanModeIsActive() && els.scanModeStatus?.dataset.readiness === "loading") {
      setScanLookupStatus("Scanner ready", "ready");
    }
  };
  if (isScanActive || options.immediate === true) build();
  else setTimeout(build, 0);
}

function setScanPerformanceMode(active) {
  state.scanModeActive = !!active;
  document.body?.classList.toggle("scan-mode-active", state.scanModeActive);
  ["full", "medium", "low"].forEach((tier) => {
    document.body?.classList.toggle(`scan-performance-${tier}`, state.scanModeActive && scanPerformanceProfile.tier === tier);
  });
  if (state.scanModeActive) {
    hideHoverTooltip();
    queueScanExactIndexBuild();
  } else {
    if (state._deferredDashboardOptions) {
      state._deferredDashboardOptions = false;
      updateFilterOptions();
      updateInventoryStateFilter();
      setDefaultDates();
    }
    if (state._deferredSharedSync) window.flushDeferredSharedSync?.();
  }
}

// Increment this whenever raw data changes to bust the SKU cache
function bumpDataStamp() {
  // After data loads, apply the 90D default if it hasn't been overridden
  clearTimeout(state._applyPresetTimer);
  state._applyPresetTimer = setTimeout(() => {
    // Only auto-apply 90D if it hasn't been manually changed by the user
    if (state.activePresetDays === 90 && state.dates.length && !scanModeIsActive()) {
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
  state._priceCheckExactIndex = null;
  state._priceCheckExactIndexStamp = 0;
  state._priceCheckRowsCache = null;
  state._priceCheckRowsStamp = 0;
  state._countSearchIndex = null; // invalidate count search index too
  state._countSetupCached = false; // FIX-D: invalidate setup dropdown cache on data load
  state._dashboardDetailsReady = false;
  state._dashboardDetailScheduled = false;
  state._dashboardDetailToken += 1;
  if (scanModeIsActive() && state._initialRestoreComplete) queueScanExactIndexBuild();
}

function clearInventorySelection() {
  if (!state.selectedInventoryCodes?.size) return;
  state.selectedInventoryCodes.clear();
  document.querySelectorAll("#inventoryBody .row-checkbox, #selectAllInventory").forEach((node) => {
    node.checked = false;
  });
}

// Debounce helper - delays fast-typing renders
function debounce(fn, ms) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), ms);
  };
}

function localIsoDate(date = new Date()) {
  const d = date instanceof Date ? date : new Date(date);
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

const DB_NAME = "posDashboardHistory_launch421";
const DB_VERSION = 1;
const DB_STORE = "app";
const DB_KEY = "state_v2";
const DB_ITEM_META_KEY = "item_meta_v3";
const LOCAL_SNAPSHOT_KEY = "posDashboardPersistedState:v2";
const IDLE_TIMEOUT_MS = 5 * 60 * 1000;
const ITEM_META_STORAGE_KEY = "posDashboardItemMeta:v1";

try {
  const legacySnapshot = localStorage.getItem(LOCAL_SNAPSHOT_KEY);
  if (legacySnapshot && legacySnapshot.length > 50000) localStorage.removeItem(LOCAL_SNAPSHOT_KEY);
  const legacyItemMeta = localStorage.getItem(ITEM_META_STORAGE_KEY);
  if (legacyItemMeta && legacyItemMeta.length > 50000) localStorage.removeItem(ITEM_META_STORAGE_KEY);
} catch (_) {
  // Ignore storage access errors on startup.
}

const els = {
  fileInput: document.querySelector("#fileInput"),
  folderInput: document.querySelector("#folderInput"),
  excelInput: document.querySelector("#excelInput"),
  dropZone: document.querySelector("#dropZone"),
  fileCount: document.querySelector("#fileCount"),
  dateCoverage: document.querySelector("#dateCoverage"),
  excelStatus: document.querySelector("#excelStatus"),
  searchInput: document.querySelector("#searchInput"),
  priceCheckSearchInput: document.querySelector("#priceCheckSearchInput"),
  priceCheckDropdown: document.querySelector("#priceCheckDropdown"),
  priceCheckSearchButton: document.querySelector("#priceCheckSearchButton"),
  priceCheckManualButton: document.querySelector("#priceCheckManualButton"),
  priceCheckClearButton: document.querySelector("#priceCheckClearButton"),
  priceCheckCameraButton: document.querySelector("#priceCheckCameraButton"),
  priceCheckStopButton: document.querySelector("#priceCheckStopButton"),
  priceCheckTorchButton: document.querySelector("#priceCheckTorchButton"),
  priceCheckVideo: document.querySelector("#priceCheckVideo"),
  priceCheckScanner: document.querySelector("#priceCheckScanner"),
  priceCheckOverlay: document.querySelector("#priceCheckOverlay"),
  priceCheckOverlayVideo: document.querySelector("#priceCheckOverlayVideo"),
  priceCheckOverlayScanner: document.querySelector("#priceCheckOverlayScanner"),
  priceCheckOverlayClose: document.querySelector("#priceCheckOverlayClose"),
  priceCheckOverlayTorchButton: document.querySelector("#priceCheckOverlayTorchButton"),
  priceCheckStatus: document.querySelector("#priceCheckStatus"),
  priceCheckResult: document.querySelector("#priceCheckResult"),
  scanModeStartButton: document.querySelector("#scanModeStartButton"),
  scanModeManualButton: document.querySelector("#scanModeManualButton"),
  scanModeStatus: document.querySelector("#scanModeStatus"),
  scanModeResult: document.querySelector("#scanModeResult"),
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
  inventoryQuickFilter: document.querySelector("#inventoryQuickFilter"),
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
  appLoadStatus: document.querySelector("#appLoadStatus"),
  columnPickerPanel: document.querySelector("#columnPickerPanel"),
  inventoryColumnPicker: document.querySelector("#inventoryColumnPicker"),
  sortMode: document.querySelector("#sortMode"),
  inventoryStateFilter: document.querySelector("#inventoryStateFilter"),
  arrangeColumnsButton: document.querySelector("#arrangeColumnsButton"),
  downloadOrder: document.querySelector("#downloadOrder"),
  downloadSku: document.querySelector("#downloadSku"),
  downloadInventory: document.querySelector("#downloadInventory"),
  downloadParents: document.querySelector("#downloadParents"),
  uploadLogBody: document.querySelector("#uploadLogBody"),
  uploadLogSummary: document.querySelector("#uploadLogSummary"),
  productPoReviewModal: document.querySelector("#productPoReviewModal"),
  productPoReviewTitle: document.querySelector("#productPoReviewTitle"),
  productPoReviewMeta: document.querySelector("#productPoReviewMeta"),
  productPoReviewPills: document.querySelector("#productPoReviewPills"),
  productPoReviewBody: document.querySelector("#productPoReviewBody"),
  multiBarcodeInput: document.querySelector("#multiBarcodeInput"),
  multiSummary: document.querySelector("#multiSummary"),
  multiSearchInput: document.querySelector("#multiSearchInput"),
  multiBody: document.querySelector("#multiBody"),
  newItemsMonthFilter: document.querySelector("#newItemsMonthFilter"),
  newItemsYearFilter: document.querySelector("#newItemsYearFilter"),
  newItemsSummary: document.querySelector("#newItemsSummary"),
  newItemsSearchInput: document.querySelector("#newItemsSearchInput"),
  newItemsVendorFilter: document.querySelector("#newItemsVendorFilter"),
  newItemsDepartmentFilter: document.querySelector("#newItemsDepartmentFilter"),
  newItemsCategoryFilter: document.querySelector("#newItemsCategoryFilter"),
  newItemsDormantFilter: document.querySelector("#newItemsDormantFilter"),
  newItemsClearButton: document.querySelector("#newItemsClearButton"),
  newItemsBody: document.querySelector("#newItemsBody"),
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
  countLaunchTitle: document.querySelector("#countLaunchTitle"),
  countLaunchDescription: document.querySelector("#countLaunchDescription"),
  countLaunchState: document.querySelector("#countLaunchState"),
  // V2.0 split launch cards
  countStartNewCard: document.querySelector("#countStartNewCard"),
  countContinueCard: document.querySelector("#countContinueCard"),
  countContinueCardTitle: document.querySelector("#countContinueCardTitle"),
  countContinueCardDesc: document.querySelector("#countContinueCardDesc"),
  // V2.0 two-column live review panel
  countReviewPanel: document.querySelector("#countReviewPanel"),
  countReviewBody: document.querySelector("#countReviewBody"),
  countReviewBadges: document.querySelector("#countReviewBadges"),
  countReviewFilters: document.querySelector("#countReviewFilters"),
  countReviewTable: document.querySelector("#countReviewTable"),
  countSyncStatus: document.querySelector("#countSyncStatus"),
  countSummaryStrip: document.querySelector("#countSummaryStrip"),
  activeCountTitle: document.querySelector("#activeCountTitle"),
  activeCountMeta: document.querySelector("#activeCountMeta"),
  activeCountSyncStatus: document.querySelector("#activeCountSyncStatus"),
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
  countAutoPlusButton: document.querySelector("#countAutoPlusButton"),
  countUndoAutoButton: document.querySelector("#countUndoAutoButton"),
  countScopeSearchInput: document.querySelector("#countScopeSearchInput"),
  countSelectedItem: document.querySelector("#countSelectedItem"),
  countQuantityDisplay: document.querySelector("#countQuantityDisplay"),
  countKeypadPanel: document.querySelector("#countSessionModal .count-panel--keypad"),
  countKeyButtons: document.querySelectorAll("[data-count-key]"),
  countSetupModal: document.querySelector("#countSetupModal"),
  countDateInput: document.querySelector("#countDateInput"),
  countVendorInput: document.querySelector("#countVendorInput"),
  countCategoryInput: document.querySelector("#countCategoryInput"),
  countAllowOutOfScopeInput: document.querySelector("#countAllowOutOfScopeInput"),
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

[els.safetyDays, els.daysOfInventory].filter(Boolean).forEach((input) => {
  const label = input.closest("label");
  if (label) {
    label.hidden = true;
    label.style.display = "none";
  }
});

const inventoryColumns = [
  ["code", "Code"],
  ["product", "Item"],
  ["plu", "PLU"],
  ["rule", "SD/DOI"],
  ["itemNumber", "Item #"],
  ["sizeAttr", "Sub"],
  ["subType", "Type"],
  ["containerAttr", "Tag"],
  ["category", "Category"],
  ["vendor", "Vendor"],
  ["state", "State"],
  ["addDate", "DATE"],
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


const PRODUCT_TABLE_HIDDEN_BY_DEFAULT = new Set(["itemNumber", "sizeAttr", "subType", "containerAttr", "inventoryCost"]);
function inventoryVisibleColumnKeys() {
  return (state.columnOrder || inventoryColumns.map(([key]) => key))
    .filter((key) => inventoryColumns.some(([colKey]) => colKey === key))
    .filter((key) => state.visibleColumns?.[key] !== false);
}
function enforceCompactProductColumns() {
  // Products tab should match the real inventory CSV fields first: CODE, ITEM, PLU,
  // SD/DOI, CATEGORY, VENDOR, STATE, DATE, STOCK, SOLD, SV/DAY, COST, PRICE.
  // The parsed style/detail columns are still available from Columns, but keeping
  // them hidden prevents the header/body offset that made Category/Vendor/Stock
  // look like they belonged over the wrong values.
  PRODUCT_TABLE_HIDDEN_BY_DEFAULT.forEach((key) => {
    if (state.visibleColumns) state.visibleColumns[key] = false;
  });
}
const orderColumns = [
  { key: "status",           label: "Status",      defaultOn: true  },
  { key: "pending",          label: "Pending",     defaultOn: true  },
  { key: "code",             label: "Code",        defaultOn: true  },
  { key: "product",          label: "Item",        defaultOn: true  },
  { key: "plu",              label: "PLU",         defaultOn: true  },
  { key: "vendor",           label: "Vendor",      defaultOn: true  },
  { key: "velocity",         label: "SV/day",      defaultOn: true  },
  { key: "units",            label: "Sold",        defaultOn: true  },
  { key: "stock",            label: "Stock",       defaultOn: true  },
  { key: "reorderMin",       label: "Min",         defaultOn: true  },
  { key: "reorderMax",       label: "Max",         defaultOn: true  },
  { key: "recommendedOrder", label: "Rec. Order",  defaultOn: true  },
  { key: "caseOrder",        label: "Cases to Order",  defaultOn: true  },
  { key: "caseSize",         label: "Case Size",   defaultOn: true  },
  { key: "unitCost",         label: "Unit Cost",   defaultOn: false },
  { key: "totalCost",        label: "Total Cost",  defaultOn: true  },
];
function placeColumnAfter(order, key, afterKey) {
  const next = (order || []).filter((entry) => entry !== key);
  const index = next.indexOf(afterKey);
  next.splice(index >= 0 ? index + 1 : next.length, 0, key);
  return next;
}
if (!state.orderVisibleColumns) {
  state.orderVisibleColumns = Object.fromEntries(orderColumns.map((c) => [c.key, c.defaultOn]));
}
if (!state.orderColumnOrder) {
  state.orderColumnOrder = orderColumns.map((c) => c.key);
}
const validOrderColumnKeys = orderColumns.map((c) => c.key);
state.orderColumnOrder = state.orderColumnOrder.filter((key) => validOrderColumnKeys.includes(key));
validOrderColumnKeys.forEach((key) => {
  if (!state.orderColumnOrder.includes(key)) state.orderColumnOrder.push(key);
  if (!(key in state.orderVisibleColumns)) state.orderVisibleColumns[key] = true;
});
state.orderColumnOrder = placeColumnAfter(state.orderColumnOrder, "product", "code");
state.orderColumnOrder = placeColumnAfter(state.orderColumnOrder, "plu", "product");
state.orderVisibleColumns.plu = true;

const hoverTooltip = document.createElement("div");
hoverTooltip.className = "row-hover-tooltip";
hoverTooltip.hidden = true;
document.body.append(hoverTooltip);

if (!state.visibleColumns) {
  // Keep operational ordering fields visible while hiding secondary item details.
  const defaultOff = new Set(["itemNumber", "sizeAttr", "subType", "containerAttr", "addDate", "inventoryCost"]);
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
state.columnOrder = placeColumnAfter(state.columnOrder, "product", "code");
state.columnOrder = placeColumnAfter(state.columnOrder, "plu", "product");
state.columnOrder = placeColumnAfter(state.columnOrder, "rule", "plu");
state.visibleColumns.plu = true;
state.visibleColumns.rule = true;
enforceCompactProductColumns();

// 2026-06-05: one-time Products layout repair. Some browsers kept an older
// local column layout that exposed secondary detail columns (Item#/Sub/Type/Tag)
// between SD/DOI and Category/Vendor, making the data look shifted. Keep the
// compact Products view aligned by default; users can still re-enable columns.
(function repairInventoryColumnLayoutOnce() {
  try {
    const flag = "posDashboardColumnLayoutRepair:20260605b";
    if (localStorage.getItem(flag) === "done") return;
    ["itemNumber", "sizeAttr", "subType", "containerAttr", "inventoryCost"].forEach((key) => {
      state.visibleColumns[key] = false;
    });
    state.columnOrder = inventoryColumns.map(([key]) => key);
    state.columnOrder = placeColumnAfter(state.columnOrder, "product", "code");
    state.columnOrder = placeColumnAfter(state.columnOrder, "plu", "product");
    state.columnOrder = placeColumnAfter(state.columnOrder, "rule", "plu");
    safeLocalStorageSet("posDashboardVisibleColumns:v4", JSON.stringify(state.visibleColumns));
    safeLocalStorageSet("posDashboardColumnOrder:v3", JSON.stringify(state.columnOrder));
    safeLocalStorageSet(flag, "done");
  } catch (_) {}
})();

const currency = new Intl.NumberFormat("en-US", { style: "currency", currency: "USD" });
const number = new Intl.NumberFormat("en-US", { maximumFractionDigits: 1 });
const svNumber = new Intl.NumberFormat("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });

function activeTabName() {
  return document.querySelector(".tab-button.active")?.dataset.tab || "dashboard";
}

function setAppLoadStatus(message = "") {
  if (!els.appLoadStatus) return;
  els.appLoadStatus.textContent = message;
  els.appLoadStatus.hidden = !message;
}

function clearAppLoadStatus(message) {
  if (!els.appLoadStatus || (message && els.appLoadStatus.textContent !== message)) return;
  setAppLoadStatus("");
}

function saveActiveTabSearch() {
  const tab = activeTabName();
  if (!["dashboard", "inventory", "ordering"].includes(tab)) return;
  state.tabSearches[tab] = els.searchInput?.value || "";
  safeLocalStorageSet("posDashboardTabSearches:v1", JSON.stringify(state.tabSearches));
}

els.fileInput.addEventListener("change", (event) => loadFiles(event.target.files));
els.folderInput.addEventListener("change", (event) => loadFiles(event.target.files));
els.excelInput.addEventListener("change", (event) => loadExcelFile(event.target.files[0]));
els.multiBarcodeInput?.addEventListener("change", (event) => loadMultiBarcodeFile(event.target.files[0]));
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

const renderDebounced = debounce(render, 40);

// Date navigation arrows + period mode
document.querySelector("#dateNavPrev")?.addEventListener("click", () => shiftDateRange(-1));
document.querySelector("#dateNavNext")?.addEventListener("click", () => shiftDateRange(1));
document.querySelector("#datePeriodMode")?.addEventListener("change", (e) => {
  state.datePeriodMode = e.target.value;
});
const renderParentsDebounced = debounce(renderParents, 120);
const syncStickyHeightsDebounced = debounce(syncStickyHeights, 60);
window.addEventListener("resize", syncStickyHeightsDebounced);
function renderInventoryControlsLight() {
  clearInventorySelection();
  syncStickyHeights();
  queueActiveTabRender();
}
// Keep search and ordering controls on the lighter active-tab render path.
const renderInventorySearchDebounced = debounce(renderInventoryControlsLight, 140);
els.searchInput?.addEventListener("input", renderInventorySearchDebounced);
[els.leadDays, els.safetyDays, els.daysOfInventory].filter(Boolean).forEach((input) => input.addEventListener("input", renderInventoryControlsLight));
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
els.searchInput?.addEventListener("click", () => {
  els.searchInput.select?.();
});
els.parentsSearch?.addEventListener("input", () => {
  const upper = els.parentsSearch.value.toUpperCase();
  if (els.parentsSearch.value !== upper) els.parentsSearch.value = upper;
});
// Selects and date pickers render immediately (single discrete change)
[els.startDate, els.endDate, els.departmentFilter, els.categoryFilter, els.vendorFilter,
 els.colorFilter, els.segmentMetric, els.segmentGroup, els.sortMode, els.inventoryStateFilter, els.inventoryQuickFilter,
 els.comparePeriod, els.compareGroup].filter(Boolean)
  .forEach((input) => input.addEventListener("input", () => {
    if ([els.departmentFilter, els.categoryFilter, els.vendorFilter, els.colorFilter, els.inventoryStateFilter, els.inventoryQuickFilter].includes(input)) {
      renderInventoryControlsLight();
      return;
    }
    render();
  }));
[els.startDate, els.endDate].filter(Boolean).forEach((input) => {
  input.addEventListener("input", () => {
    state.activePresetDays = null;
    clearInventorySelection();
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

[els.departmentFilter, els.categoryFilter, els.vendorFilter, els.colorFilter, els.inventoryStateFilter, els.inventoryQuickFilter].filter(Boolean).forEach((select) => {
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
document.querySelector("#chooseMultiButton")?.addEventListener("click", () => els.multiBarcodeInput?.click());
document.querySelector("#chooseMultiButtonInline")?.addEventListener("click", () => els.multiBarcodeInput?.click());
document.querySelector("#orderArrangeColumnsButton")?.addEventListener("click", () => {
  state.orderArrangeColumns = !state.orderArrangeColumns;
  if (!state.orderArrangeColumns) state.orderArrangeSource = "";
  document.querySelector("#orderArrangeColumnsButton").classList.toggle("active", state.orderArrangeColumns);
  document.querySelector("#orderArrangeColumnsButton").textContent = state.orderArrangeColumns ? "Lock columns" : "Arrange columns";
  renderOrders();
});
els.createPoShortcut?.addEventListener("click", openProductPoReviewModal);
els.countStartNewCard?.addEventListener("click", () => startNewCount());
els.countContinueCard?.addEventListener("click", () => resumeActiveCount());
els.countReviewFilters?.addEventListener("click", (e) => {
  const chip = e.target.closest("[data-review-filter]");
  if (!chip) return;
  state._countReviewFilter = chip.dataset.reviewFilter;
  renderCountReviewPanel();
});
els.countReviewTable?.addEventListener("click", (e) => {
  const th = e.target.closest("th[data-review-sort]");
  if (!th) return;
  cycleCountReviewSort(th.dataset.reviewSort);
});
els.closeCountSessionButton?.addEventListener("click", closeActiveCountSession);
els.countReviewButton?.addEventListener("click", () => openCountReport(state.activeCountSession?.id, "input"));
els.countSaveSessionButton?.addEventListener("click", saveCountSession);
document.querySelector("#countSoundToggleButton")?.addEventListener("click", toggleCountSound);
updateCountSoundToggle();
els.countDeleteSessionButton?.addEventListener("click", deleteCountSession);
els.countInputReportButton?.addEventListener("click", () => openCountReport(state.activeCountSession?.id, "input"));
els.countComparisonReportButton?.addEventListener("click", () => openCountReport(state.activeCountSession?.id, "comparison"));

function placeCountEntriesInDesktopGrid() {
  const grid = document.querySelector("#countSessionModal .count-workspace__grid");
  const entrySection = document.querySelector("#countEntryTable")?.closest("section");
  if (!grid || !entrySection || entrySection.parentElement === grid) return;
  entrySection.classList.add("count-workspace__entries");
  grid.appendChild(entrySection);
}
placeCountEntriesInDesktopGrid();
els.countStartButton?.addEventListener("click", (event) => {
  event.preventDefault();
  // Start Inventory must still launch the count session. Pass 12 accidentally left
  // only the layout helper here, which made the button appear non-responsive.
  startCountSessionFromModal();
  setTimeout(placeCountEntriesInDesktopGrid, 0);
});

els.countCancelButton?.addEventListener("click", closeCountSetupModal);
els.countSetupModal?.addEventListener("click", (event) => {
  if (event.target === els.countSetupModal) closeCountSetupModal();
  const addButton = event.target.closest?.("[data-count-add-filter]");
  if (addButton) {
    event.preventDefault();
    addCountSetupFilter(addButton.getAttribute("data-count-add-filter"));
  }
});
els.countSearchButton?.addEventListener("click", handleCountLookup);
els.countClearSearchButton?.addEventListener("click", clearCountLookup);
let countAutoCommitTimer = null;
function clearCountAutoCommitTimer() {
  if (countAutoCommitTimer) {
    clearTimeout(countAutoCommitTimer);
    countAutoCommitTimer = null;
  }
}
function scheduleCountAutoCommit(delay = 250) {
  if (!state.countAutoPlusMode || !els.countSearchInput || !state.activeCountSession) return;
  const raw = cleanCell(els.countSearchInput.value || '').trim();
  if (!raw || raw.length < 4) return;
  clearCountAutoCommitTimer();
  countAutoCommitTimer = setTimeout(() => {
    countAutoCommitTimer = null;
    if (!state.countAutoPlusMode || !state.activeCountSession || !els.countSearchInput) return;
    const value = cleanCell(els.countSearchInput.value || '').trim();
    if (!value || value.length < 4) return;
    // Safety net for scanners/browsers that occasionally miss the Enter key.
    // Wait for the scanner burst to finish, then exact-match only so partial
    // barcodes never fall through to a random text-search item.
    handleCountAutoPlusLookup();
  }, delay);
}
els.countSearchInput?.addEventListener("keydown", (event) => {
  if (
    event.key.length === 1
    && !event.ctrlKey
    && !event.altKey
    && !event.metaKey
    && state._replaceCountInputOnNextKey
  ) {
    els.countSearchInput.value = "";
    state._replaceCountInputOnNextKey = false;
  }
  if (state.countAutoPlusMode && event.key === "Enter") {
    event.preventDefault();
    event.stopPropagation();
    clearCountAutoCommitTimer();
    hideCountDropdown();
    handleCountAutoPlusLookup();
    return;
  }
  const dropdown = document.querySelector("#countSearchDropdown");
  if (dropdown && !dropdown.hidden) {
    const items = [...dropdown.querySelectorAll(".count-dd-item[data-code]")];
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
    state._replaceCountInputOnNextKey = true;
    handleCountLookup();
    return;
  }
  if (event.key === "Delete") {
    event.preventDefault();
    els.countSearchInput.value = "";
    hideCountDropdown();
  }
});
// Debounced count search - avoids scanning 25k items on every keystroke
const renderCountDropdownDebounced = debounce((val) => renderCountDropdown(val), 120);
els.countSearchInput?.addEventListener("input", () => {
  // In +1 Auto mode the scan box must act like a scanner-only lane.
  // Do not open the typeahead dropdown because scanner Enter can select a
  // suggestion instead of committing the +1, causing every-other-scan misses.
  if (state.countAutoPlusMode) {
    hideCountDropdown();
    scheduleCountAutoCommit();
    return;
  }
  renderCountDropdownDebounced(els.countSearchInput.value);
});
// KBD-A: On touch/mobile (Kindle, Android) do NOT auto-select on focus/click.
// That fights the soft keyboard and causes jumps mid-entry.
// On desktop keep select-on-focus so a barcode scanner's output replaces the prior value.
els.countSearchInput?.addEventListener("focus", () => {
  if (!scanPerformanceProfile.isMobile && !scanPerformanceProfile.isSilk) {
    els.countSearchInput.select?.();
  }
});
els.countSearchInput?.addEventListener("click", () => {
  // If an item is waiting for qty and the operator clicks the barcode box,
  // treat that as an intentional rescan/override. The next scanner burst
  // replaces the field instead of appending to Entered Qty.
  if (state.activeCountSession && state.countStage === "qty") {
    state.countStage = "search";
    state.selectedCountItemCode = "";
    state.countQtyBuffer = "0";
    state.pendingDuplicateMode = null;
  closeCountKeypadPanel();
    els.countSearchInput.value = "";
    hideCountDropdown();
    renderSelectedCountItem();
    renderCountQuantity();
  }
  if (!scanPerformanceProfile.isMobile && !scanPerformanceProfile.isSilk) {
    els.countSearchInput.select?.();
  }
});
// KBD-B: Global keystroke-to-search relay — desktop only.
// On touch/Kindle devices (isMobile, isSilk) the soft keyboard fires 'input' events, not
// 'keydown' events routed via document. Intercepting here causes double-entry, focus battles,
// and keyboard flicker. The countSearchInput's own 'input' listener (above) handles all
// soft-keyboard typing correctly without this relay.
document.addEventListener("keydown", (event) => {
  // Skip entirely on touch/mobile — let the browser and soft keyboard work normally.
  if (scanPerformanceProfile.isMobile || scanPerformanceProfile.isSilk) return;
  if (els.countSessionModal?.hidden || !state.activeCountSession || state.countStage !== "search") return;
  if (els.countDuplicateModal && !els.countDuplicateModal.hidden) return;
  if (!els.countSearchInput || event.target === els.countSearchInput) return;
  const target = event.target;
  if (target?.closest?.("textarea, select")) return;
  if (event.key === "Enter") {
    event.preventDefault();
    event.stopPropagation();
    handleCountLookup();
    return;
  }
  if (event.key === "Backspace") {
    event.preventDefault();
    event.stopPropagation();
    els.countSearchInput.value = els.countSearchInput.value.slice(0, -1);
    renderCountDropdownDebounced(els.countSearchInput.value);
    focusCountSearch();
    return;
  }
  if (event.key.length !== 1 || event.ctrlKey || event.altKey || event.metaKey) return;
  event.preventDefault();
  event.stopPropagation();
  const now = performanceNow();
  if (!state._lastCountScannerKeyAt || now - state._lastCountScannerKeyAt > 180) {
    els.countSearchInput.value = "";
  }
  state._lastCountScannerKeyAt = now;
  els.countSearchInput.value = `${els.countSearchInput.value || ""}${event.key}`;
  renderCountDropdownDebounced(els.countSearchInput.value);
  focusCountSearch();
}, true);
els.priceCheckSearchButton?.addEventListener("click", () => {
  const query = cleanCell(els.priceCheckSearchInput?.value || "");
  if (!query && prefersPhoneBarcodeScanner()) {
    startPriceCheckCamera({ fullscreen: true });
    return;
  }
  handlePriceCheckLookup();
});
els.priceCheckManualButton?.addEventListener("click", () => {
  stopPriceCheckCamera();
  focusPriceCheckSearch();
});
els.scanModeStartButton?.addEventListener("click", () => startPriceCheckCamera({ fullscreen: true, forceFullscreen: true }));
els.scanModeManualButton?.addEventListener("click", () => switchTab("pricecheck"));

// Scan mode: Bluetooth/manual scanners use the exact barcode index first, then refocus immediately.
document.querySelector("#scanModeLookupButton")?.addEventListener("click", () => {
  const inp = document.querySelector("#scanModeInput");
  if (!inp) return;
  handleScanModeLookup(inp.value);
});
document.querySelector("#scanModeClearButton")?.addEventListener("click", () => {
  const inp = document.querySelector("#scanModeInput");
  if (inp) { inp.value = ""; inp.focus(); }
  renderPriceCheckResult(null);
  setScanLookupStatus(state._priceCheckExactIndex ? "Scanner ready" : "Loading item lookup...", state._priceCheckExactIndex ? "ready" : "loading");
  if (!state._priceCheckExactIndex) queueScanExactIndexBuild();
});
document.querySelector("#scanModeInput")?.addEventListener("keydown", (e) => {
  if (e.key !== "Enter") return;
  e.preventDefault(); e.stopPropagation();
  const inp = e.target;
  handleScanModeLookup(inp.value);
});
document.querySelector("#scanModeInput")?.addEventListener("focus", (e) => {
  setTimeout(() => e.target.select?.(), 0);
});
els.priceCheckClearButton?.addEventListener("click", clearPriceCheckSearch);
els.priceCheckCameraButton?.addEventListener("click", () => startPriceCheckCamera({ fullscreen: prefersPhoneBarcodeScanner() }));
els.priceCheckStopButton?.addEventListener("click", stopPriceCheckCamera);
els.priceCheckTorchButton?.addEventListener("click", togglePriceCheckTorch);
els.priceCheckOverlayClose?.addEventListener("click", stopPriceCheckCamera);
els.priceCheckOverlayTorchButton?.addEventListener("click", togglePriceCheckTorch);
document.querySelector("#logoutButton")?.addEventListener("click", () => lockApp("Logged out."));
document.querySelector("#lockLogoutButton")?.addEventListener("click", () => lockApp("Logged out."));
document.querySelector("#metricsHoverZone .metrics-peek-bar")?.addEventListener("click", () => {
  state.metricsPinned = !state.metricsPinned;
  safeLocalStorageSet("posDashboardMetricsPinned:v1", JSON.stringify(state.metricsPinned));
  updateMetricsSummaryMode();
});
els.priceCheckSearchInput?.addEventListener("keydown", (event) => {
  if (event.key === "Enter") {
    event.preventDefault();
    handlePriceCheckLookup({ refocus: true });
  }
});
els.priceCheckSearchInput?.addEventListener("focus", () => els.priceCheckSearchInput.select?.());
els.priceCheckSearchInput?.addEventListener("click", () => els.priceCheckSearchInput.select?.());
els.priceCheckSearchInput?.addEventListener("input", () => {
  if (els.priceCheckStatus) els.priceCheckStatus.textContent = "Ready for next scan.";
  renderPriceCheckDropdown(els.priceCheckSearchInput?.value || "");
});
els.priceCheckSearchInput?.addEventListener("focus", () => renderPriceCheckDropdown(els.priceCheckSearchInput?.value || ""));
els.priceCheckSearchInput?.addEventListener("blur", () => setTimeout(hidePriceCheckDropdown, 120));
els.countKeyButtons?.forEach((button) => button.addEventListener("click", () => handleCountKey(button.dataset.countKey)));
els.countAutoPlusButton?.addEventListener("click", toggleCountAutoPlusMode);
els.countUndoAutoButton?.addEventListener("click", undoLastCountAutoPlus);
els.countDuplicateModal?.addEventListener("click", (event) => {
  if (event.target === els.countDuplicateModal) closeDuplicateCountModal();
});

// When the scan/count modal is open, clicking anywhere that isn't an interactive element
// forces focus back to the search/scan bar so scanning always works immediately
els.countSessionModal?.addEventListener("click", (event) => {
  if (els.countDuplicateModal && !els.countDuplicateModal.hidden) return;
  if (els.countReportModal && !els.countReportModal.hidden) return;
  // When a report/history action re-opens the count workspace, the same tap/click
  // can bubble into this backdrop handler and instantly close it. Ignore that
  // one backdrop close for a short moment.
  if (event.target === els.countSessionModal && Date.now() < Number(state._ignoreNextCountBackdropCloseUntil || 0)) {
    event.preventDefault();
    event.stopPropagation();
    return;
  }
  if (event.target === els.countSessionModal) {
    // Keep physical count open on accidental backdrop taps/clicks. This prevents
    // report-history Continue from opening and immediately closing the workspace.
    focusCountSearch();
    return;
  }
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
document.querySelector("#countContinueButton")?.addEventListener("click", (event) => {
  event.preventDefault();
  event.stopPropagation();
  if (!event.isTrusted) return;
  void continueCountFromReport(event);
});
document.querySelector("#productPoCloseButton")?.addEventListener("click", closeProductPoReviewModal);
document.querySelector("#productPoSendButton")?.addEventListener("click", () => sendProductPoSelection(false));
document.querySelector("#productPoSendAllButton")?.addEventListener("click", () => sendProductPoSelection(true));
document.querySelector("#productPoExcelButton")?.addEventListener("click", () => exportProductReviewCsv());
document.querySelector("#productPoPdfButton")?.addEventListener("click", () => exportProductReviewPdf());
document.querySelector("#productPoReviewSortSelect")?.addEventListener("change", (event) => {
  state.productPoReviewSort = event.target.value || "rec-desc";
  renderProductPoReviewModal();
});
["#logsInventoryFilter", "#logsSalesFilter", "#logsDataFilter"].forEach((selector) => {
  document.querySelector(selector)?.addEventListener("change", () => renderUploadLogs());
});
document.querySelector("#adjustLogMonthFilter")?.addEventListener("change", () => renderAdjustLog());
document.querySelector("#clearAllPendingPoButton")?.addEventListener("click", () => clearAllPendingPo());
document.querySelector("#reportBoxLogs")?.addEventListener("click", () => {
  document.querySelector("#reportLogsModal").hidden = false;
  renderUploadLogs();
});
document.querySelector("#poHistoryDetailCloseButton")?.addEventListener("click", () => {
  document.querySelector("#poHistoryDetailModal").hidden = true;
});
document.querySelector("#poHistoryDetailModal")?.addEventListener("click", (event) => {
  if (event.target === document.querySelector("#poHistoryDetailModal")) document.querySelector("#poHistoryDetailModal").hidden = true;
});
els.newItemsMonthFilter?.addEventListener("change", () => renderNewItems());
els.newItemsYearFilter?.addEventListener("change", () => {
  state.newItemsYear = els.newItemsYearFilter.value || "";
  renderNewItems();
});
els.newItemsVendorFilter?.addEventListener("change", () => {
  state.newItemsVendor = els.newItemsVendorFilter.value || "";
  renderNewItems();
});
els.newItemsDepartmentFilter?.addEventListener("change", () => {
  state.newItemsDepartment = els.newItemsDepartmentFilter.value || "";
  renderNewItems();
});
els.newItemsCategoryFilter?.addEventListener("change", () => {
  state.newItemsCategory = els.newItemsCategoryFilter.value || "";
  renderNewItems();
});
els.newItemsSearchInput?.addEventListener("input", () => {
  state.newItemsSearch = cleanCell(els.newItemsSearchInput.value || "");
  renderNewItems();
});
els.newItemsDormantFilter?.addEventListener("change", () => {
  state.newItemsDormantMonths = els.newItemsDormantFilter.value || "";
  renderNewItems();
});
els.newItemsClearButton?.addEventListener("click", () => {
  state.newItemsYear = "";
  state.newItemsVendor = "";
  state.newItemsDepartment = "";
  state.newItemsCategory = "";
  state.newItemsSearch = "";
  state.newItemsDormantMonths = "";
  if (els.newItemsYearFilter) els.newItemsYearFilter.value = "";
  if (els.newItemsVendorFilter) els.newItemsVendorFilter.value = "";
  if (els.newItemsDepartmentFilter) els.newItemsDepartmentFilter.value = "";
  if (els.newItemsCategoryFilter) els.newItemsCategoryFilter.value = "";
  if (els.newItemsSearchInput) els.newItemsSearchInput.value = "";
  if (els.newItemsDormantFilter) els.newItemsDormantFilter.value = "";
  renderNewItems();
});
els.multiSearchInput?.addEventListener("input", () => {
  state.multiBarcodeSearch = cleanCell(els.multiSearchInput.value || "");
  renderMultiBarcodes();
});
els.productPoReviewModal?.addEventListener("click", (event) => {
  if (event.target === els.productPoReviewModal) closeProductPoReviewModal();
});
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
  safeLocalStorageSet("posDashboardAdjustLog:v1", "[]");
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
  openCountReport(state.countReportOpenId || state.activeCountSession?.id || state.countSessions[0]?.id, "input");
});
els.countComparisonViewButton?.addEventListener("click", () => {
  state.countReportMode = "comparison";
  openCountReport(state.countReportOpenId || state.activeCountSession?.id || state.countSessions[0]?.id, "comparison");
});
els.arrangeColumnsButton.addEventListener("click", () => {
  state.arrangeColumns = !state.arrangeColumns;
  els.arrangeColumnsButton.classList.toggle("active-edit", state.arrangeColumns);
  els.arrangeColumnsButton.textContent = state.arrangeColumns ? "Done arranging" : "Arrange columns";
  renderInventory();
});
els.inventoryColumnPicker?.querySelector("summary")?.addEventListener("click", (event) => {
  event.preventDefault();
  const nextOpen = !els.inventoryColumnPicker.open;
  els.inventoryColumnPicker.open = nextOpen;
  event.currentTarget.setAttribute("aria-expanded", String(nextOpen));
  if (nextOpen) renderColumnPicker();
});
document.addEventListener("click", (event) => {
  if (!els.inventoryColumnPicker?.open || els.inventoryColumnPicker.contains(event.target)) return;
  els.inventoryColumnPicker.open = false;
  els.inventoryColumnPicker.querySelector("summary")?.setAttribute("aria-expanded", "false");
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
  if (document.querySelector("#poHistoryDetailModal") && !document.querySelector("#poHistoryDetailModal").hidden) {
    if (event.key === "Escape") { document.querySelector("#poHistoryDetailModal").hidden = true; return; }
  }
  if (els.productPoReviewModal && !els.productPoReviewModal.hidden) {
    if (event.key === "Escape") { closeProductPoReviewModal(); return; }
  }
  // Report modals
  for (const id of ["reportAdjustModal","reportPoModal","reportCountModal","reportLogsModal"]) {
    const el = document.querySelector("#" + id);
    if (el && !el.hidden && event.key === "Escape") { el.hidden = true; return; }
  }
  if (document.querySelector("#vendorRuleModal") && !document.querySelector("#vendorRuleModal").hidden) {
    if (event.key === "Escape") { document.querySelector("#vendorRuleModal").hidden = true; return; }
    if (event.key === "Enter") { event.preventDefault(); saveVendorRule(); return; }
  }
  if (document.querySelector("#sessionHistoryModal") && !document.querySelector("#sessionHistoryModal").hidden) {
    if (event.key === "Escape") {
      // If count report is open inside, close that first
      if (!els.countReportModal.hidden) { closeCountReport(); return; }
      document.querySelector("#sessionHistoryModal").hidden = true;
      return;
    }
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
  // Enter in setup modal navigates: date -> vendor -> category -> status -> start
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
      // V2.0: ESC simply closes the workspace VIEW. The session stays active in
      // Report History (Active / Live), so nothing is lost — Continue reopens it.
      state._countSessionOpen = false;
      els.countSessionModal.hidden = true;
      renderCountsWorkspace();
      return;
    }
  }
  if (!els.countReportModal.hidden && event.key === "Escape") {
    closeCountReport();
    return;
  }
  if (!els.countDuplicateModal.hidden) return;
  if (
    !els.countSessionModal.hidden &&
    state.activeCountSession &&
    state.countStage === "search" &&
    event.target !== els.countSearchInput &&
    !event.target.closest?.("input, textarea, select, button") &&
    (event.key === "Enter" || event.key.length === 1)
  ) {
    els.countSearchInput?.focus();
    if (event.key === "Enter") {
      event.preventDefault();
      handleCountLookup();
      return;
    }
  }
  if (!els.countSessionModal.hidden && event.target === els.countSearchInput && event.key === "Enter") return;
  if (!els.countReportModal.hidden) return;
  if (!els.countSessionModal.hidden && state.activeCountSession && state.countStage === "qty") {
    const target = event.target;
    const isTypingInSearch = target === els.countSearchInput;
    if (isTypingInSearch) {
      if (event.key === "Enter") {
        event.preventDefault();
        event.stopPropagation();
        handleCountLookup();
        return;
      }
      return;
    }
    // Qty-first logic: after an item is found, digits go to Entered Qty.
    // To rescan/override, click inside the barcode/search box first.
    if (/^\d$/.test(event.key)) {
      event.preventDefault();
      event.stopPropagation();
      if (noteQtyDigitForScannerGuard(event.key)) return;
      handleCountKey(event.key);
      return;
    }
    if (event.key === "Backspace") {
      event.preventDefault();
      event.stopPropagation();
      handleCountKey("back");
      return;
    }
    if (event.key === ".") {
      event.preventDefault();
      event.stopPropagation();
      handleCountKey(".");
      return;
    }
    if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
      const bufferedScan = state._countQtyScannerBuffer || "";
      if (looksLikeBarcodeValue(bufferedScan) && redirectQtyBarcodeToSearch(bufferedScan)) {
        showToast("Barcode scan caught in qty field — moved back to scan.", 2200, "warning");
        return;
      }
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

document.addEventListener("change", (event) => {
  const recInput = event.target.closest(".order-rec-input");
  if (recInput) {
    commitOrderRecommendationInput(recInput);
    return;
  }
  const itemField = event.target.closest("[data-item-field]");
  if (itemField) {
    if (itemField.dataset.itemField === "state") {
      itemField.className = itemField.className
        .replace(/\bstate-select[a-z-]*\b/g, "")
        .trim();
      itemField.classList.add(`state-select-${normalizeItemState(itemField.value).toLowerCase().replace(/\s+/g, "")}`);
    }
    commitItemFieldInput(itemField);
    return;
  }
  const input = event.target.closest("[data-reorder-field]");
  if (input) commitReorderFieldInput(input);
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape") {
    closeDatePickerPopup();
    closeOrderVendorMenu();
    closeInventoryBulkActionsModal();
    closeAppConfirmModal();
  }
  const editable = event.target.closest?.(".mini-input, .order-rec-input, [data-item-field]");
  if (!editable) return;
  if (event.key === "Escape") {
    event.preventDefault();
    if (editable.matches("[data-item-field]")) revertEditableInput(editable);
    else revertEditableInput(editable);
    editable.blur();
    return;
  }
  if (editable.matches("[data-reorder-field]") && event.key === "Tab" && !event.shiftKey) {
    if (editable.dataset.reorderField === "min") {
      event.preventDefault();
      commitReorderFieldInput(editable, { render: false });
      moveReorderAcrossRow(editable, "next");
      return;
    }
  }
  if (event.key !== "Enter") return;
  event.preventDefault();
  if (editable.classList.contains("order-rec-input")) {
    commitOrderRecommendationInput(editable);
  } else if (editable.matches("[data-item-field='caseSize']")) {
    const code = editable.dataset.code;
    commitItemFieldInput(editable, { render: false });
    patchInventoryRowFromCache(code);
    renderInventorySummary(state.inventoryRows || []);
    if (activeTabName() === "ordering") renderOrders();
    requestAnimationFrame(() => focusNextColumnField('[data-item-field="caseSize"]', code, event.shiftKey ? "prev" : "next"));
    return;
  } else if (editable.matches("[data-item-field]")) {
    commitItemFieldInput(editable);
  } else {
    const code = editable.dataset.code;
    const field = editable.dataset.reorderField;
    commitReorderFieldInput(editable, { render: false });
    patchInventoryRowFromCache(code);
    renderInventorySummary(state.inventoryRows || []);
    if (activeTabName() === "ordering") renderOrders();
    requestAnimationFrame(() => focusNextColumnField(`[data-reorder-field="${field}"]`, code, event.shiftKey ? "prev" : "next"));
    return;
  }
  if (editable.matches("[data-reorder-field]")) {
    moveReorderByColumn(editable, event.shiftKey ? "prev" : "next");
    return;
  }
  moveEditableFocus(editable, event.shiftKey ? "prev" : "next");
});

document.addEventListener("focusout", (event) => {
  const editable = event.target.closest?.(".mini-input, .order-rec-input, [data-item-field]");
  if (!editable) return;
  if (editable.classList.contains("order-rec-input")) commitOrderRecommendationInput(editable, { render: false });
  else if (editable.matches("[data-item-field]")) commitItemFieldInput(editable, { render: false });
  else commitReorderFieldInput(editable, { render: false });
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
  if (
    activeTabName() === "pricecheck" &&
    (state.priceCheckStream || state.priceCheckScanner) &&
    event.target.closest(".price-check-camera-shell, #priceCheckVideo, #priceCheckScanner")
  ) {
    return;
  }
  if (
    !interactiveTarget &&
    activeTabName() === "pricecheck" &&
    event.target.closest(".app, .panel, .metrics, .controls, .tab-view, .sticky-pills")
  ) {
    focusPriceCheckSearch(false);
  }
  document.querySelectorAll(".column-picker[open]").forEach((detail) => {
    if (!detail.contains(event.target)) detail.removeAttribute("open");
  });
  document.querySelectorAll(".detail-picker[open]").forEach((detail) => {
    if (!detail.contains(event.target)) detail.removeAttribute("open");
  });
  if (!event.target.closest(".order-vendor-menu") && !event.target.closest("[data-order-vendor-pill]")) {
    closeOrderVendorMenu();
  }
  if (!event.target.closest("#datePresets")) {
    closeDatePickerPopup();
  }
  if (!event.target.closest("#itemRulePopover") && !event.target.closest("td[data-col='rule']")) {
    closeRulePopover();
  }
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
    renderInventorySummary(state.inventoryRows || []);
    return;
  }
  if (event.target.closest(".checkbox-col")) return;
  const copyButton = event.target.closest(".copy-code");
  if (copyButton) {
    event.stopPropagation();
    copyText(copyButton.dataset.copyCode || copyButton.textContent.trim(), copyButton);
    return;
  }
  const ruleCell = event.target.closest("td[data-col='rule']");
  if (ruleCell) {
    event.stopPropagation();
    openRulePopover(ruleCell.dataset.code || ruleCell.closest("tr[data-item-code]")?.dataset.itemCode, ruleCell);
    return;
  }
  if (event.target.closest(".mini-input, .reset-override, [data-item-field]")) return;
  // Stock cell click -> open stock adjust modal
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
  if (event.target.closest(".mini-input, .copy-code, .reset-override, [data-item-field], .checkbox-col, .row-checkbox")) {
    event.stopPropagation();
  }
});
els.inventoryBody?.addEventListener("mousemove", (event) => {
  const row = event.target.closest("tr[data-item-code]");
  if (!row) { hideHoverTooltip(); return; }
  const hoverCell = event.target.closest("td[data-col]");
  const hoverableCols = new Set(["code", "product", "plu", "vendor"]);
  if (!hoverCell || !hoverableCols.has(hoverCell.dataset.col || "")) {
    hideHoverTooltip();
    return;
  }
  // Build tooltip lazily on first hover (deferred flag set during render)
  if (row.dataset.tooltipDeferred) {
    delete row.dataset.tooltipDeferred;
    const item = state._inventoryRowIndex?.get(codeKey(row.dataset.itemCode));
    if (item) {
      const windows = salesWindowsFor(item.code).filter((e) => ["7D","30D","60D","90D","6M","365D"].includes(e.label));
      const rule = effectiveItemRule(item);
      row.dataset.tooltipHtml = `
        <div class="row-hover-tooltip__title">${escapeHtml(item.product)}</div>
        <div class="row-hover-tooltip__line">Vendor: ${escapeHtml(item.vendor||"-")}</div>
        <div class="row-hover-tooltip__line">Sold: ${number.format(item.units)}</div>
        <div class="row-hover-tooltip__line">Stock: ${number.format(item.stock)}</div>
        <div class="row-hover-tooltip__line">Cost: ${currency.format(item.unitCost)}</div>
        <div class="row-hover-tooltip__line">Price: ${currency.format(item.price)}</div>
        <div class="row-hover-tooltip__line">Order rule: ${rule.mode === "custom" ? "Custom" : rule.mode === "category" ? "Category" : "Vendor"} (${number.format(rule.safetyDays || 0)} / ${number.format(rule.daysOfInventory || 0)})</div>
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
els.detailDrawer?.addEventListener("click", (event) => event.stopPropagation());
els.detailDrawer?.addEventListener("mousedown", (event) => event.stopPropagation());

function commitOrderRecommendationInput(input, options = {}) {
  if (!input) return;
  const code = input.dataset.code;
  const val = Math.max(0, Math.round(toNumber(input.value) || 0));
  if (input.dataset.lastCommittedValue === String(val)) {
    input.value = String(val);
    return;
  }
  if (!state._orderRecOverrides) state._orderRecOverrides = new Map();
  state._orderRecOverrides.set(codeKey(code), val);
  input.value = String(val);
  input.dataset.lastCommittedValue = String(val);
  if (options.render !== false) renderOrders();
}

function commitReorderFieldInput(input, options = {}) {
  if (!input) return;
  if (isUserRole()) return;
  const code = input.dataset.code;
  const codeId = codeKey(code);
  const field = input.dataset.reorderField;
  if (!codeId || !field) return;
  const val = input.value.trim();
  const nextCommitted = val === "" ? "" : String(Math.max(0, Math.round(toNumber(val) || 0)));
  if (input.dataset.lastCommittedValue === nextCommitted) {
    if (val !== nextCommitted) input.value = nextCommitted;
    return;
  }
  if (val === "") {
    if (state.reorderOverrides[codeId]) {
      delete state.reorderOverrides[codeId][field];
      if (!Object.keys(state.reorderOverrides[codeId]).length) delete state.reorderOverrides[codeId];
    }
  } else {
    state.reorderOverrides[codeId] = state.reorderOverrides[codeId] || {};
    state.reorderOverrides[codeId][field] = Math.max(0, Math.round(toNumber(val) || 0));
    input.value = String(state.reorderOverrides[codeId][field]);
  }
  input.dataset.lastCommittedValue = val === "" ? "" : input.value;
  safeLocalStorageSet("posDashboardReorderOverrides:v1", JSON.stringify(state.reorderOverrides));
  scheduleSharedProductSync(codeId);
  const fastPath = !(
    (els.inventoryQuickFilter?.value || "").includes("needs") ||
    (els.inventoryQuickFilter?.value || "").includes("overrides") ||
    ["reorderMin", "reorderMax", "recommendedOrder", "needs"].includes(state.inventorySort.key)
  );
  if (options.render !== false) {
      if (fastPath) {
        patchInventoryRowFromCache(code);
        renderInventorySummary(state.inventoryRows || []);
        if (activeTabName() === "ordering") renderOrders();
      } else {
      bumpDataStamp();
      scheduleLightInventoryRefresh(code);
    }
  }
}

function sharedProductVisualHash() {
  const inventoryHash = [...state.latestInventory.values()]
    .map((item) => `${item.code}:${item.stock}:${item.state}:${item.addDate}:${item.caseSize}`)
    .join("|");
  const overrideHash = Object.entries(state.reorderOverrides || {})
    .sort(([a], [b]) => compareDisplayValue(a, b))
    .map(([code, override]) => `${code}:${override?.min ?? ""}:${override?.max ?? ""}`)
    .join("|");
  return `${inventoryHash}||${overrideHash}`;
}

function revertEditableInput(input) {
  if (!input) return;
  const field = input.dataset.itemField || input.dataset.reorderField || "";
  const prev = input.dataset.prevValue ?? "";
  if (field === "caseSize") input.value = String(Math.max(1, Math.round(toNumber(prev) || 1)));
  else input.value = prev;
}

function refreshAfterInventoryEdit(options = {}) {
  if (options.refreshStates) updateInventoryStateFilter();
  syncStickyHeights();
  const active = activeTabName();
  if (active === "inventory" && options.lightInventoryCode) {
    const nextItem = patchInventoryRow(options.lightInventoryCode);
    if (nextItem) {
      const key = codeKey(nextItem.code);
      const existing = state.inventoryRows || [];
      const nextRows = existing.map((entry) => codeKey(entry.code) === key ? nextItem : entry);
      state.inventoryRows = nextRows;
      state._inventoryRowIndex = new Map(nextRows.map((entry) => [codeKey(entry.code), entry]));
      renderInventorySummary(nextRows);
    } else {
      renderInventorySummary(currentInventoryRows());
    }
  }
  else if (active === "inventory") renderInventory();
  else if (active === "ordering") renderOrders();
  else if (active === "newitems") renderNewItems();
  else queueActiveTabRender();
  refreshDetailDrawer();
}

function scheduleLightInventoryRefresh(code, options = {}) {
  const key = codeKey(code || "");
  clearTimeout(pendingInventoryRefreshTimers.get(key));
  pendingInventoryRefreshTimers.set(key, setTimeout(() => {
    pendingInventoryRefreshTimers.delete(key);
    refreshAfterInventoryEdit({ lightInventoryCode: code, ...options });
  }, 120));
}

function moveEditableFocus(current, direction = "next") {
  if (!current) return;
  const selector = current.matches("[data-item-field]")
    ? `[data-item-field="${current.dataset.itemField || ""}"]`
    : current.matches("[data-reorder-field]")
      ? "[data-reorder-field]"
      : "";
  if (!selector) return;
  const inputs = [...document.querySelectorAll(selector)].filter((node) => !node.disabled && node.offsetParent !== null);
  const index = inputs.indexOf(current);
  if (index < 0) return;
  const next = direction === "prev"
    ? inputs[Math.max(index - 1, 0)]
    : inputs[Math.min(index + 1, inputs.length - 1)];
  if (next && next !== current) {
    next.focus();
    next.select?.();
  }
}

function moveItemFieldByColumn(current, direction = "next") {
  if (!current?.matches?.("[data-item-field]")) return;
  const field = current.dataset.itemField || "";
  const selector = `[data-item-field="${field}"]`;
  const inputs = [...document.querySelectorAll(selector)].filter((node) => !node.disabled && node.offsetParent !== null);
  const index = inputs.indexOf(current);
  if (index < 0) return;
  const next = direction === "prev"
    ? inputs[Math.max(index - 1, 0)]
    : inputs[Math.min(index + 1, inputs.length - 1)];
  if (next && next !== current) {
    next.focus();
    next.select?.();
  }
}

function focusNextColumnField(selector, currentCode, direction = "next") {
  const inputs = [...document.querySelectorAll(selector)].filter((node) => !node.disabled && node.offsetParent !== null);
  const index = inputs.findIndex((node) => codeKey(node.dataset.code || "") === codeKey(currentCode));
  if (index < 0) return;
  const next = direction === "prev"
    ? inputs[Math.max(index - 1, 0)]
    : inputs[Math.min(index + 1, inputs.length - 1)];
  if (next && next !== inputs[index]) {
    next.focus();
    next.select?.();
  }
}

function moveReorderByColumn(current, direction = "next") {
  if (!current?.matches?.("[data-reorder-field]")) return;
  const field = current.dataset.reorderField || "min";
  const selector = `[data-reorder-field="${field}"]`;
  const inputs = [...document.querySelectorAll(selector)].filter((node) => !node.disabled && node.offsetParent !== null);
  const index = inputs.indexOf(current);
  if (index < 0) return;
  const next = direction === "prev"
    ? inputs[Math.max(index - 1, 0)]
    : inputs[Math.min(index + 1, inputs.length - 1)];
  if (next && next !== current) {
    next.focus();
    next.select?.();
  }
}

function moveReorderAcrossRow(current, direction = "next") {
  const row = current?.closest?.("tr");
  if (!row) return;
  const field = current.dataset.reorderField || "min";
  const targetField = direction === "prev"
    ? (field === "max" ? "min" : "min")
    : (field === "min" ? "max" : "max");
  const next = row.querySelector(`[data-reorder-field="${targetField}"]`);
  if (next && next !== current) {
    next.focus();
    next.select?.();
  }
}

function commitItemFieldInput(input, options = {}) {
  if (!input) return;
  const code = input.dataset.code;
  const field = input.dataset.itemField;
  if (!code || !field) return;
  if (isUserRole()) return;
  if (field === "caseSize") {
    const nextValue = Math.max(1, Math.round(toNumber(input.value) || 1));
    if (input.dataset.lastCommittedValue === String(nextValue)) {
      input.value = String(nextValue);
      return;
    }
    input.value = String(nextValue);
    input.dataset.prevValue = String(nextValue);
    input.dataset.lastCommittedValue = String(nextValue);
    setItemMeta(code, { caseSize: nextValue, caseSizeManual: true });
    if (options.render !== false) {
      patchInventoryRowFromCache(code);
      renderInventorySummary(state.inventoryRows || []);
      if (activeTabName() === "ordering") renderOrders();
    }
    return;
  } else if (field === "state") {
    const nextValue = normalizeItemState(input.value);
    if (input.dataset.lastCommittedValue === nextValue) {
      input.value = nextValue;
      return;
    }
    input.value = nextValue;
    input.dataset.prevValue = nextValue;
    input.dataset.lastCommittedValue = nextValue;
    setItemMeta(code, { state: nextValue, stateManual: true });
  }
  bumpDataStamp();
  if (options.render !== false) scheduleLightInventoryRefresh(code, { refreshStates: field === "state" });
}
Warning: truncated output (original token count: 105009)
Total output lines: 2000


document.addEventListener("input", (event) => {
  const recInput = event.target.closest(".order-rec-input");
  if (recInput) {
    const normalized = String(Math.max(0, Math.round(toNumber(recInput.value) || 0)));
    if (recInput.value !== normalized) recInput.value = normalized;
    return;
  }
  const itemInput = event.target.closest("[data-item-field='caseSize']");
  if (itemInput) {
    itemInput.value = itemInput.value === "" ? "" : String(Math.max(1, Math.round(toNumber(itemInput.value) || 1)));
    return;
  }
  const input = event.target.closest("[data-reorder-field]");
  if (!input) return;
  input.value = input.value === "" ? "" : String(Math.max(0, Math.round(toNumber(input.value) || 0)));
});

document.addEventListener("focusin", (event) => {
  const editable = event.target.closest(".mini-input, .order-rec-input, [data-item-field]");
  if (!editable) return;
  setTimeout(() => editable.select?.(), 0);
});

document.addEventListener("dblclick", (event) => {
  const editable = event.target.closest(".mini-input, .order-rec-input, [data-item-field]");
  if (!editable) return;
  editable.select?.();
});

// Reset-override button: click Auto to clear that field's override
document.addEventListener("click", (event) => {
  const resetBtn = event.target.closest(".reset-override");
  if (resetBtn) {
    event.stopPropagation();
    const { code, field } = resetBtn.dataset;
    const codeId = codeKey(code);
    if (state.reorderOverrides[codeId]) {
      if (field === "all") {
        delete state.reorderOverrides[codeId];
      } else {
        delete state.reorderOverrides[codeId][field];
        if (!Object.keys(state.reorderOverrides[codeId]).length) delete state.reorderOverrides[codeId];
      }
      safeLocalStorageSet("posDashboardReorderOverrides:v1", JSON.stringify(state.reorderOverrides));
      scheduleSharedProductSync(codeId);
    }
    showToast(`${field === "all" ? "Min / Max" : field === "min" ? "Min" : "Max"} restored to auto for ${code}`);
    patchInventoryRowFromCache(code);
    renderInventorySummary(state.inventoryRows || currentInventoryRows());
    refreshDetailDrawer();
    if (activeTabName() === "ordering") renderOrders();
    const heavyOrderingContext =
      (els.inventoryQuickFilter?.value || "").includes("needs") ||
      (els.inventoryQuickFilter?.value || "").includes("overrides") ||
      ["reorderMin", "reorderMax", "recommendedOrder", "needs"].includes(state.inventorySort.key);
    if (heavyOrderingContext) {
      bumpDataStamp();
      scheduleLightInventoryRefresh(code);
    }
    return;
  }
  const editableCell = event.target.closest("td.order-col, td.order-highlight");
  if (editableCell) {
    const innerInput = editableCell.querySelector("input");
    if (innerInput) {
      innerInput.focus();
      innerInput.select?.();
      event.stopPropagation();
      return;
    }
  }
  const inventoryEditableCell = event.target.closest("td.inventory-edit-cell");
  if (inventoryEditableCell) {
    const innerControl = inventoryEditableCell.querySelector("input, select");
    if (innerControl) {
      innerControl.focus();
      innerControl.select?.();
      event.stopPropagation();
      return;
    }
  }
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
    const previousCodes = new Set([...state.latestInventory.keys()]);
    const previousStockByCode = new Map([...state.latestInventory.entries()].map(([key, row]) => [key, toNumber(row.stock)]));
    const touchedSalesDates = new Set();
    let skippedDuplicates = 0;
    let processedFiles = 0;
    for (const file of files) {
      const signature = fileSignature(file);
      const rows = parseCsv(await file.text());
      const isCurrentInventoryFile = /(^|[^a-z])(current[_ -]?inventory|all[_ -]?inventory|inventory[_ -]?snapshot)/i.test(file.name)
        || looksLikeInventoryRows(rows);
      const date = dateFromFileName(file.name) || dateFromRows(rows) || (isCurrentInventoryFile ? todayFileDate() : null);
      if (!date) continue;
      if (isCurrentInventoryFile) {
        const normalized = mergeInventoryRowsByCode(rows.map((row) => normalizeInventoryRow(row, date)).filter((row) => row.code || row.product));
        if (normalized.length && (!selectedInventory || date.iso >= selectedInventory.date)) {
          selectedInventory = { date: date.iso, rows: normalized };
        }
        logUploadedFile({ filename: file.name, type: "Current_Inventory", status: "Success" });
      } else {
        state.rawSales = state.rawSales.filter((row) => row.date !== date.iso);
        state.rawSales.push(...rows.map((row) => normalizeSalesRow(row, date)).filter(Boolean));
        touchedSalesDates.add(date.iso);
        logUploadedFile({ filename: file.name, type: "Daily_Sale", status: "Success" });
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
      // An explicit current-inventory import should always become the active
      // inventory snapshot, even if a previously synced shared snapshot has a
      // newer timestamp such as a test/debug row.
      state.inventories = new Map([[selectedInventory.date, selectedInventory.rows]]);
      registerInventorySnapshotMeta(selectedInventory.date, selectedInventory.rows, previousCodes);
      let receivedPendingClears = 0;
      selectedInventory.rows.forEach((row) => {
        const key = codeKey(row.code);
        if (!key || !previousStockByCode.has(key)) return;
        if (recordPendingReceiptIfNeeded(key, previousStockByCode.get(key), row.stock, { render: false, patch: false, toast: false, persist: false })) {
          receivedPendingClears += 1;
        }
      });
      if (receivedPendingClears) {
        showToast(`Cleared ${number.format(receivedPendingClears)} PO pending item${receivedPendingClears === 1 ? "" : "s"} from current inventory receipt.`, 3400, "success");
      }
    }
    state.dates = [...new Set(state.rawSales.map((row) => row.date))].sort();
    buildLatestInventory();
    ensureVendorRulesFromData();
    bumpDataStamp();
    updateFilterOptions();
    setDefaultDates();
    state._localImportHoldUntil = Date.now() + 120000;
    render();
    renderUploadLogs();
    await savePersistedState();
      const syncOk = await syncSharedDataToSupabase({ silent: true });
    render();
    renderUploadLogs();
    if (!syncOk) {
      showToast("Shared sync did not complete. Check Supabase tables/policies.", 4200, "warning");
    }
    if (skippedDuplicates) {
      els.fileCount.textContent = `${fileSummary()} - skipped ${skippedDuplicates} duplicate file${skippedDuplicates === 1 ? "" : "s"}`;
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
    const excelRows = [...state.excelItems.values()];
    seedItemMetaFromExcelRows(excelRows);
    applyExcelStateToCurrentInventoryRows(excelRows);
    state._loadedFileSignatures.add(signature);
    ensureVendorRulesFromData();
    updateFilterOptions();
    updateInventoryStateFilter();
    bumpDataStamp();
    els.excelStatus.textContent = `${number.format(state.excelItems.size)} Excel items imported for ordering fields.`;
    logUploadedFile({ filename: file.name, type: "Excel_Product", status: "Success" });
    await savePersistedState();
    await syncSharedMetaSnapshotToSupabase({ silent: true, includeVendorRules: true });
    render();
    renderUploadLogs();
  } catch (error) {
    console.error("Excel import failed", error);
    els.excelStatus.textContent = `Excel import failed: ${error.message || error}`;
  }
}


// BEGIN BUILTIN MULTI BARCODES
// Generated from C:/Users/vinni/OneDrive/Desktop/MULTI BARCODE/multi barcode.xlsx.
// Keep this in-code so future app uploads preserve scanner aliases without requiring a manual re-import.
const BUILTIN_MULTI_BARCODE_ROWS = Object.freeze([["000003747007","GAIN LAUNDRY DETERGENT 10OZ",["03747007"]],["000294317723","ST1746 BROWN/MOCHA",["000294317983"]],["008660008633","ZIG ZAG - ULTRA THIN CONES 70S SIZE *ORIG: $1.99",["008660008640"]],["00866149","ZIG ZAG ROLLING PAPERS - ORIGINAL",["008660007025"]],["00866246","ZIG ZAG ROLLING PAPERS - FRENCH 1 1/4",["008660007247"]],["009300000383","PICKLE DILL BURGER CHIPS 16 OZ **FINAL SALE,MARKDO",["013000005791"]],["009800000524","KINDER JOY .7 OZ",["00985234"]],["009800000555","KINDER JOY BUENO 1.5 OZ",["009800552108"]],["009800000739","KINDER JOY GO WILD .7 OZ",["00987339"]],["009800800056","NUTELLA & GO W/ BREADSTICKS 1.8OZ **FINAL SALE,MAR",["009800800162"]],["00987834","KINDER JOY 0.7 OZ [LIMITED EDITION]",["009800570218"]],["010228104244","DISINFECTANT SPRAY 6OZ",["010228104237","010228104039"]],["010700519528","JOLLY RANCHER FRUIT CHEWS",["010700519511"]],["010700804228","ZERO EA",["010700904140"]],["010700807229","PAYDAY PEANUT CARAMEL BAR 1.85 OZ (52G)",["010700963239"]],["010900000406","REYNOLDS ALUMINUM WRAP",["10010900610404"]],["011111261341","DOVE BAR SOAP - ORIGINAL WHITE",["011111018723"]],["011111720756","CARESS SILKENING BEAUTY BAR DAILY SILK (3.75 OZ)",["011111017412"]],["011313012758","FANTASIA POLISH SPRITZ 10 OZ",["011313043066"]],["011313030158","FANTASIA HEAT PROTECTOR SERUM 6OZ",["011313030165"]],["011313030202","FANTASIA HEAT PROTECTOR STRAIGHT SPRAY 6OZ",["20011313030206"]],["011313030240","FANTASIA HEAT PROTECTOR STYLING FOAM 8.5OZ",["10011313031244"]],["012000004223","MUG CREAM 20OZ",["01242203"]],["012000130311","MOUNTAIN DEW - BAJA BLAST (TROPICAL LIME/CITRUS) C",["012000207464","012000250804","012000240768"]],["012000130922","LIPTON PEACH ICED TEA 16.9OZ",["012000023118"]],["012000170768","PEPSI WILD CHERRY ZERO",["012000170775"]],["012000180583","GAME FUEL - RED",["012000180620"]],["012000180590","GAME FUEL - BLUE",["012000180637"]],["012000182372","MOUNTAIN DEW BOTTLE - 16.9OZ",["012000014079"]],["012000182563","PEPSI BOTTLE - 16.9OZ",["012000014093"]],["012000204968","MOUNTAIN DEW -MAJOR MELON 12OZ CAN",["012000204975"]],["01201303","PEPSI CAN - 12OZ",["012000100109"]],["012044044643","OLD SPICE ULTIMATE 4-IN-1",["012044045046"]],["01206201","MOUNTAIN DEW CODE RED CAN - 12OZ",["012000809989"]],["01208500","MOUNTAIN DEW CAN - 12OZ",["012000100499"]],["01215982481","ULTRA HD WIG CAP [2PCS]- NUDE",["021070100902"]],["01216606","DIET MT DEW CAN - 12OZ",["012000813863"]],["01223004","PEPSI 2L",["012000140525","012000002304"]],["012300703994","WINSTON ULTRA LIGHT KING BOX",["09056234"]],["01231003","WILD CHERRY PEPSI CAN 12OZ",["012000809996"]],["0123450005","USB-C 20W POWER ADAPTER",["190198886699","657479759658"]],["01239332","CAMEL CRUSH MENTHOL KING BOX",["012300193139"]],["01241000","STARBUCKS FRAPPUCINO - VANILLA",["012000045523"]],["012546011075","TRIDENT SPEARMINT",["012546011839"]],["012546011112","TRIDENT GUM - TROPICAL TWIST",["012546011853","073390050474","073390055172"]],["012546011136","TRIDENT SUGAR FREE WATERMELON TWIST CHEWING GUM",["012546011860"]],["012546011440","TRIDENT - BUBBLEGUM SINGLE",["012546011877"]],["012546011471","TRIDENT CINNAMON GUM",["012546011884"]],["012546600026","TRIDENT LAYER SBERRY + CITRS",["012546001359"]],["01264904","STARBUCKS FRAPPUCINO MOCHA - 9.5OZ",["012000103131"]],["01275900","LIPTON BRISK LEMON ICED TEA",["012000810091"]],["01280504","DIET PEPSI CAN - 12OZ",["012000171802"]],["012823927265","BOLD HOLD EXTREME LACE FRONT GLUE 1.3OZ",["850010743011"]],["012993441142","LA CROIX - GUAVA/BEACH/BLACK",["012993812492","012993441128","012993441135"]],["013000001243","KETCHUP HEINZ SQUEEZE 14 OZ",["01312403"]],["013000002691","STEAK SAUCE HEINZ 57 5 OZ ** D/C **",["01326901"]],["013000006408","KETCHUP HEINZ SQUEEZE 20 OZ  ** KW D/C **",["01364008"]],["013087249569","OTIS SPUNKMEYER SEASONAL MUFFIN",["091752001315","013087148114","013087148121"]],["013087901122","OTIS SPUNKMEYER LEMON/MRBLE/CRUMB",["013087901139"]],["013300005019","FLOUR MARTHA WHITE PLAIN 2 LB",["013300055014"]],["013534113467","BETTY DAIN GOLD/BLK STYLING CAPE",["013534102201"]],["013800470089","STOUFFERS CHICKEN BACON RANCH BOWL 14OZ",["013800380937"]],["01402227","SUGAR DADDY CARAMEL POP",["014200302208"]],["01406922","CHARMS BLOW POPS CHERRY LOLLIPOPS - 48 LOLLIPOPS/B",["014200338696"]],["014100054016","MILANO - DOUBLE DARK CHOCOLATE .98OZ **FINAL SALE,",["014100054214"]],["014100077602","GOLDFISH CHEDDAR",["014100046134","014100050544","014100050551","014100075233"]],["014100096016","PEPPERIDGE FARM GOLDFISH 0.9OZ",["014100049999","014100086079"]],["014200004409","FLUFFY STUFF COTTON CANDY .6 OZ",["014200304400"]],["014200532100","SUGAR DADDY - CARAMEL POP 1.7OZ",["014200332106"]],["014200533107","SUGAR BABIES - CARAMEL MILK CHOCOLATE 1.7OZ",["014200333103"]],["014200612208","PARADE COTTON CNDY 2 OZ",["014200306282","60212","720464080362","720464080201","014200602124","720464080386","720464080379"]],["01425626","CHARMS SUPER BLOW POPS ASST",["014200328567"]],["015228101002","APHOGEE DEEP MOISTURE SHAMPOO 16OZ",["10015228101009"]],["015300014954","CHEETOS MACN CHEESE",["015300200289"]],["015499400114","SMART & GO CHARGING CABLE 3.3 FT (USB/TYPE-C/IPHON",["015499400113","015499400112"]],["016000148796","FRUIT BY THE FOOT",["016000234048"]],["016000166097","GARDETTOS SNAKENS ORIGINAL RECIPE SNACK MIX - 1.75",["016000494480"]],["016000172616","FRT GUSHER TROPICAL PEG 4.25 OZ",["016000170315"]],["016000310520","PILLS.SB CONFETTI COOKIES 3 OZ",["016000187030"]],["016000442818","NAUTRE VALLEY GRANOLA - SWEET AND SALTY",["016000104136"]],["016000442825","NATURE VALLEY SWEET & SALTY NUT GRANOLA BAR",["016000168794"]],["016000502963","FRUIT ROLL-UPS (1 CT) 0.5 OZ (14G)",["016000177758","016000233485"]],["016000515284","FRT GUSHER MIX FLAVOR PEG 4.25 OZ",["016000493629"]],["016463243687","7DAY MAGIC CARROT SOAP",["016463243717","016463243700"]],["016571910303","SPARKLING ICE ASSORTED",["016571955397","016571952679","016571950842","016571940355"]],["016571952105","SPARKLING ICE - BLACK CHERRY",["016571955403"]],["017000013237","ARMOUR TREET 12 OZ",["10017000031900"]],["017000024011","DIAL GOLD ANTIBACTERIAL SOAP",["017000141091"]],["017000039497","DIAL SPRING WATER SOAP",["017000141060"]],["017000205809","GOT2B INSTA FLEX HOLD SPRAY 9.1OZ",["00017000205823"]],["017000300450","GOT2B GLUED FREEZE SPRAY 2OZ",["00017000300474","00017000300450"]],["017000313078","GOT2B GLUED SPRAY WAX 8OZ",["00017000313092"]],["017082000033","J LINK ORG BEEF STEAK 1 OZ",["017082873637"]],["017082600035","J LINK JUMBO BEEF STEAK 2 OZ",["017082006035"]],["017082890412","JACK LINKS BIG BEEF STICKS",["017082874009"]],["018000136735","CINNAMON TOAST CRUNCH MINI - SOFT BAKED COOKIE CIN",["018000135639","018000138876","018000133772"]],["018000140497","PILSBURY MINI - CINNAMON ROLL",["018000139613"]],["018000234516","PILSBURY MINI - CHOCOLATE CHIP",["018000116850"]],["018000510689","PILLS.SB CHOC CHIP COOKIE 3 OZ",["018000322749"]],["01814781","NATTY JEWELRY CLEANER 4OZ",["793283822857"]],["018397000015","GUZZLER ISLAND PUNCH 20 OZ",["018397005058"]],["018397000039","GUZZLER CONCORD GRAPES 20 OZ",["018397005027"]],["018397000060","GUZZLER LEMONADE 20OZ",["018397005041"]],["018397000091","GUZZLER ALPINE FREEZE 20 OZ",["018397005096"]],["019320015441","LORNA DOONE SHORTBREAD COOKIES -",["044000013141"]],["02040335","MISTY MENTHOL SILVER LIGHT BOX",["020400052683"]],["02043536","MISTY MEN 120 BOX",["020400053581"]],["021000023226","VELVEETA SHELL MAC N CHEESE CUP 2.39OZ",["021000053551"]],["021000658930","VELVEETA SHELLS & CHEESE ORIGINAL 12OZ BOX",["021000078578"]],["021131501112","MARIE CALLENDERS CHICKEN POT PIE 10OZ",["021131302153","021131375676","021131301538","021131905828"]],["021306210290","ISOPLUS NEUTRALIZING SHAMPOO 8OZ",["10021306210297"]],["021306210566","ISOPLUS NEUTRALIZING SHAMPOO 16OZ",["10021306210563"]],["021306210696","ISOPLUS TEA TREE SHEEN SPRAY[NATURAL REMEDY] 2OZ",["10021306210693"]],["021306210702","ISOPLUS OLIVE OIL SHEEN SPRAY[NATURAL REMEDY] 2OZ",["10021306210709"]],["021306211396","ISOPLUS 24 HOUR HOLDING HAIR SPRAY 2OZ",["10021306121395"]],["021306211426","ISOPLUS SHEEN SPRAY [REGULAR] 2OZ",["10021306481420"]],["022000006653","WRIGLEYS - SPEARMINT PP EA",["02266503","022000106964"]],["022000006660","DOUBLEMINT PP EA",["022000106971","02266600","022000127174"]],["022000006677","JUICY FRUIT PP EA",["02266707"]],["022000008985","EXTRA POLAR ICE PP EA",["02289805","022000159533","022000127051"]],["022000015532","WRIGLEY 5 GUM - WINTERMINT ASCENT (15 STICKS)",["022000127020"]],["022000017352","SWEET AND SOUR SKITTLES",["022000124739"]],["022000021748","STARBURST DUOS FRUIT CHEWS",["022000127242"]],["022000024367","WRIGLEY 5 GUM - SPEARMINT (15 STICKS)",["022000005144","022000127013","022000127006","022000105080"]],["022000123749","ORBIT BUBBLEMINT GUM BOTTLE",["022000014757"]],["022000126986","WRIGLEYS EXTRA GUM - WINTERFRESH",["02284004"]],["022000282385","STARBURST - SWIRLERS ORIGINAL 2.96OZ **FINAL SALE,",["022000282392"]],["022000284792","SKITTLES SMOOTHIES",["022000284808"]],["022000293305","WRIGLEY SPEARMINT GUM 50C",["022000293312"]],["022000298669","SKITTLES GUMMIES :: [ORIGINAL] SHARE SIZE 3OZ",["022000298683"]],["022000298706","SKITTLES GUMMIES :: [SOUR] SHARE SIZE 3OZ",["022000298713"]],["022000330512","STARBURST :: SOUR FRUIT CHEW2.07OZ",["022000330529"]],["022110079806","BUBBLE TAPE AWESOME ORIG EA",["022000017727","022000002785","022000016935"]],["02217109","WRIGLEYS SPEARMINT CHEWING GUM",["022000115713","022000190215"]],["02217206","WRIGLEYS DOUBLEMINT GUM",["022000115720","022000190222"]],["02217303","WRIGLEYS JUICY FRUIT CHEWING GUM 5 STICKS",["022000190239","022000115737"]],["02217808","WRIGLEYS WINTERFRESH CHEWING GUM 5 PIECES",["022000115782","022000190253"]],["022200940214","DEODERANT",["033200191247"]],["02243405","5 FLAVOR LIFE SAVERS",["022000104281"]],["02248400","ORBIT SPEARMINT EA",["022000004840"]],["02248604","ORBIT PEPPERMINT EA",["022000004864"]],["022609000106","MAJESTIC AMONIA 64OZ",["022609000823"]],["022609000243","MAJESTIC BLEACH FLORAL 1 GALLON",["10022609000240"]],["022609000250","MAJESTIC BLEACH LEMON 1 GALLON",["10022609000257","054200044241"]],["02276908","ALTOIDS - ARCTIC MINT",["022000117311"]],["022796610348","OGX MOROCCAN ARGAN DRY SHAMPOO 5OZ",["022796640918"]],["022796910059","OGX COCONUT MILK SHAMPOO 13OZ",["022796918918"]],["022796910066","OGX COCONUT MILK CONDITIONER 13OZ",["022796918925"]],["022796910141","OGX TEA TREE MINT SHAMPOO 13OZ",["022796918932"]],["022796910158","OGX TEA TREE MINT CONDITIONER 13OZ",["022796918949"]],["022796916013","OGX BRAZILIAN KERATIN SHAMPOO 13OZ",["10022796916010"]],["022796916129","OGX MOROCCAN ARGAN OIL [EXTRA] CONDITIONER 13OZ",["022796900111","022796918123"]],["022796918116","ORGANIX MOROCCAN ARGAN OIL SHAMPOO 19.5 OZ BONUS",["10022796918113"]],["02283801","WRIGLEYS EXTRA WATERMELON",["022000159519","022000127143"]],["02284509","WRIGLEYS EXTRA CLASSIC BUBBLE -",["022000126993"]],["02289106","WRIGLEYS EXTRA GUM - PEPPERMINT",["022000127150","022000159526"]],["02289902","EXTRA SPEARMINT PP EA",["022000127044","022000159540","022000008992"]],["024000163022","DEL MONTE WHOLE KERNEL CORN 15.25 OZ",["024000553816"]],["024100000234","CHEEZ-IT PUFFD CHEDDAR 3 OZ",["02412304"]],["024100114993","CHEEZ IT SNAPD - CHEDDAR SOUR CREAM/ DOUBLE CHEESE",["024100115013","024100114962"]],["024100116263","CHEEZ IT EXTRA TOASTY 3 OZ",["024100116256"]],["024100122615","CHEEZ-IT CHEDDAR - 1.5OZ",["024100717170"]],["024100123209","CHEEZ IT EXTRA CRUNCHY 3 OZ",["024100123216"]],["024100126606","CHEEZ-IT WHITE CHEDDAR - 1.5OZ",["024100108930"]],["024100372775","CHEEZ IT HOT & SPICY 3 OZ",["024100372768"]],["024300761195","SUNBELT GRANOLA :: SWEET & SALTY PEANUT 2OZ",["024300141195"]],["024300833601","LITTLE DEBBIE MINI MUFFIN - (CHOCOLATE CHIP) 1.7OZ",["024300044410"]],["024300833724","LITTLE DEBBIE MINI MUFFIN - (BLUEBERRY) 1.7OZ",["024300044427"]],["024300835322","LIL DEBBIE STAR CRUNCH -",["024300831324","024300841323"]],["024300835919","LITTLE DEBBIE - FUDGE ROUND",["024300808913"]],["024300835933","LITTLE DEBBIE ZEBRA CAKE -",["024300831935","024300808845"]],["024300839382","LITTLE DEBBIE HONEY BUN - 2.5OZ",["024300808388"]],["024300841354","LITTLE DEBBIE :: COSMIC BROWNIE (COOKIES & CREAM)",["024300874352"]],["024500730038","PERSONA 5+1 BLADE",["024500602144","843382104313","843382104634","843382104627","843382104320"]],["025000058073","MINUTE MAID LEMONADE - 12OZ",["025000058387"]],["025000058080","MINUTE MAID FRUIT PUNCH 12 GM OZ",["025000058196"]],["025000058806","MINUTE MAID PINK LEMONADE 12 OZ",["025000058240"]],["025900179113","SWISHER SWEETS WRAP - IRISH CREAM 3 FOR $2.49",["025900379117","025900479176","025900279172"]],["025900179137","SWISHER SWEETS WRAP - PEACH BRANDY 3 FOR $2.49",["025900379131","025900379162","025900479169"]],["025900204204","SWISHER SWEET - BLACK",["025900294205"]],["025900204334","SWISHER SWEETS - STRAWBERRY 2 FOR $1.19 **FINAL SA",["025900284336"]],["025900205362","SWISHER SWEETS - DIAMONDS",["025900405366"]],["025900209735","SWISHER SWEET - SILVER",["025900309732"]],["025900220105","SWISHER GREEN SAVE MORE ON 4 (2PK)",["025900320102"]],["025900220556","SWISHER MINI 3 PK - SWEET",["025900240554"]],["025900220563","SWISHER MINI 3 PK - DIAMOND",["025900240561"]],["025900220570","SWISHER MINI 3 PK - GREEN",["02590040578"]],["025900220587","SWISHER MINI 3 PK - STICKY SWEET",["025900240585"]],["025900220792","SWISHER MINI DIAMOND 3 FOR 2",["025900240790"]],["025900220808","SWISHER MINI SWEET 3 FOR 2",["025900240806"]],["025900225803","SWISHER OUTLAWZ - ORGNL",["025900325800"]],["025900225865","SWISHER OUTLAWZ - SWEET",["025900325862"]],["025900226824","SWISHER SWEETS BANANA SMASH",["025900426828"]],["025900226855","SWISCHER SWEET - SWERVE",["025900426859"]],["025900226862","SWISCHER SWEET CREAM",["025900426866"]],["025900227265","SWISHER SWEET - RED CIGARILLOS",["025900294236"]],["025900227272","SWISHER SWEET - GRAPE",["025900294212"]],["025900227296","SWISHER SWEET - BLUEBERRY",["025900294229"]],["025900227319","SWISHER SWEET - TROPICAL FUSION",["025900296339"]],["025900227333","SWISHER SWEETS - RED 2 FOR $1.39",["025900287337","025900345525","025900445522"]],["025900227395","SWISHER SWEETS - GREEN 2 FOR $1.39",["025900296346","025900296353","025900227425","025900335557","025900445553"]],["025900227401","SWISHER BLK CHERRY 2/99 2 PK",["025900297404","25900227401"]],["025900227418","SWISHER BLK - SMOOTH",["025900297411","25900227418"]],["025900227630","SWISHER SWEET - STICKY SWEET",["025900327637"]],["025900229832","SWISHER SWEETS - DIAMONDS 2 FOR $1.39",["025900329839","025900442538","025900345532"]],["025900229856","SWISHER SWEET - SILVER 2 FOR $1.39",["025900329853","025900345570","025900445577"]],["025900230036","SWISHER SWEETS - ARCTIC ICE",["025900330033"]],["025900236502","SWISHER BLK SMOOTH 2/$1.49 2 PK",["025900336707","025900325718","025900236878","025900336875","025900297435","025900227432"]],["025900236519","SWISHER BLK CHERRY 2/$1.29 2 PK ** DO NOT USE **",["025900336714"]],["025900236885","SWISHER BLK CHERRY 2/$1.49 2 PK",["025900325725","025900336882"]],["025900243012","OPTIMO - SWEET (RED)",["025900443016"]],["025900243036","OPTIMO - GREEN",["025900443030"]],["025900243050","OPTIMO - SILVER 2 FOR $1.29",["025900443054","025900103712","025900243715"]],["025900243074","OPTIMO - BLUE",["025900443078"]],["025900243098","OPTIMO - MANGO",["025900443092"]],["025900244149","OPTIMO - DIAMOND",["025900444143"]],["025900244163","OPTIMO - CREAM",["025900444167"]],["025900259006","SWISHER SWEETS LEAF - ORIGINAL",["025900359003","025900359300","025900259303"]],["025900259013","SWISHER SWEETS LEAF - HONEY",["025900359010","025900359317","025900259310","025900359621","025900259624"]],["025900259020","SWISHER SWEETS LEAF - SWEET AROMATIC",["025900359027","025900359324","025900259327","025900359614","025900259617","025900359522","025900259525"]],["025900259341","SWISHER SWEETS LEAF - PEACH BRANDY",["025900359348"]],["025900259518","SWISHER SWEETS LEAF - HONEY",["025900359515"]],["025900259532","SWISHER SWEETS LEAF - COGNAC",["025900359539","025900359638","025900259631"]],["025900259549","SWISHER SWEETS LEAF - PEACH BRANDY",["025900359546","025900359041","025900259044"]],["025900259600","SWISHER SWEETS LEAF - ORIGINAL",["025900359607","025900359508","025900259501"]],["025900259648","SWISHER SWEETS LEAF - PEACH BRANDY",["025900359645"]],["025900259655","SWISHER SWEETS LEAF - IRISH CREAM",["025900359652","025900359553","025900259556","025900359058","025900259051"]],["025900259662","SWISHER SWEETS LEAF - DARK STOUT",["025900359669","025900359560","025900259563"]],["025900263133","SWISHER SWEETS - DIAMONDS [LEGEND] **FINAL SALE,MA",["025900463137","025900472368"]],["025900279103","SWISHER SWEETS WRAP - ORIGINAL 3 FOR $2.49",["025900379100","025900479152","025900379155"]],["025900279318","SWISHER SWEETS - GOLD 2 FOR $1.19 **FINAL SALE,MAR",["025900379315"]],["026100016802","NEWPORT MENTHOL (BOOST) KING BOX",["026100016819"]],["026100016888","NEWPORT KING BOX ($1 OFF PROMO)",["026100016895"]],["026200001296","SLIM JIM MONST HOT AF 1.94 OZ",["026200001364"]],["026200001357","SLIM JIM GIANT HOT AF EA",["026200001289"]],["026200001371","SLIM JIM GIANT TERIY EA",["026200001302"]],["026200006123","GIANT SLIM JIM - BUFFALO",["026200001128"]],["026200006130","GIANT SLIM JIM - TACO SEASONED .97OZ",["026200001135"]],["026200006147","DAVID SUNFLOWER SEED - ORIGINAL 0.8OZ",["026200012155"]],["026200006840","DAVID SUNFLOWER SEED - NACHO CHEESE 0.8OZ",["026200012452"]],["026200101927","SLIM JIM SMOKED SNACK .28 OZ",["026200362151","026200006215","20026200006219"]],["026200114200","SLIM JIM PEPP/CHEESE EA",["026200114255"]],["026200115092","SLIM JIM ORIGINAL 0.44 OZ",["026200615097"]],["026200115900","SLIM JIM GIANT TWINPK EA",["026200119151"]],["026200117003","SLIM JIM GIANT EA",["026200117058"]],["026200117300","SLIM JIM GIANT TABASC EA",["026200117355"]],["026200117409","SLIM JIM GIANT NACHO EA",["026200117454"]],["026200117508","SLIM JIM GIANT MILD EA",["026200117553"]],["026200140391","SLIM JIM MONST HONEY BBQ EA",["026200140414"]],["026200140506","SLIM JIM MONSTER TABA EA",["026200140513"]],["026200140551","GIANT SLIM JIM - HABANERO",["026200141558"]],["026200140605","SLIM JIM MONSTER ORIG EA",["026200140612"]],["026200230641","GIANT SLIM JIM - TANGY BBQ 0.97OZ",["026200230634"]],["026200230665","GIANT SLIM JIM - SWEET & HOT .97OZ",["026200230658"]],["026200321103","FIRE CRACKER SMOKE SAUSAGE",["026200636009","10026200636006","011002620063600632020005471721122821721152105015","011002620063600632020005471722022121721207283209","02620063600100","011002620063600632020005471722041821721263229509"]],["026200322001","FIRE CRACKER GIANT EA",["026200322056"]],["026200391908","TIJUANA MAMA SAUSAGE EA",["026200391953"]],["026200392905","BIG MAMA SAUSAGE EA",["026200392950"]],["026200471730","ANDY CAPPS HOT FRIES 0.85 OZ (24) G",["4015067247967","00026200471679","026200231891"]],["02657318","NEWPORT MENTHOL BOX 100 PK",["2657318","026100805734","026100005738"]],["02657512","NEWPORT MENTHOL BOX PK",["026100805758","026100005752"]],["02657716","NEWPORT MENTHOL 100 SOFT",["026100805772","026100005776"]],["02657813","NEWPORT MENTHOL KING SOFT",["026100805789","026100005783"]],["02665618","NEWPORT KINGS RED",["026100806564"]],["02666617","NEWPORT CIGARETTES 100 BOX PROMO 50C OFF",["026100865745"]],["02666714","NEWPORT KING BOX PROMO 50C OFF",["026100865752"]],["02667014","NEWPORT SMOOTH SELECT KING BOX",["026100806700"]],["02667616","NEWPORT GOLD 100 BOX",["026100806762","026100806618","02666112"]],["02667713","NEWPORT GOLD KING BOX",["026100806779","02666015","026100806601"]],["027000380406","HUNTS TOMATOES DICED 14.5 OZ",["027000379882"]],["027800062700","SOFT BATCH CHOC CHIP COOK 2.2 OZ",["030100199434"]],["027800073577","KEEB SB SUGAR WAFER KG EA",["027800173574"]],["027800100068","MOTHERS CIRCUS ANIMAL COOKIES 1OZ **FINAL SALE,MAR",["027800100259"]],["02786229","PALL MALL BLUE 100 BOX",["027200018628"]],["02786520","PALL MALL MEN 100 BOX",["027200018659"]],["028000010775","BABY RUTH EA",["028000015992","099900647682","099900625468"]],["028000131708","NESTLE CRUNCH - 1.55OZ",["028000131791","099900703876","099900908363"]],["028000202033","BUTTERFINGER SINGLE CANDY BARS (PACK OF 36)",["028000476359"]],["028200304520","BASIC MENTHOL LIGHTS 100 BOX",["028200304513"]],["028200305329","BASIC MENTHOL KING BOX",["028200305312"]],["02835727","MARLBORO RED KING BOX",["028200135704"]],["02838423","MARLBORO GOLD KING BOX",["028200138408","028200003843"]],["028400324878","LAYS JALAPE\u00d1O KETTLE COOKED CHIPS - 2.25OZ",["028400770262"]],["028400653411","SUNFLOWER SEEDS RANCH",["028400706162"]],["028400670104","FRITOS NATU PLANTAIN CHIPS - ORIGINAL",["028400738224"]],["028400706094","MUNCHIES FLAMIN HOT PEANUTS 2 7/8OZ",["028400726405"]],["028400734042","LAYS - SWEET & SPICY HONEY 2 1/2OZ",["028400770224"]],["028400751360","FUNYUNS - STEAKHOUSE ONION 2 OZ",["028400768221"]],["02873921","VIRGINIA SLIMS MENTHOL LIGHT 120",["028200173904"]],["02877422","VIRGINIA SLIMS MENTHOL 100 BOX",["028200177407"]],["029200002126","MUELLERS SPAGHETTI ENRICHED 16 OZ",["20029200002120"]],["029900212559","GOODSENSE/FOODSENSE/OXLEY SANDWICH BAGS",["10029900212556","898391002525","00860006052911","860006052904","793072626222"]],["030100103417","KEEBLER M&MS BITE SIZE COOKIES",["030100103332","027800172720","027800072723"]],["030100113218","KEEBLER CHOCOLATE SUGAR WAFER ** D/C **",["030100113201"]],["030100125150","CRACKER SAND CHEESE/PB 1.8 OZ",["030100211655"]],["030100125181","CRACKER SAND TOAST/PB 1.8 OZ",["030100211679"]],["030100125211","CRACKER SAND CLUB/CHEDDAR 1.8 OZ",["030100211631","253963"]],["030100125877","KEEB VANILLA SUGAR WAFER 2.75 OZ",["030100125884","030100464679"]],["030100125990","KEEB STRAWBRY SUGAR WAFER 2.75 OZ",["030100126027"]],["030100211488","CRACKER SAND CHS/CHEDDAR 1.8 OZ",["030100211471","253997"]],["030100211518","KEEBLER PB & J SANDWICH CRACKER 1.8OZ ** D/C **",["030100211501"]],["030100217688","KEEBLER MINI FUDGE STRIPES 2 OZ",["030100239765","027800063165","027800262728"]],["030100511595","KEEBLER PEPPER JACK CRACKERS ** KW D/C **",["030100511588"]],["030400778520","ANGEL SOFT BATH TISSUE - DOUBLE ROLL 4.00 CT",["10030400779524","030400774379"]],["030772081242","DAWN ULTR DISHWASH LIQ 7.5 OZ",["20192565"]],["031000007225","BANQUET - SPAGHETTI & MEATBALLS 10OZ",["013800103932"]],["031000007300","BANQUET - SALISBURY STEAK 11.88OZ",["031000007454"]],["031146011933","NONGSHIM SHIN RAMYN - SPICY BEEF STIR FRY 4.76OZ",["031146013395"]],["031146014408","NONGSHIM CUP - (SHIN) BLACK RAMYUN NOODLE SOUP 3.5",["031146022205"]],["031146055012","NONGSHIM BOWL (SHIN) RAMEN INSTANT 4.02OZ",["031146055029"]],["031146055418","K ARMY STEW INSTANT NOODLE",["00031146055432"]],["031146058105","NONGSHIM TONKOTSU RAMEN 3.56OZ **FINAL SALE,MARKDO",["031146060542"]],["031146063451","NONGSHIM - UDON W/ SHRIMP TEMPURA",["031146063475"]],["031146066353","NONGSHIM - VEGAN CURRY NOODLES 3.56OZ",["031146066377"]],["031146067855","NONGSHIM - PREMIUM GOLD W/ CHICKEN BROTH",["031146068135"]],["031146150601","NONGSHIM (SHIN) RAMYUM NOODLE SOUP 4.23OZ BAG ** C",["031146022854","031146039081"]],["031146156917","NONGSHIM UDON 9.73OZ",["031146852321"]],["031146250103","NONGSHIM BOWL NOODLE HOT & SPICY SOUP",["031146852598","031146071012"]],["031146250301","NONGSHIM INSTANT NOODLE - KIMCHI FLAVOR 86G",["031146852628"]],["031146262571","NONGSHIM BOWL NOODLE SOUP - SPICY CHICKEN",["031146054121","031146007622","00031146007615","00031146852307"]],["031146270606","NONGSHIM CUP - SHIN NOODLE SOUP 2.64 OZ (75G)",["00031146852406"]],["031200001085","OCEAN SPRAY VARIETY PACK 18CASE CRAN-GRAPE/APPLE",["031200000583","031200000613","031200005694"]],["031700044377","WHITE OWL - WHITE GRAPE",["031700901083"]],["031700044759","WHITE OWL - GREEN SWEETS",["031700901885"]],["031700044902","WHITE OWL - SILVER 2 FOR $1.19",["031700009406","031700237700","031700057827"]],["031700048559","WHITE OWL - SWEETS",["031700901557"]],["031700049136","WHITE OWL - PINEAPPLE",["031700009765"]],["031700049174","GAME - PINEAPPLE 2 FOR $1.29",["031700009802","031700238004","031700058121"]],["031700049181","GAME - NATURAL SILVER",["031700009833"]],["031700049204","GAME - BLACK SWEET",["031700009857"]],["031700049228","GAME - WHITE GRAPE",["031700009871"]],["031700049235","GAME - BLUE",["031700009888"]],["031700049242","GAME - GREEN 2 FOR $1.29",["031700009895","31700049242","031700237946","031700058060"]],["031700049389","GAME - RED CIGAR",["031700420058"]],["031700049907","WHITE OWL - MANGO",["031700232279"]],["031700049976","GAME - MANGO",["031700232255"]],["031700054215","WHITE OWL - WHITE PEACH",["031700233894"]],["031700054529","WHITE OWL-EMERALD",["031700234211"]],["031700055144","GAME NAT LEAF 2 FOR $1.29",["031700234853"]],["031700055151","GAME LEAF - SWEET 2 FOR $1.49",["031700234860","031700238059","031700058176"]],["031700055168","GAME LEAF WILD BERRY",["031700234877"]],["031700055175","GAME LEAF - TOBACCO",["031700234884"]],["031700055441","GAME - DIAMOND",["031700235164"]],["031700055731","GAME LEAF - WHITE RUSSIAN 2 FOR $1.29",["031700235454"]],["031700055915","W OWL MINI WT GRP 3/1.29 3 PK  ** D/C **",["031700235645"]],["031700056219","WHITE OWL - WHITE RUSSIAN ** D/C **",["031700235966"]],["031700056400","W OWL MINI SLV 3/1.29 3 PK  ** D/C **",["031700236154"]],["031700056653","GAME - BLACK SAPPHIRE",["031700236451"]],["031700056813","GAME MINI - SWEET 3 FOR $1.29",["031700236611"]],["031700056820","GAME MINI - BLUE 3 FOR $1.29  ** D/C **",["031700236628"]],["031700057018","GAME MINI - DIAMOND 3 FOR $1.29",["031700236840","31700057018"]],["031700057087","GAME MINI - DIAMOND",["031700236918"]],["031700057810","WHITE OWL - SWEETS 2 FOR $1.19",["031700237694"]],["031700057834","WHITE OWL - WHITE GRAPE 2 FOR $1.19",["031700237717"]],["031700057865","WHITE OWL - GREEN SWEETS 2 FOR $1.19",["031700237748"]],["031700058077","GAME - BLUE 2 FOR $1.29",["031700237953"]],["031700058114","GAME - DIAMOND 2 FOR $1.29",["031700237991"]],["031700058138","GAME - NATURAL SILVER 2 FOR $1.29",["031700238011"]],["031700058145","GAME - RED CIGAR 2 FOR $1.29",["031700238028"]],["031700058152","GAME - BLACK SWEETS 2 FOR $1.29",["031700238035"]],["031700058169","GAME LEAF - WHITE RUSSIAN 2 FOR $1.49",["031700238042"]],["031700058183","GAME LEAF - NATURAL 2 FOR $1.49",["031700238066"]],["031700059210","GAME - RED RUBY (LIMITED EDITION)",["031700239100"]],["031700059890","GAME LEAF - NATURAL 3 FOR $2.19",["031700239797"]],["032023003157","B&B TROPICAL ROOTS SCULPTING FOAM 8.5OZ",["032023031518"]],["032023007049","AFRICAN ROYALE HOT SIX OIL 8OZ",["032023007056"]],["032023016188","B&B PUMP IT UP FOAM WRAP LOTION 8.5OZ",["032023161819"]],["032023016362","B&B PUMP IT UP SPRITZ [SUP] 8OZ",["032023163615"]],["032207000057","FREESTONE PICKLE",["20032207000051"]],["034000000524","REESE FASTBREAK EA",["034000152513","03405204"]],["034000004409","REESE CUP EA",["03444009","034000440009","034000702329"]],["034000238118","KIT KAT VANILLA EA",["034000238101"]],["034000245246","KITKAT BIG KAT 1.5 OZ",["034000245253","034000245376"]],["034000250103","REESE BIG CUP CARAMEL EA",["034000250097"]],["034000318940","KIT KAT - CHOCOLATE FROSTED DONUT 1.5OZ",["034000318933"]],["034000402939","KITKAT KEY LIME PIE",["034000402922"]],["034000404292","KIT KAT - BLUEBERRY MUFFIN",["034000404285"]],["034000567317","TWIZZLER MINI SOUR",["034000567300"]],["034000936922","REESES FRANKEN-CUP",["034000936915"]],["034000943135","REESE CUP [CRUNCHY]",["034000943128"]],["034000953424","REESE BIG CUP LAVA KING EA",["034000953417"]],["034000997190","REESES - SNACK CAKE 2PC",["034000997183"]],["03400500","REESES NUTRAGEOUS",["034000109302"]],["03400704","ICE BREAKERS SUGAR FREE MINTS COOLMINT 1.5OZ **FIN",["034000720606"]],["03400908","ICE BREAKERS MINT - WINTERGREEN **FINAL SALE,MARKD",["034000720620"]],["034044123029","HIGH BEAMS BLACK #20",["20034044123023"]],["034044123036","HIGH BEAMS BURGUNDY #40",["20034044123030"]],["034044123081","HIGH BEAMS BROWN BLACK #30",["20034044123085"]],["034044123173","HIGH BEAMS BLACK #11",["20034044123177"]],["034044123210","HIGH BEAMS WHITE #21",["20034044123214"]],["03404603","KIT KAT DARK CHOCOLATE",["034000236435"]],["03405408","KIT KAT WHITE CHOCOLATE 1.5 OZ",["034000236107"]],["03423909","HERSHEY COOKIES & CREAM CANDY BAR 1.55 OUNCE BARS",["034000239009"]],["03424005","HERSHEYS MILK CHOCOLATE BAR 1.55-OUNCE BARS (PACK",["034000240005","034000249855","034000249862"]],["03424102","HERSHEYS MILK CHOCOLATE WITH ALMONDS -",["034000241002"]],["03424306","MR. GOODBAR CANDY BAR PEANUTS IN MILK CHOCOLATE **",["034000243006"]],["03424403","ROLO EA",["034000002443"]],["03424500","HERSHEYS SPECIAL DARK CHOCOLATE BAR 1.45 OZ (41G)",["034000245000"]],["03424607","KIT KAT CANDY BAR CRISP WAFERS IN MILK CHOCOLATE",["034000246007"]],["03424704","WHATCHAMACALLIT CANDY BAR",["034000247004"]],["03424801","REESEE PIECES",["034000248513"]],["034285526122","A/BEST KIDS ORIGINAL 2-IN-1 DETANGLER 12OZ",["10034285260139"]],["034285530006","A/BEST NO-LYE KIT REGULAR",["30034285117031","MAIN047632"]],["034285555047","A/BEST GROWTH OIL (COCONUT) 4OZ",["10034285555044"]],["034285556044","A/BEST GROWTH OIL (TEA TREE) 4OZ",["10034285556041"]],["034285557041","A/BEST GROWTH OIL (OLIVE) 4OZ",["10034285557048"]],["034285559045","A/BEST GROWTH OIL (ARGAN) 4OZ",["10034285559042"]],["034285560041","A/BEST GROWTH OIL (JAMAICAN BLACK CASTOR) 4OZ",["10034285560048"]],["034285567125","A/BEST KIDS ORIGINAL DETANGLING LOTION 12OZ",["034285509125"]],["03432004","HERSHEYS ALMOND JOY",["034000903207"]],["03433401","REESES WHITE PEANUT BUTTER CUPS",["034000433001"]],["03450204","TWIZZLER STRAWBERRY - 2.5OZ",["034000530045","034000544127"]],["03451300","TWIZZ NIBS 2.25OZ",["034000530137"]],["03484706","ICEBREAKERS ICE CUBES - SPEARMINT 40PCS **FINAL SA",["034000700264"]],["034856002420","SOUR JACKS - ORIGINAL",["034856022428"]],["034856002437","SQWORMS SOUR GUMMI WORMS",["034856022435"]],["034856002444","SOUR JACKS - WATERMELON",["034856022442"]],["034856002789","WELCHS FRUIT SNACK 0.5OZ [LIMITED VALENTINE]",["034856585060"]],["034856005773","WELCHS FRUIT SNACK 0.5OZ",["034856585091"]],["034856008187","WELCHS FRUIT SNACKS 0.8OZ",["034856890089"]],["034856008248","WELCHS :: JUICEFULS FUSIONS 0.8OZ",["034856842446"]],["034856022497","SOUR JACKS WILDBERRY",["034856122494"]],["034856121664","WELCHS MIXED FRUIT SNACKS (0.9 OZ 66 PK.)",["034856001751"]],["035000446671","AJAX TRIPLE ACTION 12.6OZ",["10035000446678","035000446282"]],["035000446688","AJAX. LEMON DISH SOAP 12.6 FL.OZ",["035000446275"]],["035000466525","IRISH SPRINGS SOAP",["827854001492"]],["035000551146","COLGATE TOOTBRUSH :: MEDIUM",["30035000551147"]],["035000566874","TOOTHPASTE",["033200000938"]],["035000766878","TOOTHPASTE",["067990008980","035000983305"]],["035600096146","FIELD HAM & CHZ 6 OZ",["0356000096146"]],["037000217374","JOY LEMON SCENT",["10037000217371","80302889"]],["037000397137","DAWN DISH ORIG DETERG 6.5 OZ",["030772011317"]],["037000537823","DAWN SIMPLY CLEAN DISHWASH SOAP",["037000537830"]],["037000929765","BOUNTY BASIC ESSENTIAL",["037000746577","10037000746574"]],["037600138727","SPAM 12 OZ",["037600001045"]],["038000052200","KELLOGGS FUN PACK 8-PK",["038000219474","038000219740","038000219634","038000219528","038000219856"]],["038000149917","PRINGLES VARIETY PACK ORGNL/CHEDDAR/SC",["03808306","03808403","03808500","038000309656"]],["038000183737","PRINGLES BBQ CHIPS 1.41 OZ",["00038000183720"]],["038000197642","RICE KRISPIE POPPERS",["038000197307","038000197628"]],["038000200984","POP TART UNICORN POWER - FROSTED CHERRY",["038000200960"]],["038000204005","POP TART BITES",["038000196874","038000223655"]],["038000209185","POP TARTS SPIDER",["038000217982"]],["038000222757","POP TART - VARIETY PK (SMORES, BLUEBERRY, CHERRY,",["038000231636","038000222559","038000222795"]],["038000223242","POP TARTS FROSTED PUMPKIN PIE - 3.5OZ",["038000265181"]],["038000232756","POP TART BITES - CONFETTI CAKE",["038000254925"]],["038000232930","POP TARTS PRETZEL",["038000196799","038000196829"]],["038000296680","KELLOGS FRUITY SNACK",["038000296710","038000296659","038000300769"]],["038000303265","RICE RKT CHC CHIP HOMESTY EA",["038000303258"]],["038000307409","RICE RKT BLISS CHC SEASAL EA",["038000307393"]],["038000310317","POPTART BLUEBERRY SINGLE 2/PK",["038000310300"]],["038000311314","POPTART CINN BROWN SUGAR 2/PK",["038000328565","038000311307","038000220906"]],["038000581717","POPTART SMORES SING SERVE 2/PK",["038000179099","038000581700"]],["038000635304","FRUIT LOOP CEREAL CUP 1.5 OZ",["038000012464"]],["038000635700","KELLOGGS FROSTED FLAKES CEREAL",["038000014680"]],["038000765414","RICE KRISPIE TREAT 1.3 OZ",["038000126710","038000110542"]],["038000845512","PRINGLES THE ORIGINAL 1.4OZ",["00038000845574","038000845574"]],["038000845529","PRINGLES SOUR CREAM POTATO CHIPS 1.4OZ",["00038000845550","038000845550"]],["038000845536","PRINGLES CHEDDAR CHEESE CHIPS 1.41 OZ",["00038000845567","038000845567"]],["038000863509","APPLE JACKS CEREAL CUP 1.5 OZ",["038000014697"]],["038200000018","VAN HOLTEN DILL PICKLES SM -",["038200041219"]],["038200000025","KOSHER PICKLES BOX 7.50 LB",["038200041233"]],["038200000063","VAN H PICKLE JUMBO DILLGF EA",["038200061217"]],["038200000070","VAN HOLTEN JUMBO KOSHER PICKLE",["038200061231"]],["038200000087","VAN HOLTEN JUMBO SOUR PICKLE",["038200061248"]],["038200000094","VAN H PICKLE JUMBO HTGF EA",["038200061224"]],["038200000162","VAN HOLTENS JUMBO PICKLE TAPATIO",["038200061293"]],["038200000841","VAN HOLTEN BLUE RASPBERRY WARHEADS",["038200000834"]],["038200000902","VAN H BG PAPA DILL PICKLE EA",["038200101210"]],["038200000926","VAN H HOT MAMA PICKLE EA",["038200101227"]],["038200000940","VAN HOLTEN JUMBO SOUR SIS (TART & TANGY) PICKLE",["038200101234"]],["038200091252","VAN HOLTEN GARLIC JOE",["038200091207"]],["038259102992","ALUMINUM FOIL 50 SQ FT",["074729120509"]],["038276005108","LUSTER PINK LOTION 32OZ",["038276905101"]],["038276005214","LUSTER PINK OIL T/TREE 2Z DL06",["038276905217"]],["038276005979","RENUTRIENTS SLICK STICK 2OZ",["30038276005970","30038276305971"]],["038276006846","LUSTER PINK KIDS EASY COMB DETANGLING 12OZ",["038276068417"]],["038276006877","LUSTER PINK KIDS DTNGL SPRAY 12OZ",["038276068714"]],["038276007218","LUSTER PCJ CHILD KIT COARSE",["038276907211"]],["038276009007","LUSTER S CURL TEXTURIZER REGULAR 15OZ",["020686533111","03827600931134"]],["038276009311","LUSTER S CURL TEXT SPRAY 8OZ",["40038276009319"]],["03827600931133","S CURL TEXTURIZER EXTRA",["847600253159"]],["038324323383","SLIM BEAUTY BRAND DIETERS TEA GREEN - 18 BAG",["781308001295","781308001288","038324321686"]],["038798032637","KID ARM FLOATIES",["038798042872","680058452398","038798014541"]],["038900009632","DOLE PINEAPPLE JUICE, 8.4 OZ",["038900773700"]],["039000008204","ORTEGA SEASONING MIX TACO 1.25 OZ",["10039000008201"]],["039400016113","BUSCHS BAKED BEAN ORGNL",["039400116110"]],["039400016120","BUSHS BAKED BEANS 16 OZ",["10039400016080"]],["040000000327","M&MS PEANUT CHOCOLATE CANDIES 1.74 OZ (49.3G)",["040000012320","04003207"]],["040000002239","M&M PEANUT BUTTER KING EA  ** D/C **",["04022309","040000380603"]],["040000002376","M&M MINIS TUBE EA",["040000032373"]],["040000002635","SNICKERS KING EA",["04026305","345207040000322528"]],["040000004059","TWIX CARAMEL KING EA",["040000353874"]],["040000004325","M&M PEANUT KING EA",["04043205"]],["040000004356","TWIX CARAMEL EA",["04043506","040000353911"]],["040000009702","STARBURST FAVE-REDS EA",["04097002040000329701","040000329701"]],["040000422068","MILKY WAY",["040000522065"]],["040000422082","3 MUSKETEERS SINGLE 1.92 OZ",["040000522089"]],["040000424314","SNICKERS EA",["040000524311"]],["040000508830","M&MS CARAMEL CHOCOLATE CANDY",["040000508823"]],["040000527060","TWIX DARK",["040000527077"]],["040000541288","M&M HAZLENUT SPREAD",["040000541295"]],["040000551829","TWIX - COOKIES & CREAM",["040000551836"]],["040000555391","M&M FUDGE BROWNIE",["040000555407"]],["04003100","M&M MILK CHOCOLATE CANDY 1.69-OZ",["040000012313","040000000310"]],["04005108","STARBURST ORIGINAL",["040000011514","040000000518"]],["040102014109","ANDIS COMB DOUBLE MAGNETIC SMALL ATTACH 5PC SET AD",["040102004872"]],["040102015571","ANDIS MASTER CLIPPER",["040102018152"]],["040102019005","ANDIS GUIDE MAGNETIC ATTACHMENT 2PK (MASTER) AD019",["040102014208"]],["040102047107","ANDIS T-OUTLINER TRIMMER",["040102047800","040102009341"]],["040102121081","ANDIS CLIPPER OIL",["50040102125015"]],["040102172359","ANDIS PROFOIL SHAVER",["040102171505"]],["040102185953","ANDIS ULTRA EASY CUT 13PC",["040102183904"]],["040102186950","ANDIS EASYSTYLE 13PC",["040102191503"]],["040102240751","ANDIS PIVOT CLIPPER TRIMMER COMBO - BLK/SIL",["040102239656","040102248108"]],["040102326905","ANDIS SLIMLINE PRO GTX CORDLESS TRIMMER",["040102322709"]],["040102730009","ANDIS CORDLESS ENVI",["040102731303"]],["040102730450","ANDIS CORDLESS ENVY",["040102730955","040102662157"]],["040102740008","ANDIS T OUTLINER CORDLESS",["040102740558"]],["040102741500","ANDIS TRIMMER GTX-EXO",["040102741005","2306102701"]],["04010508","SNICKER ALMOND",["040000001058"]],["04014407","M&M PEANUT BUTTER",["040000012443","040000514510"]],["04015600","STARBURST TROPICAL",["040000011569"]],["04016007","SKITTLES ORIGINAL",["040000011606","040000001607"]],["04016201","SKITTLES WILD BERRY",["040000011620","040000001621"]],["04016308","SKITTLES TROPICAL",["040000011637","040000001638"]],["04026606","SOUR SKITTLES",["040000012665"]],["04040101","MILKY WAY KING EA",["040000004011"]],["04043108","M&M ORIGINAL KING EA",["040000004318"]],["04046008","SKITTLES KING EA",["040000004608"]],["04046600","SKITTLES SOUR KING SIZE EA  ** D/C **",["040000004660"]],["04060309","3 MUSKETEERS KING EA",["040000006039","040000246039"]],["04062705","SNICKERS ALMOND KG EA",["040000322474","040000006275"]],["040913000018","KAISER HOT DILL SM -",["10040913100012"]],["040913000025","KAISER DILL PICKLES",["10040913100029"]],["040913000032","KAISER PICKLES KOSHER DILL",["10040913100036"]],["040913000063","KAISER JUMBO HOT PICKLE",["10040913100067"]],["040913000070","KAISER FAMOUS - JUMBO DILL PICKLE",["10040913100074"]],["040913000094","KAISER JUMBO SOUR PICKLE",["10040913100098"]],["041116005190","RING POP EA",["041116011894"]],["041116005275","JUMBO JUMPIN PUSH POP EA",["041116006265"]],["041116005343","PUSH POP EA",["041116009020"]],["041116005824","JUICY DROP TAFFY EA",["041116105821"]],["041116005923","JUICY DROP POP EA",["041116009150"]],["041116005954","BABY BOTTLE POP EA",["041116106347","041116006340","041116004773"]],["041116037016","JUICY DROP :: SOUR GEL PENS (ASST) 0.5OZ",["041116137013"]],["041165416206","CUPS RED PLASTIC 18OZ 20 PK",["050700008387"]],["041168006466","ATK CHICK-O-STICK 1.6 OZ",["041168106463"]],["041168631675","CHICK-O-STICK 0.7OZ",["041168625001"]],["041215111808","TRASH BAGS GOOD SENSE 26GAL",["10041215111805"]],["041294401876","PINESOL CLEAN LEMON FRESH 28 OZ",["055500402946"]],["041364012391","SOUR PUNCH ASSORTED 1PC/MINI (10G)",["041364085036"]],["041364080499","SOUR PUNCH WATERMELON",["041364380490"]],["041364080512","SOUR PUNCH BLUE RASPBERRY STRAWS",["041364380513"]],["041364080536","SOUR PUNCH STRAWBERRY",["041364380537"]],["041364080581","SOUR PUNCH CHERRY STRAW",["041364380582"]],["041364087924","SOUR PUNCH RAINBOW",["041364057927"]],["041364089904","SOUR PUNCH - GRAPE",["041364054346"]],["04138860","BLISTEX MEDICATED [ASST]",["04138877","04138884","00041388951065"]],["041419420058","COMBO CHD/CHZ PRETZEL 6.3 OZ",["430686"]],["041419474303","COMBOS BUFFALO BLUE CHEESE PRETZELS 1.8 OZ.",["041419574300"]],["041419714713","COMBOS SNACKS CHEDDAR CHEESE - 1.8OZ",["041419715710"]],["041419714737","COMBOS PEPPRNI PIZZA CRACKER 1.70 OZ  ** D/C **",["041419715734"]],["041419714744","COMBOS CHEDDAR CHEESE",["041419715741"]],["041419714751","COMBOS PIZZERIA PRETZEL FLAVOR - 1.8OZ",["041419715758"]],["041420001536","BOSTON BAKED BEANS .75OZ",["041420011535"]],["041420014000","JAW BUSTERS JAW BREAKERS CAND  ** KW D/C **",["041420024009"]],["041420020322","CHEWY CANDIES",["041420030321"]],["041420020612","LEMONHEAD - FIERCE CITRUS CHEWY",["041420021619"]],["041420021077","NOW & LATER -STRAWBERRY KIWI",["041420021084"]],["041420021091","NOW & LATER -MANDARIN ORANGE",["041420021121"]],["041420021503","NOW & LATER - MANGO GUAVA",["041420021510"]],["041420028878","ALEXANDER THE GRAPE",["041420028601"]],["041420029165","NOW & LATER ORIGINAL 2.44 OZ ** D/C **",["041420029554"]],["041420042515","NOW & LATER - CHERRY MANGO MORPHS",["041420042522"]],["041420042539","NOW & LATER - MORPHS (LARGE) 2.44 OZ",["041420042546"]],["041420052613","LEMONHEAD - ORIGINAL LEMON",["041420062612"]],["041420063008","LEMONHEAD - TROPICAL CHEWY",["041420163005"]],["041420075865","LEMONHEAD - PINK LEMONADE",["041420085864"]],["041420078170","LEMONHEAD - GRAPEHEAD",["041420088179"]],["041420081491","LEMONHEAD - BERRY CHEWY",["041420091490"]],["041420100420","REDHOTS",["041420200427"]],["041420102653","NOW & LATER - MANGO CHEWY",["041420103650"]],["041420110993","TROLLI SOUR BRITE CRAWLERS MINIS GUMMY CANDY 2OZ",["041420030970"]],["041420113307","NOW & LATER - PINEAPPLE CHEWY",["041420123306"]],["041420117831","LEMONHEAD - APPLEHEAD",["041420127830"]],["041420117862","LEMONHEAD - BERRY CHEWY",["041420127861"]],["041420120862","TROLLI EXTREME SOUR BITES 0.8OZ",["041420110863"]],["041420122361","BRACHS SWEET STRIPES SOFT CANDY PEPPERMINT",["041420077784","041420084492"]],["041420519031","NOW & LATER - PINK LEMONADE CHEWY",["041420529030"]],["041420528903","NOW & LATER SOFT CHERRY 0.93 OZ",["041420518904"]],["041420528910","NOW & LATER - STRAWBERRY CHEWY",["041420518911"]],["041420528965","NOW & LATER - WATERMELON CHEWY",["041420518966"]],["041420528972","NOW & LATER - BANANA CHEWY",["041420518973"]],["041420528989","NOW & LATER - APPLE CHEWY",["041420518980"]],["041420528996","NOW & LATER - GRAPE CHEWY",["041420518997"]],["041420741852","BLACK FOREST - GUMMY BEARS",["041420745836"]],["041420821738","NOW & LATER - XTREME SOUR WATERMELON",["041420831737"]],["041420821745","NOW & LATER -XTREME SOUR APPLE",["041420831744"]],["041420821752","NOW & LATER - XTREME SOUR CHERRY",["041420831751"]],["041420900198","NOW & LATER - LEMON LIME SPLIT",["041420910197"]],["041420900211","NOW & LATER SPLITS LEMON-BLUE",["041420910210"]],["041420900402","NOW & LATER - BLUE RASPBERRY CHEWY",["041420910401"]],["041500805023","HOT SAUCE FRANKS HOT WING 12 OZ",["10041500805020","041500818436"]],["041548816678","OUTSHINE FRUIT BAR POPSICLE",["041548612041","041548610047","041548750958","041548244044"]],["041729143722","SKYLINE CHILI SPAGHETTI - 2 WAY (FROZEN) 13OZ",["041729143944","041729143937","041729143715"]],["041755000518","LANGERS ORANGE 16OZ",["041755150077","041755150091"]],["041789001215","MARUCHAN [CUP] RAMEN :: CHICKEN",["00041789901218"]],["041789001253","MARUCHAN [CUP] RAMEN :: SHRIMP",["00041789901256"]],["041789001420","MARUCHAN [CUP] RAMEN :: HOT & SPICY BEEF 2.25OZ",["00041789901423"]],["041789001437","MARUCHAN [CUP] RAMEN :: HOT & SPICY CHICKEN",["00041789901430","00041789901508","041789001505"]],["041789001444","MARUCHAN INSTANT HOT & SPICY SHRIMP CUP",["00041789901447"]],["041789001499","MARUCHAN INSTANT - CHICKEN TORTILLA",["00041789901492"]],["041789001574","MARUCHAN [CUP] RAMEN :: ROAST CHICKEN 2.25OZ",["00041789901577"]],["041789002113","MARUCHAN RAMEN :: CHICKEN 3OZ",["00041789902116"]],["041789002120","MARUCHAN RAMEN :: BEEF 3OZ",["00070662023129","00041789901225","00041789902123"]],["041789002137","MARUCHAN RAMEN NOODLE SOUP PORK FLAVOR 3 OZ (85G)",["00041789902130"]],["041789002144","MARUCHAN RAMEN ORIENTAL NOODLE SOUP - ORIENTAL FLA",["00041789902147"]],["041789002168","MARUCHAN RAMEN :: CHILI 3OZ",["00041789902161"]],["041789002175","MARUCHAN RAMEN :: SHRIMP 3OZ",["00041789902178"]],["041789002465","MARUCHAN RAMEN NOODLE CHICKEN MUSHROOM 3 OZ (85G)",["00041789902468"]],["041789002519","MARUCHAN RAMEN :: CREAMY CHICKEN 3OZ",["00041789902512"]],["041789002571","MARUCHAN RAMEN :: ROAST CHICKEN 3OZ",["00041789902574"]],["041789002588","MARUCHAN RAMEN :: ROAST BEEF 3OZ",["00041789902581"]],["041789002601","MARUCHAN - SPICY BEEF FIRE NOODLE 3OZ",["00041789902604"]],["041789002625","MARUCHAN RAMEN :: FIRE NOODLE SPICY BEEF 3OZ",["00041789902628"]],["041789002632","MARUCHAN RAMEN :: SAUCY NOODS SPICY CHICKEN 3.3OZ",["00041789902635"]],["041789002816","MARUCHAN RAMEN :: HOT & SPICY CHICKEN 3OZ",["00041789902819"]],["041789003011","MARUCHAN RAMEN :: [ORANGE] BOWL CHICKEN 3.31OZ",["00041789903014"]],["041789003028","MARUCHAN RAMEN [YELLOW] BOWL - HOT & SPICY CHICKEN",["00041789903021"]],["041789003035","MARUCHAN - HOT & SPICY SHRIMP BOWL 3.32OZ",["00041789903038"]],["041789003172","RAMEN NOODLE SOUP LIME CHILI SHRIMP FLAVOR 3 OZ (8",["00041789903175"]],["041789003318","MARUCHAN - WONTON RAMEN CHICKEN 3.93OZ",["00041789903311"]],["041789003325","MARUCHAN - WONTON RAMEN HOT & SPICY 3.93OZ",["00041789903328"]],["041789003561","MARUCHAN FIRE BOWL RAMEN - SPICY BEEF 3.49OZ",["00041789903564"]],["041789007019","MARUCHAN RAMEN :: YAKISOBA CHICKEN FLAVOR 4OZ",["041789907012","00041789907012"]],["041789007439","MARUCHAN RAMEN :: YAKISOBA [SPICY] CHICKEN FLAVOR",["00041789907432"]],["041789007668","MARUCHAN RAMEN :: YAKISOBA CHEDDAR CHEESE 4OZ",["00041789907661"]],["041800301232","WELCHS 100% JUICE MIX",["041800316007","041800354009","041800301249"]],["042037103590","STANBACK (2 POWDERS)",["042037104573","042037111069","042037111052"]],["042037103606","STANBACK POWDER 6-PK",["042037104566","042037111083","042037111076"]],["042037111151","BC POWDER - 6PK",["042037111168"]],["042238301832","HARIBO GOLD BEARS 2OZ",["00042238301887"]],["042238302211","HARIBO GUMMI GOLD BEARS 5 OZ",["042238302266"]],["043000007846","COLLIDERS TWISTED REESES 7 OZ  ** D/C **",["04378406"]],["043000008409","COLLIDERS CHOP HERSH CHOC 7 OZ  ** D/C **",["04384009"]],["043000008423","COLLIDERS REESE LAYER DES 7 OZ  ** D/C **",["04384203"]],["043000008430","COLLIDERS ROLO LAYER DESE 7 OZ  ** D/C **",["04384300"]],["043100753575","40CT ENVELOPE",["071723114213"]],["043695000115","HOT POCKET PEPP PIZZA 4 OZ",["043695630633","043695033007"]],["043695000122","HOT POCKET HAM & CHEDDAR 4OZ",["043695739985"]],["043917108629","WAHL CORDLESS MAGIC CLIPPER",["043917814803"]],["043917311463","WAHL #1 CLIPPER GUARD",["043917226057"]],["044000009298","RITZ BITS CHEESE -",["044000013097"]],["044000036584","NUTTER BUTTER KING 3.5 OZ",["044000036577"]],["044000037444","FIG NEWTON SINGLE SERVE 2 OZ ** D/C **",["044000088446"]],["044000037451","NUTTER BUTTER COOKIES - 1.9OZ",["044000010461"]],["044000043148","CHIPS AHOY! ORIGINAL",["044000043131"]],["044000044299","NABISCO\u00ae OREO\u00ae COOKIES THINS",["044000044367"]],["044000047009","OREO SINGLE SERVE 6PK - 8OZ",["044000047030"]],["044000060718","OREO 2PK - DOUBLE STUFF/SPACE DUNK",["044000077181","044000077334"]],["044000069957","OREO - CAKESTER 2.02OZ",["044000073077"]],["044000080150","CHIPS AHOY CHC CRML BIG C EA",["044000080143"]],["044000080174","CHIP AHOY CHEWY CHOC CHIP - 2.5 OZ",["044000080167"]],["044000080198","CHIP AHOY CHEWY BROWNIE - 2.5 OZ",["044000080181"]],["044000082116","OREO DOUBLE STUF KING 4.15 OZ",["044000082086","044000082109"]],["044600307688","CLOROX BLEACH 30OZ",["10044600307685"]],["044700024119","LUNCHABLES - EXTRA CHEESEY PIZZA 4.2OZ",["044700101100","044700024102"]],["044700360354","LUNCHABLES - NACHO WITH CHEESE DIP & SALSA 4.4OZ",["027815118294","044700006795"]],["045635966017","ALLURE BODY MIST - PARIS HILTON 1.7 OZ",["045635966079"]],["045836005256","HOLLYWOOD OIL OLIVE 2OZ",["30045836005257"]],["045836005904","HOLLYWOOD OIL TEA TREE 2OZ",["20045836005908","30045836005905"]],["045836009254","HOLLYWOOD OIL CASTOR 2OZ",["20045836009258"]],["045836009957","HOLLYWOOD BEAUTY ALOE 2OZ",["20045836009951"]],["046100001134","SARGENTO - SHARP CHEDDAR 11 SLICES 8OZ",["046100002339"]],["046100001219","SARGENTO - COLBY JACK 11 SLICES 7.5OZ",["046100002681"]],["046100354322","SARGENTO BALANCE BREAK - CHEESE & CRACKER/NUTS 9OZ",["046100302682"]],["046500216138","RAID ANT & ROACH SPRY 17.5 OZ",["046500216121"]],["047400664685","GILLETTE CLEAR ADVANCED 5IN1 COOL WAVE / GEL",["047400561434"]],["047400668751","GILLETTE SENSITIVE SENSIBLE SHAVE GEL ISLAND BREEZ",["047400668843"]],["047677001510","SNICKERS ICE CREAM 2OZ",["047677254862"]],["047800000151","UNDERWOOD HAM SPREAD 4.25 OZ",["04781531"]],["048155905641","PERSONAL CARE AIR FRESHENER ROSE 9.10 OZ",["00000163","10048155905648"]],["048155908222","BABY LOVE- 80 BABY WIPES",["892166000839"]],["048500001745","TROPICANA - ORANGE JUICE 10 OZ",["048500758236"]],["048500001769","TROPICANA 100% APPLE JUICE 10OZ",["048500758243"]],["048500008607","TROPICANA ORANGE JUICE 15.2OZ",["048500500736"]],["048500017753","TROPICANA - OJ 12OZ",["048500206034"]],["048500017791","TROPICANA - APPLE 12OZ",["048500205983"]],["048500203194","TROPICANA - PINEAPPLE MANGO 12OZ",["048500206140"]],["049000003710","POWERADE THIRST QUENCHER",["049000007909","049000045260","049000050967"]],["049000024685","COCA COLA CLASSIC BOTTLE - 16OZ",["049000035551"]],["049000024692","DIET COKE BOTTLE - 16OZ",["049000035575"]],["049000024708","SPRITE BOTTLE - 16.9OZ",["049000037562"]],["049000029901","PIBB XTRA 12OZ CAN",["049000030778"]],["049000042566","COCA COLA ZERO - 12OZ CAN",["049000058499"]],["049000050103","COCA COLA -2L",["049000054026"]],["049000050110","DIET COKE - 2 LITER",["049000054040"]],["049000050158","SPRITE - 2LITER",["049000054057"]],["049000050745","POWERADE ZERO ASSRT FLAVR",["049000056433","049000050752"]],["049000072792","SPRITE CHERRY 12OZ CAN",["049000072808"]],["049000532036","COCA COLA WITH COFFEE",["049000543674"]],["04904403","COKE PLASTIC BOTTLES 20 OZ  ** D/C **",["049000000443"]],["04904500","DIET COKE PLASTIC BOTTLES 20 OZ  ** D/C **",["049000000450"]],["04913207","SPRITE 12OZ CAN",["049000058482"]],["049200007228","DOMINO SUGAR PURE CANE 1 LB",["10049200007225"]],["049500698331","GOIPLE HAIR WAX STICK - ORIGINAL 2.7OZ",["850055729001","850062580244"]],["04963406","COCA COLA CAN - 12 OZ",["049000058468"]],["04965802","DIET COKE CAN - 12OZ",["049000058475"]],["04971502","DIET SPRITE ZERO 12OZ CAN",["049000543520"]],["04976400","SPRITE PLASTIC BOTTLES 20 OZ  ** D/C **",["049000007640"]],["050000010110","CARNATION EVAPORATED MILK - VITAMIN D 12 OZ",["050000015313"]],["050000301522","COFFEEMATE CREAMER POWDER 11OZ",["050000903672"]],["050100407926","MARIE CALLENDERS CHICKEN POT PIE/DUMPLING SOUP 14.",["050100405755","050100407803"]],["050700009087","DBL ROLL BATH TISSUE 4-PK",["050700010571"]],["050700271088","IODIZED SALT 26 OZ",["041331038157"]],["051000000118","CAMPBELLS - TOMATO SOUP 10.75 OZ",["051000148735"]],["051000010315","CAMPBELL CREAM OF CHICKEN 10.5 OZ",["051000232755","051000197238"]],["051000012616","CAMPBELLS CREAM OF MUSHROOM SOUP",["051000218926","051000277923","051000286192"]],["051000012920","V8 VEGETABLE JUICE 11.5 FL OZ. 69.1 OZ",["051000185600"]],["051000038852","CAMPBELL CHUNKY CLASSIC CHICKEN NOODLE 18.6 OZ",["051000271297"]],["051000148704","CAMPBELLS CHICKEN NOODLE SOUP - 10.75OZ",["051000148711","051000012517"]],["051000200471","V8 ENERGY - PEACH MANGO / POMEGRANATE BLUEBERY",["051000245434","051000279286","051000200488"]],["051000217684","V8 SPLASH JUICES",["051000221971","051000217691","051000217745"]],["051000233141","SPAGHETTIOS PASTA WITH MEATBALLS 15.6OZ",["051000272898"]],["051500050163","SMUCKERS UNCRUSTABLES - PBJ & GRAPE 2.8OZ",["051500826010"]],["051500239131","CRISCO SHORTENING 1 LB",["196005242504"]],["051721002224","KLOSTERMAN BURGER & HOTDOG BUNS",["051721002286","051721007120"]],["051721002286","KLOSTERMAN HONEY HOTDOG BUN",["051721001456"]],["051721007137","KLOSTERMAN WHITE BREAD",["051721001050"]],["051721007144","KLOSTERMAN HOMESTYLE WHEAT BREAD",["051721004013"]],["052000001693","PROPEL ZERO",["052000014860","052000001716","052000014396"]],["052000103151","GATORADE - 20 OZ",["052000104622","052000104615","052000122473"]],["052000135138","GATORADE - FRUIT PUNCH 28OZ",["052000338751"]],["052000324860","GATORADE - 20 OZ",["052000204087","052000204094"]],["052000328660","GATORADE - 20 OZ",["052000104141"]],["052100034911","MCCORMICK MILD TACO SEASONING",["10052100034918"]],["052100098609","MCCORMICK GRAVY MIX BROWN 0.87 OZ",["10052100098606"]],["052336331457","GOT2B GLUED SPIKING GLUE 6OZ",["00052336914216"]],["052336331495","GOT2B GLUED ULTRA STY GEL 6OZ",["10052336331492"]],["052336341456","GOT2B GLUED SPIKING GLUE 1.25OZ",["40052336341454"]],["052336608009","SMTH N SHINE CURL MOUSS     9OZ",["10052336608006"]],["052336608108","SMTH N SHINE STYLE MOUSSE  9 OZ",["10052336608105"]],["052336631793","SMTH N SHINE SILK WRP LOT 8.5OZ",["10052336631790"]],["052336909298","GOT2B GLUED ULTRA STY GEL 1.25OZ",["00052336909304"]],["052336912014","GOT2B GLUED FREEZE SPRAY 12OZ",["00052336912021"]],["052336913516","GOT2B VOLUMANIAC BODIFYING MOUSSE 8OZ",["017000211060"]],["052336925991","GOT2B GLUED BONDING GLUE 6OZ",["00052336926011"]],["052336990678","SMTH N SHINE OLVE/T MOUSSE 8.5Z",["00052336990692"]],["053400354310","JIMMY DEAN SAUSAGE, EGG & CHEESE CROISSANT .54OZ",["077900310980","077900002663"]],["054100794130","VIENNA SAUSAGE 4.6OZ",["054100183187"]],["054100937025","VIENNA HOT SPICY SAUSAGE 4.6 OZ",["10054100937022"]],["054100952011","ARMOUR VIENNA SAUSAGE JALAPENO 4.6 OZ",["10054100952018"]],["054200044241","101 BLEACH - LEMON 128 FL OZ",["10054200044248"]],["055500017263","CLOROX BLEACH 2.4L",["10055500017260"]],["055500017317","CLOROX BLEACH 2.4L",["10055500017314"]],["058847028503","FIRST FORCE ALL PURPOSE LAVENDAR 28OZ",["XXXXX"]],["058847828066","XTRA DETERGENT - OXY BLUE",["094514000960"]],["058847950200","FIRST FORCE - ALL PURPOSE PINE CLEANER",["058847028206"]],["064144047178","CHEF BOY SPAG/MTBL MICRO 7.5 OZ",["064144605507","064144047024","064144047093","064144047307"]],["065333009076","XTRA DETERGENT - MOUNTAIN RAIN",["094514000915","094514020050"]],["065333009083","XTRA DETERGENT - TROPICAL PASSION",["10065333009080","094514000908","094514419601","094514020043"]],["065333009236","XTRA DETERGENT - SCENTSATIONS",["094514419625"]],["066197780293","YOWIE [KISKO] GIANT FREEZE POPS 5.5OZ ASSORTED",["066197782402","064331748604","064331748154"]],["070018102690","CLAIROL BASIC WHITE BW2 32OZ",["4064666814728"]],["070018106452","WELLA COLOR CHARM TONER - PALE BLONDE T10",["381519047695"]],["070018113757","CLAIROL KALEIDOCOLORS - BLUE 8OZ",["381510001818"]],["070018114358","CLAIROL PURE WHITE 20VOL 16OZ",["381519049149"]],["070018114419","CLAIROL PURE WHTIE 30V 16OZ",["381519049156"]],["070018114433","CLAIROL PURE WHTIE 40V 16OZ",["381519049163"]],["070018114457","CLAIROL SHIMMER LIGHTS SHAMPOO 2OZ",["4064666572949"]],["070100016447","JOHN MORRELL FRANK BUN LENGTH",["070100018144"]],["070137000181","BLACK & MILD CIGARS 5 CT",["070137839620"]],["070137112501","BLACK & MILD - CASINO :: 2 FOR $1.19",["070137512509","070137112778","070137512776","070137514879","070137114871"]],["070137214854","BLACK & MILD FT - SWEET :::: 5PK",["070137833123","070137213727"]],["070137500186","BLACK & MILD (10 PER PACK) (5/EA/PK) ::",["070137839651"]],["070137514855","BLACK & MILD FT - SWEET 10PK/30 [BOX] :: **UNPRICE",["70137214854","70137000181","070137833154","00070137015031","00070137013723","070137513728"]],["070137831327","BLACK & MILD FT - ORIGINAL :::: 5PK",["070137831426","070137213376"]],["070137831358","BLACK & MILD FT - ORIGINAL 10PK/30 [BOX] :: **UNPR",["00070137014997","070137831457","00070137015017","00070137013372","070137513377"]],["070221007201","OREO MILKA CHOCOLATE BAR 1.44 OZ",["070221009359"]],["070330600065","BIC FULL SIZE LIGHTER",["070330622401","070330644809","070330601277","070330647770","070330604995","070330646391","070330646148","070330670594","070330670532"]],["07033245","BIC MINITURE LIGHTER",["070330644090","070330644793","070330602670","070330646483","070330643321"]],["070462003956","SWEDISH FISH TROPICAL",["070462003949"]],["070462004045","SOUR PATCH - TROPICAL",["070462004038"]],["070462098501","SOUR PATCH KIDS CANDY 2 OZ",["070462012347"]],["070462098525","SOUR PATCH KID WATERMELON 2 OZ",["070462012323"]],["070462431445","SOUR PATCH KIDS - EXTREME",["070462041033"]],["070462542509","SWEDISH FISH",["070462012361"]],["070524078229","VELVET - TOILET PAPER 4 ROLL",["10070524078226","10070524380534","070524380537"]],["070622000054","GOETZE CARAMEL CREAM",["070622251012"]],["070622431018","GOETZ COW TALE EA",["070622801019"]],["070622431025","GOETZE COW TALES - CARAMEL BROWNIE",["070622801026"]],["070662010037","NISSIN TOP RAMEN CHICKEN",["070662024133","00070662013038"]],["070662030035","NISSIN CUP RAMEN NOODLES SOUP CHICKEN - 2.25OZ",["00070662048030","070662048030","070662038031"]],["070662087237","CHOW MEIN ORIGINAL STYLE TERIYAKI BEEF 4 OZ",["070662187739","070662082737","00070662082737","070662182734"]],["070686002360","UNIVERSAL REMOTE CONTROL",["044476083754"]],["070847012474","MONSTER ENERGY ULTRA - WHITE ZERO 16OZ",["070847036524"]],["070847036920","MONSTER - JUICE (PIPELINE, MANGO LOCO, PAPILLON) V",["070847895121","070847022909","070847029014"]],["070847037989","MONSTER - ULTRA (WATERMELON, PEACH, S.BERRY) VARIE",["070847893325","070847029427","070847891857"]],["070847811169","MONSTER ENERGY BEVERAGE (GREEN) 16OZ",["070847811237"]],["070847811268","MONSTER ENERGY LO-CARB 16 OZ",["070847811275"]],["070847812715","JAVA MONSTER ENERGY DRINK",["070847812609","070847014430"]],["070847898146","MONSTER ENERGY (ZERO SUGAR) - ULTRA VICE GUAVA 16O",["070847898153"]],["070847898245","MONSTER ENERGY (ZERO SUGAR) - ULTRA BLUE HAWAIIAN",["070847898252"]],["070970448379","MIKE AND IKE - REDRAGEOUS (RED)",["070970458378","070970475702","070970485701"]],["070970449932","MIKE AND IKE - ORGINAL FRUIT (GREEN)",["070970459931"]],["070970449949","HOT TAMALES",["070970459948"]],["070970449956","MIKE AND IKE - BERRY BLAST (BLUE)",["070970459955"]],["070970449963","MIKE AND IKE - JOLLY JOE (PURPLE)",["070970459962"]],["070970449970","MIKE AND IKE - TROPICAL TYPHOON",["070970459979"]],["070970471056","MIKE AND IKE - SOURLICIOUS ZOURS",["070970481055"]],["070970471070","MIKE AND IKE - SOUR LEMON",["070970481079"]],["070970471087","MIKE AND IKE - SOURLICIOUS FRUIT PUNCH",["070970481086"]],["070970471094","MIKE AND IKE - SOURLICIOUS GREEN APPLE",["070970481093","070970481192","070970471193"]],["070970475719","MIKE AND IKE - SOUR WATERMELON",["070970485718"]],["071130000512","GENTLE TREATMENT KIT RELAXER NO-LYE [REG]",["802535203013"]],["071130000581","GENTLE TREATMENT KIT SINGLE SUP",["802535204010"]],["071130000802","ULTRA SHN CONDITION REG 2.25OZ",["802535211018"]],["071130000833","ULTRA SHN CONDITION XTRA 2.25OZ",["802535210011"]],["071130000857","ULTRA SHN CONDITION REG 8 OZ",["802535209015"]],["071130001908","GENTLE TREATMENT KIT RELAXER NO-LYE [REG] TWIN",["802535201026"]],["071130001915","GENTLE TREATMENT KIT RELAXER NO-LYE [SUP] TWIN PK",["802535202023"]],["071164321027","HASK PLACENTA VIALS SUPER 5/8OZ",["20071164321021"]],["071164321034","HASK PLACENTA VIALS W/OLVE DL18",["20071164321038"]],["071164332016","HASK TEA TREE ROSEMARY CONDITIONING PK 1.75OZ",["20071164332010"]],["071164333112","HASK COCONUT MILK & HONEY PACKET",["20071164333116"]],["071164373286","HASK COCONUT DRY SHAMPOO 4.3OZ",["071164371282"]],["071301047872","SUNBEAM WHITE BREAD",["071301047179","071301047377","071301047407"]],["071314003919","AUNT MILLIE WHOLE WHEAT BREAD",["071314102216"]],["071314040044","AUNT MILLIES STADIUM HAMBURGER BUN",["071314043786","071314008556"]],["071314040051","AUNT MILLIES STADIUM HOT DOG BUNS 16PK",["071314008563"]],["071314042703","JOHNSONVILLE - BRAT HOTDOG BUNS 14OZ",["071314069618"]],["071409543245","MANE N TAIL COND BNS 16OZ",["071409543221"]],["071592007746","PNDTCH MUSHROOM PCS STEMSDO 4 OZ",["071592817741"]],["071610202658","BACKWOODS - BLACK RUSSIAN 5PK *CLEARANCE*",["071610340169"]],["071610202665","BACKWOODS - BLACK RUSSIASN 1PK *CLEARANCE*",["071610340176","071610440036"]],["071610202740","DUTCH - RED",["071610440111"]],["071610202757","DUTCH - HONEY FUSION GOLDEN",["071610440128"]],["071610202764","DUTCH - IRISH FUSION 2 FOR $1.29",["071610440135","071610495265","071610495241","071610495258","071610580060","071610340060","071610202559","071610341227","071610440746","071610203679"]],["071610202870","DUTCH - ROYAL HAZE",["071610440203","071610340329"]],["071610202900","DUTCH - GREEN ENVY",["071610202924","071610440258","071610440234"]],["071610203037","DUTCH - LEAF REAL SWEET ** D/C **",["071610440333"]],["071610203051","DUTCH - LEAF PURE TOBACCO ** D/C **",["071610440357","071610340527"]],["071610203075","DUTCH - LEAF RIPE BERRY",["071610340541","071610440371"]],["071610203211","BACKWOODS - BATCH NO 002",["071610340695"]],["071610203273","DUTCH - OG FUSION 2 FOR $1.29",["071610340763","071610440524","071610440531","071610203280","071610341241","071610440760","071610203693"]],["071610203358","BACKWOODS - VANILLA 5PK",["071610340848"]],["071610203525","DUTCH - PLATINUM FUSION 2 FOR $1.29",["071610341081","071610440609"]],["071610203532","BACKWOODS - GRAPE 5PK",["071610341098"]],["071610203624","DUTCH - ROYAL HAZE 2 FOR $1.29",["071610341173","071610440692"]],["071610203648","DUTCH - BERRY FUSION 2 FOR $1.29",["071610341197","071610440715"]],["071610203655","DUTCH - BLUE DREAM FUSION 2 FOR $1.29",["071610341203","071610440722","071610340213","071610440074","071610202702"]],["071610203686","DUTCH - GOLD FUSION 2 FOR $1.29",["071610341234","071610440753","071610340749","071610440500","071610203259"]],["071610203785","DUTCH - DIAMOND FUSION 2 FOR $1.29",["071610341333","071610440845"]],["071610203839","DUTCH - TROPICAL 2 FOR $1.29",["071610440906","071610341470","071610341463","071610440890","071610203822"]],["071610203853","DUTCH - SUNSET DREAM 2 FOR $1.29",["071610440920","071610341500","071610440913","071610203846"]],["071610204225","BACKWOODS - CHOCOLATE 5PK",["071610341517"]],["071610212466","BACKWOODS - RUM CAKE 5PK",["071610341920"]],["071610299603","BACKWOODS - SWEET AROMATIC 3PK",["071610299610"]],["071610301863","BACKWOODS - HONEY BERRY 5PK",["071610301870"]],["071610301894","BACKWOODS - ORIGINAL 5PK *CLEARANCE*",["071610301900"]],["071610301924","BACKWOODS - SWEET AROMATIC 5PK",["071610301931"]],["071610301955","BACKWOODS - ORIGINAL 1PK",["071610301962","071610302211"]],["071610301986","BACKWOODS - SWEET AROMATIC - 1PK",["071610301993","071610302235"]],["071610302037","BACKWOODS - HONEY BERRY 1PK **FINAL SALE,MARKDOWN",["071610302044","071610302228"]],["071610302792","BACKWOODS - RUSSIAN CREAM 5PK",["071610302808","71610302792"]],["071610303065","BACKWOODS - RUSSIAN CREAM 1PK",["071610303089","071610303072"]],["071610340565","BACKWOODS - BANANA 5PK",["071610203099"]],["071610496088","DUTCH - HONEY FUSION 2 FOR $1.29",["071610496101","071610496095","071610580107","071610341210","071610440739","071610203662","071610496132","071610496125","071610496149"]],["071610496682","DUTCH - BERRY FUSION",["071610496705","071610496699"]],["071610505643","DUTCH - BLEND SILVER 2 FOR $1.29",["071610505650","071610505667","071610202733","071610341166","071610440685","071610203617"]],["071610505728","DUTCH - SWEET FUSION 2 FOR $1.29",["071610505742","071610505735","071610341180","071610440708","071610203631"]],["07172144","CARAMEL APPLE POP EACH",["071720305010"]],["07172746","TOOTSIE ROLL POPS ASST - 100",["071720305089"]],["071817000293","BIG RED SODA 12OZ CAN",["071817033291"]],["072030013688","LITTLE BITES FUDGE B",["072030018393"]],["072030013848","LITTLE BITES BLUEBERRY MUFFINS",["072030018485"]],["072030014821","LITTLE BITES CHOCO CHIP",["072030018379","072030960371","072030960418"]],["072030021621","LITTLE BITE PARTY CAKES -",["072030022086","072030960265","072030960142"]],["072030023588","LITTLE BITE SOFT BAKED COOKIE",["072030023595"]],["072103001864","FILI GREE - ASST",["YNV0512"]],["072105011601","SELECT VALUE PAPER TOWEL",["202585"]],["072108010021","MOON PIE CHOCOLATE DD EA",["072108248011","072108720012","072108810010"]],["072108010120","MOON PIE VANILLA D.D. EA",["072108810027"]],["072108010229","MOON PIE BANANA D.D. EA",["072108810034"]],["072180749093","RED BARON DEEP DISH SINGLE PIZZA",["072180749246","072180749086","072180761743"]],["072250008204","MRS. FRESHLEYS CREME FILLED CHOCOLATE CUPCAKES -",["072250912167"]],["072250011273","WONDER ICED HONEYBUN EA",["072250011495"]],["072250037068","NATURES OWN HONEY WHEAT BREAD - 20OZ",["072250005395"]],["072250087957","HERSHEY SWISS ROLLS",["072250088084"]],["072275100105","ALE81 SODA - 12OZ",["072275100143"]],["07235748","YOO-HOO CHOCOLATE DRINK, CANS 11 OZ",["072350240030"]],["072554000287","OREO COOKIE SANDIWCH MINI 3PC",["072554464669","072554191930"]],["072554001628","NESTLE DRUMSTICK VARIETY PACK",["072554110962","072554001697","072554001604"]],["072554252679","OREO BARS SANDWICH POPSICLE",["072554349317","072554415166"]],["072554761829","KIT KAT CHOCOLATE WAFER ICE CREAM CONE",["43915845"]],["072740963716","VELVET 4PK TOILET PAPER ** DO NOT USE **",["10072740963713"]],["072790000164","MAGIC SHAVE POWDER RED 5OZ",["50072790000169"]],["072790102707","STA SOF FRO HAIR SPRAY COMBOUT [EXTRA DRY] 8OZ",["616762102700"]],["072790615559","LETS JAM SHINE GEL REG 4.4OZ",["072790790577"]],["072790790584","LETS JAM SHINE XTRA HOLD 4.4OZ /5.5OZ",["072790615658","072790790577"]],["072799030148","MAMBA FRUIT CHEWS 6 CHEWS",["072799056100","072799056117","072799056131","072799056148"]],["072799035136","MAMBA FRUIT CHEWS (LARGE)",["072799035112","072799055646","072799055639","072799055677","072799055660"]],["072799111670","MAMBA SOUR",["072799111687"]],["072799934057","MAMBA - BERRY TASTY 2.8OZ",["072799118341"]],["072979004143","SEAGRAMSS GINGER ALE 12 FL OZ",["072979004167"]],["073096500235","AA BATTERY 4PK",["30073096500236"]],["07320114","BRUTON SNUFF",["073100002021"]],["073321038762","ICEE RED WHITE AND BOOM",["073321377458","073321377465","073321038755","073321037765","073321037925"]],["073390000110","MENTOS MINT EA",["073390005764","073390007805","073390007812"]],["073390000165","MENTOS RAINBOW EA",["073390005771","073390005948","073390026424"]],["073390008383","AIRHEAD EXTREME SOUR BELT EA",["073390008376"]],["073390009038","AIRHEADS XTREMES BITES",["073390010034"]],["073390009045","AIRHEADS XTREMES (RAINBOW) SHAREABLE",["073390010041"]],["073390009168","AIRHEADS XTREMES SOURFULS",["073390018412"]],["073390011000","AIRHEAD BITES - PARADISE BLENDS 2OZ",["073390011017"]],["073390013134","AIRHEAD SOUR 5PK BARS EA",["073390614584"]],["073390013325","AIRHEADS WATERMELON GUM",["073390020958"]],["073390014360","AIRHEAD BITES FRUIT 2 OZ",["073390030001","073390030087","073390030094","073390014483"]],["073390014377","AIRHEAD BITES BERRY",["073390029999"]],["073390014841","AIRHEADS FILLED ROPES **FINAL SALE,MARKDOWN !!**",["073390014872"]],["073675261205","CAR PRIDE BOLOGNA 5 OZ",["073675261250"]],["073800001577","FAYGO CHERRY COLA CAN - 12OZ",["073800005315","073800009979"]],["073800003908","FAYGO LEMON ICED TEA - 12OZ",["073800010166"]],["073800004523","FAYGO BLACK CHERRY 12 OZ",["073800004639","073800009962"]],["073800004608","FAYGO RED POP 24OZ.",["073800010760"]],["073800004660","FAYGO VANILLA CREME SODA 24 OZ",["00073800460138"]],["073800004738","FAYGO OHANA KIWI STRAWBERRY - CAN 12OZ",["073800004745","073800010180"]],["073800005292","FAYGO CREAM SODA 12 OZ CAN",["073800005308","073800009986"]],["073800005896","FAYGO RASPBERRY LEMONADE - 12 OZ CAN",["073800005902"]],["07387835","FAYGO RED POP - 12OZ",["073800001072"]],["07388038","FAYGO ROOT BEER - 12OZ",["073800001058","073800006596","073800006602","073800010036"]],["07388339","FAYGO GINGER ALE - 12OZ",["073800009993"]],["07388436","FAYGO TWIST - 12OZ",["073800003847","073800010029","7388436"]],["07389231","FAYGO PEACH - 12OZ",["073800004431","073800010074","7389231"]],["074108381866","BABYLISS TRIMMER (ROSEFX) CORDLESS TRIMMER",["074108415394"]],["074108422453","BABYLISS T-BLADE REPLACEMENT BLADE (GOLD) - FX707G",["074108408181"]],["074108448880","BABYLISS TRIMMER GOLDFX",["074108474995"]],["074157033105","TOILET BOWL CLEANER 32 OZ",["082901177333"]],["074300081489","VISINE ADVANCED REDNESS RELIEF",["312547493833"]],["074603005212","SAMYANG RAMEN 120G",["10746030052127"]],["074704100007","MURRAYS POMADE REGULAR 3OZ",["10074704100127"]],["074704260008","MURRAYS BEESWAX YELLOW 4OZ",["10074704260128"]],["074704271004","MURRAYS EDGEWAX 4OZ",["10074704271001"]],["074704273008","MURRAYS EDGEWAX EXTREME HOLD 4OZ",["10074704273005"]],["074743522007","SUMMIT LIV CREME HAIRDRESSING CONDITIONER 15OZ",["074743433203","074743422108"]],["074764301314","ARDELL ADHESIVE INDIVIDUAL - CLEAR 0.125OZ",["074764650580"]],["074764614841","ARDELL DBL INDIVIDUALS FLARE KNOT :: FREE SHORT",["70074764614840"]],["074764614964","ARDELL DBL INDIVIDUALS FLARE KNOT FREE :: LONG",["70074764614963"]],["074764650542","ARDELL KNOT FREE LONG - BLACK",["70074764650541"]],["074764656940","ARDELL KNOT FREE - TRIPLE INDIVIDUAL :: SHORT BLAC",["70074764656949"]],["074764664938","ARDELL KNOTTED FLARE TRIOS 3 IN 1 - SHORT",["074764621573","70074764664937"]],["074764664945","ARDELL KNOTTED FLARE TRIOS 3 IN 1 - MED",["074764621580","70074764664944"]],["074764664952","ARDELL KNOTTED FLARE TRIOS 3 IN 1 - LONG",["074764621597","70074764664951"]],["074764682901","ARDELL DOUBLE INDIVIDUALS FLARE SHORT BLK",["074764664983"]],["074780000703","PERRIER SPARKLING WATER",["074780004022"]],["074806001585","BIG HUG ORANGE",["074806011584"]],["074806001592","BIG HUG GRAPE 16 OZ",["074806011591"]],["074806001608","BIG HUG FRUIT PUNCH",["074806011607"]],["074806001615","BIG HUG BLUE RASPBERRY",["074806011614"]],["074806001622","BIG HUG KIWI STRAWBERRY",["074806011621"]],["074806001639","BIG HUG LEMONADE",["074806011638"]],["074806001691","B…65009 tokens truncated…SOFTEN & REPAIR UREA FOOT CREAM (4.06OZ)",["888432983616"]],["888432983975","AH ULTRA SOFT WOMENS RAZOR 3 BLADE (4PCS)",["888432986464"]],["888432985665","ABNY CALLUS REMOVER (6.76OZ)",["888432985696"]],["888432985672","ABNY CALLUS ELIMINATOR (5.07OZ)",["888432985702"]],["888432985955","ABNY MATTE LIPSTICK BUTTERSCOTCH",["888432986051"]],["888432988741","AH WIG GLUE SUPREME MAX HOLD 1.3OZ/38ML",["888432988758"]],["888432991284","AH BEARD PEN W/ BRUSH BLACK",["888432987942"]],["888432991499","P&I EYELASH EXTENSION SHAMPOO",["888432991505"]],["888432997385","MASTER LIFT BROW GLUE/LASH GEL",["888432997392"]],["888432997767","P&I STRONG HOLD INDIVIDUAL EYELASH BLACK",["888432997781"]],["888432997774","P&I STRONG HOLD INDIVIDUAL EYELASH ADHESIVE",["888432997798"]],["888849016723","QUEST PB&JELLY BAR 2.12 OZ",["888849016747"]],["890246000472","KRAZY GLUE",["070158000054","073754123042","073754151779","070158572483","073754190525","10890246000479"]],["893164000951","SAN2GO HAND SANITIZER SPRAY 2OZ",["893164000982"]],["894700010045","CHOBANI YOGURT - CHERRY/SBERRY/PEACH/B.BERRY 5.3OZ",["818290018120","894700010052","894700010168","894700010069"]],["896586355425","KIMAKALON LARGE 20 - 1B",["655302314685"]],["897127000026","DR.MIRACLES ANTI-BRK CREME 4OZ",["802535226012"]],["897127000187","DR.MIRACLES TEMPLE BALM SUP 4Z",["802535228016"]],["897127000224","DR.MIRACLES NO-LYE RELAXER KIT REGULAR",["60897127000226","802535215016"]],["897127000415","DR.MIRACLES DEEP CONDITIONING TREATMENT PACK",["897127000484","30802535219039","802535219038"]],["897127000491","DR.MIRACLES NEW GROWTH KIT REG",["80253521701"]],["897127000774","DR.MIRACLES COND SHAMPOO 12OZ",["802535000773"]],["897127000781","DR.MIRACLES LEAVE IN COND 8OZ",["802535225015"]],["898391002310","HEALTH SMART BABY POWDER 14OZ",["676979860428"]],["898940001344","TOXIC WASTE SMOG BALLS NO 3 OZ",["10898940001341"]],["898999000022","VITA COCO :: COCONUT WATER 11.1OZ",["898999010526"]],["90512121","DAWN SIMPLE CLEAN 12.6OZ",["10037000827891"]],["9555755800043","ROYAL HONEY - THE ULTIMATE POWER SOURCE",["9555755800036"]],["96907064023","NAIROBI MOISTURE-SIL HAIR LOTION 8OZ",["096907064023"]],["984326529844","Y2K SPLIT",["984326529845"]],["A1J18343","A1 JORDAN KEYCHAIN",["091244455665"]],["ABT320BLACK","GUCCI ASSORTED BELT- ABT",["ABT320RED"]],["ABT323CLEAR","GUCCI ASSORTED DESIGNED- ABT323CLEAR",["ABT322CLEAR"]],["APPLEPIE","JJS BAKERY - APPLE PIE",["011284023067"]],["ASC0279-S","AFRICAN PRINT HAIR BUN CAP ASST",["ASC0280S"]],["BF0143BFTT7","HALO PLW-GRACE :: TT1B/REDWINE",["804725030089"]],["BF1579BF9","HALO JUDY :: 99J",["804725031055"]],["BF1579BFDB1","HALO JUDY :: DARK BLUE",["804725031086"]],["BF1579BFTT2","HALO LW JUDY :: T4/BLONDE",["804725031109"]],["BH-216D","DEGINER HAT",["BH221IV","BH221BK","BH221BR","BH216A","BH216C","BH216B"]],["BHP-021YE","REVERSIBLE HAT",["BHP021RBL","BHP021WT"]],["DEL008","DO NOT USE",["DEL9533662"]],["DELE00122473","GATORADE - 20 OZ",["052000122480","052000124859","052000122428"]],["DELE025900221133","SWISHER SWEETS - DIAMONDS [LEGEND] ** DO NOT USE *",["025900421137"]],["DELE025900259334","SWISHER LEAF - COGNAC (DELETE DO NOT USE)",["025900359331","025900359638"]],["DELE0317316","DELETE",["038000317309","038000259197"]],["DELE0411686002","CHICK-O-STICK $0.35",["041168602057","041168636311"]],["DELE043000047699","KOOL AID",["043000089545","043000080917","043000078051"]],["DELE070330608825","** DELETE **",["070330618893"]],["DELE070330643321","** DELETE **",["070330643550"]],["DELE12312312","NOW & LATER - PINEAPPLE",["HTITUKGJK"]],["DELE1509824032","KC EDGE GLUED 100 ML - GRAPE  ** D/C **",["30731509824033"]],["DELE1509824452","KC EDGE GLUED 100 ML - RED APPLE  ** D/C **",["30731509824453"]],["DELE1509824469","KC EDGE GLUED 100 ML - STRAWBERRY ACAI  ** D/C **",["30731509824460"]],["DELE1509824490","KC EDGE GLUED 100 ML - CHERRY BLOSSOM  ** D/C **",["30731509824491"]],["DELE192087301","DELETE",["DELE908120342130217"]],["DELE21892421","DELETE",["030100199359"]],["DELE31414","OTIS SPUNKMEYER MUFFIN - CHOC/BB/BANANA",["DELE","DELE12321","DELE1232"]],["DELE509824421","KC EDGE GLUED 100 ML - WATERMELON  ** D/C **",["30731509824422"]],["DELE509824476","KC EDGE GLUED 100 ML - VERY CHERRY  ** D/C **",["30731509824477"]],["DELE509824483","KC EDGE GLUED 100 ML - ACACIA  ** D/C **",["30731509824484"]],["DELE693507131362","TOYO/SKY LIGHTER",["DELE"]],["DELE7398478190413U","DR.MIRACLES DP CND TRT 12",["DELE12321"]],["DELE7861","** DELETE *",["1234567890"]],["DELE80173412","DELE",["DELE12LIDSF342123121312"]],["DELE8127321","DUTCH - SILVER",["071610440104"]],["DELE824414","KC EDGE GLUED 100 ML - PINEAPPLE  ** D/C **",["30731509824415"]],["DELE8RO","TEXT & TONES COLOR 8RO FLAMING DESRE NEW",["070018108654"]],["DELEI21374932","** DELETE **",["705372047133"]],["DELE\\799056100","MAMBA FRUIT CHEW .93OZ (SMALL)",["DSFD"]],["DHN-0414","TURBAN HEADWRAP",["721027004337"]],["DHN-0457","CRYSTAL COLLECTION BUN CAP",["636227597223"]],["DHN-0515","TURAN VELVET SOLID BK",["DHN0515"]],["EEM-1391","LABUBU EARMUFFS",["EEM1390","EEM1391"]],["HTTPS://COKEURL.COM/Q/2020-00067","HONEST KID JUICE BOX",["HTTPSCOKEURLCOMQ202000015"]],["HTTPWWWJANETCOLLECTIONCOMQR","EXPRESSION 3X BRAID 85- 4",["884593131286"]],["MAGNUM01","TROJAN MAGNUM CONDOM",["0874"]],["MT38247","IPOD EARPHONES",["611434274433"]],["NAILGLUE01","NAIL GLUE",["713836352006"]],["NARO019","NAIROBI HYDRATING DETANG SHAMP 8OZ",["096907062029"]],["NARO1016","NAIROBI DETOXIFYING SHAMPOO 8OZ",["096907062005"]],["PC0112BLACK","PC0111/112/113 BAGS",["PC0113BLACK","PC0111BLACK"]],["RBG-0465B","CLEAR BAG/SQUARE BK",["RBG0003"]],["RBG-0468B","CLEAR BACK BAG",["RBG0468R9"]],["RD0721807490","RED BARON PIZZA WHOLE",["072180760685"]],["REUSE-757528005047","TAKI FUEGO SINGLE (1OZ)",["10757528022010","757528022013"]],["RNT-1038","FASHION SOCKS",["RNT1039"]],["TIK00003066","LOVFUL NAILS-SMALL-H26/SQUARE/MID LENGTH",["TIK00108231"]],["TIK00003068","LOVFUL NAILS-LARGE-H26/SQUARE/MID LENGTH",["TIK00108233"]],["TIK00004237","LOVFUL NAILS-SMALL-H144/DUCK/CLASSIC",["TIK00108530"]],["TIK00004238","LOVFUL NAILS-MEDIUM-H144/DUCK/CLASSIC",["TIK00108531"]],["TIK00004239","LOVFUL NAILS-LARGE-H144/DUCK/CLASSIC",["TIK00108532"]],["TIK00008306","LOVFUL NAILS-SMALL-H103/COFFIN/CLASSIC",["TIK00034010"]],["TIK00008318","LOVFUL NAILS-MEDIUM-H103/COFFIN/CLASSIC",["TIK00034011"]],["TIK00008330","LOVFUL NAILS-LARGE-H103/COFFIN/CLASSIC",["TIK00034012"]],["TWEEZER01","TWEEZER",["671331777246"]],["VARIETY611269331240","RED BULL - COLOR EDITION (VARIETY PK) 8.4OZ",["611269001709","611269163452","611269001280","611269001297","611269454482","611269331240"]],["YNV-0704","BRAID ACCESSORY",["YNV0680"]],["YSU12340015","WOMEN DESIGNER WALLET",["023856458454"]],["YSU12340106","DRESS WITH FLOWER",["742042547913"]],["YSU12340112","BUTTERFLY WITH FLOWER T-SHIRT",["0917777543322"]],["YSU12340116","SHIRT ISURVIVED",["0917777543322"]],["YSU12340118","WOMENS PRINT SHIRT",["742042550708"]],["YSU12340119","TEDDY BEAR SHIRT 2",["0917777543322"]],["YSU12340120","CAT SHIRT",["0917777543322"]],["YSU12340121","NNIGHT DRESS",["742042546169"]],["YSU12340122","NIGHT SED",["742042546312"]],["YSU12340123","DRESS SED FLOWERS",["742042547968"]],["YSU12340124","NIKE TOPS",["753359033346"]],["YSU12340125","DENIM JEANS",["091277754439"]],["YSU12340129","NIKE 2 PC",["092223111133"]],["YSU12340130","CHANEL SCRUNCH BAND",["294756923829"]],["YSU12340132","CHANEL HAIRPIN",["9348652986367"]],["YSU12340134","NIKE HEAD BAND",["235973259257"]],["YSU12340135","GUCCI MULTI-PURPOSE",["934865298465"]],["YSU12340136","DESIGN BONNETS (MEDIUM )",["85683299999","1212201599106"]],["YSU12340137","DESIGN SCARF",["098111177772"]],["YSU12340138","DESIGN SCAFE",["098111177772"]],["YSU12340147","LV BAG RED -BROWN",["986543627779"]],["YSU12340148","CRIATIAN DIOR BAG",["091777334447"]],["YSU12340149","LV BAG",["482238568235"]],["YSU12340152","BIG OPEN LV PURSE",["348756387300"]],["YSU12340153","FLORAL DRESS FY-3092",["742042544578"]],["YSU12340154","FLORAL DRESS FY-3152",["742042547692"]]]);

function builtinMultiBarcodeMasters() {
  return BUILTIN_MULTI_BARCODE_ROWS.map(([masterCode, product, aliases]) => ({
    masterCode,
    product,
    vendor: "",
    plu: "",
    itemNumber: "",
    aliases,
  }));
}

function ensureBuiltinMultiBarcodesLoaded() {
  const builtins = builtinMultiBarcodeMasters();
  if (!builtins.length) return;
  const byMaster = new Map();
  (state.multiBarcodeMasters || []).forEach((row) => {
    const key = codeKey(row.masterCode);
    if (!key) return;
    byMaster.set(key, { ...row, aliases: [...(row.aliases || [])] });
  });
  let changed = false;
  builtins.forEach((row) => {
    const key = codeKey(row.masterCode);
    if (!key) return;
    const existing = byMaster.get(key);
    const mergedAliases = [...new Set([...(existing?.aliases || []), ...(row.aliases || [])]
      .map(normalizeCode)
      .filter(Boolean))];
    if (!existing) {
      byMaster.set(key, { ...row, aliases: mergedAliases });
      changed = true;
      return;
    }
    if (mergedAliases.length !== (existing.aliases || []).length) {
      existing.aliases = mergedAliases;
      changed = true;
    }
    if (!existing.product && row.product) {
      existing.product = row.product;
      changed = true;
    }
  });
  if (changed || !Object.keys(state.multiBarcodeMap || {}).length) {
    state.multiBarcodeMasters = [...byMaster.values()].sort((a, b) => compareDisplayValue(a.masterCode, b.masterCode));
    state.multiBarcodeFileName = state.multiBarcodeFileName || "Built-in multi barcode list";
    rebuildMultiBarcodeLookup();
  }
}
// END BUILTIN MULTI BARCODES


async function loadMultiBarcodeFile(file) {
  if (!file) return;
  try {
    els.excelStatus.textContent = `Reading ${file.name}...`;
    const xlsx = await ensureXlsxReader();
    if (!xlsx) {
      els.excelStatus.textContent = "Excel reader could not load.";
      return;
    }
    const workbook = xlsx.read(await file.arrayBuffer(), { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const matrix = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const masters = [];
    matrix.slice(1).forEach((row) => {
      const masterCode = normalizeCode(row[0]);
      const masterKey = rawCodeKey(masterCode);
      if (!masterCode || !masterKey) return;
      const aliases = row.slice(2, 18).map((value) => normalizeCode(value)).filter(Boolean);
      if (!aliases.length) return;
      const item = findCurrentItemByCode(masterCode) || findExcelFor({ code: masterCode }) || {};
      masters.push({
        masterCode,
        masterKey,
        product: item.product || "",
        vendor: item.vendor || "",
        plu: item.plu || "",
        itemNumber: item.itemNumber || "",
        aliases,
      });
    });
    state.multiBarcodeMasters = masters.sort((a, b) => compareDisplayValue(a.product || a.masterCode, b.product || b.masterCode));
    state.multiBarcodeFileName = file.name || "";
    rebuildMultiBarcodeLookup();
    els.excelStatus.textContent = `${number.format(state.multiBarcodeMasters.length)} multi-barcode master items loaded.`;
    logUploadedFile({ filename: file.name, type: "Multi_Barcode", status: "Success" });
    refreshDetailDrawer();
    renderInventory();
    renderOrders();
  } catch (error) {
    console.error(error);
    els.excelStatus.textContent = `Multi barcode import failed: ${error?.message || error}`;
    logUploadedFile({ filename: file?.name || "multi barcode", type: "Multi_Barcode", status: "Failed" });
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

function saveMultiBarcodeState() {
  safeLocalStorageSet("posDashboardMultiBarcodeMap:v1", JSON.stringify(state.multiBarcodeMap || {}));
  safeLocalStorageSet("posDashboardMultiBarcodeMasters:v1", JSON.stringify(state.multiBarcodeMasters || []));
  safeLocalStorageSet("posDashboardMultiBarcodeFileName:v1", state.multiBarcodeFileName || "");
  safeLocalStorageSet("posDashboardManualMultiBarcodes:v1", JSON.stringify(state.manualMultiBarcodes || {}));
}

function multiAliasesForCode(code) {
  const key = codeKey(code);
  if (!key) return [];
  const imported = (state.multiBarcodeMasters || []).find((row) => codeKey(row.masterCode) === key)?.aliases || [];
  const manual = state.manualMultiBarcodes?.[key] || [];
  return [...new Set([...imported.map(normalizeCode), ...manual.map(normalizeCode)].filter(Boolean))];
}

function rebuildMultiBarcodeLookup() {
  const map = {};
  (state.multiBarcodeMasters || []).forEach((row) => {
    const masterKey = rawCodeKey(row.masterCode);
    (row.aliases || []).forEach((alias) => {
      const aliasKey = normalizeCode(alias);
      if (!aliasKey) return;
      const resolvedKey = rawCodeKey(aliasKey);
      if (resolvedKey && resolvedKey !== masterKey) map[resolvedKey] = masterKey;
    });
  });
  Object.entries(state.manualMultiBarcodes || {}).forEach(([masterCode, aliases]) => {
    const masterKey = rawCodeKey(masterCode);
    (aliases || []).forEach((alias) => {
      const aliasKey = normalizeCode(alias);
      if (!aliasKey) return;
      const resolvedKey = rawCodeKey(aliasKey);
      if (resolvedKey && resolvedKey !== masterKey) map[resolvedKey] = masterKey;
    });
  });
  state.multiBarcodeMap = map;
  state._multiBarcodeIndexStamp += 1;
  state._priceCheckExactIndex = null;
  state._priceCheckExactIndexStamp = 0;
  saveMultiBarcodeState();
}

ensureBuiltinMultiBarcodesLoaded();

function saveItemMeta() {
  try {
    safeLocalStorageSet("posDashboardItemMetaMeta:v1", JSON.stringify({
      savedAt: new Date().toISOString(),
      items: Object.keys(state.itemMeta || {}).length,
    }));
    localStorage.removeItem(ITEM_META_STORAGE_KEY);
  } catch (_) {
    // Ignore localStorage quota issues; IndexedDB is the source of truth.
  }
  scheduleItemMetaPersist();
}

let itemMetaPersistTimer = 0;

function scheduleItemMetaPersist() {
  clearTimeout(itemMetaPersistTimer);
  itemMetaPersistTimer = setTimeout(() => {
    persistItemMetaToDb();
  }, 180);
}

async function persistItemMetaToDb() {
  try {
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readwrite");
      tx.objectStore(DB_STORE).put(state.itemMeta || {}, DB_ITEM_META_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
  } catch (error) {
    console.warn("Could not persist item metadata", error);
  }
}

function scheduleSharedProductSync(code) {
  if (!ENABLE_SHARED_SYNC) return;
  const normalizedCode = codeKey(code || "");
  if (normalizedCode) pendingSharedProductCodes.add(normalizedCode);
  clearTimeout(sharedSyncTimer);
  sharedSyncTimer = setTimeout(() => {
    const codes = [...pendingSharedProductCodes];
    pendingSharedProductCodes.clear();
    if (codes.length) {
      syncSharedProductsByCodes(codes, { silent: true }).catch(() => {});
      return;
    }
    syncSharedMetaSnapshotToSupabase({ silent: true, includeVendorRules: false }).catch(() => {});
  }, 600);
}

function scheduleSharedVendorRulesSync() {
  if (!ENABLE_SHARED_SYNC) return;
  clearTimeout(sharedVendorRulesTimer);
  sharedVendorRulesTimer = setTimeout(() => {
    syncSharedVendorRulesToSupabase(true).catch(() => {});
  }, 900);
}

let sharedImportLogsTimer = 0;
let sharedCountSessionsTimer = 0;
let sharedCountRetryTimer = 0;

function scheduleSharedImportLogsSync() {
  if (!ENABLE_SHARED_SYNC || !sharedImportLogsAvailable) return;
  clearTimeout(sharedImportLogsTimer);
  sharedImportLogsTimer = setTimeout(() => {
    syncSharedImportLogsToSupabase(true).catch(() => {});
  }, 900);
}

function scheduleSharedCountSessionsSync() {
  if (!ENABLE_SHARED_SYNC || !sharedCountSessionsAvailable) return;
  clearTimeout(sharedCountSessionsTimer);
  sharedCountSessionsTimer = setTimeout(() => {
    syncSharedCountSessionsToSupabase(true).catch(() => {});
  }, 650);
}

function scheduleCountSyncRetry() {
  if (!ENABLE_SHARED_SYNC || !sharedCountSessionsAvailable || navigator.onLine === false) return;
  clearTimeout(sharedCountRetryTimer);
  sharedCountRetryTimer = setTimeout(() => {
    syncSharedCountSessionsToSupabase(true).catch(() => {});
  }, COUNT_SYNC_RETRY_MS);
}

window.addEventListener("online", () => {
  if (state.activeCountSession?.localSyncPending || state.countSessions.some(countSessionHasUnsyncedEdits)) {
    scheduleSharedCountSessionsSync();
  }
});

function allowedItemStates() {
  return ["Active", "Force Order", "Disabled", "Discontinued"];
}

function normalizeItemState(value) {
  const raw = cleanCell(value);
  if (!raw) return "";
  const normalized = raw.toLowerCase().replace(/[_-]+/g, " ").replace(/\s+/g, " ").trim();
  if (normalized === "active") return "Active";
  if (
    normalized === "force order" ||
    normalized === "forceorder" ||
    normalized === "forced order" ||
    normalized === "forcedorder" ||
    normalized === "force reorder" ||
    normalized === "forcereorder" ||
    normalized === "forced reorder" ||
    normalized === "forcedreorder" ||
    normalized === "force"
  ) return "Force Order";
  if (normalized === "disabled") return "Disabled";
  if (normalized === "discontinued") return "Discontinued";
  return raw;
}

function normalizeItemDate(value) {
  const raw = cleanCell(value);
  if (!raw) return "";
  const iso = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (iso) return raw;
  const us = raw.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (us) {
    const [, mm, dd, yyyy] = us;
    return `${yyyy}-${mm.padStart(2, "0")}-${dd.padStart(2, "0")}`;
  }
  const parsed = Date.parse(raw);
  if (Number.isFinite(parsed)) {
    const d = new Date(parsed);
    return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}-${String(d.getDate()).padStart(2, "0")}`;
  }
  return raw;
}

function itemMetaFor(code) {
  return state.itemMeta?.[codeKey(code)] || {};
}

function isWigDefaultOrderingCategory(item = {}) {
  const category = normalizeSearchText(item.category || "");
  return [
    "WIG",
    "WIGS",
    "HUMAN WIG",
    "HUMAN WIGS",
    "HUMAN HAIR WIG",
    "HUMAN HAIR WIGS",
  ].includes(category);
}

function categoryDefaultOrderingTarget(item = {}) {
  if (!isWigDefaultOrderingCategory(item)) return null;
  return {
    min: 1,
    max: 1,
    safetyDays: 1,
    daysOfInventory: 1,
    categoryDefault: true,
  };
}

function orderingTargetsForItem({ item = {}, velocity = 0, safetyDays = 0, daysOfInventory = 0 } = {}) {
  const categoryTarget = categoryDefaultOrderingTarget(item);
  if (categoryTarget) return categoryTarget;
  return {
    ...orderingTargets({ velocity, safetyDays, daysOfInventory }),
    safetyDays,
    daysOfInventory,
    categoryDefault: false,
  };
}

function effectiveItemRule(item = {}) {
  const code = item.code || item.itemCode || "";
  const meta = itemMetaFor(code);
  const vendorName = (item.vendor || "").toUpperCase();
  const vendorRule = state.vendorRules.find((r) => r.vendor?.toUpperCase() === vendorName && r.status === "Active") || null;
  const mode = cleanCell(meta.orderingRuleMode).toLowerCase() === "custom" ? "custom" : "vendor";
  const vendorSafety = vendorRule
    ? Math.max(0, toNumber(vendorRule.safetyDays) || DEFAULT_ORDER_SAFETY_DAYS)
    : DEFAULT_ORDER_SAFETY_DAYS;
  const vendorDoi = vendorRule
    ? Math.max(0, toNumber(vendorRule.daysOfInventory) || DEFAULT_ORDER_DAYS_OF_INVENTORY)
    : DEFAULT_ORDER_DAYS_OF_INVENTORY;
  const customSafety = meta.safetyDaysOverride != null && meta.safetyDaysOverride !== ""
    ? Math.max(0, toNumber(meta.safetyDaysOverride) || 0)
    : vendorSafety;
  const customDoi = meta.daysOfInventoryOverride != null && meta.daysOfInventoryOverride !== ""
    ? Math.max(0, toNumber(meta.daysOfInventoryOverride) || 0)
    : vendorDoi;
  const safetyDays = mode === "custom" ? customSafety : vendorSafety;
  const daysOfInventory = mode === "custom" ? customDoi : vendorDoi;
  const categoryTarget = categoryDefaultOrderingTarget(item);
  if (categoryTarget) {
    return {
      mode: "category",
      vendorRule,
      safetyDays: categoryTarget.safetyDays,
      daysOfInventory: categoryTarget.daysOfInventory,
      customSafety,
      customDoi,
      vendorSafety,
      vendorDoi,
      categoryDefault: true,
    };
  }
  return {
    mode,
    vendorRule,
    safetyDays,
    daysOfInventory,
    customSafety,
    customDoi,
    vendorSafety,
    vendorDoi,
  };
}

function setItemMeta(code, patch = {}) {
  const key = codeKey(code);
  if (!key) return;
  state.itemMeta = state.itemMeta || {};
  state.itemMeta[key] = { ...(state.itemMeta[key] || {}), ...patch };
  saveItemMeta();
  scheduleSharedProductSync(key);
}

function seedItemMetaFromExcelRows(rows = []) {
  const snapshotDate = latestInventoryDate();
  let changed = false;
  rows.forEach((item) => {
    const key = codeKey(item.code);
    if (!key) return;
    const existing = state.itemMeta[key] || {};
    const excelAddDate = normalizeItemDate(item.addDate || "");
    const existingAddDate = normalizeItemDate(existing.addDate || "");
    const seededInventoryDate = existingAddDate && snapshotDate && existingAddDate === snapshotDate;
    const excelCaseSize = Math.max(1, Math.round(toNumber(item.caseSize) || 1));
    const hasExplicitExcelCaseSize = excelCaseSize > 1;
    const existingCaseSize = Math.max(1, Math.round(toNumber(existing.caseSize) || 1));
    const excelState = normalizeItemState(item.state || "");
    const existingState = normalizeItemState(existing.state || "");
    const merged = {
      ...existing,
      addDate: excelAddDate || (seededInventoryDate ? "" : existingAddDate) || "",
      // Data-file imports are the source of truth for item state.  Previously,
      // an old/manual Active value could block a later Disabled/Discontinued
      // value from the uploaded data file, leaving Products and count scope stale.
      state: excelState || (existing.stateManual ? existingState : ((existingState === "Active" && !existing.stateManual) ? "" : existingState) || ""),
      ...(excelState ? { stateManual: false } : {}),
      caseSize: hasExplicitExcelCaseSize ? excelCaseSize : (existing.caseSizeManual ? existingCaseSize : ((excelCaseSize > 1 || existingCaseSize <= 1) ? excelCaseSize : existingCaseSize)),
    };
    if (JSON.stringify(existing) !== JSON.stringify(merged)) {
      state.itemMeta[key] = merged;
      changed = true;
    }
  });
  if (changed) saveItemMeta();
}

function applyExcelStateToCurrentInventoryRows(rows = []) {
  // STATE-IMPORT PERF FIX ONLY:
  // The uploaded data workbook can contain 80k+ rows. The old version looped
  // every Excel row through every inventory snapshot row, which could lock the
  // tablet/browser and prevent the state update from finishing. Keep the same
  // behavior, but index the data-file states once and update matching inventory
  // rows in a single pass.
  const stateByCode = new Map();
  rows.forEach((item) => {
    const key = codeKey(item.code);
    const excelState = normalizeItemState(item.state || "");
    if (key && excelState) stateByCode.set(key, excelState);
  });
  if (!stateByCode.size) return false;

  let changed = false;
  state.latestInventory.forEach((inventory, key) => {
    const excelState = stateByCode.get(key);
    if (excelState && normalizeItemState(inventory.state || "") !== excelState) {
      inventory.state = excelState;
      changed = true;
    }
  });

  state.inventories.forEach((snapshotRows) => {
    snapshotRows.forEach((row) => {
      const excelState = stateByCode.get(codeKey(row.code));
      if (excelState && normalizeItemState(row.state || "") !== excelState) {
        row.state = excelState;
        changed = true;
      }
    });
  });
  return changed;
}

function registerInventorySnapshotMeta(snapshotDate, rows = [], previousCodes = new Set()) {
  let changed = false;
  rows.forEach((row) => {
    const key = codeKey(row.code);
    if (!key) return;
    const existing = state.itemMeta[key] || {};
    const isNew = !previousCodes.has(key) && !existing.addDate;
    const rowCaseSize = Math.max(1, Math.round(toNumber(row.caseSize) || 1));
    const existingCaseSize = Math.max(1, Math.round(toNumber(existing.caseSize) || 1));
    const merged = {
      ...existing,
      addDate: normalizeItemDate(existing.addDate || (isNew ? snapshotDate : "")),
      state: normalizeItemState(existing.state || row.state || ""),
      caseSize: rowCaseSize > 1 ? rowCaseSize : (existing.caseSizeManual ? existingCaseSize : (existing.caseSize ?? (toNumber(row.caseSize) > 0 ? toNumber(row.caseSize) : undefined))),
      firstSeenDate: existing.firstSeenDate || (isNew ? snapshotDate : ""),
      lastSeenDate: snapshotDate || existing.lastSeenDate || "",
    };
    if (JSON.stringify(existing) !== JSON.stringify(merged)) {
      state.itemMeta[key] = merged;
      changed = true;
    }
  });
  if (changed) saveItemMeta();
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
  const field = (...names) => {
    for (const name of names) {
      if (Object.prototype.hasOwnProperty.call(row, name) && cleanCell(row[name]) !== "") return row[name];
    }
    return "";
  };
  const code = normalizeCode(row.code ?? row.CODE);
  return {
    code,
    product: cleanCell(field("item_name", "ITEM NAME", "Item Name", "NAME", "Product", "PRODUCT")),
    vendor: cleanCell(field("vendor_name", "VENDOR NAME", "Vendor Name", "vendor", "VENDOR")),
    processingTime: toNumber(field("processing_time", "PROCESSING TIME", "Processing Time")),
    leadTime: toNumber(field("lead_time", "LEAD TIME", "Lead Time")),
    safetyStock: toNumber(field("safety_stock", "SAFETY STOCK", "Safety Stock")),
    daysOfInventory: toNumber(field("days_of_inventory", "DAYS OF INVENTORY", "Days of Inventory")),
    saleWindowSum: toNumber(field("sale_window_sum", "SALE WINDOW SUM", "Sale Window Sum")),
    saleVelocity: toNumber(field("sale_velocity", "SALE VELOCITY", "Sale Velocity")),
    stock: toNumber(field("stock", "STOCK")),
    reorderMin: toNumber(field("reorder_qty_min", "REORDER QTY MIN", "Reorder Qty Min", "MIN", "Min")),
    reorderMax: toNumber(field("reorder_qty_max", "REORDER QTY MAX", "Reorder Qty Max", "MAX", "Max")),
    daysBeforeRestock: toNumber(field("days_before_restock", "DAYS BEFORE RESTOCK", "Days Before Restock")),
    state: normalizeItemState(field("state", "STATE", "Status", "ITEM STATUS", "Item Status", "ITEM STATE", "Item State")),
    plu: cleanCell(field("PLU", "plu")),
    itemNumber: cleanCell(field("item_number", "ITEM NUMBER", "Item Number", "itemnum", "ITEMNUM")),
    category: cleanCell(field("category", "CATEGORY")),
    addDate: normalizeItemDate(field("add_date", "ADD DATE", "Add Date", "addDate", "DATE ADDED", "Date Added")),
    cost: toNumber(field("cost", "COST", "unit_cost", "UNIT COST", "Unit Cost")),
    price: toNumber(field("price", "PRICE")),
    caseSize: toNumber(field("case_size", "CASE SIZE", "Case Size", "case size", "CASE", "CASE QTY", "Case Qty")) || 1,
    maxOrderQty: toNumber(field("max_order_qty", "MAX ORDER QTY", "Max Order Qty")),
    qtyNeeded: toNumber(field("qty_needed", "QTY NEEDED", "Qty Needed")),
    orderPendingId: cleanCell(field("order_pending_id", "ORDER PENDING ID", "Order Pending Id")),
    orderPendingStale: cleanCell(field("order_pending_stale", "ORDER PENDING STALE", "Order Pending Stale")),
    poPendingClearedTimes: toNumber(field("PO_pending_cleared_times", "PO PENDING CLEARED TIMES", "PO Pending Cleared Times")),
    overrideValues: cleanCell(field("override_values", "OVERRIDE VALUES", "Override Values")),
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
  const field = (...names) => {
    for (const name of names) {
      if (Object.prototype.hasOwnProperty.call(row, name) && cleanCell(row[name]) !== "") return row[name];
    }
    return "";
  };
  const code = normalizeCode(field("CODE", "Code", "code", "UPC", "Barcode"));
  const product = cleanCell(field("PRODUCT", "Product", "product", "NAME", "Name", "name", "Item Name", "ITEM NAME"));
  const department = cleanCell(field("DEPARTMENT", "Department", "department", "DEPT", "Dept"));
  const category = cleanCell(field("::CAT::", "CATEGORY", "Category", "category"));
  const vendor = cleanCell(field("::VENDOR::", "VENDOR", "Vendor", "vendor"));
  const summaryLabel = `${product} ${department} ${category} ${vendor}`.toUpperCase();
  if (!code) return null;
  if (/\b(TOTAL|SUBTOTAL|SUMMARY|GRAND TOTAL|NET SALES|TAX|DISCOUNT)\b/.test(summaryLabel)) return null;
  return {
    date: date.iso,
    code,
    product,
    department: department || "Unassigned",
    category: category || "Unassigned",
    vendor: vendor || "Unassigned",
    units: toNumber(field("QTY", "Qty", "qty", "UNITS", "Units", "units")),
    sales: toNumber(field("SBTLS", "Subtotal", "SUBTOTAL", "Sales", "sales")),
    cost: toNumber(field("COST", "Cost", "cost", "Cost Sold", "costSold")),
    profit: toNumber(field("PRF", "Profit", "profit")),
  };
}

function normalizeInventoryRow(row, date) {
  const field = (...names) => {
    for (const name of names) {
      if (Object.prototype.hasOwnProperty.call(row, name) && cleanCell(row[name]) !== "") return row[name];
    }
    return "";
  };
  return {
    date: date.iso,
    code: normalizeCode(field("CODE", "code")),
    category: cleanCell(field("CATEGORY", "category")),
    product: cleanCell(field("NAME", "name", "ITEM NAME", "Item Name", "item_name", "product", "Product", "PRODUCT")),
    plu: cleanCell(field("PLU", "plu")),
    itemNumber: cleanCell(field("ITEM NUMBER", "item number", "ITEMNUM", "itemnum", "itemNumber", "Item Number")),
    price: toNumber(field("PRICE", "price")),
    cost: toNumber(field("COST", "cost", "unitCost", "Unit Cost", "unit_cost")),
    stock: toNumber(field("STOCK", "stock")),
    vendor: cleanCell(field("VENDOR", "vendor")),
    vendorCode: cleanCell(field("VENDOR CODE", "vendor code")),
    color: cleanCell(field("COLOR", "color")),
    size: cleanCell(field("SIZE", "size")),
    length: cleanCell(field("LENGTH", "length")),
    manufacture: cleanCell(field("MANUFACTURE", "manufacture")),
    memo: cleanCell(field("MEMO", "memo")),
    state: normalizeItemState(field("STATE", "state", "Status", "ITEM STATUS", "Item Status", "ITEM STATE", "Item State")),
    addDate: normalizeItemDate(field("ADD DATE", "add_date", "Add Date", "addDate", "DATE ADDED", "Date Added")),
    caseSize: toNumber(field("CASE SIZE", "case_size", "caseSize", "Case Size", "CASE", "CASE QTY", "Case Qty")) || 1,
  };
}

function mergeInventoryRowsByCode(rows = []) {
  const merged = new Map();
  rows.forEach((row) => {
    const key = codeKey(row.code);
    if (!key) return;
    const existing = merged.get(key);
    if (!existing) {
      merged.set(key, { ...row, stock: toNumber(row.stock) || 0 });
      return;
    }
    merged.set(key, {
      ...existing,
      ...Object.fromEntries(Object.entries(row).filter(([, value]) => cleanCell(value) !== "")),
      code: existing.code || row.code,
      product: existing.product || row.product,
      vendor: existing.vendor || row.vendor,
      category: existing.category || row.category,
      department: existing.department || row.department,
      stock: (toNumber(existing.stock) || 0) + (toNumber(row.stock) || 0),
    });
  });
  return [...merged.values()];
}

function buildLatestInventory() {
  state.latestInventory = new Map();
  [...state.inventories.entries()].sort(([a], [b]) => a.localeCompare(b)).forEach(([, rows]) => {
    mergeInventoryRowsByCode(rows).forEach((row) => {
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
  const pills = document.querySelector("#datePresets");
  if (commandBar) root.style.setProperty("--command-bar-height", `${commandBar.offsetHeight}px`);
  if (metrics) root.style.setProperty("--metrics-height", `${metrics.offsetHeight}px`);
  if (filters) root.style.setProperty("--filters-height", `${filters.offsetHeight}px`);
  if (pills) root.style.setProperty("--date-pills-height", `${pills.offsetHeight}px`);
}

function mountInventoryQuickTools() {
  if (!els.inventoryQuickTools || els.inventoryQuickTools.dataset.mounted === "true") return;
  const source = document.querySelector("[data-inventory-tools]");
  if (!source) return;
  const row = document.createElement("div");
  row.className = "inventory-quick-tools__row";
  els.inventoryQuickTools.append(row);
  els.inventoryQuickTools.dataset.mounted = "true";
  source.hidden = true;
  renderSharedQuickTools(activeTabName());
}

function ensureResizeColumnsButton() {
  let button = document.querySelector("#resizeColumnsButton");
  if (button) return button;
  button = document.createElement("button");
  button.type = "button";
  button.id = "resizeColumnsButton";
  button.className = "secondary-button";
  button.textContent = "Resize columns";
  button.addEventListener("click", () => {
    state.resizeColumns = !state.resizeColumns;
    button.classList.toggle("active-edit", state.resizeColumns);
    button.textContent = state.resizeColumns ? "Done resizing" : "Resize columns";
    renderInventoryHeader();
  });
  return button;
}

function ensureInventoryTotalsInline() {
  let node = document.querySelector("#inventoryTotalsInline");
  if (node) return node;
  node = document.createElement("div");
  node.id = "inventoryTotalsInline";
  node.className = "inventory-totals-inline";
  return node;
}

function ensureInventoryQuickChipsInline() {
  let node = document.querySelector("#inventoryQuickChipsInline");
  if (node) return node;
  node = document.createElement("div");
  node.id = "inventoryQuickChipsInline";
  node.className = "inventory-quick-chips-inline";
  return node;
}

function renderSharedQuickTools(tab = activeTabName()) {
  if (!els.inventoryQuickTools) return;
  const row = els.inventoryQuickTools.querySelector(".inventory-quick-tools__row");
  if (!row) return;
  const resizeColumnsButton = ensureResizeColumnsButton();
  const inventoryColumnPicker = els.inventoryColumnPicker;
  const orderColumnPicker = document.querySelector("#orderColumnPicker");
  const orderArrangeButton = document.querySelector("#orderArrangeColumnsButton");
  const exportPoExcelButton = document.querySelector("#exportPoExcel");
  const exportPoPdfButton = document.querySelector("#exportPoPdf");
  const downloadOrderButton = document.querySelector("#downloadOrder");
  const orderActions = document.querySelector(".po-actions");
  const formulaNote = document.querySelector("#formulaNote");
  const isOrdering = tab === "ordering";
  const isInventory = tab === "inventory";

  els.inventoryQuickTools.hidden = !(isOrdering || isInventory);
  els.inventoryQuickTools.classList.toggle("inventory-quick-tools--ordering", isOrdering);
  els.inventoryQuickTools.classList.toggle("inventory-quick-tools--inventory", isInventory);
  if (orderActions) {
    orderActions.hidden = true;
    orderActions.style.display = "none";
  }
  if (formulaNote) {
    formulaNote.hidden = true;
    formulaNote.style.display = "none";
  }

  if (isOrdering) {
    const orderingControls = [
      els.inventoryStateFilter,
      orderArrangeButton,
      orderColumnPicker,
      exportPoExcelButton,
      exportPoPdfButton,
      downloadOrderButton,
    ].filter(Boolean);
    orderingControls.forEach((control) => {
      control.hidden = false;
      control.style.display = "";
    });
    row.replaceChildren(...orderingControls);
    return;
  }

  if (isInventory) {
    resizeColumnsButton.classList.toggle("active-edit", state.resizeColumns);
    resizeColumnsButton.textContent = state.resizeColumns ? "Done resizing" : "Resize columns";
    const quickChips = ensureInventoryQuickChipsInline();
    row.replaceChildren(...[
      els.inventoryStateFilter,
      els.arrangeColumnsButton,
      resizeColumnsButton,
      inventoryColumnPicker,
      els.downloadInventory,
      els.createPoShortcut,
      quickChips,
    ].filter(Boolean));
    return;
  }

  row.replaceChildren();
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

function buildPriceCheckEntry(inventory = {}, excel = {}) {
  const itemCode = inventory.code || excel.code || "";
  const meta = itemMetaFor(itemCode);
  const stateLabel = normalizeItemState(meta.stateManual ? meta.state : (inventory.state || excel.state || meta.state || "")) || "Active";
  const rule = effectiveItemRule({
    code: itemCode,
    vendor: inventory.vendor || excel.vendor || "",
  });
  return {
    code: itemCode,
    product: bestItemName(inventory.product, excel.product, "", inventory.plu, itemCode),
    department: inventory.department || excel.department || "",
    category: inventory.category || excel.category || "",
    vendor: inventory.vendor || excel.vendor || "",
    plu: inventory.plu || excel.plu || "",
    itemNumber: inventory.itemNumber || excel.itemNumber || "",
    color: inventory.color || excel.color || "",
    state: stateLabel,
    itemState: stateLabel.toLowerCase(),
    addDate: cleanCell(meta.addDate || excel.addDate || inventory.addDate || meta.firstSeenDate || ""),
    snapshotDate: inventory.date || "",
    stock: Number(inventory.stock ?? excel.stock ?? 0),
    units: 0,
    sales: 0,
    costSold: 0,
    profit: 0,
    velocity: 0,
    unitCost: pickNumber(inventory.cost, excel.cost),
    price: pickNumber(inventory.price, excel.price),
    caseSize: effectiveCaseSizeFor(inventory, excel, meta),
    reorderMin: toNumber(reorderOverrideForCode(itemCode)?.min) || 0,
    reorderMax: toNumber(reorderOverrideForCode(itemCode)?.max) || 0,
    recommendedOrder: 0,
    caseOrder: 0,
    inventoryCost: Number(inventory.stock ?? excel.stock ?? 0) * pickNumber(inventory.cost, excel.cost),
    effectiveSafetyDays: rule.safetyDays,
    effectiveDaysOfInventory: rule.daysOfInventory,
    orderingRuleMode: rule.mode,
  };
}

function ensurePriceCheckExactIndex() {
  const stamp = `${state._dataCacheStamp}:${state._multiBarcodeIndexStamp || 0}`;
  if (state._priceCheckExactIndex && state._priceCheckExactIndexStamp === stamp) return;
  const startedAt = performanceNow();
  const exact = new Map();
  const indexEntry = (entry) => {
    [entry.code, entry.plu, entry.itemNumber].forEach((value) => {
      const key = codeKey(value);
      if (key && !exact.has(key)) exact.set(key, entry);
    });
  };
  [...state.latestInventory.values()].forEach((inventory) => {
    const entry = buildPriceCheckEntry(inventory, findExcelFor(inventory));
    if (entry.code || entry.plu || entry.itemNumber) indexEntry(entry);
  });
  [...state.excelItems.values()].forEach((excel) => {
    const probeKey = codeKey(excel.code || excel.plu || excel.itemNumber || "");
    if (probeKey && exact.has(probeKey)) return;
    const entry = buildPriceCheckEntry({}, excel);
    if (entry.code || entry.plu || entry.itemNumber) indexEntry(entry);
  });
  Object.entries(state.multiBarcodeMap || {}).forEach(([alias, masterCode]) => {
    const aliasKey = codeKey(alias);
    const masterKey = codeKey(masterCode);
    const item = masterKey ? exact.get(masterKey) : null;
    if (aliasKey && item && !exact.has(aliasKey)) exact.set(aliasKey, item);
  });
  state._priceCheckExactIndex = exact;
  state._priceCheckExactIndexStamp = stamp;
  console.info("[AppPerf] barcode index ready", {
    ms: Number((performanceNow() - startedAt).toFixed(2)),
    entries: exact.size,
  });
}

function priceCheckExactMatch(query) {
  const keyed = codeKey(query);
  if (!keyed) return null;
  ensurePriceCheckExactIndex();
  return state._priceCheckExactIndex?.get(keyed) || null;
}

function priceCheckRows() {
  if (state._priceCheckRowsCache && state._priceCheckRowsStamp === state._dataCacheStamp) {
    return state._priceCheckRowsCache;
  }
  const rowMap = new Map();
  [...state.latestInventory.values()].forEach((inventory) => {
    const itemCode = inventory.code || "";
    const key = codeKey(itemCode || inventory.plu || inventory.itemNumber || inventory.product);
    if (!key) return;
    rowMap.set(key, buildPriceCheckEntry(inventory, findExcelFor(inventory)));
  });
  [...state.excelItems.values()].forEach((excel) => {
    const key = codeKey(excel.code || excel.plu || excel.itemNumber || excel.product);
    if (!key || rowMap.has(key)) return;
    rowMap.set(key, buildPriceCheckEntry({}, excel));
  });
  const rows = [...rowMap.values()];
  state._priceCheckRowsCache = rows;
  state._priceCheckRowsStamp = state._dataCacheStamp;
  ensurePriceCheckExactIndex();
  return rows;
}

function ensureZxingReader() {
  if (window.ZXingBrowser) return Promise.resolve(window.ZXingBrowser);
  return new Promise((resolve) => {
    const existing = document.querySelector("script[data-zxing-fallback]");
    const script = existing || document.createElement("script");
    const done = () => resolve(window.ZXingBrowser || null);
    script.onload = done;
    script.onerror = () => resolve(null);
    if (!existing) {
      script.dataset.zxingFallback = "true";
      script.src = "https://cdn.jsdelivr.net/npm/@zxing/browser@0.1.5/umd/index.min.js";
      document.head.append(script);
    }
    setTimeout(done, 10000);
  });
}

function ensureHtml5Qrcode() {
  if (window.Html5Qrcode) return Promise.resolve(window.Html5Qrcode);
  return new Promise((resolve) => {
    const existing = document.querySelector("script[data-html5-qrcode]");
    const script = existing || document.createElement("script");
    const done = () => resolve(window.Html5Qrcode || null);
    script.onload = done;
    script.onerror = () => resolve(null);
    if (!existing) {
      script.dataset.html5Qrcode = "true";
      script.src = "https://unpkg.com/html5-qrcode@2.3.8/html5-qrcode.min.js";
      document.head.append(script);
    }
    setTimeout(done, 10000);
  });
}

function prefersPhoneBarcodeScanner() {
  const ua = navigator.userAgent || "";
  return /android|iphone|ipad|ipod|mobile/i.test(ua) || (window.matchMedia?.("(max-width: 900px)")?.matches ?? false);
}

function priceCheckVideoEl() {
  return state.priceCheckFullscreen ? els.priceCheckOverlayVideo : els.priceCheckVideo;
}

function priceCheckScannerEl() {
  return state.priceCheckFullscreen ? els.priceCheckOverlayScanner : els.priceCheckScanner;
}

function setPriceCheckFullscreen(open) {
  state.priceCheckFullscreen = !!open;
  if (els.priceCheckOverlay) els.priceCheckOverlay.hidden = !open;
}

function setPriceCheckTorchButtons(supported) {
  [els.priceCheckTorchButton, els.priceCheckOverlayTorchButton].forEach((button) => {
    if (!button) return;
    button.hidden = !supported;
    button.textContent = state.priceCheckTorchOn ? "Light Off" : "Light";
  });
}

function setSharedScanStatus(message) {
  if (els.priceCheckStatus) els.priceCheckStatus.textContent = message;
  if (els.scanModeStatus) {
    els.scanModeStatus.textContent = message;
    delete els.scanModeStatus.dataset.readiness;
  }
}

async function togglePriceCheckTorch() {
  const track = state.priceCheckStream?.getVideoTracks?.()?.[0];
  const caps = track?.getCapabilities?.();
  if (!track || !caps?.torch) return;
  try {
    state.priceCheckTorchOn = !state.priceCheckTorchOn;
    await track.applyConstraints({ advanced: [{ torch: state.priceCheckTorchOn }] });
    setPriceCheckTorchButtons(true);
  } catch (error) {
    state.priceCheckTorchOn = false;
    setPriceCheckTorchButtons(false);
    showToast("This camera does not allow flash control here.", 2600, "warning");
  }
}

function priceCheckMatches(query, limit = 12) {
  const raw = cleanCell(query);
  if (!raw) return [];
  const keyed = codeKey(raw);
  const exact = priceCheckExactMatch(raw);
  if (exact) return [exact];
  const rows = priceCheckRows();
  const search = raw.toLowerCase();
  const scored = [];
  for (const item of rows) {
    const code = cleanCell(item.code);
    const plu = cleanCell(item.plu);
    const itemNumber = cleanCell(item.itemNumber);
    const product = cleanCell(item.product);
    const vendor = cleanCell(item.vendor);
    const category = cleanCell(item.category);
    const department = cleanCell(item.department);
    const keys = [code, plu, itemNumber].map(codeKey).filter(Boolean);
    let score = -1;
    if (keyed && keys.includes(keyed)) score = 0;
    else if ([code, plu, itemNumber].some((value) => value && value.toLowerCase().startsWith(search))) score = 1;
    else if (product && product.toLowerCase().startsWith(search)) score = 2;
    else if ([code, product, plu, itemNumber, vendor, category, department].some((value) => value && value.toLowerCase().includes(search))) score = 3;
    if (score >= 0) scored.push({ item, score, product });
  }
  return scored
    .sort((a, b) => a.score - b.score || a.product.localeCompare(b.product))
    .slice(0, limit)
    .map((entry) => entry.item);
}

function findPriceCheckItem(query) {
  return priceCheckMatches(query, 1)[0] || null;
}

function hidePriceCheckDropdown() {
  if (els.priceCheckDropdown) els.priceCheckDropdown.hidden = true;
}

function selectPriceCheckDropdownItem(key) {
  const item = priceCheckRows().find((entry) => {
    const itemKey = codeKey(entry.code || entry.plu || entry.itemNumber || entry.product);
    return itemKey === codeKey(key);
  });
  if (!item) return;
  if (els.priceCheckSearchInput) {
    els.priceCheckSearchInput.value = cleanCell(item.code || item.plu || item.itemNumber || item.product);
  }
  hidePriceCheckDropdown();
  renderPriceCheckResult(item);
  if (els.priceCheckStatus) els.priceCheckStatus.textContent = `Loaded ${item.code || item.product}. Ready for next scan.`;
  focusPriceCheckSearch();
}

function renderPriceCheckDropdown(query) {
  const dd = els.priceCheckDropdown;
  if (!dd) return;
  const raw = cleanCell(query).trim();
  if (!raw) {
    dd.hidden = true;
    return;
  }
  const matches = priceCheckMatches(raw, 10);
  if (!matches.length) {
    dd.hidden = true;
    return;
  }
  dd.innerHTML = matches.map((item, index) => {
    const itemKey = cleanCell(item.code || item.plu || item.itemNumber || item.product);
    const meta = [
      cleanCell(item.code),
      cleanCell(item.plu) ? `PLU ${cleanCell(item.plu)}` : "",
      cleanCell(item.vendor),
    ].filter(Boolean).join("  -  ");
    return `<div class="price-check-dd-item${index === 0 ? " price-check-dd-item--active" : ""}" data-key="${escapeHtml(itemKey)}">
      <span class="price-check-dd-name">${escapeHtml(item.product || item.code || "Unknown item")}</span>
      <span class="price-check-dd-meta">${escapeHtml(meta)}</span>
    </div>`;
  }).join("");
  dd.hidden = false;
  dd.querySelectorAll(".price-check-dd-item").forEach((node) => {
    node.addEventListener("mousedown", (event) => {
      event.preventDefault();
      selectPriceCheckDropdownItem(node.dataset.key);
    });
  });
}

function renderPriceCheckResult(item) {
  if (!els.priceCheckResult && !els.scanModeResult) return;
  const renderEmpty = () => `<div class="price-check-result__empty">Search an item to show price, stock, vendor, PLU, and item details.</div>`;
  if (!item) {
    if (els.priceCheckResult) {
      els.priceCheckResult.className = "price-check-result empty";
      els.priceCheckResult.innerHTML = renderEmpty();
    }
    if (els.scanModeResult) {
      els.scanModeResult.className = "price-check-result empty";
      els.scanModeResult.innerHTML = renderEmpty();
    }
    return;
  }
  const inventory = state.latestInventory.get(codeKey(item.code)) || item;
  const excel = findExcelFor(item);
  const price = Number(item.price || excel.price || 0);
  const stock = Number(inventory.stock || excel.stock || 0);
  const cost = Number(item.unitCost || excel.cost || 0);
  const vendor = cleanCell(item.vendor || excel.vendor) || "Unassigned";
  const plu = cleanCell(item.plu || excel.plu) || "-";
  const itemNumber = cleanCell(item.itemNumber || excel.itemNumber) || "-";
  const category = cleanCell(item.category || excel.category) || "Unassigned";
  const department = cleanCell(item.department || excel.department) || "Unassigned";
  const showCost = isAdmin();
  const html = `
    <div class="price-check-result__hero">
      <div>
        <p class="eyebrow">Item found</p>
        <h3>${escapeHtml(item.product || excel.product || item.code || "Unknown item")}</h3>
      </div>
      <div class="price-check-result__price">${currency.format(price)}</div>
    </div>
    <div class="price-check-result__grid">
      <article><span>In stock</span><strong>${number.format(stock)}</strong></article>
      <article><span>Code</span><strong>${escapeHtml(item.code || "-")}</strong></article>
      <article><span>Vendor</span><strong>${escapeHtml(vendor)}</strong></article>
      <article><span>PLU</span><strong>${escapeHtml(plu)}</strong></article>
      <article><span>Item #</span><strong>${escapeHtml(itemNumber)}</strong></article>
      <article><span>Category</span><strong>${escapeHtml(category)}</strong></article>
      <article><span>Department</span><strong>${escapeHtml(department)}</strong></article>
      ${showCost ? `<article><span>Cost</span><strong>${currency.format(cost)}</strong></article>` : ""}
    </div>
  `;
  if (els.priceCheckResult) {
    els.priceCheckResult.className = "price-check-result";
    els.priceCheckResult.innerHTML = html;
  }
  if (els.scanModeResult) {
    els.scanModeResult.className = "price-check-result";
    els.scanModeResult.innerHTML = html;
  }
}

function focusPriceCheckSearch(select = true) {
  if (!els.priceCheckSearchInput) return;
  els.priceCheckSearchInput.focus();
  if (select) setTimeout(() => els.priceCheckSearchInput?.select?.(), 0);
}

function clearPriceCheckSearch() {
  if (els.priceCheckSearchInput) els.priceCheckSearchInput.value = "";
  hidePriceCheckDropdown();
  renderPriceCheckResult(null);
  setSharedScanStatus("Ready for next scan.");
  focusPriceCheckSearch();
}

function handlePriceCheckLookup(options = {}) {
  const { refocus = false, silentNotFound = false } = options;
  const query = els.priceCheckSearchInput?.value || "";
  const matches = priceCheckMatches(query, 10);
  const item = matches[0] || null;
  if (matches.length > 1) renderPriceCheckDropdown(query);
  else hidePriceCheckDropdown();
  if (!item) {
    renderPriceCheckResult(null);
    setSharedScanStatus("Item not found. Try barcode, PLU, item #, or name.");
    if (!silentNotFound) showToast("Price check item not found.", 2600, "warning");
    if (refocus) focusPriceCheckSearch();
    return null;
  }
  renderPriceCheckResult(item);
  setSharedScanStatus(`Loaded ${item.code || item.product}. Ready for next scan.`);
  if (refocus) focusPriceCheckSearch();
  return item;
}

function processPriceCheckScan(code) {
  const cleanCode = cleanCell(code || "");
  const now = Date.now();
  if (!cleanCode) return false;
  if (cleanCode === state.priceCheckLastCode && now - state.priceCheckLastScanAt <= 1200) return false;
  state.priceCheckLastCode = cleanCode;
  state.priceCheckLastScanAt = now;
  const inScanMode = activeTabName() === "scanmode";
  let item = null;
  if (inScanMode) {
    const scanInput = document.querySelector("#scanModeInput");
    if (scanInput) scanInput.value = cleanCode;
    item = handleScanModeLookup(cleanCode, { refocus: false, silentNotFound: true });
  } else {
    if (els.priceCheckSearchInput) els.priceCheckSearchInput.value = cleanCode;
    item = handlePriceCheckLookup({ refocus: false, silentNotFound: true });
  }
  if (inScanMode) stopPriceCheckCamera();
  setSharedScanStatus(item
    ? (inScanMode
      ? `Loaded ${item.code || item.product}. Tap Start Scanner for next item.`
      : `Loaded ${item.code || item.product}. Scan next item.`)
    : (inScanMode
      ? `No match for ${cleanCode}. Tap Start Scanner to try again.`
      : `No match for ${cleanCode}. Scan next item.`));
  return true;
}

async function startPriceCheckCamera(options = {}) {
  const { fullscreen = false, forceFullscreen = false } = options;
  if (!window.isSecureContext) {
    stopPriceCheckCamera();
    showToast("Camera scanning requires the secure live website. Open the GitHub Pages URL, not a local file.", 4200, "warning");
    return;
  }
  if (!navigator.mediaDevices?.getUserMedia) {
    stopPriceCheckCamera();
    showToast("Camera scanning is not supported on this device/browser.", 3200, "warning");
    return;
  }
  try {
    stopPriceCheckCamera();
    setPriceCheckFullscreen(!!forceFullscreen || (fullscreen && prefersPhoneBarcodeScanner()));
    els.priceCheckSearchInput?.blur?.();
    setSharedScanStatus("Starting camera...");
    const tryStream = async (videoOptions) => navigator.mediaDevices.getUserMedia({ video: videoOptions, audio: false });
    const cameraProfile = scanPerformanceProfile.camera;
    const optimizedVideo = {
      facingMode: { ideal: "environment" },
      width: { ideal: cameraProfile.width, max: cameraProfile.width },
      height: { ideal: cameraProfile.height, max: cameraProfile.height },
      frameRate: { ideal: cameraProfile.fps, max: cameraProfile.fps },
    };
    console.info("[ScanPerf] camera", {
      tier: scanPerformanceProfile.tier,
      width: cameraProfile.width,
      height: cameraProfile.height,
      fps: cameraProfile.fps,
    });
    try {
      state.priceCheckStream = await tryStream({ ...optimizedVideo, facingMode: { exact: "environment" } });
    } catch (primaryError) {
      state.priceCheckStream = await tryStream(optimizedVideo);
    }
    state.priceCheckTorchOn = false;
    const videoEl = priceCheckVideoEl();
    if (videoEl) {
      videoEl.setAttribute("playsinline", "");
      videoEl.setAttribute("autoplay", "");
      videoEl.muted = true;
      videoEl.srcObject = state.priceCheckStream;
      await new Promise((resolve) => {
        if (videoEl.readyState >= 1) {
          resolve();
          return;
        }
        videoEl.onloadedmetadata = () => resolve();
        setTimeout(resolve, 800);
      });
      await videoEl.play();
    }
    const activeTrack = state.priceCheckStream?.getVideoTracks?.()?.[0];
    const caps = activeTrack?.getCapabilities?.();
    setPriceCheckTorchButtons(!!caps?.torch);
    const scannerEl = priceCheckScannerEl();
    if (scannerEl) {
      scannerEl.hidden = true;
      scannerEl.innerHTML = "";
    }
    setSharedScanStatus("Camera live. Point at a barcode.");
    if (els.priceCheckStopButton) els.priceCheckStopButton.hidden = false;
    if (els.priceCheckCameraButton) els.priceCheckCameraButton.hidden = true;
    if ("BarcodeDetector" in window) {
      state.priceCheckDetector = new BarcodeDetector({
        formats: ["upc_a", "upc_e", "ean_13", "ean_8", "code_128"],
      });
      if (els.priceCheckStatus || els.scanModeStatus) {
        setSharedScanStatus(prefersPhoneBarcodeScanner()
          ? "Camera live. Hold the barcode inside the box and move slightly closer."
          : "Camera live. Point at a barcode.");
      }
      scanPriceCheckFrame();
      return;
    }
    const Html5Qrcode = await ensureHtml5Qrcode();
    if (Html5Qrcode && scannerEl) {
      try {
        state.priceCheckStream?.getTracks().forEach((track) => track.stop());
        state.priceCheckStream = null;
        setPriceCheckTorchButtons(false);
        if (videoEl) videoEl.srcObject = null;
        scannerEl.hidden = false;
        scannerEl.innerHTML = "";
        const scannerId = state.priceCheckFullscreen ? "priceCheckOverlayScanner" : "priceCheckScanner";
        state.priceCheckScanner = new Html5Qrcode(scannerId);
        const supportedFormats = window.Html5QrcodeSupportedFormats
          ? [
              window.Html5QrcodeSupportedFormats.UPC_A,
              window.Html5QrcodeSupportedFormats.UPC_E,
              window.Html5QrcodeSupportedFormats.EAN_13,
              window.Html5QrcodeSupportedFormats.EAN_8,
              window.Html5QrcodeSupportedFormats.CODE_128,
            ].filter(Boolean)
          : undefined;
        await state.priceCheckScanner.start(
          { facingMode: "environment" },
          {
            fps: cameraProfile.fps,
            qrbox: cameraProfile.qrbox,
            rememberLastUsedCamera: true,
            aspectRatio: 1.7778,
            disableFlip: true,
            formatsToSupport: supportedFormats,
            videoConstraints: optimizedVideo,
            experimentalFeatures: { useBarCodeDetectorIfSupported: true },
          },
          (decodedText) => {
            processPriceCheckScan(decodedText);
          }
        );
        setSharedScanStatus("Camera live. Hold the barcode inside the box and move slightly closer.");
        return;
      } catch (error) {
        state.priceCheckScanner = null;
        if (scannerEl) {
          scannerEl.hidden = true;
          scannerEl.innerHTML = "";
        }
      }
    }
    const ZXingBrowser = await ensureZxingReader();
    if (ZXingBrowser) {
      state.priceCheckReader = new ZXingBrowser.BrowserMultiFormatReader();
      state.priceCheckReader.decodeFromVideoDevice(undefined, videoEl, (result, error, controls) => {
        if (controls) state.priceCheckReaderControls = controls;
        processPriceCheckScan(result?.getText?.() || "");
      });
      return;
    }
    stopPriceCheckCamera();
    showToast("Camera preview opened, but this browser could not start barcode decoding. Manual search is still available.", 4200, "warning");
  } catch (error) {
    stopPriceCheckCamera();
    showToast("Camera could not start. Allow camera access and try again.", 3600, "warning");
  }
}

function stopPriceCheckCamera() {
  if (state.priceCheckRaf) {
    cancelAnimationFrame(state.priceCheckRaf);
    state.priceCheckRaf = 0;
  }
  if (state.priceCheckScanTimer) {
    clearTimeout(state.priceCheckScanTimer);
    state.priceCheckScanTimer = 0;
  }
  if (state.priceCheckStream) {
    state.priceCheckStream.getTracks().forEach((track) => track.stop());
    state.priceCheckStream = null;
  }
  if (state.priceCheckReaderControls?.stop) {
    try { state.priceCheckReaderControls.stop(); } catch (error) {}
  }
  state.priceCheckReaderControls = null;
  if (state.priceCheckReader?.reset) {
    try { state.priceCheckReader.reset(); } catch (error) {}
  }
  state.priceCheckReader = null;
  state.priceCheckDetector = null;
  if (state.priceCheckScanner) {
    try {
      const stop = typeof state.priceCheckScanner.stop === "function" ? state.priceCheckScanner.stop() : null;
      if (stop?.catch) stop.catch(() => {});
      if (typeof state.priceCheckScanner.clear === "function") state.priceCheckScanner.clear();
    } catch (error) {}
  }
  state.priceCheckScanner = null;
  [els.priceCheckVideo, els.priceCheckOverlayVideo].forEach((videoEl) => {
    if (!videoEl) return;
    videoEl.pause?.();
    videoEl.srcObject = null;
  });
  [els.priceCheckScanner, els.priceCheckOverlayScanner].forEach((scannerEl) => {
    if (!scannerEl) return;
    scannerEl.hidden = true;
    scannerEl.innerHTML = "";
  });
  state.priceCheckTorchOn = false;
  setPriceCheckTorchButtons(false);
  setPriceCheckFullscreen(false);
  if (els.priceCheckStopButton) els.priceCheckStopButton.hidden = true;
  if (els.priceCheckCameraButton) els.priceCheckCameraButton.hidden = false;
  if (els.priceCheckStatus && activeTabName() === "pricecheck") els.priceCheckStatus.textContent = "Ready for next scan.";
}

async function scanPriceCheckFrame() {
  const interval = scanPerformanceProfile.detectIntervalMs || 67;
  const videoEl = priceCheckVideoEl();
  if (!state.priceCheckDetector || !videoEl || videoEl.readyState < 2) {
    state.priceCheckScanTimer = setTimeout(scanPriceCheckFrame, interval);
    return;
  }
  try {
    const codes = await state.priceCheckDetector.detect(videoEl);
    const match = codes.find((entry) => cleanCell(entry.rawValue));
    if (match) {
      processPriceCheckScan(match.rawValue);
    }
  } catch (error) {
    // keep looping silently; browsers may throw transient detect errors
  }
  state.priceCheckScanTimer = setTimeout(scanPriceCheckFrame, interval);
}

function latestExcelAddDate() {
  return [...state.excelItems.values()]
    .map((item) => cleanCell(item.addDate))
    .filter(Boolean)
    .sort(compareDateValue)
    .at(-1) || "";
}

function currentOrderRows(options = {}) {
  pruneOrderSubmissionDrafts();
  if (!options.ignoreSubmissionDrafts && state.orderSubmissionVendors?.length && Object.keys(state.orderSubmissionDrafts || {}).length) {
    const vendorKeys = state.orderSubmissionActiveVendor ? [state.orderSubmissionActiveVendor] : state.orderSubmissionVendors;
    return vendorKeys
      .flatMap((vendor) => (state.orderSubmissionDrafts[vendor.toUpperCase()] || []).map((item) => applyOrderOverride({ ...item })))
      .filter((item) => !isPendingOrder(item.code))
      .sort((a, b) => (b.recommendedOrder || b.qtyNeeded || 0) - (a.recommendedOrder || a.qtyNeeded || 0));
  }
  const vendorFilter = getOrderVendorFilter ? getOrderVendorFilter() : "Active";
  const manualVendorSelected = cleanCell(els.vendorFilter?.value || "").toUpperCase() !== "ALL";
  return buildInventoryRows(options)
    .map(applyOrderOverride)
    .filter((item) => {
      if (item.isOrderable === false) return false;
      if (isPendingOrder(item.code)) return false;
      if (item.recommendedOrder <= 0 && !item.qtyNeeded) return false;
      if (state.vendorRules.length && vendorFilter !== "" && !manualVendorSelected) {
        const vendorName = (item.vendor || "").toUpperCase();
        const rule = state.vendorRules.find((r) => r.vendor?.toUpperCase() === vendorName);
        if (vendorFilter === "Active" && rule && rule.status !== "Active") return false;
        if (vendorFilter === "Disabled" && (!rule || rule.status !== "Disabled")) return false;
      }
      if (state.orderSubmissionVendors?.length) {
        const submitted = state.orderSubmissionVendors.map((vendor) => vendor.toUpperCase());
        if (!submitted.includes((item.vendor || "").toUpperCase())) return false;
      }
      if (state.orderSubmissionActiveVendor && (item.vendor || "").toUpperCase() !== state.orderSubmissionActiveVendor.toUpperCase()) return false;
      return true;
    })
    .sort((a, b) => (b.recommendedOrder || b.qtyNeeded || 0) - (a.recommendedOrder || a.qtyNeeded || 0));
}

function buildOrderSubmissionDrafts(vendors, rows) {
  const drafts = { ...(state.orderSubmissionDrafts || {}) };
  vendors.forEach((vendor) => {
    drafts[vendor.toUpperCase()] = rows
      .filter((item) => (item.vendor || "").toUpperCase() === vendor.toUpperCase())
      .map((item) => ({ ...item }));
  });
  return drafts;
}

function pruneOrderSubmissionDrafts() {
  const drafts = { ...(state.orderSubmissionDrafts || {}) };
  const vendors = (state.orderSubmissionVendors || []).filter((vendor) => {
    const key = String(vendor || "").toUpperCase();
    const hasOpenRows = (drafts[key] || []).some((item) => !isPendingOrder(item.code));
    if (!hasOpenRows) delete drafts[key];
    return hasOpenRows;
  });
  if (vendors.length !== (state.orderSubmissionVendors || []).length) {
    state.orderSubmissionVendors = vendors;
    state.orderSubmissionDrafts = drafts;
    if (state.orderSubmissionActiveVendor && !vendors.some((vendor) => vendor.toUpperCase() === state.orderSubmissionActiveVendor.toUpperCase())) {
      state.orderSubmissionActiveVendor = "";
    }
  }
}

function removeSubmittedVendor(vendorName, options = {}) {
  const normalized = String(vendorName || "").trim().toUpperCase();
  if (!normalized) return;
  state.orderSubmissionVendors = (state.orderSubmissionVendors || []).filter((vendor) => vendor.toUpperCase() !== normalized);
  const drafts = { ...(state.orderSubmissionDrafts || {}) };
  delete drafts[normalized];
  state.orderSubmissionDrafts = drafts;
  if ((state.orderSubmissionActiveVendor || "").toUpperCase() === normalized) {
    state.orderSubmissionActiveVendor = "";
  }
  if (options.render !== false) renderOrders();
}

function saveDismissedOrderVendors() {
  safeLocalStorageSet("posDashboardDismissedOrderVendors:v1", JSON.stringify(state.orderDismissedVendors || []));
}

function dismissOrderVendor(vendorName) {
  const normalized = String(vendorName || "").trim().toUpperCase();
  if (!normalized) return;
  const next = new Set((state.orderDismissedVendors || []).map((vendor) => String(vendor || "").toUpperCase()).filter(Boolean));
  next.add(normalized);
  state.orderDismissedVendors = [...next];
  saveDismissedOrderVendors();
}

function undismissOrderVendor(vendorName) {
  const normalized = String(vendorName || "").trim().toUpperCase();
  if (!normalized) return;
  state.orderDismissedVendors = (state.orderDismissedVendors || []).filter((vendor) => String(vendor || "").toUpperCase() !== normalized);
  saveDismissedOrderVendors();
}

function cancelVendorPo(vendorName) {
  const normalized = String(vendorName || "").trim().toUpperCase();
  if (!normalized) return false;
  const staged = (state.orderSubmissionVendors || []).some((vendor) => vendor.toUpperCase() === normalized);
  const stagedDraft = !!state.orderSubmissionDrafts?.[normalized]?.length;
  const pending = (state.pendingOrders || []).some((po) => !po.cleared && (po.vendor || "").toUpperCase() === normalized);
  if (staged || stagedDraft) removeSubmittedVendor(vendorName);
  if (pending) clearVendorPending(vendorName);
  dismissOrderVendor(vendorName);
  return true;
}

function productReviewRows() {
  return buildInventoryRows()
    .map(applyOrderOverride)
    .filter((item) => item.isOrderable !== false && (item.recommendedOrder > 0 || item.qtyNeeded))
    .filter((item) => !isPendingOrder(item.code))
    .filter((item) => {
      if (!state.productPoReviewVendor) return true;
      return (item.vendor || "").toUpperCase() === state.productPoReviewVendor.toUpperCase();
    })
    .sort((a, b) => (b.recommendedOrder || 0) - (a.recommendedOrder || 0));
}

function closeProductPoReviewModal() {
  if (els.productPoReviewModal) els.productPoReviewModal.hidden = true;
}

function renderProductPoReviewModal() {
  if (!els.productPoReviewBody || !els.productPoReviewMeta || !els.productPoReviewPills) return;
  const allRows = buildInventoryRows()
    .map(applyOrderOverride)
    .filter((item) => item.isOrderable !== false && (item.recommendedOrder > 0 || item.qtyNeeded))
    .filter((item) => !isPendingOrder(item.code))
    .sort((a, b) => (b.recommendedOrder || 0) - (a.recommendedOrder || 0));
  const vendors = [...new Set(allRows.map((item) => cleanCell(item.vendor)).filter(Boolean))].sort(compareDisplayValue);
  if (state.productPoReviewVendor && !vendors.some((vendor) => vendor.toUpperCase() === state.productPoReviewVendor.toUpperCase())) {
    state.productPoReviewVendor = "";
  }
  const rows = productReviewRows().sort((a, b) => {
    const mode = state.productPoReviewSort || "rec-desc";
    if (mode === "item-asc") return compareDisplayValue(a.product, b.product);
    if (mode === "item-desc") return compareDisplayValue(b.product, a.product);
    if (mode === "code-asc") return compareDisplayValue(a.code, b.code);
    if (mode === "code-desc") return compareDisplayValue(b.code, a.code);
    if (mode === "rec-asc") return (a.recommendedOrder || 0) - (b.recommendedOrder || 0);
    return (b.recommendedOrder || 0) - (a.recommendedOrder || 0);
  });
  const totalQty = rows.reduce((sum, item) => sum + Math.max(0, toNumber(item.recommendedOrder) || 0), 0);
  const totalCost = rows.reduce((sum, item) => sum + orderLineCost(item), 0);
  const reviewSortSelect = document.querySelector("#productPoReviewSortSelect");
  if (reviewSortSelect) reviewSortSelect.value = state.productPoReviewSort || "rec-desc";
  els.productPoReviewTitle.textContent = "Create PO / Order";
  els.productPoReviewMeta.innerHTML = `
    <span><b>${number.format(rows.length)}</b> items</span>
    <span><b>${number.format(vendors.length)}</b> vendors</span>
    <span><b>${number.format(totalQty)}</b> total qty</span>
    <span><b>${currency.format(totalCost)}</b> total cost</span>
    <span><b>${escapeHtml(els.searchInput?.value || "All search results")}</b></span>`;
  els.productPoReviewPills.innerHTML = `
    <button type="button" class="order-vendor-pill${state.productPoReviewVendor ? "" : " active"}" data-product-review-vendor="">All</button>
    ${vendors.map((vendor) => `<button type="button" class="order-vendor-pill${state.productPoReviewVendor.toUpperCase() === vendor.toUpperCase() ? " active" : ""}" data-product-review-vendor="${escapeHtml(vendor)}">${escapeHtml(vendor)}</button>`).join("")}
  `;
  els.productPoReviewPills.querySelectorAll("[data-product-review-vendor]").forEach((btn) => {
    btn.addEventListener("click", () => {
      state.productPoReviewVendor = btn.dataset.productReviewVendor || "";
      renderProductPoReviewModal();
    });
  });
  if (!rows.length) {
    els.productPoReviewBody.innerHTML = `<tr><td colspan="10" class="empty-cell">No orderable items match the current Products filters.</td></tr>`;
    return;
  }
  els.productPoReviewBody.innerHTML = rows.map((item) => `<tr>
    <td>${escapeHtml(item.vendor || "-")}</td>
    <td>${escapeHtml(item.code)}</td>
    <td class="sku-name" title="${escapeHtml(item.product)}"><span class="sku-name-text">${escapeHtml(item.product)}</span></td>
    <td>${escapeHtml(item.plu || "-")}</td>
    <td class="num">${number.format(item.stock || 0)}</td>
    <td class="num">${number.format(item.reorderMin || 0)}</td>
    <td class="num">${number.format(item.reorderMax || 0)}</td>
    <td class="num"><input type="number" class="order-rec-input mini-input" data-code="${escapeHtml(item.code)}" value="${Math.max(0, toNumber(item.recommendedOrder) || 0)}" min="0" style="width:4rem;text-align:center;font-weight:700" /></td>
    <td class="num"><b>${number.format(calcCaseOrder(item.recommendedOrder || 0, item.caseSize || 1))}</b></td>
    <td class="num">${currency.format(orderLineCost(item))}</td>
  </tr>`).join("");
  els.productPoReviewBody.querySelectorAll(".order-rec-input").forEach((input) => {
    input.addEventListener("input", () => {
      const val = Math.max(0, Math.round(toNumber(input.value) || 0));
      input.value = String(val);
      const row = input.closest("tr");
      const code = codeKey(input.dataset.code);
      if (!state._orderRecOverrides) state._orderRecOverrides = new Map();
      state._orderRecOverrides.set(code, val);
      const item = rows.find((entry) => codeKey(entry.code) === code);
      if (!row || !item) return;
      const caseCount = calcCaseOrder(val, item.caseSize || 1);
      row.cells[8].innerHTML = `<b>${number.format(caseCount)}</b>`;
      row.cells[9].textContent = currency.format((toNumber(item.unitCost) || 0) * caseCount * Math.max(1, toNumber(item.caseSize) || 1));
    });
  });
  repairMojibakeText(els.productPoReviewModal || document.body);
}

function openProductPoReviewModal() {
  state.productPoReviewVendor = "";
  renderProductPoReviewModal();
  if (els.productPoReviewModal) els.productPoReviewModal.hidden = false;
}

function sendProductPoSelection(sendAll = false) {
  const baseRows = buildInventoryRows()
    .map(applyOrderOverride)
    .filter((item) => item.isOrderable !== false && (item.recommendedOrder > 0 || item.qtyNeeded));
  const allVendors = [...new Set(baseRows.map((item) => cleanCell(item.vendor)).filter(Boolean))];
  if (!sendAll && !state.productPoReviewVendor && allVendors.length > 1) {
    showToast("Choose a vendor pill first, or use Send All.", 2600, "warning");
    return;
  }
  const vendors = sendAll
    ? allVendors
    : (state.productPoReviewVendor ? [state.productPoReviewVendor] : allVendors.slice(0, 1));
  if (!vendors.length) {
    showToast("No vendors available for this PO review.", 2600, "warning");
    return;
  }
  const existing = [...(state.orderSubmissionVendors || [])];
  vendors.forEach((vendor) => {
    if (!existing.some((entry) => entry.toUpperCase() === vendor.toUpperCase())) existing.push(vendor);
    undismissOrderVendor(vendor);
  });
  state.orderSubmissionVendors = existing;
  state.orderSubmissionDrafts = buildOrderSubmissionDrafts(vendors, baseRows);
  state.orderSubmissionActiveVendor = sendAll ? "" : vendors[0];
  closeProductPoReviewModal();
  switchTab("ordering");
  renderOrders();
}

function exportProductReviewCsv() {
  downloadCsv("products-po-review.csv", productReviewRows().map((item) => ({
    vendor: item.vendor,
    code: item.code,
    product: item.product,
    stock: item.stock,
    reorderMin: item.reorderMin,
    reorderMax: item.reorderMax,
    recommendedOrder: item.recommendedOrder,
    caseOrder: calcCaseOrder(item.recommendedOrder || 0, item.caseSize || 1),
    totalCost: orderLineCost(item),
  })));
}

function exportProductReviewPdf() {
  const body = els.productPoReviewBody;
  const rows = [...(body?.querySelectorAll("tr") || [])];
  if (!rows.length) return;
  const htmlRows = rows.map((tr) => `<tr>${[...tr.cells].map((cell, index) => `<td class="${index >= 3 ? "num" : ""}">${cell.textContent.trim()}</td>`).join("")}</tr>`).join("");
  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Products PO Review</title>
  <style>
    body { font-family: Arial, sans-serif; font-size: 11px; color: #1c2320; padding: 18px; }
    h1 { font-size: 20px; margin: 0 0 10px; }
    table { width: 100%; border-collapse: collapse; }
    th, td { padding: 5px 6px; border-bottom: 1px solid #ddd; }
    th { background: #eef3ef; text-transform: uppercase; font-size: 9px; }
    .num { text-align: right; }
  </style></head><body>
  <h1>Products PO Review</h1>
  <table><thead><tr><th>Vendor</th><th>Code</th><th>PLU</th><th>Item</th><th>Stock</th><th>Min</th><th>Max</th><th>Rec. Order</th><th>Cases</th><th>Total</th></tr></thead>
  <tbody>${htmlRows}</tbody></table></body></html>`;
  const win = window.open("", "_blank", "width=960,height=720");
  if (!win) { showToast("Pop-up blocked.", 3000, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 350);
}

function countSessionUpdatedMs(session = {}) {
  return Date.parse(session.updatedAt || session.submittedAt || session.savedAt || session.startedAt || "") || 0;
}

function makeCountIdentifier(prefix = "count") {
  const id = globalThis.crypto?.randomUUID?.()
    || `${Date.now()}-${Math.random().toString(36).slice(2, 12)}`;
  return `${prefix}-${id}`;
}

function countDeviceId() {
  let id = localStorage.getItem(COUNT_DEVICE_ID_KEY) || "";
  if (!id) {
    id = makeCountIdentifier("device");
    safeLocalStorageSet(COUNT_DEVICE_ID_KEY, id);
  }
  return id;
}

function countDeviceLabel() {
  const ua = String(navigator?.userAgent || "").toLowerCase();
  const device = ua.includes("silk") || ua.includes("kindle")
    ? "Kindle"
    : /iphone|ipad|android|mobile/.test(ua) ? "Mobile" : "Desktop";
  return `${currentAuditUser()} (${device})`;
}

function countEntryIdentity(entry = {}, sessionId = "", index = 0) {
  if (cleanCell(entry.entryId || entry.entry_id)) return cleanCell(entry.entryId || entry.entry_id);
  return `legacy-${cleanCell(sessionId) || "session"}-${cleanCell(entry.recordedAt || entry.timestamp) || index}-${codeKey(entry.code)}-${cleanCell(entry.mode || "reset")}-${Number(entry.inputQty ?? entry.qty ?? 0)}`;
}

function normalizeCountEntry(entry = {}, session = {}, statusFallback = "pending", index = 0) {
  const deviceId = cleanCell(entry.deviceId || entry.device_id || session.deviceId || session.device_id);
  return {
    ...entry,
    entryId: countEntryIdentity(entry, session.id, index),
    sessionId: cleanCell(entry.sessionId || entry.session_id || session.id),
    deviceId,
    deviceLabel: cleanCell(entry.deviceLabel || entry.device_label || session.deviceLabel || session.device_label) || "Legacy device",
    user: cleanCell(entry.user || session.user) || "System",
    code: cleanCell(entry.code),
    plu: cleanCell(entry.plu),
    itemNumber: cleanCell(entry.itemNumber || entry.item_number),
    product: cleanCell(entry.product),
    qty: Number(entry.qty ?? entry.inputQty ?? 0),
    inputQty: Number(entry.inputQty ?? entry.qty ?? 0),
    mode: cleanCell(entry.mode || "reset").toLowerCase() === "add" ? "add" : "reset",
    recordedAt: entry.recordedAt || entry.timestamp || new Date().toISOString(),
    timestamp: entry.timestamp || entry.recordedAt || new Date().toISOString(),
    syncStatus: cleanCell(entry.syncStatus || entry.sync_status || statusFallback).toLowerCase() || statusFallback,
  };
}

function recomputeCountEntryTotals(entries = []) {
  const totals = new Map();
  return entries
    .slice()
    .sort((a, b) => String(a.recordedAt || "").localeCompare(String(b.recordedAt || "")) || String(a.entryId || "").localeCompare(String(b.entryId || "")))
    .map((entry) => {
      const key = codeKey(entry.code);
      const before = totals.has(key) ? totals.get(key) : Number(entry.originalQty || 0);
      const countedQty = entry.mode === "add"
        ? Math.max(0, before + Number(entry.inputQty || 0))
        : Math.max(0, Number(entry.inputQty || 0));
      totals.set(key, countedQty);
      return { ...entry, countedQty };
    });
}

function countSessionUndoIds(session = {}) {
  return new Set([
    ...(Array.isArray(session.undoEntryIds) ? session.undoEntryIds : []),
    ...(Array.isArray(session.deletedEntryIds) ? session.deletedEntryIds : []),
  ].map((id) => cleanCell(id)).filter(Boolean));
}

function filterUndoneCountEntries(entries = [], session = {}) {
  const undone = countSessionUndoIds(session);
  if (!undone.size) return entries;
  return entries.filter((entry) => !undone.has(cleanCell(entry.entryId)));
}


// SAFETY: localStorage quota should never break the count UI.
// If the browser storage is full, keep the app running and compact old count history.
function safeLocalStorageSet(key, value) {
  try {
    localStorage.setItem(key, value);
    return true;
  } catch (error) {
    console.warn('[CountStorage] localStorage set failed', key, error);
    if (String(key) === 'posDashboardCountSessions:v1') {
      try {
        const sessions = JSON.parse(value || '[]');
        const compact = (Array.isArray(sessions) ? sessions : [])
          .sort((a, b) => new Date(b.updatedAt || b.savedAt || b.createdAt || 0) - new Date(a.updatedAt || a.savedAt || a.createdAt || 0))
          .slice(0, 40)
          .map((session) => {
            if (session && !session.isActiveLive && (session.savedAt || session.submittedAt) && Array.isArray(session.entries) && session.entries.length > 250) {
              return { ...session, entries: session.entries.slice(-250), storageCompacted: true };
            }
            return session;
          });
        localStorage.setItem(key, JSON.stringify(compact));
        return true;
      } catch (secondError) {
        console.warn('[CountStorage] compact retry failed', secondError);
      }
    }
    return false;
  }
}
function normalizeCountSession(session = {}, statusFallback = "pending") {
  if (!session?.id) return session;
  const undoEntryIds = [...countSessionUndoIds(session)];
  const entries = filterUndoneCountEntries(session.entries || [], { undoEntryIds })
    .map((entry, index) => normalizeCountEntry(entry, session, statusFallback, index));
  return {
    ...session,
    undoEntryIds,
    deviceId: cleanCell(session.deviceId || session.device_id) || countDeviceId(),
    deviceLabel: cleanCell(session.deviceLabel || session.device_label) || countDeviceLabel(),
    entries: recomputeCountEntryTotals(entries),
  };
}

function mergeCountSessions(remoteSession, localSession) {
  if (!remoteSession) return normalizeCountSession(localSession, "pending");
  if (!localSession) return normalizeCountSession(remoteSession, "synced");
  const remote = normalizeCountSession(remoteSession, "synced");
  const local = normalizeCountSession(localSession, "pending");
  const newest = countSessionUpdatedMs(local) > countSessionUpdatedMs(remote) ? local : remote;
  const undoEntryIds = [...new Set([
    ...countSessionUndoIds(remote),
    ...countSessionUndoIds(local),
  ])];
  const mergedEntries = new Map(filterUndoneCountEntries(remote.entries, { undoEntryIds }).map((entry) => [entry.entryId, { ...entry, syncStatus: "synced" }]));
  filterUndoneCountEntries(local.entries, { undoEntryIds }).forEach((entry) => {
    const remoteEntry = mergedEntries.get(entry.entryId);
    if (!remoteEntry || ["pending", "failed"].includes(entry.syncStatus)) mergedEntries.set(entry.entryId, entry);
  });
  const hasPending = [...mergedEntries.values()].some((entry) => entry.syncStatus !== "synced");
  return {
    ...newest,
    undoEntryIds,
    entries: recomputeCountEntryTotals([...mergedEntries.values()]),
    syncVersion: Math.max(Number(remote.syncVersion || 0), Number(local.syncVersion || 0)),
    localSyncPending: hasPending || local.localSyncPending === true,
    lastSyncAt: newest.lastSyncAt || remote.lastSyncAt || "",
  };
}

function countEntriesAwaitingSync(session = {}) {
  return (session.entries || []).filter((entry) => cleanCell(entry.syncStatus).toLowerCase() !== "synced");
}

function countSessionSyncStats(session = {}) {
  const entries = session.entries || [];
  const failed = entries.filter((entry) => entry.syncStatus === "failed").length;
  const pending = entries.filter((entry) => entry.syncStatus !== "synced" && entry.syncStatus !== "failed").length;
  const deviceCount = new Set(entries.map((entry) => entry.deviceId || entry.deviceLabel).filter(Boolean)).size;
  return { pending, failed, synced: Math.max(0, entries.length - pending - failed), deviceCount };
}

function renderActiveCountSyncStatus(session = state.activeCountSession) {
  if (!els.activeCountSyncStatus) return;
  if (!session) {
    els.activeCountSyncStatus.innerHTML = "";
    return;
  }
  const stats = countSessionSyncStats(session);
  const status = stats.failed ? "Failed" : stats.pending || session.localSyncPending ? "Pending" : "Synced";
  const cls = status === "Failed" ? "sync-failed" : status === "Pending" ? "sync-pending" : "";
  const lastSync = session.lastSyncAt || state._countLastSyncAt;
  els.activeCountSyncStatus.innerHTML = `
    <span class="${cls}"><b>${escapeHtml(status)}</b>${stats.pending ? ` - ${stats.pending} pending` : ""}${stats.failed ? ` - ${stats.failed} failed` : ""}</span>
    <span>Last sync: ${escapeHtml(lastSync ? new Date(lastSync).toLocaleString() : "Not synced yet")}</span>
    <span>Devices: ${number.format(stats.deviceCount || (session.deviceId ? 1 : 0))}</span>`;
}

function markCountSessionDirty(session) {
  if (!session) return null;
  const stamp = new Date().toISOString();
  session.updatedAt = stamp;
  session.syncVersion = Math.max(0, Number(session.syncVersion || 0)) + 1;
  session.localSyncPending = true;
  return session;
}

function countSessionHasUnsyncedEdits(session) {
  return session?.localSyncPending === true;
}

function countSessionForRemote(session = {}) {
  const { localSyncPending, syncError, ...payload } = normalizeCountSession(session, "pending");
  payload.undoEntryIds = [...countSessionUndoIds(payload)];
  payload.entries = filterUndoneCountEntries(payload.entries || [], payload).map((entry) => ({ ...entry, syncStatus: "synced" }));
  return payload;
}

function preferredCountSession(existing, candidate) {
  if (!existing) return candidate;
  const existingSaved = !!(existing.submittedAt || existing.savedAt);
  const candidateSaved = !!(candidate.submittedAt || candidate.savedAt);
  if (existingSaved !== candidateSaved) return candidateSaved ? candidate : existing;
  return countSessionUpdatedMs(candidate) >= countSessionUpdatedMs(existing) ? candidate : existing;
}

function canonicalCountSessions(sessions = []) {
  const byId = new Map();
  sessions.filter(Boolean).forEach((session) => {
    const id = cleanCell(session.id);
    if (!id || countSessionIsDeleted(id)) return;
    byId.set(id, preferredCountSession(byId.get(id), session));
  });
  return [...byId.values()].sort((a, b) => countSessionUpdatedMs(b) - countSessionUpdatedMs(a));
}

function persistCountSessions(options = {}) {
  state.countSessions = canonicalCountSessions(state.countSessions);
  safeLocalStorageSet("posDashboardCountSessions:v1", JSON.stringify(state.countSessions));
  safeLocalStorageSet("posDashboardActiveCountSession:v1", JSON.stringify(state.activeCountSession));
  if (options.scheduleSync !== false) scheduleSharedCountSessionsSync();
}

function handleScanModeLookup(value, options = {}) {
  const input = document.querySelector("#scanModeInput");
  const query = cleanCell(value ?? input?.value ?? "");
  if (!query) {
    input?.focus?.();
    input?.select?.();
    return null;
  }
  const start = performance.now();
  const exact = priceCheckExactMatch(query);
  const item = exact || priceCheckMatches(query, 1)[0] || null;
  const elapsed = Number((performance.now() - start).toFixed(2));
  console.debug("[ScanPerf] lookup", { ms: elapsed, exact: !!exact, found: !!item });
  renderPriceCheckResult(item);
  if (els.scanModeStatus) {
    els.scanModeStatus.textContent = item
      ? `\u2713 ${item.product || item.code} \u2014 ready for next scan.`
      : `No match for "${query}". Try again.`;
    delete els.scanModeStatus.dataset.readiness;
  }
  if (!item && !options.silentNotFound) showToast("Item not found.", 2000, "warning");
  if (options.refocus !== false) {
    requestAnimationFrame(() => {
      input?.focus?.();
      input?.select?.();
    });
  }
  return item;
}

function persistDeletedCountSessionIds() {
  safeLocalStorageSet("posDashboardDeletedCountSessionIds:v1", JSON.stringify([...state._deletedCountSessionIds].filter(Boolean).slice(-500)));
}

function countSessionIsDeleted(sessionId) {
  const id = cleanCell(sessionId);
  return !!id && !!state._deletedCountSessionIds?.has?.(id);
}

function markCountSessionDeleted(sessionId) {
  const id = cleanCell(sessionId);
  if (!id) return "";
  state._deletedCountSessionIds.add(id);
  persistDeletedCountSessionIds();
  return id;
}

function pruneDeletedCountSessions(options = {}) {
  const { persist = false } = options;
  state.countSessions = (state.countSessions || []).filter((session) => !countSessionIsDeleted(session?.id));
  if (state.activeCountSession?.id && countSessionIsDeleted(state.activeCountSession.id)) {
    state.activeCountSession = null;
    state.selectedCountItemCode = "";
    state.countQtyBuffer = "0";
    state.countStage = "search";
    state.pendingDuplicateCount = null;
    state.pendingDuplicateMode = null;
  }
  if (persist) {
    safeLocalStorageSet("posDashboardCountSessions:v1", JSON.stringify(state.countSessions));
    safeLocalStorageSet("posDashboardActiveCountSession:v1", JSON.stringify(state.activeCountSession));
  }
}

function archiveCountSessionEntriesToAdjustmentLog(session) {
  if (!session?.id || !(session.entries || []).length) return 0;
  let added = 0;
  const previousByCode = { ...(session.preCountSnapshot || {}) };
  (session.entries || []).forEach((entry) => {
    const key = codeKey(entry.code);
    if (!key) return;
    const duplicate = (state.adjustmentLog || []).some((log) =>
      cleanCell(log.countSessionId) === cleanCell(session.id) &&
      codeKey(log.code) === key &&
      cleanCell(log.recordedAt) === cleanCell(entry.recordedAt)
    );
    if (duplicate) return;
    const before = previousByCode[key] ?? entry.originalQty ?? 0;
    const after = Number(entry.countedQty || 0);
    previousByCode[key] = after;
    state.adjustmentLog.unshift({
      recordedAt: entry.recordedAt || session.updatedAt || new Date().toISOString(),
      user: entry.user || session.user || currentAuditUser(),
      code: entry.code,
      product: entry.product || "",
      vendor: entry.vendor || "",
      category: entry.category || "",
      action: "COUNT ENTRY",
      qtyChange: after - Number(before || 0),
      qtyBefore: Number(before || 0),
      qtyAfter: after,
      reason: `Archived physical count: ${countSessionLabel(session)} (${entry.mode || "reset"})`,
      countSessionId: session.id,
    });
    added++;
  });
  if (added) safeLocalStorageSet("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  return added;
}

function persistActiveCountSession() {
  markCountSessionDirty(state.activeCountSession);
  if (state.activeCountSession?.id) {
    const activeId = cleanCell(state.activeCountSession.id);
    const liveCopy = { ...state.activeCountSession, isActiveLive: true };
    // Keep the visible/session-history copy in lockstep with the active live session.
    // Without this, shared/local sync can re-merge an older countSessions[] copy and
    // make undone +1 rows appear to come back.
    state.countSessions = [
      liveCopy,
      ...(state.countSessions || []).filter((session) => cleanCell(session?.id) !== activeId),
    ];
  }
  // The local journal is immediate; only the remote transfer is debounced.
  safeLocalStorageSet("posDashboardActiveCountSession:v1", JSON.stringify(state.activeCountSession));
  safeLocalStorageSet("posDashboardCountSessions:v1", JSON.stringify(state.countSessions));
  renderActiveCountSyncStatus();
  scheduleSharedCountSessionsSync();
}

function allCountCandidateRows() {
  if (state._countCandidateCache && state._countCandidateStamp === state._dataCacheStamp) {
    return state._countCandidateCache;
  }
  // Physical count scanner must be able to find the full item master, including
  // Disabled/Discontinued items. Those items are filtered out of comparison/null-zero
  // scope later unless they were actually scanned.
  state._countCandidateCache = buildInventoryRows({ ignoreQuery: true, ignoreFilters: true, ignoreStateFilter: true });
  state._countCandidateStamp = state._dataCacheStamp;
  return state._countCandidateCache;
}

function filteredCountCandidateRows(session = state.activeCountSession) {
  if (!session) return allCountCandidateRows();
  const cacheKey = [session.id, session.vendor, session.department, session.category, JSON.stringify(session.vendorFilters || []), JSON.stringify(session.categoryFilters || []), session.status, session.searchFilter, state._dataCacheStamp].join("|");
  if (state._filteredCountCache && state._filteredCountCacheKey === cacheKey) {
    return state._filteredCountCache;
  }
  const statusFilter = (session.status || "").trim().toLowerCase();
  const sourceRows = allCountCandidateRows();
  const rows = sourceRows.filter((item) => {
    const vendorFilters = countSessionFilterList(session, "vendor", "vendorFilters").map((value) => value.toUpperCase());
    if (vendorFilters.length && !vendorFilters.includes((item.vendor || "").trim().toUpperCase())) return false;
    const catFilters = countSessionFilterList(session, "category", "categoryFilters").map((value) => value.toUpperCase());
    if (catFilters.length && !catFilters.includes((item.category || "").trim().toUpperCase())) return false;
    if (statusFilter && (item.state || "").trim().toLowerCase() !== statusFilter) return false;
    const scopeTerms = countScopeTerms(session.searchFilter || "");
    if (scopeTerms.length) {
      const scopeHaystack = [item.product, item.plu, item.itemNumber, item.code]
        .map((value) => String(value || "").toLowerCase()).join("|");
      const scopeCodeHaystack = [item.plu, item.itemNumber, item.code].map((value) => codeKey(value)).join("|");
      const matchesAnyScope = scopeTerms.some((scopeNeedle) => {
        const scopeCodeNeedle = codeKey(scopeNeedle);
        return scopeHaystack.includes(scopeNeedle) || (scopeCodeNeedle && scopeCodeHaystack.includes(scopeCodeNeedle));
      });
      if (!matchesAnyScope) return false;
    }
    return true;
  });
  state._filteredCountCache = rows;
  state._filteredCountCacheKey = cacheKey;
  return rows;
}

function countSessionLabel(session) {
  if (!session) return "Physical count";
  const vendorList = countSessionFilterList(session, "vendor", "vendorFilters");
  const categoryList = countSessionFilterList(session, "category", "categoryFilters");
  const vendorLabel = vendorList.length ? vendorList.join(", ") : (session.department || "All vendors");
  const categoryLabel = categoryList.length ? categoryList.join(", ") : "All categories";
  const parts = [session.date || "No date", vendorLabel, categoryLabel];
  if (session.searchFilter) parts.push(`Filter: ${session.searchFilter}`);
  return parts.join("  -  ");
}

// ARCH-B: Human-readable status label for count session history rows.
// Shows: Active | Completed | Paused | Synced | Local Only
function countSessionStatusLabel(session) {
  if (!session) return "-";
  // FIX-B/C: isActiveLive is set when a session is in countSessions[] AND currently active.
  // Also fall back to comparing with activeCountSession.id in case flag wasn't set.
  const isActive = session.isActiveLive || state.activeCountSession?.id === session.id;
  if (isActive) return '<span class="state-badge state-active">Active / Live</span>';
  if (session.submittedAt) return '<span class="state-badge state-synced">Completed</span>';
  const synced = !session.localSyncPending && !countSessionHasUnsyncedEdits(session);
  if (synced && session.lastSyncAt) return '<span class="state-badge state-synced">Synced</span>';
  if (session.savedAt) return '<span class="state-badge">Paused</span>';
  if (countSessionHasUnsyncedEdits(session)) return '<span class="state-badge state-badge--warn">Local only</span>';
  return '<span class="state-badge">In progress</span>';
}

function countSessionCanDelete(session) {
  if (!session) return false;
  if (isAdmin()) return true;
  const userOwnsSession = cleanCell(session.user) && cleanCell(session.user) === currentAuditUser();
  const deviceOwnsSession = cleanCell(session.deviceId) && cleanCell(session.deviceId) === countDeviceId();
  return !!(userOwnsSession || deviceOwnsSession);
}


function updateCountSoundToggle() {
  const button = document.querySelector("#countSoundToggleButton");
  if (!button) return;
  button.setAttribute("aria-pressed", state.countSoundEnabled ? "true" : "false");
  button.classList.toggle("is-on", !!state.countSoundEnabled);
  button.textContent = state.countSoundEnabled ? "Sound On" : "Sound Off";
}

function toggleCountSound() {
  state.countSoundEnabled = !state.countSoundEnabled;
  safeLocalStorageSet("posDashboardCountSoundEnabled:v1", state.countSoundEnabled ? "true" : "false");
  updateCountSoundToggle();
  if (state.countSoundEnabled) playCountConfirmBeep({ soft: true });
}

function playCountConfirmBeep(options = {}) {
  try {
    // Audio routing fix: force Add Qty / +1 to use the uploaded kaching asset,
    // not the old generated beep or cached ding path.
    const audio = state._countConfirmAudio || new Audio("add-qty-kaching-376867.mp3?v=376867");
    state._countConfirmAudio = audio;
    audio.preload = "auto";
    audio.volume = 1;
    audio.currentTime = 0;
    const playPromise = audio.play();
    if (playPromise?.catch) playPromise.catch(() => {});
  } catch (_) {}
}

function openCountSetupModal() {
  els.countDateInput.value = new Date().toISOString().slice(0, 10);
  resetCountSetupExtraFilters();
  // Instant Start UI: open the setup modal first, then populate heavy dropdowns
  // after paint so slower tablets/Kindles feel responsive immediately.
  els.countSetupModal.hidden = false;
  els.countDateInput.focus();
  requestAnimationFrame(() => {
    populateCountSetupOptions();
    // Always reset to "All" for a fresh count after options exist.
    els.countVendorInput.value = "";
    els.countCategoryInput.value = "";
    const statusEl = document.querySelector("#countStatusInput");
    if (statusEl) statusEl.value = "";
    if (els.countScopeSearchInput) els.countScopeSearchInput.value = "";
    if (els.countAllowOutOfScopeInput) els.countAllowOutOfScopeInput.checked = false;
  });
}

function closeCountSetupModal() {
  els.countSetupModal.hidden = true;
}

function countFilterValues(kind) {
  const selector = kind === "category" ? ".count-category-select" : ".count-vendor-select";
  const seen = new Set();
  return [...document.querySelectorAll(selector)]
    .map((select) => cleanCell(select.value || ""))
    .filter((value) => {
      if (!value) return false;
      const key = value.toUpperCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function countSessionFilterList(session, field, pluralField) {
  const explicit = Array.isArray(session?.[pluralField]) ? session[pluralField] : [];
  const values = explicit.length ? explicit : String(session?.[field] || "").split(/[,|]/);
  const seen = new Set();
  return values
    .map((value) => cleanCell(value || ""))
    .filter((value) => {
      if (!value) return false;
      const key = value.toUpperCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
}

function countScopeTerms(searchFilter) {
  const seen = new Set();
  return String(searchFilter || "")
    .split(",")
    .map((value) => cleanCell(value).toLowerCase())
    .filter((value) => {
      if (!value) return false;
      if (seen.has(value)) return false;
      seen.add(value);
      return true;
    });
}

function fillCountSetupSelects(kind) {
  const values = kind === "category" ? (state._countCategoryOptions || []) : (state._countVendorOptions || []);
  const selector = kind === "category" ? ".count-category-select" : ".count-vendor-select";
  document.querySelectorAll(selector).forEach((select) => {
    const current = select.value || "";
    fillSelect(select, values);
    if (current && [...select.options].some((option) => option.value === current)) select.value = current;
  });
}

function addCountSetupFilter(kind) {
  const group = document.querySelector(kind === "category" ? "#countCategoryFilterGroup" : "#countVendorFilterGroup");
  if (!group) return;
  const select = document.createElement("select");
  select.className = kind === "category" ? "count-category-select count-extra-filter" : "count-vendor-select count-extra-filter";
  select.setAttribute("aria-label", kind === "category" ? "Additional category" : "Additional vendor");
  group.appendChild(select);
  fillCountSetupSelects(kind);
  select.focus();
}

function resetCountSetupExtraFilters() {
  document.querySelectorAll(".count-extra-filter").forEach((node) => node.remove());
}

function populateCountSetupOptions() {
  // FIX-D: Cache vendor/category/status lists keyed to the data stamp.
  // Previously this rebuilt from the full inventory Map on every renderCountsWorkspace call.
  // On Kindle (10k+ items) that was 100-300ms of synchronous work on each poll tick.
  if (state._countSetupCacheStamp === state._dataCacheStamp && state._countSetupCached) {
    fillCountSetupSelects("vendor");
    fillCountSetupSelects("category");
    return; // lists are still current
  }
  const allRows = [...state.latestInventory.values()];
  const excelRows = [...state.excelItems.values()];
  const combined = allRows.length ? allRows : excelRows;
  state._countVendorOptions = unique(combined.map((r) => r.vendor).filter(Boolean));
  state._countCategoryOptions = unique(combined.map((r) => r.category).filter(Boolean));
  fillCountSetupSelects("vendor");
  fillCountSetupSelects("category");
  const statusEl = document.querySelector("#countStatusInput");
  if (statusEl) {
    const knownStates = ["Active", "Disabled", "Discontinued", "Force Order"];
    const dataStates = unique(combined.map((r) => r.state).filter(Boolean));
    fillSelect(statusEl, unique([...knownStates, ...dataStates]));
  }
  state._countSetupCacheStamp = state._dataCacheStamp;
  state._countSetupCached = true;
}



function startCountSessionFromModal() {
  const statusEl = document.querySelector("#countStatusInput");
  const deviceId = countDeviceId();
  const vendorFilters = countFilterValues("vendor");
  const categoryFilters = countFilterValues("category");
  const session = {
    id: makeCountIdentifier("count"),
    date: els.countDateInput.value || new Date().toISOString().slice(0, 10),
    vendor: vendorFilters.join(", "),
    vendorFilters,
    category: categoryFilters.join(", "),
    categoryFilters,
    status: statusEl ? (statusEl.value || "") : "",
    searchFilter: cleanCell(els.countScopeSearchInput?.value || ""),
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    deviceId,
    deviceLabel: countDeviceLabel(),
    user: currentAuditUser(),
    allowOutOfScope: !isUserRole() && !!els.countAllowOutOfScopeInput?.checked,
    syncVersion: 0,
    localSyncPending: true,
    entries: [],
  };
  // Instant Start UI: create/open the count screen first. The read-only starting
  // quantity snapshot is filled in right after paint so slow devices are not stuck
  // before the workspace appears.
  session.countStartSnapshot = {};

  // FIX-B: Register in countSessions[] immediately with isActiveLive:true so Report History
  // shows it right away. activeCountSession points to the same object.
  // isActiveLive is read by countSessionStatusLabel() to show "Active / Live" in the table.
  const liveSession = { ...session, isActiveLive: true, _instantStartLoading: true };
  // V2.0: if another count was active on THIS device, set it aside (Paused) instead of
  // leaving a second "Active / Live" row in history. Single active session per device.
  const prevId = cleanCell(state.activeCountSession?.id);
  if (prevId && prevId !== cleanCell(session.id)) {
    state.countSessions = state.countSessions.map((s) =>
      cleanCell(s.id) === prevId
        ? { ...s, isActiveLive: false, savedAt: s.savedAt || new Date().toISOString() }
        : s,
    );
  }
  clearContinuedCountLock();
  state.activeCountSession = liveSession;
  state._continuingCountId = cleanCell(session.id); // marks the currently-open count so background sync won't drop it
  state._countSessionOpen = true;
  state.countQtyBuffer = "0";
  state.selectedCountItemCode = "";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  state.pendingDuplicateMode = null;
  state.countAutoPlusMode = false;
  state.countAutoUndoStack = [];
  state.countSessions = [liveSession, ...state.countSessions.filter((s) => s.id !== session.id)];
  persistActiveCountSession();
  persistCountSessions({ scheduleSync: false });
  closeCountSetupModal();
  requestAnimationFrame(() => {
    setTimeout(() => {
      const target = state.activeCountSession?.id === liveSession.id ? state.activeCountSession : liveSession;
      if (!target || Object.keys(target.countStartSnapshot || {}).length) return;
      const snapshot = {};
      currentCountSessionCandidates(target).forEach((item) => {
        const key = codeKey(item.code);
        if (key) snapshot[key] = Number(item.stock || 0);
      });
      target.countStartSnapshot = snapshot;
      target._instantStartLoading = false;
      const stored = state.countSessions.find((s) => s.id === target.id);
      if (stored) {
        stored.countStartSnapshot = snapshot;
        stored._instantStartLoading = false;
      }
      persistActiveCountSession();
      persistCountSessions({ scheduleSync: false });
      renderCountReviewPanel();
    }, 40);
  });
  // FIX-D: Build the search index lazily — defer until after the UI paints.
  // On Kindle this was a synchronous ~400ms block before the count screen appeared.
  // findCountMatch()/findAnyCountMatch() also auto-build on first use (lines 4307/4323).
  setTimeout(() => buildCountSearchIndex(), 80);
  // V2.0: route the open through the one canonical show path so a new count reliably
  // appears (and any leftover setup/report overlay is torn down).
  showCountWorkspace();
  setTimeout(showCountWorkspace, 120);
  setTimeout(() => { showCountWorkspace(); focusCountSearch(); }, 400);
  showToast(
    state._countSyncUnavailable
      ? `Started local count: ${countSessionLabel(session)}. Sync will retry when available.`
      : `Started count: ${countSessionLabel(session)}`,
    3600,
    state._countSyncUnavailable ? "warning" : "success",
  );
}

function closeActiveCountSession() {
  if (!state.activeCountSession) return;
  // V2.0: closing just hides the workspace VIEW. The session remains active in
  // Report History so it can be continued; nothing is deleted.
  state._countSessionOpen = false;
  state.selectedCountItemCode = "";
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  state.pendingDuplicateMode = null;
  persistCountSessions();
  renderCountsWorkspace();
}

async function saveCountSession() {
  if (!state.activeCountSession) return;
  clearContinuedCountLock();
  const syncCheckedSession = await ensureCountSessionReadyToApply(state.activeCountSession, "saving");
  if (!syncCheckedSession) return;
  state._countSessionOpen = false;
  const session = markCountSessionDirty({
    ...syncCheckedSession,
    savedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    isActiveLive: false,   // FIX-B: clear live flag — session is now saved/paused
  });
  state.countSessions = [session, ...state.countSessions.filter((item) => item.id !== session.id)];

  // Apply physical counts: update stock values in latestInventory to reflect counted quantities
  const latestByCode = new Map();
  (session.entries || []).forEach((entry) => {
    latestByCode.set(codeKey(entry.code), entry);
  });
  let updatedCount = 0;
  latestByCode.forEach((entry, key) => {
    const existing = state.latestInventory.get(key);
    state._localProductMetaEdits.set(key, Date.now());
    if (existing) {
      const beforeStock = Number(existing.stock || 0);
      const afterStock = Number(entry.countedQty || 0);
      if (beforeStock !== afterStock) {
        state.adjustmentLog.unshift({
          recordedAt: entry.recordedAt || new Date().toISOString(),
          user: currentAuditUser(),
          code: existing.code || entry.code,
          product: existing.product || entry.product || "",
          vendor: existing.vendor || entry.vendor || "",
          category: existing.category || entry.category || "",
          action: "COUNT ENTRY",
          qtyChange: afterStock - beforeStock,
          qtyBefore: beforeStock,
          qtyAfter: afterStock,
          reason: `Physical count saved: ${countSessionLabel(session)} (${entry.mode || "reset"})`,
          countSessionId: session.id,
        });
      }
      existing.stock = afterStock;
      state.latestInventory.set(key, existing);
      updatedCount++;
    } else {
      // Item may be keyed differently - try to find it
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
  state.pendingDuplicateMode = null;
  // PATCH A: Do NOT call bumpDataStamp() or render() here.
  // Only count state changed — rebuilding all sales/inventory caches is unnecessary and causes
  // the visible freeze + cascading render loop on save.
  // Invalidate only the count-specific caches and affected inventory row index entries.
  state._countCandidateCache = null;
  state._filteredCountCache = null;
  state._countCandidateStamp = -1;
  [...latestByCode.keys()].forEach((key) => state._inventoryRowIndex.delete(key));
  safeLocalStorageSet("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  persistCountSessions();
  renderCountsWorkspace();
  // Sync in background. Delay product/inventory fact sync slightly so the UI is not blocked.
  void syncSharedCountSessionsToSupabase(true);
  setTimeout(() => {
    // Push the stock quantities as a product-meta update too, so Products/Scan Mode
    // on every device pulls the same inventory counts after a physical-count save.
    void syncSharedInventoryFactsByCodes([...latestByCode.keys()], { silent: true, updateSyncState: true });
  }, 400);
  showToast(`Count saved · ${updatedCount} item stock quantities updated`, 3200, "success");
}

function deleteCountSession() {
  if (!state.activeCountSession) return;
  clearContinuedCountLock();
  state._countSessionOpen = false;
  const label = countSessionLabel(state.activeCountSession);
  archiveCountSessionEntriesToAdjustmentLog(state.activeCountSession);
  const deletedId = markCountSessionDeleted(state.activeCountSession.id);
  state.activeCountSession = null;
  state.selectedCountItemCode = "";
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  state.pendingDuplicateMode = null;
  persistCountSessions();
  renderCountsWorkspace();
  void (async () => {
    if (deletedId && ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
      await supabaseDeleteRowsByValues("count_sessions", "id", [deletedId, `active:${deletedId}`]);
      await updateSharedSyncState("counts");
    }
    await syncSharedCountSessionsToSupabase(true);
  })().catch(() => {});
  showToast(`Deleted unsaved count: ${label}`, 3200, "warning");
}


function restoreContinuedCountSession() {
  const id = cleanCell(state._continuingCountId || state.countReportOpenId || "");
  if (!id) return false;
  const session = findCountSessionById(id) || state.countSessions.find((s) => cleanCell(s?.id) === id);
  if (!session) return false;
  state.activeCountSession = markCountSessionDirty({ ...session, isActiveLive: true });
  state.countSessions = [
    state.activeCountSession,
    ...(state.countSessions || []).filter((saved) => cleanCell(saved?.id) !== id),
  ];
  state._countSessionOpen = true;
  persistActiveCountSession();
  persistCountSessions({ scheduleSync: false });
  return true;
}

function stopContinueCountKeepOpenGuard() {
  if (state._continueCountKeepOpenTimer) {
    clearInterval(state._continueCountKeepOpenTimer);
    state._continueCountKeepOpenTimer = null;
  }
}

// V2.0 FIX: The old keep-open guard ran a 300ms interval that constantly forced the
// workspace open, re-asserted the active session, and stole focus — only stopping on
// Submit. That polling fought every deterministic action: it re-popped the count when
// switching to the Inventory tab, re-showed the old count over the New-count filter,
// churned the DOM so Save/Delete clicks didn't land, and clobbered header sorting.
// We now rely on a single source of truth: state._countSessionOpen, flipped ONLY by
// explicit user actions (openCount / start new / save / delete / close). This function
// just records which count is currently open (so background sync won't drop it) and
// makes sure no stray interval is left running.
function startContinueCountKeepOpenGuard(sessionId) {
  const id = cleanCell(sessionId || state._continuingCountId || "");
  if (!id) return;
  state._continuingCountId = id;
  stopContinueCountKeepOpenGuard();
}

function clearContinuedCountLock() {
  state._continuingCountId = "";
  stopContinueCountKeepOpenGuard();
}


function isCompactCountKeypadMode() {
  return !!(window.matchMedia && (window.matchMedia("(max-width: 1100px)").matches || window.matchMedia("(pointer: coarse)").matches));
}

function openCountKeypadPanel() {
  if (!els.countKeypadPanel) return;
  renderCountKeypadItemSummary();
  els.countKeypadPanel.classList.add("keypad-open");
  els.countKeypadPanel.removeAttribute("aria-hidden");
  // Tablet/Kindle qty entry should behave like a popup, not a bottom panel.
  // Do not scroll the workspace down after scan; CSS centers the keypad modal.
}

function closeCountKeypadPanel() {
  if (!els.countKeypadPanel) return;
  els.countKeypadPanel.classList.remove("keypad-open");
  if (isCompactCountKeypadMode()) els.countKeypadPanel.setAttribute("aria-hidden", "true");
  else els.countKeypadPanel.removeAttribute("aria-hidden");
}

function syncCountKeypadVisibility() {
  if (!els.countKeypadPanel) return;
  if (state.activeCountSession && state._countSessionOpen && state.countStage === "qty" && currentSelectedCountItem()) {
    openCountKeypadPanel();
  } else if (isCompactCountKeypadMode()) {
    closeCountKeypadPanel();
  } else {
    els.countKeypadPanel.classList.remove("keypad-open");
    els.countKeypadPanel.removeAttribute("aria-hidden");
  }
}

function enforceCountQtyLimit(nextValue) {
  const numeric = Number(nextValue || 0);
  if (!Number.isFinite(numeric)) return { ok: false, value: "0" };
  // No hard 99 cap: allow real bulk counts/cases. Scanner safety is handled
  // separately by the large-quantity confirmation before committing.
  return { ok: true, value: String(nextValue) };
}

function looksLikeBarcodeValue(value) {
  const digits = cleanCell(value).replace(/\D/g, "");
  return digits.length >= 8 && digits.length <= 18;
}

function redirectQtyBarcodeToSearch(barcode) {
  const code = cleanCell(barcode).replace(/\D/g, "");
  if (!looksLikeBarcodeValue(code)) return false;
  // Scanner safety: a long scanner burst should never become quantity.
  // Treat it as an accidental second scan/rescan instead.
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.selectedCountItemCode = "";
  state.pendingDuplicateMode = null;
  state._countQtyScannerBuffer = "";
  if (state._countQtyScannerTimer) {
    clearTimeout(state._countQtyScannerTimer);
    state._countQtyScannerTimer = null;
  }
  renderSelectedCountItem();
  renderCountQuantity();
  if (els.countSearchInput) {
    els.countSearchInput.value = code;
    els.countSearchInput.focus();
    els.countSearchInput.select?.();
  }
  hideCountDropdown();
  setTimeout(() => handleCountLookup(), 0);
  return true;
}

function noteQtyDigitForScannerGuard(digit) {
  const now = performanceNow();
  if (!state._countQtyScannerLastAt || now - state._countQtyScannerLastAt > 150) {
    state._countQtyScannerBuffer = "";
  }
  state._countQtyScannerLastAt = now;
  state._countQtyScannerBuffer = `${state._countQtyScannerBuffer || ""}${digit}`;
  const bufferedNow = state._countQtyScannerBuffer || "";
  // Immediate catch: once a rapid numeric burst reaches barcode length, do not
  // let more digits keep building in Entered Qty. A real hand-entered qty should
  // almost never be 8+ digits, and human typing won't hit this rapid-burst path.
  if (bufferedNow.length >= 8 && looksLikeBarcodeValue(bufferedNow)) {
    if (state._countQtyScannerTimer) {
      clearTimeout(state._countQtyScannerTimer);
      state._countQtyScannerTimer = null;
    }
    redirectQtyBarcodeToSearch(bufferedNow);
    showToast("Barcode scan caught in qty field — moved back to scan.", 2200, "warning");
    return true;
  }
  if (state._countQtyScannerTimer) clearTimeout(state._countQtyScannerTimer);
  state._countQtyScannerTimer = setTimeout(() => {
    const buffered = state._countQtyScannerBuffer || "";
    state._countQtyScannerBuffer = "";
    state._countQtyScannerTimer = null;
    if (state.countStage === "qty" && looksLikeBarcodeValue(buffered)) {
      redirectQtyBarcodeToSearch(buffered);
      showToast("Barcode scan caught in qty field — moved back to scan.", 2200, "warning");
    }
  }, 120);
  return false;
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
    const next = state.countQtyBuffer === "0" ? key : `${state.countQtyBuffer}${key}`;
    const limited = enforceCountQtyLimit(next);
    state.countQtyBuffer = limited.value;
  }
  renderCountQuantity();
}

function renderCountQuantity() {
  if (els.countQuantityDisplay) els.countQuantityDisplay.textContent = state.countQtyBuffer || "0";
  renderCountAutoControls();
}

function renderCountAutoControls() {
  const on = !!state.countAutoPlusMode;
  if (els.countAutoPlusButton) {
    els.countAutoPlusButton.textContent = on ? "+1 Auto On" : "+1 Auto Off";
    els.countAutoPlusButton.classList.toggle("is-on", on);
    els.countAutoPlusButton.setAttribute("aria-pressed", on ? "true" : "false");
  }
  if (els.countUndoAutoButton) {
    const stack = state.countAutoUndoStack || [];
    els.countUndoAutoButton.disabled = !stack.length;
    els.countUndoAutoButton.textContent = stack.length ? `Undo +1 (${stack.length})` : "Undo +1";
  }
}

function toggleCountAutoPlusMode() {
  state.countAutoPlusMode = !state.countAutoPlusMode;
  renderCountAutoControls();
  focusCountSearch();
}

function undoLastCountAutoPlus() {
  const session = state.activeCountSession;
  const stack = [...(state.countAutoUndoStack || [])];
  const last = stack.pop();
  if (!session || !last) {
    state.countAutoUndoStack = stack;
    renderCountAutoControls();
    return;
  }
  const entries = [...(session.entries || [])];
  const undoneId = cleanCell(last.entryId);
  const idx = entries.findIndex((entry) => cleanCell(entry.entryId) === undoneId);
  if (idx >= 0 || undoneId) {
    if (idx >= 0) entries.splice(idx, 1);
    const undoEntryIds = [...new Set([...(session.undoEntryIds || []), undoneId].filter(Boolean))];
    const updated = markCountSessionDirty({
      ...session,
      entries: filterUndoneCountEntries(entries, { undoEntryIds }),
      undoEntryIds,
      updatedAt: new Date().toISOString(),
      localSyncPending: true,
    });
    state.activeCountSession = updated;
    state.countAutoUndoStack = stack;
    if (updated.id) {
      const activeId = cleanCell(updated.id);
      state.countSessions = [
        { ...updated, isActiveLive: true },
        ...(state.countSessions || []).filter((saved) => cleanCell(saved?.id) !== activeId),
      ];
    }
    persistActiveCountSession();
    renderCountEntryRows(false);
    updateCountSummaryStrip();
    renderSelectedCountItem();
    renderCountQuantity();
    showToast(`Undid +1 for ${last.code}`, 1800, "warning");
    void syncSharedCountSessionsToSupabase(true).catch(() => {
      // Local state is already correct; the normal queued sync will retry.
      scheduleSharedCountSessionsSync();
    });
  } else {
    state.countAutoUndoStack = stack;
  }
  renderCountAutoControls();
  focusCountSearch();
}

function countSessionAllowsOutOfScope(session = state.activeCountSession) {
  return !isUserRole() && session?.allowOutOfScope === true;
}

function countItemIsInScope(item, session = state.activeCountSession) {
  if (!item) return false;
  return filteredCountCandidateRows(session).some((candidate) => codeKey(candidate.code) === codeKey(item.code));
}

function findAnyCountMatch(query) {
  const raw = cleanCell(query).trim();
  if (!raw) return null;
  const needle = raw.toLowerCase();
  const normalizedNeedle = codeKey(raw);
  if (!state._countSearchIndex || state._countIndexStamp !== state._dataCacheStamp) buildCountSearchIndex();
  return state._countExactIndex?.get(normalizedNeedle)
    || state._countExactIndex?.get(needle)
    || allCountCandidateRows().find((item) =>
      [item.code, item.product, item.plu, item.itemNumber, item.vendor, item.category]
        .some((value) => String(value || "").toLowerCase().includes(needle))
    )
    || null;
}

function findCountExactMatch(query) {
  const raw = cleanCell(query).trim();
  if (!raw) return null;
  const needle = raw.toLowerCase();
  const normalizedNeedle = codeKey(raw);
  if (!state._countSearchIndex || state._countIndexStamp !== state._dataCacheStamp) buildCountSearchIndex();
  const exactMatch = state._countExactIndex?.get(normalizedNeedle)
    || state._countExactIndex?.get(needle)
    || null;
  if (exactMatch && (countItemIsInScope(exactMatch) || countSessionAllowsOutOfScope())) return exactMatch;
  return null;
}

function findCountMatch(query) {
  const raw = cleanCell(query).trim();
  if (!raw) return null;
  const needle = raw.toLowerCase();
  const normalizedNeedle = codeKey(raw); // strips leading zeros for numeric matching

  if (!state._countSearchIndex || state._countIndexStamp !== state._dataCacheStamp) buildCountSearchIndex();
  // Scanner submissions use the exact map before any text search.
  const exactMatch = state._countExactIndex?.get(normalizedNeedle)
    || state._countExactIndex?.get(needle)
    || null;
  if (exactMatch && (countItemIsInScope(exactMatch) || countSessionAllowsOutOfScope())) return exactMatch;

  // Second: try filtered pool with partial text match
  const filtered = countSessionAllowsOutOfScope() ? allCountCandidateRows() : filteredCountCandidateRows();
  return filtered.find((item) =>
    [item.code, item.product, item.plu, item.itemNumber, item.vendor, item.category]
      .some((value) => String(value || "").toLowerCase().includes(needle))
  );
}

function currentSelectedCountItem() {
  if (!state.selectedCountItemCode) return null;
  const selectedKey = codeKey(state.selectedCountItemCode);
  return filteredCountCandidateRows().find((item) => codeKey(item.code) === selectedKey)
    || (countSessionAllowsOutOfScope() ? allCountCandidateRows().find((item) => codeKey(item.code) === selectedKey) : null)
    || null;
}

function renderSelectedCountItem() {
  const item = currentSelectedCountItem();
  if (!els.countSelectedItem) return;
  if (!item) {
    els.countSelectedItem.classList.remove("is-ready");
    els.countSelectedItem.innerHTML = `<p class="muted">Scan or search an item to begin counting.</p>`;
    renderCountKeypadItemSummary();
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


function renderCountKeypadItemSummary() {
  if (!els.countKeypadPanel) return;
  let summary = els.countKeypadPanel.querySelector(".count-keypad-item-summary");
  const item = currentSelectedCountItem();
  if (!summary) {
    summary = document.createElement("div");
    summary.className = "count-keypad-item-summary";
    els.countKeypadPanel.insertBefore(summary, els.countKeypadPanel.firstChild);
  }
  if (!item) {
    summary.hidden = true;
    summary.innerHTML = "";
    return;
  }
  summary.hidden = false;
  summary.innerHTML = `<strong>${escapeHtml(item.product || "-")}</strong><span>Code ${escapeHtml(item.code || "-")} · PLU ${escapeHtml(item.plu || "-")}</span>`;
}

function handleCountAutoPlusLookup() {
  if (!state.activeCountSession && !restoreContinuedCountSession()) {
    showToast("Start a physical count first.", 3000, "warning");
    return;
  }
  const query = cleanCell(els.countSearchInput?.value || "").trim();
  if (!query) return;
  const now = performanceNow();
  const signature = `${query}|${state.activeCountSession?.id || ""}`;
  if (state._countAutoLastCommitSig === signature && now - (state._countAutoLastCommitAt || 0) < 350) {
    if (els.countSearchInput) els.countSearchInput.value = "";
    focusCountSearch();
    return;
  }
  const match = findCountExactMatch(query);
  if (!match) {
    // In +1 Auto, do not use partial/text search. A partial scanner burst or
    // concatenated barcode must fail safely instead of committing a random item.
    hideCountDropdown();
    if (els.countSearchInput) {
      els.countSearchInput.classList.add("count-search-error");
      els.countSearchInput.value = "";
      setTimeout(() => els.countSearchInput && els.countSearchInput.classList.remove("count-search-error"), 900);
    }
    showToast("Barcode not exact-matched. Nothing was added.", 1800, "warning");
    focusCountSearch();
    return;
  }
  state._countAutoLastCommitSig = signature;
  state._countAutoLastCommitAt = now;
  commitCountEntry(match, 1, "add", { autoPlus: true });
  hideCountDropdown();
  if (els.countSearchInput) els.countSearchInput.value = "";
  state._replaceCountInputOnNextKey = false;
  focusCountSearch();
}

function handleCountLookup() {
  if (!state.activeCountSession && !restoreContinuedCountSession()) {
    showToast("Start a physical count first.", 3000, "warning");
    return;
  }
  const query = (els.countSearchInput.value || "").trim();
  if (!query) return;
  const match = findCountMatch(query);
  if (!match) {
    state.selectedCountItemCode = "";
    state.countStage = "search";
    state.countQtyBuffer = "0";
    hideCountDropdown();
    renderSelectedCountItem();
    renderCountQuantity();
    // Immediately reset the scan field so the next barcode replaces the failed one
    // instead of appending to it (ex: 1234 + next scan should not become 12341234).
    if (els.countSearchInput) {
      els.countSearchInput.classList.add("count-search-error");
      els.countSearchInput.value = "";
      state._replaceCountInputOnNextKey = false;
      setTimeout(() => els.countSearchInput && els.countSearchInput.classList.remove("count-search-error"), 1200);
    }
    const existsOutsideScope = !!findAnyCountMatch(query);
    showToast(
      existsOutsideScope
        ? "Item is outside this count scope. A manager must enable the override to count it."
        : "Item not found in this count scope.",
      4200,
      "warning",
    );
    focusCountSearch();
    return;
  }
  const inScope = filteredCountCandidateRows().some((item) => codeKey(item.code) === codeKey(match.code));
  if (!inScope && countSessionAllowsOutOfScope()) {
    showToast("Manager override: this item is outside the selected count scope.", 4200, "warning");
  }
  if (state.countAutoPlusMode) {
    handleCountAutoPlusLookup();
    return;
  }
  state.selectedCountItemCode = match.code;
  state.countStage = "qty";
  state.countQtyBuffer = "0";
  state.pendingDuplicateMode = null;
  hideCountDropdown();
  if (els.countSearchInput) els.countSearchInput.value = "";
  state._replaceCountInputOnNextKey = false;
  renderSelectedCountItem();
  renderCountQuantity();
  openCountKeypadPanel();
  // Normal count flow: after a successful scan, stay in QTY entry mode.
  // Operators can use the keypad/keyboard for qty, then Done/Enter returns to scan.
  // If the scan was wrong, clicking the barcode box intentionally switches back
  // to search/rescan mode.
  if (els.countSearchInput) els.countSearchInput.blur();
}

function clearCountLookup() {
  if (els.countSearchInput) els.countSearchInput.value = "";
  state._replaceCountInputOnNextKey = false;
  hideCountDropdown();
  state.selectedCountItemCode = "";
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.pendingDuplicateMode = null;
  renderCountQuantity();
  renderSelectedCountItem();
  syncCountKeypadVisibility();
  focusCountSearch();
}

function openDuplicateCountModal(item, qty, existing, options = {}) {
  // +1 Auto mode must be truly hands-free: duplicate scans add +1 immediately
  // and must not open the duplicate add/reset confirmation modal.
  if (state.countAutoPlusMode || options.autoPlus) {
    commitCountEntry(item, Number(qty || 1), "add", { autoPlus: true });
    return;
  }
  state.pendingDuplicateCount = { item, qty, existing, chooseBeforeQty: !!options.chooseBeforeQty };
  if (els.countDuplicateMessage) {
    els.countDuplicateMessage.textContent = `${item.product} was already counted as ${number.format(existing?.countedQty || 0)}. Add this new quantity or reset it?`;
  }
  els.countDuplicateModal.hidden = false;
  setTimeout(() => els.countDuplicateAddButton?.focus(), 0);
}

function closeDuplicateCountModal() {
  state.pendingDuplicateCount = null;
  state.pendingDuplicateMode = null;
  els.countDuplicateModal.hidden = true;
  state.countStage = "search";
  state.selectedCountItemCode = "";
  state.countQtyBuffer = "0";
  renderSelectedCountItem();
  renderCountQuantity();
  closeCountKeypadPanel();
  focusCountSearch();
}

function playInventoryCommitSound() {
  // Use the uploaded ding asset for successful +1 Auto / Add Qty commits.
  // Save/finish count intentionally stays silent.
  playCountConfirmBeep();
}

function commitCountEntry(item, qty, mode, options = {}) {
  clearCountAutoCommitTimer?.();
  const session = {
    ...state.activeCountSession,
    updatedAt: new Date().toISOString(),
    entries: [...(state.activeCountSession.entries || [])],
  };
  const matchingEntries = session.entries.filter((entry) => codeKey(entry.code) === codeKey(item.code));
  const existing = matchingEntries.at(-1);
  const autoAlreadyCounted = matchingEntries
    .filter((entry) => entry?.autoPlus === true)
    .reduce((sum, entry) => sum + Math.max(0, Number(entry.inputQty ?? entry.qty ?? 1) || 0), 0);
  const countedQty = options.autoPlus
    ? Math.max(0, autoAlreadyCounted + qty)
    : mode === "add"
      ? Math.max(0, Number(existing?.countedQty || 0) + qty)
      : qty;
  if (qty > 99 && !options.largeQtyConfirmed) {
    const currentQty = mode === "add" ? Number(existing?.countedQty || 0) : 0;
    showAppConfirm({
      title: "⚠ Large Quantity",
      message: `Item: ${item.product} | Current Qty: ${number.format(currentQty)} | Adding: ${number.format(qty)} | New Total: ${number.format(countedQty)}. Confirm?`,
      cancelText: "Cancel",
      confirmText: "Confirm",
      onConfirm: () => commitCountEntry(item, qty, mode, { ...options, largeQtyConfirmed: true }),
    });
    return;
  }
  const entryId = makeCountIdentifier("entry");
  session.entries.push({
    entryId,
    sessionId: session.id,
    deviceId: countDeviceId(),
    deviceLabel: countDeviceLabel(),
    user: currentAuditUser(),
    code: item.code,
    plu: item.plu || "",
    itemNumber: item.itemNumber || "",
    product: item.product,
    vendor: item.vendor || "",
    category: item.category || "",
    originalQty: item.stock || 0,
    inputQty: qty,
    qty,
    countedQty,
    mode,
    autoPlus: !!options.autoPlus,
    unitCost: item.unitCost || 0,
    recordedAt: new Date().toISOString(),
    timestamp: new Date().toISOString(),
    syncStatus: "pending",
  });
  if (options.autoPlus) {
    state.countAutoUndoStack = [...(state.countAutoUndoStack || []), { entryId, code: item.code }].slice(-50);
  }
  state.activeCountSession = session;
  if (session.id) {
    const activeId = cleanCell(session.id);
    state.countSessions = [
      { ...session, isActiveLive: true },
      ...(state.countSessions || []).filter((saved) => cleanCell(saved?.id) !== activeId),
    ];
  }
  state._countSessionOpen = true;
  state.countQtyBuffer = "0";
  state.countStage = "search";
  state.selectedCountItemCode = "";
  state.pendingDuplicateMode = null;
  persistActiveCountSession();
  if (options.autoPlus || mode === "add") playInventoryCommitSound();
  // Fast path: prepend new row, update summary, reset UI - no full workspace rebuild
  renderCountEntryRows(true);
  renderSelectedCountItem();
  renderCountQuantity();
  updateCountSummaryStrip();
  if (els.countSearchInput) els.countSearchInput.value = "";
  state._replaceCountInputOnNextKey = false;
  hideCountDropdown();
  focusCountSearch();
  if (mode === "add") {
    requestAnimationFrame(() => showToast(
      `${options.autoPlus ? "+1 Auto" : "Added " + number.format(qty)} to ${item.code}`,
      1600,
      "success",
    ));
  }
  renderCountAutoControls();
  if (!options.autoPlus) closeCountKeypadPanel();
}

function resolveDuplicateCount(mode) {
  const pending = state.pendingDuplicateCount;
  if (!pending) return;
  els.countDuplicateModal.hidden = true;
  state.pendingDuplicateCount = null;
  const resolvedMode = mode === "add" ? "add" : "reset";
  if (pending.qty == null) return;
  state.pendingDuplicateMode = null;
  commitCountEntry(pending.item, pending.qty, resolvedMode);
}

function applyCountEntry() {
  if (!state.activeCountSession && !restoreContinuedCountSession()) {
    showToast("Start a physical count first.", 3000, "warning");
    return;
  }
  if (looksLikeBarcodeValue(state.countQtyBuffer || "") && redirectQtyBarcodeToSearch(state.countQtyBuffer)) {
    showToast("That looks like a barcode, not a quantity.", 2200, "warning");
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
    // In +1 Auto mode, repeated scans are continuous additions.
    // Normal/manual mode still uses the duplicate add/reset modal.
    if (state.countAutoPlusMode) {
      commitCountEntry(item, 1, "add", { autoPlus: true });
      return;
    }
    const chosenMode = state.pendingDuplicateMode && codeKey(state.pendingDuplicateMode.code) === codeKey(item.code)
      ? state.pendingDuplicateMode.mode
      : "";
    if (chosenMode) {
      commitCountEntry(item, qty, chosenMode);
      return;
    }
    openDuplicateCountModal(item, qty, existing);
    return;
  }
  commitCountEntry(item, qty, "reset");
}

function focusCountSearch() {
  if (!els.countSearchInput) return;
  // KBD-C: On touch/Kindle do NOT call select() -- it fights the soft keyboard.
  // On desktop keep select-all so barcode scanner bursts replace prior text cleanly.
  setTimeout(() => {
    if (!els.countSearchInput) return;
    els.countSearchInput.focus();
    if (!scanPerformanceProfile.isMobile && !scanPerformanceProfile.isSilk) {
      els.countSearchInput.select?.();
    }
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
  state._countExactIndex = new Map();
  allRows.forEach((item) => {
    [item.code, item.plu, item.itemNumber].forEach((value) => {
      const rawKey = cleanCell(value).toLowerCase();
      const normalizedKey = codeKey(value);
      if (rawKey && !state._countExactIndex.has(rawKey)) state._countExactIndex.set(rawKey, item);
      if (normalizedKey && !state._countExactIndex.has(normalizedKey)) state._countExactIndex.set(normalizedKey, item);
    });
  });
  // Pre-compute filtered scope codes
  state._countFilteredCodes = new Set(
    filteredCountCandidateRows().map((r) => codeKey(r.code))
  );
  state._countIndexStamp = state._dataCacheStamp;
}

function renderCountDropdown(query) {
  const dd = document.querySelector("#countSearchDropdown");
  if (!dd) return;
  if (state.countAutoPlusMode) { dd.hidden = true; return; }
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
  const scopeLabel = [session?.vendor, session?.category, session?.status].filter(Boolean).join("  -  ") || "All items";
  const counted = new Set((session?.entries || []).map((e) => codeKey(e.code)));

  let html = "";
  if (inScopeMatches.length) {
    html += `<div class="count-dd-group-label">✓ In scope - ${escapeHtml(scopeLabel)}</div>`;
    html += inScopeMatches.map(({ item }) => {
      const alreadyCounted = counted.has(codeKey(item.code));
      return `<div class="count-dd-item${alreadyCounted ? " count-dd-counted" : ""}" data-code="${escapeHtml(item.code)}">
        <span class="count-dd-name">${escapeHtml(item.product)}</span>
        <span class="count-dd-meta">${escapeHtml(item.code)}${item.vendor ? `  -  ${escapeHtml(item.vendor)}` : ""}${alreadyCounted ? "  -  <b>Counted</b>" : ""}</span>
      </div>`;
    }).join("");
  }
  if (outScopeMatches.length && countSessionAllowsOutOfScope()) {
    html += `<div class="count-dd-group-label count-dd-out-label">Outside scope - manager override enabled</div>`;
    html += outScopeMatches.map(({ item }) => `
      <div class="count-dd-item count-dd-out" data-code="${escapeHtml(item.code)}" title="Manager override allows this out-of-scope item">
        <span class="count-dd-name">${escapeHtml(item.product)}</span>
        <span class="count-dd-meta">${escapeHtml(item.code)}${item.vendor ? `  -  ${escapeHtml(item.vendor)}` : ""}</span>
      </div>`).join("");
  }

  dd.innerHTML = html;
  dd.hidden = false;

  dd.querySelectorAll(".count-dd-item[data-code]").forEach((el) => {
    el.addEventListener("mousedown", (e) => {
      e.preventDefault();
      selectCountDropdownItem(el.dataset.code);
    });
  });
}

function selectCountDropdownItem(code) {
  hideCountDropdown();
  if (!state.activeCountSession) return;
  const key = codeKey(code);
  const filteredItem = filteredCountCandidateRows().find((r) => codeKey(r.code) === key);
  const item = filteredItem || (countSessionAllowsOutOfScope() ? allCountCandidateRows().find((r) => codeKey(r.code) === key) : null);
  if (!item) { showToast("Item not found in the item master.", 2800, "warning"); return; }
  if (!filteredItem) showToast("Manager override: counting an item outside the selected scope.", 3600, "warning");
  if (els.countSearchInput) els.countSearchInput.value = "";
  if (state.countAutoPlusMode) {
    commitCountEntry(item, 1, "add", { autoPlus: true });
    focusCountSearch();
    return;
  }
  state.selectedCountItemCode = item.code;
  state.countStage = "qty";
  state.countQtyBuffer = "0";
  state.pendingDuplicateMode = null;
  els.countSearchInput?.blur();
  renderSelectedCountItem();
  renderCountQuantity();
  openCountKeypadPanel();
  if (els.countSearchInput) els.countSearchInput.blur();
}

function findCountSessionById(sessionId) {
  if (!sessionId) return null;
  if (countSessionIsDeleted(sessionId)) return null;
  if (state.activeCountSession?.id === sessionId) return state.activeCountSession;
  return state.countSessions.find((session) => session.id === sessionId) || null;
}

function currentCountSessionCandidates(session) {
  const rows = filteredCountCandidateRows(session);
  const scannedCodes = new Set((session?.entries || []).map((entry) => codeKey(entry.code)));
  return rows.filter((item) => {
    const st = (item.state || "").trim().toLowerCase();
    if ((st === "disabled" || st === "discontinued") && !scannedCodes.has(codeKey(item.code))) return false;
    return true;
  });
}


function activeReportEntries(session = {}) {
  const normalized = normalizeCountSession(session, "pending");
  return filterUndoneCountEntries(normalized.entries || [], normalized);
}


function countInputLogDisplayRows(entries = []) {
  const runningByCode = new Map();
  const sorted = sortedCountEntries(entries);
  const displayByEntryId = new Map();
  sorted.forEach((entry) => {
    const key = codeKey(entry?.code);
    const priorSessionAdd = runningByCode.has(key) ? Number(runningByCode.get(key) || 0) : 0;
    const originalQty = Number(entry?.originalQty || 0);
    const countedQty = Number(entry?.countedQty || 0);
    let displayBefore = originalQty;
    let displayAfter = countedQty;
    let displayVariance = displayAfter - displayBefore;

    // Input Log display only: Add / Auto +1 entries are stored as cumulative
    // countedQty against the item's real stock. For the audit view, show the
    // session add action itself: 0 -> 1, 1 -> 2, etc. This does NOT affect
    // comparison/final inventory math.
    const mode = String(entry?.mode || "").toLowerCase();
    const isAddAction = entry?.autoPlus === true || mode === "add";
    if (isAddAction) {
      const cumulativeAdded = Math.max(0, countedQty - originalQty);
      let delta = Number(entry.inputQty ?? entry.qty ?? 0) || 0;
      if (!delta) delta = Math.max(1, cumulativeAdded - priorSessionAdd);
      displayBefore = priorSessionAdd;
      displayAfter = priorSessionAdd + delta;
      displayVariance = delta;
    }

    if (key) {
      const nextRunning = isAddAction ? displayAfter : Math.max(0, countedQty - originalQty);
      runningByCode.set(key, nextRunning);
    }
    const row = {
      ...entry,
      inputLogQtyBefore: displayBefore,
      inputLogQtyAfter: displayAfter,
      inputLogVariance: displayVariance,
    };
    displayByEntryId.set(cleanCell(entry?.entryId || `${entry?.code || ''}|${entry?.recordedAt || entry?.timestamp || ''}`), row);
  });
  return entries.map((entry) => {
    const id = cleanCell(entry?.entryId || `${entry?.code || ''}|${entry?.recordedAt || entry?.timestamp || ''}`);
    return displayByEntryId.get(id) || entry;
  });
}

function latestCountEntryByCode(entries = []) {
  const map = new Map();
  [...entries]
    .sort((a, b) => String(a.recordedAt || a.timestamp || "").localeCompare(String(b.recordedAt || b.timestamp || "")) || String(a.entryId || "").localeCompare(String(b.entryId || "")))
    .forEach((entry) => map.set(codeKey(entry.code), entry));
  return map;
}


function sortedCountEntries(entries = []) {
  return [...entries].sort((a, b) =>
    String(a.recordedAt || a.timestamp || "").localeCompare(String(b.recordedAt || b.timestamp || "")) ||
    String(a.entryId || "").localeCompare(String(b.entryId || ""))
  );
}

function effectiveCountFinalEntry(list = []) {
  const sorted = sortedCountEntries(list).filter(Boolean);
  if (!sorted.length) return { entry: null, countedQty: 0 };
  const latest = sorted[sorted.length - 1];
  if (latest?.autoPlus === true) {
    let total = 0;
    for (const entry of sorted) {
      if (entry?.autoPlus === true) {
        total += Math.max(0, Number(entry.inputQty ?? entry.qty ?? 1) || 0);
      } else {
        total = 0;
      }
    }
    return { entry: latest, countedQty: total };
  }
  return { entry: latest, countedQty: Number(latest?.countedQty ?? latest?.inputQty ?? 0) || 0 };
}

function countReportStatus(entry, qtyDiff) {
  if (!entry) return "Not scanned";
  return Number(qtyDiff || 0) === 0 ? "PASS" : "QTY DIFF";
}

function countStartSnapshotQty(session = {}, item = {}, firstEntry = null) {
  const keys = [codeKey(item.code), codeKey(firstEntry?.code), codeKey(item.plu), codeKey(item.itemNumber)]
    .filter(Boolean);
  const snapshots = [session.countStartSnapshot, session.startSnapshot, session.preCountSnapshot]
    .filter((snap) => snap && typeof snap === "object");
  for (const snap of snapshots) {
    for (const key of keys) {
      if (Object.prototype.hasOwnProperty.call(snap, key)) return Number(snap[key] || 0);
    }
  }
  if (firstEntry && Number.isFinite(Number(firstEntry.originalQty))) return Number(firstEntry.originalQty || 0);
  if (Number.isFinite(Number(item.stock))) return Number(item.stock || 0);
  return 0;
}

// REPORT RECALC SAFETY:
// Build comparison rows from the saved session entries first, then fill in any
// unscanned in-scope items from the current item master. This lets older saved
// counts re-open with corrected PASS/QTY DIFF logic without mutating the saved
// session or touching sync/session flow.
function countReportComparisonRows(session = {}) {
  const normalized = normalizeCountSession(session, "synced");
  const entries = activeReportEntries(normalized);
  const entriesByCode = new Map();
  entries
    .slice()
    .sort((a, b) => String(a.recordedAt || a.timestamp || "").localeCompare(String(b.recordedAt || b.timestamp || "")) || String(a.entryId || "").localeCompare(String(b.entryId || "")))
    .forEach((entry) => {
      const key = codeKey(entry.code);
      if (!key) return;
      if (!entriesByCode.has(key)) entriesByCode.set(key, []);
      entriesByCode.get(key).push(entry);
    });

  const firstByCode = new Map();
  const latestByCode = new Map();
  const finalQtyByCode = new Map();
  entriesByCode.forEach((list, key) => {
    const sorted = sortedCountEntries(list);
    firstByCode.set(key, sorted[0]);
    // FINAL COMPARISON RULE:
    // Input Log stays the full audit trail. Comparison/PDF/Excel use the
    // latest action only. For +1 Auto, the final action is the cumulative
    // physical scan count from 0, not original QOH + scans.
    const effective = effectiveCountFinalEntry(sorted);
    latestByCode.set(key, effective.entry);
    finalQtyByCode.set(key, effective.countedQty);
  });

  const itemByCode = new Map();
  currentCountSessionCandidates(normalized).forEach((item) => {
    const key = codeKey(item.code);
    if (key) itemByCode.set(key, item);
  });

  // Ensure scanned items are always visible even if the current product snapshot
  // no longer matches the old session filter or the item changed later.
  entries.forEach((entry) => {
    const key = codeKey(entry.code);
    if (!key || itemByCode.has(key)) return;
    itemByCode.set(key, {
      code: entry.code,
      plu: entry.plu || entry.itemNumber || "",
      itemNumber: entry.itemNumber || "",
      product: entry.product || "",
      vendor: entry.vendor || "",
      category: entry.category || "",
      stock: Number(entry.originalQty || 0),
      unitCost: Number(entry.unitCost || 0),
    });
  });

  return [...itemByCode.values()].map((item) => {
    const key = codeKey(item.code);
    const firstEntry = firstByCode.get(key);
    const latestEntry = latestByCode.get(key);
    const entry = latestEntry;
    // Report truth table:
    // - Input Log = every action.
    // - Comparison/PDF/Excel = final state from the latest saved entry.
    // - A missing entry means the item was NOT SCANNED, so Qty After is 0 and
    //   Qty Diff is 0 - Qty Before.
    const originalQty = countStartSnapshotQty(normalized, item, entry || firstEntry);
    const finalQty = entry ? Number(finalQtyByCode.get(key) ?? entry.countedQty ?? entry.inputQty ?? 0) : 0;
    const isNull = !entry;
    const qtyDiff = finalQty - originalQty;
    const unitCost = Number(entry?.unitCost ?? item.unitCost ?? 0);
    const costDiff = qtyDiff * unitCost;
    const status = isNull ? "Not scanned" : qtyDiff === 0 ? "PASS" : "QTY DIFF";
    return { item, entry, firstEntry, originalQty, finalQty, isNull, qtyDiff, costDiff, status };
  }).sort((a, b) => {
    // Wrong / missing items first for manager review, PASS last.
    const rank = (row) => row.isNull ? 0 : row.qtyDiff !== 0 ? 1 : 2;
    return rank(a) - rank(b) || Math.abs(b.qtyDiff) - Math.abs(a.qtyDiff) || String(a.item.product || "").localeCompare(String(b.item.product || ""));
  });
}

function countReportRowClass({ isNull, qtyDiff, status } = {}) {
  if (isNull || String(status || "").toLowerCase().includes("not scanned")) return "count-report-row-null";
  const diff = Number(qtyDiff || 0);
  if (diff < 0) return "count-report-row-negative";
  if (diff > 0) return "count-report-row-positive";
  return "count-report-row-zero";
}

function countReportSummary(rows = []) {
  const pass = rows.filter((r) => !r.isNull && Number(r.qtyDiff || 0) === 0).length;
  const qtyDiff = rows.filter((r) => !r.isNull && Number(r.qtyDiff || 0) !== 0).length;
  const notScanned = rows.filter((r) => r.isNull).length;
  return { pass, qtyDiff, notScanned, totalIssues: qtyDiff + notScanned };
}

function openCountReport(sessionId = state.activeCountSession?.id, mode = state.countReportMode || "input") {
  const session = findCountSessionById(sessionId);
  if (!session) {
    showToast("No saved count report found yet.", 3000, "warning");
    return;
  }
  state.countReportMode = mode;
  state.countReportOpenId = sessionId;
  els.countContinueButton?.blur();
  if (els.countReportTitle) {
    els.countReportTitle.textContent = mode === "comparison"
      ? `${countSessionLabel(session)}  -  Comparison report`
      : `${countSessionLabel(session)}  -  Input log`;
  }
  if (els.countReportHead) {
    els.countReportHead.innerHTML = mode === "comparison"
      ? `<tr><th data-count-report-sort="code">Code</th><th data-count-report-sort="plu">PLU</th><th data-count-report-sort="item">Item</th><th data-count-report-sort="vendor">Vendor</th><th data-count-report-sort="category">Category</th><th data-count-report-sort="originalQty">Qty before</th><th data-count-report-sort="finalQty">Qty after</th><th data-count-report-sort="isNull">NULL</th><th data-count-report-sort="qtyDiff">Qty diff</th><th data-count-report-sort="costDiff">Cost diff</th><th data-count-report-sort="status">Status</th></tr>`
      : `<tr><th data-count-report-sort="code">Code</th><th data-count-report-sort="plu">PLU</th><th data-count-report-sort="item">Item</th><th data-count-report-sort="vendor">Vendor</th><th data-count-report-sort="category">Category</th><th data-count-report-sort="originalQty">Qty before</th><th data-count-report-sort="finalQty">Qty after</th><th data-count-report-sort="qtyDiff">Variance</th><th data-count-report-sort="mode">Mode</th><th data-count-report-sort="time">Date/Time</th></tr>`;
  }
  if (els.countReportMeta) {
    const vendorLabel = session.vendor || session.department || "All vendors";
    const totalCandidates = currentCountSessionCandidates(session).length;
    const syncStats = countSessionSyncStats(session);
    els.countReportMeta.innerHTML = `
      <span><b>Date</b> ${escapeHtml(session.date || "-")}</span>
      <span><b>Vendor</b> ${escapeHtml(vendorLabel)}</span>
      <span><b>Category</b> ${escapeHtml(session.category || "All")}</span>
      <span><b>Entries</b> ${number.format((session.entries || []).length)}</span>
      <span><b>Items in scope</b> ${number.format(totalCandidates)}</span>
      <span><b>Started</b> ${escapeHtml(new Date(session.startedAt).toLocaleString())}</span>
      <span><b>Sync</b> ${syncStats.failed ? `${number.format(syncStats.failed)} failed` : syncStats.pending || session.localSyncPending ? `${number.format(syncStats.pending)} pending` : "Synced"}</span>`;
  }
  renderCountReportRows(session, mode);
  setupCountReportHeaderSorting();
  els.countReportModal.hidden = false;
}

function closeCountReport() {
  els.countReportModal.hidden = true;
}

function currentCountReportSessionForExport() {
  // Export the report the manager is actually viewing.
  // The old order used activeCountSession / first history row, which could grab
  // a blank duplicate session instead of the open 3:41 scanned session.
  const preferredIds = [
    state.countReportOpenId,
    state.activeCountSession?.id,
    state._continuingCountId,
  ].map(cleanCell).filter(Boolean);
  for (const id of preferredIds) {
    const session = findCountSessionById(id);
    if (session) return session;
  }
  return state.countSessions.find((session) => !countSessionIsDeleted(session?.id) && (session.entries || []).length)
    || state.countSessions.find((session) => !countSessionIsDeleted(session?.id))
    || null;
}

function exportCountReportPdf() {
  const session = currentCountReportSessionForExport();
  if (!session) { showToast("No count report to export.", 3000, "warning"); return; }
  state.countReportOpenId = session.id;
  const mode = state.countReportMode || "input";
  const dateStr = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const vendorLabel = session.vendor || session.department || "All vendors";
  const entries = activeReportEntries(session);
  const syncStats = countSessionSyncStats(session);
  const syncWarning = syncStats.pending || syncStats.failed || session.localSyncPending
    ? `WARNING: ${syncStats.pending} pending and ${syncStats.failed} failed sync entries. Verify before POS import.`
    : "All count entries synced.";

  let tableHtml = "";
  if (mode === "comparison") {
    const comparisonRows = countReportComparisonRows(session);
    tableHtml = `<table>
      <thead><tr><th>Code</th><th>PLU</th><th>Item</th><th>Vendor</th><th>Category</th><th class="num">Qty Before</th><th class="num">Qty After</th><th>NULL</th><th class="num">Qty Diff</th><th class="num">Cost Diff</th><th>Status</th></tr></thead>
      <tbody>${comparisonRows.map(({ item, entry, originalQty, finalQty, isNull, qtyDiff, costDiff, status }) => {
        const cellCls = isNull ? "cell-null" : qtyDiff > 0 ? "cell-positive" : qtyDiff < 0 ? "cell-negative" : "";
        return `<tr><td>${escapeHtml(item.code)}</td><td>${escapeHtml(item.plu || entry?.plu || "-")}</td><td>${escapeHtml(item.product)}</td><td>${escapeHtml(item.vendor || "-")}</td><td>${escapeHtml(item.category || "-")}</td>
          <td class="num">${number.format(originalQty)}</td>
          <td class="num">${isNull ? "-" : number.format(finalQty)}</td>
          <td class="${isNull ? "cell-null" : ""}">${isNull ? "NULL" : ""}</td>
          <td class="num ${cellCls}">${qtyDiff > 0 ? `+${number.format(qtyDiff)}` : number.format(qtyDiff)}</td>
          <td class="num ${cellCls}">${currency.format(costDiff)}</td>
          <td class="${cellCls}">${status}</td></tr>`;
      }).join("")}</tbody></table>`;
  } else {
    tableHtml = `<table>
      <thead><tr><th>Code</th><th>PLU</th><th>Item</th><th>Vendor</th><th>Category</th><th class="num">Qty Before</th><th class="num">Counted</th><th class="num">Variance</th><th>Mode</th><th>Time</th></tr></thead>
      <tbody>${countInputLogDisplayRows(entries).reverse().map((entry) => {
        const variance = Number(entry.inputLogVariance ?? (Number(entry.countedQty || 0) - Number(entry.originalQty || 0)));
        const cls = variance > 0 ? "var-up" : variance < 0 ? "var-down" : "";
        return `<tr class="${cls}"><td>${escapeHtml(entry.code)}</td><td>${escapeHtml(entry.plu || "-")}</td><td>${escapeHtml(entry.product)}</td><td>${escapeHtml(entry.vendor || "-")}</td><td>${escapeHtml(entry.category || "-")}</td>
          <td class="num">${number.format(entry.inputLogQtyBefore ?? entry.originalQty ?? 0)}</td>
          <td class="num">${number.format(entry.inputLogQtyAfter ?? entry.countedQty ?? 0)}</td>
          <td class="num">${variance > 0 ? `+${number.format(variance)}` : number.format(variance)}</td>
          <td>${escapeHtml(entry.mode || "set")}</td>
          <td>${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</td></tr>`;
      }).join("")}</tbody></table>`;
  }

  const html = `<!DOCTYPE html><html><head><meta charset="UTF-8">
  <title>Physical Count Report - ${escapeHtml(session.date || "-")}</title>
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
    .cell-negative { background: #fde2e2; color: #991b1b; font-weight: 700; }
    .cell-positive { background: #dbeafe; color: #1d4ed8; font-weight: 700; }
    .cell-null { background: #eeeeee; color: #6b7280; font-weight: 700; }
    @media print { body { padding: 8px; } .cell-negative, .cell-positive, .cell-null { -webkit-print-color-adjust: exact; print-color-adjust: exact; } }
  </style></head><body>
  <h1>Physical Count Report</h1>
  <div class="meta">
    <span><b>Date</b> ${escapeHtml(session.date || "-")}</span>
    <span><b>Vendor</b> ${escapeHtml(vendorLabel)}</span>
    <span><b>Category</b> ${escapeHtml(session.category || "All")}</span>
    <span><b>Entries</b> ${number.format(entries.length)}</span>
    <span><b>Mode</b> ${mode === "comparison" ? "Comparison" : "Input Log"}</span>
    <span><b>Sync</b> ${escapeHtml(syncWarning)}</span>
    <span><b>Generated</b> ${dateStr}</span>
  </div>
  ${tableHtml}
  </body></html>`;

  const win = window.open("", "_blank", "width=1000,height=750");
  if (!win) { showToast("Pop-up blocked - allow pop-ups to export PDF.", 3500, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 500);
  showToast("Print dialog opened - use 'Save as PDF'.", 3200, "success");
}

async function exportCountReportExcel() {
  const session = currentCountReportSessionForExport();
  if (!session) { showToast("No count report to export.", 3000, "warning"); return; }
  state.countReportOpenId = session.id;
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available.", 3000, "warning"); return; }

  const vendorLabel = session.vendor || session.department || "All vendors";
  const entries = activeReportEntries(session);
  const syncStats = countSessionSyncStats(session);
  const syncNote = syncStats.pending || syncStats.failed || session.localSyncPending
    ? `WARNING: ${syncStats.pending} pending / ${syncStats.failed} failed sync entries. Verify before POS import.`
    : "Sync: all entries confirmed.";
  const wb = xlsx.utils.book_new();
  const applyCountReportCellFill = (ws, rowIndexZeroBased, colIndexZeroBased, hex) => {
    // Safe no-op in basic SheetJS builds that ignore styles; supported builds
    // will preserve these fills in the exported Excel workbook.
    const addr = xlsx.utils.encode_cell({ r: rowIndexZeroBased, c: colIndexZeroBased });
    if (!ws[addr]) ws[addr] = { t: "s", v: "" };
    ws[addr].s = Object.assign({}, ws[addr].s || {}, { fill: { fgColor: { rgb: hex } } });
  };

  // Input log sheet
  const inputData = [
    ["Physical Count - Input Log", "", "", `Date: ${session.date || "-"}`, `Vendor: ${vendorLabel}`, `Category: ${session.category || "All"}`, syncNote],
    [],
    ["Code", "PLU", "Item", "Vendor", "Category", "Qty Before", "Counted", "Variance", "Mode", "Time"],
    ...countInputLogDisplayRows(entries).reverse().map((entry) => {
      const variance = Number(entry.inputLogVariance ?? (Number(entry.countedQty || 0) - Number(entry.originalQty || 0)));
      return [entry.code, entry.plu || "", entry.product, entry.vendor || "", entry.category || "", entry.inputLogQtyBefore ?? entry.originalQty ?? 0, entry.inputLogQtyAfter ?? entry.countedQty ?? 0, variance, entry.mode || "set", new Date(entry.recordedAt).toLocaleString()];
    }),
  ];
  const wsInput = xlsx.utils.aoa_to_sheet(inputData);
  wsInput["!cols"] = [12, 12, 32, 14, 14, 10, 10, 10, 8, 20].map((w) => ({ wch: w }));
  countInputLogDisplayRows(entries).reverse().forEach((entry, i) => {
    const variance = Number(entry.inputLogVariance ?? (Number(entry.countedQty || 0) - Number(entry.originalQty || 0)));
    const fill = variance < 0 ? "FDE2E2" : variance > 0 ? "DBEAFE" : "FFFFFF";
    if (variance !== 0) applyCountReportCellFill(wsInput, i + 3, 7, fill);
  });
  xlsx.utils.book_append_sheet(wb, wsInput, "Input Log");

  // Comparison sheet
  const comparisonRows = countReportComparisonRows(session);
  const compData = [
    ["Physical Count - Comparison", "", "", `Date: ${session.date || "-"}`, `Vendor: ${vendorLabel}`, `Category: ${session.category || "All"}`, syncNote],
    [],
    ["Code", "PLU", "Item", "Vendor", "Category", "Qty Before", "Qty After", "NULL", "Qty Diff", "Cost Diff", "Status"],
    ...comparisonRows.map(({ item, entry, originalQty, finalQty, isNull, qtyDiff, costDiff, status }) => {
      return [item.code, item.plu || entry?.plu || "", item.product, item.vendor || "", item.category || "", originalQty, finalQty, isNull ? "NULL" : "", qtyDiff, costDiff, status];
    }),
  ];
  const wsComp = xlsx.utils.aoa_to_sheet(compData);
  wsComp["!cols"] = [12, 12, 32, 14, 14, 10, 10, 8, 10, 10, 12].map((w) => ({ wch: w }));
  comparisonRows.forEach((row, i) => {
    const diff = Number(row.qtyDiff || 0);
    const fill = row.isNull ? "EEEEEE" : diff < 0 ? "FDE2E2" : diff > 0 ? "DBEAFE" : "FFFFFF";
    if (row.isNull) applyCountReportCellFill(wsComp, i + 3, 7, fill);
    if (row.isNull || diff !== 0) {
      applyCountReportCellFill(wsComp, i + 3, 8, fill);
      applyCountReportCellFill(wsComp, i + 3, 9, fill);
      applyCountReportCellFill(wsComp, i + 3, 10, fill);
    }
  });
  xlsx.utils.book_append_sheet(wb, wsComp, "Comparison");

  xlsx.writeFile(wb, `PhysicalCount_${session.date || "report"}.xlsx`);
  showToast(`Count report exported - ${entries.length} entries, ${comparisonRows.length} comparison rows`, 3200, "success");
}


function countReportSortValue(row, key, mode) {
  if (mode === "comparison") {
    const item = row.item || {};
    const entry = row.entry || {};
    if (key === "code") return cleanCell(item.code || entry.code);
    if (key === "plu") return cleanCell(item.plu || entry.plu || entry.itemNumber);
    if (key === "item") return cleanCell(item.product || entry.product);
    if (key === "vendor") return cleanCell(item.vendor || entry.vendor);
    if (key === "category") return cleanCell(item.category || entry.category);
    if (key === "originalQty") return Number(row.originalQty || 0);
    if (key === "finalQty") return Number(row.finalQty || 0);
    if (key === "isNull") return row.isNull ? 1 : 0;
    if (key === "qtyDiff") return Number(row.qtyDiff || 0);
    if (key === "costDiff") return Number(row.costDiff || 0);
    if (key === "status") return cleanCell(row.status || "");
  }
  const entry = row || {};
  if (key === "code") return cleanCell(entry.code);
  if (key === "plu") return cleanCell(entry.plu || entry.itemNumber);
  if (key === "item") return cleanCell(entry.product);
  if (key === "vendor") return cleanCell(entry.vendor);
  if (key === "category") return cleanCell(entry.category);
  if (key === "originalQty") return Number(entry.inputLogQtyBefore ?? entry.originalQty ?? 0);
  if (key === "finalQty") return Number(entry.inputLogQtyAfter ?? entry.countedQty ?? 0);
  if (key === "qtyDiff") return Number(entry.inputLogVariance ?? (Number(entry.countedQty || 0) - Number(entry.originalQty || 0)));
  if (key === "mode") return cleanCell(entry.mode || "set");
  if (key === "time") return new Date(entry.recordedAt || entry.timestamp || 0).getTime() || 0;
  return "";
}

function sortCountReportRows(rows = [], mode = state.countReportMode || "input") {
  const sort = state.countReportSort || {};
  if (!sort.key || !sort.dir) return rows;
  const dir = sort.dir === "desc" ? -1 : 1;
  return [...rows].sort((a, b) => {
    const av = countReportSortValue(a, sort.key, mode);
    const bv = countReportSortValue(b, sort.key, mode);
    if (typeof av === "number" || typeof bv === "number") return ((Number(av) || 0) - (Number(bv) || 0)) * dir;
    return String(av || "").localeCompare(String(bv || ""), undefined, { numeric: true, sensitivity: "base" }) * dir;
  });
}

function setupCountReportHeaderSorting() {
  const table = document.querySelector("#countReportTable");
  if (!table || table.dataset.sortReady === "1") return;
  table.dataset.sortReady = "1";
  table.addEventListener("click", (event) => {
    const th = event.target.closest("th[data-count-report-sort]");
    if (!th || !table.contains(th)) return;
    const key = th.dataset.countReportSort;
    const current = state.countReportSort || {};
    const dir = current.key !== key ? "asc" : current.dir === "asc" ? "desc" : current.dir === "desc" ? "" : "asc";
    state.countReportSort = dir ? { key, dir } : { key: "", dir: "" };
    const session = findCountSessionById(state.countReportOpenId || state.activeCountSession?.id);
    if (session) renderCountReportRows(session, state.countReportMode || "input");
  });
}


function markCountReportSortHeaders() {
  const headers = document.querySelectorAll("#countReportTable th[data-count-report-sort]");
  const sort = state.countReportSort || {};
  headers.forEach((th) => {
    const active = sort.key && th.dataset.countReportSort === sort.key;
    th.classList.toggle("is-sorted", !!active);
    th.dataset.sortDir = active ? sort.dir : "";
  });
}

function renderCountReportRows(session, mode = state.countReportMode || "input") {
  if (!els.countReportBody) return;
  const entries = activeReportEntries(session || {});
  if (mode === "comparison") {
    const comparisonRows = sortCountReportRows(countReportComparisonRows(session), "comparison");
    markCountReportSortHeaders();
    if (!comparisonRows.length) {
      els.countReportBody.innerHTML = `<tr><td colspan="11" class="empty-cell">No items matched this count criteria.</td></tr>`;
      return;
    }
    els.countReportBody.innerHTML = comparisonRows
      .map(({ item, entry, originalQty, finalQty, isNull, qtyDiff, costDiff, status }) => {
        const qtyClass = qtyDiff > 0 ? "variance-up" : qtyDiff < 0 ? "variance-down" : "variance-flat";
        const costClass = costDiff > 0 ? "variance-up" : costDiff < 0 ? "variance-down" : "variance-flat";
        const statusHtml = !entry
          ? `<span class="muted">Not scanned</span>`
          : status === "PASS"
            ? `<span class="count-pass-badge">PASS</span> <span class="muted">${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</span>`
            : `<span class="count-diff-badge">QTY DIFF</span> <span class="muted">${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</span>`;
        return `
          <tr>
            <td>${escapeHtml(item.code || "-")}</td>
            <td>${escapeHtml(item.plu || entry?.plu || "-")}</td>
            <td>${escapeHtml(item.product || "-")}</td>
            <td>${escapeHtml(item.vendor || "-")}</td>
            <td>${escapeHtml(item.category || "-")}</td>
            <td class="num">${number.format(originalQty)}</td>
            <td class="num">${number.format(finalQty)}</td>
            <td class="${isNull ? "variance-null" : ""}">${isNull ? `<span>NULL</span>` : ""}</td>
            <td class="num ${isNull ? "variance-null" : qtyClass}">${qtyDiff > 0 ? `+${number.format(qtyDiff)}` : number.format(qtyDiff)}</td>
            <td class="${isNull ? "variance-null" : costClass}">${currency.format(costDiff)}</td>
            <td class="${isNull ? "variance-null" : qtyDiff > 0 ? "variance-up" : qtyDiff < 0 ? "variance-down" : "variance-flat"}">${statusHtml}</td>
          </tr>`;
      })
      .join("");
    return;
  }
  if (!entries.length) {
    markCountReportSortHeaders();
    els.countReportBody.innerHTML = `<tr><td colspan="10" class="empty-cell">No items counted yet.</td></tr>`;
    return;
  }
  markCountReportSortHeaders();
  const displayEntries = countInputLogDisplayRows(entries);
  const inputRows = state.countReportSort?.key ? sortCountReportRows(displayEntries, "input") : displayEntries.slice().reverse();
  els.countReportBody.innerHTML = inputRows
    .map((entry) => {
      const variance = Number(entry.inputLogVariance ?? (Number(entry.countedQty || 0) - Number(entry.originalQty || 0)));
      const varianceClass = variance > 0 ? "variance-up" : variance < 0 ? "variance-down" : "variance-flat";
      const varianceLabel = variance > 0 ? `+${number.format(variance)}` : number.format(variance);
      return `
        <tr class="${variance !== 0 ? "count-report-row-diff" : "count-report-row-pass"}">
          <td>${escapeHtml(entry.code || "-")}</td>
          <td>${escapeHtml(entry.plu || "-")}</td>
          <td>${escapeHtml(entry.product || "-")}</td>
          <td>${escapeHtml(entry.vendor || "-")}</td>
          <td>${escapeHtml(entry.category || "-")}</td>
          <td class="num">${number.format(entry.inputLogQtyBefore ?? entry.originalQty ?? 0)}</td>
          <td class="num">${number.format(entry.inputLogQtyAfter ?? entry.countedQty ?? 0)}</td>
          <td class="num ${varianceClass}">${varianceLabel}</td>
          <td>${escapeHtml(entry.mode || "set")}</td>
          <td>${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</td>
        </tr>`;
    })
    .join("");
}

async function openSessionHistoryModal() {
  if (els.sessionHistoryModal) els.sessionHistoryModal.hidden = false;
  if (els.countSessionBody) els.countSessionBody.innerHTML = `<tr><td colspan="13" class="empty-cell">Loading report history...</td></tr>`;
  if (ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
    await restoreSharedCountSessionsOnlyFromSupabase({ silent: true, history: true });
  }
  // Populate vendor filter
  const vendors = [...new Set(state.countSessions.map((s) => s.vendor || s.department || "").filter(Boolean))];
  if (els.sessionHistoryVendorFilter) {
    const cur = els.sessionHistoryVendorFilter.value;
    els.sessionHistoryVendorFilter.innerHTML = `<option value="">All vendors</option>` +
      vendors.map((v) => `<option value="${escapeHtml(v)}" ${v === cur ? "selected" : ""}>${escapeHtml(v)}</option>`).join("");
  }
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

function countSessionDisplayKey(session) {
  if (!session) return "";
  return [
    cleanCell(session.date),
    cleanCell(session.vendor || session.department || "All"),
    cleanCell(session.category || "All"),
    cleanCell(session.status || "All"),
    cleanCell(session.startedAt),
  ].join("|");
}

function remoteActiveCountSessionIsFresh(session) {
  if (!session?.remoteActive) return true;
  const stamp = Date.parse(session.updatedAt || session.submittedAt || session.startedAt || "");
  return Number.isFinite(stamp) && Date.now() - stamp <= COUNT_REMOTE_ACTIVE_HISTORY_TTL_MS;
}

function dedupeCountSessionsForDisplay(sessions = []) {
  return canonicalCountSessions(sessions);
}

function renderCountSessionRows() {
  if (!els.countSessionBody) return;
  const employeeMode = isUserRole();
  const vendorFilter = els.sessionHistoryVendorFilter?.value || "";
  const periodFilter = els.sessionHistoryPeriodFilter?.value || "";

  // FIX-B/C: Always include activeCountSession in the history display, even if it wasn't
  // persisted into countSessions[] yet (e.g. loaded from localStorage on an older build).
  const activeSess = state.activeCountSession;
  const sessionsWithActive = activeSess && !state.countSessions.some((s) => s.id === activeSess.id)
    ? [{ ...activeSess, isActiveLive: true }, ...state.countSessions]
    : state.countSessions;

  const filtered = dedupeCountSessionsForDisplay(sessionsWithActive)
    .slice()
    .sort((a, b) => String(b.updatedAt || "").localeCompare(String(a.updatedAt || "")))
    .filter((s) => {
      if (countSessionIsDeleted(s.id)) return false;
      if (!remoteActiveCountSessionIsFresh(s)) return false;
      // Show all sessions that have been started, even if 0 entries.
      if (!s.startedAt && !(s.entries || []).length && !s.savedAt && !s.submittedAt) return false;
      // FIX-C: In employee mode, show saved/paused/submitted sessions — they all represent
      // completed work. Only hide sessions that are still actively being counted with no
      // entries yet (i.e. just opened, nothing scanned, not yet saved).
      if (employeeMode && !s.submittedAt && !s.savedAt && !(s.entries || []).length) return false;
      if (vendorFilter && (s.vendor || s.department || "") !== vendorFilter) return false;
      if (!sessionMatchesPeriodFilter(s, periodFilter)) return false;
      return true;
    });

  if (!filtered.length) {
    els.countSessionBody.innerHTML = `<tr><td colspan="13" class="empty-cell">${state.countSessions.length ? "No sessions match filters." : "No physical count sessions saved yet."}</td></tr>`;
    return;
  }
  els.countSessionBody.innerHTML = filtered
    .map((session) => `
      <tr>
        <td>${escapeHtml(session.date || "-")}</td>
        <td>${escapeHtml(session.vendor || session.department || "All")}</td>
        <td>${escapeHtml(session.category || "All")}</td>
        <td>${escapeHtml(session.status || "All")}</td>
        <td>${escapeHtml(session.searchFilter || "-")}</td>
        <td>${escapeHtml(new Date(session.startedAt).toLocaleString())}</td>
        <td class="num">${number.format((session.entries || []).length)}</td>
        <td>${escapeHtml(new Date(session.updatedAt || session.startedAt).toLocaleString())}${session.remoteActive ? ` <span class="muted">(Live)</span>` : ""}</td>
        <td>${countSessionStatusLabel(session)}</td>
        <td><button type="button" class="secondary-button count-inline-report-button" data-count-report="${escapeHtml(session.id)}">View</button></td>
        <td>${session.submittedAt ? `<button type="button" class="secondary-button final-report-btn" data-final-report="${escapeHtml(session.id)}">Final Report</button>` : `<span class="muted">Not submitted</span>`}</td>
        <td>${!employeeMode && session.preCountSnapshot ? (session.restoredAt ? `<span class="muted">Restored</span>` : `<button type="button" class="restore-count-btn" data-restore-session="${escapeHtml(session.id)}">Restore</button>`) : ""}</td>
        <td>${countSessionCanDelete(session) ? `<button type="button" class="delete-session-btn" data-delete-session="${escapeHtml(session.id)}">Delete</button>` : ""}</td>
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
  const entries = filterUndoneCountEntries(state.activeCountSession?.entries || [], state.activeCountSession);
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
    renderCountReviewPanel();
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
  renderCountReviewPanel();
}

// ===========================================================================
// V2.0 TWO-COLUMN LIVE REVIEW (Pass 2)
// Read-only right-hand list of the in-scope items for the active count.
//  - Unscanned items show RED (NULL / still needs scanning).
//  - Scanned items gray out, drop to the bottom, and show PASS (green, qty matches)
//    or DIFF (orange, qty mismatch).
//  - The "Remaining" filter literally diminishes the list as you scan.
//  - Sortable headers cycle asc -> desc -> default (Excel-style).
// This panel NEVER mutates count state — it only reads candidates + entries.
// ===========================================================================
function computeCountReviewRows(session) {
  const candidates = currentCountSessionCandidates(session);
  const entriesByCode = new Map();
  (session.entries || []).forEach((e) => {
    const key = codeKey(e.code);
    if (!key) return;
    if (!entriesByCode.has(key)) entriesByCode.set(key, []);
    entriesByCode.get(key).push(e);
  });
  return candidates.map((item) => {
    const list = entriesByCode.get(codeKey(item.code)) || [];
    const effective = effectiveCountFinalEntry(list);
    const entry = effective.entry;
    const before = Number(item.stock || 0);
    const scanned = !!entry;
    const after = scanned ? Number(effective.countedQty || 0) : null;
    const diff = scanned ? after - before : null;
    const status = !scanned ? "null" : diff === 0 ? "pass" : "diff";
    return {
      code: item.code,
      plu: item.plu || item.itemNumber || item.code || "",
      product: item.product || "",
      before, after, diff, scanned, status,
    };
  });
}

function renderCountReviewPanel() {
  if (!els.countReviewBody) return;
  const session = state.activeCountSession;
  if (!session || !state._countSessionOpen) {
    els.countReviewBody.innerHTML = `<tr><td colspan="6" class="empty-cell">Start or continue a count to see the live review list.</td></tr>`;
    if (els.countReviewBadges) els.countReviewBadges.innerHTML = "";
    return;
  }
  if (session._instantStartLoading) {
    els.countReviewBody.innerHTML = `<tr><td colspan="6" class="empty-cell">Opening count now... loading the scope list in the background.</td></tr>`;
    if (els.countReviewBadges) els.countReviewBadges.innerHTML = `<span class="review-badge">Loading</span>`;
    return;
  }
  const rows = computeCountReviewRows(session);

  const nullCount = rows.filter((r) => r.status === "null").length;
  const diffCount = rows.filter((r) => r.status === "diff").length;
  const passCount = rows.filter((r) => r.status === "pass").length;
  if (els.countReviewBadges) {
    els.countReviewBadges.innerHTML =
      `<span class="review-badge review-badge--null">Needs scan ${number.format(nullCount)}</span>` +
      `<span class="review-badge review-badge--diff">Qty diff ${number.format(diffCount)}</span>` +
      `<span class="review-badge review-badge--pass">PASS ${number.format(passCount)}</span>`;
  }

  const filter = state._countReviewFilter || "remaining";
  let view = rows.filter((r) => {
    if (filter === "remaining" || filter === "null") return r.status === "null";
    if (filter === "qtydiff") return r.status === "diff";
    if (filter === "pass") return r.status === "pass";
    return true; // all
  });

  const sort = state._countReviewSort || null;
  const cmpDefault = (a, b) => {
    // Default order: NULL first, then biggest |diff| desc, then PASS last.
    const rank = (r) => (r.status === "null" ? 0 : r.status === "diff" ? 1 : 2);
    if (rank(a) !== rank(b)) return rank(a) - rank(b);
    if (a.status === "diff" && b.status === "diff") return Math.abs(b.diff) - Math.abs(a.diff);
    return String(a.product).localeCompare(String(b.product));
  };
  if (sort && sort.dir) {
    const dir = sort.dir === "desc" ? -1 : 1;
    const val = (r) => {
      switch (sort.key) {
        case "plu": return String(r.plu).toLowerCase();
        case "item": return String(r.product).toLowerCase();
        case "before": return r.before;
        case "after": return r.after == null ? -Infinity : r.after;
        case "diff": return r.diff == null ? -Infinity : r.diff;
        case "status": return r.status;
        default: return 0;
      }
    };
    view = view.slice().sort((a, b) => {
      const va = val(a), vb = val(b);
      if (va < vb) return -1 * dir;
      if (va > vb) return 1 * dir;
      return 0;
    });
  } else {
    view = view.slice().sort(cmpDefault);
  }

  if (!view.length) {
    const msg = filter === "remaining" || filter === "null"
      ? "Nothing left to scan in this scope."
      : "No items match this filter.";
    els.countReviewBody.innerHTML = `<tr><td colspan="6" class="empty-cell">${msg}</td></tr>`;
  } else {
    els.countReviewBody.innerHTML = view.map((r) => {
      const cls = r.status === "null" ? "review-row--null"
        : r.status === "diff" ? "review-row--diff" : "review-row--pass";
      const scannedCls = r.scanned ? " review-row--scanned" : "";
      const afterTxt = r.after == null ? "—" : number.format(r.after);
      const diffTxt = r.diff == null ? "—" : (r.diff > 0 ? `+${number.format(r.diff)}` : number.format(r.diff));
      const statusTxt = r.status === "null" ? "NULL" : r.status === "diff" ? "DIFF" : "PASS";
      return `<tr class="${cls}${scannedCls}">
        <td>${escapeHtml(String(r.plu))}</td>
        <td>${escapeHtml(r.product)}</td>
        <td class="num">${number.format(r.before)}</td>
        <td class="num">${afterTxt}</td>
        <td class="num">${diffTxt}</td>
        <td><span class="review-status review-status--${r.status}">${statusTxt}</span></td>
      </tr>`;
    }).join("");
  }

  if (els.countReviewFilters) {
    els.countReviewFilters.querySelectorAll("[data-review-filter]").forEach((chip) => {
      chip.classList.toggle("is-active", chip.dataset.reviewFilter === filter);
    });
  }
  if (els.countReviewTable) {
    els.countReviewTable.querySelectorAll("th[data-review-sort]").forEach((th) => {
      th.classList.remove("sort-asc", "sort-desc");
      if (sort && sort.key === th.dataset.reviewSort && sort.dir) {
        th.classList.add(sort.dir === "desc" ? "sort-desc" : "sort-asc");
      }
    });
  }
}

function cycleCountReviewSort(key) {
  const cur = state._countReviewSort;
  if (!cur || cur.key !== key) state._countReviewSort = { key, dir: "asc" };
  else if (cur.dir === "asc") state._countReviewSort = { key, dir: "desc" };
  else state._countReviewSort = null; // back to default order
  renderCountReviewPanel();
}

function currentAuditUser() {
  return cleanCell(state.currentUser?.name || "") || "System";
}

function updateCountSummaryStrip() {
  if (!state.activeCountSession || !els.countSummaryStrip) return;
  const entries = state.activeCountSession.entries || [];
  const date = state.activeCountSession.date || "-";
  const vendor = state.activeCountSession.vendor || "All vendors";
  const category = state.activeCountSession.category || "All categories";
  const refine = state.activeCountSession.searchFilter || "No extra filter";
  els.countSummaryStrip.innerHTML = `
    <span><b>${escapeHtml(date)}</b></span>
    <span><b>${escapeHtml(vendor)}</b></span>
    <span><b>${escapeHtml(category)}</b></span>
    <span><b>Filter</b> ${escapeHtml(refine)}</span>
    <span><b>${number.format(entries.length)}</b> items counted</span>`;
}

function setCountSyncLoading(loading, text = "Loading latest count...") {
  const wasLoading = state._countSyncLoading;
  if (loading && !wasLoading) {
    state._countSessionLoadStartedAt = performanceNow();
    console.info("[AppPerf] count session load start");
  }
  state._countSyncLoading = !!loading;
  if (!loading && wasLoading && state._countSessionLoadStartedAt != null) {
    logPerformanceEnd("count session load", state._countSessionLoadStartedAt);
    state._countSessionLoadStartedAt = null;
  }
  if (!els.countSyncStatus) return;
  els.countSyncStatus.hidden = !loading;
  els.countSyncStatus.textContent = text;
}

async function refreshLatestCountSessions(options = {}) {
  // PERF-E: Only show the loading/blank state if there's genuinely nothing to display yet.
  const hasLocalData = state.countSessions.length > 0 || state.activeCountSession != null;
  if (!hasLocalData) {
    setCountSyncLoading(true);
    renderCountsWorkspace({ loading: true });
  }
  try {
    if (ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
      // Pass render:false — we will call renderCountsWorkspace ourselves in the finally block,
      // avoiding a double render when applySharedCountSessionRows also tries to render.
      const pulled = await restoreSharedCountSessionsOnlyFromSupabase({ silent: true, history: options.history === true, render: false });
      if (!pulled) {
        state._countSyncUnavailable = true;
        persistCountSessions({ scheduleSync: false });
      }
    }
  } finally {
    setCountSyncLoading(false);
    // V2.0 FIX: do NOT force the workspace open here. Whether it is open is owned by
    // _countSessionOpen (explicit user action). Re-opening on a background sync was a
    // cause of the count popping up over the Inventory landing page.
    if (activeTabName() === "counts") renderCountsWorkspace();
  }
}

async function openLatestCountOrSetup() {
  // PERF-D: Don't await the remote fetch before deciding what to show.
  // On Kindle this turned a tap on "Start / Continue Count" into a ~30s blank wait.
  // Instead: use local state immediately. If there's already an active session, open it.
  // If not, open the setup modal right away. The background sync will update history rows
  // once it completes, but the user is never blocked waiting for it.
  if (state.activeCountSession) {
    state._countSessionOpen = true;
    renderCountsWorkspace();
    // Still sync in background to pick up entries from other devices.
    if (ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
      const delay = scanPerformanceProfile.tier === "low" ? 2000 : 600;
      setTimeout(() => refreshLatestCountSessions(), delay);
    }
    return;
  }
  // No local active session — check remote before opening setup, but with a timeout.
  // On low-tier devices cap the wait so the setup modal opens within 4s regardless.
  if (ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
    const timeout = scanPerformanceProfile.tier === "low" ? 4000
      : scanPerformanceProfile.tier === "medium" ? 2500 : 1500;
    await Promise.race([
      refreshLatestCountSessions(),
      new Promise((resolve) => setTimeout(resolve, timeout)),
    ]);
  }
  if (state.activeCountSession) {
    state._countSessionOpen = true;
    renderCountsWorkspace();
    return;
  }
  openCountSetupModal();
}

function renderCountsWorkspace(options = {}) {
  if (options.loading || state._countSyncLoading) {
    // Do not blank/close a count workspace that is ALREADY open during background sync.
    // But never force it open here — _countSessionOpen is owned by explicit user actions.
    const keepOpen = state._countSessionOpen && !!state.activeCountSession;
    if (!keepOpen && els.countSessionModal) els.countSessionModal.hidden = true;
    if (els.countWorkspaceEmpty) els.countWorkspaceEmpty.hidden = true;
    if (els.countSummaryStrip) els.countSummaryStrip.hidden = true;
    if (!keepOpen) return;
  }
  // PERF-C: Only rebuild setup dropdowns when the setup modal is actually open or explicitly requested.
  // Previously this ran on every poll-triggered render, re-querying all vendors/categories
  // from the full inventory on every 2s tick — very expensive on Kindle.
  if (options.populateSetup || (els.countSetupModal && !els.countSetupModal.hidden)) {
    populateCountSetupOptions();
  }
  if (state.activeCountSession) {
    state.activeCountSession = normalizeCountSession(
      state.activeCountSession,
      countSessionHasUnsyncedEdits(state.activeCountSession) ? "pending" : "synced",
    );
  }
  const active = state.activeCountSession;
  if (els.countSummaryStrip) els.countSummaryStrip.hidden = false;
  if (els.countSessionModal) els.countSessionModal.hidden = !active || !state._countSessionOpen;
  els.countWorkspaceEmpty.hidden = false;
  if (els.countLaunchCard) els.countLaunchCard.hidden = false;
  if (els.countLaunchTitle) els.countLaunchTitle.textContent = active ? "Continue Count" : "Start New Count";
  // V2.0 split cards: Start New is always available; Continue reflects whether an
  // active count exists on this device.
  if (els.countContinueCard) {
    els.countContinueCard.classList.toggle("count-card-disabled", !active);
    els.countContinueCard.setAttribute("aria-disabled", active ? "false" : "true");
  }
  if (els.countContinueCardTitle) {
    els.countContinueCardTitle.textContent = active ? "Continue Count" : "No active count";
  }
  if (els.countContinueCardDesc) {
    els.countContinueCardDesc.textContent = active
      ? `Resume: ${countSessionLabel(active)}`
      : "Nothing active on this device. Use Report History to continue a saved count.";
  }
  if (els.countLaunchDescription) {
    els.countLaunchDescription.textContent = active
      ? "Resume the latest active synced count."
      : state._countSyncUnavailable
        ? "Sync is unavailable. Start locally; entries stay on this device and retry automatically."
        : "No active count exists. Choose the scope and begin a new physical count.";
  }
  if (els.countLaunchState) {
    els.countLaunchState.textContent = active
      ? "Active count found - Continue Count"
      : state._countSyncUnavailable
        ? "Sync unavailable - Start local count with warning"
        : "No active count - Start New Count";
  }
  if (els.closeCountSessionButton) els.closeCountSessionButton.hidden = true;
  if (els.countReviewButton) els.countReviewButton.hidden = true;
  if (active) {
    els.activeCountTitle.textContent = countSessionLabel(active);
    const vendorLabel = active.vendor || active.department || "All vendors";
    // PATCH E: Show active count criteria clearly, including whether disabled items are excluded.
    const statusLabel = active.status
      ? escapeHtml(active.status)
      : "<strong>Active items only</strong> (disabled &amp; discontinued excluded)";
    els.activeCountMeta.innerHTML = `
      <span>${escapeHtml(active.date || "-")}</span>
      <span>${escapeHtml(vendorLabel)}</span>
      <span>${escapeHtml(active.category || "All categories")}</span>
      <span>Status: ${statusLabel}</span>
      <span>${number.format((active.entries || []).length)} entries</span>
      <span>${active.allowOutOfScope && !isUserRole() ? "Manager scope override enabled" : "Scope locked"}</span>`;
    renderActiveCountSyncStatus(active);
  } else if (els.activeCountMeta) {
    els.activeCountMeta.innerHTML = "";
    renderActiveCountSyncStatus(null);
  }
  renderCountQuantity();
  renderSelectedCountItem();
  // Only rebuild entry rows when the count session is actually open.
  // These are no-ops when the modal is hidden but add measurable parse cost on Kindle.
  if (active && state._countSessionOpen) {
    renderCountEntryRows();
  }
  renderCountReviewPanel();
  renderCountSessionRows();
  if (active && state.countStage === "search") focusCountSearch();
}

function renderActiveTab() {
  hideHoverTooltip();
  const activeTab = document.querySelector(".tab-button.active")?.dataset.tab || "dashboard";
  if (activeTab === "scanmode") {
    if (els.inventoryQuickTools) els.inventoryQuickTools.hidden = true;
    return;
  }
  if (els.inventoryQuickTools) {
    els.inventoryQuickTools.hidden = !["inventory", "ordering"].includes(activeTab);
  }
  renderSharedQuickTools(activeTab);
  if (activeTab === "dashboard") {
    if (state._dashboardDetailsReady) {
      renderTrend();
      renderBars();
    } else {
      scheduleDashboardDetails();
    }
  } else if (activeTab === "inventory") {
    renderInventory();
  } else if (activeTab === "newitems") {
    renderNewItems();
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
  } else if (activeTab === "parents") {
    renderParents();
  }
  repairMojibakeText(document.body);
  refreshDetailDrawer();
}

function scheduleDashboardDetails() {
  if (state._dashboardDetailsReady || state._dashboardDetailScheduled || activeTabName() !== "dashboard") return;
  state._dashboardDetailScheduled = true;
  const token = ++state._dashboardDetailToken;
  const run = () => {
    state._dashboardDetailScheduled = false;
    if (token !== state._dashboardDetailToken || activeTabName() !== "dashboard") return;
    setAppLoadStatus("Building product index");
    renderTrend();
    renderBars();
    state._dashboardDetailsReady = true;
    clearAppLoadStatus("Building product index");
  };
  if ("requestIdleCallback" in window) window.requestIdleCallback(run, { timeout: 1400 });
  else setTimeout(run, 180);
}

function applyRoleRestrictions(force = false) {
  const userMode = isUserRole();
  if (!force && state._lastRoleRestrictionMode === userMode && state._roleRestrictionApplied) return;
  state._lastRoleRestrictionMode = userMode;
  state._roleRestrictionApplied = true;
  // Hide admin-only elements for basic users
  const adminOnly = [
    "#downloadInventoryCsvBtn", ".download-inventory-csv", "[data-admin-only]",
    "#exportPoExcel", "#exportPoPdf", "#downloadOrder",
    "#exportAdjustPdfButton", "#exportAdjustExcelButton",
    "#clearAdjustLogButton", ".arrange-columns-btn", "#arrangeColumnsButton",
    ".column-picker", "#downloadInventory", "#createPoShortcut"
  ];
  adminOnly.forEach((sel) => {
    document.querySelectorAll(sel).forEach((el) => { el.style.display = userMode ? "none" : ""; });
  });
  const metricsZone = document.querySelector("#metricsHoverZone");
  if (metricsZone) metricsZone.style.display = userMode ? "none" : "";
  // Hide cost columns for user role
  if (userMode) {
    document.querySelectorAll("#inventory th[data-col='unitCost'], #inventoryBody td[data-col='unitCost'], #inventory th[data-col='inventoryCost'], #inventoryBody td[data-col='inventoryCost']").forEach((el) => { el.style.display = "none"; });
    document.querySelectorAll(".order-table th[data-order-sort='unitCost'], .order-table td[data-col='unitCost'], .order-table th[data-order-sort='totalCost'], .order-table td[data-col='totalCost']").forEach((el) => { el.style.display = "none"; });
    // Hide metrics strip cost info
    document.querySelectorAll("#costSold, #grossProfit").forEach((el) => {
      const article = el?.closest("article");
      if (article) article.style.display = "none";
    });
    // Hide summary strip cost items
    document.querySelectorAll(".inventory-summary .cost-item, [data-cost-summary]").forEach((el) => { el.style.display = "none"; });
    document.querySelectorAll("[data-units-summary]").forEach((el) => { el.style.display = "none"; });
    // Only show products and inventory tabs
    document.querySelectorAll(".tab-button").forEach((btn) => {
      const tab = btn.dataset.tab;
      if (!tab) {
        btn.style.display = "";
        return;
      }
      if (!["inventory","counts","scanmode"].includes(tab)) btn.style.display = "none";
    });
  } else {
    document.querySelectorAll("#inventory th[data-col='unitCost'], #inventoryBody td[data-col='unitCost'], #inventory th[data-col='inventoryCost'], #inventoryBody td[data-col='inventoryCost']").forEach((el) => { el.style.display = ""; });
    document.querySelectorAll(".order-table th[data-order-sort='unitCost'], .order-table td[data-col='unitCost'], .order-table th[data-order-sort='totalCost'], .order-table td[data-col='totalCost']").forEach((el) => { el.style.display = ""; });
    document.querySelectorAll("#costSold, #grossProfit").forEach((el) => {
      const article = el?.closest("article");
      if (article) article.style.display = "";
    });
    document.querySelectorAll(".inventory-summary .cost-item, [data-cost-summary]").forEach((el) => { el.style.display = ""; });
    document.querySelectorAll("[data-units-summary]").forEach((el) => { el.style.display = ""; });
    document.querySelectorAll(".tab-button").forEach((btn) => { btn.style.display = ""; });
  }
}

const ACTIVE_TAB_RENDER_GAP_MS = 72;

function queueActiveTabRender() {
  state._activeTabRenderRequested = true;
  if (state._activeTabRenderHandle) return;

  const elapsed = Date.now() - (state._activeTabRenderLastAt || 0);
  const wait = Math.max(0, ACTIVE_TAB_RENDER_GAP_MS - elapsed);
  state._activeTabRenderHandle = setTimeout(() => {
    state._activeTabRenderHandle = 0;
    if (!state._activeTabRenderRequested) return;
    if (state._activeTabRenderRunning) {
      queueActiveTabRender();
      return;
    }

    state._activeTabRenderRequested = false;
    state._activeTabRenderRunning = true;
    try {
      renderActiveTab();
      state._activeTabRenderLastAt = Date.now();
    } finally {
      state._activeTabRenderRunning = false;
      if (state._activeTabRenderRequested) queueActiveTabRender();
    }
  }, wait);
}

function render() {
  if (els.countSessionModal && !els.countSessionModal.hidden) return;
  const skuRows = buildSkuRows();
  state.filteredSkus = sortSkuRows(skuRows);
  state._filteredSkuIndex = new Map(state.filteredSkus.map((item) => [codeKey(item.code), item]));
  const metricRows = els.inventoryStateFilter?.value ? buildSkuRows({ ignoreStateFilter: true }) : skuRows;
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
  repairMojibakeText(document.body);
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
  const safetyDays = Math.max(0, toNumber(els.safetyDays.value) || DEFAULT_ORDER_SAFETY_DAYS);
  const daysOfInventory = Math.max(0, toNumber(els.daysOfInventory?.value) || DEFAULT_ORDER_DAYS_OF_INVENTORY);

  // Cache busts on date range OR any ordering parameter change
  const cacheKey = `${start}|${end}|${leadDays}|${safetyDays}|${daysOfInventory}|ordering-math-v2`;
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
    if (query && !matchesSearchQuery(sku, query)) return false;
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
    merged._haystack = buildSearchHaystack([
      merged.code, merged.product, merged.department, merged.category, merged.vendor,
      merged.color, merged.size, merged.parent, merged.subType,
      merged.plu, merged.itemNumber, merged.state, merged.addDate, merged.caseSize,
    ]);

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
    const velocity = Math.max(0, sku.units / Math.max(dayCount, 1));
    const override = reorderOverrideForCode(sku.code);
    const isOverridden = override.min != null || override.max != null;
    const rule = effectiveItemRule(sku);
    const useRuleTargets = rule.mode === "custom" || rule.categoryDefault || !!rule.vendorRule;
    const targetSafetyDays = useRuleTargets ? rule.safetyDays : safetyDays;
    const targetDaysOfInventory = useRuleTargets ? rule.daysOfInventory : daysOfInventory;
    const dynamic = orderingTargetsForItem({
      item: sku,
      velocity,
      safetyDays: targetSafetyDays,
      daysOfInventory: targetDaysOfInventory,
    });
    const dynamicMaxWithDoi = dynamic.max;
    const manualMin = override.min ?? dynamic.min;
    const manualMax = override.max ?? dynamicMaxWithDoi;
    const recommendedOrder = recommendedOrderQty({
      stock: sku.stock,
      min: manualMin,
      max: manualMax,
      caseSize: sku.caseSize,
    });
    const caseOrder = calcCaseOrder(recommendedOrder, sku.caseSize);
    const daysSupply = velocity > 0 ? sku.stock / velocity : Infinity;
    const margin = sku.sales > 0 ? sku.profit / sku.sales : 0;
    const inventoryCost = sku.stock * sku.unitCost;
    const full = {
      ...sku, velocity, daysSupply, recommendedOrder, caseOrder, margin,
      reorderMin: manualMin, reorderMax: manualMax, inventoryCost,
      isOverridden, dynamicMin: dynamic.min, dynamicMax: dynamicMaxWithDoi,
      effectiveSafetyDays: dynamic.safetyDays,
      effectiveDaysOfInventory: dynamic.daysOfInventory,
      orderingRuleMode: rule.mode,
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
  // Drill one level deeper: if grouped by dept->show categories; by category->show items; by vendor->show items
  const subKey = groupKey === "department" ? "category" : "product";
  const subGroups = groupKey === "department"
    ? groupBy(rows, "category", metric).slice(0, 8)
    : groupBy(rows, "product", metric).slice(0, 10);

  target.innerHTML = `
    <div class="drill-heading">
      <strong>${groupValue ? escapeHtml(groupValue) : `Click a ${groupKey}`}</strong>
      <span>${groupKey === "department" ? "-> categories -> items" : "-> top items"}</span>
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
  const fmtChg = (v) => v == null ? "-" : (v >= 0 ? `▲ +${v.toFixed(1)}%` : `▼ ${v.toFixed(1)}%`);
  const chgCls = (v) => v == null ? "" : v >= 0 ? "compare-up" : "compare-down";

  target.innerHTML = `
    <p class="compare-period-label">Current ${periodLabel} vs prior same-length period  -  grouped by ${groupKey}</p>
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
  if (!panel) return;
  const orderedColumns = state.orderColumnOrder
    .map((key) => orderColumns.find((column) => column.key === key))
    .filter(Boolean);
  panel.innerHTML = `<div class="column-picker-grid order-column-grid">${orderedColumns.map((c, index) => `
    <div class="column-choice">
      <label>
        <input type="checkbox" data-order-col="${c.key}" ${state.orderVisibleColumns[c.key] ? "checked" : ""} />
        <span>${c.label}</span>
      </label>
    </div>`).join("")}</div>`;
  panel.querySelectorAll("[data-order-col]").forEach((cb) => {
    cb.addEventListener("change", () => {
      state.orderVisibleColumns[cb.dataset.orderCol] = cb.checked;
      safeLocalStorageSet("posOrderColumns:v1", JSON.stringify(state.orderVisibleColumns));
      renderOrders();
      document.querySelector("#orderColumnPicker")?.setAttribute("open", "open");
    });
  });
}

function moveOrderColumn(from, to) {
  if (!from || !to || from === to) return;
  const order = state.orderColumnOrder.filter((key) => key !== from);
  order.splice(order.indexOf(to), 0, from);
  state.orderColumnOrder = order;
  state.orderArrangeSource = "";
  safeLocalStorageSet("posOrderColumnOrder:v1", JSON.stringify(state.orderColumnOrder));
  renderOrders();
}

function moveOrderColumnRelative(key, direction = 1) {
  const order = [...(state.orderColumnOrder || [])];
  const index = order.indexOf(key);
  if (index < 0) return;
  const nextIndex = Math.max(0, Math.min(order.length - 1, index + direction));
  if (nextIndex === index) return;
  const [column] = order.splice(index, 1);
  order.splice(nextIndex, 0, column);
  state.orderColumnOrder = order;
  safeLocalStorageSet("posOrderColumnOrder:v1", JSON.stringify(state.orderColumnOrder));
  renderOrders();
}

function renderOrders() {
  renderSharedQuickTools("ordering");
  pruneOrderSubmissionDrafts();
  const arrangeButton = document.querySelector("#orderArrangeColumnsButton");
  if (arrangeButton) {
    arrangeButton.classList.toggle("active", state.orderArrangeColumns);
    arrangeButton.textContent = state.orderArrangeColumns ? "Lock columns" : "Arrange columns";
  }
  const manualVendorSelected = cleanCell(els.vendorFilter?.value || "").toUpperCase() !== "ALL";
  const orderRows = buildInventoryRows()
    .map(applyOrderOverride)
    .filter((item) => {
      if (item.isOrderable === false) return false;
      if (isPendingOrder(item.code)) return false;
      if (item.recommendedOrder <= 0 && !item.qtyNeeded) return false;
      const vendorFilter = getOrderVendorFilter ? getOrderVendorFilter() : "Active";
      if (state.vendorRules.length && vendorFilter !== "" && !manualVendorSelected) {
        const vendorName = (item.vendor || "").toUpperCase();
        const rule = state.vendorRules.find((r) => r.vendor?.toUpperCase() === vendorName);
        if (vendorFilter === "Active" && rule && rule.status !== "Active") return false;
        if (vendorFilter === "Disabled" && (!rule || rule.status !== "Disabled")) return false;
      }
      return true;
    });
  state._lastRenderedOrderRows = orderRows.map((item) => ({ ...item }));
  state.orderVendorQuickFilter = "";
  const vendorOrderStats = new Map();
  orderRows.forEach((item) => {
    const vendorName = cleanCell(item.vendor).toUpperCase();
    if (!vendorName) return;
    const current = vendorOrderStats.get(vendorName) || { total: 0, count: 0 };
    current.total += orderLineCost(item);
    current.count += 1;
    vendorOrderStats.set(vendorName, current);
  });
  // Show live formula in header
  const safety = toNumber(els.safetyDays.value) || DEFAULT_ORDER_SAFETY_DAYS;
  const doi = toNumber(els.daysOfInventory?.value) || DEFAULT_ORDER_DAYS_OF_INVENTORY;
  const note = document.getElementById("formulaNote");
  if (note) note.textContent = `Min = ceil(SV x ${safety}d)  |  Max = ceil(Min${doi ? ` + (SV x ${doi}d DOI)` : ""})  |  Case rounds order only`;

  // Today's ordering alert
  const today = new Date().toLocaleDateString("en-US", { weekday: "long" }).toLowerCase();
  const pendingVendors = pendingVendorNames();
  const dismissedVendors = new Set((state.orderDismissedVendors || []).map((vendor) => String(vendor || "").toUpperCase()).filter(Boolean));
  const todayVendors = state.vendorRules.filter((r) => {
    if (r.status !== "Active" || !(r.orderDays || []).includes(today)) return false;
    if (pendingVendors.has((r.vendor || "").toUpperCase())) return false;
    if (dismissedVendors.has((r.vendor || "").toUpperCase())) return false;
    const vendorStats = vendorOrderStats.get((r.vendor || "").toUpperCase()) || { total: 0, count: 0 };
    return vendorStats.count > 0 && (!r.minOrder || vendorStats.total >= r.minOrder);
  }).sort((a, b) => compareDisplayValue(a.vendor, b.vendor));
  state._todayOrderVendors = todayVendors.map((rule) => cleanCell(rule.vendor)).filter(Boolean);
  // Write alert to the dedicated banner, not into orderCards (so it persists)
  const banner = document.querySelector("#orderAlertBanner");
  if (banner) {
    let alertHtml = "";
    if (todayVendors.length) {
      alertHtml += `<div class="order-day-alert">
        Order today: <b>Order today:</b>
        ${todayVendors.map((r) => {
          const isPending = state.pendingOrders?.some((po) => po.vendor === r.vendor && !po.cleared);
          return `<span class="order-vendor-chip-wrap"><button type="button" class="order-vendor-chip-btn${isPending?" order-pending-chip":""}" data-vendor-order="${escapeHtml(r.vendor)}">${escapeHtml(r.vendor)}${isPending?" Pending":""}</button><button type="button" class="order-vendor-chip-dismiss" data-dismiss-order-vendor="${escapeHtml(r.vendor)}" title="Dismiss ${escapeHtml(r.vendor)}">×</button></span>`;
        }).join("")}
        <button type="button" id="submitAllPoButton" class="count-submit-btn" style="margin-left:auto">Submit All PO</button>
      </div>`;
    }
    banner.innerHTML = alertHtml;
    // Wire chips
    banner.querySelectorAll(".order-vendor-chip-btn").forEach((btn) => {
      btn.addEventListener("click", () => openVendorAnalysisPanel(btn.dataset.vendorOrder));
    });
    banner.querySelectorAll("[data-dismiss-order-vendor]").forEach((btn) => {
      btn.addEventListener("click", (event) => {
        event.stopPropagation();
        dismissOrderVendor(btn.dataset.dismissOrderVendor || "");
        renderOrders();
      });
    });
    document.querySelector("#submitAllPoButton")?.addEventListener("click", () => submitAllPo());
  }

  const vendorBar = document.querySelector(".ordering-vendor-filter-bar");
  if (vendorBar) {
    vendorBar.querySelector(".order-vendor-actions")?.remove();
    vendorBar.querySelector(".order-vendor-chips")?.remove();
    closeOrderVendorMenu();
    if (state.orderSubmissionVendors?.length) {
      const pillsHtml = `
        <div class="order-vendor-actions">
          <button type="button" class="count-submit-btn order-vendor-submit-all" id="orderingSubmitCurrentButton">Submit PO</button>
          <button type="button" class="count-submit-btn order-vendor-submit-all" id="orderingSubmitAllButton">Submit All PO</button>
        </div>
        <div class="order-vendor-chips">
          <button type="button" class="order-vendor-pill${state.orderSubmissionActiveVendor ? "" : " active"}" data-order-submission-vendor="">All sent</button>
          ${state.orderSubmissionVendors.map((vendor) => `<button type="button" class="order-vendor-pill${state.orderSubmissionActiveVendor.toUpperCase() === vendor.toUpperCase() ? " active" : ""}" data-order-submission-vendor="${escapeHtml(vendor)}">${escapeHtml(vendor)}</button><button type="button" class="order-vendor-pill-remove" data-order-remove-vendor="${escapeHtml(vendor)}" title="Clear ${escapeHtml(vendor)} order">×</button>`).join("")}
        </div>`;
      vendorBar.insertAdjacentHTML("beforeend", pillsHtml);
      vendorBar.querySelectorAll("[data-order-submission-vendor]").forEach((btn) => {
        btn.addEventListener("click", () => {
          const vendorName = btn.dataset.orderSubmissionVendor || "";
          state.orderSubmissionActiveVendor = vendorName;
          renderOrders();
          if (vendorName) openVendorAnalysisPanel(vendorName);
        });
      });
      vendorBar.querySelector("#orderingSubmitCurrentButton")?.addEventListener("click", () => {
        const vendorName = state.orderSubmissionActiveVendor || state.orderSubmissionVendors[0];
        if (!vendorName) return;
        openVendorAnalysisPanel(vendorName);
      });
      vendorBar.querySelector("#orderingSubmitAllButton")?.addEventListener("click", () => {
        if (!state.orderSubmissionVendors?.length) return;
        if (!confirm(`Submit all ${state.orderSubmissionVendors.length} sent vendor PO${state.orderSubmissionVendors.length === 1 ? "" : "s"}?`)) return;
        state.orderSubmissionVendors.forEach((vendor) => submitVendorPo(vendor));
      });
      vendorBar.querySelectorAll("[data-order-remove-vendor]").forEach((btn) => {
        btn.addEventListener("click", () => removeSubmittedVendor(btn.dataset.orderRemoveVendor || ""));
      });
    }
  }
  applyRoleRestrictions(true);
  // Render column picker
  renderOrderColumnPicker();

  const visibleCols = state.orderColumnOrder
    .map((key) => orderColumns.find((column) => column.key === key))
    .filter((column) => column && state.orderVisibleColumns[column.key]);
  const sortKey = state.orderSort?.key || "recommendedOrder";
  const sortDir = state.orderSort?.dir || "desc";
  const widthByColumn = {
    status: "6rem",
    pending: "2rem",
    code: "7.2rem",
    product: "24rem",
    vendor: "6.2rem",
    plu: "5.4rem",
    velocity: "4.2rem",
    units: "4.1rem",
    stock: "4.2rem",
    reorderMin: "3.6rem",
    reorderMax: "3.6rem",
    recommendedOrder: "5.3rem",
    caseOrder: "4.8rem",
    caseSize: "4.6rem",
    unitCost: "4.8rem",
    totalCost: "6.4rem",
  };

  // Sort orders
  const sorted = [...orderRows].sort((a, b) => {
    const av = a[sortKey] ?? 0, bv = b[sortKey] ?? 0;
    if (typeof av === "string") return sortDir === "asc" ? av.localeCompare(bv) : bv.localeCompare(av);
    return sortDir === "asc" ? av - bv : bv - av;
  }).slice(0, 120);

  els.orderCards.innerHTML = "";
  if (!sorted.length) {
    els.orderCards.innerHTML = `<p class="muted">No reorder recommendations. Discontinued, disabled, and vendor-disabled items are excluded.</p>`;
    return;
  }

  const colgroupHtml = visibleCols.map((c) => `<col style="width:${widthByColumn[c.key] || "6rem"}">`).join("");
  const headerHtml = visibleCols.map((c) => {
    const isActive = c.key === sortKey;
    const arrow = isActive ? (sortDir === "asc" ? " ↑" : " ↓") : "";
    return `<th class="order-sortable-th${isActive ? " sort-active" : ""}${state.orderArrangeColumns ? " arrange-column-active" : ""}" data-col="${c.key}" data-order-sort="${c.key}" data-order-col-header="${c.key}" draggable="${state.orderArrangeColumns ? "true" : "false"}"><span class="order-arrange-label">${c.label}${state.orderArrangeColumns ? "" : arrow}</span></th>`;
  }).join("");

  els.orderCards.innerHTML = `
    <div class="table-wrap order-table-wrap" style="overflow-x:auto;max-height:none">
      <table class="order-table order-table--fixed">
        <colgroup><col style="width:2rem">${colgroupHtml}</colgroup>
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
              status:           `<td data-col="status">${stateBadgeHtml(sku)}</td>`,
              pending:          `<td data-col="pending" class="pending-col">${isPend ? `<button type="button" class="pending-clear-btn pending-inline-btn" data-code="${escapeHtml(sku.code)}" title="Click to clear pending"><span class="pending-icon">&#x1F550;</span></button>` : ""}</td>`,
              code:             `<td data-col="code" class="copy-code" data-copy-code="${escapeHtml(sku.code)}" style="color:#2470c4;font-weight:700;cursor:pointer">${escapeHtml(sku.code)}</td>`,
              product:          `<td data-col="product" class="sku-name" title="${escapeHtml(sku.product)}"><span class="sku-name-text">${escapeHtml(sku.product)}</span></td>`,
              vendor:           `<td data-col="vendor">${escapeHtml(sku.vendor || "-")}</td>`,
              plu:              `<td data-col="plu">${escapeHtml(sku.plu || "-")}</td>`,
              velocity:         `<td data-col="velocity" class="num order-velocity-col">${formatVelocity(sku.velocity || 0)}</td>`,
              units:            `<td data-col="units" class="num order-sold-col">${number.format(sku.units || 0)}</td>`,
              stock:            `<td data-col="stock" class="num order-stock-col ${(sku.stock||0) < 0 ? "entry-negative" : ""}">${number.format(sku.stock || 0)}</td>`,
              reorderMin:       `<td data-col="reorderMin" class="num">${number.format(sku.reorderMin || 0)}</td>`,
              reorderMax:       `<td data-col="reorderMax" class="num">${number.format(sku.reorderMax || 0)}</td>`,
              recommendedOrder: `<td data-col="recommendedOrder" class="num order-highlight"><input type="number" class="order-rec-input mini-input" data-code="${escapeHtml(sku.code)}" value="${sku.recommendedOrder || sku.qtyNeeded || 0}" min="0" style="width:3.8rem;text-align:center;font-weight:700" /></td>`,
              caseOrder:        `<td data-col="caseOrder" class="num order-highlight"><b>${number.format(sku.caseOrder || 0)}</b></td>`,
              caseSize:         `<td data-col="caseSize" class="num">${number.format(sku.caseSize || 1)}</td>`,
              unitCost:         `<td data-col="unitCost" class="num">${currency.format(sku.unitCost || 0)}</td>`,
              totalCost:        `<td data-col="totalCost" class="num">${currency.format(orderLineCost(sku))}</td>`,
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
    th.style.cursor = state.orderArrangeColumns ? "grab" : "pointer";
    th.addEventListener("dragstart", (event) => {
      if (!state.orderArrangeColumns) return;
      event.dataTransfer.setData("text/plain", th.dataset.orderColHeader);
      th.classList.add("arrange-column-selected");
    });
    th.addEventListener("dragend", () => {
      th.classList.remove("arrange-column-selected");
    });
    th.addEventListener("dragover", (event) => {
      if (state.orderArrangeColumns) event.preventDefault();
    });
    th.addEventListener("drop", (event) => {
      if (!state.orderArrangeColumns) return;
      event.preventDefault();
      moveOrderColumn(event.dataTransfer.getData("text/plain"), th.dataset.orderColHeader);
    });
    th.addEventListener("click", () => {
      if (state.orderArrangeColumns) return;
      const key = th.dataset.orderSort;
      state.orderSort = { key, dir: state.orderSort?.key === key && state.orderSort?.dir === "asc" ? "desc" : "asc" };
      renderOrders();
    });
  });
  els.orderCards.querySelectorAll("[data-order-shift]").forEach((button) => {
    button.addEventListener("click", (event) => {
      event.stopPropagation();
      moveOrderColumnRelative(button.dataset.orderShift || "", Number(button.dataset.orderShiftDir || 0));
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
  // Wire editable Rec Order inputs - stopPropagation prevents opening detail drawer
  els.orderCards.querySelectorAll(".order-rec-input").forEach((input) => {
    input.addEventListener("click", (e) => e.stopPropagation());
    input.addEventListener("focus", () => input.select?.());
    input.addEventListener("dblclick", () => input.select?.());
    input.addEventListener("input", () => {
      const val = Math.max(0, Math.round(toNumber(input.value) || 0));
      input.value = String(val);
    });
  });
  // Wire per-item clear-pending buttons
  els.orderCards.querySelectorAll(".pending-clear-btn").forEach((btn) => {
    btn.addEventListener("click", (e) => {
      e.stopPropagation();
      const code = btn.dataset.code;
      clearPendingForCode(code, { reason: "manual", backorder: true, render: false });
      renderOrders();
      patchInventoryRow(code);
      showToast("Pending cleared for " + code, 2000, "success");
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
      <td class="sku-name" title="${escapeHtml(sku.product)}"><span class="sku-name-text">${escapeHtml(sku.product)}</span></td>
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
  const productsRenderStartedAt = performanceNow();
  console.info("[AppPerf] products table render start");
  renderSharedQuickTools("inventory");
  const needsIndexBuild = !state._inventoryCache || state._inventoryCacheStamp !== state._dataCacheStamp;
  if (needsIndexBuild) setAppLoadStatus("Building product index");
  const rows = currentInventoryRows();
  const filterKey = JSON.stringify({
    search: els.searchInput?.value || "",
    state: els.inventoryStateFilter?.value || "",
    quick: els.inventoryQuickFilter?.value || "",
    department: els.departmentFilter?.value || "",
    category: els.categoryFilter?.value || "",
    vendor: els.vendorFilter?.value || "",
    color: els.colorFilter?.value || "",
    start: els.startDate?.value || "",
    end: els.endDate?.value || "",
    order: state.columnOrder,
    sort: state.inventorySort,
  });
  if (state._inventoryVisibleFilterKey !== filterKey) {
    state._inventoryVisibleFilterKey = filterKey;
    state._inventoryVisibleLimit = 250;
  }
  state._inventoryFilteredCount = rows.length;
  const renderToken = (state._inventoryRenderToken || 0) + 1;
  state._inventoryRenderToken = renderToken;
  els.inventoryBody.innerHTML = "";
  renderInventorySummary(rows);
  renderInventoryHeader();
  wireInventoryInfiniteScroll();
  if (!rows.length) {
    els.inventoryBody.innerHTML = `<tr><td colspan="${inventoryVisibleColumnKeys().length + 1}" class="empty-cell">Load CSV inventory or the Excel import to view all items.</td></tr>`;
    clearAppLoadStatus("Building product index");
    logPerformanceEnd("products table render", productsRenderStartedAt, { rows: 0 });
    return;
  }

  // Chunked rendering - build DOM in batches via requestAnimationFrame so the
  // browser stays responsive. Tooltip HTML is deferred until mouseover.
  const CHUNK = 125;
  // Keep the DOM small and responsive. More rows are appended in pages as the
  // user approaches the bottom of the Products table.
  const visible = rows.slice(0, state._inventoryVisibleLimit || 250);
  let offset = 0;

  function renderChunk() {
    if (renderToken !== state._inventoryRenderToken) return;
    const fragment = document.createDocumentFragment();
    const end = Math.min(offset + CHUNK, visible.length);
    for (let i = offset; i < end; i++) {
      const item = visible[i];
      fragment.append(buildInventoryRowNode(item));
    }
    els.inventoryBody.append(fragment);
    offset = end;
    if (offset < visible.length) {
      requestAnimationFrame(renderChunk);
    } else {
      if (renderToken !== state._inventoryRenderToken) return;
      els.inventoryBody.querySelectorAll("tr[data-item-code]").forEach((entry) => wireInventoryRowInteractions(entry));
      applyColumnVisibility();
      applyInventoryColumnWidths();
      syncInventoryHeaderOffset();
      clearAppLoadStatus("Building product index");
      logPerformanceEnd("products table render", productsRenderStartedAt, { rows: visible.length, available: rows.length });
    }
  }
  requestAnimationFrame(renderChunk);
}

function syncInventoryHeaderOffset() {
  const root = document.documentElement;
  root.style.setProperty("--inventory-summary-height", "0px");
}

function applyInventoryColumnWidths() {
  let style = document.querySelector("#inventoryColumnWidthStyle");
  if (!style) {
    style = document.createElement("style");
    style.id = "inventoryColumnWidthStyle";
    document.head.append(style);
  }
  const rules = Object.entries(state.columnWidths || {})
    .filter(([, width]) => Number(width) >= 42)
    .map(([key, width]) => {
      const safeKey = CSS.escape(key);
      const px = Math.round(Number(width));
      return `#inventory th[data-col="${safeKey}"],#inventoryBody td[data-col="${safeKey}"]{width:${px}px!important;min-width:${px}px!important;max-width:${px}px!important;}`;
    });
  style.textContent = rules.join("\n");
}

function autoFitInventoryColumn(key) {
  if (!key) return;
  const header = document.querySelector(`#inventory th[data-col="${CSS.escape(key)}"]`);
  const cells = [...document.querySelectorAll(`#inventoryBody td[data-col="${CSS.escape(key)}"]`)].slice(0, 120);
  const sampleText = [header?.textContent || "", ...cells.map((cell) => cell.textContent || "")]
    .reduce((longest, text) => (text.trim().length > longest.length ? text.trim() : longest), "");
  const width = Math.max(48, Math.min(520, Math.ceil(sampleText.length * 7.4) + 34));
  state.columnWidths[key] = width;
  safeLocalStorageSet("posDashboardColumnWidths:v2", JSON.stringify(state.columnWidths));
  applyInventoryColumnWidths();
}

function startInventoryColumnResize(event, key) {
  event.preventDefault();
  event.stopPropagation();
  const header = event.currentTarget.closest("th");
  if (!header || !key) return;
  const pointer = event.touches?.[0] || event;
  const startX = pointer.clientX;
  const startWidth = header.getBoundingClientRect().width;
  const onMove = (moveEvent) => {
    moveEvent.preventDefault?.();
    const movePointer = moveEvent.touches?.[0] || moveEvent;
    const nextWidth = Math.max(42, Math.min(620, Math.round(startWidth + movePointer.clientX - startX)));
    state.columnWidths[key] = nextWidth;
    applyInventoryColumnWidths();
  };
  const onUp = () => {
    safeLocalStorageSet("posDashboardColumnWidths:v2", JSON.stringify(state.columnWidths));
    document.removeEventListener("mousemove", onMove);
    document.removeEventListener("mouseup", onUp);
    document.removeEventListener("pointermove", onMove);
    document.removeEventListener("pointerup", onUp);
    document.removeEventListener("touchmove", onMove);
    document.removeEventListener("touchend", onUp);
  };
  document.addEventListener("mousemove", onMove);
  document.addEventListener("mouseup", onUp, { once: true });
  document.addEventListener("pointermove", onMove);
  document.addEventListener("pointerup", onUp, { once: true });
  document.addEventListener("touchmove", onMove, { passive: false });
  document.addEventListener("touchend", onUp, { once: true });
}

function wireInventoryInfiniteScroll() {
  const scroller = els.inventoryBody?.closest(".table-wrap") || els.inventoryBody?.parentElement;
  if (!scroller || scroller.dataset.infiniteInventoryWired === "true") return;
  scroller.dataset.infiniteInventoryWired = "true";
  scroller.addEventListener("scroll", () => {
    const totalRows = state._inventoryFilteredCount || 0;
    if ((state._inventoryVisibleLimit || 250) >= totalRows) return;
    const remaining = scroller.scrollHeight - scroller.scrollTop - scroller.clientHeight;
    if (remaining > 220) return;
    showInventoryLoadingMore();
    state._inventoryVisibleLimit = Math.min(totalRows, (state._inventoryVisibleLimit || 250) + 250);
    renderInventory();
  }, { passive: true });
}

function showInventoryLoadingMore() {
  let note = document.querySelector("#inventoryLoadingMore");
  if (!note) {
    note = document.createElement("div");
    note.id = "inventoryLoadingMore";
    note.className = "inventory-loading-more";
    note.textContent = "Loading more...";
    els.inventoryBody?.closest(".table-wrap")?.append(note);
  }
  note.hidden = false;
  clearTimeout(state._inventoryLoadingMoreTimer);
  state._inventoryLoadingMoreTimer = setTimeout(() => { note.hidden = true; }, 900);
}


function ensureInventoryColGroup() {
  const table = document.querySelector("#inventory");
  if (!table) return;
  let colgroup = table.querySelector("colgroup[data-inventory-cols]");
  const key = JSON.stringify({ order: state.columnOrder, visible: state.visibleColumns });
  if (colgroup?.dataset.key === key) return;
  if (!colgroup) {
    colgroup = document.createElement("colgroup");
    colgroup.dataset.inventoryCols = "true";
    table.prepend(colgroup);
  }
  colgroup.dataset.key = key;
  colgroup.innerHTML = `<col class="checkbox-col" />` + inventoryVisibleColumnKeys()
    .map((colKey) => `<col data-col="${escapeHtml(colKey)}" />`)
    .join("");
}

function renderInventoryHeader() {
  const row = document.querySelector("#inventory thead tr");
  if (!row) return;
  const headerKey = JSON.stringify({
    order: inventoryVisibleColumnKeys(),
    visible: state.visibleColumns,
    sort: state.inventorySort,
    arrange: state.arrangeColumns,
    resize: state.resizeColumns,
  });
  if (state._inventoryHeaderKey === headerKey && row.children.length > 1) {
    applyColumnVisibility();
    applyInventoryColumnWidths();
    updateSortHeaders();
    return;
  }
  state._inventoryHeaderKey = headerKey;
  const labels = Object.fromEntries(inventoryColumns);
  // Always start with the checkbox th, then data columns
  ensureInventoryColGroup();
  row.innerHTML = `<th class="checkbox-col"><input type="checkbox" id="selectAllInventory" title="Select / deselect all" /></th>` +
    inventoryVisibleColumnKeys().map((key) => `
    <th data-col="${key}" data-sort="${key}" class="${state.resizeColumns ? "is-resizable" : ""}" ${state.arrangeColumns ? 'draggable="true"' : ""}>
      <span class="inventory-header-label">${escapeHtml(labels[key] || key)}</span>
      <span class="column-resize-handle" data-resize-col="${key}" title="Drag to resize. Double-click to fit."></span>
    </th>`).join("");
  // Re-wire select-all after rebuild
  const cb = row.querySelector("#selectAllInventory");
  cb?.addEventListener("change", (e) => {
    document.querySelectorAll("#inventoryBody .row-checkbox").forEach((c) => {
      c.checked = e.target.checked;
      if (e.target.checked) state.selectedInventoryCodes.add(c.dataset.code);
      else state.selectedInventoryCodes.delete(c.dataset.code);
    });
    renderInventorySummary(state.inventoryRows || []);
  });
  row.querySelectorAll("[data-sort]").forEach((header) => {
    header.addEventListener("click", () => {
      if (state.arrangeColumns || state.resizeColumns) return;
      state._pinnedAdjustCode = null;
      const key = header.dataset.sort;
      state.inventorySort = {
        key,
        dir: state.inventorySort.key === key && state.inventorySort.dir === "asc" ? "desc" : "asc",
      };
      renderInventory();
    });
    header.addEventListener("dblclick", (event) => {
      if (!state.resizeColumns) return;
      event.preventDefault();
      event.stopPropagation();
      autoFitInventoryColumn(header.dataset.col);
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
  row.querySelectorAll("[data-resize-col]").forEach((handle) => {
    handle.addEventListener("mousedown", (event) => startInventoryColumnResize(event, handle.dataset.resizeCol));
    handle.addEventListener("pointerdown", (event) => startInventoryColumnResize(event, handle.dataset.resizeCol));
    handle.addEventListener("touchstart", (event) => startInventoryColumnResize(event, handle.dataset.resizeCol), { passive: false });
    handle.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopPropagation();
    });
    handle.addEventListener("dblclick", (event) => {
      event.preventDefault();
      event.stopPropagation();
      autoFitInventoryColumn(handle.dataset.resizeCol);
    });
  });
  applyColumnVisibility();
  applyInventoryColumnWidths();
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

function inventoryStateSelectHtml(item) {
  const current = normalizeItemState(item.state || "");
  const stateClass = current ? `state-select-${current.toLowerCase().replace(/\s+/g, "")}` : "";
  const options = allowedItemStates()
    .map((value) => `<option value="${escapeHtml(value)}"${value === current ? " selected" : ""}>${escapeHtml(value || "Blank")}</option>`)
    .join("");
  return `<select class="inventory-edit-select inventory-edit-select--state ${stateClass}" data-item-field="state" data-code="${escapeHtml(item.code)}" data-prev-value="${escapeHtml(current)}" ${isUserRole() ? "disabled" : ""}>${options}</select>`;
}

function inventoryCaseSizeInputHtml(item) {
  return `<input type="number" class="inventory-edit-input mini-input" data-item-field="caseSize" data-code="${escapeHtml(item.code)}" data-prev-value="${escapeHtml(item.caseSize || 1)}" value="${escapeHtml(Math.max(1, Math.round(toNumber(item.caseSize) || 1)))}" min="1" step="1" ${isUserRole() ? "disabled" : ""} />`;
}

function inventoryCellHtml(key, item) {
  const override = reorderOverrideForCode(item.code);
  const minOverridden = override.min != null;
  const maxOverridden = override.max != null;
  const anyOverride = minOverridden || maxOverridden;
  const hasManualRule = item.orderingRuleMode === "custom";
  const hasCategoryRule = item.orderingRuleMode === "category";
  const ruleText = `${number.format(item.effectiveSafetyDays || 0)}/${number.format(item.effectiveDaysOfInventory || 0)}`;
  const ruleTitle = hasManualRule
    ? `Manual safety/DOI override: ${ruleText}`
    : hasCategoryRule
      ? `Category default safety/DOI rule: ${ruleText}`
      : `Vendor/default safety/DOI rule: ${ruleText}`;
  const backorderCount = poBackorderCountForCode(item.code);
  const values = {
    code: `<td data-col="code" class="copy-code full-code-cell" data-copy-code="${escapeHtml(item.code)}" title="Click code to copy">${escapeHtml(item.code)}</td>`,
    product: `<td data-col="product" class="sku-name" title="${escapeHtml(item.product)}"><span class="sku-name-text">${escapeHtml(item.product)}</span></td>`,
    plu: `<td data-col="plu">${escapeHtml(item.plu || "-")}</td>`,
    rule: `<td data-col="rule" class="rule-col${hasManualRule ? " rule-col--manual" : ""}" data-code="${escapeHtml(item.code)}" title="${escapeHtml(ruleTitle)}">${ruleText}</td>`,
    itemNumber: `<td data-col="itemNumber">${escapeHtml(item.itemNumber || "-")}</td>`,
    subType: `<td data-col="subType">${escapeHtml(item.subType || "-")}</td>`,
    sizeAttr: `<td data-col="sizeAttr">${escapeHtml(item.sizeAttr || "-")}</td>`,
    containerAttr: `<td data-col="containerAttr">${escapeHtml(item.containerAttr || "-")}</td>`,
    category: `<td data-col="category">${escapeHtml(item.category || "-")}</td>`,
    vendor: `<td data-col="vendor">${escapeHtml(item.vendor || "-")}</td>`,
    state: `<td data-col="state" class="inventory-edit-cell inventory-state-cell">${inventoryStateSelectHtml(item)}${backorderCount >= 3 ? `<span class="state-badge state-backorder" title="Backordered ${number.format(backorderCount)} times">B/O</span>` : ""}${backorderCount ? `<span class="po-backorder-count" title="Backordered ${number.format(backorderCount)} time${backorderCount === 1 ? "" : "s"}">${number.format(backorderCount)}</span>` : ""}</td>`,
    addDate: `<td data-col="addDate">${escapeHtml(formatShortDisplayDate(item.addDate))}</td>`,
    stock: `<td data-col="stock" class="num stock-col stock-clickable" title="Click to adjust stock">${number.format(item.stock)}</td>`,
    units: `<td data-col="units" class="num sold-col">${number.format(item.units)}</td>`,
    velocity: `<td data-col="velocity" class="num">${formatVelocity(item.velocity)}</td>`,
    unitCost: `<td data-col="unitCost" class="num">${currency.format(item.unitCost)}</td>`,
    price: `<td data-col="price" class="num">${currency.format(item.price)}</td>`,
    inventoryCost: `<td data-col="inventoryCost" class="num">${currency.format(item.inventoryCost)}</td>`,
    caseSize: `<td data-col="caseSize" class="num inventory-edit-cell">${inventoryCaseSizeInputHtml(item)}</td>`,
    reorderMin: `<td data-col="reorderMin" class="num order-col ${minOverridden ? "override-cell" : ""}">
      <input class="mini-input ${minOverridden ? "overridden" : ""}" data-code="${escapeHtml(item.code)}" data-reorder-field="min"
        value="${escapeHtml(item.reorderMin)}" placeholder="${number.format(item.dynamicMin)}" title="${minOverridden ? `Auto: ${number.format(item.dynamicMin)} - manual override` : `Auto: SV x Safety = ${number.format(item.dynamicMin)}`}" ${isUserRole() ? "disabled" : ""} />
    </td>`,
    reorderMax: `<td data-col="reorderMax" class="num order-col ${maxOverridden ? "override-cell" : ""}">
      <div class="minmax-cell">
        <input class="mini-input ${maxOverridden ? "overridden" : ""}" data-code="${escapeHtml(item.code)}" data-reorder-field="max"
          value="${escapeHtml(item.reorderMax)}" placeholder="${number.format(item.dynamicMax)}" title="${maxOverridden ? `Auto: ${number.format(item.dynamicMax)} - manual override` : `Auto: Min + (SV x DOI) = ${number.format(item.dynamicMax)}`}" ${isUserRole() ? "disabled" : ""} />
        ${anyOverride && !isUserRole() ? `<button class="reset-override" data-code="${escapeHtml(item.code)}" data-field="all" title="Clear manual min/max overrides"><span aria-hidden="true">&#128465;</span></button>` : ""}
      </div>
    </td>`,
    needs: (() => {
      if (isPendingOrder(item.code)) {
        return `<td data-col="needs" class="num needs-col pending-needs-cell"><button type="button" class="pending-clear-btn pending-inline-btn" data-code="${escapeHtml(item.code)}" title="PO pending - click to clear"><span class="pending-icon">&#x1F550;</span></button></td>`;
      }
      const needed = item.recommendedOrder || 0;
      if (!needed || item.isOrderable === false) return `<td data-col="needs" class="num"></td>`;
      return `<td data-col="needs" class="num needs-col needs-alert">${number.format(needed)}</td>`;
    })(),
  };
  return values[key] || "";
}

function buildInventoryRowNode(item) {
  const tr = document.createElement("tr");
  tr.dataset.itemCode = item.code;
  tr.dataset.tooltipDeferred = "1";
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
    renderInventorySummary(state.inventoryRows || []);
  });
  cbTd.append(cbInput);
  tr.append(cbTd);
  tr.insertAdjacentHTML("beforeend", inventoryVisibleColumnKeys().map((key) => inventoryCellHtml(key, item)).join(""));
  return tr;
}

function wireInventoryRowInteractions(row) {
  if (!row) return;
  row.querySelectorAll(".pending-clear-btn").forEach((btn) => {
    btn.addEventListener("click", (event) => {
      event.stopPropagation();
      const code = btn.dataset.code;
      clearPendingForCode(code, { reason: "manual", backorder: true, render: false });
      patchInventoryRow(code);
      renderOrders();
      showToast("Pending cleared for " + code, 2000, "success");
    });
  });
}

function closeRulePopover() {
  document.querySelector("#itemRulePopover")?.remove();
}

function openRulePopover(code, anchor) {
  closeRulePopover();
  const item = findCurrentItemByCode(code) || state._inventoryRowIndex?.get?.(codeKey(code));
  if (!item || isUserRole()) return;
  const rule = effectiveItemRule(item);
  const popover = document.createElement("div");
  popover.id = "itemRulePopover";
  popover.className = "item-rule-popover";
  popover.dataset.mode = rule.mode === "custom" ? "custom" : "vendor";
  popover.innerHTML = `
    <div class="item-rule-popover__title">SD / DOI</div>
    <div class="item-rule-popover__item">${escapeHtml(item.product || item.code)}</div>
    <div class="item-rule-popover__grid">
      <label><span>SD</span><input type="number" min="0" step="1" data-popover-safety value="${escapeHtml(String(rule.safetyDays || 0))}" /></label>
      <label><span>DOI</span><input type="number" min="0" step="1" data-popover-doi value="${escapeHtml(String(rule.daysOfInventory || 0))}" /></label>
    </div>
    <div class="item-rule-popover__hint">Vendor/default: ${number.format(rule.vendorSafety || 0)} / ${number.format(rule.vendorDoi || 0)}</div>
    <div class="item-rule-popover__actions">
      <button type="button" class="count-submit-btn" data-popover-save>Save</button>
      <button type="button" class="secondary-button" data-popover-vendor>Use vendor</button>
    </div>`;
  document.body.append(popover);
  const rect = anchor.getBoundingClientRect();
  popover.style.left = `${Math.min(window.innerWidth - 260, Math.max(8, rect.left))}px`;
  popover.style.top = `${Math.min(window.innerHeight - 190, rect.bottom + 8)}px`;
  const refreshAfterRuleChange = (message) => {
    const quickValue = els.inventoryQuickFilter?.value || "";
    const needsFullRuleRefresh = quickValue.includes("ruleOverrides") || state.inventorySort.key === "rule";
    const refreshedItem = refreshVisibleInventoryRowByCode(item.code, { fullRefresh: needsFullRuleRefresh }) || patchInventoryRowFromCache(item.code) || item;
    if (state._activeDetailCode && codeKey(state._activeDetailCode) === codeKey(item.code)) showDetail(refreshedItem, { preserveTab: true });
    closeRulePopover();
    showToast(message, 2200, "success");
  };
  const saveRuleFromPopover = () => {
    const safety = Math.max(0, Math.round(toNumber(popover.querySelector("[data-popover-safety]")?.value) || 0));
    const doi = Math.max(0, Math.round(toNumber(popover.querySelector("[data-popover-doi]")?.value) || 0));
    if (popover.dataset.mode === "vendor") {
      setItemMeta(item.code, {
        orderingRuleMode: "vendor",
        safetyDaysOverride: null,
        daysOfInventoryOverride: null,
      });
      refreshAfterRuleChange(`Reverted ${item.code} to vendor SD/DOI`);
      return;
    }
    setItemMeta(item.code, {
      orderingRuleMode: "custom",
      safetyDaysOverride: safety,
      daysOfInventoryOverride: doi,
    });
    refreshAfterRuleChange(`Saved SD/DOI for ${item.code}`);
  };
  popover.querySelector("[data-popover-save]")?.addEventListener("click", saveRuleFromPopover);
  popover.querySelector("[data-popover-vendor]")?.addEventListener("click", () => {
    popover.dataset.mode = "vendor";
    const safetyInput = popover.querySelector("[data-popover-safety]");
    const doiInput = popover.querySelector("[data-popover-doi]");
    if (safetyInput) safetyInput.value = String(Math.max(0, Math.round(rule.vendorSafety || 0)));
    if (doiInput) doiInput.value = String(Math.max(0, Math.round(rule.vendorDoi || 0)));
    popover.querySelector("[data-popover-save]")?.focus();
  });
  popover.querySelectorAll("input").forEach((input) => {
    input.addEventListener("focus", () => input.select?.());
    input.addEventListener("click", () => input.select?.());
    input.addEventListener("input", () => { popover.dataset.mode = "custom"; });
  });
  popover.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      event.stopPropagation();
      saveRuleFromPopover();
    } else if (event.key === "Escape") {
      event.preventDefault();
      closeRulePopover();
    }
  });
  const firstInput = popover.querySelector("input");
  firstInput?.focus();
  firstInput?.select?.();
}

function patchInventoryRow(code) {
  const row = [...(els.inventoryBody?.querySelectorAll("tr[data-item-code]") || [])]
    .find((entry) => codeKey(entry.dataset.itemCode || "") === codeKey(code));
  if (!row) {
    renderInventory();
    return;
  }
  const nextItem = buildInventoryRows().find((entry) => codeKey(entry.code) === codeKey(code));
  if (!nextItem) {
    row.remove();
    return null;
  }
  state._inventoryRowIndex.set(codeKey(nextItem.code), nextItem);
  const nextRow = buildInventoryRowNode(nextItem);
  row.replaceWith(nextRow);
  wireInventoryRowInteractions(nextRow);
  applyColumnVisibility();
  return nextItem;
}

function recomputeInventoryItemForLocalEdit(item) {
  if (!item) return null;
  const key = codeKey(item.code);
  const excel = findExcelFor(item);
  const meta = itemMetaFor(item.code);
  const override = reorderOverrideForCode(item.code || key);
  const rule = effectiveItemRule({ code: item.code, category: item.category || excel.category || "", vendor: item.vendor || excel.vendor || "" });
  const dynamic = orderingTargetsForItem({
    item,
    velocity: item.velocity || 0,
    safetyDays: rule.safetyDays,
    daysOfInventory: rule.daysOfInventory,
  });
  const stateLabel = normalizeItemState(meta.stateManual ? meta.state : (excel.state || meta.state || item.state || "")) || item.state || "Active";
  const caseSize = effectiveCaseSizeFor(item, excel, meta);
  item.state = stateLabel;
  item.itemState = stateLabel.toLowerCase();
  item.isOrderable = !["discontinued", "disabled"].includes(item.itemState);
  item.caseSize = caseSize;
  item.dynamicMin = item.isOrderable ? dynamic.min : 0;
  item.dynamicMax = item.isOrderable ? dynamic.max : 0;
  item.effectiveSafetyDays = rule.safetyDays;
  item.effectiveDaysOfInventory = rule.daysOfInventory;
  item.orderingRuleMode = rule.mode;
  item.isOverridden = override.min != null || override.max != null;
  item.reorderMin = item.isOrderable ? (override.min ?? item.dynamicMin ?? 0) : 0;
  item.reorderMax = item.isOrderable ? (override.max ?? item.dynamicMax ?? 0) : 0;
  item.recommendedOrder = item.isOrderable ? recommendedOrderQty({
    stock: item.stock,
    min: item.reorderMin,
    max: item.reorderMax,
    caseSize: item.caseSize,
  }) : 0;
  item.caseOrder = calcCaseOrder(item.recommendedOrder, item.caseSize);
  item.inventoryCost = item.stock * item.unitCost;
  return item;
}

function patchCachedInventoryItem(code) {
  const targetKey = codeKey(code);
  if (!targetKey) return null;
  let patched = null;
  const patchList = (list) => {
    if (!Array.isArray(list)) return;
    for (let i = 0; i < list.length; i += 1) {
      if (codeKey(list[i]?.code) !== targetKey) continue;
      list[i] = recomputeInventoryItemForLocalEdit({ ...list[i] });
      patched = list[i];
    }
  };
  patchList(state._inventoryCache);
  patchList(state.inventoryRows);
  if (!patched) {
    const existing = state._inventoryRowIndex?.get?.(targetKey);
    if (existing) patched = recomputeInventoryItemForLocalEdit({ ...existing });
  }
  if (patched) {
    state._inventoryRowIndex.set(targetKey, patched);
  }
  return patched;
}

function patchInventoryRowFromCache(code) {
  const key = codeKey(code);
  const row = [...(els.inventoryBody?.querySelectorAll("tr[data-item-code]") || [])]
    .find((entry) => codeKey(entry.dataset.itemCode || "") === key);
  const nextItem = patchCachedInventoryItem(code);
  if (!nextItem) return null;
  if (!row) return nextItem;
  const nextRow = buildInventoryRowNode(nextItem);
  row.replaceWith(nextRow);
  wireInventoryRowInteractions(nextRow);
  applyColumnVisibility();
  return nextItem;
}

function refreshVisibleInventoryRowByCode(code, options = {}) {
  let refreshed = patchInventoryRowFromCache(code);
  if (options.fullRefresh) {
    bumpDataStamp();
    refreshed = patchInventoryRow(code) || findCurrentItemByCode(code) || refreshed;
  } else if (!refreshed) {
    refreshed = patchInventoryRow(code) || findCurrentItemByCode(code);
  }
  renderInventorySummary(state.inventoryRows || []);
  if (activeTabName() === "ordering") renderOrders();
  return refreshed;
}

function patchVisibleStockCell(code, stock) {
  const key = codeKey(code);
  [...(els.inventoryBody?.querySelectorAll("tr[data-item-code]") || [])]
    .filter((row) => codeKey(row.dataset.itemCode || "") === key)
    .forEach((row) => {
      const cell = row.querySelector('[data-col="stock"]');
      if (cell) cell.textContent = number.format(stock);
    });
}

function patchInventoryModelStock(code, stock) {
  const key = codeKey(code);
  const patch = (item) => {
    if (item && codeKey(item.code) === key) item.stock = stock;
    return item;
  };
  state.inventoryRows = (state.inventoryRows || []).map(patch);
  state.filteredSkus = (state.filteredSkus || []).map(patch);
  if (state._inventoryCache) state._inventoryCache = state._inventoryCache.map(patch);
  if (state._skuCache instanceof Map) {
    const cached = state._skuCache.get(key);
    if (cached) cached.stock = stock;
  }
  const indexed = state._inventoryRowIndex?.get?.(key);
  if (indexed) indexed.stock = stock;
}

function moveColumn(from, to) {
  if (!from || !to || from === to) return;
  const order = state.columnOrder.filter((key) => key !== from);
  order.splice(order.indexOf(to), 0, from);
  state.columnOrder = order;
  safeLocalStorageSet("posDashboardColumnOrder:v3", JSON.stringify(state.columnOrder));
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
  safeLocalStorageSet("posDashboardDetailOrder:v1", JSON.stringify(order));
}

function renderInventorySummary(rows) {
  const summaryBaseRows = buildInventoryRows({ ignoreQuickFilter: true });
  const selectedKeys = new Set([...(state.selectedInventoryCodes || [])].map((code) => codeKey(code)));
  const selectedCount = rows.filter((item) => selectedKeys.has(codeKey(item.code))).length;
  const quickValue = els.inventoryQuickFilter?.value || "";
  const showOverridesOnly = quickValue.includes("overrides");
  const showRuleOverridesOnly = quickValue.includes("ruleOverrides");
  const showNeedsOnly = quickValue.includes("needs");
  const showPendingOnly = quickValue.includes("pending");
  const showBackorderOnly = quickValue.includes("backorder");
  const pendingCount = summaryBaseRows.filter((item) => isPendingOrder(item.code)).length;
  const backorderItemCount = summaryBaseRows.filter((item) => poBackorderCountForCode(item.code) > 0).length;
  const ruleOverrideCount = summaryBaseRows.filter((item) => item.orderingRuleMode === "custom").length;
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
  const totalsHtml = `
    <span><b>${number.format(totals.items)}</b> items showing</span>
    <span><b>${number.format(totals.stock)}</b> stock</span>
    <span data-units-summary><b>${number.format(totals.sold)}</b> sold</span>
    <span><b>${formatVelocity(totals.velocity)}</b> avg/day</span>
    <span data-cost-summary><b>${currency.format(totals.unitCost)}</b> cost sum</span>
    <span data-cost-summary><b>${currency.format(totals.inventoryCost)}</b> stock cost</span>`;
  const chipsHtml = `
    <label class="inventory-toggle-chip"><input type="checkbox" id="inventoryQuickOverrides"${showOverridesOnly ? " checked" : ""} />Overrides</label>
    <label class="inventory-toggle-chip"><input type="checkbox" id="inventoryQuickRuleOverrides"${showRuleOverridesOnly ? " checked" : ""} />SD/DOI${ruleOverrideCount ? ` (${number.format(ruleOverrideCount)})` : ""}</label>
    <label class="inventory-toggle-chip"><input type="checkbox" id="inventoryQuickNeeds"${showNeedsOnly ? " checked" : ""} />Order needed</label>
    <label class="inventory-toggle-chip"><input type="checkbox" id="inventoryQuickPending"${showPendingOnly ? " checked" : ""} />PO pending${pendingCount ? ` (${number.format(pendingCount)})` : ""}</label>
    <label class="inventory-toggle-chip"><input type="checkbox" id="inventoryQuickBackorder"${showBackorderOnly ? " checked" : ""} />B/O${backorderItemCount ? ` (${number.format(backorderItemCount)})` : ""}</label>
    ${pendingCount ? `<button type="button" class="secondary-button inventory-clear-pending-button" id="inventoryClearAllPendingButton">Clear all PO pending</button>` : ""}
    ${selectedCount ? `<button type="button" class="secondary-button inventory-selected-button" id="inventorySelectedActionsButton">Selected (${number.format(selectedCount)})</button>` : ""}`;
  const totalsTarget = ensureInventoryTotalsInline();
  const chipsTarget = ensureInventoryQuickChipsInline();
  const datePresets = document.querySelector("#datePresets");
  if (datePresets && totalsTarget.parentElement !== datePresets) {
    datePresets.append(totalsTarget);
  }
  totalsTarget.innerHTML = totalsHtml;
  chipsTarget.innerHTML = chipsHtml;
  els.inventorySummary.innerHTML = "";
  els.inventorySummary.hidden = true;
  applyRoleRestrictions(true);
  document.querySelector("#inventorySelectedActionsButton")?.addEventListener("click", openInventoryBulkActionsModal);
  document.querySelector("#inventoryClearAllPendingButton")?.addEventListener("click", clearAllPendingPo);
  const syncQuickChecks = () => {
    const overrides = !!document.querySelector("#inventoryQuickOverrides")?.checked;
    const ruleOverrides = !!document.querySelector("#inventoryQuickRuleOverrides")?.checked;
    const needsBox = document.querySelector("#inventoryQuickNeeds");
    let needs = !!needsBox?.checked;
    const pending = !!document.querySelector("#inventoryQuickPending")?.checked;
    const backorder = !!document.querySelector("#inventoryQuickBackorder")?.checked;
    if (pending && needsBox) {
      needsBox.checked = false;
      needs = false;
    }
    if (els.inventoryQuickFilter) {
      const parts = [];
      if (overrides) parts.push("overrides");
      if (ruleOverrides) parts.push("ruleOverrides");
      if (needs) parts.push("needs");
      if (pending) parts.push("pending");
      if (backorder) parts.push("backorder");
      els.inventoryQuickFilter.value = parts.join("+");
    }
    renderInventory();
  };
  document.querySelector("#inventoryQuickOverrides")?.addEventListener("change", syncQuickChecks);
  document.querySelector("#inventoryQuickRuleOverrides")?.addEventListener("change", syncQuickChecks);
  document.querySelector("#inventoryQuickNeeds")?.addEventListener("change", syncQuickChecks);
  document.querySelector("#inventoryQuickPending")?.addEventListener("change", syncQuickChecks);
  document.querySelector("#inventoryQuickBackorder")?.addEventListener("change", syncQuickChecks);
}

function buildInventoryRows(options = {}) {
  const query = options.ignoreQuery ? "" : els.searchInput.value.trim().toLowerCase();
  const stateFilter = (options.ignoreFilters || options.ignoreStateFilter) ? "" : els.inventoryStateFilter.value;
  const quickFilter = options.ignoreQuickFilter ? "" : (els.inventoryQuickFilter?.value || "");
  const start = els.startDate.value || "0000-00-00";
  const end = els.endDate.value || "9999-99-99";
  const dayCount = rangeDayCount(start, end);
  const leadDays = 0;
  const safetyDays = Math.max(0, toNumber(els.safetyDays.value) || DEFAULT_ORDER_SAFETY_DAYS);
  const daysOfInventory = Math.max(0, toNumber(els.daysOfInventory?.value) || DEFAULT_ORDER_DAYS_OF_INVENTORY);
  const cacheKey = `${start}|${end}|${leadDays}|${safetyDays}|${daysOfInventory}|ordering-math-v2`;
  if (!state._inventoryCache || state._inventoryCacheStamp !== state._dataCacheStamp || state._inventoryCacheKey !== cacheKey) {
    const inventoryIndex = state.latestInventory;
    const salesByCode = new Map(buildSkuRows({ ignoreQuery: true, ignoreFilters: true }).map((sku) => [codeKey(sku.code), sku]));
    const hasCurrentInventory = inventoryIndex.size > 0;
    // Products page must be a TRUE item master: include current inventory rows,
    // Excel/product-master rows, AND any items that appear in loaded daily sales.
    // Previous logic used only latestInventory whenever any inventory existed,
    // which made product search look partial (ex: sold NEWPORT SKUs missing
    // unless they were also present in the current inventory file).
    const codes = new Set([
      ...inventoryIndex.keys(),
      ...[...state.excelItems.values()].map((item) => item.code).filter(Boolean).map(codeKey),
      ...[...salesByCode.values()].map((sku) => sku.code).filter(Boolean).map(codeKey),
    ].filter(Boolean));

    state._inventoryCache = [...codes].map((code) => {
      const inventory = inventoryIndex.get(codeKey(code)) || {};
      const excel = findExcelFor(inventory.code ? inventory : { code });
      const sales = salesByCode.get(codeKey(code)) || {};
      const itemCode = inventory.code || code;
      const meta = itemMetaFor(itemCode);
      const override = reorderOverrideForCode(itemCode || code);
      const isOverridden = override.min != null || override.max != null;
      const velocity = Math.max(0, toNumber(sales.units || 0) / Math.max(dayCount, 1));
      const caseSize = effectiveCaseSizeFor(inventory, excel, meta, sales);
      const displayState = meta.stateManual
        ? meta.state
        : (normalizeItemState(excel.state) || normalizeItemState(meta.state) || normalizeItemState(inventory.state) || "");
      const stateLabel = normalizeItemState(displayState || "Active") || "Active";
      const itemState = stateLabel.toLowerCase();
      const isOrderable = !["discontinued", "disabled"].includes(itemState);
      const rule = effectiveItemRule({
        code: itemCode,
        category: inventory.category || excel.category || sales.category || "",
        vendor: inventory.vendor || excel.vendor || sales.vendor || "",
      });
      const useRuleTargets = rule.mode === "custom" || rule.categoryDefault || !!rule.vendorRule;
      const effectiveSafety = useRuleTargets ? rule.safetyDays : safetyDays;
      const effectiveDoi = useRuleTargets ? rule.daysOfInventory : daysOfInventory;
      const dynamic = orderingTargetsForItem({
        item: {
          code: itemCode,
          category: inventory.category || excel.category || sales.category || "",
        },
        velocity,
        safetyDays: effectiveSafety,
        daysOfInventory: effectiveDoi,
      });
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
        state: stateLabel,
        itemState,
        isOrderable,
        isOverridden,
        addDate: cleanCell(meta.addDate || excel.addDate || inventory.addDate || meta.firstSeenDate || ""),
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
        effectiveSafetyDays: dynamic.safetyDays,
        effectiveDaysOfInventory: dynamic.daysOfInventory,
        orderingRuleMode: rule.mode,
        saleWindowSum: sales.units || 0,
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
      item.caseOrder = calcCaseOrder(item.recommendedOrder, item.caseSize);
      item.daysSupply = velocity > 0 ? stock / velocity : Infinity;
      item._haystack = buildSearchHaystack([
        item.code, item.product, item.vendor, item.category, item.color, item.plu, item.itemNumber,
        item.state, item.parent, item.subType, item.sizeAttr, item.containerAttr, item.otherAttrs,
        item.subGroup, item.typeGroup, item.department,
      ]);
      return item;
    });
    state._inventoryCacheStamp = state._dataCacheStamp;
    state._inventoryCacheKey = cacheKey;
  }

  return state._inventoryCache.filter((item) => {
    // Force Order is treated as Active - always include it when filtering by Active
    if (stateFilter) {
      const itemState = item.state || "";
      if (stateFilter === "Active") {
        if (itemState !== "Active" && itemState.toLowerCase() !== "force order") return false;
      } else if (itemState !== stateFilter) {
        return false;
      }
    }
    if (query && !matchesSearchQuery(item, query)) return false;
    if (!options.ignoreFilters) {
      if (els.departmentFilter.value && item.department && item.department !== els.departmentFilter.value) return false;
      if (els.categoryFilter.value && item.category !== els.categoryFilter.value) return false;
      if (els.vendorFilter.value && item.vendor !== els.vendorFilter.value) return false;
      if (els.colorFilter.value && item.color !== els.colorFilter.value && item.subType !== els.colorFilter.value) return false;
    }
    const pendingOnly = quickFilter.includes("pending");
    const backorderOnly = quickFilter.includes("backorder");
    if (quickFilter.includes("overrides") && !item.isOverridden) return false;
    if (quickFilter.includes("ruleOverrides") && item.orderingRuleMode !== "custom") return false;
    if (!pendingOnly && quickFilter.includes("needs") && !(item.recommendedOrder > 0)) return false;
    if (pendingOnly && !isPendingOrder(item.code)) return false;
    if (backorderOnly && poBackorderCountForCode(item.code) <= 0) return false;
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
  const excelRows = [...state.excelItems.values()];
  fillSelect(els.departmentFilter, unique(salesWithInventory.map((row) => row.department).concat(excelRows.map((row) => row.department))));
  fillSelect(els.categoryFilter, unique(salesWithInventory.map((row) => row.category || row.inventory.category).concat(inventoryRows.map((row) => row.category), excelRows.map((row) => row.category))));
  fillSelect(els.vendorFilter, unique(salesWithInventory.map((row) => (row.vendor !== "Unassigned" ? row.vendor : row.inventory.vendor)).concat(inventoryRows.map((row) => row.vendor), excelRows.map((row) => row.vendor))));
  fillSelect(els.colorFilter, unique(salesWithInventory.map((row) => row.inventory.color).concat(inventoryRows.map((row) => row.color), excelRows.map((row) => row.color))));
  updateInventoryStateFilter();
}

function updateInventoryStateFilter() {
  fillSelect(els.inventoryStateFilter, allowedItemStates());
  const options = [...els.inventoryStateFilter.options].map((option) => option.value);
  if (options.includes("Active") && (!els.inventoryStateFilter.value || els.inventoryStateFilter.value === "All")) {
    els.inventoryStateFilter.value = "Active";
    return;
  }
  if (!options.includes(els.inventoryStateFilter.value)) {
    els.inventoryStateFilter.value = options.includes("Active") ? "Active" : "";
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

function downloadBlob(fileName, blob) {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
}

// -- PO Export helpers --------------------------------------------------------

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
      totalCost:        orderLineCost(item),
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

  // Summary sheet - all items
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
      [`Purchase Order - ${vendor}`, "", `Date: ${new Date().toLocaleDateString()}`],
      [],
    ["Code", "Product", "PLU", "Item #", "Case", "Order Qty", "Case Order", "Unit Cost", "Total Cost", "Stock", "SV/Day", "Override"],
    ...vRows.map((r) => [r.code, r.product, r.plu, r.itemNumber, r.caseSize, r.orderQty, r.caseOrderQty, r.unitCost, r.totalCost, r.currentStock, r.svPerDay, r.isOverridden]),
      [],
      ["", "", "", "", "", "TOTAL", "", vRows.reduce((s, r) => s + r.totalCost, 0)],
    ];
    const ws = xlsx.utils.aoa_to_sheet(vData);
    ws["!cols"] = [12,34,10,12,6,9,10,10,8,8,8].map((w) => ({ wch: w }));
    // Safe sheet name - Excel limits to 31 chars, no special chars
    const sheetName = vendor.replace(/[\\/*?:[\]]/g, "").slice(0, 31) || "Vendor";
    xlsx.utils.book_append_sheet(wb, ws, sheetName);
  });

  xlsx.writeFile(wb, `PO_${new Date().toISOString().slice(0, 10)}.xlsx`);
  showToast(`PO exported - ${rows.length} items across ${vendors.length} vendor(s)`);
}

function exportPoPdf() {
  const today = new Date().toLocaleDateString("en-US", { weekday: "long" }).toLowerCase();
  const pendingVendors = pendingVendorNames();
  const dismissedVendors = new Set((state.orderDismissedVendors || []).map((vendor) => String(vendor || "").toUpperCase()).filter(Boolean));
  const allowedVendors = state.vendorRules
    .filter((r) => r.status === "Active" && (r.orderDays || []).includes(today))
    .map((r) => String(r.vendor || "").toUpperCase())
    .filter((vendor) => vendor && !pendingVendors.has(vendor) && !dismissedVendors.has(vendor));
  const rows = buildPoRows().filter((row) => allowedVendors.includes(String(row.vendor || "").toUpperCase()));
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
  <div class="meta">Generated ${dateStr} &nbsp; - &nbsp; ${rows.length} items &nbsp; - &nbsp; ${vendors.length} vendor(s)</div>`;

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
        <td>${escapeHtml(r.product)}${r.isOverridden === "Yes" ? ' <span class="lock override" title="Manual override">Manual</span>' : ""}</td>
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
  if (!win) { showToast("Pop-up blocked - allow pop-ups first.", 3000, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 400);
  // After PDF review, offer to submit POs
  setTimeout(() => {
    if (confirm("PO PDF opened for review.\n\nSubmit POs now? Items will be marked Pending for 7 days.")) {
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

function dateFromRows(rows = []) {
  const sample = rows.find((row) => row && typeof row === "object") || {};
  const value = cleanCell(
    sample.snapshotDate || sample.snapshot_date || sample["Snapshot Date"] || sample["SNAPSHOT DATE"] ||
    sample.inventoryDate || sample.inventory_date || sample["Inventory Date"] || sample.date || sample.Date,
  );
  const iso = normalizeItemDate(value);
  if (!iso) return null;
  const [yyyy, mm, dd] = iso.split("-");
  return { iso, compact: `${mm}${dd}${yyyy}` };
}

function todayFileDate() {
  const iso = new Date().toISOString().slice(0, 10);
  const [yyyy, mm, dd] = iso.split("-");
  return { iso, compact: `${mm}${dd}${yyyy}` };
}

function looksLikeInventoryRows(rows = []) {
  const sample = rows.find((row) => row && typeof row === "object") || {};
  const keys = Object.keys(sample).map((key) => key.toLowerCase().replace(/[^a-z0-9]/g, ""));
  const hasCode = keys.some((key) => ["code", "barcode", "sku", "upc"].includes(key));
  const hasStock = keys.some((key) => ["stock", "stockonhand", "qtyonhand", "quantityonhand", "onhand"].includes(key));
  const hasSalesOnlyMarker = keys.some((key) => ["sbtls", "subtotal", "prf", "profit", "soldqty"].includes(key));
  return hasCode && hasStock && !hasSalesOnlyMarker;
}

function fileSignature(file) {
  return [file.name || "", file.size || 0, file.lastModified || 0].join("|");
}

function switchTab(tab) {
  saveActiveTabSearch();
  setScanPerformanceMode(tab === "scanmode");
  if (["dashboard", "inventory", "ordering"].includes(tab) && els.searchInput) {
    els.searchInput.value = state.tabSearches[tab] || "";
  }
  if (!["pricecheck", "scanmode"].includes(tab)) stopPriceCheckCamera();
  document.querySelectorAll("[data-tab]").forEach((button) => button.classList.toggle("active", button.dataset.tab === tab));
  document.querySelectorAll("[data-tab-view]").forEach((view) => view.classList.toggle("active", view.dataset.tabView === tab));
  document.body.dataset.activeTab = tab;
  // V2.0: count workspace/setup/report are global overlays. When leaving the Inventory
  // (counts) tab, hide them so they never float over Dashboard/Products/etc.
  if (tab !== "counts") {
    if (els.countSessionModal) els.countSessionModal.hidden = true;
    if (els.countSetupModal) els.countSetupModal.hidden = true;
    if (els.countReportModal) els.countReportModal.hidden = true;
    const histModal = document.querySelector("#sessionHistoryModal");
    if (histModal) histModal.hidden = true;
  }
  if (["pricecheck", "scanmode"].includes(tab)) {
    if (!els.priceCheckResult?.innerHTML) renderPriceCheckResult(null);
    if (tab === "pricecheck") focusPriceCheckSearch();
    if (tab === "scanmode") {
      setTimeout(() => {
        const input = document.querySelector("#scanModeInput");
        input?.focus?.();
        input?.select?.();
      }, 30);
    }
  } else if (tab === "inventory") {
    state._productsOpened = true;
  }
  applyRoleRestrictions(true);
  if (tab === "scanmode") return;
  if (tab === "counts") {
    state._countSessionOpen = false;
    // PERF-A: Render immediately from local state — never block the counts UI on a remote fetch.
    // Previously refreshLatestCountSessions() ran synchronously here: it awaited Supabase,
    // showed a loading spinner, and blocked the entire tab for ~30s on Kindle/slow wifi.
    // Now: paint local state first, then sync in the background after a tier-appropriate delay.
    renderCountsWorkspace();
    if (ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
      const delay = scanPerformanceProfile.tier === "low" ? 3500
        : scanPerformanceProfile.tier === "medium" ? 1200
        : 600;
      setTimeout(() => {
        // Only sync if still on the counts tab when the timer fires.
        if (activeTabName() === "counts") refreshLatestCountSessions();
      }, delay);
    }
    return;
  }
  renderSharedQuickTools(tab);
  queueActiveTabRender();
}

function clearFilters() {
  state._pinnedAdjustCode = null;
  state.selectedInventoryCodes.clear();
  els.searchInput.value = "";
  saveActiveTabSearch();
  if (els.parentsSearch) els.parentsSearch.value = "";
  els.departmentFilter.value = "";
  els.categoryFilter.value = "";
  els.vendorFilter.value = "";
  els.colorFilter.value = "";
  els.inventoryStateFilter.value = "Active";
  if (els.inventoryQuickFilter) els.inventoryQuickFilter.value = "";
  applyDatePreset(90);
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
  if (item) showDetail(item, { preserveTab: true });
}

function saveManualMultiBarcodes(code, aliases = []) {
  const key = rawCodeKey(code);
  if (!key) return;
  const cleanAliases = [...new Set((aliases || []).map(normalizeCode).filter(Boolean).filter((alias) => rawCodeKey(alias) !== key))];
  state.manualMultiBarcodes = state.manualMultiBarcodes || {};
  if (cleanAliases.length) state.manualMultiBarcodes[key] = cleanAliases;
  else delete state.manualMultiBarcodes[key];
  rebuildMultiBarcodeLookup();
}

function showDetail(item, options = {}) {
  if (!item) return;
  const preserveTab = options.preserveTab === true;
  if (!preserveTab || !state._activeDetailCode || codeKey(state._activeDetailCode) !== codeKey(item.code)) {
    state.detailDrawerTab = "fields";
  }
  state._activeDetailCode = item.code;
  const userMode = isUserRole();
  const sales = buildSkuRows({ ignoreQuery: true, ignoreFilters: true }).find((sku) => codeKey(sku.code) === codeKey(item.code)) || item;
  const windows = salesWindowsFor(item.code);
  const activeDetailTab = userMode && state.detailDrawerTab === "sales" ? "fields" : (state.detailDrawerTab || "fields");
  const itemRule = effectiveItemRule({
    ...item,
    vendor: item.vendor || sales.vendor || "",
    code: item.code,
  });
  const aliases = multiAliasesForCode(item.code);
  const salesHistory = (state.rawSales || [])
    .filter((row) => codeKey(row.code) === codeKey(item.code))
    .sort((a, b) => String(b.date || "").localeCompare(String(a.date || "")))
    .slice(0, 90);
  const detailCountSessions = [
    ...(state.activeCountSession?.id ? [state.activeCountSession] : []),
    ...(state.countSessions || []).filter((session) => session?.id !== state.activeCountSession?.id),
  ];
  const countHistoryRows = detailCountSessions.flatMap((session) => {
    let previousForCode = session.preCountSnapshot?.[codeKey(item.code)];
    return (session.entries || []).flatMap((entry) => {
      if (codeKey(entry.code) !== codeKey(item.code)) return [];
      const before = previousForCode ?? entry.originalQty ?? 0;
      const after = Number(entry.countedQty || 0);
      previousForCode = after;
      return [{
        recordedAt: entry.recordedAt || session.submittedAt || session.updatedAt || "",
        code: entry.code || item.code,
        product: entry.product || item.product || "",
        vendor: entry.vendor || item.vendor || "",
        category: entry.category || item.category || "",
        action: session.submittedAt ? "PHYSICAL COUNT" : "COUNT ENTRY",
        qtyChange: after - Number(before || 0),
        qtyBefore: before,
        qtyAfter: after,
        user: entry.user || session.user || "System",
        reason: `${session.submittedAt ? "Submitted physical count" : "Saved count entry"}: ${countSessionLabel(session)} (${entry.mode || "set"})`,
      }];
    });
  });
  const historyRows = [...(state.adjustmentLog || []), ...countHistoryRows]
    .filter((entry) => codeKey(entry.code) === codeKey(item.code))
    .sort((a, b) => String(b.recordedAt || "").localeCompare(String(a.recordedAt || "")))
    .slice(0, 90);
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
    ["velocity", "Sales velocity", "sales", `<b>${formatVelocity(sales.velocity || item.velocity || 0)} / day</b>`],
    ["vendor", "Vendor", "inventory", `<b>${escapeHtml(item.vendor || sales.vendor || "-")}</b>`],
    ["category", "Category", "inventory", `<b>${escapeHtml(item.category || sales.category || "-")}</b>`],
    ["color", "Color", "inventory", `<b>${escapeHtml(item.color || sales.color || "-")}</b>`],
    ["state", "State", "inventory", `<b>${escapeHtml(item.state || sales.state || "-")}</b>`],
    ["addDate", "Add date", "inventory", `<b>${escapeHtml(item.addDate || sales.addDate || "-")}</b>`],
    ["stock", "Stock", "inventory", `<b>${number.format(item.stock ?? sales.stock ?? 0)}</b>`],
    ["price", "Price", "inventory", `<b>${currency.format(item.price || sales.price || 0)}</b>`],
    ["caseSize", "Case size", "ordering", `<b>${number.format(item.caseSize || sales.caseSize || 1)}</b>`],
    ["minMax", "Min / Max", "ordering", `<b>${number.format(item.reorderMin || sales.reorderMin || 0)} / ${number.format(item.reorderMax || sales.reorderMax || 0)}</b>`],
    ["recommendedOrder", "Recommended order", "ordering", `<b>${number.format(sales.recommendedOrder || item.recommendedOrder || item.qtyNeeded || 0)}</b>`],
    ["caseOrder", "Case order", "ordering", `<b>${number.format(item.caseOrder || 0)}</b>`],
  ];
  if (!userMode) {
    fields.splice(9, 0,
      ["sales", "Sold total price", "sales", `<b>${currency.format(sales.sales || 0)}</b>`],
      ["costSold", "Sold total cost", "sales", `<b>${currency.format(sales.costSold || ((sales.units || item.units || 0) * (item.unitCost || sales.unitCost || 0)))}</b>`],
    );
    fields.splice(fields.findIndex(([key]) => key === "price"), 0, ["unitCost", "Cost", "inventory", `<b>${currency.format(item.unitCost || sales.unitCost || 0)}</b>`]);
    fields.splice(fields.findIndex(([key]) => key === "caseSize"), 0, ["inventoryCost", "Stock cost total", "inventory", `<b>${currency.format(item.inventoryCost || sales.inventoryCost || 0)}</b>`]);
  }
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
    <div class="detail-tab-row">
      <button type="button" class="detail-tab${activeDetailTab === "fields" ? " active" : ""}" data-detail-tab="fields">Fields</button>
      <button type="button" class="detail-tab${activeDetailTab === "multi" ? " active" : ""}" data-detail-tab="multi">Multi</button>
      ${userMode ? "" : `<button type="button" class="detail-tab${activeDetailTab === "sales" ? " active" : ""}" data-detail-tab="sales">Sales</button>`}
      <button type="button" class="detail-tab${activeDetailTab === "history" ? " active" : ""}" data-detail-tab="history">History</button>
    </div>
    <section class="detail-panel"${activeDetailTab === "fields" ? "" : " hidden"}>
      <details class="detail-picker">
        <summary>Detail fields</summary>
        <div class="detail-filter">${fields.map(([key, label]) => `
          <label><input type="checkbox" ${visibleFields.has(key) ? "checked" : ""} data-detail-filter="${key}" />${label}</label>`).join("")}</div>
      </details>
      <div class="summary-strip">
        <span><b>${itemRule.mode === "custom" ? "Custom rule" : itemRule.mode === "category" ? "Category rule" : "Vendor rule"}</b> ordering mode</span>
        <span><b>${number.format(itemRule.safetyDays || 0)}</b> safety days</span>
        <span><b>${number.format(itemRule.daysOfInventory || 0)}</b> days of inventory</span>
      </div>
      <div class="detail-grid">
        ${sortedFields.map(([key, label, section, value]) => `<span draggable="false" data-detail-key="${key}" data-detail-section="${section}" style="${visibleFields.has(key) ? "" : "display:none;"}">${label} ${value}</span>`).join("")}
      </div>
    </section>
    <section class="detail-panel detail-multi-panel"${activeDetailTab === "multi" ? "" : " hidden"}>
      <div class="detail-multi-tools">
        <button type="button" class="secondary-button" data-open-multi-upload>Load Multi Excel</button>
        <span class="muted">${escapeHtml(state.multiBarcodeFileName || "No multi barcode workbook loaded")}</span>
      </div>
      <div class="detail-multi-primary"><b>Primary code</b> ${escapeHtml(item.code)}</div>
      <div class="detail-multi-grid">
        ${Array.from({ length: 16 }, (_, index) => {
          const value = aliases[index] || "";
          const slot = index + 2;
          return `
            <label class="detail-multi-field">
              <span>Barcode #${slot}</span>
              <input type="text" value="${escapeHtml(value)}" data-multi-slot="${index}" />
            </label>`;
        }).join("")}
      </div>
      <div class="detail-multi-actions">
        <button type="button" class="count-submit-btn" data-save-multi>Save Multi</button>
        <button type="button" class="secondary-button" data-clear-multi>Clear Multi</button>
      </div>
    </section>
    <section class="detail-panel"${!userMode && activeDetailTab === "sales" ? "" : " hidden"}>
      <div class="summary-strip">
        <span><b>${number.format(salesHistory.reduce((sum, row) => sum + (row.units || 0), 0))}</b> qty in history</span>
        <span><b>${currency.format(salesHistory.reduce((sum, row) => sum + (row.sales || 0), 0))}</b> sales</span>
        <span><b>${number.format(item.stock ?? sales.stock ?? 0)}</b> stock left</span>
      </div>
      <div class="table-wrap">
        <table class="count-report-table inner-auto-table">
          <thead><tr><th>Date</th><th>Qty</th><th>Sales $</th><th>Vendor</th></tr></thead>
          <tbody>${salesHistory.length
            ? salesHistory.map((row) => `<tr><td>${escapeHtml(row.date || "-")}</td><td class="num">${number.format(row.units || 0)}</td><td class="num">${currency.format(row.sales || 0)}</td><td>${escapeHtml(row.vendor || item.vendor || "-")}</td></tr>`).join("")
            : `<tr><td colspan="4" class="empty-cell">No daily sales history loaded for this item.</td></tr>`}
          </tbody>
        </table>
      </div>
    </section>
    <section class="detail-panel"${activeDetailTab === "history" ? "" : " hidden"}>
      <div class="summary-strip">
        <span><b>${number.format(historyRows.length)}</b> stock events</span>
        <span><b>${escapeHtml(historyRows[0]?.recordedAt ? new Date(historyRows[0].recordedAt).toLocaleString() : "-")}</b> latest</span>
      </div>
      <div class="table-wrap">
        <table class="count-report-table inner-auto-table">
          <thead><tr><th>Date/Time</th><th>Action</th><th>Change</th><th>Before</th><th>After</th><th>User</th><th>Reason</th></tr></thead>
          <tbody>${historyRows.length
            ? historyRows.map((entry) => `<tr><td>${escapeHtml(new Date(entry.recordedAt).toLocaleString())}</td><td>${escapeHtml(entry.action || "-")}</td><td class="num">${number.format(entry.qtyChange || 0)}</td><td class="num">${number.format(entry.qtyBefore || 0)}</td><td class="num">${number.format(entry.qtyAfter || 0)}</td><td>${escapeHtml(entry.user || "System")}</td><td>${escapeHtml(entry.reason || "-")}</td></tr>`).join("")
            : `<tr><td colspan="7" class="empty-cell">No stock history recorded for this item yet.</td></tr>`}
          </tbody>
        </table>
      </div>
    </section>`;
  const multiDraftSignature = (values) => JSON.stringify(
    Array.from({ length: 16 }, (_, index) => normalizeCode(values[index] || ""))
  );
  const readMultiDraftValues = () => [...els.detailDrawer.querySelectorAll("[data-multi-slot]")].map((input) => normalizeCode(input.value));
  const initialMultiSignature = multiDraftSignature(aliases);
  const multiDraftChanged = () => multiDraftSignature(readMultiDraftValues()) !== initialMultiSignature;
  const persistMultiDraft = () => {
    const nextAliases = readMultiDraftValues();
    saveManualMultiBarcodes(item.code, nextAliases);
    showToast(`Saved multi barcodes for ${item.code}`, 2200, "success");
  };
  const confirmMultiDraftBeforeLeave = () => {
    if (state.detailDrawerTab !== "multi" || !multiDraftChanged()) return true;
    if (confirm("Save multi barcode changes for this item?")) {
      persistMultiDraft();
    }
    return true;
  };
  els.detailDrawer.querySelector(".drawer-close").addEventListener("click", () => {
    if (!confirmMultiDraftBeforeLeave()) return;
    els.detailDrawer.hidden = true;
    state._activeDetailCode = "";
  });
  els.detailDrawer.querySelectorAll("[data-detail-tab]").forEach((button) => {
    button.addEventListener("click", () => {
      if (!confirmMultiDraftBeforeLeave()) return;
      state.detailDrawerTab = button.dataset.detailTab || "fields";
      showDetail(item, { preserveTab: true });
    });
  });
  els.detailDrawer.querySelector("[data-open-multi-upload]")?.addEventListener("click", () => els.multiBarcodeInput?.click());
  els.detailDrawer.querySelector("[data-save-multi]")?.addEventListener("click", () => {
    if (!confirm("Save multi barcode changes for this item?")) return;
    persistMultiDraft();
    refreshDetailDrawer();
  });
  els.detailDrawer.querySelector("[data-clear-multi]")?.addEventListener("click", () => {
    if (!confirm("Clear all multi barcode entries for this item?")) return;
    saveManualMultiBarcodes(item.code, []);
    showToast(`Cleared manual multi barcodes for ${item.code}`, 2200, "success");
    refreshDetailDrawer();
  });
  els.detailDrawer.querySelector(".copy-code").addEventListener("click", (event) => copyText(item.code, event.currentTarget));
  els.detailDrawer.querySelectorAll("[data-detail-filter]").forEach((input) => {
    input.addEventListener("change", () => {
      const visible = new Set([...els.detailDrawer.querySelectorAll("[data-detail-filter]:checked")].map((item) => item.dataset.detailFilter));
      state.detailFilters = [...visible];
      safeLocalStorageSet("posDashboardDetailFilters:v1", JSON.stringify(state.detailFilters));
      els.detailDrawer.querySelectorAll("[data-detail-key]").forEach((node) => {
        const show = visible.has(node.dataset.detailKey);
        node.hidden = !show;
        node.style.display = show ? "" : "none";
      });
    });
  });
  const itemRuleCustomInput = els.detailDrawer.querySelector("[data-item-rule-custom]");
  const itemRuleSafetyInput = els.detailDrawer.querySelector("[data-item-rule-safety]");
  const itemRuleDoiInput = els.detailDrawer.querySelector("[data-item-rule-doi]");
  const syncItemRuleInputs = () => {
    const enabled = !!itemRuleCustomInput?.checked;
    if (itemRuleSafetyInput) itemRuleSafetyInput.disabled = !enabled;
    if (itemRuleDoiInput) itemRuleDoiInput.disabled = !enabled;
  };
  syncItemRuleInputs();
  itemRuleCustomInput?.addEventListener("change", syncItemRuleInputs);
  els.detailDrawer.querySelector("[data-save-item-rule]")?.addEventListener("click", () => {
    const useCustom = !!itemRuleCustomInput?.checked;
    const safetyValue = Math.max(0, Math.round(toNumber(itemRuleSafetyInput?.value) || 0));
    const doiValue = Math.max(0, Math.round(toNumber(itemRuleDoiInput?.value) || 0));
    setItemMeta(item.code, useCustom ? {
      orderingRuleMode: "custom",
      safetyDaysOverride: safetyValue,
      daysOfInventoryOverride: doiValue,
    } : {
      orderingRuleMode: "vendor",
      safetyDaysOverride: null,
      daysOfInventoryOverride: null,
    });
    const refreshedItem = refreshVisibleInventoryRowByCode(item.code) || findCurrentItemByCode(item.code) || item;
    showDetail(refreshedItem, { preserveTab: true });
    showToast(`Saved item ordering rule for ${item.code}`, 2400, "success");
  });
  els.detailDrawer.querySelector("[data-clear-item-rule]")?.addEventListener("click", () => {
    setItemMeta(item.code, {
      orderingRuleMode: "vendor",
      safetyDaysOverride: null,
      daysOfInventoryOverride: null,
    });
    const refreshedItem = refreshVisibleInventoryRowByCode(item.code) || findCurrentItemByCode(item.code) || item;
    showDetail(refreshedItem, { preserveTab: true });
    showToast(`Reverted ${item.code} to vendor ordering rules`, 2400, "success");
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
  // Match standalone length numbers 8-30 (even), possibly prefixed with space or parenthesis
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
  safeLocalStorageSet("posDashboardAttributeRules:v1", JSON.stringify(state.attributeRules));
  els.attributeRuleValue.value = "";
  els.attributeRuleAliases.value = "";
  render();
}

function removeAttributeRule(id) {
  state.attributeRules = state.attributeRules.filter((rule) => rule.id !== id);
  safeLocalStorageSet("posDashboardAttributeRules:v1", JSON.stringify(state.attributeRules));
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
  safeLocalStorageSet("posDashboardParentRules:v1", JSON.stringify(state.parentRules));
  els.parentRuleName.value = "";
  els.parentRuleAliases.value = "";
  render();
}

function removeParentRule(parent) {
  state.parentRules = state.parentRules.filter((rule) => rule.parent !== parent);
  safeLocalStorageSet("posDashboardParentRules:v1", JSON.stringify(state.parentRules));
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
  return Math.max(size, Math.ceil(qty / size) * size);
}

function roundInventoryTargetToCase(stock, targetQty, caseSize) {
  const size = Math.max(1, toNumber(caseSize) || 1);
  const desiredTotal = Math.max(Math.ceil(toNumber(stock) || 0), Math.ceil(toNumber(targetQty) || 0));
  return Math.ceil(desiredTotal / size) * size;
}

function recommendedOrderQty({ stock, min, max }) {
  const currentStock = Math.ceil(toNumber(stock) || 0);
  const minQty = Math.max(0, Math.ceil(toNumber(min) || 0));
  const maxQty = Math.max(minQty, Math.ceil(toNumber(max) || 0));
  if (currentStock >= minQty) return 0;
  const needed = Math.max(0, Math.ceil(maxQty - currentStock));
  if (needed <= 0) return 0;
  return needed;
}

function renderColumnPicker() {
  const sortedColumns = [...inventoryColumns].sort((a, b) => compareDisplayValue(a[1], b[1]));
  els.columnPickerPanel.innerHTML = `<div class="column-picker-grid order-cp-grid">${sortedColumns.map(([key, label]) => `
    <label class="column-choice">
      <input type="checkbox" data-column-toggle="${key}" ${state.visibleColumns[key] ? "checked" : ""} />
      <span>${label}</span>
    </label>`).join("")}</div>`;
  els.columnPickerPanel.querySelectorAll("[data-column-toggle]").forEach((input) => {
    input.addEventListener("change", () => {
      state.visibleColumns[input.dataset.columnToggle] = input.checked;
      safeLocalStorageSet("posDashboardVisibleColumns:v4", JSON.stringify(state.visibleColumns));
      applyColumnVisibility();
    });
  });
}

function applyColumnVisibility() {
  ensureInventoryColGroup();
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

function closeAppConfirmModal() {
  document.querySelector("#appConfirmModal")?.remove();
}

function showAppConfirm({ title = "Confirm", message = "", confirmText = "Confirm", cancelText = "Cancel", onConfirm } = {}) {
  closeAppConfirmModal();
  const modal = document.createElement("div");
  modal.id = "appConfirmModal";
  modal.className = "count-modal";
  modal.innerHTML = `
    <div class="count-modal__scrim" data-close-confirm></div>
    <div class="count-modal__dialog count-modal__dialog--vendor">
      <div class="count-modal__header">
        <div><p class="eyebrow">Confirmation</p><h3>${escapeHtml(title)}</h3></div>
        <div class="table-tools"><button type="button" class="secondary-button" data-close-confirm>&#x2715;</button></div>
      </div>
      <p class="muted">${escapeHtml(message)}</p>
      <div class="detail-multi-actions">
        <button type="button" class="secondary-button" data-close-confirm>${escapeHtml(cancelText)}</button>
        <button type="button" class="count-submit-btn" id="appConfirmAcceptButton">${escapeHtml(confirmText)}</button>
      </div>
    </div>`;
  document.body.append(modal);
  modal.hidden = false;
  modal.querySelectorAll("[data-close-confirm]").forEach((node) => node.addEventListener("click", closeAppConfirmModal));
  modal.querySelector("#appConfirmAcceptButton")?.addEventListener("click", () => {
    closeAppConfirmModal();
    onConfirm?.();
  });
}

function selectedInventoryItems() {
  const selectedKeys = new Set([...state.selectedInventoryCodes].map((code) => codeKey(code)));
  return currentInventoryRows().filter((item) => selectedKeys.has(codeKey(item.code)));
}

function closeInventoryBulkActionsModal() {
  document.querySelector("#inventoryBulkActionsModal")?.remove();
}

async function applyBulkInventoryChange(action, rawValue, options = {}) {
  if (isUserRole()) {
    showToast("Bulk changes are admin only.", 2400, "warning");
    return;
  }
  const items = selectedInventoryItems();
  if (!items.length) {
    showToast("Select at least one item first.", 2400, "warning");
    return;
  }
  const user = state.currentUser?.name || "System";
  if (action === "state") {
    const nextState = normalizeItemState(rawValue);
    if (!nextState) return;
    items.forEach((item) => setItemMeta(item.code, { state: nextState, stateManual: true }));
  } else if (action === "caseSize") {
    const nextValue = Math.max(1, Math.round(toNumber(rawValue) || 0));
    if (!nextValue) return;
    items.forEach((item) => setItemMeta(item.code, { caseSize: nextValue, caseSizeManual: true }));
  } else if (action === "price" || action === "cost") {
    const nextValue = Number(toNumber(rawValue) || 0);
    items.forEach((item) => {
      const key = codeKey(item.code);
      const inventory = state.latestInventory.get(key);
      if (!inventory) return;
      if (action === "price") inventory.price = nextValue;
      else inventory.cost = nextValue;
    });
  }
  if (action === "price" || action === "cost" || action === "caseSize" || action === "state") {
    items.forEach((item) => {
      state.adjustmentLog.unshift({
        recordedAt: new Date().toISOString(),
        code: item.code,
        product: item.product,
        vendor: item.vendor || "",
        category: item.category || "",
        action: `Bulk ${action}`,
        qtyChange: 0,
        qtyBefore: item.stock ?? 0,
        qtyAfter: item.stock ?? 0,
        user,
        reason: `Bulk update (${action})`,
      });
    });
  }
  await savePersistedState();
  bumpDataStamp();
  if (!options.deferRender) {
    renderInventory();
    if (activeTabName() === "ordering") renderOrders();
  }
  if (!options.keepOpen) closeInventoryBulkActionsModal();
  showToast(`Updated ${items.length} item${items.length === 1 ? "" : "s"} (${action})`, 2600, "success");
}

function openInventoryBulkActionsModal() {
  if (isUserRole()) {
    showToast("Bulk changes are admin only.", 2400, "warning");
    return;
  }
  const items = selectedInventoryItems();
  if (!items.length) {
    showToast("Select at least one item first.", 2400, "warning");
    return;
  }
  closeInventoryBulkActionsModal();
  const modal = document.createElement("div");
  modal.id = "inventoryBulkActionsModal";
  modal.className = "count-modal";
  modal.innerHTML = `
    <div class="count-modal__scrim" data-close-bulk></div>
    <div class="count-modal__dialog inventory-bulk-dialog">
      <div class="count-modal__header inventory-bulk-header">
        <div><p class="eyebrow">Bulk actions</p><h3>${number.format(items.length)} selected item${items.length === 1 ? "" : "s"}</h3></div>
        <div class="table-tools"><button type="button" class="secondary-button inventory-bulk-close" data-close-bulk>&#x2715;</button></div>
      </div>
      <div class="detail-multi-grid">
        <label class="detail-multi-field"><span>Change price</span><input type="number" step="0.01" placeholder="0.00" data-bulk-input="price" /></label>
        <label class="detail-multi-field"><span>Change cost</span><input type="number" step="0.01" placeholder="0.00" data-bulk-input="cost" /></label>
        <label class="detail-multi-field"><span>Change case size</span><input type="number" step="1" min="1" data-bulk-input="caseSize" /></label>
        <label class="detail-multi-field"><span>Change state</span><select data-bulk-input="state"><option value="">No change</option>${allowedItemStates().map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(value)}</option>`).join("")}</select></label>
      </div>
      <div class="detail-multi-actions">
        <button type="button" class="count-submit-btn" id="inventoryBulkApplyButton">Apply</button>
      </div>
    </div>`;
  document.body.append(modal);
  modal.hidden = false;
  modal.querySelectorAll("[data-close-bulk]").forEach((node) => node.addEventListener("click", closeInventoryBulkActionsModal));
  const dialog = modal.querySelector(".inventory-bulk-dialog");
  const header = modal.querySelector(".inventory-bulk-header");
  let drag = null;
  header?.addEventListener("mousedown", (event) => {
    if (event.target.closest("button")) return;
    const rect = dialog.getBoundingClientRect();
    drag = { startX: event.clientX, startY: event.clientY, left: rect.left, top: rect.top };
    dialog.style.position = "fixed";
    dialog.style.margin = "0";
    dialog.style.left = `${rect.left}px`;
    dialog.style.top = `${rect.top}px`;
  });
  const moveBulkModal = (event) => {
    if (!drag) return;
    dialog.style.left = `${drag.left + (event.clientX - drag.startX)}px`;
    dialog.style.top = `${drag.top + (event.clientY - drag.startY)}px`;
  };
  const stopMoveBulkModal = () => {
    drag = null;
  };
  window.addEventListener("mousemove", moveBulkModal);
  window.addEventListener("mouseup", stopMoveBulkModal);
  modal.querySelector("#inventoryBulkApplyButton")?.addEventListener("click", async () => {
    const price = modal.querySelector('[data-bulk-input="price"]')?.value?.trim();
    const cost = modal.querySelector('[data-bulk-input="cost"]')?.value?.trim();
    const caseSize = modal.querySelector('[data-bulk-input="caseSize"]')?.value?.trim();
    const stateValue = modal.querySelector('[data-bulk-input="state"]')?.value?.trim();
    const actions = [];
    if (price !== "") actions.push(["price", price]);
    if (cost !== "") actions.push(["cost", cost]);
    if (caseSize !== "") actions.push(["caseSize", caseSize]);
    if (stateValue) actions.push(["state", stateValue]);
    if (!actions.length) {
      showToast("Enter at least one bulk change first.", 2400, "warning");
      return;
    }
    showAppConfirm({
      title: "Apply Changes",
      message: `Apply ${actions.length} change${actions.length === 1 ? "" : "s"} to ${items.length} selected item${items.length === 1 ? "" : "s"}?`,
      confirmText: "Apply all",
      onConfirm: async () => {
        for (const [action, value] of actions) {
          await applyBulkInventoryChange(action, value, { keepOpen: true, deferRender: true });
        }
        clearInventorySelection();
        closeInventoryBulkActionsModal();
        renderInventory();
        if (activeTabName() === "ordering") renderOrders();
      },
    });
  });
  modal.addEventListener("keydown", (event) => {
    if (event.key === "Enter") {
      event.preventDefault();
      modal.querySelector("#inventoryBulkApplyButton")?.click();
    }
  });
}

// Date preset helpers - sets start/end inputs and re-renders
function applyDatePreset(days) {
  if (!state.dates.length) return;
  clearInventorySelection();
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

// Navigate date range by period - arrows follow the active preset button
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

function closeDatePickerPopup() {
  document.getElementById("datePickerPopup")?.remove();
}

function renderDatePresets() {
  const container = document.getElementById("datePresets");
  if (!container) return;
  const presets = [
    { label: "1D", days: 1 },
    { label: "7D", days: 7 },
    { label: "14D", days: 14 },
    { label: "30D", days: 30 },
    { label: "60D", days: 60 },
    { label: "90D", days: 90 },
    { label: "6M", days: 183 },
    { label: "1Y", days: 365 },
    { label: "YTD", days: "ytd" },
    { label: "ALL", days: "all" },
  ];
  if (!state._datePresetsReady) {
    container.innerHTML =
      `<div class="date-preset-chip-row">${presets.map((p) => `<button type="button" class="preset-chip" data-preset-days="${p.days}">${p.label}</button>`).join("")}</div>` +
      `<div class="date-custom-range-group">` +
      `<button type="button" class="date-arrow-btn" id="datePresetPrev" title="Previous period">&#8249;</button>` +
      `<span class="date-picker-anchor"><button type="button" class="date-range-label" id="dateRangeLabel" title="Click to pick custom dates"></button></span>` +
      `<button type="button" class="date-arrow-btn" id="datePresetNext" title="Next period">&#8250;</button>` +
      `</div>`;

    container.querySelectorAll("[data-preset-days]").forEach((btn) => {
      btn.addEventListener("click", () => {
        const presetValue = btn.dataset.presetDays === "all" ? "all"
          : btn.dataset.presetDays === "ytd" ? "ytd"
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

    document.getElementById("datePresetPrev")?.addEventListener("click", () => shiftDateRange(-1));
    document.getElementById("datePresetNext")?.addEventListener("click", () => shiftDateRange(1));

    // Click date label -> open calendar popup
    document.getElementById("dateRangeLabel")?.addEventListener("click", (e) => {
      e.stopPropagation();
      let popup = document.getElementById("datePickerPopup");
      if (popup) { popup.remove(); return; }
      popup = document.createElement("div");
      popup.id = "datePickerPopup";
      popup.className = "date-picker-popup";
      popup.innerHTML = `<div class="date-picker-popup__row">
        <label>From <input type="date" id="popStartDate" value="${els.startDate.value}" /></label>
        <label>To&nbsp;&nbsp;<input type="date" id="popEndDate" value="${els.endDate.value}" /></label>
        <button type="button" id="popApply" class="date-picker-apply">Apply</button>
        <button type="button" id="popClose" class="date-picker-close">&times;</button>
      </div>`;
      const anchor = container.querySelector(".date-picker-anchor");
      anchor?.append(popup);
      popup.querySelector("#popApply").addEventListener("click", () => {
        const s = popup.querySelector("#popStartDate").value;
        const e2 = popup.querySelector("#popEndDate").value;
        if (s) els.startDate.value = s;
        if (e2) els.endDate.value = e2;
        state.activePresetDays = null;
        popup.remove();
        render();
      });
      popup.querySelector("#popClose").addEventListener("click", () => popup.remove());
    });

    state._datePresetsReady = true;
  }

  // Update active pill highlight
  container.querySelectorAll(".preset-chip").forEach((chip) => {
    const value = chip.dataset.presetDays === "all" ? "all"
      : chip.dataset.presetDays === "ytd" ? "ytd"
      : Number(chip.dataset.presetDays);
    chip.classList.toggle("active", value === state.activePresetDays);
  });

  // Update date range label
  const lbl = document.getElementById("dateRangeLabel");
  if (lbl) {
    const s = els.startDate.value;
    const e = els.endDate.value;
    const fmt = (iso) => { if (!iso) return "-"; const [y,m,d] = iso.split("-"); return `${m}/${d}/${y}`; };
    lbl.textContent = (s && e && s !== e) ? `${fmt(s)} - ${fmt(e)}` : fmt(s || e);
  }
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
  const calculatedMin = Math.max(0, Math.ceil(rawMin));
  const min = calculatedMin > 0 ? Math.max(2, calculatedMin) : 0;
  const max = Math.max(min, Math.ceil(min + (velocity * Math.max(0, daysOfInventory || 0))));
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

function normalizedCaseSize(value) {
  return Math.max(1, Math.round(toNumber(value) || 1));
}

function effectiveCaseSizeFor(inventory = {}, excel = {}, meta = {}, fallback = {}) {
  const excelSize = normalizedCaseSize(excel.caseSize);
  const metaSize = normalizedCaseSize(meta.caseSize);
  const inventorySize = normalizedCaseSize(inventory.caseSize);
  const fallbackSize = normalizedCaseSize(fallback.caseSize);
  if (excelSize > 1) return excelSize;
  if (meta.caseSizeManual) return metaSize;
  if (metaSize > 1) return metaSize;
  if (inventorySize > 1) return inventorySize;
  if (fallbackSize > 1) return fallbackSize;
  return 1;
}

function cleanHeader(value) {
  return String(value ?? "").trim();
}

function cleanCell(value) {
  return String(value ?? "").replace(/^="/, "").replace(/"$/, "").trim();
}

function stableVendorRuleId(vendor = "") {
  const normalized = cleanCell(vendor).toUpperCase();
  if (!normalized) return "";
  return `vr-${normalized.replace(/[^A-Z0-9]+/g, "-").replace(/^-+|-+$/g, "")}`;
}

function normalizeSearchText(value) {
  return cleanCell(value).toUpperCase().replace(/[^A-Z0-9]+/g, " ").trim();
}

function buildSearchHaystack(values) {
  const raw = values.filter(Boolean).map((value) => cleanCell(value)).join(" ");
  const normalized = values.filter(Boolean).map((value) => normalizeSearchText(value)).join(" ");
  return `${raw} ${normalized}`.toLowerCase().trim();
}

function normalizeCode(value) {
  return cleanCell(value).replace(/^=/, "").replace(/^"|"$/g, "");
}

function rawCodeKey(value) {
  const code = normalizeCode(value);
  if (!code) return "";
  return /^\d+$/.test(code) ? (code.replace(/^0+/, "") || "0") : code.toUpperCase();
}

function codeKey(value) {
  const rawKey = rawCodeKey(value);
  if (!rawKey) return "";
  return state.multiBarcodeMap?.[rawKey] || rawKey;
}

function normalizeReorderOverridesState() {
  const normalized = {};
  Object.entries(state.reorderOverrides || {}).forEach(([rawKey, override]) => {
    const key = codeKey(rawKey);
    if (!key || !override || typeof override !== "object") return;
    normalized[key] = {
      ...(override.min != null ? { min: Math.max(0, Math.round(toNumber(override.min) || 0)) } : {}),
      ...(override.max != null ? { max: Math.max(0, Math.round(toNumber(override.max) || 0)) } : {}),
    };
    if (!Object.keys(normalized[key]).length) delete normalized[key];
  });
  state.reorderOverrides = normalized;
}

function reorderOverrideForCode(code) {
  return state.reorderOverrides?.[codeKey(code)] || {};
}

normalizeReorderOverridesState();

function matchesSearchQuery(item, query) {
  const rawQuery = cleanCell(query);
  const normalizedQuery = normalizeSearchText(rawQuery);
  if (!rawQuery && !normalizedQuery) return true;
  const exactFields = [item.code, item.plu, item.itemNumber];
  const exactMatch = exactFields.some((value) => {
    const raw = cleanCell(value);
    return raw && (
      raw.toLowerCase() === rawQuery.toLowerCase() ||
      codeKey(raw) === codeKey(rawQuery) ||
      normalizeSearchText(raw) === normalizedQuery
    );
  });
  if (exactMatch) return true;
  const haystack = String(item._haystack || "").toLowerCase();
  const rawWords = rawQuery.toLowerCase().split(/\s+/).filter(Boolean);
  const normalizedWords = normalizedQuery.toLowerCase().split(/\s+/).filter(Boolean);
  const words = [...new Set([...rawWords, ...normalizedWords])];
  if (!words.length) return true;
  return words.every((word) => haystack.includes(word));
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
  const recentLogs = state.uploadLogs || [];
  const loggedSalesDates = recentLogs
    .filter((entry) => entry.type === "Daily_Sale" && entry.status === "Success")
    .map((entry) => extractDateFromFilename(entry.filename))
    .filter(Boolean);
  const loggedInventoryDates = recentLogs
    .filter((entry) => entry.type === "Current_Inventory" && entry.status === "Success")
    .map((entry) => extractDateFromFilename(entry.filename))
    .filter(Boolean);
  const salesDays = new Set([...(state.dates || []), ...loggedSalesDates]).size;
  const inventoryDays = new Set([...state.inventories.keys(), ...loggedInventoryDates]).size;
  const excelItems = state.excelItems.size;
  if (!salesDays && !inventoryDays && !excelItems) return "No files loaded";
  return `${salesDays} sales days  -  ${inventoryDays} inventory snapshots  -  ${number.format(excelItems)} Excel items`;
}

function coverageSummary() {
  const loggedInventoryDates = (state.uploadLogs || [])
    .filter((entry) => entry.type === "Current_Inventory" && entry.status === "Success")
    .map((entry) => extractDateFromFilename(entry.filename))
    .filter(Boolean);
  const inventoryDates = [...new Set([...state.inventories.keys(), ...loggedInventoryDates])].sort();
  if (!state.dates.length && !inventoryDates.length) return "Choose POS sales and inventory CSV exports.";
  const salesText = state.dates.length ? `Sales ${state.dates[0]} through ${state.dates[state.dates.length - 1]}` : "No sales days loaded";
  const inventoryText = inventoryDates.length ? `current inventory ${inventoryDates[inventoryDates.length - 1]}` : "no inventory snapshot";
  return `${salesText}; ${inventoryText}`;
}

function saveUploadLogs() {
  safeLocalStorageSet("posDashboardUploadLogs:v1", JSON.stringify((state.uploadLogs || []).slice(0, 5000)));
}

function logUploadedFile({ filename, type, status }) {
  state.uploadLogs = state.uploadLogs || [];
  state.uploadLogs.unshift({
    id: `upl-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    recordedAt: new Date().toISOString(),
    filename: filename || "-",
    type: type || "File",
    status: status || "Success",
  });
  state.uploadLogs = state.uploadLogs.slice(0, 5000);
  saveUploadLogs();
  scheduleSharedImportLogsSync();
}

function extractDateFromFilename(filename = "") {
  const raw = String(filename || "");
  const compact = raw.match(/(?<!\d)(\d{8})(?!\d)/);
  if (compact) {
    const token = compact[1];
    const mm = Number(token.slice(0, 2));
    const dd = Number(token.slice(2, 4));
    const yyyy = Number(token.slice(4, 8));
    if (mm >= 1 && mm <= 12 && dd >= 1 && dd <= 31 && yyyy >= 2000 && yyyy <= 2100) {
      return `${yyyy}-${String(mm).padStart(2, "0")}-${String(dd).padStart(2, "0")}`;
    }
  }
  const iso = raw.match(/(20\d{2})[-_](\d{2})[-_](\d{2})/);
  if (iso) return `${iso[1]}-${iso[2]}-${iso[3]}`;
  return "";
}

function formatShortDisplayDate(value) {
  const iso = normalizeItemDate(value);
  if (!iso) return "-";
  const match = iso.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!match) return iso;
  const [, yyyy, mm, dd] = match;
  return `${mm}/${dd}/${yyyy.slice(-2)}`;
}

function renderUploadLogs() {
  if (!els.uploadLogBody || !els.uploadLogSummary) return;
  const recentLogs = (state.uploadLogs || []).slice(0, 5000);
  const loggedSalesDates = recentLogs
    .filter((entry) => entry.type === "Daily_Sale" && entry.status === "Success")
    .map((entry) => extractDateFromFilename(entry.filename))
    .filter(Boolean);
  const loggedInventoryDates = recentLogs
    .filter((entry) => entry.type === "Current_Inventory" && entry.status === "Success")
    .map((entry) => extractDateFromFilename(entry.filename))
    .filter(Boolean);
  const loggedExcelDates = recentLogs
    .filter((entry) => entry.type === "Excel_Product" && entry.status === "Success")
    .map((entry) => extractDateFromFilename(entry.filename))
    .filter(Boolean);
  const salesDates = [...new Set([...(state.dates || []), ...loggedSalesDates])].sort();
  const latestInventorySnapshot = [...new Set([...state.inventories.keys(), ...loggedInventoryDates])].sort().at(-1) || "";
  const latestExcelSnapshot = [...new Set([latestExcelAddDate(), ...loggedExcelDates])].filter(Boolean).sort().at(-1) || "";
  const inventoryDates = latestInventorySnapshot ? [latestInventorySnapshot] : [];
  const dataDates = latestExcelSnapshot ? [latestExcelSnapshot] : [];
  const todayIso = localIsoDate(new Date());
  const knownDates = [...new Set([...salesDates, ...inventoryDates, ...dataDates, todayIso])].filter(Boolean).sort();
  const excelLoaded = !!latestExcelSnapshot || state.excelItems.size > 0;
  const inventoryFilter = document.querySelector("#logsInventoryFilter")?.value || "";
  const salesFilter = document.querySelector("#logsSalesFilter")?.value || "";
  const dataFilter = document.querySelector("#logsDataFilter")?.value || "";
  const dayRows = [];
  if (knownDates.length) {
    let current = new Date(`${knownDates[0]}T00:00:00`);
    const end = new Date(`${knownDates[knownDates.length - 1]}T00:00:00`);
    const loadedSales = new Set(salesDates);
    const loadedInventory = new Set(inventoryDates);
    while (current <= end) {
      const iso = current.toISOString().slice(0, 10);
      const isSunday = current.getDay() === 0;
      dayRows.push({
        iso,
        inventory: isSunday ? "Closed" : (loadedInventory.has(iso) ? "Loaded" : "Missing"),
        sales: isSunday ? "Closed" : (loadedSales.has(iso) ? "Loaded" : "Missing"),
        data: isSunday ? "Closed" : (dataDates.includes(iso) ? "Loaded" : "Missing"),
      });
      current.setDate(current.getDate() + 1);
    }
  }
  const visibleRows = dayRows.filter((row) => {
    if (inventoryFilter && row.inventory !== inventoryFilter) return false;
    if (salesFilter && row.sales !== salesFilter) return false;
    if (dataFilter && row.data !== dataFilter) return false;
    return true;
  }).sort((a, b) => String(b.iso).localeCompare(String(a.iso)));
  const salesCount = dayRows.filter((row) => row.sales === "Loaded").length;
  const inventoryCount = dayRows.filter((row) => row.inventory === "Loaded").length;
  const dataCount = excelLoaded ? 1 : 0;
  const missingDates = dayRows.filter((row) => row.sales === "Missing").map((row) => row.iso);
  els.uploadLogSummary.innerHTML = `
    <span><b>${inventoryCount}</b> inventory CSV days loaded</span>
    <span><b>${salesCount}</b> daily sales days loaded</span>
    <span><b>${dataCount}</b> data file loaded</span>
    <span><b>${missingDates.length}</b> sales days missing</span>
    <span><b>${dayRows.filter((row) => row.sales === "Closed").length}</b> Sundays closed</span>`;
  if (!visibleRows.length) {
    els.uploadLogBody.innerHTML = `<tr><td colspan="4" class="empty-cell">No files imported yet.</td></tr>`;
    return;
  }
  const statusCell = (status) => `<span class="logs-status-chip logs-status-${String(status).toLowerCase()}">${escapeHtml(status)}</span>`;
  els.uploadLogBody.innerHTML = visibleRows.map((entry) => `<tr>
    <td>${escapeHtml(entry.iso)}</td>
    <td>${statusCell(entry.inventory)}</td>
    <td>${statusCell(entry.sales)}</td>
    <td>${statusCell(entry.data)}</td>
  </tr>`).join("");
  repairMojibakeText(document.querySelector("#reportLogsModal") || document.body);
}

function newItemsRows() {
  const codes = new Set([
    ...Object.keys(state.itemMeta || {}),
    ...state.latestInventory.keys(),
    ...[...state.excelItems.values()].map((item) => codeKey(item.code)).filter(Boolean),
  ]);
  return [...codes].map((key) => {
    const inventory = state.latestInventory.get(key) || {};
    const excel = findExcelFor(inventory.code ? inventory : { code: key }) || {};
    const meta = itemMetaFor(inventory.code || excel.code || key);
    const cachedSku = state._skuCache?.get?.(key) || {};
    const item = {
      code: inventory.code || excel.code || key,
      product: bestItemName(inventory.product, excel.product, "", inventory.plu, key),
      vendor: inventory.vendor || excel.vendor || "",
      state: normalizeItemState(meta.state || excel.state || inventory.state || ""),
      addDate: cleanCell(meta.addDate || excel.addDate || inventory.addDate || meta.firstSeenDate || ""),
      category: inventory.category || excel.category || cachedSku.category || "",
      department: cachedSku.department || "",
      plu: inventory.plu || excel.plu || "",
      itemNumber: inventory.itemNumber || excel.itemNumber || "",
      color: inventory.color || "",
    };
    return { ...item, ...parentPartsFor(item) };
  }).filter((item) => item.addDate);
}

function renderNewItems() {
  if (!els.newItemsBody || !els.newItemsMonthFilter || !els.newItemsSummary) return;
  const rows = newItemsRows().filter((item) => /^\d{4}-\d{2}-\d{2}$/.test(item.addDate));
  const yearOptions = [...new Set(rows.map((item) => item.addDate.slice(0, 4)))].sort().reverse();
  fillSelect(els.newItemsYearFilter, yearOptions);
  if (yearOptions.includes(state.newItemsYear)) els.newItemsYearFilter.value = state.newItemsYear;
  else state.newItemsYear = els.newItemsYearFilter?.value || yearOptions[0] || "";
  if (els.newItemsYearFilter) els.newItemsYearFilter.value = state.newItemsYear;
  const rowsForYear = state.newItemsYear ? rows.filter((item) => item.addDate.startsWith(`${state.newItemsYear}-`)) : rows;
  const monthNumbers = [...new Set(rowsForYear.map((item) => item.addDate.slice(5, 7)))].sort((a, b) => Number(a) - Number(b));
  els.newItemsMonthFilter.innerHTML = monthNumbers.map((month) => {
    const labelDate = new Date(2000, Number(month) - 1, 1);
    return `<option value="${month}">${labelDate.toLocaleDateString("en-US", { month: "short" })}</option>`;
  }).join("");
  const currentMonth = els.newItemsMonthFilter.value;
  if (monthNumbers.includes(currentMonth || "")) {
    els.newItemsMonthFilter.value = currentMonth;
  } else if (monthNumbers.length) {
    els.newItemsMonthFilter.value = monthNumbers[monthNumbers.length - 1];
  }
  const monthFilter = els.newItemsMonthFilter.value || "";
  const monthRows = rowsForYear.filter((item) => !monthFilter || item.addDate.slice(5, 7) === monthFilter);
  const vendors = [...new Set(monthRows.map((item) => cleanCell(item.vendor)).filter(Boolean))].sort(compareDisplayValue);
  const departments = [...new Set(monthRows.map((item) => cleanCell(item.department)).filter(Boolean))].sort(compareDisplayValue);
  const categories = [...new Set(monthRows.map((item) => cleanCell(item.category)).filter(Boolean))].sort(compareDisplayValue);
  fillSelect(els.newItemsVendorFilter, vendors);
  fillSelect(els.newItemsDepartmentFilter, departments);
  fillSelect(els.newItemsCategoryFilter, categories);
  if (vendors.includes(state.newItemsVendor)) els.newItemsVendorFilter.value = state.newItemsVendor;
  if (departments.includes(state.newItemsDepartment)) els.newItemsDepartmentFilter.value = state.newItemsDepartment;
  if (categories.includes(state.newItemsCategory)) els.newItemsCategoryFilter.value = state.newItemsCategory;
  if (state.newItemsVendor && !vendors.some((vendor) => vendor.toUpperCase() === state.newItemsVendor.toUpperCase())) state.newItemsVendor = "";
  if (state.newItemsDepartment && !departments.includes(state.newItemsDepartment)) state.newItemsDepartment = "";
  if (state.newItemsCategory && !categories.includes(state.newItemsCategory)) state.newItemsCategory = "";
  const search = normalizeSearchText(state.newItemsSearch || "");
  const dormantMonths = Number(state.newItemsDormantMonths || 0);
  const filteredRows = monthRows.filter((item) => {
    if (state.newItemsVendor && (item.vendor || "").toUpperCase() !== state.newItemsVendor.toUpperCase()) return false;
    if (state.newItemsDepartment && item.department !== state.newItemsDepartment) return false;
    if (state.newItemsCategory && item.category !== state.newItemsCategory) return false;
    if (search) {
      const haystack = normalizeSearchText([item.parent, item.product, item.subType, item.color, item.vendor, item.category, item.code].join(" "));
      if (!haystack.includes(search)) return false;
    }
    return true;
  });
  const groupedRows = [...filteredRows.reduce((map, item) => {
    const key = `${item.vendor}||${item.department}||${item.category}||${item.parent || item.product}`;
    const entry = map.get(key) || {
      vendor: item.vendor,
      department: item.department || "",
      category: item.category || "",
      parent: item.parent || item.product,
      addDate: item.addDate,
      variants: new Set(),
      units: 0,
      sales: 0,
      stock: 0,
      velocity: 0,
      recommendedOrder: 0,
      units30: 0,
      units90: 0,
      units182: 0,
      units365: 0,
    };
    const windows = salesWindowsFor(item.code);
    entry.variants.add(item.subType || item.color || item.itemNumber || item.code || item.product);
    entry.units += toNumber(item.units) || 0;
    entry.sales += toNumber(item.sales) || 0;
    entry.stock += toNumber(item.stock) || 0;
    entry.velocity += toNumber(item.velocity) || 0;
    entry.recommendedOrder += toNumber(item.recommendedOrder) || 0;
    entry.units30 += toNumber(windows.find((entry) => entry.label === "30D")?.units) || 0;
    entry.units90 += toNumber(windows.find((entry) => entry.label === "90D")?.units) || 0;
    entry.units182 += toNumber(windows.find((entry) => entry.label === "6M")?.units) || 0;
    entry.units365 += toNumber(windows.find((entry) => entry.label === "365D")?.units) || 0;
    map.set(key, entry);
    return map;
  }, new Map()).values()]
    .map((entry) => {
      let badge = "New";
      if (entry.units365 <= 0) badge = "Discontinue";
      else if (entry.units182 <= 0) badge = "Slow mover";
      else if (entry.recommendedOrder > 0) badge = "Reorder";
      else if (entry.units30 <= 0 && entry.stock > 0) badge = "Watch";
      return { ...entry, badge };
    })
    .filter((entry) => {
      if (!dormantMonths) return true;
      if (dormantMonths === 1) return entry.units30 <= 0;
      if (dormantMonths === 3) return entry.units90 <= 0;
      if (dormantMonths === 6) return entry.units182 <= 0;
      if (dormantMonths === 12) return entry.units365 <= 0;
      return true;
    });
  const monthLabel = monthFilter ? new Date(2000, Number(monthFilter) - 1, 1).toLocaleDateString("en-US", { month: "short" }) : "-";
  els.newItemsSummary.innerHTML = `
    <span><b>${escapeHtml(monthLabel)}</b> month</span>
    <span><b>${escapeHtml(state.newItemsYear || "-")}</b> year</span>
    <span><b>${vendors.length}</b> vendors</span>
    <span><b>${filteredRows.length}</b> new SKUs</span>
    <span><b>${groupedRows.length}</b> parent groups</span>`;
  if (!groupedRows.length) {
    els.newItemsBody.innerHTML = `<tr><td colspan="10" class="empty-cell">No new items found for that month/filter.</td></tr>`;
    return;
  }
  els.newItemsBody.innerHTML = groupedRows
    .sort((a, b) => compareDisplayValue(a.vendor, b.vendor) || compareDisplayValue(a.parent, b.parent))
    .map((item) => {
      const monthLabel = item.addDate
        ? new Date(Number(item.addDate.slice(0, 4)), Number(item.addDate.slice(5, 7)) - 1, 1).toLocaleDateString("en-US", { month: "short" })
        : "-";
      return `<tr>
      <td>${escapeHtml(item.vendor || "-")}</td>
      <td>${escapeHtml(item.department || "-")}</td>
      <td>${escapeHtml(item.category || "-")}</td>
      <td class="sku-name"><details class="new-items-parent-details"><summary><b>${escapeHtml(item.parent || "-")}</b></summary><div class="new-items-variant-list">${[...item.variants].sort(compareDisplayValue).map((variant) => `<span>${escapeHtml(variant)}</span>`).join("")}</div></details></td>
      <td>${number.format(item.variants.size)}</td>
      <td class="num">${number.format(item.units)}</td>
      <td class="num">${currency.format(item.sales)}</td>
      <td class="num">${number.format(item.stock)}</td>
      <td>${escapeHtml(monthLabel)}</td>
      <td>${escapeHtml(item.badge)}</td>
    </tr>`;
    }).join("");
}

function renderMultiBarcodes() {
  if (!els.multiBody || !els.multiSummary) return;
  const query = normalizeSearchText(state.multiBarcodeSearch || "");
  const rows = (state.multiBarcodeMasters || []).filter((row) => {
    if (!query) return true;
    const haystack = normalizeSearchText([row.masterCode, row.product, row.vendor, row.plu, row.itemNumber, ...(row.aliases || [])].join(" "));
    return haystack.includes(query);
  });
  els.multiSummary.innerHTML = `
    <span><b>${number.format(rows.length)}</b> master items</span>
    <span><b>${number.format((state.multiBarcodeMasters || []).reduce((sum, row) => sum + (row.aliases?.length || 0), 0))}</b> alternate barcodes linked</span>
    <span><b>File</b> ${escapeHtml(state.multiBarcodeFileName || "Not loaded")}</span>`;
  if (!rows.length) {
    els.multiBody.innerHTML = `<tr><td colspan="6" class="empty-cell">Load the multi barcode workbook to link alternate barcodes to the master item code.</td></tr>`;
    return;
  }
  els.multiBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${escapeHtml(row.masterCode)}</td>
      <td class="sku-name" title="${escapeHtml(row.product || "")}">${escapeHtml(row.product || "-")}</td>
      <td>${escapeHtml(row.vendor || "-")}</td>
      <td>${escapeHtml(row.plu || "-")}</td>
      <td>${number.format(row.aliases?.length || 0)}</td>
      <td>${escapeHtml((row.aliases || []).join(", "))}</td>
    </tr>`).join("");
}

function resetUiCriteriaOnStartup() {
  state.tabSearches = { dashboard: "", inventory: "", ordering: "" };
  safeLocalStorageSet("posDashboardTabSearches:v1", JSON.stringify(state.tabSearches));
  if (els.searchInput) els.searchInput.value = "";
  if (els.parentsSearch) els.parentsSearch.value = "";
  if (els.departmentFilter) els.departmentFilter.value = "";
  if (els.categoryFilter) els.categoryFilter.value = "";
  if (els.vendorFilter) els.vendorFilter.value = "";
  if (els.colorFilter) els.colorFilter.value = "";
  if (els.inventoryStateFilter) els.inventoryStateFilter.value = "Active";
  state.productPoReviewVendor = "";
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

function repairMojibakeText(root = document.body) {
  if (!root) return;
  const replacements = [
    [" - ", "·"],
    [" - ", "·"],
    ["->", "->"],
    ["✓", "Submit"],
    ["x", "Close"],
    ["-", "-"],
    ["-", "-"],
    ["-", "-"],
    ["-", "·"],
    ["▼", "▼"],
    ["↑", "↑"],
    ["↓", "↓"],
    ["WARNING:", "Alert:"],
    ["WARNING:", "min"],
    ["Order today:", "Order today:"],
    ["Pending", "Pending"],
    ["Lock", "Auto"],
    ["NULL->0", "NULL -> 0"],
    ["-> 0", "-> 0"],
  ];
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  const nodes = [];
  while (walker.nextNode()) nodes.push(walker.currentNode);
  nodes.forEach((node) => {
    let next = node.nodeValue || "";
    replacements.forEach(([from, to]) => {
      if (next.includes(from)) next = next.split(from).join(to);
    });
    if (next !== node.nodeValue) node.nodeValue = next;
  });
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

async function readDbValue(key) {
  try {
    const db = await openDb();
    return await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readonly");
      const request = tx.objectStore(DB_STORE).get(key);
      request.onsuccess = () => resolve(request.result ?? null);
      request.onerror = () => reject(request.error);
    });
  } catch (_) {
    return null;
  }
}

async function readPersistedState() {
  try {
    const db = await openDb();
    const dbState = await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readonly");
      const request = tx.objectStore(DB_STORE).get(DB_KEY);
      request.onsuccess = () => resolve(request.result || null);
      request.onerror = () => reject(request.error);
    });
    if (dbState) return dbState;
  } catch (_) {
    // Fall back to the tiny local snapshot metadata only.
  }
  return null;
}

async function savePersistedState() {
  const payload = {
    rawSales: state.rawSales,
    inventoryDate: latestInventoryDate(),
    inventoryRows: [...state.latestInventory.values()],
    excelRows: [...state.excelItems.values()],
    loadedFileSignatures: [...state._loadedFileSignatures],
    pendingOrders: state.pendingOrders || [],
    poBackorderCounts: state.poBackorderCounts || {},
  };
  try {
    safeLocalStorageSet(LOCAL_SNAPSHOT_KEY, JSON.stringify({
      savedAt: new Date().toISOString(),
      salesRows: payload.rawSales.length,
      inventoryRows: payload.inventoryRows.length,
      excelRows: payload.excelRows.length,
    }));
  } catch (_) {
    // Ignore quota errors here; IndexedDB is the primary persisted store.
  }
  try {
    const db = await openDb();
    await new Promise((resolve, reject) => {
      const tx = db.transaction(DB_STORE, "readwrite");
      tx.objectStore(DB_STORE).put(payload, DB_KEY);
      tx.oncomplete = () => resolve();
      tx.onerror = () => reject(tx.error);
    });
    await persistItemMetaToDb();
    return;
  } catch (_) {
    return;
  }
}

async function restorePersistedState() {
  const startedAt = performanceNow();
  console.info("[AppPerf] data restore start");
  try {
    setAppLoadStatus("Loading sales");
    const saved = await readPersistedState();
    if (!saved) {
      clearAppLoadStatus("Loading sales");
      return;
    }
    normalizeReorderOverridesState();
    state.rawSales = saved.rawSales || [];
    state.dates = [...new Set(state.rawSales.map((row) => row.date))].sort();
    setAppLoadStatus("Loading inventory");
    state.inventories = new Map();
    if (saved.inventoryDate && saved.inventoryRows?.length) {
      state.inventories.set(saved.inventoryDate, saved.inventoryRows);
    }
    buildLatestInventory();
    state.excelItems = new Map();
    (saved.excelRows || []).forEach((item) => addExcelIndex(item));
    rebuildExcelIndexes();
    if (Array.isArray(saved.pendingOrders)) {
      state.pendingOrders = saved.pendingOrders;
      savePendingOrders();
    }
    if (saved.poBackorderCounts && typeof saved.poBackorderCounts === "object") {
      state.poBackorderCounts = saved.poBackorderCounts;
      savePoBackorderCounts();
    }
    const dbItemMeta = await readDbValue(DB_ITEM_META_KEY);
    let legacyItemMeta = {};
    if (!dbItemMeta) {
      try {
        legacyItemMeta = JSON.parse(localStorage.getItem(ITEM_META_STORAGE_KEY) || "{}");
      } catch (_) {
        legacyItemMeta = {};
      }
    }
    state.itemMeta = dbItemMeta || saved.itemMeta || legacyItemMeta || {};
    if (!dbItemMeta && Object.keys(state.itemMeta || {}).length) saveItemMeta();
    state._loadedFileSignatures = new Set(saved.loadedFileSignatures || []);
    if (els.excelStatus) {
      els.excelStatus.textContent = saved.excelRows?.length
        ? `${number.format(saved.excelRows.length)} Excel items loaded from local app data.`
        : "Excel not loaded.";
    }
    bumpDataStamp();
    if (scanModeIsActive()) {
      state._deferredDashboardOptions = true;
    } else {
      updateFilterOptions();
      updateInventoryStateFilter();
      setDefaultDates();
    }
    clearAppLoadStatus("Loading inventory");
  } catch (error) {
    setAppLoadStatus("");
    console.warn("Could not restore POS dashboard history", error);
  } finally {
    logPerformanceEnd("data restore", startedAt, {
      inventory: state.latestInventory.size,
      sales: state.rawSales.length,
    });
  }
}

function sharedSnapshotIsOlderThanCurrent(productRows = [], salesRows = []) {
  const currentSalesRows = state.rawSales.length;
  const currentSalesDays = state.dates.length;
  const currentProductRows = Math.max(state.latestInventory.size, state.excelItems.size);
  const sharedSalesDays = new Set(salesRows.map((row) => row.sales_date).filter(Boolean)).size;
  const sharedSalesCount = salesRows.length;
  const sharedProductCount = productRows.length;
  const currentLatestSales = state.dates.at(-1) || "";
  const sharedLatestSales = [...new Set(salesRows.map((row) => row.sales_date).filter(Boolean))].sort().at(-1) || "";
  const currentLatestInventory = latestInventoryDate() || latestExcelAddDate() || "";
  const sharedLatestInventory = productRows
    .map((row) => cleanCell(row.updated_at || row.add_date))
    .filter(Boolean)
    .map((value) => String(value).slice(0, 10))
    .sort()
    .at(-1) || "";

  const currentHasData = currentSalesRows || currentProductRows;
  if (!currentHasData) return false;
  if (currentSalesDays > sharedSalesDays) return true;
  if (currentSalesRows > sharedSalesCount) return true;
  if (currentProductRows > sharedProductCount) return true;
  if (currentLatestSales && sharedLatestSales && currentLatestSales > sharedLatestSales) return true;
  if (currentLatestInventory && sharedLatestInventory && currentLatestInventory > sharedLatestInventory) return true;
  return false;
}

function supabaseRestUrl(tableName) {
  return `${SUPABASE_URL}/rest/v1/${tableName}`;
}

async function supabaseSelectRows(tableName, query = {}) {
  const requestedLimit = Math.max(0, Math.round(toNumber(query.limit) || 0));
  const pageSize = requestedLimit ? Math.min(1000, requestedLimit) : 1000;
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
    if (requestedLimit && rows.length >= requestedLimit) break;
    if (page.length < pageSize) break;
    offset += pageSize;
  }
  return requestedLimit ? rows.slice(0, requestedLimit) : rows;
}

async function supabaseSelectRowsSafe(tableName, query = {}) {
  try {
    return await supabaseSelectRows(tableName, query);
  } catch (_) {
    return [];
  }
}

function supabaseDeleteAllFilterColumn(tableName) {
  // Different shared tables do not all have an id column. Using id=not.is.null
  // made full imports fail silently/partially on products and daily_sales, which left
  // other devices on old local snapshots. Delete against a known required column.
  const filters = {
    products: "code",
    product_meta: "code",
    daily_sales: "sales_date",
    vendor_rules: "id",
    import_logs: "id",
    count_sessions: "id",
    sync_state: "id",
    app_users: "id",
  };
  return filters[tableName] || "id";
}

async function supabaseDeleteAllRows(tableName) {
  const url = new URL(supabaseRestUrl(tableName));
  url.searchParams.set(supabaseDeleteAllFilterColumn(tableName), "not.is.null");
  const response = await fetch(url.toString(), {
    method: "DELETE",
    headers: supabaseHeaders(),
  });
  if (!response.ok) {
    let details = "";
    try { details = await response.text(); } catch (_) {}
    throw new Error(`${tableName} delete failed (${response.status}) ${details}`.trim());
  }
}

async function supabaseDeleteRowsByValues(tableName, columnName, values) {
  const uniqueValues = [...new Set((values || []).map((value) => cleanCell(value)).filter(Boolean))];
  if (!uniqueValues.length) return;
  const chunkSize = 40;
  for (let i = 0; i < uniqueValues.length; i += chunkSize) {
    const chunk = uniqueValues.slice(i, i + chunkSize).map((value) => `"${String(value).replace(/"/g, '\\"')}"`);
    const url = new URL(supabaseRestUrl(tableName));
    url.searchParams.set(columnName, `in.(${chunk.join(",")})`);
    const response = await fetch(url.toString(), {
      method: "DELETE",
      headers: supabaseHeaders(),
    });
    if (!response.ok) throw new Error(`${tableName} filtered delete failed (${response.status})`);
  }
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

async function supabaseUpsertRows(tableName, rows, onConflict = "code") {
  if (!rows.length) return;
  const chunkSize = 500;
  for (let i = 0; i < rows.length; i += chunkSize) {
    const url = new URL(supabaseRestUrl(tableName));
    url.searchParams.set("on_conflict", onConflict);
    const headers = { ...supabaseHeaders(true), Prefer: "resolution=merge-duplicates,return=minimal" };
    const response = await fetch(url.toString(), { method: "POST", headers, body: JSON.stringify(rows.slice(i, i + chunkSize)) });
    if (!response.ok) throw new Error(`${tableName} upsert failed (${response.status})`);
  }
}

function productRowForSupabase(item) {
  const excel = findExcelFor(item);
  const normalizedCode = normalizeCode(item.code || excel.code || "");
  const meta = itemMetaFor(normalizedCode);
  const syncState = normalizeItemState(meta.stateManual ? meta.state : (meta.state || excel.state || item.state || ""));
  const syncCaseSize = meta.caseSizeManual
    ? Math.max(1, Math.round(toNumber(meta.caseSize) || 1))
    : Math.max(1, Math.round(toNumber(excel.caseSize || item.caseSize || meta.caseSize) || 1));
  const syncAddDate = normalizeItemDate(meta.addDate || excel.addDate || item.addDate || meta.firstSeenDate || "");
  return {
    code: normalizedCode,
    product: item.product || excel.product || "",
    plu: item.plu || excel.plu || "",
    item_number: item.itemNumber || excel.itemNumber || "",
    vendor: item.vendor || excel.vendor || "",
    category: item.category || excel.category || "",
    department: item.department || excel.department || "",
    color: item.color || excel.color || "",
    state: syncState || "",
    stock: Number(item.stock ?? excel.stock ?? 0),
    price: Number(item.price ?? excel.price ?? 0),
    unit_cost: Number(item.unitCost ?? item.cost ?? excel.cost ?? 0),
    case_size: Number(syncCaseSize || 1),
    add_date: syncAddDate || null,
    updated_at: new Date().toISOString(),
  };
}

function productRowFromExcelForSupabase(item = {}) {
  const normalizedCode = normalizeCode(item.code || "");
  const meta = itemMetaFor(normalizedCode);
  const syncState = normalizeItemState(meta.stateManual ? meta.state : (meta.state || item.state || ""));
  const syncCaseSize = effectiveCaseSizeFor({}, item, meta);
  const syncAddDate = normalizeItemDate(meta.addDate || item.addDate || meta.firstSeenDate || "");
  return {
    code: normalizedCode,
    product: item.product || "",
    plu: item.plu || "",
    item_number: item.itemNumber || "",
    vendor: item.vendor || "",
    category: item.category || "",
    department: item.department || "",
    color: item.color || "",
    state: syncState || "",
    stock: Number(item.stock ?? 0),
    price: Number(item.price ?? 0),
    unit_cost: Number(item.cost ?? 0),
    case_size: Number(syncCaseSize || 1),
    add_date: syncAddDate || null,
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

function vendorRuleRowForSupabase(rule = {}) {
  const vendorName = cleanCell(rule.vendor || "").toUpperCase();
  return {
    id: stableVendorRuleId(vendorName) || cleanCell(rule.id) || `vr-${Date.now()}`,
    vendor: vendorName,
    status: cleanCell(rule.status) || "Active",
    safety_days: Math.max(0, toNumber(rule.safetyDays) || 0),
    days_of_inventory: Math.max(0, toNumber(rule.daysOfInventory) || 0),
    min_order: Math.max(0, toNumber(rule.minOrder) || 0),
    email: cleanCell(rule.email),
    notes: cleanCell(rule.notes),
    order_days: Array.isArray(rule.orderDays) ? rule.orderDays : [],
    updated_at: rule.updatedAt || new Date().toISOString(),
  };
}

function importLogRowForSupabase(entry = {}) {
  return {
    id: cleanCell(entry.id) || `upl-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    recorded_at: entry.recordedAt || new Date().toISOString(),
    filename: cleanCell(entry.filename) || "-",
    type: cleanCell(entry.type) || "File",
    status: cleanCell(entry.status) || "Success",
    updated_at: entry.recordedAt || new Date().toISOString(),
  };
}

function hydrateImportLogFromSupabase(row = {}) {
  return {
    id: cleanCell(row.id) || `upl-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    recordedAt: row.recorded_at || row.updated_at || new Date().toISOString(),
    filename: cleanCell(row.filename) || "-",
    type: cleanCell(row.type) || "File",
    status: cleanCell(row.status) || "Success",
  };
}

function countSessionRowForSupabase(session = {}, active = false) {
  return {
    id: active ? `active:${cleanCell(session.id) || "session"}` : cleanCell(session.id),
    session_kind: active ? "active" : "saved",
    session_payload: countSessionForRemote(session),
    updated_at: session.updatedAt || session.submittedAt || new Date().toISOString(),
  };
}

function hydrateCountSessionFromSupabase(row = {}) {
  if (!row.session_payload || typeof row.session_payload !== "object") return null;
  return normalizeCountSession({
    ...row.session_payload,
    updatedAt: row.session_payload.updatedAt || row.updated_at || new Date().toISOString(),
    localSyncPending: false,
  }, "synced");
}

function hydrateVendorRuleFromSupabase(row = {}) {
  return createVendorRule(row.vendor || "", {
    id: row.id,
    status: cleanCell(row.status) || "Active",
    safetyDays: Math.max(0, toNumber(row.safety_days) || 0),
    daysOfInventory: Math.max(0, toNumber(row.days_of_inventory) || 0),
    minOrder: Math.max(0, toNumber(row.min_order) || 0),
    email: cleanCell(row.email),
    notes: cleanCell(row.notes),
    orderDays: Array.isArray(row.order_days) ? row.order_days : [],
    updatedAt: row.updated_at || new Date().toISOString(),
  });
}

function buildProductMetaRowsSnapshot() {
  const productMetaRowMap = new Map();
  [...state.excelItems.values()]
    .filter((item) => item?.code || item?.product || item?.plu || item?.itemNumber)
    .forEach((item) => {
      const row = productMetaRowForSupabase(item);
      const key = codeKey(row.code);
      if (key) productMetaRowMap.set(key, row);
    });
  [...state.latestInventory.values()]
    .filter((item) => item.code || item.product)
    .forEach((item) => {
      const row = productMetaRowForSupabase(item);
      const key = codeKey(row.code);
      if (key) productMetaRowMap.set(key, row);
    });
  return [...productMetaRowMap.values()].filter((row) => cleanCell(row.code));
}

function vendorRuleSyncKey(rule = {}) {
  return cleanCell(rule.vendor || "").toUpperCase() || cleanCell(rule.id || "");
}

function vendorRuleUpdatedMs(rule = {}) {
  return Date.parse(rule.updatedAt || rule.updated_at || "") || 0;
}

function applySharedVendorRuleRows(rows = []) {
  if (!rows?.length) return 0;
  const localByVendor = new Map(
    (state.vendorRules || [])
      .filter((rule) => rule?.vendor)
      .map((rule) => [vendorRuleSyncKey(rule), rule])
  );
  const mergedByVendor = new Map();
  rows
    .map(hydrateVendorRuleFromSupabase)
    .filter((rule) => rule.vendor)
    .forEach((remoteRule) => {
      const key = vendorRuleSyncKey(remoteRule);
      const localRule = localByVendor.get(key);
      const localMs = vendorRuleUpdatedMs(localRule);
      const remoteMs = vendorRuleUpdatedMs(remoteRule);
      mergedByVendor.set(key, localRule && localMs > remoteMs + 1000 ? localRule : remoteRule);
    });
  localByVendor.forEach((localRule, key) => {
    if (!mergedByVendor.has(key)) mergedByVendor.set(key, localRule);
  });
  state.vendorRules = [...mergedByVendor.values()]
    .filter((rule) => rule.vendor)
    .sort((a, b) => compareDisplayValue(a.vendor, b.vendor));
  safeLocalStorageSet("posDashboardVendorRules:v1", JSON.stringify(state.vendorRules));
  bumpDataStamp();
  if (!scanModeIsActive()) {
    renderVendorRules();
    queueActiveTabRender();
  }
  return state.vendorRules.length;
}

function productMetaRowForSupabase(source = {}) {
  const excel = findExcelFor(source);
  const normalizedCode = normalizeCode(source.code || excel.code || "");
  const sku = state._skuCache?.get?.(codeKey(normalizedCode)) || {};
  const meta = itemMetaFor(normalizedCode);
  const override = reorderOverrideForCode(normalizedCode);
  const stateValue = normalizeItemState(meta.stateManual ? meta.state : (meta.state || source.state || excel.state || sku.state || ""));
  const caseSize = effectiveCaseSizeFor(source, excel, meta, sku);
  const addDate = normalizeItemDate(meta.addDate || source.addDate || excel.addDate || meta.firstSeenDate || "");
  return {
    code: normalizedCode,
    product: source.product || excel.product || sku.product || "",
    plu: source.plu || excel.plu || sku.plu || "",
    item_number: source.itemNumber || excel.itemNumber || sku.itemNumber || "",
    vendor: source.vendor || excel.vendor || sku.vendor || "",
    category: source.category || excel.category || sku.category || "",
    department: source.department || excel.department || sku.department || "",
    color: source.color || excel.color || sku.color || "",
    state: stateValue || "",
    stock: Number(source.stock ?? excel.stock ?? sku.stock ?? 0),
    price: Number(source.price ?? excel.price ?? sku.price ?? 0),
    unit_cost: Number(source.unitCost ?? source.cost ?? excel.cost ?? sku.unitCost ?? 0),
    sold_units: Number(source.units ?? sku.units ?? 0),
    sale_velocity: Number(source.velocity ?? sku.velocity ?? 0),
    case_size: Number(caseSize || 1),
    add_date: addDate || null,
    ordering_rule_mode: cleanCell(meta.orderingRuleMode) || "vendor",
    safety_days_override: meta.safetyDaysOverride != null && meta.safetyDaysOverride !== "" ? Number(meta.safetyDaysOverride) : null,
    days_of_inventory_override: meta.daysOfInventoryOverride != null && meta.daysOfInventoryOverride !== "" ? Number(meta.daysOfInventoryOverride) : null,
    reorder_min_override: override.min != null ? Number(override.min) : null,
    reorder_max_override: override.max != null ? Number(override.max) : null,
    snapshot_date: latestInventoryDate() || null,
    updated_at: new Date().toISOString(),
  };
}

function mergeSharedProductRows(productRows = [], productMetaRows = []) {
  const merged = new Map();
  productRows.forEach((row) => {
    const code = normalizeCode(row.code);
    if (!code) return;
    merged.set(codeKey(code), { ...row, code });
  });
  productMetaRows.forEach((row) => {
    const code = normalizeCode(row.code);
    if (!code) return;
    const key = codeKey(code);
    const base = merged.get(key) || { code };
    const hasOrderingRuleMode = Object.prototype.hasOwnProperty.call(row, "ordering_rule_mode");
    const hasSafetyOverride = Object.prototype.hasOwnProperty.call(row, "safety_days_override");
    const hasDoiOverride = Object.prototype.hasOwnProperty.call(row, "days_of_inventory_override");
    const hasMinOverride = Object.prototype.hasOwnProperty.call(row, "reorder_min_override");
    const hasMaxOverride = Object.prototype.hasOwnProperty.call(row, "reorder_max_override");
    merged.set(key, {
      ...base,
      code,
      product: row.product ?? base.product ?? "",
      plu: row.plu ?? base.plu ?? "",
      item_number: row.item_number ?? base.item_number ?? "",
      vendor: row.vendor ?? base.vendor ?? "",
      category: row.category ?? base.category ?? "",
      department: row.department ?? base.department ?? "",
      color: row.color ?? base.color ?? "",
      state: row.state ?? base.state ?? "",
      stock: row.stock ?? base.stock ?? 0,
      price: row.price ?? base.price ?? 0,
      unit_cost: row.unit_cost ?? base.unit_cost ?? 0,
      case_size: row.case_size ?? base.case_size ?? 1,
      add_date: row.add_date ?? base.add_date ?? null,
      ordering_rule_mode: hasOrderingRuleMode ? row.ordering_rule_mode : (base.ordering_rule_mode ?? "vendor"),
      safety_days_override: hasSafetyOverride ? row.safety_days_override : (base.safety_days_override ?? null),
      days_of_inventory_override: hasDoiOverride ? row.days_of_inventory_override : (base.days_of_inventory_override ?? null),
      reorder_min_override: hasMinOverride ? row.reorder_min_override : (base.reorder_min_override ?? null),
      reorder_max_override: hasMaxOverride ? row.reorder_max_override : (base.reorder_max_override ?? null),
      updated_at: row.updated_at ?? base.updated_at ?? null,
      snapshot_date: row.snapshot_date ?? base.snapshot_date ?? null,
      sold_units: row.sold_units ?? base.sold_units ?? 0,
      sale_velocity: row.sale_velocity ?? base.sale_velocity ?? 0,
    });
  });
  return [...merged.values()];
}

function hydrateInventoryFromSupabase(row) {
  return {
    date: cleanCell(row.snapshot_date) || (row.updated_at ? String(row.updated_at).slice(0, 10) : new Date().toISOString().slice(0, 10)),
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
    saleVelocity: row.sale_velocity,
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


function refreshInventorySnapshotFromLatestInventory() {
  const rows = [...state.latestInventory.values()];
  if (!rows.length) return;
  const snapshotDate = latestInventoryDate()
    || rows.map((row) => row.date || "").filter(Boolean).sort().at(-1)
    || new Date().toISOString().slice(0, 10);
  state.inventories.set(snapshotDate, rows);
  state._inventoryCache = null;
  state._inventoryCacheStamp = "";
}

function applySharedProductRows(rows) {
  if (!rows?.length) return 0;
  const snapshotDate = rows.map((row) => row.updated_at ? String(row.updated_at).slice(0, 10) : "").filter(Boolean).sort().at(-1)
    || latestInventoryDate()
    || new Date().toISOString().slice(0, 10);
  const mergedInventory = new Map([...state.latestInventory.entries()]);
  const nextOverrides = { ...(state.reorderOverrides || {}) };
  state.itemMeta = state.itemMeta || {};
  rows.forEach((row) => {
    const inventoryItem = hydrateInventoryFromSupabase(row);
    const excelItem = hydrateExcelFromSupabase(row);
    const key = codeKey(inventoryItem.code || excelItem.code);
    if (!key) return;
    mergedInventory.set(key, inventoryItem);
    addExcelIndex(excelItem);
    const meta = state.itemMeta[key] || {};
    state.itemMeta[key] = {
      ...meta,
      ...(row.state != null ? { state: cleanCell(row.state), stateManual: true } : {}),
      ...(row.case_size != null ? { caseSize: normalizedCaseSize(row.case_size), caseSizeManual: meta.caseSizeManual === true } : {}),
      ...(row.add_date ? { addDate: normalizeItemDate(row.add_date) } : {}),
      orderingRuleMode: cleanCell(row.ordering_rule_mode) || meta.orderingRuleMode || "vendor",
      safetyDaysOverride: row.safety_days_override != null ? Math.max(0, Math.round(toNumber(row.safety_days_override) || 0)) : null,
      daysOfInventoryOverride: row.days_of_inventory_override != null ? Math.max(0, Math.round(toNumber(row.days_of_inventory_override) || 0)) : null,
    };
    const overrideMin = row.reorder_min_override;
    const overrideMax = row.reorder_max_override;
    if (overrideMin != null || overrideMax != null) {
      nextOverrides[key] = {
        ...(overrideMin != null ? { min: toNumber(overrideMin) } : {}),
        ...(overrideMax != null ? { max: toNumber(overrideMax) } : {}),
      };
    } else {
      delete nextOverrides[key];
    }
  });
  state.latestInventory = mergedInventory;
  saveItemMeta();
  state.reorderOverrides = nextOverrides;
  safeLocalStorageSet("posDashboardReorderOverrides:v1", JSON.stringify(state.reorderOverrides));
  state.inventories.set(snapshotDate, [...mergedInventory.values()]);
  rebuildExcelIndexes();
  bumpDataStamp();
  return rows.length;
}

function applySharedProductMetaRowsLight(rows = []) {
  if (!rows?.length) return [];
  const changedCodes = new Set();
  const nextOverrides = { ...(state.reorderOverrides || {}) };
  state.itemMeta = state.itemMeta || {};
  rows.forEach((row) => {
    const code = normalizeCode(row.code);
    const key = codeKey(code);
    if (!key) return;
    const localEditMs = state._localProductMetaEdits?.get?.(key) || 0;
    const remoteEditMs = Date.parse(row.updated_at || row.snapshot_date || "") || 0;
    if (localEditMs && (!remoteEditMs || remoteEditMs < localEditMs - 1000)) return;
    if (localEditMs && remoteEditMs >= localEditMs - 1000) state._localProductMetaEdits.delete(key);
    const beforeInventory = state.latestInventory.get(key) || {};
    const beforeMeta = state.itemMeta[key] || {};
    const beforeOverride = reorderOverrideForCode(key);
    const nextInventory = { ...beforeInventory };
    if (!nextInventory.code) nextInventory.code = code;
    if (row.product != null) nextInventory.product = cleanCell(row.product);
    if (row.plu != null) nextInventory.plu = cleanCell(row.plu);
    if (row.item_number != null) nextInventory.itemNumber = cleanCell(row.item_number);
    if (row.vendor != null) nextInventory.vendor = cleanCell(row.vendor);
    if (row.category != null) nextInventory.category = cleanCell(row.category);
    if (row.department != null) nextInventory.department = cleanCell(row.department);
    if (row.color != null) nextInventory.color = cleanCell(row.color);
    if (row.stock != null) nextInventory.stock = toNumber(row.stock);
    if (row.price != null) nextInventory.price = toNumber(row.price);
    if (row.unit_cost != null) nextInventory.cost = toNumber(row.unit_cost);
    if (row.snapshot_date != null || row.updated_at != null) {
      nextInventory.date = cleanCell(row.snapshot_date) || (row.updated_at ? String(row.updated_at).slice(0, 10) : nextInventory.date);
    }
    const nextMeta = {
      ...beforeMeta,
      ...(row.state != null ? { state: normalizeItemState(row.state), stateManual: true } : {}),
      ...(row.case_size != null ? { caseSize: normalizedCaseSize(row.case_size), caseSizeManual: beforeMeta.caseSizeManual === true } : {}),
      ...(row.add_date ? { addDate: normalizeItemDate(row.add_date) } : {}),
      orderingRuleMode: cleanCell(row.ordering_rule_mode) || beforeMeta.orderingRuleMode || "vendor",
      safetyDaysOverride: row.safety_days_override != null ? Math.max(0, Math.round(toNumber(row.safety_days_override) || 0)) : null,
      daysOfInventoryOverride: row.days_of_inventory_override != null ? Math.max(0, Math.round(toNumber(row.days_of_inventory_override) || 0)) : null,
    };
    if (row.reorder_min_override != null || row.reorder_max_override != null) {
      nextOverrides[key] = {
        ...(row.reorder_min_override != null ? { min: toNumber(row.reorder_min_override) } : {}),
        ...(row.reorder_max_override != null ? { max: toNumber(row.reorder_max_override) } : {}),
      };
    } else {
      delete nextOverrides[key];
    }
    const nextOverride = nextOverrides[key] || {};
    const inventoryChanged = JSON.stringify([
      beforeInventory.stock,
      beforeInventory.price,
      beforeInventory.cost,
      beforeInventory.vendor,
      beforeInventory.product,
      beforeInventory.category,
      beforeInventory.department,
    ]) !== JSON.stringify([
      nextInventory.stock,
      nextInventory.price,
      nextInventory.cost,
      nextInventory.vendor,
      nextInventory.product,
      nextInventory.category,
      nextInventory.department,
    ]);
    const metaChanged = JSON.stringify(beforeMeta) !== JSON.stringify(nextMeta);
    const overrideChanged = JSON.stringify(beforeOverride) !== JSON.stringify(nextOverride);
    if (inventoryChanged || metaChanged || overrideChanged) {
      state.latestInventory.set(key, nextInventory);
      state.itemMeta[key] = nextMeta;
      changedCodes.add(key);
    }
  });
  if (!changedCodes.size) return [];
  state.reorderOverrides = nextOverrides;
  try {
    safeLocalStorageSet("posDashboardReorderOverrides:v1", JSON.stringify(state.reorderOverrides));
  } catch (_) {}
  saveItemMeta();
  refreshInventorySnapshotFromLatestInventory();
  bumpDataStamp();
  return [...changedCodes];
}

function applySharedImportLogRows(rows = []) {
  if (!rows?.length) return 0;
  const merged = new Map((state.uploadLogs || [])
    .filter((entry) => entry?.id)
    .map((entry) => [entry.id, entry]));
  rows
    .map(hydrateImportLogFromSupabase)
    .filter((entry) => entry.id)
    .forEach((entry) => merged.set(entry.id, entry));
  state.uploadLogs = [...merged.values()]
    .sort((a, b) => String(b.recordedAt || "").localeCompare(String(a.recordedAt || "")))
    .slice(0, 5000);
  saveUploadLogs();
  renderUploadLogs();
  return state.uploadLogs.length;
}

function applySharedCountSessionRows(rows = [], options = {}) {
  pruneDeletedCountSessions();
  // PATCH B: capture a hash of current count state so we can skip render if nothing changed.
  const _preApplyHash = [
    state.activeCountSession?.id || "",
    state.activeCountSession?.updatedAt || "",
    (state.countSessions || []).map((s) => `${s.id}:${s.updatedAt}`).join(","),
  ].join("|");
  const localSavedById = new Map((state.countSessions || [])
    .filter((session) => cleanCell(session?.id) && !countSessionIsDeleted(session.id))
    .map((session) => [cleanCell(session.id), normalizeCountSession(session, countSessionHasUnsyncedEdits(session) ? "pending" : "synced")]));
  let localActive = state.activeCountSession
    ? normalizeCountSession(state.activeCountSession, countSessionHasUnsyncedEdits(state.activeCountSession) ? "pending" : "synced")
    : null;
  if (!rows?.length) {
    const keepContinuedActive = localActive && cleanCell(localActive.id) === cleanCell(state._continuingCountId || "");
    state.countSessions = [...localSavedById.values()].filter((session) =>
      countSessionHasUnsyncedEdits(session) || cleanCell(session?.id) === cleanCell(state._continuingCountId || "")
    );
    state.activeCountSession = (countSessionHasUnsyncedEdits(localActive) || keepContinuedActive) ? localActive : null;
    // V2.0: preserve the active session across sync, but do NOT force the view open —
    // visibility is owned by explicit user action (_countSessionOpen).
    state._countRemoteLoaded = true;
    persistCountSessions({ scheduleSync: false });
    if (options.render !== false && activeTabName() === "counts") {
      const _postHash = [state.activeCountSession?.id || "", state.activeCountSession?.updatedAt || "", (state.countSessions || []).map((s) => `${s.id}:${s.updatedAt}`).join(",")].join("|");
      if (_postHash !== _preApplyHash) renderCountsWorkspace();
    }
    return state.countSessions.length + (state.activeCountSession ? 1 : 0);
  }
  const remoteSavedById = new Map();
  let remoteActive = null;
  rows.forEach((row) => {
    const session = hydrateCountSessionFromSupabase(row);
    if (!session) return;
    const sessionId = cleanCell(session.id);
    if (!sessionId || countSessionIsDeleted(sessionId)) return;
    if (cleanCell(row.session_kind).toLowerCase() === "active") {
      if (!remoteActive || countSessionUpdatedMs(session) >= countSessionUpdatedMs(remoteActive)) remoteActive = session;
    } else {
      const previous = remoteSavedById.get(sessionId);
      if (!previous || countSessionUpdatedMs(session) >= countSessionUpdatedMs(previous)) remoteSavedById.set(sessionId, session);
    }
  });
  if (remoteActive?.id) {
    const submittedCopy = remoteSavedById.get(cleanCell(remoteActive.id));
    if (submittedCopy && (submittedCopy.savedAt || submittedCopy.submittedAt)) {
      const staleActiveId = `active:${cleanCell(remoteActive.id)}`;
      remoteActive = null;
      if (ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
        supabaseDeleteRowsByValues("count_sessions", "id", [staleActiveId]).catch(() => {});
      }
    }
  }
  const mergedSavedById = new Map();
  remoteSavedById.forEach((remoteSession, id) => {
    mergedSavedById.set(id, mergeCountSessions(remoteSession, localSavedById.get(id)));
  });
  localSavedById.forEach((localSession, id) => {
    if (!mergedSavedById.has(id) && countSessionHasUnsyncedEdits(localSession)) mergedSavedById.set(id, localSession);
  });
  const remoteActiveId = cleanCell(remoteActive?.id);
  const localActiveId = cleanCell(localActive?.id);
  const submittedRemoteForLocalActive = localActiveId ? remoteSavedById.get(localActiveId) : null;
  if (localActive && submittedRemoteForLocalActive?.submittedAt) {
    const remoteEntryIds = new Set((submittedRemoteForLocalActive.entries || []).map((entry) => entry.entryId));
    const hasLateLocalEntries = (localActive.entries || []).some((entry) => !remoteEntryIds.has(entry.entryId));
    if (hasLateLocalEntries) {
      const mergedLateSession = mergeCountSessions(submittedRemoteForLocalActive, localActive);
      mergedLateSession.lateEntriesAfterSubmit = true;
      mergedLateSession.localSyncPending = true;
      mergedSavedById.set(localActiveId, mergedLateSession);
    }
    localActive = null;
  }
  let nextActive = remoteActive && remoteActiveCountSessionIsFresh({ ...remoteActive, remoteActive: true })
    ? mergeCountSessions(remoteActive, remoteActiveId === localActiveId ? localActive : null)
    : null;
  // FIX-B: If the local device has already saved or submitted this session, the remote
  // active:${id} row is stale. Never let it become nextActive — that would undo the save.
  // Schedule a background delete of the stale remote row so other devices also converge.
  if (nextActive?.id) {
    const localSavedCopy = localSavedById.get(cleanCell(nextActive.id));
    if (localSavedCopy?.savedAt || localSavedCopy?.submittedAt) {
      if (ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
        supabaseDeleteRowsByValues("count_sessions", "id", [`active:${cleanCell(nextActive.id)}`]).catch(() => {});
      }
      nextActive = null;
    }
  }
  // Continue Count is an explicit operator action. While that session is being continued,
  // keep the local active workspace open even if the remote pull only returns the saved
  // history row and no active:${id} row yet. This prevents the modal from closing after
  // a delayed background sync/poll.
  const continuingId = cleanCell(state._continuingCountId || "");
  const keepContinuedActive = localActive && continuingId && cleanCell(localActive.id) === continuingId
    && !localActive.savedAt && !localActive.submittedAt;
  if ((keepContinuedActive || countSessionHasUnsyncedEdits(localActive)) && (!nextActive || localActiveId !== remoteActiveId)) {
    if (nextActive?.id && nextActive.id !== localActive.id) mergedSavedById.set(nextActive.id, nextActive);
    nextActive = { ...localActive, isActiveLive: true };
    // V2.0: keep the session active across sync, but do NOT force the workspace open.
  }
  // FIX-C: Do NOT remove nextActive from mergedSavedById.
  // Previously this deleted the active session from the sessions list entirely, making it
  // invisible in Report History. Instead keep it in countSessions[] with isActiveLive:true
  // so history can show "Active / Live" status while it's still in progress.
  if (nextActive?.id) {
    mergedSavedById.set(cleanCell(nextActive.id), { ...nextActive, isActiveLive: true });
  }
  state.countSessions = canonicalCountSessions([...mergedSavedById.values()])
    .filter((session) => !countSessionIsDeleted(session.id))
    .sort((a, b) => String(b.updatedAt || b.submittedAt || "").localeCompare(String(a.updatedAt || a.submittedAt || "")));
  state.activeCountSession = nextActive;
  state._countRemoteLoaded = true;
  state._countSyncUnavailable = false;
  persistCountSessions({ scheduleSync: false });
  if (nextActive && countSessionHasUnsyncedEdits(nextActive)) scheduleSharedCountSessionsSync();
  if (options.render !== false && activeTabName() === "counts" && !state._countSyncLoading) {
    const _postHash = [state.activeCountSession?.id || "", state.activeCountSession?.updatedAt || "", (state.countSessions || []).map((s) => `${s.id}:${s.updatedAt}`).join(",")].join("|");
    if (_postHash !== _preApplyHash) renderCountsWorkspace();
  }
  return state.countSessions.length + (state.activeCountSession ? 1 : 0);
}

async function restoreSharedProductsOnlyFromSupabase(options = {}) {
  if (!ENABLE_SHARED_SYNC) return false;
  const { silent = false } = options;
  try {
    // The GKPOS PowerShell bridge writes live inventory facts to `products` while
    // the app keeps operator-managed ordering fields in `product_meta`. Pull and
    // merge both tables so a CSV delta can update stock/name/price without
    // overwriting case-size rules, min/max overrides, or other app edits.
    const [productRows, productMetaRows] = await Promise.all([
      supabaseSelectRowsSafe("products", {
        select: "*",
        order: "updated_at.desc",
      }),
      supabaseSelectRowsSafe("product_meta", {
        select: "*",
        order: "updated_at.desc",
      }),
    ]);
    if (!productMetaRows.length && SHARED_PRODUCT_META_PULL_LIMIT) {
      const limitedMetaRows = await supabaseSelectRowsSafe("product_meta", {
        select: "*",
        order: "updated_at.desc",
        limit: String(SHARED_PRODUCT_META_PULL_LIMIT),
      });
      productMetaRows.push(...limitedMetaRows);
    }
    const liveProductsByCode = new Map(productRows.map((row) => [codeKey(row.code), row]));
    const mergedRows = mergeSharedProductRows(productRows, productMetaRows).map((row) => {
      const live = liveProductsByCode.get(codeKey(row.code));
      if (!live) return row;
      // GKPOS is authoritative for live product/inventory facts. Ordering-rule
      // overrides still come from product_meta through mergeSharedProductRows().
      return {
        ...row,
        product: live.product ?? row.product,
        plu: live.plu ?? row.plu,
        item_number: live.item_number ?? row.item_number,
        vendor: live.vendor ?? row.vendor,
        category: live.category ?? row.category,
        department: live.department ?? row.department,
        color: live.color ?? row.color,
        state: live.state ?? row.state,
        stock: live.stock ?? row.stock,
        price: live.price ?? row.price,
        unit_cost: live.unit_cost ?? row.unit_cost,
        case_size: live.case_size ?? row.case_size,
        add_date: live.add_date ?? row.add_date,
        snapshot_date: live.snapshot_date ?? row.snapshot_date,
        updated_at: live.updated_at ?? row.updated_at,
      };
    });
    const changedCodes = applySharedProductMetaRowsLight(mergedRows);
    if (!changedCodes.length) return false;
    if (!state.activeCountSession) {
      if (activeTabName() === "inventory") {
        // Full Products rerender keeps header/body rows, summary chips, filtered rows,
        // and stock cells from using an old local table snapshot.
        renderInventory();
        refreshDetailDrawer();
      } else if (activeTabName() === "ordering") {
        renderOrders();
      }
    }
    if (!silent) showToast(`Shared product updates loaded (${number.format(changedCodes.length)} items)`, 1800, "success");
    return true;
  } catch (error) {
    if (!silent) showToast("Shared product refresh failed.", 2600, "warning");
    return false;
  }
}

async function restoreSharedVendorRulesOnlyFromSupabase(options = {}) {
  if (!ENABLE_SHARED_SYNC || !sharedVendorRulesAvailable) return false;
  const { silent = false } = options;
  try {
    const vendorRuleRows = await supabaseSelectRowsSafe("vendor_rules", { select: "*" });
    if (!vendorRuleRows.length) return false;
    applySharedVendorRuleRows(vendorRuleRows);
    if (!silent) showToast(`Shared vendor rules loaded (${number.format(vendorRuleRows.length)} vendors)`, 2400, "success");
    return true;
  } catch (error) {
    if (!silent) showToast("Shared vendor rules refresh failed.", 2600, "warning");
    return false;
  }
}

async function restoreSharedImportLogsOnlyFromSupabase(options = {}) {
  if (!ENABLE_SHARED_SYNC || !sharedImportLogsAvailable) return false;
  const { silent = false } = options;
  try {
    const rows = await supabaseSelectRowsSafe("import_logs", { select: "*", order: "recorded_at.desc" });
    applySharedImportLogRows(rows);
    if (!silent) showToast(`Shared import logs loaded (${number.format(rows.length)} rows)`, 2400, "success");
    return true;
  } catch (error) {
    if (!silent) showToast("Shared import log refresh failed.", 2600, "warning");
    return false;
  }
}

async function restoreSharedCountSessionsOnlyFromSupabase(options = {}) {
  if (!ENABLE_SHARED_SYNC || !sharedCountSessionsAvailable) return false;
  // render option: false means applySharedCountSessionRows will not call renderCountsWorkspace.
  // Used by refreshLatestCountSessions which renders in its own finally block to avoid double render.
  const { silent = false, history = false, fromSync = false, render: renderOpt } = options;
  const shouldRender = renderOpt === false ? false : !fromSync;
  try {
    const rows = await supabaseSelectRows("count_sessions", {
      select: "*",
      order: "updated_at.desc",
      limit: history ? "60" : "25",
    });
    applySharedCountSessionRows(rows, { render: shouldRender });
    state._countSyncUnavailable = false;
    if (!silent) showToast("Shared count sessions loaded.", 2400, "success");
    return true;
  } catch (error) {
    state._countSyncUnavailable = true;
    if (!silent) showToast("Shared count session refresh failed.", 2600, "warning");
    return false;
  }
}

async function updateSharedSyncState(kind = "products") {
  if (!ENABLE_SHARED_SYNC) return false;
  try {
    const stamp = new Date().toISOString();
    const payload = [{
      id: SYNC_STATE_ROW_ID,
      last_sync_kind: kind,
      updated_at: stamp,
      latest_inventory_date: latestInventoryDate() || null,
      latest_sales_date: state.dates.at(-1) || null,
      product_count: Math.max(state.latestInventory.size, state.excelItems.size),
      sales_count: state.rawSales.length,
    }];
    await supabaseUpsertRows("sync_state", payload, "id");
    lastLocalSharedSyncAt = stamp;
    return true;
  } catch (_) {
    return false;
  }
}

async function syncSharedVendorRulesToSupabase(silent = false) {
  if (!ENABLE_SHARED_SYNC || !sharedVendorRulesAvailable) return false;
  try {
    const rows = [...new Map((state.vendorRules || [])
      .map(vendorRuleRowForSupabase)
      .filter((row) => row.id && row.vendor)
      .map((row) => [cleanCell(row.vendor).toUpperCase(), row])).values()];
    if (!rows.length) return true;
    await supabaseUpsertRows("vendor_rules", rows, "id");
    await updateSharedSyncState("vendor-rules");
    return true;
  } catch (error) {
    if (String(error?.message || "").includes("(401)") || String(error?.message || "").includes("(403)")) {
      sharedVendorRulesAvailable = false;
    }
    if (!silent) showToast("Shared vendor rules sync failed.", 3200, "warning");
    return false;
  }
}

async function syncSharedMetaSnapshotToSupabase(options = {}) {
  if (!ENABLE_SHARED_SYNC) return false;
  const { silent = false, includeVendorRules = true } = options;
  try {
    const productMetaRows = buildProductMetaRowsSnapshot();
    if (productMetaRows.length && sharedProductMetaAvailable) {
      await supabaseUpsertRows("product_meta", productMetaRows, "code");
    }
    if (includeVendorRules) {
      await syncSharedVendorRulesToSupabase(true);
    }
    await syncSharedImportLogsToSupabase(true);
    await syncSharedCountSessionsToSupabase(true);
    await updateSharedSyncState("product-meta");
    if (!silent) showToast("Shared product meta updated.", 2200, "success");
    return true;
  } catch (error) {
    if (String(error?.message || "").includes("product_meta")) {
      sharedProductMetaAvailable = false;
    }
    if (!silent) showToast("Shared product meta sync failed.", 2600, "warning");
    return false;
  }
}

async function syncSharedImportLogsToSupabase(silent = false) {
  if (!ENABLE_SHARED_SYNC || !sharedImportLogsAvailable) return false;
  try {
    const rows = (state.uploadLogs || []).map(importLogRowForSupabase).filter((row) => row.id);
    if (rows.length) {
      await supabaseUpsertRows("import_logs", rows, "id");
    }
    await updateSharedSyncState("imports");
    return true;
  } catch (error) {
    if (String(error?.message || "").includes("(401)") || String(error?.message || "").includes("(403)")) {
      sharedImportLogsAvailable = false;
    }
    if (!silent) showToast("Shared import log sync failed.", 2600, "warning");
    return false;
  }
}

async function syncSharedCountSessionsToSupabase(silent = false) {
  if (!ENABLE_SHARED_SYNC || !sharedCountSessionsAvailable) return false;
  if (state._countSyncInFlight) return state._countSyncInFlight;
  clearTimeout(sharedCountSessionsTimer);
  sharedCountSessionsTimer = 0;
  state._countSyncInFlight = (async () => {
    try {
      // FIX-A: Delete stale remote active:${id} rows for any locally-saved sessions BEFORE
      // pulling remote. If we pull first, the remote active row wins the merge and resurrects
      // the session as active — undoing the save on the saving device and all other devices.
      const savedSessionIds = (state.countSessions || [])
        .filter((session) => cleanCell(session?.id) && !countSessionIsDeleted(session.id) && session.savedAt)
        .map((session) => cleanCell(session.id))
        .filter(Boolean);
      if (savedSessionIds.length) {
        await supabaseDeleteRowsByValues("count_sessions", "id", savedSessionIds.map((id) => `active:${id}`));
      }
      const pulled = await restoreSharedCountSessionsOnlyFromSupabase({ silent: true, history: true, fromSync: true });
      if (!pulled) throw new Error("count sessions remote read failed");
    const deletedIds = [...(state._deletedCountSessionIds || [])].filter(Boolean).slice(-500);
    if (deletedIds.length) {
      await supabaseDeleteRowsByValues("count_sessions", "id", deletedIds.flatMap((id) => [id, `active:${id}`]));
    }
    const dirtySaved = (state.countSessions || [])
      .filter((session) =>
        cleanCell(session?.id)
        && !countSessionIsDeleted(session.id)
        && countSessionHasUnsyncedEdits(session)
        && !(session.id === state.activeCountSession?.id && session.isActiveLive === true)
      );
    const dirtyActive = state.activeCountSession?.id
      && !countSessionIsDeleted(state.activeCountSession.id)
      && countSessionHasUnsyncedEdits(state.activeCountSession)
      ? state.activeCountSession
      : null;
    if (dirtySaved.length) {
      await supabaseDeleteRowsByValues("count_sessions", "id", dirtySaved.map((session) => `active:${session.id}`));
    }
    if (dirtyActive) {
      await supabaseDeleteRowsByValues("count_sessions", "id", [dirtyActive.id]);
    }
    const rows = [
      ...dirtySaved
        .map((session) => countSessionRowForSupabase(session, false)),
      ...(dirtyActive
        ? [countSessionRowForSupabase(dirtyActive, true)]
        : []),
    ];
    if (rows.length) {
      // PATCH C: stamp lastLocalSharedSyncAt BEFORE the await so the 2s poll guard
      // fires immediately and the saving device does not re-pull its own just-pushed data.
      const syncedAt = new Date().toISOString();
      lastLocalSharedSyncAt = syncedAt;
      await supabaseUpsertRows("count_sessions", rows, "id");
      dirtySaved.forEach((session) => {
        session.entries = (session.entries || []).map((entry) => ({ ...entry, syncStatus: "synced" }));
        session.localSyncPending = false;
        session.lastSyncAt = syncedAt;
      });
      if (dirtyActive) {
        dirtyActive.entries = (dirtyActive.entries || []).map((entry) => ({ ...entry, syncStatus: "synced" }));
        dirtyActive.localSyncPending = false;
        dirtyActive.lastSyncAt = syncedAt;
      }
      state._countLastSyncAt = syncedAt;
      state._countSyncUnavailable = false;
      persistCountSessions({ scheduleSync: false });
      renderActiveCountSyncStatus();
    }
    if (rows.length || deletedIds.length) await updateSharedSyncState("counts");
    return true;
    } catch (error) {
      [state.activeCountSession, ...(state.countSessions || [])].filter(Boolean).forEach((session) => {
        if (!countSessionHasUnsyncedEdits(session)) return;
        session.entries = (session.entries || []).map((entry) => entry.syncStatus === "synced"
          ? entry
          : { ...entry, syncStatus: "failed" });
        session.syncError = "Shared sync failed";
      });
      state._countSyncUnavailable = true;
      persistCountSessions({ scheduleSync: false });
      renderActiveCountSyncStatus();
      if (String(error?.message || "").includes("(401)") || String(error?.message || "").includes("(403)")) {
        sharedCountSessionsAvailable = false;
      } else {
        scheduleCountSyncRetry();
      }
      if (!silent) showToast("Shared count session sync failed. Local entries are preserved and will retry.", 3600, "warning");
      return false;
    } finally {
      state._countSyncInFlight = null;
    }
  })();
  return state._countSyncInFlight;
}

function productRowByCodeForSupabase(code) {
  const normalized = codeKey(code || "");
  if (!normalized) return null;
  const inventory = state.latestInventory.get(normalized);
  if (inventory?.code || inventory?.product) return productRowForSupabase(inventory);
  const excel = findExcelFor({ code: normalized }) || state.excelItems.get(normalized);
  if (excel?.code || excel?.product || excel?.plu || excel?.itemNumber) return productRowFromExcelForSupabase(excel);
  return null;
}

function productMetaRowByCodeForSupabase(code) {
  const normalized = codeKey(code || "");
  if (!normalized) return null;
  const inventory = state.latestInventory.get(normalized);
  if (inventory?.code || inventory?.product) return productMetaRowForSupabase(inventory);
  const excel = findExcelFor({ code: normalized }) || state.excelItems.get(normalized);
  if (excel?.code || excel?.product || excel?.plu || excel?.itemNumber) return productMetaRowForSupabase(excel);
  return null;
}

async function syncSharedProductsByCodes(codes = [], options = {}) {
  if (!ENABLE_SHARED_SYNC) return false;
  const { silent = false, updateSyncState = true } = options;
  try {
    const uniqueCodes = [...new Set((codes || []).map((code) => codeKey(code)).filter(Boolean))];
    const productMetaRows = uniqueCodes
      .map(productMetaRowByCodeForSupabase)
      .filter((row) => cleanCell(row?.code));
    if (!productMetaRows.length) return true;
    await supabaseUpsertRows("product_meta", productMetaRows, "code");
    if (updateSyncState) await updateSharedSyncState("product-meta");
    if (!silent) showToast(`Shared ${productMetaRows.length === 1 ? "item" : "items"} synced.`, 2200, "success");
    return true;
  } catch (error) {
    if (String(error?.message || "").includes("product_meta")) {
      sharedProductMetaAvailable = false;
    }
    if (!silent) showToast("Shared product meta sync failed.", 2600, "warning");
    return false;
  }
}

async function syncSharedInventoryFactsByCodes(codes = [], options = {}) {
  if (!ENABLE_SHARED_SYNC) return false;
  const { silent = false, updateSyncState = true } = options;
  try {
    const uniqueCodes = [...new Set((codes || []).map((code) => codeKey(code)).filter(Boolean))];
    const productRows = uniqueCodes
      .map(productRowByCodeForSupabase)
      .filter((row) => cleanCell(row?.code));
    const productMetaRows = uniqueCodes
      .map(productMetaRowByCodeForSupabase)
      .filter((row) => cleanCell(row?.code));
    if (productMetaRows.length) await supabaseUpsertRows("product_meta", productMetaRows, "code");
    if (updateSyncState && productMetaRows.length) await updateSharedSyncState("product-meta");
    if (productRows.length) {
      await supabaseUpsertRows("products", productRows, "code");
      if (updateSyncState) await updateSharedSyncState("products");
    }
    if (!silent) showToast(`Shared ${uniqueCodes.length === 1 ? "item" : "items"} synced.`, 2200, "success");
    return true;
  } catch (error) {
    if (String(error?.message || "").includes("product_meta")) sharedProductMetaAvailable = false;
    if (!silent) showToast("Shared inventory sync failed.", 2600, "warning");
    return false;
  }
}

async function restoreSharedDataFromSupabase(options = {}) {
  if (!ENABLE_SHARED_SYNC) return false;
  const { silent = false, preferCurrentState = false } = options;
  try {
    const [productRows, productMetaRows, salesRows, vendorRuleRows, importLogRows, countSessionRows] = await Promise.all([
      supabaseSelectRows("products", { select: "*" }),
      supabaseSelectRowsSafe("product_meta", { select: "*" }),
      supabaseSelectRows("daily_sales", { select: "*", order: "sales_date.asc" }),
      supabaseSelectRowsSafe("vendor_rules", { select: "*" }),
      supabaseSelectRowsSafe("import_logs", { select: "*", order: "recorded_at.desc" }),
      supabaseSelectRowsSafe("count_sessions", { select: "*" }),
    ]);
    let mergedProductRows = mergeSharedProductRows(productRows, productMetaRows);
    if (!mergedProductRows.length && !salesRows.length && !vendorRuleRows.length && !importLogRows.length && !countSessionRows.length) return false;
    const localImportProtected = Date.now() < (state._localImportHoldUntil || 0);
    const shouldKeepLocalData = (preferCurrentState || localImportProtected) && sharedSnapshotIsOlderThanCurrent(mergedProductRows, salesRows);
    if (shouldKeepLocalData) {
      if (vendorRuleRows.length) applySharedVendorRuleRows(vendorRuleRows);
      if (importLogRows.length) applySharedImportLogRows(importLogRows);
      if (countSessionRows.length) applySharedCountSessionRows(countSessionRows);
      if (!silent) showToast("Kept newer local data instead of older shared snapshot.", 2800, "success");
      return "kept-local";
    }
    mergedProductRows = mergedProductRows.map((row) => {
      const key = codeKey(row.code);
      const localEditMs = state._localProductMetaEdits?.get?.(key) || 0;
      const remoteEditMs = Date.parse(row.updated_at || row.snapshot_date || "") || 0;
      if (!localEditMs || (remoteEditMs && remoteEditMs >= localEditMs - 1000)) return row;
      const localItem = state.latestInventory.get(key);
      if (!localItem) return row;
      return { ...row, stock: localItem.stock, price: localItem.price ?? row.price, unit_cost: localItem.cost ?? row.unit_cost };
    });

    state.rawSales = salesRows.map(hydrateSalesFromSupabase).filter((row) => row.code || row.product);
    state.dates = [...new Set(state.rawSales.map((row) => row.date))].sort();

    state.excelItems = new Map();
    state.excelByPlu = new Map();
    state.excelByItemNumber = new Map();
    const nextOverrides = {};
    const nextItemMeta = { ...(state.itemMeta || {}) };
    mergedProductRows.map(hydrateExcelFromSupabase).forEach((item) => addExcelIndex(item));
    mergedProductRows.forEach((row) => {
      const code = normalizeCode(row.code);
      if (!code) return;
      const existingMeta = nextItemMeta[code] || {};
      nextItemMeta[code] = {
        ...existingMeta,
        ...(row.state != null ? { state: cleanCell(row.state), stateManual: true } : {}),
        ...(row.case_size != null ? { caseSize: normalizedCaseSize(row.case_size), caseSizeManual: existingMeta.caseSizeManual === true } : {}),
        ...(row.add_date ? { addDate: normalizeItemDate(row.add_date) } : {}),
        orderingRuleMode: cleanCell(row.ordering_rule_mode) || existingMeta.orderingRuleMode || "vendor",
        safetyDaysOverride: row.safety_days_override != null ? Math.max(0, Math.round(toNumber(row.safety_days_override) || 0)) : null,
        daysOfInventoryOverride: row.days_of_inventory_override != null ? Math.max(0, Math.round(toNumber(row.days_of_inventory_override) || 0)) : null,
      };
      const overrideMin = row.reorder_min_override;
      const overrideMax = row.reorder_max_override;
      if (overrideMin != null || overrideMax != null) {
        nextOverrides[code] = {
          ...(overrideMin != null ? { min: toNumber(overrideMin) } : {}),
          ...(overrideMax != null ? { max: toNumber(overrideMax) } : {}),
        };
      }
    });
    state.itemMeta = nextItemMeta;
    saveItemMeta();
    state.reorderOverrides = nextOverrides;
    safeLocalStorageSet("posDashboardReorderOverrides:v1", JSON.stringify(state.reorderOverrides));
    rebuildExcelIndexes();

    const inventoryRows = mergedProductRows.map(hydrateInventoryFromSupabase).filter((row) => row.code || row.product);
    state.inventories = new Map();
    if (inventoryRows.length) {
      const inventoryDate = inventoryRows.map((row) => row.date).sort().at(-1) || new Date().toISOString().slice(0, 10);
      state.inventories.set(inventoryDate, inventoryRows);
    }
    buildLatestInventory();
      if (vendorRuleRows.length) {
        applySharedVendorRuleRows(vendorRuleRows);
      } else {
        const seededRules = ensureVendorRulesFromData();
        if (seededRules && ENABLE_SHARED_SYNC) {
          syncSharedVendorRulesToSupabase(true).catch(() => {});
        }
      }
      applySharedImportLogRows(importLogRows);
      applySharedCountSessionRows(countSessionRows);
    state._loadedFileSignatures = new Set(["supabase-shared"]);
    bumpDataStamp();
    if (scanModeIsActive()) {
      state._deferredDashboardOptions = true;
    } else {
      updateFilterOptions();
      updateInventoryStateFilter();
      setDefaultDates();
    }
    if (!silent) showToast(`Loaded shared data - ${number.format(mergedProductRows.length)} products, ${number.format(salesRows.length)} sales rows`, 3200, "success");
    return true;
  } catch (error) {
    if (!silent) showToast("Supabase shared data could not be loaded.", 3200, "warning");
    return false;
  }
}

async function syncSharedDataToSupabase(options = {}) {
  if (!ENABLE_SHARED_SYNC) return false;
  const { productsOnly = false, silent = false, salesDates = null } = options;
  try {
    const productRowMap = new Map();
    [...state.excelItems.values()]
      .filter((item) => item?.code || item?.product || item?.plu || item?.itemNumber)
      .forEach((item) => {
        const row = productRowFromExcelForSupabase(item);
        const key = codeKey(row.code);
        if (key) {
          productRowMap.set(key, row);
        }
      });
    [...state.latestInventory.values()]
      .filter((item) => item.code || item.product)
      .forEach((item) => {
        const row = productRowForSupabase(item);
        const key = codeKey(row.code);
        if (key) {
          productRowMap.set(key, row);
        }
      });
    const productRows = [...productRowMap.values()].filter((row) => cleanCell(row.code));
    const productMetaRows = buildProductMetaRowsSnapshot();
    if (!productsOnly) {
      // Full import = one authoritative master snapshot for every device.
      // Clear the shared snapshot tables first, then write the current laptop data,
      // and only then bump sync_state so tablets/personal laptops pull a complete set.
      await supabaseDeleteAllRows("products");
      await supabaseDeleteAllRows("product_meta");
      await supabaseDeleteAllRows("daily_sales");
      if (productRows.length) await supabaseInsertRows("products", productRows);
      if (productMetaRows.length) await supabaseInsertRows("product_meta", productMetaRows);
    } else {
      if (productRows.length) await supabaseUpsertRows("products", productRows, "code");
      if (productMetaRows.length) await supabaseUpsertRows("product_meta", productMetaRows, "code");
    }
    await syncSharedVendorRulesToSupabase(true);

    if (!productsOnly) {
      const salesRows = state.rawSales
        .filter((row) => row.code || row.product)
        .map(salesRowForSupabase);
      if (salesRows.length) await supabaseInsertRows("daily_sales", salesRows);
    }
    await syncSharedImportLogsToSupabase(true);
    await syncSharedCountSessionsToSupabase(true);
    await updateSharedSyncState(productsOnly ? "product-meta" : "full-snapshot");
    if (!productsOnly) await savePersistedState();
    if (!silent) showToast("Shared Supabase data updated.", 2800, "success");
    return true;
  } catch (error) {
    console.warn("Supabase shared data sync failed", error);
    if (!silent) showToast("Supabase shared data sync failed. Check table columns/policies.", 4200, "warning");
    return false;
  }
}

async function initApp() {
  const bootStartedAt = performanceNow();
  console.info("[AppPerf] app boot start");
  try {
    state._countConfirmAudio = new Audio("add-qty-kaching-376867.mp3?v=376867");
    state._countConfirmAudio.preload = "auto";
    state._countConfirmAudio.volume = 1;
    state._countConfirmAudio.load?.();
  } catch (_) {}
  mountInventoryQuickTools();
  if (ENABLE_SHARED_SYNC) {
    refreshUsersFromSupabase({ silent: true }).catch(() => {});
  }
  resetUiCriteriaOnStartup();
  switchTab(state.currentUser && isUserRole() ? "scanmode" : "dashboard");
  updateMetricsSummaryMode();
  if (!scanModeIsActive()) {
    render();
    renderUploadLogs();
  }
  document.body.dataset.activeTab = activeTabName();
  renderPriceCheckResult(null);
  appInitDone = true;
  try {
    await restorePersistedState();
    state._initialRestoreComplete = true;
    queueScanExactIndexBuild({ prioritize: true, immediate: true });
    pruneDeletedCountSessions({ persist: true });
    ensureVendorRulesFromData();
    if (ENABLE_SHARED_SYNC && state.vendorRules.length) {
      syncSharedVendorRulesToSupabase(true).catch(() => {});
    }
    if (state.dates.length && !scanModeIsActive()) {
      state.activePresetDays = 90;
      applyDatePreset(90);
    } else if (!scanModeIsActive()) {
      render();
    } else {
      queueScanExactIndexBuild();
    }
    if (!scanModeIsActive()) renderUploadLogs();
  } catch (error) {
    state._initialRestoreComplete = true;
    console.warn("Background local restore failed", error);
  }
  if (ENABLE_SHARED_SYNC) {
    setTimeout(async () => {
      try {
        const hasLocalDataset = !!(state.latestInventory.size || state.excelItems.size || state.rawSales.length);
        if (hasLocalDataset && scanModeIsActive() && scanPerformanceProfile.deferRestoresWhileScanning) {
          state._deferredSharedSync = true;
          return;
        }
        // Shared master snapshot wins on normal devices. Previously, any device that
        // already had local data only pulled product deltas, so daily_sales/data-file
        // fields like case size, disabled/discontinued, etc. could stay stale.
        const restoredShared = await restoreSharedDataFromSupabase({
          silent: true,
          preferCurrentState: window.location.protocol === "file:",
        });
        if (restoredShared === "kept-local") {
          await syncSharedMetaSnapshotToSupabase({ silent: true, includeVendorRules: true });
        } else if (restoredShared) {
          await savePersistedState();
          queueScanExactIndexBuild({ prioritize: true, immediate: true });
          if (scanModeIsActive()) {
            queueScanExactIndexBuild();
          } else {
            queueActiveTabRender();
            renderUploadLogs();
          }
        }
      } catch (error) {
        console.warn("Background shared restore failed", error);
      }
    }, 150);
  }
  logPerformanceEnd("app boot", bootStartedAt, { tab: activeTabName() });
}

function bootAppIfNeeded() {
  if (appInitPromise) return appInitPromise;
  appInitPromise = initApp().catch((error) => {
    appInitPromise = null;
    throw error;
  });
  return appInitPromise;
}

// -- Confirm delete saved session ------------------------------------------
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
  const session = state.countSessions.find((s) => s.id === id)
    || (state.activeCountSession?.id === id ? state.activeCountSession : null);
  if (!countSessionCanDelete(session)) {
    showToast("Only the count owner or an admin can delete this session.", 3000, "warning");
    return;
  }
  archiveCountSessionEntriesToAdjustmentLog(session);
  markCountSessionDeleted(id);
  state.countSessions = state.countSessions.filter((s) => s.id !== id);
  if (state.activeCountSession?.id === id) {
    state.activeCountSession = null;
    state._countSessionOpen = false;
  }
  persistCountSessions();
  renderCountsWorkspace();
  void (async () => {
    if (ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
      await supabaseDeleteRowsByValues("count_sessions", "id", [id, `active:${id}`]);
      await updateSharedSyncState("counts");
    }
    await syncSharedCountSessionsToSupabase(true);
  })().catch(() => {});
  showToast("Physical count deleted locally and queued for shared removal.", 2800, "warning");
}

// -- Continue count from report (re-open the session as active) -------------
// ===========================================================================
// V2.0 SESSION MANAGER — single canonical open path.
// Every resume/continue path routes through openCount(sessionId).
// showCountWorkspace() is the one place that forces the workspace visible and
// tears down any container overlay (report / history / setup) that could sit on
// top and intercept clicks. No path should hand-roll its own open logic.
// ===========================================================================

// Force the active count workspace visible and remove every container overlay
// that is NOT the workspace itself. Safe to call repeatedly.
function showCountWorkspace() {
  if (els.countReportModal) els.countReportModal.hidden = true;
  const reportCountModal = document.querySelector("#reportCountModal");
  if (reportCountModal) reportCountModal.hidden = true;
  const sessionHistoryModal = document.querySelector("#sessionHistoryModal");
  if (sessionHistoryModal) sessionHistoryModal.hidden = true;
  if (els.countSetupModal) els.countSetupModal.hidden = true;
  renderCountsWorkspace();
  if (els.countSessionModal) {
    // Force it fully visible and in front. removeAttribute is belt-and-suspenders in case
    // a stale [hidden] attribute lingers and defeats the :not([hidden]) z-index rule.
    els.countSessionModal.hidden = false;
    els.countSessionModal.removeAttribute("hidden");
    els.countSessionModal.classList.add("count-session-forced-open");
    // Land at the TOP of the count screen (search box), not scrolled down to the keypad.
    els.countSessionModal.scrollTop = 0;
  }
}

// THE single function for opening/resuming an existing saved count by id.
// Used by Continue (from report), Resume (active card), and any future resume path.
// Returns true if the workspace was opened.
function openCount(sessionId) {
  const id = cleanCell(sessionId);
  const session = findCountSessionById(id)
    || (cleanCell(state.activeCountSession?.id) === id ? state.activeCountSession : null);
  if (!session) { showToast("Session not found.", 3000, "warning"); return false; }
  // FIX-B: keep it in countSessions[] with isActiveLive:true so history still shows it.
  const liveSession = markCountSessionDirty({ ...session, savedAt: "", submittedAt: "", isActiveLive: true });
  state.countSessions = [liveSession, ...state.countSessions.filter((s) => s.id !== id)];
  state.activeCountSession = liveSession;
  state._countSessionOpen = true;
  state.countQtyBuffer = "0";
  state.selectedCountItemCode = "";
  state.countStage = "search";
  state.pendingDuplicateCount = null;
  state.pendingDuplicateMode = null;
  state._continuingCountId = id;
  state._ignoreNextCountBackdropCloseUntil = Date.now() + 4000;
  persistActiveCountSession();
  persistCountSessions({ scheduleSync: false });
  startContinueCountKeepOpenGuard(id);
  // Open immediately, then re-assert twice to win against any late render/sync tick
  // that would otherwise close the workspace right after it opened.
  showCountWorkspace();
  setTimeout(showCountWorkspace, 120);
  setTimeout(() => { showCountWorkspace(); focusCountSearch(); }, 400);
  return true;
}

async function continueCountFromReport(event = null) {
  event?.preventDefault?.();
  event?.stopPropagation?.();
  // PERF-F: find locally first; only fall back to a timed remote pull if missing.
  const sessionId = state.countReportOpenId;
  let session = findCountSessionById(sessionId);
  if (!session && ENABLE_SHARED_SYNC && sharedCountSessionsAvailable) {
    const timeout = scanPerformanceProfile.tier === "low" ? 4000 : 2000;
    await Promise.race([
      refreshLatestCountSessions({ history: true }),
      new Promise((resolve) => setTimeout(resolve, timeout)),
    ]);
    session = findCountSessionById(sessionId);
  }
  if (!session) { showToast("Session not found.", 3000, "warning"); return; }
  if (openCount(sessionId)) {
    // Do not block the reopened workspace on sync; queue it after the UI is stable.
    setTimeout(() => syncSharedCountSessionsToSupabase(true).catch(() => scheduleSharedCountSessionsSync()), 800);
    showToast(`Continuing count: ${countSessionLabel(session)}`, 2800, "success");
  }
}

// V2.0: Start New always prompts the scope filters and begins a FRESH session.
// Guard against silently discarding an in-progress count on THIS device.
function startNewCount() {
  const active = state.activeCountSession;
  const hasUnsavedWork = active && (active.entries || []).length && !active.savedAt && !active.submittedAt;
  if (hasUnsavedWork) {
    const ok = confirm(
      `A count is already in progress on this device:\n${countSessionLabel(active)}\n` +
      `(${(active.entries || []).length} item(s) scanned)\n\n` +
      `Starting a NEW count sets that one aside — it stays in Report History and you can ` +
      `Continue it later. Start a new count now?`,
    );
    if (!ok) return;
  }
  // Close any open workspace FIRST so renderCountsWorkspace cannot re-show the old count
  // on top of the New-count filter modal.
  state._countSessionOpen = false;
  if (els.countSessionModal) els.countSessionModal.hidden = true;
  openCountSetupModal();
}

// V2.0: Resume the active count on this device. If there is none, send the user
// to Report History (where Continue lives) rather than silently doing nothing.
function resumeActiveCount() {
  if (state.activeCountSession) {
    openCount(state.activeCountSession.id);
    return;
  }
  showToast("No active count on this device. Open Report History to continue a saved count, or Start New Count.", 4200, "warning");
  if (typeof openSessionHistoryModal === "function") openSessionHistoryModal();
}

async function ensureCountSessionReadyToApply(session, action = "submitting") {
  if (!session) return null;
  let current = findCountSessionById(session.id) || session;
  if (current.localSyncPending || countEntriesAwaitingSync(current).length) {
    await syncSharedCountSessionsToSupabase(true);
    current = findCountSessionById(session.id) || current;
  }
  const pending = countEntriesAwaitingSync(current);
  if (!current.localSyncPending && !pending.length) return current;
  const failedCount = pending.filter((entry) => entry.syncStatus === "failed").length;
  const message = `Cannot finish ${action}: ${pending.length} count entr${pending.length === 1 ? "y is" : "ies are"} not confirmed in shared sync (${failedCount} failed). Local entries are preserved.`;
  if (isAdmin() && confirm(`${message}\n\nManager override: continue anyway and record an unsynced warning?`)) {
    current.managerSyncOverrideAt = new Date().toISOString();
    current.managerSyncOverrideBy = currentAuditUser();
    current.finalizedWithUnsyncedEntries = true;
    persistCountSessions({ scheduleSync: false });
    return current;
  }
  showToast(message, 5200, "warning");
  renderActiveCountSyncStatus(current);
  return null;
}

// -- Submit & apply count (from report modal) -------------------------------
async function openConfirmSubmitCount() {
  const sessionId = state.countReportOpenId;
  state.pendingSubmitSessionId = sessionId;
  const session = await ensureCountSessionReadyToApply(findCountSessionById(sessionId), "submitting");
  if (!session) {
    state.pendingSubmitSessionId = null;
    return;
  }
  const entryCount = (session?.entries || []).length;
  const candidates = currentCountSessionCandidates(session || {});
  const nullCount = candidates.length - entryCount;
  if (els.confirmSubmitCountMessage) {
    els.confirmSubmitCountMessage.innerHTML =
      `<b>${entryCount}</b> scanned item${entryCount !== 1 ? "s" : ""} will be updated to physical counts.<br>` +
      (nullCount > 0 ? `<b>${nullCount}</b> unscanned item${nullCount !== 1 ? "s" : ""} in scope will be set to <b>0</b>.<br>` : "") +
      `<br>A pre-count snapshot will be saved so you can restore if needed.<br><br>WARNING: Are you sure you want to apply?`;
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
      user: currentAuditUser(),
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
  markCountSessionDirty(session);
  Object.keys(session.preCountSnapshot || {}).forEach((code) => state._localProductMetaEdits.set(codeKey(code), Date.now()));
  persistCountSessions();
  safeLocalStorageSet("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  bumpDataStamp();
  renderCountSessionRows();
  renderCountsWorkspace();
  render();
  renderAdjustLog();
  void syncSharedCountSessionsToSupabase(true);
  void syncSharedProductsByCodes(Object.keys(session.preCountSnapshot || {}), { silent: true });
  void syncSharedInventoryFactsByCodes(Object.keys(session.preCountSnapshot || {}), { silent: true, updateSyncState: true });
  showToast(`Restored - ${restored} stock values reverted`, 3200, "success");
}

async function submitAndApplyCount() {
  clearContinuedCountLock();
  const sessionId = state.pendingSubmitSessionId;
  state.pendingSubmitSessionId = null;
  document.querySelector("#confirmSubmitCountModal").hidden = true;
  const candidateSession = findCountSessionById(sessionId);
  if (!candidateSession) { showToast("Session not found.", 3000, "warning"); return; }
  const session = await ensureCountSessionReadyToApply(candidateSession, "submitting");
  if (!session) return;
  const entries = session.entries || [];
  const latestByCode = new Map();
  entries.forEach((entry) => latestByCode.set(codeKey(entry.code), entry));

  // Save pre-count snapshot for restore
  const snapshot = {};
  state.latestInventory.forEach((item, key) => { snapshot[key] = item.stock; });
  const savedSession = markCountSessionDirty({ ...session, preCountSnapshot: snapshot, submittedAt: new Date().toISOString() });
  state.countSessions = state.countSessions.map((s) => s.id === sessionId ? savedSession : s);

  // All items in scope - scanned items get their count, null/unscanned items get 0
  const candidates = currentCountSessionCandidates(session);
  const scopeCodes = new Set(candidates.map((item) => codeKey(item.code)));
  let updated = 0;
  state.latestInventory.forEach((item, key) => {
    if (!scopeCodes.has(key) && !scopeCodes.has(codeKey(item.code))) return;
    const entry = latestByCode.get(key) || latestByCode.get(codeKey(item.code));
    const before = item.stock;
    const after = entry ? Number(entry.countedQty || 0) : 0; // NULL -> 0
    if (before === after) return;
    item.stock = after;
    state.latestInventory.set(key, item);
    state._localProductMetaEdits.set(codeKey(item.code), Date.now());
    updated++;
    state.adjustmentLog.unshift({
      recordedAt: new Date().toISOString(),
      user: currentAuditUser(),
      code: item.code,
      product: item.product,
      vendor: item.vendor || "",
      category: item.category || "",
      action: "COUNT SUBMIT",
      qtyChange: after - before,
      qtyBefore: before,
      qtyAfter: after,
      reason: entry ? `Physical count: ${countSessionLabel(session)}` : `NULL -> 0: ${countSessionLabel(session)}`,
    });
  });
  safeLocalStorageSet("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  persistCountSessions();
  bumpDataStamp();
  closeCountReport();
  renderCountsWorkspace();
  render();
  renderAdjustLog();
  generateFinalCountExport(session, candidates, latestByCode, snapshot);
  void syncSharedCountSessionsToSupabase(true);
  void syncSharedProductsByCodes([...scopeCodes], { silent: true });
  void syncSharedInventoryFactsByCodes([...scopeCodes], { silent: true, updateSyncState: true });
  showToast(`Submitted - ${updated} stock values updated`, 3200, "success");
}

// -- Zero negatives ---------------------------------------------------------
function openConfirmZeroNeg() {
  const negItems = [...state.latestInventory.values()].filter((item) => (item.stock || 0) < 0);
  if (els.confirmZeroNegMessage) {
    els.confirmZeroNegMessage.textContent = negItems.length
      ? `${negItems.length} item${negItems.length !== 1 ? "s" : ""} have negative stock and will be set to 0. A report entry will be saved for each.`
      : "No items currently have negative stock.";
  }
  document.querySelector("#confirmZeroNegModal").hidden = false;
}

async function downloadZeroedNegativesReport(items) {
  if (!items?.length) return;
  const rows = items.map((item) => ({
    code: item.code,
    item: item.product,
    vendor: item.vendor || "",
    category: item.category || "",
    qty_before: item.before,
    qty_after: 0,
    adjusted_at: new Date().toISOString(),
    reason: "ZERO NEGATIVES",
  }));
  const xlsx = await ensureXlsxReader();
  if (xlsx) {
    const wb = xlsx.utils.book_new();
    const ws = xlsx.utils.json_to_sheet(rows);
    xlsx.utils.book_append_sheet(wb, ws, "Zeroed Negatives");
    xlsx.writeFile(wb, `Zeroed_Negatives_${new Date().toISOString().slice(0, 10)}.xlsx`);
    return;
  }
  downloadCsv(`Zeroed_Negatives_${new Date().toISOString().slice(0, 10)}.csv`, rows);
}

async function applyZeroNegatives() {
  document.querySelector("#confirmZeroNegModal").hidden = true;
  const negItems = [...state.latestInventory.values()].filter((item) => (item.stock || 0) < 0);
  if (!negItems.length) { showToast("No negative stock values found.", 2400); return; }
  const zeroedRows = [];
  negItems.forEach((item) => {
    const before = item.stock;
    item.stock = 0;
    state.latestInventory.set(codeKey(item.code), item);
    zeroedRows.push({ ...item, before });
    state.adjustmentLog.unshift({
      recordedAt: new Date().toISOString(),
      user: currentAuditUser(),
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
  safeLocalStorageSet("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  bumpDataStamp();
  render();
  renderAdjustLog();
  void syncSharedProductsByCodes(negItems.map((item) => item.code), { silent: true });
  await downloadZeroedNegativesReport(zeroedRows);
  showToast(`${negItems.length} negative stock values set to 0`, 2800, "success");
}

// -- Stock Adjust Modal (Products tab) --------------------------------------
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
  const action = state.stockAdjustAction;
  const inventoryItem = state.latestInventory.get(codeKey(item.code));
  if (!inventoryItem) { showToast("Item not found in inventory.", 3000, "warning"); closeStockAdjustModal(); return; }
  const before = inventoryItem.stock;
  let after;
  if (action === "add") after = before + qty;
  else if (action === "remove") after = before - qty;
  else after = qty; // set
  inventoryItem.stock = after;
  state.latestInventory.set(codeKey(item.code), inventoryItem);
  // Also update any duplicate keys
  state.latestInventory.forEach((inv, k) => {
    if (codeKey(inv.code) === codeKey(item.code)) inv.stock = after;
  });
  const key = codeKey(item.code);
  state._localProductMetaEdits.set(key, Date.now());
  patchInventoryModelStock(item.code, after);
  patchVisibleStockCell(item.code, after);
  // Record in log
  state.adjustmentLog.unshift({
    recordedAt: new Date().toISOString(),
    user: currentAuditUser(),
    code: item.code,
    product: item.product,
    vendor: item.vendor || "",
    category: item.category || "",
    action: String(action || "set").toUpperCase(),
    qtyChange: after - before,
    qtyBefore: before,
    qtyAfter: after,
    reason,
  });
  safeLocalStorageSet("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  state._pinnedAdjustCode = item.code;
  closeStockAdjustModal();
  patchVisibleStockCell(item.code, after);
  if (state.inventoryRows?.length) renderInventorySummary(state.inventoryRows);
  if (activeTabName() === "ordering") setTimeout(() => renderOrders(), 0);
  else if (activeTabName() !== "inventory") setTimeout(() => queueActiveTabRender(), 0);
  renderAdjustLog();
  void savePersistedState();
  void syncSharedProductsByCodes([item.code], { silent: true });
  void syncSharedInventoryFactsByCodes([item.code], { silent: true, updateSyncState: false });
  showToast(`${action === "add" ? "Added" : action === "remove" ? "Removed" : "Set"} ${number.format(qty)} - ${item.code} now ${number.format(after)}`, 3000, "success");
}

// -- Adjust log rendering ---------------------------------------------------
function monthKeyForDate(value) {
  const date = new Date(value);
  if (Number.isNaN(date.getTime())) return "";
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
}

function renderAdjustMonthFilter() {
  const select = document.querySelector("#adjustLogMonthFilter");
  if (!select) return;
  const current = select.value;
  const months = [...new Set((state.adjustmentLog || []).map((entry) => monthKeyForDate(entry.recordedAt)).filter(Boolean))].sort().reverse();
  select.innerHTML = `<option value="">All months</option>${months.map((month) => {
    const [year, mm] = month.split("-");
    const start = new Date(Number(year), Number(mm) - 1, 1);
    const end = new Date(Number(year), Number(mm), 0);
    return `<option value="${month}">${start.toLocaleDateString("en-US")} - ${end.toLocaleDateString("en-US")}</option>`;
  }).join("")}`;
  select.value = months.includes(current) ? current : "";
}

function filteredAdjustLogEntries() {
  const month = document.querySelector("#adjustLogMonthFilter")?.value || "";
  return (state.adjustmentLog || []).filter((entry) => !month || monthKeyForDate(entry.recordedAt) === month);
}

function renderAdjustLog() {
  if (!els.adjustLogBody) return;
  renderAdjustMonthFilter();
  const entries = filteredAdjustLogEntries();
  if (!entries.length) {
    els.adjustLogBody.innerHTML = `<tr><td colspan="10" class="empty-cell">No stock adjustments recorded yet.</td></tr>`;
    return;
  }

  // Separate normal vs. auto-generated (NULL->0 and RESTORE) entries
  const normalEntries = entries.filter((e) => !["NULL -> 0", "RESTORE", "ZERO NEGATIVES"].some((t) => (e.reason || "").toUpperCase().includes(t) || (e.action || "").toUpperCase().includes(t)));
  const autoEntries   = entries.filter((e) =>  ["NULL -> 0", "RESTORE", "ZERO NEGATIVES"].some((t) => (e.reason || "").toUpperCase().includes(t) || (e.action || "").toUpperCase().includes(t)));

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
          <summary>${autoEntries.length} auto-generated entr${autoEntries.length === 1 ? "y" : "ies"} (NULL->0 / RESTORE / ZERO NEGATIVES)</summary>
          <table class="count-report-table inner-auto-table">
            <thead><tr><th>Date/Time</th><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th>Action</th><th>Change</th><th>Before</th><th>After</th><th>Reason</th></tr></thead>
            <tbody>${autoEntries.map(rowHtml).join("")}</tbody>
          </table>
        </details>
      </td></tr>`
    : "";

  els.adjustLogBody.innerHTML = normalHtml + autoHtml;
}

// Reason -> color map for reports
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
  const exportRows = filteredAdjustLogEntries();
  if (!exportRows.length) { showToast("No adjustment records to export.", 3000, "warning"); return; }
  const dateStr = new Date().toLocaleDateString("en-US", { year: "numeric", month: "long", day: "numeric" });
  const rows = exportRows.map((entry) => {
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
  <div class="meta">Generated ${dateStr} &nbsp; - &nbsp; ${exportRows.length} records</div>
  <table><thead><tr><th>Date/Time</th><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th>Action</th><th>Change</th><th>Qty Before</th><th>Qty After</th><th>Reason</th></tr></thead>
  <tbody>${rows}</tbody></table></body></html>`;
  const win = window.open("", "_blank", "width=1100,height=750");
  if (!win) { showToast("Pop-up blocked - please allow pop-ups.", 3500, "warning"); return; }
  win.document.write(html);
  win.document.close();
  setTimeout(() => win.print(), 500);
}

async function exportAdjustLogExcel() {
  const exportRows = filteredAdjustLogEntries();
  if (!exportRows.length) { showToast("No adjustment records to export.", 3000, "warning"); return; }
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available.", 3000, "warning"); return; }
  const wb = xlsx.utils.book_new();
  const data = [
    ["Stock Adjustment Log", "", "", "", "", "", "", "", "", `Generated: ${new Date().toLocaleDateString()}`],
    [],
    ["Date/Time", "Code", "Item", "Vendor", "Category", "Action", "Qty Change", "Qty Before", "Qty After", "Reason"],
    ...exportRows.map((entry) => [
      new Date(entry.recordedAt).toLocaleString(),
      entry.code || "", entry.product || "", entry.vendor || "", entry.category || "",
      entry.action || "", entry.qtyChange || 0, entry.qtyBefore ?? 0, entry.qtyAfter ?? 0, (entry.reason || "").toUpperCase(),
    ]),
  ];
  const ws = xlsx.utils.aoa_to_sheet(data);
  ws["!cols"] = [20, 12, 32, 14, 14, 12, 10, 10, 10, 26].map((w) => ({ wch: w }));
  xlsx.utils.book_append_sheet(wb, ws, "Adjustment Log");
  xlsx.writeFile(wb, `StockAdjustments_${new Date().toISOString().slice(0, 10)}.xlsx`);
  showToast(`Exported ${exportRows.length} adjustment records`, 2800, "success");
}

// -- Final Count Report (submitted sessions only) ----------------------------
function openFinalCountReport(sessionId) {
  const session = state.countSessions.find((s) => s.id === sessionId);
  if (!session?.preCountSnapshot) {
    showToast("No submitted snapshot found for this session.", 3000, "warning");
    return;
  }
  state.finalReportSessionId = sessionId;
  if (els.finalReportTitle) els.finalReportTitle.textContent = `Final Count - ${countSessionLabel(session)}`;
  if (els.finalReportMeta) {
    const syncStats = countSessionSyncStats(session);
    els.finalReportMeta.innerHTML = `
      <span><b>Date</b> ${escapeHtml(session.date || "-")}</span>
      <span><b>Vendor</b> ${escapeHtml(session.vendor || "All")}</span>
      <span><b>Category</b> ${escapeHtml(session.category || "All")}</span>
      <span><b>Status filter</b> ${escapeHtml(session.status || "All")}</span>
      <span><b>Submitted</b> ${escapeHtml(new Date(session.submittedAt).toLocaleString())}</span>
      <span><b>Entries</b> ${number.format((session.entries || []).length)}</span>
      <span><b>Sync</b> ${syncStats.pending || syncStats.failed || session.finalizedWithUnsyncedEntries ? "WARNING - unconfirmed entries" : "All entries synced"}</span>`;
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
    const status = entry ? "Scanned" : "Not scanned (-> 0)";
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
  const syncStats = countSessionSyncStats(session);
  const syncWarning = syncStats.pending || syncStats.failed || session.finalizedWithUnsyncedEntries
    ? "WARNING - export includes entries not confirmed by shared sync."
    : "All entries synced.";
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
      <td>${entry ? "" : "NULL"}</td>
      <td class="num">${variance > 0 ? "+" : ""}${number.format(variance)}</td>
      <td class="num">${currency.format(costVar)}</td>
      <td>${entry ? "Scanned" : "Not scanned (-> 0)"}</td>
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
    <span><b>Sync</b> ${escapeHtml(syncWarning)}</span>
    <span><b>Generated</b> ${dateStr}</span>
  </div>
  <table><thead><tr><th>Code</th><th>Item</th><th>Vendor</th><th>Category</th><th>Qty Start</th><th>Qty End</th><th>NULL</th><th>Variance</th><th>Cost Var</th><th>Status</th></tr></thead>
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
  const syncStats = countSessionSyncStats(session);
  const syncNote = syncStats.pending || syncStats.failed || session.finalizedWithUnsyncedEntries
    ? "WARNING: export includes entries not confirmed by shared sync."
    : "Sync: all entries confirmed.";
  const wb = xlsx.utils.book_new();
  const data = [
    ["Final Physical Count Report", "", `Date: ${session.date || "-"}`, `Vendor: ${session.vendor || "All"}`, `Submitted: ${new Date(session.submittedAt).toLocaleString()}`, syncNote],
    [],
    ["Code", "Item", "Vendor", "Category", "Qty Start", "Qty End", "NULL", "Variance", "Cost Variance", "Status"],
    ...candidates.map((item) => {
      const key = codeKey(item.code);
      const qtyStart = snapshot[key] ?? Number(item.stock ?? 0);
      const entry = latestByCode.get(key);
      const qtyEnd = entry ? Number(entry.countedQty || 0) : 0;
      const variance = qtyEnd - qtyStart;
      const costVar = variance * Number(item.unitCost || 0);
      return [item.code, item.product, item.vendor || "", item.category || "", qtyStart, qtyEnd, entry ? "" : "NULL", variance, costVar, entry ? "Scanned" : "Not scanned (-> 0)"];
    }),
  ];
  const ws = xlsx.utils.aoa_to_sheet(data);
  ws["!cols"] = [14, 34, 14, 14, 10, 10, 8, 10, 12, 18].map((w) => ({ wch: w }));
  applyXlsxTextToCodeColumns(ws, data, [0]); // code column = index 0
  xlsx.utils.book_append_sheet(wb, ws, "Final Count");
  xlsx.writeFile(wb, `FinalCount_${session.date || "report"}.xlsx`);
  showToast(`Final count exported - ${candidates.length} items`, 2800, "success");
}

// -- Vendor Rules (Vendors tab) ----------------------------------------------
function persistVendorRules() {
  safeLocalStorageSet("posDashboardVendorRules:v1", JSON.stringify(state.vendorRules));
  bumpDataStamp();
  queueActiveTabRender();
  scheduleSharedVendorRulesSync();
}

const DAYS_OF_WEEK = ["monday","tuesday","wednesday","thursday","friday","saturday","sunday"];
const DAY_SHORT = { monday:"Mon",tuesday:"Tue",wednesday:"Wed",thursday:"Thu",friday:"Fri",saturday:"Sat",sunday:"Sun" };
const DEFAULT_VENDOR_RULE = Object.freeze({
  status: "Active",
  safetyDays: DEFAULT_ORDER_SAFETY_DAYS,
  daysOfInventory: DEFAULT_ORDER_DAYS_OF_INVENTORY,
  minOrder: 0,
  email: "",
  notes: "",
  orderDays: [],
});

function createVendorRule(vendorName, overrides = {}) {
  return {
    id: overrides.id || `vr-${Date.now()}-${Math.random().toString(36).slice(2,6)}`,
    vendor: String(vendorName || "").trim().toUpperCase(),
    status: overrides.status ?? DEFAULT_VENDOR_RULE.status,
    safetyDays: overrides.safetyDays ?? DEFAULT_VENDOR_RULE.safetyDays,
    daysOfInventory: overrides.daysOfInventory ?? DEFAULT_VENDOR_RULE.daysOfInventory,
    minOrder: overrides.minOrder ?? DEFAULT_VENDOR_RULE.minOrder,
    email: overrides.email ?? DEFAULT_VENDOR_RULE.email,
    notes: overrides.notes ?? DEFAULT_VENDOR_RULE.notes,
    orderDays: [...(overrides.orderDays || DEFAULT_VENDOR_RULE.orderDays)],
    updatedAt: overrides.updatedAt || new Date().toISOString(),
  };
}

function vendorNamesFromLoadedData() {
  const inventoryRows = [...state.latestInventory.values()];
  const excelRows = [...state.excelItems.values()];
  const salesVendors = state.rawSales.map((r) => r.vendor).filter(Boolean);
  return unique([
    ...inventoryRows.map((r) => r.vendor),
    ...excelRows.map((r) => r.vendor),
    ...salesVendors,
  ].filter((v) => v && v !== "Unassigned" && v !== "-"));
}

function ensureVendorRulesFromData() {
  const allVendors = vendorNamesFromLoadedData();
  if (!allVendors.length) return 0;
  const existingNames = new Set(state.vendorRules.map((r) => r.vendor?.toUpperCase()));
  const newVendors = allVendors.filter((vendor) => !existingNames.has(vendor.toUpperCase()));
  newVendors.forEach((vendor) => {
    state.vendorRules.push(createVendorRule(vendor));
  });
  if (newVendors.length) persistVendorRules();
  return newVendors.length;
}

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
      ${todayVendors.length ? `<span class="vendor-today-alert">WARNING: <b>${todayVendors.length}</b> vendor${todayVendors.length>1?"s":""} to order TODAY: ${todayVendors.map(v=>v.vendor).join(", ")}</span>` : ""}`;
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

  // Inline number/text/email inputs - save on blur or Enter
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

  const rule = createVendorRule(vendor, {
    id: state.vendorRuleEditId || `vr-${Date.now()}`,
    status: document.querySelector("#vrStatus").value,
    safetyDays: toNumber(document.querySelector("#vrSafetyDays").value) || 21,
    daysOfInventory: toNumber(document.querySelector("#vrDaysOfInventory").value) || 7,
    minOrder: toNumber(document.querySelector("#vrMinOrder").value) || 0,
    email: document.querySelector("#vrEmail").value.trim(),
    notes: document.querySelector("#vrNotes").value.trim(),
    orderDays: [...state.vendorRuleSelectedDays],
    updatedAt: new Date().toISOString(),
  });

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

async function importVendorRulesFile(file) {
  if (!file) return;
  const xlsx = await ensureXlsxReader();
  if (!xlsx) { showToast("Excel library not available.", 3000, "warning"); return; }
  try {
    const workbook = xlsx.read(await file.arrayBuffer(), { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = xlsx.utils.sheet_to_json(sheet, { header: 1, defval: "" });
    const headerIndex = rows.findIndex((row) => {
      const cells = Array.isArray(row) ? row.map((cell) => cleanCell(cell).toLowerCase()) : [];
      return cells.includes("vendor") && (
        cells.includes("status")
        || cells.includes("safety days")
        || cells.includes("safety stock days")
        || cells.includes("days of inventory")
        || cells.includes("min order $")
        || cells.includes("min order amount ($)")
        || cells.includes("order days")
      );
    });
    if (headerIndex < 0) {
      showToast("Could not find vendor-rule columns in that file.", 3200, "warning");
      return;
    }
    const headers = rows[headerIndex].map((cell) => cleanCell(cell).toLowerCase());
    const col = (...names) => headers.findIndex((header) => names.includes(header));
    const vendorCol = col("vendor");
    const statusCol = col("status");
    const safetyCol = col("safety days", "safety stock days", "safety stock");
    const doiCol = col("days of inventory", "days inventory");
    const minOrderCol = col("min order $", "min order", "minimum order", "min order amount ($)");
    const orderDaysCol = col("order days");
    const emailCol = col("email");
    const notesCol = col("notes");
    const dataRows = rows.slice(headerIndex + 1).filter((row) => cleanCell(row?.[vendorCol >= 0 ? vendorCol : 0]));
    if (!dataRows.length) {
      showToast("No vendor rules found in that file.", 3200, "warning");
      return;
    }
    const existingByVendor = new Map((state.vendorRules || []).map((rule) => [String(rule.vendor || "").toUpperCase(), rule]));
    const importedRules = dataRows.map((row) => {
      const vendor = cleanCell(row[vendorCol >= 0 ? vendorCol : 0]).toUpperCase();
      const existing = existingByVendor.get(vendor);
      const orderDays = cleanCell(row[orderDaysCol >= 0 ? orderDaysCol : 5])
        .split(",")
        .map((value) => cleanCell(value).toLowerCase())
        .filter((value) => DAYS_OF_WEEK.includes(value));
      return createVendorRule(vendor, {
        id: existing?.id,
        status: cleanCell(row[statusCol >= 0 ? statusCol : 1]) || existing?.status || "Active",
        safetyDays: Math.max(0, toNumber(row[safetyCol >= 0 ? safetyCol : 2]) || existing?.safetyDays || DEFAULT_VENDOR_RULE.safetyDays),
        daysOfInventory: Math.max(0, toNumber(row[doiCol >= 0 ? doiCol : 3]) || existing?.daysOfInventory || DEFAULT_VENDOR_RULE.daysOfInventory),
        minOrder: Math.max(0, toNumber(row[minOrderCol >= 0 ? minOrderCol : 4]) || existing?.minOrder || DEFAULT_VENDOR_RULE.minOrder),
        orderDays: orderDays.length ? orderDays : (existing?.orderDays || []),
        email: emailCol >= 0 ? (cleanCell(row[emailCol]) || existing?.email || "") : (existing?.email || ""),
        notes: notesCol >= 0 ? (cleanCell(row[notesCol]) || existing?.notes || "") : (existing?.notes || ""),
        updatedAt: new Date().toISOString(),
      });
    }).filter((rule) => rule.vendor);
    state.vendorRules = importedRules.sort((a, b) => compareDisplayValue(a.vendor, b.vendor));
    persistVendorRules();
    renderVendorRules();
    const synced = await syncSharedVendorRulesToSupabase(true);
    showToast(
      synced
        ? `Imported ${number.format(importedRules.length)} vendor rules and synced them.`
        : `Imported ${number.format(importedRules.length)} vendor rules. Shared sync still needs attention.`,
      3200,
      synced ? "success" : "warning"
    );
  } catch (error) {
    console.error("Vendor rule import failed", error);
    showToast("Vendor rule import failed.", 3200, "warning");
  }
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
document.querySelector("#importVendorRulesButton")?.addEventListener("click", () => document.querySelector("#vendorRulesImportInput")?.click());
document.querySelector("#exportVendorRulesButton")?.addEventListener("click", () => exportVendorRules());
document.querySelector("#vendorRuleSaveButton")?.addEventListener("click", () => saveVendorRule());
document.querySelector("#preloadVendorsButton")?.addEventListener("click", () => preloadVendorsFromInventory());
document.querySelector("#vendorStatusFilter")?.addEventListener("change", () => renderVendorRules());
document.querySelector("#vendorRulesImportInput")?.addEventListener("change", async (event) => {
  const file = event.target.files?.[0];
  await importVendorRulesFile(file);
  event.target.value = "";
});
const vendorImportBox = document.querySelector("#vendorImportBox");
if (vendorImportBox) {
  ["dragenter", "dragover"].forEach((eventName) => vendorImportBox.addEventListener(eventName, (event) => {
    event.preventDefault();
    vendorImportBox.classList.add("dragging");
  }));
  ["dragleave", "drop"].forEach((eventName) => vendorImportBox.addEventListener(eventName, (event) => {
    event.preventDefault();
    vendorImportBox.classList.remove("dragging");
  }));
  vendorImportBox.addEventListener("drop", async (event) => {
    const file = [...(event.dataTransfer?.files || [])][0];
    if (file) await importVendorRulesFile(file);
  });
  vendorImportBox.addEventListener("click", () => document.querySelector("#vendorRulesImportInput")?.click());
}

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
  const updatedAt = new Date().toISOString();
  state.vendorRules = state.vendorRules.map((r) => ids.has(r.id) ? { ...r, status, updatedAt } : r);
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
  const allVendors = vendorNamesFromLoadedData();

  if (!allVendors.length) {
    showToast("No inventory or sales data loaded yet - load your CSV files first.", 3500, "warning");
    return;
  }

  const existingNames = new Set(state.vendorRules.map((r) => r.vendor?.toUpperCase()));
  const newVendors = allVendors.filter((v) => !existingNames.has(v.toUpperCase()));

  if (!newVendors.length) {
    showToast("All vendors from your data already have rules.", 2800);
    return;
  }

  newVendors.forEach((v) => state.vendorRules.push(createVendorRule(v)));

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

// -- Pending Orders --------------------------------------------------------
function loadPendingOrders() {
  return JSON.parse(localStorage.getItem("posPendingOrders:v1") || "[]");
}
function savePendingOrders() {
  safeLocalStorageSet("posPendingOrders:v1", JSON.stringify(state.pendingOrders));
}
if (!state.pendingOrders) state.pendingOrders = loadPendingOrders();

function savePoBackorderCounts() {
  state.poBackorderCounts = state.poBackorderCounts || {};
  safeLocalStorageSet("posDashboardPoBackorderCounts:v1", JSON.stringify(state.poBackorderCounts));
}

function pendingVendorNames() {
  return new Set(
    (state.pendingOrders || [])
      .filter((po) => !po.cleared)
      .map((po) => String(po.vendor || "").toUpperCase())
      .filter(Boolean),
  );
}

function isPendingOrder(code) {
  if (!state.pendingOrders?.length) return false;
  const key = codeKey(code);
  return state.pendingOrders.some((po) => !po.cleared && (po.codes || []).map(codeKey).includes(key));
}

function poBackorderCountForCode(code) {
  const count = Math.round(toNumber(state.poBackorderCounts?.[codeKey(code)]) || 0);
  return count > 0 ? count : 0;
}

function incrementPoBackorderForCodes(codes = [], options = {}) {
  state.poBackorderCounts = state.poBackorderCounts || {};
  [...new Set(codes.map(codeKey).filter(Boolean))].forEach((key) => {
    state.poBackorderCounts[key] = poBackorderCountForCode(key) + 1;
  });
  if (options.save !== false) savePoBackorderCounts();
}

function resetPoBackorderForCode(code, options = {}) {
  const key = codeKey(code);
  if (!key || !state.poBackorderCounts?.[key]) return;
  delete state.poBackorderCounts[key];
  if (options.save !== false) savePoBackorderCounts();
}

function clearPendingForCode(code, options = {}) {
  const key = codeKey(code);
  if (!key || !state.pendingOrders?.length) return false;
  let changed = false;
  state.pendingOrders = state.pendingOrders.map((po) => {
    if (po.cleared) return po;
    const oldCodes = (po.codes || []).map(codeKey).filter(Boolean);
    if (!oldCodes.includes(key)) return po;
    changed = true;
    const nextCodes = oldCodes.filter((entry) => entry !== key);
    return {
      ...po,
      codes: nextCodes,
      cleared: nextCodes.length === 0,
      clearedAt: nextCodes.length === 0 ? new Date().toISOString() : po.clearedAt,
      clearReason: nextCodes.length === 0 ? (options.reason || "manual") : po.clearReason,
    };
  });
  if (!changed) return false;
  if (options.backorder) incrementPoBackorderForCodes([key], { save: false });
  if (options.received) resetPoBackorderForCode(key, { save: false });
  savePendingOrders();
  savePoBackorderCounts();
  if (options.persist !== false) void savePersistedState();
  if (options.render) {
    patchInventoryRow(code);
    renderOrders();
  }
  return true;
}

function recordPendingReceiptIfNeeded(code, before, after, options = {}) {
  if ((toNumber(after) - toNumber(before)) <= 2 || !isPendingOrder(code)) return false;
  const cleared = clearPendingForCode(code, { reason: "received", received: true, backorder: false, render: false, persist: options.persist });
  if (cleared) {
    if (options.patch !== false) patchInventoryRow(code);
    if (options.render !== false && activeTabName() === "ordering") renderOrders();
    if (options.toast !== false) showToast(`PO pending cleared for ${code} - stock received.`, 2600, "success");
  }
  return cleared;
}

function closeOrderVendorMenu() {
  document.querySelector(".order-vendor-menu")?.remove();
}

function exportVendorPoExcel(vendorName, items) {
  if (!items?.length) {
    showToast(`No items to export for ${vendorName}.`, 2600, "warning");
    return;
  }
  const rows = currentVendorPoEmailRows(vendorName, items);
  const grandTotal = rows.reduce((sum, row) => sum + (toNumber(row.totalCost) || 0), 0);
  const today = new Date().toLocaleDateString("en-US", { year: "numeric", month: "2-digit", day: "2-digit" });
  const fileVendor = poSafeFilenamePart(vendorName);
  const html = buildPoExcelHtml(vendorName, rows, { today, grandTotal, poNumber: "" });
  downloadBlob(`PO-${fileVendor}-${localIsoDate(new Date())}.xls`, new Blob([html], { type: "application/vnd.ms-excel;charset=utf-8" }));
}

function submitShownVendorPo() {
  const vendors = [...new Set(currentOrderRows().map((item) => cleanCell(item.vendor)).filter(Boolean))];
  if (!vendors.length) {
    showToast("No vendor orders shown for the current filters.", 3000, "warning");
    return;
  }
  if (!confirm(`Submit POs for all ${vendors.length} vendor${vendors.length === 1 ? "" : "s"} shown?`)) return;
  vendors.forEach((vendor) => submitVendorPo(vendor));
}

function openOrderVendorMenu(vendorName, anchor) {
  if (!vendorName || !anchor) return;
  const existing = document.querySelector(".order-vendor-menu");
  if (existing && existing.dataset.vendorName === vendorName) {
    existing.remove();
    return;
  }
  closeOrderVendorMenu();
  const items = currentOrderRows().filter((item) => (item.vendor || "").toUpperCase() === vendorName.toUpperCase());
  const menu = document.createElement("div");
  menu.className = "order-vendor-menu";
  menu.dataset.vendorName = vendorName;
  menu.innerHTML = `
    <button type="button" class="count-submit-btn" data-vendor-menu-action="submit">Submit PO</button>
    <button type="button" class="excel-report-btn" data-vendor-menu-action="excel">Excel</button>
    <button type="button" class="pdf-report-btn" data-vendor-menu-action="pdf">PDF</button>
  `;
  anchor.insertAdjacentElement("afterend", menu);
  menu.querySelector('[data-vendor-menu-action="submit"]')?.addEventListener("click", () => {
    submitVendorPo(vendorName);
    closeOrderVendorMenu();
  });
  menu.querySelector('[data-vendor-menu-action="excel"]')?.addEventListener("click", () => {
    exportVendorPoExcel(vendorName, items);
    closeOrderVendorMenu();
  });
  menu.querySelector('[data-vendor-menu-action="pdf"]')?.addEventListener("click", () => {
    exportVendorPoPdf(vendorName, items);
    closeOrderVendorMenu();
  });
}

function submitVendorPo(vendorName, options = {}) {
  const draftItems = state.orderSubmissionDrafts?.[String(vendorName || "").toUpperCase()] || [];
  const items = (draftItems.length ? draftItems : currentOrderRows({ ignoreSubmissionDrafts: true }))
    .filter((item) => (item.vendor||"").toUpperCase() === vendorName.toUpperCase())
    .filter((item) => !isPendingOrder(item.code))
    .map((item) => applyOrderOverride({ ...item }));
  if (!items.length) { showToast(`No items to order for ${vendorName}.`, 2800, "warning"); return; }
  const clearAt = Date.now() + 7 * 24 * 60 * 60 * 1000; // 7 days
  const totalCost = items.reduce((sum, item) => sum + orderLineCost(item), 0);
  const poNumber = nextPoNumber();
  const po = {
    id: `po-${Date.now()}`,
    poNumber,
    vendor: vendorName,
    codes: items.map((item) => codeKey(item.code)),
    items: items.map((item) => ({
      code: item.code,
      plu: item.plu || "",
      product: item.product,
      vendor: item.vendor,
      recommendedOrder: item.recommendedOrder,
      caseOrder: item.caseOrder,
      caseSize: item.caseSize,
      totalCost: orderLineCost(item),
    })),
    totalCost,
    submittedAt: new Date().toISOString(),
    clearAt,
    cleared: false,
  };
  state.pendingOrders = [...(state.pendingOrders||[]), po];
  savePendingOrders();
  void savePersistedState();
  state.adjustmentLog.unshift({
    recordedAt: new Date().toISOString(),
    user: currentAuditUser(),
    code: "-",
    product: "PO #" + poNumber + " submitted: " + vendorName + " (" + items.length + " items)",
    vendor: vendorName,
    category: "-",
    action: "PO SUBMIT",
    qtyChange: 0,
    qtyBefore: 0,
    qtyAfter: 0,
    reason: currency.format(totalCost) + " total",
  });
  safeLocalStorageSet("posDashboardAdjustLog:v1", JSON.stringify(state.adjustmentLog));
  if (options.email !== false) {
    emailVendorPo(vendorName, { items, silent: options.silentEmail, poNumber });
  }
  removeSubmittedVendor(vendorName, { render: false });
  showToast("PO submitted for " + vendorName + " - " + items.length + " items pending", 3200, "success");
  renderOrders();
  renderAdjustLog();
}

function submitAllPo() {
  if (state.orderSubmissionVendors?.length) {
    showAppConfirm({
      title: "Submit All PO",
      message: `Submit all ${state.orderSubmissionVendors.length} vendor PO${state.orderSubmissionVendors.length === 1 ? "" : "s"} currently in Ordering?`,
      confirmText: "Submit all",
      onConfirm: () => state.orderSubmissionVendors.forEach((vendor) => submitVendorPo(vendor)),
    });
    return;
  }
  const todayVendors = [...new Set((state._todayOrderVendors || []).map((vendor) => cleanCell(vendor)).filter(Boolean))];
  if (!todayVendors.length) { showToast("No vendors currently qualify for ordering today.", 3000, "warning"); return; }
  showAppConfirm({
    title: "Submit Today POs",
    message: `Submit POs for all ${todayVendors.length} vendor${todayVendors.length > 1 ? "s" : ""} shown for today?`,
    confirmText: "Submit PO",
    onConfirm: () => todayVendors.forEach((vendor) => submitVendorPo(vendor)),
  });
}

function clearExpiredPendingOrders() {
  if (!state.pendingOrders?.length) return;
  const now = Date.now();
  let changed = false;
  const expiredCodes = [];
  state.pendingOrders = state.pendingOrders.map((po) => {
    if (!po.cleared && po.clearAt && now > po.clearAt) {
      changed = true;
      expiredCodes.push(...(po.codes || []));
      return { ...po, cleared: true, clearedAt: new Date().toISOString(), clearReason: "expired" };
    }
    return po;
  });
  if (expiredCodes.length) incrementPoBackorderForCodes(expiredCodes, { save: false });
  if (changed) {
    savePendingOrders();
    savePoBackorderCounts();
  }
}
// Run on load
clearExpiredPendingOrders();

// -- Vendor Analysis Panel (inline in ordering tab) ------------------------
function exportVendorPoPdfLegacy(vendorName, items) {
  const today = new Date().toLocaleDateString("en-US", { year:"numeric", month:"long", day:"numeric" });
  const body = document.querySelector("#vpmBody");
  const rows = [...(body?.querySelectorAll("tr[data-vpm-code]") || [])];
  let grandTotal = 0;
  const rowsHtml = rows.map((tr) => {
    const qtyInput = tr.querySelector(".vpm-qty-input");
    const qty = Math.max(0, toNumber(qtyInput?.value||0));
    const uc = parseFloat(qtyInput?.dataset.unitCost||0);
    const caseSize = Math.max(1, toNumber(tr.querySelector(".vpm-line-cost")?.dataset.caseSize || 1));
    const lineCost = qty * caseSize * uc;
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
  <h1>Purchase Order - ${escapeHtml(vendorName)}</h1>
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
document.querySelector("#vendorPoModal")?.addEventListener("click", (event) => {
  if (event.target === document.querySelector("#vendorPoModal")) {
    document.querySelector("#vendorPoModal").hidden = true;
  }
});
// Also wire Esc in the global keydown handler via the existing modals check

function openVendorAnalysisPanel(vendorName) {
  const modal = document.querySelector("#vendorPoModal");
  if (!modal) return;
  const normalizedVendor = String(vendorName || "").trim();
  const draftItems = state.orderSubmissionDrafts?.[normalizedVendor.toUpperCase()] || [];
  const renderedRows = state._lastRenderedOrderRows?.length ? state._lastRenderedOrderRows.map((item) => applyOrderOverride({ ...item })) : [];
  const baseRows = draftItems.length
    ? draftItems.map((item) => applyOrderOverride({ ...item }))
    : (renderedRows.length ? renderedRows : currentOrderRows());
  const sortMode = state.vendorPoSort || "item-asc";
  const sortItems = (rows) => rows.sort((a, b) => {
    if (sortMode === "item-desc") return compareDisplayValue(b.product, a.product);
    if (sortMode === "code-asc") return compareDisplayValue(a.code, b.code);
    if (sortMode === "code-desc") return compareDisplayValue(b.code, a.code);
    if (sortMode === "stock-asc") return (a.stock || 0) - (b.stock || 0);
    if (sortMode === "stock-desc") return (b.stock || 0) - (a.stock || 0);
    if (sortMode === "velocity-asc") return (a.velocity || 0) - (b.velocity || 0);
    if (sortMode === "velocity-desc") return (b.velocity || 0) - (a.velocity || 0);
    if (sortMode === "min-asc") return (a.reorderMin || 0) - (b.reorderMin || 0);
    if (sortMode === "min-desc") return (b.reorderMin || 0) - (a.reorderMin || 0);
    if (sortMode === "max-asc") return (a.reorderMax || 0) - (b.reorderMax || 0);
    if (sortMode === "max-desc") return (b.reorderMax || 0) - (a.reorderMax || 0);
    if (sortMode === "rec-asc") return (a.recommendedOrder || 0) - (b.recommendedOrder || 0);
    if (sortMode === "rec-desc") return (b.recommendedOrder || 0) - (a.recommendedOrder || 0);
    if (sortMode === "case-asc") return (a.caseOrder || 0) - (b.caseOrder || 0);
    if (sortMode === "case-desc") return (b.caseOrder || 0) - (a.caseOrder || 0);
    if (sortMode === "caseSize-asc") return (a.caseSize || 0) - (b.caseSize || 0);
    if (sortMode === "caseSize-desc") return (b.caseSize || 0) - (a.caseSize || 0);
    if (sortMode === "unitCost-asc") return (a.unitCost || 0) - (b.unitCost || 0);
    if (sortMode === "unitCost-desc") return (b.unitCost || 0) - (a.unitCost || 0);
    if (sortMode === "totalCost-asc") return orderLineCost(a) - orderLineCost(b);
    if (sortMode === "totalCost-desc") return orderLineCost(b) - orderLineCost(a);
    return compareDisplayValue(a.product, b.product);
  });
  const items = sortItems((normalizedVendor
    ? baseRows.filter((item) => (item.vendor||"").toUpperCase() === normalizedVendor.toUpperCase())
    : baseRows.slice()));
  const rule = normalizedVendor ? state.vendorRules.find((r) => r.vendor && r.vendor.toUpperCase() === normalizedVendor.toUpperCase()) : null;
  const totalCost = items.reduce((s, item) => s + orderLineCost(item), 0);
  const minOk = !rule || !rule.minOrder || totalCost >= rule.minOrder;
  const isPending = normalizedVendor && (state.pendingOrders||[]).some((po) => po.vendor === normalizedVendor && !po.cleared);

  const titleEl = document.querySelector("#vpmTitle");
  if (titleEl) titleEl.textContent = normalizedVendor || "All sent vendors";
  const metaEl = document.querySelector("#vpmMeta");
  if (metaEl) metaEl.innerHTML =
    "<span><b>" + items.length + "</b> items to order</span>" +
    "<span><b>Total</b> " + currency.format(totalCost) + "</span>" +
    (rule && rule.minOrder ? "<span class='" + (minOk ? "" : "order-min-warn-text") + "'><b>Min order</b> " + currency.format(rule.minOrder) + " " + (minOk ? "&#10003;" : "&#9888; below min") + "</span>" : "") +
    (isPending ? "<span class='pending-badge'>&#x1F550; PO pending</span>" : "");
  const clearBtn = document.querySelector("#vpmClearPendingButton");
  if (clearBtn) {
    clearBtn.hidden = !isPending;
    clearBtn.onclick = function() { clearVendorPending(normalizedVendor); modal.hidden = true; };
  }
  const cancelBtn = document.querySelector("#vpmCancelButton");
  if (cancelBtn) {
    cancelBtn.hidden = true;
    cancelBtn.onclick = null;
  }
  const syncModalOverrides = function() {
    const body = document.querySelector("#vpmBody");
    if (!body) return;
    body.querySelectorAll(".vpm-qty-input").forEach(function(input) {
      const code = codeKey(input.dataset.code);
      const row = input.closest("tr");
      const caseSize = Math.max(1, toNumber(row?.querySelector(".vpm-line-cost")?.dataset.caseSize || 1));
      const caseQty = Math.max(0, Math.round(toNumber(input.value) || 0));
      if (!state._orderRecOverrides) state._orderRecOverrides = new Map();
      state._orderRecOverrides.set(code, caseQty * caseSize);
    });
  };
  const submitBtn = document.querySelector("#vpmSubmitButton");
  if (submitBtn) submitBtn.onclick = function() {
    syncModalOverrides();
    renderOrders();
    if (normalizedVendor) submitVendorPo(normalizedVendor);
    else submitAllPo();
    modal.hidden = true;
  };
  const excelBtn = document.querySelector("#vpmExcelButton");
  if (excelBtn) excelBtn.onclick = function() {
    syncModalOverrides();
    exportVendorPoExcel(normalizedVendor || "ALL_SENT", items);
  };
  const pdfBtn = document.querySelector("#vpmPdfButton");
  if (pdfBtn) pdfBtn.onclick = function() {
    syncModalOverrides();
    exportVendorPoPdf(normalizedVendor || "ALL_SENT", items);
  };
  const emailBtn = document.querySelector("#vpmEmailButton");
  if (emailBtn) emailBtn.onclick = function() {
    syncModalOverrides();
    emailVendorPo(normalizedVendor || "ALL_SENT");
  };

  const body = document.querySelector("#vpmBody");
  if (body) {
    body.innerHTML = items.map((sku) => {
      const pend = isPendingOrder(sku.code);
      return "<tr data-vpm-code='" + sku.code + "'>" +
        "<td style='color:#2470c4;font-weight:700;white-space:nowrap'>" + escapeHtml(sku.code) + "</td>" +
        "<td style='white-space:nowrap'>" + escapeHtml(sku.plu || "") + "</td>" +
        "<td class='sku-name' title='" + escapeHtml(sku.product) + "'><span class='sku-name-text'>" + escapeHtml(sku.product) + "</span></td>" +
        "<td class='num " + ((sku.stock||0)<0?"entry-negative":"") + "'>" + number.format(sku.stock||0) + "</td>" +
        "<td class='num'>" + formatVelocity(sku.velocity||0) + "</td>" +
        "<td class='num'>" + number.format(sku.reorderMin||0) + "</td>" +
        "<td class='num'>" + number.format(sku.reorderMax||0) + "</td>" +
        "<td class='num'>" + number.format(sku.recommendedOrder||0) + "</td>" +
        "<td class='num order-highlight'><input type='number' class='vpm-qty-input mini-input' data-code='" + sku.code + "' data-unit-cost='" + (sku.unitCost||0) + "' value='" + (sku.caseOrder||0) + "' min='0' style='width:4rem;text-align:center;font-weight:700' /></td>" +
        "<td class='num'>" + currency.format(sku.unitCost||0) + "</td>" +
        "<td class='num vpm-line-cost' data-unit-cost='" + (sku.unitCost||0) + "' data-case-size='" + (sku.caseSize||1) + "'>" + currency.format(orderLineCost(sku)) + "</td>" +
        "<td style='text-align:center'>" + (pend ? "&#x1F550;" : "") + "</td>" +
        "</tr>";
    }).join("") || "<tr><td colspan='12' class='empty-cell'>No items to order for this vendor right now.</td></tr>";

    body.querySelectorAll(".vpm-qty-input").forEach(function(input) {
      input.addEventListener("input", function() {
        var row = input.closest("tr");
        var uc = parseFloat(input.dataset.unitCost || 0);
        var caseSize = Math.max(1, toNumber(row.querySelector(".vpm-line-cost")?.dataset.caseSize || 1));
        var qty = Math.max(0, toNumber(input.value)||0);
        var lineCell = row.querySelector(".vpm-line-cost");
        if (lineCell) lineCell.textContent = currency.format(qty * caseSize * uc);
        if (!state._orderRecOverrides) state._orderRecOverrides = new Map();
        state._orderRecOverrides.set(codeKey(input.dataset.code), qty * caseSize);
        var total = 0;
        body.querySelectorAll(".vpm-qty-input").forEach(function(inp) {
          var parentRow = inp.closest("tr");
          var lineCaseSize = Math.max(1, toNumber(parentRow?.querySelector(".vpm-line-cost")?.dataset.caseSize || 1));
          total += Math.max(0, toNumber(inp.value)||0) * lineCaseSize * parseFloat(inp.dataset.unitCost||0);
        });
        var dispEl = document.querySelector("#vpmTotalDisplay");
        if (dispEl) dispEl.textContent = currency.format(total);
        var gt = document.querySelector("#vpmGrandTotal");
        if (gt) gt.textContent = "Grand Total: " + currency.format(total);
      });
    });
  }
  modal.querySelectorAll("[data-vpm-sort]").forEach((th) => {
    th.style.cursor = "pointer";
    th.onclick = () => {
      const base = String(th.dataset.vpmSort || "item-asc").replace(/-(asc|desc)$/, "");
      const current = state.vendorPoSort || "";
      state.vendorPoSort = current === `${base}-asc` ? `${base}-desc` : `${base}-asc`;
      openVendorAnalysisPanel(normalizedVendor);
    };
  });

  var gt = document.querySelector("#vpmGrandTotal");
  if (gt) gt.textContent = "Grand Total: " + currency.format(totalCost);
  state._vaVendorName = normalizedVendor;
  modal.hidden = false;
}

function exportVendorPoPdf(vendorName, items) {
  var today = new Date().toLocaleDateString("en-US", { year:"numeric", month:"long", day:"numeric" });
  var body = document.querySelector("#vpmBody");
  var rows = body ? Array.from(body.querySelectorAll("tr[data-vpm-code]")) : [];
  var grandTotal = 0;
  var rowsHtml = rows.map(function(tr) {
    var qtyInput = tr.querySelector(".vpm-qty-input");
    var cases = Math.max(0, toNumber(qtyInput ? qtyInput.value : 0));
    var uc = parseFloat(qtyInput ? qtyInput.dataset.unitCost : 0);
    var caseSize = Math.max(1, toNumber(tr.querySelector(".vpm-line-cost")?.dataset.caseSize || 1));
    var qty = caseSize <= 1 ? cases : cases * caseSize;
    var lineCost = qty * uc;
    grandTotal += lineCost;
    var cells = Array.from(tr.cells);
    return "<tr><td>" + (cells[0]?cells[0].textContent.trim():"") + "</td>" +
      "<td>" + (cells[1]?cells[1].textContent.trim():"") + "</td>" +
      "<td>" + (cells[2]?cells[2].textContent.trim():"") + "</td>" +
      "<td class='num'><b>" + qty + "</b></td>" +
      "<td class='num'>" + cases + "</td>" +
      "<td class='num'>" + (cells[9]?cells[9].textContent.trim():"") + "</td>" +
      "<td class='num'><b>" + currency.format(lineCost) + "</b></td></tr>";
  }).join("");
  var html = "<!DOCTYPE html><html><head><meta charset='UTF-8'><title>PO - " + escapeHtml(vendorName) + "</title>" +
    "<style>body{font-family:Arial,sans-serif;font-size:11px;color:#111;padding:20px;margin:0;background:#fff}" +
    ".head{display:flex;align-items:center;justify-content:space-between;border-bottom:3px solid #bd8c22;padding-bottom:10px;margin-bottom:10px}" +
    ".brand{display:flex;align-items:center;gap:12px}.brand img{width:74px;height:auto}.brand h1{font-size:22px;margin:0;color:#111;letter-spacing:.3px}" +
    ".po-label{color:#6f5014;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.7px}.meta{display:flex;gap:10px;flex-wrap:wrap;margin:10px 0 14px;font-size:10px;color:#111}" +
    ".meta span{background:#fff7dd;border:1px solid #ead28e;padding:4px 9px;border-radius:999px}" +
    "table{width:100%;border-collapse:collapse}th{background:#111;color:#f5d57b;text-align:left;padding:6px 7px;font-size:9px;text-transform:uppercase;border-bottom:2px solid #bd8c22}" +
    "td{padding:5px 7px;border-bottom:1px solid #eee;color:#111}.num{text-align:right}" +
    ".grand{font-size:15px;font-weight:800;text-align:right;padding:12px 6px;border-top:2px solid #111;color:#111}" +
    "@media print{body{padding:8px}}</style></head><body>" +
    "<div class='head'><div class='brand'><img src='data:image/jpeg;base64," + PO_LOGO_GOLD_JPEG_BASE64 + "' alt='Cee Kay'><h1>Cee Kay Order</h1></div><div class='po-label'>Purchase Order</div></div>" +
    "<h1 style='font-size:18px;margin:0 0 6px;color:#111'>" + escapeHtml(vendorName) + "</h1>" +
    "<div class='meta'><span><b>Date</b> " + today + "</span><span><b>Items</b> " + rows.length + "</span><span><b>Total</b> " + currency.format(grandTotal) + "</span></div>" +
    "<table><thead><tr><th>Code</th><th>PLU</th><th>Item</th><th>Qty</th><th>Cases</th><th>Unit Cost</th><th>Total</th></tr></thead>" +
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

function clearVendorPending(vendorName, options = {}) {
  if (!state.pendingOrders?.length) return;
  const target = String(vendorName || "").toUpperCase();
  const clearedCodes = [];
  let changed = false;
  state.pendingOrders = state.pendingOrders.map((po) => {
    if (po.cleared || String(po.vendor || "").toUpperCase() !== target) return po;
    changed = true;
    clearedCodes.push(...(po.codes || []));
    return { ...po, cleared: true, clearedAt: new Date().toISOString(), clearReason: options.reason || "manual" };
  });
  if (!changed) return;
  if (options.backorder !== false) incrementPoBackorderForCodes(clearedCodes, { save: false });
  savePendingOrders();
  savePoBackorderCounts();
  void savePersistedState();
  renderInventory();
  renderOrders();
  showToast(`Pending cleared for ${vendorName}`, 2400, "success");
}

// -- Final count export (called after Submit & Apply) ----------------------
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
  const syncStats = countSessionSyncStats(session);
  const syncWarning = syncStats.pending || syncStats.failed || session.finalizedWithUnsyncedEntries
    ? "WARNING: unconfirmed shared-sync entries are included."
    : "";
  state.finalCountReports.unshift({ sessionId: session.id, label: countSessionLabel(session), date: dateStr, submittedAt: new Date().toISOString(), entries, syncWarning });
  safeLocalStorageSet("posFinalCountReports:v1", JSON.stringify(state.finalCountReports.slice(0,30)));
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
    ["Final Inventory Count - " + report.label, "", "", "", "", report.date, report.syncWarning || "Sync: all entries confirmed."],
    [],
    ["Code", "Item", "Vendor", "Category", "Qty Before", "QTY", "NULL", "Variance", "Scanned"],
    ...report.entries.map((e) => [e.code, e.product, e.vendor, e.category, e.qtyStart, e.qty, e.scanned ? "" : "NULL", e.variance, e.scanned ? "Yes" : "No (0)"]),
  ];
  const ws = xlsx.utils.aoa_to_sheet(data);
  ws["!cols"] = [16,32,14,14,10,10,8,10,10].map((w)=>({wch:w}));
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

// -- Reports tab - load saved final count reports ---------------------------
if (!state.finalCountReports) {
  state.finalCountReports = JSON.parse(localStorage.getItem("posFinalCountReports:v1") || "[]");
}

// -- User/PIN auth system ---------------------------------------------------
const AUTH_KEY = "posAuthUsers:v1";
const SUPABASE_URL = "https://mqkrgieotabpptsbosdh.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_i0n6EMZnW-E40Dx8Od_8mg_eD9GChBy";
const SYNC_STATE_ROW_ID = "pos-dashboard-main";
if (!state.authUsers) state.authUsers = [];
if (!state.authUsersLoaded) state.authUsersLoaded = false;

function defaultAuthUsers() {
  return ENABLE_SHARED_SYNC
    ? []
    : [{ name: "Admin", pin: "0000", role: "admin", id: "default-admin", active: true }];
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
  safeLocalStorageSet(AUTH_KEY, JSON.stringify(normalized));
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
    if (!silent) showToast("Using saved login users - Supabase sync unavailable.", 3200, "warning");
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

// -- Render Settings tab -----------------------------------------------------
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

// -- Lock screen ------------------------------------------------------------
function showLockScreen() {
  const overlay = document.querySelector("#lockScreen");
  if (overlay) {
    overlay.classList.remove("lock-dismissed");
    overlay.tabIndex = -1;
    setTimeout(() => overlay.focus?.(), 0);
  }
  state.currentUser = null;
  state._roleRestrictionApplied = false;
  resetUiCriteriaOnStartup();
}

function lockApp(message = "Session locked.") {
  showLockScreen();
  clearTimeout(state._idleLogoutTimer);
  if (message) showToast(message, 2200, "warning");
}

function updateMetricsSummaryMode() {
  const zone = document.querySelector("#metricsHoverZone");
  const bar = zone?.querySelector(".metrics-peek-bar");
  if (!zone || !bar) return;
  zone.classList.toggle("pinned", !!state.metricsPinned);
  bar.textContent = state.metricsPinned
    ? "Sales summary - pinned (click to return to hover)"
    : "Sales summary - hover to expand (click to pin)";
}

function resetIdleLogoutTimer() {
  clearTimeout(state._idleLogoutTimer);
  const overlay = document.querySelector("#lockScreen");
  if (!state.currentUser || !overlay || !overlay.classList.contains("lock-dismissed")) return;
  state._idleLogoutTimer = setTimeout(() => lockApp("Logged out after 5 minutes of inactivity."), IDLE_TIMEOUT_MS);
}

let authUsersRefreshPromise = null;

function ensureAuthUsersLoadedForLock() {
  if (!ENABLE_SHARED_SYNC) return Promise.resolve(loadUsers());
  if (state.authUsersLoaded) return Promise.resolve(loadUsers());
  if (!authUsersRefreshPromise) {
    authUsersRefreshPromise = refreshUsersFromSupabase({ silent: true })
      .catch(() => loadUsers())
      .finally(() => { authUsersRefreshPromise = null; });
  }
  return authUsersRefreshPromise;
}

async function tryUnlock(pin) {
  let user = verifyPin(pin);
  if (!user) {
    const disp = document.querySelector("#lockPinDisplay");
    if (disp && ENABLE_SHARED_SYNC && !state.authUsersLoaded) {
      disp.textContent = "Loading users...";
      disp.style.color = "";
    }
    await ensureAuthUsersLoadedForLock();
    user = verifyPin(pin);
  }
  if (!user) {
    const disp = document.querySelector("#lockPinDisplay");
    if (disp) { disp.textContent = "Wrong PIN"; disp.style.color = "#c0392b"; setTimeout(() => { disp.textContent = ""; disp.style.color = ""; }, 1000); }
    return false;
  }
  state.currentUser = user;
  const overlay = document.querySelector("#lockScreen");
  const disp = document.querySelector("#lockPinDisplay");
  if (disp) disp.textContent = "";
  if (overlay) overlay.classList.add("lock-dismissed");
  applyRoleRestrictions(true);
  updateMetricsSummaryMode();
  resetIdleLogoutTimer();
  if (appInitDone) switchTab(isUserRole() ? "scanmode" : "dashboard");
  else showToast("Loading your dashboard...", 1600, "success");
  Promise.resolve(bootAppIfNeeded())
    .then(() => {
      switchTab(isUserRole() ? "scanmode" : "dashboard");
      showToast("Welcome, " + user.name + "!", 1200, "success");
    })
    .catch((error) => {
      console.error("App boot failed after unlock", error);
      showToast("App loading hit an error. Refresh and try again.", 3200, "warning");
    });
  return true;
}

// -- Render ordering vendor filter ------------------------------------------
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

// -- PO Modal: case order auto-adjusts when rec order changes --------------
function calcCaseOrder(recOrder, caseSize) {
  const orderQty = Math.max(0, toNumber(recOrder) || 0);
  const size = Math.max(1, toNumber(caseSize) || 1);
  if (!orderQty) return 0;
  if (size <= 1) return orderQty;
  if (orderQty <= size) return 1;
  return Math.max(1, Math.ceil(orderQty / size));
}

function orderLineUnits(item) {
  const cases = Math.max(0, toNumber(item.caseOrder) || 0);
  const size = Math.max(1, toNumber(item.caseSize) || 1);
  return size <= 1 ? cases : cases * size;
}

function orderLineCost(item) {
  return orderLineUnits(item) * (toNumber(item.unitCost) || 0);
}

function nextPoNumber() {
  const existingMax = Math.max(0, ...(state.pendingOrders || []).map((po) => Number(po.poNumber || 0)).filter(Number.isFinite));
  const storedNext = Number(localStorage.getItem("posDashboardNextPoNumber:v1") || "1") || 1;
  const poNumber = Math.max(1, existingMax + 1, storedNext);
  safeLocalStorageSet("posDashboardNextPoNumber:v1", String(poNumber + 1));
  return poNumber;
}

function vendorRuleForName(vendorName) {
  const key = cleanCell(vendorName).toUpperCase();
  if (!key) return null;
  return state.vendorRules.find((rule) => cleanCell(rule.vendor).toUpperCase() === key) || null;
}

function vendorEmailForName(vendorName) {
  return cleanCell(vendorRuleForName(vendorName)?.email || "");
}

function itemsToPoEmailRows(items) {
  return (items || []).map((item) => {
    const caseSize = Math.max(1, toNumber(item.caseSize) || 1);
    const cases = Math.max(0, Math.round(toNumber(item.caseOrder) || 0));
    const qty = caseSize <= 1 ? cases : cases * caseSize;
    const unitCost = toNumber(item.unitCost) || 0;
    return {
      code: item.code || "",
      plu: item.plu || "",
      product: item.product || "",
      qty,
      cases,
      unitCost,
      totalCost: qty * unitCost,
    };
  }).filter((row) => row.cases > 0 || row.qty > 0);
}

function currentVendorPoEmailRows(vendorName, sourceItems = null) {
  if (Array.isArray(sourceItems)) return itemsToPoEmailRows(sourceItems);
  const body = document.querySelector("#vpmBody");
  const domRows = Array.from(body?.querySelectorAll("tr[data-vpm-code]") || []);
  if (domRows.length) {
    return domRows.map((tr) => {
      const cells = Array.from(tr.cells);
      const input = tr.querySelector(".vpm-qty-input");
      const caseSize = Math.max(1, toNumber(tr.querySelector(".vpm-line-cost")?.dataset.caseSize || 1));
      const cases = Math.max(0, Math.round(toNumber(input?.value || cells[8]?.textContent || 0)));
      const unitCost = toNumber(input?.dataset.unitCost || cells[9]?.textContent || 0);
      const qty = caseSize <= 1 ? cases : cases * caseSize;
      return {
        code: cells[0]?.textContent.trim() || "",
        plu: cells[1]?.textContent.trim() || "",
        product: cells[2]?.textContent.trim() || "",
        qty,
        cases,
        unitCost,
        totalCost: qty * unitCost,
      };
    }).filter((row) => row.cases > 0 || row.qty > 0);
  }
  const vendorKey = cleanCell(vendorName).toUpperCase();
  return itemsToPoEmailRows(currentOrderRows()
    .filter((item) => !vendorKey || cleanCell(item.vendor).toUpperCase() === vendorKey)
    .map((item) => applyOrderOverride({ ...item })));
}

function csvValue(value) {
  const text = String(value ?? "");
  return /[",\n\r]/.test(text) ? `"${text.replace(/"/g, '""')}"` : text;
}

function base64FromUtf8(value) {
  const bytes = new TextEncoder().encode(String(value ?? ""));
  let binary = "";
  bytes.forEach((byte) => { binary += String.fromCharCode(byte); });
  return btoa(binary);
}

function poSafeFilenamePart(value) {
  return cleanCell(value || "Vendor").replace(/[^A-Za-z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "Vendor";
}

function poAscii(value) {
  return String(value ?? "").normalize("NFKD").replace(/[^\x20-\x7E]/g, "").replace(/\s+/g, " ").trim();
}

function pdfTextEscape(value) {
  return poAscii(value).replace(/\\/g, "\\\\").replace(/\(/g, "\\(").replace(/\)/g, "\\)");
}

const PO_LOGO_GOLD_JPEG_BASE64 = "/9j/4AAQSkZJRgABAQEAYABgAAD/2wBDAAUEBAQEAwUEBAQGBQUGCA0ICAcHCBALDAkNExAUExIQEhIUFx0ZFBYcFhISGiMaHB4fISEhFBkkJyQgJh0gISD/2wBDAQUGBggHCA8ICA8gFRIVICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICAgICD/wAARCAEWAWgDASIAAhEBAxEB/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBAQAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBSExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkpOUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwD7LooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigDL1PxBpOkSLFfXYSVhkIqlmx64HSqd14x0C2sTcpfJcHHyxR8ux9Mdvxrm/iFZFb2zv1HEiGJj7g5H6E/lXCstfK43NsRh606SitNv6uetQwdOpBTbZ7F4b19tf0+W7azNqI5CnL7g3GeDx61f0vUE1TTkvohiORnCe6hioP44zXOWSnSPhe0iDbI1s0mf8Aafof1FW/A0gfwfbRj/lkzp/48T/WvVw2Im506VR3bjd/ev8AgnJVpxtKcdk7HS0UUV6pyBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAY/ibTv7T8P3EKrmVB5sf+8O34jI/GvHnX5CR6V7zXkvifSv7O1ueNFxDN+9j+h6j8Dn9K+Vz7DfDXj6P9D1sBV3ps7PXl2/D4xr0EMK/hlazfh9dYhvdPY8qwmUexGD/ACH51qXpF58PN685tUb8sE/yrh9Hvm0nWYL3nYp2yAd0PX/H8KMTXVDGUavTlX3O4qUOejOHW567RTUdJI1kjYMjAMrDoQadX1R5QUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABWF4n0g6rpRMK5uYMvH/teq/j/MCt2isa1KNam6c9mXCbhJSXQ5PwpML/AMLT6ax+eLfFg/3WyR/M/lXEvERkMORwa9HbS30/Xf7UsFzFcfLcwj9HX8eo+tcnrll9l1m4QDCu3mL9Dz/PNfJZjQnChTU94e76ro/66nq4epFzly9df8y34X8QrZhdNv3xAT+6lPRP9k+3v2rvAQRkcivIHj9q1dK8R3+khYW/0m2H/LNzyv8Aunt9K0y/NvZRVKvstn2FiMLzvnhuelUViWXijR70BTcC2kP8E3y/r0/WtpHSRQ6MGU9CpyK+ppVqdVXpyTPLlCUHaSsLRRRWxAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABSMdqFgpbAzgdTS1HDKk8CTR52uMjNJtXsPzMya41+Q/6Lp9vCv/AE2lyfyFUJ77xVaKZJNPt5kHXy8n+Rz+ldNRXFPCylqqsk/l+VjVVUvso5O28aREhbyyeP1aNt2PwOKfrLWWsaet/YTLK8H3wPvBT6jrx/jWlqnh+x1PMhXybg/8tUHX6jvWNo2gX2neIN06K9v5bKZFPytnHBHWvKrU8Y/9nre/CWl7befy/pnVGVH+JDRroc48dV3j9q6DVtONheFAD5L8xt7en1FZTx18tVpSpTcJrVHpQmpK6Mt489qbFNc2jb7W4lgPrG5X+VXnjqBkAOSuR6HvWSundGl77lqHxZ4gtgB9rWdR2mQH9Rg1oxfEK9T/AI+dMhk943K/zzWfb6fpd8RGt+1jOeizqGQn2YYx+IqW58Fa1GMwpDcr1Bjkxn88V61KrmCjzUpOS8tfw1aOeUcO3aaS/A2I/iNYH/X6bcx/7hVv6irsPj/w5KwEk08HvJCf6Zrz660TVbXJn064QDv5ZI/MVkyR4JB4Poav+2MbSdp2+at/kH1OhP4fwZ7nY6xpWpD/AEHUIJz/AHUcbvy61er53ZCrBlyrDkEcEV0Ok+N9d0oqksv2+3H/ACznPzAezdfzzXoYfPot2rxt5r/L/hznqZc1rTdz2eiszRNcsde04XlkxGDtkjb70beh/wAe9adfSwnGpFTg7pnlyi4u0twoooqyQooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKpaYwFqbVv9ZbsY2H8j+IxV2qbRhdailHG+FlPuQQR/M1hUupRmvT77f8AALjqmiKXWLGJzFcmSA9PnjYZrEsPEC2btazb5rVWIjk/jC9sjvXUyxRTRmOaNZEPUMMisG88MRSMXs5vKJ/gfkfgetebi6eNUlUotO3yf52f4HRSlSs4z6mvDqNhOm+K7iI92AI/A1PHLFKC0UiSAd1INcvb+FJGmBvJk8odo85P4npXQWmm2NiS1pbrEzDBI6kV0YWriqmtWCivXX7v+CZ1I04/C7jr6yhv7RreYYzyrDqp9RXB3VtNZXTW1yuGXoezD1Fei1R1LTYNStvLk+WReUkHVT/hWGY5esTHnh8a/HyLw9f2bs9jgGQEZFQPHV25t7iwuWt7lNrDoezD1FR4Vxx1r4mUXFuMlZo9hO6ujNeP2q7p2t6lpRCwS+ZCP+WMnK/h6fhSPHVd4/anTqTpS5qbsxtKStJXO80zxRp2oFYpG+y3B/gkPB+h6GtW4sbK6BFzaQzA/wDPRA3868meLPatTS/Eeo6Uyxs5ubYcGKQ8gf7J7fyr6LDZyn7mKV/P/Nf5fccNTB/apM6e+8D6BeKTHA1o5/ihbA/I5FcHr/gzUdGja5Rhd2i9ZEGGQf7Q/qK9X07UbXVLQXNpJuU8FTwVPoRViVolhdpyoiAO4v0x3zmvSr5bhcTDngkr7Nf1Y56eKrUpWevkzx3wPey2Hi23iVj5V2DFIvY8EqfwI/U17LXmHhLR0uvFsuo2yH+zrSRzE5HDE5CgfQHP5V6fWeSRnHDtS2u7f16l4+UXUVuwUUUV7p54UUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFIVVipIBKnIPpS0UAFFFFABRRSMyohd2CqBkknAFAC0UAggEHIPeigClqWnQ6laGGQYcco/dTXn0sctrcyW067ZIzgivTq5XxfYgwR6lGPmQiOTHcHofwP868DN8GqlP28V7y381/wDuwlW0uR7M50OGGDTWjquklTpJXx56tiF46rPHWkVDDIqF46VhpjdI1GXSNSS4QkxN8sqf3l/xFej3Om6fqIWS6gW4QgEBmJU/hnFeZPHXpeiljoNiX6+Sv8AKvpcknz89CavHc8/GK1px0ZciiigiWKGNY414CoMAfhT6KK+qStojywooopgFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFI6JIjI6hlYYIIyCKWuf1rxBLYM1tFaSJMfuySAbfqPWufEV6dCDnU2NKcJTdom5HDHDAIIl2RqMAAngUsUUcESxRLtRegrjIbq8ljEo16FVYfvFeYhgfYY4/CrWka3dC7+xy77xWbCOv3v1xkfWvNp5nRcoxcbX0W36G8sPOzd7nWVR1mET6HexnvCxH1AyP5VeqhrMy2+g30zHAWB/zxgV6la3spc21mc8L8yseVxy5ANWUkrLjkwAKsJJX5mmfSNGoklTBgw5rOST3qwklUmQ0TSR8HFekWUXkafbw4xsjVf0rhdJgN7qcEGMrnc3sBya9Br6rIqTtOr8v6/A8zGS2iFFFFfTHnBRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQBi6VrqX2q6hpcyrHc2krBQD/rEzwfr61f1DT7fUrQ29wpx1Vh1U+orhINNutR8TeIL7TJ/J1CxulMJJ+V8ghlb64rqNK8TWl6/wBjvl/s/Uk4kt5vlyf9knqK8fD4lVIulieraTezV2rev/DnZUpcr5qfS1/LQzpPCt95qFLyCRU4UyJg498da2dI0VNMDSPL507DG7GAo9AK1qoajrFjpiD7RLmVuEhT5nc+gFaxwWFw0vbbW7sh1qtRchannWBASMux2ovdj6Vxvj7V1t9Oh0hHBnuMPLjsgP8AU/yNXtQ1hdEsW1zWwFupAUtbNTyvt9em49q8kvNSudT1Ca+u5N80zZPoPQD2A4rzM0zDlpOlHeX4Lz8327HZhMNzS53svz/4BaSSrKSVlpJVhJPevkEz2GjUSSrCSVmws8kixxqzuxwqqMkn2Fei+G/C5tCl/qagz9Y4eoj9z6n+Vd2EwlTFT5YbdX2OWtVjSV5Gj4b0prCzM9wuLiYZIP8AAvYfX1rdoor9AoUY0Kapw2R8/ObnJyYUUUVuQFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUV4z8dvjrY/B/SrO2tbFNU8Q6iC1tauxWOOMHBkkI5xngAdSDyMV4xc/tDftCeDbez8R+O/hvaR+Hrp1GTayW5G7kKH3tsYjOA60AfZtFcXpfj/TvE/wgk+IPhoiSCTTpruFJxykiK2Y3APUMpU4Pbg96+V/C37S37QXjb7WPCfgPR9Z+x7PP+y2czeVuztz++77W/I0AfWHgyzvon1m9v7WS2kvLxnVZBglex/WtvVNE0zWIgl/arIy/dkHDr9CK8k+DnjX40+J/Eeo2vxN8DweH9PitfMt5o7aSIyS7wNuWkbPyknp2rybxH+0j8YE+M+vfD/wd4W0fV5rK/uLa1hFtK80iRk8nEoBO1STwK5aeGhCl7GWq8/vNZVZSnzrRn00PBnlfJb6/qUUP/PMSdKLmxtfC1m13puj3er6i/CEAu2fUt0UfrXivgb4k/tJat480fTfFnwytdN0O4nCXl0lnKhiTB+bJlIHbsaPjn8evGHw0+LOieE9C0/Srixv7SGeR7uKRpAzzOhwVdRjCjt1rL6jRjrTVn0e9vS5ft5v4ndG5qmneNta1F7/AFHRr6SVuABEQqL/AHVHYVXXwz4lH/MCvf8Av3UH7QH7Q1z8MdV07wv4Qs7PU/Ec+JblLlWkS3jPCLtRgS7HkDPAHT5hXr3w+ufGt94Hsb/4gW9jaa7dL50lpZRsiWyn7sbbmYlwPvc4ycdsnypZDTnJylUbb9DtWYySsoo8vXw54jHXQ7z/AL91qaX4O1+/uPLktGsox96W4GMfQdSa5n9or9oDWPhPrWi6F4XtNOu9QuoHuroXqO4ij3bY8BWXkkP1/uj1r1L4R+PV+JXwr0bxYyRRXdwhju4YshY50JVwASSASMgEnhhThkFBNOUm18iZZjUa0SOg0Tw1p2hoGhUzXJGGnk+9+HoPpW1XzJqnx88beDf2loPh14107SI/Dt3dLHb30EMiSGGXIhkJLleGIV+MfK3TFaf7QPx71z4Z+JNC8LeDdPsdS1q/QzTx3UbybVZtkSqEZTuZg35D1r3qVKFKKhTVkedKcpvmk7s+iKKzdAOtHw3pzeJPsw1loEa8FqpWJZSMsqZJOAeBzzitKtSQooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigD4P8A2g83f7Z/hq2vxvtRJpkQV/u+WZssPpktX2t4q0Lw34k8NXOkeLLS3u9ImKGaK5bahKsCuTkY+YDvXh37R/wK1j4jTab4w8Fyxp4m0uMQmB5BH9pjDFk2ueFdWLYzgEHqMDPlniXSv2sPixoUHgbxH4at7DTfMQ3Ny4igWYqchpGDnIBwcIvUDigD6gn8PeFvDHwc8RaV4Os7a00pbC7dYraQugcxNu5JNfDH7PupfGfT5PEX/CodFs9TaQW/2/7T5f7vHmeXje69cv0z0r7Y8K/DVfAXwBn8BaW/2+8GnXKvKAF+0XEiMSRnoNxAGewGa+Uvhp4L/ae+FDao3hTwHEDqYiE/2p7eX/V7tuP3ox98/pQB9YfCC/8Aixf6BqMnxa0m103UVuQLVLby8NFtGSdjsPvZr4vuLrxzZ/tneJbj4cWMN94lXVr/AOzQz7djAhw+dzKOE3HrX1h8Hdb+POqeIdRi+LPh210vTUtQ1q8KxgvLvHB2yN/DntXgOs/Df48eHP2j/EPxF8EeDPtJfUruWzmmlgaOSOXcu7aZAeVbvigD3b4T6v8AtEX3jSWH4qeHrDTtCFo7JLb+TuM+5No+SRjjG/t2r57/AGx7l7L48+H7yNVZ7fSIJVVuhK3ExAPtxXsXgbxN+1ReePdHtvGnhCxs/D0k+L2eNIQyR4PIxKT1x0Brmf2kfhN498c/GrQNb8OeGJNV0m3sIIbiUSxKoInkZlIdgT8rDt3oA8e+BevaTq37UNprvxPjkn1PVJWntJrj5US9kw0LMp/hIOE7AlCOAMfo2SFBJIAHJJ7V8uftOfArWPGVzpPjH4f6X5+uW220uraB0iMkQ5jkUkgZQjb1zgr/AHa7e41H4tan+zDfWt34Tuo/iBJaHTGgE0WZS2ENyGD7R8hLdR8wIxjFAHzJp+peFPjD+1V4l13xjrmnaf4bSC5jtpNQuEijdFj8iDbuI+b5hLgdwTXbfsaeLW03xH4n+Gt7dRyK5N9aFHDI0iERy7SODuXYRjshNT/CL9knT9R8I3N58V9P1HT9Xa6ZYLW3u0G2EKuGO3cCSxbv0Aqlc/Ajxx8Lf2i9G8S/DHw7fat4as5IZi7XUW8Iw2TxHcykkruI4/iHpQB6J+118OP+Em+HEXjXTYc6r4bJeUr96S1Y/OP+AHDj0G/1ryr9nTRNa+MXxyvPin4ycXaaEkJDFcLJciMJCAP9lULn/a2nvX1v8UtI1HX/AIQeLNF0i1a61C+0yeC3hVgpkdkIAySAOfU15N+yn8P/ABh8P/CfiKy8YaJJpU93exywo8iPvUJgn5GPegD6KooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooA+QP2xvEfiTTNf8FaVo2v3+l210k7SraTtFvbdGoZtpGcAnGfU+teVfFbwn4x+HGjyalD8S/FlyIdUbSzHqImtDcEIX86AiZxJFxgnjBI9a9v8A2rPhp468bax4S1bwdoT6wunLMkyROoZGLIykhiMg7T09PpXk3jjwj8ffG+nX1g/wgg0eLUtRGq37WLEtdXAVlDMZJn2gB24XA5oA574oaJ4r+HNhAw8e+OryeVoQtzPFJBYzb4hIRHN57FiM4xsHRvSsrxZp3j/wv8O/DnitviNrF3LqQUX1kt9MG01pE822Vzv58yH5xwMYI5rrPGnw8+MvjicX+pfA62tdVZofP1C2mkEk6xxiMKytcFACqrnaoPHXrWj4n0T9prxfo2uaJrnw7E2mao8Tw2yRRIunmJgU8kh88LlPmLfKT9aAOSvND8Z2kmtoPidr7/2X4WtfEgP2mUeYZhB+5/1nAHnfe5zt6DPEMEhf4PyeO5fi541S4jul01rJY2K/a2haUAP9p/1fyEbtuf8AZrq77wj+0FfeEptFb4TRRX1zpsGj3OsR4FzPZwlSkTAy+WP9WgLBQxCjJrml+EPx3X4eS+CB8OLj7BLqaaqZdyeaJFiaILnfjbhiemc96AG6LpHirV/h8vjw/FbXIdEtLS6bVGN1KZLS7jKiG3A8z5zN5kW1uP48/c59S/Zktdc1Hwb40+IniiTWvFFtpsDQWGnPdSzGeRE82QIpYguf3ag4/iIry6z+EHx3svA2reEIvhzctZapd295LIzp5itCJAoX95jB8w5yD0FfX3w18O23hP4Fab4CvdI1+2uHs2F9LZ28kcqTy5aQpIvIKs2Aw/uigDP8LeK9N134feJfFlx4Q0ZotH09r6JtN1c3McrLE0jW8gKrJDKu0BgyY+bIzitnw/q+ga34t0jQm8JwwLqHhaLxGZRcM3ll5FTycYGQN2d2fwpNA8JaNpTeI5tWXxX4iuvEVolhfXF9YqjPbqrIqfukQE4dsucsc9ad4N8K6V4Q1U6oz+Ltbu49Oj0e1k1CzX/RbNG3LEojRc84yzZY4HNAHK/DP4haH4613RtLvvB+k2p1qxmvbf8As7VjeS2oiKhluo9qmEncNpyQSCKTwj8QtE8Q+NtN0C88HaRFDql7e2MH2PVzcXcDWxky89vtUojCM4YE9V9a9A8C6V4b+H3hez0LRPDmrkW0Qia7k00Ce4AJOZGVRuPNHg7SvDfgmxmg0rw5q8lxPcT3El5Lpg89zLK0pVnVQSoLYAPYCgCr4dfTfE+keL5tO8JWov8AQtVvNKtoJLpgt08IG0s2Pk3FgOhx71g+EPEug+P7fVZtF8LNYWmmaWkl3dtPIktrqLKWe1UYGWiABZsjBZRitjw9oSeGvFmpa1p2peMfsOp38+pXGkPp8Rt2mlHzHPleYAOCBv7D3q/o9jY+HdH8SWOj6Rr9xLrt/d6k4ntNoE0/VQcABQcYz07mgDvtLlkn0aymlbfJJAjMx7kqCTVuq2nQPa6VaW0uPMihRGx0yFANWaACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAiuEmkgZbebyZf4XKhgPwri9U1XxVpkm2eSMKT8sgiBVvxruaZLFHNE0U0ayIwwVYZBrhxWGnWX7ubi/J6fcb0qig/eimjy2bxh4lTO25h/78itHRPiEVcW3iCMLk4F1EvA/wB5e31H5Vf1vwgAjXOmAsByYDyR/un+lcFcWfX5a+Tq18fgqv7yTfrqn/X3nrwhh68dF/me3QzQ3MCT28qSxOMq6HIYexqSvFND1y/8N3oeJmks2bM1uTwR3I9D/k17SjrJGsiHKsAQfavqMBj44yDaVpLdHl4jDui+6Y6iiivSOUKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigArz/AMR2UQ1qcRADcA5A9SOa6vWNf07RLcvdzKZSMpCp+d/w7D3rjtPtdX8QXj3zRmKOVtxlcEKB6L6189m81WUcPTXNO99Oh6GFi4XqS0RlWmhyanqMdpGvyscu391e5r1pFVEVFGFUYA9qqWGnW+nQeXCuWP33PVjVyuzLcD9Ug+b4nv8A5GOJr+1atsgooor1TlCiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAooooAKKKKACiiigAqrqEF3c2TQWd6bKVuPOEYcqO+AeM+5z9KtUUpJSVmNOzuc/p/g/RbGc3UkUl/dsdzXF4/mMT688fpXQAADA4FFFZ06NOkrU1YqU5Td5O4UUUVqQFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAFFFFABRRRQAUUUUAf/9k=";

function buildSimplePoPdfBase64(vendorName, rows, meta) {
  const rowsPerPage = 31;
  const chunks = [];
  for (let i = 0; i < rows.length; i += rowsPerPage) chunks.push(rows.slice(i, i + rowsPerPage));
  if (!chunks.length) chunks.push([]);
  const poLabel = meta.poNumber ? `PO# ${meta.poNumber}` : "Purchase Order";
  const text = (value, x, y, size = 8, font = "F1", color = "0 0 0") =>
    `${color} rg BT /${font} ${size} Tf ${x} ${y} Td (${pdfTextEscape(value)}) Tj ET`;
  const rightText = (value, rightX, y, size = 8, font = "F1", color = "0 0 0") => {
    const str = poAscii(value);
    return text(str, rightX - (str.length * size * 0.48), y, size, font, color);
  };
  const line = (x1, y1, x2, y2, color = "0.78 0.83 0.80") => `${color} RG ${x1} ${y1} m ${x2} ${y2} l S`;
  const rect = (x, y, w, h, color) => `${color} rg ${x} ${y} ${w} ${h} re f`;
  const logoBinary = atob(PO_LOGO_GOLD_JPEG_BASE64);
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica-Bold >>",
    `<< /Type /XObject /Subtype /Image /Width 360 /Height 278 /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /DCTDecode /Length ${logoBinary.length} >>\nstream\n${logoBinary}\nendstream`,
  ];
  const pageObjectNumbers = [];
  chunks.forEach((chunk, pageIndex) => {
    const pageObj = objects.length + 1;
    const contentObj = objects.length + 2;
    pageObjectNumbers.push(pageObj);
    const commands = [
      "q 82 0 0 63 38 711 cm /Logo Do Q",
      text("CEE KAY BEAUTY SUPPLY", 136, 748, 14, "F2", "0 0 0"),
      text("Purchase Order", 136, 732, 10, "F1", "0.66 0.47 0.12"),
      rightText(poLabel, 570, 748, 10, "F2", "0 0 0"),
      text(`Vendor: ${vendorName || "Vendor"}`, 136, 718, 9, "F2", "0 0 0"),
      line(36, 704, 576, 704, "0.76 0.56 0.16"),
      text(`Date ${meta.today}`, 46, 686, 8, "F1"),
      text(`Items ${rows.length}`, 126, 686, 8, "F1"),
      text(`Total ${currency.format(meta.grandTotal)}`, 190, 686, 8, "F1"),
      pageIndex ? text(`Page ${pageIndex + 1}`, 522, 686, 8, "F1") : "",
      rect(36, 656, 540, 18, "0.95 0.90 0.77"),
      text("CODE", 42, 662, 7, "F2"),
      text("PLU", 100, 662, 7, "F2"),
      text("ITEM", 148, 662, 7, "F2"),
      rightText("QTY", 390, 662, 7, "F2"),
      rightText("CASES", 438, 662, 7, "F2"),
      rightText("UNIT COST", 510, 662, 7, "F2"),
      rightText("TOTAL", 570, 662, 7, "F2"),
      line(36, 656, 576, 656, "0.76 0.56 0.16"),
    ].filter(Boolean);
    chunk.forEach((row, index) => {
      const y = 639 - (index * 16);
      if (index % 2) commands.push(rect(36, y - 4, 540, 16, "0.98 0.99 0.98"));
      commands.push(text(poAscii(row.code).slice(0, 14), 42, y, 7));
      commands.push(text(poAscii(row.plu).slice(0, 10), 100, y, 7));
      commands.push(text(poAscii(row.product).slice(0, 42), 148, y, 7, "F2"));
      commands.push(rightText(number.format(row.qty || 0), 390, y, 7));
      commands.push(rightText(number.format(row.cases || 0), 438, y, 7, "F2"));
      commands.push(rightText(currency.format(row.unitCost || 0), 510, y, 7));
      commands.push(rightText(currency.format(row.totalCost || 0), 570, y, 7, "F2"));
      commands.push(line(36, y - 7, 576, y - 7, "0.86 0.89 0.87"));
    });
    const totalY = Math.max(72, 639 - (chunk.length * 16) - 16);
    if (pageIndex === chunks.length - 1) {
      commands.push(line(390, totalY + 16, 576, totalY + 16, "0 0 0"));
      commands.push(rightText(`Grand Total: ${currency.format(meta.grandTotal)}`, 570, totalY, 12, "F2"));
    }
    const content = commands.join("\n");
    objects.push(`<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 3 0 R /F2 4 0 R >> /XObject << /Logo 5 0 R >> >> /Contents ${contentObj} 0 R >>`);
    objects.push(`<< /Length ${content.length} >>\nstream\n${content}\nendstream`);
  });
  objects[1] = `<< /Type /Pages /Count ${pageObjectNumbers.length} /Kids [${pageObjectNumbers.map((num) => `${num} 0 R`).join(" ")}] >>`;
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  offsets.slice(1).forEach((offset) => {
    pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  });
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return btoa(pdf);
}

function buildPoCsv(vendorName, rows, meta) {
  const table = [
    ["Purchase Order", vendorName || "Vendor"],
    ["Date", meta.today],
    ["Items", rows.length],
    ["Grand Total", currency.format(meta.grandTotal)],
    [],
    ["CODE", "PLU", "ITEM", "QTY", "CASES", "UNIT COST", "TOTAL COST"],
    ...rows.map((row) => [
      row.code,
      row.plu || "",
      row.product,
      row.qty || 0,
      row.cases || 0,
      row.unitCost || 0,
      row.totalCost || 0,
    ]),
  ];
  return table.map((line) => line.map(csvValue).join(",")).join("\r\n");
}

function buildPoExcelHtml(vendorName, rows, meta) {
  const poLabel = meta.poNumber ? `PO# ${escapeHtml(meta.poNumber)}` : "Purchase Order";
  return `<!doctype html><html><head><meta charset="utf-8">
    <style>
      body{font-family:Arial,sans-serif;color:#111!important;background:#fff;}
      .head{border-bottom:4px solid #caa03d;padding:10px 0 12px;margin-bottom:12px;}
      .logo{height:70px;vertical-align:middle;margin-right:16px;}
      .title{color:#111!important;font-size:24px;font-weight:800;vertical-align:middle;}
      .sub{font-size:16px;font-weight:700;margin:12px 0 8px;color:#8a6518!important;}
      .chip{display:inline-block;border:1px solid #e2d3aa;background:#fff8e6;color:#111!important;padding:6px 10px;margin:0 6px 12px 0;border-radius:6px;}
      table{border-collapse:collapse;width:100%;}
      th{background:#f3e7c3;color:#111!important;border-bottom:2px solid #caa03d;padding:8px;text-align:left;font-size:12px;}
      td{border-bottom:1px solid #e8dfc6;color:#111!important;padding:7px;font-size:12px;vertical-align:top;}
      tr:nth-child(even) td{background:#fbfdfb;}
      .num{text-align:right;}
      .grand{font-size:18px;font-weight:700;text-align:right;border-top:2px solid #17231f;padding-top:10px;margin-top:10px;}
    </style></head><body>
    <div class="head"><img class="logo" src="data:image/jpeg;base64,${PO_LOGO_GOLD_JPEG_BASE64}" alt="Cee Kay"><span class="title">Cee Kay Order - ${escapeHtml(vendorName || "Vendor")}</span></div>
    <div class="sub">${poLabel}</div>
    <span class="chip"><b>Date</b> ${escapeHtml(meta.today)}</span>
    <span class="chip"><b>Items</b> ${number.format(rows.length)}</span>
    <span class="chip"><b>Total</b> ${escapeHtml(currency.format(meta.grandTotal))}</span>
    <table>
      <thead><tr><th>CODE</th><th>PLU</th><th>ITEM</th><th class="num">QTY</th><th class="num">CASES</th><th class="num">UNIT COST</th><th class="num">TOTAL COST</th></tr></thead>
      <tbody>${rows.map((row) => `<tr>
        <td style="mso-number-format:'\\@';">${escapeHtml(row.code)}</td>
        <td style="mso-number-format:'\\@';">${escapeHtml(row.plu || "")}</td>
        <td>${escapeHtml(row.product)}</td>
        <td class="num">${escapeHtml(number.format(row.qty || 0))}</td>
        <td class="num">${escapeHtml(number.format(row.cases || 0))}</td>
        <td class="num">${escapeHtml(currency.format(row.unitCost || 0))}</td>
        <td class="num">${escapeHtml(currency.format(row.totalCost || 0))}</td>
      </tr>`).join("")}</tbody>
    </table>
    <div class="grand">Grand Total: ${escapeHtml(currency.format(meta.grandTotal))}</div>
  </body></html>`;
}

function buildVendorPoEmail(vendorName, rows, options = {}) {
  const today = new Date().toLocaleDateString("en-US", { year: "numeric", month: "2-digit", day: "2-digit" });
  const grandTotal = rows.reduce((sum, row) => sum + (toNumber(row.totalCost) || 0), 0);
  const dateForFile = localIsoDate(new Date());
  const fileVendor = poSafeFilenamePart(vendorName);
  const poNumber = Number(options.poNumber || 0) || "";
  const poLabel = poNumber ? `PO# ${poNumber}` : "PO";
  const header = ["CODE", "PLU", "ITEM", "QTY", "CASES", "UNIT COST", "TOTAL COST"];
  const lines = rows.map((row) => [
    row.code,
    row.plu || "",
    row.product,
    number.format(row.qty || 0),
    number.format(row.cases || 0),
    currency.format(row.unitCost || 0),
    currency.format(row.totalCost || 0),
  ].join(" | "));
  return {
    subject: `[${vendorName || "Vendor"}] Cee Kay Order (${poLabel})`,
    body: [
      `Purchase Order: ${vendorName || "Vendor"}`,
      poNumber ? `PO#: ${poNumber}` : "",
      `Date: ${today}`,
      `Items: ${rows.length}`,
      `Grand Total: ${currency.format(grandTotal)}`,
      "",
      header.join(" | "),
      ...lines,
      "",
      `Copy sent to: ${DEFAULT_PO_COPY_EMAIL}`,
    ].join("\n"),
    html: `
      <p><b>Purchase Order:</b> ${escapeHtml(vendorName || "Vendor")}</p>
      <p>${poNumber ? `<b>PO#:</b> ${escapeHtml(poNumber)}<br>` : ""}<b>Date:</b> ${escapeHtml(today)}<br><b>Items:</b> ${number.format(rows.length)}<br><b>Grand Total:</b> ${escapeHtml(currency.format(grandTotal))}</p>
      <table border="1" cellpadding="6" cellspacing="0" style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:13px;color:#111;">
        <thead><tr>${header.map((label) => `<th style="background:#f3e7c3;color:#111;">${escapeHtml(label)}</th>`).join("")}</tr></thead>
        <tbody>${rows.map((row) => `<tr>
          <td style="color:#111;">${escapeHtml(row.code)}</td>
          <td style="color:#111;">${escapeHtml(row.plu || "")}</td>
          <td style="color:#111;">${escapeHtml(row.product)}</td>
          <td style="color:#111;">${escapeHtml(number.format(row.qty || 0))}</td>
          <td style="color:#111;">${escapeHtml(number.format(row.cases || 0))}</td>
          <td style="color:#111;">${escapeHtml(currency.format(row.unitCost || 0))}</td>
          <td style="color:#111;">${escapeHtml(currency.format(row.totalCost || 0))}</td>
        </tr>`).join("")}</tbody>
      </table>
      <p>Copy sent to: ${escapeHtml(DEFAULT_PO_COPY_EMAIL)}</p>`,
    attachments: [
      {
        filename: `PO-${fileVendor}-${dateForFile}.xls`,
        contentType: "application/vnd.ms-excel",
        content: base64FromUtf8(buildPoExcelHtml(vendorName, rows, { today, grandTotal, poNumber })),
      },
      {
        filename: `PO-${fileVendor}-${dateForFile}.pdf`,
        contentType: "application/pdf",
        content: buildSimplePoPdfBase64(vendorName, rows, { today, grandTotal, poNumber }),
      },
    ],
  };
}

async function sendPoEmailBackend(vendorName, rows, options = {}) {
  const vendorEmail = vendorEmailForName(vendorName);
  const toEmail = vendorEmail || DEFAULT_PO_COPY_EMAIL;
  const cc = toEmail.toLowerCase() === DEFAULT_PO_COPY_EMAIL.toLowerCase() ? [] : [DEFAULT_PO_COPY_EMAIL];
  const email = buildVendorPoEmail(vendorName, rows, { poNumber: options.poNumber });
  const payload = {
    to: [toEmail],
    cc,
    vendor: vendorName || "Vendor",
    subject: email.subject,
    text: email.body,
    html: email.html,
    attachments: email.attachments,
  };
  const response = await fetch(`${SUPABASE_URL}/functions/v1/${PO_EMAIL_FUNCTION_NAME}`, {
    method: "POST",
    headers: supabaseHeaders(),
    body: JSON.stringify(payload),
  });
  if (!response.ok) {
    const detail = await response.text().catch(() => "");
    throw new Error(`PO email backend failed (${response.status}) ${detail}`);
  }
  return response.json().catch(() => ({ ok: true }));
}

function openPoEmailDraft(vendorName, rows, options = {}) {
  const vendorEmail = vendorEmailForName(vendorName);
  const toEmail = vendorEmail || DEFAULT_PO_COPY_EMAIL;
  const { subject, body } = buildVendorPoEmail(vendorName, rows, { poNumber: options.poNumber });
  const cc = toEmail.toLowerCase() === DEFAULT_PO_COPY_EMAIL.toLowerCase() ? "" : DEFAULT_PO_COPY_EMAIL;
  const gmailUrl = "https://mail.google.com/mail/?view=cm&fs=1" +
    `&to=${encodeURIComponent(toEmail)}` +
    (cc ? `&cc=${encodeURIComponent(cc)}` : "") +
    `&su=${encodeURIComponent(subject)}` +
    `&body=${encodeURIComponent(body)}`;
  if (!vendorEmail && !options.silent) showToast(`No vendor email saved for ${vendorName}. Opening draft to ${DEFAULT_PO_COPY_EMAIL}.`, 3600, "warning");
  if (gmailUrl.length > 7000 && navigator.clipboard?.writeText) {
    navigator.clipboard.writeText(body).catch(() => {});
    showToast("PO is large, so the email body was copied too.", 3600, "warning");
  }
  const win = window.open(gmailUrl, "_blank", "noopener,noreferrer");
  if (!win) window.location.href = gmailUrl;
  return true;
}

async function emailVendorPo(vendorName, options = {}) {
  const rows = currentVendorPoEmailRows(vendorName, options.items || null);
  if (!rows.length) {
    if (!options.silent) showToast("No PO lines to email for this vendor.", 2600, "warning");
    return false;
  }
  try {
    await sendPoEmailBackend(vendorName, rows, { poNumber: options.poNumber });
    if (!options.silent) showToast(`PO emailed for ${vendorName || "vendor"} with Excel/PDF attachments.`, 3200, "success");
    return true;
  } catch (error) {
    console.warn("Backend PO email failed", error);
    if (!options.silent) {
      const message = (error?.message || "Unknown email error").replace(/\s+/g, " ").slice(0, 220);
      showToast(`PO email was not sent: ${message}`, 7600, "warning");
    }
    return false;
  }
}

function applyOrderOverride(item) {
  const overrideQty = state._orderRecOverrides?.get(codeKey(item.code));
  if (overrideQty == null) return item;
  return {
    ...item,
    recommendedOrder: overrideQty,
    caseOrder: calcCaseOrder(overrideQty, item.caseSize),
  };
}


// -- Reports tab box buttons -------------------------------------------------
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

// -- Report modal Esc + click-outside ----------------------------------------
["reportAdjustModal","reportPoModal","reportCountModal","reportLogsModal"].forEach((id) => {
  const el = document.querySelector("#" + id);
  if (!el) return;
  el.addEventListener("click", (e) => { if (e.target === el) el.hidden = true; });
});

// -- PO history rendering -----------------------------------------------------
function renderPoHistory() {
  const body = document.querySelector("#poHistoryBody");
  if (!body) return;
  const pos = (state.pendingOrders||[]).filter((po) => po.submittedAt);
  if (!pos.length) { body.innerHTML = '<tr><td colspan="7" class="empty-cell">No POs submitted yet.</td></tr>'; return; }
  body.innerHTML = pos.sort((a,b) => b.submittedAt.localeCompare(a.submittedAt)).map((po) => `<tr>
    <td><button type="button" class="text-link-button po-history-open" data-po-id="${escapeHtml(po.id)}">${new Date(po.submittedAt).toLocaleDateString()}</button></td>
    <td><button type="button" class="text-link-button po-history-open" data-po-id="${escapeHtml(po.id)}"><b>${po.poNumber ? `#${escapeHtml(po.poNumber)}` : "-"}</b></button></td>
    <td><button type="button" class="text-link-button po-history-open" data-po-id="${escapeHtml(po.id)}"><b>${escapeHtml(po.vendor)}</b></button></td>
    <td>${(po.codes||[]).length}</td>
    <td>${currency.format(toNumber(po.totalCost) || (po.items || []).reduce((sum, item) => sum + (toNumber(item.totalCost) || 0), 0))}</td>
    <td>${po.cleared ? '<span class="state-badge state-disabled">Cleared</span>' : '<span class="pending-badge">Pending</span>'}</td>
    <td><button type="button" class="secondary-button po-history-open" data-po-id="${escapeHtml(po.id)}">Review</button></td>
  </tr>`).join("");
  body.querySelectorAll(".po-history-open").forEach((button) => {
    button.addEventListener("click", () => openPoHistoryDetail(button.dataset.poId));
  });
  repairMojibakeText(document.querySelector("#reportPoModal") || document.body);
}

function clearAllPendingPo() {
  const activePos = (state.pendingOrders || []).filter((po) => !po.cleared);
  if (!activePos.length) {
    showToast("No pending purchase orders to clear.", 2400, "warning");
    return;
  }
  if (!confirm(`Clear ${activePos.length} pending PO${activePos.length === 1 ? "" : "s"}?`)) return;
  const clearedCodes = activePos.flatMap((po) => po.codes || []);
  state.pendingOrders = (state.pendingOrders || []).map((po) => po.cleared ? po : { ...po, cleared: true, clearedAt: new Date().toISOString(), clearReason: "manual-all" });
  if (clearedCodes.length) incrementPoBackorderForCodes(clearedCodes, { save: false });
  savePendingOrders();
  savePoBackorderCounts();
  void savePersistedState();
  renderPoHistory();
  renderInventory();
  renderOrders();
  showToast("All pending PO statuses cleared.", 2400, "success");
}

function openPoHistoryDetail(poId) {
  const po = (state.pendingOrders || []).find((entry) => entry.id === poId);
  const modal = document.querySelector("#poHistoryDetailModal");
  const body = document.querySelector("#poHistoryDetailBody");
  if (!po || !modal || !body) return;
  const fallbackItems = (po.items && po.items.length)
    ? po.items
    : (buildInventoryRows({ ignoreFilters: true, ignoreQuery: true, ignoreStateFilter: true })
        .map(applyOrderOverride)
        .filter((item) => (po.codes || []).includes(codeKey(item.code)))
        .map((item) => ({
          code: item.code,
          plu: item.plu || "",
          product: item.product,
          vendor: item.vendor,
          recommendedOrder: item.recommendedOrder,
          caseOrder: item.caseOrder,
          caseSize: item.caseSize,
          totalCost: orderLineCost(item),
        })));
  const totalCost = toNumber(po.totalCost) || fallbackItems.reduce((sum, item) => sum + (toNumber(item.totalCost) || 0), 0);
  document.querySelector("#poHistoryDetailTitle").textContent = `${po.vendor} PO`;
  document.querySelector("#poHistoryDetailMeta").innerHTML = `
    <span><b>PO #</b> ${po.poNumber ? escapeHtml(po.poNumber) : "-"}</span>
    <span><b>Date</b> ${escapeHtml(new Date(po.submittedAt).toLocaleString())}</span>
    <span><b>Items</b> ${fallbackItems.length}</span>
    <span><b>Total</b> ${currency.format(totalCost)}</span>
    <span><b>Status</b> ${po.cleared ? "Cleared" : "Pending"}</span>`;
  body.innerHTML = fallbackItems.map((item) => `<tr>
    <td>${escapeHtml(item.code)}</td>
    <td>${escapeHtml(item.plu || "-")}</td>
    <td class="sku-name" title="${escapeHtml(item.product)}"><span class="sku-name-text">${escapeHtml(item.product)}</span></td>
    <td>${escapeHtml(item.vendor || po.vendor)}</td>
    <td class="num">${number.format(item.recommendedOrder || 0)}</td>
    <td class="num">${number.format(item.caseOrder || 0)}</td>
    <td class="num">${currency.format(item.totalCost || 0)}</td>
  </tr>`).join("") || `<tr><td colspan="7" class="empty-cell">No stored PO detail for this order.</td></tr>`;
  modal.hidden = false;
  repairMojibakeText(modal);
}

// -- Count report history rendering ------------------------------------------
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

// -- Ordering vendor filter ----------------------------------------------------
document.querySelector("#orderVendorFilterSelect")?.addEventListener("change", () => {
  state.orderVendorFilter = document.querySelector("#orderVendorFilterSelect").value;
  renderOrders();
});

// -- Report Esc key support ----------------------------------------------------


// -- Lock screen - no lag, visual feedback, keyboard + click -----------------
(function() {
  const overlay = document.querySelector("#lockScreen");
  if (!overlay) return;
  if (state.authRequired) overlay.classList.remove("lock-dismissed");
  else { overlay.classList.add("lock-dismissed"); return; }
  ensureAuthUsersLoadedForLock().catch(() => {});

  let pin = "";
  let lastTouchKey = "";
  let lastTouchAt = 0;

  function draw() {
    const disp = document.querySelector("#lockPinDisplay");
    if (!disp) return;
    const dots = [];
    for (let i = 0; i < 4; i++) {
      dots.push(`<span class="pin-dot${i < pin.length ? " filled" : ""}"></span>`);
    }
    disp.innerHTML = dots.join("");
    disp.style.borderColor = "var(--green,#16835b)";
  }

  function press(k) {
    if (k === "clear") { pin = ""; draw(); return; }
    if (k === "del" || k === "Backspace") { pin = pin.slice(0, -1); draw(); return; }
    if (/^[0-9]$/.test(k) && pin.length < 4) {
      pin += k; draw();
      if (pin.length === 4) {
        const p = pin;
        setTimeout(async () => {
          const unlocked = await tryUnlock(p);
          pin = "";
          if (!unlocked) draw();
        }, 0);
      }
    }
  }

  function handleLockKeyEvent(e) {
    const btn = e.target.closest("[data-lock-key]");
    overlay.focus?.();
    if (!btn) return;
    e.preventDefault();
    e.stopPropagation();
    const key = btn.dataset.lockKey || "";
    const now = Date.now();
    if (key && key === lastTouchKey && now - lastTouchAt < 250) return;
    lastTouchKey = key;
    lastTouchAt = now;
    btn.classList.add("lock-key--pressed");
    setTimeout(() => btn.classList.remove("lock-key--pressed"), 120);
    press(key);
  }

  overlay.addEventListener("pointerdown", handleLockKeyEvent);
  overlay.addEventListener("touchstart", handleLockKeyEvent, { passive: false });
  overlay.addEventListener("click", handleLockKeyEvent);

  // Keyboard - capture phase, fires before other handlers
  document.addEventListener("keydown", (e) => {
    if (overlay.classList.contains("lock-dismissed")) return;
    if (e.key !== "Tab" && e.key !== "F5" && e.key !== "F12") {
      e.preventDefault(); e.stopImmediatePropagation();
    }
    press(e.key);
  }, true);

  draw();
})();
["pointerdown", "keydown", "input", "focusin"].forEach((eventName) => {
  document.addEventListener(eventName, () => resetIdleLogoutTimer(), true);
});
if (!state.authRequired || !loadUsers().length) {
  if (!state.authRequired) bootAppIfNeeded();
}

// Multi-device sync: poll a tiny shared sync_state row so devices can refresh
// products-only changes quickly without constantly refetching full sales data.
(function startSyncPoller() {
  if (!ENABLE_SHARED_SYNC) return;
  let _lastHash = "";
  let pollTimer = 0;
  // PERF-B: Use countPollMs when Count Mode is active — Kindle gets 30s, medium 10s, full 4s.
  // This prevents the 2s Supabase hammer from interrupting count entry on slow devices.
  const nextPollDelay = () => {
    if (scanModeIsActive()) return scanPerformanceProfile.scanPollMs;
    if (activeTabName() === "counts" && scanPerformanceProfile.countPollMs) return scanPerformanceProfile.countPollMs;
    return scanPerformanceProfile.normalPollMs;
  };
  async function poll(options = {}) {
    const forceRestore = options.forceRestore === true;
    try {
      const url = new URL(`${SUPABASE_URL}/rest/v1/sync_state`);
      url.searchParams.set("select", "id,last_sync_kind,updated_at,latest_inventory_date,latest_sales_date,product_count,sales_count");
      url.searchParams.set("id", `eq.${SYNC_STATE_ROW_ID}`);
      url.searchParams.set("limit", "1");
      const resp = await fetch(url.toString(), { headers: supabaseHeaders() });
      if (!resp.ok) return;
      const rows = await resp.json();
      if (!rows.length) return;
      const syncRow = rows[0];
      if (syncRow.updated_at && syncRow.updated_at === lastLocalSharedSyncAt) {
        _lastHash = [
          syncRow.id,
          syncRow.last_sync_kind,
          syncRow.updated_at,
          syncRow.latest_inventory_date,
          syncRow.latest_sales_date,
          syncRow.product_count,
          syncRow.sales_count,
        ].join("|");
        state._deferredSharedSync = false;
        return;
      }
      const hash = [
        syncRow.id,
        syncRow.last_sync_kind,
        syncRow.updated_at,
        syncRow.latest_inventory_date,
        syncRow.latest_sales_date,
        syncRow.product_count,
        syncRow.sales_count,
      ].join("|");
      if (hash === _lastHash) return;
      if (scanModeIsActive() && scanPerformanceProfile.deferRestoresWhileScanning && !forceRestore) {
        state._deferredSharedSync = true;
        return;
      }
      _lastHash = hash;
      state._deferredSharedSync = false;
      const syncKind = String(syncRow.last_sync_kind || "").toLowerCase();
      const isProductsOnly = syncKind === "products" || syncKind === "product-meta";
      const isVendorRulesOnly = syncKind === "vendor-rules";
      const isImportsOnly = syncKind === "imports";
      const isCountsOnly = syncKind === "counts";
      const restored = isVendorRulesOnly
        ? await restoreSharedVendorRulesOnlyFromSupabase({ silent: true })
        : isImportsOnly
          ? await restoreSharedDataFromSupabase({ silent: true, preferCurrentState: false })
          : isCountsOnly
            ? await restoreSharedCountSessionsOnlyFromSupabase({ silent: true })
        : isProductsOnly
          ? await restoreSharedProductsOnlyFromSupabase({ silent: true })
          : await restoreSharedDataFromSupabase({ silent: true, preferCurrentState: false });
      if (restored) {
        if (!scanModeIsActive() && !isProductsOnly && !isCountsOnly && !state.activeCountSession) {
          showToast(
            isVendorRulesOnly
              ? "\u21ba Synced shared vendor rules from another device"
              : isImportsOnly
                ? "\u21ba Synced shared import updates from another device"
                : isCountsOnly
                  ? "\u21ba Synced shared count updates from another device"
                  : "\u21ba Synced shared sales and inventory updates from another device",
            1800,
            "success"
          );
        }
      }
    } catch (_) { /* silent fail on file:// */ }
  }
  function schedule(delay = nextPollDelay()) {
    clearTimeout(pollTimer);
    pollTimer = setTimeout(async () => {
      await poll();
      schedule();
    }, delay);
  }
  window.flushDeferredSharedSync = async () => {
    if (!state._deferredSharedSync) return;
    state._deferredSharedSync = false;
    await poll({ forceRestore: true });
    schedule();
  };
  schedule(SHARED_SYNC_INITIAL_DELAY_MS);
})();
