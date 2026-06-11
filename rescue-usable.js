/* 20260610: Count workflow rescue layer — one count screen + filtered review list. */
(function () {
  function boot() {
    if (typeof state === "undefined" || typeof els === "undefined") {
      setTimeout(boot, 60);
      return;
    }

    const originalStart = typeof startCountSessionFromModal === "function" ? startCountSessionFromModal : null;
    const originalSave = typeof saveCountSession === "function" ? saveCountSession : null;
    const originalDelete = typeof deleteCountSession === "function" ? deleteCountSession : null;
    const originalRenderWorkspace = typeof renderCountsWorkspace === "function" ? renderCountsWorkspace : null;

    state._countReviewFilter = state._countReviewFilter || "all";
    state._countReviewVendor = state._countReviewVendor || "";
    state._countReviewCategory = state._countReviewCategory || "";
    state._countReviewSearch = state._countReviewSearch || "";
    state._countReviewSort = state._countReviewSort || { key: "default", dir: "desc" };

    function cc(value) { return typeof cleanCell === "function" ? cleanCell(value) : String(value || "").trim(); }
    function ck(value) { return typeof codeKey === "function" ? codeKey(value) : cc(value).replace(/\D/g, ""); }
    function esc(value) { return typeof escapeHtml === "function" ? escapeHtml(value) : String(value ?? "").replace(/[&<>"']/g, (m) => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m])); }
    function fmt(value) { return typeof number !== "undefined" && number?.format ? number.format(Number(value || 0)) : String(Number(value || 0)); }

    function countCandidatesForReview(session) {
      try {
        if (typeof currentCountSessionCandidates === "function") return currentCountSessionCandidates(session || state.activeCountSession) || [];
        if (typeof filteredCountCandidateRows === "function") return filteredCountCandidateRows(session || state.activeCountSession) || [];
      } catch (_) {}
      return [];
    }

    function latestEntryByCode(session) {
      const map = new Map();
      const entries = typeof filterUndoneCountEntries === "function"
        ? filterUndoneCountEntries(session?.entries || [], session)
        : (session?.entries || []);
      entries.forEach((entry) => map.set(ck(entry.code), entry));
      return map;
    }

    function buildReviewRows(session = state.activeCountSession) {
      if (!session) return [];
      const entryMap = latestEntryByCode(session);
      const seen = new Set();
      const rows = [];
      const addRow = (item, entry = null) => {
        const code = item?.code || entry?.code || "";
        const key = ck(code);
        if (!key || seen.has(key)) return;
        seen.add(key);
        const before = Number(entry?.originalQty ?? item?.stock ?? 0) || 0;
        const counted = entry ? (Number(entry.countedQty ?? 0) || 0) : null;
        const diff = entry ? counted - before : null;
        const status = !entry ? "null" : diff === 0 ? "pass" : "diff";
        rows.push({
          code,
          plu: item?.plu || entry?.plu || "",
          itemNumber: item?.itemNumber || entry?.itemNumber || "",
          product: item?.product || entry?.product || "",
          vendor: item?.vendor || entry?.vendor || "",
          category: item?.category || entry?.category || "",
          before,
          entered: entry ? Number(entry.inputQty ?? entry.qty ?? 0) || 0 : null,
          counted,
          mode: entry?.mode || "",
          diff,
          status,
          entry,
          item: item || entry,
        });
      };
      countCandidatesForReview(session).forEach((item) => addRow(item, entryMap.get(ck(item.code)) || null));
      // Include scanned out-of-scope items too, so nothing counted disappears.
      entryMap.forEach((entry, key) => {
        if (seen.has(key)) return;
        addRow(entry, entry);
      });
      return rows;
    }

    function reviewCounts(rows) {
      return rows.reduce((acc, row) => {
        if (row.status === "null") acc.null += 1;
        else if (row.status === "diff") acc.diff += 1;
        else acc.pass += 1;
        return acc;
      }, { null: 0, diff: 0, pass: 0 });
    }

    function ensureCountReviewControls() {
      const table = document.querySelector("#countEntryTable");
      if (!table) return;
      const section = table.closest("section") || table.parentElement;
      if (!section) return;
      let toolbar = document.querySelector("#countReviewFilterBar");
      if (!toolbar) {
        toolbar = document.createElement("div");
        toolbar.id = "countReviewFilterBar";
        toolbar.className = "count-review-filterbar";
        toolbar.innerHTML = `
          <div id="countReviewBadges" class="count-review-badges"></div>
          <div class="count-review-controls">
            <label>Status
              <select id="countReviewStatusFilter">
                <option value="all">All</option>
                <option value="needs">Needs Review</option>
                <option value="null">NULL / Not Scanned</option>
                <option value="diff">Qty Diff</option>
                <option value="pass">PASS</option>
              </select>
            </label>
            <label>Vendor <select id="countReviewVendorFilter"><option value="">All</option></select></label>
            <label>Category <select id="countReviewCategoryFilter"><option value="">All</option></select></label>
            <label>Search <input id="countReviewSearchFilter" type="search" placeholder="Code, PLU, item..." autocomplete="off" /></label>
          </div>`;
        const heading = section.querySelector(".panel-heading");
        if (heading) heading.insertAdjacentElement("afterend", toolbar);
        else section.prepend(toolbar);

        toolbar.querySelector("#countReviewStatusFilter")?.addEventListener("change", (e) => {
          state._countReviewFilter = e.target.value || "all";
          renderCountEntryRows(false);
        });
        toolbar.querySelector("#countReviewVendorFilter")?.addEventListener("change", (e) => {
          state._countReviewVendor = e.target.value || "";
          renderCountEntryRows(false);
        });
        toolbar.querySelector("#countReviewCategoryFilter")?.addEventListener("change", (e) => {
          state._countReviewCategory = e.target.value || "";
          renderCountEntryRows(false);
        });
        toolbar.querySelector("#countReviewSearchFilter")?.addEventListener("input", (e) => {
          state._countReviewSearch = e.target.value || "";
          renderCountEntryRows(false);
        });
      }

      const head = table.querySelector("thead tr");
      if (head && !head.dataset.reviewHeaders) {
        head.dataset.reviewHeaders = "1";
        head.innerHTML = `
          <th data-review-sort="code">Code</th>
          <th data-review-sort="plu">PLU</th>
          <th data-review-sort="product">Item</th>
          <th data-review-sort="vendor">Vendor</th>
          <th data-review-sort="category">Category</th>
          <th data-review-sort="before">POS Stock</th>
          <th data-review-sort="entered">Entered Qty</th>
          <th data-review-sort="counted">Physical Count</th>
          <th data-review-sort="diff">Qty Diff</th>
          <th data-review-sort="status">Status</th>`;
        head.querySelectorAll("[data-review-sort]").forEach((th) => {
          th.classList.add("sortable-count-head");
          th.title = "Click to sort";
          th.addEventListener("click", () => {
            const key = th.dataset.reviewSort;
            const cur = state._countReviewSort || { key: "default", dir: "desc" };
            state._countReviewSort = cur.key === key
              ? { key, dir: cur.dir === "asc" ? "desc" : "asc" }
              : { key, dir: key === "diff" ? "desc" : "asc" };
            renderCountEntryRows(false);
          });
        });
      }
      const colgroup = table.querySelector("colgroup");
      if (colgroup) colgroup.remove();
    }

    function updateReviewDropdownOptions(allRows) {
      const vendorSel = document.querySelector("#countReviewVendorFilter");
      const catSel = document.querySelector("#countReviewCategoryFilter");
      if (!vendorSel || !catSel) return;
      const fill = (select, values, selected) => {
        const current = selected || "";
        const opts = [`<option value="">All</option>`, ...values.map((value) => `<option value="${esc(value)}">${esc(value)}</option>`)];
        const html = opts.join("");
        if (select.dataset.lastOptions !== html) {
          select.innerHTML = html;
          select.dataset.lastOptions = html;
        }
        select.value = values.includes(current) ? current : "";
      };
      const vendors = [...new Set(allRows.map((r) => r.vendor).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
      const cats = [...new Set(allRows.map((r) => r.category).filter(Boolean))].sort((a,b)=>a.localeCompare(b));
      fill(vendorSel, vendors, state._countReviewVendor || "");
      fill(catSel, cats, state._countReviewCategory || "");
      const status = document.querySelector("#countReviewStatusFilter");
      if (status) status.value = state._countReviewFilter || "all";
      const search = document.querySelector("#countReviewSearchFilter");
      if (search && search.value !== (state._countReviewSearch || "")) search.value = state._countReviewSearch || "";
    }

    function filteredAndSortedReviewRows(rows) {
      const filter = state._countReviewFilter || "all";
      const vendor = (state._countReviewVendor || "").toUpperCase();
      const category = (state._countReviewCategory || "").toUpperCase();
      const needle = (state._countReviewSearch || "").toLowerCase().trim();
      const codeNeedle = ck(needle);
      let out = rows.filter((row) => {
        if (filter === "needs" && row.status === "pass") return false;
        if (filter === "null" && row.status !== "null") return false;
        if (filter === "diff" && row.status !== "diff") return false;
        if (filter === "pass" && row.status !== "pass") return false;
        if (vendor && String(row.vendor || "").toUpperCase() !== vendor) return false;
        if (category && String(row.category || "").toUpperCase() !== category) return false;
        if (needle) {
          const hay = [row.code, row.plu, row.itemNumber, row.product, row.vendor, row.category].join("|").toLowerCase();
          const codeHay = [row.code, row.plu, row.itemNumber].map(ck).join("|");
          if (!hay.includes(needle) && !(codeNeedle && codeHay.includes(codeNeedle))) return false;
        }
        return true;
      });
      const sort = state._countReviewSort || { key: "default", dir: "desc" };
      const dir = sort.dir === "asc" ? 1 : -1;
      const statusRank = { null: 0, diff: 1, pass: 2 };
      if (!sort.key || sort.key === "default") {
        return out.sort((a, b) => {
          const rank = statusRank[a.status] - statusRank[b.status];
          if (rank) return rank;
          if (a.status === "diff" || b.status === "diff") return Math.abs(Number(b.diff || 0)) - Math.abs(Number(a.diff || 0));
          return String(a.product || "").localeCompare(String(b.product || ""));
        });
      }
      out.sort((a, b) => {
        let av = a[sort.key], bv = b[sort.key];
        if (["before", "entered", "counted", "diff"].includes(sort.key)) {
          av = Number(av == null ? -999999 : av);
          bv = Number(bv == null ? -999999 : bv);
          return (av - bv) * dir;
        }
        return String(av || "").localeCompare(String(bv || "")) * dir;
      });
      return out;
    }

    function renderReviewBadges(allRows, visibleRows) {
      const badges = document.querySelector("#countReviewBadges");
      if (!badges) return;
      const counts = reviewCounts(allRows);
      const needs = counts.null + counts.diff;
      badges.innerHTML = `
        <span class="count-review-badge badge-needs">Needs Review: <b>${fmt(needs)}</b></span>
        <span class="count-review-badge badge-null">NULL: <b>${fmt(counts.null)}</b></span>
        <span class="count-review-badge badge-diff">Qty Diff: <b>${fmt(counts.diff)}</b></span>
        <span class="count-review-badge badge-pass">PASS: <b>${fmt(counts.pass)}</b></span>
        <span class="count-review-badge">Showing: <b>${fmt(visibleRows.length)}</b></span>`;
    }

    renderCountEntryRows = function renderCountEntryRowsRescue() {
      if (!els.countEntryBody) return;
      ensureCountReviewControls();
      const session = state.activeCountSession;
      if (!session) {
        els.countEntryBody.innerHTML = `<tr><td colspan="10" class="empty-cell">Start or continue a physical count first.</td></tr>`;
        return;
      }
      const allRows = buildReviewRows(session);
      updateReviewDropdownOptions(allRows);
      const rows = filteredAndSortedReviewRows(allRows);
      renderReviewBadges(allRows, rows);
      if (!rows.length) {
        els.countEntryBody.innerHTML = `<tr><td colspan="10" class="empty-cell">No rows match the current review filters.</td></tr>`;
        return;
      }
      els.countEntryBody.innerHTML = rows.map((row) => {
        const diffLabel = row.diff == null ? "-" : (row.diff > 0 ? `+${fmt(row.diff)}` : fmt(row.diff));
        const statusLabel = row.status === "null" ? "NULL" : row.status === "diff" ? "QTY DIFF" : "PASS";
        const diffClass = row.status === "null" ? "entry-null" : row.diff > 0 ? "entry-positive" : row.diff < 0 ? "entry-negative" : "entry-exact";
        return `<tr class="count-review-row count-review-${row.status}" data-count-review-code="${esc(row.code)}" title="Click row to scan/recount this item">
          <td>${esc(row.code || "-")}</td>
          <td>${esc(row.plu || "-")}</td>
          <td>${esc(row.product || "-")}</td>
          <td>${esc(row.vendor || "-")}</td>
          <td>${esc(row.category || "-")}</td>
          <td class="num">${fmt(row.before)}</td>
          <td class="num">${row.entered == null ? "-" : fmt(row.entered)}</td>
          <td class="num">${row.counted == null ? "-" : fmt(row.counted)}</td>
          <td class="num ${diffClass}">${diffLabel}</td>
          <td><span class="review-status-pill review-status-${row.status}">${statusLabel}</span></td>
        </tr>`;
      }).join("");
      els.countEntryBody.querySelectorAll("[data-count-review-code]").forEach((row) => {
        row.addEventListener("click", () => {
          const code = row.dataset.countReviewCode || "";
          try {
            if (typeof selectCountDropdownItem === "function") selectCountDropdownItem(code);
            else if (els.countSearchInput) { els.countSearchInput.value = code; handleCountLookup?.(); }
          } catch (_) {
            if (els.countSearchInput) { els.countSearchInput.value = code; els.countSearchInput.focus(); }
          }
        });
      });
    };

    function forceCountWorkspaceFront() {
      if (els.countSetupModal) {
        els.countSetupModal.hidden = true;
        els.countSetupModal.style.pointerEvents = "none";
      }
      [els.countReportModal, document.querySelector("#reportCountModal"), document.querySelector("#sessionHistoryModal")].forEach((modal) => {
        if (modal) { modal.hidden = true; modal.style.pointerEvents = "none"; }
      });
      if (els.countSessionModal) {
        els.countSessionModal.hidden = false;
        els.countSessionModal.style.zIndex = "20000";
        els.countSessionModal.style.pointerEvents = "auto";
        const dialog = els.countSessionModal.querySelector(".count-modal__dialog");
        if (dialog) dialog.style.pointerEvents = "auto";
      }
      state._countSessionOpen = !!state.activeCountSession;
      try { originalRenderWorkspace?.(); } catch (_) {}
      ensureCountReviewControls();
      try { renderCountEntryRows(false); } catch (_) {}
      if (typeof focusCountSearch === "function") setTimeout(focusCountSearch, 40);
    }

    function resetReviewFiltersForNewSession() {
      state._countReviewFilter = "all";
      state._countReviewVendor = "";
      state._countReviewCategory = "";
      state._countReviewSearch = "";
      state._countReviewSort = { key: "default", dir: "desc" };
    }

    startCountSessionFromModal = function startCountSessionFromModalRescue() {
      if (state._startingCountNow) return;
      state._startingCountNow = true;
      try {
        resetReviewFiltersForNewSession();
        originalStart?.();
        state._countSessionOpen = !!state.activeCountSession;
        forceCountWorkspaceFront();
      } finally {
        setTimeout(() => { state._startingCountNow = false; }, 800);
      }
    };

    continueCountFromReport = async function continueCountFromReportRescue(event = null) {
      event?.preventDefault?.();
      event?.stopPropagation?.();
      const id = cc(state.countReportOpenId || event?.target?.dataset?.continueSession || "");
      let session = id && typeof findCountSessionById === "function" ? findCountSessionById(id) : null;
      if (!session && id && typeof refreshLatestCountSessions === "function") {
        try { await Promise.race([refreshLatestCountSessions({ history: true }), new Promise((r) => setTimeout(r, 1500))]); } catch (_) {}
        session = findCountSessionById(id);
      }
      if (!session) { showToast?.("Count session not found yet. Reopen history and try again.", 2600, "warning"); return; }
      const live = typeof markCountSessionDirty === "function"
        ? markCountSessionDirty({ ...session, savedAt: "", submittedAt: "", isActiveLive: true })
        : { ...session, savedAt: "", submittedAt: "", isActiveLive: true };
      state.activeCountSession = live;
      state.countSessions = [live, ...(state.countSessions || []).filter((s) => cc(s?.id) !== cc(live.id))];
      state._continuingCountId = live.id;
      state._countSessionOpen = true;
      state.countQtyBuffer = "0";
      state.selectedCountItemCode = "";
      state.countStage = "search";
      state.pendingDuplicateCount = null;
      state.pendingDuplicateMode = null;
      try { persistActiveCountSession?.(); persistCountSessions?.({ scheduleSync: false }); } catch (_) {}
      forceCountWorkspaceFront();
      setTimeout(forceCountWorkspaceFront, 200);
      showToast?.(`Continuing count: ${typeof countSessionLabel === "function" ? countSessionLabel(live) : live.id}`, 2200, "success");
    };

    saveCountSession = async function saveCountSessionRescue() {
      if (!state.activeCountSession) return;
      try { await originalSave?.(); }
      finally {
        if (els.countSessionModal) els.countSessionModal.hidden = true;
        if (els.countSetupModal) els.countSetupModal.hidden = true;
        state._countSessionOpen = false;
        state._continuingCountId = "";
        try { originalRenderWorkspace?.(); } catch (_) {}
      }
    };

    deleteCountSession = function deleteCountSessionRescue() {
      try { originalDelete?.(); }
      finally {
        if (els.countSessionModal) els.countSessionModal.hidden = true;
        state._countSessionOpen = false;
        state._continuingCountId = "";
        try { originalRenderWorkspace?.(); } catch (_) {}
      }
    };

    renderCountsWorkspace = function renderCountsWorkspaceRescue(options = {}) {
      originalRenderWorkspace?.(options);
      const launch = els.countLaunchCard;
      if (launch) {
        launch.style.display = "";
        launch.hidden = false;
      }
      if (els.countLaunchTitle) els.countLaunchTitle.textContent = "Start New Count";
      if (els.countLaunchDescription) els.countLaunchDescription.textContent = "Open the setup wizard for a new physical count.";
      if (els.countLaunchState) els.countLaunchState.textContent = "New physical count";
      document.querySelectorAll(".count-launch-card").forEach((card) => {
        if (card.id && card.id !== "countLaunchCard" && card.id !== "openSessionHistoryButton") card.style.display = "none";
      });
      if (state.activeCountSession && state._countSessionOpen) {
        ensureCountReviewControls();
      }
    };

    function bindCapture(id, handler) {
      const node = document.querySelector(id);
      if (!node || node.dataset.rescueBound === "1") return;
      node.dataset.rescueBound = "1";
      node.addEventListener("click", (event) => {
        event.preventDefault();
        event.stopPropagation();
        event.stopImmediatePropagation();
        handler(event);
      }, true);
    }

    bindCapture("#countStartButton", () => startCountSessionFromModal());
    bindCapture("#countSaveSessionButton", () => { void saveCountSession(); });
    bindCapture("#countDeleteSessionButton", () => deleteCountSession());
    bindCapture("#countContinueButton", (e) => { void continueCountFromReport(e); });
    // Override old Review button path: same count screen, just switch list to needs-review filter.
    bindCapture("#countReviewButton", () => {
      state._countReviewFilter = "needs";
      forceCountWorkspaceFront();
    });

    document.addEventListener("keydown", (event) => {
      if (event.key !== "Escape") return;
      if (els.countDuplicateModal && !els.countDuplicateModal.hidden) return;
      if (els.countSessionModal && !els.countSessionModal.hidden) {
        event.preventDefault();
        event.stopPropagation();
        els.countSessionModal.hidden = true;
        state._countSessionOpen = false;
        try { originalRenderWorkspace?.(); } catch (_) {}
      }
    }, true);

    const style = document.createElement("style");
    style.textContent = `
      #countSetupModal[hidden], #countReportModal[hidden], #sessionHistoryModal[hidden], #reportCountModal[hidden] { pointer-events: none !important; }
      #countSessionModal { z-index: 20000 !important; }
      .count-review-filterbar { display:grid; gap:.55rem; margin:.5rem 0 .75rem; padding:.65rem; border:1px solid #dce3df; border-radius:10px; background:#fbfdfb; }
      .count-review-badges { display:flex; flex-wrap:wrap; gap:.4rem; align-items:center; }
      .count-review-badge { display:inline-flex; gap:.25rem; align-items:center; padding:.25rem .5rem; border-radius:999px; border:1px solid #dce3df; background:#fff; font-size:.78rem; font-weight:800; }
      .badge-needs { border-color:#e85f4c; color:#9b2418; background:#fff5f3; }
      .badge-null { border-color:#e85f4c; color:#9b2418; background:#fff0ee; }
      .badge-diff { border-color:#d79b25; color:#8a5a00; background:#fff7e8; }
      .badge-pass { border-color:#16835b; color:#116144; background:#eefaf4; }
      .count-review-controls { display:grid; grid-template-columns: minmax(9rem, .8fr) minmax(9rem, 1fr) minmax(9rem, 1fr) minmax(12rem, 1.4fr); gap:.55rem; align-items:end; }
      .count-review-controls label { font-size:.68rem; }
      .sortable-count-head { cursor:pointer; user-select:none; }
      .sortable-count-head:hover { background:#e7f4ed !important; }
      .count-review-row { cursor:pointer; }
      .count-review-null { background:#fff4f2 !important; }
      .count-review-diff { background:#fff9e8 !important; }
      .count-review-pass { background:#f0fbf5 !important; color:#40524b; }
      .review-status-pill { display:inline-block; border-radius:999px; padding:.16rem .45rem; font-size:.72rem; font-weight:900; white-space:nowrap; }
      .review-status-null { background:#e85f4c; color:#fff; }
      .review-status-diff { background:#d79b25; color:#1c2320; }
      .review-status-pass { background:#16835b; color:#fff; }
      #countEntryTable th, #countEntryTable td { white-space:nowrap; }
      #countEntryTable td:nth-child(3) { white-space:normal; min-width:18rem; }
      @media (max-width: 900px) { .count-review-controls { grid-template-columns: 1fr 1fr; } }
    `;
    document.head.appendChild(style);

    // Initial cleanup.
    setTimeout(() => {
      try { renderCountsWorkspace(); } catch (_) {}
    }, 100);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})();

// 2026-06-11 final session/review stabilization patch.
// Keeps the one-list review UI, but fixes launch/continue/modal focus issues without touching sync/imports.
(function finalCountSessionStabilizer() {
  const PATCH = "20260611-session-launch-review-stabilizer";
  function $id(id) { return document.getElementById(id); }
  function val(id) { return ($id(id)?.value || "").trim(); }
  function norm(v) { return String(v || "").trim(); }
  function safeCall(fn, ...args) { try { return typeof fn === "function" ? fn(...args) : undefined; } catch (e) { console.warn(`[${PATCH}]`, e); return undefined; } }
  function hardHide(modal) {
    if (!modal) return;
    modal.hidden = true;
    modal.style.pointerEvents = "none";
    modal.style.display = "";
  }
  function hardShowCountModal() {
    const modal = $id("countSessionModal");
    if (!modal) return false;
    modal.hidden = false;
    modal.style.pointerEvents = "auto";
    modal.style.display = "";
    modal.style.zIndex = "30000";
    modal.classList.add("count-session-forced-open");
    const dialog = modal.querySelector(".count-modal__dialog");
    if (dialog) {
      dialog.style.pointerEvents = "auto";
      dialog.style.zIndex = "30001";
    }
    return true;
  }
  function closeBlockingModals() {
    ["countSetupModal", "countReportModal", "sessionHistoryModal", "reportCountModal", "finalCountReportModal"].forEach((id) => hardHide($id(id)));
  }
  function focusScanNoScroll(delay = 40) {
    setTimeout(() => {
      if (Date.now() < Number(window.__countReviewNoFocusUntil || 0)) return;
      const input = $id("countSearchInput");
      if (!input || input.closest("[hidden]")) return;
      try { input.focus({ preventScroll: true }); }
      catch (_) { input.focus(); }
      try { input.select?.(); } catch (_) {}
    }, delay);
  }
  function makeSessionObject() {
    const statusEl = $id("countStatusInput");
    const id = safeCall(window.makeCountIdentifier, "count") || `count-${Date.now()}-${Math.random().toString(16).slice(2)}`;
    return {
      id,
      date: val("countDateInput") || new Date().toISOString().slice(0, 10),
      vendor: val("countVendorInput"),
      category: val("countCategoryInput"),
      status: statusEl ? (statusEl.value || "") : "",
      searchFilter: val("countScopeSearchInput"),
      startedAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
      deviceId: safeCall(window.countDeviceId) || localStorage.getItem("posDashboardCountDeviceId:v1") || "local-device",
      deviceLabel: safeCall(window.countDeviceLabel) || navigator.userAgent.slice(0, 48),
      user: safeCall(window.currentAuditUser) || "System",
      allowOutOfScope: !safeCall(window.isUserRole) && !!$id("countAllowOutOfScopeInput")?.checked,
      syncVersion: 0,
      localSyncPending: true,
      isActiveLive: true,
      entries: [],
    };
  }
  function installActiveSession(session) {
    if (!window.state) return false;
    const existing = Array.isArray(state.countSessions) ? state.countSessions : [];
    state.activeCountSession = session;
    state.countSessions = [session, ...existing.filter((s) => norm(s?.id) !== norm(session.id))];
    state._countSessionOpen = true;
    state._countHardOpenUntil = Date.now() + 4500;
    state._continuingCountId = session.id;
    state.countQtyBuffer = "0";
    state.selectedCountItemCode = "";
    state.countStage = "search";
    state.pendingDuplicateCount = null;
    state.pendingDuplicateMode = null;
    state.countAutoPlusMode = false;
    state.countAutoUndoStack = [];
    safeCall(window.persistActiveCountSession);
    safeCall(window.persistCountSessions, { scheduleSync: false });
    return true;
  }
  function openActiveCountNow(session, opts = {}) {
    if (!session || !window.state) return false;
    installActiveSession(session);
    closeBlockingModals();
    safeCall(window.renderCountsWorkspace, { suppressLoadingClose: true });
    hardShowCountModal();
    safeCall(window.renderSelectedCountItem);
    safeCall(window.renderCountQuantity);
    safeCall(window.renderCountEntryRows, false);
    if (opts.review) {
      state._countReviewFilter = state._countReviewFilter || "needs";
      safeCall(window.renderCountEntryRows, false);
    }
    focusScanNoScroll(80);
    return true;
  }
  function beginNewCountOnce() {
    if (!window.state || !$id("countSessionModal")) {
      safeCall(window.showToast, "Count screen is not ready yet. Refresh and try again.", 2800, "warning");
      return;
    }
    if (state._startingCountNow) return;
    state._startingCountNow = true;
    const btn = $id("countStartButton");
    if (btn) btn.disabled = true;
    try {
      const session = makeSessionObject();
      const opened = openActiveCountNow(session, { review: false });
      if (!opened) return;
      safeCall(window.showToast, `Started count: ${safeCall(window.countSessionLabel, session) || session.searchFilter || session.vendor || "New count"}`, 2400, "success");
      // Defer optional index build until after the UI is open.
      setTimeout(() => safeCall(window.buildCountSearchIndex), 120);
    } finally {
      setTimeout(() => {
        state._startingCountNow = false;
        if (btn) btn.disabled = false;
      }, 900);
    }
  }
  function findSessionForContinue(event) {
    const fromDataset = event?.target?.closest?.("[data-count-report],[data-session-id],[data-continue-session]");
    const id = norm(state?.countReportOpenId || fromDataset?.dataset?.countReport || fromDataset?.dataset?.sessionId || fromDataset?.dataset?.continueSession || "");
    if (id && typeof window.findCountSessionById === "function") return window.findCountSessionById(id);
    if (id) return (state.countSessions || []).find((s) => norm(s?.id) === id);
    return state?.activeCountSession || (state?.countSessions || [])[0] || null;
  }
  function continueSelectedSession(event, review = false) {
    const session = findSessionForContinue(event);
    if (!session) {
      safeCall(window.showToast, "Count session not found yet. Close and reopen Report History.", 2800, "warning");
      return;
    }
    const live = { ...session, savedAt: "", submittedAt: "", isActiveLive: true, updatedAt: new Date().toISOString() };
    if (review) state._countReviewFilter = "needs";
    openActiveCountNow(live, { review });
    safeCall(window.showToast, `Continuing count: ${safeCall(window.countSessionLabel, live) || live.id}`, 2200, "success");
  }

  // Replace the global functions too, so older handlers that still fire take the same path.
  window.startCountSessionFromModal = beginNewCountOnce;
  window.continueCountFromReport = (event) => { event?.preventDefault?.(); event?.stopPropagation?.(); continueSelectedSession(event, false); };
  window.openCountSessionDirect = (sessionId, opts = {}) => {
    const session = (typeof window.findCountSessionById === "function" ? window.findCountSessionById(sessionId) : null)
      || (state?.countSessions || []).find((s) => norm(s?.id) === norm(sessionId));
    return openActiveCountNow(session, opts);
  };

  // One authoritative capture handler. This runs before old target listeners.
  document.addEventListener("click", (event) => {
    const reviewControl = event.target.closest?.(".count-review-filterbar, .sortable-count-head, #countEntryTable th, #countEntryTable select, #countEntryTable input");
    if (reviewControl) window.__countReviewNoFocusUntil = Date.now() + 900;

    const tabCounts = event.target.closest?.('.tab-button[data-tab="counts"]');
    if (tabCounts && !state?._startingCountNow && !state?._countHardOpenUntil) {
      // Opening Inventory tab should show the workspace page, not auto-popup an old modal.
      state._countSessionOpen = false;
      hardHide($id("countSessionModal"));
      return;
    }

    const startBtn = event.target.closest?.("#countStartButton");
    if (startBtn) {
      event.preventDefault(); event.stopPropagation(); event.stopImmediatePropagation();
      beginNewCountOnce();
      return;
    }
    const launchCard = event.target.closest?.("#countLaunchCard");
    if (launchCard) {
      event.preventDefault(); event.stopPropagation(); event.stopImmediatePropagation();
      safeCall(window.openCountSetupModal);
      const m = $id("countSetupModal");
      if (m) { m.hidden = false; m.style.pointerEvents = "auto"; m.style.zIndex = "30010"; }
      return;
    }
    const continueBtn = event.target.closest?.("#countContinueButton");
    if (continueBtn) {
      event.preventDefault(); event.stopPropagation(); event.stopImmediatePropagation();
      continueSelectedSession(event, false);
      return;
    }
    const reviewBtn = event.target.closest?.("#countContinueReviewButton, #countReviewButton");
    if (reviewBtn) {
      event.preventDefault(); event.stopPropagation(); event.stopImmediatePropagation();
      continueSelectedSession(event, true);
      return;
    }
    const saveBtn = event.target.closest?.("#countSaveSessionButton");
    if (saveBtn) {
      event.preventDefault(); event.stopPropagation(); event.stopImmediatePropagation();
      Promise.resolve(safeCall(window.saveCountSession)).finally(() => {
        state._countSessionOpen = false;
        state._continuingCountId = "";
        hardHide($id("countSessionModal"));
        closeBlockingModals();
        safeCall(window.renderCountsWorkspace);
      });
      return;
    }
  }, true);

  document.addEventListener("change", (event) => {
    if (event.target.closest?.(".count-review-filterbar")) window.__countReviewNoFocusUntil = Date.now() + 900;
  }, true);
  document.addEventListener("keydown", (event) => {
    if (event.key !== "Escape") return;
    const setup = $id("countSetupModal");
    const count = $id("countSessionModal");
    if (setup && !setup.hidden) { event.preventDefault(); event.stopPropagation(); hardHide(setup); return; }
    if (count && !count.hidden) {
      event.preventDefault(); event.stopPropagation();
      state._countSessionOpen = false;
      state._continuingCountId = "";
      hardHide(count);
      safeCall(window.renderCountsWorkspace);
    }
  }, true);

  // Make focus safe: scanner search remains usable, but filters/headers do not yank scroll to top.
  const oldFocus = window.focusCountSearch;
  window.focusCountSearch = function focusCountSearchNoScroll() {
    if (Date.now() < Number(window.__countReviewNoFocusUntil || 0)) return;
    const active = document.activeElement;
    if (active?.closest?.(".count-review-filterbar, #countEntryTable")) return;
    const input = $id("countSearchInput");
    if (!input || input.closest("[hidden]")) return;
    try { input.focus({ preventScroll: true }); }
    catch (_) { oldFocus ? safeCall(oldFocus) : input.focus(); }
  };

  // Periodic cleanup only; do not force-open unless a user just started/continued.
  setInterval(() => {
    ["countSetupModal", "countReportModal", "sessionHistoryModal", "reportCountModal"].forEach((id) => {
      const m = $id(id);
      if (m && m.hidden) m.style.pointerEvents = "none";
    });
    if (state?.activeCountSession && state?._countSessionOpen && Date.now() < Number(state._countHardOpenUntil || 0)) {
      hardShowCountModal();
    }
  }, 600);

  console.info(`[${PATCH}] loaded`);
})();
