/* 20260612: Count session-flow rescue layer.
   Goal: one reliable count screen.
   - New Count creates ONE session and immediately opens the workspace.
   - Report Continue opens that existing session directly.
   - Review lives inside the same scanned-items list with filters/sort.
   - No import/sync architecture changes. */
(function () {
  "use strict";

  function boot() {
    if (typeof state === "undefined" || typeof els === "undefined") {
      setTimeout(boot, 60);
      return;
    }

    const originalRenderWorkspace = typeof renderCountsWorkspace === "function" ? renderCountsWorkspace : null;
    const originalSave = typeof saveCountSession === "function" ? saveCountSession : null;
    const originalDelete = typeof deleteCountSession === "function" ? deleteCountSession : null;
    const originalFocusCountSearch = typeof focusCountSearch === "function" ? focusCountSearch : null;

    state._countReviewFilter = state._countReviewFilter || "all";
    state._countReviewVendor = state._countReviewVendor || "";
    state._countReviewCategory = state._countReviewCategory || "";
    state._countReviewSearch = state._countReviewSearch || "";
    state._countReviewSort = state._countReviewSort || { key: "default", dir: "desc" };

    function cc(value) { return typeof cleanCell === "function" ? cleanCell(value) : String(value || "").trim(); }
    function ck(value) { return typeof codeKey === "function" ? codeKey(value) : cc(value).replace(/\D/g, ""); }
    function esc(value) { return typeof escapeHtml === "function" ? escapeHtml(value) : String(value ?? "").replace(/[&<>"']/g, (m) => ({"&":"&amp;","<":"&lt;",">":"&gt;","\"":"&quot;","'":"&#39;"}[m])); }
    function fmt(value) { return typeof number !== "undefined" && number?.format ? number.format(Number(value || 0)) : String(Number(value || 0)); }
    function toast(message, ms = 2600, type = "success") { try { showToast?.(message, ms, type); } catch (_) {} }

    function suppressAutoFocus(ms = 500) { state._countSuppressAutoFocusUntil = Date.now() + ms; }
    function shouldSuppressAutoFocus() {
      const active = document.activeElement;
      return Date.now() < Number(state._countSuppressAutoFocusUntil || 0)
        || !!active?.closest?.("#countReviewFilterBar, #countEntryTable");
    }
    if (originalFocusCountSearch) {
      focusCountSearch = function focusCountSearchRescue() {
        if (shouldSuppressAutoFocus()) return;
        return originalFocusCountSearch();
      };
    }

    function closeModal(modal) {
      if (!modal) return;
      modal.hidden = true;
      modal.classList.remove("count-modal-front", "count-session-forced-open");
      modal.style.pointerEvents = "none";
    }
    function openModal(modal) {
      if (!modal) return;
      modal.hidden = false;
      modal.style.pointerEvents = "auto";
      modal.style.zIndex = "30000";
      modal.classList.add("count-modal-front");
      const dialog = modal.querySelector(".count-modal__dialog");
      if (dialog) dialog.style.pointerEvents = "auto";
    }
    function closeBlockingModals() {
      closeModal(els.countSetupModal);
      closeModal(els.countReportModal);
      closeModal(document.querySelector("#sessionHistoryModal"));
      closeModal(document.querySelector("#reportCountModal"));
      closeModal(document.querySelector("#finalCountReportModal"));
    }

    function setActiveSession(session, { live = true } = {}) {
      if (!session?.id) return null;
      const normalized = typeof normalizeCountSession === "function" ? normalizeCountSession(session, session.localSyncPending ? "pending" : "synced") : session;
      const active = typeof markCountSessionDirty === "function"
        ? markCountSessionDirty({ ...normalized, isActiveLive: !!live, submittedAt: "" })
        : { ...normalized, isActiveLive: !!live, submittedAt: "" };
      state.activeCountSession = active;
      state.countSessions = [active, ...(state.countSessions || []).filter((s) => cc(s?.id) !== cc(active.id))];
      state._countSessionOpen = true;
      state.countQtyBuffer = "0";
      state.selectedCountItemCode = "";
      state.countStage = "search";
      state.pendingDuplicateCount = null;
      state.pendingDuplicateMode = null;
      try { persistActiveCountSession?.(); persistCountSessions?.({ scheduleSync: false }); } catch (_) {}
      return active;
    }

    function openActiveCountScreen({ focus = true } = {}) {
      if (!state.activeCountSession) return false;
      closeBlockingModals();
      state._countSessionOpen = true;
      try { originalRenderWorkspace?.({ populateSetup: false }); } catch (_) {}
      openModal(els.countSessionModal);
      ensureCountReviewControls();
      try { renderCountEntryRows(false); } catch (_) {}
      if (focus) setTimeout(() => { try { focusCountSearch?.(); } catch (_) {} }, 80);
      return true;
    }

    function resetReviewFiltersForNewSession() {
      state._countReviewFilter = "all";
      state._countReviewVendor = "";
      state._countReviewCategory = "";
      state._countReviewSearch = "";
      state._countReviewSort = { key: "default", dir: "desc" };
    }

    function makeSessionFromSetup() {
      const statusEl = document.querySelector("#countStatusInput");
      const id = typeof makeCountIdentifier === "function" ? makeCountIdentifier("count") : `count-${Date.now()}`;
      return {
        id,
        date: els.countDateInput?.value || new Date().toISOString().slice(0, 10),
        vendor: els.countVendorInput?.value || "",
        category: els.countCategoryInput?.value || "",
        status: statusEl ? (statusEl.value || "") : "",
        searchFilter: cc(els.countScopeSearchInput?.value || ""),
        startedAt: new Date().toISOString(),
        updatedAt: new Date().toISOString(),
        deviceId: typeof countDeviceId === "function" ? countDeviceId() : "local",
        deviceLabel: typeof countDeviceLabel === "function" ? countDeviceLabel() : "This device",
        user: typeof currentAuditUser === "function" ? currentAuditUser() : "",
        allowOutOfScope: !(typeof isUserRole === "function" && isUserRole()) && !!els.countAllowOutOfScopeInput?.checked,
        syncVersion: 0,
        localSyncPending: true,
        entries: [],
        isActiveLive: true,
      };
    }

    function startFreshCountFromSetup(event = null) {
      event?.preventDefault?.();
      event?.stopPropagation?.();
      event?.stopImmediatePropagation?.();
      if (state._hardStartingCount) return;
      if (!els.countSessionModal) { toast("Count screen is not available.", 3000, "warning"); return; }
      state._hardStartingCount = true;
      const button = els.countStartButton || document.querySelector("#countStartButton");
      if (button) button.disabled = true;
      try {
        resetReviewFiltersForNewSession();
        const session = makeSessionFromSetup();
        setActiveSession(session, { live: true });
        openActiveCountScreen({ focus: true });
        try { buildCountSearchIndex?.(); } catch (_) { setTimeout(() => { try { buildCountSearchIndex?.(); } catch (_) {} }, 80); }
        setTimeout(() => { try { syncSharedCountSessionsToSupabase?.(true); } catch (_) {} }, 900);
        toast(`Started count: ${typeof countSessionLabel === "function" ? countSessionLabel(session) : session.id}`, 2800, "success");
      } finally {
        setTimeout(() => { state._hardStartingCount = false; if (button) button.disabled = false; }, 900);
      }
    }

    function openNewCountSetup(event = null) {
      event?.preventDefault?.();
      event?.stopPropagation?.();
      event?.stopImmediatePropagation?.();
      closeBlockingModals();
      // This button must ALWAYS start a fresh count wizard. It should not resume old sessions.
      if (els.countDateInput) els.countDateInput.value = new Date().toISOString().slice(0, 10);
      try { populateCountSetupOptions?.(); } catch (_) {}
      if (els.countVendorInput) els.countVendorInput.value = "";
      if (els.countCategoryInput) els.countCategoryInput.value = "";
      const statusEl = document.querySelector("#countStatusInput");
      if (statusEl) statusEl.value = "";
      if (els.countScopeSearchInput) els.countScopeSearchInput.value = cc(els.searchInput?.value || "");
      if (els.countAllowOutOfScopeInput) els.countAllowOutOfScopeInput.checked = false;
      openModal(els.countSetupModal);
      setTimeout(() => els.countScopeSearchInput?.focus?.(), 60);
    }

    async function continueExistingCountFromReport(event = null) {
      event?.preventDefault?.();
      event?.stopPropagation?.();
      event?.stopImmediatePropagation?.();
      const reportModal = els.countReportModal;
      const id = cc(state.countReportOpenId || reportModal?.dataset?.sessionId || event?.target?.dataset?.continueSession || "");
      let session = id && typeof findCountSessionById === "function" ? findCountSessionById(id) : null;
      if (!session && id && typeof refreshLatestCountSessions === "function") {
        try { await Promise.race([refreshLatestCountSessions({ history: true }), new Promise((resolve) => setTimeout(resolve, 1800))]); } catch (_) {}
        try { session = findCountSessionById(id); } catch (_) {}
      }
      if (!session) { toast("Could not find that saved count yet. Reopen Report History and try again.", 3200, "warning"); return; }
      setActiveSession({ ...session, savedAt: "", submittedAt: "" }, { live: true });
      state._continuingCountId = session.id;
      openActiveCountScreen({ focus: true });
      toast(`Continuing count: ${typeof countSessionLabel === "function" ? countSessionLabel(session) : session.id}`, 2400, "success");
      setTimeout(() => { try { syncSharedCountSessionsToSupabase?.(true); } catch (_) {} }, 900);
    }

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
      entryMap.forEach((entry, key) => { if (!seen.has(key)) addRow(entry, entry); });
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
        if (heading) heading.insertAdjacentElement("afterend", toolbar); else section.prepend(toolbar);
        toolbar.querySelector("#countReviewStatusFilter")?.addEventListener("change", (e) => { suppressAutoFocus(); state._countReviewFilter = e.target.value || "all"; renderCountEntryRows(false); });
        toolbar.querySelector("#countReviewVendorFilter")?.addEventListener("change", (e) => { suppressAutoFocus(); state._countReviewVendor = e.target.value || ""; renderCountEntryRows(false); });
        toolbar.querySelector("#countReviewCategoryFilter")?.addEventListener("change", (e) => { suppressAutoFocus(); state._countReviewCategory = e.target.value || ""; renderCountEntryRows(false); });
        toolbar.querySelector("#countReviewSearchFilter")?.addEventListener("input", (e) => { suppressAutoFocus(); state._countReviewSearch = e.target.value || ""; renderCountEntryRows(false); });
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
          th.addEventListener("click", (event) => {
            event.preventDefault();
            suppressAutoFocus();
            const key = th.dataset.reviewSort;
            const cur = state._countReviewSort || { key: "default", dir: "desc" };
            state._countReviewSort = cur.key === key ? { key, dir: cur.dir === "asc" ? "desc" : "asc" } : { key, dir: key === "diff" ? "desc" : "asc" };
            renderCountEntryRows(false);
          });
        });
      }
      table.querySelector("colgroup")?.remove();
    }

    function updateReviewDropdownOptions(allRows) {
      const vendorSel = document.querySelector("#countReviewVendorFilter");
      const catSel = document.querySelector("#countReviewCategoryFilter");
      if (!vendorSel || !catSel) return;
      const fill = (select, values, selected) => {
        const html = [`<option value="">All</option>`, ...values.map((value) => `<option value="${esc(value)}">${esc(value)}</option>`)].join("");
        if (select.dataset.lastOptions !== html) { select.innerHTML = html; select.dataset.lastOptions = html; }
        select.value = values.includes(selected || "") ? selected : "";
      };
      fill(vendorSel, [...new Set(allRows.map((r) => r.vendor).filter(Boolean))].sort((a,b)=>a.localeCompare(b)), state._countReviewVendor || "");
      fill(catSel, [...new Set(allRows.map((r) => r.category).filter(Boolean))].sort((a,b)=>a.localeCompare(b)), state._countReviewCategory || "");
      const status = document.querySelector("#countReviewStatusFilter");
      if (status) status.value = state._countReviewFilter || "all";
      const search = document.querySelector("#countReviewSearchFilter");
      if (search && search.value !== (state._countReviewSearch || "")) search.value = state._countReviewSearch || "";
    }
    function filteredAndSortedReviewRows(rows) {
      const filter = state._countReviewFilter || "all";
      const vendor = String(state._countReviewVendor || "").toUpperCase();
      const category = String(state._countReviewCategory || "").toUpperCase();
      const needle = String(state._countReviewSearch || "").toLowerCase().trim();
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
      return out.sort((a, b) => {
        let av = a[sort.key], bv = b[sort.key];
        if (["before", "entered", "counted", "diff"].includes(sort.key)) return (Number(av == null ? -999999 : av) - Number(bv == null ? -999999 : bv)) * dir;
        return String(av || "").localeCompare(String(bv || "")) * dir;
      });
    }
    function renderReviewBadges(allRows, visibleRows) {
      const badges = document.querySelector("#countReviewBadges");
      if (!badges) return;
      const counts = reviewCounts(allRows);
      badges.innerHTML = `
        <span class="count-review-badge badge-needs">Needs Review: <b>${fmt(counts.null + counts.diff)}</b></span>
        <span class="count-review-badge badge-null">NULL: <b>${fmt(counts.null)}</b></span>
        <span class="count-review-badge badge-diff">Qty Diff: <b>${fmt(counts.diff)}</b></span>
        <span class="count-review-badge badge-pass">PASS: <b>${fmt(counts.pass)}</b></span>
        <span class="count-review-badge">Showing: <b>${fmt(visibleRows.length)}</b></span>`;
    }

    renderCountEntryRows = function renderCountEntryRowsRescue() {
      if (!els.countEntryBody) return;
      ensureCountReviewControls();
      const session = state.activeCountSession;
      if (!session) { els.countEntryBody.innerHTML = `<tr><td colspan="10" class="empty-cell">Start or continue a physical count first.</td></tr>`; return; }
      const allRows = buildReviewRows(session);
      updateReviewDropdownOptions(allRows);
      const rows = filteredAndSortedReviewRows(allRows);
      renderReviewBadges(allRows, rows);
      if (!rows.length) { els.countEntryBody.innerHTML = `<tr><td colspan="10" class="empty-cell">No rows match the current review filters.</td></tr>`; return; }
      els.countEntryBody.innerHTML = rows.map((row) => {
        const diffLabel = row.diff == null ? "-" : (row.diff > 0 ? `+${fmt(row.diff)}` : fmt(row.diff));
        const statusLabel = row.status === "null" ? "NULL" : row.status === "diff" ? "QTY DIFF" : "PASS";
        const diffClass = row.status === "null" ? "entry-null" : row.diff > 0 ? "entry-positive" : row.diff < 0 ? "entry-negative" : "entry-exact";
        return `<tr class="count-review-row count-review-${row.status}" data-count-review-code="${esc(row.code)}" title="Click row to scan/recount this item">
          <td>${esc(row.code || "-")}</td><td>${esc(row.plu || "-")}</td><td>${esc(row.product || "-")}</td><td>${esc(row.vendor || "-")}</td><td>${esc(row.category || "-")}</td>
          <td class="num">${fmt(row.before)}</td><td class="num">${row.entered == null ? "-" : fmt(row.entered)}</td><td class="num">${row.counted == null ? "-" : fmt(row.counted)}</td>
          <td class="num ${diffClass}">${diffLabel}</td><td><span class="review-status-pill review-status-${row.status}">${statusLabel}</span></td>
        </tr>`;
      }).join("");
      els.countEntryBody.querySelectorAll("[data-count-review-code]").forEach((row) => {
        row.addEventListener("click", () => {
          const code = row.dataset.countReviewCode || "";
          try { selectCountDropdownItem?.(code); }
          catch (_) { if (els.countSearchInput) { els.countSearchInput.value = code; try { handleCountLookup?.(); } catch (_) {} } }
        });
      });
    };

    renderCountsWorkspace = function renderCountsWorkspaceRescue(options = {}) {
      const wasOpen = !!state._countSessionOpen && !!state.activeCountSession;
      originalRenderWorkspace?.(options);
      // Inventory tab must not automatically pop an old count unless an explicit action opened it.
      if (!wasOpen && !state._explicitCountOpenNow && els.countSessionModal) els.countSessionModal.hidden = true;
      if (els.countLaunchCard) {
        els.countLaunchCard.hidden = false;
        els.countLaunchCard.style.display = "";
      }
      if (els.countLaunchTitle) els.countLaunchTitle.textContent = "Start New Count";
      if (els.countLaunchDescription) els.countLaunchDescription.textContent = "Open a fresh vendor/category/name setup wizard.";
      if (els.countLaunchState) els.countLaunchState.textContent = "New physical count";
      if (state.activeCountSession && state._countSessionOpen && !els.countSessionModal?.hidden) ensureCountReviewControls();
    };

    saveCountSession = async function saveCountSessionRescue(event = null) {
      event?.preventDefault?.(); event?.stopPropagation?.(); event?.stopImmediatePropagation?.();
      if (!state.activeCountSession) return;
      try { await originalSave?.(); }
      finally {
        state._countSessionOpen = false;
        state._continuingCountId = "";
        closeModal(els.countSessionModal);
        closeModal(els.countSetupModal);
        try { originalRenderWorkspace?.({ populateSetup: false }); } catch (_) {}
      }
    };
    deleteCountSession = function deleteCountSessionRescue(event = null) {
      event?.preventDefault?.(); event?.stopPropagation?.(); event?.stopImmediatePropagation?.();
      try { originalDelete?.(); }
      finally {
        state._countSessionOpen = false;
        state._continuingCountId = "";
        closeModal(els.countSessionModal);
        try { originalRenderWorkspace?.({ populateSetup: false }); } catch (_) {}
      }
    };

    function replaceAndBind(id, handler, propName) {
      const oldNode = document.querySelector(id);
      if (!oldNode) return null;
      const node = oldNode.cloneNode(true);
      oldNode.replaceWith(node);
      if (propName && els) els[propName] = node;
      node.addEventListener("click", handler, true);
      return node;
    }

    replaceAndBind("#countLaunchCard", openNewCountSetup, "countLaunchCard");
    replaceAndBind("#countStartButton", startFreshCountFromSetup, "countStartButton");
    replaceAndBind("#countSaveSessionButton", (e) => void saveCountSession(e), "countSaveSessionButton");
    replaceAndBind("#countDeleteSessionButton", (e) => deleteCountSession(e), "countDeleteSessionButton");
    replaceAndBind("#countContinueButton", (e) => void continueExistingCountFromReport(e), "countContinueButton");
    replaceAndBind("#countCancelButton", (e) => { e.preventDefault(); closeModal(els.countSetupModal); }, "countCancelButton");

    // Old Review button now uses the same list — no separate review mode screen.
    replaceAndBind("#countReviewButton", (e) => {
      e.preventDefault();
      state._countReviewFilter = "needs";
      openActiveCountScreen({ focus: false });
    }, "countReviewButton");

    document.addEventListener("click", (event) => {
      // If user opens Inventory tab, show the page only. Do not auto-open a stale count modal.
      const tab = event.target?.closest?.('.tab-button[data-tab="counts"]');
      if (tab) {
        state._explicitCountOpenNow = false;
        state._countSessionOpen = false;
        closeModal(els.countSessionModal);
        closeModal(els.countSetupModal);
      }
    }, true);

    document.addEventListener("keydown", (event) => {
      if (event.key !== "Escape") return;
      if (els.countDuplicateModal && !els.countDuplicateModal.hidden) return;
      if (els.countSetupModal && !els.countSetupModal.hidden) { event.preventDefault(); closeModal(els.countSetupModal); return; }
      if (els.countSessionModal && !els.countSessionModal.hidden) {
        event.preventDefault();
        state._countSessionOpen = false;
        closeModal(els.countSessionModal);
        try { originalRenderWorkspace?.({ populateSetup: false }); } catch (_) {}
      }
    }, true);

    const style = document.createElement("style");
    style.textContent = `
      #countSetupModal[hidden], #countReportModal[hidden], #sessionHistoryModal[hidden], #reportCountModal[hidden], #finalCountReportModal[hidden] { pointer-events:none !important; }
      #countSetupModal.count-modal-front, #countSessionModal.count-modal-front { z-index:30000 !important; pointer-events:auto !important; }
      #countSessionModal .count-modal__dialog--workspace { width:min(98vw, 116rem) !important; max-width:116rem !important; max-height:94vh !important; overflow:auto !important; }
      @media (min-width: 980px) { #countSessionModal .count-workspace__grid { display:grid !important; grid-template-columns:minmax(30rem, .95fr) minmax(26rem, .75fr) !important; gap:1rem !important; align-items:start !important; } }
      .count-review-filterbar { display:grid; gap:.55rem; margin:.5rem 0 .75rem; padding:.65rem; border:1px solid #dce3df; border-radius:10px; background:#fbfdfb; }
      .count-review-badges { display:flex; flex-wrap:wrap; gap:.4rem; align-items:center; }
      .count-review-badge { display:inline-flex; gap:.25rem; align-items:center; padding:.25rem .5rem; border-radius:999px; border:1px solid #dce3df; background:#fff; font-size:.78rem; font-weight:800; }
      .badge-needs { border-color:#e85f4c; color:#9b2418; background:#fff5f3; } .badge-null { border-color:#e85f4c; color:#9b2418; background:#fff0ee; } .badge-diff { border-color:#d79b25; color:#8a5a00; background:#fff7e8; } .badge-pass { border-color:#16835b; color:#116144; background:#eefaf4; }
      .count-review-controls { display:grid; grid-template-columns:minmax(9rem,.8fr) minmax(9rem,1fr) minmax(9rem,1fr) minmax(12rem,1.4fr); gap:.55rem; align-items:end; } .count-review-controls label { font-size:.68rem; }
      .sortable-count-head { cursor:pointer; user-select:none; } .sortable-count-head:hover { background:#e7f4ed !important; }
      .count-review-row { cursor:pointer; } .count-review-null { background:#fff4f2 !important; } .count-review-diff { background:#fff9e8 !important; } .count-review-pass { background:#f0fbf5 !important; color:#40524b; }
      .review-status-pill { display:inline-block; border-radius:999px; padding:.16rem .45rem; font-size:.72rem; font-weight:900; white-space:nowrap; }
      .review-status-null { background:#e85f4c; color:#fff; } .review-status-diff { background:#d79b25; color:#1c2320; } .review-status-pass { background:#16835b; color:#fff; }
      #countEntryTable th, #countEntryTable td { white-space:nowrap; } #countEntryTable td:nth-child(3) { white-space:normal; min-width:18rem; }
      @media (max-width:900px) { .count-review-controls { grid-template-columns:1fr 1fr; } }
    `;
    document.head.appendChild(style);

    // Initial dashboard cleanup: no automatic count popup on page/tab load.
    state._countSessionOpen = false;
    closeModal(els.countSessionModal);
    closeModal(els.countSetupModal);
    setTimeout(() => { try { renderCountsWorkspace({ populateSetup: false }); } catch (_) {} }, 120);
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})();
