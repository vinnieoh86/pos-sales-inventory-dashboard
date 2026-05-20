(function () {
  "use strict";

  const debounce = (fn, ms) => {
    let timer = 0;
    return function debounced(...args) {
      window.clearTimeout(timer);
      timer = window.setTimeout(() => fn.apply(this, args), ms);
    };
  };

  function renderActiveTab() {
    if (typeof window.clearInventorySelection === "function") window.clearInventorySelection();
    if (typeof window.syncStickyHeights === "function") window.syncStickyHeights();
    if (typeof window.queueActiveTabRender === "function") {
      window.queueActiveTabRender();
      return;
    }
    if (document.body.dataset.activeTab === "inventory" && typeof window.renderInventory === "function") {
      window.renderInventory();
    }
    if (document.body.dataset.activeTab === "ordering" && typeof window.renderOrdering === "function") {
      window.renderOrdering();
    }
  }

  const renderSearch = debounce(renderActiveTab, 160);

  function installSearchRescue() {
    const search = document.querySelector("#searchInput");
    if (!search || search.dataset.rescueUsable === "1") return;
    search.dataset.rescueUsable = "1";
    search.removeAttribute("readonly");
    search.readOnly = false;
    search.type = "search";
    search.name = "posInventoryLookup";
    search.setAttribute("autocomplete", "off");
    search.setAttribute("data-lpignore", "true");
    search.setAttribute("data-1p-ignore", "true");
    search.setAttribute("data-form-type", "other");

    search.addEventListener("input", (event) => {
      event.stopImmediatePropagation();
      const upper = search.value.toUpperCase();
      if (search.value !== upper) search.value = upper;
      if (typeof window.saveActiveTabSearch === "function") window.saveActiveTabSearch();
      renderSearch();
    }, true);

    search.addEventListener("focus", () => window.setTimeout(() => search.select(), 0));
    search.addEventListener("click", () => search.select());
    search.addEventListener("keydown", (event) => {
      if (event.key === "Enter") {
        event.preventDefault();
        search.select();
        renderActiveTab();
      }
    });
  }

  function forceVisibleHeaders() {
    document.querySelectorAll("#inventory thead, .order-table thead").forEach((thead) => {
      thead.style.display = "table-header-group";
      thead.style.visibility = "visible";
    });
  }

  function bootRescue() {
    installSearchRescue();
    forceVisibleHeaders();
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", bootRescue);
  } else {
    bootRescue();
  }
  window.addEventListener("focus", bootRescue);
})();
