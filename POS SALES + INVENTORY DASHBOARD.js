(function () {
  "use strict";

  const VERSION = "20260603-physical-count-modal-auto-undo-fix";

  function makeSearchEditable() {
    const search = document.querySelector("#searchInput");
    if (!search) return;
    search.removeAttribute("readonly");
    search.readOnly = false;
    search.type = "search";
    search.name = "posInventoryLookup";
    search.autocomplete = "off";
    search.setAttribute("data-lpignore", "true");
    search.setAttribute("data-1p-ignore", "true");
    search.setAttribute("data-form-type", "other");
  }

  function loadScript(src) {
    return new Promise((resolve, reject) => {
      const script = document.createElement("script");
      script.src = src;
      script.async = false;
      script.onload = resolve;
      script.onerror = () => reject(new Error(`Failed to load ${src}`));
      document.body.appendChild(script);
    });
  }

  makeSearchEditable();
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", makeSearchEditable);
  }

  loadScript(`POS SALES + INVENTORY DASHBOARD.core.js?v=${VERSION}`)
    .then(() => {
      makeSearchEditable();
      return loadScript(`rescue-usable.js?v=${VERSION}`);
    })
    .then(makeSearchEditable)
    .catch((error) => {
      console.error(error);
      makeSearchEditable();
    });
})();
