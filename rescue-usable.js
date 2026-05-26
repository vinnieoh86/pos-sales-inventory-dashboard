(function () {
  "use strict";

  function forceVisibleHeaders() {
    document.querySelectorAll("#inventory thead, .order-table thead").forEach((thead) => {
      thead.style.display = "table-header-group";
      thead.style.visibility = "visible";
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", forceVisibleHeaders, { once: true });
  } else {
    forceVisibleHeaders();
  }
})();
