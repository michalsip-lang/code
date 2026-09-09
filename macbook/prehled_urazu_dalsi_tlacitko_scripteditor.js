(function () {
    "use strict";

    var SCRIPT_URL = "/homeZVK/SiteAssets/prehled_urazu_dalsi_tlacitko.js";
    var SCRIPT_ATTR = "data-ti-dalsi-button-loader";

    function runInitIfReady() {
        if (window.initDalsiUrazyButton && typeof window.initDalsiUrazyButton === "function") {
            window.initDalsiUrazyButton();
        }
    }

    function loadScriptOnce(onDone) {
        var existing = document.querySelector('script[' + SCRIPT_ATTR + '="1"]');
        if (existing) {
            if (existing.getAttribute("data-ti-loaded") === "1") {
                onDone();
                return;
            }
            existing.addEventListener("load", onDone);
            return;
        }

        var script = document.createElement("script");
        script.type = "text/javascript";
        script.src = SCRIPT_URL;
        script.setAttribute(SCRIPT_ATTR, "1");
        script.onload = function () {
            script.setAttribute("data-ti-loaded", "1");
            onDone();
        };
        script.onerror = function () {
            console.error("Nepodařilo se načíst skript:", SCRIPT_URL);
        };
        document.head.appendChild(script);
    }

    function init() {
        loadScriptOnce(runInitIfReady);
    }

    if (document.readyState === "loading") {
        document.addEventListener("DOMContentLoaded", init);
    } else {
        init();
    }
})();
