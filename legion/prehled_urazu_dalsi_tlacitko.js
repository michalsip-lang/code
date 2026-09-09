(function () {
    "use strict";

    var BUTTON_ID = "ti-dalsi-urazy-button";

    var CONFIG = {
        sourceWebUrl: "http://portal.samohyl.cz/homeZVK",
        sourceListServerRelativeUrl: "/homeZVK/Lists/seznam_prodejen",
        sourceLocationField: "lokace_prodejny",

        webUrl: "http://portal.samohyl.cz/bozp_po",
        listServerRelativeUrl: "/bozp_po/Lists/zamestnanci",
        itemDispFormUrlBase: "http://portal.samohyl.cz/bozp_po/Lists/zamestnanci/DispForm.aspx?ID=",
        launchButtonText: "Přehled zaměstnanců BOZP/PO",
        launchButtonLoadingText: "Načítám...",
        launchButtonContainerId: "ti-urazy-button-container",
        top: 5000,
        employeeCenterCandidates: ["stredisko_zamestnance", "stredisko_x005f_zamestnance", "StrediskoZamestnance"],
        endDateCandidates: ["datum_ukonceni", "datum_x005f_ukonceni", "DatumUkonceni"],
        titleCandidates: ["Title"],
        employeeNumberCandidates: ["osobni_cislo_zamestnance", "osobni_x005f_cislo_x005f_zamestnance", "OsobniCisloZamestnance"],
        firstNameCandidates: ["jmeno_zamestnance", "jmeno_x005f_zamestnance", "JmenoZamestnance"],
        lastNameCandidates: ["prijmeni_zamestance", "prijmeni_zamestnance", "prijmeni_x005f_zamestance", "prijmeni_x005f_zamestnance", "PrijmeniZamestance", "PrijmeniZamestnance"],
        bozpStatusCandidates: ["stav_skoleni_bozp", "stav_x005f_skoleni_x005f_bozp", "StavSkoleniBOZP"],
        bozpLastDateCandidates: ["vstupni_skoleni_bozp", "vstupni_x005f_skoleni_x005f_bozp", "VstupniSkoleniBOZP"],
        bozpLinkCandidates: ["odkaz_test_bozp", "odkaz_x005f_test_x005f_bozp", "OdkazTestBOZP"],
        poStatusCandidates: ["stav_skoleni_po", "stav_x005f_skoleni_x005f_po", "StavSkoleniPO"],
        poLastDateCandidates: ["opakovane_skoleni_po", "opakovane_x005f_skoleni_x005f_po", "OpakovaneSkoleniPO"],
        poLinkCandidates: ["odkaz_na_skoleni_po", "odkaz_x005f_na_x005f_skoleni_x005f_po", "OdkazNaSkoleniPO"],
        bozpCandidates: ["skoleni_bozp", "skoleni_x005f_bozp", "BOZP", "datum_skoleni_bozp", "platnost_bozp"],
        poCandidates: ["skoleni_po", "skoleni_x005f_po", "PO", "datum_skoleni_po", "platnost_po"]
    };

    function getQueryParam(name) {
        var match = RegExp("[?&]" + name + "=([^&]*)", "i").exec(window.location.search);
        return match ? decodeURIComponent(match[1]) : null;
    }

    function safeHtml(val) {
        if (val === null || val === undefined) return "";
        return String(val)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/\"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    function toScalarString(value) {
        if (value === null || value === undefined) return "";
        if (Array.isArray(value)) {
            return value.map(function (v) { return toScalarString(v); }).filter(Boolean).join(", ");
        }
        if (typeof value === "object") {
            if (Object.prototype.hasOwnProperty.call(value, "results") && Array.isArray(value.results)) {
                return value.results.map(function (v) { return toScalarString(v); }).filter(Boolean).join(", ");
            }
            if (Object.prototype.hasOwnProperty.call(value, "Url")) return String(value.Url || "");
            if (Object.prototype.hasOwnProperty.call(value, "url")) return String(value.url || "");
            if (Object.prototype.hasOwnProperty.call(value, "Value")) return String(value.Value);
            if (Object.prototype.hasOwnProperty.call(value, "Label")) return String(value.Label);
            if (Object.prototype.hasOwnProperty.call(value, "Title")) return String(value.Title);
            if (Object.prototype.hasOwnProperty.call(value, "Name")) return String(value.Name);
            if (Object.prototype.hasOwnProperty.call(value, "LookupValue")) return String(value.LookupValue);
            return "";
        }
        return String(value);
    }

    function cleanValue(value) {
        var raw = toScalarString(value).trim();
        if (!raw) return "";

        var parts = raw.split(/;#|[;,\t\n\r]+/);
        var out = [];
        var seen = {};
        for (var i = 0; i < parts.length; i++) {
            var p = (parts[i] || "").trim();
            if (!p) continue;
            if (/^\d+$/.test(p)) continue;
            var key = p.toLowerCase();
            if (!seen[key]) {
                seen[key] = true;
                out.push(p);
            }
        }
        return out.length ? out.join(", ") : raw;
    }

    function formatDateTimeCz(value) {
        var text = toScalarString(value).trim();
        if (!text) return "";

        var match = /^\/Date\((\d+)(?:[+-]\d+)?\)\/$/.exec(text);
        var date = match ? new Date(parseInt(match[1], 10)) : new Date(text);
        if (isNaN(date.getTime())) return cleanValue(text);

        var d = date.getDate();
        var m = date.getMonth() + 1;
        var y = date.getFullYear();
        var h = date.getHours();
        var min = date.getMinutes();
        var hh = h < 10 ? "0" + h : String(h);
        var mm = min < 10 ? "0" + min : String(min);
        return d + ". " + m + ". " + y + " " + hh + ":" + mm;
    }

    function getUrlFromValue(value) {
        if (value === null || value === undefined) return "";

        if (typeof value === "object") {
            if (Object.prototype.hasOwnProperty.call(value, "Url")) return String(value.Url || "");
            if (Object.prototype.hasOwnProperty.call(value, "url")) return String(value.url || "");
            if (Object.prototype.hasOwnProperty.call(value, "Value")) {
                var inner = String(value.Value || "");
                if (/^https?:\/\//i.test(inner) || inner.indexOf("/") === 0) return inner;
            }
        }

        var text = toScalarString(value).trim();
        if (!text) return "";

        if (/^https?:\/\//i.test(text) || text.indexOf("/") === 0) return text;

        var split = text.split(/\s*[,;]\s*/);
        for (var i = 0; i < split.length; i++) {
            var candidate = (split[i] || "").trim();
            if (/^https?:\/\//i.test(candidate) || candidate.indexOf("/") === 0) {
                return candidate;
            }
        }

        return "";
    }

    function parseSharePointDate(value) {
        var text = toScalarString(value).trim();
        if (!text) return null;

        var match = /^\/Date\((\d+)(?:[+-]\d+)?\)\/$/.exec(text);
        var date = match ? new Date(parseInt(match[1], 10)) : new Date(text);
        return isNaN(date.getTime()) ? null : date;
    }

    function isEndDateAllowed(value) {
        var raw = cleanValue(value);
        if (!raw) return true;

        var date = parseSharePointDate(value);
        if (!date) return true;

        var dateOnly = new Date(date.getFullYear(), date.getMonth(), date.getDate());
        var now = new Date();
        var todayOnly = new Date(now.getFullYear(), now.getMonth(), now.getDate());

        return dateOnly > todayOnly;
    }

    function normalizeCompareValue(value) {
        var raw = toScalarString(value).trim();
        if (!raw) return "";

        if (raw.indexOf(";#") >= 0 || raw.indexOf("#") >= 0) {
            var parts = raw.split(/;#|#/);
            var nonEmpty = [];
            for (var p = 0; p < parts.length; p++) {
                var part = (parts[p] || "").trim();
                if (part) nonEmpty.push(part);
            }

            for (var n = 0; n < nonEmpty.length; n++) {
                var numCandidate = nonEmpty[n].replace(/\s+/g, "").replace(",", ".");
                if (/^[-+]?\d+(\.\d+)?$/.test(numCandidate)) {
                    var parsed = Number(numCandidate);
                    if (!isNaN(parsed) && isFinite(parsed)) return String(parsed);
                }
            }

            if (nonEmpty.length) raw = nonEmpty[0];
        }

        var compact = raw.replace(/\s+/g, "");
        var numericLike = compact.replace(",", ".");
        if (/^[-+]?\d+(\.\d+)?$/.test(numericLike)) {
            var parsed2 = Number(numericLike);
            if (!isNaN(parsed2) && isFinite(parsed2)) return String(parsed2);
        }

        return compact.replace(/^0+/, "").toLowerCase();
    }

    function getCurrentListServerRelativeUrl() {
        if (window._spPageContextInfo && _spPageContextInfo.webServerRelativeUrl && _spPageContextInfo.listUrl) {
            var web = _spPageContextInfo.webServerRelativeUrl.replace(/\/$/, "");
            var listUrl = _spPageContextInfo.listUrl;
            if (listUrl.indexOf("/") !== 0) listUrl = "/" + listUrl;
            return web + listUrl;
        }

        var path = window.location.pathname || "";
        var lower = path.toLowerCase();
        var idx = lower.indexOf("/dispform.aspx");
        return idx > 0 ? path.substring(0, idx) : CONFIG.sourceListServerRelativeUrl;
    }

    function buildGetListItemUrl(webUrl, listServerRelativeUrl, itemId, selectFields) {
        return webUrl +
            "/_api/web/GetList(@list)/items(" + itemId + ")" +
            "?@list='" + encodeURIComponent(listServerRelativeUrl) + "'" +
            "&$select=" + selectFields.join(",");
    }

    function httpGetJson(url, success, fail) {
        var xhr = new XMLHttpRequest();
        xhr.open("GET", url, true);
        xhr.setRequestHeader("Accept", "application/json;odata=verbose");
        xhr.onreadystatechange = function () {
            if (xhr.readyState !== 4) return;
            if (xhr.status >= 200 && xhr.status < 300) {
                var parsed;
                try {
                    parsed = JSON.parse(xhr.responseText);
                } catch (e) {
                    fail("Neplatná JSON odpověď: " + e.message);
                    return;
                }

                try {
                    success(parsed);
                } catch (e2) {
                    fail("Chyba zpracování dat: " + e2.message);
                }
            } else {
                fail("HTTP chyba " + xhr.status + " pro URL: " + url + " | Odpověď: " + (xhr.responseText || ""));
            }
        };
        xhr.send();
    }

    function getFieldValue(item, candidates, fallbackWords) {
        if (!item || !candidates || !candidates.length) return "";
        fallbackWords = fallbackWords || [];

        var i;
        for (i = 0; i < candidates.length; i++) {
            var key = candidates[i];
            if (Object.prototype.hasOwnProperty.call(item, key)) {
                var val = cleanValue(item[key]);
                if (val || getUrlFromValue(item[key])) return item[key];
            }
        }

        var keys = Object.keys(item || {});
        for (i = 0; i < keys.length; i++) {
            var k = keys[i].toLowerCase();
            var all = true;
            for (var j = 0; j < fallbackWords.length; j++) {
                if (k.indexOf(fallbackWords[j]) === -1) {
                    all = false;
                    break;
                }
            }
            if (!all) continue;
            var value = cleanValue(item[keys[i]]);
            if (value || getUrlFromValue(item[keys[i]])) return item[keys[i]];
        }
        return "";
    }

    function showModal(items, locationValue) {
        var html = [];
        html.push("<div style='padding:8px;font-family:Segoe UI,Arial,sans-serif;width:100%;max-width:100%;box-sizing:border-box;'>");
        html.push("<div style='margin-bottom:8px;'><b>Středisko:</b> " + safeHtml(cleanValue(locationValue || "")) + "</div>");
        html.push("<div style='margin-bottom:8px;'><b>Nalezeno zaměstnanců:</b> " + items.length + "</div>");
        html.push("<div style='width:100%;max-width:100%;overflow:hidden;'>");
        html.push("<table style='width:100%;max-width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px;line-height:1.25;'>");
        html.push("<thead>");
        html.push("<tr>");
        html.push("<th style='border:1px solid #ddd;padding:4px;' colspan='3'></th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;text-align:center;' colspan='3'>Školení BOZP</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;text-align:center;' colspan='3'>Školení PO</th>");
        html.push("</tr>");
        html.push("<tr>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:12%;'>Odkaz</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:10%;'>Osobní číslo</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:16%;'>Jméno a příjmení</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:10%;'>Stav</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:14%;'>Datum posledního školení</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:12%;'>Odkaz na školení BOZP</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:10%;'>Stav</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:14%;'>Datum posledního školení</th>");
        html.push("<th style='border:1px solid #ddd;padding:4px;width:12%;'>Odkaz na školení PO</th>");
        html.push("</tr></thead><tbody>");

        for (var i = 0; i < items.length; i++) {
            var it = items[i];
            var titleValue = getFieldValue(it, CONFIG.titleCandidates, ["title"]) || it.Title || "Zobrazit";
            var personalNumber = getFieldValue(it, CONFIG.employeeNumberCandidates, ["osobni", "cislo"]);
            var firstName = getFieldValue(it, CONFIG.firstNameCandidates, ["jmeno", "zamest"]);
            var lastName = getFieldValue(it, CONFIG.lastNameCandidates, ["prijmeni", "zamest"]);
            var fullName = (cleanValue(firstName) + " " + cleanValue(lastName)).replace(/\s+/g, " ").trim();
            var itemUrl = CONFIG.itemDispFormUrlBase + encodeURIComponent(it.Id);
            var bozpStatus = getFieldValue(it, CONFIG.bozpStatusCandidates, ["stav", "bozp"]);
            var bozpLastDate = getFieldValue(it, CONFIG.bozpLastDateCandidates, ["datum", "posled", "bozp"]);
            var bozpLinkValue = getFieldValue(it, CONFIG.bozpLinkCandidates, ["odkaz", "bozp"]);
            var bozpLinkUrl = getUrlFromValue(bozpLinkValue);
            var bozpLastDateText = formatDateTimeCz(bozpLastDate) || cleanValue(bozpLastDate);
            var poStatus = getFieldValue(it, CONFIG.poStatusCandidates, ["stav", "po"]);
            var poLastDate = getFieldValue(it, CONFIG.poLastDateCandidates, ["datum", "posled", "po"]);
            var poLinkValue = getFieldValue(it, CONFIG.poLinkCandidates, ["odkaz", "po"]);
            var poLinkUrl = getUrlFromValue(poLinkValue);
            var poLastDateText = formatDateTimeCz(poLastDate) || cleanValue(poLastDate);

            html.push("<tr>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'><a href='" + safeHtml(itemUrl) + "' target='_blank' style='word-break:break-word;overflow-wrap:anywhere;'>" + safeHtml(cleanValue(titleValue)) + "</a></td>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'>" + safeHtml(cleanValue(personalNumber)) + "</td>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'>" + safeHtml(fullName) + "</td>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'>" + safeHtml(cleanValue(bozpStatus)) + "</td>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'>" + safeHtml(bozpLastDateText) + "</td>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'>" + (bozpLinkUrl ? "<a href='" + safeHtml(bozpLinkUrl) + "' target='_blank' style='word-break:break-word;overflow-wrap:anywhere;'>Otevřít školení</a>" : "") + "</td>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'>" + safeHtml(cleanValue(poStatus)) + "</td>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'>" + safeHtml(poLastDateText) + "</td>");
            html.push("<td style='border:1px solid #ddd;padding:4px;word-break:break-word;overflow-wrap:anywhere;'>" + (poLinkUrl ? "<a href='" + safeHtml(poLinkUrl) + "' target='_blank' style='word-break:break-word;overflow-wrap:anywhere;'>Otevřít školení</a>" : "") + "</td>");
            html.push("</tr>");
        }

        html.push("</tbody></table>");
        html.push("</div>");
        html.push("</div>");

        var container = document.createElement("div");
        container.innerHTML = html.join("");

        var viewportWidth = window.innerWidth || document.documentElement.clientWidth || 1280;
        var viewportHeight = window.innerHeight || document.documentElement.clientHeight || 800;
        var modalWidth = Math.max(760, Math.floor(viewportWidth * 0.98));
        var modalHeight = Math.max(560, Math.floor(viewportHeight * 0.92));

        if (window.SP && SP.UI && SP.UI.ModalDialog && typeof SP.UI.ModalDialog.showModalDialog === "function") {
            SP.UI.ModalDialog.showModalDialog({
                title: "Přehled zaměstnanců a školení BOZP/PO",
                html: container,
                allowMaximize: true,
                showClose: true,
                autoSize: false,
                width: modalWidth,
                height: modalHeight
            });
            return;
        }

        alert("Přehled zaměstnanců a školení BOZP/PO je připraven.");
    }

    window.runDalsiUrazyAkce = function () {
        var itemId = getQueryParam("ID");
        if (!itemId) {
            console.error("ID položky není v URL.");
            return;
        }

        var sourceUrl = buildGetListItemUrl(
            CONFIG.sourceWebUrl,
            getCurrentListServerRelativeUrl(),
            itemId,
            ["Id", CONFIG.sourceLocationField]
        );

        httpGetJson(sourceUrl, function (sourceData) {
            var sourceItem = sourceData && sourceData.d ? sourceData.d : null;
            if (!sourceItem) {
                console.error("Nelze načíst zdrojovou položku prodejny.");
                return;
            }

            var locationValue = sourceItem[CONFIG.sourceLocationField];
            var locationNorm = normalizeCompareValue(locationValue);
            if (!locationNorm) {
                console.error("Pole " + CONFIG.sourceLocationField + " je prázdné.");
                return;
            }

            var url = CONFIG.webUrl +
                "/_api/web/GetList(@list)/items" +
                "?@list='" + encodeURIComponent(CONFIG.listServerRelativeUrl) + "'" +
                "&$top=" + CONFIG.top;

            httpGetJson(url, function (data) {
                var items = (data && data.d && data.d.results) ? data.d.results : [];
                var filtered = [];

                for (var i = 0; i < items.length; i++) {
                    var centerValue = getFieldValue(items[i], CONFIG.employeeCenterCandidates, ["stredisko", "zamestnance"]);
                    var centerNorm = normalizeCompareValue(centerValue);
                    var endDateValue = getFieldValue(items[i], CONFIG.endDateCandidates, ["datum", "ukonceni"]);
                    var endDateAllowed = isEndDateAllowed(endDateValue);

                    if (centerNorm && centerNorm === locationNorm && endDateAllowed) {
                        filtered.push(items[i]);
                    }
                }

                showModal(filtered, locationValue);
            }, function (err2) {
                console.error("Chyba načtení přehledu zaměstnanců:", err2);
            });
        }, function (err1) {
            console.error("Chyba načtení položky prodejny:", err1);
        });
    };

    function createActionButton(text, loadingText) {
        var button = document.createElement("button");
        button.type = "button";
        button.id = BUTTON_ID;
        button.textContent = text;
        button.setAttribute("data-original-text", text);
        button.setAttribute("data-loading-text", loadingText || "Načítám...");
        button.style.padding = "8px 12px";
        button.style.cursor = "pointer";
        button.style.border = "1px solid #0078d4";
        button.style.background = "#0078d4";
        button.style.color = "#fff";
        button.style.borderRadius = "2px";
        return button;
    }

    function ensureLaunchButtonContainer() {
        var existing = document.getElementById(CONFIG.launchButtonContainerId);
        if (existing) return existing;

        var container = document.createElement("div");
        container.id = CONFIG.launchButtonContainerId;
        container.style.margin = "8px 0";
        container.style.display = "flex";
        container.style.gap = "8px";
        container.style.flexWrap = "wrap";

        var scriptNode = document.currentScript;
        if (scriptNode && scriptNode.parentNode) {
            scriptNode.parentNode.insertBefore(container, scriptNode.nextSibling);
        } else {
            document.body.appendChild(container);
        }

        return container;
    }

    function renderLaunchButton() {
        var path = (window.location.pathname || "").toLowerCase();
        if (path.indexOf("/dispform.aspx") === -1) return;

        if (document.getElementById(BUTTON_ID)) return;

        var container = ensureLaunchButtonContainer();
        if (!container) return;

        if (!container.style.display) container.style.display = "flex";
        if (!container.style.gap) container.style.gap = "8px";
        if (!container.style.flexWrap) container.style.flexWrap = "wrap";

        var button = createActionButton(CONFIG.launchButtonText, CONFIG.launchButtonLoadingText);
        button.addEventListener("click", function () {
            button.disabled = true;
            var original = button.getAttribute("data-original-text") || button.textContent;
            button.textContent = button.getAttribute("data-loading-text") || "Načítám...";

            try {
                window.runDalsiUrazyAkce();
            } finally {
                button.disabled = false;
                button.textContent = original;
            }
        });

        container.appendChild(button);
    }

    window.initDalsiUrazyButton = renderLaunchButton;
})();
