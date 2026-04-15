(function() {

  function onReady(fn) {
    if (document.readyState !== "loading") fn();
    else document.addEventListener("DOMContentLoaded", fn);
  }

  var DEBUG_LOKACE_PRO_ZBOZI = true;

  function debugLog() {
    if (!DEBUG_LOKACE_PRO_ZBOZI || !window.console || !console.log) return;
    try {
      console.log.apply(console, arguments);
    } catch (e) {}
  }

  // ===================== HELPERS: field find/lock/set =====================

  function getTisaRowByInternalName(internalName) {
    return document.querySelector("tr.tisa-form-row.tisa-form-row" + internalName);
  }

  function getControlsByInternalName(internalName) {
    var result = [];

    // 1) TISA row
    var tisaRow = getTisaRowByInternalName(internalName);
    if (tisaRow) result = tisaRow.querySelectorAll("input, select, textarea");

    // 2) SP classic *_FieldName
    if (!result || result.length === 0) {
      var hidden = document.querySelector("input[id$='_FieldName'][value='" + internalName + "']");
      if (hidden) {
        var row = hidden.closest("tr");
        if (row) result = row.querySelectorAll("input, select, textarea");
      }
    }

    // 3) data-field
    if (!result || result.length === 0) {
      var cell = document.querySelector("td.ms-formbody[data-field='" + internalName + "']");
      if (cell) result = cell.querySelectorAll("input, select, textarea");
    }

    // 4) name/id contains internalName
    if (!result || result.length === 0) {
      result = document.querySelectorAll(
        "input[name*='" + internalName + "'],select[name*='" + internalName + "'],textarea[name*='" + internalName + "']," +
        "input[id*='" + internalName + "'],select[id*='" + internalName + "'],textarea[id*='" + internalName + "']"
      );
    }

    return Array.prototype.slice.call(result || []);
  }

  function lockField(internalName) {
    var controls = getControlsByInternalName(internalName);
    if (!controls || controls.length === 0) return;

    controls.forEach(function(ctrl) {
      if (ctrl.tagName === "INPUT" || ctrl.tagName === "TEXTAREA") ctrl.setAttribute("readonly", "readonly");
      ctrl.setAttribute("disabled", "disabled");
      ctrl.style.pointerEvents = "none";
      ctrl.style.backgroundColor = "#f3f2f1";
      ctrl.style.borderColor = "#e1dfdd";
    });
  }

  function setChoiceLikeField(internalName, value) {
    var controls = getControlsByInternalName(internalName);
    debugLog("[lokace_pro_zbozi] nalezene controly pro", internalName, controls);
    if (!controls || controls.length === 0) return false;

    var f = null;

    // Pro Choice pole preferujeme SELECT, aby se správně propsala UI i interní hodnota.
    for (var i = 0; i < controls.length; i++) {
      if (controls[i].tagName === "SELECT") {
        f = controls[i];
        break;
      }
    }

    // Fallback pro text/textarea pole.
    if (!f) {
      for (var k = 0; k < controls.length; k++) {
        var t = controls[k].tagName;
        if (t === "INPUT" || t === "TEXTAREA") {
          f = controls[k];
          break;
        }
      }
    }

    if (!f) return false;

    var matched = false;
    debugLog("[lokace_pro_zbozi] vybrany control:", f);

    if (f.tagName === "SELECT") {
      var availableOptions = [];
      var wanted = normalizeChoiceMatchValue(value);
      for (var j = 0; j < f.options.length; j++) {
        var rawText = trimSafe(f.options[j].text || "");
        var rawValue = trimSafe(f.options[j].value || "");
        var optText = normalizeChoiceMatchValue(rawText);
        var optValue = normalizeChoiceMatchValue(rawValue);
        availableOptions.push(rawText + " | value=" + rawValue + " | normalized=" + optText + "/" + optValue);
        if (optText === wanted || optValue === wanted) {
          f.selectedIndex = j;
          matched = true;
          break;
        }
      }

      debugLog("[lokace_pro_zbozi] hledana hodnota raw:", value, "normalized:", wanted);
      debugLog("[lokace_pro_zbozi] options v SELECTu:", availableOptions);
    } else {
      f.value = value;
      matched = true;
    }

    debugLog("[lokace_pro_zbozi] pokus o nastaveni hodnoty:", value, "uspech:", matched);
    if (!matched) return false;

    try {
      var ev = document.createEvent("HTMLEvents");
      ev.initEvent("change", true, false);
      f.dispatchEvent(ev);
    } catch (e) {}

    return true;
  }

  function fireInputEvents(el) {
    try {
      ["input", "change", "blur"].forEach(function(name) {
        var e = document.createEvent("HTMLEvents");
        e.initEvent(name, true, false);
        el.dispatchEvent(e);
      });
    } catch (e) {}
  }

  function removeDiacritics(v) {
    var s = String(v == null ? "" : v);
    try {
      return s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    } catch (e) {
      return s;
    }
  }

  function normalizeDecisionValue(v) {
    return trimSafe(removeDiacritics(v)).toLowerCase();
  }

  function toCanonicalDecisionValue(v) {
    var normalized = normalizeDecisionValue(v)
      .replace(/[\s\-]+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "");

    if (normalized.indexOf("nepotvrzen") !== -1) return "rozdil_nepotvrzen";

    if (normalized.indexOf("trva_na_rozdilu") !== -1) return "rozdil_potvrzen";

    if (normalized === "souhlasi" || normalized === "rozdil_potvrzen" || normalized === "potvrzen") {
      return "rozdil_potvrzen";
    }

    if (normalized === "nesouhlasi" || normalized === "rozdil_nepotvrzen" || normalized === "nepotvrzen") {
      return "rozdil_nepotvrzen";
    }

    return "";
  }

  function getDecisionOptionsByRole(role) {
    if (role === "prodejna") {
      return [
        { label: "Prodejna trvá na rozdílu", color: "#107c10" },
        { label: "Prodejna rozdíl nepotvrdila", color: "#a80000" }
      ];
    }

    return [
      { label: "Rozdíl potvrzen skladem", color: "#107c10" },
      { label: "Rozdíl skladem nepotvrzen", color: "#a80000" }
    ];
  }

  // ===================== ADVANCED TABLE (Treeinfo) =========================

  var TOC_FIELD = "rozdil_v_toc";
  var TOC_HTML_FIELD = "rozdil_v_toc_html"; // skryté pole pro HTML do e-mailu
  var TOC_COLS = 5;

  var COL_GOODS = 0;
  var COL_LOCATION = 1;
  var COL_DL = 2;
  var COL_FIZ = 3;
  var COL_DIFF = 4;
  var TASKFORM_COL_STAV_SKLAD = 5; // 6. sloupec v Edit TaskForm
  var TASKFORM_COL_STAV_PRODEJNA = 6; // 7. sloupec v Edit TaskForm

  var LOKACE_CHOICE_SINGLE_BRAMB = "BRAMB";
  var LOKACE_CHOICE_SINGLE_CENTR = "CENTR";
  var LOKACE_CHOICE_SINGLE_OSTATNI = "OSTATNI";
  var LOKACE_CHOICE_COMBINED = "BRAMR_CENTR";

  // Lookup do jiného SharePoint seznamu podle "Číslo zboží".
  var LOOKUP_LIST_URL = "http://portal.samohyl.cz/wms/Lists/detail_o_zbozi/";
  var LOOKUP_KEY_FIELD = "Title";
  var LOOKUP_INFO_FIELD = "Sloupec2";
  var LOOKUP_LOCATION_FIELD = "lokace_zbozi";
  var LOOKUP_SEPARATOR = " | ";
  var LOOKUP_BATCH_SIZE = 20;

  function getAdvancedTablePanel() {
    var row = getTisaRowByInternalName(TOC_FIELD);
    if (!row) return null;
    return row.querySelector(".tisa-AdvancedTable");
  }

  function getDataRows(panel) {
    if (!panel) return [];
    return Array.prototype.slice.call(panel.querySelectorAll("tr.dataRow"));
  }

  function clickAddRow(panel) {
    if (!panel) return false;
    var add = panel.querySelector("span.add");
    if (add) {
      add.click();
      return true;
    }
    var add2 = panel.querySelector("th.cAddRow, .cAddRow");
    if (add2) {
      add2.click();
      return true;
    }
    return false;
  }

  function clickRemoveRow(tr) {
    if (!tr) return;
    var rem = tr.querySelector("span.remove");
    if (rem) rem.click();
  }

  function normalizeCells(cells) {
    var arr = Array.isArray(cells) ? cells.slice(0) : [];
    while (arr.length < TOC_COLS) arr.push("");
    if (arr.length > TOC_COLS) arr = arr.slice(0, TOC_COLS);
    return arr;
  }

  function trimSafe(v) {
    return String(v == null ? "" : v).trim();
  }

  function normalizeChoiceMatchValue(v) {
    return trimSafe(v)
      .toUpperCase()
      .replace(/[;,/|]+/g, "_")
      .replace(/\s+/g, "_")
      .replace(/_+/g, "_")
      .replace(/^_+|_+$/g, "");
  }

  function stripLookupInfo(goodsCellValue) {
    var raw = trimSafe(goodsCellValue);
    if (!raw) return "";
    var idx = raw.indexOf(LOOKUP_SEPARATOR);
    return idx === -1 ? raw : raw.substring(0, idx).trim();
  }

  function normalizeGoodsKey(goodsValue) {
    return trimSafe(goodsValue).toUpperCase();
  }

  function encodeODataLiteral(s) {
    return String(s == null ? "" : s).replace(/'/g, "''");
  }

  function getLookupApiBaseUrl() {
    var a = document.createElement("a");
    a.href = LOOKUP_LIST_URL;

    var listPath = (a.pathname || "").replace(/\/$/, "");
    if (!listPath) return "";

    var webPath = listPath.replace(/\/Lists\/[^/]+$/i, "");
    if (!webPath) return "";

    var origin = a.protocol + "//" + a.host;
    var escapedListPath = listPath.replace(/'/g, "''");

    return origin + webPath + "/_api/web/GetList('" + escapedListPath + "')";
  }

  function sharePointGetJson(url, onSuccess, onError) {
    var xhr = new XMLHttpRequest();
    xhr.open("GET", url, true);
    xhr.setRequestHeader("Accept", "application/json;odata=verbose");
    xhr.onreadystatechange = function() {
      if (xhr.readyState !== 4) return;
      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          onSuccess(JSON.parse(xhr.responseText || "{}"));
        } catch (e) {
          if (onError) onError(e);
        }
      } else {
        if (onError) onError(new Error("HTTP " + xhr.status));
      }
    };
    xhr.send();
  }

  function buildLookupUrlForBatch(batchValues) {
    var apiBase = getLookupApiBaseUrl();
    if (!apiBase) return "";

    var select = "$select=" + LOOKUP_KEY_FIELD + "," + LOOKUP_INFO_FIELD + "," + LOOKUP_LOCATION_FIELD;

    var filters = [];
    for (var i = 0; i < batchValues.length; i++) {
      filters.push(LOOKUP_KEY_FIELD + " eq '" + encodeODataLiteral(batchValues[i]) + "'");
    }
    var filter = "$filter=" + encodeURIComponent(filters.join(" or "));
    var top = "$top=" + Math.max(batchValues.length, 1);

    return apiBase + "/items?" + select + "&" + filter + "&" + top;
  }

  function fetchLookupMapByGoodsNumbers(goodsNumbers, done) {
    var unique = [];
    var seen = {};

    for (var i = 0; i < goodsNumbers.length; i++) {
      var key = normalizeGoodsKey(stripLookupInfo(goodsNumbers[i]));
      if (!key || seen[key]) continue;
      seen[key] = true;
      unique.push(key);
    }

    if (unique.length === 0) {
      done({});
      return;
    }

    var map = {};
    var index = 0;

    function nextBatch() {
      if (index >= unique.length) {
        done(map);
        return;
      }

      var batch = unique.slice(index, index + LOOKUP_BATCH_SIZE);
      index += LOOKUP_BATCH_SIZE;

      var url = buildLookupUrlForBatch(batch);
      if (!url) {
        done(map);
        return;
      }

      sharePointGetJson(
        url,
        function(data) {
          var results = data && data.d && data.d.results ? data.d.results : [];
          for (var r = 0; r < results.length; r++) {
            var row = results[r] || {};
            var k = normalizeGoodsKey(row[LOOKUP_KEY_FIELD]);
            if (!k || map[k]) continue;

            map[k] = {
              info: trimSafe(row[LOOKUP_INFO_FIELD]),
              location: trimSafe(row[LOOKUP_LOCATION_FIELD])
            };
          }
          nextBatch();
        },
        function() {
          // Nezablokuje uložení formuláře při selhání lookupu.
          nextBatch();
        }
      );
    }

    nextBatch();
  }

  function enrichRowsFromLookup(rows, done) {
    var arr = Array.isArray(rows) ? rows : [];
    var goodsNumbers = [];

    for (var i = 0; i < arr.length; i++) {
      var cells = normalizeCells(arr[i] && arr[i].Cells ? arr[i].Cells : []);
      goodsNumbers.push(cells[0]);
    }

    fetchLookupMapByGoodsNumbers(goodsNumbers, function(lookupMap) {
      var enriched = [];

      for (var j = 0; j < arr.length; j++) {
        var inCells = normalizeCells(arr[j] && arr[j].Cells ? arr[j].Cells : []);
        var baseGoods = stripLookupInfo(inCells[COL_GOODS]);
        var lookupItem = lookupMap[normalizeGoodsKey(baseGoods)] || {};
        var lookupInfo = trimSafe(lookupItem.info);
        var lookupLocation = trimSafe(lookupItem.location);

        if (baseGoods && lookupInfo) inCells[COL_GOODS] = baseGoods + LOOKUP_SEPARATOR + lookupInfo;
        else inCells[COL_GOODS] = baseGoods;

        inCells[COL_LOCATION] = lookupLocation || trimSafe(inCells[COL_LOCATION]);

        enriched.push({
          Cells: inCells
        });
      }

      done(enriched);
    });
  }

  // ===================== Lokace pro zboží – aggregate ==================

  function getBaseLocationKey(rawLocation) {
    var loc = trimSafe(rawLocation).toUpperCase();
    if (!loc) return "";

    if (loc.indexOf("CENTR") !== -1) return "CENTR";
    if (loc.indexOf("BRAMB") !== -1) return "BRAMB";
    if (loc.indexOf("OSTATNI") !== -1) return "OSTATNI";
    return "";
  }

  function computeLokaceProZbozi(enrichedRows) {
    var orderedUnique = [];
    var seen = {};
    var rawLocations = [];

    for (var i = 0; i < enrichedRows.length; i++) {
      var cells = normalizeCells(enrichedRows[i] && enrichedRows[i].Cells ? enrichedRows[i].Cells : []);
      rawLocations.push(cells[COL_LOCATION]);
      var key = getBaseLocationKey(cells[COL_LOCATION]);
      if (!key || seen[key]) continue;

      seen[key] = true;
      orderedUnique.push(key);
    }

    if (orderedUnique.length === 0) return "";

    // Do Choice pole se zapisují skutečné hodnoty nakonfigurované v SharePointu.
    debugLog("[lokace_pro_zbozi] raw lokace:", rawLocations);
    debugLog("[lokace_pro_zbozi] unikatni lokace:", orderedUnique);

    if (orderedUnique.length === 1) {
      if (orderedUnique[0] === "BRAMB") {
        debugLog("[lokace_pro_zbozi] vysledna hodnota:", LOKACE_CHOICE_SINGLE_BRAMB);
        return LOKACE_CHOICE_SINGLE_BRAMB;
      }
      if (orderedUnique[0] === "CENTR") {
        debugLog("[lokace_pro_zbozi] vysledna hodnota:", LOKACE_CHOICE_SINGLE_CENTR);
        return LOKACE_CHOICE_SINGLE_CENTR;
      }
      if (orderedUnique[0] === "OSTATNI") {
        debugLog("[lokace_pro_zbozi] vysledna hodnota:", LOKACE_CHOICE_SINGLE_OSTATNI);
        return LOKACE_CHOICE_SINGLE_OSTATNI;
      }
    }

    debugLog("[lokace_pro_zbozi] vysledna hodnota:", LOKACE_CHOICE_COMBINED);
    return LOKACE_CHOICE_COMBINED;
  }

  function applyRowsToTisaAdvancedTable(rows) {
    var panel = getAdvancedTablePanel();
    if (!panel) return false;

    var desired = rows || [];

    // add rows
    var safety = 0;
    while (getDataRows(panel).length < desired.length && safety < 50) {
      if (!clickAddRow(panel)) break;
      safety++;
    }

    var trs = getDataRows(panel);

    // remove extra
    for (var d = trs.length - 1; d >= desired.length; d--) {
      clickRemoveRow(trs[d]);
    }

    trs = getDataRows(panel);

    // fill UI inputs
    for (var i = 0; i < desired.length; i++) {
      var cells = normalizeCells(desired[i].Cells);
      var tr = trs[i];
      if (!tr) continue;

      var inputs = tr.querySelectorAll("input.textInput");
      if (!inputs || inputs.length < TOC_COLS) continue;

      for (var c = 0; c < TOC_COLS; c++) {
        inputs[c].value = String(cells[c] == null ? "" : cells[c]);
        fireInputEvents(inputs[c]);
      }
    }

    return true;
  }

  function getTisaControlHiddenValue(internalName) {
    var row = getTisaRowByInternalName(internalName);
    if (!row) return "[]";
    var h = row.querySelector("input[type='hidden'][id^='tisa_controlvalue_']");
    return h ? (h.value || "[]") : "[]";
  }

  function parseExistingRowsFromHidden() {
    var raw = getTisaControlHiddenValue(TOC_FIELD);
    try {
      var parsed = JSON.parse(raw || "[]");
      if (!Array.isArray(parsed)) return [];
      return parsed.filter(function(r) {
        return r && Array.isArray(r.Cells);
      });
    } catch (e) {
      return [];
    }
  }

  // ===================== HTML generator for email ==========================

  function getAdvancedTableHtmlHeaders(columnCount) {
    var baseHeaders = [
      "Číslo zboží",
      "Lokace",
      "Množství na DL",
      "Fyzické dodané množství",
      "Rozdíl",
      "Vyjádření sklad",
      "Vyjádření prodejna"
    ];

    var count = Math.max(TOC_COLS, columnCount || 0);
    var headers = [];

    for (var i = 0; i < count; i++) {
      headers.push(baseHeaders[i] || ("Sloupec " + (i + 1)));
    }

    return headers;
  }

  function getPaddedCellsForHtml(cells, desiredCount) {
    var arr = Array.isArray(cells) ? cells.slice(0, desiredCount) : [];
    while (arr.length < desiredCount) arr.push("");
    return arr;
  }

  function advancedTableRowsToEmailHtml(rows) {
    var arr = Array.isArray(rows) ? rows : [];
    var colCount = TOC_COLS;

    for (var cIdx = 0; cIdx < arr.length; cIdx++) {
      var cLen = arr[cIdx] && Array.isArray(arr[cIdx].Cells) ? arr[cIdx].Cells.length : 0;
      if (cLen > colCount) colCount = cLen;
    }

    var headers = getAdvancedTableHtmlHeaders(colCount);

    function esc(s) {
      return String(s == null ? "" : s)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    var html = "";
    html += "<table cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse;width:100%;font-family:Segoe UI, Arial, sans-serif;font-size:13px;'>";
    html += "<thead><tr>";
    for (var h = 0; h < headers.length; h++) {
      html += "<th style='border:1px solid #d0d0d0;background:#f3f2f1;text-align:left;padding:8px;'>" + esc(headers[h]) + "</th>";
    }
    html += "</tr></thead><tbody>";

    for (var i = 0; i < arr.length; i++) {
      var cells = getPaddedCellsForHtml(arr[i] && arr[i].Cells, colCount);
      html += "<tr>";
      for (var c = 0; c < colCount; c++) {
        html += "<td style='border:1px solid #d0d0d0;padding:8px;'>" + esc(cells[c] || "") + "</td>";
      }
      html += "</tr>";
    }

    html += "</tbody></table>";
    return html;
  }

  function setFieldHtml(internalName, html) {
    var row = getTisaRowByInternalName(internalName);
    if (!row) {
      var hidden = document.querySelector("input[id$='_FieldName'][value='" + internalName + "']");
      row = hidden ? hidden.closest("tr") : null;
    }
    if (!row) return false;

    var ta = row.querySelector("textarea");
    if (ta) {
      ta.value = html;
      fireInputEvents(ta);
      return true;
    }

    var inp = row.querySelector("input[type='text'],input[type='hidden']");
    if (inp) {
      inp.value = html;
      fireInputEvents(inp);
      return true;
    }

    return false;
  }

  function collectAdvancedTableRowsFromPanel(panel) {
    var p = panel || getAdvancedTablePanel();
    if (!p) return [];

    var rows = [];
    var trs = getDataRows(p);

    for (var r = 0; r < trs.length; r++) {
      var controls = getEditableAdvancedControlsInRow(trs[r]);
      if (!controls || controls.length === 0) continue;

      var cells = [];
      var hasValue = false;

      for (var i = 0; i < controls.length; i++) {
        var raw = typeof controls[i].value === "string"
          ? controls[i].value
          : (controls[i].textContent || "");
        var value = trimSafe(raw);
        cells.push(value);
        if (value) hasValue = true;
      }

      if (!hasValue) continue;

      rows.push({
        Cells: cells
      });
    }

    return rows;
  }

  function syncTocHtmlFromAdvancedTable(panel) {
    var rows = collectAdvancedTableRowsFromPanel(panel);
    var htmlTable = advancedTableRowsToEmailHtml(rows);
    return setFieldHtml(TOC_HTML_FIELD, htmlTable);
  }

  function getWritableFieldControl(internalName) {
    var controls = getControlsByInternalName(internalName);
    if (!controls || controls.length === 0) return null;

    for (var i = 0; i < controls.length; i++) {
      if (controls[i].tagName === "TEXTAREA") return controls[i];
    }

    for (var j = 0; j < controls.length; j++) {
      var tag = controls[j].tagName;
      var type = (controls[j].type || "").toLowerCase();
      if (tag === "INPUT" && type !== "hidden" && type !== "button" && type !== "submit") return controls[j];
    }

    for (var k = 0; k < controls.length; k++) {
      if (controls[k].tagName === "INPUT") return controls[k];
    }

    return controls[0] || null;
  }

  function setFieldValueSmart(internalName, value) {
    if (setChoiceLikeField(internalName, value)) return true;

    var ctrl = getWritableFieldControl(internalName);
    if (!ctrl) return false;

    ctrl.value = value;
    fireInputEvents(ctrl);
    return true;
  }

  function updateStatusByLocationDecision(locationText, decisionValue) {
    var key = getBaseLocationKey(locationText);
    var isConfirmed = decisionValue === "rozdil_potvrzen";
    var statusValue = isConfirmed ? "Sklad souhlasí" : "Sklad nesouhlasí";

    if (key === "BRAMB") {
      return setFieldValueSmart("stav_bramborarna", statusValue);
    }

    if (key === "CENTR") {
      return setFieldValueSmart("stav_centrala", statusValue);
    }

    return false;
  }

  function appendKomentRozdilNote(itemLabel, columnLabel, decisionLabel, actualState, noteText) {
    var ctrl = getWritableFieldControl("koment_rozdil");
    if (!ctrl) return false;

    var current = trimSafe(ctrl.value || "");
    var parts = [
      itemLabel + " – " + columnLabel + " (" + decisionLabel + ")"
    ];

    var decisionValue = toCanonicalDecisionValue(decisionLabel);
    var actualStateText = trimSafe(actualState);
    if (decisionValue !== "rozdil_potvrzen" && actualStateText) {
      parts.push("Skutečný napočítaný stav: " + actualStateText);
    }

    var note = trimSafe(noteText);
    if (note) parts.push("Komentář: " + note);

    var block = parts.join("\r\n");

    ctrl.value = current ? current + "\r\n\r\n" + block : block;
    fireInputEvents(ctrl);
    return true;
  }

  function getEditableAdvancedControlsInRow(tr) {
    if (!tr) return [];
    var nodes = tr.querySelectorAll("select, textarea, input");
    return Array.prototype.slice.call(nodes || []).filter(function(el) {
      var type = (el.type || "").toLowerCase();
      if (type === "hidden" || type === "button" || type === "submit") return false;
      // Exclude original TISA controls that are hidden behind our display picker
      if (el.dataset && el.dataset.decisionPickerHidden === "1") return false;
      return true;
    });
  }

  function normalizeDecisionDisplayValue(rawValue, role) {
    var value = trimSafe(rawValue);
    var canonical = toCanonicalDecisionValue(value);
    if (canonical === "rozdil_potvrzen") {
      return role === "prodejna" ? "Prodejna trvá na rozdílu" : "Rozdíl potvrzen skladem";
    }
    if (canonical === "rozdil_nepotvrzen") {
      return role === "prodejna" ? "Prodejna rozdíl nepotvrdila" : "Rozdíl skladem nepotvrzen";
    }
    return value;
  }

  function isTrackedDecisionValue(v) {
    return v === "rozdil_potvrzen" || v === "rozdil_nepotvrzen";
  }

  function decisionValueToLabel(v, role) {
    if (v === "rozdil_potvrzen") {
      return role === "prodejna" ? "Prodejna trvá na rozdílu" : "Rozdíl potvrzen skladem";
    }
    if (v === "rozdil_nepotvrzen") {
      return role === "prodejna" ? "Prodejna rozdíl nepotvrdila" : "Rozdíl skladem nepotvrzen";
    }
    return "";
  }

  function setDecisionControlValue(control, label, fireEventsToo) {
    if (!control) return;

    var role = control && control.dataset ? (control.dataset.decisionRole || "") : "";
    var displayValue = normalizeDecisionDisplayValue(label, role);
    control.value = displayValue;
    control.defaultValue = displayValue;
    control.setAttribute("value", displayValue);
    control.dataset.lastDecisionValue = displayValue;

    var hiddenSrc = control.__decisionHiddenSource;
    if (hiddenSrc) {
      if (hiddenSrc.tagName === "SELECT") {
        var targetDecision = toCanonicalDecisionValue(displayValue);
        for (var k = 0; k < hiddenSrc.options.length; k++) {
          if (toCanonicalDecisionValue(hiddenSrc.options[k].text) === targetDecision ||
              toCanonicalDecisionValue(hiddenSrc.options[k].value) === targetDecision) {
            hiddenSrc.selectedIndex = k;
            break;
          }
        }
      } else {
        hiddenSrc.value = displayValue;
        hiddenSrc.defaultValue = displayValue;
        hiddenSrc.setAttribute("value", displayValue);
      }

      if (fireEventsToo) fireInputEvents(hiddenSrc);
    }

    if (fireEventsToo) fireInputEvents(control);
    setDecisionVisualState(control);
  }

  function closeDecisionPickerMenu() {
    if (!window.__taskFormDecisionPickerMenu) return;
    if (window.__taskFormDecisionPickerMenu.parentNode) {
      window.__taskFormDecisionPickerMenu.parentNode.removeChild(window.__taskFormDecisionPickerMenu);
    }
    window.__taskFormDecisionPickerMenu = null;
  }

  function ensureDecisionPickerGlobalCloseBinding() {
    if (window.__taskFormDecisionPickerGlobalBound) return;
    window.__taskFormDecisionPickerGlobalBound = true;

    document.addEventListener("mousedown", function(e) {
      var menu = window.__taskFormDecisionPickerMenu;
      if (!menu) return;

      var anchor = menu.__anchorControl;
      var t = e.target;
      if (menu.contains(t)) return;
      if (anchor && anchor.contains && anchor.contains(t)) return;

      closeDecisionPickerMenu();
    }, true);

    window.addEventListener("resize", closeDecisionPickerMenu);
    window.addEventListener("scroll", closeDecisionPickerMenu, true);
  }

  function openDecisionPickerMenu(anchorControl) {
    if (!anchorControl || anchorControl.disabled) return;

    ensureDecisionPickerGlobalCloseBinding();
    closeDecisionPickerMenu();

    var menu = document.createElement("div");
    menu.style.cssText =
      "position:absolute;min-width:170px;background:#fff;border:1px solid #c8c6c4;border-radius:4px;" +
      "box-shadow:0 4px 12px rgba(0,0,0,0.2);z-index:10020;padding:4px;";

    function addOption(label, color) {
      var btn = document.createElement("button");
      btn.type = "button";
      btn.textContent = label;
      btn.style.cssText =
        "display:block;width:100%;text-align:left;padding:7px 10px;border:none;background:#fff;" +
        "cursor:pointer;border-radius:3px;font-size:13px;color:" + color + ";";

      btn.onmouseenter = function() { btn.style.background = "#f3f2f1"; };
      btn.onmouseleave = function() { btn.style.background = "#fff"; };
      btn.onclick = function() {
        setDecisionControlValue(anchorControl, label, true);
        closeDecisionPickerMenu();
      };

      menu.appendChild(btn);
    }

    var options = getDecisionOptionsByRole(anchorControl.dataset && anchorControl.dataset.decisionRole);
    for (var o = 0; o < options.length; o++) {
      addOption(options[o].label, options[o].color);
    }

    document.body.appendChild(menu);

    var rect = anchorControl.getBoundingClientRect();
    menu.style.left = (rect.left + window.pageXOffset) + "px";
    menu.style.top = (rect.bottom + window.pageYOffset + 2) + "px";
    menu.__anchorControl = anchorControl;
    window.__taskFormDecisionPickerMenu = menu;
  }

  function bindDecisionPickerControl(control) {
    if (!control || control.__decisionPickerBound) return;
    control.__decisionPickerBound = true;

    control.readOnly = true;
    control.style.cursor = control.disabled ? "default" : "pointer";
    control.setAttribute("autocomplete", "off");
    control.setAttribute("placeholder", "Klikněte pro výběr");

    var initialValue = trimSafe(control.value || control.getAttribute("value") || "");
    if (initialValue) {
      control.defaultValue = initialValue;
      control.dataset.lastDecisionValue = initialValue;
      control.setAttribute("value", initialValue);
      var initialDecision = toCanonicalDecisionValue(initialValue);
      if (isTrackedDecisionValue(initialDecision)) control.dataset.lastCommittedDecision = initialDecision;
    }

    control.addEventListener("click", function() {
      openDecisionPickerMenu(control);
    });

    control.addEventListener("focus", function() {
      openDecisionPickerMenu(control);
    });

    control.addEventListener("keydown", function(e) {
      var key = e.key || "";
      if (key === "Enter" || key === " " || key === "ArrowDown") {
        e.preventDefault();
        openDecisionPickerMenu(control);
      }
    });

    control.addEventListener("change", function() {
      var val = trimSafe(control.value || "");
      control.defaultValue = val;
      control.dataset.lastDecisionValue = val;
      control.setAttribute("value", val);
    });
  }

  function getSourceCurrentValue(src) {
    if (!src) return "";
    if (src.tagName === "SELECT") {
      if (src.selectedIndex >= 0 && src.options[src.selectedIndex]) {
        return trimSafe(src.options[src.selectedIndex].text);
      }
      return trimSafe(src.value || "");
    }
    return trimSafe(src.value || src.getAttribute("value") || "");
  }

  function createDecisionPickerControl(sourceControl) {
    if (!sourceControl) return null;

    var input = document.createElement("input");
    input.type = "text";
    var attrs = sourceControl.attributes || [];

    for (var i = 0; i < attrs.length; i++) {
      var attr = attrs[i];
      if (!attr || !attr.name) continue;
      // Skip type/value/readonly/name – name would cause duplicate field submission
      if (attr.name === "type" || attr.name === "value" || attr.name === "readonly" || attr.name === "name") continue;
      input.setAttribute(attr.name, attr.value);
    }

    input.style.cssText = sourceControl.style.cssText || "";
    input.dataset.decisionPicker = "1";
    input.removeAttribute("name");

    var currentValue = normalizeDecisionDisplayValue(
      getSourceCurrentValue(sourceControl) ||
      sourceControl.dataset.lastDecisionValue ||
      "",
      sourceControl.dataset && sourceControl.dataset.decisionRole
    );
    input.value = currentValue;
    input.defaultValue = currentValue;
    input.setAttribute("value", currentValue);
    if (currentValue) input.dataset.lastDecisionValue = currentValue;

    if (sourceControl.disabled) input.disabled = true;

    // Hide the original TISA control in DOM – TISA still finds it for serialization
    sourceControl.dataset.decisionPickerHidden = "1";
    sourceControl.style.cssText = "display:none;";
    sourceControl.__decisionDisplayInput = input;

    // Insert display input immediately before the hidden source
    if (sourceControl.parentNode) {
      sourceControl.parentNode.insertBefore(input, sourceControl);
    }

    input.__decisionHiddenSource = sourceControl;
    bindDecisionPickerControl(input);
    return input;
  }

  function ensureTaskFormDecisionSelects(panel) {
    var rows = getDataRows(panel);

    for (var r = 0; r < rows.length; r++) {
      var controls = getEditableAdvancedControlsInRow(rows[r]);
      var indexes = [TASKFORM_COL_STAV_SKLAD, TASKFORM_COL_STAV_PRODEJNA];

      for (var i = 0; i < indexes.length; i++) {
        var control = controls[indexes[i]];
        if (!control) continue;

        control.dataset.decisionRole = indexes[i] === TASKFORM_COL_STAV_PRODEJNA ? "prodejna" : "sklad";

        if (control.dataset.decisionPicker !== "1") {
          control = createDecisionPickerControl(control);
          if (control) control.dataset.decisionRole = indexes[i] === TASKFORM_COL_STAV_PRODEJNA ? "prodejna" : "sklad";
        } else {
          // Sync display value from hidden TISA source in case TISA set it externally
          var hiddenSrc = control.__decisionHiddenSource;
          if (hiddenSrc) {
            var srcVal = getSourceCurrentValue(hiddenSrc);
            if (srcVal) {
              var displayVal = normalizeDecisionDisplayValue(srcVal, control.dataset.decisionRole);
              if (displayVal !== trimSafe(control.value || "")) {
                control.value = displayVal;
                control.defaultValue = displayVal;
                control.setAttribute("value", displayVal);
                control.dataset.lastDecisionValue = displayVal;
              }
            }
          }
          bindDecisionPickerControl(control);
        }

        setDecisionVisualState(control);
      }
    }
  }

  function setDecisionVisualState(el) {
    if (!el) return;

    var rawValue = typeof el.value === "string" ? el.value : (el.textContent || "");
    var val = toCanonicalDecisionValue(rawValue);

    el.style.color = "";
    el.style.fontWeight = "";

    if (val === "rozdil_potvrzen") {
      el.style.color = "#107c10";
      el.style.fontWeight = "600";
      return;
    }

    if (val === "rozdil_nepotvrzen") {
      el.style.color = "#a80000";
      el.style.fontWeight = "600";
      return;
    }
  }

  function getDecisionTargetInRow(tr, colIndex) {
    if (!tr) return null;

    var controls = getEditableAdvancedControlsInRow(tr);
    if (controls[colIndex]) return controls[colIndex];

    var cells = tr.querySelectorAll("td");
    return cells && cells[colIndex] ? cells[colIndex] : null;
  }

  function refreshTaskFormDecisionColors(panel) {
    var rows = getDataRows(panel);
    for (var r = 0; r < rows.length; r++) {
      var skladTarget = getDecisionTargetInRow(rows[r], TASKFORM_COL_STAV_SKLAD);
      var prodejnaTarget = getDecisionTargetInRow(rows[r], TASKFORM_COL_STAV_PRODEJNA);

      if (skladTarget) setDecisionVisualState(skladTarget);
      if (prodejnaTarget) setDecisionVisualState(prodejnaTarget);
    }
  }

  function initDisplayFormDecisionColors(attempt) {
    attempt = attempt || 0;

    var panel = getAdvancedTablePanel();
    if (!panel) {
      if (attempt < 20) {
        setTimeout(function() {
          initDisplayFormDecisionColors(attempt + 1);
        }, 400);
      }
      return;
    }

    refreshTaskFormDecisionColors(panel);
  }

  function openTaskFormNoteModal(options) {
    options = options || {};

    var decisionLabel = trimSafe(options.decisionLabel || "");
    var decisionValue = toCanonicalDecisionValue(decisionLabel);
    var modalTitle = decisionValue === "rozdil_potvrzen"
      ? "Rozdíl potvrzen – doplň skutečný stav"
      : "Rozdíl nepotvrzen – doplň skutečný stav";

    var existing = document.getElementById("taskFormNoteBackdrop");
    if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

    var backdrop = document.createElement("div");
    backdrop.id = "taskFormNoteBackdrop";
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:10001;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:18px 20px;border-radius:6px;width:520px;max-width:94vw;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 3px 12px rgba(0,0,0,0.35);";

    dialog.innerHTML =
      "<h2 style='margin:0 0 8px 0;font-size:18px;'>" + escapeHtml(modalTitle) + "</h2>" +
      "<div style='font-size:13px;color:#323130;margin-bottom:10px;'>" +
      "Řádek <strong>" + escapeHtml(options.rowNumber || "") + "</strong>, sloupec <strong>" + escapeHtml(options.columnLabel || "") + "</strong>, hodnota <strong>" + escapeHtml(options.decisionLabel || "") + "</strong>." +
      "</div>" +
      "<label style='display:block;font-size:12px;font-weight:600;margin-bottom:4px;'>Skutečný napočítaný stav</label>" +
      "<input type='text' id='taskFormActualState' style='width:100%;box-sizing:border-box;padding:8px;border:1px solid #c8c6c4;border-radius:4px;'>" +
      "<label style='display:block;font-size:12px;font-weight:600;margin:10px 0 4px 0;'>Komentář (volitelné)</label>" +
      "<textarea id='taskFormNoteText' style='width:100%;min-height:100px;box-sizing:border-box;padding:8px;border:1px solid #c8c6c4;border-radius:4px;resize:vertical;'></textarea>" +
      "<div id='taskFormNoteErr' style='display:none;margin-top:8px;color:#a80000;font-size:12px;'></div>" +
      "<div style='display:flex;justify-content:flex-end;gap:8px;margin-top:14px;'>" +
      "  <button type='button' id='taskFormNoteCancel' style='padding:8px 12px;border-radius:4px;border:1px solid #8a8886;background:#fff;cursor:pointer;'>Zrušit</button>" +
      "  <button type='button' id='taskFormNoteSave' style='padding:8px 12px;border-radius:4px;border:1px solid #107c10;background:#107c10;color:#fff;cursor:pointer;'>Zapsat</button>" +
      "</div>";

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    var actual = dialog.querySelector("#taskFormActualState");
    var ta = dialog.querySelector("#taskFormNoteText");
    var err = dialog.querySelector("#taskFormNoteErr");

    dialog.querySelector("#taskFormNoteCancel").onclick = function() {
      backdrop.remove();
    };

    dialog.querySelector("#taskFormNoteSave").onclick = function() {
      var actualState = trimSafe(actual.value || "");
      var note = trimSafe(ta.value || "");

      if (!actualState) {
        err.textContent = "Doplňte skutečný napočítaný stav.";
        err.style.display = "block";
        actual.focus();
        return;
      }

      var ok = true;
      if (typeof options.onSave === "function") ok = options.onSave(actualState, note);

      if (ok === false) {
        err.textContent = "Nepodařilo se zapsat údaje do pole koment_rozdil.";
        err.style.display = "block";
        return;
      }

      backdrop.remove();
    };

    setTimeout(function() {
      try { actual.focus(); } catch (e) {}
    }, 0);
  }

  function openDecisionChangeConfirmModal(options) {
    options = options || {};

    var existing = document.getElementById("taskFormDecisionChangeBackdrop");
    if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

    var backdrop = document.createElement("div");
    backdrop.id = "taskFormDecisionChangeBackdrop";
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:10003;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:18px 20px;border-radius:6px;width:560px;max-width:94vw;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 3px 12px rgba(0,0,0,0.35);";

    dialog.innerHTML =
      "<h2 style='margin:0 0 8px 0;font-size:18px;'>Potvrzení změny rozhodnutí</h2>" +
      "<div style='font-size:13px;color:#323130;margin-bottom:8px;'>" +
      "Měníte rozhodnutí u <strong>" + escapeHtml(options.itemLabel || "") + "</strong>, sloupec <strong>" + escapeHtml(options.columnLabel || "") + "</strong>." +
      "</div>" +
      "<div style='font-size:13px;color:#323130;margin-bottom:12px;'>" +
      "Původní hodnota: <strong>" + escapeHtml(options.fromLabel || "") + "</strong> → Nová hodnota: <strong>" + escapeHtml(options.toLabel || "") + "</strong>." +
      "</div>" +
      "<div style='font-size:12px;color:#605e5c;margin-bottom:14px;'>" +
      "Potvrďte, že jde o zamýšlenou opravu." +
      "</div>" +
      "<div style='display:flex;justify-content:flex-end;gap:8px;'>" +
      "  <button type='button' id='taskFormDecisionChangeCancel' style='padding:8px 12px;border-radius:4px;border:1px solid #8a8886;background:#fff;cursor:pointer;'>Ne, vrátit zpět</button>" +
      "  <button type='button' id='taskFormDecisionChangeConfirm' style='padding:8px 12px;border-radius:4px;border:1px solid #0078d4;background:#0078d4;color:#fff;cursor:pointer;'>Ano, potvrdit změnu</button>" +
      "</div>";

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    dialog.querySelector("#taskFormDecisionChangeCancel").onclick = function() {
      backdrop.remove();
      if (typeof options.onCancel === "function") options.onCancel();
    };

    dialog.querySelector("#taskFormDecisionChangeConfirm").onclick = function() {
      backdrop.remove();
      if (typeof options.onConfirm === "function") options.onConfirm();
    };
  }

  function initTaskFormCommentBinding(attempt) {
    attempt = attempt || 0;

    var panel = getAdvancedTablePanel();
    if (!panel) {
      if (attempt < 20) {
        setTimeout(function() {
          initTaskFormCommentBinding(attempt + 1);
        }, 400);
      }
      return;
    }

    ensureTaskFormDecisionSelects(panel);
    refreshTaskFormDecisionColors(panel);
    syncTocHtmlFromAdvancedTable(panel);

    // Delayed re-syncs catch TISA setting source values after initial DOM render
    var lateSync = 0;
    function lateDecisionSync() {
      lateSync++;
      ensureTaskFormDecisionSelects(panel);
      refreshTaskFormDecisionColors(panel);
      syncTocHtmlFromAdvancedTable(panel);
      if (lateSync < 6) setTimeout(lateDecisionSync, 300);
    }
    setTimeout(lateDecisionSync, 200);

    if (panel.__taskFormCommentBound) return;
    panel.__taskFormCommentBound = true;

    if (!panel.__taskFormDecisionObserver) {
      panel.__taskFormDecisionObserver = new MutationObserver(function() {
        ensureTaskFormDecisionSelects(panel);
        refreshTaskFormDecisionColors(panel);
        syncTocHtmlFromAdvancedTable(panel);
      });

      panel.__taskFormDecisionObserver.observe(panel, {
        childList: true,
        subtree: true
      });
    }

    function getTrackedColumnIndex(target, tr) {
      var controls = getEditableAdvancedControlsInRow(tr);
      if (!controls || controls.length === 0) return -1;

      for (var i = 0; i < controls.length; i++) {
        if (controls[i] === target) return i;
      }

      return -1;
    }

    function handleDecisionInteraction(target, tr) {
      if (target.__decisionInteractionSuppressed) return;

      var colIndex = getTrackedColumnIndex(target, tr);
      if (colIndex !== TASKFORM_COL_STAV_SKLAD && colIndex !== TASKFORM_COL_STAV_PRODEJNA) return;

      setDecisionVisualState(target);

      var decisionLabel = trimSafe(target.value || "");
      var decisionValue = toCanonicalDecisionValue(decisionLabel);
      if (!isTrackedDecisionValue(decisionValue)) return;

      var rowNumber = getDataRows(panel).indexOf(tr) + 1;
      if (rowNumber <= 0) rowNumber = 1;

      var dedupeKey = rowNumber + "|" + colIndex + "|" + decisionValue;
      var now = Date.now();
      if (panel.__decisionModalKey === dedupeKey && now - (panel.__decisionModalTs || 0) < 300) return;
      panel.__decisionModalKey = dedupeKey;
      panel.__decisionModalTs = now;

      var controls = getEditableAdvancedControlsInRow(tr);
      var goodsValue = controls[COL_GOODS]
        ? trimSafe(controls[COL_GOODS].value || controls[COL_GOODS].textContent || "")
        : "";
      var goodsNumber = stripLookupInfo(goodsValue);
      var locationText = controls[COL_LOCATION]
        ? trimSafe(controls[COL_LOCATION].value || controls[COL_LOCATION].textContent || "")
        : "";

      var itemLabel = goodsNumber
        ? "Číslo zboží " + goodsNumber
        : "Řádek " + rowNumber;

      updateStatusByLocationDecision(locationText, decisionValue);

      var columnLabel = colIndex === TASKFORM_COL_STAV_SKLAD
        ? "Vyjádření sklad"
        : "Vyjádření prodejna";

      var previousDecision = toCanonicalDecisionValue(target.dataset.lastCommittedDecision || "");
      var requiresDecisionChangeConfirmation =
        isTrackedDecisionValue(previousDecision) &&
        previousDecision !== decisionValue;

      function commitDecisionAction() {
        target.dataset.lastCommittedDecision = decisionValue;
        syncTocHtmlFromAdvancedTable(panel);

        if (decisionValue === "rozdil_potvrzen") {
          appendKomentRozdilNote(itemLabel, columnLabel, decisionLabel, "", "");
          return;
        }

        openTaskFormNoteModal({
          rowNumber: rowNumber,
          columnLabel: columnLabel,
          decisionLabel: decisionLabel,
          onSave: function(actualState, noteText) {
            return appendKomentRozdilNote(itemLabel, columnLabel, decisionLabel, actualState, noteText);
          }
        });
      }

      if (requiresDecisionChangeConfirmation) {
        if (target.__decisionChangeConfirmOpen) return;
        target.__decisionChangeConfirmOpen = true;

        var previousLabel = decisionValueToLabel(previousDecision, target.dataset.decisionRole);

        openDecisionChangeConfirmModal({
          itemLabel: itemLabel,
          columnLabel: columnLabel,
          fromLabel: previousLabel,
          toLabel: decisionLabel,
          onConfirm: function() {
            target.__decisionChangeConfirmOpen = false;
            appendKomentRozdilNote(
              itemLabel,
              columnLabel,
              "Oprava rozhodnutí",
              "",
              "Uživatel potvrdil změnu vyjádření z \"" + previousLabel + "\" na \"" + decisionLabel + "\"."
            );
            commitDecisionAction();
          },
          onCancel: function() {
            target.__decisionChangeConfirmOpen = false;
            target.__decisionInteractionSuppressed = true;
            setDecisionControlValue(target, previousLabel, false);
            syncTocHtmlFromAdvancedTable(panel);
            setTimeout(function() {
              target.__decisionInteractionSuppressed = false;
            }, 0);
          }
        });

        return;
      }

      commitDecisionAction();
    }

    panel.addEventListener("input", function(e) {
      var target = e.target;
      if (!target) return;

      var tr = target.closest ? target.closest("tr.dataRow") : null;
      if (!tr) return;

      handleDecisionInteraction(target, tr);
    }, true);

    panel.addEventListener("change", function(e) {
      var target = e.target;
      if (!target) return;

      var tr = target.closest ? target.closest("tr.dataRow") : null;
      if (!tr) return;

      handleDecisionInteraction(target, tr);
    }, true);
  }

  // ===================== Modal UI: type chooser + table editor ==============

  function toNum(v) {
    var n = parseFloat(String(v || "").replace(",", "."));
    return isNaN(n) ? 0 : n;
  }

  function calcDiff(dl, fiz) {
    return toNum(fiz) - toNum(dl);
  }

  function escapeHtml(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function createTocEditorModal() {
    var existing = parseExistingRowsFromHidden();

    var backdrop = document.createElement("div");
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:10000;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:18px 20px;border-radius:6px;width:860px;max-width:96vw;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 3px 12px rgba(0,0,0,0.35);";

    dialog.innerHTML =
      "<div style='display:flex;align-items:center;justify-content:space-between;gap:12px;'>" +
      "  <div>" +
      "    <h2 style='margin:0;font-size:18px;'>Rozdíl v transferu – položky</h2>" +
      "    <div style='margin-top:4px;font-size:12px;color:#605e5c;'>Vyplň řádky. Rozdíl se dopočítá.</div>" +
      "  </div>" +
      "  <button type='button' id='tocAddRow' style='padding:6px 10px;border-radius:4px;border:1px solid #0078d4;background:#0078d4;color:#fff;cursor:pointer;'>Přidat řádek</button>" +
      "</div>" +
      "<div style='margin-top:12px;overflow:auto;max-height:55vh;border:1px solid #edebe9;border-radius:4px;'>" +
      "  <table style='border-collapse:collapse;width:100%;min-width:760px;'>" +
      "    <thead><tr style='background:#faf9f8;'>" +
      "      <th style='text-align:left;border-bottom:1px solid #edebe9;padding:8px;'>Číslo zboží</th>" +
      "      <th style='text-align:left;border-bottom:1px solid #edebe9;padding:8px;'>Lokace</th>" +
      "      <th style='text-align:left;border-bottom:1px solid #edebe9;padding:8px;'>Množství na DL</th>" +
      "      <th style='text-align:left;border-bottom:1px solid #edebe9;padding:8px;'>Fyzické dodané množství</th>" +
      "      <th style='text-align:left;border-bottom:1px solid #edebe9;padding:8px;'>Rozdíl</th>" +
      "      <th style='text-align:left;border-bottom:1px solid #edebe9;padding:8px;'></th>" +
      "    </tr></thead>" +
      "    <tbody id='tocBody'></tbody>" +
      "  </table>" +
      "</div>" +
      "<div id='tocErr' style='display:none;margin-top:10px;color:#a80000;font-size:12px;'></div>" +
      "<div style='display:flex;justify-content:flex-end;gap:8px;margin-top:14px;'>" +
      "  <button type='button' id='tocCancel' style='padding:8px 12px;border-radius:4px;border:1px solid #8a8886;background:#fff;cursor:pointer;'>Zrušit</button>" +
      "  <button type='button' id='tocSave' style='padding:8px 12px;border-radius:4px;border:1px solid #107c10;background:#107c10;color:#fff;cursor:pointer;'>Uložit</button>" +
      "</div>";

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    function showError(msg) {
      var err = document.getElementById("tocErr");
      err.textContent = msg;
      err.style.display = "block";
    }

    function clearError() {
      var err = document.getElementById("tocErr");
      err.textContent = "";
      err.style.display = "none";
    }

    function renderRowFromCells(cells) {
      var c = normalizeCells(cells);

      var tr = document.createElement("tr");
      tr.style.borderBottom = "1px solid #edebe9";

      tr.innerHTML =
        "<td style='padding:8px;'><input type='text' class='toc-cislo' style='width:100%;box-sizing:border-box;padding:6px;' value='" + escapeHtml(c[0]) + "'></td>" +
        "<td style='padding:8px;'><input type='text' class='toc-lokace' style='width:100%;box-sizing:border-box;padding:6px;background:#f3f2f1;' value='" + escapeHtml(c[1]) + "' readonly disabled></td>" +
        "<td style='padding:8px;'><input type='text' class='toc-dl' style='width:100%;box-sizing:border-box;padding:6px;' value='" + escapeHtml(c[2]) + "'></td>" +
        "<td style='padding:8px;'><input type='text' class='toc-fiz' style='width:100%;box-sizing:border-box;padding:6px;' value='" + escapeHtml(c[3]) + "'></td>" +
        "<td style='padding:8px;'><input type='text' class='toc-rozdil' style='width:100%;box-sizing:border-box;padding:6px;background:#f3f2f1;' value='" + escapeHtml(c[4]) + "' disabled></td>" +
        "<td style='padding:8px;'><button type='button' class='toc-del' style='padding:6px 10px;border-radius:4px;border:1px solid #8a8886;background:#fff;cursor:pointer;'>Smazat</button></td>";

      function recalc() {
        var dl = tr.querySelector(".toc-dl").value;
        var fiz = tr.querySelector(".toc-fiz").value;
        tr.querySelector(".toc-rozdil").value = String(calcDiff(dl, fiz));
      }

      tr.querySelector(".toc-dl").addEventListener("input", recalc);
      tr.querySelector(".toc-fiz").addEventListener("input", recalc);
      tr.querySelector(".toc-del").addEventListener("click", function() {
        tr.remove();
      });

      recalc();
      return tr;
    }

    function addEmptyRow() {
      document.getElementById("tocBody").appendChild(renderRowFromCells(["", "", "", "", ""]));
    }

    if (existing.length > 0) {
      for (var i = 0; i < existing.length; i++) {
        document.getElementById("tocBody").appendChild(renderRowFromCells(existing[i].Cells));
      }
    } else {
      addEmptyRow();
    }

    document.getElementById("tocAddRow").addEventListener("click", addEmptyRow);

    document.getElementById("tocCancel").addEventListener("click", function() {
      backdrop.remove();
    });

    document.getElementById("tocSave").addEventListener("click", function() {
      clearError();

      var saveBtn = document.getElementById("tocSave");
      var oldSaveText = saveBtn.textContent;
      saveBtn.disabled = true;
      saveBtn.textContent = "Načítám data...";

      var rows = [];
      var trs = document.querySelectorAll("#tocBody tr");

      for (var i = 0; i < trs.length; i++) {
        var tr = trs[i];

        var cislo = (tr.querySelector(".toc-cislo").value || "").trim();
        var lokace = (tr.querySelector(".toc-lokace").value || "").trim();
        var dl = (tr.querySelector(".toc-dl").value || "").trim();
        var fiz = (tr.querySelector(".toc-fiz").value || "").trim();
        var rozdil = (tr.querySelector(".toc-rozdil").value || "").trim();

        if (!cislo && !dl && !fiz) continue;
        if (!cislo) {
          saveBtn.disabled = false;
          saveBtn.textContent = oldSaveText;
          showError("Řádek " + (i + 1) + ": chybí Číslo zboží.");
          return;
        }

        rows.push({
          Cells: normalizeCells([cislo, lokace, dl, fiz, rozdil])
        });
      }

      enrichRowsFromLookup(rows, function(enrichedRows) {
        var ok = applyRowsToTisaAdvancedTable(enrichedRows);
        if (!ok) {
          saveBtn.disabled = false;
          saveBtn.textContent = oldSaveText;
          showError("Nepodařilo se najít/naplnit Treeinfo Advanced Table na formuláři.");
          return;
        }

        // HTML vždy synchronizovat podle aktuálního stavu v Advanced Table na formuláři.
        syncTocHtmlFromAdvancedTable();

        // Zapsat agregovanou lokaci do pole lokace_pro_zbozi
        var lokaceValue = computeLokaceProZbozi(enrichedRows);
        if (lokaceValue) {
          debugLog("[lokace_pro_zbozi] enrichedRows:", enrichedRows);
          var setOk = setChoiceLikeField("lokace_pro_zbozi", lokaceValue);
          debugLog("[lokace_pro_zbozi] vysledek nastaveni primary:", setOk);

          // Kompatibilita pro starší/opravené varianty názvu kombinované volby.
          if (!setOk && lokaceValue === LOKACE_CHOICE_COMBINED) {
            var fallbackOk = setChoiceLikeField("lokace_pro_zbozi", "BRAMB_CENTR");
            debugLog("[lokace_pro_zbozi] vysledek nastaveni fallback BRAMB_CENTR:", fallbackOk);

            if (!fallbackOk) {
              fallbackOk = setChoiceLikeField("lokace_pro_zbozi", "CENTR_BRAMB");
              debugLog("[lokace_pro_zbozi] vysledek nastaveni fallback CENTR_BRAMB:", fallbackOk);
            }
          }
        }

        backdrop.remove();
      });
    });
  }

  function createTypeModal() {
    var backdrop = document.createElement("div");
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:9999;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:22px 26px;border-radius:6px;width:420px;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 3px 12px rgba(0,0,0,0.35);";

    dialog.innerHTML =
      "<h2 style='margin-top:0;margin-bottom:12px;font-size:18px;'>Vyber typ chyby</h2>" +
      "<p style='margin:0 0 12px 0;font-size:13px;line-height:1.4;'>Zvol typ chyby pro tento záznam transferu.</p>" +
      "<div style='display:flex;flex-direction:column;gap:10px;margin-top:10px;'>" +
      "  <button id='btRozdil' style='padding:10px 12px;border-radius:4px;border:1px solid #0078d4;background:#0078d4;color:white;font-size:14px;cursor:pointer;text-align:left;'>Rozdíl v transferu</button>" +
      "  <button id='btProblem' style='padding:10px 12px;border-radius:4px;border:1px solid #666;background:#f3f2f1;color:#3231v30;font-size:14px;cursor:pointer;text-align:left;'>Problém s transferem</button>" +
      "</div>";

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    // === PROHOZENÉ ZAMYKÁNÍ POLÍ ===

    // Rozdíl v transferu -> zamknout problem pole + otevřít AdvancedTable modal
    document.getElementById("btRozdil").onclick = function() {
      setChoiceLikeField("typ_chyby", "Rozdíl v transferu");

      [
        "identifikace_problemu",
        "ostatni_problem",
        "podrobny_popis",
        "priloha_transfer"
      ].forEach(lockField);

      backdrop.remove();
      createTocEditorModal();
    };

    // Problém s transferem -> zamknout sadu pro rozdíl
    document.getElementById("btProblem").onclick = function() {
      setChoiceLikeField("typ_chyby", "Problém s transferem");

      [
        "datum_prijmu_transferu",
        "cislo_dl",
        "cislo_zbozi",
        "mnozstvi_dl",
        "fyzicke_mnozstvi",
        "rozdil_dl_fyz",
        "cislo_transferu"
      ].forEach(lockField);

      backdrop.remove();
    };
  }

  // ===================== INIT ===============================

  function isNewForm() {
    var p = (location.pathname || "").toLowerCase();
    if (p.indexOf("newform.aspx") !== -1) return true;
    return (location.search || "").toLowerCase().indexOf("formtype=new") !== -1;
  }

  function isEditTaskForm() {
    var p = (location.pathname || "").toLowerCase();
    var q = (location.search || "").toLowerCase();
    return p.indexOf("editform.aspx") !== -1 && q.indexOf("istaskform") !== -1;
  }

  function isDispForm() {
    var p = (location.pathname || "").toLowerCase();
    return p.indexOf("dispform.aspx") !== -1 || p.indexOf("displayform.aspx") !== -1;
  }

  function hideHtmlFieldRow() {
    var r = getTisaRowByInternalName(TOC_HTML_FIELD);
    if (r) r.style.display = "none";

    var hidden = document.querySelector("input[id$='_FieldName'][value='" + TOC_HTML_FIELD + "']");
    var row = hidden ? hidden.closest("tr") : null;
    if (row) row.style.display = "none";
  }

  onReady(function() {
    if (isNewForm()) {
      hideHtmlFieldRow();
      createTypeModal();
      return;
    }

    if (isEditTaskForm()) {
      initTaskFormCommentBinding();
      return;
    }

    if (isDispForm()) {
      initDisplayFormDecisionColors();
    }
  });

})();