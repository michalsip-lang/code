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

  function lockAdvancedTablePanel(panel) {
    if (!panel) return false;

    // Block any user interaction in the Treeinfo table (+, remove, edit).
    panel.style.pointerEvents = "none";
    panel.style.opacity = "0.85";

    var controls = panel.querySelectorAll("input, select, textarea, button");
    Array.prototype.forEach.call(controls, function(ctrl) {
      var tag = ctrl.tagName;
      if (tag === "INPUT" || tag === "TEXTAREA") ctrl.setAttribute("readonly", "readonly");
      ctrl.setAttribute("disabled", "disabled");
      ctrl.style.pointerEvents = "none";
      ctrl.style.backgroundColor = "#f3f2f1";
      ctrl.style.borderColor = "#e1dfdd";
    });

    if (!panel.__lockedObserver) {
      panel.__lockedObserver = new MutationObserver(function() {
        lockAdvancedTablePanel(panel);
      });

      panel.__lockedObserver.observe(panel, {
        childList: true,
        subtree: true
      });
    }

    return true;
  }

  function ensureAdvancedTableLocked(internalName, attempt) {
    attempt = attempt || 0;

    var row = getTisaRowByInternalName(internalName);
    var panel = row ? row.querySelector(".tisa-AdvancedTable") : null;
    if (panel) {
      lockAdvancedTablePanel(panel);
      return;
    }

    if (attempt < 20) {
      setTimeout(function() {
        ensureAdvancedTableLocked(internalName, attempt + 1);
      }, 200);
    }
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

  function fireHtmlEvent(el, eventName) {
    if (!el || !eventName) return;
    try {
      var ev = document.createEvent("HTMLEvents");
      ev.initEvent(eventName, true, false);
      el.dispatchEvent(ev);
    } catch (e) {}
  }

  function fireKeyboardEvent(el, eventName, keyValue) {
    if (!el || !eventName) return;

    try {
      var keyboardEvent = new KeyboardEvent(eventName, {
        bubbles: true,
        cancelable: true,
        key: keyValue || "",
        code: keyValue || "",
        keyCode: keyValue && keyValue.length === 1 ? keyValue.charCodeAt(0) : 0,
        which: keyValue && keyValue.length === 1 ? keyValue.charCodeAt(0) : 0
      });
      el.dispatchEvent(keyboardEvent);
      return;
    } catch (keyboardError) {}

    try {
      var legacyEvent = document.createEvent("Event");
      legacyEvent.initEvent(eventName, true, true);
      legacyEvent.keyCode = keyValue && keyValue.length === 1 ? keyValue.charCodeAt(0) : 0;
      legacyEvent.which = legacyEvent.keyCode;
      el.dispatchEvent(legacyEvent);
    } catch (legacyError) {}
  }

  function triggerLookupFieldBehavior(el, value) {
    if (!el) return false;

    var normalizedValue = trimSafe(value || "");
    var lastChar = normalizedValue ? normalizedValue.charAt(normalizedValue.length - 1) : "";

    try { el.focus(); } catch (focusError) {}
    try { el.click(); } catch (clickError) {}

    if (typeof el.value !== "undefined" && el.value !== normalizedValue) {
      el.value = normalizedValue;
    }

    fireHtmlEvent(el, "focus");
    fireHtmlEvent(el, "click");
    fireHtmlEvent(el, "input");
    fireKeyboardEvent(el, "keydown", lastChar);
    fireKeyboardEvent(el, "keypress", lastChar);
    fireKeyboardEvent(el, "keyup", lastChar);
    fireHtmlEvent(el, "change");
    triggerJQueryAutocompleteSearch(el, normalizedValue);

    // SharePoint lookup/autocomplete often reacts only after the initial input settles.
    setTimeout(function() {
      if (!document.body.contains(el)) return;
      fireHtmlEvent(el, "input");
      fireKeyboardEvent(el, "keyup", lastChar);
      fireHtmlEvent(el, "change");
      triggerJQueryAutocompleteSearch(el, normalizedValue);
      debugLog("[newform_center] lookup retriggered for field", el.id || el.name || el.title || "?", "value", el.value || "");
    }, 250);

    setTimeout(function() {
      if (!document.body.contains(el)) return;
      fireHtmlEvent(el, "blur");
    }, 500);

    return true;
  }

  function exposeNewFormLookupDebug() {
    window.__debugIdentifikaceProdejny2 = function() {
      var ctrl = getWritableFieldControl("identifikace_prodejny_2");
      var allControls = getControlsByInternalName("identifikace_prodejny_2");
      var hiddenCtrl = null;
      var suggestionNodes = [];
      var suggestionTexts = [];
      var i;

      if (!ctrl) {
        console.log("[newform_center] Field identifikace_prodejny_2 not found.");
        return null;
      }

      for (i = 0; i < allControls.length; i++) {
        if ((allControls[i].type || "").toLowerCase() === "hidden") {
          hiddenCtrl = allControls[i];
          break;
        }
      }

      suggestionNodes = getAutocompleteSuggestionNodes();
      for (i = 0; i < suggestionNodes.length; i++) {
        suggestionTexts.push(trimSafe(suggestionNodes[i].textContent || suggestionNodes[i].innerText || ""));
      }

      var info = {
        id: ctrl.id || "",
        name: ctrl.name || "",
        title: ctrl.title || "",
        tagName: ctrl.tagName || "",
        type: ctrl.type || "",
        value: ctrl.value || "",
        className: ctrl.className || "",
        hiddenValue: hiddenCtrl ? (hiddenCtrl.value || "") : "",
        hiddenId: hiddenCtrl ? (hiddenCtrl.id || "") : "",
        suggestionCount: suggestionNodes.length,
        suggestions: suggestionTexts,
        hasJQuery: typeof window.jQuery !== "undefined",
        hasUiAutocomplete: !!(window.jQuery && window.jQuery.ui && window.jQuery.ui.autocomplete)
      };

      console.log("[newform_center] identifikace_prodejny_2 snapshot", info, ctrl);
      return info;
    };
  }

  function getHiddenLookupControl(internalName) {
    var controls = getControlsByInternalName(internalName);
    for (var i = 0; i < controls.length; i++) {
      if ((controls[i].type || "").toLowerCase() === "hidden") return controls[i];
    }
    return null;
  }

  function getAutocompleteSuggestionNodes() {
    var selectors = [
      ".ui-autocomplete li",
      ".ui-menu-item",
      ".ui-autocomplete .ui-menu-item",
      "ul.ui-autocomplete li a",
      ".autocomplete ul li",
      ".searchBox + ul li"
    ];

    for (var i = 0; i < selectors.length; i++) {
      var nodes = document.querySelectorAll(selectors[i]);
      if (nodes && nodes.length) return Array.prototype.slice.call(nodes);
    }

    return [];
  }

  function getAutocompleteWidgetInstance(el) {
    if (!el || !window.jQuery) return null;

    try {
      var $el = window.jQuery(el);
      if (!$el || !$el.length || !$el.autocomplete) return null;

      return $el.data("ui-autocomplete") || $el.data("autocomplete") || null;
    } catch (e) {
      return null;
    }
  }

  function triggerJQueryAutocompleteSearch(el, value) {
    if (!el || !window.jQuery) return false;

    try {
      var $el = window.jQuery(el);
      if (!$el || !$el.length || !$el.autocomplete) return false;

      $el.autocomplete("search", value || "");
      debugLog("[newform_center] jQuery autocomplete search triggered", value || "", el.id || el.name || "?");
      return true;
    } catch (e) {
      debugLog("[newform_center] jQuery autocomplete search failed", e);
      return false;
    }
  }

  function selectFirstAutocompleteSuggestionViaJquery(el, expectedText) {
    if (!el || !window.jQuery) return false;

    try {
      var $el = window.jQuery(el);
      if (!$el || !$el.length || !$el.autocomplete) return false;

      var instance = getAutocompleteWidgetInstance(el);
      var widget = $el.autocomplete("widget");
      var firstItem = widget && widget.length ? widget.find("li:visible").first() : null;
      var itemData = firstItem && firstItem.length
        ? (firstItem.data("ui-autocomplete-item") || firstItem.data("item.autocomplete") || null)
        : null;

      if (!firstItem || !firstItem.length) return false;

      debugLog(
        "[newform_center] jQuery autocomplete first suggestion",
        expectedText || "",
        itemData ? (itemData.label || itemData.value || "") : trimSafe(firstItem.text() || "")
      );

      if (instance && itemData && typeof instance._trigger === "function") {
        instance._trigger("select", null, { item: itemData });
      }

      try { firstItem.trigger("mouseenter"); } catch (mouseenterError) {}
      try { firstItem.trigger("click"); } catch (clickError) {}
      return true;
    } catch (e) {
      debugLog("[newform_center] jQuery autocomplete select failed", e);
      return false;
    }
  }

  function trySelectFirstAutocompleteSuggestion(internalName, expectedText, attempt) {
    attempt = attempt || 0;

    var fieldCtrl = getWritableFieldControl(internalName);
    if (fieldCtrl && selectFirstAutocompleteSuggestionViaJquery(fieldCtrl, expectedText)) {
      setTimeout(function() {
        var hiddenAfterJquery = getHiddenLookupControl(internalName);
        debugLog(
          "[newform_center] hidden lookup after jQuery select",
          hiddenAfterJquery ? trimSafe(hiddenAfterJquery.value || "") : ""
        );
      }, 250);
      return true;
    }

    var nodes = getAutocompleteSuggestionNodes();
    if (!nodes.length) {
      if (attempt < 4) {
        setTimeout(function() {
          trySelectFirstAutocompleteSuggestion(internalName, expectedText, attempt + 1);
        }, 200);
      }
      return false;
    }

    var firstNode = nodes[0];
    var hiddenBefore = getHiddenLookupControl(internalName);
    var beforeValue = hiddenBefore ? trimSafe(hiddenBefore.value || "") : "";

    try { firstNode.click(); } catch (clickError) {}
    fireHtmlEvent(firstNode, "mousedown");
    fireHtmlEvent(firstNode, "mouseup");
    fireHtmlEvent(firstNode, "click");

    setTimeout(function() {
      var hiddenAfter = getHiddenLookupControl(internalName);
      var afterValue = hiddenAfter ? trimSafe(hiddenAfter.value || "") : "";
      debugLog(
        "[newform_center] autocomplete selection attempt",
        expectedText || "",
        "beforeHidden",
        beforeValue,
        "afterHidden",
        afterValue,
        "firstSuggestion",
        trimSafe(firstNode.textContent || firstNode.innerText || "")
      );
    }, 250);

    return true;
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
    if (normalized.indexOf("nepotvrdil") !== -1 || normalized.indexOf("nepotvrdila") !== -1) return "rozdil_nepotvrzen";
    if (normalized.indexOf("nesouhlasi") !== -1) return "rozdil_nepotvrzen";

    if (normalized.indexOf("trva_na_rozdilu") !== -1) return "rozdil_potvrzen";
    if (normalized.indexOf("potvrzen") !== -1 || normalized.indexOf("potvrdil") !== -1 || normalized.indexOf("potvrdila") !== -1) {
      return "rozdil_potvrzen";
    }
    if (normalized.indexOf("souhlasi") !== -1) return "rozdil_potvrzen";

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
  var TASKFORM_COL_OSTATNI_ASSIGN = 5; // 6. sloupec v Edit TaskForm (novy)
  var TASKFORM_COL_STAV_SKLAD = 6; // 7. sloupec v Edit TaskForm
  var TASKFORM_COL_STAV_PRODEJNA = 7; // 8. sloupec v Edit TaskForm

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
  var EMPLOYEE_LIST_URL = "http://portal.samohyl.cz/bozp_po/Lists/zamestnanci/";
  var EMPLOYEE_IDENT_FIELD = "identifikace";
  var EMPLOYEE_CENTER_FIELD = "stredisko_zamestnance";
  var EMPLOYEE_FIRST_NAME_FIELD = "jmeno_zamestnance";
  var EMPLOYEE_LAST_NAME_FIELD = "prijmeni_zamestance";
  var CENTER_CODE_CENTR = "411";
  var CENTER_CODE_BRAMB = "413";

  function resolveTaskFormColumnIndex(controls, logicalIndex) {
    var len = controls && controls.length ? controls.length : 0;

    // New structure: goods, location, dl, fiz, diff, assign, sklad, prodejna
    if (len >= 8) {
      return logicalIndex;
    }

    // Legacy structure without assignment column: goods, location, dl, fiz, diff, sklad, prodejna
    if (logicalIndex === TASKFORM_COL_OSTATNI_ASSIGN) return -1;
    if (logicalIndex === TASKFORM_COL_STAV_SKLAD) return 5;
    if (logicalIndex === TASKFORM_COL_STAV_PRODEJNA) return 6;

    return logicalIndex;
  }

  function mapPhysicalToLogicalTaskFormColumnIndex(controls, physicalIndex) {
    var len = controls && controls.length ? controls.length : 0;

    if (len >= 8) return physicalIndex;

    if (physicalIndex === 5) return TASKFORM_COL_STAV_SKLAD;
    if (physicalIndex === 6) return TASKFORM_COL_STAV_PRODEJNA;
    return physicalIndex;
  }

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

  function getListApiBaseUrlFromUrl(listUrl) {
    var a = document.createElement("a");
    a.href = listUrl;

    var listPath = (a.pathname || "").replace(/\/$/, "");
    if (!listPath) return "";

    var webPath = listPath.replace(/\/Lists\/[^/]+$/i, "");
    if (!webPath) return "";

    var origin = a.protocol + "//" + a.host;
    var escapedListPath = listPath.replace(/'/g, "''");

    return origin + webPath + "/_api/web/GetList('" + escapedListPath + "')";
  }

  function getLookupApiBaseUrl() {
    return getListApiBaseUrlFromUrl(LOOKUP_LIST_URL);
  }

  function getCurrentWebAbsoluteUrl() {
    if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) {
      return window._spPageContextInfo.webAbsoluteUrl.replace(/\/$/, "");
    }

    return (window.location.origin || "").replace(/\/$/, "");
  }

  function getCurrentUserEmail(done) {
    var contextEmail = trimSafe(window._spPageContextInfo && window._spPageContextInfo.userEmail);
    if (contextEmail) {
      done(null, contextEmail);
      return;
    }

    var apiUrl = getCurrentWebAbsoluteUrl() + "/_api/web/currentuser?$select=Email";
    sharePointGetJson(
      apiUrl,
      function(data) {
        var email = trimSafe(data && data.d && data.d.Email);
        if (!email) {
          done(new Error("Aktuální uživatel nemá dostupný email."));
          return;
        }

        done(null, email);
      },
      function(err) {
        done(err || new Error("Nepodařilo se načíst aktuálního uživatele."));
      }
    );
  }

  function normalizePersonName(value) {
    return trimSafe(removeDiacritics(value)).toLowerCase().replace(/\s+/g, " ");
  }

  function splitFullName(fullName) {
    var normalized = trimSafe(fullName).replace(/\s+/g, " ");
    if (!normalized) {
      return {
        firstName: "",
        lastName: ""
      };
    }

    var parts = normalized.split(" ");
    if (parts.length === 1) {
      return {
        firstName: parts[0],
        lastName: ""
      };
    }

    return {
      firstName: parts[0],
      lastName: parts.slice(1).join(" ")
    };
  }

  function getCurrentUserInfo(done) {
    var contextEmail = trimSafe(window._spPageContextInfo && window._spPageContextInfo.userEmail);
    var contextTitle = trimSafe(window._spPageContextInfo && window._spPageContextInfo.userDisplayName);
    if (contextEmail && contextTitle) {
      var splitContext = splitFullName(contextTitle);
      done(null, {
        email: contextEmail,
        displayName: contextTitle,
        firstName: splitContext.firstName,
        lastName: splitContext.lastName
      });
      return;
    }

    var apiUrl = getCurrentWebAbsoluteUrl() + "/_api/web/currentuser?$select=Email,Title";
    sharePointGetJson(
      apiUrl,
      function(data) {
        var currentUser = data && data.d ? data.d : {};
        var email = trimSafe(currentUser.Email || contextEmail);
        var displayName = trimSafe(currentUser.Title || contextTitle);
        var splitName = splitFullName(displayName);

        if (!email) {
          done(new Error("Aktuální uživatel nemá dostupný email."));
          return;
        }

        done(null, {
          email: email,
          displayName: displayName,
          firstName: splitName.firstName,
          lastName: splitName.lastName
        });
      },
      function(err) {
        done(err || new Error("Nepodařilo se načíst aktuálního uživatele."));
      }
    );
  }

  function buildEmployeeExactLookupUrl(email) {
    var apiBase = getListApiBaseUrlFromUrl(EMPLOYEE_LIST_URL);
    if (!apiBase) return "";

    var select = "$select=Id,Title," + EMPLOYEE_IDENT_FIELD + "," + EMPLOYEE_CENTER_FIELD + "," + EMPLOYEE_FIRST_NAME_FIELD + "," + EMPLOYEE_LAST_NAME_FIELD;
    var filter = "$filter=" + encodeURIComponent(EMPLOYEE_IDENT_FIELD + " eq '" + encodeODataLiteral(email) + "'");
    return apiBase + "/items?" + select + "&" + filter + "&$top=50";
  }

  function buildEmployeesListUrl() {
    var apiBase = getListApiBaseUrlFromUrl(EMPLOYEE_LIST_URL);
    if (!apiBase) return "";

    return apiBase + "/items?$select=Id,Title," + EMPLOYEE_IDENT_FIELD + "," + EMPLOYEE_CENTER_FIELD + "," + EMPLOYEE_FIRST_NAME_FIELD + "," + EMPLOYEE_LAST_NAME_FIELD + "&$top=5000";
  }

  function normalizeEmployeeIdentifier(value) {
    return trimSafe(value).toLowerCase();
  }

  function employeeMatchesCurrentUser(item, currentUserInfo) {
    if (!item || !currentUserInfo) return false;

    var itemEmail = normalizeEmployeeIdentifier(item[EMPLOYEE_IDENT_FIELD]);
    if (itemEmail !== normalizeEmployeeIdentifier(currentUserInfo.email)) return false;

    var userFirstName = normalizePersonName(currentUserInfo.firstName);
    var userLastName = normalizePersonName(currentUserInfo.lastName);
    var itemFirstName = normalizePersonName(item[EMPLOYEE_FIRST_NAME_FIELD]);
    var itemLastName = normalizePersonName(item[EMPLOYEE_LAST_NAME_FIELD]);

    if (!userFirstName && !userLastName) return true;

    var directMatch = itemFirstName === userFirstName && itemLastName === userLastName;
    var reversedMatch = itemFirstName === userLastName && itemLastName === userFirstName;
    return directMatch || reversedMatch;
  }

  function getEmployeeCenterByUser(currentUserInfo, done) {
    var email = currentUserInfo && currentUserInfo.email ? currentUserInfo.email : "";
    var normalizedEmail = normalizeEmployeeIdentifier(email);
    if (!normalizedEmail) {
      done(new Error("Chybí email pro dohledání zaměstnance."));
      return;
    }

    var exactUrl = buildEmployeeExactLookupUrl(email);
    if (!exactUrl) {
      done(new Error("Nepodařilo se sestavit URL pro seznam zaměstnanců."));
      return;
    }

    sharePointGetJson(
      exactUrl,
      function(data) {
        var exactItems = data && data.d && data.d.results ? data.d.results : [];
        for (var exactIndex = 0; exactIndex < exactItems.length; exactIndex++) {
          if (!employeeMatchesCurrentUser(exactItems[exactIndex], currentUserInfo)) continue;

          done(null, {
            email: email,
            displayName: currentUserInfo.displayName || "",
            firstName: currentUserInfo.firstName || "",
            lastName: currentUserInfo.lastName || "",
            identifikace: trimSafe(exactItems[exactIndex][EMPLOYEE_IDENT_FIELD]),
            center: trimSafe(exactItems[exactIndex][EMPLOYEE_CENTER_FIELD])
          });
          return;
        }

        var listUrl = buildEmployeesListUrl();
        if (!listUrl) {
          done(new Error("Nepodařilo se sestavit fallback URL pro seznam zaměstnanců."));
          return;
        }

        sharePointGetJson(
          listUrl,
          function(listData) {
            var items = listData && listData.d && listData.d.results ? listData.d.results : [];
            for (var i = 0; i < items.length; i++) {
              if (!employeeMatchesCurrentUser(items[i], currentUserInfo)) continue;

              done(null, {
                email: email,
                displayName: currentUserInfo.displayName || "",
                firstName: currentUserInfo.firstName || "",
                lastName: currentUserInfo.lastName || "",
                identifikace: trimSafe(items[i][EMPLOYEE_IDENT_FIELD]),
                center: trimSafe(items[i][EMPLOYEE_CENTER_FIELD])
              });
              return;
            }

            done(new Error("V seznamu zaměstnanců nebyla nalezena shoda pro email " + email + " a jméno " + (currentUserInfo.displayName || "") + "."));
          },
          function(listErr) {
            done(listErr || new Error("Nepodařilo se načíst seznam zaměstnanců."));
          }
        );
      },
      function(err) {
        done(err || new Error("Nepodařilo se načíst zaměstnance podle emailu."));
      }
    );
  }

  function mapCenterToLocationBranch(centerCode) {
    var center = trimSafe(centerCode);
    if (center === CENTER_CODE_CENTR) return "CENTR";
    if (center === CENTER_CODE_BRAMB) return "BRAMB";
    return "";
  }

  function resolveUserLocationBranch(done) {
    getCurrentUserInfo(function(userErr, currentUserInfo) {
      if (userErr) {
        done(userErr);
        return;
      }

      getEmployeeCenterByUser(currentUserInfo, function(employeeErr, employeeInfo) {
        if (employeeErr) {
          done(employeeErr);
          return;
        }

        done(null, {
          email: currentUserInfo.email,
          displayName: currentUserInfo.displayName || "",
          firstName: currentUserInfo.firstName || "",
          lastName: currentUserInfo.lastName || "",
          center: employeeInfo && employeeInfo.center ? employeeInfo.center : "",
          branch: mapCenterToLocationBranch(employeeInfo && employeeInfo.center)
        });
      });
    });
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
      var missingGoods = [];
      var missingSeen = {};

      for (var j = 0; j < arr.length; j++) {
        var inCells = normalizeCells(arr[j] && arr[j].Cells ? arr[j].Cells : []);
        var baseGoods = stripLookupInfo(inCells[COL_GOODS]);
        var lookupKey = normalizeGoodsKey(baseGoods);
        var lookupItem = lookupMap[lookupKey] || {};
        var lookupInfo = trimSafe(lookupItem.info);
        var lookupLocation = trimSafe(lookupItem.location);

        if (baseGoods && !lookupMap[lookupKey] && !missingSeen[lookupKey]) {
          missingSeen[lookupKey] = true;
          missingGoods.push(baseGoods);
        }

        if (baseGoods && lookupInfo) inCells[COL_GOODS] = baseGoods + LOOKUP_SEPARATOR + lookupInfo;
        else inCells[COL_GOODS] = baseGoods;

        inCells[COL_LOCATION] = lookupLocation || trimSafe(inCells[COL_LOCATION]);

        enriched.push({
          Cells: inCells
        });
      }

      done(enriched, missingGoods);
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
      "Zařazení OSTATNÍ",
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

  function getKomentRozdilValue() {
    var ctrl = getWritableFieldControl("koment_rozdil");
    return ctrl ? trimSafe(ctrl.value || "") : "";
  }

  function komentRozdilToEmailHtml(commentText) {
    var text = trimSafe(commentText);
    if (!text) return "";

    function esc(s) {
      return String(s == null ? "" : s)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    return "" +
      "<div style='margin-top:14px;font-family:Segoe UI, Arial, sans-serif;font-size:13px;'>" +
      "<div style='font-weight:600;margin-bottom:6px;'>Komentář k rozdílu</div>" +
      "<div style='border:1px solid #d0d0d0;padding:8px;white-space:pre-wrap;'>" +
      esc(text) +
      "</div>" +
      "</div>";
  }

  function syncTocHtmlFromAdvancedTable(panel) {
    var rows = collectAdvancedTableRowsFromPanel(panel);
    var htmlTable = advancedTableRowsToEmailHtml(rows);
    var komentarHtml = komentRozdilToEmailHtml(getKomentRozdilValue());
    return setFieldHtml(TOC_HTML_FIELD, htmlTable + komentarHtml);
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
    syncTocHtmlFromAdvancedTable();
    return true;
  }

  function isRozdilVTransferuType() {
    var ctrl = getWritableFieldControl("typ_chyby");
    if (!ctrl) return false;

    var raw = "";
    if (ctrl.tagName === "SELECT") {
      if (ctrl.selectedIndex >= 0 && ctrl.options[ctrl.selectedIndex]) {
        raw = ctrl.options[ctrl.selectedIndex].text || ctrl.value || "";
      } else {
        raw = ctrl.value || "";
      }
    } else {
      raw = ctrl.value || ctrl.getAttribute("value") || "";
    }

    var normalized = normalizeDecisionValue(raw);
    return normalized.indexOf("rozdil v transferu") !== -1;
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

  function sanitizeDecisionRawValue(raw) {
    var value = trimSafe(raw);
    if (!value) return "";

    var normalized = normalizeDecisionValue(value);

    // Placeholder/default option must not be treated as a real decision.
    if (normalized.indexOf("kliknete pro vyber") !== -1) return "";

    return value;
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
        return sanitizeDecisionRawValue(src.options[src.selectedIndex].text || src.value || "");
      }
      return sanitizeDecisionRawValue(src.value || "");
    }
    return sanitizeDecisionRawValue(src.value || src.getAttribute("value") || "");
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
        var resolvedIndex = resolveTaskFormColumnIndex(controls, indexes[i]);
        var control = resolvedIndex >= 0 ? controls[resolvedIndex] : null;
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
    var resolvedIndex = resolveTaskFormColumnIndex(controls, colIndex);
    if (resolvedIndex >= 0 && controls[resolvedIndex]) return controls[resolvedIndex];

    var cells = tr.querySelectorAll("td");
    return cells && cells[resolvedIndex] ? cells[resolvedIndex] : null;
  }

  function normalizeOstatniAssignmentValue(rawValue) {
    var normalized = normalizeDecisionValue(rawValue);
    if (!normalized) return "";

    if (normalized.indexOf("bramb") !== -1) return "JE BRAMB";
    if (normalized.indexOf("centr") !== -1) return "JE CENTR";
    return "";
  }

  function assignmentValueToBranch(rawValue) {
    var normalized = normalizeOstatniAssignmentValue(rawValue);
    if (normalized === "JE BRAMB") return "BRAMB";
    if (normalized === "JE CENTR") return "CENTR";
    return "";
  }

  function setOstatniAssignmentControlValue(control, label, fireEventsToo) {
    if (!control) return false;

    var target = normalizeOstatniAssignmentValue(label);
    if (!target) return false;

    var matched = false;

    if (control.tagName === "SELECT") {
      for (var i = 0; i < control.options.length; i++) {
        var option = control.options[i] || {};
        var byText = normalizeOstatniAssignmentValue(option.text || "");
        var byValue = normalizeOstatniAssignmentValue(option.value || "");
        if (byText === target || byValue === target) {
          control.selectedIndex = i;
          matched = true;
          break;
        }
      }
    } else {
      control.value = target;
      control.defaultValue = target;
      control.setAttribute("value", target);
      matched = true;
    }

    if (!matched) return false;
    if (fireEventsToo) fireInputEvents(control);
    return true;
  }

  function getOstatniAssignmentBranchFromRow(tr) {
    var assignTarget = getDecisionTargetInRow(tr, TASKFORM_COL_OSTATNI_ASSIGN);
    if (!assignTarget) return "";

    var raw = "";
    if (assignTarget.tagName === "SELECT") {
      if (assignTarget.selectedIndex >= 0 && assignTarget.options[assignTarget.selectedIndex]) {
        raw = assignTarget.options[assignTarget.selectedIndex].text || assignTarget.value || "";
      } else {
        raw = assignTarget.value || "";
      }
    } else {
      raw = assignTarget.value || assignTarget.textContent || "";
    }

    return assignmentValueToBranch(raw);
  }

  function getDecisionControlRawValue(control) {
    if (!control) return "";

    var raw = sanitizeDecisionRawValue(control.value || control.textContent || "");
    if (raw) return raw;

    var hiddenSrc = control.__decisionHiddenSource;
    if (!hiddenSrc) return "";

    if (hiddenSrc.tagName === "SELECT") {
      if (hiddenSrc.selectedIndex >= 0 && hiddenSrc.options[hiddenSrc.selectedIndex]) {
        return sanitizeDecisionRawValue(hiddenSrc.options[hiddenSrc.selectedIndex].text || hiddenSrc.value || "");
      }
      return sanitizeDecisionRawValue(hiddenSrc.value || "");
    }

    return sanitizeDecisionRawValue(hiddenSrc.value || hiddenSrc.getAttribute("value") || "");
  }

  function isDecisionControlFilled(control) {
    var raw = getDecisionControlRawValue(control);
    return isTrackedDecisionValue(toCanonicalDecisionValue(raw));
  }

  function setDecisionControlEditable(control, isEditable) {
    if (!control) return;

    control.disabled = !isEditable;
    control.setAttribute("aria-disabled", isEditable ? "false" : "true");
    control.style.cursor = isEditable ? "pointer" : "not-allowed";

    if (!isEditable) closeDecisionPickerMenu();

    var hiddenSrc = control.__decisionHiddenSource;
    if (hiddenSrc) {
      hiddenSrc.disabled = !isEditable;
      hiddenSrc.setAttribute("aria-disabled", isEditable ? "false" : "true");
    }
  }

  function enforceProdejnaDependsOnSklad(panel) {
    var rows = getDataRows(panel);
    for (var r = 0; r < rows.length; r++) {
      var skladTarget = getDecisionTargetInRow(rows[r], TASKFORM_COL_STAV_SKLAD);
      var prodejnaTarget = getDecisionTargetInRow(rows[r], TASKFORM_COL_STAV_PRODEJNA);
      if (!prodejnaTarget) continue;

      var skladFilled = isDecisionControlFilled(skladTarget);
      setDecisionControlEditable(prodejnaTarget, skladFilled);

      if (!skladFilled) {
        prodejnaTarget.title = "Nejprve vyplňte vyjádření skladu.";
      } else {
        prodejnaTarget.removeAttribute("title");
      }
    }
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
    enforceProdejnaDependsOnSklad(panel);
    refreshTaskFormDecisionColors(panel);
    syncTocHtmlFromAdvancedTable(panel);

    // Delayed re-syncs catch TISA setting source values after initial DOM render
    var lateSync = 0;
    function lateDecisionSync() {
      lateSync++;
      ensureTaskFormDecisionSelects(panel);
      enforceProdejnaDependsOnSklad(panel);
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
        enforceProdejnaDependsOnSklad(panel);
        refreshTaskFormDecisionColors(panel);
        syncTocHtmlFromAdvancedTable(panel);
      });

      panel.__taskFormDecisionObserver.observe(panel, {
        childList: true,
        subtree: true
      });
    }

    var htmlSyncTimer = null;
    function scheduleTocHtmlSync() {
      if (htmlSyncTimer) clearTimeout(htmlSyncTimer);
      htmlSyncTimer = setTimeout(function() {
        htmlSyncTimer = null;
        syncTocHtmlFromAdvancedTable(panel);
      }, 80);
    }

    var komentRozdilCtrl = getWritableFieldControl("koment_rozdil");
    if (komentRozdilCtrl && !komentRozdilCtrl.__tocHtmlSyncBound) {
      komentRozdilCtrl.__tocHtmlSyncBound = true;
      komentRozdilCtrl.addEventListener("input", scheduleTocHtmlSync, true);
      komentRozdilCtrl.addEventListener("change", scheduleTocHtmlSync, true);
    }

    function getTrackedColumnIndex(target, tr) {
      var controls = getEditableAdvancedControlsInRow(tr);
      if (!controls || controls.length === 0) return -1;

      for (var i = 0; i < controls.length; i++) {
        if (controls[i] === target) return mapPhysicalToLogicalTaskFormColumnIndex(controls, i);
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

        // Pro typ "Rozdíl v transferu" nechceme při nesouhlasu skladu vyžadovat skutečný napočítaný stav.
        var skipActualStateModal =
          decisionValue === "rozdil_nepotvrzen" &&
          colIndex === TASKFORM_COL_STAV_SKLAD &&
          isRozdilVTransferuType();

        if (skipActualStateModal) {
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
      enforceProdejnaDependsOnSklad(panel);
      scheduleTocHtmlSync();
    }, true);

    panel.addEventListener("change", function(e) {
      var target = e.target;
      if (!target) return;

      var tr = target.closest ? target.closest("tr.dataRow") : null;
      if (!tr) return;

      handleDecisionInteraction(target, tr);
      enforceProdejnaDependsOnSklad(panel);
      scheduleTocHtmlSync();
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

  function showCenteredWarningModal(title, message) {
    var existing = document.getElementById("taskFormCenteredWarningBackdrop");
    if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

    var backdrop = document.createElement("div");
    backdrop.id = "taskFormCenteredWarningBackdrop";
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:10050;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:18px 20px;border-radius:8px;width:560px;max-width:94vw;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 10px 24px rgba(0,0,0,0.35);" +
      "border:2px solid #d83b01;";

    dialog.innerHTML =
      "<div style='display:flex;align-items:center;gap:8px;margin-bottom:8px;'>" +
      "  <span style='display:inline-block;width:10px;height:10px;border-radius:50%;background:#d83b01;'></span>" +
      "  <h2 style='margin:0;font-size:18px;color:#a4262c;line-height:1.2;'>" + escapeHtml(title || "Upozornění") + "</h2>" +
      "</div>" +
      "<div style='font-size:13px;color:#323130;line-height:1.45;white-space:pre-wrap;'>" + escapeHtml(message || "") + "</div>" +
      "<div style='display:flex;justify-content:flex-end;margin-top:14px;'>" +
      "  <button type='button' id='taskFormCenteredWarningOk' style='padding:8px 14px;border-radius:4px;border:1px solid #d83b01;background:#d83b01;color:#fff;cursor:pointer;font-weight:600;'>Rozumím</button>" +
      "</div>";

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    var close = function() {
      if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
    };

    var okBtn = dialog.querySelector("#taskFormCenteredWarningOk");
    if (okBtn) okBtn.onclick = close;

    backdrop.addEventListener("click", function(e) {
      if (e.target === backdrop) close();
    });
  }

  function isTaskDecisionControlEditable(ctrl) {
    if (!ctrl) return false;
    if (ctrl.disabled) return false;

    var ariaDisabled = (ctrl.getAttribute("aria-disabled") || "").toLowerCase();
    if (ariaDisabled === "true") return false;

    return true;
  }

  function detectTaskRoleLabelFromPanel(panel) {
    if (!panel) return "";

    var rows = getDataRows(panel);
    if (!rows || rows.length === 0) return "";

    var hasAnySklad = false;
    var hasAnyMissingSklad = false;

    for (var i = 0; i < rows.length; i++) {
      var skladTarget = getDecisionTargetInRow(rows[i], TASKFORM_COL_STAV_SKLAD);

      if (!skladTarget) continue;
      hasAnySklad = true;

      if (!isDecisionControlFilled(skladTarget)) {
        hasAnyMissingSklad = true;
        break;
      }
    }

    if (!hasAnySklad) return "";
    if (hasAnyMissingSklad) return "inventura_skladu";
    return "vyjadreni_prodejny";
  }

  function detectTaskLocationBranchFromPanel(panel) {
    if (!panel) return "";

    var rows = getDataRows(panel);
    var hasBramb = false;
    var hasCentr = false;

    for (var i = 0; i < rows.length; i++) {
      var controls = getEditableAdvancedControlsInRow(rows[i]);
      var locationControl = controls[COL_LOCATION];
      var locationText = locationControl
        ? trimSafe(locationControl.value || locationControl.textContent || "")
        : "";
      var key = getBaseLocationKey(locationText);

      if (key === "BRAMB") hasBramb = true;
      if (key === "CENTR") hasCentr = true;
    }

    if (hasBramb && !hasCentr) return "BRAMB";
    if (hasCentr && !hasBramb) return "CENTR";
    return "";
  }

  function collectGoodsForLocationBranch(panel, branch) {
    if (!panel || !branch) return [];

    var allowed = { OSTATNI: true };
    if (branch === "BRAMB") allowed.BRAMB = true;
    if (branch === "CENTR") allowed.CENTR = true;

    var rows = getDataRows(panel);
    var goods = [];
    var seen = {};

    for (var i = 0; i < rows.length; i++) {
      var controls = getEditableAdvancedControlsInRow(rows[i]);
      var goodsControl = controls[COL_GOODS];
      var locationControl = controls[COL_LOCATION];

      var goodsValue = goodsControl
        ? trimSafe(goodsControl.value || goodsControl.textContent || "")
        : "";
      var goodsNumber = stripLookupInfo(goodsValue);

      var locationText = locationControl
        ? trimSafe(locationControl.value || locationControl.textContent || "")
        : "";
      var locationKey = getBaseLocationKey(locationText);
      var assignedBranch = getOstatniAssignmentBranchFromRow(rows[i]);

      if (!allowed[locationKey]) continue;
      if (locationKey === "OSTATNI" && assignedBranch && assignedBranch !== branch) continue;
      if (!goodsNumber || seen[goodsNumber]) continue;

      seen[goodsNumber] = true;
      goods.push(goodsNumber);
    }

    return goods;
  }

  function collectInventoryRowsForLocationBranch(panel, branch) {
    if (!panel || !branch) return [];

    ensureTaskFormDecisionSelects(panel);

    var allowed = { OSTATNI: true };
    if (branch === "BRAMB") allowed.BRAMB = true;
    if (branch === "CENTR") allowed.CENTR = true;

    var rows = getDataRows(panel);
    var items = [];

    for (var i = 0; i < rows.length; i++) {
      var controls = getEditableAdvancedControlsInRow(rows[i]);
      var goodsControl = controls[COL_GOODS];
      var locationControl = controls[COL_LOCATION];

      var goodsValue = goodsControl
        ? trimSafe(goodsControl.value || goodsControl.textContent || "")
        : "";
      var goodsNumber = stripLookupInfo(goodsValue);

      var locationText = locationControl
        ? trimSafe(locationControl.value || locationControl.textContent || "")
        : "";
      var locationKey = getBaseLocationKey(locationText);
      var assignedBranch = getOstatniAssignmentBranchFromRow(rows[i]);
      var skladTarget = getDecisionTargetInRow(rows[i], TASKFORM_COL_STAV_SKLAD);

      if (!allowed[locationKey]) continue;
      if (locationKey === "OSTATNI" && assignedBranch && assignedBranch !== branch) continue;
      if (!goodsNumber) continue;
      if (isDecisionControlFilled(skladTarget)) continue;

      items.push({
        rowIndex: i,
        goodsNumber: goodsNumber,
        locationText: locationText,
        locationKey: locationKey,
        assignmentBranch: assignedBranch
      });
    }

    return items;
  }

  function applySkladDecisionsToRows(panel, decisions) {
    if (!panel || !Array.isArray(decisions) || decisions.length === 0) return false;

    ensureTaskFormDecisionSelects(panel);

    var rows = getDataRows(panel);
    var hadFailure = false;

    for (var i = 0; i < decisions.length; i++) {
      var decision = decisions[i] || {};
      var tr = rows[decision.rowIndex];
      var assignmentTarget = getDecisionTargetInRow(tr, TASKFORM_COL_OSTATNI_ASSIGN);
      var target = getDecisionTargetInRow(tr, TASKFORM_COL_STAV_SKLAD);
      if (!target && !assignmentTarget) {
        hadFailure = true;
        continue;
      }

      if (decision.assignmentLabel) {
        // V legacy struktuře nemusí existovat sloupec pro zařazení OSTATNÍ.
        // V takovém případě nepovažujeme zápis assignmentu za chybu celého uložení.
        if (assignmentTarget) {
          var assignmentOk = setOstatniAssignmentControlValue(assignmentTarget, decision.assignmentLabel, true);
          if (!assignmentOk) hadFailure = true;
        }
      }

      if (!decision.label) {
        continue;
      }

      if (!target) {
        hadFailure = true;
        continue;
      }

      setDecisionControlValue(target, decision.label, false);

      var canonical = toCanonicalDecisionValue(decision.label);
      target.dataset.lastCommittedDecision = canonical;
      target.dataset.lastDecisionValue = decision.label;
      setDecisionVisualState(target);

      var hiddenSrc = target.__decisionHiddenSource;
      if (hiddenSrc) {
        hiddenSrc.dataset.lastCommittedDecision = canonical;
        hiddenSrc.dataset.lastDecisionValue = decision.label;
        // TISA serializuje hodnoty ze zdrojového controlu až po změnových eventech.
        fireInputEvents(hiddenSrc);
      }

      updateStatusByLocationDecision(decision.locationText, canonical);
    }

    enforceProdejnaDependsOnSklad(panel);
    refreshTaskFormDecisionColors(panel);
    syncTocHtmlFromAdvancedTable(panel);
    return !hadFailure;
  }

  function showCenteredInfoModal(title, message) {
    var existing = document.getElementById("taskFormCenteredInfoBackdrop");
    if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

    var backdrop = document.createElement("div");
    backdrop.id = "taskFormCenteredInfoBackdrop";
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:10040;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:18px 20px;border-radius:8px;width:560px;max-width:94vw;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 10px 24px rgba(0,0,0,0.35);" +
      "border:2px solid #0078d4;";

    dialog.innerHTML =
      "<div style='display:flex;align-items:center;gap:8px;margin-bottom:8px;'>" +
      "  <span style='display:inline-block;width:10px;height:10px;border-radius:50%;background:#0078d4;'></span>" +
      "  <h2 style='margin:0;font-size:18px;color:#004578;line-height:1.2;'>" + escapeHtml(title || "Informace") + "</h2>" +
      "</div>" +
      "<div style='font-size:13px;color:#323130;line-height:1.45;white-space:pre-wrap;'>" + escapeHtml(message || "") + "</div>" +
      "<div style='display:flex;justify-content:flex-end;margin-top:14px;'>" +
      "  <button type='button' id='taskFormCenteredInfoOk' style='padding:8px 14px;border-radius:4px;border:1px solid #0078d4;background:#0078d4;color:#fff;cursor:pointer;font-weight:600;'>Rozumím</button>" +
      "</div>";

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    var close = function() {
      if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
    };

    var okBtn = dialog.querySelector("#taskFormCenteredInfoOk");
    if (okBtn) okBtn.onclick = close;

    backdrop.addEventListener("click", function(e) {
      if (e.target === backdrop) close();
    });
  }

  function parseDatabaseCsvRows(csvText) {
    var normalizedCsv = trimSafe(csvText || "");
    if (!normalizedCsv) return [];

    var lines = normalizedCsv.split(";");
    var rows = [];
    for (var i = 0; i < lines.length; i++) {
      var rowText = trimSafe(lines[i]);
      if (!rowText) continue;

      var rawCells = rowText.split(",");
      var cells = [];
      for (var c = 0; c < rawCells.length; c++) {
        cells.push(trimSafe(rawCells[c] || ""));
      }

      rows.push(cells);
    }

    return rows;
  }

  function normalizeDsuGoodsNumber(value) {
    return trimSafe(value).toUpperCase();
  }

  function getSingleTransferNumberFromDatabase(csvText) {
    var rows = parseDatabaseCsvRows(csvText);
    if (!rows.length) {
      return {
        ok: false,
        message: "DSU neobsahuje žádná data."
      };
    }

    var transferNumber = "";

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i] || [];
      if (row.length < 5) {
        return {
          ok: false,
          message: "Řádek " + (i + 1) + " v DSU není kompletní (očekáváno 5 sloupců)."
        };
      }

      for (var c = 0; c < 5; c++) {
        if (!trimSafe(row[c] || "")) {
          return {
            ok: false,
            message: "Řádek " + (i + 1) + " v DSU není kompletní (sloupec " + (c + 1) + " je prázdný)."
          };
        }
      }

      var current = trimSafe(row[2] || "");
      if (!current) {
        return {
          ok: false,
          message: "V DSU chybí číslo transferu na řádku " + (i + 1) + "."
        };
      }

      if (!transferNumber) {
        transferNumber = current;
        continue;
      }

      if (current !== transferNumber) {
        return {
          ok: false,
          message: "V DSU je více čísel transferu (např. " + transferNumber + " a " + current + ")."
        };
      }
    }

    return {
      ok: true,
      value: transferNumber
    };
  }

  function ensureDatabasePreviewState() {
    if (!window.__databasePreviewState) {
      window.__databasePreviewState = {
        hidden: true,
        top: 96,
        left: null,
        right: 14
      };
    }

    return window.__databasePreviewState;
  }

  function clampDatabasePreviewPosition(state, panelWidth, panelHeight) {
    var viewportWidth = Math.max(document.documentElement.clientWidth || 0, window.innerWidth || 0);
    var viewportHeight = Math.max(document.documentElement.clientHeight || 0, window.innerHeight || 0);
    var margin = 8;

    if (typeof state.top !== "number" || isNaN(state.top)) state.top = 96;
    state.top = Math.min(Math.max(state.top, margin), Math.max(margin, viewportHeight - panelHeight - margin));

    if (typeof state.left === "number" && !isNaN(state.left)) {
      state.left = Math.min(Math.max(state.left, margin), Math.max(margin, viewportWidth - panelWidth - margin));
      state.right = null;
      return;
    }

    if (typeof state.right !== "number" || isNaN(state.right)) state.right = 14;
    state.right = Math.min(Math.max(state.right, margin), Math.max(margin, viewportWidth - panelWidth - margin));
  }

  function positionDatabasePreviewPanel(panel) {
    if (!panel) return;

    var state = ensureDatabasePreviewState();
    var panelWidth = panel.offsetWidth || 520;
    var panelHeight = panel.offsetHeight || 320;

    clampDatabasePreviewPosition(state, panelWidth, panelHeight);

    panel.style.top = state.top + "px";
    if (typeof state.left === "number" && !isNaN(state.left)) {
      panel.style.left = state.left + "px";
      panel.style.right = "auto";
    } else {
      panel.style.left = "auto";
      panel.style.right = state.right + "px";
    }
  }

  function ensureDatabasePreviewToggleButton() {
    var state = ensureDatabasePreviewState();
    var btn = document.getElementById("databaseTablePreviewToggle");

    if (state.hidden) {
      if (!btn) {
        btn = document.createElement("button");
        btn.id = "databaseTablePreviewToggle";
        btn.type = "button";
        btn.style.cssText =
          "position:fixed;right:14px;top:96px;z-index:10060;padding:8px 12px;border-radius:999px;" +
          "border:1px solid #0078d4;background:#0078d4;color:#fff;font-family:Segoe UI, sans-serif;" +
          "font-size:12px;cursor:pointer;box-shadow:0 6px 16px rgba(0,0,0,0.22);";
        btn.textContent = "Zobrazit DSU";
        btn.onclick = function() {
          ensureDatabasePreviewState().hidden = false;
          var panel = document.getElementById("databaseTablePreviewPanel");
          if (panel) panel.style.display = "block";
          ensureDatabasePreviewToggleButton();
        };
        document.body.appendChild(btn);
      }
      return btn;
    }

    if (btn && btn.parentNode) btn.parentNode.removeChild(btn);
    return null;
  }

  function enableDatabasePreviewDragging(panel, dragHandle) {
    if (!panel || !dragHandle) return;

    if (panel.__dragHandleBound && panel.__dragHandleBound !== dragHandle && panel.__dragMouseDownHandler) {
      panel.__dragHandleBound.removeEventListener("mousedown", panel.__dragMouseDownHandler, false);
    }

    if (panel.__dragHandleBound === dragHandle && panel.__dragMouseDownHandler) return;

    dragHandle.style.cursor = "move";

    var onMouseDown = function(e) {
      if (!e) return;
      if (typeof e.button === "number" && e.button !== 0 && e.button !== 1) return;
      if (e.target && e.target.closest && e.target.closest("button")) return;

      var state = ensureDatabasePreviewState();
      var rect = panel.getBoundingClientRect();
      var startX = e.clientX;
      var startY = e.clientY;
      var startLeft = rect.left;
      var startTop = rect.top;

      function onMove(moveEvent) {
        var nextLeft = startLeft + (moveEvent.clientX - startX);
        var nextTop = startTop + (moveEvent.clientY - startY);
        state.left = nextLeft;
        state.right = null;
        state.top = nextTop;
        positionDatabasePreviewPanel(panel);
      }

      function onUp() {
        document.removeEventListener("mousemove", onMove, true);
        document.removeEventListener("mouseup", onUp, true);
      }

      document.addEventListener("mousemove", onMove, true);
      document.addEventListener("mouseup", onUp, true);
      e.preventDefault();
    };

    dragHandle.addEventListener("mousedown", onMouseDown, false);
    panel.__dragHandleBound = dragHandle;
    panel.__dragMouseDownHandler = onMouseDown;
  }

  function renderDatabaseTablePreview(csvText) {
    var state = ensureDatabasePreviewState();
    var rows = parseDatabaseCsvRows(csvText);
    if (!rows.length) {
      var existingPanel = document.getElementById("databaseTablePreviewPanel");
      if (existingPanel && existingPanel.parentNode) existingPanel.parentNode.removeChild(existingPanel);
      var existingToggle = document.getElementById("databaseTablePreviewToggle");
      if (existingToggle && existingToggle.parentNode) existingToggle.parentNode.removeChild(existingToggle);
      return;
    }

    var headers = [
      "Číslo dodávky transferu",
      "Číslo dodávky ze skladu",
      "Číslo transferu",
      "Číslo zboží",
      "Počet"
    ];

    var tableRows = [];
    for (var i = 0; i < rows.length; i++) {
      var cells = rows[i] || [];
      var rowGoods = normalizeDsuGoodsNumber(cells[3] || "");
      var cellsHtml = "";
      for (var j = 0; j < headers.length; j++) {
        cellsHtml += "<td style='padding:5px 10px;border:1px solid #c8c6c4;white-space:pre-wrap;'>" + escapeHtml(trimSafe(cells[j] || "")) + "</td>";
      }
      tableRows.push("<tr data-dsu-row='1' data-dsu-goods='" + escapeHtml(rowGoods) + "'>" + cellsHtml + "</tr>");
    }

    var headerHtml = "<tr>";
    for (var h = 0; h < headers.length; h++) {
      headerHtml += "<th style='padding:7px 10px;border:1px solid #c8c6c4;background:#f3f2f1;text-align:left;'>" + escapeHtml(headers[h]) + "</th>";
    }
    headerHtml += "</tr>";

    var tableHtml = tableRows.length
      ? "<table style='border-collapse:collapse;font-size:13px;width:100%;'><thead>" + headerHtml + "</thead><tbody>" + tableRows.join("") + "</tbody></table>"
      : "<em style='color:#888;'>Žádná data.</em>";

    var panel = document.getElementById("databaseTablePreviewPanel");
    if (!panel) {
      panel = document.createElement("aside");
      panel.id = "databaseTablePreviewPanel";
      panel.style.cssText =
        "position:fixed;right:14px;top:96px;width:520px;max-width:42vw;max-height:78vh;overflow:auto;" +
        "background:#fff;border:2px solid #0078d4;border-radius:8px;padding:14px;" +
        "font-family:Segoe UI, sans-serif;box-shadow:0 10px 24px rgba(0,0,0,0.25);z-index:10030;";
      document.body.appendChild(panel);
    }

    panel.innerHTML =
      "<div id='databaseTablePreviewHeader' style='display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:10px;user-select:none;'>" +
      "  <div style='display:flex;align-items:center;gap:8px;min-width:0;'>" +
      "    <span style='display:inline-block;width:10px;height:10px;border-radius:50%;background:#0078d4;flex:0 0 auto;'></span>" +
      "    <h2 style='margin:0;font-size:16px;color:#004578;line-height:1.2;'>Náhled na DSU</h2>" +
      "  </div>" +
      "  <div style='display:flex;align-items:center;gap:8px;flex:0 0 auto;'>" +
      "    <span style='font-size:11px;color:#605e5c;white-space:nowrap;'>Přetáhni panel</span>" +
      "    <button type='button' id='databaseTablePreviewHide' style='padding:4px 10px;border-radius:999px;border:1px solid #0078d4;background:#fff;color:#0078d4;font-size:12px;cursor:pointer;'>Skrýt</button>" +
      "  </div>" +
      "</div>" +
      "<div style='font-size:12px;color:#605e5c;margin-bottom:8px;'>Oddělovače: sloupce čárka, řádky středník.</div>" +
      "<div style='overflow-x:auto;'>" + tableHtml + "</div>";

    panel.style.display = state.hidden ? "none" : "block";
    positionDatabasePreviewPanel(panel);
    enableDatabasePreviewDragging(panel, panel.querySelector("#databaseTablePreviewHeader"));

    var hideBtn = panel.querySelector("#databaseTablePreviewHide");
    if (hideBtn && !hideBtn.__hideBound) {
      hideBtn.__hideBound = true;
      hideBtn.onclick = function() {
        ensureDatabasePreviewState().hidden = true;
        panel.style.display = "none";
        ensureDatabasePreviewToggleButton();
      };
    }

    ensureDatabasePreviewToggleButton();
  }

  function initCurrentUserCenterForNewForm(attempt) {
    attempt = attempt || 0;
    if (window.__newFormCenterInitStarted) return;

    exposeNewFormLookupDebug();

    var ctrl = getWritableFieldControl("identifikace_prodejny_2");
    if (!ctrl) {
      if (attempt < 25) {
        setTimeout(function() {
          initCurrentUserCenterForNewForm(attempt + 1);
        }, 250);
      }
      return;
    }

    var existingValue = trimSafe(ctrl.value || "");
    if (existingValue) return;

    window.__newFormCenterInitStarted = true;

    resolveUserLocationBranch(function(resolveErr, userInfo) {
      window.__newFormCenterInitStarted = false;

      if (resolveErr) {
        debugLog("[newform_center] nepodarilo se dohledat stredisko", resolveErr);
        return;
      }

      var centerValue = trimSafe(userInfo && userInfo.center ? userInfo.center : "");
      if (!centerValue) {
        debugLog("[newform_center] uzivatel nema dohledane stredisko", userInfo);
        return;
      }

      var currentCtrl = getWritableFieldControl("identifikace_prodejny_2");
      var hiddenBefore = getHiddenLookupControl("identifikace_prodejny_2");
      var hiddenBeforeValue = hiddenBefore ? trimSafe(hiddenBefore.value || "") : "";
      if (!currentCtrl) return;
      if (trimSafe(currentCtrl.value || "")) return;

      var setOk = setFieldValueSmart("identifikace_prodejny_2", centerValue);
      debugLog("[newform_center] stredisko zapsano do identifikace_prodejny_2", centerValue, "setOk", setOk);
      triggerLookupFieldBehavior(currentCtrl, centerValue);

      setTimeout(function() {
        var hiddenAfter = getHiddenLookupControl("identifikace_prodejny_2");
        var hiddenAfterValue = hiddenAfter ? trimSafe(hiddenAfter.value || "") : "";

        debugLog(
          "[newform_center] hidden lookup state after trigger",
          "before",
          hiddenBeforeValue,
          "after",
          hiddenAfterValue,
          "jQuery",
          typeof window.jQuery !== "undefined",
          "uiAutocomplete",
          !!(window.jQuery && window.jQuery.ui && window.jQuery.ui.autocomplete)
        );

        if (hiddenAfterValue === hiddenBeforeValue) {
          trySelectFirstAutocompleteSuggestion("identifikace_prodejny_2", centerValue, 0);
        }
      }, 450);
    });
  }

  function showNewFormDeliveryNoteModal(options) {
    options = options || {};

    var existing = document.getElementById("newFormDeliveryNoteBackdrop");
    if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

    var backdrop = document.createElement("div");
    backdrop.id = "newFormDeliveryNoteBackdrop";
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:10070;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:18px 20px;border-radius:8px;width:520px;max-width:94vw;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 10px 24px rgba(0,0,0,0.35);" +
      "border:2px solid #0078d4;";

    dialog.innerHTML =
      "<h2 style='margin:0 0 8px 0;font-size:18px;color:#004578;'>Zadejte číslo dodacího listu (DSU)</h2>" +
      "<div style='font-size:13px;color:#323130;line-height:1.45;margin-bottom:10px;'>" +
      "Než budete pokračovat, doplňte číslo dodacího listu. Hodnota se zapíše do pole cislo_dl a spustí se lookup." +
      "</div>" +
      "<label style='display:block;font-size:12px;font-weight:600;margin-bottom:4px;'>Číslo dodacího listu</label>" +
      "<input type='text' id='newFormDeliveryNoteValue' style='width:100%;box-sizing:border-box;padding:8px;border:1px solid #c8c6c4;border-radius:4px;'>" +
      "<div id='newFormDeliveryNoteErr' style='display:none;margin-top:8px;color:#a80000;font-size:12px;'></div>" +
      "<div style='display:flex;justify-content:flex-end;margin-top:14px;'>" +
      "  <button type='button' id='newFormDeliveryNoteSave' style='padding:8px 12px;border-radius:4px;border:1px solid #0078d4;background:#0078d4;color:#fff;cursor:pointer;font-weight:600;'>Potvrdit</button>" +
      "</div>";

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    var valueInput = dialog.querySelector("#newFormDeliveryNoteValue");
    var err = dialog.querySelector("#newFormDeliveryNoteErr");
    var saveBtn = dialog.querySelector("#newFormDeliveryNoteSave");

    saveBtn.onclick = function() {
      var deliveryNumber = trimSafe(valueInput.value || "");
      if (!deliveryNumber) {
        err.textContent = "Zadejte číslo dodacího listu.";
        err.style.display = "block";
        valueInput.focus();
        return;
      }

      err.style.display = "none";

      var ctrl = getWritableFieldControl("cislo_dl");
      var hiddenBefore = getHiddenLookupControl("cislo_dl");
      var hiddenBeforeValue = hiddenBefore ? trimSafe(hiddenBefore.value || "") : "";
      if (!ctrl) {
        err.textContent = "Nepodařilo se najít pole cislo_dl na formuláři.";
        err.style.display = "block";
        return;
      }

      var setOk = setFieldValueSmart("cislo_dl", deliveryNumber);
      if (!setOk) {
        err.textContent = "Nepodařilo se zapsat hodnotu do pole cislo_dl.";
        err.style.display = "block";
        return;
      }

      triggerLookupFieldBehavior(ctrl, deliveryNumber);
      window.__newFormDeliveryNoteConfirmed = true;
      window.__pendingContinueAfterDeliveryNote = true;
      window.__newFormDeliveryLookupFailed = false;

      if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
      showNewFormDeliveryNoteLoadingModal();

      ensureDeliveryLookupAndWaitDatabase(deliveryNumber, {
        onSuccess: function() {
          var hiddenAfter = getHiddenLookupControl("cislo_dl");
          var hiddenAfterValue = hiddenAfter ? trimSafe(hiddenAfter.value || "") : "";

          debugLog(
            "[newform_delivery_note] lookup+database ready",
            "beforeHidden",
            hiddenBeforeValue,
            "afterHidden",
            hiddenAfterValue,
            "jQuery",
            typeof window.jQuery !== "undefined",
            "uiAutocomplete",
            !!(window.jQuery && window.jQuery.ui && window.jQuery.ui.autocomplete)
          );

          if (typeof options.onSaved === "function") options.onSaved(deliveryNumber);
        },
        onTimeout: function() {
          showCenteredWarningModal(
            "Lookup DSU se nespustil",
            "Pole database se po zadání DSU nenaplnilo. Zkuste znovu potvrdit číslo DSU nebo ručně kliknout do pole cislo_dl a vybrat hodnotu z našeptávače."
          );
        }
      });
    };

    setTimeout(function() {
      try { valueInput.focus(); } catch (e) {}
    }, 0);
  }

  function showNewFormDeliveryNoteLoadingModal() {
    var existingPrompt = document.getElementById("newFormDeliveryNoteBackdrop");
    if (existingPrompt && existingPrompt.parentNode) return existingPrompt;

    var existing = document.getElementById("newFormDeliveryNoteLoadingBackdrop");
    if (existing && existing.parentNode) return existing;

    var backdrop = document.createElement("div");
    backdrop.id = "newFormDeliveryNoteLoadingBackdrop";
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:10065;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:18px 20px;border-radius:8px;width:420px;max-width:92vw;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 10px 24px rgba(0,0,0,0.35);" +
      "border:2px solid #0078d4;text-align:center;";

    dialog.innerHTML =
      "<div style='display:flex;align-items:center;justify-content:center;gap:10px;margin-bottom:10px;'>" +
      "  <span style='width:14px;height:14px;border:2px solid #c7e0f4;border-top-color:#0078d4;border-radius:50%;display:inline-block;animation:newFormDeliverySpin 0.8s linear infinite;'></span>" +
      "  <strong style='font-size:16px;color:#004578;'>Probíhá načítání DSU dat</strong>" +
      "</div>" +
      "<div style='font-size:13px;color:#323130;line-height:1.45;'>Čekám na vyplnění pole <strong>database</strong> z lookupu podle zadaného čísla DSU.</div>";

    if (!document.getElementById("newFormDeliveryNoteLoadingStyle")) {
      var style = document.createElement("style");
      style.id = "newFormDeliveryNoteLoadingStyle";
      style.textContent = "@keyframes newFormDeliverySpin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }";
      document.head.appendChild(style);
    }

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);
    return backdrop;
  }

  function hideNewFormDeliveryNoteLoadingModal() {
    var existing = document.getElementById("newFormDeliveryNoteLoadingBackdrop");
    if (existing && existing.parentNode) existing.parentNode.removeChild(existing);
  }

  function clearNewFormDeliveryLookupWaiter() {
    if (window.__newFormDeliveryLookupWaitTimer) {
      clearInterval(window.__newFormDeliveryLookupWaitTimer);
      window.__newFormDeliveryLookupWaitTimer = null;
    }
  }

  function ensureDeliveryLookupAndWaitDatabase(deliveryNumber, options) {
    options = options || {};

    clearNewFormDeliveryLookupWaiter();

    var ctrl = getWritableFieldControl("cislo_dl");
    var databaseCtrl = getWritableFieldControl("database");
    var initialDatabaseValue = trimSafe(databaseCtrl && databaseCtrl.value ? databaseCtrl.value : "");
    var startedAt = Date.now();
    var maxWaitMs = typeof options.maxWaitMs === "number" ? options.maxWaitMs : 18000;
    var attempt = 0;

    if (!ctrl) {
      hideNewFormDeliveryNoteLoadingModal();
      if (typeof options.onTimeout === "function") options.onTimeout();
      return;
    }

    function succeed(databaseValue) {
      clearNewFormDeliveryLookupWaiter();
      hideNewFormDeliveryNoteLoadingModal();
      window.__newFormDeliveryLookupFailed = false;
      if (typeof options.onSuccess === "function") options.onSuccess(databaseValue || "");
    }

    function fail() {
      clearNewFormDeliveryLookupWaiter();
      hideNewFormDeliveryNoteLoadingModal();
      window.__pendingContinueAfterDeliveryNote = false;
      window.__newFormDeliveryLookupFailed = true;
      if (typeof options.onTimeout === "function") options.onTimeout();
    }

    function tick() {
      attempt++;

      var dbCtrlNow = getWritableFieldControl("database") || databaseCtrl;
      var currentDatabaseValue = trimSafe(dbCtrlNow && dbCtrlNow.value ? dbCtrlNow.value : "");
      if (currentDatabaseValue && (currentDatabaseValue !== initialDatabaseValue || !initialDatabaseValue)) {
        succeed(currentDatabaseValue);
        return;
      }

      if (Date.now() - startedAt >= maxWaitMs) {
        fail();
        return;
      }

      // Lookup in this form is flaky; periodically retrigger and try selecting first suggestion.
      if (attempt % 2 === 1) {
        setFieldValueSmart("cislo_dl", deliveryNumber);
        triggerLookupFieldBehavior(ctrl, deliveryNumber);
      }

      if (attempt % 3 === 0) {
        trySelectFirstAutocompleteSuggestion("cislo_dl", deliveryNumber, 0);
      }
    }

    tick();
    window.__newFormDeliveryLookupWaitTimer = setInterval(tick, 350);
  }

  function initNewFormDeliveryNoteModal(attempt) {
    attempt = attempt || 0;

    var ctrl = getWritableFieldControl("cislo_dl");
    if (!ctrl) {
      if (!window.__newFormDeliveryNoteFieldObserver && document.body) {
        window.__newFormDeliveryNoteFieldObserver = new MutationObserver(function() {
          if (!getWritableFieldControl("cislo_dl")) return;

          try {
            window.__newFormDeliveryNoteFieldObserver.disconnect();
          } catch (observerError) {}
          window.__newFormDeliveryNoteFieldObserver = null;
          initNewFormDeliveryNoteModal(0);
        });

        window.__newFormDeliveryNoteFieldObserver.observe(document.body, {
          childList: true,
          subtree: true
        });
      }

      if (attempt < 60) {
        setTimeout(function() {
          initNewFormDeliveryNoteModal(attempt + 1);
        }, 150);
      }
      return;
    }

    if (window.__newFormDeliveryNoteFieldObserver) {
      try {
        window.__newFormDeliveryNoteFieldObserver.disconnect();
      } catch (observerDisconnectError) {}
      window.__newFormDeliveryNoteFieldObserver = null;
    }

    if (trimSafe(ctrl.value || "")) {
      window.__newFormDeliveryNoteConfirmed = true;
      return;
    }

    if (window.__newFormDeliveryNoteConfirmed) return;
    if (document.getElementById("newFormDeliveryNoteBackdrop")) return;

    showNewFormDeliveryNoteModal({
      onSaved: function() {
        var databaseCtrl = getWritableFieldControl("database");
        if (!databaseCtrl) return;

        if (typeof databaseCtrl.__syncPreview === "function") databaseCtrl.__syncPreview();
        if (typeof databaseCtrl.__tryContinueAfterDsu === "function") {
          databaseCtrl.__tryContinueAfterDsu(true);
          return;
        }

        fireHtmlEvent(databaseCtrl, "change");
      }
    });
  }

  function initDatabaseField(attempt) {
    attempt = attempt || 0;
    var ctrl = getWritableFieldControl("database");
    if (!ctrl) {
      if (attempt < 25) {
        setTimeout(function() {
          initDatabaseField(attempt + 1);
        }, 250);
      }
      return;
    }
    if (ctrl.__databasePreviewBound) return;
    ctrl.__databasePreviewBound = true;

    function ensureTocIsReady() {
      var tocRow = getTisaRowByInternalName(TOC_FIELD);
      if (tocRow) return true;

      var hidden = document.querySelector("input[id$='_FieldName'][value='" + TOC_FIELD + "']");
      return !!hidden;
    }

    function tryContinueAfterDsu(showWarnings) {
      if (window.__typeModalShownAfterDsu) return true;

      if (!window.__newFormDeliveryNoteConfirmed) {
        hideNewFormDeliveryNoteLoadingModal();
        initNewFormDeliveryNoteModal(0);
        if (showWarnings) {
          showCenteredWarningModal(
            "Nejprve zadejte číslo dodacího listu",
            "Pro pokračování nejdříve potvrďte číslo DSU do pole cislo_dl."
          );
        }
        return false;
      }

      var databaseValue = trimSafe(ctrl.value || "");
      if (!databaseValue) {
        if (window.__newFormDeliveryLookupFailed) {
          hideNewFormDeliveryNoteLoadingModal();
          return false;
        }
        showNewFormDeliveryNoteLoadingModal();
        return false;
      }

      hideNewFormDeliveryNoteLoadingModal();

      renderDatabaseTablePreview(databaseValue);

      if (!ensureTocIsReady()) {
        if (showWarnings) {
          showCenteredWarningModal(
            "TOC není připraven",
            "Nepodařilo se najít pole rozdil_v_toc (TOC)."
          );
        }
        return false;
      }

      var transferCheck = getSingleTransferNumberFromDatabase(databaseValue);
      if (!transferCheck.ok) {
        if (showWarnings) {
          showCenteredWarningModal(
            "Neplatné DSU",
            transferCheck.message || "Nepodařilo se určit jediné číslo transferu z DSU."
          );
        }
        return false;
      }

      var transferSetOk = setFieldValueSmart("cislo_transferu", transferCheck.value || "");
      if (!transferSetOk) {
        if (showWarnings) {
          showCenteredWarningModal(
            "Nelze doplnit číslo transferu",
            "Nepodařilo se vyplnit pole cislo_transferu hodnotou " + (transferCheck.value || "") + "."
          );
        }
        return false;
      }

      window.__typeModalShownAfterDsu = true;
      window.__pendingContinueAfterDeliveryNote = false;
      hideNewFormDeliveryNoteLoadingModal();
      createTypeModal();
      return true;
    }

    function syncPreview() {
      renderDatabaseTablePreview(ctrl.value || "");
    }

    var lastValue = null;
    function syncIfChanged() {
      var currentValue = ctrl.value || "";
      if (currentValue === lastValue) return;
      lastValue = currentValue;
      renderDatabaseTablePreview(currentValue);
      tryContinueAfterDsu(false);
    }

    ctrl.addEventListener("input", syncPreview);
    ctrl.addEventListener("change", function() {
      syncPreview();
      tryContinueAfterDsu(true);
    });
    ctrl.addEventListener("blur", function() {
      syncPreview();
      tryContinueAfterDsu(true);
    });
    ctrl.__syncPreview = syncPreview;
    ctrl.__tryContinueAfterDsu = tryContinueAfterDsu;

    if (window.__pendingContinueAfterDeliveryNote) {
      setTimeout(function() {
        tryContinueAfterDsu(true);
      }, 0);
    }

    ctrl.__databasePreviewInterval = setInterval(syncIfChanged, 250);
    syncIfChanged();
  }

  function showTaskInventoryAssignmentModal(options) {
    options = options || {};

    var existing = document.getElementById("taskInventoryAssignmentBackdrop");
    if (existing && existing.parentNode) existing.parentNode.removeChild(existing);

    var items = Array.isArray(options.items) ? options.items : [];
    var backdrop = document.createElement("div");
    backdrop.id = "taskInventoryAssignmentBackdrop";
    backdrop.style.cssText =
      "position:fixed;top:0;left:0;width:100%;height:100%;background:rgba(0,0,0,0.35);" +
      "z-index:10045;display:flex;align-items:center;justify-content:center;";

    var dialog = document.createElement("div");
    dialog.style.cssText =
      "background:#fff;padding:18px 20px;border-radius:8px;width:760px;max-width:96vw;" +
      "max-height:88vh;overflow:auto;font-family:Segoe UI, sans-serif;" +
      "box-shadow:0 10px 24px rgba(0,0,0,0.35);border:2px solid #0078d4;";

    var rowsHtml = [];
    for (var i = 0; i < items.length; i++) {
      var item = items[i] || {};
      var rowName = "inventoryDecision_" + item.rowIndex;

      rowsHtml.push(
        "<tr>" +
        "<td style='border:1px solid #d0d0d0;padding:8px;'>" + escapeHtml(item.goodsNumber || "") + "</td>" +
        "<td style='border:1px solid #d0d0d0;padding:8px;'>" + escapeHtml(item.locationKey || item.locationText || "") + "</td>" +
        "<td style='border:1px solid #d0d0d0;padding:8px;white-space:nowrap;'>" +
        "<label style='display:inline-flex;align-items:center;gap:6px;margin-right:14px;cursor:pointer;'><input type='radio' name='" + rowName + "' value='souhlas'>Souhlas</label>" +
        "<label style='display:inline-flex;align-items:center;gap:6px;cursor:pointer;'><input type='radio' name='" + rowName + "' value='nesouhlas'>Nesouhlas</label>" +
        "</td>" +
        "</tr>"
      );
    }

    dialog.innerHTML =
      "<div style='display:flex;align-items:flex-start;justify-content:space-between;gap:12px;margin-bottom:12px;'>" +
      "  <div>" +
      "    <h2 style='margin:0 0 6px 0;font-size:18px;color:#004578;'>Inventura skladu</h2>" +
      "    <div style='font-size:13px;color:#323130;line-height:1.45;'>Vyjádři se pouze ke zboží pro tvoje středisko. Uživatel: <strong>" + escapeHtml(options.displayName || "") + "</strong>, email: <strong>" + escapeHtml(options.email || "") + "</strong>, středisko: <strong>" + escapeHtml(options.center || "") + "</strong>.</div>" +
      "  </div>" +
      "</div>" +
      "<div style='font-size:12px;color:#605e5c;margin-bottom:10px;'>Zobrazené zboží: <strong>" + escapeHtml(options.branch || "") + "</strong> + <strong>OSTATNÍ</strong>.</div>" +
      "<div style='border:1px solid #d0d0d0;border-radius:6px;overflow:hidden;'>" +
      "  <table style='width:100%;border-collapse:collapse;font-size:13px;'>" +
      "    <thead><tr style='background:#f3f2f1;'><th style='border:1px solid #d0d0d0;padding:8px;text-align:left;'>Číslo zboží</th><th style='border:1px solid #d0d0d0;padding:8px;text-align:left;'>Lokace</th><th style='border:1px solid #d0d0d0;padding:8px;text-align:left;'>Vyjádření skladu</th></tr></thead>" +
      "    <tbody>" + rowsHtml.join("") + "</tbody>" +
      "  </table>" +
      "</div>" +
      "<div id='taskInventoryAssignmentErr' style='display:none;margin-top:10px;color:#a80000;font-size:12px;'></div>" +
      "<div style='display:flex;justify-content:flex-end;gap:8px;margin-top:14px;'>" +
      "  <button type='button' id='taskInventoryAssignmentCancel' style='padding:8px 12px;border-radius:4px;border:1px solid #8a8886;background:#fff;cursor:pointer;'>Zrušit</button>" +
      "  <button type='button' id='taskInventoryAssignmentSave' style='padding:8px 12px;border-radius:4px;border:1px solid #0078d4;background:#0078d4;color:#fff;cursor:pointer;font-weight:600;'>Zapsat do inventury</button>" +
      "</div>";

    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    function close() {
      if (backdrop.parentNode) backdrop.parentNode.removeChild(backdrop);
    }

    dialog.querySelector("#taskInventoryAssignmentCancel").onclick = close;

    dialog.querySelector("#taskInventoryAssignmentSave").onclick = function() {
      var err = dialog.querySelector("#taskInventoryAssignmentErr");
      var decisions = [];

      for (var i = 0; i < items.length; i++) {
        var item = items[i] || {};
        var selected = dialog.querySelector("input[name='inventoryDecision_" + item.rowIndex + "']:checked");
        if (!selected) continue;

        decisions.push({
          rowIndex: item.rowIndex,
          goodsNumber: item.goodsNumber,
          locationText: item.locationText,
          label: selected ? (selected.value === "souhlas" ? "Rozdíl potvrzen skladem" : "Rozdíl skladem nepotvrzen") : ""
        });
      }

      if (decisions.length === 0) {
        err.textContent = "Vyberte alespoň jedno zboží k uložení.";
        err.style.display = "block";
        return;
      }

      err.style.display = "none";

      var ok = true;
      if (typeof options.onSave === "function") ok = options.onSave(decisions);
      if (ok === false) {
        err.textContent = "Nepodařilo se zapsat vyjádření do tabulky inventury.";
        err.style.display = "block";
        return;
      }

      close();
    };

    backdrop.addEventListener("click", function(e) {
      if (e.target === backdrop) close();
    });
  }

  function showTaskRoleModalOnTaskForm(attempt) {
    attempt = attempt || 0;
    if (window.__taskRoleIntroShown || window.__taskRoleIntroPending) return;

    var panel = getAdvancedTablePanel();
    var role = detectTaskRoleLabelFromPanel(panel);

    if (!role) {
      if (attempt < 20) {
        setTimeout(function() {
          showTaskRoleModalOnTaskForm(attempt + 1);
        }, 300);
      }
      return;
    }

    window.__taskRoleIntroPending = true;

    resolveUserLocationBranch(function(resolveErr, userInfo) {
      window.__taskRoleIntroPending = false;
      window.__taskRoleIntroShown = true;

      var userBranch = userInfo && userInfo.branch ? userInfo.branch : "";
      var detectedBranch = detectTaskLocationBranchFromPanel(panel);
      var effectiveBranch = userBranch || detectedBranch;
      var goodsForUser = collectGoodsForLocationBranch(panel, effectiveBranch);

      if (role === "inventura_skladu") {
        if (resolveErr && !effectiveBranch) {
          showCenteredWarningModal(
            "Nepodařilo se rozpoznat středisko uživatele",
            resolveErr.message || "Nepodařilo se načíst email nebo středisko přihlášeného uživatele."
          );
          return;
        }

        var inventoryItems = collectInventoryRowsForLocationBranch(panel, effectiveBranch);
        if (!effectiveBranch || inventoryItems.length === 0) {
          showCenteredWarningModal(
            "Všechno zboží je již zpracováno",
            "Pro přihlášeného uživatele už nejsou žádné položky k vyjádření ve větvi " + (effectiveBranch || "?") + "."
          );
          return;
        }

        showTaskInventoryAssignmentModal({
          displayName: userInfo && userInfo.displayName ? userInfo.displayName : "",
          email: userInfo && userInfo.email ? userInfo.email : "",
          center: userInfo && userInfo.center ? userInfo.center : "",
          branch: effectiveBranch,
          items: inventoryItems,
          onSave: function(decisions) {
            return applySkladDecisionsToRows(panel, decisions);
          }
        });
        return;
      }

      var locationInfo = "";
      if (effectiveBranch === "BRAMB") {
        locationInfo = "\nZobrazené zboží pro řešení: BRAMB a OSTATNÍ.";
      } else if (effectiveBranch === "CENTR") {
        locationInfo = "\nZobrazené zboží pro řešení: CENTR a OSTATNÍ.";
      }

      var goodsInfo = "";
      if (goodsForUser.length > 0) {
        goodsInfo = "\n\nZboží:\n- " + goodsForUser.join("\n- ");
      }

      var userInfoText = "";
      if (userInfo && (userInfo.email || userInfo.center)) {
        userInfoText = "\nUživatel: " + (userInfo.email || "") + (userInfo.center ? ", středisko " + userInfo.center : "") + ".";
      } else if (resolveErr) {
        userInfoText = "\nNepodařilo se dohledat středisko uživatele, použita byla lokace z úkolu.";
      }

      if (role === "vyjadreni_prodejny") {
        showCenteredInfoModal(
          "Typ otevřeného úkolu",
          "Byl rozpoznán úkol: vyjádření prodejny." + userInfoText + locationInfo + goodsInfo
        );
        return;
      }

      showCenteredInfoModal(
        "Typ otevřeného úkolu",
        "Byl rozpoznán úkol: inventura skladu." + userInfoText + locationInfo + goodsInfo
      );
    });
  }

  function createTocEditorModal() {
    var existing = parseExistingRowsFromHidden();
    var databaseCtrl = getWritableFieldControl("database");
    var databaseRows = parseDatabaseCsvRows(databaseCtrl && databaseCtrl.value ? databaseCtrl.value : "");
    var dsuByGoods = {};

    for (var d = 0; d < databaseRows.length; d++) {
      var dsuCells = databaseRows[d] || [];
      var dsuGoods = normalizeDsuGoodsNumber(dsuCells[3] || "");
      if (!dsuGoods || dsuByGoods[dsuGoods]) continue;

      dsuByGoods[dsuGoods] = {
        quantityDl: trimSafe(dsuCells[4] || "")
      };
    }

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

    function refreshDsuPreviewHighlightsFromToc() {
      var panel = document.getElementById("databaseTablePreviewPanel");
      if (!panel) return;

      var highlightedGoods = {};
      var tocGoodsInputs = dialog.querySelectorAll(".toc-cislo");
      for (var i = 0; i < tocGoodsInputs.length; i++) {
        var key = normalizeDsuGoodsNumber(tocGoodsInputs[i].value || "");
        if (key) highlightedGoods[key] = true;
      }

      var dsuRows = panel.querySelectorAll("tr[data-dsu-row='1']");
      for (var r = 0; r < dsuRows.length; r++) {
        var row = dsuRows[r];
        var goods = normalizeDsuGoodsNumber(row.getAttribute("data-dsu-goods") || "");
        if (goods && highlightedGoods[goods]) {
          row.style.backgroundColor = "#fde7e9";
        } else {
          row.style.backgroundColor = "";
        }
      }
    }

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

      function syncGoodsFromDsu() {
        var goodsInput = tr.querySelector(".toc-cislo");
        var dlInput = tr.querySelector(".toc-dl");
        if (!goodsInput || !dlInput) return;

        var goodsKey = normalizeDsuGoodsNumber(goodsInput.value || "");
        if (!goodsKey) {
          goodsInput.style.borderColor = "";
          goodsInput.style.backgroundColor = "";
          goodsInput.removeAttribute("title");
          return;
        }

        var dsuMatch = dsuByGoods[goodsKey];
        if (!dsuMatch) {
          goodsInput.style.borderColor = "#a4262c";
          goodsInput.style.backgroundColor = "#fff4f4";
          goodsInput.title = "Zboží nebylo nalezeno v DSU.";
          return;
        }

        goodsInput.style.borderColor = "#c8c6c4";
        goodsInput.style.backgroundColor = "";
        goodsInput.removeAttribute("title");

        dlInput.value = dsuMatch.quantityDl || "";
      }

      function recalc() {
        var dl = tr.querySelector(".toc-dl").value;
        var fiz = tr.querySelector(".toc-fiz").value;
        tr.querySelector(".toc-rozdil").value = String(calcDiff(dl, fiz));
      }

      function onGoodsChanged() {
        syncGoodsFromDsu();
        recalc();
        refreshDsuPreviewHighlightsFromToc();
      }

      tr.querySelector(".toc-dl").addEventListener("input", recalc);
      tr.querySelector(".toc-fiz").addEventListener("input", recalc);
      tr.querySelector(".toc-cislo").addEventListener("input", onGoodsChanged);
      tr.querySelector(".toc-cislo").addEventListener("change", onGoodsChanged);
      tr.querySelector(".toc-del").addEventListener("click", function() {
        tr.remove();
        refreshDsuPreviewHighlightsFromToc();
      });

      syncGoodsFromDsu();
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
    refreshDsuPreviewHighlightsFromToc();

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

        if (!dl || !fiz) {
          saveBtn.disabled = false;
          saveBtn.textContent = oldSaveText;

          var missingCols = [];
          if (!dl) missingCols.push("Množství na DL");
          if (!fiz) missingCols.push("Fyzicky dodané množství");

          showCenteredWarningModal(
            "Neúplný řádek",
            "Řádek " + (i + 1) + " (Číslo zboží: " + cislo + ") nemá vyplněno: " + missingCols.join(" a ") + "."
          );
          return;
        }

        rows.push({
          Cells: normalizeCells([cislo, lokace, dl, fiz, rozdil])
        });
      }

      enrichRowsFromLookup(rows, function(enrichedRows, missingGoods) {
        if (missingGoods && missingGoods.length > 0) {
          saveBtn.disabled = false;
          saveBtn.textContent = oldSaveText;
          showCenteredWarningModal(
            "Číslo zboží nenalezeno",
            "Nepodařilo se vyhledat tato čísla zboží:\n" + missingGoods.join(", ")
          );
          return;
        }

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
      "background:#fff;padding:22px 26px;border-radius:6px;width:560px;max-width:96vw;" +
      "font-family:Segoe UI, sans-serif;box-shadow:0 3px 12px rgba(0,0,0,0.35);";

    dialog.innerHTML =
      "<h2 style='margin-top:0;margin-bottom:12px;font-size:18px;'>Vyber typ chyby</h2>" +
      "<p style='margin:0 0 12px 0;font-size:13px;line-height:1.4;'>Zvol typ chyby pro tento záznam transferu.</p>" +
      "<div style='margin:0 0 12px 0;padding:10px;border:1px solid #edebe9;border-radius:4px;background:#faf9f8;font-size:12px;line-height:1.45;color:#323130;'>" +
      "  <div style='font-weight:600;margin-bottom:4px;'>Rozdíl v transferu</div>" +
      "  <div style='margin-bottom:8px;'>Manko nebo přebytek na zboží, které dorazilo na prodejnu.</div>" +
      "  <div style='font-weight:600;margin-bottom:4px;'>Problém s transferem</div>" +
      "  <ul style='margin:0 0 0 18px;padding:0;'>" +
      "    <li>Přišla paleta/krabice pro jinou lokaci</li>" +
      "    <li>Chyba v zabalení palety/krabice</li>" +
      "    <li>Cizí objekt v dodávce</li>" +
      "    <li>Poničené zboží/krabice</li>" +
      "    <li>Těžká krabice</li>" +
      "    <li>Nevhodná kombinace obsahu krabice</li>" +
      "  </ul>" +
      "</div>" +
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
        "rozdil_v_toc"
      ].forEach(lockField);

      ensureAdvancedTableLocked("rozdil_v_toc");

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
      initCurrentUserCenterForNewForm(0);
      initNewFormDeliveryNoteModal(0);
      initDatabaseField(0);
      return;
    }

    if (isEditTaskForm()) {
      initTaskFormCommentBinding();
      showTaskRoleModalOnTaskForm();
      return;
    }

    if (isDispForm()) {
      initDisplayFormDecisionColors();
    }
  });

})();
