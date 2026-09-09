(function () {
  "use strict";

  var FIELD_EMPLOYEE = "zamestnanec";
  var FIELD_MANAGER = "nadrizeny_dovolena";
  var FIELD_START = "datum_zacatek";
  var FIELD_END = "datum_konec";
  var FIELD_STATUS = "stav_dovolene";
  var FIELD_NOTE = "poznamka_dovolena";
  var EMPLOYEE_WEB_URL = "http://portal.samohyl.cz/bozp_po";
  var EMPLOYEE_LIST_TITLE = "zamestnanci";
  var FULL_NAME_CANDIDATES = ["jmeno_prijmeni", "jmeno_x005f_prijmeni", "JmenoPrijmeni", "Title"];
  var FIRST_NAME_CANDIDATES = ["jmeno_zamestnance", "jmeno_x005f_zamestnance", "JmenoZamestnance"];
  var LAST_NAME_CANDIDATES = ["prijmeni_zamestance", "prijmeni_zamestnance", "prijmeni_x005f_zamestance", "prijmeni_x005f_zamestnance", "PrijmeniZamestance", "PrijmeniZamestnance"];
  var MANAGER_CANDIDATES = ["eskalace_nadrizeny", "eskalace_x005f_nadrizeny", "EskalaceNadrizeny"];
  var MODAL_ID = "nef-current-user-modal";
  var STATUS_ID = "nef-manager-status";
  var booted = false;
  var employeeLookupTimer = null;
  var lastResolvedEmployeeName = "";

  function log() {
    if (!window.console || !console.log) return;
    var args = Array.prototype.slice.call(arguments);
    args.unshift("[DOVOLENA]");
    console.log.apply(console, args);
  }

  function fireEvent(element, eventName) {
    if (!element) return;
    var event;

    if (typeof Event === "function") {
      event = new Event(eventName, { bubbles: true });
    } else {
      event = document.createEvent("Event");
      event.initEvent(eventName, true, true);
    }

    element.dispatchEvent(event);
  }

  function fireInputEvents(element) {
    fireEvent(element, "input");
    fireEvent(element, "change");
    fireEvent(element, "blur");
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&#39;");
  }

  function toScalarString(value) {
    if (value === null || value === undefined) return "";
    if (Array.isArray(value)) {
      return value.map(function (item) { return toScalarString(item); }).filter(Boolean).join(", ");
    }
    if (typeof value === "object") {
      if (Object.prototype.hasOwnProperty.call(value, "results") && Array.isArray(value.results)) {
        return value.results.map(function (item) { return toScalarString(item); }).filter(Boolean).join(", ");
      }
      if (Object.prototype.hasOwnProperty.call(value, "Title")) return String(value.Title || "");
      if (Object.prototype.hasOwnProperty.call(value, "Name")) return String(value.Name || "");
      if (Object.prototype.hasOwnProperty.call(value, "LookupValue")) return String(value.LookupValue || "");
      if (Object.prototype.hasOwnProperty.call(value, "Label")) return String(value.Label || "");
      if (Object.prototype.hasOwnProperty.call(value, "Value")) return String(value.Value || "");
    }
    return String(value);
  }

  function cleanValue(value) {
    var raw = toScalarString(value).trim();
    if (!raw) return "";

    var parts = raw.split(/;#|[;\n\r\t]+/);
    var out = [];
    var seen = {};
    var i;

    for (i = 0; i < parts.length; i += 1) {
      var part = (parts[i] || "").trim();
      var key;
      if (!part) continue;
      if (/^\d+$/.test(part)) continue;
      key = part.toLowerCase();
      if (!seen[key]) {
        seen[key] = true;
        out.push(part);
      }
    }

    return out.length ? out.join(", ") : raw;
  }

  function normalizeText(value) {
    return cleanValue(value)
      .toLowerCase()
      .replace(/\s+/g, " ")
      .trim();
  }

  function getFieldValue(item, candidates) {
    var i;
    for (i = 0; i < candidates.length; i += 1) {
      if (Object.prototype.hasOwnProperty.call(item, candidates[i])) {
        var value = cleanValue(item[candidates[i]]);
        if (value) return value;
      }
    }
    return "";
  }

  function splitFullName(fullName) {
    var normalized = cleanValue(fullName).replace(/\s+/g, " ").trim();
    var parts;
    if (!normalized) {
      return { firstName: "", lastName: "", normalized: "" };
    }

    parts = normalized.split(" ");
    if (parts.length === 1) {
      return { firstName: parts[0], lastName: "", normalized: normalized };
    }

    return {
      firstName: parts[0],
      lastName: parts.slice(1).join(" "),
      normalized: normalized
    };
  }

  function findFieldElement(internalName) {
    return document.querySelector(
      "input[id^='" + internalName + "_'], textarea[id^='" + internalName + "_'], select[id^='" + internalName + "_'], " +
      "[id*='" + internalName + "_$TextField'], [id*='" + internalName + "_$ClientPeoplePicker'], [title='" + internalName + "']"
    );
  }

  function waitForField(internalName, callback) {
    var attempts = 0;
    var timer = window.setInterval(function () {
      attempts += 1;
      var element = findFieldElement(internalName);

      if (element) {
        window.clearInterval(timer);
        callback(true);
        return;
      }

      if (attempts >= 120) {
        window.clearInterval(timer);
        callback(false);
      }
    }, 250);
  }

  function getStatusHost() {
    var employeeField = findFieldElement(FIELD_EMPLOYEE);
    if (!employeeField) return null;
    return employeeField.parentNode || employeeField;
  }

  function ensureStatusBox() {
    var existing = document.getElementById(STATUS_ID);
    var host;

    if (existing) return existing;

    host = getStatusHost();
    if (!host || !host.parentNode) return null;

    existing = document.createElement("div");
    existing.id = STATUS_ID;
    existing.style.cssText = [
      "margin-top:8px",
      "padding:10px 12px",
      "border-radius:8px",
      "font-size:13px",
      "line-height:1.5",
      "background:#f1f5f9",
      "border:1px solid #cbd5e1",
      "color:#0f172a",
      "display:none"
    ].join(";");

    if (host.nextSibling) {
      host.parentNode.insertBefore(existing, host.nextSibling);
    } else {
      host.parentNode.appendChild(existing);
    }

    return existing;
  }

  function showStatus(message, tone) {
    var box = ensureStatusBox();
    if (!box) return;

    box.textContent = message || "";
    box.style.display = message ? "block" : "none";
    box.style.background = tone === "error" ? "#fef2f2" : (tone === "success" ? "#ecfdf5" : "#f1f5f9");
    box.style.borderColor = tone === "error" ? "#fecaca" : (tone === "success" ? "#a7f3d0" : "#cbd5e1");
    box.style.color = tone === "error" ? "#991b1b" : (tone === "success" ? "#065f46" : "#0f172a");
  }

  function findPeoplePicker(internalName) {
    if (typeof SPClientPeoplePicker === "undefined" || !SPClientPeoplePicker.SPClientPeoplePickerDict) {
      return null;
    }

    var dict = SPClientPeoplePicker.SPClientPeoplePickerDict;
    var keys = Object.keys(dict);
    var search = internalName.toLowerCase();
    var i;

    for (i = 0; i < keys.length; i += 1) {
      if (keys[i].toLowerCase().indexOf(search + "_") === 0) {
        return dict[keys[i]];
      }
    }

    for (i = 0; i < keys.length; i += 1) {
      if (keys[i].toLowerCase().indexOf(search) !== -1) {
        return dict[keys[i]];
      }
    }

    return null;
  }

  function setPeopleField(internalName, loginName, displayName) {
    var picker = findPeoplePicker(internalName);
    if (picker) {
      try {
        picker.DeleteProcessedUser();
      } catch (deleteError) {}

      try {
        picker.AddUserKeys(loginName || displayName || "");
        if (picker.ResolveAllUsers) picker.ResolveAllUsers();
        return true;
      } catch (pickerError) {
        log("PeoplePicker set failed", pickerError);
      }
    }

    var input = document.querySelector(
      "input[id^='" + internalName + "_'][type='text'], textarea[id^='" + internalName + "_'], [id*='" + internalName + "_$TextField']"
    );

    if (!input) return false;

    input.value = displayName || loginName || "";
    fireInputEvents(input);
    return true;
  }

  function getRequestDigest() {
    var digest = document.getElementById("__REQUESTDIGEST");
    return digest ? digest.value : "";
  }

  function getWebAbsoluteUrl() {
    if (window._spPageContextInfo && window._spPageContextInfo.webAbsoluteUrl) {
      return window._spPageContextInfo.webAbsoluteUrl;
    }

    return window.location.origin;
  }

  function getCurrentUser(callback) {
    var apiUrl = getWebAbsoluteUrl().replace(/\/$/, "") + "/_api/web/currentuser?$select=Id,Title,LoginName,Email";
    var xhr = new XMLHttpRequest();

    xhr.open("GET", apiUrl, true);
    xhr.setRequestHeader("Accept", "application/json;odata=verbose");

    if (getRequestDigest()) {
      xhr.setRequestHeader("X-RequestDigest", getRequestDigest());
    }

    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;

      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          var data = JSON.parse(xhr.responseText);
          callback(null, data.d);
        } catch (parseError) {
          callback(parseError);
        }
        return;
      }

      callback(new Error("Nepodařilo se načíst aktuálního uživatele. HTTP " + xhr.status));
    };

    xhr.send();
  }

  function loadEmployees(callback) {
    var apiUrl = EMPLOYEE_WEB_URL.replace(/\/$/, "") + "/_api/web/lists/getbytitle('" + encodeURIComponent(EMPLOYEE_LIST_TITLE) + "')/items?$top=5000";
    var xhr = new XMLHttpRequest();

    xhr.open("GET", apiUrl, true);
    xhr.setRequestHeader("Accept", "application/json;odata=verbose");

    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) return;

      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          var data = JSON.parse(xhr.responseText);
          callback(null, data.d && data.d.results ? data.d.results : []);
        } catch (parseError) {
          callback(parseError);
        }
        return;
      }

      callback(new Error("Nepodařilo se načíst seznam zaměstnanců. HTTP " + xhr.status));
    };

    xhr.send();
  }

  function findEmployeeRecord(employeeName, callback) {
    var target = splitFullName(employeeName);
    var targetFull = normalizeText(target.normalized);
    var targetFirst = normalizeText(target.firstName);
    var targetLast = normalizeText(target.lastName);

    loadEmployees(function (error, items) {
      var i;
      if (error) {
        callback(error);
        return;
      }

      for (i = 0; i < items.length; i += 1) {
        var item = items[i];
        var combinedName = getFieldValue(item, FULL_NAME_CANDIDATES);
        var firstName = getFieldValue(item, FIRST_NAME_CANDIDATES);
        var lastName = getFieldValue(item, LAST_NAME_CANDIDATES);
        var normalizedCombined = normalizeText(combinedName);
        var fullNameA = normalizeText(firstName + " " + lastName);
        var fullNameB = normalizeText(lastName + " " + firstName);
        var sameFullName = targetFull && (targetFull === fullNameA || targetFull === fullNameB || targetFull === normalizedCombined);
        var sameSplitName = targetFirst && targetLast && (
          (targetFirst === normalizeText(firstName) && targetLast === normalizeText(lastName)) ||
          (targetFirst === normalizeText(lastName) && targetLast === normalizeText(firstName))
        );

        if (sameFullName || sameSplitName) {
          callback(null, item);
          return;
        }
      }

      callback(null, null);
    });
  }

  function getManagerNameFromRecord(record) {
    return getFieldValue(record || {}, MANAGER_CANDIDATES);
  }

  function closeModal() {
    var modal = document.getElementById(MODAL_ID);
    if (modal && modal.parentNode) {
      modal.parentNode.removeChild(modal);
    }
  }

  function showConfirmModal(user, managerName, onYes, onNo) {
    closeModal();

    var overlay = document.createElement("div");
    overlay.id = MODAL_ID;
    overlay.style.cssText = [
      "position:fixed",
      "inset:0",
      "background:rgba(0,0,0,0.45)",
      "display:flex",
      "align-items:center",
      "justify-content:center",
      "z-index:999999"
    ].join(";");

    var card = document.createElement("div");
    card.style.cssText = [
      "width:min(92vw,520px)",
      "background:#ffffff",
      "border-radius:14px",
      "box-shadow:0 18px 50px rgba(0,0,0,0.28)",
      "padding:24px",
      "font-family:Segoe UI,Arial,sans-serif",
      "color:#1f2937"
    ].join(";");

    card.innerHTML = [
      "<div style='font-size:22px;font-weight:700;margin-bottom:12px;'>Nová žádost o dovolenou</div>",
      "<div style='font-size:14px;line-height:1.55;margin-bottom:16px;'>",
      "Přihlášený uživatel:<br>",
      "<strong>" + escapeHtml(user.Title || "") + "</strong><br>",
      escapeHtml(user.Email || ""),
      "</div>",
      "<div style='font-size:14px;line-height:1.55;margin-bottom:16px;'>",
      "Nadřízený:<br>",
      "<strong>" + escapeHtml(managerName || "Nenalezen v seznamu zaměstnanců") + "</strong>",
      "</div>",
      "<div style='font-size:15px;line-height:1.55;margin-bottom:22px;'>Bude žádost založena pro tohoto uživatele?</div>",
      "<div style='display:flex;gap:12px;justify-content:flex-end;'>",
      "<button type='button' data-role='no' style='min-width:110px;padding:10px 16px;border:1px solid #cbd5e1;background:#fff;border-radius:10px;cursor:pointer;font-size:14px;'>Ne</button>",
      "<button type='button' data-role='yes' style='min-width:110px;padding:10px 16px;border:1px solid #0f766e;background:#0f766e;color:#fff;border-radius:10px;cursor:pointer;font-size:14px;'>Ano</button>",
      "</div>"
    ].join("");

    overlay.appendChild(card);
    document.body.appendChild(overlay);

    card.querySelector("[data-role='yes']").onclick = function () {
      closeModal();
      if (typeof onYes === "function") onYes();
    };

    card.querySelector("[data-role='no']").onclick = function () {
      closeModal();
      if (typeof onNo === "function") onNo();
    };
  }

  function showInfo(message) {
    window.alert(message);
  }

  function fillManagerField(managerName) {
    if (!managerName) return false;
    return setPeopleField(FIELD_MANAGER, managerName, managerName);
  }

  function resolveManagerForEmployee(employeeName, callback) {
    if (!cleanValue(employeeName)) {
      callback(null, "", null);
      return;
    }

    findEmployeeRecord(employeeName, function (error, record) {
      if (error) {
        callback(error);
        return;
      }

      callback(null, getManagerNameFromRecord(record), record);
    });
  }

  function getSelectedEmployeeName() {
    var picker = findPeoplePicker(FIELD_EMPLOYEE);
    var info;
    var input;

    if (picker && typeof picker.GetAllUserInfo === "function") {
      info = picker.GetAllUserInfo();
      if (info && info.length) {
        return cleanValue(info[0].DisplayText || info[0].Title || info[0].Key || "");
      }
    }

    input = document.querySelector(
      "input[id^='" + FIELD_EMPLOYEE + "_'][type='text'], textarea[id^='" + FIELD_EMPLOYEE + "_'], [id*='" + FIELD_EMPLOYEE + "_$TextField']"
    );

    return input ? cleanValue(input.value) : "";
  }

  function syncManagerForEmployee(employeeName, silent) {
    var normalizedName = normalizeText(employeeName);

    if (!normalizedName || normalizedName === lastResolvedEmployeeName) return;
    lastResolvedEmployeeName = normalizedName;
    showStatus("Dohledávám nadřízeného pro zaměstnance: " + cleanValue(employeeName), "info");

    resolveManagerForEmployee(employeeName, function (error, managerName) {
      if (error) {
        log(error.message || error);
        showStatus("Nepodařilo se dohledat nadřízeného v seznamu zaměstnanců.", "error");
        if (!silent) showInfo("Nepodařilo se dohledat nadřízeného v seznamu zaměstnanců.");
        return;
      }

      if (!managerName) {
        log("Nadřízený nebyl nalezen pro zaměstnance", employeeName);
        showStatus("Nadřízený pro zaměstnance '" + cleanValue(employeeName) + "' nebyl nalezen.", "error");
        return;
      }

      fillManagerField(managerName);
      showStatus("Nadřízený: " + managerName, "success");
      log("Dohledán nadřízený", managerName, "pro zaměstnance", employeeName);
    });
  }

  function attachEmployeeWatcher() {
    var field = findFieldElement(FIELD_EMPLOYEE);
    var scheduleLookup = function () {
      if (employeeLookupTimer) {
        window.clearTimeout(employeeLookupTimer);
      }

      employeeLookupTimer = window.setTimeout(function () {
        syncManagerForEmployee(getSelectedEmployeeName(), true);
      }, 350);
    };

    if (!field) return;

    field.addEventListener("change", scheduleLookup);
    field.addEventListener("blur", scheduleLookup);
    field.addEventListener("input", scheduleLookup);
  }

  function fillCurrentUser(user, managerName) {
    var ok = setPeopleField(FIELD_EMPLOYEE, user.LoginName, user.Title);

    if (!ok) {
      showInfo("Pole '" + FIELD_EMPLOYEE + "' nebylo nalezeno nebo se jej nepodařilo vyplnit.");
      showStatus("Pole '" + FIELD_EMPLOYEE + "' nebylo nalezeno nebo se jej nepodařilo vyplnit.", "error");
      return;
    }

    if (managerName) {
      fillManagerField(managerName);
      showStatus("Nadřízený: " + managerName, "success");
    }

    log("Vyplněn uživatel do pole", FIELD_EMPLOYEE, user.LoginName || user.Title);

    window.setTimeout(function () {
      lastResolvedEmployeeName = "";
      syncManagerForEmployee(getSelectedEmployeeName() || user.Title || "", false);
    }, 800);
  }

  function boot() {
    if (booted) return;
    booted = true;

    log("Inicializace", {
      employee: FIELD_EMPLOYEE,
      manager: FIELD_MANAGER,
      start: FIELD_START,
      end: FIELD_END,
      status: FIELD_STATUS,
      note: FIELD_NOTE
    });

    waitForField(FIELD_EMPLOYEE, function (found) {
      if (!found) {
        log("Form field not found", FIELD_EMPLOYEE);
        return;
      }

      attachEmployeeWatcher();
      showStatus("Načítám informace o přihlášeném uživateli.", "info");

      getCurrentUser(function (error, user) {
        if (error) {
          log(error.message || error);
          showInfo("Nepodařilo se načíst informace o přihlášeném uživateli.");
          showStatus("Nepodařilo se načíst informace o přihlášeném uživateli.", "error");
          return;
        }

        resolveManagerForEmployee(user.Title || "", function (managerError, managerName) {
          if (managerError) {
            log(managerError.message || managerError);
            showStatus("Přihlášený uživatel načten, ale nadřízený se nepodařil dohledat.", "error");
          } else if (managerName) {
            showStatus("Nadřízený: " + managerName, "success");
          } else {
            showStatus("Přihlášený uživatel načten. Nadřízený zatím nebyl nalezen.", "info");
          }

          showConfirmModal(
            user,
            managerName,
            function () {
              fillCurrentUser(user, managerName);
            },
            function () {
              log("Uživatel zvolil ruční vyplnění formuláře.");
            }
          );
        });
      });
    });
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", boot);
  } else {
    boot();
  }
})();