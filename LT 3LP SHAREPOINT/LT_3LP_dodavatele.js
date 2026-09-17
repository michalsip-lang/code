(function () {
    "use strict";

    var CONFIG = {
        fieldInternalName: "dodavatel_3pl",
        alternateFieldInternalNames: ["dodavatele_3lp"],
        purchaserTargetFieldInternalNames: ["nakupci_3lp"],
        conditionFieldInternalNames: ["listace_jineho_dodavatele"],
        attachmentSourceFieldInternalNames: ["priloha_podklad", "priloha_podkald"],
        attachmentTargetFieldInternalNames: ["priloha_final"],
        taskTitlesForAttachmentMove: [
            "Příprava LT pro Alza",
            "Příprava LT pro Dr. Max"
        ],
        attachmentHint: "Vložte XLS soubor se seznamem zboží včetně cenotvorby / slevotvorby / kompenzace / forecastu / bere na sklad nebo trade / předpokládané datum odběru, pokud bere na sklad.",
        historyFieldInternalNames: ["historie"],
        workflowHistoryFieldInternalNames: ["historie_workflow"],
        siteUrl: "http://portal.samohyl.cz/nakup",
        listTitle: "Dodavatele",
        listRelativeUrl: "/nakup/Lists/Dodavatele",
        displayField: "Title",
        buyerFieldInternalName: "N_x00e1_kup_x010d__x00ed_",
        buyerColumnLabel: "Nákupčí krátce",
        purchaserListRelativeUrl: "/nakup/Lists/Purchasers",
        purchaserTitleFieldInternalName: "Title",
        purchaserEmailFieldInternalName: "Email",
        purchaserEmailColumnLabel: "Nákupčí email",
        separator: "; ",
        requestTimeoutMs: 30000
    };

    var dialogState = null;
    var observer = null;
    var fallbackInterval = null;
    var styleAdded = false;
    var delegatedClickBound = false;
    var currentUserName = "";
    var auditBound = false;
    var auditAttachmentCounts = {};
    var auditAttachmentNames = {};
    var auditFieldSnapshots = {};
    var auditFieldId = 0;
    var writingHistory = false;

    function addEvent(element, eventName, handler) {
        if (!element) {
            return;
        }

        if (element.addEventListener) {
            element.addEventListener(eventName, handler, false);
        } else if (element.attachEvent) {
            element.attachEvent("on" + eventName, handler);
        }
    }

    function removeEvent(element, eventName, handler) {
        if (!element) {
            return;
        }

        if (element.removeEventListener) {
            element.removeEventListener(eventName, handler, false);
        } else if (element.detachEvent) {
            element.detachEvent("on" + eventName, handler);
        }
    }

    function getAttribute(element, name) {
        if (!element || !element.getAttribute) {
            return "";
        }

        return element.getAttribute(name) || "";
    }

    function isTextField(element) {
        var tagName;
        var type;
        var contentEditable;
        var role;

        if (!element || !element.tagName) {
            return false;
        }

        tagName = element.tagName.toLowerCase();

        if (tagName === "textarea") {
            return true;
        }

        if (tagName !== "input") {
            contentEditable = getAttribute(element, "contenteditable").toLowerCase();
            role = getAttribute(element, "role").toLowerCase();

            return contentEditable === "true" || role === "textbox";
        }

        type = (getAttribute(element, "type") || "text").toLowerCase();

        return type === "text" ||
            type === "search";
    }

    function containsIgnoreCase(value, search) {
        return String(value).toLowerCase().indexOf(String(search).toLowerCase()) !== -1;
    }

    function matchesFieldInternalName(value) {
        var names = CONFIG.alternateFieldInternalNames || [];
        var i;

        if (containsIgnoreCase(value, CONFIG.fieldInternalName)) {
            return true;
        }

        for (i = 0; i < names.length; i += 1) {
            if (containsIgnoreCase(value, names[i])) {
                return true;
            }
        }

        return false;
    }

    function findFieldByAttribute(attributeName, expectedValue, partial) {
        var elements = document.getElementsByTagName("*");
        var i;
        var value;

        for (i = 0; i < elements.length; i += 1) {
            value = getAttribute(elements[i], attributeName);

            if (partial) {
                if (matchesFieldInternalName(value) && isTextField(elements[i])) {
                    return elements[i];
                }
            } else if (matchesFieldInternalName(value) && isTextField(elements[i])) {
                return elements[i];
            }
        }

        return null;
    }

    function findTextFieldInContainer(container) {
        var fields;
        var i;

        if (!container || !container.getElementsByTagName) {
            return null;
        }

        fields = container.getElementsByTagName("input");

        for (i = 0; i < fields.length; i += 1) {
            if (isTextField(fields[i])) {
                return fields[i];
            }
        }

        fields = container.getElementsByTagName("textarea");

        if (fields.length) {
            return fields[0];
        }

        fields = container.getElementsByTagName("*");

        for (i = 0; i < fields.length; i += 1) {
            if (isTextField(fields[i])) {
                return fields[i];
            }
        }

        return null;
    }

    function hasTargetReference(element) {
        return matchesFieldInternalName(getAttribute(element, "id")) ||
            matchesFieldInternalName(getAttribute(element, "name")) ||
            matchesFieldInternalName(getAttribute(element, "data-field-internal-name")) ||
            matchesFieldInternalName(getAttribute(element, "data-field-name")) ||
            matchesFieldInternalName(getAttribute(element, "data-name"));
    }

    function findTargetField() {
        var field;
        var elements;
        var i;
        var title;
        var ariaLabel;
        var dataName;
        var dataFieldName;
        var parent;

        field = document.getElementById(CONFIG.fieldInternalName);

        if (isTextField(field)) {
            return field;
        }

        if (CONFIG.alternateFieldInternalNames) {
            for (i = 0; i < CONFIG.alternateFieldInternalNames.length; i += 1) {
                field = document.getElementById(CONFIG.alternateFieldInternalNames[i]);

                if (isTextField(field)) {
                    return field;
                }
            }
        }

        field = findFieldByAttribute("name", CONFIG.fieldInternalName, true);

        if (field) {
            return field;
        }

        field = findFieldByAttribute("id", CONFIG.fieldInternalName, true);

        if (field) {
            return field;
        }

        field = findFieldByAttribute(
            "data-field-internal-name",
            CONFIG.fieldInternalName,
            false
        );

        if (field) {
            return field;
        }

        elements = document.getElementsByTagName("*");

        for (i = 0; i < elements.length; i += 1) {
            if (!isTextField(elements[i])) {
                continue;
            }

            title = getAttribute(elements[i], "title");
            ariaLabel = getAttribute(elements[i], "aria-label");
            dataName = getAttribute(elements[i], "data-name");
            dataFieldName = getAttribute(elements[i], "data-field-name");

            if (matchesFieldInternalName(title) ||
                matchesFieldInternalName(ariaLabel) ||
                matchesFieldInternalName(dataName) ||
                matchesFieldInternalName(dataFieldName)) {
                return elements[i];
            }
        }

        /* Některé verze SharePointu mají interní název pouze na wrapperu pole. */
        for (i = 0; i < elements.length; i += 1) {
            if (!hasTargetReference(elements[i])) {
                continue;
            }

            field = findTextFieldInContainer(elements[i]);

            if (field) {
                return field;
            }
        }

        /* Další možnost je wrapper několik úrovní nad viditelným vstupem. */
        for (i = 0; i < elements.length; i += 1) {
            if (!isTextField(elements[i])) {
                continue;
            }

            parent = elements[i].parentNode;

            while (parent && parent !== document.body) {
                if (hasTargetReference(parent)) {
                    return elements[i];
                }

                parent = parent.parentNode;
            }
        }

        return null;
    }

    function findFieldByInternalNames(names) {
        var elements = document.getElementsByTagName("*");
        var i;
        var j;
        var value;

        for (i = 0; i < elements.length; i += 1) {
            if (!isTextField(elements[i])) {
                continue;
            }

            for (j = 0; j < names.length; j += 1) {
                value = getAttribute(elements[i], "id") + " " +
                    getAttribute(elements[i], "name") + " " +
                    getAttribute(elements[i], "data-field-internal-name") + " " +
                    getAttribute(elements[i], "data-field-name");

                if (containsIgnoreCase(value, names[j])) {
                    return elements[i];
                }
            }
        }

        return null;
    }

    function findHistoryField() {
        var elements = document.getElementsByTagName("*");
        var fallback = null;
        var i;
        var j;
        var reference;
        var tagName;

        for (i = 0; i < elements.length; i += 1) {
            reference = getAttribute(elements[i], "id") + " " +
                getAttribute(elements[i], "name") + " " +
                getAttribute(elements[i], "data-field-internal-name") + " " +
                getAttribute(elements[i], "data-field-name") + " " +
                getAttribute(elements[i], "title");

            for (j = 0; j < CONFIG.historyFieldInternalNames.length; j += 1) {
                if (containsIgnoreCase(reference, CONFIG.historyFieldInternalNames[j]) &&
                    elements[i].tagName.toLowerCase() !== "label") {
                    tagName = elements[i].tagName.toLowerCase();

                    if (tagName === "textarea" ||
                        (tagName === "input" &&
                            String(getAttribute(elements[i], "type")).toLowerCase() !== "hidden")) {
                        return elements[i];
                    }

                    if (!fallback) {
                        fallback = elements[i];
                    }
                }
            }
        }

        return fallback;
    }

    function findHistoryDisplayElement() {
        var elements = document.getElementsByTagName("*");
        var i;
        var j;
        var reference;

        for (i = 0; i < elements.length; i += 1) {
            reference = getAttribute(elements[i], "id") + " " +
                getAttribute(elements[i], "name") + " " +
                getAttribute(elements[i], "data-field-internal-name") + " " +
                getAttribute(elements[i], "data-field-name");

            for (j = 0; j < CONFIG.historyFieldInternalNames.length; j += 1) {
                if (containsIgnoreCase(reference, CONFIG.historyFieldInternalNames[j]) &&
                    elements[i].tagName.toLowerCase() !== "label") {
                    return elements[i];
                }
            }
        }

        return null;
    }

    /* Volitelné pole, do ktereho zapisuje pouze workflow; drží historii oddelene od pole "historie". */
    function findWorkflowHistoryField() {
        var elements = document.getElementsByTagName("*");
        var names = CONFIG.workflowHistoryFieldInternalNames || [];
        var i;
        var j;
        var reference;
        var tagName;

        for (i = 0; i < elements.length; i += 1) {
            reference = getAttribute(elements[i], "id") + " " +
                getAttribute(elements[i], "name") + " " +
                getAttribute(elements[i], "data-field-internal-name") + " " +
                getAttribute(elements[i], "data-field-name") + " " +
                getAttribute(elements[i], "title");

            for (j = 0; j < names.length; j += 1) {
                if (containsIgnoreCase(reference, names[j]) &&
                    elements[i].tagName.toLowerCase() !== "label") {
                    tagName = elements[i].tagName.toLowerCase();

                    if (tagName === "textarea" ||
                        (tagName === "input" &&
                            String(getAttribute(elements[i], "type")).toLowerCase() !== "hidden")) {
                        return elements[i];
                    }
                }
            }
        }

        return null;
    }

    function parseCzTimestamp(value) {
        var match = String(value || "").match(
            /(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})/
        );

        if (!match) {
            return null;
        }

        return new Date(
            parseInt(match[3], 10),
            parseInt(match[2], 10) - 1,
            parseInt(match[1], 10),
            parseInt(match[4], 10),
            parseInt(match[5], 10),
            parseInt(match[6], 10)
        );
    }

    /* Slouci hlavni historii se zaznamy z odděleneho pole workflow, seřazeno chronologicky. */
    function mergeHistoryValues(primaryValue, workflowValue) {
        var lines = [];

        function collect(value) {
            String(value || "").split(/\r?\n/).forEach(function (line) {
                var trimmed = line.replace(/^\s+|\s+$/g, "");

                if (trimmed) {
                    lines.push({
                        text: trimmed,
                        date: parseCzTimestamp(parseHistoryLine(trimmed).when)
                    });
                }
            });
        }

        collect(primaryValue);
        collect(workflowValue);

        lines.sort(function (a, b) {
            if (a.date && b.date) {
                return a.date - b.date;
            }

            return 0;
        });

        return lines.map(function (line) {
            return line.text;
        }).join("\n");
    }

    function getCurrentUserName() {
        if (currentUserName) {
            return currentUserName;
        }

        if (typeof _spPageContextInfo !== "undefined" &&
            _spPageContextInfo &&
            _spPageContextInfo.userDisplayName) {
            return _spPageContextInfo.userDisplayName;
        }

        return "Neznámý uživatel";
    }

    function loadCurrentUser() {
        var request = new XMLHttpRequest();
        var url = CONFIG.siteUrl + "/_api/web/currentuser";

        if (typeof _spPageContextInfo !== "undefined" &&
            _spPageContextInfo &&
            _spPageContextInfo.userDisplayName) {
            currentUserName = _spPageContextInfo.userDisplayName;
        }

        try {
            request.open("GET", url, true);
            request.setRequestHeader("Accept", "application/json;odata=verbose");
            request.onreadystatechange = function () {
                var data;
                var user;

                if (request.readyState !== 4 ||
                    request.status < 200 || request.status >= 300) {
                    return;
                }

                try {
                    data = JSON.parse(request.responseText);
                    user = data && data.d ? data.d : data;

                    if (user && (user.Title || user.LoginName || user.Email)) {
                        currentUserName = user.Title || user.LoginName || user.Email;
                    }
                } catch (ignore) {
                    /* Page context zůstává záložním zdrojem identity. */
                }
            };
            request.send(null);
        } catch (ignoreRequest) {
            /* Načtení identity nesmí zablokovat formulář. */
        }
    }

    function getHistoryTimestamp() {
        var now = new Date();

        try {
            return now.toLocaleString("cs-CZ");
        } catch (error) {
            return now.toLocaleString();
        }
    }

    function appendHistoryEntry(action) {
        var historyField = findHistoryField();
        var row;
        var storedField;
        var oldValue;
        var entry;

        if (!historyField || !action || writingHistory) {
            return;
        }

        writingHistory = true;

        oldValue = String(historyField.value || "");

        if (!oldValue && historyField.closest) {
            row = historyField.closest("tr");
            storedField = row && row.querySelector ? row.querySelector(
                "textarea, input:not([type='hidden'])"
            ) : null;
            oldValue = String(storedField && storedField.value || "");
        }

        if (!oldValue) {
            oldValue = String(historyField.textContent || "");
        }
        entry = getCurrentUserName() +
            " | " + getHistoryTimestamp() +
            " | " + action;

        if (historyField.className.indexOf("dodavatel-picker-history-field") === -1) {
            historyField.className += " dodavatel-picker-history-field";
        }

        setTargetFieldValue(
            historyField,
            oldValue ? oldValue + "\n" + entry : entry
        );

        if (dialogState && dialogState.historyElement) {
            renderHistoryTable(dialogState.historyElement, historyField.value);
        }

        writingHistory = false;
    }

    function getAuditFieldName(field) {
        var row;
        var label;
        var title;

        if (!field) {
            return "neznámé pole";
        }

        title = getAttribute(field, "title") ||
            getAttribute(field, "data-field-name") ||
            getAttribute(field, "name") ||
            getAttribute(field, "id");

        row = field.closest ? field.closest("tr") : field.parentNode;
        label = row && row.querySelector ? row.querySelector("td.ms-formlabel, label") : null;

        return String(label ? (label.innerText || label.textContent || "") : title)
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "") || "neznámé pole";
    }

    function getAuditValue(field) {
        var row;
        var options;
        var visibleFields;
        var value;
        var i;

        if (!field) {
            return "";
        }

        if (String(field.tagName).toLowerCase() === "select" &&
            field.selectedIndex >= 0 && field.options[field.selectedIndex]) {
            return String(field.options[field.selectedIndex].text || "")
                .replace(/^\s+|\s+$/g, "");
        }

        row = field.closest ? field.closest("tr") : field.parentNode;
        visibleFields = row && row.querySelectorAll ? row.querySelectorAll(
            "input[type='text'], textarea, select, [contenteditable='true'], " +
            "a.ms-entity-respicker, span.ms-entity-respicker, .ms-lookup, " +
            ".sp-peoplepicker-topLevel"
        ) : [];

        for (i = 0; i < visibleFields.length; i += 1) {
            if (visibleFields[i] === field ||
                (visibleFields[i].offsetParent === null &&
                    visibleFields[i].tagName.toLowerCase() !== "textarea")) {
                continue;
            }

            if (String(visibleFields[i].tagName).toLowerCase() === "select" &&
                visibleFields[i].selectedIndex >= 0 &&
                visibleFields[i].options[visibleFields[i].selectedIndex]) {
                value = visibleFields[i].options[
                    visibleFields[i].selectedIndex
                ].text;
            } else {
                value = typeof visibleFields[i].value === "string" ?
                    visibleFields[i].value :
                    (visibleFields[i].innerText || visibleFields[i].textContent || "");
            }

            if (String(value || "").replace(/\s+/g, "").length) {
                return String(value);
            }
        }

        value = typeof field.value === "string" ? field.value :
            (field.textContent || field.innerHTML || "");

        return String(value)
            .replace(/<[^>]*>/g, " ")
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "")
            .substring(0, 160);
    }

    function isHistoryField(field) {
        var reference = getAttribute(field, "id") + " " +
            getAttribute(field, "name") + " " +
            getAttribute(field, "data-field-internal-name") + " " +
            getAttribute(field, "data-field-name") + " " +
            getAttribute(field, "title");

        return CONFIG.historyFieldInternalNames.some(function (name) {
            return containsIgnoreCase(reference, name);
        }) || CONFIG.workflowHistoryFieldInternalNames.some(function (name) {
            return containsIgnoreCase(reference, name);
        });
    }

    function isInternalAuditUi(element) {
        var current = element;

        while (current && current !== document.body) {
            if (containsIgnoreCase(getAttribute(current, "class"), "dodavatel-picker-")) {
                return true;
            }

            current = current.parentNode;
        }

        return false;
    }

    function isEditorChromeField(field) {
        var reference = getAttribute(field, "id") + " " +
            getAttribute(field, "name") + " " +
            getAttribute(field, "title") + " " +
            getAttribute(field, "class");

        return containsIgnoreCase(reference, "FontFamilyStyleValue") ||
            containsIgnoreCase(reference, "FontSizeStyleValue") ||
            containsIgnoreCase(reference, "Ribbon.") ||
            containsIgnoreCase(reference, "ms-rte") ||
            containsIgnoreCase(reference, "rteStyle");
    }

    function hasAuditFormContext(field) {
        var row;

        if (getAttribute(field, "data-field-internal-name") ||
            getAttribute(field, "data-field-name")) {
            return true;
        }

        row = field.closest ? field.closest("tr") : null;

        return !!(row && row.querySelector && row.querySelector(
            "td.ms-formlabel, th.ms-formlabel, label, nobr"
        ));
    }

    function isAuditableField(field) {
        var tagName;
        var type;

        if (!field || !field.tagName || isHistoryField(field) ||
            isInternalAuditUi(field) || isEditorChromeField(field) ||
            !hasAuditFormContext(field)) {
            return false;
        }

        tagName = field.tagName.toLowerCase();

        if (tagName === "textarea" || tagName === "select") {
            return true;
        }

        if (tagName === "input") {
            type = (getAttribute(field, "type") || "text").toLowerCase();

            return type === "text" ||
                type === "search" ||
                type === "checkbox";
        }

        return getAttribute(field, "contenteditable").toLowerCase() === "true" ||
            getAttribute(field, "role").toLowerCase() === "textbox";
    }

    function getAuditFieldKey(field) {
        var key;

        key = getAttribute(field, "data-dodavatel-audit-key") ||
            getAttribute(field, "data-field-internal-name") ||
            getAttribute(field, "data-field-name") ||
            getAttribute(field, "name") ||
            getAttribute(field, "id") ||
            getAttribute(field, "title");

        if (!key) {
            auditFieldId += 1;
            key = "dodavatel-audit-field-" + auditFieldId;
            field.setAttribute("data-dodavatel-audit-key", key);
        }

        return key;
    }

    function getAuditComparableValue(field) {
        var tagName;
        var type;

        if (!field) {
            return "";
        }

        tagName = String(field.tagName || "").toLowerCase();

        if (tagName === "input") {
            type = (getAttribute(field, "type") || "text").toLowerCase();

            if (type === "checkbox") {
                return field.checked ? "zaškrtnuto" : "nezaškrtnuto";
            }
        }

        return getAuditValue(field);
    }

    function updateAuditFieldSnapshot(field) {
        if (!isAuditableField(field)) {
            return;
        }

        auditFieldSnapshots[getAuditFieldKey(field)] = getAuditComparableValue(field);
    }

    function auditFieldValueChanges() {
        var fields;
        var field;
        var key;
        var value;
        var previous;
        var i;

        if (writingHistory || !document.querySelectorAll) {
            return;
        }

        fields = document.querySelectorAll(
            "input, textarea, select, [contenteditable='true'], [role='textbox']"
        );

        for (i = 0; i < fields.length; i += 1) {
            field = fields[i];

            if (!isAuditableField(field)) {
                continue;
            }

            key = getAuditFieldKey(field);
            value = getAuditComparableValue(field);
            previous = auditFieldSnapshots[key];
            auditFieldSnapshots[key] = value;

            if (typeof previous === "undefined" || previous === value) {
                continue;
            }

            appendHistoryEntry(
                "Automaticky změněno pole " + getAuditFieldName(field) +
                (value ? ": " + value : " (vymazáno)")
            );
        }
    }

    function getAttachmentItemCount(container) {
        if (!container || !container.querySelectorAll) {
            return 0;
        }

        return container.querySelectorAll(
            ".tispMultipleUploadFT .containerItems > *"
        ).length;
    }

    function getAttachmentItemNames(container) {
        var controls;
        var data;
        var fileUrl;
        var items;
        var names = {};
        var i;
        var item;
        var name;

        if (!container || !container.querySelectorAll) {
            return names;
        }

        controls = container.querySelectorAll("input[id^='tisa_controlvalue_']");

        for (i = 0; i < controls.length; i += 1) {
            try {
                data = JSON.parse(controls[i].value || "[]");
            } catch (ignore) {
                data = [];
            }

            if (!data || Object.prototype.toString.call(data) !== "[object Array]") {
                continue;
            }

            for (item = 0; item < data.length; item += 1) {
                fileUrl = String(data[item] && data[item].FileUrl || "");
                name = fileUrl.substring(fileUrl.lastIndexOf("/") + 1);

                if (name) {
                    try {
                        name = decodeURIComponent(name);
                    } catch (ignoreDecode) {
                        /* Ponecháme neupravený název souboru. */
                    }
                    names[name] = true;
                }
            }
        }

        if (Object.keys(names).length) {
            return names;
        }

        items = container.querySelectorAll(
            ".tispMultipleUploadFT .containerItems > *"
        );

        for (i = 0; i < items.length; i += 1) {
            item = items[i];
            name = getAttachmentFileName(item);

            if (name) {
                names[name] = true;
            }
        }

        return names;
    }

    function getAttachmentFileName(item) {
        var text;
        var match;

        if (!item) {
            return "";
        }

        text = String(item.textContent || item.innerText || "")
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "");

        match = text.match(/[\w().%+-]+\.(?:pdf|xls|xlsx|csv|doc|docx|zip|rar|txt|xml)(?=\s|$)/i);

        return match ? match[0] : "";
    }

    function getAttachmentAuditName(container) {
        var text = container && container.querySelector ?
            container.querySelector(".ms-formlabel, h3, nobr") : null;

        return String(text ? (text.innerText || text.textContent || "") : "přílohy")
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "");
    }

    function auditAttachmentChanges() {
        var names = [
            CONFIG.attachmentSourceFieldInternalNames,
            CONFIG.attachmentTargetFieldInternalNames
        ];
        var i;
        var container;
        var key;
        var count;
        var previous;
        var namesSnapshot;
        var previousNames;
        var name;

        for (i = 0; i < names.length; i += 1) {
            container = findAttachmentFieldContainer(names[i]);

            if (!container) {
                continue;
            }

            key = names[i].join("|");
            count = getAttachmentItemCount(container);
            namesSnapshot = getAttachmentItemNames(container);
            previous = auditAttachmentCounts[key];
            previousNames = auditAttachmentNames[key] || {};
            auditAttachmentCounts[key] = count;
            auditAttachmentNames[key] = namesSnapshot;

            if (typeof previous === "undefined") {
                continue;
            }

            for (name in namesSnapshot) {
                if (namesSnapshot[name] && !previousNames[name]) {
                    appendHistoryEntry(
                        "Přidána příloha v poli " +
                        getAttachmentAuditName(container) + ": " + name
                    );
                }
            }

            for (name in previousNames) {
                if (previousNames[name] && !namesSnapshot[name]) {
                    appendHistoryEntry(
                        "Odstraněna příloha v poli " +
                        getAttachmentAuditName(container) + ": " + name
                    );
                }
            }
        }
    }

    function scheduleAttachmentAudit() {
        [250, 1000, 2500, 5000].forEach(function (delay) {
            window.setTimeout(auditAttachmentChanges, delay);
        });
    }

    function scheduleFieldAudit() {
        [250, 1000, 2500, 5000].forEach(function (delay) {
            window.setTimeout(auditFieldValueChanges, delay);
        });
    }

    function handleAuditFieldChange(event) {
        var field = event && (event.target || event.srcElement);

        if (!field || writingHistory || !isAuditableField(field)) {
            return;
        }

        appendHistoryEntry(
            "Vyplněno/upraveno pole " + getAuditFieldName(field) +
            (getAuditComparableValue(field) ? ": " + getAuditComparableValue(field) : " (vymazáno)")
        );
        updateAuditFieldSnapshot(field);
    }

    function isSaveAction(element) {
        var text = String(
            element && (element.value || element.innerText || element.textContent || "")
        ).replace(/^\s+|\s+$/g, "").toLowerCase();

        return text === "uložit" || text === "ulozit" || text === "save" ||
            containsIgnoreCase(getAttribute(element, "title"), "uložit") ||
            containsIgnoreCase(getAttribute(element, "title"), "save");
    }

    function handleAuditClick(event) {
        var target = event && (event.target || event.srcElement);
        var current = target;

        while (current && current !== document.body) {
            if (containsIgnoreCase(getAttribute(current, "class"), "addnewattachment")) {
                scheduleAttachmentAudit();
                break;
            }

            if (isSaveAction(current)) {
                auditFieldValueChanges();
                appendHistoryEntry("Uživatel klikl na Uložit");
                break;
            }

            if (containsIgnoreCase(getAttribute(current, "class"), "delete") ||
                containsIgnoreCase(getAttribute(current, "class"), "remove") ||
                containsIgnoreCase(getAttribute(current, "class"), "deleteall")) {
                appendHistoryEntry("Uživatel odstranil přílohu nebo požádal o její odstranění");
                break;
            }

            current = current.parentNode;
        }
    }

    function initializeAuditLogging() {
        if (auditBound) {
            return;
        }

        addEvent(document, "change", handleAuditFieldChange);
        addEvent(document, "click", handleAuditClick);
        auditBound = true;
        auditFieldValueChanges();
        scheduleFieldAudit();
        auditAttachmentChanges();
    }

    function escapeHtml(value) {
        return String(value || "")
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/\"/g, "&quot;")
            .replace(/'/g, "&#39;");
    }

    function parseHistoryLine(line) {
        var parts = String(line || "").split(" | ");
        var result = { who: "", when: "", what: "" };
        var i;

        if (parts.length >= 3 &&
            parts[0].indexOf("KDO:") !== 0 &&
            parts[1].indexOf("KDY:") !== 0 &&
            parts[2].indexOf("CO:") !== 0) {
            result.who = parts[0];
            result.when = parts[1];
            result.what = parts.slice(2).join(" | ");
            return result;
        }

        for (i = 0; i < parts.length; i += 1) {
            if (parts[i].indexOf("KDO:") === 0) {
                result.who = parts[i].replace(/^KDO:\s*/, "");
            } else if (parts[i].indexOf("KDY:") === 0) {
                result.when = parts[i].replace(/^KDY:\s*/, "");
            } else if (parts[i].indexOf("CO:") === 0) {
                result.what = parts[i].replace(/^CO:\s*/, "");
            }
        }

        if (!result.what && line) {
            result.what = line;
        }

        return result;
    }

    function renderHistoryTable(container, value) {
        var lines = String(value || "").split(/\r?\n/).filter(function (line) {
            return line.replace(/^\s+|\s+$/g, "") !== "";
        });
        var table;
        var head;
        var row;
        var cell;
        var tableRow;
        var item;
        var i;

        while (container.firstChild) {
            container.removeChild(container.firstChild);
        }

        if (!lines.length) {
            container.appendChild(createElement(
                "div",
                "dodavatel-picker-history-empty",
                "Zatím nejsou evidované žádné změny."
            ));
            return;
        }

        table = createElement("table", "dodavatel-picker-history-table");
        head = document.createElement("thead");
        row = document.createElement("tr");

        ["Kdo", "Kdy", "Co provedl"].forEach(function (text) {
            row.appendChild(createElement("th", "", text));
        });

        head.appendChild(row);
        table.appendChild(head);
        row = document.createElement("tbody");

        for (i = lines.length - 1; i >= 0; i -= 1) {
            item = parseHistoryLine(lines[i]);
            tableRow = document.createElement("tr");
            tableRow.appendChild(createElement("td", "", item.who));
            tableRow.appendChild(createElement("td", "", item.when));
            tableRow.appendChild(createElement("td", "", item.what));
            row.appendChild(tableRow);
        }

        table.appendChild(row);
        container.appendChild(table);
    }

    function initializeHistoryDisplay() {
        var historyElement = findHistoryDisplayElement();
        var workflowField = findWorkflowHistoryField();
        var sourceValue;
        var workflowValue;

        if (!historyElement ||
            getAttribute(historyElement, "data-dodavatel-history-rendered") === "true") {
            return;
        }

        sourceValue = historyElement.value || historyElement.innerHTML || historyElement.textContent || "";

        if (workflowField) {
            workflowValue = workflowField.value || workflowField.textContent || "";
            sourceValue = mergeHistoryValues(sourceValue, workflowValue);
        }

        injectStyles();
        renderHistoryTable(historyElement, sourceValue);
        historyElement.setAttribute("data-dodavatel-history-rendered", "true");
    }

    function findConditionField() {
        var elements = document.getElementsByTagName("input");
        var i;
        var j;
        var value;

        for (i = 0; i < elements.length; i += 1) {
            if (String(elements[i].type).toLowerCase() !== "checkbox") {
                continue;
            }

            value = getAttribute(elements[i], "id") + " " +
                getAttribute(elements[i], "name") + " " +
                getAttribute(elements[i], "data-field-internal-name");

            for (j = 0; j < CONFIG.conditionFieldInternalNames.length; j += 1) {
                if (containsIgnoreCase(value, CONFIG.conditionFieldInternalNames[j])) {
                    return elements[i];
                }
            }
        }

        return null;
    }

    function isSupplierPickerEnabled() {
        var conditionField = findConditionField();

        return !conditionField || conditionField.checked === true;
    }

    function updateConditionalFieldAvailability() {
        var enabled = isSupplierPickerEnabled();
        var supplierField = findTargetField();
        var purchaserField = findFieldByInternalNames(
            CONFIG.purchaserTargetFieldInternalNames
        );

        if (supplierField) {
            supplierField.disabled = !enabled;
            supplierField.setAttribute("aria-disabled", enabled ? "false" : "true");
        }

        if (purchaserField) {
            purchaserField.disabled = !enabled;
            purchaserField.setAttribute("aria-disabled", enabled ? "false" : "true");
        }
    }

    function handleConditionFieldChange(event) {
        var conditionField;
        var supplierField;
        var purchaserField;

        event = event || window.event;
        conditionField = event && (event.target || event.srcElement);

        if (conditionField && conditionField.checked === false) {
            supplierField = findTargetField();
            purchaserField = findFieldByInternalNames(
                CONFIG.purchaserTargetFieldInternalNames
            );

            setTargetFieldValue(supplierField, "");
            setTargetFieldValue(purchaserField, "");
            appendHistoryEntry(
                "Vypnuta evidence jiného dodavatele; vymazána pole dodavatelů a nákupčích"
            );
        }

        updateConditionalFieldAvailability();
    }

    function bindConditionField() {
        var conditionField = findConditionField();

        if (!conditionField) {
            return;
        }

        if (getAttribute(conditionField, "data-dodavatel-condition-bound") !== "true") {
            addEvent(conditionField, "change", handleConditionFieldChange);
            conditionField.setAttribute("data-dodavatel-condition-bound", "true");
        }

        updateConditionalFieldAvailability();
    }

    function hasFieldNameReference(element, names) {
        var reference;
        var i;

        if (!element || !names || !names.length) {
            return false;
        }

        reference = getAttribute(element, "id") + " " +
            getAttribute(element, "name") + " " +
            getAttribute(element, "data-field-internal-name") + " " +
            getAttribute(element, "data-field-name") + " " +
            getAttribute(element, "data-name") + " " +
            getAttribute(element, "class") + " " +
            getAttribute(element, "title");

        for (i = 0; i < names.length; i += 1) {
            if (containsIgnoreCase(reference, names[i])) {
                return true;
            }
        }

        return false;
    }

    function findAttachmentFieldContainer(names) {
        var elements = document.getElementsByTagName("*");
        var element;
        var parent;
        var i;

        for (i = 0; i < elements.length; i += 1) {
            element = elements[i];

            if (!hasFieldNameReference(element, names)) {
                continue;
            }

            parent = element;

            while (parent && parent !== document.body) {
                if (String(parent.tagName).toLowerCase() === "tr" ||
                    containsIgnoreCase(getAttribute(parent, "class"), "ms-formfield")) {
                    return parent;
                }

                parent = parent.parentNode;
            }

            return element;
        }

        return null;
    }

    function hasAttachmentContent(container) {
        var fieldElements;
        var fileInputs;
        var links;
        var attachmentMarkers;
        var treeInfoItems;
        var treeInfoValues;
        var value;
        var i;

        if (!container) {
            return false;
        }

        treeInfoItems = container.querySelectorAll ?
            container.querySelectorAll(".tispMultipleUploadFT .containerItems > *") : [];

        if (treeInfoItems.length) {
            return true;
        }

        treeInfoValues = container.querySelectorAll ?
            container.querySelectorAll("input[id^='tisa_controlvalue_']") : [];

        for (i = 0; i < treeInfoValues.length; i += 1) {
            if (String(treeInfoValues[i].value || "").replace(/\s+/g, "") !== "") {
                return true;
            }
        }

        if (hasFieldNameReference(
            container,
            CONFIG.attachmentSourceFieldInternalNames
        ) && typeof container.value === "string" &&
            container.value.replace(/\s+/g, "") !== "") {
            return true;
        }

        fieldElements = container.querySelectorAll ?
            container.querySelectorAll("textarea, input, [contenteditable='true'], [data-field-internal-name], [data-field-name]") : [];

        for (i = 0; i < fieldElements.length; i += 1) {
            if (!hasFieldNameReference(
                fieldElements[i],
                CONFIG.attachmentSourceFieldInternalNames
            )) {
                continue;
            }

            value = typeof fieldElements[i].value === "string" ?
                fieldElements[i].value :
                (fieldElements[i].textContent || fieldElements[i].innerHTML || "");

            if (String(value)
                .replace(/&nbsp;|<br\s*\/?>(\s*)/gi, " ")
                .replace(/<[^>]*>/g, "")
                .replace(/\s+/g, "") !== "") {
                return true;
            }
        }

        fileInputs = container.querySelectorAll ?
            container.querySelectorAll("input[type='file']") : [];

        for (i = 0; i < fileInputs.length; i += 1) {
            if ((fileInputs[i].files && fileInputs[i].files.length) ||
                fileInputs[i].value ||
                getAttribute(fileInputs[i], "value")) {
                return true;
            }
        }

        attachmentMarkers = container.querySelectorAll ?
            container.querySelectorAll("[data-attachment-name], [data-attachment], .ms-fileField, .tispMultipleUploadFT .containerItems a") : [];

        if (attachmentMarkers.length) {
            return true;
        }

        links = container.getElementsByTagName ? container.getElementsByTagName("a") : [];

        for (i = 0; i < links.length; i += 1) {
            if (containsIgnoreCase(getAttribute(links[i], "href"), "/attachments/")) {
                return true;
            }
        }

        return false;
    }

    function updateAttachmentFieldAvailability() {
        var sourceContainer = findAttachmentFieldContainer(
            CONFIG.attachmentSourceFieldInternalNames
        );
        var targetContainer = findAttachmentFieldContainer(
            CONFIG.attachmentTargetFieldInternalNames
        );
        var available = hasAttachmentContent(sourceContainer);

        if (!targetContainer) {
            return;
        }

        targetContainer.style.display = available ? "" : "none";
        targetContainer.hidden = !available;
        targetContainer.setAttribute("aria-hidden", available ? "false" : "true");
    }

    function bindAttachmentSourceChanges() {
        var sourceContainer = findAttachmentFieldContainer(
            CONFIG.attachmentSourceFieldInternalNames
        );
        var fileInputs;
        var i;

        if (!sourceContainer || !sourceContainer.querySelectorAll) {
            return;
        }

        if (getAttribute(sourceContainer, "data-attachment-availability-bound") !== "true") {
            addEvent(sourceContainer, "change", updateAttachmentFieldAvailability);
            addEvent(sourceContainer, "input", updateAttachmentFieldAvailability);
            sourceContainer.setAttribute("data-attachment-availability-bound", "true");
        }

        if (hasFieldNameReference(
            sourceContainer,
            CONFIG.attachmentSourceFieldInternalNames
        ) && typeof sourceContainer.value === "string" &&
            getAttribute(sourceContainer, "data-attachment-availability-bound") !== "true") {
            addEvent(sourceContainer, "change", updateAttachmentFieldAvailability);
            addEvent(sourceContainer, "input", updateAttachmentFieldAvailability);
            sourceContainer.setAttribute("data-attachment-availability-bound", "true");
        }

        fileInputs = sourceContainer.querySelectorAll(
            "textarea, input, [contenteditable='true']"
        );

        for (i = 0; i < fileInputs.length; i += 1) {
            if (!hasFieldNameReference(
                fileInputs[i],
                CONFIG.attachmentSourceFieldInternalNames
            )) {
                continue;
            }

            if (getAttribute(fileInputs[i], "data-attachment-availability-bound") === "true") {
                continue;
            }

            addEvent(fileInputs[i], "change", updateAttachmentFieldAvailability);
            addEvent(fileInputs[i], "input", updateAttachmentFieldAvailability);
            fileInputs[i].setAttribute("data-attachment-availability-bound", "true");
        }
    }

    function customizeAttachmentControls() {
        var names = [
            CONFIG.attachmentSourceFieldInternalNames,
            CONFIG.attachmentTargetFieldInternalNames
        ];
        var i;
        var container;
        var fieldBody;
        var buttons;
        var removeButtons;
        var hint;

        for (i = 0; i < names.length; i += 1) {
            container = findAttachmentFieldContainer(names[i]);

            if (!container || !container.querySelectorAll) {
                continue;
            }

            fieldBody = container.querySelector("td.ms-formbody") || container;
            buttons = container.querySelectorAll("a.addNewAttachment");
            removeButtons = container.querySelectorAll(
                "a.deleteAll, .deleteAll"
            );

            Array.prototype.forEach.call(buttons, function (button) {
                if (button.textContent !== "Vložit přílohy") {
                    button.textContent = "Vložit přílohy";
                }

                if (String(button.className).indexOf(
                    "dodavatel-picker-attachment-button"
                ) === -1) {
                    button.className += " dodavatel-picker-attachment-button";
                }

                if (getAttribute(button, "role") !== "button") {
                    button.setAttribute("role", "button");
                }
            });

            Array.prototype.forEach.call(removeButtons, function (button) {
                button.style.display = "none";
                button.style.visibility = "hidden";
                button.setAttribute("aria-hidden", "true");
            });

            if (i === 0 && !fieldBody.querySelector(
                ".dodavatel-picker-attachment-hint"
            )) {
                hint = createElement(
                    "div",
                    "dodavatel-picker-attachment-hint",
                    CONFIG.attachmentHint
                );
                fieldBody.insertBefore(hint, fieldBody.firstChild);
            }
        }
    }

    function isTargetTaskForAttachmentMove() {
        var taskTitle = document.getElementById("lblTaskName");
        var currentTitle;

        if (!taskTitle) {
            return false;
        }

        currentTitle = String(taskTitle.textContent || taskTitle.innerText || "")
                .replace(/\s+/g, " ")
                .replace(/^\s+|\s+$/g, "");

        return CONFIG.taskTitlesForAttachmentMove.indexOf(currentTitle) !== -1;
    }

    function moveTaskAttachmentsAboveOutcome() {
        var outcomeRow;
        var parent;
        var targetRow;
        var taskContainer;

        if (!isTargetTaskForAttachmentMove()) {
            return;
        }

        taskContainer = document.getElementById("tblMainContainer");

        if (taskContainer && String(taskContainer.className).indexOf(
            "dodavatel-picker-task-styled"
        ) === -1) {
            taskContainer.className += " dodavatel-picker-task-styled";
        }

        outcomeRow = document.getElementById("TaskFormTrOutcomes");

        if (!outcomeRow || !outcomeRow.parentNode) {
            return;
        }

        parent = outcomeRow.parentNode;
        targetRow = findAttachmentFieldContainer(
            CONFIG.attachmentTargetFieldInternalNames
        );

        if (targetRow && targetRow.parentNode !== parent) {
            parent.insertBefore(targetRow, outcomeRow);
        }

        if (targetRow) {
            if (String(targetRow.className).indexOf(
                "dodavatel-picker-task-attachment-row"
            ) === -1) {
                targetRow.className += " dodavatel-picker-task-attachment-row";
            }
        }
    }

    function injectStyles() {
        var style;
        var css =
            ".dodavatel-picker-overlay{" +
                "position:fixed;z-index:2147483646;top:0;right:0;bottom:0;left:0;" +
                "background:rgba(20,22,47,.46);overflow:auto;padding:24px;box-sizing:border-box;" +
            "}" +
            ".dodavatel-picker-dialog{" +
                "position:relative;width:100%;max-width:700px;margin:24px auto;background:#ffffff;" +
                "border:1px solid rgba(41,41,130,.16);box-shadow:0 12px 30px rgba(41,41,130,.22);" +
                "font-family:Segoe UI,Arial,sans-serif;color:#14162f;box-sizing:border-box;" +
            "}" +
            ".dodavatel-picker-header{" +
                "position:relative;padding:16px 52px 14px 20px;" +
                "background:#292982;" +
                "color:#ffffff;border-bottom:0;" +
            "}" +
            ".dodavatel-picker-title{margin:0;color:#ffffff;font-size:20px;font-weight:600;}" +
            ".dodavatel-picker-body{padding:20px;background:#ffffff;}" +
            ".dodavatel-picker-search{" +
                "display:block;width:100%;height:36px;padding:7px 10px;border:1px solid #808184;" +
                "box-sizing:border-box;font-size:14px;" +
            "}" +
            ".dodavatel-picker-search:focus{border-color:#292982;outline:1px solid #292982;}" +
            ".dodavatel-picker-status{min-height:22px;margin:10px 0;font-size:13px;color:#808184;}" +
            ".dodavatel-picker-error{padding:10px;border:1px solid #e01b37;background:#fff0f2;color:#a61f2c;line-height:1.4;}" +
            ".dodavatel-picker-list{height:320px;overflow-y:auto;border:1px solid rgba(41,41,130,.16);background:#f3f4f8;}" +
            ".dodavatel-picker-list-header{display:grid;grid-template-columns:24px minmax(160px,1.35fr) minmax(130px,1fr) minmax(170px,1.1fr);padding:7px 12px;border-bottom:1px solid rgba(41,41,130,.16);background:#f3f4f8;color:#808184;font-size:12px;font-weight:600;}" +
            ".dodavatel-picker-list-header span:first-child{visibility:hidden;}" +
            ".dodavatel-picker-item{display:grid;grid-template-columns:24px minmax(160px,1.35fr) minmax(130px,1fr) minmax(170px,1.1fr);align-items:center;padding:9px 12px;border-bottom:1px solid rgba(41,41,130,.12);cursor:pointer;font-size:14px;line-height:20px;background:#ffffff;}" +
            ".dodavatel-picker-item:hover{background:#f3f4f8;}" +
            ".dodavatel-picker-item input{margin:0 9px 0 0;vertical-align:middle;}" +
            ".dodavatel-picker-supplier-name{min-width:0;padding-right:12px;overflow-wrap:break-word;}" +
            ".dodavatel-picker-buyer-name{min-width:0;color:#808184;overflow-wrap:break-word;}" +
            ".dodavatel-picker-email{min-width:0;color:#292982;overflow-wrap:break-word;}" +
            ".dodavatel-picker-history-field{border-color:#292982 !important;background:#f0f1fb !important;line-height:1.5;}" +
            ".dodavatel-picker-history-card{margin-top:16px;border-left:4px solid #292982;border-top:1px solid rgba(41,41,130,.12);border-right:1px solid rgba(41,41,130,.12);border-bottom:1px solid rgba(41,41,130,.12);background:#ffffff;}" +
            ".dodavatel-picker-history-title{padding:10px 12px;margin:0;color:#292982;font-size:15px;font-weight:700;border-bottom:1px solid rgba(41,41,130,.12);}" +
            ".dodavatel-picker-history-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:12px;}" +
            ".dodavatel-picker-history-table th,.dodavatel-picker-history-table td{padding:8px 10px;text-align:left;vertical-align:top;border:1px solid rgba(41,41,130,.12);overflow-wrap:anywhere;}" +
            ".dodavatel-picker-history-table th{background:#f0f1fb;color:#292982;font-weight:700;}" +
            ".dodavatel-picker-history-table td{color:#14162f;}" +
            ".dodavatel-picker-history-empty{padding:12px;color:#808184;font-size:12px;}" +
            ".dodavatel-picker-empty{padding:18px;color:#808184;background:#ffffff;}" +
            ".dodavatel-picker-footer{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 20px;border-top:1px solid rgba(41,41,130,.12);background:#f3f4f8;}" +
            ".dodavatel-picker-count{font-size:13px;color:#808184;}" +
            ".dodavatel-picker-actions{display:flex;gap:8px;}" +
            ".dodavatel-picker-button{min-width:120px;height:34px;padding:0 14px;border:1px solid #808184;background:#ffffff;color:#14162f;font-size:14px;cursor:pointer;}" +
            ".dodavatel-picker-button-primary{border-color:#292982;background:#292982;color:#ffffff;}" +
            ".dodavatel-picker-button:hover,.dodavatel-picker-button:focus{outline:1px solid #292982;outline-offset:1px;}" +
            ".dodavatel-picker-button-primary:hover,.dodavatel-picker-button-primary:focus{background:#1f1f63;}" +
            ".dodavatel-picker-attachment-button{" +
                "display:inline-block;min-width:120px;height:34px;padding:0 14px;" +
                "box-sizing:border-box;border:1px solid #292982;background:#292982;" +
                "color:#ffffff !important;-webkit-text-fill-color:#ffffff !important;" +
                "font-family:Segoe UI,Arial,sans-serif;" +
                "font-size:14px;font-weight:400;line-height:32px;" +
                "text-align:center;text-decoration:none;cursor:pointer;" +
            "}" +
            ".dodavatel-picker-attachment-button:visited,.dodavatel-picker-attachment-button:hover," +
            ".dodavatel-picker-attachment-button:focus,.dodavatel-picker-attachment-button:active{" +
                "background:#1f1f63;color:#ffffff !important;" +
                "-webkit-text-fill-color:#ffffff !important;outline:1px solid #292982;" +
                "outline-offset:1px;" +
            "}" +
            ".dodavatel-picker-attachment-button *{" +
                "color:#ffffff !important;-webkit-text-fill-color:#ffffff !important;" +
            "}" +
            ".dodavatel-picker-attachment-hint{" +
                "margin:0 0 10px 0;padding:8px 10px;border-left:3px solid #808184;" +
                "background:#f3f4f8;color:#808184;font-size:12px;line-height:1.5;" +
            "}" +
            ".dodavatel-picker-attachment-button{display:inline-block !important;visibility:visible !important;}" +
            ".dodavatel-picker-task-styled{" +
                "border-color:#292982 !important;background:#ffffff;" +
            "}" +
            ".dodavatel-picker-task-styled .ui-widget-header{" +
                "background:#292982;color:#ffffff;border-color:#292982;" +
            "}" +
            ".dodavatel-picker-task-attachment-row > td{" +
                "padding-top:10px;padding-bottom:10px;vertical-align:top;" +
            "}" +
            ".dodavatel-picker-task-attachment-row > td.ms-formlabel{" +
                "width:180px;padding-left:12px;padding-right:18px;color:#292982;font-weight:600;" +
                "white-space:nowrap;" +
            "}" +
            ".dodavatel-picker-task-attachment-row > td.ms-formbody{" +
                "width:auto;padding-left:18px;padding-right:0;text-align:left;" +
            "}" +
            ".dodavatel-picker-task-attachment-row .tispMultipleUploadFT{" +
                "max-width:100%;" +
            "}" +
            ".dodavatel-picker-task-attachment-row .tispMultipleUploadFT .buttons{" +
                "margin-top:8px;text-align:left;" +
            "}" +
            "@media screen and (max-width:520px){" +
                ".dodavatel-picker-overlay{padding:8px;}" +
                ".dodavatel-picker-dialog{margin:8px auto;}" +
                ".dodavatel-picker-body{padding:14px;}" +
                ".dodavatel-picker-footer{display:block;padding:12px 14px;}" +
                ".dodavatel-picker-count{display:block;margin-bottom:10px;}" +
                ".dodavatel-picker-actions{display:flex;}" +
                ".dodavatel-picker-button{flex:1;min-width:0;}" +
                ".dodavatel-picker-list-header,.dodavatel-picker-item{grid-template-columns:24px minmax(110px,1.2fr) minmax(90px,1fr) minmax(130px,1.1fr);}" +
            "}";

        if (styleAdded) {
            return;
        }

        style = document.createElement("style");
        style.type = "text/css";

        if (style.styleSheet) {
            style.styleSheet.cssText = css;
        } else {
            style.appendChild(document.createTextNode(css));
        }

        document.getElementsByTagName("head")[0].appendChild(style);
        styleAdded = true;
    }

    function normalizeSearchValue(value) {
        return String(value || "")
            .replace(/\s+/g, " ")
            .replace(/^\s+|\s+$/g, "")
            .toLowerCase();
    }

    function normalizeSupplierName(value) {
        return normalizeSearchValue(value);
    }

    function escapeODataString(value) {
        return String(value).replace(/'/g, "''");
    }

    function getInitialRestUrl() {
        var escapedListUrl = escapeODataString(CONFIG.listRelativeUrl);
        var select = "Id," + CONFIG.displayField + "," + CONFIG.buyerFieldInternalName;

        return CONFIG.siteUrl +
            "/_api/web/GetList('" +
            escapedListUrl +
            "')/items" +
            "?%24select=" + encodeURIComponent(select) +
            "&%24orderby=" + encodeURIComponent(CONFIG.displayField + " asc");
    }

    function getPurchaserRestUrl() {
        var escapedListUrl = escapeODataString(CONFIG.purchaserListRelativeUrl);
        var select = "Id," + CONFIG.purchaserTitleFieldInternalName + "," +
            CONFIG.purchaserEmailFieldInternalName;

        return CONFIG.siteUrl +
            "/_api/web/GetList('" +
            escapedListUrl +
            "')/items" +
            "?%24select=" + encodeURIComponent(select) +
            "&%24orderby=" + encodeURIComponent(
                CONFIG.purchaserTitleFieldInternalName + " asc"
            );
    }

    function getBuyerName(item) {
        var value;

        if (!item) {
            return "";
        }

        value = item[CONFIG.buyerFieldInternalName];

        if (value === null || typeof value === "undefined") {
            return "";
        }

        if (typeof value === "object") {
            return String(value.Title || value.Name || value.Description || "")
                .replace(/^\s+|\s+$/g, "");
        }

        return String(value).replace(/^\s+|\s+$/g, "");
    }

    function loadSupplierPage(url, callback) {
        var request = new XMLHttpRequest();
        var completed = false;
        var timeoutId;

        function finish(error, result) {
            if (completed) {
                return;
            }

            completed = true;
            window.clearTimeout(timeoutId);
            callback(error, result);
        }

        request.open("GET", url, true);
        request.setRequestHeader("Accept", "application/json;odata=verbose");

        request.onreadystatechange = function () {
            var data;
            var results;
            var nextUrl;

            if (request.readyState !== 4) {
                return;
            }

            if (request.status < 200 || request.status >= 300) {
                finish(new Error("REST API vrátilo HTTP status " + request.status));
                return;
            }

            try {
                data = JSON.parse(request.responseText);

                if (!data || !data.d || !data.d.results) {
                    throw new Error("Odpověď REST API nemá očekávaný formát.");
                }

                results = data.d.results;
                nextUrl = data.d.__next || null;

                finish(null, {
                    results: results,
                    nextUrl: nextUrl
                });
            } catch (error) {
                finish(error);
            }
        };

        request.onerror = function () {
            finish(new Error("REST API není dostupné."));
        };

        request.ontimeout = function () {
            finish(new Error("Požadavek na REST API vypršel."));
        };

        timeoutId = window.setTimeout(function () {
            try {
                request.abort();
            } catch (ignore) {
                /* Starší prohlížeč nemusí abort podporovat dokonale. */
            }

            finish(new Error("Požadavek na REST API vypršel."));
        }, CONFIG.requestTimeoutMs);

        try {
            request.send(null);
        } catch (error) {
            finish(error);
        }
    }

    function loadAllSuppliers(callback) {
        var suppliers = [];
        var seen = {};

        function loadNext(url) {
            loadSupplierPage(url, function (error, page) {
                var i;
                var item;
                var title;
                var buyer;
                var key;

                if (error) {
                    callback(error);
                    return;
                }

                for (i = 0; i < page.results.length; i += 1) {
                    item = page.results[i];
                    title = item ? item[CONFIG.displayField] : "";

                    if (title === null || typeof title === "undefined") {
                        if (window.console && window.console.error) {
                            window.console.error("Položka dodavatele nemá vyplněné pole Title.", item);
                        }
                        continue;
                    }

                    title = String(title).replace(/^\s+|\s+$/g, "");

                    if (!title) {
                        if (window.console && window.console.error) {
                            window.console.error("Položka dodavatele má prázdné pole Title.", item);
                        }
                        continue;
                    }

                    key = normalizeSupplierName(title);
                    buyer = getBuyerName(item);

                    if (!seen[key]) {
                        seen[key] = true;
                        suppliers.push({
                            id: item.Id,
                            title: title,
                            buyer: buyer
                        });
                    }
                }

                if (page.nextUrl) {
                    loadNext(page.nextUrl);
                } else {
                    callback(null, suppliers);
                }
            });
        }

        loadNext(getInitialRestUrl());
    }

    function loadAllPurchasers(callback) {
        var purchasers = {};

        function loadNext(url) {
            loadSupplierPage(url, function (error, page) {
                var i;
                var item;
                var shortName;
                var email;

                if (error) {
                    callback(error);
                    return;
                }

                for (i = 0; i < page.results.length; i += 1) {
                    item = page.results[i];
                    shortName = item ? item[CONFIG.purchaserTitleFieldInternalName] : "";
                    email = item ? item[CONFIG.purchaserEmailFieldInternalName] : "";

                    if (shortName === null || typeof shortName === "undefined") {
                        continue;
                    }

                    shortName = String(shortName).replace(/^\s+|\s+$/g, "");

                    if (typeof email === "object" && email !== null) {
                        email = email.EMail || email.Email || email.Value || "";
                    }

                    email = String(email || "").replace(/^\s+|\s+$/g, "");

                    if (shortName) {
                        purchasers[normalizeSupplierName(shortName)] = email;
                    }
                }

                if (page.nextUrl) {
                    loadNext(page.nextUrl);
                } else {
                    callback(null, purchasers);
                }
            });
        }

        loadNext(getPurchaserRestUrl());
    }

    function createElement(tagName, className, text) {
        var element = document.createElement(tagName);

        if (className) {
            element.className = className;
        }

        if (typeof text !== "undefined") {
            element.textContent = text;
        }

        return element;
    }

    function readExistingSelection(field) {
        var selected = {};
        var parts;
        var i;
        var name;

        if (!field) {
            return selected;
        }

        parts = String(field.value || "").split(";");

        for (i = 0; i < parts.length; i += 1) {
            name = parts[i].replace(/^\s+|\s+$/g, "");

            if (name) {
                selected[normalizeSupplierName(name)] = true;
            }
        }

        return selected;
    }

    function updateSelectedCount() {
        var state = dialogState;
        var count = 0;
        var i;

        if (!state || !state.countElement) {
            return;
        }

        for (i = 0; i < state.suppliers.length; i += 1) {
            if (state.selected[normalizeSupplierName(state.suppliers[i].title)]) {
                count += 1;
            }
        }

        state.countElement.textContent = "Vybráno dodavatelů: " + count;
    }

    function renderSupplierList() {
        var state = dialogState;
        var search;
        var supplier;
        var label;
        var checkbox;
        var visibleCount = 0;
        var i;

        if (!state || !state.listElement) {
            return;
        }

        search = normalizeSearchValue(state.searchElement.value);
        while (state.listElement.firstChild) {
            state.listElement.removeChild(state.listElement.firstChild);
        }

        if (state.suppliers.length) {
            label = createElement("div", "dodavatel-picker-list-header");
            label.appendChild(createElement("span", "", ""));
            label.appendChild(createElement("span", "", "Dodavatel"));
            label.appendChild(createElement("span", "", CONFIG.buyerColumnLabel));
            label.appendChild(createElement("span", "", CONFIG.purchaserEmailColumnLabel));
            state.listElement.appendChild(label);
        }

        for (i = 0; i < state.suppliers.length; i += 1) {
            supplier = state.suppliers[i];

            if (search && normalizeSearchValue(supplier.title).indexOf(search) === -1) {
                continue;
            }

            label = createElement("label", "dodavatel-picker-item");
            checkbox = document.createElement("input");
            checkbox.type = "checkbox";
            checkbox.value = supplier.title;
            checkbox.checked = !!state.selected[normalizeSupplierName(supplier.title)];

            addEvent(checkbox, "change", function () {
                var key = normalizeSupplierName(this.value);

                if (this.checked) {
                    state.selected[key] = true;
                } else {
                    delete state.selected[key];
                }

                updateSelectedCount();
            });

            label.appendChild(checkbox);
            label.appendChild(createElement("span", "dodavatel-picker-supplier-name", supplier.title));
            label.appendChild(createElement(
                "span",
                "dodavatel-picker-buyer-name",
                supplier.buyer || ""
            ));
            label.appendChild(createElement(
                "span",
                "dodavatel-picker-email",
                supplier.email || ""
            ));
            state.listElement.appendChild(label);
            visibleCount += 1;
        }

        if (!visibleCount) {
            state.listElement.appendChild(createElement(
                "div",
                "dodavatel-picker-empty",
                state.suppliers.length ?
                    "Vyhledávání neodpovídá žádnému dodavateli." :
                    "Seznam dodavatelů neobsahuje žádné položky."
            ));
        }

        updateSelectedCount();
    }

    function filterSupplierList() {
        renderSupplierList();
    }

    function showLoadingState() {
        if (!dialogState) {
            return;
        }

        dialogState.statusElement.className = "dodavatel-picker-status";
        dialogState.statusElement.textContent = "Načítám dodavatele…";

        while (dialogState.listElement.firstChild) {
            dialogState.listElement.removeChild(dialogState.listElement.firstChild);
        }
    }

    function showError(error) {
        var technicalMessage = error && error.message ? error.message : String(error || "Neznámá chyba.");

        if (window.console && window.console.error) {
            window.console.error("Chyba při načítání dodavatelů:", technicalMessage);
        }

        if (!dialogState) {
            return;
        }

        dialogState.statusElement.className = "dodavatel-picker-status dodavatel-picker-error";
        dialogState.statusElement.textContent = "Dodavatele se nepodařilo načíst. Zkontrolujte připojení a svá oprávnění k seznamu.";

        while (dialogState.listElement.firstChild) {
            dialogState.listElement.removeChild(dialogState.listElement.firstChild);
        }
    }

    function dispatchCompatibleEvent(element, eventName) {
        var eventObject;

        if (!element) {
            return;
        }

        try {
            if (document.createEvent) {
                eventObject = document.createEvent("Event");
                eventObject.initEvent(eventName, true, true);
                element.dispatchEvent(eventObject);
            } else if (document.createEventObject) {
                eventObject = document.createEventObject();
                eventObject.bubbles = true;
                eventObject.cancelable = true;
                element.fireEvent("on" + eventName, eventObject);
            }
        } catch (error) {
            if (window.console && window.console.error) {
                window.console.error("Nepodařilo se vyvolat událost " + eventName + ".", error);
            }
        }
    }

    function setTargetFieldValue(field, value) {
        if (!field) {
            return;
        }

        field.value = value;
        dispatchCompatibleEvent(field, "input");
        dispatchCompatibleEvent(field, "change");
        dispatchCompatibleEvent(field, "blur");
    }

    function confirmSelection() {
        var state = dialogState;
        var result = [];
        var emailResult = [];
        var seen = {};
        var emailSeen = {};
        var purchaserField;
        var i;
        var supplier;
        var key;

        if (!state || !state.suppliers) {
            return;
        }

        for (i = 0; i < state.suppliers.length; i += 1) {
            supplier = state.suppliers[i];
            key = normalizeSupplierName(supplier.title);

            if (state.selected[key] && !seen[key]) {
                seen[key] = true;
                result.push(supplier.title);

                if (supplier.email && !emailSeen[normalizeSearchValue(supplier.email)]) {
                    emailSeen[normalizeSearchValue(supplier.email)] = true;
                    emailResult.push(supplier.email);
                }
            }
        }

        setTargetFieldValue(state.field, result.join(CONFIG.separator));

        purchaserField = findFieldByInternalNames(
            CONFIG.purchaserTargetFieldInternalNames
        );

        if (purchaserField) {
            setTargetFieldValue(
                purchaserField,
                emailResult.join(CONFIG.separator)
            );
        } else if (window.console && window.console.error) {
            window.console.error(
                "Pole nakupci_3lp nebylo nalezeno; emaily nebyly zapsány."
            );
        }

        appendHistoryEntry(
            "Potvrzen výběr dodavatelů: " +
            (result.length ? result.join(CONFIG.separator) : "bez dodavatele") +
            "; nákupčí emaily: " +
            (emailResult.length ? emailResult.join(CONFIG.separator) : "bez emailu")
        );

        closeSupplierDialog();
    }

    function createFocusTrap(event) {
        var state = dialogState;
        var inputs;
        var buttons;
        var all = [];
        var first;
        var last;
        var i;

        if (!state || !state.dialogElement) {
            return;
        }

        event = event || window.event;

        if (event.keyCode === 27) {
            closeSupplierDialog();
            return;
        }

        if (event.keyCode !== 9) {
            return;
        }

        inputs = state.dialogElement.getElementsByTagName("input");
        buttons = state.dialogElement.getElementsByTagName("button");

        for (i = 0; i < inputs.length; i += 1) {
            if (!inputs[i].disabled) {
                all.push(inputs[i]);
            }
        }

        for (i = 0; i < buttons.length; i += 1) {
            if (!buttons[i].disabled) {
                all.push(buttons[i]);
            }
        }

        if (!all.length) {
            return;
        }

        first = all[0];
        last = all[all.length - 1];

        if (event.shiftKey && document.activeElement === first) {
            event.preventDefault();
            last.focus();
        } else if (!event.shiftKey && document.activeElement === last) {
            event.preventDefault();
            first.focus();
        }
    }

    function closeSupplierDialog() {
        var state = dialogState;

        if (!state) {
            return;
        }

        removeEvent(document, "keydown", state.keydownHandler);

        if (state.overlayElement && state.overlayElement.parentNode) {
            state.overlayElement.parentNode.removeChild(state.overlayElement);
        }

        dialogState = null;

        if (state.field && state.field.focus) {
            try {
                state.field.focus();
            } catch (ignore) {
                /* Pole může být mezitím odstraněno částečným načtením. */
            }
        }
    }

    function openSupplierDialog(field) {
        var overlay;
        var dialog;
        var header;
        var title;
        var body;
        var search;
        var status;
        var list;
        var footer;
        var count;
        var actions;
        var cancelButton;
        var confirmButton;
        var keydownHandler;

        if (dialogState || !field || !isSupplierPickerEnabled()) {
            return;
        }

        injectStyles();

        overlay = createElement("div", "dodavatel-picker-overlay");
        dialog = createElement("div", "dodavatel-picker-dialog");
        dialog.setAttribute("role", "dialog");
        dialog.setAttribute("aria-modal", "true");
        dialog.setAttribute("aria-labelledby", "dodavatel-picker-title");

        header = createElement("div", "dodavatel-picker-header");
        title = createElement("h2", "dodavatel-picker-title", "Výběr dodavatelů");
        title.id = "dodavatel-picker-title";

        header.appendChild(title);

        body = createElement("div", "dodavatel-picker-body");
        search = createElement("input", "dodavatel-picker-search");
        search.type = "search";
        search.placeholder = "Vyhledat dodavatele…";
        search.setAttribute("aria-label", "Vyhledat dodavatele");

        status = createElement("div", "dodavatel-picker-status");
        list = createElement("div", "dodavatel-picker-list");
        list.setAttribute("aria-label", "Seznam dodavatelů");

        footer = createElement("div", "dodavatel-picker-footer");
        count = createElement("div", "dodavatel-picker-count", "Vybráno dodavatelů: 0");
        actions = createElement("div", "dodavatel-picker-actions");

        cancelButton = createElement("button", "dodavatel-picker-button", "Zrušit");
        cancelButton.type = "button";
        confirmButton = createElement("button", "dodavatel-picker-button dodavatel-picker-button-primary", "Potvrdit výběr");
        confirmButton.type = "button";

        actions.appendChild(cancelButton);
        actions.appendChild(confirmButton);
        footer.appendChild(count);
        footer.appendChild(actions);
        body.appendChild(search);
        body.appendChild(status);
        body.appendChild(list);
        dialog.appendChild(header);
        dialog.appendChild(body);
        dialog.appendChild(footer);
        overlay.appendChild(dialog);
        document.body.appendChild(overlay);

        keydownHandler = createFocusTrap;

        dialogState = {
            field: field,
            overlayElement: overlay,
            dialogElement: dialog,
            searchElement: search,
            statusElement: status,
            listElement: list,
            countElement: count,
            suppliers: [],
            selected: readExistingSelection(field),
            keydownHandler: keydownHandler
        };

        addEvent(cancelButton, "click", closeSupplierDialog);
        addEvent(confirmButton, "click", confirmSelection);
        addEvent(search, "input", filterSupplierList);
        addEvent(search, "keyup", filterSupplierList);
        addEvent(document, "keydown", keydownHandler);

        showLoadingState();
        search.focus();

        loadAllPurchasers(function (purchaserError, purchaserEmails) {
            if (!dialogState || dialogState.overlayElement !== overlay) {
                return;
            }

            if (purchaserError) {
                showError(purchaserError);
                return;
            }

            loadAllSuppliers(function (supplierError, suppliers) {
                var i;

                if (!dialogState || dialogState.overlayElement !== overlay) {
                    return;
                }

                if (supplierError) {
                    showError(supplierError);
                    return;
                }

                for (i = 0; i < suppliers.length; i += 1) {
                    suppliers[i].email = purchaserEmails[
                        normalizeSupplierName(suppliers[i].buyer)
                    ] || "";
                }

                dialogState.suppliers = suppliers;
                dialogState.statusElement.className = "dodavatel-picker-status";
                dialogState.statusElement.textContent = suppliers.length ?
                    "Vyberte jednoho nebo více dodavatelů." :
                    "Seznam dodavatelů neobsahuje žádné položky.";
                renderSupplierList();
            });
        });
    }

    function bindTargetField(field) {
        var clickHandler;
        var keydownHandler;

        if (!field || getAttribute(field, "data-dodavatel-picker-bound") === "true") {
            return;
        }

        field.setAttribute("data-dodavatel-picker-bound", "true");

        clickHandler = function (event) {
            event = event || window.event;

            if (dialogState) {
                return;
            }

            if (event.preventDefault) {
                event.preventDefault();
            }

            openSupplierDialog(field);
        };

        keydownHandler = function (event) {
            event = event || window.event;

            if (event.keyCode === 13 || event.keyCode === 32) {
                if (event.preventDefault) {
                    event.preventDefault();
                }

                openSupplierDialog(field);
            }
        };

        addEvent(field, "click", clickHandler);
        addEvent(field, "keydown", keydownHandler);
    }

    function handleDelegatedClick(event) {
        var target;
        var field;
        var current;

        if (dialogState) {
            return;
        }

        event = event || window.event;
        target = event.target || event.srcElement;
        field = findTargetField();

        if (!target || !field) {
            return;
        }

        current = target;

        while (current && current !== document.body) {
            if (current === field) {
                if (event.preventDefault) {
                    event.preventDefault();
                }

                openSupplierDialog(field);
                return;
            }

            current = current.parentNode;
        }
    }

    function isSupportedFormPage() {
        var path = String(window.location.pathname || "").toLowerCase();
        var query = String(window.location.search || "").toLowerCase();

        return /\/(newform|editform|dispform)\.aspx$/.test(path) ||
            (/\/listform\.aspx$/.test(path) &&
                /(?:^|[?&])pagetype=(?:6|8)(?:&|$)/.test(query));
    }

    function isDisplayFormPage() {
        return /\/dispform\.aspx$/.test(
            String(window.location.pathname || "").toLowerCase()
        );
    }

    function initializeSupplierPicker() {
        // Picker se aktivuje pouze na formulářích pro nový nebo upravovaný záznam.
        if (!isSupportedFormPage()) {
            return;
        }

        if (isDisplayFormPage()) {
            initializeHistoryDisplay();
            return;
        }

        loadCurrentUser();

        if (!delegatedClickBound) {
            addEvent(document, "click", handleDelegatedClick);
            delegatedClickBound = true;
        }

        var field = findTargetField();

        if (field) {
            bindTargetField(field);
        }

        injectStyles();
        bindConditionField();
        bindAttachmentSourceChanges();
        updateAttachmentFieldAvailability();
        customizeAttachmentControls();
        moveTaskAttachmentsAboveOutcome();
        initializeAuditLogging();

        if (!observer && window.MutationObserver && document.body) {
            observer = new MutationObserver(function () {
                var currentField = findTargetField();

                if (currentField) {
                    bindTargetField(currentField);
                }

                bindConditionField();
                bindAttachmentSourceChanges();
                updateAttachmentFieldAvailability();
                customizeAttachmentControls();
                moveTaskAttachmentsAboveOutcome();
                auditFieldValueChanges();
                auditAttachmentChanges();
            });

            observer.observe(document.body, {
                childList: true,
                subtree: true
            });
        }

        if (!fallbackInterval) {
            fallbackInterval = window.setInterval(function () {
                var currentField = findTargetField();

                if (currentField) {
                    bindTargetField(currentField);
                }

                bindConditionField();
                bindAttachmentSourceChanges();
                updateAttachmentFieldAvailability();
                customizeAttachmentControls();
                moveTaskAttachmentsAboveOutcome();
                auditFieldValueChanges();
                auditAttachmentChanges();
            }, 1000);

            window.setTimeout(function () {
                if (fallbackInterval) {
                    window.clearInterval(fallbackInterval);
                    fallbackInterval = null;
                }
            }, 30000);
        }
    }

    if (document.readyState === "loading") {
        addEvent(document, "DOMContentLoaded", initializeSupplierPicker);
    } else {
        initializeSupplierPicker();
    }

    addEvent(window, "load", initializeSupplierPicker);

    if (window.Sys && window.Sys.Application && window.Sys.Application.add_load) {
        window.Sys.Application.add_load(initializeSupplierPicker);
    }

    if (window._spBodyOnLoadFunctionNames) {
        window._spBodyOnLoadFunctionNames.push("initializeSupplierPicker");
        window.initializeSupplierPicker = initializeSupplierPicker;
    }
})();
