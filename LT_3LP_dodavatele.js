(function () {
    "use strict";

    var CONFIG = {
        fieldInternalName: "dodavatel_3pl",
        alternateFieldInternalNames: ["dodavatele_3lp"],
        siteUrl: "http://portal.samohyl.cz/nakup",
        listTitle: "Dodavatele",
        listRelativeUrl: "/nakup/Lists/Dodavatele",
        displayField: "Title",
        separator: "; ",
        requestTimeoutMs: 30000
    };

    var dialogState = null;
    var observer = null;
    var fallbackInterval = null;
    var styleAdded = false;
    var delegatedClickBound = false;

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
                "background:linear-gradient(90deg,#292982 0%,#34349a 62%,#808184 100%);" +
                "color:#ffffff;border-bottom:1px solid #1f1f63;" +
            "}" +
            ".dodavatel-picker-title{margin:0;font-size:20px;font-weight:600;}" +
            ".dodavatel-picker-close{" +
                "position:absolute;top:10px;right:12px;width:34px;height:34px;border:0;" +
                "background:transparent;color:#fff;font-size:26px;line-height:30px;cursor:pointer;" +
            "}" +
            ".dodavatel-picker-close:hover,.dodavatel-picker-close:focus{" +
                "background:#1f1f63;outline:1px solid #fff;" +
            "}" +
            ".dodavatel-picker-body{padding:20px;background:#ffffff;}" +
            ".dodavatel-picker-search{" +
                "display:block;width:100%;height:36px;padding:7px 10px;border:1px solid #808184;" +
                "box-sizing:border-box;font-size:14px;" +
            "}" +
            ".dodavatel-picker-search:focus{border-color:#292982;outline:1px solid #292982;}" +
            ".dodavatel-picker-status{min-height:22px;margin:10px 0;font-size:13px;color:#808184;}" +
            ".dodavatel-picker-error{padding:10px;border:1px solid #e01b37;background:#fff0f2;color:#a61f2c;line-height:1.4;}" +
            ".dodavatel-picker-list{height:320px;overflow-y:auto;border:1px solid rgba(41,41,130,.16);background:#f3f4f8;}" +
            ".dodavatel-picker-item{display:block;padding:9px 12px;border-bottom:1px solid rgba(41,41,130,.12);cursor:pointer;font-size:14px;line-height:20px;background:#ffffff;}" +
            ".dodavatel-picker-item:hover{background:#f3f4f8;}" +
            ".dodavatel-picker-item input{margin:0 9px 0 0;vertical-align:middle;}" +
            ".dodavatel-picker-empty{padding:18px;color:#808184;background:#ffffff;}" +
            ".dodavatel-picker-footer{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:14px 20px;border-top:1px solid rgba(41,41,130,.12);background:#f3f4f8;}" +
            ".dodavatel-picker-count{font-size:13px;color:#808184;}" +
            ".dodavatel-picker-actions{display:flex;gap:8px;}" +
            ".dodavatel-picker-button{min-width:120px;height:34px;padding:0 14px;border:1px solid #808184;background:#ffffff;color:#14162f;font-size:14px;cursor:pointer;}" +
            ".dodavatel-picker-button-primary{border-color:#292982;background:#292982;color:#ffffff;}" +
            ".dodavatel-picker-button:hover,.dodavatel-picker-button:focus{outline:1px solid #292982;outline-offset:1px;}" +
            ".dodavatel-picker-button-primary:hover,.dodavatel-picker-button-primary:focus{background:#1f1f63;}" +
            "@media screen and (max-width:520px){" +
                ".dodavatel-picker-overlay{padding:8px;}" +
                ".dodavatel-picker-dialog{margin:8px auto;}" +
                ".dodavatel-picker-body{padding:14px;}" +
                ".dodavatel-picker-footer{display:block;padding:12px 14px;}" +
                ".dodavatel-picker-count{display:block;margin-bottom:10px;}" +
                ".dodavatel-picker-actions{display:flex;}" +
                ".dodavatel-picker-button{flex:1;min-width:0;}" +
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

        return CONFIG.siteUrl +
            "/_api/web/GetList('" +
            escapedListUrl +
            "')/items" +
            "?%24select=Id%2CTitle&%24orderby=Title%20asc";
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

                    if (!seen[key]) {
                        seen[key] = true;
                        suppliers.push({
                            id: item.Id,
                            title: title
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
            label.appendChild(document.createTextNode(supplier.title));
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
        var seen = {};
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
            }
        }

        setTargetFieldValue(state.field, result.join(CONFIG.separator));
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
        var closeButton;
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

        if (dialogState || !field) {
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

        closeButton = createElement("button", "dodavatel-picker-close", "×");
        closeButton.type = "button";
        closeButton.setAttribute("aria-label", "Zavřít");
        addEvent(closeButton, "click", closeSupplierDialog);

        header.appendChild(title);
        header.appendChild(closeButton);

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

        loadAllSuppliers(function (error, suppliers) {
            if (!dialogState || dialogState.overlayElement !== overlay) {
                return;
            }

            if (error) {
                showError(error);
                return;
            }

            dialogState.suppliers = suppliers;
            dialogState.statusElement.className = "dodavatel-picker-status";
            dialogState.statusElement.textContent = suppliers.length ?
                "Vyberte jednoho nebo více dodavatelů." :
                "Seznam dodavatelů neobsahuje žádné položky.";
            renderSupplierList();
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

        return /\/(newform|editform)\.aspx$/.test(path) ||
            (/\/listform\.aspx$/.test(path) &&
                /(?:^|[?&])pagetype=(?:6|8)(?:&|$)/.test(query));
    }

    function initializeSupplierPicker() {
        // Picker se aktivuje pouze na formulářích pro nový nebo upravovaný záznam.
        if (!isSupportedFormPage()) {
            return;
        }

        if (!delegatedClickBound) {
            addEvent(document, "click", handleDelegatedClick);
            delegatedClickBound = true;
        }

        var field = findTargetField();

        if (field) {
            bindTargetField(field);
        }

        if (!observer && window.MutationObserver && document.body) {
            observer = new MutationObserver(function () {
                var currentField = findTargetField();

                if (currentField) {
                    bindTargetField(currentField);
                }
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
