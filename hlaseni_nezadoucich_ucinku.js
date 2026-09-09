(function () {
	"use strict";

	if (typeof window === "undefined" || typeof document === "undefined") {
		return;
	}

	var SUKL_FIELDS_TO_LOCK = [
		"nazev_sarze_exspirace_pripravku",
		"registracni_cislo_pripravku",
		"jmeno_prijmeni_odesilatele",
		"adresa_odesilatele",
		"datum_odeslani_odesilatel",
		"jmeno_prijmeni_odhaleni_nu",
		"adresa_odhaleni_nu",
		"datum_odhaleni_nu",
		"jmeno_prijmeni_lekar",
		"adresa_lekar",
		"datum_informovani_lekare",
		"drzitel_informovan",
		"duvod_pouziti",
		"pocet_zvirat_pouzity_pripravek",
		"pocet_zvirat_s_nu",
		"pocet_uhynulych_zvirat",
		"podane_mnozstvi",
		"podavajici_osoba_pripravek",
		"datum_podani_pripravku",
		"delka_lecby",
		"misto_a_cesta_podani",
		"byl_pripravek_podan_jiz_drive",
		"pokud_ano_kolikrat",
		"datum_vyskytu_nu_tab",
		"druh_plemeno_tab",
		"hmotnost_kg_tab",
		"vek",
		"pohlavi_tab",
		"povaha_nu_tab",
		"lecba_nu",
		"predchozi_vakcinace",
		"stari_zvirete",
		"primovakcinace",
		"_x0031__reaktivace",
		"dalsi_revakcinace",
		"postmortalni_vysetreni",
		"nestaci_prostor"
	];

	function isNewFormPage() {
		return (
			window.location &&
			typeof window.location.pathname === "string" &&
			/newform\.aspx/i.test(window.location.pathname)
		);
	}

	function forEachNode(nodeList, callback) {
		var i;
		for (i = 0; i < nodeList.length; i += 1) {
			callback(nodeList[i], i);
		}
	}

	function removeElement(element) {
		if (element && element.parentNode) {
			element.parentNode.removeChild(element);
		}
	}

	function normalizeText(value) {
		if (!value) {
			return "";
		}

		var text = String(value).toLowerCase();
		if (typeof text.normalize === "function") {
			text = text.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
		}

		return text.replace(/\s+/g, " ").trim();
	}

	function getOwnText(element) {
		var text = "";
		var childNodes = element.childNodes;

		forEachNode(childNodes, function (node) {
			if (node && node.nodeType === 3) {
				text += node.nodeValue || "";
			}
		});

		return normalizeText(text);
	}

	function strikeMainHeader() {
		var target = "formular k hlaseni nu";
		var candidates = document.querySelectorAll(
			"h1, h2, h3, h4, .ms-core-pageTitle, .ms-core-pageTitle a, .ms-standardheader, .ms-breadcrumb-title, nobr"
		);
		var bestMatch = null;
		var bestLength = 999999;

		forEachNode(candidates, function (el) {
			var ownText = getOwnText(el);
			var fullText = normalizeText(el.textContent || el.innerText || "");
			var checkText = ownText || fullText;

			if (
				checkText.indexOf(target) !== -1 ||
				(checkText.indexOf("formular") !== -1 &&
					checkText.indexOf("hlaseni") !== -1 &&
					checkText.indexOf("nu") !== -1)
			) {
				if (checkText.length < bestLength) {
					bestMatch = el;
					bestLength = checkText.length;
				}
			}
		});

		if (bestMatch) {
			bestMatch.classList.add("sukl-main-header-locked");
		}
	}

	function lockFieldsWithoutSukl() {
		function lockControl(field) {
			var tagName = (field.tagName || "").toLowerCase();

			if (tagName === "select") {
				field.disabled = true;
			} else if (tagName === "input" || tagName === "textarea") {
				field.readOnly = true;
				field.disabled = true;
			}

			field.classList.add("sukl-field-locked");
			field.setAttribute("data-sukl-locked", "1");
		}

		function disableDatePickerButtons(container) {
			var datePickerButtons = container.querySelectorAll('a[onclick*="clickDatePicker"]');
			forEachNode(datePickerButtons, function (button) {
				button.removeAttribute("onclick");
				button.setAttribute("href", "javascript:void(0)");
				button.setAttribute("tabindex", "-1");
				button.classList.add("sukl-datepicker-locked");
			});
		}

		SUKL_FIELDS_TO_LOCK.forEach(function (internalName) {
			var fields = document.querySelectorAll(
				'[title="' + internalName + '"], [id^="' + internalName + '_"]'
			);

			forEachNode(fields, function (field) {
				lockControl(field);
			});

			var row = document.querySelector(".tisa-form-row" + internalName);
			if (row) {
				row.classList.add("sukl-row-locked");

				var rowFields = row.querySelectorAll("input, select, textarea");
				forEachNode(rowFields, function (rowField) {
					lockControl(rowField);
				});

				disableDatePickerButtons(row);
			}
		});
	}

	function injectStyles() {
		if (document.getElementById("sukl-modal-styles")) {
			return;
		}

		var style = document.createElement("style");
		style.id = "sukl-modal-styles";
		style.type = "text/css";
		style.textContent =
			"#sukl-modal-overlay{" +
			"position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99998;" +
			"display:flex;align-items:center;justify-content:center;padding:16px;box-sizing:border-box;" +
			"}" +
			"#sukl-info-overlay{" +
			"position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99999;" +
			"display:flex;align-items:center;justify-content:center;padding:16px;box-sizing:border-box;" +
			"}" +
			"#sukl-modal{" +
			"width:100%;max-width:520px;background:#fff;border-radius:10px;" +
			"box-shadow:0 18px 45px rgba(0,0,0,.25);padding:24px 20px 18px;font-family:Segoe UI,Tahoma,Arial,sans-serif;" +
			"}" +
			"#sukl-info-modal{" +
			"width:100%;max-width:520px;background:#fff;border-radius:10px;" +
			"box-shadow:0 18px 45px rgba(0,0,0,.25);padding:24px 20px 18px;font-family:Segoe UI,Tahoma,Arial,sans-serif;" +
			"}" +
			"#sukl-modal h2{margin:0 0 10px;font-size:20px;line-height:1.3;color:#1f1f1f;}" +
			"#sukl-modal p{margin:0 0 16px;font-size:14px;color:#444;line-height:1.45;}" +
			"#sukl-info-modal h2{margin:0 0 10px;font-size:20px;line-height:1.3;color:#1f1f1f;}" +
			"#sukl-info-modal p{margin:0 0 16px;font-size:14px;color:#444;line-height:1.45;}" +
			"#sukl-modal-actions{display:flex;gap:10px;flex-wrap:wrap;}" +
			"#sukl-info-actions{display:flex;gap:10px;flex-wrap:wrap;}" +
			"#sukl-modal-actions button{" +
			"border:0;cursor:pointer;border-radius:6px;padding:10px 14px;font-size:14px;font-weight:600;" +
			"}" +
			"#sukl-info-actions button{" +
			"border:0;cursor:pointer;border-radius:6px;padding:10px 14px;font-size:14px;font-weight:600;" +
			"}" +
			".sukl-field-locked{background:#f3f3f3 !important;color:#6a6a6a !important;cursor:not-allowed !important;}" +
			".sukl-datepicker-locked{pointer-events:none !important;opacity:.45 !important;cursor:not-allowed !important;}" +
			".sukl-row-locked .ms-formlabel h3,.sukl-row-locked .ms-standardheader{color:#6a6a6a !important;text-decoration:line-through !important;}" +
			".sukl-main-header-locked{text-decoration:line-through !important;color:#6a6a6a !important;}" +
			"#sukl-with-report{background:#0c6a3e;color:#fff;}" +
			"#sukl-without-report{background:#d83b01;color:#fff;}" +
			"#sukl-info-ok{background:#005a9e;color:#fff;}";

		document.head.appendChild(style);
	}

	function createModal() {
		if (document.getElementById("sukl-modal-overlay")) {
			return;
		}

		var overlay = document.createElement("div");
		overlay.id = "sukl-modal-overlay";

		var modal = document.createElement("div");
		modal.id = "sukl-modal";
		modal.setAttribute("role", "dialog");
		modal.setAttribute("aria-modal", "true");
		modal.setAttribute("aria-labelledby", "sukl-modal-title");

		var title = document.createElement("h2");
		title.id = "sukl-modal-title";
		title.textContent = "Vyberte typ hlášení";

		var description = document.createElement("p");
		description.textContent =
			"Zvolte, jestli se jedná o formulář s hlášením SÚKLu, nebo bez hlášení SÚKLu.";

		var actions = document.createElement("div");
		actions.id = "sukl-modal-actions";

		var withReportButton = document.createElement("button");
		withReportButton.id = "sukl-with-report";
		withReportButton.type = "button";
		withReportButton.textContent = "S hlášením SÚKLu";

		var withoutReportButton = document.createElement("button");
		withoutReportButton.id = "sukl-without-report";
		withoutReportButton.type = "button";
		withoutReportButton.textContent = "Bez hlášení SÚKLu";

		withReportButton.addEventListener("click", function () {
			removeElement(overlay);
		});

		withoutReportButton.addEventListener("click", function () {
			lockFieldsWithoutSukl();
			strikeMainHeader();
			removeElement(overlay);
			createNoSuklInfoModal();
		});

		actions.appendChild(withReportButton);
		actions.appendChild(withoutReportButton);

		modal.appendChild(title);
		modal.appendChild(description);
		modal.appendChild(actions);
		overlay.appendChild(modal);

		document.body.appendChild(overlay);
	}

	function createNoSuklInfoModal() {
		if (document.getElementById("sukl-info-overlay")) {
			return;
		}

		var overlay = document.createElement("div");
		overlay.id = "sukl-info-overlay";

		var modal = document.createElement("div");
		modal.id = "sukl-info-modal";
		modal.setAttribute("role", "dialog");
		modal.setAttribute("aria-modal", "true");
		modal.setAttribute("aria-labelledby", "sukl-info-title");

		var title = document.createElement("h2");
		title.id = "sukl-info-title";
		title.textContent = "Bez hlášení SÚKLu";

		var message = document.createElement("p");
		message.textContent =
			"Byla zvolena možnost Bez hlášení SÚKLu. Formulář k hlášení NÚ nelze vyplnit.";

		var actions = document.createElement("div");
		actions.id = "sukl-info-actions";

		var okButton = document.createElement("button");
		okButton.id = "sukl-info-ok";
		okButton.type = "button";
		okButton.textContent = "OK";

		okButton.addEventListener("click", function () {
			removeElement(overlay);
		});

		actions.appendChild(okButton);
		modal.appendChild(title);
		modal.appendChild(message);
		modal.appendChild(actions);
		overlay.appendChild(modal);
		document.body.appendChild(overlay);
	}

	function init() {
		if (!isNewFormPage()) {
			return;
		}

		injectStyles();
		createModal();
	}

	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", init);
	} else {
		init();
	}
})();

