(function () {
	"use strict";

	function isDisplayForm() {
		var path = (window.location && window.location.pathname ? window.location.pathname : "").toLowerCase();
		return path.indexOf("dispform.aspx") !== -1;
	}

	function getAttachmentsField() {
		return document.getElementById("SPFieldAttachments");
	}

	function injectStyle() {
		if (document.getElementById("ti-attachments-legacy-style")) {
			return;
		}

		var style = document.createElement("style");
		style.id = "ti-attachments-legacy-style";
		style.type = "text/css";
		style.textContent =
			"#SPFieldAttachments.ti-attachments-legacy{" +
			"background:#f3f3f3;" +
			"border:1px solid #c8c8c8;" +
			"border-radius:4px;" +
			"padding:10px 12px;" +
			"}" +
			"#ti-attachments-legacy-note{" +
			"margin:0 0 8px 0;" +
			"padding:6px 8px;" +
			"background:#e9e9e9;" +
			"border-left:3px solid #8a8a8a;" +
			"font-size:12px;" +
			"line-height:1.4;" +
			"color:#333;" +
			"font-style:italic;" +
			"}";

		document.head.appendChild(style);
	}

	function ensureVisible() {
		var row = document.getElementById("idAttachmentsRow");
		var field = getAttachmentsField();
		var table = document.getElementById("idAttachmentsTable");

		if (row) {
			row.style.display = "table-row";
			row.style.visibility = "visible";
		}

		if (field) {
			field.style.display = "";
			field.style.visibility = "visible";
			field.hidden = false;
		}

		if (table) {
			table.style.display = "";
			table.style.visibility = "visible";
			table.hidden = false;
		}
	}

	function addLegacyNote() {
		var field = getAttachmentsField();
		if (!field) {
			return false;
		}

		field.classList.add("ti-attachments-legacy");

		if (document.getElementById("ti-attachments-legacy-note")) {
			return true;
		}

		var note = document.createElement("div");
		note.id = "ti-attachments-legacy-note";
		note.textContent = "Poznamka: Jedna se o starsi proces s prilohami.";

		if (field.firstChild) {
			field.insertBefore(note, field.firstChild);
		} else {
			field.appendChild(note);
		}

		return true;
	}

	function init() {
		if (!isDisplayForm()) {
			return;
		}

		injectStyle();

		var tries = 0;
		var timer = window.setInterval(function () {
			tries += 1;
			ensureVisible();
			if (addLegacyNote() || tries >= 20) {
				window.clearInterval(timer);
			}
		}, 250);

		ensureVisible();
		addLegacyNote();
	}

	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", init);
	} else {
		init();
	}
})();
