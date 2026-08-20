(function () {
	"use strict";

	function isViewFormPage() {
		var path = (window.location && window.location.pathname ? window.location.pathname : "").toLowerCase();
		return path.indexOf("dispform.aspx") !== -1;
	}

	function forEachNode(nodeList, callback) {
		for (var i = 0; i < nodeList.length; i += 1) {
			callback(nodeList[i], i);
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

	function findAttachmentRow() {
		return document.getElementById("SPFieldAttachments") || document.getElementById("idAttachmentsTable") || null;
	}

	function injectStyles() {
		if (document.getElementById("legacy-attachments-style")) {
			return;
		}

		var style = document.createElement("style");
		style.id = "legacy-attachments-style";
		style.type = "text/css";
		style.textContent =
			"#SPFieldAttachments.legacy-attachments-box{" +
			"background:#f3f3f3;border:1px solid #c8c8c8;border-radius:4px;padding:10px 12px;" +
			"}" +
			"#legacy-attachments-note{" +
			"margin:0 0 8px 0;padding:6px 8px;background:#e9e9e9;border-left:3px solid #8a8a8a;" +
			"font-size:12px;line-height:1.4;color:#333;font-style:italic;" +
			"}";

		document.head.appendChild(style);
	}

	function addLegacyNote() {
		var field = document.getElementById("SPFieldAttachments");
		if (!field) {
			return;
		}

		field.classList.add("legacy-attachments-box");

		if (document.getElementById("legacy-attachments-note")) {
			return;
		}

		var note = document.createElement("div");
		note.id = "legacy-attachments-note";
		note.textContent = "Poznámka: Jedná se o starší proces s přílohami.";

		if (field.firstChild) {
			field.insertBefore(note, field.firstChild);
		} else {
			field.appendChild(note);
		}
	}

	function showElement(element) {
		if (!element) {
			return;
		}

		element.style.display = "";
		element.style.visibility = "visible";
		element.style.height = "auto";
		element.style.overflow = "visible";
		element.hidden = false;
	}

	function renameAttachmentLabel(row) {
		if (!row) {
			return;
		}

		var candidates = row.querySelectorAll("th, td, span, a, label, nobr");
		forEachNode(candidates, function (node) {
			var text = normalizeText(node.textContent || node.innerText || "");
			if (text === "attachments") {
				node.textContent = "Přílohy";
			}
		});
	}

	function revealAttachments() {
		var row = findAttachmentRow();
		if (!row) {
			return false;
		}

		showElement(row);
		showElement(row.parentNode);
		showElement(row.parentNode ? row.parentNode.parentNode : null);
		showElement(document.getElementById("SPFieldAttachments"));
		showElement(document.getElementById("idAttachmentsTable"));
		renameAttachmentLabel(row);
		addLegacyNote();

		return true;
	}

	function init() {
		if (!isViewFormPage()) {
			return;
		}

		injectStyles();

		var attempts = 0;
		var timer = window.setInterval(function () {
			attempts += 1;
			if (revealAttachments() || attempts >= 20) {
				window.clearInterval(timer);
			}
		}, 250);

		revealAttachments();
	}

	if (document.readyState === "loading") {
		document.addEventListener("DOMContentLoaded", init);
	} else {
		init();
	}
})();