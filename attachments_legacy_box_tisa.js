function TiSaForm_Customize(cfg) {
  function applyLegacyAttachmentsBox() {
    var row = document.getElementById("idAttachmentsRow");
    var field = document.getElementById("SPFieldAttachments");
    var table = document.getElementById("idAttachmentsTable");

    if (!row || !field) {
      return false;
    }

    row.style.display = "table-row";
    row.style.visibility = "visible";

    field.style.display = "table-cell";
    field.style.visibility = "visible";
    field.style.backgroundColor = "#f3f3f3";
    field.style.border = "1px solid #c8c8c8";
    field.style.borderRadius = "4px";
    field.style.padding = "10px 12px";

    if (table) {
      table.style.display = "table";
      table.style.visibility = "visible";
      table.style.marginTop = "6px";
    }

    var note = document.getElementById("legacyAttachmentsNote");
    if (!note) {
      note = document.createElement("div");
      note.id = "legacyAttachmentsNote";
      note.style.margin = "0 0 8px 0";
      note.style.padding = "6px 8px";
      note.style.backgroundColor = "#e9e9e9";
      note.style.borderLeft = "3px solid #8a8a8a";
      note.style.fontSize = "12px";
      note.style.lineHeight = "1.4";
      note.style.color = "#333";
      note.style.fontStyle = "italic";
      note.innerHTML = "Poznamka: Jedna se o starsi proces s prilohami.";

      if (field.firstChild) {
        field.insertBefore(note, field.firstChild);
      } else {
        field.appendChild(note);
      }
    }

    return true;
  }

  this.OnFormInitComplete = function(form) {
    if (!form || !form.IsDisplayForm) {
      return;
    }

    var attempts = 0;
    var timer = window.setInterval(function() {
      attempts += 1;
      if (applyLegacyAttachmentsBox() || attempts >= 60) {
        window.clearInterval(timer);
      }
    }, 250);

    applyLegacyAttachmentsBox();
  };
}
