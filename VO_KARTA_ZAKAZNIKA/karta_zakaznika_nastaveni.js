$(document).ready(function () {

    // =====================================================
    // POUZE MICHAL ŠÍP
    // =====================================================

    var path = String(window.location && window.location.pathname ? window.location.pathname : "").toLowerCase();
    var query = String(window.location && window.location.search ? window.location.search : "").toLowerCase();
    var isEditFormPage = /\/editform\.aspx$/.test(path) ||
        (/\/listform\.aspx$/.test(path) && /(?:^|[?&])pagetype=(?:6|8)(?:&|$)/.test(query));

    if (!isEditFormPage) {
        return;
    }

    var ctx = window._spPageContextInfo || {};
    var loginName = String(
        ctx.userLoginName ||
        ctx.userName ||
        ctx.systemUserKey ||
        ""
    ).toLowerCase();
    var isMichalSip =
        loginName.indexOf("michal.sip") !== -1 ||
        loginName.indexOf("i:0#.w|sam\\michal.sip") !== -1 ||
        Number(ctx.userId) === 1119;

    if (!isMichalSip) {
        return;
    }

    // =====================================================
    // DEFINICE SEKCÍ
    // =====================================================

    var sections = {

        zakladni: [
            "Zákazník název",
            "Zákazník fakturační adresa",
            "Zákazník IČO",
            "Zákazník DIČ",
            "VO_web"
        ],

        kontakty: [
            "Zákazník telefon",
            "Zákazník mobil",
            "Zákazník email",
            "Jednatel",
            "Jednatel - email",
            "Jednatel - telefon",
            "odpovedna_osoba_zakaznika"
        ],

        fakturace: [
            "Zákazník splatnost",
            "Zákazník perioda faktury",
            "Adresa pro korespondenci",
            "Zákazník odeslání faktur",
            "vo_zakaznik_varianta tisku faktury",
            "vo_zakaznik_celk_castka_na_radku_fa"
        ],

        slevy: [
            "Zákazník fakturační sleva (%)",
            "Zákazník skupina slev",
            "Zákazník dod. slevy",
            "Zákazník slevy SG"
        ],

        doprava: [
            "VO_Zakaznik_Dodaci_adresa",
            "Zákazník doprava časové okno",
            "Zákazník doprava pracovní doba",
            "Zákazník doprava klíč",
            "Zákazník doprava mrtvá schránka",
            "VO_Telefon_pro_dopravu",
            "Zákazník doprava volat",
            "Zákazník doprava samostatný PPL",
            "Zákazník doprava společné místo vykládky"
        ],

        b2b: [
            "Zákazník B2B Login",
            "Zákazník B2B Heslo",
            "Zákazník B2B sklad"
        ],

        dodaciListy: [
            "VO_zakazník_export_cen",
            "Zákazník dodací list email",
            "Zákazník dodací list formát",
            "Zákazník dodací list rabaty",
            "Zákazník dodací list export šarží",
            "Zákazník dodací list varianta"
        ],

        restrikce: [
            "Zákazník skupina restrikce prodeje",
            "Restrikce prodeje",
            "zakaz_prodejni_objednavky_1",
            "zakaz_prodejni_objednavky_2"
        ],

        vse: []
    };

    // =====================================================
    // VZHLED
    // =====================================================

    $("<style>")
        .prop("type", "text/css")
        .html(`
            #msWizardOverlay{
                position:fixed;
                left:0;
                top:0;
                width:100%;
                height:100%;
                background:rgba(0,0,0,.55);
                z-index:999999;
            }

            #msWizardBox{
                width:700px;
                background:#fff;
                margin:80px auto;
                padding:25px;
                border-radius:10px;
                box-shadow:0 0 20px rgba(0,0,0,.3);
            }

            #msWizardBox h2{
                margin-top:0;
                margin-bottom:20px;
            }

            .msSectionBtn{
                width:100%;
                text-align:left;
                padding:12px;
                margin-bottom:8px;
                cursor:pointer;
                border:1px solid #ccc;
                background:#f8f8f8;
                font-size:14px;
            }

            .msSectionBtn:hover{
                background:#eaeaea;
            }

            .msSectionBtn.is-active{
                background:#d9ecff;
                border-color:#0078d4;
            }

            .msWizardActions{
                margin-top:14px;
                display:flex;
                gap:10px;
            }

            .msWizardApply,
            .msWizardCancel{
                padding:10px 14px;
                border:1px solid #ccc;
                background:#f8f8f8;
                cursor:pointer;
            }

            .msWizardApply{
                background:#0078d4;
                border-color:#0078d4;
                color:#fff;
            }

            #msChangeSection{
                position:fixed;
                top:10px;
                right:10px;
                z-index:99999;
                padding:8px 15px;
                display:none;
            }

            .msSectionInfo{
                position:fixed;
                top:50px;
                right:10px;
                background:#0078d4;
                color:white;
                padding:6px 14px;
                z-index:99999;
                border-radius:4px;
                display:none;
            }
        `)
        .appendTo("head");


    // =====================================================
    // FUNKCE
    // =====================================================

    function hideAllRows() {

        $(".ms-formtable tr").hide();

        $("#idAttachmentsRow").show();
    }

    function showAllRows() {

        $(".ms-formtable tr").show();
    }

    function normalizeText(value) {

        if (!value) {
            return "";
        }

        var text = String(value).toLowerCase();

        if (typeof text.normalize === "function") {
            text = text.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
        }

        return text
            .replace(/\s+/g, " ")
            .replace(/[:*]/g, "")
            .trim();
    }

    function getRowLabel(row) {

        return normalizeText(
            $(row).find("h3:first").text() ||
            $(row).find("th h3:first").text() ||
            $(row).find("th nobr:first").text() ||
            $(row).find("th:first").text()
        );
    }

    function showSections(sectionNames) {

        if (!Array.isArray(sectionNames)) {
            sectionNames = [sectionNames];
        }

        if (sectionNames.indexOf("vse") >= 0) {

            showAllRows();

            $("#msCurrentSection")
                .text("Zobrazen celý formulář");

            return;
        }

        hideAllRows();

        var wantedBySection = {};
        sectionNames.forEach(function (sectionName) {
            (sections[sectionName] || []).forEach(function (name) {
                wantedBySection[normalizeText(name)] = true;
            });
        });

        var rows = Object.keys(wantedBySection);
        var shown = 0;

        $(".ms-formtable tr").each(function () {

            var label = getRowLabel(this);
            var match = rows.some(function (wanted) {
                return label === wanted ||
                    label.indexOf(wanted) >= 0 ||
                    wanted.indexOf(label) >= 0;
            });

            if (match) {
                $(this).show();
                shown += 1;
            }
        });

        if (shown === 0) {
            showAllRows();
            $("#msCurrentSection")
                .text("Nenalezena pole pro sekci, zobrazen celý formulář");
            return;
        }

        var caption = sectionNames.join(", ");
        $("#msCurrentSection")
            .text("Aktivní oblasti: " + caption);
    }


    function openWizard() {

        $("#msWizardOverlay").remove();

        $("body").append(`

            <div id="msWizardOverlay">

                <div id="msWizardBox">

                    <h2>Co chcete upravit?</h2>

                    <button class="msSectionBtn" data-section="zakladni">
                        Základní údaje
                    </button>

                    <button class="msSectionBtn" data-section="kontakty">
                        Kontakty
                    </button>

                    <button class="msSectionBtn" data-section="fakturace">
                        Fakturace
                    </button>

                    <button class="msSectionBtn" data-section="slevy">
                        Slevy
                    </button>

                    <button class="msSectionBtn" data-section="doprava">
                        Doprava
                    </button>

                    <button class="msSectionBtn" data-section="b2b">
                        B2B
                    </button>

                    <button class="msSectionBtn" data-section="dodaciListy">
                        Dodací listy
                    </button>

                    <button class="msSectionBtn" data-section="restrikce">
                        Restrikce
                    </button>

                    <button class="msSectionBtn" data-section="vse">
                        Zobrazit celý formulář
                    </button>

                    <div class="msWizardActions">
                        <button class="msWizardApply" type="button">Použít výběr</button>
                        <button class="msWizardCancel" type="button">Zrušit</button>
                    </div>

                </div>

            </div>

        `);

        $(".msSectionBtn").on("click", function () {

            var selected = String($(this).data("section") || "");

            if (selected === "vse") {
                $(".msSectionBtn").removeClass("is-active");
                $(this).addClass("is-active");
                return;
            }

            $(".msSectionBtn[data-section='vse']").removeClass("is-active");
            $(this).toggleClass("is-active");
        });

        $(".msWizardApply").on("click", function () {

            var selected = [];

            $(".msSectionBtn.is-active").each(function () {
                selected.push(String($(this).data("section") || ""));
            });

            if (!selected.length) {
                alert("Vyberte alespoň jednu oblast.");
                return;
            }

            showSections(selected);

            $("#msWizardOverlay").remove();

            $("#msChangeSection").show();

            $("#msCurrentSection").show();
        });

        $(".msWizardCancel").on("click", function () {
            $("#msWizardOverlay").remove();
        });
    }


    // =====================================================
    // TLAČÍTKO
    // =====================================================

    $("body").append(`
        <button id="msChangeSection">
            Změnit oblast
        </button>

        <div id="msCurrentSection" class="msSectionInfo"></div>
    `);

    $("#msChangeSection").on("click", function () {

        openWizard();

    });


    // =====================================================
    // START
    // =====================================================

    openWizard();

});