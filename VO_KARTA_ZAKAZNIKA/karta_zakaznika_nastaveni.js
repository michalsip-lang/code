$(document).ready(function () {

    // =====================================================
    // POUZE MICHAL ŠÍP
    // =====================================================

    if (!_spPageContextInfo || _spPageContextInfo.userId !== 1119) {
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

    function showSection(sectionName) {

        if (sectionName === "vse") {

            showAllRows();

            $("#msCurrentSection")
                .text("Zobrazen celý formulář");

            return;
        }

        hideAllRows();

        var rows = sections[sectionName];

        $(".ms-formtable tr").each(function () {

            var label = $(this)
                .find("h3")
                .text()
                .trim();

            if (rows.indexOf(label) >= 0) {
                $(this).show();
            }
        });

        $("#msCurrentSection")
            .text("Aktivní oblast: " + sectionName);
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

                </div>

            </div>

        `);

        $(".msSectionBtn").click(function () {

            var selected = $(this).data("section");

            showSection(selected);

            $("#msWizardOverlay").remove();

            $("#msChangeSection").show();

            $("#msCurrentSection").show();

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