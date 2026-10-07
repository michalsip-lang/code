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

    var groups = [
        {
            key: "nastaveni",
            name: "Nastavení",
            fields: [
                { title: "Jedná se o nového zákazníka nebo změnu?", internalName: "VO_OP_novy_zakaznik" },
                { title: "Vyhledat existujícího zákazníka", internalName: "zakaznik_ciselnik" },
                { title: "Nastavení a kontrolu potřebných dokumentů k založení karty provedl", internalName: "vo_zakaznik_kontrolu_nastaveni_k" },
                { title: "Datum kontroly nastavení karty", internalName: "datum_kontroly_nastaveni_karty" },
                { title: "Název obecného kontaktu", internalName: "nazev_obecneho_kontaktu" }
            ]
        },
        {
            key: "obecne",
            name: "Obecné",
            fields: [
                { title: "Obchodní zástupce", internalName: "VO_Obchodni_zastupce" },
                { title: "Pracovní odpovědnost", internalName: "VO_Pracovni_odpovednost" },
                { title: "Jednatel", internalName: "VO_Jednatel" },
                { title: "Jednatel - email", internalName: "VO_Jednatel_email" },
                { title: "Jednatel - telefonní spojení", internalName: "VO_Jednatel_telefon" },
                { title: "Jednatel - poznámka", internalName: "VO_Jednatel_Poznamka" },
                { title: "Odpovědná osoba za veterinární léčiva (KVL)", internalName: "odpovedna_osoba_zakaznika" }
            ]
        },
        {
            key: "zakaznik_fakturacni_udaje",
            name: "Zákazník - fakturační údaje",
            fields: [
                { title: "Název", internalName: "VO_Zakaznik" },
                { title: "Fakturační adresa", internalName: "VO_FA_Adresa" },
                { title: "Zákázat prodejní objednávky", internalName: "zakaz_prodejni_objednavky_1" },
                { title: "Adresa pro korespondenci", internalName: "VO_ZAK_Faktury_adresa" },
                { title: "IČO", internalName: "VO_ICO" },
                { title: "DIČ", internalName: "VO_ZAK_DIC" },
                { title: "Mobil", internalName: "VO_ZAK_Mobil" },
                { title: "Telefon", internalName: "VO_ZAK_Telefon" },
                { title: "Email", internalName: "VO_ZAK_EMail" },
                { title: "Web", internalName: "VO_web" },
                { title: "Bankovní spojení", internalName: "VO_ZAK_Ucet" },
                { title: "Restrikce prodeje", internalName: "VO_Restrikce_prodeje" },
                { title: "Profil odesílání FA, PD", internalName: "VO_ZAK_Odeslani_FA" },
                { title: "Profil odeslání FA, PD - email", internalName: "profil_odeslani_FA_PD" },
                { title: "", internalName: "VO_ZAK_B2B_Login" },
                { title: "", internalName: "VO_ZAK_B2B_Heslo" }
            ]
        },
        {
            key: "program_nkp",
            name: "Program NKP",
            fields: [
                { title: "Program NKP", internalName: "VO_NKP" },
                { title: "OZ za Norbrook", internalName: "VO_Oz_Noorbrook" },
                { title: "Info o stavu konta - SMS", internalName: "VO_NBR_Stav_konta_sms" },
                { title: "Info o stavu konta - Email", internalName: "VO_NBR_Stav_konta_email" },
                { title: "skupinový kontakt", internalName: "nkp_skupinovy_kontakt" }
            ]
        },
        {
            key: "program_pvp",
            name: "Program PVP",
            fields: [
                { title: "Program PVP", internalName: "VO_PVP" },
                { title: "Info o stavu konta - Email", internalName: "VO_NBR_Stav_konta_email1" },
                { title: "Skupinový kontakt", internalName: "pvp_skupinovy_kontakt" }
            ]
        },
        {
            key: "zakaznik_dodaci_udaje",
            name: "Zákazník - dodací údaje",
            fields: [
                { title: "Dodací adresa", internalName: "VO_Zakaznik_Dodaci_adresa" },
                { title: "Zákázat prodejní objednávky", internalName: "zakaz_prodejni_objednavky_2" },
                { title: "Telefon pro dopravu", internalName: "VO_Telefon_pro_dopravu" },
                { title: "Otevírací doba", internalName: "VO_ZAK_DOP_Pracovni_doba" },
                { title: "Časové okno", internalName: "VO_ZAK_DOP_Casove_okno" },
                { title: "Klíč", internalName: "VO_ZAK_DOP_Klic" },
                { title: "Mrtvá schránka", internalName: "VO_ZAK_DOP_MS" },
                { title: "Popis mrtvé schránky", internalName: "VO_Popis_MS" },
                { title: "Mrtvá schránka - SMS", internalName: "VO_MS_SMS" },
                { title: "Společné místo vykládky", internalName: "VO_ZAK_DOP_Spolecne_misto_vyklad" },
                { title: "Kontakt pro společné místo vykládky", internalName: "VO_Kontakt_spol_mista_vykladky" },
                { title: "Samostatný PPL", internalName: "VO_ZAK_DOP_Samostatny_PPL" },
                { title: "Volat?", internalName: "VO_ZAK_DOP_Volat" }
            ]
        },
        {
            key: "zakaznik_dodaci_udaje_2",
            name: "Zákazník - dodací údaje 2",
            fields: [
                { title: "Dodací adresa", internalName: "vo_dodaci_adresa_2" },
                { title: "Telefon pro dopravu", internalName: "vo_telefon_pro_dopravu_2" },
                { title: "Otevírací doba", internalName: "vo_oteviraci_doba_2" },
                { title: "Časové okno", internalName: "vo_casove_okno_2" },
                { title: "Klíč", internalName: "vo_klic_2" },
                { title: "Mrtvá schránka", internalName: "vo_mrtva_schranka_2" },
                { title: "Popis mrtvé schránky", internalName: "vo_popis_mrtve_schranky_2" },
                { title: "Mrtvá schránka - sms", internalName: "vo_mrtva_schranka_sms" },
                { title: "Společné místo vykládky", internalName: "spolecne_misto_vykladky_2" },
                { title: "Kontakt pro společné místo vykládky", internalName: "vo_spolecne_misto_vykladky_konta" },
                { title: "Samostatný PPL", internalName: "vo_samostatny_ppl_2" },
                { title: "Volat?", internalName: "vo_volat_2" }
            ]
        },
        {
            key: "odpady",
            name: "Odpady",
            fields: [
                { title: "Odpady mailing RZ", internalName: "VO_ZAK_Odpady_RZ" },
                { title: "Email", internalName: "VO_Likvidace_odpadu" }
            ]
        },
        {
            key: "parametry_nastaveni",
            name: "Parametry nastavení",
            fields: [
                { title: "Aktivace vykazování dat", internalName: "aktivace_vykazovani_dat" },
                { title: "Nastavení výchozích kódů", internalName: "nastaveni_vychozich_kodu" },
                { title: "Kombinovat dle kontaktu", internalName: "VO_Kombinace_dle_x0020_kontaktu" },
                { title: "Kombinovat dle objednávky", internalName: "VO_kombinace_dle_objednavky" },
                { title: "Varianty tisku dodávky", internalName: "VO_ZAK_DL_Varianta" },
                { title: "Profil odeslání dokladu - DL", internalName: "VO_Profil_odeslani_dokladu" },
                { title: "Doporučená cena na dodávce", internalName: "VO_ZAK_DL_Dop_ceny" },
                { title: "Skrýt benefity na dodávce - zboží", internalName: "VO_benefity_zbozi" },
                { title: "Skrýt benefity na dodávce - poplatky", internalName: "VO_benefity_poplatky" },
                { title: "Na dodávce katalogová cena bez DPH", internalName: "VO_DL_katalog_cena_bez_dph" },
                { title: "Na dodávce pouze fakturační sleva", internalName: "VO_DL_Pouze_FA_Sleva" },
                { title: "Mezisoučty za dodávky na faktuře", internalName: "VO_mezisoucty_za_dodavky" },
                { title: "Typ exportního souboru (dodávka)", internalName: "VO_ZAK_DL_Format" },
                { title: "Export slev", internalName: "VO_export_slev" },
                { title: "Export cen (dodávka)", internalName: "VO_ZAK_DL_Ceny" },
                { title: "Export cen rabatů (dodávky)", internalName: "VO_ZAK_DL_Rabaty" },
                { title: "Export šarží dodávky", internalName: "VO_ZAK_DL_Sarze_export" },
                { title: "Varianta tisku faktury", internalName: "vo_zakaznik_varianta_x0020_tisku" },
                { title: "Celková částka na řádku faktury", internalName: "vo_zakaznik_celk_castka_na_radku" }
            ]
        },
        {
            key: "zakaznik_obecne",
            name: "Zákazník - obecné",
            fields: [
                { title: "Poznámka", internalName: "obecne_poznamky" },
                { title: "Přílohy", internalName: "vo_zakaznik_prilohy" }
            ]
        }
    ];

    var groupByKey = {};
    groups.forEach(function (group) {
        groupByKey[group.key] = group;
    });

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

    function getAttributeText(element, attrName) {

        var value = $(element).attr(attrName);
        return normalizeText(value || "");
    }

    function rowMatchesInternalName(row, internalName) {

        var wanted = normalizeText(internalName);
        if (!wanted) {
            return false;
        }

        if (getAttributeText(row, "id").indexOf(wanted) >= 0) {
            return true;
        }

        var found = false;
        $(row).find("*").each(function () {
            var probes = [
                getAttributeText(this, "id"),
                getAttributeText(this, "name"),
                getAttributeText(this, "title"),
                getAttributeText(this, "data-sp-field-internal-name")
            ];

            if (probes.some(function (probe) { return probe.indexOf(wanted) >= 0; })) {
                found = true;
                return false;
            }
        });

        return found;
    }

    function rowMatchesTitle(row, title) {

        var wanted = normalizeText(title);
        if (!wanted) {
            return false;
        }

        var label = getRowLabel(row);
        return label === wanted || label.indexOf(wanted) >= 0 || wanted.indexOf(label) >= 0;
    }

    function findRowForField(field) {

        var match = null;

        $(".ms-formtable tr").each(function () {
            if (rowMatchesInternalName(this, field.internalName)) {
                match = $(this);
                return false;
            }
        });

        if (match) {
            return match;
        }

        $(".ms-formtable tr").each(function () {
            if (rowMatchesTitle(this, field.title)) {
                match = $(this);
                return false;
            }
        });

        return match;
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

        var fields = [];
        sectionNames.forEach(function (sectionKey) {
            var group = groupByKey[sectionKey];
            if (!group) {
                return;
            }

            group.fields.forEach(function (field) {
                fields.push(field);
            });
        });

        var table = $(".ms-formtable:first");
        var parent = table.find("tbody:first");
        if (!parent.length) {
            parent = table;
        }

        var usedRowIds = {};
        var shown = 0;

        fields.forEach(function (field) {
            var row = findRowForField(field);
            if (!row || !row.length) {
                return;
            }

            var rowKey = row.attr("id") || String(row.index());
            if (usedRowIds[rowKey]) {
                return;
            }

            usedRowIds[rowKey] = true;
            row.show();
            parent.append(row);
            shown += 1;
        });

        if (shown === 0) {
            showAllRows();
            $("#msCurrentSection")
                .text("Nenalezena pole pro sekci, zobrazen celý formulář");
            return;
        }

        var caption = sectionNames.map(function (sectionKey) {
            return groupByKey[sectionKey] ? groupByKey[sectionKey].name : sectionKey;
        }).join(", ");
        $("#msCurrentSection")
            .text("Aktivní oblasti: " + caption);
    }


    function openWizard() {

        $("#msWizardOverlay").remove();

        $("body").append(`

            <div id="msWizardOverlay">

                <div id="msWizardBox">

                    <h2>Co chcete upravit?</h2>

                    <div id="msSectionList"></div>

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

        groups.forEach(function (group) {
            $("#msSectionList").append(
                '<button class="msSectionBtn" data-section="' + group.key + '">' + group.name + '</button>'
            );
        });

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