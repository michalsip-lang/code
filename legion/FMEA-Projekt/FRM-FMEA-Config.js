/**
 * FMEA Formulář - Konfigurační soubor
 * 
 * Upravte tyto hodnoty podle vašeho SharePoint prostředí
 */

window.FMEA_CONFIG = {
    
    // ============= SHAREPOINT KONFIGURACE =============
    
    // Název SharePoint listy kde se budou data ukládat
    // Vytvořte seznam se sloupci: Proces_FMEA, Selhani, Nasledky, Pricina,
    // KontrolaPrevence, KontrolaDetekce, Zavaznost, Vyskytu, Detekce,
    // Opatreni, VlastnikOpatreni, TerminOpatreni, DukazUcinnosti,
    // ZavaznostPo, VyskytuPo, DetekcePo, Historie
    listName: 'FMEA_Analyza',
    
    // Site URL - obvykle se automaticky zjistí z _spPageContextInfo
    // Pokud je problém, zadejte ručně (bez koncového /)
    // Příklad: 'https://sharepoint.company.com/sites/MyProject'
    siteUrl: null, // null = auto-detect
    
    // Jazykové locale pro práci se SharePointem
    // 'cs' = Čeština, 'en' = Angličtina
    locale: 'cs',
    
    
    // ============= VÝVOJOVÉ REŽIMY =============
    
    // Použít localStorage místo SharePointu (pro vývoj a testování)
    useLocalStorage: false,
    
    // Debug mode - vypisy do konzole
    debug: true,
    
    // Simulovat pomalé připojení (pro testování loading stavu)
    simulateSlowNetwork: false,
    
    
    // ============= VZHLED =============
    
    // Počet záznamů na stránku (0 = bez stránkování)
    pageSize: 0,
    
    // Automatické třídění
    defaultSort: 'rpn', // 'proces', 'zavaznost', 'rpn'
    defaultSortDirection: 'desc',
    
    // Zobrazit RPN barvy
    showRPNHighlight: true,
    
    // Limity pro RPN barevné zvýraznění
    rpnThresholds: {
        high: 28,    // 28-125 = vysoké riziko, opatření povinné
        medium: 13   // 13-27 = střední riziko
    },
    
    
    // ============= MAPOVÁNÍ POLÍ SHAREPOINT =============
    
    // Mapujte názvům sloupců v SharePointu (jejich interní jména)
    fields: {
        id: 'ID',
        proces: 'Proces_FMEA',
        selhani: 'Selhani',
        nasledky: 'Nasledky',
        pricina: 'Pricina',
        kontrolaPrevence: 'KontrolaPrevence',
        kontrolaDetekce: 'KontrolaDetekce',
        zavaznost: 'Zavaznost',
        vyskytu: 'Vyskytu',
        detekce: 'Detekce',
        opatreni: 'Opatreni',
        vlastnikOpatreni: 'VlastnikOpatreni',
        terminOpatreni: 'TerminOpatreni',
        dukazUcinnosti: 'DukazUcinnosti',
        zavaznostPo: 'ZavaznostPo',
        vyskytuPo: 'VyskytuPo',
        detekcePo: 'DetekcePo',
        historie: 'Historie',
        modified: 'Modified',        // Povinné pro SharePoint
        modifiedBy: 'Editor'          // Povinné pro SharePoint
    },
    
    
    // ============= VALIDACE =============
    
    // Validace hodnot Závažnost/Výskyt/Odhalitelnost dle MET-GDP/GMP-FMEA-ALL-002
    validation: {
        numeric: {
            min: 1,
            max: 5,
            required: true
        },
        text: {
            minLength: 3,
            maxLength: 500,
            required: true
        }
    },
    
    
    // ============= PRÁVA A OPRÁVNĚNÍ =============
    
    // Vždy povoleno - nevaliduje práva
    allowDeleteWithoutPrompt: false,
    
    // Vyžadovat potvrzení před smazáním
    requireDeleteConfirmation: true,
    
    // Povolit tisk a PDF výstup tabulky
    allowPrintOutput: true,
    
    // Skrýt smazat tlačítko (pro povolené čtení pouze)
    hideDeleteButton: false,
    
    
    // ============= NOTIFIKACE A ZPRÁVY =============
    
    // Jak dlouho zobrazovat zprávu (ms)
    messageTimeout: 5000,
    
    // Jazyk zpráv
    messages: {
        recordAdded: 'Záznam přidán',
        recordUpdated: 'Záznam aktualizován',
        recordDeleted: 'Záznam smazán',
        recordDeletedSharePoint: 'Záznam smazán ze SharePointu',
        confirmDelete: 'Opravdu chcete smazat tento záznam?',
        fillAllFields: 'Vyplňte prosím všechna povinná pole',
        invalidRange: 'Hodnota musí být v rozsahu 1-5',
        loading: 'Načítám záznamy...',
        noRecords: 'Zatím nejsou žádné záznamy. Přidejte první záznam formulářem výše.',
        error: 'Chyba: ',
        loadingError: 'Chyba při načítání. Zkontrolujte konzoli.',
        saveError: 'Chyba při ukládání'
    }
};

/**
 * Funkce pro načtení konfigurace
 * Kombinuje výchozí nastavení s vlastním
 */
function getFMEAConfig() {
    
    // Získej site URL z SharePointu
    if (!window.FMEA_CONFIG.siteUrl && typeof _spPageContextInfo !== 'undefined') {
        window.FMEA_CONFIG.siteUrl = _spPageContextInfo.webAbsoluteUrl;
    }
    
    // Fallback
    if (!window.FMEA_CONFIG.siteUrl) {
        window.FMEA_CONFIG.siteUrl = window.location.origin + '/sites/default';
    }
    
    return window.FMEA_CONFIG;
}

// Export
if (typeof module !== 'undefined' && module.exports) {
    module.exports = getFMEAConfig;
}
