# FMEA Formulář pro SharePoint 2013

Webová aplikace pro FMEA analýzu (Failure Mode and Effects Analysis) se automatickým výpočtem RPN (Risk Priority Number).

## Soubory

- **FRM-FMEA-Formular.html** - HTML struktura a styly
- **FRM-FMEA-Formular.js** - JavaScript logika a CSOM integrace
- **README.md** - Tato dokumentace

## Funkce

- ✅ Přidávání nových FMEA záznamů
- ✅ Inline editace dat v tabulce
- ✅ Automatický výpočet RPN = Závažnost × Výskyt × Detekce
- ✅ Barevné zvýraznění podle kritičnosti (vysoké RPN = červená)
- ✅ Mazání záznamů s potvrzením
- ✅ Integrace se SharePoint listy (REST API)
- ✅ Fallback na localStorage pro vývoj
- ✅ Validace vstupních dat
- ✅ Responzivní design

## FMEA Pole

| Pole | Typ | Rozsah | Popis |
|------|-----|--------|-------|
| Proces | Text | - | Identifikace procesu nebo funkce |
| FM | Text | - | Způsob neshody / co se může pokazit |
| Účinek | Text | - | Dopad na produkt, pacienta nebo shodu |
| Příčina | Text | - | Mechanismus nebo faktor vedoucí k neshodě |
| Kontrola prevence | Text | - | Opatření snižující pravděpodobnost příčiny |
| Kontrola detekce | Text | - | Opatření odhalující neshodu před dopadem |
| Závažnost | Číslo | 1-5 | 1=bez dopadu, 5=kritický dopad na pacienta, jakost nebo shodu |
| Výskyt | Číslo | 1-5 | 1=velmi vzácné, 5=velmi časté |
| Odhalitelnost | Číslo | 1-5 | 1=téměř jisté odhalení, 5=prakticky neodhalitelné |
| Opatření | Text | - | Povinné při RPN 28-125 nebo závažnosti 5 |
| Vlastník / termín / důkaz | Text, datum | - | Trasovatelnost realizace a ověření účinnosti |

## RPN Interpretace

- **RPN 1-12** - Zelená - Mírné riziko, bez povinného opatření
- **RPN 13-27** - Žlutá - Střední riziko, opatření není povinné
- **RPN 28-125** - Červená - Vysoké riziko, opatření je povinné
- **Závažnost = 5** - Opatření je povinné bez ohledu na RPN

## Nasazení do SharePointu 2013

### 1. Příprava SharePoint listy

V SharePointu si vytvořte nový seznam s těmito sloupci:

```
Sloupec          | Typ      | Povinný
-----------------+----------+--------
Proces           | Text     | Ano
Selhani          | Text     | Ano
Nasledky         | Text     | Ano
Zavaznost        | Číslo    | Ano
Vyskytu          | Číslo    | Ano
Detekce          | Číslo    | Ano
Pricina          | Více řádků textu | Ano
KontrolaPrevence | Více řádků textu | Ano
KontrolaDetekce  | Více řádků textu | Ano
Opatreni         | Více řádků textu | Podmíněně
VlastnikOpatreni | Text     | Podmíněně
TerminOpatreni   | Datum    | Podmíněně
DukazUcinnosti   | Více řádků textu | Ne
ZavaznostPo      | Číslo    | Ne
VyskytuPo        | Číslo    | Ne
DetekcePo        | Číslo    | Ne
Historie         | Více řádků textu | Ne
```

Pole `Historie` ukládá audit změn ve formátu JSON: kdo, kdy a co změnil.

### 2. Vytvoření webpartu v SharePoint

1. Otevřete stránku v režimu **Edit**
2. Klikněte na **Insert > Web Part**
3. Vyberte **Script Editor** (starší SharePoint) nebo **Content Editor Web Part**
4. Do editorů vložte HTML soubor (bez `<html>`, `<head>`, `<body>` tagů)

### 3. Konfiguraci v JavaScript souboru

V souboru `FRM-FMEA-Formular.js` upravte konfiguraci:

```javascript
var config = {
    listName: 'FMEA_Analyza',           // Název vaší SharePoint listy
    siteUrl: _spPageContextInfo.webAbsoluteUrl, // Automaticky detectováno
    useLocalStorage: false,              // true = vývoj, false = SharePoint
    fields: { ... }
};
```

### 4. Vložení do SharePoint

**Možnost A: Script Editor Web Part**
```html
<!-- Obsah FRM-FMEA-Formular.html bez <!DOCTYPE> a <html> tagů -->
```

**Možnost B: Content Editor Web Part + referenci na JS**
```html
<div id="fmeaApp"></div>
<script src="/sites/YourSite/Shared%20Documents/FRM-FMEA-Formular.js"></script>
```

## Vývoj a testování

### Lokální testování (bez SharePointu)

Nastavte v konfiguraci:
```javascript
config.useLocalStorage = true;
```

Data se budou ukládat v localStorage prohlížeče.

### Ladění

Otevřete DevTools (F12) v prohlížeči:

```javascript
// Zobrazení všech záznamů
console.log(fmeaApp.data.records);

// Ruční přidání testovacího záznamu
localStorage.setItem('fmeaRecords', JSON.stringify([
    {
        id: 1,
        proces: 'Testovací proces',
        selhani: 'Testovací selhání',
        nasledky: 'Testovací následky',
        pricina: 'Testovací příčina',
        kontrolaPrevence: 'Testovací prevence',
        kontrolaDetekce: 'Testovací detekce',
        zavaznost: 5,
        vyskytu: 3,
        detekce: 2,
        opatreni: 'Testovací opatření',
        vlastnikOpatreni: 'QA',
        terminOpatreni: '2026-06-30'
    }
]));
fmeaApp.loadRecords();
```

## Pokročilé funkce

### Tisk a PDF

Aplikace podporuje tisk FMEA tabulky a uložení do PDF přes tiskový dialog prohlížeče. Výstup obsahuje hlavičku samohýl group a samotnou FMEA tabulku.
```html
<button class="btn-secondary" onclick="fmeaApp.printReport()">Tisk FMEA tabulky</button>
<button class="btn-primary" onclick="fmeaApp.saveAsPDF()">PDF</button>
```

### Filtrování vysokých RPN

```javascript
var highRiskRecords = data.records.filter(function(r) {
    var rpn = calculateRPN(r.zavaznost, r.vyskytu, r.detekce);
    return rpn >= 28 || r.zavaznost === 5;
});
```

## Bezpečnost

- HTML je escapován pro ochranu proti XSS
- Všechny vstupy jsou validovány
- CSRF token se automaticky čte z SharePointu (`__REQUESTDIGEST`)
- REST API volání mají správný `Content-Type` header

## Browser Kompatibilita

- ✅ Chrome/Edge 60+
- ✅ Firefox 55+
- ✅ IE 11 (bez Promise - potřebuje polyfill)
- ✅ Safari 12+

Pro IE 11 přidejte polyfill na začátek HTML:
```html
<script src="https://cdn.jsdelivr.net/npm/fetch-ie8@1.5.0/fetch.js"></script>
