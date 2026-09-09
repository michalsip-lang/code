# Doporučené nasazení FMEA aplikace na SharePoint

Tento dokument popisuje doporučený způsob nasazení FMEA aplikace v prostředí SharePointu, kdy uživatelé aplikaci otevírají odkazem na HTML stránku uloženou v knihovně dokumentů.

## Doporučená architektura

Doporučené řešení:

- aplikace jako statické soubory v dokumentové knihovně SharePointu,
- data FMEA v SharePoint seznamu `FMEA_Analyza`,
- historie změn buď ve sloupci `Historie`, nebo lépe v samostatném seznamu `FMEA_Historie`,
- testovací/offline režim pouze pro vývoj, ne jako hlavní produkční úložiště.

Nedoporučuji ukládat všechna data do jednoho JSON souboru v knihovně dokumentů. U více uživatelů hrozí přepsání změn, složitější zamykání souboru, horší audit a horší filtrování.

## Doporučené soubory v SharePointu

Vytvořte dokumentovou knihovnu nebo složku například:

```text
SiteAssets/FMEA/
```

Do ní nahrajte:

```text
FRM-FMEA-Formular.html
FRM-FMEA-Formular.js
FRM-FMEA-Config.js
fmea-sw.js
```

Testovací soubor `FRM-FMEA-Formular-TEST.html` do produkce nedávejte jako hlavní odkaz. Slouží jen pro lokální testování.

Produkční odkaz pro uživatele bude například:

```text
https://sharepoint/sites/kvalita/SiteAssets/FMEA/FRM-FMEA-Formular.html
```

## Datové úložiště

### Varianta A: jeden seznam se sloupcem Historie

Vhodné pro první produkční nasazení.

Seznam:

```text
FMEA_Analyza
```

Výhody:

- jednodušší nasazení,
- aplikace už s tímto modelem počítá,
- všechny informace jsou u jednoho záznamu.

Nevýhody:

- historie je uložená jako JSON ve sloupci,
- hůře se nad historií filtruje a reportuje.

### Varianta B: samostatný seznam historie

Doporučená cílová varianta pro delší provoz.

Seznamy:

```text
FMEA_Analyza
FMEA_Historie
```

Výhody:

- lepší audit změn,
- jednodušší filtrování kdo/kdy/co,
- historie může růst nezávisle na hlavním záznamu,
- lepší pro budoucí reporty.

Nevýhody:

- vyžaduje drobnou úpravu aplikace, aby historii zapisovala do samostatného seznamu.

Pro aktuální verzi doporučuji začít variantou A a později přejít na variantu B, pokud bude historie důležitá pro audity.

## Sloupce seznamu FMEA_Analyza

Vytvořte vlastní seznam `FMEA_Analyza` s těmito sloupci. Interní názvy doporučuji ponechat přesně bez diakritiky.

| Interní název | Typ | Povinné | Poznámka |
|---|---|---|---|
| Title | Jednořádkový text | Ne | Lze plnit hodnotou procesu nebo nechat technicky |
| Proces_FMEA | Více řádků textu | Ano | Krok procesu / funkce; zobrazovaný název může být „Proces“ |
| Selhani | Více řádků textu | Ano | FM, co se může pokazit |
| Nasledky | Více řádků textu | Ano | Účinek / dopad |
| Pricina | Více řádků textu | Ano | Příčina selhání |
| KontrolaPrevence | Více řádků textu | Ano | Kontrola prevence |
| KontrolaDetekce | Více řádků textu | Ano | Kontrola detekce |
| Zavaznost | Číslo | Ano | S, rozsah 1-5 |
| Vyskytu | Číslo | Ano | O, rozsah 1-5 |
| Detekce | Číslo | Ano | D, rozsah 1-5 |
| Opatreni | Více řádků textu | Podmíněně | Povinné při RPN 28-125 nebo S=5 |
| VlastnikOpatreni | Jednořádkový text nebo Osoba | Podmíněně | Vlastník realizace |
| TerminOpatreni | Datum | Podmíněně | Termín realizace |
| DukazUcinnosti | Více řádků textu | Ne | Odkaz na záznam, trend, školení, validaci |
| ZavaznostPo | Číslo | Ne | S' po opatření, rozsah 1-5 |
| VyskytuPo | Číslo | Ne | O' po opatření, rozsah 1-5 |
| DetekcePo | Číslo | Ne | D' po opatření, rozsah 1-5 |
| Historie | Více řádků textu | Ne | JSON historie změn |

Poznámka: SharePoint si někdy vytvoří interní název podle prvního názvu sloupce. Proto je lepší sloupec nejdřív vytvořit bez diakritiky a až potom případně změnit zobrazovaný název.

## Volitelný seznam FMEA_Historie

Pokud se rozhodnete pro samostatnou historii, doporučená struktura je:

| Interní název | Typ | Poznámka |
|---|---|---|
| Title | Jednořádkový text | Např. `FMEA změna` |
| FMEA_ID | Číslo nebo Lookup | Vazba na položku ve `FMEA_Analyza` |
| Zmenil | Osoba nebo text | Uživatel, který změnu provedl |
| DatumZmeny | Datum a čas | Čas změny |
| Pole | Jednořádkový text | Název změněného pole |
| PuvodniHodnota | Více řádků textu | Původní hodnota |
| NovaHodnota | Více řádků textu | Nová hodnota |
| TypZmeny | Volba | Vytvoření / Úprava / Smazání |

## Doporučené oprávnění

Doporučené skupiny:

| Skupina | Oprávnění | Použití |
|---|---|---|
| FMEA Čtenáři | Čtení | Běžné nahlížení |
| FMEA Editoři | Přispívání / úprava položek | Vytváření a editace záznamů |
| FMEA Správci | Úplné řízení | Správa seznamů, sloupců a aplikace |

Doporučení:

- soubory aplikace v `SiteAssets/FMEA` mění jen správci,
- běžní uživatelé neupravují HTML/JS soubory,
- data v seznamu upravují jen schválení editoři,
- mazání záznamů omezit na správce nebo vybrané role.

## Konfigurace aplikace

V souboru `FRM-FMEA-Formular.js` zkontrolujte:

```javascript
var config = {
    listName: 'FMEA_Analyza',
    siteUrl: (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo) ? _spPageContextInfo.webAbsoluteUrl : '',
    useLocalStorage: (typeof _spPageContextInfo === 'undefined'),
    fields: { ... }
};
```

Pro produkci má aplikace běžet nad SharePointem. Lokální `localStorage` je určený jen jako fallback/testovací režim mimo SharePoint.

## Doporučený postup nasazení

1. Vytvořit knihovnu nebo složku `SiteAssets/FMEA`.
2. Nahrát aplikační soubory:
   - `FRM-FMEA-Formular.html`,
   - `FRM-FMEA-Formular.js`,
   - `FRM-FMEA-Config.js`,
   - `fmea-sw.js`.
3. Vytvořit seznam `FMEA_Analyza` se sloupci uvedenými výše.
4. Nastavit oprávnění na seznam a knihovnu.
5. Otevřít `FRM-FMEA-Formular.html` přes URL v prohlížeči.
6. Vytvořit testovací FMEA záznam.
7. Ověřit tlačítko `Detail`.
8. Ověřit `Změnit záznam` a zápis historie.
9. Ověřit tisk FMEA tabulky a uložení do PDF.
10. Ověřit přístup běžným uživatelem bez správcovských práv.

## Kontrola po nasazení

Po nasazení ověřte:

- stránka se otevře z odkazu bez chyby,
- načtou se existující záznamy ze seznamu,
- nový záznam se uloží do seznamu,
- detail záznamu se otevře přes tlačítko `Detail`,
- změna záznamu se uloží,
- historie ukáže kdo/kdy/co změnil,
- RPN se počítá podle S x O x D,
- vysoké riziko 28-125 vyžaduje opatření,
- závažnost 5 vyžaduje opatření bez ohledu na RPN,
- tisk/PDF obsahuje hlavičku samohýl group.

## Cache a aktualizace verze

Aplikace používá verzování JS v HTML odkazu, například:

```html
<script src="FRM-FMEA-Formular.js?v=20260907-detail-escape-fix"></script>
```

Při každé změně JavaScriptu doporučuji zvýšit hodnotu za `?v=`, například:

```html
<script src="FRM-FMEA-Formular.js?v=20260907-2"></script>
```

Pokud se používá service worker, zvyšte také `CACHE_NAME` v `fmea-sw.js`. Jinak může prohlížeč držet starou verzi aplikace.

Po aktualizaci doporučte uživatelům tvrdé obnovení stránky:

```text
Ctrl + F5
```

## Bezpečnostní doporučení

- Aplikaci provozovat pouze přes HTTPS.
- Uživatelé nesmí mít zbytečně právo měnit soubory aplikace.
- Zápis do seznamu řídit SharePoint oprávněními.
- Sloupce s historií nemazat ani ručně nepřepisovat.
- Pravidelně zálohovat seznam `FMEA_Analyza`.
- Pro auditní provoz zvážit samostatný seznam `FMEA_Historie`.

## Doporučení pro produkční start

Pro první ostré spuštění doporučuji:

1. nasadit aplikaci do samostatné složky `SiteAssets/FMEA`,
2. data ukládat do seznamu `FMEA_Analyza`,
3. historii zatím ukládat do sloupce `Historie`,
4. po ověření provozu zvážit přesun historie do samostatného seznamu,
5. ponechat `FRM-FMEA-Formular-TEST.html` pouze mimo produkční odkazy.

Toto řešení je nejjednodušší na správu, využívá SharePoint oprávnění a zároveň nechává aplikaci jako jednoduchý HTML odkaz pro uživatele.