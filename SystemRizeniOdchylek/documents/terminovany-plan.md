# Termínovaný plán zavedení systému řízení odchylek

Pracovní návrh ze dne 20. 10. 2026 podle `zadani.txt` a existujících návrhů formulářů.

Předpoklady: zahájení 1. 10. 2026, dostupný vývojář a správce SharePoint/Treeinfo/WFE, průběžná součinnost QA/QP a interního auditora. Data jsou plánovací cíle, nikoliv potvrzené závazky. Jména odpovědných osob a jejich kapacity je potřeba doplnit při zahájení. Plán zahrnuje vytvoření systému i organizační přípravu procesu.

## Harmonogram

| Termín | Etapa a činnosti | Odpovědnost | Výstup a podmínka dokončení |
|---|---|---|---|
| 1.–7. 10. 2026 | Potvrzení rozsahu, rolí, zastupování, klasifikace, lhůt a eskalací; inventura otevřených odchylek; ověření přístupů do testovacího prostředí | QA/QP, interní auditor, vedoucí středisek, IT | Schválený rozsah první verze, matice rolí a seznam otevřených rozhodnutí s vlastníky. Určena oprávněná zastupující osoba pro odborná rozhodnutí a střety rolí. |
| 8.–14. 10. 2026 | Návrh datových seznamů, příloh, auditu, oprávnění a přechodů workflow; ověření rozhraní WFE a FMEA; návrh SOP a akceptačních scénářů | Vývojář, IT, QA/QP | Odsouhlasený datový model, technicky ověřené napojení a kritéria převzetí. Bez tohoto milníku nelze slíbit termín firemního nasazení. |
| 15.–21. 10. 2026 | Lokální prototyp SQLite: hlášení, hlavní karta, posouzení rizika, zajištění produktu/procesu, vyšetřování, role a historie | Vývojář; QA/QP a auditor ověřují | Průchod od založení do ukončení vyšetřování na modelových případech včetně procesní odchylky bez produktu. |
| 22.–29. 10. 2026 | CAPA, revize FMEA, neměnné snímky rizika, ověření účinnosti, schválení a uzavření; simulace workflow | Vývojář, QA/QP, vlastník FMEA | Úplný průchod odchylky; blokace uzavření bez potřebných důkazů, schválení a vypořádání revize FMEA. |
| 30. 10.–5. 11. 2026 | Balíčky kontrol a hromadné založení, statistiky, termíny a eskalace; předvedení a připomínky k prototypu | Vývojář, interní auditor, vedoucí středisek | Funkční prototyp pokrývající zadání; evidence připomínek a potvrzený rozsah pilotu. |
| 6.–12. 11. 2026 | Napojení na testovací SharePoint 2013, Treeinfo a WFE Sapiens; oprávnění, řízené přílohy, audit a vazby FMEA | IT, vývojář, správce WFE | Celý proces funguje v testovacím firemním prostředí; oprávnění jsou vynucena serverem a chybové stavy nezpůsobují ztrátu či duplicitu záznamů. |
| 13.–20. 11. 2026 | Uživatelské a bezpečnostní ověření, náprava vad, zkouška obnovy, schválení SOP a školení uživatelů | QA/QP, auditor, IT, vedoucí středisek | Doložené výsledky ověření; žádná otevřená kritická vada. Připravený postup podpory, zálohování a návratu při selhání. |
| 23.–27. 11. 2026 | Omezený pilot v jednom středisku; kontrolované převzetí otevřených případů; denní vyhodnocení vad a termínů | Vedoucí pilotního střediska, QA/QP, IT | Vyhodnocený pilot a rozhodnutí o pokračování. Pokud nejsou splněna kritéria, pilot se prodlouží a plošné spuštění odloží. |
| 30. 11.–4. 12. 2026 | Rezerva na opravy, opakované ověření a připravenost ostatních středisek | Vývojář, IT, QA/QP | Uzavření vad bránících nasazení a schválená připravenost k provozu. |
| 7. 12. 2026 | Cílové plošné spuštění | QA/QP schvaluje proces, IT technickou připravenost; ředitel rozhoduje o potřebných zdrojích | Zahájení běžného provozu při splnění podmínek převzetí. |
| 21. 12. 2026 | První provozní přezkoumání | QA/QP, auditor, vedoucí středisek | Přehled otevřených případů, prodlení, rizik, chyb systému a potřebných úprav. |

Etapy kolem 28. 10. a 17. 11. počítají se státními svátky jako s nepracovními dny. Provozní přezkoumání 21. 12. proběhne před vánočními svátky; dostupnost podpory během svátků je nutné dohodnout. Harmonogram vyžaduje potvrzení kapacit do 7. 10. 2026.

## Milníky

- **7. 10. 2026:** potvrzen proces, odpovědnosti a rozsah.
- **14. 10. 2026:** schválen návrh a ověřena proveditelnost integrace.
- **5. 11. 2026:** předán kompletní lokální prototyp.
- **20. 11. 2026:** systém připraven k pilotu.
- **27. 11. 2026:** vyhodnocen pilot.
- **7. 12. 2026:** cílové spuštění do běžného provozu.

## Pravidla řízení plánu

- Každé pondělí proběhne krátká kontrola postupu: dokončené výstupy, překážky, nejbližší termíny a rozhodnutí.
- Každý úkol dostane konkrétního vlastníka; vedoucí středisek určí řešitele a zajistí kapacitu pro vyšetřování a CAPA.
- Zpoždění předchozího milníku se promítne do navazujících termínů. Největší závislostí je dostupnost a chování WFE Sapiens a serverových oprávnění.
- Rozšíření rozsahu po 14. 10. se posoudí podle dopadu na pilot a nasazení; termíny se upraví vědomým rozhodnutím.
- Produkční spuštění vyžaduje ověřené oprávnění a audit, funkční eskalace, zálohování a obnovu, schválený postup práce a proškolené uživatele. Lokální prototyp sám tuto podmínku nesplňuje.

## Termíny jednotlivých odchylek

Níže jsou pouze dosavadní návrhové interní cíle ze zadání; QA/QP je má potvrdit do 7. 10. 2026. Nenahrazují konkrétní lhůty stanovené pro daný případ.

| Klasifikace | Návrhový cíl od zjištění |
|---|---|
| Kritická | Okamžité zajištění a informování QA/QP; předběžné posouzení do 1 pracovního dne. Další termíny stanoví QA podle případu. |
| Vysoká | Zajištění a oznámení QA/QP v den zjištění; vyšetřování do 10 pracovních dnů; plán CAPA do 20 pracovních dnů. |
| Střední | Posouzení QA do 3 pracovních dnů; cílové uzavření do 30 kalendářních dnů. |
| Nízká | Dokumentace a přiměřené opatření; cílové uzavření do 30 kalendářních dnů. |

Termín realizace CAPA a ověření účinnosti se určí jednotlivě podle rizika a potřebné doby sledování. Případ se neuzavírá jen kvůli splnění cílové lhůty. Prodloužení vyžaduje zdůvodnění, nový termín a příslušné schválení; prodlení se eskaluje QA. Konkrétní závazná lhůta z protokolu či rozhodnutí má při plánování případu přednost.
