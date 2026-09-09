# 📊 FMEA Projekt - Přehled Souborů

## 🚀 Jak Spustit?

### **Testovací Verze (Offline - Teď Hned)**
1. Otevřete: [FRM-FMEA-Formular-TEST.html](FRM-FMEA-Formular-TEST.html)
2. Přidávejte záznam FMEA
3. Data se uchovávají v localStorage

### **Produkční Verze (SharePoint)**
1. Vytvořte SharePoint seznam s poli:
    - Proces, Selhani, Nasledky, Pricina, KontrolaPrevence, KontrolaDetekce, Zavaznost, Vyskytu, Detekce, Opatreni
2. Vložte obsah [FRM-FMEA-Formular.html](FRM-FMEA-Formular.html) do Script Editor webpartu
3. Vložte [FRM-FMEA-Formular.js](FRM-FMEA-Formular.js) na SharePoint
4. Upravte konfiguraci v [FRM-FMEA-Config.js](FRM-FMEA-Config.js)

---

## 📁 Struktura Souborů

```
FMEA-Projekt/
├── INDEX.md                              ← Tento soubor
├── README-FMEA.md                        ← Dokumentace nasazení
├── README-SHAREPOINT-NASAZENI.md         ← Doporučené produkční nasazení
├── SECURITY-REPORT.md                    ← Bezpečnostní audit
├── BEZPECNOST-OFFLINE.md                 ← Offline režim detaily
│
├── FRM-FMEA-Formular-TEST.html          ← 🟢 SPUSTIT - Testovací verze
├── FRM-FMEA-Formular.html               ← Pro SharePoint Script Editor
├── FRM-FMEA-Formular.js                 ← Hlavní aplikace
├── FRM-FMEA-Config.js                   ← Konfigurace
└── fmea-sw.js                           ← Service Worker (offline)
```

---

## 🎯 Soubory - Podrobný Přehled

### 🟢 **Ke Spuštění Hned**

#### [FRM-FMEA-Formular-TEST.html](FRM-FMEA-Formular-TEST.html)
**Testovací verze pro vývoj a testování bez SharePointu**
- ✅ Plně funkční offline
- ✅ Hezký UI s statistikou
- ✅ Data v localStorage
- 👉 **Doporučuji: Otevřít v prohlížeči** HNED!

---

### 📄 **Dokumentace**

#### [README-FMEA.md](README-FMEA.md)
Kompletní dokumentace:
- Jak nasadit do SharePointu
- Konfigurace
- Vylepšení a rozšíření
- Kompatibilita prohlížečů

#### [README-SHAREPOINT-NASAZENI.md](README-SHAREPOINT-NASAZENI.md)
Doporučený produkční postup:
- Kam uložit HTML/JS soubory
- Jak založit SharePoint seznamy
- Jak nastavit oprávnění
- Jak ověřit nasazení
- Jak řešit cache a aktualizace

#### [SECURITY-REPORT.md](SECURITY-REPORT.md)
Bezpečnostní audit:
- ✅ XSS ochrana
- ✅ CSRF token
- ✅ Input validace
- ✅ SQL injection prevence
- Testovací výsledky

#### [BEZPECNOST-OFFLINE.md](BEZPECNOST-OFFLINE.md)
Offline a bezpečnostní detaily:
- Offline režim jak funguje
- Service Worker
- Doporučení vylepšení
- FAQ

---

### 💻 **Zdrojový Kód**

#### [FRM-FMEA-Formular.html](FRM-FMEA-Formular.html)
**HTML pro SharePoint Script Editor webpart**
- Struktura formuláře
- Tabulka
- Styly
- ~230 řádků

#### [FRM-FMEA-Formular.js](FRM-FMEA-Formular.js)
**Hlavní JavaScript logika**
- CRUD operace (Create, Read, Update, Delete)
- CSOM/REST API
- Offline detekce
- RPN výpočty
- ~450 řádků

#### [FRM-FMEA-Config.js](FRM-FMEA-Config.js)
**Konfigurace**
- Seznamová jména
- Site URL
- Validace
- Messages
- Lehce upravitelné

#### [fmea-sw.js](fmea-sw.js)
**Service Worker**
- Offline caching
- Cache-first strategie
- Auto-sync
- ~100 řádků

---

## 🔄 Quick Start - 3 Kroky

### 1️⃣ **Test**
```bash
# Otevřete v prohlížeči:
FRM-FMEA-Formular-TEST.html
```

### 2️⃣ **Přidejte záznam**
- Pole: Proces, Selhání, Následky
- Čísla: 1-5 (Závažnost, Výskyt, Odhalitelnost)
- Klikni "Přidat záznam"

### 3️⃣ **Vidíte tabulku?**
✅ Vše funguje!

---

## 🔧 Konfiguraci pro SharePoint

Editujte [FRM-FMEA-Config.js](FRM-FMEA-Config.js):

```javascript
window.FMEA_CONFIG = {
    listName: 'FMEA_Analyza',  // ← Váš seznam
    siteUrl: null,              // ← Auto-detect
    useLocalStorage: false,     // ← false = SharePoint
    // ... další nastavení
};
```

---

## 🌐 Offline Režim

Aplikace **automaticky** pracuje:
- **Online**: Data se ukládají do SharePointu
- **Offline**: Data se ukládají do localStorage
- **Bez internetu**: Aplikace stále funguje!

---

## 📊 FMEA Pola

| Pole | Typ | Rozsah | Příklad |
|------|-----|--------|---------|
| Proces | Text | - | "Výroba součástky" |
| Selhání | Text | - | "Prasklina na povrchu" |
| Následky | Text | - | "Vrácení od zákazníka" |
| Závažnost | Číslo | 1-5 | 5 (kritické) |
| Výskyt | Číslo | 1-5 | 3 (občasné) |
| Odhalitelnost | Číslo | 1-5 | 2 (vysoká šance odhalení) |
| **RPN** | **Vypočítáno** | **1-125** | **5×3×2=30** |

---

## 🎨 Barvy RPN

- 🔴 **Červená** (RPN 28-125) - Vysoké riziko, opatření povinné
- 🟡 **Žlutá** (RPN 13-27) - Střední riziko
- 🟢 **Zelená** (RPN 1-12) - Mírné riziko

---

## ⚙️ Technologie

- **HTML5** - Struktura
- **CSS3** - Styling  
- **Vanilla JavaScript** - Bez frameworků
- **REST API** - SharePoint integrace
- **Service Worker** - Offline podpora
- **localStorage** - Lokální ukládání

---

## ✅ Kontrola

- ✅ Bezpečné (XSS, CSRF protected)
- ✅ Offline funkční (Service Worker)
- ✅ Kompatibilní (Chrome, Firefox, Edge, Safari, IE11)
- ✅ Bez externích závislostí (vanillaJS)
- ✅ SharePoint 2013 kompatibilní

---

## 🚀 Nasazení - Checklist

- [ ] Vytvořit seznam v SharePointu
- [ ] Nastavit sloupce (Proces, Selhani, atd.)
- [ ] Konfigurovat FRM-FMEA-Config.js
- [ ] Vložit HTML do Script Editor webpartu
- [ ] Nahrát JS soubory na SharePoint
- [ ] Registrovat Service Worker (optional)
- [ ] Testovat v prohlížeči
- [ ] Live nasazení

---

## 📞 Support

Viz dokumentační soubory:
- **Jak to funguje?** → [README-FMEA.md](README-FMEA.md)
- **Je to bezpečné?** → [SECURITY-REPORT.md](SECURITY-REPORT.md)
- **Offline?** → [BEZPECNOST-OFFLINE.md](BEZPECNOST-OFFLINE.md)

---

## 📅 Verze

- **v1.0** - Základní FMEA s offline
- **v1.1** - Offline sync pending
- **v2.0** - Import řízených dat, audit logy

---

**🟢 PROJEKT JE PŘIPRAVEN K NASAZENÍ!**

Začínat: [FRM-FMEA-Formular-TEST.html](FRM-FMEA-Formular-TEST.html) 👈
