# 🔒 FMEA Formulář - Bezpečnostní Report

## ✅ Bezpečnost - Shrnutí

### Implementované ochrany

| Hrozba | Ochrana | Status | Detaily |
|--------|---------|--------|---------|
| **XSS (JavaScript injection)** | HTML escapování | ✅ | Všechny vstupy `<>\"'&` → entity |
| **CSRF (Cross-Site Request Forgery)** | SharePoint token | ✅ | X-RequestDigest v headers |
| **SQL Injection** | REST API | ✅ | Bez SQL - parametrizované queries |
| **Nevalidované vstupy** | Validace | ✅ | Range 1-5, type checking |
| **Neoprávněný přístup** | SharePoint RBAC | ✅ | Automatická kontrola práv |
| **Man-in-the-Middle** | HTTPS (SharePoint) | ✅ | Vyžaduje HTTPS |
| **Sensitive data exposure** | Lokální storage | ✅ | Data nejsou citlivá |
| **Broken access control** | SharePoint auth | ✅ | AD/OAuth integrován |

---

## 🌐 Offline Režim - Jak Funguje

### Detekce prostředí

```javascript
// Automaticky detekuje:
useLocalStorage = (typeof _spPageContextInfo === 'undefined')

// Výsledek:
// Bez SharePointu (offline/test) → localStorage
// V SharePointu → REST API + localStorage fallback
```

### Strategie offline

```
┌─────────────────────────┐
│   Pokus se připojit     │
│   k SharePoint API      │
└────────────┬────────────┘
             │
         ✅ úspěch → Uložit do SharePoint
         ❌ chyba ↓
             │
        localStorage (backup)
             │
    ✅ Data zůstávají, 
       app funguje offline
```

### Service Worker

**Soubor:** `fmea-sw.js`

```javascript
// Cache-first strategie:
1. Najdi v cache
2. Pokud není → stáhni ze sítě
3. Pokud nefunguje síť → offline fallback

// API volání (SharePoint):
- Vždy zkus síť
- Pokud selhane → vrať error (offline)
```

**Výhody:**
- ✅ Práce bez internetu
- ✅ Rychlejší načítání (cache)
- ✅ Otevírá se offline

**Nutné:**
- ✅ HTTPS (povinné pro SW)
- ✅ Browser podpora (Chrome, Firefox, Edge, Safari 12+)

---

## 📊 Testovací Výsledky

### Test 1: XSS Prevention ✅
```
Vstup: <script>alert('XSS')</script>
Výstup: &lt;script&gt;alert('XSS')&lt;/script&gt;
Výsledek: ✅ Script se neexekutuje
```

### Test 2: Offline Režim ✅
```
1. Zavřít internet
2. Otevřít FRM-FMEA-Formular-TEST.html
3. Přidat záznam
4. Zavřít prohlížeč
5. Znovu otevřít
Výsledek: ✅ Data se zachovala (localStorage)
```

### Test 3: Synchronizace (Budoucí)
```
1. Offline - 5 záznamů
2. Online - synchronizuj
3. SharePoint - zobrazit
Výsledek: ⏳ Implementovat v další verzi
```

---

## 🚀 Nasazení - Bezpečnostní Checklist

### Před nasazením do SharePointu

- [ ] ✅ Ověřit **HTTPS** (povinné)
- [ ] ✅ Ověřit **CORS** settings
- [ ] ✅ Nastavit **CSP** headers (optional)
- [ ] ⚠️ Testovat na **IE 11** (bez Promise)
- [ ] ✅ Audit přístupu (SharePoint permissions)
- [ ] ⚠️ Aktivovat Service Worker (optional)

### Konfigurační soubor

**FRM-FMEA-Config.js** - upravit dle prostředí:

```javascript
listName: 'FMEA_Analyza',  // ← Váš seznam
siteUrl: null,              // ← Auto-detect
useLocalStorage: false,     // ← false pro SharePoint
```

---

## 🔐 Doporučené Vylepšení

### P1 - Vysoká priorita (nie je kritické, ale doporučujeme)
- [ ] Přidat Content-Security-Policy header
- [ ] Logování akcí do SharePoint audit
- [ ] Šifrování localStorage (pro budoucí citlivé data)

### P2 - Střední priorita
- [ ] Import/Export s heslem
- [ ] Rate limiting API volání
- [ ] Backup funkce (automatic)
- [ ] Verze kontrola dat

### P3 - Nízká priorita
- [ ] JWT authentication
- [ ] Two-factor authentication (AD handles)
- [ ] Encryption at rest (not needed yet)

---

## 📱 Kompatibilita

| Browser | Offline | Service Worker | Poznámka |
|---------|---------|-----------------|----------|
| Chrome 90+ | ✅ | ✅ | Plná podpora |
| Firefox 88+ | ✅ | ✅ | Plná podpora |
| Edge 90+ | ✅ | ✅ | Plná podpora |
| Safari 12+ | ✅ | ✅ | Plná podpora |
| IE 11 | ✅ localStorage | ❌ SW | Bez SW, ale funguje |

---

## 🎯 Offline vs Online

### Offline (localStorage)
```
Pro: ✅ Funguje bez internetu
     ✅ Rychle
     ✅ Bez centrálního úložiště
     
Proti: ❌ Pouze lokální
       ❌ Lze smazat (cache clear)
       ❌ Bez synchronizace
```

### Online (SharePoint)
```
Pro: ✅ Centrální úložiště
     ✅ Backup
     ✅ Sdílení mezi uživateli
     ✅ Audit trail
     
Proti: ❌ Vyžaduje síť
       ❌ Zpožděná komunikace
```

### Hybrid (Doporučujeme)
```
✅ Online = primární (SharePoint)
✅ Offline cache = fallback (localStorage)
✅ Auto-sync při obnovení

Jak funguje:
1. Pokud je síť → uložit do SharePoint
2. Pokud není → uložit lokálně
3. Když se síť vrátí → sync
```

---

## 🧠 Bezpečnostní FAQ

**Q: Jsou data bezpečná?**
A: ✅ Ano. HTML je escapován, CSRF je chráněný, SharePoint kontroluje přístup.

**Q: Funguje bez internetu?**
A: ✅ Ano. localStorage fallback + Service Worker caching.

**Q: Mohu mazat data?**
A: ✅ Ano. Tlačítko "Smazat" v tabulce. Data jdou do SharePoint (pokud máte oprávnění).

**Q: Co když vypadne internet?**
A: ✅ App pokračuje. Data se ukládají lokálně. Sync při obnovení (budoucí verze).

**Q: Jsou data šifrovaná?**
A: ⚠️ Ne, ale nejsou citlivá (FMEA výstupy nejsou top-secret).

**Q: Kdo má přístup k datům?**
A: SharePoint RBAC. Jen uživatelé s oprávněním k seznamu.

---

## 📝 Bezpečnostní Politika

1. **Vstupy jsou vždy escapovány** - Nikdy nevěřit uživateli
2. **CSRF protection** - SharePoint tokens pro každý API call
3. **Offline fallback** - Vždy funguje (lokálně)
4. **No external dependencies** - Vanilla JS, bez CDN (lokálně)
5. **Audit-friendly** - SharePoint logs all changes

---

## 🛡️ Bezpečnostní Kontakty

Pokud najdete bezpečnostní chybu:
- Neposílat veřejně
- Kontaktovat IT bezpečnost
- Uvést typ chyby a kroky k reprodukci

---

**Finální Verdikt: 🟢 BEZPEČNÉ pro vnitřní firemní použití**

Aplikace splňuje bezpečnostní standardy pro malé až střední SharePoint projekty.
Pro vyšší bezpečnost → viz P1/P2 doporučení vylepšení.
