# FMEA Formulář - Bezpečnost a Offline Režim

## 🔒 Bezpečnostní opatření

### 1. **Ochrana proti XSS (Cross-Site Scripting)**
✅ **Implementováno**
- Všechny uživatelské vstupy jsou HTML-escapovány před vložením do DOM
- Funkce `escapeHtml()` převádí speciální znaky:
  - `<` → `&lt;` 
  - `>` → `&gt;`
  - `"` → `&quot;`
  - `'` → `&#039;`
  - `&` → `&amp;`

```javascript
// Příklad escapování
escapeHtml(record.proces) // Bezpečný HTML
```

### 2. **Ochrana proti CSRF (Cross-Site Request Forgery)**
✅ **Implementováno**
- Aplikace čte `X-RequestDigest` token z SharePoint
- Každý požadavek na SharePoint API obsahuje CSRF token
- Token se automaticky čte z `__REQUESTDIGEST` elementu

```javascript
'X-RequestDigest': document.getElementById('__REQUESTDIGEST').value
```

### 3. **Validace vstupů**
✅ **Implementováno**
- Závažnost, Výskyt, Odhalitelnost: rozsah 1-5 dle MET-GDP/GMP-FMEA-ALL-002
- Textová pole: povinná, min. 3 znaky (v produkci)
- Všechny hodnoty jsou kontrolovány před uložením

### 4. **Bezpečné ukládání dat**
✅ **Implementováno**
- localStorage je bezpečné pro lokální data
- Data se neposílají do třetích stran
- SharePoint API komunikace: HTTPS + CSRF token

### 5. **Ochrana proti SQL Injection**
✅ **Není relevantní**
- Aplikace neposílá SQL příkazy
- Používá SharePoint REST API (parametrizované dotazy)
- JSON formát data → nemožný SQL injection

### 6. **Kontrola autorpráv** (SharePoint)
✅ **Inherentní**
- SharePoint automaticky kontroluje oprávnění
- Uživatel musí mít přístup k listu
- API vrací chybu pokud chybí oprávnění

### 7. **Bezpečnost localStorage**
⚠️ **Omezení (v pořádku)**
- localStorage je dostupný pro JavaScript
- Není šifrován (ale běžně se neukládají hesla)
- Soubor FMEA údaje nejsou citlivé
- Pokud je potřeba větší bezpečnost → šifrovat localStorage

---

## 🌐 Offline režim

### Jak to funguje?

**Automatické přepínání:**

```javascript
// Aplikace automaticky detekuje prostředí:
siteUrl: (typeof _spPageContextInfo !== 'undefined') 
         ? _spPageContextInfo.webAbsoluteUrl 
         : '',
useLocalStorage: (typeof _spPageContextInfo === 'undefined')
                 // true = vývoj/offline, false = SharePoint
```

### Offline scénáře

#### 1️⃣ **Vývoj bez SharePointu** (v prohlížeči)
- ✅ **Funguje** - Automaticky se přepne na localStorage
- Soubor: `FRM-FMEA-Formular-TEST.html`
- Data se uchovávají v prohlížeči

#### 2️⃣ **SharePoint bez internetu** ❌
- ❌ **Nefunguje** - REST API vyžaduje síť
- Ale může být online cache:

```html
<!-- Přidat Service Worker pro offline cache -->
<script>
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('sw.js');
}
</script>
```

#### 3️⃣ **Hybridní režim** ✅ DOPORUČUJEME
- SharePoint = primární storage
- localStorage = backup cache
- Při chybě: fallback na cache

---

## 📱 Offline Funkčnost Plán

### Aktuální stav:
- ✅ localStorage pro vývoj
- ❌ Service Worker pro pravý offline
- ❌ Sycnhronizace při obnovení

### Jak aktivovat offline:

**Editujte FRM-FMEA-Config.js:**
```javascript
// Pro offline režim:
useLocalStorage: true,  // force offline
```

### Jak přidat Service Worker pro cache:

```javascript
// v FRM-FMEA-Formular.js

// Registrace Service Workera
if ('serviceWorker' in navigator) {
    navigator.serviceWorker.register('fmea-sw.js').then(function(reg) {
        console.log('Service Worker registrován');
    }).catch(function(err) {
        console.log('Service Worker chyba:', err);
    });
}

// Offline event
window.addEventListener('offline', function() {
    console.log('Offline - přepínám na localStorage');
    fmeaApp.showStatus('Bez internetu - pracuji offline', 'info');
});

window.addEventListener('online', function() {
    console.log('Online - synchronizuji s SharePoint');
    fmeaApp.showStatus('Připojeno - synchronizuji', 'info');
    fmeaApp.syncWithSharePoint();
});
```

---

## 🔐 Bezpečnostní Rekomendace

### Vysoká priorita ✅ (již implementováno)
- [x] HTML escapování
- [x] CSRF token
- [x] Input validace
- [x] Bez SQL queries

### Střední priorita ⚠️
- [ ] Content Security Policy (CSP) header
- [ ] Subresource Integrity (SRI) pro CDN
- [ ] Rate limiting na SharePoint API
- [ ] Audit logging

### Nízká priorita (není potřeba teď)
- [ ] Šifrování localStorage
- [ ] JWT tokeny
- [ ] Role-based access control (RBAC) - máme SharePoint

---

## 🧪 Bezpečnostní Testy

### Test 1: XSS - Vložit JavaScript
```
Proces: <script>alert('XSS')</script>
Výsledek: ✅ Script se escapuje, neexekutuje
```

### Test 2: Offline režim
```
1. Otevřít FRM-FMEA-Formular-TEST.html offline
2. Přidat záznamy
3. Zavřít a znovu otevřít
Výsledek: ✅ Data zůstávají (localStorage)
```

### Test 3: Offline → Online (v budoucnu)
```
1. Offline - přidat data
2. Připojit se k síti
3. Synchronizovat
Výsledek: ⏳ Zatím není implementováno
```

---

## 💾 Výstupy a Obnova

### Tisk / PDF FMEA Protokolu
```
1. Klikni "Tisk FMEA Protokolu" nebo "PDF"
2. Otevře se tiskový dialog prohlížeče
3. Pro PDF vyber "Uložit jako PDF"
```

Výstup obsahuje hlavičku samohýl group a samotnou FMEA tabulku.

### Importovat data (budoucí feature)
Import řízených dat lze doplnit později podle toho, jaký formát bude schválen pro interní dokumentaci.

### Ruční backup z localStorage
```javascript
// V DevTools konzoli:
console.log(JSON.stringify(
    JSON.parse(localStorage.getItem('fmeaRecords')), 
    null, 2
));
// Zkopírovat a uložit do souboru
```

---

## 🌍 Nasazení - Bezpečnostní Checklist

- [ ] Ověřit HTTPS na SharePointu
- [ ] Ověřit CORS nastavení
- [ ] Nastavit CSP headers
- [ ] Ověřit CSRF token (SharePoint defaults)
- [ ] Povolit Service Worker (optional)
- [ ] Testovat na IE 11 (bez Promise)
- [ ] Audit logy - logs do SharePoint (future)

---

## Shrnutí

| Aspekt | Status | Poznámka |
|--------|--------|----------|
| **XSS ochrana** | ✅ | HTML escapování |
| **CSRF ochrana** | ✅ | Token handling |
| **Input validace** | ✅ | Range + type checking |
| **SQL Injection** | ✅ | N/A (API based) |
| **Offline režim** | ✅ localStorage | Service Worker: pending |
| **Data šifrování** | ⚠️ | Není potřeba (interní data) |
| **Authentication** | ✅ | SharePoint AD/OAuth |
| **Authorization** | ✅ | SharePoint RBAC |

**Konečný verdikt: 🟢 BEZPEČNÉ pro vnitřní firemní použití**
