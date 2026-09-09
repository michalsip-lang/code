// FMEA Formulář - JavaScript aplikace
// Pro SharePoint 2013 - Script Editor webpart

var fmeaApp = (function() {
    'use strict';
    
    // Konfiguraceace
    var ratingMin = 1;
    var ratingMax = 5;
    var rpnThresholds = {
        medium: 13,
        high: 28
    };
    var config = {
        listName: 'FMEA_Analyza', // Název SharePoint listy
        siteUrl: (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo) ? _spPageContextInfo.webAbsoluteUrl : '',
        useLocalStorage: (typeof _spPageContextInfo === 'undefined'), // Pro vývojové testování
        fields: {
            proces: 'Proces',
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
            historie: 'Historie'
        }
    };

    var formFields = ['proces', 'selhani', 'nasledky', 'pricina', 'kontrolaPrevence', 'kontrolaDetekce', 'zavaznost', 'vyskytu', 'detekce', 'opatreni', 'vlastnikOpatreni', 'terminOpatreni', 'dukazUcinnosti', 'zavaznostPo', 'vyskytuPo', 'detekcePo'];
    var fieldLabels = {
        proces: 'Krok procesu',
        selhani: 'FM',
        nasledky: 'Účinek',
        pricina: 'Příčina',
        kontrolaPrevence: 'Kontrola prevence',
        kontrolaDetekce: 'Kontrola detekce',
        zavaznost: 'Závažnost',
        vyskytu: 'Výskyt',
        detekce: 'Odhalitelnost',
        opatreni: 'Opatření',
        vlastnikOpatreni: 'Vlastník opatření',
        terminOpatreni: 'Termín opatření',
        dukazUcinnosti: 'Důkaz účinnosti',
        zavaznostPo: 'Závažnost po opatření',
        vyskytuPo: 'Výskyt po opatření',
        detekcePo: 'Odhalitelnost po opatření'
    };
    var editState = {
        id: null,
        original: null
    };
    var detailClickBound = false;
    
    // Uložiště dat
    var data = {
        records: [],
        isDirty: false
    };
    
    // Inicializace - zavolá se při načtení stránky
    var init = function() {
        console.log('FMEA aplikace inicializace...');
        bindRecordDetailClicks();
        document.addEventListener('keydown', function(event) {
            if (event.key === 'Escape') {
                closeRecordDetail();
            }
        });
        loadRecords();
    };

    var bindRecordDetailClicks = function() {
        if (detailClickBound) return;
        detailClickBound = true;
        document.addEventListener('click', function(event) {
            var target = getRecordDetailClickTarget(event.target);
            if (!target) return;
            var id = target.getAttribute('data-record-id');
            if (!isNaN(id)) {
                event.preventDefault();
                openRecordDetail(id);
            }
        });
    };

    var getRecordDetailClickTarget = function(target) {
        while (target && target !== document) {
            if (target.getAttribute && target.getAttribute('data-action') === 'open-record-detail') {
                return target;
            }
            target = target.parentNode;
        }
        return null;
    };
    
    // Načtení záznamů ze SharePoint nebo localStorage
    var loadRecords = function() {
        showStatus('Načítám záznamy...', 'info');
        
        if (config.useLocalStorage) {
            loadFromLocalStorage();
        } else if (config.siteUrl) {
            loadFromSharePoint();
        } else {
            // Fallback na lokální storage pro testování
            loadFromLocalStorage();
        }
        
        // Zavolej updateStats pokud existuje (pro testovací prostředí)
        if (typeof window.fmeaAppUpdateStats === 'function') {
            setTimeout(window.fmeaAppUpdateStats, 100);
        }
    };
    
    // Načtení z localStorage (vývojové testování)
    var loadFromLocalStorage = function() {
        var stored = localStorage.getItem('fmeaRecords');
        data.records = stored ? JSON.parse(stored) : [];
        data.records = data.records.map(function(record) {
            record.zavaznost = normalizeRating(record.zavaznost);
            record.vyskytu = normalizeRating(record.vyskytu);
            record.detekce = normalizeRating(record.detekce);
            record.pricina = record.pricina || '';
            record.kontrolaPrevence = record.kontrolaPrevence || '';
            record.kontrolaDetekce = record.kontrolaDetekce || '';
            record.opatreni = record.opatreni || '';
            record.vlastnikOpatreni = record.vlastnikOpatreni || '';
            record.terminOpatreni = record.terminOpatreni || '';
            record.dukazUcinnosti = record.dukazUcinnosti || '';
            record.zavaznostPo = normalizeOptionalRating(record.zavaznostPo);
            record.vyskytuPo = normalizeOptionalRating(record.vyskytuPo);
            record.detekcePo = normalizeOptionalRating(record.detekcePo);
            record.historie = normalizeHistory(record.historie);
            return record;
        });
        if (stored) {
            saveToLocalStorage();
        }
        renderTable();
        showStatus(data.records.length + ' záznamů načteno', 'success');
    };
    
    // Načtení ze SharePoint REST API
    var loadFromSharePoint = function() {
        var url = config.siteUrl + "/_api/web/lists/getbytitle('" + config.listName + "')/items?$select=*";
        
        fetch(url, {
            method: 'GET',
            headers: {
                'Accept': 'application/json',
                'X-RequestDigest': document.getElementById('__REQUESTDIGEST') ? 
                                   document.getElementById('__REQUESTDIGEST').value : ''
            }
        })
        .then(function(response) {
            if (!response.ok) {
                throw new Error('Chyba při načítání: ' + response.status);
            }
            return response.json();
        })
        .then(function(result) {
            data.records = result.value.map(function(item) {
                return {
                    id: item.ID,
                    proces: item.Proces || '',
                    selhani: item.Selhani || '',
                    nasledky: item.Nasledky || '',
                    pricina: item.Pricina || '',
                    kontrolaPrevence: item.KontrolaPrevence || '',
                    kontrolaDetekce: item.KontrolaDetekce || '',
                    zavaznost: normalizeRating(item.Zavaznost),
                    vyskytu: normalizeRating(item.Vyskytu),
                    detekce: normalizeRating(item.Detekce),
                    opatreni: item.Opatreni || '',
                    vlastnikOpatreni: item.VlastnikOpatreni || '',
                    terminOpatreni: item.TerminOpatreni || '',
                    dukazUcinnosti: item.DukazUcinnosti || '',
                    zavaznostPo: normalizeOptionalRating(item.ZavaznostPo),
                    vyskytuPo: normalizeOptionalRating(item.VyskytuPo),
                    detekcePo: normalizeOptionalRating(item.DetekcePo),
                    historie: normalizeHistory(item.Historie)
                };
            });
            renderTable();
            showStatus(data.records.length + ' záznamů načteno ze SharePointu', 'success');
        })
        .catch(function(error) {
            console.error('Chyba načítání:', error);
            showStatus('Chyba: ' + error.message, 'error');
            // Fallback na localStorage
            loadFromLocalStorage();
        });
    };
    
    // Přidání nového záznamu
    var addRecord = function() {
        var record = readFormRecord(editState.id || Date.now());
        if (!validateRecord(record)) return;

        if (editState.id !== null) {
            saveEditedRecord(record);
            return;
        }

        record.historie = [{
            kdy: getTimestamp(),
            kdo: getCurrentUser(),
            co: 'Záznam vytvořen'
        }];
        
        if (config.useLocalStorage) {
            data.records.push(record);
            saveToLocalStorage();
            showStatus('Záznam přidán', 'success');
        } else if (config.siteUrl) {
            saveToSharePoint(record);
        }
        
        clearForm();
        renderTable();
    };
    
    // Uložení do localStorage
    var saveToLocalStorage = function() {
        localStorage.setItem('fmeaRecords', JSON.stringify(data.records));
    };
    
    // Uložení do SharePoint
    var saveToSharePoint = function(record) {
        var url = config.siteUrl + "/_api/web/lists/getbytitle('" + config.listName + "')/items";
        
        var itemData = {
            Proces: record.proces,
            Selhani: record.selhani,
            Nasledky: record.nasledky,
            Pricina: record.pricina,
            KontrolaPrevence: record.kontrolaPrevence,
            KontrolaDetekce: record.kontrolaDetekce,
            Zavaznost: record.zavaznost.toString(),
            Vyskytu: record.vyskytu.toString(),
            Detekce: record.detekce.toString(),
            Opatreni: record.opatreni,
            VlastnikOpatreni: record.vlastnikOpatreni,
            TerminOpatreni: record.terminOpatreni,
            DukazUcinnosti: record.dukazUcinnosti,
            ZavaznostPo: record.zavaznostPo ? record.zavaznostPo.toString() : '',
            VyskytuPo: record.vyskytuPo ? record.vyskytuPo.toString() : '',
            DetekcePo: record.detekcePo ? record.detekcePo.toString() : '',
            Historie: JSON.stringify(record.historie || [])
        };
        
        fetch(url, {
            method: 'POST',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'X-RequestDigest': document.getElementById('__REQUESTDIGEST') ? 
                                   document.getElementById('__REQUESTDIGEST').value : ''
            },
            body: JSON.stringify(itemData)
        })
        .then(function(response) {
            if (!response.ok) {
                throw new Error('Chyba při ukládání: ' + response.status);
            }
            showStatus('Záznam uložen do SharePointu', 'success');
            loadRecords();
        })
        .catch(function(error) {
            console.error('Chyba ukládání:', error);
            showStatus('Chyba: ' + error.message, 'error');
        });
    };
    
    // Úprava záznamu
    var updateRecord = function(id, field, value) {
        var record = data.records.find(function(r) { return r.id === id; });
        if (!record) return;
        var original = cloneRecord(record);
        
        var requiredRatingFields = ['zavaznost', 'vyskytu', 'detekce'];
        var optionalRatingFields = ['zavaznostPo', 'vyskytuPo', 'detekcePo'];
        if (requiredRatingFields.indexOf(field) !== -1 || optionalRatingFields.indexOf(field) !== -1) {
            if (optionalRatingFields.indexOf(field) !== -1 && value.trim() === '') {
                record[field] = '';
            } else {
                var numValue = parseInt(value, 10);
                if (!isRatingValid(numValue)) {
                    showStatus('Hodnota musí být v rozsahu 1-5', 'error');
                    return;
                }
                record[field] = numValue;
            }
        } else {
            record[field] = value;
        }
        addHistory(record, describeChanges(original, record));
        
        if (config.useLocalStorage) {
            saveToLocalStorage();
        } else if (config.siteUrl && record.id) {
            updateInSharePoint(record);
        }
        
        renderTable();
        showStatus('Záznam aktualizován', 'success');
    };
    
    // Aktualizace v SharePoint
    var updateInSharePoint = function(record) {
        var url = config.siteUrl + "/_api/web/lists/getbytitle('" + config.listName + "')/items(" + record.id + ")";
        
        var itemData = {
            Proces: record.proces,
            Selhani: record.selhani,
            Nasledky: record.nasledky,
            Pricina: record.pricina,
            KontrolaPrevence: record.kontrolaPrevence,
            KontrolaDetekce: record.kontrolaDetekce,
            Zavaznost: record.zavaznost.toString(),
            Vyskytu: record.vyskytu.toString(),
            Detekce: record.detekce.toString(),
            Opatreni: record.opatreni,
            VlastnikOpatreni: record.vlastnikOpatreni,
            TerminOpatreni: record.terminOpatreni,
            DukazUcinnosti: record.dukazUcinnosti,
            ZavaznostPo: record.zavaznostPo ? record.zavaznostPo.toString() : '',
            VyskytuPo: record.vyskytuPo ? record.vyskytuPo.toString() : '',
            DetekcePo: record.detekcePo ? record.detekcePo.toString() : '',
            Historie: JSON.stringify(record.historie || [])
        };
        
        fetch(url, {
            method: 'PATCH',
            headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json',
                'X-RequestDigest': document.getElementById('__REQUESTDIGEST') ? 
                                   document.getElementById('__REQUESTDIGEST').value : ''
            },
            body: JSON.stringify(itemData)
        })
        .catch(function(error) {
            console.error('Chyba aktualizace:', error);
            showStatus('Chyba při aktualizaci: ' + error.message, 'error');
        });
    };
    
    // Smazání záznamu
    var deleteRecord = function(id) {
        if (!confirm('Opravdu chcete smazat tento záznam?')) return;
        
        var index = data.records.findIndex(function(r) { return r.id === id; });
        if (index !== -1) {
            var record = data.records[index];
            data.records.splice(index, 1);
            
            if (config.useLocalStorage) {
                saveToLocalStorage();
                showStatus('Záznam smazán', 'success');
            } else if (config.siteUrl && record.id) {
                deleteFromSharePoint(record.id);
            }
            
            renderTable();
        }
    };
    
    // Smazání ze SharePoint
    var deleteFromSharePoint = function(itemId) {
        var url = config.siteUrl + "/_api/web/lists/getbytitle('" + config.listName + "')/items(" + itemId + ")";
        
        fetch(url, {
            method: 'DELETE',
            headers: {
                'X-RequestDigest': document.getElementById('__REQUESTDIGEST') ? 
                                   document.getElementById('__REQUESTDIGEST').value : ''
            }
        })
        .then(function(response) {
            if (!response.ok) {
                throw new Error('Chyba při mazání: ' + response.status);
            }
            showStatus('Záznam smazán ze SharePointu', 'success');
        })
        .catch(function(error) {
            console.error('Chyba mazání:', error);
            showStatus('Chyba: ' + error.message, 'error');
        });
    };
    
    // Výpočet RPN (Risk Priority Number)
    var calculateRPN = function(zavaznost, vyskytu, detekce) {
        return zavaznost * vyskytu * detekce;
    };

    var isRatingValid = function(value) {
        return !isNaN(value) && value >= ratingMin && value <= ratingMax;
    };

    var normalizeRating = function(value) {
        var parsed = parseInt(value, 10);
        if (isNaN(parsed)) return ratingMax;
        return Math.max(ratingMin, Math.min(ratingMax, parsed));
    };

    var normalizeOptionalRating = function(value) {
        if (value === null || value === undefined || value === '') return '';
        var parsed = parseInt(value, 10);
        if (isNaN(parsed)) return '';
        return Math.max(ratingMin, Math.min(ratingMax, parsed));
    };

    var normalizeHistory = function(value) {
        if (!value) return [];
        if (Object.prototype.toString.call(value) === '[object Array]') return value;
        try {
            var parsed = JSON.parse(value);
            return Object.prototype.toString.call(parsed) === '[object Array]' ? parsed : [];
        } catch (error) {
            return [];
        }
    };

    var getInputValue = function(id) {
        var element = document.getElementById(id);
        return element ? element.value.trim() : '';
    };

    var readFormRecord = function(id) {
        return {
            id: id,
            proces: getInputValue('proces'),
            selhani: getInputValue('selhani'),
            nasledky: getInputValue('nasledky'),
            pricina: getInputValue('pricina'),
            kontrolaPrevence: getInputValue('kontrolaPrevence'),
            kontrolaDetekce: getInputValue('kontrolaDetekce'),
            zavaznost: parseInt(getInputValue('zavaznost'), 10),
            vyskytu: parseInt(getInputValue('vyskytu'), 10),
            detekce: parseInt(getInputValue('detekce'), 10),
            opatreni: getInputValue('opatreni'),
            vlastnikOpatreni: getInputValue('vlastnikOpatreni'),
            terminOpatreni: getInputValue('terminOpatreni'),
            dukazUcinnosti: getInputValue('dukazUcinnosti'),
            zavaznostPo: normalizeOptionalRating(getInputValue('zavaznostPo')),
            vyskytuPo: normalizeOptionalRating(getInputValue('vyskytuPo')),
            detekcePo: normalizeOptionalRating(getInputValue('detekcePo'))
        };
    };

    var validateRecord = function(record) {
        clearValidationHighlights();
        var missingFields = [];
        ['proces', 'selhani', 'nasledky', 'pricina', 'kontrolaPrevence', 'kontrolaDetekce'].forEach(function(field) {
            if (!record[field]) missingFields.push(field);
        });

        if (missingFields.length > 0) {
            showValidationError('Doplňte povinná pole: ' + missingFields.map(function(field) { return fieldLabels[field]; }).join(', '), missingFields);
            return false;
        }

        if (!isRatingValid(record.zavaznost) || !isRatingValid(record.vyskytu) || !isRatingValid(record.detekce)) {
            showValidationError('Hodnoty pro závažnost, výskyt a odhalitelnost musí být v rozsahu 1-5', ['zavaznost', 'vyskytu', 'detekce']);
            return false;
        }

        var rpn = calculateRPN(record.zavaznost, record.vyskytu, record.detekce);
        if (requiresAction(record, rpn) && (!record.opatreni || !record.vlastnikOpatreni || !record.terminOpatreni)) {
            var actionFields = [];
            if (!record.opatreni) actionFields.push('opatreni');
            if (!record.vlastnikOpatreni) actionFields.push('vlastnikOpatreni');
            if (!record.terminOpatreni) actionFields.push('terminOpatreni');
            showValidationError('U vysokého rizika nebo závažnosti 5 doplňte: ' + actionFields.map(function(field) { return fieldLabels[field]; }).join(', '), actionFields);
            return false;
        }

        return true;
    };

    var clearValidationHighlights = function() {
        formFields.forEach(function(field) {
            var element = document.getElementById(field);
            if (element) {
                element.className = element.className.replace(/\bis-invalid\b/g, '').trim();
            }
        });
    };

    var showValidationError = function(message, fields) {
        showStatus(message, 'error');
        fields.forEach(function(field) {
            var element = document.getElementById(field);
            if (element && element.className.indexOf('is-invalid') === -1) {
                element.className += (element.className ? ' ' : '') + 'is-invalid';
            }
        });
        var first = document.getElementById(fields[0]);
        if (first) {
            first.focus();
            if (first.scrollIntoView) {
                first.scrollIntoView({ behavior: 'smooth', block: 'center' });
            }
        }
    };

    var cloneRecord = function(record) {
        return JSON.parse(JSON.stringify(record));
    };

    var getCurrentUser = function() {
        if (typeof _spPageContextInfo !== 'undefined' && _spPageContextInfo && _spPageContextInfo.userDisplayName) {
            return _spPageContextInfo.userDisplayName;
        }
        return 'Lokální uživatel';
    };

    var getTimestamp = function() {
        return new Date().toLocaleString('cs-CZ');
    };

    var describeChanges = function(before, after) {
        var changes = [];
        formFields.forEach(function(field) {
            var oldValue = before[field] === undefined || before[field] === null ? '' : before[field].toString();
            var newValue = after[field] === undefined || after[field] === null ? '' : after[field].toString();
            if (oldValue !== newValue) {
                changes.push({
                    pole: fieldLabels[field],
                    puvodne: oldValue,
                    nove: newValue
                });
            }
        });
        return changes;
    };

    var addHistory = function(record, changes) {
        if (!changes || (Object.prototype.toString.call(changes) === '[object Array]' && changes.length === 0)) return;
        record.historie = normalizeHistory(record.historie);
        var isStructured = Object.prototype.toString.call(changes) === '[object Array]';
        record.historie.push({
            kdy: getTimestamp(),
            kdo: getCurrentUser(),
            co: isStructured ? 'Změna záznamu' : changes,
            zmeny: isStructured ? changes : parseHistoryText(changes)
        });
    };

    var saveEditedRecord = function(formRecord) {
        var record = data.records.find(function(r) { return r.id.toString() === editState.id.toString(); });
        if (!record) return;

        var changes = describeChanges(editState.original || record, formRecord);
        if (changes.length === 0) {
            showStatus('Nejsou žádné změny k uložení', 'info');
            return;
        }

        formFields.forEach(function(field) {
            record[field] = formRecord[field];
        });
        addHistory(record, changes);

        if (config.useLocalStorage) {
            saveToLocalStorage();
        } else if (config.siteUrl && record.id) {
            updateInSharePoint(record);
        }

        editState.id = null;
        editState.original = null;
        updateEditControls(false);
        clearForm();
        renderTable();
        showStatus('Záznam změněn a historie uložena', 'success');
    };
    
    // Získání CSS třídy podle RPN hodnoty
    var getRPNClass = function(rpn) {
        if (rpn >= rpnThresholds.high) return 'rpn-high';
        if (rpn >= rpnThresholds.medium) return 'rpn-medium';
        return 'rpn-low';
    };

    var getRiskLabel = function(rpn) {
        if (rpn >= rpnThresholds.high) return 'Vysoké';
        if (rpn >= rpnThresholds.medium) return 'Střední';
        return 'Mírné';
    };

    var requiresAction = function(record, rpn) {
        return record.zavaznost === ratingMax || rpn >= rpnThresholds.high;
    };
    
    // Vykreslení tabulky
    var renderTable = function() {
        var tableContent = document.getElementById('tableContent');
        var recordCount = document.getElementById('recordCount');
        if (recordCount) {
            recordCount.textContent = data.records.length;
        }
        
        if (data.records.length === 0) {
            if (tableContent) {
                tableContent.innerHTML = '<p style="text-align: center; color: #999; padding: 40px;">Zatím nejsou žádné záznamy. Přidejte první záznam formulářem výše.</p>';
            }
            return;
        }

        var html = '<table><thead><tr>' +
                   '<th style="width: 10%">Krok procesu</th>' +
                   '<th style="width: 12%">FM</th>' +
                   '<th style="width: 12%">Účinek</th>' +
                   '<th style="width: 12%">Příčina</th>' +
                   '<th style="width: 12%">Prevence</th>' +
                   '<th style="width: 12%">Detekční kontrola</th>' +
                   '<th style="width: 8%">Závažnost</th>' +
                   '<th style="width: 8%">Výskyt</th>' +
                   '<th style="width: 8%">Odhalitelnost</th>' +
                   '<th style="width: 8%">RPN</th>' +
                   '<th style="width: 12%">Riziko / opatření</th>' +
                   '<th style="width: 12%">Vlastník / termín</th>' +
                   '<th style="width: 10%">Po opatření</th>' +
                   '<th style="width: 12%">Důkaz účinnosti</th>' +
                   '<th style="width: 8%">Akce</th>' +
                   '</tr></thead><tbody>';
        
        data.records.forEach(function(record) {
            var rpn = calculateRPN(record.zavaznost, record.vyskytu, record.detekce);
            var rpnClass = getRPNClass(rpn);
            var actionText = requiresAction(record, rpn) ? 'Opatření povinné' : 'Bez povinného opatření';
                var postRpn = (record.zavaznostPo && record.vyskytuPo && record.detekcePo) ? calculateRPN(record.zavaznostPo, record.vyskytuPo, record.detekcePo) : '';
            
                html += '<tr>' +
                    '<td>' + 
                    escapeHtml(record.proces) + '</td>' +
                    '<td>' + 
                    escapeHtml(record.selhani) + '</td>' +
                    '<td>' + 
                    escapeHtml(record.nasledky) + '</td>' +
                    '<td>' + 
                    escapeHtml(record.pricina || '') + '</td>' +
                    '<td>' + 
                    escapeHtml(record.kontrolaPrevence || '') + '</td>' +
                    '<td>' + 
                    escapeHtml(record.kontrolaDetekce || '') + '</td>' +
                    '<td style="text-align: center;">' + 
                    record.zavaznost + '</td>' +
                    '<td style="text-align: center;">' + 
                    record.vyskytu + '</td>' +
                    '<td style="text-align: center;">' + 
                    record.detekce + '</td>' +
                    '<td class="rpn-value ' + rpnClass + '">' + rpn + '</td>' +
                    '<td><strong>' + getRiskLabel(rpn) + '</strong><br><span class="decision-text">' + actionText + '</span><br>' + escapeHtml(record.opatreni || '') + '</td>' +
                    '<td>' + escapeHtml(record.vlastnikOpatreni || '') + '<br><span class="decision-text">' + escapeHtml(record.terminOpatreni || '') + '</span></td>' +
                    '<td>S\'/O\'/D\': ' + (record.zavaznostPo || '-') + '/' + (record.vyskytuPo || '-') + '/' + (record.detekcePo || '-') + '<br>RPN\': ' + (postRpn || '-') + '</td>' +
                    '<td>' + escapeHtml(record.dukazUcinnosti || '') + '</td>' +
                    '<td class="actions"><button type="button" class="btn-secondary btn-small" data-action="open-record-detail" data-record-id="' + record.id + '">Detail</button> <button type="button" class="btn-danger btn-small" onclick="fmeaApp.deleteRecord(' + record.id + ')">Smazat</button></td>' +
                    '</tr>';
        });
        
        html += '</tbody></table>';
        if (tableContent) {
            tableContent.innerHTML = html;
            bindRenderedDetailButtons(tableContent);
        }
        
        // Zavolej updateStats pokud existuje (pro testovací prostředí)
        if (typeof window.fmeaAppUpdateStats === 'function') {
            setTimeout(window.fmeaAppUpdateStats, 50);
        }
    };

    var bindRenderedDetailButtons = function(container) {
        var buttons = container.querySelectorAll ? container.querySelectorAll('[data-action="open-record-detail"]') : [];
        for (var i = 0; i < buttons.length; i++) {
            buttons[i].onclick = function(event) {
                event.preventDefault();
                event.stopPropagation();
                openRecordDetail(this.getAttribute('data-record-id'));
            };
        }
    };
    
    // Spuštění editace inline
    var startEdit = function(id, field) {
        var record = data.records.find(function(r) { return r.id === id; });
        if (!record) return;
        
        var value = record[field];
        var ratingFields = ['zavaznost', 'vyskytu', 'detekce', 'zavaznostPo', 'vyskytuPo', 'detekcePo'];
        var longTextFields = ['selhani', 'nasledky', 'pricina', 'kontrolaPrevence', 'kontrolaDetekce', 'opatreni', 'dukazUcinnosti'];
        var inputType = (ratingFields.indexOf(field) !== -1) ? 'number' : 'text';
        var inputTag = (longTextFields.indexOf(field) !== -1) ? 'textarea' : 'input';
        
        var input = document.createElement(inputTag);
        input.type = inputType;
        input.value = value;
        if (inputType === 'number') {
            input.min = ratingMin;
            input.max = ratingMax;
        }
        
        if (inputTag === 'textarea') {
            input.style.minHeight = '60px';
        }
        
        var cell = event.target;
        cell.innerHTML = '';
        cell.appendChild(input);
        input.focus();
        
        input.onblur = function() {
            var newValue = input.value;
            if (newValue !== value.toString()) {
                updateRecord(id, field, newValue);
            } else {
                renderTable();
            }
        };
        
        input.onkeydown = function(e) {
            if (e.key === 'Enter' && inputTag !== 'textarea') {
                input.blur();
            }
            if (e.key === 'Escape') {
                renderTable();
            }
        };
    };

    var renderHistory = function(history) {
        history = normalizeHistory(history);
        if (history.length === 0) return '<span class="decision-text">Bez historie</span>';
        var visible = history.slice(Math.max(history.length - 3, 0));
        return '<div class="history-list">' + visible.map(function(item) {
            var changes = getHistoryChanges(item);
            var summary = changes.length > 0 ? changes.length + ' změn' : (item.co || 'Záznam');
            return '<div><strong>' + escapeHtml(item.kdo || '') + '</strong><br><span class="decision-text">' + escapeHtml(item.kdy || '') + '</span><br>' + escapeHtml(summary) + '</div>';
        }).join('') + '</div>';
    };

    var renderFullHistory = function(history) {
        history = normalizeHistory(history);
        if (history.length === 0) return '<p class="decision-text">Zatím nejsou evidované žádné změny.</p>';
        return '<div class="history-list detail-history-list">' + history.slice().reverse().map(function(item) {
            return '<div class="history-entry"><div class="history-entry-head"><strong>' + escapeHtml(item.kdo || '') + '</strong><span>' + escapeHtml(item.kdy || '') + '</span></div>' + renderHistoryChanges(item) + '</div>';
        }).join('') + '</div>';
    };

    var getHistoryChanges = function(item) {
        if (item && Object.prototype.toString.call(item.zmeny) === '[object Array]') return item.zmeny;
        return parseHistoryText(item ? item.co : '');
    };

    var parseHistoryText = function(text) {
        if (!text || text.indexOf(' -> ') === -1) return [];
        return text.split('; ').map(function(part) {
            var match = part.match(/^([^:]+):\s*"([\s\S]*)"\s*->\s*"([\s\S]*)"$/);
            if (!match) return null;
            return {
                pole: match[1],
                puvodne: match[2],
                nove: match[3]
            };
        }).filter(function(change) { return change !== null; });
    };

    var renderHistoryChanges = function(item) {
        var changes = getHistoryChanges(item);
        if (changes.length === 0) return '<p class="history-note">' + escapeHtml(item.co || 'Záznam vytvořen') + '</p>';
        return '<table class="history-change-table"><thead><tr><th>Pole</th><th>Původně</th><th>Nově</th></tr></thead><tbody>' + changes.map(function(change) {
            return '<tr><td>' + escapeHtml(change.pole || '') + '</td><td>' + escapeHtml(change.puvodne || '-') + '</td><td>' + escapeHtml(change.nove || '-') + '</td></tr>';
        }).join('') + '</tbody></table>';
    };

    var detailRow = function(label, value) {
        return '<div class="detail-row"><strong>' + escapeHtml(label) + '</strong><span>' + escapeHtml(value || '-') + '</span></div>';
    };

    var openRecordDetail = function(id) {
        var record = data.records.find(function(r) { return r.id.toString() === id.toString(); });
        var modal = document.getElementById('recordDetailModal');
        var content = document.getElementById('recordDetailContent');
        var history = document.getElementById('recordHistoryContent');
        var editButton = document.getElementById('modalEditButton');
        if (!record || !modal || !content || !history || !editButton) return;

        var rpn = calculateRPN(record.zavaznost, record.vyskytu, record.detekce);
        var postRpn = (record.zavaznostPo && record.vyskytuPo && record.detekcePo) ? calculateRPN(record.zavaznostPo, record.vyskytuPo, record.detekcePo) : '';
        content.innerHTML =
            '<div class="detail-grid">' +
            '<div class="detail-card"><h3>Identifikace</h3>' +
            detailRow('Krok procesu', record.proces) +
            detailRow('FM', record.selhani) +
            detailRow('Účinek', record.nasledky) +
            detailRow('Příčina', record.pricina) +
            '</div>' +
            '<div class="detail-card"><h3>Kontroly a hodnocení</h3>' +
            detailRow('Kontrola prevence', record.kontrolaPrevence) +
            detailRow('Kontrola detekce', record.kontrolaDetekce) +
            detailRow('S/O/D', record.zavaznost + '/' + record.vyskytu + '/' + record.detekce) +
            detailRow('RPN', rpn + ' - ' + getRiskLabel(rpn)) +
            '</div>' +
            '<div class="detail-card"><h3>Opatření</h3>' +
            detailRow('Návrh opatření', record.opatreni) +
            detailRow('Vlastník', record.vlastnikOpatreni) +
            detailRow('Termín', record.terminOpatreni) +
            detailRow('Důkaz účinnosti', record.dukazUcinnosti) +
            '</div>' +
            '<div class="detail-card"><h3>Po opatření</h3>' +
            detailRow('S\'/O\'/D\'', (record.zavaznostPo || '-') + '/' + (record.vyskytuPo || '-') + '/' + (record.detekcePo || '-')) +
            detailRow('RPN\'', postRpn || '-') +
            '</div>' +
            '</div>';
        history.innerHTML = renderFullHistory(record.historie);
        editButton.onclick = function() {
            closeRecordDetail();
            loadRecordToForm(id);
        };
        modal.style.display = 'flex';
    };

    var closeRecordDetail = function() {
        var modal = document.getElementById('recordDetailModal');
        if (modal) modal.style.display = 'none';
    };

    var loadRecordToForm = function(id) {
        var record = data.records.find(function(r) { return r.id.toString() === id.toString(); });
        if (!record) return;
        formFields.forEach(function(field) {
            setInputValue(field, record[field] || '');
        });
        editState.id = record.id;
        editState.original = cloneRecord(record);
        updateEditControls(true);
        showStatus('Záznam načten do formuláře. Po úpravě klikněte na Změnit záznam.', 'info');
        var proces = document.getElementById('proces');
        if (proces) proces.focus();
    };

    var updateEditControls = function(isEditing) {
        var button = document.getElementById('saveRecordButton');
        var cancelButton = document.getElementById('cancelEditButton');
        var title = document.getElementById('formTitle');
        if (button) button.textContent = isEditing ? 'Změnit záznam' : 'Přidat záznam';
        if (cancelButton) cancelButton.style.display = isEditing ? 'inline-block' : 'none';
        if (title) title.textContent = isEditing ? 'Úprava záznamu' : 'Přidání nového záznamu';
    };

    var cancelEdit = function() {
        editState.id = null;
        editState.original = null;
        updateEditControls(false);
        clearForm();
        showStatus('Úprava zrušena', 'info');
    };
    
    // Vyčistit formulář
    var clearForm = function() {
        document.getElementById('proces').value = '';
        document.getElementById('selhani').value = '';
        document.getElementById('nasledky').value = '';
        setInputValue('pricina', '');
        setInputValue('kontrolaPrevence', '');
        setInputValue('kontrolaDetekce', '');
        document.getElementById('zavaznost').value = '3';
        document.getElementById('vyskytu').value = '3';
        document.getElementById('detekce').value = '3';
        setInputValue('opatreni', '');
        setInputValue('vlastnikOpatreni', '');
        setInputValue('terminOpatreni', '');
        setInputValue('dukazUcinnosti', '');
        setInputValue('zavaznostPo', '');
        setInputValue('vyskytuPo', '');
        setInputValue('detekcePo', '');
        editState.id = null;
        editState.original = null;
        updateEditControls(false);
        document.getElementById('proces').focus();
    };

    var setInputValue = function(id, value) {
        var element = document.getElementById(id);
        if (element) element.value = value;
    };

    var printReport = function() {
        document.body.className = document.body.className.replace(/\bprint-mode\b/g, '').trim();
        document.body.className += (document.body.className ? ' ' : '') + 'print-mode';
        window.print();
    };

    var saveAsPDF = function() {
        printReport();
        showStatus('V tiskovém dialogu vyberte možnost Uložit jako PDF', 'info');
    };
    
    // Zobrazení statusové zprávy
    var showStatus = function(message, type) {
        var element = document.getElementById('statusMessage');
        element.textContent = message;
        element.className = 'status-message ' + type;
        
        if (type === 'success' || type === 'info') {
            setTimeout(function() {
                element.className = 'status-message';
            }, 5000);
        }
    };
    
    // Escapování HTML
    var escapeHtml = function(text) {
        var map = {
            '&': '&amp;',
            '<': '&lt;',
            '>': '&gt;',
            '"': '&quot;',
            "'": '&#039;'
        };
        return String(text === null || text === undefined ? '' : text).replace(/[&<>"']/g, function(m) { return map[m]; });
    };
    
    // Veřejné rozhraní
    return {
        init: init,
        addRecord: addRecord,
        deleteRecord: deleteRecord,
        updateRecord: updateRecord,
        startEdit: startEdit,
        openRecordDetail: openRecordDetail,
        closeRecordDetail: closeRecordDetail,
        loadRecordToForm: loadRecordToForm,
        cancelEdit: cancelEdit,
        clearForm: clearForm,
        renderTable: renderTable,
        loadRecords: loadRecords,
        showStatus: showStatus,
        printReport: printReport,
        saveAsPDF: saveAsPDF,
        data: data,  // Exposuj data pro vnější přístup
        calculateRPN: calculateRPN,  // Exposuj RPN kalkulátor
        getRPNClass: getRPNClass  // Exposuj CSS třídu helper
    };
})();

// Inicializace při načtení DOM
if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', fmeaApp.init);
} else {
    fmeaApp.init();
}
