# Nasazení Ribbon tlačítka (SharePoint 2013)

Soubor: `ribbon_tlacitko_seznam_prodejen.xml`

## Co to dělá
- Přidá tlačítko **Přehled BOZP/PO** do Ribbonu na **Display Form** (`DispForm.aspx`).
- Po kliknutí načte skript ze `SiteAssets` a zavolá funkci `runDalsiUrazyAkce()`.

## Důležité
- Cesta ke skriptu je nastavená na:
  - `/homeZVK/SiteAssets/prehled_urazu_dalsi_tlacitko.js`
- Pokud máš skript jinde, uprav v XML hodnotu `url` v `CommandAction`.

## Nasazení
Tento XML soubor nasazuj jako SharePoint Feature (`Elements.xml`) ve WSP řešení.

## Poznámka k rozsahu
- `RegistrationType="List"` + `RegistrationId="100"` znamená generic listy.
- Pokud chceš tlačítko jen pro jeden konkrétní seznam, dej vědět a připravím variantu podle konkrétního List GUID.
