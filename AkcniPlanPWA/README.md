# Akcni plan PWA (iPad)

Tato verze je samostatna aplikace bez backendu. Ukoly se ukladaji lokalne do IndexedDB primo v zarizeni.

## Co umi

- instalace na plochu iPadu (Safari -> Sdileni -> Pridat na plochu)
- offline rezim (service worker)
- lokalni uloziste ukolu (bez serveru)
- filtrace dle stavu, prepinani stavu, mazani

## Obsah slozky

- index.html
- styles.css
- app.js
- service-worker.js
- manifest.webmanifest
- icons/

## Nasazeni, aby to fungovalo bez Macu

Je potreba jen staticky hosting (bez backendu), napriklad:

- Cloudflare Pages
- Netlify
- GitHub Pages

Nahrat cely obsah slozky `AkcniPlanPWA` jako staticky web.

## GitHub Pages (bez Netlify)

V repozitari je pripraveny workflow pro automaticky deploy na GitHub Pages:

- `.github/workflows/pwa-pages.yml`

Postup:

1. Nahrajte zmeny do vetve `main`.
2. V GitHub repu otevrete `Settings -> Pages`.
3. U `Source` zvolte `GitHub Actions`.
4. Pockejte na dokonceni workflow `Deploy AkcniPlanPWA to GitHub Pages`.
5. Otevrite URL, kterou GitHub Pages vypise (typicky `https://michalsip-lang.github.io/code/`).

## Instalace na iPad

1. Otevri URL nasazene aplikace v Safari.
2. Klepni na Sdileni.
3. Zvol Pridat na plochu.
4. Spoustej z ikony jako bezna aplikace.

## Poznamky

- Data zustavaji jen v tom iPadu (a v tom prohlizeci).
- Kdyz smazes data webu v Safari, smazou se i ukoly.
- Pro synchronizaci mezi vice zarizenimi by bylo nutne doplnit cloud backend.

Posledni aktualizace nasazeni: 2026-09-03.
