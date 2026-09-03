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

## Synchronizace mezi zarizenimi (Supabase zdarma)

Do aplikace je pridany cloud sync panel na Dashboardu.

### 1) Vytvor Supabase projekt

- registrace: https://supabase.com
- vytvor novy projekt (free plan)

### 2) V SQL editoru spust tento skript

```sql
create table if not exists public.tasks_sync (
	id bigint generated always as identity primary key,
	profile_id text not null,
	task_id text not null,
	updated_at timestamptz not null default now(),
	task jsonb not null
);

create index if not exists idx_tasks_sync_profile on public.tasks_sync (profile_id);

alter table public.tasks_sync enable row level security;

drop policy if exists "tasks_sync_open" on public.tasks_sync;
drop policy if exists "tasks_sync_owner" on public.tasks_sync;
create policy "tasks_sync_owner"
on public.tasks_sync
for all
to authenticated
using (profile_id = auth.uid()::text)
with check (profile_id = auth.uid()::text);
```

Poznamka: tato politika je bezpecnejsi. Data uvidi jen prihlaseny vlastnik.

### 3) V aplikaci vypln cloud sync

- Supabase URL (`https://xxxx.supabase.co`)
- Anon key (`Project Settings -> API -> anon public`)
- E-mail + heslo pro prihlaseni
- V aplikaci klikni `Vytvorit ucet` (jen poprve) a potom `Prihlasit`

### 3b) Prihlaseni pres GitHub (volitelne)

1. V Supabase otevri `Authentication -> Providers -> GitHub` a zapni provider.
2. V GitHub Developers vytvor OAuth App:
	- Homepage URL: `https://michalsip-lang.github.io/code/`
	- Authorization callback URL: `https://<TVUJ-PROJECT-REF>.supabase.co/auth/v1/callback`
3. Client ID + Client Secret vloz zpet do Supabase GitHub provideru.
4. V Supabase otevri `Authentication -> URL Configuration`:
	- `Site URL`: `https://michalsip-lang.github.io/code/`
	- `Redirect URLs`: pridej `https://michalsip-lang.github.io/code/`
5. V appce klikni `Prihlasit pres GitHub`.

### 4) Synchronizace

- `Ulozit nastaveni`
- `Nahrat do cloudu` na prvnim zarizeni
- na druhem zarizeni se prihlas stejnym uctem a klikni `Nacist z cloudu`

Po kazde vetsi zmene staci jednou kliknout na `Nahrat do cloudu`.
