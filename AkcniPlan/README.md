# Akcni plan (.NET 8, ASP.NET Core MVC, SQLite)

Jednouzivatelska lokalni webova aplikace pro rizeni ukolu, priorit, denni planovani a osobni analytiku.

## Pouzite technologie

- .NET 8
- ASP.NET Core MVC
- Entity Framework Core
- SQLite
- Bootstrap 5
- Chart.js

## Architektura

- `Controllers/`
  - `HomeController`: dashboard a prehled KPI
  - `TasksController`: CRUD ukolu, auto-vytvareni z textu, stavy
  - `Api/AnalyticsController`: JSON endpointy pro grafy
- `Data/`
  - `AppDbContext`: databazovy kontext a mapovani relaci
  - `SeedData`: inicialni data
- `Models/`
  - `TaskItem`, `TaskTag`, `TaskItemTag`, `TaskDependency`, `TaskItemStatus`
- `Services/`
  - `PriorityScoringService`: vypocet priority 0-100
  - `AnalyticsService`: metriky a casove rady
  - `RecommendationService`: doporuceni podle dat
  - `DashboardService`: orchestrace dashboardu
  - `TaskAutomationService`: auto-vytvareni ukolu z textu
- `Views/` + `wwwroot/js/dashboard.js`: UI + vizualizace

## Databazove schema

### TaskItem
- `Id` (PK)
- `Title` (required)
- `Description`
- `CreatedAt`
- `DueDate`
- `EstimatedHours`
- `ActualHours`
- `Importance` (1-5)
- `Area` (`Svp`, `Sdp`, `Bozp`, `Po`, `Jine`)
- `Status` (`Todo`, `InProgress`, `Done`, `Blocked`)
- `PriorityScore` (0-100)
- `CompletedAt`

### TaskTag
- `Id` (PK)
- `Name` (unique)

### TaskItemTag (M:N)
- `TaskItemId` (FK)
- `TaskTagId` (FK)
- composite PK (`TaskItemId`, `TaskTagId`)

### TaskDependency (self M:N)
- `TaskItemId` (FK)
- `DependsOnTaskItemId` (FK)
- composite PK (`TaskItemId`, `DependsOnTaskItemId`)

## Prioritizace (0-100)

Priority score kombinuje:
- blizkost terminu (35%)
- odhad pracnosti (15%)
- dulezitost (30%)
- dny po terminu (10%)
- pocet zavislosti (10%)

Vypocet je implementovan v `PriorityScoringService`.

## Dashboard

Obsahuje:
- ukoly po terminu
- ukoly na dnesek
- ukoly na tento tyden
- dokoncene ukoly
- rozpracovane ukoly
- KPI plneni

## Analytika

Vyhodnocuje:
- dokoncene ukoly za den
- dokoncene ukoly za tyden
- dokoncene ukoly za mesic
- prumernou dobu dokonceni
- plneni terminu
- osobni produktivitu
- trend vykonnosti

## Grafy

- Line chart: denni a mesicni trend
- Bar chart: tydenni porovnani
- Donut chart: rozlozeni stavu
- Heatmap: kalendar produktivity za poslednich 120 dni

## Doporuceni

Aplikace navrhuje:
- ktere ukoly resit jako prvni
- pretizeni v nasledujicich dnech
- realistickou denni kapacitu
- odhad terminu dokonceni backlogu

## Spusteni

1. Nainstalujte .NET 8 SDK.
2. V root slozce projektu:

```bash
dotnet restore
dotnet run
```

3. Otevrete URL vypisane v terminalu (typicky `https://localhost:xxxx` nebo `http://localhost:xxxx`).

Databaze `akcni-plan.db` se vytvori automaticky pri prvnim spusteni.

## API

- `GET /api/analytics/series`
  - vraci casove rady pro den/tyden/mesic, stavy a heatmap data
- `GET /api/analytics/kpi`
  - vraci KPI hodnoty

## Proc MVC + SQLite

Pro lokalni single-user app je MVC + SQLite nejefektivnejsi:
- jednoduche nasazeni bez slozite infrastruktury
- rychly vyvoj i provoz
- snadne rozsireni o API, auth nebo sync v budoucnu

Alternativa Blazor Server je vhodna, pokud chcete vice interaktivni SPA UX bez psani JS. V tomto reseni je interaktivita pokryta kombinaci MVC + Chart.js s mensi slozitosti.
