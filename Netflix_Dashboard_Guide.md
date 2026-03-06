# Power BI Dashboard Implementation Guide
## Netflix Content Insights — 8-Visual Dashboard

---

## Step 1: Data Import

1. Open Power BI Desktop → **Get Data → Text/CSV**
2. Load `netflix_titles.csv` (Kaggle: "Netflix Movies and TV Shows")
3. Open **Power Query Editor**

### Power Query Transformations

```powerquery
// Split genre column (listed_in) into rows
= Table.TransformColumns(Source, {{"listed_in", each Text.Split(_, ", ")}})
= Table.ExpandListColumn(Source, "listed_in")
= Table.RenameColumns(Source, {{"listed_in", "genre"}})

// Parse date_added
= Table.TransformColumnTypes(Source, {{"date_added", type date}})

// Extract year and month
= Table.AddColumn(Source, "year_added", each Date.Year([date_added]))
= Table.AddColumn(Source, "month_added", each Date.Month([date_added]))
```

---

## Step 2: Data Model (Star Schema)

Create 3 dimension tables in Power Query:

### dim_date
```powerquery
let
    StartDate = #date(2019, 1, 1),
    EndDate = #date(2022, 12, 31),
    NumberOfDays = Duration.Days(EndDate - StartDate) + 1,
    Source = List.Dates(StartDate, NumberOfDays, #duration(1,0,0,0)),
    #"Table" = Table.FromList(Source, Splitter.SplitByNothing()),
    #"Renamed Columns" = Table.RenameColumns(#"Table", {{"Column1", "Date"}}),
    #"Changed Type" = Table.TransformColumnTypes(#"Renamed Columns", {{"Date", type date}}),
    #"Added Year" = Table.AddColumn(#"Changed Type", "Year", each Date.Year([Date])),
    #"Added Month" = Table.AddColumn(#"Added Year", "Month", each Date.Month([Date])),
    #"Added MonthName" = Table.AddColumn(#"Added Month", "MonthName", each Date.MonthName([Date])),
    #"Added Quarter" = Table.AddColumn(#"Added MonthName", "Quarter", each Date.QuarterOfYear([Date]))
in
    #"Added Quarter"
```

### dim_genre (from split genres)
- Distinct values of `genre` column after splitting

### dim_country (from country column)
- Distinct values of `country` (handle multi-country comma-splits)

---

## Step 3: Relationships

```
dim_date[Date]    → netflix_titles[date_added]    (1:Many)
dim_genre[genre]  → netflix_genres[genre]         (1:Many)
dim_country[country] → netflix_titles[country]    (1:Many)
```

---

## Step 4: DAX Measures

Paste these into the **Model** tab → **New Measure**:

### Core KPIs
```dax
Total Titles = COUNTROWS(NetflixTitles)

Total Movies = CALCULATE([Total Titles], NetflixTitles[content_type] = "Movie")

Total TV Shows = CALCULATE([Total Titles], NetflixTitles[content_type] = "TV Show")

Movie Share % = DIVIDE([Total Movies], [Total Titles], 0)
```

### Growth Measures
```dax
Content Growth Rate =
VAR CurrentYear = [Total Titles]
VAR PriorYear = CALCULATE([Total Titles],
    DATEADD(dim_date[Date], -1, YEAR))
RETURN DIVIDE(CurrentYear - PriorYear, PriorYear, 0)

Cumulative Titles =
CALCULATE([Total Titles],
    FILTER(ALL(dim_date),
        dim_date[Year] <= MAX(dim_date[Year])))

3Y CAGR =
VAR End = [Total Titles]
VAR Start = CALCULATE([Total Titles],
    DATEADD(dim_date[Date], -3, YEAR))
RETURN POWER(DIVIDE(End, Start, 0), 1/3) - 1
```

### Genre Intelligence
```dax
Genre Share % =
DIVIDE(
    COUNTROWS(NetflixGenres),
    CALCULATE(COUNTROWS(NetflixGenres), ALL(NetflixGenres[genre])),
    0)

Recent Growth % =
VAR Recent = CALCULATE([Total Titles], dim_date[Year] >= 2021)
RETURN DIVIDE(Recent, [Total Titles], 0)

Underserved Score =
VAR RecentShare = [Recent Growth %]
VAR CatalogShare = [Genre Share %]
RETURN RecentShare - CatalogShare
```

### Engagement Proxies
```dax
Engagement Proxy =
VAR MultiSeason = CALCULATE([Total Titles],
    NetflixTitles[duration_seasons] > 2)
VAR Movies = [Total Movies]
RETURN DIVIDE(MultiSeason + Movies * 0.6, [Total Titles], 0)

Avg Seasons =
CALCULATE(
    AVERAGE(NetflixTitles[duration_seasons]),
    NetflixTitles[content_type] = "TV Show")

Catalog Freshness =
AVERAGEX(
    NetflixTitles,
    NetflixTitles[year_added] - NetflixTitles[release_year])
```

### Audience Segmentation
```dax
Mature Content % =
DIVIDE(
    CALCULATE([Total Titles],
        NetflixTitles[rating] IN {"TV-MA", "R", "NC-17"}),
    [Total Titles], 0)

Family Content % =
DIVIDE(
    CALCULATE([Total Titles],
        NetflixTitles[rating] IN {"TV-G", "TV-Y", "TV-Y7", "TV-PG", "G", "PG"}),
    [Total Titles], 0)

Monthly Velocity =
DIVIDE([Total Titles],
    DISTINCTCOUNT(dim_date[YearMonth]))
```

---

## Step 5: Visual Layout

### Page 1: Overview Dashboard

| Position | Visual | Fields |
|----------|--------|--------|
| Top Row | KPI Cards ×6 | Total Titles, Movies, TV Shows, Growth Rate, Countries, Genres |
| Left Center | Line Chart | Year → Total Titles (Movies + TV Shows series) |
| Right Center | Donut Chart | content_type → Title Count |
| Bottom Left | Bar Chart | Genre → Title Count (Top 10) |
| Bottom Right | Map | Country → Title Count (bubble size) |

### Page 2: Genre Deep Dive

| Position | Visual | Fields |
|----------|--------|--------|
| Left | Stacked Bar | Genre × content_type |
| Right Top | Scatter Plot | Genre Share % vs Recent Growth % → Underserved Score (size) |
| Right Bottom | Table | Segment, Score, Priority, Recommendation |

### Slicers (right panel)
- Year slicer (2019-2022)
- Content Type slicer (Movie / TV Show)
- Rating slicer

---

## Step 6: Formatting

### Theme Colors
```json
{
  "name": "Netflix Dark",
  "dataColors": ["#E50914", "#F5A623", "#00A8E1", "#46D369", "#7B68EE", "#B20710"],
  "background": "#141414",
  "foreground": "#FFFFFF",
  "tableAccent": "#E50914"
}
```

### Chart Formatting
- Canvas background: `#141414` (Netflix black)
- Visual background: `#221F1F` (dark card)
- Text color: `#F5F5F1` (off-white)
- Accent color: `#E50914` (Netflix red)
- Grid lines: `#2D2D2D` (subtle)

---

## Expected Dashboard Output

After completing all steps, your dashboard should show:

- **8 interactive visuals** with cross-filtering enabled
- **18 DAX measures** in the Fields pane
- **Star schema** with 1 Fact + 3 Dim tables
- **Drill-through** from overview → genre detail page
- **Consistent dark Netflix theme** throughout
