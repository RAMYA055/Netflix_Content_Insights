# Netflix Content Insights Dashboard

![SQL](https://img.shields.io/badge/SQL-MySQL%208+-4479A1?style=for-the-badge&logo=mysql&logoColor=white)
![Power BI](https://img.shields.io/badge/Power%20BI-Dashboard-F2C811?style=for-the-badge&logo=powerbi&logoColor=black)
![DAX](https://img.shields.io/badge/DAX-18%20Measures-E50914?style=for-the-badge)
![PowerPoint](https://img.shields.io/badge/PowerPoint-10%20Slides-B7472A?style=for-the-badge&logo=microsoftpowerpoint&logoColor=white)

---

## Project Overview

End-to-end analytics project analyzing Netflix's content catalog (2019–2022) to identify growth patterns, underserved genre segments, and strategic content investment opportunities.

> **Resume Bullets:**
> - Authored 15+ optimized SQL queries to extract and aggregate Netflix content metadata; built an 8-visual Power BI dashboard using DAX measures for content growth rate, genre distribution, and engagement proxies across 3 years of catalog data
> - Identified 3 underserved genre segments with high viewership potential (Stand-Up Comedy, Docuseries, International Thrillers) and documented recommendations in an executive PowerPoint brief for non-technical stakeholders

---

## Key Findings

| Insight | Finding |
|---------|---------|
| Catalog Size | 10,060+ titles across 2019–2022 |
| Growth Peak | +22.5% YoY in 2021 (post-pandemic pipeline surge) |
| Top Genre | International Movies (27.4% catalog share) |
| Critical Segment | International Thrillers — 0.91 opportunity score |
| Maturity Trend | Mature content (TV-MA) grew +4.2pp from 2019→2022 |
| Family Gap | Family content at 13.4% — 21pp below Disney+ benchmark |

---

## Repository Structure

```
Netflix_Content_Insights/
│
├── README.md
│
├── sql/
│   └── netflix_queries.sql          <- 17 SQL queries + schema design
│
├── powerbi/
│   └── Netflix_Dashboard_Guide.md   <- Power BI implementation guide
│
└── Netflix_Content_Insights_Dashboard.xlsx   <- 7-tab Excel workbook
    Netflix_Content_Insights.pptx             <- 10-slide executive brief
```

---

## Project Components

### SQL (17 Queries)

| Query | Purpose | Technique |
|-------|---------|-----------|
| Q01 | Executive KPI Summary | Aggregation |
| Q02 | Annual Trend | LAG() + Window Functions |
| Q03 | Growth by Type | Partition + YoY |
| Q04 | Genre Distribution | GROUP BY + JOIN |
| Q05 | Genre Growth | Pivot-style |
| Q06 | **Underserved Detection** | CTE + Scoring |
| Q07 | Genre × Type Matrix | Conditional Aggregation |
| Q08 | Rating Segmentation | CASE + GROUP BY |
| Q09 | Rating Trend | Time-series |
| Q10 | Top Countries | Geographic filter |
| Q11 | Director Analysis | HAVING + Career span |
| Q12 | Movie Duration | CAST + Buckets |
| Q13 | TV Season Depth | Engagement proxy |
| Q14 | Catalog Freshness | Date arithmetic |
| Q15 | Monthly Seasonality | WINDOW functions |
| Q16 | Underserved Deep Dive | UNION ALL |
| Q17 | Composite Opportunity Score | CTE + Weighted formula |

### Power BI Dashboard (8 Visuals)

| Visual | Type | Key Metric |
|--------|------|------------|
| VIZ 1 | KPI Cards (×6) | Total titles, growth, countries |
| VIZ 2 | Line Chart | Annual content addition trend |
| VIZ 3 | Stacked Bar | Genre distribution by content type |
| VIZ 4 | Donut Chart | Ratings audience segmentation |
| VIZ 5 | Map Visual | Geographic content origins |
| VIZ 6 | Scatter Plot | Opportunity score matrix |
| VIZ 7 | Treemap | Genre hierarchy |
| VIZ 8 | Table Matrix | Underserved segment detail |

### DAX Measures (18 total)

Organized across 5 categories:
- **Core KPIs** — Total Titles, Movies, TV Shows, Movie Share %
- **Growth** — Content Growth Rate, Cumulative Titles, 3-Year CAGR
- **Genre** — Genre Title Count, Genre Share %, Underserved Score
- **Engagement** — Engagement Proxy, Avg Seasons
- **Time Intelligence** — Recent Growth %, Monthly Velocity
- **Audience** — Mature Content %, Family Content %, Catalog Freshness

### 3 Underserved Segments

| Segment | Catalog Share | Recent Growth | Score | Priority |
|---------|--------------|--------------|-------|---------|
| International Thrillers | 2.5% | +35.1% | 0.91 | CRITICAL |
| Stand-Up Comedy | 3.9% | +31.4% | 0.87 | CRITICAL |
| Docuseries | 3.8% | +28.7% | 0.82 | HIGH |

### PowerPoint (10 Slides)

1. Title Slide
2. Project Overview & Methodology
3. Executive KPI Summary + Data Model
4. Content Growth Trends (bar chart)
5. Genre Distribution Analysis
6. Ratings & Audience Segmentation
7. Geographic Distribution
8. **3 Underserved Segments** (key insight slide)
9. Power BI Technical Architecture + DAX
10. Strategic Recommendations

---

## How to Use

### SQL
1. Create schema using `CREATE TABLE` blocks at top of `netflix_queries.sql`
2. Import Netflix dataset (available on Kaggle: "Netflix Movies and TV Shows")
3. Run queries sequentially — Q01 first for baseline KPIs

### Power BI
1. Import raw Netflix CSV into Power BI Desktop
2. Create star schema (Fact + 3 Dims) per Q01-Q03 structure
3. Copy DAX measures from Excel Tab 6 into Power BI measure editor
4. Apply visual layout from Excel Tab 7 (SQL Query Index for visual mapping)

### Excel Workbook
- Tab 1: Executive Dashboard — KPI overview
- Tab 2: Content Growth — trend data
- Tab 3: Genre Analysis — distribution + underserved segments
- Tab 4: Ratings & Audience — segmentation
- Tab 5: Geographic Analysis — country breakdown
- Tab 6: DAX Measures — 18 production-ready formulas
- Tab 7: SQL Query Index — all 17 queries catalogued

---

## Google Drive Upload Structure

```
Netflix_Content_Insights/        <- Root folder
├── README.md
├── sql/
│   └── netflix_queries.sql
└── [root level]
    ├── Netflix_Content_Insights_Dashboard.xlsx
    └── Netflix_Content_Insights.pptx
```

---

## License

Portfolio project for educational purposes. Netflix dataset sourced from public Kaggle repository.
