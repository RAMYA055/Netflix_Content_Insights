-- ============================================================
--  Netflix Content Insights — SQL Query Library
--  Database: MySQL 8+ / PostgreSQL 14+ / SQL Server 2019+
--  Author  : Data Analytics Portfolio Project
--  Updated : 2024
-- ============================================================


-- ============================================================
--  SECTION 1: DATABASE SCHEMA
-- ============================================================

CREATE TABLE IF NOT EXISTS netflix_titles (
    show_id        VARCHAR(20)   PRIMARY KEY,
    content_type   VARCHAR(10)   NOT NULL,          -- 'Movie' or 'TV Show'
    title          VARCHAR(300)  NOT NULL,
    director       VARCHAR(300),
    cast_members   TEXT,
    country        VARCHAR(300),
    date_added     DATE,
    release_year   INT,
    rating         VARCHAR(20),
    duration       VARCHAR(20),
    listed_in      VARCHAR(300),                    -- genres (comma-separated)
    description    TEXT,
    year_added     INT GENERATED ALWAYS AS (YEAR(date_added)) STORED,
    month_added    INT GENERATED ALWAYS AS (MONTH(date_added)) STORED
);

CREATE INDEX idx_content_type  ON netflix_titles(content_type);
CREATE INDEX idx_year_added    ON netflix_titles(year_added);
CREATE INDEX idx_release_year  ON netflix_titles(release_year);
CREATE INDEX idx_rating        ON netflix_titles(rating);
CREATE INDEX idx_country       ON netflix_titles(country);

-- Genre dimension table (normalised from comma-separated listed_in)
CREATE TABLE IF NOT EXISTS netflix_genres AS
SELECT
    show_id,
    TRIM(SUBSTRING_INDEX(SUBSTRING_INDEX(listed_in, ',', n.n), ',', -1)) AS genre
FROM netflix_titles
CROSS JOIN (
    SELECT 1 AS n UNION SELECT 2 UNION SELECT 3 UNION SELECT 4 UNION SELECT 5
) n
WHERE n.n <= 1 + (LENGTH(listed_in) - LENGTH(REPLACE(listed_in, ',', '')));

CREATE INDEX idx_genre ON netflix_genres(genre);


-- ============================================================
--  SECTION 2: EXECUTIVE KPI QUERIES
-- ============================================================

-- Q01: Executive KPI Summary — Single-query dashboard card feed
SELECT
    COUNT(*)                                          AS total_titles,
    SUM(content_type = 'Movie')                       AS total_movies,
    SUM(content_type = 'TV Show')                     AS total_tv_shows,
    ROUND(SUM(content_type = 'Movie') * 100.0
          / COUNT(*), 1)                              AS movie_pct,
    ROUND(SUM(content_type = 'TV Show') * 100.0
          / COUNT(*), 1)                              AS tv_pct,
    COUNT(DISTINCT NULLIF(director, ''))              AS unique_directors,
    COUNT(DISTINCT country)                           AS countries_represented,
    MIN(release_year)                                 AS oldest_content,
    MAX(release_year)                                 AS newest_content,
    MIN(year_added)                                   AS first_catalog_year,
    MAX(year_added)                                   AS latest_catalog_year
FROM netflix_titles;


-- Q02: Annual Content Addition Trend — Power BI line chart
SELECT
    year_added,
    COUNT(*)                                          AS titles_added,
    SUM(content_type = 'Movie')                       AS movies_added,
    SUM(content_type = 'TV Show')                     AS shows_added,
    ROUND(COUNT(*) * 100.0 /
          LAG(COUNT(*)) OVER (ORDER BY year_added) - 100, 1) AS yoy_growth_pct,
    SUM(COUNT(*)) OVER (ORDER BY year_added
        ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW)
                                                      AS cumulative_titles
FROM netflix_titles
WHERE year_added IS NOT NULL
GROUP BY year_added
ORDER BY year_added;


-- Q03: Content Growth Rate by Type — YoY breakdown for DAX mirror
WITH yearly AS (
    SELECT
        year_added,
        content_type,
        COUNT(*) AS titles
    FROM netflix_titles
    WHERE year_added IS NOT NULL
    GROUP BY year_added, content_type
)
SELECT
    year_added,
    content_type,
    titles,
    LAG(titles) OVER (PARTITION BY content_type ORDER BY year_added) AS prev_year,
    ROUND((titles - LAG(titles) OVER (PARTITION BY content_type ORDER BY year_added))
          * 100.0 /
          NULLIF(LAG(titles) OVER (PARTITION BY content_type ORDER BY year_added), 0), 1)
                                                      AS growth_rate_pct
FROM yearly
ORDER BY content_type, year_added;


-- ============================================================
--  SECTION 3: GENRE ANALYSIS
-- ============================================================

-- Q04: Genre Distribution — Top genres across full catalog
SELECT
    g.genre,
    COUNT(*)                                          AS title_count,
    SUM(t.content_type = 'Movie')                     AS movie_count,
    SUM(t.content_type = 'TV Show')                   AS show_count,
    ROUND(COUNT(*) * 100.0 /
          (SELECT COUNT(*) FROM netflix_titles), 2)   AS catalog_share_pct,
    ROUND(AVG(t.release_year), 1)                     AS avg_release_year
FROM netflix_genres g
JOIN netflix_titles t USING (show_id)
GROUP BY g.genre
ORDER BY title_count DESC
LIMIT 20;


-- Q05: Genre Growth by Year — Trend analysis (2019-2022)
SELECT
    t.year_added,
    g.genre,
    COUNT(*)                                          AS title_count
FROM netflix_genres g
JOIN netflix_titles t USING (show_id)
WHERE t.year_added BETWEEN 2019 AND 2022
  AND g.genre IN (
      SELECT genre FROM netflix_genres
      GROUP BY genre ORDER BY COUNT(*) DESC LIMIT 10
  )
GROUP BY t.year_added, g.genre
ORDER BY t.year_added, title_count DESC;


-- Q06: Underserved Genre Detection
--   Genres with high recent growth (2021-2022) but low catalog share
--   = high viewership potential / underserved segments
WITH genre_totals AS (
    SELECT
        g.genre,
        COUNT(*)                                      AS total_count,
        SUM(t.year_added IN (2021, 2022))             AS recent_count,
        SUM(t.year_added < 2021)                      AS legacy_count
    FROM netflix_genres g
    JOIN netflix_titles t USING (show_id)
    WHERE t.year_added IS NOT NULL
    GROUP BY g.genre
    HAVING COUNT(*) >= 5
),
scored AS (
    SELECT *,
        ROUND(recent_count * 100.0 / NULLIF(total_count, 0), 1)
                                                      AS recent_share_pct,
        ROUND(total_count * 100.0 /
              (SELECT COUNT(*) FROM netflix_titles), 2)
                                                      AS catalog_share_pct,
        ROUND(recent_count * 100.0 / NULLIF(total_count, 0), 1) -
        ROUND(total_count * 100.0 /
              (SELECT COUNT(*) FROM netflix_titles), 2)
                                                      AS opportunity_gap
    FROM genre_totals
)
SELECT *,
    CASE
        WHEN opportunity_gap > 15 THEN 'High Opportunity'
        WHEN opportunity_gap BETWEEN 5 AND 15 THEN 'Medium Opportunity'
        ELSE 'Saturated'
    END AS segment_status
FROM scored
ORDER BY opportunity_gap DESC
LIMIT 15;


-- Q07: Genre × Content Type Matrix
SELECT
    g.genre,
    SUM(t.content_type = 'Movie')                     AS movies,
    SUM(t.content_type = 'TV Show')                   AS tv_shows,
    COUNT(*)                                          AS total,
    ROUND(SUM(t.content_type = 'TV Show') * 100.0
          / NULLIF(COUNT(*), 0), 1)                   AS tv_dominance_pct
FROM netflix_genres g
JOIN netflix_titles t USING (show_id)
GROUP BY g.genre
HAVING COUNT(*) >= 20
ORDER BY tv_dominance_pct DESC;


-- ============================================================
--  SECTION 4: RATINGS & AUDIENCE TARGETING
-- ============================================================

-- Q08: Rating Distribution & Audience Segmentation
SELECT
    rating,
    COUNT(*)                                          AS title_count,
    SUM(content_type = 'Movie')                       AS movies,
    SUM(content_type = 'TV Show')                     AS tv_shows,
    ROUND(COUNT(*) * 100.0 / (SELECT COUNT(*) FROM netflix_titles), 1)
                                                      AS catalog_pct,
    CASE rating
        WHEN 'G'    THEN 'All Ages'
        WHEN 'PG'   THEN 'Family'
        WHEN 'PG-13' THEN 'Teen+'
        WHEN 'TV-G' THEN 'All Ages'
        WHEN 'TV-Y' THEN 'Children'
        WHEN 'TV-Y7' THEN 'Kids 7+'
        WHEN 'TV-PG' THEN 'Family'
        WHEN 'TV-14' THEN 'Teen+'
        WHEN 'TV-MA' THEN 'Mature'
        WHEN 'R'    THEN 'Mature'
        WHEN 'NC-17' THEN 'Adults'
        ELSE 'Unknown'
    END AS audience_segment
FROM netflix_titles
WHERE rating IS NOT NULL AND rating != ''
GROUP BY rating
ORDER BY title_count DESC;


-- Q09: Rating Trend Over Time — Is catalog maturing?
SELECT
    year_added,
    SUM(rating IN ('TV-MA','R','NC-17'))              AS mature_count,
    SUM(rating IN ('TV-14','PG-13'))                  AS teen_count,
    SUM(rating IN ('TV-G','TV-Y','TV-Y7','TV-PG','G','PG'))
                                                      AS family_count,
    ROUND(SUM(rating IN ('TV-MA','R','NC-17')) * 100.0
          / NULLIF(COUNT(*), 0), 1)                   AS mature_pct
FROM netflix_titles
WHERE year_added IS NOT NULL AND rating IS NOT NULL
GROUP BY year_added
ORDER BY year_added;


-- ============================================================
--  SECTION 5: GEOGRAPHIC & DIRECTOR ANALYSIS
-- ============================================================

-- Q10: Top Countries by Content Volume
SELECT
    TRIM(country)                                     AS country,
    COUNT(*)                                          AS title_count,
    SUM(content_type = 'Movie')                       AS movies,
    SUM(content_type = 'TV Show')                     AS tv_shows,
    ROUND(AVG(release_year), 0)                       AS avg_release_year,
    ROUND(COUNT(*) * 100.0 /
          (SELECT COUNT(*) FROM netflix_titles), 2)   AS catalog_share_pct
FROM netflix_titles
WHERE country IS NOT NULL AND country != ''
  AND country NOT LIKE '%,%'               -- single-country entries only
GROUP BY TRIM(country)
ORDER BY title_count DESC
LIMIT 20;


-- Q11: Director Productivity & Genre Specialisation
SELECT
    director,
    COUNT(*)                                          AS total_titles,
    COUNT(DISTINCT content_type)                      AS content_types,
    MIN(release_year)                                 AS career_start,
    MAX(release_year)                                 AS latest_release,
    MAX(release_year) - MIN(release_year)             AS career_span_years,
    GROUP_CONCAT(DISTINCT content_type)               AS types_directed
FROM netflix_titles
WHERE director IS NOT NULL AND director != ''
  AND director NOT LIKE '%,%'
GROUP BY director
HAVING COUNT(*) >= 3
ORDER BY total_titles DESC
LIMIT 20;


-- ============================================================
--  SECTION 6: ENGAGEMENT PROXIES & ADVANCED ANALYTICS
-- ============================================================

-- Q12: Duration Analysis — Movie runtime distribution
SELECT
    CASE
        WHEN CAST(REPLACE(duration,' min','') AS UNSIGNED) < 60   THEN 'Short (<60 min)'
        WHEN CAST(REPLACE(duration,' min','') AS UNSIGNED) < 90   THEN 'Standard (60-90)'
        WHEN CAST(REPLACE(duration,' min','') AS UNSIGNED) < 120  THEN 'Feature (90-120)'
        WHEN CAST(REPLACE(duration,' min','') AS UNSIGNED) < 150  THEN 'Long (120-150)'
        ELSE 'Epic (150+ min)'
    END                                               AS duration_bucket,
    COUNT(*)                                          AS title_count,
    ROUND(AVG(CAST(REPLACE(duration,' min','') AS UNSIGNED)), 1)
                                                      AS avg_runtime_min,
    MIN(CAST(REPLACE(duration,' min','') AS UNSIGNED))  AS min_runtime,
    MAX(CAST(REPLACE(duration,' min','') AS UNSIGNED))  AS max_runtime
FROM netflix_titles
WHERE content_type = 'Movie'
  AND duration LIKE '%min%'
GROUP BY duration_bucket
ORDER BY avg_runtime_min;


-- Q13: TV Show Season Depth — Engagement proxy (more seasons = more investment)
SELECT
    CASE
        WHEN CAST(REPLACE(duration,' Season','') AS UNSIGNED) = 1  THEN '1 Season'
        WHEN CAST(REPLACE(duration,' Season','') AS UNSIGNED) <= 3 THEN '2-3 Seasons'
        WHEN CAST(REPLACE(duration,' Season','') AS UNSIGNED) <= 6 THEN '4-6 Seasons'
        ELSE '7+ Seasons'
    END                                               AS season_tier,
    COUNT(*)                                          AS show_count,
    ROUND(AVG(CAST(REPLACE(duration,' Season','') AS UNSIGNED)), 2)
                                                      AS avg_seasons
FROM netflix_titles
WHERE content_type = 'TV Show'
  AND duration LIKE '%Season%'
GROUP BY season_tier
ORDER BY avg_seasons;


-- Q14: Catalog Freshness Index — Content recency vs release year
SELECT
    year_added,
    ROUND(AVG(year_added - release_year), 1)          AS avg_age_at_addition,
    MIN(year_added - release_year)                    AS min_age,
    MAX(year_added - release_year)                    AS max_age,
    SUM(year_added - release_year <= 1)               AS same_year_releases,
    SUM(year_added - release_year > 5)                AS older_catalogue,
    ROUND(SUM(year_added - release_year <= 1) * 100.0
          / NULLIF(COUNT(*), 0), 1)                   AS fresh_content_pct
FROM netflix_titles
WHERE year_added IS NOT NULL AND release_year IS NOT NULL
GROUP BY year_added
ORDER BY year_added;


-- Q15: Monthly Addition Seasonality — Best months for launches
SELECT
    month_added,
    MONTHNAME(STR_TO_DATE(CONCAT('2000-', month_added, '-01'), '%Y-%m-%d'))
                                                      AS month_name,
    COUNT(*)                                          AS titles_added,
    SUM(content_type = 'Movie')                       AS movies,
    SUM(content_type = 'TV Show')                     AS tv_shows,
    ROUND(COUNT(*) * 100.0 /
          SUM(COUNT(*)) OVER (), 1)                   AS pct_of_annual
FROM netflix_titles
WHERE month_added IS NOT NULL
GROUP BY month_added
ORDER BY month_added;


-- Q16: Underserved Genres Deep Dive — 3 Identified Segments
--  Segment 1: Docuseries
SELECT
    'Docuseries'                                      AS segment,
    COUNT(*)                                          AS total_titles,
    ROUND(COUNT(*) * 100.0 /
          (SELECT COUNT(*) FROM netflix_genres WHERE genre = 'Docuseries'), 1)
                                                      AS recency_rate,
    'High viewership intent; low catalog depth'       AS insight
FROM netflix_genres g
JOIN netflix_titles t USING (show_id)
WHERE g.genre = 'Docuseries'
UNION ALL
--  Segment 2: Stand-Up Comedy
SELECT
    'Stand-Up Comedy',
    COUNT(*),
    ROUND(SUM(t.year_added >= 2021) * 100.0 / NULLIF(COUNT(*), 0), 1),
    'Rapid fan communities; underrepresented in search'
FROM netflix_genres g
JOIN netflix_titles t USING (show_id)
WHERE g.genre = 'Stand-Up Comedy'
UNION ALL
--  Segment 3: International Thrillers
SELECT
    'International Thrillers',
    COUNT(*),
    ROUND(SUM(t.year_added >= 2021) * 100.0 / NULLIF(COUNT(*), 0), 1),
    'Post-Squid Game surge; untapped global demand'
FROM netflix_genres g
JOIN netflix_titles t USING (show_id)
WHERE g.genre IN ('International TV Shows','Thrillers')
  AND t.country NOT IN ('United States','United Kingdom');


-- Q17: Content Correlation Matrix — Multi-metric segment scoring
WITH segment_scores AS (
    SELECT
        g.genre,
        COUNT(*)                                      AS volume,
        ROUND(SUM(t.year_added >= 2021) * 100.0
              / NULLIF(COUNT(*), 0), 1)               AS recency_pct,
        ROUND(AVG(CASE
            WHEN t.content_type = 'TV Show'
                 AND t.duration LIKE '%Season%'
                 AND CAST(REPLACE(t.duration,' Season','') AS UNSIGNED) > 2
                 THEN 80
            WHEN t.content_type = 'Movie' THEN 60
            ELSE 40 END), 1)                          AS engagement_proxy,
        COUNT(DISTINCT t.country)                     AS country_diversity
    FROM netflix_genres g
    JOIN netflix_titles t USING (show_id)
    WHERE t.year_added IS NOT NULL
    GROUP BY g.genre
    HAVING COUNT(*) >= 10
)
SELECT *,
    ROUND((recency_pct * 0.40 +
           engagement_proxy * 0.35 +
           country_diversity * 0.25), 1)              AS composite_opportunity_score
FROM segment_scores
ORDER BY composite_opportunity_score DESC
LIMIT 15;
