"""
Generate a realistic Netflix content metadata CSV dataset (8,000+ rows, 2021-2023)
"""
import pandas as pd
import numpy as np
import random
from datetime import date, timedelta

np.random.seed(42)
random.seed(42)

GENRES = [
    "Drama", "Comedy", "Action", "Thriller", "Documentary",
    "Horror", "Romance", "Sci-Fi", "Animation", "Crime",
    "Fantasy", "Biography", "Family", "Mystery", "Reality"
]

COUNTRIES = [
    "United States", "United Kingdom", "India", "South Korea", "Spain",
    "France", "Germany", "Japan", "Brazil", "Canada",
    "Mexico", "Italy", "Australia", "Turkey", "Nigeria"
]

RATINGS = ["G", "PG", "PG-13", "TV-14", "TV-MA", "R", "TV-G", "TV-PG", "NC-17"]

LANGUAGES = ["English", "Spanish", "Hindi", "Korean", "French", "German", "Japanese", "Portuguese", "Turkish", "Italian"]

DIRECTORS = [f"Director_{i}" for i in range(1, 201)]
CAST_POOL  = [f"Actor_{i}" for i in range(1, 501)]

def rand_date(start=date(2021, 1, 1), end=date(2023, 12, 31)):
    return start + timedelta(days=random.randint(0, (end - start).days))

# Genre weights: some underserved
GENRE_WEIGHTS = {
    "Drama": 0.18, "Comedy": 0.13, "Action": 0.12, "Thriller": 0.10,
    "Documentary": 0.07, "Horror": 0.06, "Romance": 0.06, "Sci-Fi": 0.06,
    "Animation": 0.04, "Crime": 0.05, "Fantasy": 0.04, "Biography": 0.03,
    "Family": 0.03, "Mystery": 0.02, "Reality": 0.01
}

# Engagement scores by genre (higher = more engaged viewers per title)
GENRE_ENGAGEMENT = {
    "Drama": 72, "Comedy": 68, "Action": 75, "Thriller": 78,
    "Documentary": 65, "Horror": 82, "Romance": 70, "Sci-Fi": 85,
    "Animation": 74, "Crime": 80, "Fantasy": 88, "Biography": 60,
    "Family": 71, "Mystery": 84, "Reality": 58
}

rows = []
content_id = 1000

for _ in range(8500):
    genre = random.choices(list(GENRE_WEIGHTS.keys()), weights=list(GENRE_WEIGHTS.values()))[0]
    content_type = random.choices(["Movie", "TV Show"], weights=[0.55, 0.45])[0]
    added_date   = rand_date()
    release_year = added_date.year - random.randint(0, 3)
    duration_min = random.randint(80, 160) if content_type == "Movie" else None
    seasons      = random.randint(1, 8) if content_type == "TV Show" else None
    episodes     = seasons * random.randint(6, 12) if seasons else None
    country      = random.choices(COUNTRIES, weights=[0.28,0.10,0.12,0.09,0.07,0.06,0.05,0.05,0.04,0.04,0.03,0.02,0.02,0.02,0.01])[0]
    language     = random.choice(LANGUAGES)
    rating       = random.choices(RATINGS, weights=[0.04,0.08,0.18,0.20,0.28,0.12,0.03,0.04,0.01])[0]
    base_eng     = GENRE_ENGAGEMENT[genre]
    engagement   = max(10, min(100, base_eng + np.random.normal(0, 12)))
    views_m      = round(max(0.1, (engagement / 100) * np.random.lognormal(1.5, 1.0)), 2)
    watchtime_h  = round(views_m * random.uniform(1.2, 2.8), 2)
    user_rating  = round(min(10.0, max(1.0, (engagement / 100) * 10 + np.random.normal(0, 0.8))), 1)
    
    rows.append({
        "content_id":      content_id,
        "title":           f"Title_{content_id}",
        "type":            content_type,
        "genre":           genre,
        "sub_genre":       random.choice(GENRES),
        "country":         country,
        "language":        language,
        "rating":          rating,
        "release_year":    release_year,
        "date_added":      added_date.strftime("%Y-%m-%d"),
        "director":        random.choice(DIRECTORS),
        "cast_count":      random.randint(3, 20),
        "duration_min":    duration_min,
        "seasons":         seasons,
        "episodes":        episodes,
        "engagement_score":round(engagement, 1),
        "views_millions":  views_m,
        "watchtime_hours_millions": watchtime_h,
        "user_rating":     user_rating,
        "is_original":     random.choices([1, 0], weights=[0.38, 0.62])[0],
        "budget_category": random.choices(["Low","Mid","High","Premium"], weights=[0.25,0.35,0.28,0.12])[0],
        "added_year":      added_date.year,
        "added_month":     added_date.month,
        "added_quarter":   f"Q{(added_date.month - 1)//3 + 1}",
    })
    content_id += 1

df = pd.DataFrame(rows)
df.to_csv("/home/claude/netflix_project/netflix_content_data.csv", index=False)
print(f"Dataset: {len(df)} rows x {len(df.columns)} columns")
print(df.dtypes)
print("\nGenre distribution:")
print(df.groupby("genre")["content_id"].count().sort_values(ascending=False))
print("\nEngagement by genre (avg):")
print(df.groupby("genre")["engagement_score"].mean().sort_values(ascending=False).round(1))
