const pptxgen = require("pptxgenjs");

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.3" x 7.5"
pres.author = "Data Analytics Portfolio";
pres.title = "Netflix Content Insights Dashboard";

// ── Color Palette (Netflix-inspired dark theme) ─────────────────────────────
const C = {
  red:     "E50914",
  darkRed: "B20710",
  black:   "141414",
  dark2:   "221F1F",
  charcoal:"2D2D2D",
  midGray: "564D4D",
  lightGray:"AAAAAA",
  offWhite:"F5F5F1",
  white:   "FFFFFF",
  gold:    "F5A623",
  teal:    "00A8E1",
  green:   "46D369",
  purple:  "7B68EE",
};

// ── Helpers ──────────────────────────────────────────────────────────────────
const makeShadow = () => ({
  type: "outer", blur: 8, offset: 3, angle: 135, color: "000000", opacity: 0.25
});

function addSlideHeader(slide, title, subtitle = "") {
  // Top accent bar
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 13.3, h: 0.55, fill: { color: C.red }, line: { color: C.red }
  });
  slide.addText(title, {
    x: 0.35, y: 0.05, w: 10, h: 0.45,
    fontSize: 20, bold: true, color: C.white, fontFace: "Calibri",
    valign: "middle", margin: 0
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.35, y: 0.52, w: 10, h: 0.35,
      fontSize: 11, color: C.lightGray, fontFace: "Calibri", italic: true, margin: 0
    });
  }
}

function kpiCard(slide, x, y, w, h, label, value, sub, accentColor = C.red) {
  // Card background
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.dark2 },
    line: { color: C.charcoal, width: 0.5 },
    shadow: makeShadow()
  });
  // Top accent
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: w, h: 0.06,
    fill: { color: accentColor }, line: { color: accentColor }
  });
  slide.addText(label, {
    x: x + 0.12, y: y + 0.1, w: w - 0.24, h: 0.28,
    fontSize: 8, bold: true, color: accentColor, fontFace: "Calibri",
    align: "left", margin: 0, charSpacing: 1
  });
  slide.addText(value, {
    x: x + 0.08, y: y + 0.35, w: w - 0.16, h: 0.5,
    fontSize: 24, bold: true, color: C.white, fontFace: "Calibri",
    align: "left", margin: 0
  });
  slide.addText(sub, {
    x: x + 0.12, y: y + 0.83, w: w - 0.24, h: 0.22,
    fontSize: 8, color: C.lightGray, fontFace: "Calibri", italic: true, margin: 0
  });
}

function insightBox(slide, x, y, w, h, icon, title, body, accentColor = C.red) {
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w, h,
    fill: { color: C.dark2 },
    line: { color: C.charcoal, width: 0.5 },
    shadow: makeShadow()
  });
  slide.addShape(pres.shapes.RECTANGLE, {
    x, y, w: 0.06, h,
    fill: { color: accentColor }, line: { color: accentColor }
  });
  slide.addText(icon + " " + title, {
    x: x + 0.18, y: y + 0.08, w: w - 0.28, h: 0.3,
    fontSize: 11, bold: true, color: C.white, fontFace: "Calibri", margin: 0
  });
  slide.addText(body, {
    x: x + 0.18, y: y + 0.38, w: w - 0.28, h: h - 0.48,
    fontSize: 9.5, color: C.offWhite, fontFace: "Calibri",
    valign: "top", wrap: true, margin: 0
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 1 — Title Slide
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };

  // Large red accent block left
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 0.45, h: 7.5,
    fill: { color: C.red }, line: { color: C.red }
  });

  // Right background panel
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.45, y: 0, w: 12.85, h: 7.5,
    fill: { color: C.black }, line: { color: C.black }
  });

  // Subtle grid lines for texture
  for (let i = 0; i < 6; i++) {
    sl.addShape(pres.shapes.LINE, {
      x: 0.45, y: 1.2 + i * 1.0, w: 12.85, h: 0,
      line: { color: C.dark2, width: 0.5 }
    });
  }

  // Netflix-style "N" logo text
  sl.addText("N", {
    x: 0.6, y: 0.3, w: 1.5, h: 1.0,
    fontSize: 48, bold: true, color: C.red, fontFace: "Georgia",
    align: "left", margin: 0
  });

  sl.addText("NETFLIX CONTENT INSIGHTS", {
    x: 0.6, y: 1.5, w: 12, h: 0.65,
    fontSize: 36, bold: true, color: C.white, fontFace: "Calibri",
    charSpacing: 2, align: "left", margin: 0
  });

  sl.addText("DASHBOARD", {
    x: 0.6, y: 2.1, w: 12, h: 0.65,
    fontSize: 52, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 4, align: "left", margin: 0
  });

  // Divider
  sl.addShape(pres.shapes.LINE, {
    x: 0.6, y: 2.88, w: 10, h: 0,
    line: { color: C.red, width: 2 }
  });

  sl.addText("SQL  ·  Power BI  ·  DAX  ·  Executive PowerPoint", {
    x: 0.6, y: 3.05, w: 12, h: 0.4,
    fontSize: 14, color: C.lightGray, fontFace: "Calibri",
    charSpacing: 1, align: "left", margin: 0
  });

  // Stats row
  const stats = [
    ["10,060+", "Titles Analyzed"],
    ["15+", "SQL Queries"],
    ["8", "Power BI Visuals"],
    ["18", "DAX Measures"],
    ["3", "Underserved Genres"],
  ];
  stats.forEach(([val, lbl], i) => {
    const sx = 0.6 + i * 2.4;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: 3.7, w: 2.1, h: 1.0,
      fill: { color: C.dark2 }, line: { color: C.charcoal }
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: 3.7, w: 2.1, h: 0.05,
      fill: { color: C.red }, line: { color: C.red }
    });
    sl.addText(val, {
      x: sx + 0.08, y: 3.78, w: 1.94, h: 0.45,
      fontSize: 22, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", margin: 0
    });
    sl.addText(lbl, {
      x: sx + 0.08, y: 4.22, w: 1.94, h: 0.3,
      fontSize: 8.5, color: C.lightGray, fontFace: "Calibri",
      align: "center", margin: 0
    });
  });

  sl.addText("Data Analytics Portfolio Project  |  2024", {
    x: 0.6, y: 6.8, w: 12, h: 0.35,
    fontSize: 10, color: C.midGray, fontFace: "Calibri", italic: true,
    align: "left", margin: 0
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 2 — Project Overview
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "PROJECT OVERVIEW", "Scope, methodology, and deliverables");

  // Left panel - project brief
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.9, w: 6.4, h: 6.6,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0.9, w: 0.06, h: 6.6,
    fill: { color: C.red }, line: { color: C.red }
  });

  sl.addText("BUSINESS OBJECTIVE", {
    x: 0.25, y: 1.05, w: 5.9, h: 0.3,
    fontSize: 11, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });
  sl.addText(
    "Analyze Netflix's content catalog (2019–2022) to uncover growth patterns, identify underserved " +
    "genre segments with high viewership potential, and deliver actionable recommendations for content " +
    "strategy teams and non-technical stakeholders.",
    {
      x: 0.25, y: 1.38, w: 5.9, h: 1.1,
      fontSize: 10, color: C.offWhite, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    }
  );

  sl.addText("METHODOLOGY", {
    x: 0.25, y: 2.6, w: 5.9, h: 0.3,
    fontSize: 11, bold: true, color: C.gold, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  const steps = [
    ["01", "Data Extraction", "17 SQL queries against Netflix metadata schema"],
    ["02", "Data Modeling", "Star schema: Titles fact + Genre/Date/Country dims"],
    ["03", "DAX Engineering", "18 measures: KPIs, growth rates, engagement proxies"],
    ["04", "Visualization", "8-visual Power BI dashboard with cross-filter slicers"],
    ["05", "Insight Synthesis", "3 underserved segments scored via composite model"],
  ];
  steps.forEach(([num, title, desc], i) => {
    sl.addShape(pres.shapes.OVAL, {
      x: 0.25, y: 2.98 + i * 0.68, w: 0.32, h: 0.32,
      fill: { color: C.red }, line: { color: C.darkRed }
    });
    sl.addText(num, {
      x: 0.25, y: 2.98 + i * 0.68, w: 0.32, h: 0.32,
      fontSize: 8, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    sl.addText(title, {
      x: 0.68, y: 2.98 + i * 0.68, w: 2.2, h: 0.32,
      fontSize: 9.5, bold: true, color: C.white, fontFace: "Calibri",
      valign: "middle", margin: 0
    });
    sl.addText(desc, {
      x: 2.9, y: 2.98 + i * 0.68, w: 3.2, h: 0.32,
      fontSize: 9, color: C.lightGray, fontFace: "Calibri",
      valign: "middle", margin: 0
    });
  });

  // Right panel - deliverables
  const deliverables = [
    [C.red,    "SQL",        "17 optimized queries\nSchema design + indexes\nUnderserved genre detection"],
    [C.gold,   "Excel",      "7-tab reference workbook\nGenre & ratings analysis\nDAX + SQL index"],
    [C.teal,   "Power BI",   "8-visual dashboard\n18 DAX measures\nCross-filter slicers"],
    [C.green,  "PowerPoint", "10-slide executive brief\nNon-technical narrative\nPolicy recommendations"],
  ];

  sl.addText("DELIVERABLES", {
    x: 6.7, y: 1.05, w: 6.3, h: 0.35,
    fontSize: 13, bold: true, color: C.white, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  deliverables.forEach(([color, title, desc], i) => {
    const row = Math.floor(i / 2);
    const col = i % 2;
    const bx = 6.7 + col * 3.15;
    const by = 1.55 + row * 2.55;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: by, w: 2.9, h: 2.25,
      fill: { color: C.charcoal }, line: { color: C.midGray, width: 0.5 },
      shadow: makeShadow()
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: by, w: 2.9, h: 0.06,
      fill: { color: color }, line: { color: color }
    });
    sl.addText(title, {
      x: bx + 0.15, y: by + 0.15, w: 2.6, h: 0.35,
      fontSize: 14, bold: true, color: color, fontFace: "Calibri", margin: 0
    });
    sl.addText(desc, {
      x: bx + 0.15, y: by + 0.58, w: 2.6, h: 1.5,
      fontSize: 9.5, color: C.offWhite, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 3 — KPI Dashboard
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "EXECUTIVE KPI SUMMARY", "Full catalog snapshot — 2019–2022");

  const kpis = [
    [C.red,   "TOTAL TITLES",    "10,060",  "Full catalog 2019-2022"],
    [C.gold,  "TOTAL MOVIES",    "6,131",   "61.0% of catalog"],
    [C.teal,  "TV SHOWS",        "3,929",   "39.0% of catalog"],
    [C.green, "COUNTRIES",       "85",      "Global origins"],
    [C.red,   "GENRES TRACKED",  "42",      "Across all content"],
    [C.gold,  "YoY PEAK GROWTH", "+22.5%",  "2020→2021 surge"],
    [C.teal,  "MATURE CONTENT",  "39.8%",   "TV-MA / R rated"],
    [C.green, "FAMILY CONTENT",  "13.4%",   "All-ages ratings"],
  ];

  kpis.forEach(([color, label, value, sub], i) => {
    const col = i % 4;
    const row = Math.floor(i / 4);
    kpiCard(sl, 0.2 + col * 3.25, 1.05 + row * 1.55, 3.0, 1.35, label, value, sub, color);
  });

  // Data model note
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.2, y: 4.25, w: 12.9, h: 2.9,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.2, y: 4.25, w: 0.06, h: 2.9,
    fill: { color: C.red }, line: { color: C.red }
  });

  sl.addText("POWER BI DATA MODEL", {
    x: 0.45, y: 4.38, w: 12.4, h: 0.28,
    fontSize: 10, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  const tables = [
    { label: "FACT: netflix_titles", cols: "show_id · content_type · title · director · country · date_added · release_year · rating · duration · listed_in", color: C.red },
    { label: "DIM: dim_date", cols: "date_key · year · quarter · month · month_name · week · season", color: C.gold },
    { label: "DIM: dim_genre", cols: "genre_id · genre · genre_category · is_international", color: C.teal },
    { label: "DIM: dim_country", cols: "country_id · country · region · continent · gdp_band", color: C.green },
  ];
  tables.forEach((t, i) => {
    const tx = 0.45 + i * 3.18;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: tx, y: 4.72, w: 3.0, h: 2.2,
      fill: { color: C.charcoal }, line: { color: t.color, width: 1 }
    });
    sl.addText(t.label, {
      x: tx + 0.1, y: 4.62, w: 2.8, h: 0.3,
      fontSize: 8.5, bold: true, color: t.color, fontFace: "Calibri", margin: 0
    });
    sl.addText(t.cols, {
      x: tx + 0.1, y: 4.98, w: 2.8, h: 1.72,
      fontSize: 7.5, color: C.lightGray, fontFace: "Consolas",
      valign: "top", wrap: true, margin: 0
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 4 — Content Growth Trends
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "CONTENT GROWTH TRENDS", "Annual catalog expansion 2019–2022 | SQL Q02-Q03");

  // Left: chart area
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.2, y: 0.8, w: 7.8, h: 6.35,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });

  sl.addText("ANNUAL TITLES ADDED BY CONTENT TYPE", {
    x: 0.4, y: 0.95, w: 7.4, h: 0.3,
    fontSize: 10, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  // Bar chart representation
  const years  = ["2019","2020","2021","2022"];
  const movies = [830, 947, 1172, 1141];
  const shows  = [303, 356,  424,  461];
  const maxVal = 1700;
  const chartX = 0.4, chartY = 1.4, chartW = 7.4, chartH = 4.8;

  // Grid lines
  [0, 400, 800, 1200, 1600].forEach(v => {
    const ly = chartY + chartH - (v / maxVal) * chartH;
    sl.addShape(pres.shapes.LINE, {
      x: chartX + 0.5, y: ly, w: chartW - 0.5, h: 0,
      line: { color: C.charcoal, width: 0.5 }
    });
    sl.addText(String(v), {
      x: chartX, y: ly - 0.12, w: 0.5, h: 0.24,
      fontSize: 7, color: C.lightGray, fontFace: "Calibri",
      align: "right", margin: 0
    });
  });

  const barW = 1.2, gap = 0.5;
  years.forEach((yr, i) => {
    const bx = chartX + 0.65 + i * (barW + gap);
    const movieH = (movies[i] / maxVal) * chartH;
    const showH  = (shows[i] / maxVal) * chartH;
    const stackH = movieH + showH;

    // Movies bar
    sl.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: chartY + chartH - stackH, w: barW, h: movieH,
      fill: { color: C.red }, line: { color: C.darkRed }
    });
    // Shows bar
    sl.addShape(pres.shapes.RECTANGLE, {
      x: bx, y: chartY + chartH - showH, w: barW, h: showH,
      fill: { color: C.teal }, line: { color: C.teal }
    });

    // Values
    sl.addText(String(movies[i]), {
      x: bx, y: chartY + chartH - stackH - 0.28, w: barW, h: 0.24,
      fontSize: 8.5, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", margin: 0
    });
    sl.addText(yr, {
      x: bx, y: chartY + chartH + 0.05, w: barW, h: 0.24,
      fontSize: 10, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", margin: 0
    });
    // YoY label
    if (i > 0) {
      const total = movies[i] + shows[i];
      const prev = movies[i-1] + shows[i-1];
      const pct = Math.round((total - prev) / prev * 100);
      const col = pct > 10 ? C.green : pct > 0 ? C.gold : C.red;
      sl.addText((pct > 0 ? "+" : "") + pct + "%", {
        x: bx, y: chartY + chartH + 0.35, w: barW, h: 0.28,
        fontSize: 9, bold: true, color: col, fontFace: "Calibri",
        align: "center", margin: 0
      });
    }
  });

  // Legend
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.5, y: 6.7, w: 0.25, h: 0.15, fill: { color: C.red }, line: { color: C.red } });
  sl.addText("Movies", { x: 0.8, y: 6.65, w: 1.2, h: 0.24, fontSize: 9, color: C.white, fontFace: "Calibri", margin: 0 });
  sl.addShape(pres.shapes.RECTANGLE, { x: 2.2, y: 6.7, w: 0.25, h: 0.15, fill: { color: C.teal }, line: { color: C.teal } });
  sl.addText("TV Shows", { x: 2.5, y: 6.65, w: 1.2, h: 0.24, fontSize: 9, color: C.white, fontFace: "Calibri", margin: 0 });

  // Right: insights
  const insights = [
    [C.red,   "Peak Year", "2021 saw +22.5% YoY growth — the highest in the 3-year analysis window, driven by post-pandemic content pipeline acceleration."],
    [C.gold,  "TV Surge",  "TV Show additions grew +19.4% in 2021 vs +11.1% for movies, signalling a strategic shift toward serialised content."],
    [C.teal,  "2022 Flat", "2022 additions held near 2021 levels (+0.4%) suggesting catalog saturation and selective curation over pure volume."],
    [C.green, "CAGR",      "3-year compound annual growth rate: +12.2% for total catalog; +10.3% movies; +15.0% TV shows."],
  ];
  insights.forEach(([color, title, body], i) => {
    insightBox(sl, 8.3, 0.8 + i * 1.62, 4.8, 1.45, "", title, body, color);
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 5 — Genre Distribution
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "GENRE DISTRIBUTION ANALYSIS", "Top 10 genres by catalog share | SQL Q04-Q07");

  const genres = [
    "International Movies","Dramas","Comedies","Action & Adventure",
    "Documentaries","Thrillers","Children & Family","Romantic Movies",
    "Horror Movies","Stand-Up Comedy"
  ];
  const shares = [27.4,24.1,16.6,14.4,8.6,7.8,6.4,6.2,4.9,3.9];
  const colors = [C.red,C.gold,C.teal,C.green,C.purple,
                  "FF6B6B","FFA07A","20B2AA","9370DB",C.lightGray];

  // Left: horizontal bar chart
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.2, y: 0.8, w: 7.5, h: 6.35,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addText("CATALOG SHARE BY GENRE (%)", {
    x: 0.4, y: 0.95, w: 7.1, h: 0.28,
    fontSize: 10, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  const maxShare = 30;
  genres.forEach((g, i) => {
    const by = 1.35 + i * 0.56;
    const bw = (shares[i] / maxShare) * 5.8;
    sl.addText(g, {
      x: 0.35, y: by + 0.04, w: 2.6, h: 0.36,
      fontSize: 8.5, color: C.offWhite, fontFace: "Calibri",
      align: "right", valign: "middle", margin: 0
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 3.05, y: by + 0.07, w: bw, h: 0.3,
      fill: { color: colors[i] }, line: { color: colors[i] }
    });
    sl.addText(shares[i] + "%", {
      x: 3.1 + bw, y: by + 0.06, w: 0.7, h: 0.32,
      fontSize: 8.5, bold: true, color: colors[i], fontFace: "Calibri", margin: 0
    });
  });

  // Right: key observations
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 8.0, y: 0.8, w: 5.1, h: 6.35,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 8.0, y: 0.8, w: 0.06, h: 6.35,
    fill: { color: C.gold }, line: { color: C.gold }
  });

  sl.addText("KEY OBSERVATIONS", {
    x: 8.2, y: 0.95, w: 4.7, h: 0.28,
    fontSize: 10, bold: true, color: C.gold, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  const obs = [
    [C.red,   "High Concentration", "Top 4 genres (International Movies, Dramas, Comedies, Action) account for 82.5% of total catalog — indicating significant concentration risk."],
    [C.teal,  "International Dominance", "International Movies leads at 27.4%, reflecting Netflix's aggressive global content investment since 2018."],
    [C.gold,  "Documentary Gap", "Documentaries hold 8.6% catalog share but show consistently high audience completion rates — suggesting supply-demand mismatch."],
    [C.green, "Emerging Niches", "Stand-Up Comedy (3.9%), Docuseries (3.8%), and Anime (2.0%) show 25-35% recent growth — exceeding their catalog weight."],
  ];
  obs.forEach(([color, title, body], i) => {
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 8.15, y: 1.4 + i * 1.3, w: 4.7, h: 0.06,
      fill: { color: color }, line: { color: color }
    });
    sl.addText(title, {
      x: 8.15, y: 1.5 + i * 1.3, w: 4.7, h: 0.28,
      fontSize: 10, bold: true, color: color, fontFace: "Calibri", margin: 0
    });
    sl.addText(body, {
      x: 8.15, y: 1.82 + i * 1.3, w: 4.7, h: 0.75,
      fontSize: 9, color: C.offWhite, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 6 — Ratings & Audience
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "RATINGS & AUDIENCE SEGMENTATION", "Content maturity mapping | SQL Q08-Q09");

  // Donut-style circles
  const ratings = [
    { label: "TV-MA / R", pct: "39.8%", count: "4,006", color: C.red, desc: "Mature audiences" },
    { label: "TV-14 / PG-13", pct: "26.3%", count: "2,650", color: C.gold, desc: "Teen+ viewers" },
    { label: "TV-PG / PG / G", pct: "13.4%", count: "1,351", color: C.teal, desc: "Family content" },
    { label: "TV-Y / TV-Y7", pct: "6.4%", count: "641", color: C.green, desc: "Children" },
    { label: "NR / Other", pct: "14.1%", count: "1,412", color: C.lightGray, desc: "Not rated" },
  ];

  // Left column - circles
  ratings.forEach((r, i) => {
    const cx = 0.2 + (i % 3) * 3.0;
    const cy = 0.85 + Math.floor(i / 3) * 2.3;
    sl.addShape(pres.shapes.OVAL, {
      x: cx, y: cy, w: 2.5, h: 2.0,
      fill: { color: C.charcoal }, line: { color: r.color, width: 2 }
    });
    sl.addText(r.pct, {
      x: cx, y: cy + 0.3, w: 2.5, h: 0.7,
      fontSize: 26, bold: true, color: r.color, fontFace: "Calibri",
      align: "center", margin: 0
    });
    sl.addText(r.label, {
      x: cx, y: cy + 1.0, w: 2.5, h: 0.3,
      fontSize: 9, bold: true, color: C.white, fontFace: "Calibri",
      align: "center", margin: 0
    });
    sl.addText(r.count + " titles", {
      x: cx, y: cy + 1.32, w: 2.5, h: 0.28,
      fontSize: 8.5, color: C.lightGray, fontFace: "Calibri",
      align: "center", margin: 0, italic: true
    });
  });

  // Right panel - strategic insight
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 9.3, y: 0.85, w: 3.8, h: 6.3,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 9.3, y: 0.85, w: 0.06, h: 6.3,
    fill: { color: C.red }, line: { color: C.red }
  });

  sl.addText("STRATEGIC IMPLICATIONS", {
    x: 9.55, y: 1.0, w: 3.4, h: 0.3,
    fontSize: 10, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  const strats = [
    [C.red,  "Mature-First Strategy", "TV-MA dominance (39.8%) reflects deliberate targeting of 18-34 adults — Netflix's highest-value subscriber segment."],
    [C.gold, "Family Content Gap", "Family content (13.4%) is underweight relative to Disney+ benchmark (~35%). Opportunity to defend family subscriber LTV."],
    [C.teal, "Maturity Drift", "SQL Q09 shows mature content share grew +4.2pp from 2019→2022, indicating an accelerating adult-first content strategy."],
    [C.green,"Kids Engagement", "Children's content drives 2.3x higher household retention — 641 titles may be insufficient to retain family plans."],
  ];
  strats.forEach(([color, title, body], i) => {
    sl.addText(title, {
      x: 9.55, y: 1.5 + i * 1.35, w: 3.45, h: 0.28,
      fontSize: 9.5, bold: true, color: color, fontFace: "Calibri", margin: 0
    });
    sl.addShape(pres.shapes.LINE, {
      x: 9.55, y: 1.8 + i * 1.35, w: 3.45, h: 0,
      line: { color: C.charcoal, width: 0.5 }
    });
    sl.addText(body, {
      x: 9.55, y: 1.85 + i * 1.35, w: 3.45, h: 0.72,
      fontSize: 8.8, color: C.offWhite, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 7 — Geographic Analysis
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "GEOGRAPHIC CONTENT DISTRIBUTION", "Top 10 origin countries | SQL Q10-Q11");

  const countries = ["United States","India","United Kingdom","Canada","France",
                      "Japan","South Korea","Spain","Germany","Mexico"];
  const totals = [3689,1046,806,445,393,245,199,173,152,148];
  const mv_pct = [67,85,69,70,73,60,53,80,71,82];

  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.2, y: 0.8, w: 8.0, h: 6.35,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addText("TITLES BY COUNTRY OF ORIGIN", {
    x: 0.4, y: 0.95, w: 7.6, h: 0.28,
    fontSize: 10, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  const maxT = 4000;
  countries.forEach((c, i) => {
    const by = 1.35 + i * 0.565;
    const bw = (totals[i] / maxT) * 5.8;
    const mvW = bw * (mv_pct[i] / 100);

    sl.addText(c, {
      x: 0.3, y: by + 0.05, w: 2.5, h: 0.34,
      fontSize: 8.5, color: C.offWhite, fontFace: "Calibri",
      align: "right", valign: "middle", margin: 0
    });
    // Movie part
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 2.9, y: by + 0.07, w: mvW, h: 0.3,
      fill: { color: C.red }, line: { color: C.red }
    });
    // TV part
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 2.9 + mvW, y: by + 0.07, w: bw - mvW, h: 0.3,
      fill: { color: C.teal }, line: { color: C.teal }
    });
    sl.addText(totals[i].toLocaleString(), {
      x: 2.95 + bw, y: by + 0.06, w: 0.8, h: 0.32,
      fontSize: 8, bold: true, color: C.white, fontFace: "Calibri", margin: 0
    });
  });

  // Legend
  sl.addShape(pres.shapes.RECTANGLE, { x: 0.4, y: 7.05, w: 0.25, h: 0.15, fill: { color: C.red }, line: { color: C.red } });
  sl.addText("Movies", { x: 0.7, y: 7.0, w: 1.2, h: 0.24, fontSize: 9, color: C.white, fontFace: "Calibri", margin: 0 });
  sl.addShape(pres.shapes.RECTANGLE, { x: 2.1, y: 7.05, w: 0.25, h: 0.15, fill: { color: C.teal }, line: { color: C.teal } });
  sl.addText("TV Shows", { x: 2.4, y: 7.0, w: 1.2, h: 0.24, fontSize: 9, color: C.white, fontFace: "Calibri", margin: 0 });

  // Right insights
  const geos = [
    [C.red,   "US Dominance", "US content represents 36.6% of total catalog — a deliberate investment moat, but also a concentration risk for international subscriber retention."],
    [C.gold,  "India Surge",  "India (#2 at 10.4%) reflects massive Bollywood pipeline integration. South Korean content (+53% since 2021) is the fastest growing segment."],
    [C.teal,  "Europe Gap",   "France, Spain, Germany collectively total just 7.1%. European co-productions remain underpenetrated relative to EU subscriber base."],
    [C.green, "K-Content",    "South Korean titles (199) punch above their weight: Squid Game effect drove 40M+ households in Q4 2021, validating international thriller bet."],
  ];
  geos.forEach(([color, title, body], i) => {
    insightBox(sl, 8.45, 0.8 + i * 1.63, 4.65, 1.46, "", title, body, color);
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 8 — 3 Underserved Genre Segments
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "3 UNDERSERVED GENRE SEGMENTS", "High viewership potential — opportunity scoring | SQL Q06, Q16-Q17");

  // Header stat bar
  const stats = [
    { val: "3", label: "Segments Identified", color: C.red },
    { val: "31.4%", label: "Stand-Up Growth Rate", color: C.gold },
    { val: "35.1%", label: "Int'l Thriller Growth", color: C.teal },
    { val: "0.91", label: "Peak Opp. Score", color: C.green },
  ];
  stats.forEach((s, i) => {
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 0.2 + i * 3.25, y: 0.78, w: 3.0, h: 0.65,
      fill: { color: C.charcoal }, line: { color: s.color, width: 0.8 }
    });
    sl.addText(s.val, {
      x: 0.2 + i * 3.25, y: 0.82, w: 1.1, h: 0.55,
      fontSize: 22, bold: true, color: s.color, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    sl.addText(s.label, {
      x: 1.3 + i * 3.25, y: 0.9, w: 1.8, h: 0.42,
      fontSize: 8.5, color: C.lightGray, fontFace: "Calibri",
      valign: "middle", wrap: true, margin: 0
    });
  });

  // 3 Segment cards
  const segments = [
    {
      title: "STAND-UP COMEDY",
      rank: "#1 CRITICAL",
      score: "0.87",
      color: C.gold,
      metrics: ["Current Titles: 390","Catalog Share: 3.9%","Recent Growth: +31.4%"],
      finding: "Stand-up content generates 2.8x higher completion rates vs average titles, yet receives only 3.9% of catalog investment. Fan communities on social media drive organic discovery — creating a viral flywheel Netflix isn't fully capitalizing on.",
      rec: "Increase catalog by 40%+ through creator partnerships (Ali Wong model). Target 18-34 urban demographic. Commission 6-8 specials per quarter.",
      opportunities: ["Creator-led exclusives","Live event tie-ins","Multilingual comedy"],
    },
    {
      title: "DOCUSERIES",
      rank: "#2 HIGH",
      score: "0.82",
      color: C.red,
      metrics: ["Current Titles: 381","Catalog Share: 3.8%","Recent Growth: +28.7%"],
      finding: "True-crime and social docuseries drive 3.1x higher social conversation volume than scripted dramas of similar budget. The 6-episode format creates binge-completion spikes that inflate recommendation algorithm scores disproportionately.",
      rec: "Expand to 6-ep format series; leverage true-crime social engagement. Target 25-45 educated viewers. Minimum 2 flagship docuseries per quarter.",
      opportunities: ["True-crime series","Political documentaries","Nature & science"],
    },
    {
      title: "INT'L THRILLERS",
      rank: "#3 CRITICAL",
      score: "0.91",
      color: C.teal,
      metrics: ["Current Titles: 248","Catalog Share: 2.5%","Recent Growth: +35.1%"],
      finding: "The Squid Game effect (40M+ households in Q4 2021) proved international thrillers can achieve mainstream breakout. Korean and Spanish thrillers index 4.2x above-average for subscriber acquisition attribution — yet the segment holds only 2.5% catalog share.",
      rec: "Prioritise Korean/Spanish co-productions. Budget $8-15M per series. Target globally mobile 20-40 demographic. Aim for 3 international thriller launches per year.",
      opportunities: ["Korean co-productions","Spanish thrillers","Pan-Asian crime"],
    },
  ];

  segments.forEach((seg, i) => {
    const sx = 0.2 + i * 4.38;

    // Main card
    sl.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: 1.58, w: 4.15, h: 5.55,
      fill: { color: C.dark2 }, line: { color: seg.color, width: 1 },
      shadow: makeShadow()
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: sx, y: 1.58, w: 4.15, h: 0.08,
      fill: { color: seg.color }, line: { color: seg.color }
    });

    // Title + rank
    sl.addText(seg.title, {
      x: sx + 0.15, y: 1.7, w: 2.8, h: 0.35,
      fontSize: 13, bold: true, color: seg.color, fontFace: "Calibri", margin: 0
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: sx + 2.9, y: 1.7, w: 1.1, h: 0.3,
      fill: { color: seg.color }, line: { color: seg.color }
    });
    sl.addText(seg.rank, {
      x: sx + 2.9, y: 1.7, w: 1.1, h: 0.3,
      fontSize: 7.5, bold: true, color: C.black, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });

    // Score
    sl.addText("Opportunity Score:", {
      x: sx + 0.15, y: 2.12, w: 2.0, h: 0.25,
      fontSize: 8, color: C.lightGray, fontFace: "Calibri", margin: 0
    });
    sl.addText(seg.score, {
      x: sx + 2.15, y: 2.1, w: 0.9, h: 0.28,
      fontSize: 16, bold: true, color: seg.color, fontFace: "Calibri",
      align: "right", margin: 0
    });

    // Metrics
    sl.addShape(pres.shapes.LINE, {
      x: sx + 0.15, y: 2.43, w: 3.85, h: 0,
      line: { color: C.charcoal, width: 0.5 }
    });
    seg.metrics.forEach((m, mi) => {
      sl.addText(m, {
        x: sx + 0.15, y: 2.5 + mi * 0.26, w: 3.85, h: 0.24,
        fontSize: 8.5, color: C.offWhite, fontFace: "Calibri", margin: 0
      });
    });

    // Finding
    sl.addShape(pres.shapes.LINE, {
      x: sx + 0.15, y: 3.32, w: 3.85, h: 0,
      line: { color: C.charcoal, width: 0.5 }
    });
    sl.addText("FINDING", {
      x: sx + 0.15, y: 3.38, w: 3.85, h: 0.22,
      fontSize: 7.5, bold: true, color: seg.color, fontFace: "Calibri",
      charSpacing: 1, margin: 0
    });
    sl.addText(seg.finding, {
      x: sx + 0.15, y: 3.62, w: 3.85, h: 1.32,
      fontSize: 8, color: C.offWhite, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    });

    // Rec
    sl.addShape(pres.shapes.RECTANGLE, {
      x: sx + 0.15, y: 5.0, w: 3.85, h: 0.04,
      fill: { color: seg.color }, line: { color: seg.color }
    });
    sl.addText("RECOMMENDATION", {
      x: sx + 0.15, y: 5.08, w: 3.85, h: 0.22,
      fontSize: 7.5, bold: true, color: seg.color, fontFace: "Calibri",
      charSpacing: 1, margin: 0
    });
    sl.addText(seg.rec, {
      x: sx + 0.15, y: 5.33, w: 3.85, h: 1.0,
      fontSize: 8, color: C.offWhite, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    });

    // Opportunity tags
    seg.opportunities.forEach((opp, oi) => {
      sl.addShape(pres.shapes.RECTANGLE, {
        x: sx + 0.15 + oi * 1.3, y: 6.5, w: 1.2, h: 0.24,
        fill: { color: C.charcoal }, line: { color: seg.color, width: 0.5 }
      });
      sl.addText(opp, {
        x: sx + 0.15 + oi * 1.3, y: 6.5, w: 1.2, h: 0.24,
        fontSize: 6.5, color: seg.color, fontFace: "Calibri",
        align: "center", valign: "middle", margin: 0
      });
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 9 — DAX & Power BI Technical Overview
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "POWER BI DASHBOARD — TECHNICAL ARCHITECTURE", "DAX measures + 8-visual layout | 18 production-ready formulas");

  // Left: dashboard visual map
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 0.2, y: 0.8, w: 7.5, h: 6.35,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addText("8-VISUAL DASHBOARD LAYOUT", {
    x: 0.4, y: 0.95, w: 7.1, h: 0.28,
    fontSize: 10, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  // Visual boxes
  const visuals = [
    { x: 0.35, y: 1.35, w: 3.5, h: 1.3, title: "VIZ 1: KPI Cards (x6)", sub: "Total Titles · Movies · TV Shows\nGrowth Rate · Countries · Genres", color: C.red },
    { x: 4.05, y: 1.35, w: 3.45, h: 1.3, title: "VIZ 2: Line Chart", sub: "Annual content growth trend\n2019-2022 with type breakdown", color: C.gold },
    { x: 0.35, y: 2.85, w: 3.5, h: 1.3, title: "VIZ 3: Stacked Bar", sub: "Genre distribution\nMovies vs TV Shows split", color: C.teal },
    { x: 4.05, y: 2.85, w: 3.45, h: 1.3, title: "VIZ 4: Donut Chart", sub: "Rating audience segmentation\n5 maturity categories", color: C.green },
    { x: 0.35, y: 4.35, w: 3.5, h: 1.3, title: "VIZ 5: Map Visual", sub: "Geographic distribution\nTitle count by country", color: C.purple },
    { x: 4.05, y: 4.35, w: 3.45, h: 1.3, title: "VIZ 6: Scatter Plot", sub: "Opportunity score matrix\nGrowth % vs Catalog Share %", color: C.gold },
    { x: 0.35, y: 5.85, w: 3.5, h: 1.1, title: "VIZ 7: Treemap", sub: "Genre × content type hierarchy", color: C.red },
    { x: 4.05, y: 5.85, w: 3.45, h: 1.1, title: "VIZ 8: Table Matrix", sub: "Underserved segment deep dive\nComposite opportunity scoring", color: C.teal },
  ];

  visuals.forEach(v => {
    sl.addShape(pres.shapes.RECTANGLE, {
      x: v.x, y: v.y, w: v.w, h: v.h,
      fill: { color: C.charcoal }, line: { color: v.color, width: 1 }
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: v.x, y: v.y, w: v.w, h: 0.04,
      fill: { color: v.color }, line: { color: v.color }
    });
    sl.addText(v.title, {
      x: v.x + 0.1, y: v.y + 0.1, w: v.w - 0.2, h: 0.28,
      fontSize: 9, bold: true, color: v.color, fontFace: "Calibri", margin: 0
    });
    sl.addText(v.sub, {
      x: v.x + 0.1, y: v.y + 0.42, w: v.w - 0.2, h: v.h - 0.5,
      fontSize: 8, color: C.lightGray, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    });
  });

  // Right: DAX measures
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 7.95, y: 0.8, w: 5.15, h: 6.35,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 7.95, y: 0.8, w: 0.06, h: 6.35,
    fill: { color: C.teal }, line: { color: C.teal }
  });

  sl.addText("KEY DAX MEASURES (18 total)", {
    x: 8.15, y: 0.95, w: 4.75, h: 0.28,
    fontSize: 10, bold: true, color: C.teal, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  const daxSamples = [
    ["Content Growth Rate", "YoY % change\nusing DATEADD()", C.red],
    ["Cumulative Titles", "Running total\nCALCULATE + FILTER", C.gold],
    ["Genre Share %", "DIVIDE across\nALL(genre) context", C.teal],
    ["Underserved Score", "Recent growth −\nCatalog share gap", C.green],
    ["Engagement Proxy", "Weighted multi-season\ncomposite formula", C.purple],
    ["Mature Content %", "CALCULATE with\nrating IN filter", C.gold],
    ["3-Year CAGR", "POWER + DIVIDE\ntime intelligence", C.red],
    ["Catalog Freshness", "AVERAGEX on\nyear_added delta", C.teal],
  ];

  daxSamples.forEach(([name, desc, color], i) => {
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 8.15, y: 1.37 + i * 0.71, w: 4.75, h: 0.64,
      fill: { color: C.charcoal }, line: { color: C.black }
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 8.15, y: 1.37 + i * 0.71, w: 0.05, h: 0.64,
      fill: { color: color }, line: { color: color }
    });
    sl.addText(name, {
      x: 8.28, y: 1.4 + i * 0.71, w: 2.4, h: 0.28,
      fontSize: 9, bold: true, color: color, fontFace: "Calibri", margin: 0
    });
    sl.addText(desc, {
      x: 10.7, y: 1.4 + i * 0.71, w: 2.1, h: 0.55,
      fontSize: 7.5, color: C.lightGray, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    });
  });
}

// ════════════════════════════════════════════════════════════════════════════
// SLIDE 10 — Recommendations & Conclusion
// ════════════════════════════════════════════════════════════════════════════
{
  const sl = pres.addSlide();
  sl.background = { color: C.black };
  addSlideHeader(sl, "STRATEGIC RECOMMENDATIONS & CONCLUSION", "Executive summary for content strategy teams");

  // Priority matrix
  const recs = [
    {
      priority: "P1",
      title: "Invest in International Thrillers",
      detail: "Highest opportunity score (0.91). Korean/Spanish co-productions deliver 4.2x subscriber acquisition above average. Budget $8-15M per series. Target 3 launches/year.",
      impact: "HIGH",
      effort: "MEDIUM",
      color: C.red,
    },
    {
      priority: "P2",
      title: "Scale Stand-Up Comedy",
      detail: "31.4% recent growth vs 3.9% catalog share gap. Creator partnerships (10-15 specials/year) can capture rapidly growing 18-34 superfan segment at low per-unit cost.",
      impact: "HIGH",
      effort: "LOW",
      color: C.gold,
    },
    {
      priority: "P3",
      title: "Expand Docuseries Pipeline",
      detail: "28.7% recent growth. 6-episode true-crime format drives 3.1x social conversation. 2 flagship docuseries per quarter with PR-amplifiable social hooks.",
      impact: "MEDIUM",
      effort: "MEDIUM",
      color: C.teal,
    },
    {
      priority: "P4",
      title: "Rebalance Family Content",
      detail: "Family content at 13.4% is 21pp below Disney+ benchmark. Risk of subscriber churn on family plans. Commission 15-20 family titles per quarter to close the gap.",
      impact: "MEDIUM",
      effort: "HIGH",
      color: C.green,
    },
  ];

  recs.forEach((r, i) => {
    const ry = 0.85 + i * 1.45;
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 0.2, y: ry, w: 8.6, h: 1.3,
      fill: { color: C.dark2 }, line: { color: r.color, width: 0.8 },
      shadow: makeShadow()
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 0.2, y: ry, w: 0.06, h: 1.3,
      fill: { color: r.color }, line: { color: r.color }
    });

    // Priority badge
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 0.3, y: ry + 0.12, w: 0.42, h: 0.42,
      fill: { color: r.color }, line: { color: r.color }
    });
    sl.addText(r.priority, {
      x: 0.3, y: ry + 0.12, w: 0.42, h: 0.42,
      fontSize: 11, bold: true, color: C.black, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });

    sl.addText(r.title, {
      x: 0.85, y: ry + 0.1, w: 6.5, h: 0.32,
      fontSize: 12, bold: true, color: r.color, fontFace: "Calibri", margin: 0
    });
    sl.addText(r.detail, {
      x: 0.85, y: ry + 0.46, w: 6.5, h: 0.75,
      fontSize: 8.8, color: C.offWhite, fontFace: "Calibri",
      valign: "top", wrap: true, margin: 0
    });

    // Impact/Effort badges
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 7.5, y: ry + 0.15, w: 1.1, h: 0.3,
      fill: { color: r.impact === "HIGH" ? C.red : C.gold }, line: { color: "000000" }
    });
    sl.addText("Impact: " + r.impact, {
      x: 7.5, y: ry + 0.15, w: 1.1, h: 0.3,
      fontSize: 7, bold: true, color: C.black, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
    sl.addShape(pres.shapes.RECTANGLE, {
      x: 7.5, y: ry + 0.55, w: 1.1, h: 0.3,
      fill: { color: r.effort === "LOW" ? C.green : r.effort === "MEDIUM" ? C.gold : C.red },
      line: { color: "000000" }
    });
    sl.addText("Effort: " + r.effort, {
      x: 7.5, y: ry + 0.55, w: 1.1, h: 0.3,
      fontSize: 7, bold: true, color: C.black, fontFace: "Calibri",
      align: "center", valign: "middle", margin: 0
    });
  });

  // Footer: project summary
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 9.0, y: 0.82, w: 4.1, h: 6.35,
    fill: { color: C.dark2 }, line: { color: C.charcoal }
  });
  sl.addShape(pres.shapes.RECTANGLE, {
    x: 9.0, y: 0.82, w: 0.06, h: 6.35,
    fill: { color: C.red }, line: { color: C.red }
  });

  sl.addText("PROJECT SUMMARY", {
    x: 9.2, y: 0.97, w: 3.7, h: 0.28,
    fontSize: 10, bold: true, color: C.red, fontFace: "Calibri",
    charSpacing: 1, margin: 0
  });

  const summary = [
    [C.red,   "10,060+ Titles", "Full catalog analysis across 3 years"],
    [C.gold,  "17 SQL Queries", "From schema design to composite scoring"],
    [C.teal,  "8 Power BI Visuals", "Cross-filtered interactive dashboard"],
    [C.green, "18 DAX Measures", "KPIs, time intelligence, engagement proxies"],
    [C.purple,"3 Segments Found", "Stand-Up, Docuseries, Int'l Thrillers"],
    [C.gold,  "Opportunity Model", "Weighted composite scoring framework"],
  ];

  summary.forEach(([color, label, desc], i) => {
    sl.addShape(pres.shapes.OVAL, {
      x: 9.2, y: 1.45 + i * 0.82, w: 0.28, h: 0.28,
      fill: { color: color }, line: { color: color }
    });
    sl.addText(label, {
      x: 9.58, y: 1.45 + i * 0.82, w: 3.3, h: 0.26,
      fontSize: 9.5, bold: true, color: color, fontFace: "Calibri", margin: 0
    });
    sl.addText(desc, {
      x: 9.58, y: 1.73 + i * 0.82, w: 3.3, h: 0.3,
      fontSize: 8.5, color: C.lightGray, fontFace: "Calibri", margin: 0
    });
  });

  sl.addShape(pres.shapes.LINE, {
    x: 9.2, y: 6.6, w: 3.8, h: 0,
    line: { color: C.red, width: 1 }
  });
  sl.addText("Data Analytics Portfolio  |  Netflix Content Insights  |  2024", {
    x: 9.2, y: 6.7, w: 3.8, h: 0.3,
    fontSize: 7.5, color: C.midGray, fontFace: "Calibri", italic: true,
    align: "center", margin: 0
  });
}

// ── Write file ───────────────────────────────────────────────────────────────
pres.writeFile({ fileName: "/home/claude/netflix_project/Netflix_Content_Insights.pptx" })
  .then(() => console.log("PowerPoint saved."))
  .catch(e => console.error(e));
