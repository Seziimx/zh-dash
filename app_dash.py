# app_dash.py
import dash
from dash import dcc, html, dash_table, Input, Output, State
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
from dash import ctx
import base64
from datetime import datetime

# ----------------------
# Load & prepare data
# ----------------------
def load_data(path="data/zhubanov_scopus_issn.xlsx", sheet="ARTICLE"):
    df = pd.read_excel(path, sheet_name=sheet, engine="openpyxl")

    rename_map = {
        "Автор (ы)": "authors_raw",
        "Author full names": "authors_full",
        "Название документа": "title",
        "Год": "year",
        "Название источника": "source",
        "Цитирования": "cited_by",
        "DOI": "doi",
        "Ссылка": "url",
        "ISSN": "issn",
        "Квартиль": "quartile",
        "Процентиль 2024": "percentile_2024",
    }
    df = df.rename(columns=rename_map)

    # Ensure types
    if "year" in df.columns:
        df["year"] = pd.to_numeric(df["year"], errors="coerce").astype("Int64")
    else:
        df["year"] = pd.NA
    if "cited_by" in df.columns:
        df["cited_by"] = pd.to_numeric(df["cited_by"], errors="coerce").fillna(0).astype(int)
    else:
        df["cited_by"] = 0
    if "percentile_2024" in df.columns:
        df["percentile_2024"] = pd.to_numeric(df["percentile_2024"], errors="coerce")
    else:
        df["percentile_2024"] = pd.NA

    # DOI link
    df["doi_link"] = df.get("doi").apply(lambda x: f"https://doi.org/{str(x).strip()}" if pd.notna(x) and str(x).strip() else None)

    # authors formatted (from authors_raw) — replace ";" with newline for display
    df["authors_fmt"] = df.get("authors_raw", "").astype(str).str.replace(";", "\n")

    # Lowercase helpers for search
    df["_title_lc"] = df["title"].astype(str).str.lower()
    df["_source_lc"] = df["source"].astype(str).str.lower()
    df["_authors_raw_lc"] = df["authors_raw"].astype(str).str.lower()

    return df

df = load_data()

# Precompute options with counts for sources and authors
source_counts = df["source"].fillna("—").value_counts()
source_options = [{"label": f"{s} ({int(source_counts[s])})", "value": s} for s in source_counts.index]

# authors list derived from authors_raw (split by ;)
authors_series = df["authors_raw"].dropna().astype(str).str.split(";").explode().str.strip()
author_counts = authors_series.value_counts()
author_options = [{"label": f"{a} ({int(author_counts[a])})", "value": a} for a in author_counts.index]

# ----------------------
# Dash App
# ----------------------
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "Zh Scopus — Жубанов"

# Helper: default year range
years_nonnull = df["year"].dropna().astype(int)
if len(years_nonnull) > 0:
    min_year, max_year = int(years_nonnull.min()), int(years_nonnull.max())
else:
    min_year, max_year = 2000, datetime.now().year

# Layout
app.layout = html.Div([
    html.Div([
        html.H2("Zh Scopus — портал публикаций", style={"marginBottom": "6px"}),
        html.Div("Университет Жубанова • Аналитика Scopus", style={"color": "#666", "marginBottom": "12px"}),
        # Sidebar filters
        html.H4("Фильтры"),
        dcc.Input(id="search", type="text", placeholder="Поиск (автор/название/источник)", debounce=True,
                  style={"width": "100%", "marginBottom": "8px"}),

        html.Label("Интервал (быстрый)"),
        dcc.RadioItems(id="preset_years",
                       options=[{"label": "Все годы", "value": "all"},
                                {"label": "Последние 5 лет", "value": "last5"},
                                {"label": "Последние 10 лет", "value": "last10"}],
                       value="all", inline=False, style={"marginBottom": "8px"}),

        html.Label("Диапазон лет"),
        dcc.RangeSlider(id="year_range", min=min_year, max=max_year, value=[min_year, max_year],
                        marks={y: str(y) for y in range(max_year, min_year-1, -5)}, step=1,
                        tooltip={"placement": "bottom", "always_visible": False}),
        html.Br(),

        html.Label("Фильтр по источникам"),
        dcc.Dropdown(id="source_filter", options=source_options, multi=True, placeholder="Выберите источники",
                     style={"marginBottom": "8px"}),

        html.Label("Фильтр по авторам"),
        dcc.Dropdown(id="author_filter", options=author_options, multi=True, placeholder="Выберите авторов",
                     style={"marginBottom": "8px"}),

        html.Label("Квартиль"),
        dcc.Checklist(id="quartile_filter",
                      options=[{"label": q, "value": q} for q in ["Q1", "Q2", "Q3", "Q4"]],
                      value=["Q1", "Q2", "Q3", "Q4"],
                      style={"marginBottom": "8px"}),

        html.Label("Процентиль 2024"),
        dcc.RangeSlider(id="percentile_range", min=0, max=100, value=[0, 100], step=1,
                        marks={0: "0", 50: "50", 100: "100"}, style={"marginBottom": "12px"}),

        html.Label("Сортировка"),
        dcc.Dropdown(id="sort_by",
                     options=[
                         {"label": "Год (новые → старые)", "value": "year_desc"},
                         {"label": "Год (старые → новые)", "value": "year_asc"},
                         {"label": "Цитирования (много → мало)", "value": "cited_desc"},
                         {"label": "Цитирования (мало → много)", "value": "cited_asc"},
                         {"label": "Процентиль (высокий → низкий)", "value": "pct_desc"},
                         {"label": "Автор (A–Z)", "value": "author_az"},
                         {"label": "Автор (Z–A)", "value": "author_za"},
                         {"label": "Источник (A–Z)", "value": "source_az"},
                         {"label": "Источник (Z–A)", "value": "source_za"},
                         {"label": "Название (A–Z)", "value": "title_az"},
                     ],
                     value="year_desc", clearable=False, style={"marginBottom": "12px"}),

        html.Button("Применить / Обновить", id="apply_btn", n_clicks=0, style={"width": "100%"}),
        html.Br(), html.Br(),

        html.Div([
            html.Button("Экспорт CSV", id="export_csv", n_clicks=0, style={"marginRight": "8px"}),
            html.Button("Экспорт Excel", id="export_xlsx", n_clicks=0)
        ], style={"textAlign": "center"}),
        dcc.Download(id="download-dataframe"),

        html.Div(style={"marginTop": "18px", "fontSize": "12px", "color": "#888"}, children=[
            "© Zh Scopus / Университет Жубанова"
        ])
    ], style={"width": "320px", "padding": "16px", "float": "left", "boxSizing": "border-box",
              "borderRight": "1px solid #eee", "height": "100vh", "overflowY": "auto"}),

    # Main content
    html.Div([
        dcc.Tabs(id="main_tabs", value="tab_table", children=[
            dcc.Tab(label="Таблица", value="tab_table"),
            dcc.Tab(label="Scopus-вид", value="tab_cards"),
            dcc.Tab(label="Топ источников", value="tab_sources"),
            dcc.Tab(label="Топ авторов", value="tab_authors"),
        ]),
        html.Div(id="tab_content", style={"padding": "16px"})
    ], style={"marginLeft": "340px", "padding": "16px"})
])

# ----------------------
# Helper: filtering logic
# ----------------------
def apply_filters(df_in,
                  search: str,
                  year_preset: str,
                  year_range: list,
                  sources: list,
                  authors: list,
                  quartiles: list,
                  percentile_range: list,
                  sort_by: str):
    df_f = df_in.copy()

    # Year preset
    max_year = df_in["year"].max() if not df_in["year"].isna().all() else year_range[1]
    if year_preset == "last5":
        df_f = df_f[df_f["year"].fillna(0) >= (int(max_year) - 4)]
    elif year_preset == "last10":
        df_f = df_f[df_f["year"].fillna(0) >= (int(max_year) - 9)]
    else:
        # range slider applied
        if year_range:
            df_f = df_f[df_f["year"].between(int(year_range[0]), int(year_range[1]))]

    # Quartile
    if quartiles:
        df_f = df_f[df_f["quartile"].astype(str).isin(quartiles)]

    # Percentile
    if percentile_range:
        p = df_f["percentile_2024"].fillna(-1)
        df_f = df_f[(p >= percentile_range[0]) & (p <= percentile_range[1])]

    # Sources
    if sources and len(sources) > 0:
        df_f = df_f[df_f["source"].isin(sources)]

    # Authors (authors list are values like "Surname, I.")
    if authors and len(authors) > 0:
        mask_auth = df_f["authors_raw"].astype(str).apply(
            lambda x: any(a.strip() in x for a in authors)
        )
        df_f = df_f[mask_auth]

    # Search (title | authors_raw | source)
    if search and str(search).strip() != "":
        q = str(search).lower()
        mask = (
            df_f["_title_lc"].str.contains(q, na=False) |
            df_f["_authors_raw_lc"].str.contains(q, na=False) |
            df_f["_source_lc"].str.contains(q, na=False)
        )
        df_f = df_f[mask]

    # Sorting
    if sort_by == "year_desc":
        df_f = df_f.sort_values(by="year", ascending=False, na_position="last")
    elif sort_by == "year_asc":
        df_f = df_f.sort_values(by="year", ascending=True, na_position="last")
    elif sort_by == "cited_desc":
        df_f = df_f.sort_values(by="cited_by", ascending=False)
    elif sort_by == "cited_asc":
        df_f = df_f.sort_values(by="cited_by", ascending=True)
    elif sort_by == "pct_desc":
        df_f = df_f.sort_values(by="percentile_2024", ascending=False, na_position="last")
    elif sort_by == "author_az":
        df_f = df_f.sort_values(by="authors_raw", ascending=True)
    elif sort_by == "author_za":
        df_f = df_f.sort_values(by="authors_raw", ascending=False)
    elif sort_by == "source_az":
        df_f = df_f.sort_values(by="source", ascending=True)
    elif sort_by == "source_za":
        df_f = df_f.sort_values(by="source", ascending=False)
    elif sort_by == "title_az":
        df_f = df_f.sort_values(by="title", ascending=True)
    return df_f

# ----------------------
# Callbacks: render tabs content & export
# ----------------------
@app.callback(
    Output("tab_content", "children"),
    Input("apply_btn", "n_clicks"),
    State("search", "value"),
    State("preset_years", "value"),
    State("year_range", "value"),
    State("source_filter", "value"),
    State("author_filter", "value"),
    State("quartile_filter", "value"),
    State("percentile_range", "value"),
    State("sort_by", "value"),
    State("main_tabs", "value"),
)
def render_tabs(n_clicks, search, preset_years, year_range, sources, authors,
                quartiles, percentile_range, sort_by, active_tab):
    # Apply filters
    filtered = apply_filters(df, search, preset_years, year_range, sources, authors,
                             quartiles, percentile_range, sort_by)

    # Always reset numbering from 1
    filtered_display = filtered.reset_index(drop=True).copy()
    filtered_display.insert(0, "№", range(1, len(filtered_display) + 1))

    # Ensure authors_fmt uses authors_raw and newline
    filtered_display["authors_fmt"] = filtered_display.get("authors_raw", "").astype(str).str.replace(";", "\n")

    # Common columns for table
    table_columns = [
        {"name": "№", "id": "№"},
        {"name": "Авторы", "id": "authors_fmt"},
        {"name": "Название", "id": "title"},
        {"name": "Год", "id": "year"},
        {"name": "Источник", "id": "source"},
        {"name": "Квартиль", "id": "quartile"},
        {"name": "Процентиль 2024", "id": "percentile_2024"},
        {"name": "Цитирования", "id": "cited_by"},
        {"name": "DOI", "id": "doi_link"},
        {"name": "Scopus ссылка", "id": "url"},
    ]
    # Keep only existing columns
    table_columns = [c for c in table_columns if c["id"] in filtered_display.columns]

    # Tab: Таблица
    if active_tab == "tab_table":
        table = dash_table.DataTable(
            id="results_table",
            columns=table_columns,
            data=filtered_display.to_dict("records"),
            page_size=20,
            style_cell={
                "whiteSpace": "pre-line",  # <newline> отображается
                "textAlign": "left",
                "padding": "6px",
                "fontFamily": "Arial, sans-serif",
                "fontSize": "13px"
            },
            style_header={"backgroundColor": "#0D1B2A", "color": "white", "fontWeight": "bold"},
            style_data_conditional=[
                {"if": {"row_index": "odd"}, "backgroundColor": "#f9fbfd"},
                {"if": {"row_index": "even"}, "backgroundColor": "white"},
                {"if": {"state": "active"}, "backgroundColor": "#112240", "color": "white"}
            ],
            style_table={"overflowX": "auto"},
        )
        return html.Div([
            html.H3(f"Результаты — {len(filtered_display)}"),
            table
        ])

    # Tab: Scopus-вид (карточки)
    if active_tab == "tab_cards":
        cards = []
        for idx, row in filtered_display.iterrows():
            num = int(row["№"])
            title = row.get("title", "Без названия")
            authors = row.get("authors_fmt", "—")
            source = row.get("source", "—")
            year = row.get("year", "—")
            quart = row.get("quartile", "—")
            pct = row.get("percentile_2024", "—")
            cites = row.get("cited_by", 0)
            doi_link = row.get("doi_link", None)
            url = row.get("url", None)

            link_block = []
            if pd.notna(url) and url:
                link_block.append(html.A("Scopus", href=url, target="_blank", style={"marginRight": "8px"}))
            if pd.notna(doi_link) and doi_link:
                link_block.append(html.A("DOI", href=doi_link, target="_blank"))

            card = html.Div([
                html.H4(f"{num}. {title}", style={"marginBottom": "6px"}),
                html.Pre(f"Авторы:\n{authors}", style={"whiteSpace": "pre-wrap", "marginBottom": "6px"}),
                html.Div([
                    html.Span(f"Источник: {source}", style={"marginRight": "12px"}),
                    html.Span(f"Год: {year}", style={"marginRight": "12px"}),
                    html.Span(f"Квартиль: {quart}", style={"marginRight": "12px"}),
                    html.Span(f"Процентиль: {pct}", style={"marginRight": "12px"}),
                ], style={"marginBottom": "6px"}),
                html.Div(f"Цитирования: {cites}", style={"marginBottom": "6px"}),
                html.Div(link_block),
                html.Hr()
            ], style={"padding": "8px 4px"})
            cards.append(card)
        return html.Div(cards)

    # Tab: Топ источников
    if active_tab == "tab_sources":
        top_sources = (filtered.groupby("source")
                       .agg(pub_count=("title", "count"), cites=("cited_by", "sum"))
                       .sort_values("pub_count", ascending=False)
                       .reset_index())
        fig = px.bar(top_sources.head(20), x="pub_count", y="source", orientation="h",
                     labels={"pub_count": "Публикаций", "source": "Источник"},
                     height=600)
        table_src = dash_table.DataTable(
            columns=[{"name": "Источник", "id": "source"}, {"name": "Публикаций", "id": "pub_count"}, {"name": "Цитирования", "id": "cites"}],
            data=top_sources.head(20).to_dict("records"),
            style_table={"overflowX": "auto"},
            style_cell={"textAlign": "left"}
        )
        return html.Div([
            html.H3("Топ источников"),
            dcc.Graph(figure=fig),
            html.H4("Таблица"),
            table_src
        ])

    # Tab: Топ авторов
    if active_tab == "tab_authors":
        exploded = (filtered.assign(_authors=filtered["authors_raw"].astype(str).str.split(";"))
                    .explode("_authors"))
        exploded["_authors"] = exploded["_authors"].str.strip()
        top_authors = (exploded[exploded["_authors"] != ""]
                       .groupby("_authors").agg(pub_count=("title", "count"), cites=("cited_by", "sum"))
                       .sort_values("pub_count", ascending=False)
                       .reset_index().rename(columns={"_authors": "author"}))
        fig2 = px.bar(top_authors.head(20), x="pub_count", y="author", orientation="h",
                      labels={"pub_count": "Публикаций", "author": "Автор"},
                      height=600)
        table_auth = dash_table.DataTable(
            columns=[{"name": "Автор", "id": "author"}, {"name": "Публикаций", "id": "pub_count"}, {"name": "Цитирования", "id": "cites"}],
            data=top_authors.head(20).to_dict("records"),
            style_table={"overflowX": "auto"},
            style_cell={"textAlign": "left"}
        )
        return html.Div([
            html.H3("Топ авторов"),
            dcc.Graph(figure=fig2),
            html.H4("Таблица"),
            table_auth
        ])

    return html.Div("No tab")

# ----------------------
# Export callback
# ----------------------
@app.callback(
    Output("download-dataframe", "data"),
    Input("export_csv", "n_clicks"),
    Input("export_xlsx", "n_clicks"),
    State("search", "value"),
    State("preset_years", "value"),
    State("year_range", "value"),
    State("source_filter", "value"),
    State("author_filter", "value"),
    State("quartile_filter", "value"),
    State("percentile_range", "value"),
    State("sort_by", "value"),
    prevent_initial_call=True
)
def export_data(n_csv, n_xlsx, search, preset_years, year_range, sources, authors, quartiles, percentile_range, sort_by):
    trigger_id = ctx.triggered_id
    filtered = apply_filters(df, search, preset_years, year_range, sources, authors, quartiles, percentile_range, sort_by)
    # Build export frame: authors from authors_raw (replace ; with newlines)
    export_df = filtered.reset_index(drop=True).copy()
    export_df.insert(0, "№", range(1, len(export_df) + 1))
    export_df["Авторы"] = export_df.get("authors_raw", "").astype(str).str.replace(";", "\n")
    # Choose columns for export
    cols = ["№", "Авторы", "title", "year", "source", "quartile", "percentile_2024", "cited_by", "doi_link", "url", "issn"]
    cols = [c for c in cols if c in export_df.columns or c == "Авторы"]
    export_df = export_df[[c for c in cols if c in export_df.columns or c == "Авторы"]]

    if trigger_id == "export_csv":
        # CSV
        csv_bytes = export_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")
        return dcc.send_bytes(lambda x: x.write(csv_bytes), filename=f"zh_scopus_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv")
    elif trigger_id == "export_xlsx":
        # Excel
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            export_df.to_excel(writer, index=False, sheet_name="Export")
        buffer.seek(0)
        return dcc.send_bytes(lambda f: f.write(buffer.getvalue()), filename=f"zh_scopus_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
    return None

# ----------------------
if __name__ == "__main__":
    app.run_server(debug=True)