# app_dash.py 
import dash
from dash import dcc, html, dash_table, Input, Output, State, ctx
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
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

    # authors formatted
    df["authors_fmt"] = df.get("authors_raw", "").astype(str).str.replace(";", "\n")

    # lowercase helpers
    df["_title_lc"] = df["title"].astype(str).str.lower()
    df["_source_lc"] = df["source"].astype(str).str.lower()
    df["_authors_raw_lc"] = df["authors_raw"].astype(str).str.lower()

    return df


df = load_data()

# Precompute options
source_counts = df["source"].fillna("—").value_counts()
source_options = [{"label": f"{s} ({int(source_counts[s])})", "value": s} for s in source_counts.index]

authors_series = df["authors_raw"].dropna().astype(str).str.split(";").explode().str.strip()
author_counts = authors_series.value_counts()
author_options = [{"label": f"{a} ({int(author_counts[a])})", "value": a} for a in author_counts.index]

# ----------------------
# Dash App
# ----------------------
app = dash.Dash(__name__, suppress_callback_exceptions=True)
app.title = "Zh Scopus — Жубанов"
server = app.server

# Year range
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
                        marks={0: "0", 50: "50", 100: "100"}),

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
def apply_filters(df_in, search, year_preset, year_range, sources, authors, quartiles, percentile_range, sort_by):
    df_f = df_in.copy()
    max_year = df_in["year"].max() if not df_in["year"].isna().all() else year_range[1]

    # Year filters
    if year_preset == "last5":
        df_f = df_f[df_f["year"].fillna(0) >= (int(max_year) - 4)]
    elif year_preset == "last10":
        df_f = df_f[df_f["year"].fillna(0) >= (int(max_year) - 9)]
    elif year_range:
        df_f = df_f[df_f["year"].between(int(year_range[0]), int(year_range[1]))]

    if quartiles:
        df_f = df_f[df_f["quartile"].astype(str).isin(quartiles)]

    if percentile_range:
        p = df_f["percentile_2024"].fillna(-1)
        df_f = df_f[(p >= percentile_range[0]) & (p <= percentile_range[1])]

    if sources:
        df_f = df_f[df_f["source"].isin(sources)]

    if authors:
        mask_auth = df_f["authors_raw"].astype(str).apply(lambda x: any(a.strip() in x for a in authors))
        df_f = df_f[mask_auth]

    if search and str(search).strip():
        q = str(search).lower()
        mask = (
            df_f["_title_lc"].str.contains(q, na=False) |
            df_f["_authors_raw_lc"].str.contains(q, na=False) |
            df_f["_source_lc"].str.contains(q, na=False)
        )
        df_f = df_f[mask]

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
# Callbacks
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
    filtered = apply_filters(df, search, preset_years, year_range, sources, authors,
                             quartiles, percentile_range, sort_by)
    filtered_display = filtered.reset_index(drop=True).copy()
    filtered_display.insert(0, "№", range(1, len(filtered_display) + 1))
    filtered_display["authors_fmt"] = filtered_display.get("authors_raw", "").astype(str).str.replace(";", "\n")

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
    table_columns = [c for c in table_columns if c["id"] in filtered_display.columns]

    # Таблица
    if active_tab == "tab_table":
        return dash_table.DataTable(
            columns=table_columns,
            data=filtered_display.to_dict("records"),
            page_size=20,
            style_cell={"whiteSpace": "pre-line", "textAlign": "left"},
            style_header={"backgroundColor": "#0D1B2A", "color": "white", "fontWeight": "bold"},
            style_table={"overflowX": "auto"},
        )

    # Scopus-вид
    if active_tab == "tab_cards":
        cards = []
        for _, row in filtered_display.iterrows():
            cards.append(html.Div([
                html.H4(f"{row['№']}. {row['title']}"),
                html.Pre(f"Авторы:\n{row['authors_fmt']}"),
                html.Div(f"Источник: {row['source']} | Год: {row['year']} | Квартиль: {row['quartile']} | Процентиль: {row['percentile_2024']}"),
                html.Div(f"Цитирования: {row['cited_by']}"),
                html.A("DOI", href=row["doi_link"], target="_blank") if row.get("doi_link") else None,
                html.Br(),
                html.A("Scopus", href=row["url"], target="_blank") if row.get("url") else None,
                html.Hr()
            ], style={"marginBottom": "12px"}))
        return html.Div(cards)

    # Топ источников
    if active_tab == "tab_sources":
        top_sources = (filtered.groupby("source")
                       .agg(pub_count=("title", "count"), cites=("cited_by", "sum"))
                       .sort_values("pub_count", ascending=False).reset_index())
        fig = px.bar(top_sources.head(20), x="pub_count", y="source", orientation="h",
                     labels={"pub_count": "Публикаций", "source": "Источник"})
        return dcc.Graph(figure=fig)

    # Топ авторов
    if active_tab == "tab_authors":
        exploded = (filtered.assign(_authors=filtered["authors_raw"].astype(str).str.split(";"))
                    .explode("_authors"))
        exploded["_authors"] = exploded["_authors"].str.strip()
        top_authors = (exploded[exploded["_authors"] != ""]
                       .groupby("_authors").agg(pub_count=("title", "count"), cites=("cited_by", "sum"))
                       .sort_values("pub_count", ascending=False).reset_index())
        fig2 = px.bar(top_authors.head(20), x="pub_count", y="_authors", orientation="h",
                      labels={"pub_count": "Публикаций", "_authors": "Автор"})
        return dcc.Graph(figure=fig2)

    return html.Div("Нет данных")

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
def export_data(n_csv, n_xlsx, search, preset_years, year_range, sources, authors,
                quartiles, percentile_range, sort_by):
    filtered = apply_filters(df, search, preset_years, year_range, sources, authors,
                             quartiles, percentile_range, sort_by)
    filtered = filtered.reset_index(drop=True)
    filtered.insert(0, "№", range(1, len(filtered) + 1))

    trigger_id = ctx.triggered_id
    if trigger_id == "export_csv":
        return dcc.send_data_frame(filtered.to_csv, f"Zh_Scopus_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", index=False)
    elif trigger_id == "export_xlsx":
        return dcc.send_data_frame(filtered.to_excel, f"Zh_Scopus_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", index=False, engine="openpyxl")

if __name__ == "__main__":
    app.run_server(debug=True)
