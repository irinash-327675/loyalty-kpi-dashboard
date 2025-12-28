import streamlit as st
import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple

import plotly.graph_objects as go
from plotly.subplots import make_subplots

KPI_FILE = "KPI_loyalty_metrics_full_ru.xlsx"

GREEN_ACCENT = "#2E8B57"          # основной акцент
NOTE_BG = "rgba(46, 139, 87, 0.08)"  # фон плашек
BAR_COLOR = "rgba(46, 139, 87, 0.18)" # цвет столбиков



# -------------------- THEME --------------------
st.set_page_config(page_title="Анализ программы лояльности", layout="wide")

st.markdown(
    f"""
<style>
:root {{ --primary-color: {GREEN_ACCENT}; }}

::selection {{ background: rgba(46, 139, 87, 0.22); }}

div[data-baseweb="select"] > div {{
  border-color: rgba(46, 139, 87, 0.28) !important;
}}
div[data-baseweb="select"] > div:hover {{
  border-color: rgba(46, 139, 87, 0.55) !important;
}}
input:focus, textarea:focus, [role="combobox"]:focus {{
  outline-color: rgba(46, 139, 87, 0.55) !important;
}}

/* KPI GRID */
.kpi-grid {{
  display: grid;
  grid-template-columns: repeat(3, minmax(0, 1fr));
  gap: 16px;
  align-items: stretch;
}}

@media (max-width: 1100px) {{
  .kpi-grid {{ grid-template-columns: repeat(2, minmax(0, 1fr)); }}
}}
@media (max-width: 700px) {{
  .kpi-grid {{ grid-template-columns: repeat(1, minmax(0, 1fr)); }}
}}

.kpi-tile {{
  background: #FFFFFF;
  border: 1px solid rgba(0,0,0,0.10);
  border-radius: 16px;
  padding: 18px 18px;
  box-shadow: 0 1px 10px rgba(0,0,0,0.04);
  height: 100%;
}}

.kpi-title {{
  font-size: 20px;
  font-weight: 850;
  line-height: 1.25;
  margin-bottom: 8px;
  color: #111111;
}}

.kpi-meta {{
  font-size: 12.5px;
  line-height: 1.3;
  color: #4B5563; /* нормальный серый, читается в любой теме */
  margin-bottom: 12px;
}}

.kpi-h {{
  font-size: 13.5px;
  font-weight: 850;
  letter-spacing: 0.2px;
  text-transform: none;
  margin: 14px 0 8px 0;
  color: #111111;
}}

.kpi-text {{
  font-size: 14.5px;
  line-height: 1.6;
  color: #111111;
}}

.kpi-code {{
  background: #F3F4F6;
  border: 1px solid rgba(0,0,0,0.06);
  border-radius: 12px;
  padding: 10px 12px;
  font-family: ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, "Liberation Mono", "Courier New", monospace;
  font-size: 12.5px;
  line-height: 1.45;
  color: #111111;
  white-space: pre-wrap;
}}

.kpi-tile ul {{
  margin: 8px 0 0 0;
  padding-left: 18px;
}}

.kpi-tile li {{
  margin: 6px 0;
  line-height: 1.55;
  color: #111111;
}}

.note {{
  background: transparent;
  border: none;
  border-left: 3px solid rgba(0,0,0,0.18);
  border-radius: 0;
  padding: 2px 0 2px 12px;
  margin: 12px 0 0 0;
}}
.note b {{
  color: rgba(0,0,0,0.88);
}}
.note ul {{
  margin: 8px 0 0 0;
  padding-left: 18px;
}}
.note li {{
  margin: 6px 0;
  line-height: 1.55;
  color: rgba(0,0,0,0.88);
}}

</style>
""",
    unsafe_allow_html=True,
)

st.title("Анализ программы лояльности")
st.caption("Пример отчета на данных за 6 месяцев и справочник KPI с рекомендациями")


# -------------------- HELPERS --------------------
def _clean_str(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return ""
    s = str(x).replace("\r\n", "\n").replace("\r", "\n").strip()
    if s.lower() == "nan":
        return ""
    return s


def _escape_html(s: str) -> str:
    return (
        s.replace("&", "&amp;")
        .replace("<", "&lt;")
        .replace(">", "&gt;")
    )


def recommendations_to_list(text: str) -> List[str]:
    raw = _clean_str(text)
    if not raw:
        return []
    lines = [ln.strip() for ln in raw.split("\n") if ln.strip()]
    if len(lines) == 1 and ";" in lines[0]:
        lines = [p.strip() for p in lines[0].split(";") if p.strip()]
    bullets: List[str] = []
    for ln in lines:
        cleaned = ln.lstrip("•▪*-— ").strip()
        if cleaned:
            bullets.append(cleaned)
    return bullets


def fmt_int(x) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "—"
    try:
        v = int(round(float(x)))
        return f"{v:,}".replace(",", " ")
    except Exception:
        return "—"


def fmt_money(x) -> str:
    return fmt_int(x)


def fmt_pct_ru(x, digits: int = 1) -> str:
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return "—"
    try:
        v = float(x)
        return f"{v:.{digits}f}".replace(".", ",") + "%"
    except Exception:
        return "—"


def safe_div(a, b):
    if b is None or (isinstance(b, float) and pd.isna(b)) or b == 0:
        return float("nan")
    return a / b


# ===================== TAB 1: METRICS =====================
@st.cache_data
def load_kpi_table(path: str) -> pd.DataFrame:
    df = pd.read_excel(path)
    expected = ["KPI", "Категория", "Описание", "Формула", "Рекомендации"]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"В KPI-файле не хватает колонок: {', '.join(missing)}")

    df = df.dropna(subset=["KPI"]).reset_index(drop=True)

    # чистим и убираем англ. хвосты в скобках у названий KPI
    df["KPI"] = df["KPI"].astype(str).apply(_clean_str).str.replace(r"\s*\([^)]*\)\s*$", "", regex=True)
    for c in ["Категория", "Описание", "Формула", "Рекомендации"]:
        df[c] = df[c].astype(str).apply(_clean_str)

    return df


# ===================== TAB 2: REPORT ======================
@st.cache_data
def build_demo_report() -> pd.DataFrame:
    """
    Демо-данные 6 месяцев (2025-07..2025-12) с “событием” в 2025-10:
    запуск игровой механики => рост активных и повторных.
    Все расчёты согласованы.
    """
    months = pd.date_range("2025-07-01", periods=6, freq="MS")
    mtxt = [d.strftime("%Y-%m") for d in months]

    total_sales = [520_000_000, 535_000_000, 550_000_000, 560_000_000, 590_000_000, 605_000_000]
    loyalty_sales = [138_000_000, 145_000_000, 152_000_000, 160_000_000, 178_000_000, 186_000_000]

    loyalty_tx = [430_000, 445_000, 460_000, 480_000, 525_000, 545_000]
    unique_members_tx = [285_000, 292_000, 298_000, 318_000, 350_000, 360_000]  # рост с октября

    loyalty_db_total = [1_180_000, 1_195_000, 1_210_000, 1_230_000, 1_250_000, 1_270_000]
    active_members = [275_000, 282_000, 288_000, 310_000, 345_000, 352_000]      # рост с октября

    # повторные покупатели (из уникальных покупавших)
    repeaters = [82_000, 85_000, 87_000, 98_000, 125_000, 130_000]

    avg_basket_non_member = [1_180, 1_190, 1_200, 1_205, 1_215, 1_220]

    df = pd.DataFrame({
        "Месяц": mtxt,
        "Общие продажи, ₽": total_sales,
        "Продажи участников ПЛ, ₽": loyalty_sales,
        "Транзакции участников ПЛ, шт": loyalty_tx,
        "Уникальные покупатели ПЛ, чел": unique_members_tx,
        "База ПЛ, всего, чел": loyalty_db_total,
        "Активные участники, чел": active_members,
        "Покупатели с ≥2 покупками, чел": repeaters,
        "Средний чек не участника, ₽": avg_basket_non_member,
    })

    # производные показатели (точно)
    df["Доля ПЛ в общих продажах, %"] = df["Продажи участников ПЛ, ₽"] / df["Общие продажи, ₽"] * 100.0
    df["Средний чек участника, ₽"] = df["Продажи участников ПЛ, ₽"] / df["Транзакции участников ПЛ, шт"]
    df["Частота покупок участника"] = df["Транзакции участников ПЛ, шт"] / df["Уникальные покупатели ПЛ, чел"]
    df["Доля повторных покупок, %"] = df["Покупатели с ≥2 покупками, чел"] / df["Уникальные покупатели ПЛ, чел"] * 100.0
    df["Доля активных в базе, %"] = df["Активные участники, чел"] / df["База ПЛ, всего, чел"] * 100.0

    # сравнение с не участниками
    df["Средний чек: разница с не участником, ₽"] = df["Средний чек участника, ₽"] - df["Средний чек не участника, ₽"]

    return df


def build_report_matrix(df: pd.DataFrame) -> pd.DataFrame:
    """
    Таблица “перевёрнута”: колонки — месяцы, строки — показатели.
    Числа отформатированы: пробелы в тысячах, проценты с запятой.
    """
    months = df["Месяц"].tolist()

    rows: List[Tuple[str, str, List, str]] = []

    # Продажи
    rows.append(("Продажи", "Продажи участников ПЛ", df["Продажи участников ПЛ, ₽"].tolist(), "money"))
    rows.append(("Продажи", "Доля ПЛ в общих продажах", df["Доля ПЛ в общих продажах, %"].tolist(), "pct"))
    rows.append(("Продажи", "Транзакции участников ПЛ", df["Транзакции участников ПЛ, шт"].tolist(), "int"))
    rows.append(("Продажи", "Уникальные покупатели ПЛ", df["Уникальные покупатели ПЛ, чел"].tolist(), "int"))
    rows.append(("Продажи", "Средний чек участника", df["Средний чек участника, ₽"].tolist(), "money"))
    rows.append(("Продажи", "Средний чек: разница с не участником", df["Средний чек: разница с не участником, ₽"].tolist(), "money"))
    rows.append(("Продажи", "Частота покупок участника", df["Частота покупок участника"].tolist(), "num2"))
    rows.append(("Продажи", "Доля повторных покупок", df["Доля повторных покупок, %"].tolist(), "pct"))

    # База
    rows.append(("База участников", "База ПЛ, всего", df["База ПЛ, всего, чел"].tolist(), "int"))
    rows.append(("База участников", "Активные участники", df["Активные участники, чел"].tolist(), "int"))
    rows.append(("База участников", "Доля активных в базе", df["Доля активных в базе, %"].tolist(), "pct"))

    idx = pd.MultiIndex.from_tuples([(b, n) for b, n, _, _ in rows], names=["Блок", "Показатель"])
    mat = pd.DataFrame(index=idx, columns=months)

    fmt_map: Dict[Tuple[str, str], str] = {}

    for b, n, values, fmt in rows:
        mat.loc[(b, n), :] = values
        fmt_map[(b, n)] = fmt

    out = mat.copy()
    for (b, n), fmt in fmt_map.items():
        vals = out.loc[(b, n), :].tolist()
        formatted = []
        for v in vals:
            if fmt == "money":
                formatted.append(fmt_money(v))
            elif fmt == "int":
                formatted.append(fmt_int(v))
            elif fmt == "pct":
                formatted.append(fmt_pct_ru(v, digits=1))
            elif fmt == "num2":
                if v is None or (isinstance(v, float) and pd.isna(v)):
                    formatted.append("—")
                else:
                    formatted.append(f"{float(v):.2f}".replace(".", ","))
            else:
                formatted.append(str(v))
        out.loc[(b, n), :] = formatted

    return out


# -------------------- UI --------------------
tabs = st.tabs(["Пример отчёта", "Метрики"])


# ===== TAB 1: METRICS =====
with tabs[1]:
    kpi_path = Path(KPI_FILE)
    if not kpi_path.exists():
        st.error(
            f"Не найден файл '{KPI_FILE}' рядом с app.py.\n\n"
            f"Положи '{KPI_FILE}' в папку проекта: {Path('.').resolve()}"
        )
    else:
        kpi_df = load_kpi_table(str(kpi_path))
        with st.container(border=True):
            c1, c2 = st.columns([1.2, 3.0])
            with c1:
                categories = ["Все категории"] + sorted([c for c in kpi_df["Категория"].unique().tolist() if c])
                selected_cat = st.selectbox("Категория", categories, index=0)
            with c2:
                query = st.text_input("Поиск", value="", placeholder="Например: выручка, активные, баллы, уровни…").strip().lower()

        
        df = kpi_df.copy()
        if selected_cat != "Все категории":
            df = df[df["Категория"] == selected_cat]

        if query:
            hay = (
                df["KPI"].astype(str).str.lower()
                + " " + df["Описание"].astype(str).str.lower()
                + " " + df["Формула"].astype(str).str.lower()
                + " " + df["Рекомендации"].astype(str).str.lower()
            )
            df = df[hay.str.contains(query, na=False)]

        df = df.reset_index(drop=True)
        if df.empty:
          st.warning("По выбранным условиям ничего не найдено.")
        else:
          show = df.copy()

    # короткие заголовки, чтобы меньше распирало по ширине
    show = show.rename(columns={
        "Категория": "Кат.",
        "Описание": "Зачем смотреть",
    })

    # порядок колонок
    show = show[["Кат.", "KPI", "Зачем смотреть", "Формула", "Рекомендации"]]

    st.dataframe(
        show,
        use_container_width=True,
        height=780,     # чтобы помещалось по вертикали и скроллилось внутри
        hide_index=True,
    )



# ===== TAB 2: REPORT =====
with tabs[0]:
    df = build_demo_report()
    last = df.iloc[-1]
    prev = df.iloc[-2]

    # KPI (как в первой версии: крупно, на верхней панели; без слов в скобках)
    st.subheader("Ключевые KPI")

    kpi_cols = st.columns(5)

    def delta_pct(curr, prv):
        if prv == 0 or pd.isna(prv) or pd.isna(curr):
            return "н/д"
        d = (curr / prv - 1.0) * 100.0
        sign = "+" if d >= 0 else ""
        return f"{sign}{d:.1f}%".replace(".", ",")

    # 1) Продажи ПЛ
    kpi_cols[0].metric(
        "Продажи участников ПЛ, ₽",
        fmt_money(last["Продажи участников ПЛ, ₽"]),
        delta=delta_pct(last["Продажи участников ПЛ, ₽"], prev["Продажи участников ПЛ, ₽"]),
    )
    # 2) Доля ПЛ в продажах
    kpi_cols[1].metric(
        "Доля ПЛ в общих продажах, %",
        fmt_pct_ru(last["Доля ПЛ в общих продажах, %"], 1),
        delta=delta_pct(last["Доля ПЛ в общих продажах, %"], prev["Доля ПЛ в общих продажах, %"]),
    )
    # 3) Активные участники
    kpi_cols[2].metric(
        "Активные участники, чел",
        fmt_int(last["Активные участники, чел"]),
        delta=delta_pct(last["Активные участники, чел"], prev["Активные участники, чел"]),
    )
    # 4) Доля активных
    kpi_cols[3].metric(
        "Доля активных в базе, %",
        fmt_pct_ru(last["Доля активных в базе, %"], 1),
        delta=delta_pct(last["Доля активных в базе, %"], prev["Доля активных в базе, %"]),
    )
    # 5) Доля повторных
    kpi_cols[4].metric(
        "Доля повторных покупок, %",
        fmt_pct_ru(last["Доля повторных покупок, %"], 1),
        delta=delta_pct(last["Доля повторных покупок, %"], prev["Доля повторных покупок, %"]),
    )

    st.markdown("---")


# ===================== CHART 1 ============================
    fig_sales = make_subplots(specs=[[{"secondary_y": True}]])

    fig_sales.add_trace(
        go.Bar(
            x=df["Месяц"],
            y=df["Продажи участников ПЛ, ₽"],
            name="Продажи участников ПЛ, ₽",
            marker=dict(
                color=BAR_COLOR,
                line=dict(color="rgba(46, 139, 87, 0.30)", width=1),
            ),
        ),
        secondary_y=False,
    )

    fig_sales.add_trace(
        go.Scatter(
            x=df["Месяц"],
            y=df["Доля ПЛ в общих продажах, %"],
            name="Доля ПЛ в общих продажах, %",
            mode="lines+markers",
            line=dict(color=GREEN_ACCENT, width=2),
        ),
        secondary_y=True,
    )

    fig_sales.update_layout(
        yaxis=dict(showgrid=True),
        template="plotly_white",
        height=420,
        margin=dict(l=20, r=20, t=40, b=20),
        title=dict(
            text=("Динамика продаж участников ПЛ"),
            x=0,
            xanchor="left",
            font=dict(size=18),
            pad=dict(b=12),
        ),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=0.96,
            xanchor="left",
            x=0,
            
        ),
    )

    fig_sales.update_yaxes(
        title_text="Продажи, ₽",
        secondary_y=False,
        showgrid=True,
        gridcolor="rgba(0,0,0,0.16)", 
        gridwidth=1.2,
        zeroline=False)
    
    fig_sales.update_yaxes(title_text="Доля ПЛ, %", secondary_y=True, showgrid=False)

    st.plotly_chart(fig_sales, use_container_width=True)

# вертикальный отступ между графиками
    st.markdown("<div style='height:30px'></div>", unsafe_allow_html=True)

# ===================== CHART 2 ============================
    BAR_COLOR_2 = "rgba(59, 130, 246, 0.28)"  # другой зелёный для столбиков 2-го графика

    fig_quality = make_subplots(specs=[[{"secondary_y": True}]])

    # столбики: активные участники (левая ось)
    fig_quality.add_trace(
        go.Bar(
            x=df["Месяц"],
            y=df["Активные участники, чел"],
            name="Активные участники, чел",
            marker=dict(color="rgba(59, 130, 246, 0.28)"),
            ),
        secondary_y=False,
    )

    # линия: доля повторных покупок (правая ось)
    fig_quality.add_trace(
        go.Scatter(
            x=df["Месяц"],
            y=df["Доля повторных покупок, %"],
            name="Доля повторных покупок, %",
            mode="lines+markers",
            line=dict(color="rgb(37, 99, 235)", width=2),
        ),
        secondary_y=True,
    )

    fig_quality.update_layout(
        template="plotly_white",
        height=420,
        margin=dict(l=20, r=20, t=52, b=20),
        title=dict(
            text="Активные участники и повторные покупки",
            x=0,
            xanchor="left",
            font=dict(size=16),
            pad=dict(b=12),
        ),
        legend=dict(
            orientation="h",
            yanchor="top",
            y=0.96,
            xanchor="left",
            x=0,
        ),
    )

    # Левая ось: с линиями (gridlines)
    fig_quality.update_yaxes(
        title_text="Активные участники, чел",
        secondary_y=False,
        showgrid=True,
        gridcolor="rgba(0,0,0,0.16)",
        gridwidth=1.2,
        zeroline=False,
    )

    # Правая ось: без линий (gridlines off)
    fig_quality.update_yaxes(
        title_text="Доля повторных покупок, %",
        secondary_y=True,
        showgrid=False,
        zeroline=False,
    )

    st.plotly_chart(fig_quality, use_container_width=True)



# ===================== COMMENTS ===========================
    # Комментарии — 2 смысловых абзаца, без зелёного фона
    st.markdown(
        """
<div class="note">
<b>Комментарий.</b> После запуска игровой механики в октябре (2025-10) в демо-данных виден рост активности. Параллельно растёт доля повторных покупок — это сигнал, что сценарии онбординга и мотивации на 2-ю покупку стали работать лучше.
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown(
        """
<div class="note">
<b>Рекомендации.</b> Зафиксируй гипотезу эффектов через контроль (когорта/регион/holdout) и проверь экономику стимулов: рост активности должен окупаться. Дальше усили 2-ю покупку — триггерами и миссиями для новых участников, и отдельной веткой для тех, кто застрял на 1-й транзакции.
</div>
""",
        unsafe_allow_html=True,
    )

    st.markdown("---")
    st.subheader("Таблица")

    matrix = build_report_matrix(df)

    from streamlit import column_config

    # ширины: "Показатель" шире, месяцы одинаковые
    month_columns = {
        col: column_config.TextColumn(label=col, width="small")
        for col in matrix.columns
    }
    month_columns = {
        "Показатель": column_config.TextColumn(label="Показатель", width="medium"),
        **month_columns,
    }

    # выводим по блокам, чтобы не было горизонтального скролла
    blocks = matrix.index.get_level_values(0).unique().tolist()

    for b in blocks:
        st.markdown(f"**{b}**")

        # превращаем "Показатель" в колонку (чтобы не терялся при hide_index=True)
        sub = matrix.loc[b].copy().reset_index()
        sub = sub.rename(columns={sub.columns[0]: "Показатель"})

        st.data_editor(
            sub,
            use_container_width=True,
            height=min(520, 42 + 34 * len(sub)),
            hide_index=True,
            disabled=True,
            column_config=month_columns,
        )

        st.markdown("<div style='height:12px'></div>", unsafe_allow_html=True)
