import streamlit as st
import pandas as pd
from pathlib import Path


@st.cache_data
def load_kpi_table(path: str) -> pd.DataFrame:
    """Загрузка справочника KPI из Excel-файла (полностью, на русском)."""
    return pd.read_excel(path)


@st.cache_data
def load_timeseries(file) -> pd.DataFrame:
    """
    Загрузка файла с динамикой показателей.
    Ожидается таблица с колонками:
    - 'Период'
    - 'Выручка участников'
    - 'Количество покупок'
    - 'Активные участники'
    - 'Участники с ≥2 покупками'
    """
    from pathlib import Path as _Path

    file_suffix = _Path(file.name).suffix.lower()
    if file_suffix == ".csv":
        df = pd.read_csv(file, sep=",")
    else:
        df = pd.read_excel(file)
    return df


st.set_page_config(
    page_title="Метрики программы лояльности",
    layout="wide"
)

st.title("Метрики программы лояльности")
st.caption("Интерактивный справочник и калькулятор KPI по программе лояльности")

st.subheader("Справочник KPI (полная версия)")

kpi_file_path = "KPI_loyalty_metrics_full_ru.xlsx"

try:
    if Path(kpi_file_path).exists():
        kpi_df = load_kpi_table(kpi_file_path)

        categories = ["Все"] + sorted(kpi_df["Категория"].dropna().unique().tolist())
        selected_cat = st.selectbox("Фильтр по категории", categories, index=0)

        if selected_cat != "Все":
            kpi_df_display = kpi_df[kpi_df["Категория"] == selected_cat]
        else:
            kpi_df_display = kpi_df

        st.dataframe(
            kpi_df_display,
            use_container_width=True
        )
    else:
        st.warning(
            f"Файл '{kpi_file_path}' не найден рядом с app.py. "
            f"Положи его в ту же папку, чтобы отобразить таблицу KPI."
        )
except Exception as e:
    st.warning("Не удалось загрузить таблицу KPI.")
    st.text(f"Ошибка: {e}")

st.markdown("---")

st.subheader("Калькулятор KPI по текущему периоду")

col_input_curr, col_output_curr = st.columns(2)

with col_input_curr:
    st.markdown("### Входные данные (текущий период)")

    revenue_curr = st.number_input(
        "Выручка от участников программы, ₽",
        min_value=0.0,
        value=120_000_000.0,
        step=100_000.0,
        format="%.2f"
    )

    tx_count_curr = st.number_input(
        "Количество покупок участников за период",
        min_value=0,
        value=400_000,
        step=1
    )

    member_count_curr = st.number_input(
        "Количество активных участников за период",
        min_value=1,
        value=250_000,
        step=1
    )

    repeat_members_curr = st.number_input(
        "Количество участников с ≥ 2 покупками",
        min_value=0,
        value=70_000,
        step=1
    )

with col_output_curr:
    st.markdown("### Расчёт KPI (текущий период)")

    if tx_count_curr > 0:
        avg_basket_curr = revenue_curr / tx_count_curr
    else:
        avg_basket_curr = 0.0

    if member_count_curr > 0:
        frequency_curr = tx_count_curr / member_count_curr
        repeat_rate_curr = repeat_members_curr / member_count_curr
    else:
        frequency_curr = 0.0
        repeat_rate_curr = 0.0

    st.metric(
        "Продажи по программе лояльности, ₽",
        f"{revenue_curr:,.0f}".replace(",", " ")
    )
    st.metric(
        "Средний чек, ₽",
        f"{avg_basket_curr:,.2f}".replace(",", " ")
    )
    st.metric(
        "Частота покупок",
        f"{frequency_curr:.2f}"
    )
    st.metric(
        "Repeat Purchase Rate, %",
        f"{repeat_rate_curr * 100:.1f}%"
    )

st.caption(
    "Пример: выручка 120 млн ₽, 400k покупок, 250k активных участников, "
    "70k участников с ≥2 покупками → средний чек 300 ₽, частота 1.60, RPR 28.0%."
)

st.markdown("---")

st.subheader("Сравнение с прошлым годом (YoY)")

st.markdown(
    "Введите показатели прошлого года, чтобы посчитать индексы роста по основным KPI."
)

col_prev_input, col_yoy_metrics = st.columns(2)

with col_prev_input:
    st.markdown("### Входные данные (прошлый год)")

    revenue_prev = st.number_input(
        "Выручка от участников программы в прошлом году, ₽",
        min_value=0.0,
        value=100_000_000.0,
        step=100_000.0,
        format="%.2f"
    )

    tx_count_prev = st.number_input(
        "Количество покупок участников в прошлом году",
        min_value=0,
        value=350_000,
        step=1
    )

    member_count_prev = st.number_input(
        "Количество активных участников в прошлом году",
        min_value=1,
        value=230_000,
        step=1
    )

    repeat_members_prev = st.number_input(
        "Количество участников с ≥ 2 покупками в прошлом году",
        min_value=0,
        value=60_000,
        step=1
    )

with col_yoy_metrics:
    st.markdown("### Индексы роста по KPI (текущий год vs прошлый)")

    if tx_count_prev > 0:
        avg_basket_prev = revenue_prev / tx_count_prev
    else:
        avg_basket_prev = 0.0

    if member_count_prev > 0:
        frequency_prev = tx_count_prev / member_count_prev
        repeat_rate_prev = repeat_members_prev / member_count_prev
    else:
        frequency_prev = 0.0
        repeat_rate_prev = 0.0

    def calc_delta(curr: float, prev: float) -> str:
        if prev <= 0:
            return "н/д"
        change = (curr / prev - 1.0) * 100.0
        sign = "+" if change >= 0 else ""
        return f"{sign}{change:.1f}%"

    delta_revenue = calc_delta(revenue_curr, revenue_prev)
    delta_avg_basket = calc_delta(avg_basket_curr, avg_basket_prev)
    delta_frequency = calc_delta(frequency_curr, frequency_prev)
    delta_repeat_rate = calc_delta(repeat_rate_curr, repeat_rate_prev)

    st.metric(
        "Продажи по программе, ₽",
        f"{revenue_curr:,.0f}".replace(",", " "),
        delta=delta_revenue
    )
    st.metric(
        "Средний чек, ₽",
        f"{avg_basket_curr:,.2f}".replace(",", " "),
        delta=delta_avg_basket
    )
    st.metric(
        "Частота покупок",
        f"{frequency_curr:.2f}",
        delta=delta_frequency
    )
    st.metric(
        "Repeat Purchase Rate, %",
        f"{repeat_rate_curr * 100:.1f}%",
        delta=delta_repeat_rate
    )

st.caption(
    "Индексы считаются как (текущий показатель / показатель прошлого года − 1) × 100%. "
    "Если в прошлом году показатель был 0, индекс не рассчитывается."
)

st.markdown("---")

st.subheader("Динамика KPI по периодам")

st.markdown(
    "Загрузите файл с динамикой показателей по периодам, чтобы построить графики. "
    "Ожидаются столбцы:\n"
    "- 'Период'\n"
    "- 'Выручка участников'\n"
    "- 'Количество покупок'\n"
    "- 'Активные участники'\n"
    "- 'Участники с ≥2 покупками'"
)

uploaded_file = st.file_uploader(
    "Загрузите CSV или Excel-файл с динамикой",
    type=["csv", "xlsx"]
)

if uploaded_file is not None:
    try:
        ts_df = load_timeseries(uploaded_file)

        required_columns = [
            "Период",
            "Выручка участников",
            "Количество покупок",
            "Активные участники",
            "Участники с ≥2 покупками"
        ]

        missing_columns = [col for col in required_columns if col not in ts_df.columns]

        if missing_columns:
            st.error("В загруженном файле не хватает колонок: " + ", ".join(missing_columns))
        else:
            ts_df = ts_df.copy()

            ts_df["Средний чек"] = ts_df["Выручка участников"] / ts_df["Количество покупок"]
            ts_df["Частота покупок"] = ts_df["Количество покупок"] / ts_df["Активные участники"]
            ts_df["Repeat Purchase Rate, %"] = (
                ts_df["Участники с ≥2 покупками"] / ts_df["Активные участники"] * 100.0
            )

            ts_df_display = ts_df[[
                "Период",
                "Выручка участников",
                "Количество покупок",
                "Активные участники",
                "Участники с ≥2 покупками",
                "Средний чек",
                "Частота покупок",
                "Repeat Purchase Rate, %"
            ]]

            st.markdown("### Таблица с рассчитанными KPI по периодам")
            st.dataframe(ts_df_display, use_container_width=True)

            ts_df_indexed = ts_df_display.set_index("Период")

            st.markdown("### График: Продажи по программе лояльности")
            st.line_chart(ts_df_indexed[["Выручка участников"]])

            st.markdown("### График: Средний чек")
            st.line_chart(ts_df_indexed[["Средний чек"]])

            st.markdown("### График: Частота покупок")
            st.line_chart(ts_df_indexed[["Частота покупок"]])

            st.markdown("### График: Repeat Purchase Rate, %")
            st.line_chart(ts_df_indexed[["Repeat Purchase Rate, %"]])

    except Exception as e:
        st.error("Не удалось обработать загруженный файл.")
        st.text(f"Ошибка: {e}")
else:
    st.info("Файл с динамикой пока не загружен. Для графиков загрузите CSV/XLSX.")
