import streamlit as st
import numpy as np
import pandas as pd
from pathlib import Path
from typing import Dict, List, Tuple
import re


import plotly.graph_objects as go
from plotly.subplots import make_subplots


KPI_FILE = "KPI_loyalty_metrics_full_ru.xlsx"

GREEN_ACCENT = "rgb(46, 139, 87)"
BAR_COLOR = "rgba(46, 139, 87, 0.28)"

BAR_COLOR_2 = "rgba(59, 130, 246, 0.24)"
LINE_COLOR_2 = "rgb(37, 99, 235)"

# -------------------- Форматирование чисел (RU) --------------------

def fmt_int_ru(x) -> str:
    """
    Целые числа с пробелами: 112726 -> '112 726'
    """
    try:
        return f"{int(round(float(x))):,}".replace(",", " ")
    except Exception:
        return "—"


def fmt_money_ru(x) -> str:
    """
    Деньги: 74501199 -> '74 501 199 ₽'
    """
    try:
        return f"{int(round(float(x))):,}".replace(",", " ") + " ₽"
    except Exception:
        return "—"


def fmt_pct_ru(x, decimals: int = 1) -> str:
    """
    Проценты: 51.63 -> '51,6 %'
    """
    try:
        return f"{float(x):.{decimals}f}".replace(".", ",") + " %"
    except Exception:
        return "—"


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
    Демо-данные на 6 месяцев (2025-07 .. 2025-12) со связями:
    - Общие продажи -> продажи ПЛ (доля) -> транзакции ПЛ -> чек/частота -> повторка
    - База/статусы
    - Бонусы: начислено/списано/сгорело/остаток + доли
    - Уровни A/B/C: выручка/скидка/структура
    - ПГ/Цели + индексы

    События:
    - 2025-10: запуск игровой механики
    - 2025-11: дедлайн/акция на списание
    """

    rng = np.random.default_rng(7)
    months = pd.period_range("2025-07", "2025-12", freq="M").astype(str)

    # --- БАЗА: общие продажи (₽) ---
    total_sales = np.array([132e6, 128e6, 136e6, 141e6, 138e6, 146e6], dtype=float)
    total_sales *= rng.normal(1.0, 0.012, size=6)

    # --- ДОЛЯ ПЛ в продажах (40–60%), до событий ---
    loyalty_share = np.array([0.47, 0.46, 0.48, 0.49, 0.49, 0.50], dtype=float)
    loyalty_share += rng.normal(0.0, 0.004, size=6)

    # --- Повторка (в долях), до событий ---
    repeat_rate = np.array([0.33, 0.32, 0.34, 0.345, 0.345, 0.35], dtype=float)
    repeat_rate += rng.normal(0.0, 0.008, size=6)

    # --- Активные участники (чел), до событий ---
    active_members = np.array([98000, 96000, 100000, 103000, 104000, 106000], dtype=float)
    active_members *= rng.normal(1.0, 0.012, size=6)

    # --- Событие 1: октябрь — запуск игровой механики ---
    i_oct = list(months).index("2025-10")
    # эффект: +8% активных в октябре, +5% в ноябре, +6% в декабре (устойчивый, но не фантастический)
    active_members[i_oct:] *= np.array([1.08, 1.05, 1.06], dtype=float)

    # доля ПЛ: +2 п.п. в окт, +1 п.п. в ноя, +1.5 п.п. в дек
    loyalty_share[i_oct:] += np.array([0.020, 0.010, 0.015], dtype=float)

    # повторка: +2.5 п.п. в окт, +1.5 п.п. в ноя, +2 п.п. в дек (в долях)
    repeat_rate[i_oct:] += np.array([0.025, 0.015, 0.020], dtype=float)

    # --- Событие 2: ноябрь — “дедлайн/акция на списание бонусов” ---
    i_nov = list(months).index("2025-11")
    loyalty_share[i_nov] += 0.005  # чуть выше доля ПЛ в ноябре

    # клипуем в реалистичные рамки
    loyalty_share = np.clip(loyalty_share, 0.40, 0.60)
    repeat_rate = np.clip(repeat_rate, 0.20, 0.55)
    active_members = np.clip(active_members, 70000, 180000).round().astype(int)

    loyalty_sales = (total_sales * loyalty_share).round().astype(int)

    # --- Средние чеки ---
    avg_basket_member = np.array([3850, 3780, 3920, 3980, 4050, 4180], dtype=float)
    avg_basket_member *= rng.normal(1.0, 0.01, size=6)

    avg_basket_non_member = avg_basket_member * rng.normal(0.88, 0.01, size=6)

    # --- Транзакции ПЛ согласуем с sales и avg basket ---
    loyalty_transactions = (loyalty_sales / avg_basket_member).round().astype(int)

    # --- Частота ---
    avg_freq_member = loyalty_transactions / np.maximum(active_members, 1)
    avg_freq_non_member = avg_freq_member * rng.normal(0.72, 0.03, size=6)

    # --- База участников и статусы ---
    db_total = np.array([510000, 515000, 522000, 532000, 538000, 545000], dtype=int)

    new_members = np.array([12000, 11500, 14000, 16000, 15000, 17000], dtype=int)
    lost_members = np.array([6500, 7000, 6800, 7200, 7100, 7400], dtype=int)
    reactivated_members = np.array([7200, 6900, 7600, 8800, 8300, 9100], dtype=int)

    inactive_members = (db_total - active_members).astype(int)

    # --- Бонусы (1 бонус = 1 ₽ условно) ---
    points_issued = (loyalty_sales * 0.02).round().astype(int)  # ~2% от продаж ПЛ
    redemption_rate = np.clip(rng.normal(0.62, 0.04, size=6), 0.45, 0.80)

    # игровой запуск слегка повышает вовлечённость в списание
    redemption_rate[i_oct:] += np.array([0.03, 0.02, 0.02], dtype=float)

    # ноябрьский дедлайн даёт разовый всплеск списаний
    redemption_rate[i_nov] += 0.08
    redemption_rate = np.clip(redemption_rate, 0.45, 0.85)

    points_redeemed = (points_issued * redemption_rate).round().astype(int)

    expiration_rate = np.clip(rng.normal(0.05, 0.012, size=6), 0.01, 0.10)
    points_expired = (points_issued * expiration_rate).round().astype(int)

    total_active_points = []
    balance = 18_000_000
    for i in range(6):
        balance = balance + int(points_issued[i]) - int(points_redeemed[i]) - int(points_expired[i])
        total_active_points.append(max(balance, 0))
    total_active_points = np.array(total_active_points, dtype=int)

    # --- Уровни (A/B/C) ---
    tier_share_members = np.array([0.15, 0.35, 0.50])  # A, B, C
    tier_members_A = (db_total * tier_share_members[0]).round().astype(int)
    tier_members_B = (db_total * tier_share_members[1]).round().astype(int)
    tier_members_C = (db_total * tier_share_members[2]).round().astype(int)

    tier_share_revenue = np.array([0.52, 0.30, 0.18])  # A, B, C
    rev_A = (loyalty_sales * tier_share_revenue[0]).round().astype(int)
    rev_B = (loyalty_sales * tier_share_revenue[1]).round().astype(int)
    rev_C = (loyalty_sales * tier_share_revenue[2]).round().astype(int)

    disc_A = np.array([5, 5, 5, 6, 6, 6], dtype=int)
    disc_B = np.array([3, 3, 3, 3, 4, 4], dtype=int)
    disc_C = np.array([1, 1, 1, 1, 1, 2], dtype=int)

    # --- Прошлый год и цели (для индексов) ---
    last_year_loyalty_sales = (loyalty_sales / rng.normal(1.06, 0.01, size=6)).round().astype(int)
    last_year_loyalty_tx = (loyalty_transactions / rng.normal(1.04, 0.01, size=6)).round().astype(int)
    last_year_avg_basket_member = (avg_basket_member / rng.normal(1.03, 0.01, size=6)).round(2)
    last_year_avg_freq_member = (avg_freq_member / rng.normal(1.02, 0.01, size=6))
    last_year_repeat_rate = np.clip(repeat_rate - rng.normal(0.015, 0.004, size=6), 0.15, 0.55)
    last_year_db_total = (db_total / rng.normal(1.04, 0.005, size=6)).round().astype(int)

    target_loyalty_sales = (loyalty_sales * rng.normal(1.03, 0.01, size=6)).round().astype(int)
    target_loyalty_tx = (loyalty_transactions * rng.normal(1.02, 0.01, size=6)).round().astype(int)
    target_avg_basket_member = (avg_basket_member * rng.normal(1.01, 0.005, size=6)).round(2)
    target_avg_freq_member = (avg_freq_member * rng.normal(1.01, 0.005, size=6))
    target_repeat_rate = np.clip(repeat_rate + rng.normal(0.01, 0.004, size=6), 0.20, 0.60)
    target_db_total = (db_total * rng.normal(1.02, 0.005, size=6)).round().astype(int)

    # Индексы как отклонение в % (дельта): +3.2% / -1.1%
    idx_sales_vs_ly = (loyalty_sales / last_year_loyalty_sales - 1) * 100
    idx_sales_vs_target = (loyalty_sales / target_loyalty_sales - 1) * 100

    idx_tx_vs_ly = (loyalty_transactions / last_year_loyalty_tx - 1) * 100
    idx_tx_vs_target = (loyalty_transactions / target_loyalty_tx - 1) * 100

    idx_db_vs_ly = (db_total / last_year_db_total - 1) * 100
    idx_db_vs_target = (db_total / target_db_total - 1) * 100


    df = pd.DataFrame({
        "Месяц": months,

        # Продажи
        "Общие продажи, ₽": total_sales.round().astype(int),
        "Продажи по ПЛ, ₽": loyalty_sales.round().astype(int),
        "Доля ПЛ в общих продажах, %": loyalty_share * 100,
        "Покупки участников ПЛ, шт": loyalty_transactions,
        "Активные участники, чел": active_members,
        "Средний чек участника, ₽": avg_basket_member.round(2),
        "Средний чек не-участника, ₽": avg_basket_non_member.round(2),
        "Средняя частота покупок участника, раз": avg_freq_member,
        "Доля повторных покупок, %": repeat_rate * 100,

        # Индексы/сравнения
        "Индекс продаж ПЛ к прошлому году, %": idx_sales_vs_ly,
        "Индекс продаж ПЛ к цели, %": idx_sales_vs_target,
        "Индекс покупок ПЛ к прошлому году, %": idx_tx_vs_ly,
        "Индекс покупок ПЛ к цели, %": idx_tx_vs_target,

        # База
        "Размер базы участников ПЛ, чел": db_total,
        "Индекс базы к прошлому году, %": idx_db_vs_ly,
        "Индекс базы к цели, %": idx_db_vs_target,
        "Новые участники, чел": new_members,
        "Неактивные участники, чел": inactive_members,
        "Потерянные участники, чел": lost_members,
        "Реактивированные участники, чел": reactivated_members,

        # Бонусы
        "Начисленные бонусы, шт": points_issued,
        "Списанные бонусы, шт": points_redeemed,
        "Сгоревшие бонусы, шт": points_expired,
        "Активные бонусы, шт": total_active_points,
        "Доля списания бонусов, %": (points_redeemed / np.maximum(points_issued, 1) * 100),
        "Доля сгорания бонусов, %": (points_expired / np.maximum(points_issued, 1) * 100),

        # Уровни
        "Выручка уровня A, ₽": rev_A,
        "Выручка уровня B, ₽": rev_B,
        "Выручка уровня C, ₽": rev_C,
        "Скидка уровня A, %": disc_A,
        "Скидка уровня B, %": disc_B,
        "Скидка уровня C, %": disc_C,
        "Участники уровня A, чел": tier_members_A,
        "Участники уровня B, чел": tier_members_B,
        "Участники уровня C, чел": tier_members_C,

        # Прошлый год
        "ПГ: Продажи по ПЛ, ₽": last_year_loyalty_sales,
        "ПГ: Покупки участников ПЛ, шт": last_year_loyalty_tx,
        "ПГ: Средний чек участника, ₽": last_year_avg_basket_member,
        "ПГ: Средняя частота покупок, раз": last_year_avg_freq_member,
        "ПГ: Доля повторных покупок, %": last_year_repeat_rate * 100,
        "ПГ: Размер базы участников ПЛ, чел": last_year_db_total,

        # Цели
        "Цель: Продажи по ПЛ, ₽": target_loyalty_sales,
        "Цель: Покупки участников ПЛ, шт": target_loyalty_tx,
        "Цель: Средний чек участника, ₽": target_avg_basket_member,
        "Цель: Средняя частота покупок, раз": target_avg_freq_member,
        "Цель: Доля повторных покупок, %": target_repeat_rate * 100,
        "Цель: Размер базы участников ПЛ, чел": target_db_total,
    })
    return df

def build_report_matrix(df: pd.DataFrame) -> pd.DataFrame:
    months = df["Месяц"].tolist()
    
    # ---- форматирование ----
    def fmt_int(x):
        if pd.isna(x): return "–"
        return f"{int(round(float(x))):,}".replace(",", " ")

    def fmt_money(x):
        if pd.isna(x): return "–"
        return f"{int(round(float(x))):,}".replace(",", " ") + " ₽"

    def fmt_pct(x):
        if pd.isna(x): return "–"
        return f"{float(x):.1f} %".replace(".", ",")

    def fmt_idx_delta(x):
        if pd.isna(x): return "–"
        return f"{float(x):+.1f} %".replace(".", ",")

    def fmt_float(x):
        if pd.isna(x): return "–"
        return f"{float(x):.2f}".replace(".", ",")

    # ---- структура отчёта + units ----
    REPORT = {
        "Продажи": [
            ("Продажи по программе лояльности", "₽"),
            ("Индекс к прошлому году (продажи ПЛ)", "%"),
            ("Индекс к цели (продажи ПЛ)", "%"),
            ("Доля ПЛ в общих продажах", "%"),
            ("Покупки участников ПЛ", "шт"),
            ("Индекс к прошлому году (покупки ПЛ)", "%"),
            ("Индекс к цели (покупки ПЛ)", "%"),
            ("Количество активных участников", "чел"),
            ("Средний чек участника", "₽"),
            ("Средний чек не-участника", "₽"),
            ("Средняя частота покупок участника", "раз"),
            ("Доля повторных покупок", "%"),
        ],
        "База участников": [
            ("Размер базы участников ПЛ", "чел"),
            ("Индекс к прошлому году (база)", "%"),
            ("Индекс к цели (база)", "%"),
            ("Новые участники", "чел"),
            ("Активные участники", "чел"),
            ("Неактивные участники", "чел"),
            ("Потерянные участники", "чел"),
            ("Реактивированные участники", "чел"),
        ],
        "Бонусная валюта": [
            ("Начисленные бонусы", "шт"),
            ("Списанные бонусы", "шт"),
            ("Сгоревшие бонусы", "шт"),
            ("Активные бонусы", "шт"),
            ("Доля списания бонусов", "%"),
            ("Доля сгорания бонусов", "%"),
        ],
        "Уровни программы": [
            ("Выручка уровня A", "₽"),
            ("Выручка уровня B", "₽"),
            ("Выручка уровня C", "₽"),
            ("Скидка уровня A", "%"),
            ("Скидка уровня B", "%"),
            ("Скидка уровня C", "%"),
            ("Участники уровня A", "чел"),
            ("Участники уровня B", "чел"),
            ("Участники уровня C", "чел"),
        ],
        "Дополнительные данные": [
            ("Общие продажи", "₽"),
        ],
    }

    # ---- маппинг: (блок, показатель) -> (колонка df, formatter) ----
    MAP = {
        ("Продажи", "Продажи по программе лояльности"): ("Продажи по ПЛ, ₽", fmt_money),
        ("Продажи", "Индекс к прошлому году (продажи ПЛ)"): ("Индекс продаж ПЛ к прошлому году, %", fmt_idx_delta),
        ("Продажи", "Индекс к цели (продажи ПЛ)"): ("Индекс продаж ПЛ к цели, %", fmt_idx_delta),
        ("Продажи", "Доля ПЛ в общих продажах"): ("Доля ПЛ в общих продажах, %", fmt_pct),
        ("Продажи", "Покупки участников ПЛ"): ("Покупки участников ПЛ, шт", fmt_int),
        ("Продажи", "Индекс к прошлому году (покупки ПЛ)"): ("Индекс покупок ПЛ к прошлому году, %", fmt_idx_delta),
        ("Продажи", "Индекс к цели (покупки ПЛ)"): ("Индекс покупок ПЛ к цели, %", fmt_idx_delta),
        ("Продажи", "Количество активных участников"): ("Активные участники, чел", fmt_int),
        ("Продажи", "Средний чек участника"): ("Средний чек участника, ₽", fmt_money),
        ("Продажи", "Средний чек не-участника"): ("Средний чек не-участника, ₽", fmt_money),
        ("Продажи", "Средняя частота покупок участника"): ("Средняя частота покупок участника, раз", fmt_float),
        ("Продажи", "Доля повторных покупок"): ("Доля повторных покупок, %", fmt_pct),

        ("База участников", "Размер базы участников ПЛ"): ("Размер базы участников ПЛ, чел", fmt_int),
        ("База участников", "Индекс к прошлому году (база)"): ("Индекс базы к прошлому году, %", fmt_idx_delta),
        ("База участников", "Индекс к цели (база)"): ("Индекс базы к цели, %", fmt_idx_delta),
        ("База участников", "Новые участники"): ("Новые участники, чел", fmt_int),
        ("База участников", "Активные участники"): ("Активные участники, чел", fmt_int),
        ("База участников", "Неактивные участники"): ("Неактивные участники, чел", fmt_int),
        ("База участников", "Потерянные участники"): ("Потерянные участники, чел", fmt_int),
        ("База участников", "Реактивированные участники"): ("Реактивированные участники, чел", fmt_int),

        ("Бонусная валюта", "Начисленные бонусы"): ("Начисленные бонусы, шт", fmt_int),
        ("Бонусная валюта", "Списанные бонусы"): ("Списанные бонусы, шт", fmt_int),
        ("Бонусная валюта", "Сгоревшие бонусы"): ("Сгоревшие бонусы, шт", fmt_int),
        ("Бонусная валюта", "Активные бонусы"): ("Активные бонусы, шт", fmt_int),
        ("Бонусная валюта", "Доля списания бонусов"): ("Доля списания бонусов, %", fmt_pct),
        ("Бонусная валюта", "Доля сгорания бонусов"): ("Доля сгорания бонусов, %", fmt_pct),

        ("Уровни программы", "Выручка уровня A"): ("Выручка уровня A, ₽", fmt_money),
        ("Уровни программы", "Выручка уровня B"): ("Выручка уровня B, ₽", fmt_money),
        ("Уровни программы", "Выручка уровня C"): ("Выручка уровня C, ₽", fmt_money),
        ("Уровни программы", "Скидка уровня A"): ("Скидка уровня A, %", fmt_pct),
        ("Уровни программы", "Скидка уровня B"): ("Скидка уровня B, %", fmt_pct),
        ("Уровни программы", "Скидка уровня C"): ("Скидка уровня C, %", fmt_pct),
        ("Уровни программы", "Участники уровня A"): ("Участники уровня A, чел", fmt_int),
        ("Уровни программы", "Участники уровня B"): ("Участники уровня B, чел", fmt_int),
        ("Уровни программы", "Участники уровня C"): ("Участники уровня C, чел", fmt_int),

        ("Дополнительные данные", "Общие продажи"): ("Общие продажи, ₽", fmt_money),
    }

    # ---- строим матрицу “из структуры” ----
    rows = []
    for block, metric_list in REPORT.items():
        for metric, unit in metric_list:
            row = {"Блок": block, "Показатель": metric, "Ед.": unit}
            for m in months:
                row[m] = "–"
            rows.append(row)

    out = pd.DataFrame(rows).set_index(["Блок", "Показатель"])

    df_i = df.set_index("Месяц")

    for (block, metric), (col, formatter) in MAP.items():
        if col not in df_i.columns:
            continue
        series = df_i[col].reindex(months)
        out.loc[(block, metric), months] = [formatter(v) for v in series.values]

    return out


# -------------------- UI --------------------
tabs = st.tabs(["Пример отчёта", "Метрики"])

with tabs[1]:
    st.subheader("Справочник KPI программы лояльности")

    # --- читаем справочник из Excel ---
    from pathlib import Path

    BASE_DIR = Path(__file__).parent
    KPI_FILE = BASE_DIR / "KPI_loyalty_metrics_full_ru.xlsx"

    df_metrics = pd.read_excel(KPI_FILE, sheet_name="Sheet1")

    # --- приводим имена колонок к аккуратному виду (убираем пробелы) ---
    df_metrics.columns = [str(c).strip() for c in df_metrics.columns]

    required_cols = ["KPI", "Категория", "Описание", "Формула", "Рекомендации"]
    missing = [c for c in required_cols if c not in df_metrics.columns]
    if missing:
        st.error(
            "В файле Excel не найдены нужные колонки: "
            + ", ".join(missing)
            + ". Проверь названия столбцов на листе Sheet1."
        )
        st.stop()

    # --- поиск + фильтр по категории ---
    col_a, col_b = st.columns([2, 1])
    with col_a:
        q = st.text_input("Поиск по KPI / описанию / формуле", value="")
    with col_b:
        categories = ["Все"] + sorted(
            [x for x in df_metrics["Категория"].dropna().astype(str).unique().tolist() if x.strip() != ""]
        )
        cat = st.selectbox("Категория", categories, index=0)

    view = df_metrics.copy()

    if cat != "Все":
        view = view[view["Категория"].astype(str).str.strip() == cat]

    if q.strip():
        q_low = q.strip().lower()
        mask = (
            view["KPI"].astype(str).str.lower().str.contains(q_low, na=False)
            | view["Описание"].astype(str).str.lower().str.contains(q_low, na=False)
            | view["Формула"].astype(str).str.lower().str.contains(q_low, na=False)
            | view["Рекомендации"].astype(str).str.lower().str.contains(q_low, na=False)
        )
        view = view[mask]

    # --- рекомендации: превращаем ячейку в буллиты ---
    def bullets_from_cell(cell) -> str:
        if pd.isna(cell):
            return "—"
        s = str(cell).strip()
        if s == "" or s in {"-", "–", "—"}:
            return "—"

        # нормализуем разделители:
        # - строки -> пункты
        # - "•" -> пункты
        # - ";" тоже часто используют
        s = s.replace("\r\n", "\n").replace("\r", "\n")
        s = s.replace("•", "\n").replace(";", "\n")

        items = []
        for line in s.split("\n"):
            t = line.strip()
            if not t:
                continue
            # убираем лидирующие маркеры вроде "- " / "— " / "• "
            t = t.lstrip("-").lstrip("—").lstrip("–").strip()
            if t:
                items.append(t)

        if not items:
            return "—"

        return "<ul>" + "".join(f"<li>{i}</li>" for i in items) + "</ul>"

    view = view.copy()
    view["Рекомендации"] = view["Рекомендации"].apply(bullets_from_cell)

    # --- порядок и названия колонок для отображения ---
    view = view[["KPI", "Категория", "Описание", "Формула", "Рекомендации"]].rename(
        columns={
            "KPI": "KPI",
            "Категория": "Категория",
            "Описание": "Зачем смотреть",
            "Формула": "Формула",
            "Рекомендации": "Рекомендации",
        }
    )

    # --- компактный HTML-рендер (CSS класс report-table у тебя уже есть) ---
    html = view.to_html(index=False, escape=False)

    st.markdown(
        f"<div class='report-table'>{html}</div>",
        unsafe_allow_html=True,
    )


# ===== TAB 2: REPORT =====
with tabs[0]:
    st.subheader("")

    df = build_demo_report()

    # ---- 5 ключевых KPI (оставляем как было по смыслу) ----
    last = df.iloc[-1]
    prev = df.iloc[-2]

    # доля активных в базе считаем на лету (без отдельной колонки)
    base_last = max(1, float(last["Размер базы участников ПЛ, чел"]))
    base_prev = max(1, float(prev["Размер базы участников ПЛ, чел"]))
    active_share_last = 100 * float(last["Активные участники, чел"]) / base_last
    active_share_prev = 100 * float(prev["Активные участники, чел"]) / base_prev

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.metric(
            "Продажи по ПЛ, ₽",
            fmt_int_ru(last["Продажи по ПЛ, ₽"]) + " ₽",
            fmt_pct_ru(100 * (last["Продажи по ПЛ, ₽"] / prev["Продажи по ПЛ, ₽"] - 1), 1),
        )
    with c2:
        st.metric(
            "Доля ПЛ в общих продажах, %",
            fmt_pct_ru(last["Доля ПЛ в общих продажах, %"], 1),
            fmt_pct_ru(last["Доля ПЛ в общих продажах, %"] - prev["Доля ПЛ в общих продажах, %"], 1),
        )
    with c3:
        st.metric(
            "Активные участники, чел",
            fmt_int_ru(last["Активные участники, чел"]),
            fmt_pct_ru(100 * (last["Активные участники, чел"] / prev["Активные участники, чел"] - 1), 1),
        )
    with c4:
        st.metric(
            "Доля активных в базе, %",
            fmt_pct_ru(active_share_last, 1),
            fmt_pct_ru(active_share_last - active_share_prev, 1),
        )
    with c5:
        st.metric(
            "Доля повторных покупок, %",
            fmt_pct_ru(last["Доля повторных покупок, %"], 1),
            fmt_pct_ru(last["Доля повторных покупок, %"] - prev["Доля повторных покупок, %"], 1),
        )


# ===================== CHART 1 ============================
    # вертикальный отступ между графиками
    st.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)

    fig_sales = make_subplots(specs=[[{"secondary_y": True}]])

    # единые имена колонок (чтобы не ломалось)
    COL = {
        "month": "Месяц",
        "sales": "Продажи по ПЛ, ₽",
        "share": "Доля ПЛ в общих продажах, %",
    }

    # столбцы — продажи по ПЛ
    fig_sales.add_trace(
        go.Bar(
            x=df[COL["month"]],
            y=df[COL["sales"]],
            name="Продажи по ПЛ, ₽",
            marker=dict(
                color=BAR_COLOR,
                line=dict(color="rgba(46, 139, 87, 0.30)", width=1),
            ),
        ),
        secondary_y=False,
    )

    # линия — доля ПЛ в общих продажах
    fig_sales.add_trace(
        go.Scatter(
            x=df[COL["month"]],
            y=df[COL["share"]],
            name="Доля ПЛ в общих продажах, %",
            mode="lines+markers",
            line=dict(color=GREEN_ACCENT, width=2),
        ),
        secondary_y=True,
    )

    fig_sales.update_layout(
        template="plotly_white",
        height=420,
        margin=dict(l=20, r=20, t=40, b=20),
        title=dict(
            text="Динамика продаж участников ПЛ",
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

    # левая ось — деньги, с сеткой
    fig_sales.update_yaxes(
        title_text="Продажи, ₽",
        secondary_y=False,
        showgrid=True,
        gridcolor="rgba(0,0,0,0.16)",
        gridwidth=1.2,
        zeroline=False,
    )

    # правая ось — проценты, без сетки
    fig_sales.update_yaxes(
        title_text="Доля ПЛ, %",
        secondary_y=True,
        showgrid=False,
        zeroline=False,
    )

    st.plotly_chart(fig_sales, use_container_width=True)

    # вертикальный отступ между графиками
    st.markdown("<div style='height:40px'></div>", unsafe_allow_html=True)


    # ===================== CHART 2 ============================
    BAR_COLOR_2 = "rgba(59, 130, 246, 0.28)"  # отдельный цвет для столбиков 2-го графика
    LINE_COLOR_2 = "rgb(37, 99, 235)"         # линия 2-го графика

    fig_quality = make_subplots(specs=[[{"secondary_y": True}]])

    # единые имена колонок (чтобы не ломалось)
    COL2 = {
        "month": "Месяц",
        "active": "Активные участники, чел",
        "repeat": "Доля повторных покупок, %",
    }

    # столбики: активные участники (левая ось)
    fig_quality.add_trace(
        go.Bar(
            x=df[COL2["month"]],
            y=df[COL2["active"]],
            name="Активные участники, чел",
            marker=dict(color=BAR_COLOR_2),
        ),
        secondary_y=False,
    )

    # линия: доля повторных покупок (правая ось)
    fig_quality.add_trace(
        go.Scatter(
            x=df[COL2["month"]],
            y=df[COL2["repeat"]],
            name="Доля повторных покупок, %",
            mode="lines+markers",
            line=dict(color=LINE_COLOR_2, width=2),
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
    st.markdown(
        """
    <div class="note"> <b>Комментарий.</b> Во втором полугодии вклад
    программы лояльности в выручку растёт: увеличиваются продажи по ПЛ и
    её доля в общих продажах, а после запуска игровой механики также
    видно расширение активности и рост доли повторных покупок. В
    совокупности это указывает на усиление вовлечения участников и
    снижение доли разовых покупателей, однако для подтверждения
    эффекта необходима дополнительная оценка эффекта через замеры с контрольной группой.

    </div>
    """,
            unsafe_allow_html=True,
        )

    st.markdown(
            """
    <div class="note"> <b>Рекомендации.</b>Зафиксировать влияние
    программы лояльности через сравнение с контрольными группами или
    периодами без активных механик. Проверить, какие механики ПЛ (миссии,
    уровни, бонусы) дают вклад в повторную покупку и средний чек, и
    скорректировать правила начисления под фактическую отдачу. Далее усилить сценарии для стимулирования второй покупки
    и удержания клиентов: отдельные триггеры для новых участников и механики
    активации для тех, кто с 1-й покупкой. </div> """,
            unsafe_allow_html=True,
        )

    st.markdown("---")

    # ---- Таблица исходных значений — только здесь ----
    st.subheader("Данные")

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


