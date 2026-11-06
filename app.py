import io
import math
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Управление запасами и план закупок",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --- 1. ПАРСЕР ВЕДОМОСТИ ДВИЖЕНИЯ ТОВАРОВ (НОВЫЙ ОТЧЕТ) ---
def parse_movement_report(file_bytes):
    df_raw = pd.read_excel(file_bytes, sheet_name=0)

    # Определение периода анализа из шапки (по умолчанию 371 день: 01.11.25 - 06.11.26)
    days_in_period = 371
    for r in range(1, 6):
        val = str(df_raw.iloc[r, 2])
        if "Период отчета" in val:
            dates = re.findall(r"\d{2}\.\d{2}\.\d{4}", val)
            if len(dates) == 2:
                d1 = pd.to_datetime(dates[0], format="%d.%m.%Y")
                d2 = pd.to_datetime(dates[1], format="%d.%m.%Y")
                days_in_period = max(1, (d2 - d1).days)

    items = []
    current_supplier = "Не указан"

    for idx in range(12, len(df_raw)):
        col0 = str(df_raw.iloc[idx, 0]).strip()
        col3 = df_raw.iloc[idx, 3]

        if col0 in ["Склад Курьяново", "Итого", "nan"]:
            continue

        # Если в колонке 3 пусто — это строка поставщика/производителя
        if pd.isna(col3) or str(col3).strip() == "":
            current_supplier = col0
            continue

        name = str(col3).strip()

        # Кратность короба из наименования (например, ': 6')
        pack_match = re.search(r":\s*(\d+)\s*,?", name)
        pack_size = int(pack_match.group(1)) if pack_match else 6

        # Данные движения
        qty_in = (
            float(df_raw.iloc[idx, 9]) if pd.notna(df_raw.iloc[idx, 9]) else 0.0
        )
        qty_out = (
            float(df_raw.iloc[idx, 11])
            if pd.notna(df_raw.iloc[idx, 11])
            else 0.0
        )
        stock_end = (
            float(df_raw.iloc[idx, 14])
            if pd.notna(df_raw.iloc[idx, 14])
            else 0.0
        )
        val_end = (
            float(df_raw.iloc[idx, 15])
            if pd.notna(df_raw.iloc[idx, 15])
            else 0.0
        )

        cost_price = (val_end / stock_end) if stock_end > 0 else 0.0
        adu = qty_out / days_in_period  # Среднедневной расход

        items.append(
            {
                "Поставщик": current_supplier,
                "Артикул": col0,
                "Наименование": name,
                "Кратность (бут)": pack_size,
                "Себестоимость (руб)": round(cost_price, 2),
                "Приход (период)": int(qty_in),
                "Расход (период)": int(qty_out),
                "Остаток (бут)": int(stock_end),
                "В пути (бут)": 0,  # Заполнится, если загружен 1-й отчет
                "Дефицит (бут)": 0,  # Заполнится, если загружен 1-й отчет
                "Расход в день (бут)": round(adu, 3),
            }
        )

    return pd.DataFrame(items), days_in_period


# --- 2. ПАРСЕР ДОСТУПНОСТИ И ТРАНЗИТА (ПЕРВЫЙ ОТЧЕТ) ---
def enrich_with_availability(df_main, file_bytes):
    df_avail = pd.read_excel(file_bytes, sheet_name=0)
    transit_map = {}
    deficit_map = {}

    for idx in range(8, len(df_avail)):
        col0 = str(df_avail.iloc[idx, 0]).strip()
        col3 = df_avail.iloc[idx, 3]

        if pd.isna(col3) or str(col3).strip() == "" or col0 in ["nan", "Итого"]:
            continue

        incoming = (
            float(df_avail.iloc[idx, 11])
            if pd.notna(df_avail.iloc[idx, 11])
            else 0.0
        )
        to_fulfill = (
            float(df_avail.iloc[idx, 15])
            if pd.notna(df_avail.iloc[idx, 15])
            else 0.0
        )

        transit_map[col0] = int(incoming)
        deficit_map[col0] = int(to_fulfill)

    df_main["В пути (бут)"] = (
        df_main["Артикул"].map(transit_map).fillna(0).astype(int)
    )
    df_main["Дефицит (бут)"] = (
        df_main["Артикул"].map(deficit_map).fillna(0).astype(int)
    )
    return df_main


# --- 3. ИНТЕРФЕЙС И ЗАГРУЗКА ---
st.title("🍷 Дашборд оборачиваемости и планирования закупок")
st.caption(
    "Склад Курьяново | Алгоритм расчета Reorder Point, страхового запаса и бюджета закупки"
)

with st.sidebar:
    st.header("📂 Входные данные")
    f_movement = st.file_uploader(
        "1. Ведомость по товарам (Движение + Цены)", type=["xlsx"]
    )
    f_avail = st.file_uploader(
        "2. Отчет по доступности (Транзит + Заказы)", type=["xlsx"]
    )

    st.header("⚙️ Логистические нормативы")
    lead_time = st.slider(
        "Срок поставки (Lead Time, дни)", min_value=3, max_value=60, value=14
    )
    target_cover = st.slider(
        "Целевой запас склада (дней)", min_value=14, max_value=90, value=45
    )
    safety_days = st.slider(
        "Страховой запас (Safety Stock, дни)",
        min_value=0,
        max_value=30,
        value=10,
    )

# Автозагрузка локальных файлов при запуске
try:
    if f_movement is not None:
        df, days_cnt = parse_movement_report(f_movement)
    else:
        df, days_cnt = parse_movement_report(
            "Остатки и движение Курьяново 01.11.25-06.11.26.xlsx"
        )

    if f_avail is not None:
        df = enrich_with_availability(df, f_avail)
    else:
        try:
            df = enrich_with_availability(df, "Остатки и доступность 06.11.26.xlsx")
        except Exception:
            pass
except Exception as e:
    st.error(f"Ошибка загрузки данных: {e}")
    st.stop()

# --- 4. МАТЕМАТИКА ОБОРАЧИВАЕМОСТИ И ЗАКУПКИ ---
# Доступный запас = остаток + транзит
df["Всего доступно (бут)"] = df["Остаток (бут)"] + df["В пути (бут)"]

# Дни покрытия запасом (если расход = 0, запас условно 999 дней)
df["Дней покрытия"] = df.apply(
    lambda r: round(r["Всего доступно (бут)"] / r["Расход в день (бут)"], 1)
    if r["Расход в день (бут)"] > 0
    else 999.0,
    axis=1,
)

# Точка заказа (Reorder Point): ROP = Спрос * (Lead Time + Safety Stock)
df["Точка заказа ROP (бут)"] = (
    df["Расход в день (бут)"] * (lead_time + safety_days)
).round(0)

# Целевой нормативный запас: Спрос * (Lead Time + Target Cover)
df["Целевой сток (бут)"] = (
    df["Расход в день (бут)"] * (lead_time + target_cover)
).round(0)


# Расчет потребности в заказе с округлением до коробок
def calculate_order(row):
    need = (row["Целевой сток (бут)"] + row["Дефицит (бут)"]) - row[
        "Всего доступно (бут)"
    ]
    if need <= 0:
        return 0
    pack = row["Кратность (бут)"]
    return int(math.ceil(need / pack) * pack)


df["К заказу (бут)"] = df.apply(calculate_order, axis=1)
df["К заказу (кор)"] = (df["К заказу (бут)"] / df["Кратность (бут)"]).astype(
    int
)
df["Бюджет закупки (руб)"] = (
    df["К заказу (бут)"] * df["Себестоимость (руб)"]
).round(2)


# Классификация статуса запаса
def evaluate_status(row):
    if row["Дефицит (бут)"] > 0:
        return "🚨 Дефицит по клиентам"
    if row["Всего доступно (бут)"] <= row["Точка заказа ROP (бут)"]:
        return "⚠️ Пора заказывать"
    if row["Дней покрытия"] > (target_cover * 2):
        return "📦 Избыток / Неликвид"
    return "✅ В норме"


df["Статус стока"] = df.apply(evaluate_status, axis=1)

# Фильтр по поставщику
suppliers_list = ["Все поставщики"] + sorted(df["Поставщик"].unique().tolist())
selected_supp = st.sidebar.selectbox("Фильтр по производителю", suppliers_list)

filtered = (
    df if selected_supp == "Все поставщики" else df[df["Поставщик"] == selected_supp]
)

# --- 5. МЕТРИКИ И KPI ДАШБОРДА ---
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric(
    "Складской сток (бут)",
    f"{filtered['Остаток (бут)'].sum():,}".replace(",", " "),
)
c2.metric("В пути (бут)", f"{filtered['В пути (бут)'].sum():,}".replace(",", " "))
c3.metric(
    "Годовой расход (бут)",
    f"{filtered['Расход (период)'].sum():,}".replace(",", " "),
)
c4.metric(
    "Потребность к заказу (бут)",
    f"{filtered['К заказу (бут)'].sum():,}".replace(",", " "),
)
c5.metric(
    "Бюджет заказа (руб)",
    f"{filtered['Бюджет закупки (руб)'].sum():,.0f}".replace(",", " ") + " ₽",
)

st.markdown("---")

# --- 6. ТАБЛИЦА ПЛАНА ЗАКУПОК И ЭКСПОРТ ---
col_head, col_toggle = st.columns([3, 1])
with col_head:
    st.subheader("Сформированный план пополнения запасов")
with col_toggle:
    only_orders = st.toggle("Только позиции к заказу", value=False)

table_df = (
    filtered[filtered["К заказу (бут)"] > 0] if only_orders else filtered
)

cols_to_show = [
    "Поставщик",
    "Наименование",
    "Статус стока",
    "Себестоимость (руб)",
    "Расход в день (бут)",
    "Остаток (бут)",
    "В пути (бут)",
    "Дней покрытия",
    "Кратность (бут)",
    "К заказу (бут)",
    "К заказу (кор)",
    "Бюджет закупки (руб)",
]

st.dataframe(
    table_df[cols_to_show].style.format(
        {
            "Себестоимость (руб)": "{:.2f} ₽",
            "Расход в день (бут)": "{:.2f}",
            "Бюджет закупки (руб)": "{:,.2f} ₽",
        }
    ),
    use_container_width=True,
    height=480,
)

# Выгрузка сформированного заказа в Excel
output = io.BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    table_df[cols_to_show].to_excel(
        writer, index=False, sheet_name="План закупок"
    )

st.download_button(
    label="📥 Экспортировать план закупок в Excel",
    data=output.getvalue(),
    file_name="План_закупок_Курьяново.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
