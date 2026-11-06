import io
import math
import os
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Управление запасами и план закупок",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --- 1. ПАРСЕР ВЕДОМОСТИ ДВИЖЕНИЯ ТОВАРОВ ---
def razobrat_vedomost_dvizheniya(fayl_bayty):
    tablica_syraya = pd.read_excel(fayl_bayty, sheet_name=0)

    # Определение периода отчета в днях из шапки
    kolichestvo_dney = 371
    for stroka_nomer in range(1, 6):
        znachenie = str(tablica_syraya.iloc[stroka_nomer, 2])
        if "Период отчета" in znachenie:
            daty = re.findall(r"\d{2}\.\d{2}\.\d{4}", znachenie)
            if len(daty) == 2:
                data_nachala = pd.to_datetime(daty[0], format="%d.%m.%Y")
                data_konca = pd.to_datetime(daty[1], format="%d.%m.%Y")
                kolichestvo_dney = max(1, (data_konca - data_nachala).days)

    stroki_tovarov = []
    tekuschiy_postavschik = "Не указан"

    for nomer in range(12, len(tablica_syraya)):
        artikul_tekst = str(tablica_syraya.iloc[nomer, 0]).strip()
        nomenklatura_tekst = tablica_syraya.iloc[nomer, 3]

        if artikul_tekst in ["Склад Курьяново", "Итого", "nan"]:
            continue

        # Строка-группа (поставщик / производитель)
        if (
            pd.isna(nomenklatura_tekst)
            or str(nomenklatura_tekst).strip() == ""
        ):
            tekuschiy_postavschik = artikul_tekst
            continue

        naimenovanie = str(nomenklatura_tekst).strip()

        # Определение кратности короба из названия (например, ': 6')
        poisk_koroba = re.search(r":\s*(\d+)\s*,?", naimenovanie)
        kratnost_koroba = int(poisk_koroba.group(1)) if poisk_koroba else 6

        # Показатели движения
        prihod_kolichestvo = (
            float(tablica_syraya.iloc[nomer, 9])
            if pd.notna(tablica_syraya.iloc[nomer, 9])
            else 0.0
        )
        rashod_kolichestvo = (
            float(tablica_syraya.iloc[nomer, 11])
            if pd.notna(tablica_syraya.iloc[nomer, 11])
            else 0.0
        )
        konechnyj_ostatok = (
            float(tablica_syraya.iloc[nomer, 14])
            if pd.notna(tablica_syraya.iloc[nomer, 14])
            else 0.0
        )
        ocenka_konechnogo_ostatka = (
            float(tablica_syraya.iloc[nomer, 15])
            if pd.notna(tablica_syraya.iloc[nomer, 15])
            else 0.0
        )

        sebestoimost_edinicy = (
            (ocenka_konechnogo_ostatka / konechnyj_ostatok)
            if konechnyj_ostatok > 0
            else 0.0
        )
        srednesutochnyy_rashod = rashod_kolichestvo / kolichestvo_dney

        stroki_tovarov.append(
            {
                "Поставщик": tekuschiy_postavschik,
                "Артикул": artikul_tekst,
                "Наименование": naimenovanie,
                "Кратность (бут)": kratnost_koroba,
                "Себестоимость (руб)": round(sebestoimost_edinicy, 2),
                "Приход за период (бут)": int(prihod_kolichestvo),
                "Расход за период (бут)": int(rashod_kolichestvo),
                "Остаток на складе (бут)": int(konechnyj_ostatok),
                "Товары в пути (бут)": 0,
                "Потребность по заказам (бут)": 0,
                "Среднесуточный расход (бут/день)": round(
                    srednesutochnyy_rashod, 3
                ),
            }
        )

    return pd.DataFrame(stroki_tovarov), kolichestvo_dney


# --- 2. ПАРСЕР ДОСТУПНОСТИ И ТРАНЗИТА ---
def obogatit_dannymi_dostupnosti(osnovnaya_tablica, fayl_bayty):
    tablica_dostupnosti = pd.read_excel(fayl_bayty, sheet_name=0)
    karta_puti = {}
    karta_potrebnosti = {}

    for nomer in range(8, len(tablica_dostupnosti)):
        artikul_tekst = str(tablica_dostupnosti.iloc[nomer, 0]).strip()
        nomenklatura_tekst = tablica_dostupnosti.iloc[nomer, 3]

        if (
            pd.isna(nomenklatura_tekst)
            or str(nomenklatura_tekst).strip() == ""
            or artikul_tekst in ["nan", "Итого"]
        ):
            continue

        ozhidaetsya_postuplenie = (
            float(tablica_dostupnosti.iloc[nomer, 11])
            if pd.notna(tablica_dostupnosti.iloc[nomer, 11])
            else 0.0
        )
        k_obespecheniyu = (
            float(tablica_dostupnosti.iloc[nomer, 15])
            if pd.notna(tablica_dostupnosti.iloc[nomer, 15])
            else 0.0
        )

        karta_puti[artikul_tekst] = int(ozhidaetsya_postuplenie)
        karta_potrebnosti[artikul_tekst] = int(k_obespecheniyu)

    osnovnaya_tablica["Товары в пути (бут)"] = (
        osnovnaya_tablica["Артикул"].map(karta_puti).fillna(0).astype(int)
    )
    osnovnaya_tablica["Потребность по заказам (бут)"] = (
        osnovnaya_tablica["Артикул"].map(karta_potrebnosti).fillna(0).astype(int)
    )
    return osnovnaya_tablica


# --- 3. ИНТЕРФЕЙС И БОКОВАЯ ПАНЕЛЬ ---
st.title("🍷 Дашборд управления запасами и расчета потребности в закупках")
st.caption(
    "Склад Курьяново | Расчет точки перезаказа, страхового запаса и бюджета поставок"
)

with st.sidebar:
    st.header("📂 Загрузка отчетов 1С")
    zagruzhennaya_vedomost = st.file_uploader(
        "1. Ведомость по товарам (движение и цены)", type=["xlsx"]
    )
    zagruzhennaya_dostupnost = st.file_uploader(
        "2. Остатки и доступность (товары в пути и резервы)", type=["xlsx"]
    )

    st.header("⚙️ Складские нормативы")
    srok_dostavki = st.slider(
        "Срок поставки (плечо доставки, дней)",
        min_value=3,
        max_value=60,
        value=14,
    )
    celevoy_zapas = st.slider(
        "Целевой запас склада (дней покрытия)",
        min_value=14,
        max_value=90,
        value=45,
    )
    strahovoy_zapas_dni = st.slider(
        "Страховой запас (дней покрытия)",
        min_value=0,
        max_value=30,
        value=10,
    )

# Поиск файлов на сервере или в локальной папке data/
fayl_dvizheniya_po_umolchaniyu = (
    "Остатки и движение Курьяново 01.11.25-06.11.26.xlsx"
    if os.path.exists("Остатки и движение Курьяново 01.11.25-06.11.26.xlsx")
    else "data/Остатки и движение Курьяново 01.11.25-06.11.26.xlsx"
)

fayl_dostupnosti_po_umolchaniyu = (
    "Остатки и доступность 06.11.26.xlsx"
    if os.path.exists("Остатки и доступность 06.11.26.xlsx")
    else "data/Остатки и доступность 06.11.26.xlsx"
)

try:
    if zagruzhennaya_vedomost is not None:
        tablica, dney_v_periode = razobrat_vedomost_dvizheniya(
            zagruzhennaya_vedomost
        )
    else:
        tablica, dney_v_periode = razobrat_vedomost_dvizheniya(
            fayl_dvizheniya_po_umolchaniyu
        )

    if zagruzhennaya_dostupnost is not None:
        tablica = obogatit_dannymi_dostupnosti(
            tablica, zagruzhennaya_dostupnost
        )
    elif os.path.exists(fayl_dostupnosti_po_umolchaniyu):
        tablica = obogatit_dannymi_dostupnosti(
            tablica, fayl_dostupnosti_po_umolchaniyu
        )
except Exception as oshibka:
    st.error(f"Ошибка при обработке файлов: {oshibka}")
    st.stop()

# --- 4. РАСЧЕТ ОБОРАЧИВАЕМОСТИ И ПОТРЕБНОСТИ В ЗАКУПКЕ ---
# Фактически доступный объем: физический остаток + груз в пути
tablica["Всего доступно (бут)"] = (
    tablica["Остаток на складе (бут)"] + tablica["Товары в пути (бут)"]
)

# Обеспеченность склада текущим остатком в днях
tablica["Обеспеченность (дней)"] = tablica.apply(
    lambda r: round(
        r["Всего доступно (бут)"] / r["Среднесуточный расход (бут/день)"], 1
    )
    if r["Среднесуточный расход (бут/день)"] > 0
    else 999.0,
    axis=1,
)

# Точка перезаказа = Среднесуточный расход * (Срок поставки + Страховой запас в днях)
tablica["Точка перезаказа (бут)"] = (
    tablica["Среднесуточный расход (бут/день)"]
    * (srok_dostavki + strahovoy_zapas_dni)
).round(0)

# Нормативный целевой запас = Среднесуточный расход * (Срок поставки + Целевой запас)
tablica["Целевой запас (бут)"] = (
    tablica["Среднесуточный расход (бут/день)"]
    * (srok_dostavki + celevoy_zapas)
).round(0)


# Расчет объема закупки с округлением вверх до целых коробов
def rasschitat_zakaz(stroka):
    chistaya_potrebnost = (
        stroka["Целевой запас (бут)"]
        + stroka["Потребность по заказам (бут)"]
    ) - stroka["Всего доступно (бут)"]

    if chistaya_potrebnost <= 0:
        return 0

    kratnost = stroka["Кратность (бут)"]
    return int(math.ceil(chistaya_potrebnost / kratnost) * kratnost)


tablica["К заказу (бут)"] = tablica.apply(rasschitat_zakaz, axis=1)
tablica["К заказу (кор)"] = (
    tablica["К заказу (бут)"] / tablica["Кратность (бут)"]
).astype(int)
tablica["Бюджет закупки (руб)"] = (
    tablica["К заказу (бут)"] * tablica["Себестоимость (руб)"]
).round(2)


# Классификация состояния товарного остатка
def opredelit_status(stroka):
    if stroka["Потребность по заказам (бут)"] > 0:
        return "🚨 Срочный дефицит"
    if stroka["Всего доступно (бут)"] <= stroka["Точка перезаказа (бут)"]:
        return "⚠️ Пора заказывать"
    if stroka["Обеспеченность (дней)"] > (celevoy_zapas * 2):
        return "📦 Избыточный запас"
    return "✅ В норме"


tablica["Статус запаса"] = tablica.apply(opredelit_status, axis=1)

# Фильтрация по производителю
spisok_postavschikov = ["Все поставщики"] + sorted(
    tablica["Поставщик"].unique().tolist()
)
vybrannyy_postavschik = st.sidebar.selectbox(
    "Фильтр по производителю", spisok_postavschikov
)

otfiltrovannaya_tablica = (
    tablica
    if vybrannyy_postavschik == "Все поставщики"
    else tablica[tablica["Поставщик"] == vybrannyy_postavschik]
)

# --- 5. КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ СКЛАДА ---
k1, k2, k3, k4, k5 = st.columns(5)
k1.metric(
    "Остаток на складе (бут)",
    f"{otfiltrovannaya_tablica['Остаток на складе (бут)'].sum():,}".replace(
        ",", " "
    ),
)
k2.metric(
    "Товары в пути (бут)",
    f"{otfiltrovannaya_tablica['Товары в пути (бут)'].sum():,}".replace(
        ",", " "
    ),
)
k3.metric(
    "Расход за период (бут)",
    f"{otfiltrovannaya_tablica['Расход за период (бут)'].sum():,}".replace(
        ",", " "
    ),
)
k4.metric(
    "Потребность к заказу (бут)",
    f"{otfiltrovannaya_tablica['К заказу (бут)'].sum():,}".replace(",", " "),
)
k5.metric(
    "Бюджет закупки (руб)",
    f"{otfiltrovannaya_tablica['Бюджет закупки (руб)'].sum():,.0f}".replace(
        ",", " "
    )
    + " ₽",
)

st.markdown("---")

# --- 6. ТАБЛИЦА ФОРМИРОВАНИЯ ЗАКУПОК И ЭКСПОРТ ---
zagolovok_stolbec, pereklyuchatel_stolbec = st.columns([3, 1])
with zagolovok_stolbec:
    st.subheader("Сформированный план пополнения склада")
with pereklyuchatel_stolbec:
    tolko_k_zakazu = st.toggle("Только позиции к заказу", value=False)

tablica_dlya_vyvoda = (
    otfiltrovannaya_tablica[otfiltrovannaya_tablica["К заказу (бут)"] > 0]
    if tolko_k_zakazu
    else otfiltrovannaya_tablica
)

vyvodimye_stolbcy = [
    "Поставщик",
    "Наименование",
    "Статус запаса",
    "Себестоимость (руб)",
    "Среднесуточный расход (бут/день)",
    "Остаток на складе (бут)",
    "Товары в пути (бут)",
    "Обеспеченность (дней)",
    "Кратность (бут)",
    "К заказу (бут)",
    "К заказу (кор)",
    "Бюджет закупки (руб)",
]

st.dataframe(
    tablica_dlya_vyvoda[vyvodimye_stolbcy].style.format(
        {
            "Себестоимость (руб)": "{:.2f} ₽",
            "Среднесуточный расход (бут/день)": "{:.2f}",
            "Бюджет закупки (руб)": "{:,.2f} ₽",
        }
    ),
    use_container_width=True,
    height=480,
)

# Экспорт в Excel
baffer_eksporta = io.BytesIO()
with pd.ExcelWriter(baffer_eksporta, engine="openpyxl") as pisatel:
    tablica_dlya_vyvoda[vyvodimye_stolbcy].to_excel(
        pisatel, index=False, sheet_name="План закупок"
    )

st.download_button(
    label="📥 Экспортировать план закупок в Excel",
    data=baffer_eksporta.getvalue(),
    file_name="План_закупок_Курьяново.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
