import io
import math
import os
import re
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Управление запасами и планирование продаж",
    layout="wide",
    initial_sidebar_state="expanded",
)


# --- 1. ПАРСЕР ВЕДОМОСТИ ДВИЖЕНИЯ ТОВАРОВ ---
def razobrat_vedomost_dvizheniya(fayl_bayty):
    tablica_syraya = pd.read_excel(fayl_bayty, sheet_name=0)

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

        if (
            pd.isna(nomenklatura_tekst)
            or str(nomenklatura_tekst).strip() == ""
        ):
            tekuschiy_postavschik = artikul_tekst
            continue

        naimenovanie = str(nomenklatura_tekst).strip()

        poisk_koroba = re.search(r":\s*(\d+)\s*,?", naimenovanie)
        kratnost_koroba = int(poisk_koroba.group(1)) if poisk_koroba else 6

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
                "План продаж на месяц (бут)": 0,
                "Тип позиции": "Действующая матрица",
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


# --- 3. ГЕНЕРАТОР ШАБЛОНА ПЛАНА ПРОДАЖ ДЛЯ МЕНЕДЖЕРОВ ---
def sozdat_shablon_plana(df_baza):
    shablon_spisok = []

    # 1. Заполняем текущую матрицу
    for _, r in df_baza.iterrows():
        shablon_spisok.append(
            {
                "Поставщик / Производитель": r["Поставщик"],
                "Артикул": r["Артикул"],
                "Наименование продукции": r["Наименование"],
                "Кратность короба (бут)": r["Кратность (бут)"],
                "Ориентировочная себестоимость (руб)": r["Себестоимость (руб)"],
                "План продаж на месяц (бут)": 0,
                "Комментарий менеджера (клиенты/сеть)": "",
            }
        )

    # 2. Добавляем примеры строк для новых поставщиков и новинок
    shablon_spisok.append(
        {
            "Поставщик / Производитель": "НОВЫЙ ПОСТАВЩИК (ПРИМЕР)",
            "Артикул": "НОВИНКА-001",
            "Наименование продукции": "Новое вино сухое белое 0.75: 6",
            "Кратность короба (бут)": 6,
            "Ориентировочная себестоимость (руб)": 350.0,
            "План продаж на месяц (бут)": 600,
            "Комментарий менеджера (клиенты/сеть)": "Ввод под новую локальную сеть",
        }
    )

    df_shablon = pd.DataFrame(shablon_spisok)
    baffer = io.BytesIO()
    with pd.ExcelWriter(baffer, engine="openpyxl") as pisatel:
        df_shablon.to_excel(pisatel, index=False, sheet_name="План продаж")
    return baffer.getvalue()


# --- 4. ОБЪЕДИНЕНИЕ БАЗЫ С ПЛАНОМ ПРОДАЖ МЕНЕДЖЕРОВ ---
def primenit_plan_prodazh(osnovnaya_tablica, fayl_plana):
    df_plan = pd.read_excel(fayl_plana)

    # Поиск нужных колонок вне зависимости от мелких различий в регистре
    kolonki = {str(c).strip().lower(): c for c in df_plan.columns}

    k_postavschik = next(
        (v for k, v in kolonki.items() if "поставщик" in k or "производ" in k),
        None,
    )
    k_artikul = next((v for k, v in kolonki.items() if "артикул" in k), None)
    k_naimenovanie = next(
        (v for k, v in kolonki.items() if "наименов" in k or "продукт" in k),
        None,
    )
    k_plan = next(
        (v for k, v in kolonki.items() if "план" in k or "продаж" in k), None
    )
    k_sebest = next(
        (v for k, v in kolonki.items() if "себестоим" in k or "цена" in k), None
    )
    k_kratnost = next((v for k, v in kolonki.items() if "кратн" in k), None)

    if not k_plan:
        return osnovnaya_tablica

    plany_po_artikulam = {}
    novye_pozicii = []
    suschestvuyuschie_artikuly = set(osnovnaya_tablica["Артикул"].astype(str))

    for _, stroka in df_plan.iterrows():
        artikul_str = str(stroka[k_artikul]).strip() if k_artikul else ""
        plan_obem = (
            float(stroka[k_plan])
            if pd.notna(stroka[k_plan]) and str(stroka[k_plan]).strip() != ""
            else 0.0
        )

        if artikul_str in suschestvuyuschie_artikuly and artikul_str != "":
            plany_po_artikulam[artikul_str] = int(plan_obem)
        else:
            # Это новая позиция или новый поставщик
            naim = (
                str(stroka[k_naimenovanie]).strip()
                if k_naimenovanie and pd.notna(stroka[k_naimenovanie])
                else "Новинка без названия"
            )
            if "ПРИМЕР" in naim.upper() and plan_obem == 0:
                continue
            if plan_obem > 0:
                post = (
                    str(stroka[k_postavschik]).strip()
                    if k_postavschik and pd.notna(stroka[k_postavschik])
                    else "Новый поставщик"
                )
                kratn = (
                    int(stroka[k_kratnost])
                    if k_kratnost and pd.notna(stroka[k_kratnost])
                    else 6
                )
                seb = (
                    float(stroka[k_sebest])
                    if k_sebest and pd.notna(stroka[k_sebest])
                    else 300.0
                )

                novye_pozicii.append(
                    {
                        "Поставщик": post,
                        "Артикул": artikul_str
                        if artikul_str and artikul_str != "nan"
                        else f"НОВ-{len(novye_pozicii)+1}",
                        "Наименование": naim,
                        "Кратность (бут)": kratn,
                        "Себестоимость (руб)": seb,
                        "Приход за период (бут)": 0,
                        "Расход за период (бут)": 0,
                        "Остаток на складе (бут)": 0,
                        "Товары в пути (бут)": 0,
                        "Потребность по заказам (бут)": 0,
                        "Среднесуточный расход (бут/день)": round(
                            plan_obem / 30.0, 3
                        ),
                        "План продаж на месяц (бут)": int(plan_obem),
                        "Тип позиции": "✨ Новинка / Новый поставщик",
                    }
                )

    # Обновляем план в основной таблице
    osnovnaya_tablica["План продаж на месяц (бут)"] = (
        osnovnaya_tablica["Артикул"]
        .map(plany_po_artikulam)
        .fillna(0)
        .astype(int)
    )

    # Если есть новинки — прикрепляем их
    if novye_pozicii:
        df_novinki = pd.DataFrame(novye_pozicii)
        osnovnaya_tablica = pd.concat(
            [osnovnaya_tablica, df_novinki], ignore_index=True
        )

    return osnovnaya_tablica


# --- 5. ИНТЕРФЕЙС И НАСТРОЙКИ ---
st.title("🍷 Дашборд управления запасами и планирования продаж")
st.caption(
    "Склад Курьяново | Учет планов менеджеров, новых клиентов и паллетных норм"
)

with st.sidebar:
    st.header("📂 Загрузка отчетов 1С")
    zagruzhennaya_vedomost = st.file_uploader(
        "1. Ведомость по товарам (движение и цены)", type=["xlsx"]
    )
    zagruzhennaya_dostupnost = st.file_uploader(
        "2. Остатки и доступность (транзит и резервы)", type=["xlsx"]
    )

    st.header("📈 План продаж менеджеров")
    zagruzhennyy_plan = st.file_uploader(
        "3. Загрузить план продаж на месяц", type=["xlsx"]
    )

    metod_prognoza = st.radio(
        "База для расчета среднесуточного расхода:",
        options=[
            "Только факт отгрузок 1С",
            "План менеджеров (План / 30 дн)",
            "Максимум (Факт vs План)",
        ],
        index=2,
    )

    st.header("🔍 Поиск и фильтрация")
    poisk_tekst = st.text_input("Поиск по названию или артикулу", "")

    st.header("📦 Кванты отгрузки поставщика")
    emkost_palleta = st.number_input(
        "Вместимость 1 паллета (бут)", min_value=60, max_value=1200, value=600, step=60
    )
    okruglyat_do_palleta = st.checkbox(
        "Округлять заказ до целых паллет", value=False
    )
    minimalnyy_zakaz_pallet = st.number_input(
        "Минимальный заказ (в паллетах)", min_value=1, max_value=33, value=1
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

# Автозагрузка исходных файлов
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
    elif os.path.exists(fayl_dvizheniya_po_umolchaniyu):
        tablica, dney_v_periode = razobrat_vedomost_dvizheniya(
            fayl_dvizheniya_po_umolchaniyu
        )
    else:
        st.warning("Загрузите ведомость движения товаров в боковом меню.")
        st.stop()

    if zagruzhennaya_dostupnost is not None:
        tablica = obogatit_dannymi_dostupnosti(
            tablica, zagruzhennaya_dostupnost
        )
    elif os.path.exists(fayl_dostupnosti_po_umolchaniyu):
        tablica = obogatit_dannymi_dostupnosti(
            tablica, fayl_dostupnosti_po_umolchaniyu
        )

    # Применение плана продаж, если он загружен
    if zagruzhennyy_plan is not None:
        tablica = primenit_plan_prodazh(tablica, zagruzhennyy_plan)

except Exception as oshibka:
    st.error(f"Ошибка при обработке файлов: {oshibka}")
    st.stop()

# --- 6. ПЕРЕСЧЕТ СРЕДНЕСУТОЧНОГО СПРОСА С УЧЕТОМ ПЛАНА ---
def opredelit_itogovyy_rashod(stroka):
    fakt_sutki = stroka["Среднесуточный расход (бут/день)"]
    plan_sutki = (
        (stroka["План продаж на месяц (бут)"] / 30.0)
        if stroka["План продаж на месяц (бут)"] > 0
        else 0.0
    )

    if metod_prognoza == "План менеджеров (План / 30 дн)":
        return plan_sutki if plan_sutki > 0 else fakt_sutki
    elif metod_prognoza == "Максимум (Факт vs План)":
        return max(fakt_sutki, plan_sutki)
    return fakt_sutki


tablica["Расчетный спрос в день (бут)"] = tablica.apply(
    opredelit_itogovyy_rashod, axis=1
).round(3)

# --- 7. РАСЧЕТ ОБОРАЧИВАЕМОСТИ И ПОТРЕБНОСТИ В ЗАКУПКЕ ---
tablica["Всего доступно (бут)"] = (
    tablica["Остаток на складе (бут)"] + tablica["Товары в пути (бут)"]
)

# Обеспеченность склада (дней)
tablica["Обеспеченность (дней)"] = tablica.apply(
    lambda r: round(
        r["Всего доступно (бут)"] / r["Расчетный спрос в день (бут)"], 1
    )
    if r["Расчетный спрос в день (бут)"] > 0
    else 999.0,
    axis=1,
)

# Точка перезаказа и нормативный целевой запас
tablica["Точка перезаказа (бут)"] = (
    tablica["Расчетный спрос в день (бут)"]
    * (srok_dostavki + strahovoy_zapas_dni)
).round(0)

tablica["Целевой запас (бут)"] = (
    tablica["Расчетный спрос в день (бут)"]
    * (srok_dostavki + celevoy_zapas)
).round(0)


# Расчет объема закупки с учетом квантов (короба, паллеты, мин. заказ)
def rasschitat_zakaz_kvanty(stroka):
    chistaya_potrebnost = (
        stroka["Целевой запас (бут)"]
        + stroka["Потребность по заказам (бут)"]
    ) - stroka["Всего доступно (бут)"]

    if chistaya_potrebnost <= 0:
        return 0

    kratnost_korob = stroka["Кратность (бут)"]
    zakaz_butylki = int(math.ceil(chistaya_potrebnost / kratnost_korob) * kratnost_korob)

    if okruglyat_do_palleta and zakaz_butylki > 0:
        zakaz_butylki = int(math.ceil(zakaz_butylki / emkost_palleta) * emkost_palleta)
    elif zakaz_butylki > 0 and minimalnyy_zakaz_pallet > 0:
        minimalnyy_obem = int(minimalnyy_zakaz_pallet * emkost_palleta)
        if zakaz_butylki < minimalnyy_obem:
            zakaz_butylki = minimalnyy_obem

    return zakaz_butylki


tablica["К заказу (бут)"] = tablica.apply(rasschitat_zakaz_kvanty, axis=1)
tablica["К заказу (кор)"] = (
    tablica["К заказу (бут)"] / tablica["Кратность (бут)"]
).astype(int)
tablica["К заказу (паллет)"] = (
    tablica["К заказу (бут)"] / emkost_palleta
).round(2)
tablica["Бюджет закупки (руб)"] = (
    tablica["К заказу (бут)"] * tablica["Себестоимость (руб)"]
).round(2)


# Расчет избыточного запаса (профицита)
def rasschitat_izbytok(stroka):
    izlishek = stroka["Всего доступно (бут)"] - stroka["Целевой запас (бут)"]
    return int(max(0, izlishek))


tablica["Избыток (бут)"] = tablica.apply(rasschitat_izbytok, axis=1)
tablica["Заморожено в избытке (руб)"] = (
    tablica["Избыток (бут)"] * tablica["Себестоимость (руб)"]
).round(2)


# Классификация состояния остатка
def opredelit_status(stroka):
    if stroka["Тип позиции"] == "✨ Новинка / Новый поставщик":
        return "✨ Ввод новинки"
    if stroka["Потребность по заказам (бут)"] > 0:
        return "🚨 Срочный дефицит"
    if stroka["Всего доступно (бут)"] <= stroka["Точка перезаказа (бут)"]:
        return "⚠️ Пора заказывать"
    if (
        stroka["Обеспеченность (дней)"] > (celevoy_zapas * 1.5)
        and stroka["Избыток (бут)"] > 0
    ):
        return "📦 Избыточный запас"
    return "✅ В норме"


tablica["Статус запаса"] = tablica.apply(opredelit_status, axis=1)

# Фильтр по поставщику
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

# Поиск по ключевому слову
if poisk_tekst:
    otfiltrovannaya_tablica = otfiltrovannaya_tablica[
        otfiltrovannaya_tablica["Наименование"].str.contains(
            poisk_tekst, case=False, na=False
        )
        | otfiltrovannaya_tablica["Артикул"].str.contains(
            poisk_tekst, case=False, na=False
        )
    ]

# --- 8. БЛОК ВЫГРУЗКИ ШАБЛОНА ДЛЯ МЕНЕДЖЕРОВ ---
with st.expander("📥 Скачать форму-шаблон для заполнения плана менеджерами", expanded=False):
    st.write(
        "Файл содержит полную матрицу действующих позиций склада, а также строки-примеры для внесения "
        "новых виноделен, линеек и прогнозов продаж под расширение клиентской базы."
    )
    dannie_shablona = sozdat_shablon_plana(tablica[tablica["Тип позиции"] == "Действующая матрица"])
    st.download_button(
        label="📥 Скачать Excel-шаблон плана продаж",
        data=dannie_shablona,
        file_name="Шаблон_плана_продаж_на_месяц.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# --- 9. КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ (KPI) ---
m1, m2, m3, m4, m5 = st.columns(5)
m1.metric(
    "Остаток на складе (бут)",
    f"{otfiltrovannaya_tablica['Остаток на складе (бут)'].sum():,}".replace(
        ",", " "
    ),
)
m2.metric(
    "Товары в пути (бут)",
    f"{otfiltrovannaya_tablica['Товары в пути (бут)'].sum():,}".replace(
        ",", " "
    ),
)
m3.metric(
    "Потребность к заказу (бут / пал)",
    f"{otfiltrovannaya_tablica['К заказу (бут)'].sum():,} бут "
    f"({otfiltrovannaya_tablica['К заказу (бут)'].sum() / emkost_palleta:.1f} пал)".replace(
        ",", " "
    ),
)
m4.metric(
    "Бюджет закупки (руб)",
    f"{otfiltrovannaya_tablica['Бюджет закупки (руб)'].sum():,.0f}".replace(
        ",", " "
    )
    + " ₽",
)
m5.metric(
    "Заморожено в профиците (руб)",
    f"{otfiltrovannaya_tablica['Заморожено в избытке (руб)'].sum():,.0f}".replace(
        ",", " "
    )
    + " ₽",
)

st.markdown("---")

# --- 10. ПЕРЕКЛЮЧАТЕЛЬ РЕЖИМА И СОРТИРОВКА ТАБЛИЦЫ ---
kolonka_perekl, kolonka_sort = st.columns([2, 2])

with kolonka_perekl:
    rezhim_otobrazheniya = st.radio(
        "Режим просмотра выборки:",
        options=[
            "Все позиции",
            "⚠️ Только дефицит и закупка",
            "📦 Только избыточный запас (профицит)",
            "✨ Новинки и новые поставщики",
        ],
        horizontal=True,
    )

with kolonka_sort:
    if rezhim_otobrazheniya == "📦 Только избыточный запас (профицит)":
        sortirovka_prozicit = st.selectbox(
            "Сортировка профицита:",
            [
                "По убыванию замороженных денег (руб)",
                "По убыванию излишка (бут)",
                "По убыванию обеспеченности (дней)",
            ],
        )
    else:
        sortirovka_prozicit = None

# Применение выбранного режима и сортировки
if rezhim_otobrazheniya == "⚠️ Только дефицит и закупка":
    tablica_vyvoda = otfiltrovannaya_tablica[
        otfiltrovannaya_tablica["К заказу (бут)"] > 0
    ].sort_values(by="К заказу (бут)", ascending=False)
elif rezhim_otobrazheniya == "📦 Только избыточный запас (профицит)":
    tablica_vyvoda = otfiltrovannaya_tablica[
        (otfiltrovannaya_tablica["Статус запаса"] == "📦 Избыточный запас")
        | (otfiltrovannaya_tablica["Избыток (бут)"] > 0)
    ]
    if sortirovka_prozicit == "По убыванию замороженных денег (руб)":
        tablica_vyvoda = tablica_vyvoda.sort_values(
            by="Заморожено в избытке (руб)", ascending=False
        )
    elif sortirovka_prozicit == "По убыванию излишка (бут)":
        tablica_vyvoda = tablica_vyvoda.sort_values(
            by="Избыток (бут)", ascending=False
        )
    else:
        tablica_vyvoda = tablica_vyvoda.sort_values(
            by="Обеспеченность (дней)", ascending=False
        )
elif rezhim_otobrazheniya == "✨ Новинки и новые поставщики":
    tablica_vyvoda = otfiltrovannaya_tablica[
        otfiltrovannaya_tablica["Тип позиции"] == "✨ Новинка / Новый поставщик"
    ]
else:
    tablica_vyvoda = otfiltrovannaya_tablica

# Информационный бейдж
kol_vo_novinok = (tablica["Тип позиции"] == "✨ Новинка / Новый поставщик").sum()
st.info(
    f"Отображается строк: **{len(tablica_vyvoda)}** из {len(otfiltrovannaya_tablica)} | "
    f"Позиций к заказу: **{(otfiltrovannaya_tablica['К заказу (бут)'] > 0).sum()}** | "
    f"Позиций в профиците: **{(otfiltrovannaya_tablica['Избыток (бут)'] > 0).sum()}** | "
    f"Загружено новинок: **{kol_vo_novinok}**"
)

# Список колонок для вывода
stolbcy_dlya_pokaza = [
    "Поставщик",
    "Наименование",
    "Статус запаса",
    "Себестоимость (руб)",
    "План продаж на месяц (бут)",
    "Расчетный спрос в день (бут)",
    "Остаток на складе (бут)",
    "Товары в пути (бут)",
    "Обеспеченность (дней)",
    "Кратность (бут)",
    "К заказу (бут)",
    "К заказу (кор)",
    "К заказу (паллет)",
    "Бюджет закупки (руб)",
    "Избыток (бут)",
    "Заморожено в избытке (руб)",
]

st.dataframe(
    tablica_vyvoda[stolbcy_dlya_pokaza].style.format(
        {
            "Себестоимость (руб)": "{:.2f} ₽",
            "Расчетный спрос в день (бут)": "{:.2f}",
            "К заказу (паллет)": "{:.2f}",
            "Бюджет закупки (руб)": "{:,.2f} ₽",
            "Заморожено в избытке (руб)": "{:,.2f} ₽",
        }
    ),
    use_container_width=True,
)

# --- 11. ВЫГРУЗКА В EXCEL ---
baffer_eksporta = io.BytesIO()
with pd.ExcelWriter(baffer_eksporta, engine="openpyxl") as pisatel:
    tablica_vyvoda[stolbcy_dlya_pokaza].to_excel(
        pisatel, index=False, sheet_name="План закупок"
    )

st.download_button(
    label="📥 Экспортировать план закупок в Excel",
    data=baffer_eksporta.getvalue(),
    file_name="План_закупок_с_учетом_плана_продаж.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
