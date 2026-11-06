import io
import math
import os
import re
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Управление и планирование запасов",
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


# --- 3. ГЕНЕРАТОР ФОРМЫ ПЛАНА ДЛЯ МЕНЕДЖЕРОВ (СТРОГО 6 ПОЛЕЙ) ---
def sozdat_formu_plana_menedzheram(df_baza):
    stroki = []

    for _, r in df_baza.iterrows():
        stroki.append(
            {
                "Поставщик / Производитель": r["Поставщик"],
                "Наименование продукции": r["Наименование"],
                "Текущий остаток (бут)": int(r["Остаток на складе (бут)"]),
                "Себестоимость (руб)": float(r["Себестоимость (руб)"]),
                "План продаж (бут)": 0,
                "План продаж (руб)": 0.0,
            }
        )

    stroki.append(
        {
            "Поставщик / Производитель": "НОВЫЙ ПОСТАВЩИК (ПРИМЕР)",
            "Наименование продукции": "Новое вино сухое белое 0.75: 6",
            "Текущий остаток (бут)": 0,
            "Себестоимость (руб)": 350.0,
            "План продаж (бут)": 600,
            "План продаж (руб)": 210000.0,
        }
    )

    df_itog = pd.DataFrame(stroki)
    baffer = io.BytesIO()

    with pd.ExcelWriter(baffer, engine="openpyxl") as pisatel:
        df_itog.to_excel(pisatel, index=False, sheet_name="План продаж")
        ws = pisatel.sheets["План продаж"]

        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.5
        ws.page_margins.right = 0.5
        ws.page_margins.top = 0.6
        ws.page_margins.bottom = 0.6

        granica_tonkaya = Border(
            left=Side(style="thin", color="CCCCCC"),
            right=Side(style="thin", color="CCCCCC"),
            top=Side(style="thin", color="CCCCCC"),
            bottom=Side(style="thin", color="CCCCCC"),
        )
        zalivka_shapki = PatternFill(
            start_color="3B4252", end_color="3B4252", fill_type="solid"
        )
        shrift_shapki = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        for col_idx in range(1, 7):
            c = ws.cell(row=1, column=col_idx)
            c.fill = zalivka_shapki
            c.font = shrift_shapki
            c.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

        ws.row_dimensions[1].height = 28

        zalivka_plana = PatternFill(
            start_color="FFFBEB", end_color="FFFBEB", fill_type="solid"
        )

        for row_idx in range(2, len(df_itog) + 2):
            ws.row_dimensions[row_idx].height = 20

            cA = ws.cell(row=row_idx, column=1)
            cA.border = granica_tonkaya
            cA.alignment = Alignment(horizontal="left", vertical="center")

            cB = ws.cell(row=row_idx, column=2)
            cB.border = granica_tonkaya
            cB.alignment = Alignment(horizontal="left", vertical="center")

            cC = ws.cell(row=row_idx, column=3)
            cC.border = granica_tonkaya
            cC.number_format = "#,##0"
            cC.alignment = Alignment(horizontal="right", vertical="center")

            cD = ws.cell(row=row_idx, column=4)
            cD.border = granica_tonkaya
            cD.number_format = "#,##0.00 ₽"
            cD.alignment = Alignment(horizontal="right", vertical="center")

            cE = ws.cell(row=row_idx, column=5)
            cE.border = granica_tonkaya
            cE.fill = zalivka_plana
            cE.number_format = "#,##0"
            cE.font = Font(name="Calibri", size=11, bold=True)
            cE.alignment = Alignment(horizontal="right", vertical="center")

            cF = ws.cell(row=row_idx, column=6)
            cF.value = f"=D{row_idx}*E{row_idx}"
            cF.border = granica_tonkaya
            cF.number_format = "#,##0.00 ₽"
            cF.alignment = Alignment(horizontal="right", vertical="center")

        ws.column_dimensions["A"].width = 28
        ws.column_dimensions["B"].width = 54
        ws.column_dimensions["C"].width = 16
        ws.column_dimensions["D"].width = 18
        ws.column_dimensions["E"].width = 18
        ws.column_dimensions["F"].width = 20

        ws.freeze_panes = "C2"

    return baffer.getvalue()


# --- 4. ПРИЕМ ПЛАНА ПРОДАЖ МЕНЕДЖЕРОВ ---
def primenit_plan_prodazh(osnovnaya_tablica, fayl_plana):
    df_plan = pd.read_excel(fayl_plana)
    kolonki = {str(c).strip().lower(): c for c in df_plan.columns}

    k_postavschik = next(
        (v for k, v in kolonki.items() if "поставщик" in k or "производ" in k),
        None,
    )
    k_naimenovanie = next(
        (v for k, v in kolonki.items() if "наименов" in k or "продукт" in k),
        None,
    )
    k_plan = next(
        (v for k, v in kolonki.items() if "план" in k and "бут" in k or "план" in k),
        None,
    )
    k_sebest = next(
        (v for k, v in kolonki.items() if "себестоим" in k or "цена" in k), None
    )

    if not k_plan or not k_naimenovanie:
        return osnovnaya_tablica

    plany_po_imenam = {}
    novye_pozicii = []
    suschestvuyuschie_naimenovaniya = set(
        osnovnaya_tablica["Наименование"].astype(str).str.strip().str.lower()
    )

    for _, stroka in df_plan.iterrows():
        naim_str = (
            str(stroka[k_naimenovanie]).strip()
            if pd.notna(stroka[k_naimenovanie])
            else ""
        )
        if not naim_str or naim_str.lower() == "nan":
            continue

        try:
            plan_obem = float(stroka[k_plan]) if pd.notna(stroka[k_plan]) else 0.0
        except Exception:
            plan_obem = 0.0

        if naim_str.lower() in suschestvuyuschie_naimenovaniya:
            plany_po_imenam[naim_str.lower()] = int(plan_obem)
        else:
            if "ПРИМЕР" in naim_str.upper() and plan_obem == 0:
                continue
            if plan_obem > 0:
                post = (
                    str(stroka[k_postavschik]).strip()
                    if k_postavschik and pd.notna(stroka[k_postavschik])
                    else "Новый поставщик"
                )
                try:
                    seb = (
                        float(stroka[k_sebest])
                        if k_sebest and pd.notna(stroka[k_sebest])
                        else 350.0
                    )
                except Exception:
                    seb = 350.0

                poisk_kratn = re.search(r":\s*(\d+)\s*,?", naim_str)
                kratn = int(poisk_kratn.group(1)) if poisk_kratn else 6

                novye_pozicii.append(
                    {
                        "Поставщик": post,
                        "Артикул": f"НОВ-{len(novye_pozicii)+1}",
                        "Наименование": naim_str,
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

    osnovnaya_tablica["План продаж на месяц (бут)"] = osnovnaya_tablica[
        "Наименование"
    ].apply(lambda x: plany_po_imenam.get(str(x).strip().lower(), 0))

    if novye_pozicii:
        df_novinki = pd.DataFrame(novye_pozicii)
        osnovnaya_tablica = pd.concat(
            [osnovnaya_tablica, df_novinki], ignore_index=True
        )

    return osnovnaya_tablica


# --- 5. ФОРМАТИРОВАНИЕ ИТОГОВОГО EXCEL ПЛАНА ЗАКУПОК ---
def sformirovat_krasivyy_excel(tablica_vyvoda, stolbcy_dlya_pokaza):
    baffer = io.BytesIO()
    with pd.ExcelWriter(baffer, engine="openpyxl") as pisatel:
        tablica_vyvoda[stolbcy_dlya_pokaza].to_excel(
            pisatel, index=False, sheet_name="План закупок"
        )
        ws = pisatel.sheets["План закупок"]

        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.paperSize = ws.PAPERSIZE_A4
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = 0.5
        ws.page_margins.right = 0.5
        ws.page_margins.top = 0.6
        ws.page_margins.bottom = 0.6

        granica_tonkaya = Border(
            left=Side(style="thin", color="D3D3D3"),
            right=Side(style="thin", color="D3D3D3"),
            top=Side(style="thin", color="D3D3D3"),
            bottom=Side(style="thin", color="D3D3D3"),
        )
        granica_shapki = Border(
            left=Side(style="thin", color="2C1820"),
            right=Side(style="thin", color="2C1820"),
            top=Side(style="medium", color="2C1820"),
            bottom=Side(style="medium", color="2C1820"),
        )

        zalivka_shapki = PatternFill(
            start_color="581825", end_color="581825", fill_type="solid"
        )
        shrift_shapki = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

        ws.row_dimensions[1].height = 32

        for col_idx in range(1, len(stolbcy_dlya_pokaza) + 1):
            c = ws.cell(row=1, column=col_idx)
            c.fill = zalivka_shapki
            c.font = shrift_shapki
            c.border = granica_shapki
            c.alignment = Alignment(
                horizontal="center", vertical="center", wrap_text=True
            )

        podsvetka_deficit = PatternFill(
            start_color="FDE8E8", end_color="FDE8E8", fill_type="solid"
        )
        podsvetka_zakaz = PatternFill(
            start_color="FEF08A", end_color="FEF08A", fill_type="solid"
        )
        podsvetka_izbytok = PatternFill(
            start_color="E0E7FF", end_color="E0E7FF", fill_type="solid"
        )
        podsvetka_novinka = PatternFill(
            start_color="DCFCE7", end_color="DCFCE7", fill_type="solid"
        )

        for row_idx in range(2, len(tablica_vyvoda) + 2):
            ws.row_dimensions[row_idx].height = 20
            status_znachenie = str(
                ws.cell(row=row_idx, column=3).value or ""
            ).lower()

            for col_idx, nazvanie_kolonki in enumerate(
                stolbcy_dlya_pokaza, start=1
            ):
                cell = ws.cell(row=row_idx, column=col_idx)
                cell.border = granica_tonkaya
                cell.font = Font(name="Calibri", size=10)

                if col_idx == 3:
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )
                    if "дефицит" in status_znachenie:
                        cell.fill = podsvetka_deficit
                        cell.font = Font(
                            name="Calibri", size=10, bold=True, color="991B1B"
                        )
                    elif "заказывать" in status_znachenie:
                        cell.fill = podsvetka_zakaz
                        cell.font = Font(
                            name="Calibri", size=10, bold=True, color="854D0E"
                        )
                    elif "избыточный" in status_znachenie:
                        cell.fill = podsvetka_izbytok
                        cell.font = Font(
                            name="Calibri", size=10, bold=True, color="3730A3"
                        )
                    elif "новинка" in status_znachenie:
                        cell.fill = podsvetka_novinka
                        cell.font = Font(
                            name="Calibri", size=10, bold=True, color="166534"
                        )
                    continue

                if nazvanie_kolonki in [
                    "Себестоимость (руб)",
                    "Бюджет закупки (руб)",
                    "Заморожено в избытке (руб)",
                ]:
                    cell.number_format = "#,##0.00 ₽"
                    cell.alignment = Alignment(
                        horizontal="right", vertical="center"
                    )
                elif nazvanie_kolonki in [
                    "Расчетный спрос в день (бут)",
                    "Среднесуточный расход (бут/день)",
                    "Обеспеченность (дней)",
                    "К заказу (паллет)",
                ]:
                    cell.number_format = "#,##0.0"
                    cell.alignment = Alignment(
                        horizontal="right", vertical="center"
                    )
                elif nazvanie_kolonki in [
                    "План продаж на месяц (бут)",
                    "Остаток на складе (бут)",
                    "Товары в пути (бут)",
                    "Кратность (бут)",
                    "К заказу (бут)",
                    "К заказу (кор)",
                    "Избыток (бут)",
                ]:
                    cell.number_format = "#,##0"
                    cell.alignment = Alignment(
                        horizontal="right", vertical="center"
                    )
                elif nazvanie_kolonki in ["Поставщик", "Наименование"]:
                    cell.alignment = Alignment(
                        horizontal="left", vertical="center"
                    )
                else:
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

        for col in ws.columns:
            dlina = max(len(str(cell.value or "")) for cell in col)
            bukva = get_column_letter(col[0].column)
            if col[0].value == "Наименование":
                ws.column_dimensions[bukva].width = 52
            elif col[0].value == "Поставщик":
                ws.column_dimensions[bukva].width = 28
            else:
                ws.column_dimensions[bukva].width = max(min(dlina + 4, 30), 13)

        ws.freeze_panes = "C2"

    return baffer.getvalue()


# --- 6. ИНТЕРФЕЙС И НАСТРОЙКИ ---
st.title("🍷 Управление и планирование запасов")
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
        "3. Загрузить заполненный план продаж", type=["xlsx"]
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

    if zagruzhennyy_plan is not None:
        tablica = primenit_plan_prodazh(tablica, zagruzhennyy_plan)

except Exception as oshibka:
    st.error(f"Ошибка при обработке файлов: {oshibka}")
    st.stop()

# --- 7. ПЕРЕСЧЕТ СРЕДНЕСУТОЧНОГО СПРОСА С УЧЕТОМ ПЛАНА ---
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

# --- 8. РАСЧЕТ ОБОРАЧИВАЕМОСТИ И ПОТРЕБНОСТИ В ЗАКУПКЕ ---
tablica["Всего доступно (бут)"] = (
    tablica["Остаток на складе (бут)"] + tablica["Товары в пути (бут)"]
)

tablica["Обеспеченность (дней)"] = tablica.apply(
    lambda r: round(
        r["Всего доступно (бут)"] / r["Расчетный спрос в день (бут)"], 1
    )
    if r["Расчетный спрос в день (бут)"] > 0
    else 999.0,
    axis=1,
)

tablica["Точка перезаказа (бут)"] = (
    tablica["Расчетный спрос в день (бут)"]
    * (srok_dostavki + strahovoy_zapas_dni)
).round(0)

tablica["Целевой запас (бут)"] = (
    tablica["Расчетный спрос в день (бут)"]
    * (srok_dostavki + celevoy_zapas)
).round(0)


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


def rasschitat_izbytok(stroka):
    izlishek = stroka["Всего доступно (бут)"] - stroka["Целевой запас (бут)"]
    return int(max(0, izlishek))


tablica["Избыток (бут)"] = tablica.apply(rasschitat_izbytok, axis=1)
tablica["Заморожено в избытке (руб)"] = (
    tablica["Избыток (бут)"] * tablica["Себестоимость (руб)"]
).round(2)


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

# ВЕСОВОЙ РАНГ СТАТУСОВ ДЛЯ ПРИОРИТЕТНОГО ВЫВОДА ДЕФИЦИТА НА ПЕРВЫЕ МЕСТА
rang_prioritetov = {
    "🚨 Срочный дефицит": 1,
    "⚠️ Пора заказывать": 2,
    "✨ Ввод новинки": 3,
    "✅ В норме": 4,
    "📦 Избыточный запас": 5,
}
tablica["Приоритет_ранг"] = tablica["Статус запаса"].map(rang_prioritetov).fillna(99)

# --- МУЛЬТИСЕЛЕКТ ПО ПРОИЗВОДИТЕЛЯМ ---
spisok_postavschikov = sorted(tablica["Поставщик"].unique().tolist())
vybrannye_postavschiki = st.sidebar.multiselect(
    "Фильтр по производителям (можно выбрать несколько)",
    options=spisok_postavschikov,
    default=[],
)

if vybrannye_postavschiki:
    otfiltrovannaya_tablica = tablica[
        tablica["Поставщик"].isin(vybrannye_postavschiki)
    ]
else:
    otfiltrovannaya_tablica = tablica

if poisk_tekst:
    otfiltrovannaya_tablica = otfiltrovannaya_tablica[
        otfiltrovannaya_tablica["Наименование"].str.contains(
            poisk_tekst, case=False, na=False
        )
        | otfiltrovannaya_tablica["Артикул"].str.contains(
            poisk_tekst, case=False, na=False
        )
    ]

# --- 9. БЛОК ВЫГРУЗКИ ФОРМЫ ДЛЯ МЕНЕДЖЕРОВ (6 ПОЛЕЙ) ---
st.markdown("### 📋 Форма плана продаж для коммерческого отдела")
kol_podskazka, kol_knopka = st.columns([3, 1])
with kol_podskazka:
    st.write(
        "Файл содержит строго 6 колонок: **Производитель, Наименование, Текущий остаток, Себестоимость, План продаж (бут) и План продаж (руб)** с зашитой формулой расчета."
    )
with kol_knopka:
    fayl_formy_menedzherov = sozdat_formu_plana_menedzheram(
        tablica[tablica["Тип позиции"] == "Действующая матрица"]
    )
    st.download_button(
        label="📥 Скачать форму плана для менеджеров (Excel)",
        data=fayl_formy_menedzherov,
        file_name="Форма_плана_продаж_менеджерам.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

st.markdown("---")

# --- 10. КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ (KPI) ---
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
    "Заморожено в избытке (руб)",
    f"{otfiltrovannaya_tablica['Заморожено в избытке (руб)'].sum():,.0f}".replace(
        ",", " "
    )
    + " ₽",
)

st.markdown("---")

# --- 11. ПЕРЕКЛЮЧАТЕЛЬ РЕЖИМА И СОРТИРОВКА ТАБЛИЦЫ ---
kolonka_perekl, kolonka_sort = st.columns([2, 2])

with kolonka_perekl:
    rezhim_otobrazheniya = st.radio(
        "Режим просмотра выборки:",
        options=[
            "Все позиции",
            "⚠️ Только дефицит и закупка",
            "📦 Только избыточный запас",
            "✨ Новинки и новые поставщики",
        ],
        horizontal=True,
    )

with kolonka_sort:
    if rezhim_otobrazheniya == "📦 Только избыточный запас":
        sortirovka_izbytok = st.selectbox(
            "Сортировка избытка:",
            [
                "По убыванию замороженных денег (руб)",
                "По убыванию излишка (бут)",
                "По убыванию обеспеченности (дней)",
            ],
        )
    else:
        sortirovka_izbytok = None

if rezhim_otobrazheniya == "⚠️ Только дефицит и закупка":
    tablica_vyvoda = otfiltrovannaya_tablica[
        otfiltrovannaya_tablica["К заказу (бут)"] > 0
    ].sort_values(by=["Приоритет_ранг", "К заказу (бут)"], ascending=[True, False])
elif rezhim_otobrazheniya == "📦 Только избыточный запас":
    tablica_vyvoda = otfiltrovannaya_tablica[
        (otfiltrovannaya_tablica["Статус запаса"] == "📦 Избыточный запас")
        | (otfiltrovannaya_tablica["Избыток (бут)"] > 0)
    ]
    if sortirovka_izbytok == "По убыванию замороженных денег (руб)":
        tablica_vyvoda = tablica_vyvoda.sort_values(
            by="Заморожено в избытке (руб)", ascending=False
        )
    elif sortirovka_izbytok == "По убыванию излишка (бут)":
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
    ].sort_values(by="К заказу (бут)", ascending=False)
else:
    # "Все позиции": Дефицитные и требующие заказа ВСЕГДА первыми, внутри них — по убыванию заказа
    tablica_vyvoda = otfiltrovannaya_tablica.sort_values(
        by=["Приоритет_ранг", "К заказу (бут)"], ascending=[True, False]
    )

kol_vo_novinok = (tablica["Тип позиции"] == "✨ Новинка / Новый поставщик").sum()
st.info(
    f"Отображается строк: **{len(tablica_vyvoda)}** из {len(otfiltrovannaya_tablica)} | "
    f"Позиций к заказу: **{(otfiltrovannaya_tablica['К заказу (бут)'] > 0).sum()}** | "
    f"Позиций в избытке: **{(otfiltrovannaya_tablica['Избыток (бут)'] > 0).sum()}** | "
    f"Загружено новинок: **{kol_vo_novinok}**"
)

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


# Функция подсветки строк по статусам
def podsvetit_stroki_po_statusu(stroka):
    status = str(stroka.get("Статус запаса", ""))
    if "Срочный дефицит" in status:
        # Мягкий тревожный красный фон с темно-красным контрастным текстом
        return ["background-color: #ffebee; color: #b71c1c; font-weight: 600;"] * len(stroka)
    elif "Пора заказывать" in status:
        # Мягкий желто-персиковый фон предупреждения
        return ["background-color: #fffde7; color: #bf360c;"] * len(stroka)
    return [""] * len(stroka)


st_table = (
    tablica_vyvoda[stolbcy_dlya_pokaza]
    .style.apply(podsvetit_stroki_po_statusu, axis=1)
    .format(
        {
            "Себестоимость (руб)": "{:.2f} ₽",
            "Расчетный спрос в день (бут)": "{:.2f}",
            "Остаток на складе (бут)": "{:,.0f}",
            "Товары в пути (бут)": "{:,.0f}",
            "К заказу (бут)": "{:,.0f}",
            "К заказу (кор)": "{:,.0f}",
            "К заказу (паллет)": "{:.2f}",
            "Избыток (бут)": "{:,.0f}",
            "Бюджет закупки (руб)": "{:,.2f} ₽",
            "Заморожено в избытке (руб)": "{:,.2f} ₽",
        }
    )
)

st.dataframe(st_table, use_container_width=True)

# --- 12. ВЫГРУЗКА ИТОГОВОГО ПЛАНА ЗАКУПОК В EXCEL ---
fayl_eksporta_bayty = sformirovat_krasivyy_excel(
    tablica_vyvoda, stolbcy_dlya_pokaza
)

st.download_button(
    label="📥 Экспортировать итоговый план закупок в Excel",
    data=fayl_eksporta_bayty,
    file_name="Итоговый_план_закупок.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
