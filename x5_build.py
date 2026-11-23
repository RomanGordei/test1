"""Сборка куба X5 Dialog для Excel: таблицы из ClickHouse (Parquet) -> проверенные таблицы CSV.

Запуск:  python x5_build.py   (после x5_export.py)
Нужно:   pip install pandas pyarrow python-calamine

Что делает:
  1. Читает x5_weekly.parquet и x5_monthly.parquet, выгруженные из клика скриптом x5_export.py.
  2. Переименовывает колонки ПО СОДЕРЖИМОМУ: в клике 15 колонок названы не тем, что в них лежит.
  3. Проверяет формулы X5 на каждой строке и останавливается, если колонки в клике поменялись.
  4. Делит данные на факты по PLU и итоги категории (строки TOTAL), убирает полные дубли, проверяет баланс.
  5. Собирает справочники и календари (недели разложены по месяцам 4-4-5), пишет build_report.txt.

Результат: CSV для Power Query в Excel (разделитель «;», десятичная запятая, UTF-8).
"""
import re
from pathlib import Path

import pandas as pd
import pyarrow.parquet as pq
from python_calamine import CalamineWorkbook

# ---------------- НАСТРОЙКИ ----------------
BASE = Path(r"R:\2.ДКП\ПЭП\02_Аналитика\03_Базы данных\01_X5 Dialog")
RAW = {"weekly": BASE / "Raw data" / "ClickHouse" / "x5_weekly.parquet",
       "monthly": BASE / "Raw data" / "ClickHouse" / "x5_monthly.parquet"}
CLASSIFIER = BASE / "Справочники" / "X5_Products.xlsx"   # твой исправленный классификатор
SHOPS = BASE / "Справочники" / "shop_tree.xlsx"          # гео-справочник с портала
OUT = BASE / "Cube"
MAX_BAD_SHARE = 0.01   # доля строк с нарушенной формулой, после которой сборка останавливается

# ---------------- КОЛОНКИ КЛИКА ----------------
# Таблицы x5.x5 (недели) и x5.x5_month (месяцы). 15 колонок в клике названы не тем, что в них лежит.
# Соответствие ниже проверено формулами X5 на строках обеих таблиц и перепроверяется при каждом запуске.
COLUMNS = {
    "period": "Период", "period_start_date": "НачалоПериода", "period_end_date": "КонецПериода",
    "network": "Сеть", "region_group": "ФО", "region": "Регион", "product_group": "ГруппаТоваров", "plu": "PLU",
    "total_sales_rub": "TOTAL_Руб", "total_sales_units": "TOTAL_Шт", "total_sales_volume": "TOTAL_КгЛ",
    "total_receipts_count": "TOTAL_Чеки", "total_stores_count": "TOTAL_Магазины",
    "total_avg_price_per_unit_rub": "TOTAL_ЦенаШт", "total_avg_price_per_volume_rub": "TOTAL_ЦенаКгЛ",
    "total_avg_receipt_rub": "TOTAL_ЧекРуб",
    "promo_sales_rub": "PROMO_Руб", "promo_sales_units": "PROMO_Шт", "promo_sales_volume": "PROMO_КгЛ",
    "promo_stores_count": "PROMO_Магазины", "promo_sales_share_rub_pct": "PROMO_ДоляРуб",
    "promo_avg_price_per_volume_rub": "PROMO_ЦенаКгЛ",
    "loyal_customers_count": "LOYAL_Покупатели", "loyal_purchase_frequency": "LOYAL_Частота",
    "loyal_customer_units": "LOYAL_ШтНаПокупателя",
    # дальше название колонки в клике не совпадает с тем, что в ней лежит
    "loyal_avg_receipt_rub": "LOYAL_Руб",
    "loyal_avg_receipt_units": "LOYAL_Шт",
    "loyal_avg_receipt_volume": "LOYAL_КгЛ",
    "total_sales_per_store_per_day_rub": "LOYAL_Чеки",
    "total_sales_per_store_per_day_volume": "LOYAL_ЦенаКгЛ",
    "total_weighted_distribution_pct": "LOYAL_ЧекРуб",
    "loyal_sales_rub": "TOTAL_ВДеньРуб",
    "loyal_receipts_count": "TOTAL_ВзвДистр",
}

# В таблицы идут только базы. Цены, доли, средние чеки и пенетрации считаются мерами из них.
BASE_COLS = ["TOTAL_Руб", "TOTAL_Шт", "TOTAL_КгЛ", "TOTAL_Чеки", "TOTAL_Магазины",
             "PROMO_Руб", "PROMO_Шт", "PROMO_КгЛ", "PROMO_Магазины",
             "LOYAL_Руб", "LOYAL_Шт", "LOYAL_КгЛ", "LOYAL_Чеки", "LOYAL_Покупатели",
             "TOTAL_ВзвДистр", "TOTAL_ДнейПродаж"]
CELL = ["НачалоПериода", "Сеть", "Регион"]   # клетка, в которой X5 считает итог категории

# Формулы X5: (готовая колонка, числитель, знаменатель, множитель)
FORMULAS = [("TOTAL_ЦенаШт", "TOTAL_Руб", "TOTAL_Шт", 1), ("TOTAL_ЦенаКгЛ", "TOTAL_Руб", "TOTAL_КгЛ", 1),
            ("TOTAL_ЧекРуб", "TOTAL_Руб", "TOTAL_Чеки", 1), ("PROMO_ДоляРуб", "PROMO_Руб", "TOTAL_Руб", 100),
            ("PROMO_ЦенаКгЛ", "PROMO_Руб", "PROMO_КгЛ", 1), ("LOYAL_ЧекРуб", "LOYAL_Руб", "LOYAL_Чеки", 1),
            ("LOYAL_ЦенаКгЛ", "LOYAL_Руб", "LOYAL_КгЛ", 1), ("LOYAL_Частота", "LOYAL_Чеки", "LOYAL_Покупатели", 1),
            ("LOYAL_ШтНаПокупателя", "LOYAL_Шт", "LOYAL_Покупатели", 1)]

report = []


def log(text=""):
    print(text)
    report.append(str(text))


def to_num(s):
    """Число из ячейки: и настоящее число, и текст с десятичной запятой."""
    if pd.api.types.is_numeric_dtype(s):
        return s.astype("float64")
    t = s.astype(str).str.replace("\u00a0", "", regex=False).str.replace(" ", "", regex=False)
    t = t.str.replace(",", ".", regex=False).replace({"None": "", "nan": ""})
    return pd.to_numeric(t.where(t != ""), errors="coerce")


def read_table(path):
    """Справочник из первого листа xlsx: все ячейки текстом, без хвоста «.0» у чисел."""
    rows = CalamineWorkbook.from_path(str(path)).get_sheet_by_index(0).to_python()
    df = pd.DataFrame(rows[1:], columns=[str(h).strip() for h in rows[0]]).astype(str)
    return df.apply(lambda col: col.str.strip().str.replace(r"\.0$", "", regex=True).replace({"None": "", "nan": ""}))


def read_source(name):
    """Таблица из клика (Parquet) с настоящими названиями колонок, проверенная формулами X5."""
    path = RAW[name]
    missing = [c for c in COLUMNS if c not in pq.ParquetFile(path).schema_arrow.names]
    if missing:
        raise SystemExit(f"СТОП. {path.name}: нет колонок {missing}. Структура таблицы в клике поменялась.")
    df = pq.read_table(path, columns=list(COLUMNS)).to_pandas().rename(columns=COLUMNS)
    log(f"  {path.name}: строк {len(df)}")

    for c in ["НачалоПериода", "КонецПериода"]:
        df[c] = pd.to_datetime(df[c])
    for c in ["Сеть", "ФО", "Регион", "ГруппаТоваров", "PLU"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["PLU"] = df["PLU"].str.replace(r"\.0$", "", regex=True)          # в месячной таблице бывает «9994.0»
    df.loc[df["ГруппаТоваров"].str.upper() == "TOTAL", "PLU"] = ""     # у итогов PLU то пустой, то «0»
    orphans = (df["ГруппаТоваров"].str.upper() != "TOTAL") & df["PLU"].isin(["", "0"])
    if orphans.any():
        log(f"  ! строк без PLU, которые не TOTAL: {int(orphans.sum())} (отброшены)")
        df = df[~orphans]
    df["ЭтоИтог"] = df["PLU"] == ""

    def check(bad, what):
        """bad — нарушения среди проверенных строк. Больше MAX_BAD_SHARE — значит, колонки переставили."""
        if bad.any():
            ex = df.loc[bad[bad].index[0]]
            msg = (f"{what}: {int(bad.sum())} строк ({bad.mean():.2%}), например {ex['Сеть']}, {ex['Регион']}, "
                   f"{ex['Период']}, PLU {ex['PLU'] or 'TOTAL'}")
            log("  ! " + msg)
            if bad.mean() > MAX_BAD_SHARE:
                raise SystemExit(f"СТОП. {path.name}: {msg}.\nФормулы X5 не сходятся: в клике поменялись колонки.")

    for ready, top, bottom, k in FORMULAS:
        ok = df[bottom] > 0
        expect = k * df.loc[ok, top] / df.loc[ok, bottom]
        check((df.loc[ok, ready] - expect).abs() > (0.02 * expect.abs()).clip(lower=0.001), f"{ready} ≠ {top}/{bottom}")

    # Две колонки без формулы проверяются по смыслу: дни продаж не больше длины периода, дистрибуция 0–100
    ok = df["TOTAL_ВДеньРуб"] != 0
    days = df["TOTAL_Руб"] / df["TOTAL_Магазины"] / df["TOTAL_ВДеньРуб"].where(ok)
    length = (df["КонецПериода"] - df["НачалоПериода"]).dt.days + 1
    check(((days < 0.5) | (days > length + 0.5))[ok], "дней продаж вне периода")
    check((df["TOTAL_ВзвДистр"] < 0) | (df["TOTAL_ВзвДистр"] > 100.5), "взвешенная дистрибуция вне 0–100")
    df["TOTAL_ДнейПродаж"] = days.round()

    keep = ["Период", "НачалоПериода", "КонецПериода", "Сеть", "ФО", "Регион", "PLU", "ГруппаТоваров", "ЭтоИтог"]
    return df[keep + BASE_COLS]


def load_contour(name):
    """Факты по PLU и итоги категории (строки TOTAL) одного контура плюс проверки дублей и баланса."""
    df = read_source(name)
    exact = df.duplicated()
    if exact.any():
        by_period = df.loc[exact, "Период"].value_counts().sort_index()
        log(f"  ! полных дублей строк: {int(exact.sum())} (убраны), по периодам: "
            + ", ".join(f"{p}: {n}" for p, n in by_period.items()))
        df = df[~exact]
    clash = df.duplicated(CELL + ["PLU"], keep=False)

    # Итог категории, загруженный в клик дважды с разными цифрами: остаётся тот, что ближе к сумме PLU своей клетки
    twice = clash & df["ЭтоИтог"]
    if twice.any():
        plu_sum = df[~df["ЭтоИтог"]].groupby(CELL)["TOTAL_Руб"].sum().rename("СуммаPLU")
        t = df[twice].join(plu_sum, on=CELL)
        t["Отклонение"] = (t["TOTAL_Руб"] - t["СуммаPLU"]).abs()
        worse = t.sort_values("Отклонение").duplicated(CELL, keep="first")
        rub = lambda v: f"{v:,.2f}".replace(",", " ")
        for (start, net, reg), g in t.groupby(CELL):
            log(f"  ! итог категории загружен {len(g)} раза: {net}, {reg}, {start:%d.%m.%Y}: "
                f"{' и '.join(rub(v) for v in g['TOTAL_Руб'])} руб; сумма PLU {rub(g['СуммаPLU'].iloc[0])}, "
                f"оставлен ближайший к ней")
        df = df.drop(index=worse[worse].index)
        clash = df.duplicated(CELL + ["PLU"], keep=False)

    if clash.any():
        ex = df[clash].sort_values(CELL + ["PLU"]).head(4)[CELL + ["PLU", "TOTAL_Руб"]]
        raise SystemExit(f"СТОП. {RAW[name].name}: разные строки с одним ключом период × сеть × регион × PLU: "
                         f"{int(clash.sum())}. Например:\n{ex.to_string(index=False)}")

    fact = df[~df["ЭтоИтог"]].drop(columns="ЭтоИтог").reset_index(drop=True)
    cat = df[df["ЭтоИтог"]].drop(columns=["ЭтоИтог", "PLU", "ГруппаТоваров"]).reset_index(drop=True)

    # Баланс: сумма PLU против строки TOTAL в каждой клетке период × сеть × регион
    both = fact.groupby(CELL, as_index=False)["TOTAL_Руб"].sum().merge(
        cat[CELL + ["TOTAL_Руб"]], on=CELL, how="outer", suffixes=("_plu", "_total"))
    log(f"  фактов {len(fact)}, итогов категории {len(cat)}, периодов {fact['НачалоПериода'].nunique()} "
        f"({fact['НачалоПериода'].min():%d.%m.%Y} – {fact['КонецПериода'].max():%d.%m.%Y})")
    log(f"  баланс: PLU дают {both['TOTAL_Руб_plu'].sum() / both['TOTAL_Руб_total'].sum():.2%} от суммы строк TOTAL; "
        f"клеток без TOTAL: {int(both['TOTAL_Руб_total'].isna().sum())}, "
        f"клеток с расхождением больше 0,5%: {int(((both['TOTAL_Руб_plu'] / both['TOTAL_Руб_total'] - 1).abs() > 0.005).sum())} "
        f"из {len(both)}")
    return fact, cat


def month_445(week):
    """Номер месяца по схеме 4-4-5: недели 1-4, 5-8, 9-13, 14-17 ... 48-52; 53-я неделя уходит в декабрь."""
    quarter, rest = divmod(min(int(week), 52) - 1, 13)
    return quarter * 3 + (1 if rest < 4 else 2 if rest < 8 else 3)


def build_calendar(fact, weekly):
    cal = fact[["НачалоПериода", "КонецПериода", "Период"]].drop_duplicates().sort_values("НачалоПериода")
    if weekly:
        iso = cal["НачалоПериода"].dt.isocalendar()
        cal["Год"], cal["Неделя"] = iso["year"].astype(int), iso["week"].astype(int)
        cal["НеделяМетка"] = cal["Год"].astype(str) + "-W" + cal["Неделя"].astype(str).str.zfill(2)
        cal["Месяц445"] = cal["Неделя"].map(month_445)
    else:
        cal["Год"], cal["Месяц445"] = cal["НачалоПериода"].dt.year, cal["НачалоПериода"].dt.month
    cal = cal.rename(columns={"Месяц445": "Месяц"})
    cal["Квартал"] = (cal["Месяц"] - 1) // 3 + 1
    cal["МесяцМетка"] = cal["Год"].astype(str) + "-" + cal["Месяц"].astype(str).str.zfill(2)
    return cal.reset_index(drop=True)


def weight_sort(label):
    """Ключ сортировки интервала веса: «<=40г» -> 0, «55,1-60г» -> 55.1, «>120г» -> 120.5."""
    text = str(label).strip()
    nums = re.findall(r"\d+(?:[.,]\d+)?", text)
    if not nums:
        return None
    first = float(nums[0].replace(",", "."))
    return 0.0 if text.startswith("<") else first + (0.5 if text.startswith(">") else 0)


def build_products(facts):
    cls = read_table(CLASSIFIER).rename(columns={"PLU ID": "PLU"})
    if cls["PLU"].duplicated().any():
        raise SystemExit(f"СТОП. В классификаторе дубли PLU: {cls.loc[cls['PLU'].duplicated(), 'PLU'].tolist()[:10]}")
    cls["Размечен"] = "да"

    # PLU из данных, которых нет в классификаторе, не теряем: иначе их продажи выпадут из итогов
    seen = pd.concat(facts).groupby("PLU", as_index=False).agg(Название=("ГруппаТоваров", "last"), Руб=("TOTAL_Руб", "sum"))
    new = seen[~seen["PLU"].isin(cls["PLU"])].sort_values("Руб", ascending=False)
    log(f"  в классификаторе {len(cls)} PLU; в данных есть, а в классификаторе нет: {len(new)}")
    for _, r in new.head(30).iterrows():
        log(f"    {r['PLU']} | {r['Название']} | {r['Руб']:,.0f} руб".replace(",", " "))
    if len(new):
        add = pd.DataFrame({"PLU": new["PLU"], "Название": new["Название"], "Размечен": "нет"})
        cls = pd.concat([cls, add], ignore_index=True).fillna("")
        own = ["Группа товаров", "Категория", "Производитель (корректный)", "Бренд (корректный)", "Продуктовая линия"]
        cls.loc[cls["Размечен"] == "нет", own] = "Не размечено"
    cls["Вес нетто (абс.)"] = to_num(cls["Вес нетто (абс.)"])
    cls["Вес (сортировка)"] = cls["Вес (интерв.)"].map(weight_sort)
    return cls


def build_regions(facts):
    """Регион -> федеральный округ: берём из самих данных, код субъекта добавляем из shop_tree."""
    reg = pd.concat(facts)[["Регион", "ФО"]].drop_duplicates()
    twice = reg[reg["Регион"].duplicated(keep=False)]
    if len(twice):
        raise SystemExit(f"СТОП. Регион записан в нескольких ФО: {sorted(twice['Регион'].unique())[:10]}")
    shops = read_table(SHOPS).drop_duplicates("Регион")
    codes = shops[["Регион", "Код региона (субъекта РФ)", "Регион (Субъект РФ)"]].set_axis(
        ["Регион", "КодСубъекта", "СубъектРФ"], axis=1)
    reg = reg.merge(codes, on="Регион", how="left").fillna("")
    log(f"  регионов {len(reg)}, федеральных округов {reg['ФО'].nunique()}; "
        f"регионов без кода субъекта: {int((reg['КодСубъекта'] == '').sum())}")
    return reg.sort_values(["ФО", "Регион"]).reset_index(drop=True)


def save(df, name):
    """CSV для Power Query: «;», десятичная запятая, UTF-8. Целые пишутся целыми, дробные округляются."""
    df = df.copy()
    for c in df.select_dtypes("number"):
        values = df[c].dropna()
        df[c] = df[c].astype("Int64") if (values % 1 == 0).all() else df[c].round(2 if c.endswith("_Руб") else 3)
    df.to_csv(OUT / f"{name}.csv", sep=";", decimal=",", index=False, encoding="utf-8-sig", date_format="%Y-%m-%d")
    log(f"  {name}.csv: строк {len(df)}, колонок {len(df.columns)}")


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    log(f"Сборка куба X5 Dialog, {pd.Timestamp.now():%d.%m.%Y %H:%M}")
    data = {}
    for name in RAW:
        log(f"\n[{name}]")
        data[name] = load_contour(name)
    facts = [fact for fact, _ in data.values()]

    log("\n[справочники]")
    products, regions = build_products(facts), build_regions(facts)
    networks = pd.concat(facts)[["Сеть"]].drop_duplicates().sort_values("Сеть")

    log("\n[файлы]")
    for name, (fact, cat) in data.items():
        save(fact[CELL + ["PLU"] + BASE_COLS], f"x5_{name}_fact")
        save(cat[CELL + BASE_COLS], f"x5_{name}_category")
        save(build_calendar(fact, weekly=(name == "weekly")), f"x5_{name}_calendar")
    save(products, "x5_products")
    save(regions, "x5_regions")
    save(networks, "x5_networks")

    (OUT / "build_report.txt").write_text("\n".join(report), encoding="utf-8")
    print(f"\nГотово. Таблицы и build_report.txt лежат в {OUT}")


if __name__ == "__main__":
    main()
