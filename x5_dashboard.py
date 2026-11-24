"""Автономный HTML-дашборд X5 Dialog по файлам куба.

Запуск:  python x5_dashboard.py   (после x5_build.py)
Нужно:   pandas и numpy (уже стоят для x5_build.py)

Что делает:
  1. Читает CSV куба (факты, итоги категории, календарь, продукты, регионы).
  2. Собирает факты по неделям (месяцам) × сети × региону × PLU (или до ФО, если REGIONS = False).
     Части взвешенной глубины промо считаются на строках регионов, поэтому мера совпадает с Excel.
  3. Упаковывает данные в сжатый бинарный блок и встраивает его в HTML вместе со страницей.
  Результат: два файла в папке куба, X5_дашборд_недели.html и X5_дашборд_месяцы.html.
  Интернет для открытия не нужен, достаточно Chrome или Edge.
"""
import base64
import gzip
import json
from pathlib import Path

import numpy as np
import pandas as pd

# ---------------- НАСТРОЙКИ ----------------
BASE = Path(r"R:\2.ДКП\ПЭП\02_Аналитика\03_Базы данных\01_X5 Dialog")
CUBE = BASE / "Cube"                      # папка, куда пишет x5_build.py
OUT = CUBE                                # куда класть HTML
OWN_PRODUCER = "Маревен Фуд Сэнтрал"      # выделяется цветом на графиках производителей
START_YEAR = None                         # например 2024, чтобы облегчить файл; у первого года не будет YA
REGIONS = True                            # False — свести данные до ФО: файл в разы легче, но без среза по регионам
CONTOURS = {"weekly": "недели", "monthly": "месяцы"}
TITLE = "X5 Dialog · Продукция быстрого приготовления"

# Только суммируемые базы (плюс магазины для TDP): из них страница считает доли, цены и динамику
FACT_COLS = ["TOTAL_Шт", "TOTAL_Руб", "TOTAL_КгЛ", "PROMO_Шт", "PROMO_Руб", "PROMO_КгЛ",
             "LOYAL_Шт", "LOYAL_Руб", "LOYAL_КгЛ", "TOTAL_Магазины"]


def read(name, **kw):
    return pd.read_csv(CUBE / f"{name}.csv", sep=";", decimal=",", encoding="utf-8-sig", **kw)


def codes(values, order):
    """Индексы значений в списке order; значения вне списка — ошибка."""
    index = {v: i for i, v in enumerate(order)}
    out = values.map(index)
    if out.isna().any():
        raise SystemExit(f"СТОП. Нет в справочнике: {sorted(values[out.isna()].unique())[:10]}")
    return out.to_numpy()


def pack(columns):
    """Колонки -> один сжатый блок. Многобайтовые колонки хранятся побайтовыми плоскостями:
    старшие байты почти всегда нули и сжимаются почти в ничто. Каждая колонка выровнена на 8 байт."""
    header, chunks, offset = [], [], 0
    for name, (array, kind) in columns.items():
        array = np.ascontiguousarray(array)
        width = array.dtype.itemsize
        raw = array.view(np.uint8).reshape(-1, width).T.tobytes() if width > 1 else array.tobytes()
        pad = (-len(raw)) % 8
        header.append({"name": name, "type": kind, "offset": offset, "length": int(len(array)), "width": width})
        chunks.append(raw + b"\0" * pad)
        offset += len(raw) + pad
    blob = gzip.compress(b"".join(chunks), compresslevel=9, mtime=0)
    return header, base64.b64encode(blob).decode("ascii"), offset


def as_int(series, name, exact=False):
    """Целые для упаковки (int32). Рубли и килограммы округляются: на дашборде это незаметно.
    exact=True — для штук и магазинов: если там есть дроби, колонка без округления уходит как float64."""
    values = series.to_numpy(dtype="f8")
    rounded = np.round(values)
    if exact and np.abs(values - rounded).max(initial=0) > 1e-6:
        print(f"  ! {name}: есть дробные значения, колонка пойдёт как float64 (файл будет тяжелее)")
        return values.astype("<f8"), "f64"
    if np.abs(rounded).max(initial=0) < 2**31 - 1:
        return rounded.astype("<i4"), "i32"
    print(f"  ! {name}: значения больше int32, колонка пойдёт как float64")
    return values.astype("<f8"), "f64"


def as_u16(series, name):
    values = series.to_numpy(dtype="f8")
    if np.array_equal(values, np.round(values)) and values.min(initial=0) >= 0 and values.max(initial=0) < 65535:
        return values.astype("<u2"), "u16"
    return as_int(series, name, exact=True)


def build(contour):
    weekly = contour == "weekly"

    cal = read(f"x5_{contour}_calendar", dtype=str)
    for c in ["НачалоПериода", "КонецПериода"]:
        cal[c] = pd.to_datetime(cal[c])
    cal["НачалоПериодаПГ"] = pd.to_datetime(cal["НачалоПериодаПГ"], errors="coerce")
    for c in ["Год", "Месяц", "Квартал"] + (["Неделя"] if weekly else []):
        cal[c] = cal[c].astype(int)
    # Сквозной номер периода: по нему считаются скользящие окна (L4W, MAT) и PP.
    # Недели: каждая неделя из данных — следующий номер, пропуск в данных — столько номеров, сколько недель прошло.
    if weekly:
        gap = cal["НачалоПериода"].diff().dt.days
        cal["Номер"] = (gap / 7).round().clip(lower=1).fillna(0).cumsum().astype(int)
        odd = cal[cal["НачалоПериода"].dt.dayofweek != 0]
        if len(odd):
            day = ["пн", "вт", "ср", "чт", "пт", "сб", "вс"]
            print(f"  Недель не с понедельника: {len(odd)} из {len(cal)}, например: " + "; ".join(
                f"{s:%d.%m.%Y} ({day[s.dayofweek]}) – {e:%d.%m.%Y} ({day[e.dayofweek]})"
                for s, e in zip(odd["НачалоПериода"].head(6), odd["КонецПериода"].head(6))))
    else:
        cal["Номер"] = cal["Год"] * 12 + cal["Месяц"] - 1
    if START_YEAR:
        cal = cal[cal["Год"] >= START_YEAR]
    cal = cal.sort_values("НачалоПериода").reset_index(drop=True)
    per_index = {d: i for i, d in enumerate(cal["НачалоПериода"])}
    pg = cal["НачалоПериодаПГ"].map(per_index).fillna(-1).astype(int)

    regions = read("x5_regions", dtype=str).fillna("")
    regions["ФО"] = regions["ФО"].replace("", "Не указан")
    reg_fo = dict(zip(regions["Регион"], regions["ФО"]))
    products = read("x5_products", dtype=str).fillna("")
    for c in ["Категория", "Производитель (корректный)", "Бренд (корректный)", "Название"]:
        products[c] = products[c].replace("", "Не указано")

    keys = {"НачалоПериода": str, "Сеть": str, "Регион": str, "PLU": str}
    fact = read(f"x5_{contour}_fact", usecols=list(keys) + FACT_COLS, dtype=keys)
    fact["НачалоПериода"] = pd.to_datetime(fact["НачалоПериода"])
    fact = fact[fact["НачалоПериода"].isin(per_index)]
    fact["ФО"] = fact["Регион"].map(reg_fo)
    if fact["ФО"].isna().any():
        raise SystemExit(f"СТОП. Регионы без ФО: {sorted(fact.loc[fact['ФО'].isna(), 'Регион'].unique())}")
    fact[FACT_COLS] = fact[FACT_COLS].fillna(0)

    # Взвешенная глубина промо: обычная цена берётся в той же строке (SKU × сеть × регион × период)
    reg_units = fact["TOTAL_Шт"] - fact["PROMO_Шт"]
    reg_rub = fact["TOTAL_Руб"] - fact["PROMO_Руб"]
    ok = (fact["PROMO_Шт"] > 0) & (reg_units > 0) & (reg_rub > 0)
    fact["ГлубинаПолн"] = np.where(ok, reg_rub / reg_units.where(ok, 1) * fact["PROMO_Шт"], 0.0)
    fact["ГлубинаОпл"] = np.where(ok, fact["PROMO_Руб"], 0.0)

    nets = sorted(fact["Сеть"].unique())
    fos = sorted(fact["ФО"].unique())
    # География данных: регионы (каждый со своим ФО) или сразу ФО
    if REGIONS:
        geo = fact[["ФО", "Регион"]].drop_duplicates().sort_values(["ФО", "Регион"])
        geo_names, geo_col = geo["Регион"].tolist(), "Регион"
        geo_fo = codes(geo["ФО"], fos).tolist()
    else:
        geo_names, geo_col, geo_fo = fos, "ФО", list(range(len(fos)))
    fact["p"] = codes(fact["НачалоПериода"], list(per_index))
    fact["n"] = codes(fact["Сеть"], nets)
    fact["g"] = codes(fact[geo_col], geo_names)
    fact["k"] = codes(fact["PLU"], list(products["PLU"]))
    values = FACT_COLS + ["ГлубинаПолн", "ГлубинаОпл"]
    agg = fact.groupby(["k", "n", "g", "p"], as_index=False, sort=True)[values].sum()
    agg = agg[(agg[values] != 0).any(axis=1)].reset_index(drop=True)

    cat = read(f"x5_{contour}_category", usecols=["НачалоПериода", "Сеть", "Регион", "TOTAL_Магазины"], dtype=keys)
    cat["НачалоПериода"] = pd.to_datetime(cat["НачалоПериода"])
    cat["ФО"] = cat["Регион"].map(reg_fo)
    cat = cat[cat["НачалоПериода"].isin(per_index) & cat["Сеть"].isin(nets) & cat[geo_col].isin(geo_names)]
    cat["p"], cat["n"], cat["g"] = codes(cat["НачалоПериода"], list(per_index)), codes(cat["Сеть"], nets), codes(cat[geo_col], geo_names)
    cagg = cat.groupby(["p", "n", "g"], as_index=False, sort=True)["TOTAL_Магазины"].sum()

    # Справочники для срезов: продукты, производители и бренды по убыванию продаж
    units_by_k = agg.groupby("k")["TOTAL_Шт"].sum().reindex(range(len(products)), fill_value=0)
    products["_units"] = units_by_k.to_numpy()
    def ranked(col):
        return products.groupby(col)["_units"].sum().sort_values(ascending=False).index.tolist()
    categories = sorted(products["Категория"].unique())
    producers, brands = ranked("Производитель (корректный)"), ranked("Бренд (корректный)")
    slots = {}
    if OWN_PRODUCER in producers:
        slots[OWN_PRODUCER] = 0
    for name in producers:
        if len(slots) >= 8:
            break
        slots.setdefault(name, len(slots))

    label = cal["НеделяМетка"] if weekly else cal["МесяцМетка"]
    meta = {
        "title": TITLE, "contour": contour, "contourName": CONTOURS[contour],
        "generated": pd.Timestamp.now().strftime("%d.%m.%Y %H:%M"),
        "periods": {"label": label.tolist(), "year": cal["Год"].tolist(),
                    "num": (cal["Неделя"] if weekly else cal["Месяц"]).tolist(),
                    "month": cal["Месяц"].tolist(), "quarter": cal["Квартал"].tolist(), "ord": cal["Номер"].tolist(),
                    "start": cal["НачалоПериода"].dt.strftime("%d.%m.%Y").tolist(),
                    "end": cal["КонецПериода"].dt.strftime("%d.%m.%Y").tolist(), "pg": pg.tolist()},
        "nets": nets, "fos": fos, "regionLevel": REGIONS, "regions": geo_names if REGIONS else [], "geoFo": geo_fo,
        "categories": categories, "producers": producers, "brands": brands,
        "products": {"plu": products["PLU"].tolist(), "name": products["Название"].tolist(),
                     "cat": codes(products["Категория"], categories).tolist(),
                     "prod": codes(products["Производитель (корректный)"], producers).tolist(),
                     "brand": codes(products["Бренд (корректный)"], brands).tolist()},
        "own": producers.index(OWN_PRODUCER) if OWN_PRODUCER in producers else -1,
        "slots": [slots.get(name, -1) for name in producers],
        "check": {"rows": int(len(agg)), "units": float(agg["TOTAL_Шт"].sum()), "catRows": int(len(cagg))},
    }
    def rub_kop(col):
        """Рубли с копейками: целые рубли (int32) + копейки (0..99) отдельной колонкой."""
        kop = np.round(agg[col].to_numpy("f8") * 100)
        rub = np.floor_divide(kop, 100)
        return as_int(pd.Series(rub), col), ((kop - rub * 100).astype("u1"), "u8")

    geo_type = ("u1", "u8") if len(geo_names) <= 255 else ("<u2", "u16")
    columns = {
        "P": (agg["p"].to_numpy("<u2"), "u16"), "N": (agg["n"].to_numpy("u1"), "u8"),
        "Geo": (agg["g"].to_numpy(geo_type[0]), geo_type[1]), "K": (agg["k"].to_numpy("<u2"), "u16"),
        "U": as_int(agg["TOTAL_Шт"], "TOTAL_Шт", exact=True),
        "G": as_int(agg["TOTAL_КгЛ"] * 1000, "TOTAL_КгЛ"),                      # кг/л хранятся в граммах
        "PU": as_int(agg["PROMO_Шт"], "PROMO_Шт", exact=True), "LU": as_int(agg["LOYAL_Шт"], "LOYAL_Шт", exact=True),
        "S": as_u16(agg["TOTAL_Магазины"], "TOTAL_Магазины"),
        "cP": (cagg["p"].to_numpy("<u2"), "u16"), "cN": (cagg["n"].to_numpy("u1"), "u8"),
        "cGeo": (cagg["g"].to_numpy(geo_type[0]), geo_type[1]), "cS": as_int(cagg["TOTAL_Магазины"], "cS", exact=True),
    }
    for name, col in [("R", "TOTAL_Руб"), ("PR", "PROMO_Руб"), ("LR", "LOYAL_Руб")]:
        columns[name], columns[name + "k"] = rub_kop(col)
    # Килограммы промо и по картам у одного SKU — это штуки × вес штуки. Если так и есть, страница считает их сама
    # и файл легче; если хоть где-то не сходится, колонки кладутся как есть.
    weight = np.divide(agg["TOTAL_КгЛ"], agg["TOTAL_Шт"], out=np.zeros(len(agg)), where=agg["TOTAL_Шт"].to_numpy() != 0)
    same = lambda kg, units: np.allclose(agg[kg], agg[units] * weight, rtol=2e-3, atol=0.005)
    if not (same("PROMO_КгЛ", "PROMO_Шт") and same("LOYAL_КгЛ", "LOYAL_Шт")):
        print("  кг промо и по картам не сводятся к штукам × вес: кладу их в файл отдельно")
        columns["PG"] = as_int(agg["PROMO_КгЛ"] * 1000, "PROMO_КгЛ")
        columns["LG"] = as_int(agg["LOYAL_КгЛ"] * 1000, "LOYAL_КгЛ")
    # Части глубины промо: на уровне регионов страница считает их из строк сама, после свёртки до ФО — только готовые
    if not REGIONS:
        columns["DF"] = as_int(agg["ГлубинаПолн"], "ГлубинаПолн")
        columns["DS"] = as_int(agg["ГлубинаПолн"] - agg["ГлубинаОпл"], "ГлубинаСкидка")
    if len(products) > 65535 or len(cal) > 65535 or len(nets) > 255 or len(geo_names) > 65535:
        raise SystemExit("СТОП. Справочники не помещаются в формат файла.")
    meta["columns"], data, raw_size = pack(columns)
    return meta, data, len(fact), raw_size


def render(meta, data):
    meta_json = json.dumps(meta, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    page = TEMPLATE.replace("__X5_TITLE__", f"{meta['title']} · {meta['contourName']}")
    return page.replace("__X5_META__", meta_json).replace("__X5_DATA__", data)


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for contour, name in CONTOURS.items():
        meta, data, fact_rows, raw = build(contour)
        path = OUT / f"X5_дашборд_{name}.html"
        path.write_text(render(meta, data), encoding="utf-8")
        size = path.stat().st_size / 1e6
        print(f"{path.name}: строк фактов {fact_rows} -> после свёртки до {'регионов' if REGIONS else 'ФО'} "
              f"{meta['check']['rows']}, данные {raw / 1e6:.0f} МБ -> файл {size:.1f} МБ")
        if size > 25 and not START_YEAR:
            print("  Файл тяжёлый для почты. Облегчить: START_YEAR = 2024 или REGIONS = False (только ФО).")
    print(f"Готово. Файлы лежат в {OUT}")


TEMPLATE = r'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__X5_TITLE__</title>
<script>try { document.documentElement.dataset.theme = localStorage.getItem('x5-theme') || 'light'; } catch (e) { document.documentElement.dataset.theme = 'light'; }</script>
<style>
:root {
  color-scheme: light;
  --page: #f9f9f7; --surface: #fcfcfb; --ink: #0b0b0b; --ink2: #52514e; --muted: #898781;
  --grid: #e1e0d9; --axis: #c3c2b7; --border: rgba(11,11,11,0.10); --wash: rgba(11,11,11,0.05);
  --good: #006300; --bad: #d03b3b; --other: #898781; --track: #f0efec;
  --s1: #2a78d6; --s2: #eb6834; --s3: #1baf7a; --s4: #eda100; --s5: #e87ba4; --s6: #008300; --s7: #6250d6; --s8: #e34948;
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --page: #0d0d0d; --surface: #1a1a19; --ink: #f0efec; --ink2: #c3c2b7; --muted: #898781;
  --grid: #2c2c2a; --axis: #383835; --border: rgba(255,255,255,0.10); --wash: rgba(255,255,255,0.07);
  --good: #0ca30c; --bad: #e66767; --other: #898781; --track: #262624;
  --s1: #3987e5; --s2: #d95926; --s3: #199e70; --s4: #c98500; --s5: #d55181; --s6: #0ca30c; --s7: #9085e9; --s8: #e66767;
}
* { box-sizing: border-box; }
[hidden] { display: none !important; }
body { margin: 0; background: var(--page); color: var(--ink); font: 14px/1.45 system-ui, -apple-system, "Segoe UI", sans-serif; }
.wrap { max-width: 1600px; margin: 0 auto; padding: 20px 24px 40px; }
h1 { font-size: 20px; font-weight: 600; margin: 0; }
.sub { color: var(--ink2); font-size: 13px; margin-top: 2px; }
.filters { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin: 16px 0; position: sticky; top: 0; z-index: 30;
  background: var(--page); padding: 8px 0; border-bottom: 1px solid var(--border); }
.filters .spacer { flex: 1; }
.layout { display: grid; grid-template-columns: 248px minmax(0, 1fr); gap: 20px; align-items: start; }
.side { position: sticky; top: 64px; max-height: calc(100vh - 76px); overflow-y: auto; display: flex; flex-direction: column; gap: 12px;
  padding-right: 4px; }
.side h3 { font-size: 12px; font-weight: 600; text-transform: uppercase; letter-spacing: .04em; color: var(--muted); margin: 6px 0 -4px; }
.fl { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 10px 0 6px; }
.fl-head { display: flex; align-items: baseline; gap: 8px; padding: 0 12px 6px; }
.fl-title { font-weight: 600; font-size: 13px; }
.fl-count { color: var(--s1); font-size: 12px; }
.fl-reset { margin-left: auto; border: 0; background: none; color: var(--ink2); font: inherit; font-size: 12px; cursor: pointer; padding: 0; }
.fl-reset:hover { color: var(--ink); text-decoration: underline; }
.fl .ms-search { margin: 0 12px 6px; width: calc(100% - 24px); }
.fl .ms-list { max-height: 228px; padding: 0 0 4px; }
.fl .ms-item { padding: 3px 12px; font-size: 13px; }
.fl .ms-item.own span { font-weight: 600; }
.main { min-width: 0; }
@media (max-width: 1100px) { .layout { grid-template-columns: minmax(0, 1fr); } .side { position: static; max-height: none; } }
.seg { display: inline-flex; border: 1px solid var(--border); border-radius: 8px; overflow: hidden; background: var(--surface); }
.seg button { border: 0; background: transparent; color: var(--ink2); padding: 6px 12px; font: inherit; cursor: pointer; }
.seg button + button { border-left: 1px solid var(--border); }
.seg button[aria-pressed="true"] { background: var(--ink); color: var(--surface); }
.ms { position: relative; }
.ms-btn, .btn { display: inline-flex; align-items: center; gap: 6px; height: 32px; padding: 0 10px; border: 1px solid var(--border);
  border-radius: 8px; background: var(--surface); color: var(--ink); font: inherit; cursor: pointer; max-width: 260px; }
.ms-btn:hover, .btn:hover { background: var(--wash); }
.ms-label { color: var(--ink2); }
.ms-val { font-weight: 500; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.ms-btn[data-active="true"] { border-color: var(--s1); }
.field { display: inline-flex; align-items: center; gap: 6px; height: 32px; padding: 0 0 0 10px; border: 1px solid var(--border);
  border-radius: 8px; background: var(--surface); }
.select { height: 30px; border: 0; background: transparent; color: var(--ink); font: inherit; font-weight: 500; padding: 0 6px; max-width: 230px; cursor: pointer; }
.select option { background: var(--surface); color: var(--ink); }
.period { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 12px; color: var(--ink2); font-size: 13px; margin: -6px 0 14px; }
.period .spacer { flex: 1; }
.card-actions { display: flex; gap: 6px; }
.status { color: var(--ink2); font-size: 12px; }
.period b { color: var(--ink); font-weight: 600; }
.period .warn { color: var(--ink); }
.ms-panel { position: absolute; top: 36px; left: 0; z-index: 40; width: 300px; max-height: 360px; display: flex; flex-direction: column;
  background: var(--surface); border: 1px solid var(--border); border-radius: 10px; box-shadow: 0 8px 24px rgba(0,0,0,.12); }
.ms-panel[hidden] { display: none; }
.ms-search { margin: 8px; height: 30px; padding: 0 8px; border: 1px solid var(--border); border-radius: 6px; background: var(--page); color: var(--ink); font: inherit; }
.ms-list { overflow: auto; padding: 4px 0 8px; }
.ms-item { display: flex; gap: 8px; align-items: center; padding: 4px 12px; cursor: pointer; }
.ms-item:hover { background: var(--wash); }
.ms-item input { margin: 0; accent-color: var(--s1); }
.ms-all { border-bottom: 1px solid var(--border); padding-bottom: 6px; margin-bottom: 4px; font-weight: 500; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(136px, 1fr)); gap: 12px; margin-bottom: 16px; }
.tile { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 12px 14px; min-width: 0; }
.tile .label { color: var(--ink2); font-size: 13px; }
.tile .value { font-size: clamp(18px, 1.55vw, 24px); font-weight: 600; margin-top: 4px; white-space: nowrap; }
.delta { font-size: 13px; margin-top: 2px; white-space: nowrap; }
.good { color: var(--good); } .bad { color: var(--bad); } .neutral { color: var(--ink2); } .none { color: var(--muted); }
.grid2 { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; margin-bottom: 16px; }
@media (max-width: 980px) { .grid2 { grid-template-columns: minmax(0, 1fr); } }
.card { position: relative; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 16px; min-width: 0; }
.card + .card { margin-top: 0; }
.full { margin-bottom: 16px; }
.card-head { display: flex; justify-content: space-between; gap: 12px; align-items: flex-start; margin-bottom: 8px; }
.card h2 { font-size: 15px; font-weight: 600; margin: 0; }
.card .note { color: var(--ink2); font-size: 12px; margin-top: 2px; }
.ghost { border: 1px solid var(--border); background: transparent; color: var(--ink2); border-radius: 8px; height: 28px; padding: 0 10px;
  font: inherit; font-size: 12px; cursor: pointer; white-space: nowrap; }
.ghost:hover { background: var(--wash); }
.chart { width: 100%; min-height: 40px; }
.chart svg { display: block; overflow: visible; }
.chart text { fill: var(--muted); font-size: 11px; font-variant-numeric: tabular-nums; }
.chart .endlabel { fill: var(--ink2); font-size: 12px; }
.legend { display: flex; flex-wrap: wrap; gap: 6px 16px; margin-top: 8px; color: var(--ink2); font-size: 12px; }
.legend span { display: inline-flex; align-items: center; gap: 6px; }
.key-line { width: 14px; height: 2px; border-radius: 1px; display: inline-block; }
.key-rect { width: 10px; height: 10px; border-radius: 2px; display: inline-block; }
.tip { position: absolute; z-index: 20; pointer-events: none; background: var(--surface); border: 1px solid var(--border); border-radius: 8px;
  padding: 8px 10px; box-shadow: 0 6px 18px rgba(0,0,0,.14); font-size: 12px; min-width: 150px; }
.tip[hidden] { display: none; }
.tip .t-title { color: var(--ink2); margin-bottom: 4px; }
.tip .t-row { display: flex; align-items: center; gap: 8px; }
.tip .t-val { font-weight: 600; color: var(--ink); font-variant-numeric: tabular-nums; }
.tip .t-name { color: var(--ink2); }
.bars { display: grid; grid-template-columns: minmax(90px, 34%) 1fr; row-gap: 2px; }
.bar-label { padding: 5px 8px 5px 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; color: var(--ink2); font-size: 13px; }
.bar-label.own { color: var(--ink); font-weight: 600; }
.bar-cell { display: flex; flex-direction: column; justify-content: center; gap: 2px; padding: 3px 0; border-radius: 4px; outline-offset: 2px; }
.bar-row { display: flex; align-items: center; gap: 6px; min-height: 12px; }
.bar { height: 16px; border-radius: 0 4px 4px 0; min-width: 1px; }
.bar.thin { height: 10px; }
.bar-val { font-size: 12px; color: var(--ink2); font-variant-numeric: tabular-nums; white-space: nowrap; }
.bar-cell:hover .bar, .bar-cell:focus .bar { filter: brightness(1.12); }
table.data { width: 100%; border-collapse: collapse; font-size: 13px; }
table.data th { text-align: left; font-weight: 500; color: var(--ink2); padding: 6px 8px; border-bottom: 1px solid var(--axis); cursor: pointer;
  user-select: none; white-space: nowrap; vertical-align: bottom; }
table.data th.num, table.data td.num { text-align: right; font-variant-numeric: tabular-nums; }
table.data th[aria-sort] { color: var(--ink); }
table.data td { padding: 5px 8px; border-bottom: 1px solid var(--grid); }
table.data td.name { max-width: 320px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
table.data tr.total td { font-weight: 600; border-top: 1px solid var(--axis); }
table.data td .own { font-weight: 600; }
.plu { color: var(--muted); font-size: 12px; margin-left: 6px; }
.table-wrap { overflow-x: auto; }
.more { margin-top: 8px; }
.loading, .error { padding: 48px; text-align: center; color: var(--ink2); }
.error { color: var(--bad); }
.empty { color: var(--muted); padding: 24px 0; text-align: center; }
footer { color: var(--ink2); font-size: 12px; line-height: 1.6; margin-top: 8px; }
footer p { margin: 4px 0; max-width: 1100px; }
.sr-only { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; }
:focus-visible { outline: 2px solid var(--s1); outline-offset: 2px; }
</style>
</head>
<body>
<div class="wrap">
  <header>
    <h1 id="title"></h1>
    <div class="sub" id="subtitle"></div>
  </header>
  <div id="app"><div class="loading">Загружаю данные…</div></div>
</div>
<script type="application/json" id="x5-meta">__X5_META__</script>
<script type="application/octet-stream" id="x5-data">__X5_DATA__</script>
<script>
(async function () {
  'use strict';
  const META = JSON.parse(document.getElementById('x5-meta').textContent);
  const app = document.getElementById('app');
  document.getElementById('title').textContent = META.title;

  // ---------- утилиты ----------
  function el(tag, attrs, ...kids) {
    const e = document.createElement(tag);
    for (const [k, v] of Object.entries(attrs || {})) {
      if (v == null || v === false) continue;
      if (k === 'class') e.className = v;
      else if (k === 'text') e.textContent = v;
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
      else e.setAttribute(k, v === true ? '' : v);
    }
    for (const c of kids) if (c != null) e.append(c);
    return e;
  }
  function svg(tag, attrs) {
    const e = document.createElementNS('http://www.w3.org/2000/svg', tag);
    for (const [k, v] of Object.entries(attrs || {})) e.setAttribute(k, v);
    return e;
  }
  const nf0 = new Intl.NumberFormat('ru-RU', { maximumFractionDigits: 0 });
  const nf1 = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const nf2 = new Intl.NumberFormat('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  const ok = v => v != null && isFinite(v);
  const fmt = {
    int: v => ok(v) ? nf0.format(v) : '—',
    compact: v => {
      if (!ok(v)) return '—';
      const a = Math.abs(v);
      if (a >= 1e9) return nf1.format(v / 1e9) + ' млрд';
      if (a >= 1e6) return nf1.format(v / 1e6) + ' млн';
      if (a >= 1e4) return nf0.format(v / 1e3) + ' тыс.';
      return nf0.format(v);
    },
    pct: v => ok(v) ? nf1.format(v * 100) + '%' : '—',
    rub: v => ok(v) ? nf2.format(v) + ' ₽' : '—',
    idx: v => ok(v) ? nf2.format(v) : '—',
    tdp: v => ok(v) ? nf1.format(v * 100) : '—',
    sales: v => ok(v) ? (Math.abs(v) >= 100 || v === 0 ? nf0 : nf1).format(v) : '—',
  };
  // Изменение: проценты (pct, v — доля), процентные пункты (pp) или пункты TDP (pt). Ноль после округления — без стрелки.
  function delta(v, kind) {
    if (!ok(v)) return null;
    const x = kind === 'pct' ? v * 100 : v, r = Math.round(x * 10) / 10;
    const unit = kind === 'pct' ? '%' : kind === 'pp' ? ' п.п.' : '';
    if (r === 0) return { text: nf1.format(0) + unit, dir: 0 };
    return { text: (r > 0 ? '+' : '−') + nf1.format(Math.abs(r)) + unit, dir: r > 0 ? 1 : -1 };
  }
  const div = (a, b) => (b ? a / b : null);
  const growth = (a, b) => (b > 0 ? a / b - 1 : null);
  const MONTHS = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь'];
  const MON3 = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек'];

  // ---------- данные ----------
  async function decode() {
    if (typeof DecompressionStream === 'undefined') {
      throw new Error('Этот браузер не умеет распаковывать данные дашборда. Откройте файл в Chrome или Edge.');
    }
    const text = document.getElementById('x5-data').textContent.trim();
    let bytes;
    if (Uint8Array.fromBase64) bytes = Uint8Array.fromBase64(text);
    else {
      const bin = atob(text);
      bytes = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
    }
    const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream('gzip'));
    const buf = await new Response(stream).arrayBuffer();
    const types = { u8: Uint8Array, u16: Uint16Array, i32: Int32Array, f32: Float32Array, f64: Float64Array };
    const out = {};
    for (const c of META.columns) {
      if (c.width === 1) { out[c.name] = new types[c.type](buf, c.offset, c.length); continue; }
      const planes = new Uint8Array(buf, c.offset, c.length * c.width), bytes = new Uint8Array(c.length * c.width);
      for (let b = 0; b < c.width; b++) {
        const base = b * c.length;
        for (let i = 0, j = b; i < c.length; i++, j += c.width) bytes[j] = planes[base + i];
      }
      out[c.name] = new types[c.type](bytes.buffer);
    }
    return out;
  }
  let D;
  try {
    D = await decode();
    let units = 0;
    for (let i = 0; i < D.U.length; i++) units += D.U[i];
    if (D.P.length !== META.check.rows || Math.abs(units - META.check.units) > 1e-6 * Math.max(1, Math.abs(META.check.units))) throw new Error('Данные в файле повреждены: контрольные суммы не сходятся.');
  } catch (e) {
    app.textContent = '';
    app.append(el('div', { class: 'error', text: e.message }));
    return;
  }

  const per = META.periods, nP = per.label.length;
  const weekly = META.contour === 'weekly';
  const nN = META.nets.length, nF = META.fos.length, nK = META.products.plu.length;
  const nProd = META.producers.length, nBrand = META.brands.length;
  const catOf = META.products.cat, prodOf = META.products.prod, brandOf = META.products.brand;
  const geoFo = META.geoFo, nGeo = geoFo.length, regionLevel = META.regionLevel;   // регион (или ФО) -> ФО
  // Свой бренд — больше половины его SKU у своего производителя (в классификаторе бывают единичные чужие SKU)
  const ownCnt = new Int32Array(nBrand), allCnt = new Int32Array(nBrand);
  for (let k = 0; k < nK; k++) { allCnt[brandOf[k]]++; if (prodOf[k] === META.own) ownCnt[brandOf[k]]++; }
  const isOwnBrand = b => META.own >= 0 && ownCnt[b] * 2 > allCnt[b];
  document.getElementById('subtitle').textContent =
    `${weekly ? 'Недели; месяцы и кварталы по 4-4-5' : 'Календарные месяцы'} · данные с ${per.label[0]} по ${per.label[nP - 1]} · собрано ${META.generated}`;

  // ---------- периоды ----------
  // ord — сквозной номер недели (месяца). По нему строятся скользящие окна и PP: пропуски в данных окна не сдвигают.
  const ord = per.ord, ordIndex = new Map(ord.map((o, p) => [o, p]));
  const SPANS = weekly
    ? [['P', 'Нед', 'Одна неделя'], ['L4', 'L4W', 'Последние 4 недели'], ['L13', 'L13W', 'Последние 13 недель'],
       ['MTD', 'MTD', 'С начала месяца 4-4-5 до выбранной недели'], ['QTD', 'QTD', 'С начала квартала 4-4-5 до выбранной недели'],
       ['YTD', 'YTD', 'С начала года до выбранной недели'], ['MAT', 'MAT', 'Последние 52 недели']]
    : [['P', 'Мес', 'Один месяц'], ['L3', 'L3M', 'Последние 3 месяца'], ['QTD', 'QTD', 'С начала квартала до выбранного месяца'],
       ['YTD', 'YTD', 'С начала года до выбранного месяца'], ['MAT', 'MAT', 'Последние 12 месяцев']];
  const ROLL = { P: 1, L4: 4, L13: 13, L3: 3, MAT: weekly ? 52 : 12 };
  const CHART_LEN = weekly ? 52 : 12, UNIT = weekly ? 'нед.' : 'мес.';
  // Подпись точки оси: неделя — из календаря (пропуск в данных без подписи), месяц — из сквозного номера
  function axisLabel(a) {
    if (weekly) {
      if (a.p < 0) return { short: '', long: 'Нет данных за эту неделю' };
      const p = a.p;
      return { short: 'Н' + per.num[p], long: `Неделя ${per.num[p]} ${per.year[p]} (${per.start[p].slice(0, 5)}–${per.end[p].slice(0, 5)})` };
    }
    const y = Math.floor(a.o / 12), mo = a.o - y * 12;
    return { short: mo === 0 ? `${MON3[0]} ${String(y).slice(2)}` : MON3[mo], long: `${MONTHS[mo]} ${y}` + (a.p < 0 ? ' · нет данных' : '') };
  }
  // Единица, в которой считаются продажи, доли, доля промо и доля по картам
  const MEASURES = [['U', 'шт', 'В штуках'], ['R', 'руб', 'В рублях'], ['G', 'кг', 'В килограммах (литрах)']];
  const mu = () => MEASURES.find(m => m[0] === state.unit)[1];
  const salesDigits = () => (state.unit === 'U' ? 0 : state.unit === 'R' ? 2 : 3);
  const periodName = p => (weekly ? `${per.label[p]} · ${per.start[p].slice(0, 5)}–${per.end[p].slice(0, 5)}`
    : `${MONTHS[per.month[p] - 1]} ${per.year[p]}`);

  // ---------- состояние и расчёт ----------
  const DEFAULTS = { span: 'YTD', cmp: 'YA', unit: 'U', block: 'T' };
  const state = { end: nP - 1, ...DEFAULTS,
    nets: new Set(), fos: new Set(), regs: new Set(), cats: new Set(), prods: new Set(), brands: new Set() };

  function mask(n, set) {
    const m = new Uint8Array(n);
    for (let i = 0; i < n; i++) m[i] = set.size === 0 || set.has(i) ? 1 : 0;
    return m;
  }

  // W — периоды выбранного окна. C — окно сравнения: YA — те же недели ISO (те же месяцы) годом раньше,
  // PP — столько же недель (месяцев) прямо перед W.
  function periodWindow(st) {
    const e = st.end, oe = ord[e], W = [];
    let lo;
    if (ROLL[st.span]) {
      lo = oe - ROLL[st.span] + 1;
      for (let p = 0; p < nP; p++) if (ord[p] >= lo && ord[p] <= oe) W.push(p);
    } else {
      for (let p = 0; p < nP; p++) {
        if (ord[p] > oe || per.year[p] !== per.year[e]) continue;
        if (st.span === 'QTD' && per.quarter[p] !== per.quarter[e]) continue;
        if (st.span === 'MTD' && per.month[p] !== per.month[e]) continue;
        W.push(p);
      }
      lo = ord[W[0]];
    }
    const len = oe - lo + 1, C = [];
    if (st.cmp === 'YA') { for (const p of W) if (per.pg[p] >= 0) C.push(per.pg[p]); }
    else for (let p = 0; p < nP; p++) if (ord[p] >= lo - len && ord[p] <= oe - len) C.push(p);
    return { W, C, lo, oe, len, missing: len - W.length, cmpMissing: (st.cmp === 'YA' ? W.length : len) - C.length };
  }

  // Маски срезов: сеть, география (ФО и регион), категория, производитель и бренд
  function masks(st) {
    const netM = mask(nN, st.nets), foM = mask(nF, st.fos), regM = mask(nGeo, st.regs);
    const geoOk = new Uint8Array(nGeo);
    for (let g = 0; g < nGeo; g++) geoOk[g] = foM[geoFo[g]] && regM[g] ? 1 : 0;
    const catM = new Uint8Array(nK), fullM = new Uint8Array(nK);
    for (let k = 0; k < nK; k++) {
      const c = st.cats.size === 0 || st.cats.has(catOf[k]);
      catM[k] = c ? 1 : 0;
      fullM[k] = c && (st.prods.size === 0 || st.prods.has(prodOf[k])) && (st.brands.size === 0 || st.brands.has(brandOf[k])) ? 1 : 0;
    }
    return { netM, geoOk, catM, fullM };
  }

  // Значение в выбранной единице. Рубли — с копейками; кг без своей колонки — штуки × вес штуки этой строки.
  function measureOf(unit, units, rub, kop, kg) {
    const U = D.U, G = D.G;
    if (unit === 'U') return i => units[i];
    if (unit === 'R') return i => rub[i] + kop[i] * 0.01;
    if (kg) return i => kg[i] * 0.001;
    return i => (U[i] ? units[i] * G[i] * 0.001 / U[i] : 0);
  }
  // Блоки метрик X5: TOTAL — все продажи, PROMO — продажи по промо, LOYAL — продажи по картам лояльности
  const BLOCKS = [['T', 'Тотал', '', 'Все продажи'], ['P', 'Промо', 'промо', 'Продажи по промо'], ['L', 'Лояльные', 'по картам', 'Продажи по картам лояльности']];
  const blockArrays = b => (b === 'T' ? [D.U, D.R, D.Rk, D.G] : b === 'P' ? [D.PU, D.PR, D.PRk, D.PG] : [D.LU, D.LR, D.LRk, D.LG]);
  const valueOf = (b, unit) => measureOf(unit, ...blockArrays(b));
  const bw = () => BLOCKS.find(b => b[0] === state.block)[2];          // '', 'промо', 'по картам'
  const bwSp = () => (bw() ? ' ' + bw() : '');

  function compute(st) {
    const win = periodWindow(st);
    const pf = new Uint8Array(nP);   // 1 — окно, 2 — сравнение, 4 — ось графиков, 8 — YA для графика
    for (const p of win.W) pf[p] |= 1;
    for (const p of win.C) pf[p] |= 2;
    const axis = [];
    for (let o = Math.min(win.lo, win.oe - CHART_LEN + 1); o <= win.oe; o++) {
      const p = ordIndex.has(o) ? ordIndex.get(o) : -1;
      axis.push({ o, p });
      if (p >= 0) { pf[p] |= 4; if (per.pg[p] >= 0) pf[per.pg[p]] |= 8; }
    }
    const { netM, geoOk, catM, fullM } = masks(st);
    // t, p, l — тотал, промо и продажи по картам в выбранной единице; m — продажи выбранного блока.
    // Цены блока — его рубли к его штукам и килограммам.
    const tOf = valueOf('T', st.unit), pOf = valueOf('P', st.unit), lOf = valueOf('L', st.unit);
    const bUof = valueOf(st.block, 'U'), bRof = valueOf(st.block, 'R'), bGof = valueOf(st.block, 'G');
    const promoRub = valueOf('P', 'R');
    const blk = st.block;
    let wM = 0, wT = 0, wP = 0, wL = 0, wU = 0, wR = 0, wG = 0, wDF = 0, wDS = 0, catV = 0, catR = 0, catG = 0;
    let cmpM = 0, cmpT = 0, cmpP = 0, cmpL = 0, cmpU = 0, cmpR = 0, cmpG = 0, cmpDF = 0, cmpDS = 0, cmpCat = 0;
    const selM_p = new Float64Array(nP), selS_p = new Float64Array(nP), catM_p = new Float64Array(nP), catS_p = new Float64Array(nP);
    const prodM = new Float64Array(nProd), prodT = new Float64Array(nProd), prodP = new Float64Array(nProd), prodL = new Float64Array(nProd);
    const prodDF = new Float64Array(nProd), prodDS = new Float64Array(nProd), prodM_p = new Float64Array(nP * nProd);
    const brM = new Float64Array(nBrand), brC = new Float64Array(nBrand), brT = new Float64Array(nBrand), brP = new Float64Array(nBrand);
    const brL = new Float64Array(nBrand), brU = new Float64Array(nBrand), brR = new Float64Array(nBrand), brG = new Float64Array(nBrand);
    const brDF = new Float64Array(nBrand), brDS = new Float64Array(nBrand), brS_p = new Float64Array(nP * nBrand);
    const foV = new Float64Array(nF), foC = new Float64Array(nF), foCat = new Float64Array(nF), foCatC = new Float64Array(nF);
    const geoV = new Float64Array(nGeo), geoC = new Float64Array(nGeo), geoCat = new Float64Array(nGeo), geoCatC = new Float64Array(nGeo);
    const nV = new Float64Array(nN), nC = new Float64Array(nN), nCat = new Float64Array(nN), nCatC = new Float64Array(nN);
    const skM = new Float64Array(nK), skC = new Float64Array(nK), skT = new Float64Array(nK), skP = new Float64Array(nK);
    const skL = new Float64Array(nK), skU = new Float64Array(nK), skR = new Float64Array(nK), skG = new Float64Array(nK);
    const P = D.P, N = D.N, GEO = D.Geo, K = D.K, U = D.U, R = D.R, Rk = D.Rk, S = D.S, PU = D.PU;
    const DFa = D.DF || null, DSa = D.DS || null;
    // части глубины промо строки: обычная цена той же строки × промо-штуки (после свёртки до ФО — готовые)
    let dF = 0, dS = 0;
    const depthOf = i => {
      if (DFa) { dF = DFa[i]; dS = DSa[i]; return; }
      const u = U[i], pu = PU[i], paid = promoRub(i), ru = u - pu, rr = R[i] + Rk[i] * 0.01 - paid;
      if (pu > 0 && ru > 0 && rr > 0) { dF = rr / ru * pu; dS = dF - paid; } else { dF = 0; dS = 0; }
    };
    for (let i = 0, n = P.length; i < n; i++) {
      const p = P[i], fl = pf[p];
      if (!fl) continue;
      const net = N[i], g = GEO[i];
      if (!netM[net] || !geoOk[g]) continue;
      const k = K[i];
      if (!catM[k]) continue;
      const tv = tOf(i), pv = pOf(i), lv = lOf(i), m = blk === 'T' ? tv : blk === 'P' ? pv : lv;
      const pr = prodOf[k], full = fullM[k];
      catM_p[p] += m; prodM_p[p * nProd + pr] += m;
      if (full) { selM_p[p] += m; selS_p[p] += S[i]; }
      if (fl & 1) {
        const f = geoFo[g], br = bRof(i), bg = bGof(i);
        depthOf(i);
        catV += m; catR += br; catG += bg; foCat[f] += m; geoCat[g] += m; nCat[net] += m;
        prodM[pr] += m; prodT[pr] += tv; prodP[pr] += pv; prodL[pr] += lv; prodDF[pr] += dF; prodDS[pr] += dS;
        if (full) {
          const b = brandOf[k], bu = bUof(i);
          wM += m; wT += tv; wP += pv; wL += lv; wU += bu; wR += br; wG += bg; wDF += dF; wDS += dS;
          brM[b] += m; brT[b] += tv; brP[b] += pv; brL[b] += lv; brU[b] += bu; brR[b] += br; brG[b] += bg;
          brDF[b] += dF; brDS[b] += dS; brS_p[p * nBrand + b] += S[i];
          foV[f] += m; geoV[g] += m; nV[net] += m;
          skM[k] += m; skT[k] += tv; skP[k] += pv; skL[k] += lv; skU[k] += bu; skR[k] += br; skG[k] += bg;
        }
      }
      if (fl & 2) {
        const f = geoFo[g];
        cmpCat += m; foCatC[f] += m; geoCatC[g] += m; nCatC[net] += m;
        if (full) {
          depthOf(i);
          cmpM += m; cmpT += tv; cmpP += pv; cmpL += lv; cmpU += bUof(i); cmpR += bRof(i); cmpG += bGof(i); cmpDF += dF; cmpDS += dS;
          brC[brandOf[k]] += m; foC[f] += m; geoC[g] += m; nC[net] += m; skC[k] += m;
        }
      }
    }
    const cPer = D.cP, cN = D.cN, cGeo = D.cGeo, cS = D.cS;
    for (let j = 0; j < cPer.length; j++) {
      const p = cPer[j];
      if ((pf[p] & 3) && netM[cN[j]] && geoOk[cGeo[j]]) catS_p[p] += cS[j];
    }
    // TDP: среднее по неделям (месяцам) окна; знаменатель — магазины с продажами всей категории БП
    const tdpOver = (periods, stores, offset, stride) => {
      let sum = 0, cnt = 0;
      for (const p of periods) if (catS_p[p] > 0) { sum += stores[p * stride + offset] / catS_p[p]; cnt++; }
      return cnt ? sum / cnt : null;
    };
    const hasCmp = win.C.length > 0;
    const dShare = (a, b) => (hasCmp && ok(a) && ok(b) ? 100 * (a - b) : null);
    const depth = (ds, df) => (df > 0 ? ds / df : null);
    const kpi = {
      sales: wM, salesCmp: cmpM, share: div(wM, catV), shareCmp: div(cmpM, cmpCat),
      promoShare: div(wP, wT), promoShareCmp: div(cmpP, cmpT), loyalShare: div(wL, wT), loyalShareCmp: div(cmpL, cmpT),
      price: div(wR, wU), priceCmp: div(cmpR, cmpU), priceKg: div(wR, wG), priceKgCmp: div(cmpR, cmpG),
      depth: depth(wDS, wDF), depthCmp: depth(cmpDS, cmpDF),
      tdp: tdpOver(win.W, selS_p, 0, 1), tdpCmp: tdpOver(win.C, selS_p, 0, 1), hasCmp,
    };
    const chart1 = { axis, band: [axis.findIndex(a => a.o >= win.lo), axis.length - 1],
      cur: axis.map(a => (a.p >= 0 ? selM_p[a.p] : null)),
      ya: axis.map(a => (a.p >= 0 && per.pg[a.p] >= 0 ? selM_p[per.pg[a.p]] : null)) };
    const prodRank = [...Array(nProd).keys()].filter(i => prodM[i] > 0).sort((a, b) => prodM[b] - prodM[a]);
    const top5 = prodRank.slice(0, 5), top10 = prodRank.slice(0, 10);
    const chart2 = { series: top5.map(pr => ({ prod: pr,
      values: axis.map(a => (a.p >= 0 && catM_p[a.p] > 0 ? prodM_p[a.p * nProd + pr] / catM_p[a.p] : null)) })) };
    const chart3 = top10.map(pr => ({ prod: pr, share: div(prodM[pr], catV) }));
    const chart4 = top10.map(pr => ({ prod: pr, promoShare: div(prodP[pr], prodT[pr]), loyalShare: div(prodL[pr], prodT[pr]),
      depth: depth(prodDS[pr], prodDF[pr]) }));
    const brands = [], fos = [], regions = [], nets = [], skus = [];
    for (let b = 0; b < nBrand; b++) {
      if (!brM[b] && !brC[b]) continue;
      const share = div(brM[b], catV), shareCmp = div(brC[b], cmpCat);
      brands.push({ brand: b, sales: brM[b], salesCmp: brC[b], growth: growth(brM[b], brC[b]), share, shareCmp,
        dShare: dShare(share, shareCmp), promoShare: div(brP[b], brT[b]), loyalShare: div(brL[b], brT[b]),
        price: div(brR[b], brU[b]), priceKg: div(brR[b], brG[b]),
        index: brG[b] > 0 && catG > 0 && catR > 0 ? (brR[b] / brG[b]) / (catR / catG) : null,
        depth: depth(brDS[b], brDF[b]), tdp: tdpOver(win.W, brS_p, b, nBrand) });
    }
    const geoRow = (v, c, cat, catC) => {
      const share = div(v, cat), shareCmp = div(c, catC);
      return { sales: v, salesCmp: c, growth: growth(v, c), share, shareCmp, dShare: dShare(share, shareCmp) };
    };
    for (let f = 0; f < nF; f++) if (foCat[f] || foCatC[f]) fos.push({ fo: f, ...geoRow(foV[f], foC[f], foCat[f], foCatC[f]) });
    if (regionLevel) for (let g = 0; g < nGeo; g++) if (geoCat[g] || geoCatC[g]) regions.push({ geo: g, fo: geoFo[g], ...geoRow(geoV[g], geoC[g], geoCat[g], geoCatC[g]) });
    for (let n = 0; n < nN; n++) if (nCat[n] || nCatC[n]) nets.push({ net: n, ...geoRow(nV[n], nC[n], nCat[n], nCatC[n]) });
    for (let k = 0; k < nK; k++) {
      if (!skM[k] && !skC[k]) continue;
      skus.push({ k, sales: skM[k], salesCmp: skC[k], growth: growth(skM[k], skC[k]), share: div(skM[k], catV),
        promoShare: div(skP[k], skT[k]), loyalShare: div(skL[k], skT[k]), price: div(skR[k], skU[k]), priceKg: div(skR[k], skG[k]) });
    }
    const total = { sales: wM, salesCmp: cmpM, growth: growth(wM, cmpM), share: kpi.share, shareCmp: kpi.shareCmp,
      dShare: dShare(kpi.share, kpi.shareCmp) };
    return { win, kpi, chart1, chart2, chart3, chart4, brands, fos, regions, nets, skus, total };
  }

  // ---------- выгрузка в CSV для Excel: «;», десятичная запятая, UTF-8 с BOM ----------
  const EXCEL_MAX = 1048575;
  const r2 = (v, d) => (ok(v) ? Math.round(v * 10 ** d) / 10 ** d : null);
  function csvCell(v) {
    if (v == null || (typeof v === 'number' && !isFinite(v))) return '';
    if (typeof v === 'number') return String(v).replace('.', ',');
    const s = String(v);
    return /[;"\r\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
  }
  function saveCsv(name, parts) {
    const url = URL.createObjectURL(new Blob(['\ufeff', ...parts], { type: 'text/csv;charset=utf-8' }));
    const a = el('a', { href: url, download: name });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 60000);
  }
  function fileName(kind) {
    const span = SPANS.find(s => s[0] === state.span)[1];
    return `X5_${META.contourName}_${kind}_${span}_${per.label[state.end]}.csv`.replace(/[\s\/\\:*?"<>|]+/g, '_');
  }
  // Строки данных за выбранный период со всеми срезами: неделя (месяц) × сеть × регион (ФО) × SKU
  function detailCsv(st) {
    const win = periodWindow(st), inW = new Uint8Array(nP);
    for (const p of win.W) inW[p] = 1;
    const { netM, geoOk, fullM } = masks(st);
    const P = D.P, N = D.N, GEO = D.Geo, K = D.K;
    let count = 0;
    for (let i = 0; i < P.length; i++) if (inW[P[i]] && netM[N[i]] && geoOk[GEO[i]] && fullM[K[i]]) count++;
    if (count > EXCEL_MAX) return { tooMany: count };
    const head = ['Период', 'Начало', 'Конец', 'Сеть', 'ФО', ...(regionLevel ? ['Регион'] : []), 'PLU', 'SKU', 'Бренд',
      'Производитель', 'Категория', 'Продажи, шт', 'Продажи, руб', 'Продажи, кг', 'Промо, шт', 'Промо, руб', 'Промо, кг',
      'По картам, шт', 'По картам, руб', 'По картам, кг', 'Магазины с продажами'];
    const cell = list => list.map(csvCell);
    const perS = per.label.map((l, p) => cell([l, per.start[p], per.end[p]]).join(';'));
    const netS = cell(META.nets), foS = cell(META.fos), regS = cell(META.regions), pluS = cell(META.products.plu);
    const nameS = cell(META.products.name), brandS = cell(META.brands), prodS = cell(META.producers), catS = cell(META.categories);
    const U = D.U, R = D.R, Rk = D.Rk, G = D.G, PU = D.PU, LU = D.LU, S = D.S;
    const promoRub = measureOf('R', PU, D.PR, D.PRk), loyalRub = measureOf('R', LU, D.LR, D.LRk);
    const promoKg = measureOf('G', PU, D.PR, D.PRk, D.PG), loyalKg = measureOf('G', LU, D.LR, D.LRk, D.LG);
    const parts = [cell(head).join(';') + '\r\n'];
    let chunk = [];
    for (let i = 0; i < P.length; i++) {
      const p = P[i], net = N[i], g = GEO[i], k = K[i];
      if (!inW[p] || !netM[net] || !geoOk[g] || !fullM[k]) continue;
      chunk.push([perS[p], netS[net], foS[geoFo[g]] + (regionLevel ? ';' + regS[g] : ''), pluS[k], nameS[k],
        brandS[brandOf[k]], prodS[prodOf[k]], catS[catOf[k]],
        csvCell(U[i]), csvCell(r2(R[i] + Rk[i] * 0.01, 2)), csvCell(r2(G[i] * 0.001, 3)),
        csvCell(PU[i]), csvCell(r2(promoRub(i), 2)), csvCell(r2(promoKg(i), 3)),
        csvCell(LU[i]), csvCell(r2(loyalRub(i), 2)), csvCell(r2(loyalKg(i), 3)), csvCell(S[i])].join(';'));
      if (chunk.length === 20000) { parts.push(chunk.join('\r\n') + '\r\n'); chunk = []; }
    }
    if (chunk.length) parts.push(chunk.join('\r\n') + '\r\n');
    return { parts, count };
  }

  // ---------- интерфейс: срезы ----------
  let renderQueued = false;
  function schedule() {
    if (renderQueued) return;
    renderQueued = true;
    requestAnimationFrame(() => { renderQueued = false; renderAll(); });
  }
  // Список с галочками и поиском для боковой панели. «Все» — когда ничего не отмечено.
  function filterList(label, set, getOptions, onChange) {
    const count = el('span', { class: 'fl-count' });
    const reset = el('button', { class: 'fl-reset', type: 'button', text: 'Сбросить' });
    const search = el('input', { class: 'ms-search', type: 'search', placeholder: 'Поиск', 'aria-label': `Поиск: ${label}` });
    const list = el('div', { class: 'ms-list', role: 'group', 'aria-label': label });
    const wrap = el('section', { class: 'fl' }, el('div', { class: 'fl-head' }, el('span', { class: 'fl-title', text: label }), count, reset), search, list);
    function head() { count.textContent = set.size ? `выбрано ${set.size}` : ''; reset.hidden = set.size === 0; }
    function fill() {
      const all = getOptions(), q = search.value.trim().toLowerCase();
      search.hidden = all.length <= 8;
      list.textContent = '';
      const allBox = el('input', { type: 'checkbox' });
      allBox.checked = set.size === 0;
      allBox.addEventListener('change', () => { set.clear(); refresh(); onChange(); });
      list.append(el('label', { class: 'ms-item ms-all' }, allBox, el('span', { text: 'Все' })));
      for (const o of all) {
        if (q && !o.text.toLowerCase().includes(q)) continue;
        const box = el('input', { type: 'checkbox' });
        box.checked = set.has(o.value);
        box.addEventListener('change', () => {
          if (box.checked) set.add(o.value); else set.delete(o.value);
          allBox.checked = set.size === 0;
          head(); onChange();
        });
        list.append(el('label', { class: 'ms-item' + (o.own ? ' own' : '') }, box, el('span', { text: o.text })));
      }
    }
    function refresh() { fill(); head(); }
    reset.addEventListener('click', () => { set.clear(); refresh(); onChange(); });
    search.addEventListener('input', fill);
    refresh();
    return { wrap, refresh };
  }

  function productOk(k, useProds) {
    if (state.cats.size && !state.cats.has(catOf[k])) return false;
    if (useProds && state.prods.size && !state.prods.has(prodOf[k])) return false;
    return true;
  }
  // Списки производителей и брендов: свои первыми, дальше по алфавиту
  const byName = (a, b) => a.text.localeCompare(b.text, 'ru');
  function availableProducers() {
    const seen = new Uint8Array(nProd);
    for (let k = 0; k < nK; k++) if (productOk(k, false)) seen[prodOf[k]] = 1;
    return META.producers.map((t, i) => ({ value: i, text: t, own: i === META.own })).filter(o => seen[o.value])
      .sort((a, b) => b.own - a.own || byName(a, b));
  }
  function availableBrands() {
    const seen = new Uint8Array(nBrand);
    for (let k = 0; k < nK; k++) if (productOk(k, true)) seen[brandOf[k]] = 1;
    return META.brands.map((t, i) => ({ value: i, text: t, own: isOwnBrand(i) })).filter(o => seen[o.value])
      .sort((a, b) => b.own - a.own || byName(a, b));
  }
  function availableRegions() {
    return META.regions.map((t, i) => ({ value: i, text: t }))
      .filter(o => state.fos.size === 0 || state.fos.has(geoFo[o.value])).sort(byName);
  }
  function prune(set, options) {
    const allowed = new Set(options.map(o => o.value));
    for (const v of [...set]) if (!allowed.has(v)) set.delete(v);
  }

  const controls = {};
  function segmented(label, options, get, set, rerender = true) {
    const wrap = el('div', { class: 'seg', role: 'group', 'aria-label': label });
    const buttons = options.map(o => {
      const b = el('button', { type: 'button', text: o.text, title: o.title || null });
      b.addEventListener('click', () => { set(o.value); refresh(); if (rerender) schedule(); });
      return b;
    });
    function refresh() { buttons.forEach((b, i) => b.setAttribute('aria-pressed', String(get() === options[i].value))); }
    wrap.append(...buttons);
    refresh();
    return { wrap, refresh };
  }
  const theme = () => document.documentElement.dataset.theme || 'light';
  function setTheme(v) {
    document.documentElement.dataset.theme = v;
    try { localStorage.setItem('x5-theme', v); } catch (e) { /* без сохранения — тема просто не запомнится */ }
  }
  // Сверху: период, длина, сравнение, блок метрик, единица, тема
  function buildTopbar() {
    const bar = el('div', { class: 'filters', role: 'region', 'aria-label': 'Период и метрики' });
    const sel = el('select', { class: 'select', 'aria-label': weekly ? 'Неделя, которой заканчивается период' : 'Месяц, которым заканчивается период' });
    for (let p = nP - 1; p >= 0; p--) sel.append(el('option', { value: String(p), text: periodName(p) }));
    sel.addEventListener('change', () => { state.end = Number(sel.value); schedule(); });
    controls.end = { wrap: el('label', { class: 'field' }, el('span', { class: 'ms-label', text: 'До' }), sel),
      refresh: () => { sel.value = String(state.end); } };
    controls.end.refresh();
    controls.span = segmented('Длина периода', SPANS.map(([value, text, title]) => ({ value, text, title })), () => state.span, v => { state.span = v; });
    controls.cmp = segmented('Сравнение', [{ value: 'YA', text: 'vs YA', title: 'С тем же периодом год назад' },
      { value: 'PP', text: 'vs PP', title: 'С предыдущим периодом той же длины' }], () => state.cmp, v => { state.cmp = v; });
    controls.block = segmented('Блок метрик', BLOCKS.map(([value, text, , title]) => ({ value, text, title })), () => state.block, v => { state.block = v; });
    controls.unit = segmented('Единица', MEASURES.map(([value, text, title]) => ({ value, text, title })), () => state.unit, v => { state.unit = v; });
    controls.theme = segmented('Тема', [{ value: 'light', text: '☀ День', title: 'Дневная тема' }, { value: 'dark', text: '☾ Ночь', title: 'Ночная тема' }],
      theme, setTheme, false);
    const reset = el('button', { class: 'btn', type: 'button', text: 'Сбросить всё' });
    reset.addEventListener('click', () => {
      Object.assign(state, { end: nP - 1 }, DEFAULTS);
      for (const k of ['nets', 'fos', 'regs', 'cats', 'prods', 'brands']) state[k].clear();
      syncControls(); schedule();
    });
    bar.append(controls.end.wrap, controls.span.wrap, controls.cmp.wrap, controls.block.wrap, controls.unit.wrap,
      el('span', { class: 'spacer' }), controls.theme.wrap, reset);
    return bar;
  }
  // Сбоку: рынки (сеть, ФО, регион) и продукт (категория, производитель, бренд)
  function buildSide() {
    const side = el('aside', { class: 'side', 'aria-label': 'Срезы' });
    controls.nets = filterList('Сеть', state.nets, () => META.nets.map((t, i) => ({ value: i, text: t })), schedule);
    controls.fos = filterList('ФО', state.fos, () => META.fos.map((t, i) => ({ value: i, text: t })), () => {
      if (regionLevel) { prune(state.regs, availableRegions()); controls.regs.refresh(); }
      schedule();
    });
    if (regionLevel) controls.regs = filterList('Регион', state.regs, availableRegions, schedule);
    controls.cats = filterList('Категория', state.cats, () => META.categories.map((t, i) => ({ value: i, text: t })), () => {
      prune(state.prods, availableProducers()); prune(state.brands, availableBrands());
      controls.prods.refresh(); controls.brands.refresh(); schedule();
    });
    controls.prods = filterList('Производитель', state.prods, availableProducers, () => {
      prune(state.brands, availableBrands()); controls.brands.refresh(); schedule();
    });
    controls.brands = filterList('Бренд', state.brands, availableBrands, schedule);
    side.append(el('h3', { text: 'Рынок' }), controls.nets.wrap, controls.fos.wrap, ...(regionLevel ? [controls.regs.wrap] : []),
      el('h3', { text: 'Продукт' }), controls.cats.wrap, controls.prods.wrap, controls.brands.wrap);
    return side;
  }
  function syncControls() { for (const c of Object.values(controls)) c.refresh(); }

  // ---------- показатели ----------
  const kpiBox = el('section', { class: 'kpis', 'aria-label': 'Ключевые показатели' });
  function setDelta(node, d, tone) {
    node.textContent = '';
    if (!d) { node.className = 'delta none'; node.textContent = `к ${state.cmp}: нет данных`; return; }
    node.className = 'delta ' + (d.dir === 0 ? 'neutral' : tone(d.dir > 0));
    if (d.dir) node.append(el('span', { 'aria-hidden': 'true', text: d.dir > 0 ? '▲ ' : '▼ ' }));
    node.append(document.createTextNode(`${d.text} к ${state.cmp}`));
  }
  const goodUp = up => (up ? 'good' : 'bad');
  const neutralTone = () => 'neutral';
  const dPct = (a, b) => (ok(a) && ok(b) && b ? delta(a / b - 1, 'pct') : null);
  const dPP = (a, b) => (ok(a) && ok(b) ? delta(100 * (a - b), 'pp') : null);
  const KPI = {
    sales: [() => `Продажи${bwSp()}, ${mu()}`, k => fmt.compact(k.sales), k => delta(growth(k.sales, k.salesCmp), 'pct'), goodUp],
    share: [() => `Доля в категории${bw() ? ` (${bw()})` : ''}, ${mu()}`, k => fmt.pct(k.share), k => dPP(k.share, k.shareCmp), goodUp],
    promoShare: [() => (state.block === 'P' ? `Доля промо в продажах, ${mu()}` : `Доля промо, ${mu()}`), k => fmt.pct(k.promoShare),
      k => dPP(k.promoShare, k.promoShareCmp), neutralTone],
    loyalShare: [() => (state.block === 'L' ? `Доля по картам в продажах, ${mu()}` : `Доля по картам, ${mu()}`), k => fmt.pct(k.loyalShare),
      k => dPP(k.loyalShare, k.loyalShareCmp), neutralTone],
    price: [() => `Цена${bwSp()} за шт`, k => fmt.rub(k.price), k => dPct(k.price, k.priceCmp), neutralTone],
    priceKg: [() => `Цена${bwSp()} за кг/л`, k => fmt.rub(k.priceKg), k => dPct(k.priceKg, k.priceKgCmp), neutralTone],
    depth: [() => 'Глубина промо', k => fmt.pct(k.depth), k => dPP(k.depth, k.depthCmp), neutralTone],
    tdp: [() => 'TDP средняя', k => fmt.tdp(k.tdp), k => (ok(k.tdp) && ok(k.tdpCmp) ? delta(100 * (k.tdp - k.tdpCmp), 'pt') : null), goodUp],
  };
  const KPI_SETS = { T: ['sales', 'share', 'promoShare', 'loyalShare', 'price', 'priceKg', 'tdp'],
    P: ['sales', 'share', 'promoShare', 'price', 'priceKg', 'depth'], L: ['sales', 'share', 'loyalShare', 'price', 'priceKg'] };
  function renderKpis(r) {
    const k = r.kpi;
    kpiBox.textContent = '';
    for (const key of KPI_SETS[state.block]) {
      const [label, value, change, tone] = KPI[key];
      const d = el('div', { class: 'delta' });
      setDelta(d, k.hasCmp ? change(k) : null, tone);
      kpiBox.append(el('div', { class: 'tile' }, el('div', { class: 'label', text: label() }), el('div', { class: 'value', text: value(k) }), d));
    }
  }

  // ---------- карточки и табличный вид ----------
  function card(title, note) {
    const body = el('div', { class: 'chart' });
    const table = el('div', { class: 'table-wrap', hidden: true });
    const tip = el('div', { class: 'tip', hidden: true, role: 'status' });
    const legend = el('div', { class: 'legend' });
    const toggle = el('button', { class: 'ghost', type: 'button', text: 'Таблица', 'aria-pressed': 'false' });
    const h2 = el('h2', { text: title });
    const node = el('section', { class: 'card' },
      el('div', { class: 'card-head' }, el('div', {}, h2, note ? el('div', { class: 'note', text: note }) : null), toggle),
      body, legend, table, tip);
    const c = { node, body, table, tip, legend, toggle, rows: null, head: null, draw: null, setTitle: s => { h2.textContent = s; } };
    toggle.addEventListener('click', () => {
      const show = table.hidden;
      table.hidden = !show; body.hidden = show; legend.hidden = show;
      toggle.textContent = show ? 'График' : 'Таблица';
      toggle.setAttribute('aria-pressed', String(show));
      if (show) fillTable(c);
    });
    return c;
  }
  function fillTable(c) {
    c.table.textContent = '';
    if (!c.rows || !c.rows.length) { c.table.append(el('div', { class: 'empty', text: 'Нет данных для выбранных срезов' })); return; }
    const t = el('table', { class: 'data' });
    t.append(el('thead', {}, el('tr', {}, ...c.head.map((h, i) => el('th', { class: i ? 'num' : '', text: h })))));
    const tb = el('tbody');
    for (const row of c.rows) tb.append(el('tr', {}, ...row.map((v, i) => el('td', { class: i ? 'num' : 'name', text: v }))));
    t.append(tb);
    c.table.append(t);
  }
  function showTip(c, x, y, title, rows) {
    const tip = c.tip;
    tip.textContent = '';
    tip.append(el('div', { class: 't-title', text: title }));
    for (const r of rows) {
      tip.append(el('div', { class: 't-row' },
        r.color ? el('span', { class: r.rect ? 'key-rect' : 'key-line', style: `background:${r.color}` }) : null,
        el('span', { class: 't-val', text: r.value }), el('span', { class: 't-name', text: r.name })));
    }
    tip.hidden = false;
    const box = c.node.getBoundingClientRect();
    const w = tip.offsetWidth, h = tip.offsetHeight;
    let left = x - box.left + 14, top = y - box.top - h - 10;
    if (left + w > box.width - 8) left = x - box.left - w - 14;
    if (top < 8) top = y - box.top + 16;
    tip.style.left = Math.max(8, left) + 'px';
    tip.style.top = top + 'px';
  }
  const hideTip = c => { c.tip.hidden = true; };
  const color = slot => (slot >= 0 ? `var(--s${slot + 1})` : 'var(--other)');
  const prodColor = pr => color(META.slots[pr]);

  function niceTicks(min, max, count) {
    const span = max - min || 1;
    const raw = span / count, mag = Math.pow(10, Math.floor(Math.log10(raw)));
    const step = [1, 2, 5, 10].map(m => m * mag).find(s => s >= raw) || raw;
    const lo = Math.floor(min / step) * step, hi = Math.ceil(max / step) * step, out = [];
    for (let v = lo; v <= hi + step / 2; v += step) out.push(Math.round(v / step) * step);
    return out;
  }

  function legendFor(c, items, rect) {
    c.legend.textContent = '';
    for (const it of items) c.legend.append(el('span', {}, el('span', { class: rect ? 'key-rect' : 'key-line', style: `background:${it.color}` }), document.createTextNode(it.name)));
  }

  function lineChart(c, cfg) {
    const host = c.body;
    host.textContent = '';
    const n = cfg.labels.length;
    if (!n || !cfg.series.some(s => s.values.some(ok))) { host.append(el('div', { class: 'empty', text: 'Нет данных для выбранных срезов' })); c.legend.textContent = ''; return; }
    const W = Math.max(320, host.clientWidth || 600), H = 262, m = { l: 58, r: 84, t: 24, b: 28 };
    const iw = W - m.l - m.r, ih = H - m.t - m.b;
    let lo = 0, hi = 0;
    for (const s of cfg.series) for (const v of s.values) if (ok(v)) { hi = Math.max(hi, v); lo = Math.min(lo, v); }
    if (hi === lo) hi = lo + 1;
    const ticks = niceTicks(lo, hi, 4), y0 = ticks[0], y1 = ticks[ticks.length - 1], step = ticks[1] - ticks[0];
    const X = i => m.l + (n === 1 ? iw / 2 : (i * iw) / (n - 1));
    const Y = v => m.t + ih - ((v - y0) / (y1 - y0)) * ih;
    const root = svg('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': cfg.aria });
    if (cfg.band && cfg.band[0] > 0) {
      const half = n > 1 ? iw / (n - 1) / 2 : 0;
      const x0 = Math.max(m.l, X(cfg.band[0]) - half), x1 = Math.min(m.l + iw, X(cfg.band[1]) + half);
      root.append(svg('rect', { x: x0, y: m.t, width: Math.max(1, x1 - x0), height: ih, fill: 'var(--wash)' }));
      if (cfg.bandLabel) {   // подпись над областью графика, чтобы не ложилась на линии
        const wide = x1 - x0 >= 48;
        const bt = svg('text', { x: wide ? x0 + 4 : x1, y: m.t - 8, 'text-anchor': wide ? 'start' : 'end' });
        bt.textContent = cfg.bandLabel;
        root.append(bt);
      }
    }
    for (const t of ticks) {
      root.append(svg('line', { x1: m.l, x2: W - m.r, y1: Y(t), y2: Y(t), stroke: t === 0 ? 'var(--axis)' : 'var(--grid)', 'stroke-width': 1 }));
      const tx = svg('text', { x: m.l - 8, y: Y(t) + 4, 'text-anchor': 'end' });
      tx.textContent = cfg.axisFmt(t, step);
      root.append(tx);
    }
    const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor(iw / 46))));
    for (let i = 0; i < n; i += every) {
      const tx = svg('text', { x: X(i), y: H - 8, 'text-anchor': 'middle' });
      tx.textContent = cfg.xLabel(i);
      root.append(tx);
    }
    const ends = [], lastOf = s => { let l = -1; s.values.forEach((v, i) => { if (ok(v)) l = i; }); return l; };
    // первая серия рисуется последней, чтобы быть сверху
    [...cfg.series].reverse().forEach(s => {
      let d = '', pen = false;
      s.values.forEach((v, i) => {
        if (!ok(v)) { pen = false; return; }
        d += (pen ? 'L' : 'M') + X(i).toFixed(1) + ' ' + Y(v).toFixed(1);
        pen = true;
      });
      root.append(svg('path', { d, fill: 'none', stroke: s.color, 'stroke-width': 2, 'stroke-linejoin': 'round', 'stroke-linecap': 'round' }));
      const last = lastOf(s);
      if (last >= 0) root.append(svg('circle', { cx: X(last), cy: Y(s.values[last]), r: 4, fill: s.color, stroke: 'var(--surface)', 'stroke-width': 2 }));
    });
    for (const s of cfg.series) {
      const last = lastOf(s);
      if (last >= 0) ends.push({ y: Y(s.values[last]), x: X(last), text: cfg.fmt(s.values[last]) });
    }
    ends.sort((a, b) => a.y - b.y);
    let prevY = -1e9;
    for (const e of ends) {
      if (e.y - prevY < 14) continue;
      const tx = svg('text', { x: e.x + 8, y: e.y + 4, class: 'endlabel' });
      tx.textContent = e.text;
      root.append(tx);
      prevY = e.y;
    }
    const cross = svg('line', { y1: m.t, y2: m.t + ih, stroke: 'var(--muted)', 'stroke-width': 1, 'stroke-dasharray': '3 3', visibility: 'hidden' });
    const dots = cfg.series.map(s => svg('circle', { r: 4, fill: s.color, stroke: 'var(--surface)', 'stroke-width': 2, visibility: 'hidden' }));
    root.append(cross, ...dots);
    const hit = svg('rect', { x: m.l - 10, y: m.t, width: iw + 20, height: ih, fill: 'transparent', tabindex: 0, 'aria-label': `${cfg.aria}. Стрелки влево и вправо — по периодам` });
    root.append(hit);
    let cur = -1;
    function show(i, cx, cy) {
      cur = i;
      cross.setAttribute('x1', X(i)); cross.setAttribute('x2', X(i)); cross.setAttribute('visibility', 'visible');
      const rows = [];
      cfg.series.forEach((s, j) => {
        const v = s.values[i];
        if (ok(v)) { dots[j].setAttribute('cx', X(i)); dots[j].setAttribute('cy', Y(v)); dots[j].setAttribute('visibility', 'visible'); }
        else dots[j].setAttribute('visibility', 'hidden');
        rows.push({ v, value: ok(v) ? cfg.fmt(v) : '—', name: s.name, color: s.color });
      });
      if (cfg.sortTip) rows.sort((a, b) => (ok(b.v) ? b.v : -Infinity) - (ok(a.v) ? a.v : -Infinity));
      showTip(c, cx, cy, cfg.tipTitle(i), rows);
    }
    function hide() { cross.setAttribute('visibility', 'hidden'); dots.forEach(d => d.setAttribute('visibility', 'hidden')); hideTip(c); }
    hit.addEventListener('pointermove', e => {
      const b = root.getBoundingClientRect();
      const x = ((e.clientX - b.left) / b.width) * W;
      const i = n === 1 ? 0 : Math.max(0, Math.min(n - 1, Math.round(((x - m.l) / iw) * (n - 1))));
      show(i, e.clientX, e.clientY);
    });
    hit.addEventListener('pointerleave', hide);
    hit.addEventListener('blur', hide);
    hit.addEventListener('keydown', e => {
      if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
      e.preventDefault();
      const i = Math.max(0, Math.min(n - 1, (cur < 0 ? n - 1 : cur + (e.key === 'ArrowRight' ? 1 : -1))));
      const b = root.getBoundingClientRect();
      show(i, b.left + (X(i) / W) * b.width, b.top + m.t + 20);
    });
    host.append(root);
    legendFor(c, cfg.series.map(s => ({ name: s.name, color: s.color })), false);
  }

  function barChart(c, cfg) {
    const host = c.body;
    host.textContent = '';
    if (!cfg.rows.length) { host.append(el('div', { class: 'empty', text: 'Нет данных для выбранных срезов' })); c.legend.textContent = ''; return; }
    let max = 0;
    for (const r of cfg.rows) for (const v of r.values) if (ok(v)) max = Math.max(max, v);
    max = max || 1;
    const grid = el('div', { class: 'bars' });
    for (const r of cfg.rows) {
      const cell = el('div', { class: 'bar-cell', tabindex: 0, 'aria-label': `${r.label}: ` + cfg.series.map((s, j) => `${s.name} ${cfg.fmt(r.values[j])}`).join(', ') });
      r.values.forEach((v, j) => {
        const w = ok(v) ? Math.max(0, (v / max) * 82) : 0;
        cell.append(el('div', { class: 'bar-row' },
          el('div', { class: 'bar' + (cfg.series.length > 1 ? ' thin' : ''), style: `width:${w}%;background:${r.colors ? r.colors[j] : cfg.series[j].color}` }),
          el('span', { class: 'bar-val', text: cfg.fmt(v) })));
      });
      const tipRows = () => cfg.series.map((s, j) => ({ value: cfg.fmt(r.values[j]), name: s.name, color: r.colors ? r.colors[j] : s.color, rect: true }));
      cell.addEventListener('pointermove', e => showTip(c, e.clientX, e.clientY, r.label, tipRows()));
      cell.addEventListener('pointerleave', () => hideTip(c));
      cell.addEventListener('focus', () => { const b = cell.getBoundingClientRect(); showTip(c, b.left + b.width / 2, b.top, r.label, tipRows()); });
      cell.addEventListener('blur', () => hideTip(c));
      grid.append(el('div', { class: 'bar-label' + (r.own ? ' own' : ''), title: r.label, text: r.label }), cell);
    }
    host.append(grid);
    if (cfg.series.length > 1) legendFor(c, cfg.series.map(s => ({ name: s.name, color: s.color })), true);
    else if (cfg.legend) legendFor(c, cfg.legend, true);
    else c.legend.textContent = '';
  }

  // ---------- таблицы ----------
  // Вид колонки определяет и показ, и выгрузку: text, sales, pct (доля), dpct (прирост), pp (изменение доли), rub, idx, tdp
  const titleOf = c => (typeof c.title === 'function' ? c.title() : c.title);
  const scaled = (x, mul, d) => (ok(x) ? r2(x * mul, d) : null);
  function exportColumns(c) {
    const t = titleOf(c), v = r => r[c.key];
    if (c.exp) return c.exp;
    switch (c.kind) {
      case 'sales': return [[t, r => scaled(v(r), 1, salesDigits())]];
      case 'pct': case 'dpct': return [[t + ', %', r => scaled(v(r), 100, 2)]];
      case 'pp': return [[t + ', п.п.', r => scaled(v(r), 1, 2)]];
      case 'rub': return [[t + ', руб', r => scaled(v(r), 1, 2)]];
      case 'idx': return [[t, r => scaled(v(r), 1, 3)]];
      case 'tdp': return [[t, r => scaled(v(r), 100, 1)]];
      default: return [[t, v]];
    }
  }
  function dataTable(title, note, columns, opts) {
    const wrapT = el('div', { class: 'table-wrap' });
    const more = el('button', { class: 'ghost more', type: 'button', hidden: true });
    const csvBtn = el('button', { class: 'ghost', type: 'button', text: 'CSV', title: 'Скачать таблицу целиком для Excel' });
    const node = el('section', { class: 'card' + (opts.full ? ' full' : '') },
      el('div', { class: 'card-head' }, el('div', {}, el('h2', { text: title }), note ? el('div', { class: 'note', text: note }) : null),
        el('div', { class: 'card-actions' }, csvBtn)), wrapT, more);
    const t = { node, rows: [], total: null, sortKey: opts.sortKey, dir: -1, expanded: false };
    more.addEventListener('click', () => { t.expanded = !t.expanded; draw(); });
    function cellOf(col, row) {
      if (col.render) return col.render(row);
      const v = row[col.key];
      if (col.kind === 'dpct' || col.kind === 'pp') {
        const d = delta(v, col.kind === 'dpct' ? 'pct' : 'pp');
        if (!d) return el('td', { class: 'num none', text: '—' });
        if (!d.dir) return el('td', { class: 'num neutral', text: d.text });
        return el('td', { class: 'num ' + (d.dir > 0 ? 'good' : 'bad') }, el('span', { 'aria-hidden': 'true', text: d.dir > 0 ? '▲ ' : '▼ ' }), document.createTextNode(d.text));
      }
      const f = { sales: fmt.sales, pct: fmt.pct, rub: fmt.rub, idx: fmt.idx, tdp: fmt.tdp }[col.kind];
      return el('td', { class: f ? 'num' : 'name', text: f ? f(v) : String(v ?? '') });
    }
    const visible = () => columns.filter(c => !c.when || c.when());
    function sortedRows() {
      const col = visible().find(x => x.key === t.sortKey) || visible().find(x => x.kind === 'sales');
      const val = r => { const v = r[col.key]; return ok(v) ? v : (typeof v === 'string' ? v : -Infinity); };
      return [...t.rows].sort((a, b) => {
        const x = val(a), y = val(b);
        return (typeof x === 'string' ? x.localeCompare(y, 'ru') : x - y) * t.dir;
      });
    }
    function draw() {
      wrapT.textContent = '';
      if (!t.rows.length) { wrapT.append(el('div', { class: 'empty', text: 'Нет данных для выбранных срезов' })); more.hidden = true; return; }
      const rows = sortedRows();
      const shown = t.expanded || !opts.limit ? rows : rows.slice(0, opts.limit);
      const table = el('table', { class: 'data' });
      const head = el('tr');
      const cols = visible();
      for (const cdef of cols) {
        const isNum = cdef.kind !== 'text';
        const th = el('th', { class: isNum ? 'num' : '', text: titleOf(cdef), scope: 'col', tabindex: 0 });
        if (cdef.key === t.sortKey) th.setAttribute('aria-sort', t.dir > 0 ? 'ascending' : 'descending');
        const sortBy = () => { if (t.sortKey === cdef.key) t.dir = -t.dir; else { t.sortKey = cdef.key; t.dir = isNum ? -1 : 1; } draw(); };
        th.addEventListener('click', sortBy);
        th.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); sortBy(); } });
        head.append(th);
      }
      table.append(el('thead', {}, head));
      const tb = el('tbody');
      for (const r of shown) tb.append(el('tr', {}, ...cols.map(cdef => cellOf(cdef, r))));
      if (t.total) tb.append(el('tr', { class: 'total' }, ...cols.map(cdef => cellOf(cdef, t.total))));
      table.append(tb);
      wrapT.append(table);
      if (opts.limit && rows.length > opts.limit) {
        more.hidden = false;
        more.textContent = t.expanded ? 'Свернуть' : `Показать все (${rows.length})`;
      } else more.hidden = true;
    }
    t.csv = () => {
      const cols = visible().flatMap(exportColumns);
      const lines = [cols.map(c => csvCell(c[0])).join(';')];
      for (const r of [...sortedRows(), ...(t.total ? [t.total] : [])]) lines.push(cols.map(c => csvCell(c[1](r))).join(';'));
      return [lines.join('\r\n') + '\r\n'];
    };
    csvBtn.addEventListener('click', () => saveCsv(fileName(opts.file), t.csv()));
    t.set = (rows, total) => { t.rows = rows; t.total = total || null; draw(); };
    return t;
  }

  // ---------- сборка страницы ----------
  app.textContent = '';
  const periodText = el('span');
  const exportStatus = el('span', { class: 'status', role: 'status' });
  const exportBtn = el('button', { class: 'ghost', type: 'button', text: 'Скачать данные',
    title: `Строки по ${weekly ? 'неделям' : 'месяцам'}, сетям, ${regionLevel ? 'регионам' : 'ФО'} и SKU за выбранный период и срезы` });
  exportBtn.addEventListener('click', () => {
    exportStatus.textContent = 'Готовлю файл…';
    setTimeout(() => {
      const res = detailCsv(state);
      if (res.tooMany) {
        exportStatus.textContent = `Строк ${nf0.format(res.tooMany)} — больше, чем помещается на лист Excel. Сузьте период или срезы.`;
        return;
      }
      saveCsv(fileName('данные'), res.parts);
      exportStatus.textContent = `Выгружено строк: ${nf0.format(res.count)}`;
    }, 30);
  });
  const periodLine = el('div', { class: 'period' }, periodText, el('span', { class: 'spacer' }), exportStatus, exportBtn);
  const main = el('main', { class: 'main' });
  app.append(buildTopbar(), el('div', { class: 'layout' }, buildSide(), main));
  main.append(periodLine, kpiBox);
  const periodWord = weekly ? 'неделям' : 'месяцам';
  const lastWord = weekly ? 'Последние 52 недели' : 'Последние 12 месяцев';
  const c1 = card('', `${lastWord} и YA · выбранный период выделен фоном`);
  const c2 = card('', `${lastWord} · топ-5 за выбранный период · не зависит от срезов производителя и бренда`);
  const c3 = card('', 'Топ-10 за выбранный период · не зависит от срезов производителя и бренда');
  const c4 = card('', 'Топ-10 производителей · без срезов производителя и бренда');
  main.append(el('div', { class: 'grid2' }, c1.node, c2.node), el('div', { class: 'grid2' }, c3.node, c4.node));
  const name = k => META.products.name[k];
  const ownMark = (text, own) => (own ? el('span', { class: 'own', text }) : document.createTextNode(text));
  const salesTitle = () => `Продажи${bwSp()}, ${mu()}`;
  const shareTitleCat = () => `Доля в категории${bw() ? ` (${bw()})` : ''}`;
  const growthTitle = () => `Прирост к ${state.cmp}`, shareTitle = () => `Изменение доли к ${state.cmp}`;
  const promoCol = { key: 'promoShare', kind: 'pct', title: () => (state.block === 'P' ? 'Доля промо в продажах' : 'Доля промо'), when: () => state.block !== 'L' };
  const loyalCol = { key: 'loyalShare', kind: 'pct', title: () => (state.block === 'L' ? 'Доля по картам в продажах' : 'Доля по картам'), when: () => state.block !== 'P' };
  const priceCols = [{ key: 'price', kind: 'rub', title: () => `Цена${bwSp()} за шт` }, { key: 'priceKg', kind: 'rub', title: () => `Цена${bwSp()} за кг/л` }];
  const tBrands = dataTable('Бренды', 'Доля и индекс цены — от категорий из среза «Категория» в тех же сетях, регионах и периоде', [
    { key: 'name', kind: 'text', title: 'Бренд', render: r => el('td', { class: 'name', title: r.name }, ownMark(r.name, r.own)) },
    { key: 'sales', kind: 'sales', title: salesTitle },
    { key: 'growth', kind: 'dpct', title: growthTitle },
    { key: 'share', kind: 'pct', title: shareTitleCat },
    { key: 'dShare', kind: 'pp', title: shareTitle },
    promoCol, loyalCol, ...priceCols,
    { key: 'index', kind: 'idx', title: 'Индекс цены за кг' },
    { key: 'depth', kind: 'pct', title: 'Глубина промо', when: () => state.block === 'P' },
    { key: 'tdp', kind: 'tdp', title: 'TDP средняя', when: () => state.block === 'T' },
  ], { sortKey: 'sales', limit: 20, full: true, file: 'бренды' });
  main.append(tBrands.node);
  const geoColumns = extra => [...extra,
    { key: 'sales', kind: 'sales', title: salesTitle },
    { key: 'growth', kind: 'dpct', title: growthTitle },
    { key: 'share', kind: 'pct', title: shareTitleCat },
    { key: 'dShare', kind: 'pp', title: shareTitle },
  ];
  const tFo = dataTable('Федеральные округа', 'Доля — в категории своего округа', geoColumns([{ key: 'name', kind: 'text', title: 'ФО' }]),
    { sortKey: 'sales', file: 'ФО' });
  const tNets = dataTable('Сети', 'Доля — в категории своей сети', geoColumns([{ key: 'name', kind: 'text', title: 'Сеть' }]),
    { sortKey: 'sales', file: 'сети' });
  const tRegions = regionLevel ? dataTable('Регионы', 'Доля — в категории своего региона',
    geoColumns([{ key: 'name', kind: 'text', title: 'Регион' }, { key: 'foName', kind: 'text', title: 'ФО' }]),
    { sortKey: 'sales', limit: 20, full: true, file: 'регионы' }) : null;
  const tSku = dataTable('SKU', 'Доля — в категориях из среза «Категория» в тех же сетях, регионах и периоде', [
    { key: 'name', kind: 'text', title: 'SKU', exp: [['SKU', r => r.name], ['PLU', r => r.plu]],
      render: r => el('td', { class: 'name', title: `${r.name} · PLU ${r.plu}` }, document.createTextNode(r.name), el('span', { class: 'plu', text: r.plu })) },
    { key: 'brandName', kind: 'text', title: 'Бренд' },
    { key: 'sales', kind: 'sales', title: salesTitle },
    { key: 'growth', kind: 'dpct', title: growthTitle },
    { key: 'share', kind: 'pct', title: shareTitleCat },
    promoCol, { ...loyalCol, when: () => state.block === 'L' }, ...priceCols,
  ], { sortKey: 'sales', limit: 20, full: true, file: 'SKU' });
  main.append(el('div', { class: 'grid2' }, tFo.node, tNets.node), ...(tRegions ? [tRegions.node] : []), tSku.node);
  main.append(el('footer', {},
    el('p', { text: `Источник: X5 Dialog, отчёт «Анализ трендов», выгрузка из ClickHouse. ${regionLevel ? 'Данные по регионам X5, ФО — по справочнику регионов.' : 'Регионы сведены в федеральные округа.'}` }),
    el('p', { text: weekly
      ? 'Период задаётся последней неделей («До») и длиной: Нед — одна неделя; L4W и L13W — последние 4 и 13 недель; MTD, QTD и YTD — с начала месяца, квартала или года до выбранной недели, месяцы и кварталы по 4-4-5; MAT — последние 52 недели.'
      : 'Период задаётся последним месяцем («До») и длиной: Мес — один месяц; L3M — последние 3 месяца; QTD и YTD — с начала квартала или года до выбранного месяца; MAT — последние 12 месяцев.' }),
    el('p', { text: weekly
      ? 'YA — те же недели ISO годом раньше; у 53-й недели пары нет, это видно в строке периода. PP — столько же недель сразу перед выбранным периодом.'
      : 'YA — те же месяцы годом раньше. PP — столько же месяцев сразу перед выбранным периодом.' }),
    el('p', { text: 'Блоки метрик как в X5: Тотал — все продажи, Промо — продажи по промо, Лояльные — продажи покупателям с картами лояльности X5. Продажи, доли в категории, прирост и цены считаются внутри выбранного блока: например, в блоке «Промо» доля в категории — промо-продажи бренда к промо-продажам категории. Доля промо и доля по картам — от всех продаж. Единица — штуки, рубли или килограммы (литры).' }),
    el('p', { text: 'Доля в категории — продажи выбранного к продажам категорий из среза «Категория» в тех же сетях, регионах и периоде; срезы производителя и бренда на знаменатель не влияют. В таблицах ФО, регионов и сетей доля считается внутри своего округа, региона или сети. Изменение доли — в процентных пунктах к периоду сравнения.' }),
    el('p', { text: 'Взвешенная глубина промо — на сколько промо-цена ниже обычной цены тех же SKU в той же сети, регионе и периоде, взвешено по промо-штукам. Строки, где SKU продавался только по промо, пропускаются.' }),
    el('p', { text: `TDP — сумма нумерических дистрибуций SKU в пунктах: для каждого SKU берётся доля магазинов с его продажами среди магазинов с продажами всей категории БП, в процентах, и доли складываются. TDP 250 — в среднем 2,5 SKU в магазине. TDP средняя — среднее по ${periodWord} периода. Магазин засчитывается, только если в нём была продажа. В Excel та же мера показана в процентах: 250% там — это 250 здесь.` }),
    el('p', { text: 'Индекс цены за кг — цена за кг бренда к цене за кг категории в том же блоке; 1,00 — на уровне категории. Чеки, покупатели и взвешенная дистрибуция здесь не показаны: для брендов из данных X5 их честно не посчитать.' }),
    el('p', { text: `Выгрузка: кнопка CSV у таблицы сохраняет её целиком, «Скачать данные» — строки по ${periodWord}, сетям, ${regionLevel ? 'регионам' : 'ФО'} и SKU за выбранный период и срезы со всеми метриками. Файлы открываются в Excel: разделитель «;», десятичная запятая.` }),
    el('p', { id: 'perf', class: 'sub' })));

  function renderPeriod(win) {
    const spanText = SPANS.find(s => s[0] === state.span)[1];
    const range = ps => (!ps.length ? 'нет данных' : ps.length === 1 ? per.label[ps[0]] : `${per.label[ps[0]]} – ${per.label[ps[ps.length - 1]]}`);
    periodText.textContent = '';
    periodText.append(el('b', { text: spanText }), document.createTextNode(` · ${range(win.W)} · ${win.W.length} ${UNIT} · ${state.cmp}: ${range(win.C)}`));
    const warn = [];
    if (win.missing > 0) warn.push(`в периоде нет данных за ${win.missing} ${UNIT}`);
    if (win.cmpMissing > 0) warn.push(state.cmp === 'YA' ? `у ${win.cmpMissing} ${UNIT} нет пары в YA` : `в PP нет данных за ${win.cmpMissing} ${UNIT}`);
    if (warn.length) periodText.append(el('span', { class: 'warn', text: ' · ⚠ ' + warn.join(', ') }));
    exportStatus.textContent = '';
  }

  let last = null;
  function drawCharts(r) {
    const ax = r.chart1.axis, labels = ax.map(axisLabel), u = mu(), b = bw();
    const xLabel = i => labels[i].short;
    const tipTitle = i => labels[i].long;
    const bandLabel = SPANS.find(s => s[0] === state.span)[1];
    const inner = b ? `${b}, ${u}` : u;
    c1.setTitle(`Продажи${bwSp()}, ${u} по ${periodWord}`);
    c2.setTitle(`Доля в категории (${inner}) по ${periodWord}: топ-5 производителей`);
    c3.setTitle(`Доля производителей в категории (${inner})`);
    c1.rows = ax.map((a, i) => [labels[i].long, fmt.sales(r.chart1.cur[i]), fmt.sales(r.chart1.ya[i])]);
    c1.head = ['Период', `Продажи${bwSp()}, ${u}`, `YA, ${u}`];
    lineChart(c1, { labels, xLabel, tipTitle, band: r.chart1.band, bandLabel, fmt: fmt.compact, axisFmt: v => fmt.compact(v), sortTip: false,
      aria: `Продажи${bwSp()} (${u}) по ${periodWord} и YA`,
      series: [{ name: 'Продажи', color: 'var(--s1)', values: r.chart1.cur }, { name: 'YA', color: 'var(--other)', values: r.chart1.ya }] });
    c2.rows = ax.map((a, i) => [labels[i].long, ...r.chart2.series.map(s => fmt.pct(s.values[i]))]);
    c2.head = ['Период', ...r.chart2.series.map(s => META.producers[s.prod])];
    const pctAxis = (v, step) => { const d = step * 100 >= 0.999 ? 0 : step * 100 >= 0.0999 ? 1 : 2;
      return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: d, maximumFractionDigits: d }).format(v * 100) + '%'; };
    lineChart(c2, { labels, xLabel, tipTitle, band: r.chart1.band, bandLabel, fmt: fmt.pct, axisFmt: pctAxis, sortTip: true,
      aria: 'Доля производителей в категории по периодам',
      series: r.chart2.series.map(s => ({ name: META.producers[s.prod], color: prodColor(s.prod), values: s.values })) });
    c3.rows = r.chart3.map(x => [META.producers[x.prod], fmt.pct(x.share)]);
    c3.head = ['Производитель', `Доля в категории (${inner})`];
    barChart(c3, { fmt: fmt.pct, series: [{ name: `Доля в категории (${inner})`, color: 'var(--other)' }],
      rows: r.chart3.map(x => ({ label: META.producers[x.prod], values: [x.share], own: x.prod === META.own,
        colors: [x.prod === META.own ? 'var(--s1)' : 'var(--other)'] })),
      legend: META.own >= 0 ? [{ name: META.producers[META.own], color: 'var(--s1)' }, { name: 'Другие производители', color: 'var(--other)' }] : null });
    if (state.block === 'L') {
      c4.setTitle('Доля по картам по производителям');
      c4.rows = r.chart4.map(x => [META.producers[x.prod], fmt.pct(x.loyalShare)]);
      c4.head = ['Производитель', `Доля по картам в продажах, ${u}`];
      barChart(c4, { fmt: fmt.pct, series: [{ name: `Доля по картам в продажах, ${u}`, color: 'var(--s3)' }],
        rows: r.chart4.map(x => ({ label: META.producers[x.prod], values: [x.loyalShare], own: x.prod === META.own })),
        legend: [{ name: `Доля по картам в продажах, ${u}`, color: 'var(--s3)' }] });
    } else {
      c4.setTitle('Промо по производителям');
      c4.rows = r.chart4.map(x => [META.producers[x.prod], fmt.pct(x.promoShare), fmt.pct(x.depth)]);
      c4.head = ['Производитель', `Доля промо в продажах, ${u}`, 'Глубина промо'];
      barChart(c4, { fmt: fmt.pct, series: [{ name: `Доля промо в продажах, ${u}`, color: 'var(--s1)' }, { name: 'Глубина промо, взвешенная', color: 'var(--s2)' }],
        rows: r.chart4.map(x => ({ label: META.producers[x.prod], values: [x.promoShare, x.depth], own: x.prod === META.own })) });
    }
    for (const c of [c1, c2, c3, c4]) if (!c.table.hidden) fillTable(c);
  }
  const tables = () => ({ brands: tBrands, fos: tFo, nets: tNets, regions: tRegions, skus: tSku });
  function renderAll() {
    const t0 = performance.now();
    const r = compute(state);
    last = r;
    renderPeriod(r.win);
    renderKpis(r);
    drawCharts(r);
    tBrands.set(r.brands.map(b => ({ ...b, name: META.brands[b.brand], own: isOwnBrand(b.brand) })));
    tFo.set(r.fos.map(f => ({ ...f, name: META.fos[f.fo] })), { ...r.total, name: 'Итого' });
    tNets.set(r.nets.map(x => ({ ...x, name: META.nets[x.net] })), { ...r.total, name: 'Итого' });
    if (tRegions) tRegions.set(r.regions.map(x => ({ ...x, name: META.regions[x.geo], foName: META.fos[x.fo] })), { ...r.total, name: 'Итого', foName: '' });
    tSku.set(r.skus.map(s => ({ ...s, name: name(s.k), plu: META.products.plu[s.k], brandName: META.brands[brandOf[s.k]] })));
    document.getElementById('perf').textContent = `Строк данных: ${nf0.format(D.P.length)} · пересчёт ${Math.round(performance.now() - t0)} мс`;
    document.body.dataset.ready = '1';
  }
  let resizeTimer = 0;
  new ResizeObserver(() => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => last && drawCharts(last), 120); }).observe(main);
  window.X5 = {
    meta: META, state, compute, periodWindow, fileName, setTheme,
    set(patch) {
      for (const [k, v] of Object.entries(patch)) state[k] = v instanceof Array ? new Set(v) : v;
      syncControls();
      renderAll();
      return last;
    },
    tableCsv: name => tables()[name].csv().join(''),
    detailCsv: () => { const res = detailCsv(state); return res.tooMany ? { tooMany: res.tooMany } : { count: res.count, text: res.parts.join('') }; },
  };
  renderAll();
})();
</script>
</body>
</html>
'''

if __name__ == "__main__":
    main()
