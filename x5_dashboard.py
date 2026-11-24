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
    if "Продуктовая линия" not in products:
        products["Продуктовая линия"] = ""
    for c in ["Категория", "Производитель (корректный)", "Бренд (корректный)", "Название", "Продуктовая линия"]:
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
    lines = sorted(products["Продуктовая линия"].unique())
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
        "categories": categories, "producers": producers, "brands": brands, "lines": lines,
        "products": {"plu": products["PLU"].tolist(), "name": products["Название"].tolist(),
                     "cat": codes(products["Категория"], categories).tolist(),
                     "prod": codes(products["Производитель (корректный)"], producers).tolist(),
                     "brand": codes(products["Бренд (корректный)"], brands).tolist(),
                     "line": codes(products["Продуктовая линия"], lines).tolist()},
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
  --page: #f4f4f1; --surface: #ffffff; --raised: #ffffff; --ink: #0b0b0b; --ink2: #52514e; --muted: #898781;
  --grid: #e8e7e1; --axis: #c3c2b7; --border: rgba(11,11,11,0.09); --border2: rgba(11,11,11,0.16); --wash: rgba(11,11,11,0.045);
  --track: #e8e7e2; --seg-on: #ffffff; --accent: #2a78d6; --accent-ink: #1d5fae; --accent-wash: rgba(42,120,214,0.09);
  --good: #006300; --bad: #d03b3b; --other: #898781;
  --shadow: 0 1px 2px rgba(16,16,14,0.05); --pop: 0 14px 36px rgba(16,16,14,0.16), 0 2px 6px rgba(16,16,14,0.06);
  --s1: #2a78d6; --s2: #eb6834; --s3: #1baf7a; --s4: #eda100; --s5: #e87ba4; --s6: #008300; --s7: #6250d6; --s8: #e34948;
}
:root[data-theme="dark"] {
  color-scheme: dark;
  --page: #0e0e0d; --surface: #191918; --raised: #222221; --ink: #f0efec; --ink2: #c3c2b7; --muted: #8f8d87;
  --grid: #2a2a28; --axis: #3a3a37; --border: rgba(255,255,255,0.09); --border2: rgba(255,255,255,0.18); --wash: rgba(255,255,255,0.06);
  --track: #262624; --seg-on: #3b3b39; --accent: #3987e5; --accent-ink: #8dbcf3; --accent-wash: rgba(57,135,229,0.17);
  --good: #0ca30c; --bad: #e66767; --other: #898781;
  --shadow: none; --pop: 0 18px 44px rgba(0,0,0,0.55);
  --s1: #3987e5; --s2: #d95926; --s3: #199e70; --s4: #c98500; --s5: #d55181; --s6: #0ca30c; --s7: #9085e9; --s8: #e66767;
}
* { box-sizing: border-box; }
[hidden] { display: none !important; }
body { margin: 0; background: var(--page); color: var(--ink); font: 14px/1.45 system-ui, -apple-system, "Segoe UI", Roboto, sans-serif;
  -webkit-font-smoothing: antialiased; }
button, input, select { font: inherit; color: inherit; }
.wrap { max-width: 1680px; margin: 0 auto; padding: 0 24px 32px; }

/* шапка */
.appbar { display: flex; align-items: center; justify-content: space-between; gap: 16px; padding: 16px 0 8px; }
.brand { min-width: 0; }
h1 { font-size: 19px; font-weight: 650; margin: 0; letter-spacing: -0.01em; }
.sub { color: var(--ink2); font-size: 12.5px; margin-top: 2px; }
.appbar-actions { display: flex; gap: 8px; flex: none; }
.icon-btn { width: 34px; height: 34px; display: inline-grid; place-items: center; padding: 0; border: 1px solid var(--border);
  border-radius: 9px; background: var(--surface); color: var(--ink2); cursor: pointer; box-shadow: var(--shadow); }
.icon-btn:hover { color: var(--ink); background: var(--wash); }
.icon-btn svg { width: 18px; height: 18px; }

/* панель управления: период, метрики, рынок */
.toolbar { position: sticky; top: 0; z-index: 30; display: flex; flex-wrap: wrap; align-items: center; gap: 8px 10px; padding: 10px 0;
  margin-bottom: 14px; background: var(--page); background: color-mix(in srgb, var(--page) 90%, transparent);
  -webkit-backdrop-filter: saturate(1.4) blur(10px); backdrop-filter: saturate(1.4) blur(10px); border-bottom: 1px solid var(--border); }
.tgroup { display: inline-flex; flex-wrap: wrap; align-items: center; gap: 6px; }
.tsep { width: 1px; height: 22px; background: var(--border2); margin: 0 2px; }
.toolbar .spacer { flex: 1; }
.seg { display: inline-flex; gap: 2px; padding: 2px; border-radius: 9px; background: var(--track); }
.seg button { height: 28px; padding: 0 11px; border: 0; border-radius: 7px; background: transparent; color: var(--ink2); cursor: pointer;
  font-size: 13px; white-space: nowrap; }
.seg button:hover { color: var(--ink); }
.seg button[aria-pressed="true"] { background: var(--seg-on); color: var(--ink); font-weight: 600; box-shadow: 0 1px 2px rgba(0,0,0,0.12); }
.field { display: inline-flex; align-items: center; gap: 2px; height: 32px; padding: 0 2px 0 10px; border: 1px solid var(--border);
  border-radius: 9px; background: var(--surface); box-shadow: var(--shadow); }
.lbl { color: var(--muted); font-size: 13px; white-space: nowrap; }
.select { height: 30px; max-width: 220px; padding: 0 4px; border: 0; background: transparent; color: var(--ink); font-weight: 600; cursor: pointer; }
.select option { background: var(--surface); color: var(--ink); }
.btn { display: inline-flex; align-items: center; gap: 6px; height: 32px; padding: 0 12px; border: 1px solid var(--border); border-radius: 9px;
  background: var(--surface); color: var(--ink); cursor: pointer; white-space: nowrap; box-shadow: var(--shadow); }
.btn:hover { background: var(--wash); }
.btn svg { width: 15px; height: 15px; }
.btn.quiet { border-color: transparent; background: transparent; color: var(--ink2); box-shadow: none; }
.btn.quiet:hover { color: var(--ink); background: var(--wash); }
.btn.small { height: 28px; padding: 0 10px; font-size: 12.5px; }

/* выпадающий список с галочками */
.dd { position: relative; }
.dd-btn { display: inline-flex; align-items: center; gap: 6px; max-width: 250px; height: 32px; padding: 0 8px 0 10px; border: 1px solid var(--border);
  border-radius: 9px; background: var(--surface); cursor: pointer; box-shadow: var(--shadow); }
.dd-btn:hover { background: var(--wash); }
.dd-val { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-weight: 600; }
.dd-btn[data-active="true"] { border-color: var(--accent); background: var(--accent-wash); }
.chev { flex: none; width: 14px; height: 14px; color: var(--muted); transition: transform 0.15s; }
.dd-btn[aria-expanded="true"] .chev { transform: rotate(180deg); }
.ms-panel { position: absolute; top: calc(100% + 6px); left: 0; z-index: 50; display: flex; flex-direction: column; width: 310px;
  max-height: min(440px, 70vh); overflow: hidden; background: var(--raised); border: 1px solid var(--border2); border-radius: 12px; box-shadow: var(--pop); }
.ms-search { height: 32px; margin: 10px 10px 6px; padding: 0 10px; border: 1px solid var(--border); border-radius: 8px; background: var(--page); color: var(--ink); }
.ms-list { overflow: auto; padding: 2px 0 6px; }
.ms-item { display: flex; align-items: center; gap: 10px; padding: 5px 12px; cursor: pointer; font-size: 13.5px; }
.ms-item:hover { background: var(--wash); }
.ms-item input { flex: none; width: 15px; height: 15px; margin: 0; accent-color: var(--accent); }
.ms-item .sec { margin-left: auto; padding-left: 8px; color: var(--muted); font-size: 12px; white-space: nowrap; }
.ms-item.own .nm { font-weight: 600; }
.ms-all { margin-bottom: 3px; padding-bottom: 7px; border-bottom: 1px solid var(--border); font-weight: 500; }
.ms-foot { display: flex; justify-content: space-between; gap: 8px; padding: 8px 10px; border-top: 1px solid var(--border); }
.ms-empty { padding: 10px 12px; color: var(--muted); font-size: 13px; }

/* строка периода и выбранные срезы */
.context { display: flex; flex-direction: column; gap: 8px; margin-bottom: 14px; }
.period { display: flex; flex-wrap: wrap; align-items: center; gap: 6px 12px; color: var(--ink2); font-size: 13px; }
.period b { color: var(--ink); font-weight: 600; }
.period .warn { color: var(--bad); }
.period .spacer { flex: 1; }
.status { color: var(--ink2); font-size: 12px; }
.chips { display: flex; flex-wrap: wrap; align-items: center; gap: 6px; }
.chip { display: inline-flex; align-items: center; gap: 4px; max-width: 380px; height: 26px; padding: 0 3px 0 10px; border-radius: 13px;
  background: var(--accent-wash); font-size: 12.5px; }
.chip .k { color: var(--ink2); white-space: nowrap; }
.chip .v { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; font-weight: 500; }
.chip button { display: grid; place-items: center; width: 20px; height: 20px; padding: 0; border: 0; border-radius: 10px; background: transparent;
  color: var(--ink2); cursor: pointer; }
.chip button:hover { background: var(--wash); color: var(--ink); }
.chip svg { width: 11px; height: 11px; }
.chips .clear { padding: 0 6px; border: 0; background: none; color: var(--accent-ink); cursor: pointer; font-size: 12.5px; }
.chips .clear:hover { text-decoration: underline; }

/* раскладка: продукт слева, данные справа */
.layout { display: grid; grid-template-columns: 236px minmax(0, 1fr); gap: 20px; align-items: start; }
.side { position: sticky; top: calc(var(--bar-h, 56px) + 14px); display: flex; flex-direction: column; gap: 10px;
  max-height: calc(100vh - var(--bar-h, 56px) - 28px); overflow-y: auto; padding: 0 4px 8px 0; scrollbar-width: thin; }
.side-title { margin: 2px 0 0 2px; color: var(--muted); font-size: 11.5px; font-weight: 650; letter-spacing: 0.06em; text-transform: uppercase; }
.fl { padding: 10px 0 6px; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; box-shadow: var(--shadow); }
.fl-head { display: flex; align-items: baseline; gap: 8px; padding: 0 12px 6px; }
.fl-title { font-size: 13px; font-weight: 600; }
.fl-count { display: inline-grid; place-items: center; min-width: 18px; height: 18px; padding: 0 6px; border-radius: 9px; background: var(--accent);
  color: #fff; font-size: 11px; font-weight: 600; }
.fl-reset { margin-left: auto; padding: 0; border: 0; background: none; color: var(--ink2); cursor: pointer; font-size: 12px; }
.fl-reset:hover { color: var(--ink); text-decoration: underline; }
.fl .ms-search { width: calc(100% - 20px); height: 30px; margin: 0 10px 6px; }
.fl .ms-list { max-height: 300px; padding: 0 0 2px; }
.fl .ms-item { padding: 4px 12px; font-size: 13px; }
.main { min-width: 0; }
@media (max-width: 1100px) {
  .layout { grid-template-columns: minmax(0, 1fr); }
  .side { position: static; display: grid; grid-template-columns: repeat(auto-fit, minmax(230px, 1fr)); max-height: none; }
  .side-title { grid-column: 1 / -1; }
}

/* плитки */
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(128px, 1fr)); gap: 10px; margin-bottom: 16px; }
.tile { min-width: 0; padding: 12px 14px; background: var(--surface); border: 1px solid var(--border); border-radius: 12px; box-shadow: var(--shadow); }
.tile .label { min-height: 2.6em; color: var(--ink2); font-size: 12.5px; line-height: 1.3; }
.tile .value { margin-top: 6px; font-size: clamp(18px, 1.5vw, 24px); font-weight: 650; letter-spacing: -0.01em; white-space: nowrap;
  font-variant-numeric: tabular-nums; }
.delta { margin-top: 2px; font-size: 12.5px; white-space: nowrap; font-variant-numeric: tabular-nums; }
.good { color: var(--good); } .bad { color: var(--bad); } .neutral { color: var(--ink2); } .none { color: var(--muted); }

/* карточки и графики */
.grid2 { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 16px; margin-bottom: 16px; }
@media (max-width: 980px) { .grid2 { grid-template-columns: minmax(0, 1fr); } }
.card { position: relative; min-width: 0; padding: 16px; background: var(--surface); border: 1px solid var(--border); border-radius: 12px;
  box-shadow: var(--shadow); }
.card-head { display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; margin-bottom: 8px; }
.card h2 { margin: 0; font-size: 15px; font-weight: 600; }
.card .note { margin-top: 2px; color: var(--ink2); font-size: 12px; }
.ghost { height: 28px; padding: 0 10px; border: 1px solid var(--border); border-radius: 8px; background: transparent; color: var(--ink2);
  cursor: pointer; font-size: 12px; white-space: nowrap; }
.ghost:hover { background: var(--wash); color: var(--ink); }
.chart { width: 100%; min-height: 40px; }
.chart svg { display: block; overflow: visible; }
.chart text { fill: var(--muted); font-size: 11px; font-variant-numeric: tabular-nums; }
.chart .endlabel { fill: var(--ink2); font-size: 12px; }
.legend { display: flex; flex-wrap: wrap; gap: 6px 16px; margin-top: 8px; color: var(--ink2); font-size: 12px; }
.legend span { display: inline-flex; align-items: center; gap: 6px; }
.key-line { display: inline-block; width: 14px; height: 2px; border-radius: 1px; }
.key-rect { display: inline-block; width: 10px; height: 10px; border-radius: 2px; }
.tip { position: absolute; z-index: 20; min-width: 150px; padding: 8px 10px; pointer-events: none; background: var(--raised);
  border: 1px solid var(--border2); border-radius: 8px; box-shadow: var(--pop); font-size: 12px; }
.tip .t-title { margin-bottom: 4px; color: var(--ink2); }
.tip .t-row { display: flex; align-items: center; gap: 8px; }
.tip .t-val { color: var(--ink); font-weight: 600; font-variant-numeric: tabular-nums; }
.tip .t-name { color: var(--ink2); }
.bars { display: grid; grid-template-columns: minmax(90px, 34%) 1fr; row-gap: 2px; }
.bar-label { padding: 5px 8px 5px 0; overflow: hidden; color: var(--ink2); font-size: 13px; text-overflow: ellipsis; white-space: nowrap; }
.bar-label.own { color: var(--ink); font-weight: 600; }
.bar-cell { display: flex; flex-direction: column; justify-content: center; gap: 2px; padding: 3px 0; border-radius: 4px; outline-offset: 2px; }
.bar-row { display: flex; align-items: center; gap: 6px; min-height: 12px; }
.bar { min-width: 1px; height: 16px; border-radius: 0 4px 4px 0; }
.bar.thin { height: 10px; }
.bar-val { color: var(--ink2); font-size: 12px; font-variant-numeric: tabular-nums; white-space: nowrap; }
.bar-cell:hover .bar, .bar-cell:focus .bar { filter: brightness(1.12); }

/* разрезы во вкладках */
.tabs-card { padding: 0; }
.tabs-head { display: flex; flex-wrap: wrap; align-items: center; gap: 4px 16px; padding: 4px 16px 0; border-bottom: 1px solid var(--border); }
.tabs { display: flex; flex: 1 1 auto; gap: 2px; min-width: 0; margin-bottom: -1px; overflow-x: auto; scrollbar-width: none; }
.tab { padding: 11px 12px 9px; border: 0; border-bottom: 2px solid transparent; background: none; color: var(--ink2); cursor: pointer;
  font-size: 13.5px; white-space: nowrap; }
.tab:hover { color: var(--ink); }
.tab[aria-selected="true"] { color: var(--ink); font-weight: 600; border-bottom-color: var(--accent); }
.tab-tools { display: flex; align-items: center; gap: 8px; padding: 6px 0; }
.tab-search { width: 210px; height: 30px; padding: 0 10px; border: 1px solid var(--border); border-radius: 8px; background: var(--page); color: var(--ink); }
.tab-note { padding: 10px 16px 0; color: var(--ink2); font-size: 12px; }
.tab-body { padding: 4px 16px 14px; }
.table-wrap { overflow-x: auto; }
table.data { width: 100%; border-collapse: collapse; font-size: 13px; }
table.data th { padding: 9px 8px 7px; border-bottom: 1px solid var(--axis); color: var(--ink2); cursor: pointer; font-size: 12px; font-weight: 500;
  line-height: 1.25; text-align: left; vertical-align: bottom; user-select: none; }
table.data th.num { min-width: 64px; }
table.data td { white-space: nowrap; }
table.data th:hover { color: var(--ink); }
table.data th[aria-sort] { color: var(--ink); }
table.data th[aria-sort="descending"]::after { content: " ↓"; color: var(--accent); }
table.data th[aria-sort="ascending"]::after { content: " ↑"; color: var(--accent); }
table.data th.num, table.data td.num { text-align: right; font-variant-numeric: tabular-nums; }
table.data td { padding: 6px 8px; border-bottom: 1px solid var(--grid); }
table.data tbody tr:hover td { background: var(--wash); }
table.data td.name { max-width: 340px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
table.data td .own, table.data tr.own td.name { font-weight: 600; }
table.data tr.own td:first-child { box-shadow: inset 3px 0 0 var(--accent); }
table.data tr.pin-end td { border-bottom-color: var(--axis); }
table.data tr.total td { border-top: 1px solid var(--axis); border-bottom: 0; font-weight: 600; }
table.data tr.line { cursor: pointer; }
table.data tr.child td { background: var(--wash); font-size: 12.5px; }
table.data tr.child td.name { padding-left: 36px; color: var(--ink2); font-weight: 400; }
table.data tr.child:hover td { background: var(--accent-wash); }
.tree-btn { display: inline-grid; place-items: center; width: 20px; height: 20px; margin: 0 6px 0 -2px; padding: 0; vertical-align: -5px;
  border: 0; border-radius: 5px; background: none; color: var(--ink2); cursor: pointer; }
.tree-btn:hover { background: var(--wash); color: var(--ink); }
.tree-btn svg { width: 14px; height: 14px; transition: transform 0.15s; }
tr[aria-expanded="true"] .tree-btn svg { transform: rotate(90deg); }
.plu { margin-left: 6px; color: var(--muted); font-size: 11.5px; }
table.data.lines td.name { max-width: 290px; }
.nm-cut { display: inline-block; max-width: 200px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; vertical-align: bottom; }
.more { margin-top: 8px; }
.empty { padding: 24px 0; color: var(--muted); text-align: center; }

/* справочник */
dialog.help { width: min(780px, calc(100vw - 32px)); max-height: calc(100vh - 64px); padding: 0; overflow: auto; background: var(--raised);
  color: var(--ink); border: 1px solid var(--border2); border-radius: 16px; box-shadow: var(--pop); }
dialog.help::backdrop { background: rgba(10,10,10,0.38); -webkit-backdrop-filter: blur(2px); backdrop-filter: blur(2px); }
.help-head { position: sticky; top: 0; display: flex; justify-content: space-between; align-items: center; gap: 12px; padding: 16px 20px;
  background: var(--raised); border-bottom: 1px solid var(--border); }
.help-head h2 { margin: 0; font-size: 17px; }
.help-body { padding: 2px 20px 20px; font-size: 13.5px; line-height: 1.55; }
.help-body h3 { margin: 18px 0 6px; font-size: 14px; }
.help-body p { margin: 6px 0; color: var(--ink2); }
.help-body dl { display: grid; grid-template-columns: 170px 1fr; gap: 6px 16px; margin: 8px 0; }
.help-body dt { font-weight: 600; }
.help-body dd { margin: 0; color: var(--ink2); }
@media (max-width: 640px) { .help-body dl { grid-template-columns: 1fr; } }

/* служебное */
.loading { display: flex; justify-content: center; align-items: center; gap: 12px; padding: 72px 0; color: var(--ink2); }
.spinner { width: 18px; height: 18px; border: 2px solid var(--border2); border-top-color: var(--accent); border-radius: 50%; animation: spin 0.8s linear infinite; }
@keyframes spin { to { transform: rotate(360deg); } }
.error { padding: 48px; color: var(--bad); text-align: center; }
footer { display: flex; flex-wrap: wrap; gap: 6px 16px; margin-top: 18px; color: var(--muted); font-size: 12px; }
.sr-only { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; }
:focus-visible { outline: 2px solid var(--accent); outline-offset: 2px; }
input:focus-visible { outline-offset: 0; border-color: var(--accent); }
@media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } }
</style>
</head>
<body>
<div class="wrap">
  <header class="appbar">
    <div class="brand"><h1 id="title"></h1><div class="sub" id="subtitle"></div></div>
    <div class="appbar-actions" id="appbar-actions"></div>
  </header>
  <div id="app"><div class="loading"><span class="spinner"></span>Загружаю данные…</div></div>
</div>
<dialog id="help" class="help" aria-labelledby="help-title"></dialog>
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
    if (kind === 'idx') {   // изменение индекса цены: два знака, без единицы
      const r = Math.round(v * 100) / 100;
      if (r === 0) return { text: nf2.format(0), dir: 0 };
      return { text: (r > 0 ? '+' : '−') + nf2.format(Math.abs(r)), dir: r > 0 ? 1 : -1 };
    }
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
    `${weekly ? 'Недели, месяцы по 4-4-5' : 'Календарные месяцы'} · данные с ${per.label[0]} по ${per.label[nP - 1]} · собрано ${META.generated}`;

  // ---------- периоды ----------
  // ord — сквозной номер недели (месяца). По нему строятся скользящие окна и PP: пропуски в данных окна не сдвигают.
  const ord = per.ord, ordIndex = new Map(ord.map((o, p) => [o, p]));
  const SPANS = weekly
    ? [['P', 'Нед', 'Одна неделя'], ['L4', 'L4W', 'Последние 4 недели'], ['L13', 'L13W', 'Последние 13 недель'],
       ['MTD', 'MTD', 'С начала месяца 4-4-5 до выбранной недели'],
       ['YTD', 'YTD', 'С начала года до выбранной недели'], ['MAT', 'MAT', 'Последние 52 недели']]
    : [['P', 'Мес', 'Один месяц'], ['L3', 'L3M', 'Последние 3 месяца'],
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
    // Магазины с продажами всей категории БП по периодам — знаменатель ND и TDP
    const catS_p = new Float64Array(nP);
    const cPer = D.cP, cN = D.cN, cGeo = D.cGeo, cS = D.cS;
    for (let j = 0; j < cPer.length; j++) {
      const p = cPer[j];
      if ((pf[p] & 3) && netM[cN[j]] && geoOk[cGeo[j]]) catS_p[p] += cS[j];
    }
    const inv = new Float64Array(nP);
    let cntW = 0, cntC = 0;
    for (const p of win.W) if (catS_p[p] > 0) { inv[p] = 1 / catS_p[p]; cntW++; }
    for (const p of win.C) if (catS_p[p] > 0) { inv[p] = 1 / catS_p[p]; cntC++; }
    // t, p, l — тотал, промо и продажи по картам в выбранной единице; m — продажи выбранного блока.
    // Цены блока — его рубли к его штукам и килограммам.
    const tOf = valueOf('T', st.unit), pOf = valueOf('P', st.unit), lOf = valueOf('L', st.unit);
    const bUof = valueOf(st.block, 'U'), bRof = valueOf(st.block, 'R'), bGof = valueOf(st.block, 'G');
    const promoRub = valueOf('P', 'R');
    const blk = st.block, own = META.own, lineOf = META.products.line, nLine = META.lines.length, nCats = META.categories.length;
    let wM = 0, wT = 0, wP = 0, wL = 0, wU = 0, wR = 0, wG = 0, wDF = 0, wDS = 0, catV = 0, catR = 0, catG = 0, catU = 0;
    let cmpM = 0, cmpT = 0, cmpP = 0, cmpL = 0, cmpU = 0, cmpR = 0, cmpG = 0, cmpDF = 0, cmpDS = 0, cmpCat = 0;
    let cmpCatR = 0, cmpCatU = 0, cmpCatG = 0, pbpV = 0, pbpC = 0;
    const selM_p = new Float64Array(nP), selS_p = new Float64Array(nP), catM_p = new Float64Array(nP);
    const prodM = new Float64Array(nProd), prodT = new Float64Array(nProd), prodP = new Float64Array(nProd), prodL = new Float64Array(nProd);
    const prodDF = new Float64Array(nProd), prodDS = new Float64Array(nProd), prodM_p = new Float64Array(nP * nProd);
    const prodC = new Float64Array(nProd), prodU = new Float64Array(nProd), prodR = new Float64Array(nProd), prodG = new Float64Array(nProd);
    const prodND = new Float64Array(nProd);
    const brM = new Float64Array(nBrand), brC = new Float64Array(nBrand), brT = new Float64Array(nBrand), brP = new Float64Array(nBrand);
    const brL = new Float64Array(nBrand), brU = new Float64Array(nBrand), brR = new Float64Array(nBrand), brG = new Float64Array(nBrand);
    const brDF = new Float64Array(nBrand), brDS = new Float64Array(nBrand), brS_p = new Float64Array(nP * nBrand);
    const foV = new Float64Array(nF), foC = new Float64Array(nF), foCat = new Float64Array(nF), foCatC = new Float64Array(nF);
    const geoV = new Float64Array(nGeo), geoC = new Float64Array(nGeo), geoCat = new Float64Array(nGeo), geoCatC = new Float64Array(nGeo);
    const nV = new Float64Array(nN), nC = new Float64Array(nN), nCat = new Float64Array(nN), nCatC = new Float64Array(nN);
    const skM = new Float64Array(nK), skC = new Float64Array(nK), skT = new Float64Array(nK), skP = new Float64Array(nK);
    const skL = new Float64Array(nK), skU = new Float64Array(nK), skR = new Float64Array(nK), skG = new Float64Array(nK);
    const skDF = new Float64Array(nK), skDS = new Float64Array(nK), skND = new Float64Array(nK), skNDc = new Float64Array(nK), skSt = new Float64Array(nK);
    const lnM = new Float64Array(nLine), lnC = new Float64Array(nLine), lnT = new Float64Array(nLine), lnP = new Float64Array(nLine);
    const lnL = new Float64Array(nLine), lnU = new Float64Array(nLine), lnR = new Float64Array(nLine), lnG = new Float64Array(nLine);
    const lnDF = new Float64Array(nLine), lnDS = new Float64Array(nLine), lnND = new Float64Array(nLine), lnNDc = new Float64Array(nLine);
    const lnSt = new Float64Array(nLine);
    // вся ПБП по категориям — без срезов продукта
    const ctM = new Float64Array(nCats), ctC = new Float64Array(nCats), ctT = new Float64Array(nCats), ctP = new Float64Array(nCats);
    const ctL = new Float64Array(nCats), ctU = new Float64Array(nCats), ctR = new Float64Array(nCats), ctG = new Float64Array(nCats);
    const ctOwn = new Float64Array(nCats), ctOwnC = new Float64Array(nCats);
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
      const tv = tOf(i), pv = pOf(i), lv = lOf(i), m = blk === 'T' ? tv : blk === 'P' ? pv : lv;
      const c = catOf[k], pr = prodOf[k], mine = pr === own;
      let bu = 0, br = 0, bg = 0;
      if (fl & 1) {
        bu = bUof(i); br = bRof(i); bg = bGof(i);
        pbpV += m; ctM[c] += m; ctT[c] += tv; ctP[c] += pv; ctL[c] += lv; ctU[c] += bu; ctR[c] += br; ctG[c] += bg;
        if (mine) ctOwn[c] += m;
      }
      if (fl & 2) { pbpC += m; ctC[c] += m; if (mine) ctOwnC[c] += m; }
      if (!catM[k]) continue;
      const full = fullM[k];
      catM_p[p] += m; prodM_p[p * nProd + pr] += m;
      if (full) { selM_p[p] += m; selS_p[p] += S[i]; }
      if (fl & 1) {
        const f = geoFo[g], nd = S[i] * inv[p];
        depthOf(i);
        catV += m; catR += br; catG += bg; catU += bu; foCat[f] += m; geoCat[g] += m; nCat[net] += m;
        prodM[pr] += m; prodT[pr] += tv; prodP[pr] += pv; prodL[pr] += lv; prodDF[pr] += dF; prodDS[pr] += dS;
        prodU[pr] += bu; prodR[pr] += br; prodG[pr] += bg; prodND[pr] += nd;
        if (full) {
          const b = brandOf[k], li = lineOf[k];
          wM += m; wT += tv; wP += pv; wL += lv; wU += bu; wR += br; wG += bg; wDF += dF; wDS += dS;
          brM[b] += m; brT[b] += tv; brP[b] += pv; brL[b] += lv; brU[b] += bu; brR[b] += br; brG[b] += bg;
          brDF[b] += dF; brDS[b] += dS; brS_p[p * nBrand + b] += S[i];
          foV[f] += m; geoV[g] += m; nV[net] += m;
          skM[k] += m; skT[k] += tv; skP[k] += pv; skL[k] += lv; skU[k] += bu; skR[k] += br; skG[k] += bg;
          skDF[k] += dF; skDS[k] += dS; skND[k] += nd; skSt[k] += S[i];
          lnM[li] += m; lnT[li] += tv; lnP[li] += pv; lnL[li] += lv; lnU[li] += bu; lnR[li] += br; lnG[li] += bg;
          lnDF[li] += dF; lnDS[li] += dS; lnND[li] += nd; lnSt[li] += S[i];
        }
      }
      if (fl & 2) {
        const f = geoFo[g], nd = S[i] * inv[p], cu = bUof(i), cr = bRof(i), cg = bGof(i);
        cmpCat += m; cmpCatU += cu; cmpCatR += cr; cmpCatG += cg; foCatC[f] += m; geoCatC[g] += m; nCatC[net] += m;
        prodC[pr] += m;
        if (full) {
          const li = lineOf[k];
          depthOf(i);
          cmpM += m; cmpT += tv; cmpP += pv; cmpL += lv; cmpU += cu; cmpR += cr; cmpG += cg; cmpDF += dF; cmpDS += dS;
          brC[brandOf[k]] += m; foC[f] += m; geoC[g] += m; nC[net] += m; skC[k] += m; skNDc[k] += nd;
          lnC[li] += m; lnNDc[li] += nd;
        }
      }
    }
    // TDP: среднее по неделям (месяцам) окна; знаменатель — магазины с продажами всей категории БП
    const tdpOver = (periods, stores, offset, stride) => {
      let sum = 0, cnt = 0;
      for (const p of periods) if (catS_p[p] > 0) { sum += stores[p * stride + offset] / catS_p[p]; cnt++; }
      return cnt ? sum / cnt : null;
    };
    const ndW = x => (cntW ? x / cntW : null), ndC = x => (cntC ? x / cntC : null);
    const hasCmp = win.C.length > 0;
    const dShare = (a, b) => (hasCmp && ok(a) && ok(b) ? 100 * (a - b) : null);
    const depth = (ds, df) => (df > 0 ? ds / df : null);
    // Индекс цены — к категориям из среза: за штуку, а при единице «кг» — за килограмм
    const kgIdx = st.unit === 'G';
    const priceIndex = (r, u, g, cr, cu, cg) => (kgIdx ? (g > 0 && cg > 0 && cr > 0 ? (r / g) / (cr / cg) : null)
      : (u > 0 && cu > 0 && cr > 0 ? (r / u) / (cr / cu) : null));
    const index = (r, u, g) => priceIndex(r, u, g, catR, catU, catG);
    const kpi = {
      sales: wM, salesCmp: cmpM, share: div(wM, catV), shareCmp: div(cmpM, cmpCat),
      sharePbp: div(wM, pbpV), sharePbpCmp: div(cmpM, pbpC),
      promoShare: div(wP, wT), promoShareCmp: div(cmpP, cmpT), loyalShare: div(wL, wT), loyalShareCmp: div(cmpL, cmpT),
      price: div(wR, wU), priceCmp: div(cmpR, cmpU), priceKg: div(wR, wG), priceKgCmp: div(cmpR, cmpG),
      index: index(wR, wU, wG), indexCmp: priceIndex(cmpR, cmpU, cmpG, cmpCatR, cmpCatU, cmpCatG),
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
    const brands = [], fos = [], regions = [], nets = [], skus = [], producers = [], categories = [], lines = [];
    for (let b = 0; b < nBrand; b++) {
      if (!brM[b] && !brC[b]) continue;
      const share = div(brM[b], catV), shareCmp = div(brC[b], cmpCat);
      brands.push({ brand: b, sales: brM[b], salesCmp: brC[b], growth: growth(brM[b], brC[b]), share, shareCmp,
        dShare: dShare(share, shareCmp), sharePbp: div(brM[b], pbpV), promoShare: div(brP[b], brT[b]), loyalShare: div(brL[b], brT[b]),
        price: div(brR[b], brU[b]), priceKg: div(brR[b], brG[b]), index: index(brR[b], brU[b], brG[b]),
        depth: depth(brDS[b], brDF[b]), tdp: tdpOver(win.W, brS_p, b, nBrand) });
    }
    // Производители — все в категориях из среза: срезы производителя и бренда не применяются
    for (let pr = 0; pr < nProd; pr++) {
      if (!prodM[pr] && !prodC[pr]) continue;
      const share = div(prodM[pr], catV), shareCmp = div(prodC[pr], cmpCat);
      producers.push({ prod: pr, sales: prodM[pr], salesCmp: prodC[pr], growth: growth(prodM[pr], prodC[pr]), share, shareCmp,
        dShare: dShare(share, shareCmp), sharePbp: div(prodM[pr], pbpV), promoShare: div(prodP[pr], prodT[pr]),
        loyalShare: div(prodL[pr], prodT[pr]), price: div(prodR[pr], prodU[pr]), priceKg: div(prodR[pr], prodG[pr]),
        index: index(prodR[pr], prodU[pr], prodG[pr]), depth: depth(prodDS[pr], prodDF[pr]), tdp: ndW(prodND[pr]) });
    }
    // Категории — вся ПБП в выбранных сетях и регионах: срезы продукта не применяются
    let tAll = 0, pAll = 0, lAll = 0, uAll = 0, rAll = 0, gAll = 0, ownAll = 0, ownAllC = 0;
    for (let c = 0; c < nCats; c++) {
      tAll += ctT[c]; pAll += ctP[c]; lAll += ctL[c]; uAll += ctU[c]; rAll += ctR[c]; gAll += ctG[c]; ownAll += ctOwn[c]; ownAllC += ctOwnC[c];
      if (!ctM[c] && !ctC[c]) continue;
      const share = div(ctM[c], pbpV), shareCmp = div(ctC[c], pbpC);
      const ownShare = own >= 0 ? div(ctOwn[c], ctM[c]) : null, ownShareCmp = own >= 0 ? div(ctOwnC[c], ctC[c]) : null;
      categories.push({ cat: c, sales: ctM[c], salesCmp: ctC[c], growth: growth(ctM[c], ctC[c]), share, shareCmp,
        dShare: dShare(share, shareCmp), ownShare, ownShareCmp, dOwnShare: dShare(ownShare, ownShareCmp),
        promoShare: div(ctP[c], ctT[c]), loyalShare: div(ctL[c], ctT[c]), price: div(ctR[c], ctU[c]), priceKg: div(ctR[c], ctG[c]) });
    }
    const pbpOwn = own >= 0 ? div(ownAll, pbpV) : null, pbpOwnC = own >= 0 ? div(ownAllC, pbpC) : null;
    const catTotal = { sales: pbpV, salesCmp: pbpC, growth: growth(pbpV, pbpC), share: pbpV > 0 ? 1 : null, dShare: null,
      ownShare: pbpOwn, dOwnShare: dShare(pbpOwn, pbpOwnC), promoShare: div(pAll, tAll), loyalShare: div(lAll, tAll),
      price: div(rAll, uAll), priceKg: div(rAll, gAll) };
    const geoRow = (v, c, cat, catC) => {
      const share = div(v, cat), shareCmp = div(c, catC);
      return { sales: v, salesCmp: c, growth: growth(v, c), share, shareCmp, dShare: dShare(share, shareCmp) };
    };
    for (let f = 0; f < nF; f++) if (foCat[f] || foCatC[f]) fos.push({ fo: f, ...geoRow(foV[f], foC[f], foCat[f], foCatC[f]) });
    if (regionLevel) for (let g = 0; g < nGeo; g++) if (geoCat[g] || geoCatC[g]) regions.push({ geo: g, fo: geoFo[g], ...geoRow(geoV[g], geoC[g], geoCat[g], geoCatC[g]) });
    for (let n = 0; n < nN; n++) if (nCat[n] || nCatC[n]) nets.push({ net: n, ...geoRow(nV[n], nC[n], nCat[n], nCatC[n]) });
    for (let k = 0; k < nK; k++) {
      if (!skM[k] && !skC[k]) continue;
      const share = div(skM[k], catV), shareCmp = div(skC[k], cmpCat), nd = ndW(skND[k]), ndCmp = ndC(skNDc[k]);
      skus.push({ k, sales: skM[k], salesCmp: skC[k], growth: growth(skM[k], skC[k]), share, shareCmp, dShare: dShare(share, shareCmp),
        sharePbp: div(skM[k], pbpV), promoShare: div(skP[k], skT[k]), loyalShare: div(skL[k], skT[k]), price: div(skR[k], skU[k]),
        priceKg: div(skR[k], skG[k]), index: index(skR[k], skU[k], skG[k]), depth: depth(skDS[k], skDF[k]),
        nd, ndCmp, dNd: hasCmp && ok(nd) && ok(ndCmp) ? 100 * (nd - ndCmp) : null, ros: div(skM[k], skSt[k]) });
    }
    // Продуктовые линии: суммы своих SKU; TDP линии — сумма ND её SKU
    const kids = new Map();
    for (const s of skus) { const li = lineOf[s.k]; if (!kids.has(li)) kids.set(li, []); kids.get(li).push(s); }
    for (let li = 0; li < nLine; li++) {
      if (!lnM[li] && !lnC[li]) continue;
      const list = kids.get(li) || [], byBrand = new Map();
      let mineCnt = 0, brand = -1, best = -Infinity;
      for (const s of list) {
        if (prodOf[s.k] === own) mineCnt++;
        byBrand.set(brandOf[s.k], (byBrand.get(brandOf[s.k]) || 0) + s.sales);
      }
      for (const [b, v] of byBrand) if (v > best) { best = v; brand = b; }
      const share = div(lnM[li], catV), shareCmp = div(lnC[li], cmpCat), tdp = ndW(lnND[li]), tdpCmp = ndC(lnNDc[li]);
      lines.push({ line: li, brand, own: own >= 0 && mineCnt * 2 > list.length, sales: lnM[li], salesCmp: lnC[li],
        growth: growth(lnM[li], lnC[li]), share, shareCmp, dShare: dShare(share, shareCmp), sharePbp: div(lnM[li], pbpV),
        promoShare: div(lnP[li], lnT[li]), loyalShare: div(lnL[li], lnT[li]), price: div(lnR[li], lnU[li]), priceKg: div(lnR[li], lnG[li]),
        index: index(lnR[li], lnU[li], lnG[li]), depth: depth(lnDS[li], lnDF[li]), tdp, tdpCmp,
        dTdp: hasCmp && ok(tdp) && ok(tdpCmp) ? 100 * (tdp - tdpCmp) : null, ros: div(lnM[li], lnSt[li]),
        skuCount: list.filter(s => s.sales > 0).length, skus: list });
    }
    const total = { sales: wM, salesCmp: cmpM, growth: growth(wM, cmpM), share: kpi.share, shareCmp: kpi.shareCmp,
      dShare: dShare(kpi.share, kpi.shareCmp) };
    return { win, kpi, chart1, chart2, chart3, chart4, brands, fos, regions, nets, skus, total, producers, categories, catTotal, lines };
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
      'Производитель', 'Категория', 'Продуктовая линия', 'Продажи, шт', 'Продажи, руб', 'Продажи, кг', 'Промо, шт', 'Промо, руб', 'Промо, кг',
      'По картам, шт', 'По картам, руб', 'По картам, кг', 'Магазины с продажами'];
    const cell = list => list.map(csvCell);
    const perS = per.label.map((l, p) => cell([l, per.start[p], per.end[p]]).join(';'));
    const netS = cell(META.nets), foS = cell(META.fos), regS = cell(META.regions), pluS = cell(META.products.plu);
    const nameS = cell(META.products.name), brandS = cell(META.brands), prodS = cell(META.producers), catS = cell(META.categories), lineS = cell(META.lines), lineOfK = META.products.line;
    const U = D.U, R = D.R, Rk = D.Rk, G = D.G, PU = D.PU, LU = D.LU, S = D.S;
    const promoRub = measureOf('R', PU, D.PR, D.PRk), loyalRub = measureOf('R', LU, D.LR, D.LRk);
    const promoKg = measureOf('G', PU, D.PR, D.PRk, D.PG), loyalKg = measureOf('G', LU, D.LR, D.LRk, D.LG);
    const parts = [cell(head).join(';') + '\r\n'];
    let chunk = [];
    for (let i = 0; i < P.length; i++) {
      const p = P[i], net = N[i], g = GEO[i], k = K[i];
      if (!inW[p] || !netM[net] || !geoOk[g] || !fullM[k]) continue;
      chunk.push([perS[p], netS[net], foS[geoFo[g]] + (regionLevel ? ';' + regS[g] : ''), pluS[k], nameS[k],
        brandS[brandOf[k]], prodS[prodOf[k]], catS[catOf[k]], lineS[lineOfK[k]],
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
  const ICONS = {
    chev: '<svg class="chev" viewBox="0 0 16 16" aria-hidden="true"><path d="M4 6l4 4 4-4" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    right: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M6 3.5l4.5 4.5L6 12.5" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    x: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M4 4l8 8M12 4l-8 8" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
    moon: '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20.5 14.5A8.5 8.5 0 0 1 9.5 3.5a8.5 8.5 0 1 0 11 11z" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linejoin="round"/></svg>',
    sun: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="4.2" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M12 2.5v2.2M12 19.3v2.2M2.5 12h2.2M19.3 12h2.2M5.3 5.3l1.6 1.6M17.1 17.1l1.6 1.6M5.3 18.7l1.6-1.6M17.1 6.9l1.6-1.6" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/></svg>',
    help: '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9.2" fill="none" stroke="currentColor" stroke-width="1.8"/><path d="M9.4 9.3a2.7 2.7 0 0 1 5.2 1c0 1.8-2.6 2.3-2.6 4" fill="none" stroke="currentColor" stroke-width="1.8" stroke-linecap="round"/><circle cx="12" cy="17.4" r="1.1" fill="currentColor"/></svg>',
    download: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M8 2.5v8M4.5 7.5L8 11l3.5-3.5M3 13.5h10" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>',
    reset: '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3.2 8a4.8 4.8 0 1 0 1.5-3.5M3 2.5v3h3" fill="none" stroke="currentColor" stroke-width="1.7" stroke-linecap="round" stroke-linejoin="round"/></svg>',
  };
  function icon(name) {
    const t = document.createElement('template');
    t.innerHTML = ICONS[name];
    return t.content.firstChild;
  }
  // Список с галочками: «Все» — когда ничего не отмечено. Один и тот же для боковой панели и выпадающих списков шапки.
  function checkList(label, set, getOptions, onChange, list, search, after) {
    function fill() {
      const all = getOptions(), q = search.value.trim().toLowerCase();
      list.textContent = '';
      const allBox = el('input', { type: 'checkbox' });
      allBox.checked = set.size === 0;
      allBox.addEventListener('change', () => { set.clear(); fill(); after(); onChange(); });
      list.append(el('label', { class: 'ms-item ms-all' }, allBox, el('span', { text: 'Все' })));
      let shown = 0;
      for (const o of all) {
        if (q && !o.text.toLowerCase().includes(q)) continue;
        shown++;
        const box = el('input', { type: 'checkbox' });
        box.checked = set.has(o.value);
        box.addEventListener('change', () => {
          if (box.checked) set.add(o.value); else set.delete(o.value);
          allBox.checked = set.size === 0;
          after(); onChange();
        });
        list.append(el('label', { class: 'ms-item' + (o.own ? ' own' : '') }, box, el('span', { class: 'nm', text: o.text }),
          o.sec ? el('span', { class: 'sec', text: o.sec }) : null));
      }
      if (!shown) list.append(el('div', { class: 'ms-empty', text: 'Ничего не найдено' }));
    }
    return fill;
  }
  // Боковая панель: категория, производитель, бренд
  function filterList(label, set, getOptions, onChange) {
    const count = el('span', { class: 'fl-count' });
    const reset = el('button', { class: 'fl-reset', type: 'button', text: 'Сбросить' });
    const search = el('input', { class: 'ms-search', type: 'search', placeholder: 'Поиск', 'aria-label': `Поиск: ${label}` });
    const list = el('div', { class: 'ms-list', role: 'group', 'aria-label': label });
    const wrap = el('section', { class: 'fl' }, el('div', { class: 'fl-head' }, el('span', { class: 'fl-title', text: label }), count, reset), search, list);
    function head() { count.textContent = set.size ? String(set.size) : ''; count.title = set.size ? `выбрано ${set.size}` : ''; count.hidden = !set.size; reset.hidden = !set.size; }
    const fill = checkList(label, set, getOptions, onChange, list, search, head);
    function refresh() { search.hidden = getOptions().length <= 8; fill(); head(); }
    reset.addEventListener('click', () => { set.clear(); refresh(); onChange(); });
    search.addEventListener('input', fill);
    refresh();
    return { wrap, refresh };
  }
  // Шапка: выпадающий список с галочками и поиском (сеть, ФО, регион)
  let openDropdown = null;
  function dropdown(label, set, getOptions, onChange) {
    const val = el('span', { class: 'dd-val' });
    const btn = el('button', { class: 'dd-btn', type: 'button', 'aria-haspopup': 'true', 'aria-expanded': 'false' },
      el('span', { class: 'lbl', text: label }), val, icon('chev'));
    const search = el('input', { class: 'ms-search', type: 'search', placeholder: 'Поиск', 'aria-label': `Поиск: ${label}` });
    const list = el('div', { class: 'ms-list', role: 'group', 'aria-label': label });
    const clear = el('button', { class: 'btn quiet small', type: 'button', text: 'Сбросить' });
    const done = el('button', { class: 'btn small', type: 'button', text: 'Готово' });
    const panel = el('div', { class: 'ms-panel', hidden: true }, search, list, el('div', { class: 'ms-foot' }, clear, done));
    const wrap = el('div', { class: 'dd' }, btn, panel);
    function head() {
      const names = getOptions().filter(o => set.has(o.value)).map(o => o.text);
      val.textContent = !set.size ? 'Все' : set.size <= 2 ? names.join(', ') : `${set.size} выбрано`;
      btn.dataset.active = String(set.size > 0);
      btn.title = set.size ? `${label}: ${names.join(', ')}` : `${label}: все`;
    }
    const fill = checkList(label, set, getOptions, onChange, list, search, head);
    function open() {
      if (openDropdown && openDropdown !== api) openDropdown.close(false);
      openDropdown = api;
      search.hidden = getOptions().length <= 8;
      fill();
      panel.hidden = false;
      btn.setAttribute('aria-expanded', 'true');
      (search.hidden ? list.querySelector('input') : search).focus();
    }
    function close(focus) {
      if (panel.hidden) return;
      panel.hidden = true;
      btn.setAttribute('aria-expanded', 'false');
      search.value = '';
      if (openDropdown === api) openDropdown = null;
      if (focus) btn.focus();
    }
    btn.addEventListener('click', () => (panel.hidden ? open() : close(false)));
    clear.addEventListener('click', () => { set.clear(); fill(); head(); onChange(); });
    done.addEventListener('click', () => close(true));
    search.addEventListener('input', fill);
    wrap.addEventListener('keydown', e => { if (e.key === 'Escape' && !panel.hidden) { e.preventDefault(); close(true); } });
    function refresh() { head(); if (!panel.hidden) fill(); }
    const api = { wrap, refresh, close };
    head();
    return api;
  }
  document.addEventListener('pointerdown', e => { if (openDropdown && !openDropdown.wrap.contains(e.target)) openDropdown.close(false); });

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
    const showFo = state.fos.size !== 1;
    return META.regions.map((t, i) => ({ value: i, text: t, sec: showFo ? META.fos[geoFo[i]] : null }))
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
  // Тема: одна кнопка в шапке, выбор запоминается в браузере
  const theme = () => document.documentElement.dataset.theme || 'light';
  const themeBtn = el('button', { class: 'icon-btn', type: 'button' });
  function paintThemeBtn() {
    const dark = theme() === 'dark';
    themeBtn.textContent = '';
    themeBtn.append(icon(dark ? 'sun' : 'moon'));
    themeBtn.title = dark ? 'Светлая тема' : 'Тёмная тема';
    themeBtn.setAttribute('aria-label', themeBtn.title);
  }
  function setTheme(v) {
    document.documentElement.dataset.theme = v;
    try { localStorage.setItem('x5-theme', v); } catch (e) { /* без сохранения — тема просто не запомнится */ }
    paintThemeBtn();
  }
  themeBtn.addEventListener('click', () => setTheme(theme() === 'dark' ? 'light' : 'dark'));
  paintThemeBtn();
  const helpBtn = el('button', { class: 'icon-btn', type: 'button', title: 'Справочник (клавиша ?)', 'aria-label': 'Справочник' }, icon('help'));
  const resetBtn = el('button', { class: 'btn quiet', type: 'button', title: 'Вернуть период, метрики и срезы по умолчанию' }, icon('reset'), 'Сбросить всё');
  resetBtn.addEventListener('click', () => {
    Object.assign(state, { end: nP - 1 }, DEFAULTS);
    for (const k of ['nets', 'fos', 'regs', 'cats', 'prods', 'brands']) state[k].clear();
    syncControls(); schedule();
  });
  document.getElementById('appbar-actions').append(resetBtn, helpBtn, themeBtn);

  function marketChanged(dim) {
    if (dim === 'fos' && regionLevel) { prune(state.regs, availableRegions()); controls.regs.refresh(); }
    schedule();
  }
  // Сверху: период, метрики и рынок
  function buildToolbar() {
    const bar = el('div', { class: 'toolbar', role: 'region', 'aria-label': 'Период, метрики и рынок' });
    const sel = el('select', { class: 'select', 'aria-label': weekly ? 'Неделя, которой заканчивается период' : 'Месяц, которым заканчивается период' });
    for (let p = nP - 1; p >= 0; p--) sel.append(el('option', { value: String(p), text: periodName(p) }));
    sel.addEventListener('change', () => { state.end = Number(sel.value); schedule(); });
    controls.end = { wrap: el('label', { class: 'field' }, el('span', { class: 'lbl', text: 'До' }), sel),
      refresh: () => { sel.value = String(state.end); } };
    controls.end.refresh();
    controls.span = segmented('Длина периода', SPANS.map(([value, text, title]) => ({ value, text, title })), () => state.span, v => { state.span = v; });
    controls.cmp = segmented('Сравнение', [{ value: 'YA', text: 'vs YA', title: 'С тем же периодом год назад' },
      { value: 'PP', text: 'vs PP', title: 'С предыдущим периодом той же длины' }], () => state.cmp, v => { state.cmp = v; });
    controls.block = segmented('Блок метрик', BLOCKS.map(([value, text, , title]) => ({ value, text, title })), () => state.block, v => { state.block = v; });
    controls.unit = segmented('Единица', MEASURES.map(([value, text, title]) => ({ value, text, title })), () => state.unit, v => { state.unit = v; });
    controls.nets = dropdown('Сеть', state.nets, () => META.nets.map((t, i) => ({ value: i, text: t })), () => marketChanged('nets'));
    controls.fos = dropdown('ФО', state.fos, () => META.fos.map((t, i) => ({ value: i, text: t })), () => marketChanged('fos'));
    if (regionLevel) controls.regs = dropdown('Регион', state.regs, availableRegions, () => marketChanged('regs'));
    const sep = () => el('span', { class: 'tsep', 'aria-hidden': 'true' });
    bar.append(
      el('div', { class: 'tgroup', role: 'group', 'aria-label': 'Период' }, controls.end.wrap, controls.span.wrap, controls.cmp.wrap), sep(),
      el('div', { class: 'tgroup', role: 'group', 'aria-label': 'Метрики' }, controls.block.wrap, controls.unit.wrap), sep(),
      el('div', { class: 'tgroup', role: 'group', 'aria-label': 'Рынок' }, controls.nets.wrap, controls.fos.wrap, ...(regionLevel ? [controls.regs.wrap] : [])));
    return bar;
  }
  // Сбоку: продукт
  function buildSide() {
    const side = el('aside', { class: 'side', 'aria-label': 'Продукт' });
    controls.cats = filterList('Категория', state.cats, () => META.categories.map((t, i) => ({ value: i, text: t })), () => {
      prune(state.prods, availableProducers()); prune(state.brands, availableBrands());
      controls.prods.refresh(); controls.brands.refresh(); schedule();
    });
    controls.prods = filterList('Производитель', state.prods, availableProducers, () => {
      prune(state.brands, availableBrands()); controls.brands.refresh(); schedule();
    });
    controls.brands = filterList('Бренд', state.brands, availableBrands, schedule);
    side.append(el('div', { class: 'side-title', text: 'Продукт' }), controls.cats.wrap, controls.prods.wrap, controls.brands.wrap);
    return side;
  }
  function syncControls() { for (const c of Object.values(controls)) c.refresh(); }

  // Выбранные срезы — плашками над показателями; крестик убирает срез
  const chipsBox = el('div', { class: 'chips', role: 'list', 'aria-label': 'Выбранные срезы', hidden: true });
  const DIMS = [['nets', 'Сеть', i => META.nets[i]], ['fos', 'ФО', i => META.fos[i]], ['regs', 'Регион', i => META.regions[i]],
    ['cats', 'Категория', i => META.categories[i]], ['prods', 'Производитель', i => META.producers[i]], ['brands', 'Бренд', i => META.brands[i]]];
  function dropFilter(dim, value) {
    if (value == null) state[dim].clear(); else state[dim].delete(value);
    if (dim === 'fos' && regionLevel) prune(state.regs, availableRegions());
    if (dim === 'cats') { prune(state.prods, availableProducers()); prune(state.brands, availableBrands()); }
    if (dim === 'prods') prune(state.brands, availableBrands());
    syncControls(); schedule();
  }
  function chip(k, v, onRemove) {
    const x = el('button', { type: 'button', title: 'Убрать', 'aria-label': `Убрать срез ${k} ${v}` }, icon('x'));
    x.addEventListener('click', onRemove);
    return el('span', { class: 'chip', role: 'listitem' }, el('span', { class: 'k', text: k + ':' }), el('span', { class: 'v', text: v, title: v }), x);
  }
  function renderChips() {
    chipsBox.textContent = '';
    let any = false;
    for (const [dim, label, nameOf] of DIMS) {
      const set = state[dim];
      if (!set.size) continue;
      any = true;
      if (set.size <= 3) for (const v of set) chipsBox.append(chip(label, nameOf(v), () => dropFilter(dim, v)));
      else chipsBox.append(chip(label, `${set.size} выбрано`, () => dropFilter(dim, null)));
    }
    if (any) {
      const clr = el('button', { class: 'clear', type: 'button', text: 'Очистить срезы' });
      clr.addEventListener('click', () => { for (const [dim] of DIMS) state[dim].clear(); syncControls(); schedule(); });
      chipsBox.append(clr);
    }
    chipsBox.hidden = !any;
  }

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
  const prodFilter = () => state.prods.size > 0 || state.brands.size > 0;
  const catFilter = () => state.cats.size > 0;
  const idxUnit = () => (state.unit === 'G' ? 'кг' : 'шт');
  const KPI = {
    sales: [() => `Продажи${bwSp()}, ${mu()}`, k => fmt.compact(k.sales), k => delta(growth(k.sales, k.salesCmp), 'pct'), goodUp,
      'Продажи выбранного блока за период'],
    share: [() => `Доля в категории${bw() ? ` (${bw()})` : ''}, ${mu()}`, k => fmt.pct(k.share), k => dPP(k.share, k.shareCmp), goodUp,
      'Продажи выбранных производителей и брендов к продажам категорий из среза в тех же сетях, регионах и периоде'],
    sharePbp: [() => `Доля в ПБП${bw() ? ` (${bw()})` : ''}, ${mu()}`, k => fmt.pct(k.sharePbp), k => dPP(k.sharePbp, k.sharePbpCmp), goodUp,
      'Продажи выбранного к продажам всей ПБП — всех категорий — в тех же сетях, регионах и периоде'],
    promoShare: [() => (state.block === 'P' ? `Доля промо в продажах, ${mu()}` : `Доля промо, ${mu()}`), k => fmt.pct(k.promoShare),
      k => dPP(k.promoShare, k.promoShareCmp), neutralTone, 'Продажи по промо к всем продажам'],
    loyalShare: [() => (state.block === 'L' ? `Доля по картам в продажах, ${mu()}` : `Доля по картам, ${mu()}`), k => fmt.pct(k.loyalShare),
      k => dPP(k.loyalShare, k.loyalShareCmp), neutralTone, 'Продажи покупателям с картами лояльности к всем продажам'],
    price: [() => `Цена${bwSp()} за шт`, k => fmt.rub(k.price), k => dPct(k.price, k.priceCmp), neutralTone, 'Рубли к штукам'],
    priceKg: [() => `Цена${bwSp()} за кг/л`, k => fmt.rub(k.priceKg), k => dPct(k.priceKg, k.priceKgCmp), neutralTone, 'Рубли к килограммам (литрам)'],
    index: [() => `Индекс цены${bwSp()} за ${idxUnit()}`, k => fmt.idx(k.index), k => (ok(k.index) && ok(k.indexCmp) ? delta(k.index - k.indexCmp, 'idx') : null),
      neutralTone, 'Цена выбранного к цене категорий из среза; 1,00 — на уровне категории. За кг — когда выбрана единица «кг»'],
    depth: [() => 'Глубина промо', k => fmt.pct(k.depth), k => dPP(k.depth, k.depthCmp), neutralTone,
      'На сколько промо-цена ниже обычной цены тех же SKU, взвешено по промо-штукам'],
    tdp: [() => 'TDP средняя', k => fmt.tdp(k.tdp), k => (ok(k.tdp) && ok(k.tdpCmp) ? delta(100 * (k.tdp - k.tdpCmp), 'pt') : null), goodUp,
      'Сумма ND выбранных SKU в пунктах, среднее по периодам'],
  };
  // Доля в категории — когда выбраны производители или бренды; доля в ПБП — когда выбраны категории
  function kpiKeys() {
    const shares = [...(prodFilter() ? ['share'] : []), ...(catFilter() ? ['sharePbp'] : [])];
    const idx = prodFilter() ? ['index'] : [];
    if (state.block === 'T') return ['sales', ...shares, 'promoShare', 'loyalShare', 'price', 'priceKg', ...idx, 'tdp'];
    if (state.block === 'P') return ['sales', ...shares, 'promoShare', 'price', 'priceKg', ...idx, 'depth'];
    return ['sales', ...shares, 'loyalShare', 'price', 'priceKg', ...idx];
  }
  function renderKpis(r) {
    const k = r.kpi;
    kpiBox.textContent = '';
    for (const key of kpiKeys()) {
      const [label, value, change, tone, hint] = KPI[key];
      const d = el('div', { class: 'delta' });
      setDelta(d, k.hasCmp ? change(k) : null, tone);
      kpiBox.append(el('div', { class: 'tile', title: hint, 'data-kpi': key }, el('div', { class: 'label', text: label() }), el('div', { class: 'value', text: value(k) }), d));
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
  // Вид колонки определяет и показ, и выгрузку: text, sales, pct (доля), dpct (прирост), pp (изменение доли), dpt (изменение TDP),
  // rub, idx, tdp, ros (продажи на магазин), int (штуки)
  const titleOf = c => (typeof c.title === 'function' ? c.title() : c.title);
  const scaled = (x, mul, d) => (ok(x) ? r2(x * mul, d) : null);
  fmt.ros = v => (ok(v) ? (Math.abs(v) >= 100 ? nf0 : Math.abs(v) >= 10 ? nf1 : nf2).format(v) : '—');
  function exportColumns(c) {
    const t = titleOf(c), v = r => r[c.key];
    if (c.exp) return c.exp;
    switch (c.kind) {
      case 'sales': return [[t, r => scaled(v(r), 1, salesDigits())]];
      case 'pct': case 'dpct': return [[t + ', %', r => scaled(v(r), 100, 2)]];
      case 'pp': return [[t + ', п.п.', r => scaled(v(r), 1, 2)]];
      case 'dpt': return [[t + ', пункты', r => scaled(v(r), 1, 1)]];
      case 'rub': return [[t + ', руб', r => scaled(v(r), 1, 2)]];
      case 'idx': return [[t, r => scaled(v(r), 1, 3)]];
      case 'tdp': return [[t, r => scaled(v(r), 100, 1)]];
      case 'ros': return [[t, r => scaled(v(r), 1, 3)]];
      default: return [[t, v]];
    }
  }
  function deltaCell(v, kind) {
    const d = delta(v, kind);
    if (!d) return el('td', { class: 'num none', text: '—' });
    if (!d.dir) return el('td', { class: 'num neutral', text: d.text });
    return el('td', { class: 'num ' + (d.dir > 0 ? 'good' : 'bad') }, el('span', { 'aria-hidden': 'true', text: d.dir > 0 ? '▲ ' : '▼ ' }), document.createTextNode(d.text));
  }
  function cellOf(col, row) {
    if (col.render) return col.render(row);
    const v = row[col.key];
    if (col.kind === 'dpct') return deltaCell(v, 'pct');
    if (col.kind === 'pp') return deltaCell(v, 'pp');
    if (col.kind === 'dpt') return deltaCell(v, 'pt');
    const f = { sales: fmt.sales, pct: fmt.pct, rub: fmt.rub, idx: fmt.idx, tdp: fmt.tdp, ros: fmt.ros, int: fmt.int }[col.kind];
    return el('td', { class: f ? 'num' : 'name', text: f ? f(v) : String(v ?? '') });
  }
  // Таблица во вкладке: сортировка по заголовку, свои строки сверху (pinOwn), поиск, раскрытие до SKU (children)
  function dataTable(columns, opts) {
    const wrapT = el('div', { class: 'table-wrap' });
    const more = el('button', { class: 'btn quiet small more', type: 'button', hidden: true });
    const node = el('div', { role: 'tabpanel', hidden: true }, wrapT, more);
    const t = { node, opts, rows: [], total: null, sortKey: opts.sortKey, dir: -1, expanded: false, query: '', open: new Set(), visible: false, dirty: true };
    more.addEventListener('click', () => { t.expanded = !t.expanded; draw(); });
    const visibleCols = () => columns.filter(c => !c.when || c.when());
    function sorter() {
      const cols = visibleCols(), col = cols.find(x => x.key === t.sortKey) || cols.find(x => x.kind === 'sales');
      const val = r => { const v = r[col.key]; return ok(v) ? v : (typeof v === 'string' ? v : -Infinity); };
      return (a, b) => {
        if (opts.pinOwn && !!a.own !== !!b.own) return a.own ? -1 : 1;
        const x = val(a), y = val(b);
        return (typeof x === 'string' ? String(x).localeCompare(y, 'ru') : x - y) * t.dir;
      };
    }
    const hit = (r, q) => (r.search || r.name || '').toLowerCase().includes(q);
    function draw() {
      if (!t.visible) { t.dirty = true; return; }
      t.dirty = false;
      wrapT.textContent = '';
      const q = t.query.trim().toLowerCase();
      let rows = q ? t.rows.filter(r => hit(r, q) || (opts.children && r.children.some(ch => hit(ch, q)))) : t.rows;
      if (!rows.length) {
        wrapT.append(el('div', { class: 'empty', text: q ? 'Ничего не найдено' : 'Нет данных для выбранных срезов' }));
        more.hidden = true;
        return;
      }
      const sort = sorter();
      rows = [...rows].sort(sort);
      const limit = q ? 0 : opts.limit;
      const shown = t.expanded || !limit ? rows : rows.slice(0, limit);
      const cols = visibleCols();
      const head = el('tr');
      for (const cdef of cols) {
        const isNum = cdef.kind !== 'text';
        const th = el('th', { class: isNum ? 'num' : '', text: titleOf(cdef), scope: 'col', tabindex: 0, title: cdef.hint || null });
        if (cdef.key === t.sortKey) th.setAttribute('aria-sort', t.dir > 0 ? 'ascending' : 'descending');
        const sortBy = () => { if (t.sortKey === cdef.key) t.dir = -t.dir; else { t.sortKey = cdef.key; t.dir = isNum ? -1 : 1; } draw(); };
        th.addEventListener('click', sortBy);
        th.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); sortBy(); } });
        head.append(th);
      }
      const tb = el('tbody');
      shown.forEach((r, i) => {
        const next = shown[i + 1], pinEnd = opts.pinOwn && r.own && next && !next.own;
        const tr = el('tr', { class: [r.own ? 'own' : '', pinEnd ? 'pin-end' : ''].join(' ').trim() || null }, ...cols.map(cdef => cellOf(cdef, r)));
        tb.append(tr);
        if (!opts.children) return;
        const opened = t.open.has(r.id) || (q && !hit(r, q));
        tr.classList.add('line');
        tr.setAttribute('aria-expanded', String(opened));
        tr.addEventListener('click', () => { if (t.open.has(r.id)) t.open.delete(r.id); else t.open.add(r.id); draw(); });
        if (!opened) return;
        const kids = q && !hit(r, q) ? r.children.filter(ch => hit(ch, q)) : r.children;
        for (const ch of [...kids].sort(sort)) tb.append(el('tr', { class: 'child' }, ...cols.map(cdef => cellOf(cdef, ch))));
      });
      if (t.total) tb.append(el('tr', { class: 'total' }, ...cols.map(cdef => cellOf(cdef, t.total))));
      wrapT.append(el('table', { class: 'data' + (opts.cls ? ' ' + opts.cls : '') }, el('thead', {}, head), tb));
      if (limit && rows.length > limit) {
        more.hidden = false;
        more.textContent = t.expanded ? 'Свернуть' : `Показать все (${rows.length})`;
      } else more.hidden = true;
    }
    t.csv = () => {
      const cols = [...(opts.children ? [['Уровень', r => (r.children ? 'линия' : 'SKU')]] : []), ...visibleCols().flatMap(exportColumns)];
      const lines = [cols.map(c => csvCell(c[0])).join(';')], sort = sorter();
      for (const r of [...t.rows].sort(sort)) {
        lines.push(cols.map(c => csvCell(c[1](r))).join(';'));
        if (opts.children) for (const ch of [...r.children].sort(sort)) lines.push(cols.map(c => csvCell(c[1](ch))).join(';'));
      }
      if (t.total) lines.push(cols.map(c => csvCell(c[1](t.total))).join(';'));
      return [lines.join('\r\n') + '\r\n'];
    };
    t.set = (rows, total) => { t.rows = rows; t.total = total || null; draw(); };
    t.show = on => { t.visible = on; node.hidden = !on; if (on && t.dirty) draw(); };
    t.setQuery = q => { t.query = q; t.expanded = false; draw(); };
    t.allOpen = () => t.rows.length > 0 && t.rows.every(r => t.open.has(r.id));
    t.openAll = on => { t.open = on ? new Set(t.rows.map(r => r.id)) : new Set(); draw(); };
    return t;
  }
  // Карточка с вкладками разрезов; последняя открытая вкладка запоминается
  function tabsCard(defs, initial) {
    const tablist = el('div', { class: 'tabs', role: 'tablist', 'aria-label': 'Разрезы' });
    const search = el('input', { class: 'tab-search', type: 'search', placeholder: 'Поиск в таблице', 'aria-label': 'Поиск в таблице' });
    const openBtn = el('button', { class: 'btn quiet small', type: 'button' });
    const csvBtn = el('button', { class: 'btn small', type: 'button', title: 'Скачать таблицу целиком для Excel' }, icon('download'), 'CSV');
    const note = el('div', { class: 'tab-note' });
    const body = el('div', { class: 'tab-body' });
    const node = el('section', { class: 'card tabs-card', 'aria-label': 'Разрезы' },
      el('div', { class: 'tabs-head' }, tablist, el('div', { class: 'tab-tools' }, search, openBtn, csvBtn)), note, body);
    let active = defs[0];
    const buttons = defs.map(d => {
      const b = el('button', { class: 'tab', type: 'button', role: 'tab', text: d.label, 'aria-selected': 'false', tabindex: -1 });
      b.addEventListener('click', () => select(d.id));
      body.append(d.table.node);
      return b;
    });
    tablist.append(...buttons);
    tablist.addEventListener('keydown', e => {
      if (e.key !== 'ArrowRight' && e.key !== 'ArrowLeft') return;
      e.preventDefault();
      const i = (defs.indexOf(active) + (e.key === 'ArrowRight' ? 1 : defs.length - 1)) % defs.length;
      select(defs[i].id);
      buttons[i].focus();
    });
    const paintOpen = () => { openBtn.hidden = !active.table.opts.children; openBtn.textContent = active.table.allOpen() ? 'Свернуть все' : 'Развернуть все'; };
    function refresh() { note.textContent = typeof active.note === 'function' ? active.note() : active.note; paintOpen(); }
    function select(id) {
      active = defs.find(d => d.id === id) || defs[0];
      defs.forEach((d, i) => {
        const on = d === active;
        buttons[i].setAttribute('aria-selected', String(on));
        buttons[i].tabIndex = on ? 0 : -1;
        d.table.show(on);
      });
      search.hidden = !active.searchable;
      search.value = active.table.query;
      refresh();
      try { localStorage.setItem('x5-tab', active.id); } catch (e) { /* без сохранения */ }
    }
    search.addEventListener('input', () => { active.table.setQuery(search.value); paintOpen(); });
    openBtn.addEventListener('click', () => { active.table.openAll(!active.table.allOpen()); paintOpen(); });
    csvBtn.addEventListener('click', () => saveCsv(fileName(active.file), active.table.csv()));
    let saved = null;
    try { saved = localStorage.getItem('x5-tab'); } catch (e) { /* по умолчанию */ }
    select(defs.some(d => d.id === saved) ? saved : initial);
    return { node, refresh };
  }

  // ---------- справочник ----------
  const help = document.getElementById('help');
  const perfText = el('span', { id: 'perf' });
  function buildHelp() {
    const sec = (title, ...kids) => el('section', {}, el('h3', { text: title }), ...kids);
    const p = text => el('p', { text });
    const dl = pairs => el('dl', {}, ...pairs.flatMap(([a, b]) => [el('dt', { text: a }), el('dd', { text: b })]));
    const close = el('button', { class: 'icon-btn', type: 'button', title: 'Закрыть', 'aria-label': 'Закрыть справочник' }, icon('x'));
    close.addEventListener('click', () => help.close());
    const periodWords = weekly ? 'неделям' : 'месяцам';
    help.append(el('div', { class: 'help-head' }, el('h2', { id: 'help-title', text: 'Справочник' }), close), el('div', { class: 'help-body' },
      sec('Как пользоваться', dl([
        ['Период', `«До» — ${weekly ? 'последняя неделя' : 'последний месяц'} периода, рядом — длина окна и база сравнения.`],
        ['Метрики', 'Блок X5 (Тотал, Промо, Лояльные) и единица: штуки, рубли или килограммы.'],
        ['Рынок', `Сети, ФО${regionLevel ? ' и регионы' : ''} — выпадающие списки вверху, можно выбрать несколько.`],
        ['Продукт', 'Категории, производители и бренды — на панели слева. Маревен всегда первым, остальные по алфавиту.'],
        ['Срезы', 'Выбранное видно плашками над показателями, крестик убирает срез. «Сбросить» возвращает всё по умолчанию.'],
        ['Таблицы', 'Вкладки внизу. Заголовок столбца сортирует, CSV выгружает таблицу целиком. Продуктовую линию можно раскрыть до SKU.'],
      ])),
      sec('Период', p(weekly
        ? 'Нед — одна неделя; L4W и L13W — последние 4 и 13 недель; MTD — с начала месяца 4-4-5 до выбранной недели; YTD — с начала года; MAT — последние 52 недели.'
        : 'Мес — один месяц; L3M — последние 3 месяца; YTD — с начала года до выбранного месяца; MAT — последние 12 месяцев.'),
        p(weekly
          ? 'YA — те же недели ISO годом раньше; у 53-й недели пары нет, это видно в строке периода. PP — столько же недель сразу перед выбранным периодом.'
          : 'YA — те же месяцы годом раньше. PP — столько же месяцев сразу перед выбранным периодом.')),
      sec('Блоки и единицы', p('Тотал — все продажи, Промо — продажи по промо, Лояльные — продажи покупателям с картами лояльности X5. Продажи, доли, прирост и цены считаются внутри выбранного блока: в блоке «Промо» доля в категории — промо-продажи бренда к промо-продажам категории. Доля промо и доля по картам — от всех продаж.')),
      sec('Метрики', dl([
        ['Доля в категории', 'Продажи выбранного к продажам категорий из среза «Категория» в тех же сетях, регионах и периоде; срезы производителя и бренда на знаменатель не влияют. Показывается, когда выбраны производители или бренды.'],
        ['Доля в ПБП', 'Продажи выбранного к продажам всей продукции быстрого приготовления — всех категорий — в тех же сетях, регионах и периоде. Показывается, когда выбраны категории.'],
        ['Изменение доли', 'В процентных пунктах к периоду сравнения. В таблицах ФО, сетей и регионов доля считается внутри своего округа, сети или региона.'],
        ['Индекс цены', 'Цена выбранного к цене категорий из среза в том же блоке: за штуку, а при единице «кг» — за килограмм. 1,00 — на уровне категории.'],
        ['Глубина промо', 'На сколько промо-цена ниже обычной цены тех же SKU в той же сети, регионе и периоде, взвешено по промо-штукам. Строки, где SKU продавался только по промо, пропускаются.'],
        ['ND', `Доля магазинов с продажами SKU среди магазинов с продажами всей категории БП, среднее по ${periodWords}. Магазин засчитывается, только если в нём была продажа.`],
        ['TDP', 'Сумма ND SKU в пунктах: TDP 250 — в среднем 2,5 SKU в магазине. TDP продуктовой линии — сумма ND её SKU. В Excel та же мера в процентах: 250% там — это 250 здесь.'],
        ['Продажи на магазин', `Продажи SKU к числу магазинов с его продажами, в среднем за ${weekly ? 'неделю' : 'месяц'}. Для линии — то же по всем её SKU.`],
        ['Чего здесь нет', 'Чеки, покупатели и взвешенная дистрибуция: для брендов и линий из данных X5 их честно не посчитать.'],
      ])),
      sec('Выгрузка', p(`CSV у таблицы сохраняет её целиком. «Скачать данные» — строки по ${periodWords}, сетям, ${regionLevel ? 'регионам' : 'ФО'} и SKU за выбранный период и срезы со всеми метриками. Файлы открываются в Excel: разделитель «;», десятичная запятая.`)),
      sec('Источник', p(`X5 Dialog, отчёт «Анализ трендов», выгрузка из ClickHouse. ${regionLevel ? 'Данные по регионам X5, ФО — по справочнику регионов.' : 'Регионы сведены в федеральные округа.'}`))));
    help.addEventListener('click', e => { if (e.target === help) help.close(); });
  }
  helpBtn.addEventListener('click', () => help.showModal());
  document.addEventListener('keydown', e => {
    if (e.key !== '?' || e.ctrlKey || e.metaKey || e.altKey || help.open) return;
    if (/^(INPUT|SELECT|TEXTAREA)$/.test(document.activeElement && document.activeElement.tagName)) return;
    e.preventDefault();
    help.showModal();
  });

  // ---------- сборка страницы ----------
  app.textContent = '';
  buildHelp();
  const periodText = el('span');
  const exportStatus = el('span', { class: 'status', role: 'status' });
  const exportBtn = el('button', { class: 'btn small', type: 'button',
    title: `Строки по ${weekly ? 'неделям' : 'месяцам'}, сетям, ${regionLevel ? 'регионам' : 'ФО'} и SKU за выбранный период и срезы` }, icon('download'), 'Скачать данные');
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
  const toolbar = buildToolbar();
  app.append(toolbar, el('div', { class: 'layout' }, buildSide(), main));
  new ResizeObserver(() => document.documentElement.style.setProperty('--bar-h', toolbar.offsetHeight + 'px')).observe(toolbar);
  main.append(el('div', { class: 'context' }, periodLine, chipsBox), kpiBox);
  const periodWord = weekly ? 'неделям' : 'месяцам';
  const lastWord = weekly ? 'Последние 52 недели' : 'Последние 12 месяцев';
  const c1 = card('', `${lastWord} и YA · выбранный период выделен фоном`);
  const c2 = card('', `${lastWord} · топ-5 за выбранный период · не зависит от срезов производителя и бренда`);
  const c3 = card('', 'Топ-10 за выбранный период · не зависит от срезов производителя и бренда');
  const c4 = card('', 'Топ-10 производителей · без срезов производителя и бренда');
  main.append(el('div', { class: 'grid2' }, c1.node, c2.node), el('div', { class: 'grid2' }, c3.node, c4.node));

  const ownMark = (text, own) => (own ? el('span', { class: 'own', text }) : document.createTextNode(text));
  const nameCol = title => ({ key: 'name', kind: 'text', title, render: r => el('td', { class: 'name', title: r.name }, ownMark(r.name, r.own)) });
  const salesTitle = () => `Продажи${bwSp()}, ${mu()}`;
  const growthTitle = () => `Прирост к ${state.cmp}`, shareTitle = () => `Изменение доли к ${state.cmp}`;
  const salesCol = { key: 'sales', kind: 'sales', title: salesTitle };
  const growthCol = { key: 'growth', kind: 'dpct', title: growthTitle };
  const shareCol = { key: 'share', kind: 'pct', title: () => `Доля в категории${bw() ? ` (${bw()})` : ''}`,
    hint: 'К продажам категорий из среза «Категория» в тех же сетях, регионах и периоде' };
  const dShareCol = { key: 'dShare', kind: 'pp', title: shareTitle };
  const sharePbpCol = { key: 'sharePbp', kind: 'pct', title: () => `Доля в ПБП${bw() ? ` (${bw()})` : ''}`, when: catFilter,
    hint: 'К продажам всей ПБП в тех же сетях, регионах и периоде' };
  const promoCol = { key: 'promoShare', kind: 'pct', title: () => (state.block === 'P' ? 'Доля промо в продажах' : 'Доля промо'), when: () => state.block !== 'L' };
  const loyalCol = { key: 'loyalShare', kind: 'pct', title: () => (state.block === 'L' ? 'Доля по картам в продажах' : 'Доля по картам'), when: () => state.block !== 'P' };
  const priceCols = [{ key: 'price', kind: 'rub', title: () => `Цена${bwSp()} за шт` }, { key: 'priceKg', kind: 'rub', title: () => `Цена${bwSp()} за кг/л` }];
  const indexCol = { key: 'index', kind: 'idx', title: () => `Индекс цены за ${idxUnit()}`, hint: '1,00 — цена на уровне категорий из среза' };
  const depthCol = { key: 'depth', kind: 'pct', title: 'Глубина промо', when: () => state.block === 'P' };
  const tdpCol = { key: 'tdp', kind: 'tdp', title: 'TDP средняя', when: () => state.block === 'T' };

  const tCats = dataTable([
    { key: 'name', kind: 'text', title: 'Категория', render: r => el('td', { class: 'name', title: r.name }, ownMark(r.name, r.sel)) },
    salesCol, growthCol,
    { key: 'share', kind: 'pct', title: () => `Доля в ПБП${bw() ? ` (${bw()})` : ''}` },
    { key: 'dShare', kind: 'pp', title: shareTitle },
    ...(META.own >= 0 ? [{ key: 'ownShare', kind: 'pct', title: () => `Доля ${META.producers[META.own]}`, hint: 'Доля Маревен в категории' },
      { key: 'dOwnShare', kind: 'pp', title: () => `Изменение доли ${META.producers[META.own]}` }] : []),
    promoCol, loyalCol, ...priceCols,
  ], { sortKey: 'sales', file: 'категории' });
  const tProds = dataTable([nameCol('Производитель'), salesCol, growthCol, shareCol, dShareCol, sharePbpCol, promoCol, loyalCol, ...priceCols,
    indexCol, depthCol, tdpCol], { sortKey: 'sales', limit: 25, pinOwn: true, file: 'производители' });
  const tBrands = dataTable([nameCol('Бренд'), salesCol, growthCol, shareCol, dShareCol, sharePbpCol, promoCol, loyalCol, ...priceCols,
    indexCol, depthCol, tdpCol], { sortKey: 'sales', limit: 25, pinOwn: true, file: 'бренды' });
  const lineName = li => META.lines[li];
  const tLines = dataTable([
    { key: 'name', kind: 'text', title: 'Продуктовая линия · SKU',
      exp: [['Продуктовая линия', r => r.lineName], ['PLU', r => r.plu || ''], ['SKU', r => r.skuName || ''], ['Бренд', r => r.brandName]],
      render: r => (r.children
        ? el('td', { class: 'name', title: `${r.name} · ${r.brandName}` }, el('button', { class: 'tree-btn', type: 'button', 'aria-label': `SKU линии ${r.name}` }, icon('right')),
          ownMark(r.name, r.own), el('span', { class: 'plu', text: r.brandName }))
        : el('td', { class: 'name', title: `${r.name} · PLU ${r.plu}` }, el('span', { class: 'nm-cut', text: r.name }), el('span', { class: 'plu', text: r.plu }))) },
    salesCol, growthCol, shareCol, dShareCol,
    { key: 'dist', kind: 'tdp', title: 'TDP · ND', hint: 'Для линии — TDP, сумма ND её SKU; для SKU — ND, доля магазинов с продажами',
      exp: [['TDP линии · ND SKU, %', r => scaled(r.dist, 100, 1)]],
      render: r => el('td', { class: 'num', text: r.children ? fmt.tdp(r.dist) : fmt.pct(r.dist) }) },
    { key: 'dDist', kind: 'dpt', title: () => `Изменение TDP · ND к ${state.cmp}`, render: r => deltaCell(r.dDist, r.children ? 'pt' : 'pp') },
    { key: 'ros', kind: 'ros', title: () => `Продажи на магазин, ${mu()}`,
      hint: `Продажи SKU к числу магазинов с его продажами, в среднем за ${weekly ? 'неделю' : 'месяц'}; для линии — по всем её SKU` },
    { key: 'skuCount', kind: 'int', title: 'SKU', hint: 'SKU линии с продажами за период', render: r => el('td', { class: 'num', text: r.children ? fmt.int(r.skuCount) : '' }) },
    promoCol, { ...loyalCol, when: () => state.block === 'L' }, ...priceCols, indexCol, sharePbpCol,
  ], { sortKey: 'sales', limit: 30, pinOwn: true, children: true, cls: 'lines', file: 'продуктовые_линии' });
  const geoColumns = extra => [...extra, salesCol, growthCol,
    { key: 'share', kind: 'pct', title: () => `Доля в категории${bw() ? ` (${bw()})` : ''}` }, dShareCol];
  const tNets = dataTable(geoColumns([{ key: 'name', kind: 'text', title: 'Сеть' }]), { sortKey: 'sales', file: 'сети' });
  const tFo = dataTable(geoColumns([{ key: 'name', kind: 'text', title: 'ФО' }]), { sortKey: 'sales', file: 'ФО' });
  const tRegions = regionLevel ? dataTable(geoColumns([{ key: 'name', kind: 'text', title: 'Регион' }, { key: 'foName', kind: 'text', title: 'ФО' }]),
    { sortKey: 'sales', limit: 30, file: 'регионы' }) : null;
  const tabs = tabsCard([
    { id: 'categories', label: 'Категории', table: tCats, file: 'категории', note: 'Вся ПБП в выбранных сетях и регионах · срезы продукта не применяются · доля — в ПБП' },
    { id: 'producers', label: 'Производители', table: tProds, file: 'производители', searchable: true,
      note: 'Производители категорий из среза · срезы производителя и бренда не применяются · Маревен сверху' },
    { id: 'brands', label: 'Бренды', table: tBrands, file: 'бренды', searchable: true,
      note: 'Доля и индекс цены — от категорий из среза «Категория» в тех же сетях, регионах и периоде · бренды Маревен сверху' },
    { id: 'lines', label: 'Продуктовые линии', table: tLines, file: 'продуктовые_линии', searchable: true,
      note: 'Нажмите на линию, чтобы увидеть её SKU с дистрибуцией · TDP линии — сумма ND её SKU · поиск находит и линии, и SKU' },
    { id: 'nets', label: 'Сети', table: tNets, file: 'сети', note: 'Доля — в категории своей сети' },
    { id: 'fos', label: 'ФО', table: tFo, file: 'ФО', note: 'Доля — в категории своего округа' },
    ...(tRegions ? [{ id: 'regions', label: 'Регионы', table: tRegions, file: 'регионы', searchable: true, note: 'Доля — в категории своего региона' }] : []),
  ], 'brands');
  main.append(tabs.node, el('footer', {}, el('span', { text: 'Источник: X5 Dialog, «Анализ трендов», выгрузка из ClickHouse' }), perfText,
    el('span', { text: 'Справочник — кнопка «?» вверху' })));

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
    // у производителя вне восьми постоянных цветов — свободный цвет этого графика, синий остаётся за Маревен
    const used = new Set(r.chart2.series.map(s => META.slots[s.prod]).filter(x => x >= 0));
    const spare = [1, 2, 3, 4, 5, 6, 7].filter(x => !used.has(x));
    const c2color = s => (META.slots[s.prod] >= 0 ? prodColor(s.prod) : spare.length ? color(spare.shift()) : 'var(--other)');
    lineChart(c2, { labels, xLabel, tipTitle, band: r.chart1.band, bandLabel, fmt: fmt.pct, axisFmt: pctAxis, sortTip: true,
      aria: 'Доля производителей в категории по периодам',
      series: r.chart2.series.map(s => ({ name: META.producers[s.prod], color: c2color(s), values: s.values })) });
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
  const tables = () => ({ categories: tCats, producers: tProds, brands: tBrands, lines: tLines, nets: tNets, fos: tFo, regions: tRegions });
  function renderAll() {
    const t0 = performance.now();
    const r = compute(state);
    last = r;
    renderPeriod(r.win);
    renderChips();
    renderKpis(r);
    drawCharts(r);
    tCats.set(r.categories.map(c => ({ ...c, name: META.categories[c.cat], sel: state.cats.has(c.cat) })), { ...r.catTotal, name: 'ПБП, всего' });
    tProds.set(r.producers.map(x => ({ ...x, name: META.producers[x.prod], own: x.prod === META.own })));
    tBrands.set(r.brands.map(b => ({ ...b, name: META.brands[b.brand], own: isOwnBrand(b.brand) })));
    tLines.set(r.lines.map(l => ({ ...l, id: l.line, name: lineName(l.line), lineName: lineName(l.line), brandName: l.brand >= 0 ? META.brands[l.brand] : '',
      dist: l.tdp, dDist: l.dTdp, children: l.skus.map(s => {
        const nm = META.products.name[s.k], plu = META.products.plu[s.k];
        return { ...s, name: nm, plu, skuName: nm, lineName: lineName(l.line), brandName: META.brands[brandOf[s.k]], dist: s.nd, dDist: s.dNd,
          search: `${nm} ${plu}` };
      }) })));
    tFo.set(r.fos.map(f => ({ ...f, name: META.fos[f.fo] })), { ...r.total, name: 'Итого' });
    tNets.set(r.nets.map(x => ({ ...x, name: META.nets[x.net] })), { ...r.total, name: 'Итого' });
    if (tRegions) tRegions.set(r.regions.map(x => ({ ...x, name: META.regions[x.geo], foName: META.fos[x.fo] })), { ...r.total, name: 'Итого', foName: '' });
    tabs.refresh();
    perfText.textContent = `Строк данных: ${nf0.format(D.P.length)} · пересчёт ${Math.round(performance.now() - t0)} мс`;
    document.body.dataset.ready = '1';
  }
  let resizeTimer = 0;
  new ResizeObserver(() => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => last && drawCharts(last), 120); }).observe(main);
  window.X5 = {
    meta: META, state, compute, periodWindow, fileName, setTheme,
    set(patch) {
      for (const [k, v] of Object.entries(patch)) {
        if (v instanceof Array) { state[k].clear(); for (const x of v) state[k].add(x); } else state[k] = v;
      }
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
