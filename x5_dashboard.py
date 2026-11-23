"""Автономный HTML-дашборд X5 Dialog по файлам куба.

Запуск:  python x5_dashboard.py   (после x5_build.py)
Нужно:   pandas и numpy (уже стоят для x5_build.py)

Что делает:
  1. Читает CSV куба (факты, итоги категории, календарь, продукты, регионы).
  2. Сворачивает факты до неделя (месяц) × сеть × ФО × PLU. Части взвешенной глубины промо считаются
     до свёртки, на строках регионов, поэтому мера совпадает с Excel.
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
START_YEAR = None                         # например 2024, чтобы облегчить файл; у первого года не будет ПГ
CONTOURS = {"weekly": "недели", "monthly": "месяцы"}
TITLE = "X5 Dialog · Продукция быстрого приготовления"

FACT_COLS = ["TOTAL_Шт", "TOTAL_Руб", "TOTAL_КгЛ", "PROMO_Шт", "PROMO_Руб", "TOTAL_Магазины"]  # промо-рубли нужны для глубины


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
    cal["НачалоПериода"] = pd.to_datetime(cal["НачалоПериода"])
    cal["НачалоПериодаПГ"] = pd.to_datetime(cal["НачалоПериодаПГ"], errors="coerce")
    for c in ["Год", "Месяц"] + (["Неделя"] if weekly else []):
        cal[c] = cal[c].astype(int)
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
    fact["p"] = codes(fact["НачалоПериода"], list(per_index))
    fact["n"] = codes(fact["Сеть"], nets)
    fact["f"] = codes(fact["ФО"], fos)
    fact["k"] = codes(fact["PLU"], list(products["PLU"]))
    values = FACT_COLS + ["ГлубинаПолн", "ГлубинаОпл"]
    agg = fact.groupby(["k", "n", "f", "p"], as_index=False, sort=True)[values].sum()
    agg = agg[(agg[values] != 0).any(axis=1)].reset_index(drop=True)

    cat = read(f"x5_{contour}_category", usecols=["НачалоПериода", "Сеть", "Регион", "TOTAL_Магазины"], dtype=keys)
    cat["НачалоПериода"] = pd.to_datetime(cat["НачалоПериода"])
    cat = cat[cat["НачалоПериода"].isin(per_index) & cat["Сеть"].isin(nets)]
    cat["ФО"] = cat["Регион"].map(reg_fo)
    cat = cat[cat["ФО"].isin(fos)]
    cat["p"], cat["n"], cat["f"] = codes(cat["НачалоПериода"], list(per_index)), codes(cat["Сеть"], nets), codes(cat["ФО"], fos)
    cagg = cat.groupby(["p", "n", "f"], as_index=False, sort=True)["TOTAL_Магазины"].sum()

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
                    "month": cal["Месяц"].tolist(), "start": cal["НачалоПериода"].dt.strftime("%d.%m.%Y").tolist(),
                    "pg": pg.tolist()},
        "nets": nets, "fos": fos, "categories": categories, "producers": producers, "brands": brands,
        "products": {"plu": products["PLU"].tolist(), "name": products["Название"].tolist(),
                     "cat": codes(products["Категория"], categories).tolist(),
                     "prod": codes(products["Производитель (корректный)"], producers).tolist(),
                     "brand": codes(products["Бренд (корректный)"], brands).tolist()},
        "own": producers.index(OWN_PRODUCER) if OWN_PRODUCER in producers else -1,
        "slots": [slots.get(name, -1) for name in producers],
        "check": {"rows": int(len(agg)), "units": float(agg["TOTAL_Шт"].sum()), "catRows": int(len(cagg))},
    }
    # Рубли храним с копейками: целые рубли (int32) + копейки (0..99) отдельной колонкой
    kop = np.round(agg["TOTAL_Руб"].to_numpy("f8") * 100)
    rub = np.floor_divide(kop, 100)
    columns = {
        "P": (agg["p"].to_numpy("<u2"), "u16"), "N": (agg["n"].to_numpy("u1"), "u8"),
        "F": (agg["f"].to_numpy("u1"), "u8"), "K": (agg["k"].to_numpy("<u2"), "u16"),
        "U": as_int(agg["TOTAL_Шт"], "TOTAL_Шт", exact=True), "R": as_int(pd.Series(rub), "TOTAL_Руб"), "Rk": ((kop - rub * 100).astype("u1"), "u8"),
        "G": as_int(agg["TOTAL_КгЛ"] * 100, "TOTAL_КгЛ"), "PU": as_int(agg["PROMO_Шт"], "PROMO_Шт", exact=True),
        "S": as_u16(agg["TOTAL_Магазины"], "TOTAL_Магазины"), "DF": as_int(agg["ГлубинаПолн"], "ГлубинаПолн"),
        "DS": as_int(agg["ГлубинаПолн"] - agg["ГлубинаОпл"], "ГлубинаСкидка"),
        "cP": (cagg["p"].to_numpy("<u2"), "u16"), "cN": (cagg["n"].to_numpy("u1"), "u8"),
        "cF": (cagg["f"].to_numpy("u1"), "u8"), "cS": as_int(cagg["TOTAL_Магазины"], "cS", exact=True),
    }
    if len(products) > 65535 or len(cal) > 65535 or len(nets) > 255 or len(fos) > 255:
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
        print(f"{path.name}: строк фактов {fact_rows} -> после свёртки до ФО {meta['check']['rows']}, "
              f"данные {raw / 1e6:.0f} МБ -> файл {size:.1f} МБ")
        if size > 25 and not START_YEAR:
            print("  Файл тяжёлый для почты. Чтобы облегчить, поставь в настройках START_YEAR, например 2024.")
    print(f"Готово. Файлы лежат в {OUT}")


TEMPLATE = r'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>__X5_TITLE__</title>
<style>
:root {
  color-scheme: light;
  --page: #f9f9f7; --surface: #fcfcfb; --ink: #0b0b0b; --ink2: #52514e; --muted: #898781;
  --grid: #e1e0d9; --axis: #c3c2b7; --border: rgba(11,11,11,0.10); --wash: rgba(11,11,11,0.05);
  --good: #006300; --bad: #d03b3b; --other: #898781; --track: #f0efec;
  --s1: #2a78d6; --s2: #eb6834; --s3: #1baf7a; --s4: #eda100; --s5: #e87ba4; --s6: #008300; --s7: #6250d6; --s8: #e34948;
}
@media (prefers-color-scheme: dark) {
  :root:where(:not([data-theme="light"])) {
    color-scheme: dark;
    --page: #0d0d0d; --surface: #1a1a19; --ink: #f0efec; --ink2: #c3c2b7; --muted: #898781;
    --grid: #2c2c2a; --axis: #383835; --border: rgba(255,255,255,0.10); --wash: rgba(255,255,255,0.07);
    --good: #0ca30c; --bad: #e66767; --other: #898781; --track: #262624;
    --s1: #3987e5; --s2: #d95926; --s3: #199e70; --s4: #c98500; --s5: #d55181; --s6: #0ca30c; --s7: #9085e9; --s8: #e66767;
  }
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
.wrap { max-width: 1440px; margin: 0 auto; padding: 20px 24px 40px; }
h1 { font-size: 20px; font-weight: 600; margin: 0; }
.sub { color: var(--ink2); font-size: 13px; margin-top: 2px; }
.filters { display: flex; flex-wrap: wrap; gap: 8px; align-items: center; margin: 16px 0; position: sticky; top: 0; z-index: 30;
  background: var(--page); padding: 8px 0; border-bottom: 1px solid var(--border); }
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
.ms-panel { position: absolute; top: 36px; left: 0; z-index: 40; width: 300px; max-height: 360px; display: flex; flex-direction: column;
  background: var(--surface); border: 1px solid var(--border); border-radius: 10px; box-shadow: 0 8px 24px rgba(0,0,0,.12); }
.ms-panel[hidden] { display: none; }
.ms-search { margin: 8px; height: 30px; padding: 0 8px; border: 1px solid var(--border); border-radius: 6px; background: var(--page); color: var(--ink); font: inherit; }
.ms-list { overflow: auto; padding: 4px 0 8px; }
.ms-item { display: flex; gap: 8px; align-items: center; padding: 4px 12px; cursor: pointer; }
.ms-item:hover { background: var(--wash); }
.ms-item input { margin: 0; accent-color: var(--s1); }
.ms-all { border-bottom: 1px solid var(--border); padding-bottom: 6px; margin-bottom: 4px; font-weight: 500; }
.kpis { display: grid; grid-template-columns: repeat(auto-fit, minmax(190px, 1fr)); gap: 12px; margin-bottom: 16px; }
.tile { background: var(--surface); border: 1px solid var(--border); border-radius: 12px; padding: 14px 16px; }
.tile .label { color: var(--ink2); font-size: 13px; }
.tile .value { font-size: 26px; font-weight: 600; margin-top: 4px; }
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
  const years = [...new Set(per.year)].sort((a, b) => a - b);
  // Свой бренд — больше половины его SKU у своего производителя (в классификаторе бывают единичные чужие SKU)
  const ownCnt = new Int32Array(nBrand), allCnt = new Int32Array(nBrand);
  for (let k = 0; k < nK; k++) { allCnt[brandOf[k]]++; if (prodOf[k] === META.own) ownCnt[brandOf[k]]++; }
  const isOwnBrand = b => META.own >= 0 && ownCnt[b] * 2 > allCnt[b];
  const lastLabel = per.label[nP - 1];
  document.getElementById('subtitle').textContent =
    `${weekly ? 'Недели, месяцы по 4-4-5' : 'Календарные месяцы'} · данные с ${per.label[0]} по ${lastLabel} · собрано ${META.generated}`;

  // ---------- состояние и расчёт ----------
  const state = { year: years[years.length - 1], months: new Set(), nets: new Set(), fos: new Set(), cats: new Set(), prods: new Set(), brands: new Set() };

  function mask(n, set) {
    const m = new Uint8Array(n);
    for (let i = 0; i < n; i++) m[i] = set.size === 0 || set.has(i) ? 1 : 0;
    return m;
  }

  function compute(st) {
    const perSel = new Uint8Array(nP), perPy = new Uint8Array(nP), sel = [];
    for (let p = 0; p < nP; p++) {
      if (per.year[p] === st.year && (st.months.size === 0 || st.months.has(per.month[p]))) { perSel[p] = 1; sel.push(p); }
    }
    for (const p of sel) if (per.pg[p] >= 0) perPy[per.pg[p]] = 1;
    const netM = mask(nN, st.nets), foM = mask(nF, st.fos);
    const catM = new Uint8Array(nK), fullM = new Uint8Array(nK);
    for (let k = 0; k < nK; k++) {
      const c = st.cats.size === 0 || st.cats.has(catOf[k]);
      catM[k] = c ? 1 : 0;
      fullM[k] = c && (st.prods.size === 0 || st.prods.has(prodOf[k])) && (st.brands.size === 0 || st.brands.has(brandOf[k])) ? 1 : 0;
    }
    let selU = 0, selR = 0, selG = 0, selPU = 0, catU = 0, catR = 0, catG = 0, pyU = 0, pyR = 0, pyPU = 0, pyCatU = 0;
    const selU_p = new Float64Array(nP), selS_p = new Float64Array(nP), pyU_p = new Float64Array(nP), pyS_p = new Float64Array(nP);
    const catU_p = new Float64Array(nP), catS_p = new Float64Array(nP);
    const prodU = new Float64Array(nProd), prodPU = new Float64Array(nProd), prodDF = new Float64Array(nProd), prodDS = new Float64Array(nProd);
    const prodU_p = new Float64Array(nP * nProd);
    const brU = new Float64Array(nBrand), brR = new Float64Array(nBrand), brG = new Float64Array(nBrand), brPU = new Float64Array(nBrand);
    const brPy = new Float64Array(nBrand), brS_p = new Float64Array(nP * nBrand);
    const foU = new Float64Array(nF), foPy = new Float64Array(nF), foCat = new Float64Array(nF), foCatPy = new Float64Array(nF);
    const nU = new Float64Array(nN), nPy = new Float64Array(nN), nCat = new Float64Array(nN), nCatPy = new Float64Array(nN);
    const skU = new Float64Array(nK), skR = new Float64Array(nK), skPU = new Float64Array(nK), skPy = new Float64Array(nK);
    const P = D.P, N = D.N, F = D.F, K = D.K, U = D.U, R = D.R, Rk = D.Rk, G = D.G, PU = D.PU, S = D.S, DF = D.DF, DS = D.DS;
    for (let i = 0, n = P.length; i < n; i++) {
      const p = P[i];
      const isSel = perSel[p];
      if (!isSel && !perPy[p]) continue;
      const f = F[i], net = N[i];
      if (!netM[net] || !foM[f]) continue;
      const k = K[i];
      if (!catM[k]) continue;
      const u = U[i];
      if (isSel) {
        const r = R[i] + Rk[i] * 0.01, g = G[i], pu = PU[i], pr = prodOf[k];
        catU += u; catR += r; catG += g; catU_p[p] += u; foCat[f] += u; nCat[net] += u;
        prodU[pr] += u; prodPU[pr] += pu; prodDF[pr] += DF[i]; prodDS[pr] += DS[i]; prodU_p[p * nProd + pr] += u;
        if (fullM[k]) {
          const s = S[i], b = brandOf[k];
          selU += u; selR += r; selG += g; selPU += pu; selU_p[p] += u; selS_p[p] += s;
          brU[b] += u; brR[b] += r; brG[b] += g; brPU[b] += pu; brS_p[p * nBrand + b] += s;
          foU[f] += u; nU[net] += u; skU[k] += u; skR[k] += r; skPU[k] += pu;
        }
      } else {
        pyCatU += u; foCatPy[f] += u; nCatPy[net] += u;
        if (fullM[k]) {
          pyU += u; pyR += R[i] + Rk[i] * 0.01; pyPU += PU[i]; pyU_p[p] += u; pyS_p[p] += S[i];
          brPy[brandOf[k]] += u; foPy[f] += u; nPy[net] += u; skPy[k] += u;
        }
      }
    }
    const cP = D.cP, cN = D.cN, cF = D.cF, cS = D.cS;
    for (let j = 0; j < cP.length; j++) {
      const p = cP[j];
      if ((perSel[p] || perPy[p]) && netM[cN[j]] && foM[cF[j]]) catS_p[p] += cS[j];
    }
    // TDP: средняя по периодам выбора; знаменатель — магазины с продажами всей категории БП
    function tdpAvg(stores, offset, stride) {
      let sum = 0, cnt = 0;
      for (const p of sel) if (catS_p[p] > 0) { sum += stores[p * stride + offset] / catS_p[p]; cnt++; }
      return cnt ? sum / cnt : null;
    }
    let tdpPySum = 0, tdpPyCnt = 0;
    for (const p of sel) {
      const q = per.pg[p];
      if (q >= 0 && catS_p[q] > 0) { tdpPySum += pyS_p[q] / catS_p[q]; tdpPyCnt++; }
    }
    const kpi = {
      units: selU, unitsPy: pyU, rub: selR, rubPy: pyR,
      share: div(selU, catU), sharePy: div(pyU, pyCatU),
      promo: div(selPU, selU), promoPy: div(pyPU, pyU),
      price: div(selR, selU), pricePy: div(pyR, pyU),
      tdp: tdpAvg(selS_p, 0, 1), tdpPy: tdpPyCnt ? tdpPySum / tdpPyCnt : null,
      hasPy: sel.some(p => per.pg[p] >= 0),
    };
    const chart1 = { labels: sel.map(p => per.label[p]), nums: sel.map(p => per.num[p]),
      cur: sel.map(p => selU_p[p]), py: sel.map(p => (per.pg[p] >= 0 ? pyU_p[per.pg[p]] : null)) };
    const prodRank = [...Array(nProd).keys()].filter(i => prodU[i] > 0).sort((a, b) => prodU[b] - prodU[a]);
    const top5 = prodRank.slice(0, 5), top10 = prodRank.slice(0, 10);
    const chart2 = { labels: chart1.labels, series: top5.map(pr => ({ prod: pr,
      values: sel.map(p => (catU_p[p] > 0 ? prodU_p[p * nProd + pr] / catU_p[p] : null)) })) };
    const chart3 = top10.map(pr => ({ prod: pr, share: div(prodU[pr], catU) }));
    const chart4 = top10.map(pr => ({ prod: pr, promo: div(prodPU[pr], prodU[pr]),
      depth: prodDF[pr] > 0 ? prodDS[pr] / prodDF[pr] : null }));
    const brands = [];
    for (let b = 0; b < nBrand; b++) {
      if (!brU[b] && !brPy[b]) continue;
      const share = div(brU[b], catU), sharePy = div(brPy[b], pyCatU);
      brands.push({ brand: b, units: brU[b], unitsPy: brPy[b], growth: growth(brU[b], brPy[b]), share, sharePy,
        dShare: ok(share) && ok(sharePy) && kpi.hasPy ? 100 * (share - sharePy) : null,
        promo: div(brPU[b], brU[b]), price: div(brR[b], brU[b]),
        index: brG[b] > 0 && catG > 0 && catR > 0 ? (brR[b] / brG[b]) / (catR / catG) : null,
        tdp: tdpAvg(brS_p, b, nBrand) });
    }
    const fos = [];
    for (let f = 0; f < nF; f++) {
      if (!foCat[f] && !foCatPy[f]) continue;
      const share = div(foU[f], foCat[f]), sharePy = div(foPy[f], foCatPy[f]);
      fos.push({ fo: f, units: foU[f], unitsPy: foPy[f], growth: growth(foU[f], foPy[f]), share, sharePy,
        dShare: ok(share) && ok(sharePy) && kpi.hasPy ? 100 * (share - sharePy) : null });
    }
    const nets = [];
    for (let n = 0; n < nN; n++) {
      if (!nCat[n] && !nCatPy[n]) continue;
      const share = div(nU[n], nCat[n]), sharePy = div(nPy[n], nCatPy[n]);
      nets.push({ net: n, units: nU[n], unitsPy: nPy[n], growth: growth(nU[n], nPy[n]), share, sharePy,
        dShare: ok(share) && ok(sharePy) && kpi.hasPy ? 100 * (share - sharePy) : null });
    }
    const skus = [];
    for (let k = 0; k < nK; k++) {
      if (!skU[k] && !skPy[k]) continue;
      skus.push({ k, units: skU[k], unitsPy: skPy[k], growth: growth(skU[k], skPy[k]), price: div(skR[k], skU[k]),
        promo: div(skPU[k], skU[k]), share: div(skU[k], catU) });
    }
    const total = { units: selU, unitsPy: pyU, growth: growth(selU, pyU), share: kpi.share, sharePy: kpi.sharePy,
      dShare: ok(kpi.share) && ok(kpi.sharePy) && kpi.hasPy ? 100 * (kpi.share - kpi.sharePy) : null };
    return { sel, kpi, chart1, chart2, chart3, chart4, brands, fos, nets, skus, total };
  }

  // ---------- интерфейс: срезы ----------
  let renderQueued = false;
  function schedule() {
    if (renderQueued) return;
    renderQueued = true;
    requestAnimationFrame(() => { renderQueued = false; renderAll(); });
  }
  const openPanels = new Set();
  document.addEventListener('pointerdown', e => {
    for (const p of [...openPanels]) if (!p.wrap.contains(e.target)) p.close();
  });
  document.addEventListener('keydown', e => { if (e.key === 'Escape') for (const p of [...openPanels]) p.close(); });

  function multiSelect(label, set, getOptions, onChange) {
    const wrap = el('div', { class: 'ms' });
    const val = el('span', { class: 'ms-val' });
    const btn = el('button', { class: 'ms-btn', type: 'button', 'aria-haspopup': 'listbox', 'aria-expanded': 'false' },
      el('span', { class: 'ms-label', text: label }), val, el('span', { 'aria-hidden': 'true', text: '▾' }));
    const panel = el('div', { class: 'ms-panel', hidden: true, role: 'listbox', 'aria-multiselectable': 'true', 'aria-label': label });
    const search = el('input', { class: 'ms-search', type: 'search', placeholder: 'Поиск', 'aria-label': `Поиск: ${label}` });
    const list = el('div', { class: 'ms-list' });
    panel.append(search, list);
    wrap.append(btn, panel);
    const api = { wrap, refresh, close };
    function textOf(v) { const o = getOptions().find(x => x.value === v); return o ? o.text : String(v); }
    function refresh() {
      val.textContent = set.size === 0 ? 'Все' : set.size === 1 ? textOf([...set][0]) : `выбрано ${set.size}`;
      btn.dataset.active = set.size > 0 ? 'true' : 'false';
    }
    function fill() {
      const q = search.value.trim().toLowerCase();
      list.textContent = '';
      const allBox = el('input', { type: 'checkbox' });
      allBox.checked = set.size === 0;
      allBox.addEventListener('change', () => { set.clear(); fill(); refresh(); onChange(); });
      list.append(el('label', { class: 'ms-item ms-all' }, allBox, el('span', { text: 'Все' })));
      for (const o of getOptions()) {
        if (q && !o.text.toLowerCase().includes(q)) continue;
        const box = el('input', { type: 'checkbox' });
        box.checked = set.has(o.value);
        box.addEventListener('change', () => {
          if (box.checked) set.add(o.value); else set.delete(o.value);
          allBox.checked = set.size === 0;
          refresh(); onChange();
        });
        list.append(el('label', { class: 'ms-item' }, box, el('span', { text: o.text })));
      }
    }
    function open() {
      for (const p of [...openPanels]) p.close();
      search.value = '';
      fill();
      panel.hidden = false; btn.setAttribute('aria-expanded', 'true'); openPanels.add(api);
      search.focus();
    }
    function close() { panel.hidden = true; btn.setAttribute('aria-expanded', 'false'); openPanels.delete(api); }
    btn.addEventListener('click', () => (panel.hidden ? open() : close()));
    search.addEventListener('input', fill);
    refresh();
    return api;
  }

  function productOk(k, useProds) {
    if (state.cats.size && !state.cats.has(catOf[k])) return false;
    if (useProds && state.prods.size && !state.prods.has(prodOf[k])) return false;
    return true;
  }
  function availableProducers() {
    const seen = new Uint8Array(nProd);
    for (let k = 0; k < nK; k++) if (productOk(k, false)) seen[prodOf[k]] = 1;
    return META.producers.map((t, i) => ({ value: i, text: t })).filter(o => seen[o.value]);
  }
  function availableBrands() {
    const seen = new Uint8Array(nBrand);
    for (let k = 0; k < nK; k++) if (productOk(k, true)) seen[brandOf[k]] = 1;
    return META.brands.map((t, i) => ({ value: i, text: t })).filter(o => seen[o.value]);
  }
  function prune(set, options) {
    const allowed = new Set(options.map(o => o.value));
    for (const v of [...set]) if (!allowed.has(v)) set.delete(v);
  }

  const controls = {};
  function buildFilters() {
    const bar = el('div', { class: 'filters', role: 'region', 'aria-label': 'Срезы' });
    const seg = el('div', { class: 'seg', role: 'group', 'aria-label': 'Год' });
    for (const y of years) {
      const b = el('button', { type: 'button', text: String(y), 'aria-pressed': String(y === state.year) });
      b.addEventListener('click', () => {
        state.year = y;
        for (const x of seg.children) x.setAttribute('aria-pressed', String(x === b));
        schedule();
      });
      seg.append(b);
    }
    controls.years = seg;
    const monthOpts = MONTHS.map((t, i) => ({ value: i + 1, text: t }));
    controls.months = multiSelect(weekly ? 'Месяц 4-4-5' : 'Месяц', state.months, () => monthOpts, schedule);
    controls.nets = multiSelect('Сеть', state.nets, () => META.nets.map((t, i) => ({ value: i, text: t })), schedule);
    controls.fos = multiSelect('ФО', state.fos, () => META.fos.map((t, i) => ({ value: i, text: t })), schedule);
    controls.cats = multiSelect('Категория', state.cats, () => META.categories.map((t, i) => ({ value: i, text: t })), () => {
      prune(state.prods, availableProducers()); prune(state.brands, availableBrands());
      controls.prods.refresh(); controls.brands.refresh(); schedule();
    });
    controls.prods = multiSelect('Производитель', state.prods, availableProducers, () => {
      prune(state.brands, availableBrands()); controls.brands.refresh(); schedule();
    });
    controls.brands = multiSelect('Бренд', state.brands, availableBrands, schedule);
    const reset = el('button', { class: 'btn', type: 'button', text: 'Сбросить' });
    reset.addEventListener('click', () => {
      state.year = years[years.length - 1];
      for (const k of ['months', 'nets', 'fos', 'cats', 'prods', 'brands']) { state[k].clear(); controls[k].refresh(); }
      for (const x of seg.children) x.setAttribute('aria-pressed', String(Number(x.textContent) === state.year));
      schedule();
    });
    bar.append(seg, controls.months.wrap, controls.nets.wrap, controls.fos.wrap, controls.cats.wrap,
      controls.prods.wrap, controls.brands.wrap, reset);
    return bar;
  }

  // ---------- показатели ----------
  const tiles = {};
  function buildKpis() {
    const box = el('section', { class: 'kpis', 'aria-label': 'Ключевые показатели' });
    const defs = [
      ['units', 'Продажи, шт'], ['rub', 'Продажи, руб'], ['share', 'Доля в категории, шт'],
      ['promo', 'Доля промо, шт'], ['price', 'Цена за шт'], ['tdp', 'TDP средняя'],
    ];
    for (const [key, label] of defs) {
      const value = el('div', { class: 'value', text: '—' }), delta = el('div', { class: 'delta' });
      tiles[key] = { value, delta };
      box.append(el('div', { class: 'tile' }, el('div', { class: 'label', text: label }), value, delta));
    }
    return box;
  }
  function setDelta(node, d, tone) {
    node.textContent = '';
    if (!d) { node.className = 'delta none'; node.textContent = 'к ПГ: нет данных'; return; }
    node.className = 'delta ' + (d.dir === 0 ? 'neutral' : tone(d.dir > 0));
    if (d.dir) node.append(el('span', { 'aria-hidden': 'true', text: d.dir > 0 ? '▲ ' : '▼ ' }));
    node.append(document.createTextNode(d.text + ' к ПГ'));
  }
  const goodUp = up => (up ? 'good' : 'bad');
  const neutralTone = () => 'neutral';
  function renderKpis(r) {
    const k = r.kpi;
    tiles.units.value.textContent = fmt.compact(k.units);
    setDelta(tiles.units.delta, k.hasPy ? delta(growth(k.units, k.unitsPy), 'pct') : null, goodUp);
    tiles.rub.value.textContent = fmt.compact(k.rub);
    setDelta(tiles.rub.delta, k.hasPy ? delta(growth(k.rub, k.rubPy), 'pct') : null, goodUp);
    tiles.share.value.textContent = fmt.pct(k.share);
    setDelta(tiles.share.delta, k.hasPy && ok(k.share) && ok(k.sharePy) ? delta(100 * (k.share - k.sharePy), 'pp') : null, goodUp);
    tiles.promo.value.textContent = fmt.pct(k.promo);
    setDelta(tiles.promo.delta, k.hasPy && ok(k.promo) && ok(k.promoPy) ? delta(100 * (k.promo - k.promoPy), 'pp') : null, neutralTone);
    tiles.price.value.textContent = fmt.rub(k.price);
    setDelta(tiles.price.delta, k.hasPy && ok(k.price) && ok(k.pricePy) ? delta(k.price / k.pricePy - 1, 'pct') : null, neutralTone);
    tiles.tdp.value.textContent = fmt.tdp(k.tdp);
    setDelta(tiles.tdp.delta, k.hasPy && ok(k.tdp) && ok(k.tdpPy) ? delta(100 * (k.tdp - k.tdpPy), 'pt') : null, goodUp);
  }

  // ---------- карточки и табличный вид ----------
  function card(title, note) {
    const body = el('div', { class: 'chart' });
    const table = el('div', { class: 'table-wrap', hidden: true });
    const tip = el('div', { class: 'tip', hidden: true, role: 'status' });
    const legend = el('div', { class: 'legend' });
    const toggle = el('button', { class: 'ghost', type: 'button', text: 'Таблица', 'aria-pressed': 'false' });
    const node = el('section', { class: 'card' },
      el('div', { class: 'card-head' }, el('div', {}, el('h2', { text: title }), note ? el('div', { class: 'note', text: note }) : null), toggle),
      body, legend, table, tip);
    const c = { node, body, table, tip, legend, toggle, rows: null, head: null, draw: null };
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
    const W = Math.max(320, host.clientWidth || 600), H = 250, m = { l: 58, r: 84, t: 12, b: 28 };
    const iw = W - m.l - m.r, ih = H - m.t - m.b;
    let lo = 0, hi = 0;
    for (const s of cfg.series) for (const v of s.values) if (ok(v)) { hi = Math.max(hi, v); lo = Math.min(lo, v); }
    if (hi === lo) hi = lo + 1;
    const ticks = niceTicks(lo, hi, 4), y0 = ticks[0], y1 = ticks[ticks.length - 1], step = ticks[1] - ticks[0];
    const X = i => m.l + (n === 1 ? iw / 2 : (i * iw) / (n - 1));
    const Y = v => m.t + ih - ((v - y0) / (y1 - y0)) * ih;
    const root = svg('svg', { width: W, height: H, viewBox: `0 0 ${W} ${H}`, role: 'img', 'aria-label': cfg.aria });
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
  function dataTable(title, note, columns, opts) {
    const wrapT = el('div', { class: 'table-wrap' });
    const more = el('button', { class: 'ghost more', type: 'button', hidden: true });
    const node = el('section', { class: 'card' + (opts.full ? ' full' : '') },
      el('div', { class: 'card-head' }, el('div', {}, el('h2', { text: title }), note ? el('div', { class: 'note', text: note }) : null)), wrapT, more);
    const t = { node, rows: [], total: null, sortKey: opts.sortKey, dir: -1, expanded: false };
    more.addEventListener('click', () => { t.expanded = !t.expanded; draw(); });
    function cellOf(col, row) {
      const v = row[col.key];
      if (col.delta) {
        const d = delta(v, col.delta);
        if (!d) return el('td', { class: 'num none', text: '—' });
        if (!d.dir) return el('td', { class: 'num neutral', text: d.text });
        return el('td', { class: 'num ' + (d.dir > 0 ? 'good' : 'bad') }, el('span', { 'aria-hidden': 'true', text: d.dir > 0 ? '▲ ' : '▼ ' }), document.createTextNode(d.text));
      }
      if (col.render) return col.render(row);
      return el('td', { class: col.num ? 'num' : 'name', text: col.fmt ? col.fmt(v) : String(v) });
    }
    function draw() {
      wrapT.textContent = '';
      if (!t.rows.length) { wrapT.append(el('div', { class: 'empty', text: 'Нет данных для выбранных срезов' })); more.hidden = true; return; }
      const col = columns.find(x => x.key === t.sortKey);
      const val = r => { const v = col.sort ? col.sort(r) : r[col.key]; return ok(v) ? v : (typeof v === 'string' ? v : -Infinity); };
      const rows = [...t.rows].sort((a, b) => {
        const x = val(a), y = val(b);
        return (typeof x === 'string' ? x.localeCompare(y, 'ru') : x - y) * t.dir;
      });
      const shown = t.expanded || !opts.limit ? rows : rows.slice(0, opts.limit);
      const table = el('table', { class: 'data' });
      const head = el('tr');
      for (const cdef of columns) {
        const th = el('th', { class: cdef.num ? 'num' : '', text: cdef.title, scope: 'col', tabindex: 0 });
        if (cdef.key === t.sortKey) th.setAttribute('aria-sort', t.dir > 0 ? 'ascending' : 'descending');
        const sortBy = () => { if (t.sortKey === cdef.key) t.dir = -t.dir; else { t.sortKey = cdef.key; t.dir = cdef.num ? -1 : 1; } draw(); };
        th.addEventListener('click', sortBy);
        th.addEventListener('keydown', e => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); sortBy(); } });
        head.append(th);
      }
      table.append(el('thead', {}, head));
      const tb = el('tbody');
      for (const r of shown) tb.append(el('tr', {}, ...columns.map(cdef => cellOf(cdef, r))));
      if (t.total) tb.append(el('tr', { class: 'total' }, ...columns.map(cdef => cellOf(cdef, t.total))));
      table.append(tb);
      wrapT.append(table);
      if (opts.limit && rows.length > opts.limit) {
        more.hidden = false;
        more.textContent = t.expanded ? 'Свернуть' : `Показать все (${rows.length})`;
      } else more.hidden = true;
    }
    t.set = (rows, total) => { t.rows = rows; t.total = total || null; draw(); };
    return t;
  }

  // ---------- сборка страницы ----------
  app.textContent = '';
  app.append(buildFilters(), buildKpis());
  const periodWord = weekly ? 'неделям' : 'месяцам';
  const c1 = card(`Продажи, шт по ${periodWord}`, 'Выбранный год и тот же период прошлого года');
  const c2 = card(`Доля в категории по ${periodWord}: топ-5 производителей`, 'Доля в штуках · не зависит от срезов производителя и бренда');
  const c3 = card('Доля производителей в категории, шт', 'Топ-10 за выбранный период · не зависит от срезов производителя и бренда');
  const c4 = card('Промо по производителям', 'Доля промо в продажах, шт и взвешенная глубина промо · топ-10 · без срезов производителя и бренда');
  app.append(el('div', { class: 'grid2' }, c1.node, c2.node), el('div', { class: 'grid2' }, c3.node, c4.node));
  const name = k => META.products.name[k];
  const ownMark = (text, own) => (own ? el('span', { class: 'own', text }) : document.createTextNode(text));
  const tBrands = dataTable('Бренды', 'Доля и индекс цены — от категорий из среза «Категория» в тех же сетях, ФО и периоде', [
    { key: 'name', title: 'Бренд', render: r => el('td', { class: 'name', title: r.name }, ownMark(r.name, r.own)) },
    { key: 'units', title: 'Продажи, шт', num: true, fmt: fmt.int },
    { key: 'growth', title: 'Прирост к ПГ', num: true, delta: 'pct' },
    { key: 'share', title: 'Доля в категории', num: true, fmt: fmt.pct },
    { key: 'dShare', title: 'Изменение доли', num: true, delta: 'pp' },
    { key: 'promo', title: 'Доля промо', num: true, fmt: fmt.pct },
    { key: 'price', title: 'Цена за шт', num: true, fmt: fmt.rub },
    { key: 'index', title: 'Индекс цены за кг', num: true, fmt: fmt.idx },
    { key: 'tdp', title: 'TDP средняя', num: true, fmt: fmt.tdp },
  ], { sortKey: 'units', limit: 20, full: true });
  app.append(tBrands.node);
  const tFo = dataTable('Федеральные округа', 'Доля — в категории своего округа', [
    { key: 'name', title: 'ФО' },
    { key: 'units', title: 'Продажи, шт', num: true, fmt: fmt.int },
    { key: 'growth', title: 'Прирост к ПГ', num: true, delta: 'pct' },
    { key: 'share', title: 'Доля в категории', num: true, fmt: fmt.pct },
    { key: 'dShare', title: 'Изменение доли', num: true, delta: 'pp' },
  ], { sortKey: 'units' });
  const tNets = dataTable('Сети', 'Доля — в категории своей сети', [
    { key: 'name', title: 'Сеть' },
    { key: 'units', title: 'Продажи, шт', num: true, fmt: fmt.int },
    { key: 'growth', title: 'Прирост к ПГ', num: true, delta: 'pct' },
    { key: 'share', title: 'Доля в категории', num: true, fmt: fmt.pct },
    { key: 'dShare', title: 'Изменение доли', num: true, delta: 'pp' },
  ], { sortKey: 'units' });
  const tSku = dataTable('SKU', 'Доля — в категориях из среза «Категория» в тех же сетях, ФО и периоде', [
    { key: 'name', title: 'SKU', render: r => el('td', { class: 'name', title: `${r.name} · PLU ${r.plu}` }, document.createTextNode(r.name), el('span', { class: 'plu', text: r.plu })) },
    { key: 'brandName', title: 'Бренд' },
    { key: 'units', title: 'Продажи, шт', num: true, fmt: fmt.int },
    { key: 'growth', title: 'Прирост к ПГ', num: true, delta: 'pct' },
    { key: 'price', title: 'Цена за шт', num: true, fmt: fmt.rub },
    { key: 'promo', title: 'Доля промо', num: true, fmt: fmt.pct },
    { key: 'share', title: 'Доля в категории', num: true, fmt: fmt.pct },
  ], { sortKey: 'units', limit: 20, full: true });
  app.append(el('div', { class: 'grid2' }, tFo.node, tNets.node), tSku.node);
  app.append(el('footer', {},
    el('p', { text: 'Источник: X5 Dialog, отчёт «Анализ трендов», выгрузка из ClickHouse. Регионы сведены в федеральные округа.' }),
    el('p', { text: 'Доля в категории — продажи выбранного в штуках к продажам категорий из среза «Категория» в тех же сетях, ФО и периоде; срезы производителя и бренда на знаменатель не влияют. В таблицах ФО и сетей доля считается внутри своего округа или своей сети. Изменение доли — в процентных пунктах к тому же периоду прошлого года.' }),
    el('p', { text: weekly
      ? 'ПГ — те же недели ISO прошлого года (у 53-й недели пары нет). Месяцы недельных данных — по схеме 4-4-5.'
      : 'ПГ — те же месяцы прошлого года.' }),
    el('p', { text: 'Взвешенная глубина промо — на сколько промо-цена ниже обычной цены тех же SKU в той же сети, регионе и периоде, взвешено по промо-штукам. Строки, где SKU продавался только по промо, пропускаются.' }),
    el('p', { text: 'TDP — сумма нумерических дистрибуций SKU в пунктах: для каждого SKU берётся доля магазинов с его продажами среди магазинов с продажами всей категории БП, в процентах, и доли складываются. TDP 250 — в среднем 2,5 SKU в магазине. TDP средняя — среднее по периодам выбора. Магазин засчитывается, только если в нём была продажа. В Excel та же мера показана в процентах: 250% там — это 250 здесь.' }),
    el('p', { text: 'Индекс цены за кг — цена за кг бренда к цене за кг категории; 1,00 — на уровне категории. Чеки, покупатели и взвешенная дистрибуция здесь не показаны: для брендов из данных X5 их честно не посчитать.' }),
    el('p', { id: 'perf', class: 'sub' })));

  let last = null;
  function drawCharts(r) {
    const y = state.year;
    const xLabel = i => (weekly ? 'Н' + r.chart1.nums[i] : MON3[r.chart1.nums[i] - 1]);
    const tipTitle = i => (weekly ? `Неделя ${r.chart1.nums[i]} (${r.chart1.labels[i]})` : MONTHS[r.chart1.nums[i] - 1] + ' ' + y);
    c1.rows = r.chart1.labels.map((l, i) => [l, fmt.int(r.chart1.cur[i]), fmt.int(r.chart1.py[i])]);
    c1.head = ['Период', `${y}, шт`, `${y - 1}, шт`];
    lineChart(c1, { labels: r.chart1.labels, xLabel, tipTitle, fmt: fmt.compact, axisFmt: v => fmt.compact(v), sortTip: false, aria: `Продажи в штуках по ${periodWord}, ${y} и ${y - 1}`,
      series: [{ name: String(y), color: 'var(--s1)', values: r.chart1.cur }, { name: `${y - 1} (ПГ)`, color: 'var(--other)', values: r.chart1.py }] });
    c2.rows = r.chart1.labels.map((l, i) => [l, ...r.chart2.series.map(s => fmt.pct(s.values[i]))]);
    c2.head = ['Период', ...r.chart2.series.map(s => META.producers[s.prod])];
    const pctAxis = (v, step) => { const d = step * 100 >= 0.999 ? 0 : step * 100 >= 0.0999 ? 1 : 2;
      return new Intl.NumberFormat('ru-RU', { minimumFractionDigits: d, maximumFractionDigits: d }).format(v * 100) + '%'; };
    lineChart(c2, { labels: r.chart1.labels, xLabel, tipTitle, fmt: fmt.pct, axisFmt: pctAxis, sortTip: true, aria: 'Доля производителей в категории по периодам',
      series: r.chart2.series.map(s => ({ name: META.producers[s.prod], color: prodColor(s.prod), values: s.values })) });
    c3.rows = r.chart3.map(x => [META.producers[x.prod], fmt.pct(x.share)]);
    c3.head = ['Производитель', 'Доля в категории, шт'];
    barChart(c3, { fmt: fmt.pct, series: [{ name: 'Доля в категории, шт', color: 'var(--other)' }],
      rows: r.chart3.map(x => ({ label: META.producers[x.prod], values: [x.share], own: x.prod === META.own,
        colors: [x.prod === META.own ? 'var(--s1)' : 'var(--other)'] })),
      legend: META.own >= 0 ? [{ name: META.producers[META.own], color: 'var(--s1)' }, { name: 'Другие производители', color: 'var(--other)' }] : null });
    c4.rows = r.chart4.map(x => [META.producers[x.prod], fmt.pct(x.promo), fmt.pct(x.depth)]);
    c4.head = ['Производитель', 'Доля промо, шт', 'Глубина промо'];
    barChart(c4, { fmt: fmt.pct, series: [{ name: 'Доля промо, шт', color: 'var(--s1)' }, { name: 'Глубина промо, взвешенная', color: 'var(--s2)' }],
      rows: r.chart4.map(x => ({ label: META.producers[x.prod], values: [x.promo, x.depth], own: x.prod === META.own })) });
    for (const c of [c1, c2, c3, c4]) if (!c.table.hidden) fillTable(c);
  }
  function renderAll() {
    const t0 = performance.now();
    const r = compute(state);
    last = r;
    renderKpis(r);
    drawCharts(r);
    tBrands.set(r.brands.map(b => ({ ...b, name: META.brands[b.brand], own: isOwnBrand(b.brand) })));
    tFo.set(r.fos.map(f => ({ ...f, name: META.fos[f.fo] })), { ...r.total, name: 'Итого' });
    tNets.set(r.nets.map(x => ({ ...x, name: META.nets[x.net] })), { ...r.total, name: 'Итого' });
    tSku.set(r.skus.map(s => ({ ...s, name: name(s.k), plu: META.products.plu[s.k], brandName: META.brands[brandOf[s.k]] })));
    document.getElementById('perf').textContent = `Строк данных: ${nf0.format(D.P.length)} · пересчёт ${Math.round(performance.now() - t0)} мс`;
    document.body.dataset.ready = '1';
  }
  let resizeTimer = 0;
  new ResizeObserver(() => { clearTimeout(resizeTimer); resizeTimer = setTimeout(() => last && drawCharts(last), 120); }).observe(app);
  window.X5 = {
    meta: META, state, compute,
    set(patch) {
      for (const [k, v] of Object.entries(patch)) state[k] = v instanceof Array ? new Set(v) : v;
      renderAll();
      return last;
    },
  };
  renderAll();
})();
</script>
</body>
</html>
'''

if __name__ == "__main__":
    main()
