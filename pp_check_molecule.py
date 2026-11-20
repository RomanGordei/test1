"""
pp_check_molecule.py — наши значения для сверки с Молекулой по базовым фактам (без словаря).

Таблица А — одна неделя: все факты должны совпасть с Молекулой один в один.
Таблица Б — периоды: кг и рубли должны совпасть с «наше»; по скидке, длине,
интенсивности и ND смотрим, с чем совпала Молекула — с «наше» (взвешенное, как
в справке Nielsen) или со «средним недель» (простое среднее).
Скидка и интенсивность — со знаком Nielsen (минус), как в Молекуле.
Читает out/pp_week.csv и out/pp_period.csv, ничего не пишет.

Запуск: python pp_check_molecule.py   (рядом с pp_extract_pd.py, после выгрузки)
"""

from pathlib import Path

import pandas as pd

OUT = Path(__file__).resolve().parent / "out"
pd.set_option("display.width", 200)

R_U = "Total Mixed chains Rural + Urban"
WEEK = (R_U, "W 2026 39")
PERIODS = [(R_U, "MAT W39 2026"), ("Discounters", "SEP 26")]

# Факт в Молекуле -> из каких наших колонок он получается
FACTS = {
    "1:Volume (in 1000)":                       lambda d: d["Объём, т"],
    "3:Volume (in 1000)":                       lambda d: d["Объём, т"] * d["Доля промо в кг, %"] / 100,
    "1:Value (in 1000 RUR)":                    lambda d: d["Продажи, тыс. руб."],
    "3:Share of Promo Volume in Total Volume":  lambda d: d["Доля промо в кг, %"],
    "Discount % - Volume weighted":             lambda d: -d["Скидка TPR, %"],
    "Length of TPR (weeks) - Volume weighted":  lambda d: d["Длина TPR, нед."],
    "Intensity of TPR-% of Price paid for TPR": lambda d: -d["Интенсивность TPR, %"],
    "1:Numeric Distribution (w)":               lambda d: d["ND, %"],
    "3:Numeric Distribution (w)":               lambda d: d["ND промо, %"],
    "1:Sellers Universe":                       lambda d: d["Магазинов с продажами"],
    "1:Number of SKUs":                         lambda d: d["SKU в продаже, шт."],
}
PERIOD_FACTS = ["1:Volume (in 1000)", "3:Volume (in 1000)", "1:Value (in 1000 RUR)",
                "Discount % - Volume weighted", "Length of TPR (weeks) - Volume weighted",
                "Intensity of TPR-% of Price paid for TPR", "3:Numeric Distribution (w)"]


def read(name):
    d = pd.read_csv(OUT / name, sep=";", decimal=",", encoding="utf-8-sig")
    d = d[(d["Уровень"] == "Тотал")]
    for c in ("Неделя с", "Неделя по"):
        d[c] = pd.to_datetime(d[c], format="%d.%m.%Y")
    return d


w, p = read("pp_week.csv"), read("pp_period.csv")

mkt, per = WEEK
r = w[(w["Рынок"] == mkt) & (w["Период"] == per)]
print(f"А. TOTAL NOODLES, {mkt}, неделя {per}"
      + (f" ({r['Неделя с'].iloc[0]:%d.%m}–{r['Неделя по'].iloc[0]:%d.%m.%Y})" if len(r) else ""))
if len(r) == 1:
    print(pd.DataFrame({"наше": {f: round(float(fn(r).iloc[0]), 4) for f, fn in FACTS.items()}}).to_string())
else:
    print(f"!!! строк {len(r)}, а нужна одна: проверь имя рынка и недели в WEEK")

for mkt, per in PERIODS:
    r = p[(p["Рынок"] == mkt) & (p["Период"] == per)]
    if len(r) != 1:
        print(f"\n!!! {mkt}, {per}: строк {len(r)}, а нужна одна: проверь PERIODS")
        continue
    d0, d1 = r["Неделя с"].iloc[0], r["Неделя по"].iloc[0]
    weeks = w[(w["Рынок"] == mkt) & (w["Неделя по"] >= d0) & (w["Неделя по"] <= d1)]
    t = pd.DataFrame({f: {"наше": FACTS[f](r).iloc[0], "среднее недель": FACTS[f](weeks).mean(),
                          "сумма недель": FACTS[f](weeks).sum()} for f in PERIOD_FACTS}).T
    print(f"\nБ. TOTAL NOODLES, {mkt}, {per} ({d0:%d.%m.%Y}–{d1:%d.%m.%Y}, недель {len(weeks)})")
    print(t.round(2).to_string())
