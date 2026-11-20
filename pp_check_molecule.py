"""
pp_check_molecule.py — наши значения для сверки с Молекулой: три среза, рядом — имя факта
в Молекуле. Читает out/pp_period.csv, ничего не пишет.
Запуск: python pp_check_molecule.py   (рядом с pp_extract_pd.py, после выгрузки)
"""

from pathlib import Path

import pandas as pd

OUT = Path(__file__).resolve().parent / "out"
pd.set_option("display.width", 250)

# Срез: рынок, уровень, начало имени узла, период
CASES = {
    "Тотал R+U MAT":      ("Total Mixed chains Rural + Urban", "Тотал", "TOTAL NOODLES", "MAT W39 2026"),
    "BigBon R+U MAT":     ("Total Mixed chains Rural + Urban", "Бренд", "BIGBON", "MAT W39 2026"),
    "Тотал Disc. SEP 26": ("Discounters", "Тотал", "TOTAL NOODLES", "SEP 26"),
}
# Наша колонка -> факт в Молекуле (имена из справки Nielsen). * — сверить обязательно.
FACTS = {
    "Неделя с": "(начало периода)",
    "Неделя по": "(конец периода)",
    "Недель": "(недель в периоде)",
    "Объём, т": "1:Volume (in 1000)",
    "Продажи, тыс. руб.": "1:Value (in 1000 RUR)",
    "Доля промо в кг, %": "* Promo Volume Share,%",
    "Доля промо в руб., %": "* Promo Value Share,%",
    "Доля TPR в кг, %": "TPR Volume Share,%",
    "Скидка TPR, %": "* Weighted TPR discount, % (у Nielsen с минусом)",
    "Длина TPR, нед.": "* Length of TPRs, weeks",
    "Интенсивность TPR, %": "* Intensity of TPR-% AGG (с минусом)",
    "Частота промо, нед. на магазин с продажами": "* Promo Frequency — вариант 1",
    "Частота промо, нед. на магазин канала": "* Promo Frequency — вариант 2",
    "ND, %": "Total Numeric Distribution,%",
    "ND промо, %": "* Promo Numeric Distribution,%",
    "SKU в продаже, шт.": "Number of Total Unique SKUs (average)",
    "Магазинов с продажами": "1: Average Selling Universe",
    "Цена рег., руб./уп.": "Average Regular Price per Item",
    "Рег. цена на полке, руб./уп.": "Average Regular Price per Unit in all Stores",
}

p = pd.read_csv(OUT / "pp_period.csv", sep=";", decimal=",", encoding="utf-8-sig")
t = pd.DataFrame(index=list(FACTS))
for name, (mkt, lvl, node, per) in CASES.items():
    r = p[(p["Рынок"] == mkt) & (p["Уровень"] == lvl) & p["Узел"].str.startswith(node) & (p["Период"] == per)]
    if len(r) != 1:
        print(f"!!! {name}: нашлось строк {len(r)}, а нужна одна. Проверь имена в CASES.")
        continue
    t[name] = [v if isinstance(v, str) else int(v) if float(v).is_integer() else round(float(v), 2)
               for v in r.iloc[0][list(FACTS)]]
t["Факт в Молекуле"] = list(FACTS.values())
print(t.to_string())
