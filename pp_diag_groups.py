"""
pp_diag_groups.py — разовая проверка: почему в fact_data не сходятся группы
(A = B + C, C = D + E). Печатает в консоль, насколько велики расхождения,
где они живут и несколько примеров. Файлов не пишет.

Считает в сырых целых, как в файле: 1 = последний знак (0,0001 тыс.).
Расхождение в 1 — округление Nielsen, его не считаем.

Запуск: python pp_diag_groups.py   (кладётся рядом с pp_extract_pd.py, та же папка data)
"""

import time
from pathlib import Path

import pandas as pd

BASE     = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"
PREFIX   = "MRVNODPP"
CHUNK    = 250_000

TAGS  = ["MKT_TAG", "PROD_TAG", "PER_TAG"]
COLS  = [f"{m}_{g}" for m in ("VOL", "VAL", "SALES") for g in "ABCDE"]
PAIRS = [(m, a, b, c) for m in ("VOL", "VAL", "SALES") for a, b, c in (("A", "B", "C"), ("C", "D", "E"))]
pd.set_option("display.width", 250)
pd.set_option("display.max_columns", 30)
t0 = time.time()


def find(name):
    return next(DATA_DIR.glob(f"{PREFIX}_{name}*.csv"))


def read_dict(name):
    return pd.read_csv(find(name), sep="|", dtype=str, keep_default_na=False, encoding="utf-8-sig")


mkt, prod = read_dict("MKT"), read_dict("PROD")
blank, space, parts, rows = dict.fromkeys(COLS, 0), dict.fromkeys(COLS, 0), [], 0

# Всё текстом и без подмены пропусков: нужно отличить пустую ячейку от пробела.
with pd.read_csv(find("FACT_DATA"), sep="|", usecols=TAGS + COLS, dtype=str, keep_default_na=False,
                 chunksize=CHUNK, encoding="utf-8-sig") as reader:
    for c in reader:
        rows += len(c)
        num = {}
        for col in COLS:
            s = c[col]
            e = s == ""
            blank[col] += int(e.sum())
            space[col] += int((~e & (s.str.strip() == "")).sum())
            num[col] = pd.to_numeric(s, errors="coerce").fillna(0)
        for m, a, b, d in PAIRS:
            A, B, C = num[f"{m}_{a}"], num[f"{m}_{b}"], num[f"{m}_{d}"]
            diff = A - B - C
            bad = diff.abs() > 1
            if bad.any():
                v = c.loc[bad, TAGS].copy()
                v["pair"] = f"{m} {a}={b}+{d}"
                v["A"], v["B"], v["C"], v["diff"] = A[bad], B[bad], C[bad], diff[bad]
                v["B_txt"], v["C_txt"] = c.loc[bad, f"{m}_{b}"], c.loc[bad, f"{m}_{d}"]
                parts.append(v)
        print(f"[{time.time() - t0:5.0f}s] прочитано {rows:,} строк", flush=True)

print("\n1. Пропуски по колонкам: пустая ячейка и пробел (строк)")
print(pd.DataFrame({"пусто": blank, "пробел": space}).T.to_string())

if not parts:
    print("\nРасхождений больше 1 в последнем знаке нет.")
    raise SystemExit

v = pd.concat(parts)
v["rel"] = v["diff"].abs() / v["A"].abs().where(v["A"] != 0)
v["pass_op"] = (v["B_txt"].str.strip() == "") | (v["C_txt"].str.strip() == "")
v = (v.merge(mkt[["TAG", "SHORT"]].rename(columns={"TAG": "MKT_TAG", "SHORT": "Рынок"}), on="MKT_TAG", how="left")
      .merge(prod[["TAG", "SHORT", "HIER_NUM", "HIER_LEVEL_NAME"]]
             .rename(columns={"TAG": "PROD_TAG", "SHORT": "Узел"}), on="PROD_TAG", how="left"))
v["Уровень"] = v["HIER_NUM"] + " " + v["HIER_LEVEL_NAME"]

g = v.groupby("pair")
q = lambda s, p: s.quantile(p)
summary = pd.DataFrame({
    "строк": g.size(),
    "% строк": (100 * g.size() / rows).round(3),
    "A>B+C": g["diff"].apply(lambda x: int((x > 0).sum())),
    "A<B+C": g["diff"].apply(lambda x: int((x < 0).sum())),
    "B/C пусты": g["pass_op"].sum(),
    "|р| медиана": g["diff"].apply(lambda x: x.abs().median()),
    "|р| 90%": g["diff"].apply(lambda x: q(x.abs(), .9)),
    "|р| макс": g["diff"].apply(lambda x: x.abs().max()),
    "отн. ≤1e-6": g["rel"].apply(lambda x: int((x <= 1e-6).sum())),
    "≤1e-4": g["rel"].apply(lambda x: int(((x > 1e-6) & (x <= 1e-4)).sum())),
    "≤1e-2": g["rel"].apply(lambda x: int(((x > 1e-4) & (x <= 1e-2)).sum())),
    ">1e-2": g["rel"].apply(lambda x: int((x > 1e-2).sum())),
})
print("\n2. Расхождения больше 1 в последнем знаке. |р| — модуль разности в сырых единицах;\n"
      "   отн. — |A − B − C| / A; «B/C пусты» — строки, где одно из слагаемых пустое или пробел")
print(summary.to_string())

print("\n3. По уровням PROD (иерархия и уровень)")
print(pd.crosstab(v["Уровень"], v["pair"]).to_string())
print("\n4. По рынкам")
print(pd.crosstab(v["Рынок"], v["pair"]).to_string())
print("\n5. По годам")
print(pd.crosstab(v["PER_TAG"].str[1:5], v["pair"]).to_string())

print("\n6. Примеры: самые большие относительные расхождения (значения в настоящих единицах)")
ex = v.sort_values("rel", ascending=False).groupby("pair").head(2)
ex = ex.assign(**{k: ex[k] / 10_000 for k in ("A", "B", "C", "diff")})
print(ex[["pair", "Рынок", "Уровень", "Узел", "PER_TAG", "A", "B", "C", "diff", "rel", "B_txt", "C_txt"]]
      .to_string(index=False))
print(f"\n[{time.time() - t0:5.0f}s] готово")
