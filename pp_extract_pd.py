"""
pp_extract_pd.py — выгрузка Nielsen Promo Pressure по лапше на чистом pandas (без DuckDB).

Делает то же, что pp_extract.py, и пишет те же три CSV с теми же колонками.
Отличается только движком. Файл 2,5 ГБ целиком в память не поместится, поэтому
читаем его кусками по CHUNK строк и только нужные колонки, а из каждого куска
сразу копим суммы по периодам. Все правила агрегации сводятся к суммам:
взвешенное среднее = сумма произведений / сумма весов, простое среднее =
сумма / число недель. Поэтому куски складываются без потерь.

Шаги:
  1. Читает MRVNODPP_FACT_DATA кусками и переводит сырые целые в настоящие
     значения: сырое / 10^PRECISION из FCT.
  2. Попутно проверяет выгрузку: группы сходятся (A = B + C, C = D + E), три
     корня TOTAL NOODLES равны, формулы Nielsen подтверждаются, нет дублей
     и чужих тегов.
  3. Собирает периоды из недель по флагам PER: месяцы 4-4-5, кварталы,
     полугодия, годы, MAT, YTD. Периоды, обрезанные началом или концом базы,
     отбрасывает.
  4. Пишет три CSV для Excel: pp_week (недели), pp_period (периоды), pp_sku (SKU).

Правила агрегации недель — из справки Nielsen по фактам Promo Pressure:
  кг, рубли, упаковки — сумма; цены и доли — из сумм;
  ND и WD группы — среднее, взвешенное на кг этой группы;
  скидка и длина TPR — среднее, взвешенное на кг TPR;
  интенсивность = скидка × доля TPR в кг (так её считает и Nielsen: проверено на данных);
  число SKU и магазинов — простое среднее по неделям с продажами;
  регулярная цена на полке — среднее, взвешенное на число магазинов;
  частота промо = сумма магазинов с TPR по неделям / среднее число магазинов.

Запуск: python pp_extract_pd.py     (нужен только pandas)
"""

import sys
import time
from collections import Counter
from pathlib import Path

print(">>> pp_extract_pd.py стартовал", flush=True)
try:
    import pandas as pd
except ImportError as e:
    print(f"!!! Не хватает пакета: {e}. Установи: pip install pandas", flush=True)
    raise SystemExit(1)

_t0 = time.time()


def log(msg):
    """Печать с отметкой времени от старта: видно, где скрипт тратит время."""
    print(f"[{time.time() - _t0:6.1f}s] {msg}", flush=True)


# ============================ CONFIG ============================
BASE     = Path(__file__).resolve().parent
DATA_DIR = BASE / "data"          # файлы MRVNODPP_*
PREFIX   = "MRVNODPP"
OUT_DIR  = BASE / "out"

# Сколько строк fact_data читать за раз. 500 тыс. строк — около 300 МБ памяти
# на кусок. Если ноутбук задыхается, уменьши до 200_000: станет медленнее, но легче.
CHUNK = 500_000

# Рынки, которые не берём (подстрока в SHORT, регистр не важен):
# состав топ-5 и топ-10 сетей Nielsen не раскрывает, объяснить их цифры нечем.
MKT_SKIP = ["TOP"]

# Узлы PROD в выгрузке: (номер иерархии, номер уровня).
#   иерархия 1 MAIN:         1 тотал, 2 производитель, 3 бренд, ..., 8 SKU
#   иерархия 2 PACKAGE TYPE: 1 тотал, 2 тип упаковки, 3 производитель, 4 бренд
# Иерархию 3 (вкусы) не берём: её корень повторяет тотал, вкусы отчёту не нужны.
WEEK_NODES   = {(1, 1), (1, 2), (1, 3), (2, 2)}
PERIOD_NODES = {(1, 1), (1, 2), (1, 3), (2, 2), (2, 3), (2, 4)}
# SKU — листья иерархии 1 (у СТМ уровня ITEM нет, её лист выше). По неделям
# и месяцам SKU не выгружаем: файл раздуется, а отчёту нужны итоги периода.
SKU_PERIODS  = ["MAT", "YTD", "Год"]
LEVEL_RU = {(1, 1): "Тотал", (1, 2): "Производитель", (1, 3): "Бренд", (2, 2): "Тип упаковки",
            (2, 3): "Упаковка × производитель", (2, 4): "Упаковка × бренд"}

# CSV как у прежних выгрузок: Excel и Power Query с русской локалью читают без настроек.
CSV = dict(sep=";", decimal=",", index=False, encoding="utf-8-sig",
           float_format="%.6f", date_format="%d.%m.%Y")
# =================================================================

TAGS = ["MKT_TAG", "PROD_TAG", "PER_TAG"]
KEYS = ["MKT_TAG", "PROD_TAG", "ptype", "period"]
# Из 82 фактов читаем только эти 27: остальные отчёту не нужны, а память не резиновая.
FACTS = ["VOL_A", "VOL_B", "VOL_C", "VOL_D", "VOL_E", "VAL_A", "VAL_B", "VAL_C", "VAL_D", "VAL_E",
         "SALES_A", "SALES_B", "SALES_C", "SALES_D", "SALES_E", "TDISCV", "TLENV", "ADOFTPR",
         "ND_A", "ND_C", "WVOLD_C", "CNT1_A", "CNT1_C", "CNT4_A", "CNT4_C", "CNT4_D", "AVGPRICEZ"]
SUMS = ["VOL_A", "VOL_B", "VOL_C", "VOL_D", "VOL_E", "VAL_A", "VAL_B", "VAL_C",
        "SALES_A", "SALES_B", "SALES_C"]
# Взвешенные средние: имя -> (факт, вес). Копим сумму факт × вес и сумму весов
# только по неделям, где факт есть: пустая скидка не должна тянуть среднее к нулю.
WAVG = {"tdisc": ("TDISCV", "VOL_D"), "tlen": ("TLENV", "VOL_D"), "nd_a": ("ND_A", "VOL_A"),
        "nd_c": ("ND_C", "VOL_C"), "wd_c": ("WVOLD_C", "VOL_C"), "shelf": ("AVGPRICEZ", "CNT4_A")}
# Сходимость групп: всё = регулярные + промо, промо = TPR + SP.
PAIRS = [(f"{m}_{a}", f"{m}_{b}", f"{m}_{c}") for m in ("VOL", "VAL", "SALES")
         for a, b, c in (("A", "B", "C"), ("C", "D", "E"))]
EMPTY = ["VOL_A", "VOL_B", "VOL_C", "VOL_D", "VOL_E", "TDISCV", "CNT4_D"]


def find(name):
    hits = sorted(DATA_DIR.glob(f"{PREFIX}_{name}*.csv"))
    if not hits:
        raise SystemExit(f"!!! Нет файла {PREFIX}_{name}*.csv в {DATA_DIR}. "
                         f"Положи туда поставку или поправь DATA_DIR в CONFIG.")
    return hits[0]


def read_dict(name):
    # Справочники маленькие, читаем целиком, всё текстом.
    # keep_default_na=False обязательно: у Nielsen «NA» — настоящее значение.
    f = find(name)
    for enc in ("utf-8-sig", "cp1251"):
        try:
            return pd.read_csv(f, sep="|", dtype=str, keep_default_na=False, encoding=enc)
        except UnicodeDecodeError:
            continue
    raise SystemExit(f"!!! Не разобрал кодировку {f.name}: ни UTF-8, ни cp1251.")


def period_map(per):
    """Неделя -> периоды, в которые она входит. Только полные периоды."""
    per = per.copy()
    per["d_to"] = pd.to_datetime(per["LONG"].str[-8:], format="%d/%m/%y")   # "Week Ending 03/09/23"
    if (per["WEEK53"] != "").any():
        log("!!! В PER заполнен WEEK53: в году 53 недели. Проверь, в какой месяц Nielsen её отнёс.")
    last_week = int(per.loc[per["d_to"].idxmax(), "WEEKNR"])
    full = {"Неделя": 1, "Квартал": 13, "Полугодие": 26, "Год": 52, "MAT": 52, "YTD": last_week}
    five = {"MARCH", "JUNE", "SEPTEMBER", "DECEMBER"}        # 4-4-5: третий месяц квартала — 5 недель
    spec = [("Неделя", "SHORT", ""), ("Месяц", "MONTHYEAR", ""), ("Квартал", "QUARTERYEAR", ""),
            ("Полугодие", "HJYEAR", ""), ("Год", "YEAR", ""),
            ("MAT", "MATYEAR", "MAT"), ("YTD", "YTDYEAR", "YTD")]
    # «№ в году» — номер недели, месяца, квартала или полугодия: по нему в сводной
    # годы ложатся друг на друга, так видна сезонность.
    months = ["JANUARY", "FEBRUARY", "MARCH", "APRIL", "MAY", "JUNE", "JULY", "AUGUST",
              "SEPTEMBER", "OCTOBER", "NOVEMBER", "DECEMBER"]
    num = {"Неделя": lambda r: int(r["WEEKNR"]), "Месяц": lambda r: months.index(r["MONTH"]) + 1,
           "Квартал": lambda r: int(r["QUARTER"][1:]), "Полугодие": lambda r: int(r["HALFYEAR"][2:])}
    rows = [(r["TAG"], t, r[col], r[rel] if rel else "", r["d_to"], r["MONTH"], int(r["YEAR"]),
             num[t](r) if t in num else None)
            for _, r in per.iterrows() for t, col, rel in spec if r[col]]
    pm = pd.DataFrame(rows, columns=["PER_TAG", "ptype", "period", "rel", "d_to", "month", "year", "num"])

    g = (pm.groupby(["ptype", "period"], sort=False)
           .agg(n=("PER_TAG", "size"), d_from=("d_to", "min"), d_to=("d_to", "max"),
                month=("month", "first"), rel=("rel", "first"), year=("year", "last"), num=("num", "last"))
           .reset_index())
    g["d_from"] = g["d_from"] - pd.Timedelta(6, unit="D")                      # неделя — пн..вс
    # Месяц, квартал, MAT и YTD часто кончаются одной неделей: порядок типов — чтобы
    # строки в файле шли одинаково при каждом запуске.
    g["tord"] = g["ptype"].map({t: i for i, (t, _, _) in enumerate(spec)})
    g["need"] = [(5 if m in five else 4) if t == "Месяц" else full[t] for t, m in zip(g.ptype, g.month)]
    cut = g[g["n"] < g["need"]]
    if len(cut):
        log("Неполные периоды, не берём: " +
            ", ".join(f"{p} ({n} из {k} нед.)" for p, n, k in zip(cut.period, cut.n, cut.need)))
    g = g[g["n"] >= g["need"]]
    pm = pm.merge(g[["ptype", "period"]], on=["ptype", "period"])
    log(f"Периоды: {len(g)} полных ({', '.join(f'{t} {k}' for t, k in g.groupby('ptype', sort=False).size().items())})")
    # Для чтения кусками нужен словарь на каждый тип периода: неделя -> метка периода.
    maps = {t: dict(zip(x["PER_TAG"], x["period"])) for t, x in pm.groupby("ptype", sort=False)}
    return maps, g[["ptype", "period", "rel", "d_from", "d_to", "n", "year", "num", "tord"]]


def pick_nodes(prod):
    p = prod.copy()
    p["ord"] = range(len(p))                      # порядок как в файле Nielsen
    p["h"] = p["HIER_NUM"].astype(int)
    p["lvl"] = p["HIER_LEVEL_NUM"].astype(int)
    key = list(zip(p["h"], p["lvl"]))
    p["in_week"] = [k in WEEK_NODES for k in key]
    p["in_period"] = [k in PERIOD_NODES for k in key]
    p["in_sku"] = (p["h"] == 1) & (p["lvl"] > 1) & ~p["TAG"].isin(set(p["PARENT_TAG"]))
    p["mareven"] = ["" if not m else "Mareven" if "MAREVEN" in m.upper() else "Конкуренты"
                    for m in p["MANUFACTURER"]]
    p["level"] = [LEVEL_RU.get(k) or ("SKU" if s else n)
                  for k, s, n in zip(key, p["in_sku"], p["HIER_LEVEL_NAME"])]
    # У СТМ нет уровня ITEM: имя листа — просто вкус («CHICKEN»). Собираем имя из пути.
    path = ["BRAND", "PRODUCT CLASS", "PACKAGE TYPE", "WEIGHT", "FLAVOUR-GENERAL FOODS"]
    p["sku_name"] = [(r["ITEM"] or " ".join(r[c] for c in path if r[c])) if r["in_sku"] else ""
                     for _, r in p.iterrows()]
    roots = p[p["lvl"] == 1]
    root = roots.loc[roots["h"] == 1, "TAG"].iloc[0]
    others = list(roots.loc[roots["h"] != 1, "TAG"])
    log(f"PROD: узлов в выгрузке {int((p.in_week | p.in_period | p.in_sku).sum())}, "
        f"из них SKU {int(p.in_sku.sum())}")
    return p[p["in_week"] | p["in_period"] | p["in_sku"]], root, others


def blocks(c):
    """Из недельных строк — слагаемые, которые можно складывать по неделям и кускам."""
    b = c[TAGS].copy()
    for f in SUMS:
        b[f] = c[f].fillna(0)
    for name, (x, w) in WAVG.items():
        ok = c[x].notna() & c[w].notna()
        b[name + "_xw"] = (c[x] * c[w]).where(ok, 0)
        b[name + "_w"] = c[w].where(ok, 0)
    b["cnt1a"], b["cnt1a_n"] = c["CNT1_A"].fillna(0), c["CNT1_A"].notna().astype(int)
    b["cnt4a"], b["cnt4a_n"] = c["CNT4_A"].fillna(0), c["CNT4_A"].notna().astype(int)
    b["cnt1c"], b["cnt4d"] = c["CNT1_C"].fillna(0), c["CNT4_D"].fillna(0)
    b["rows"] = 1
    return b


def read_fact(fct, maps, nodes, mkt_on, all_roots):
    src = find("FACT_DATA")
    with open(src, encoding="utf-8-sig") as fh:
        head = fh.readline().rstrip("\n").split("|")
    lost = [c for c in TAGS + FACTS if c not in head]
    if lost:
        raise SystemExit("!!! В fact_data нет колонок: " + ", ".join(lost) +
                         ". Поставка не та или Nielsen поменял состав фактов.")
    prec = dict(zip(fct["TAG"], fct["PRECISION"].astype(int)))

    week = set(nodes.loc[nodes["in_week"], "TAG"])
    period = set(nodes.loc[nodes["in_period"], "TAG"])
    period_sku = period | set(nodes.loc[nodes["in_sku"], "TAG"])
    wanted = week | period_sku
    mk = set(mkt_on["TAG"])

    def scan(robust):
        st = dict(rows=0, dup=0, bad=Counter(), empty=Counter(), odd=Counter(), odd_ex=set(),
                  mkt=set(), prod=set(), per=set())
        parts, roots = [], []
        if robust:
            # Медленный путь: факты текстом, в числа переводим сами. Нечисловое не роняет
            # скрипт и не теряется молча: считаем и показываем в проверках.
            opts = dict(dtype=str)
        else:
            # Быстрый путь. skipinitialspace: пустой факт Nielsen пишет то пустой строкой,
            # то пробелом, а на пробеле pandas падает («could not convert string to float: ' '»).
            opts = dict(dtype={**{c: str for c in TAGS}, **{c: "float64" for c in FACTS}},
                        skipinitialspace=True)
        log(f"читаю {src.name} ({src.stat().st_size // 2**20} МБ) кусками по {CHUNK:,} строк ...")
        with pd.read_csv(src, sep="|", usecols=TAGS + FACTS, chunksize=CHUNK,
                         encoding="utf-8-sig", **opts) as reader:
            for c in reader:
                for f in FACTS:
                    if robust:
                        num = pd.to_numeric(c[f], errors="coerce")
                        odd = c[f][num.isna() & c[f].notna()]
                        odd = odd[odd.str.strip() != ""]
                        if len(odd):
                            st["odd"][f] += len(odd)
                            st["odd_ex"].update(odd.head(3))
                        c[f] = num
                    c[f] = c[f] / 10 ** prec[f]

                # проверки — по всем строкам, до фильтров
                st["rows"] += len(c)
                st["dup"] += int(c.duplicated(TAGS).sum())
                for x, y, z in PAIRS:
                    # Каждый факт Nielsen округляет отдельно: 1 в последнем знаке — норма.
                    st["bad"][f"{x}={y}+{z}"] += int(((c[x].fillna(0) - c[y].fillna(0)
                                                       - c[z].fillna(0)).abs() > 1.5e-4).sum())
                for f in EMPTY:
                    st["empty"][f] += int(c[f].isna().sum())
                for k, t in (("mkt", "MKT_TAG"), ("prod", "PROD_TAG"), ("per", "PER_TAG")):
                    st[k].update(c[t].unique())
                roots.append(c[c["PROD_TAG"].isin(all_roots)])

                # суммы по периодам — только нужные рынки и узлы
                b = blocks(c[c["MKT_TAG"].isin(mk) & c["PROD_TAG"].isin(wanted)])
                for ptype, m in maps.items():
                    nodes_here = week if ptype == "Неделя" else period_sku if ptype in SKU_PERIODS else period
                    s = b[b["PROD_TAG"].isin(nodes_here)]
                    lab = s["PER_TAG"].map(m)
                    s = s[lab.notna()].drop(columns="PER_TAG").assign(ptype=ptype, period=lab[lab.notna()])
                    parts.append(s.groupby(KEYS, sort=False).sum())
                log(f"  прочитано {st['rows']:,} строк")

        # Узел-период может попасть в два соседних куска: складываем ещё раз.
        agg = pd.concat(parts).groupby(level=KEYS, sort=False).sum().reset_index()
        log(f"  строк в агрегате: {len(agg):,}")
        return agg, pd.concat(roots), st

    try:
        return scan(robust=False)
    except ValueError as e:
        if "could not convert" not in str(e):
            raise
        log(f"!!! В fact_data нечисловое значение ({e}). Перечитываю медленнее, факты текстом, "
            f"чтобы найти и посчитать все такие ячейки ...")
        return scan(robust=True)


def checks(st, roots, root, other_roots, dicts):
    log(f"fact_data: {st['rows']:,} строк; рынков {len(st['mkt'])}, узлов PROD {len(st['prod'])}, "
        f"недель {len(st['per'])} ({min(st['per'])} … {max(st['per'])})")
    log(("!!! " if st["dup"] else "") + f"Дублей строк внутри кусков (рынок × узел × неделя): {st['dup']}")
    alien = [len(st[k] - set(dicts[k]["TAG"])) for k in ("mkt", "prod", "per")]
    log(("!!! " if any(alien) else "") +
        f"Тегов вне справочников: MKT {alien[0]}, PROD {alien[1]}, PER {alien[2]}")
    if st["odd"]:
        log("!!! Нечисловые значения в фактах (взяты как пропуск): " +
            ", ".join(f"{f} {n:,}" for f, n in st["odd"].items()) +
            "; примеры: " + ", ".join(repr(x) for x in sorted(st["odd_ex"])[:10]))
    else:
        log("Нечисловых значений в фактах нет (пустые и пробелы — это пропуски)")
    if any(st["bad"].values()):
        log("!!! Группы не сходятся (строк): " + ", ".join(f"{k}: {v}" for k, v in st["bad"].items() if v))
    else:
        log("Группы сходятся: A = B + C и C = D + E по кг, рублям и упаковкам (с точностью до округления)")

    r1 = roots[roots["PROD_TAG"] == root]
    if other_roots:
        j = r1.merge(roots[roots["PROD_TAG"].isin(other_roots)], on=["MKT_TAG", "PER_TAG"], suffixes=("", "_o"))
        nbad = int((((j["VOL_A"] - j["VOL_A_o"]).abs() > 1.5e-4) | ((j["VAL_A"] - j["VAL_A_o"]).abs() > 1.5e-4)
                    | ((j["VOL_C"].fillna(0) - j["VOL_C_o"].fillna(0)).abs() > 1.5e-4)).sum())
        log(("!!! " if nbad or not len(j) else "") + f"Корни TOTAL NOODLES в иерархиях: сравнено {len(j):,} пар, "
            f"расхождений {nbad}")

    # Формулы Nielsen на строках тотала: интенсивность = скидка × доля TPR;
    # магазины / ND одинаковы у всех групп (это вселенная канала — нужна для частоты).
    r = r1[(r1["VOL_A"] > 0) & (r1["VOL_D"] > 0) & (r1["ND_A"] > 0) & (r1["ND_C"] > 1)]
    d_int = (r["ADOFTPR"] - r["TDISCV"] * r["VOL_D"] / r["VOL_A"]).abs().max()
    d_univ = 100 * ((r["CNT4_C"] / r["ND_C"]) / (r["CNT4_A"] / r["ND_A"]) - 1).abs().max()
    d_int, d_univ = (0 if pd.isna(d_int) else d_int), (0 if pd.isna(d_univ) else d_univ)
    log(("!!! " if d_int > 1e-3 else "") +
        f"Интенсивность = скидка × доля TPR: макс. расхождение {d_int:.5f} п.п.")
    log(("!!! " if d_univ > 0.1 else "") +
        f"Вселенная магазинов (магазины / ND) по группам: расхождение до {d_univ:.4f}%")
    log("Пустые ячейки, % строк: " + ", ".join(f"{f} {100 * st['empty'][f] / st['rows']:.1f}" for f in EMPTY))
    if st["empty"]["VOL_A"]:
        log("!!! Есть строки без общего объёма VOL_A — посмотри их в fact_data")


def channel_universe(roots, root, maps):
    """Вселенная канала (все его магазины) из строки тотала: ND = магазины / вселенная * 100.
    Это свойство рынка, поэтому среднее берём по всем неделям периода."""
    u = roots[(roots["PROD_TAG"] == root) & (roots["ND_A"] > 0)]
    u = u.assign(univ=u["CNT4_A"] / u["ND_A"] * 100)[["MKT_TAG", "PER_TAG", "univ"]]
    parts = [u.assign(ptype=t, period=u["PER_TAG"].map(m)).dropna(subset=["period"])
             for t, m in maps.items() if t != "Неделя"]
    return pd.concat(parts).groupby(["MKT_TAG", "ptype", "period"], as_index=False)["univ"].mean()


def build(agg, nodes, mkt_on, periods, univ):
    # Из справочника узлов — только нужные колонки: агрегат длинный, лишние раздувают память.
    keep = ["TAG", "SHORT", "MANUFACTURER", "BRAND", "PACKAGE TYPE", "sku_name", "mareven", "level",
            "ord", "in_period", "in_sku"]
    a = (agg.merge(nodes[keep], left_on="PROD_TAG", right_on="TAG")
            .merge(mkt_on[["TAG", "SHORT", "ord"]].rename(columns={"TAG": "MKT_TAG", "SHORT": "mkt", "ord": "mord"}),
                   on="MKT_TAG")
            .merge(periods.rename(columns={"n": "weeks"}), on=["ptype", "period"])
            .merge(univ, on=["MKT_TAG", "ptype", "period"], how="left")
            .sort_values(["mord", "ord", "d_to", "tord"], kind="stable"))
    nz = lambda s: s.where(s != 0)                               # как nullif(x, 0) в SQL
    wavg = lambda k: a[k + "_xw"] / nz(a[k + "_w"])
    tdisc = wavg("tdisc")
    st_a = a["cnt4a"] / nz(a["cnt4a_n"])
    not_week = a["ptype"] != "Неделя"                            # частота валидна от месяца
    o = pd.DataFrame({
        "Рынок": a["mkt"], "Уровень": a["level"], "Узел": a["SHORT"],
        "Производитель": a["MANUFACTURER"], "Бренд": a["BRAND"], "Тип упаковки": a["PACKAGE TYPE"],
        "SKU": a["sku_name"], "Mareven": a["mareven"],
        "Тип периода": a["ptype"], "Период": a["period"], "Отн. период": a["rel"],
        "Неделя с": a["d_from"], "Неделя по": a["d_to"], "Недель": a["weeks"],
        "Год": a["year"], "№ в году": a["num"].astype("Int64"),
        "Объём, т": a["VOL_A"], "Продажи, тыс. руб.": a["VAL_A"], "Упаковки, тыс. шт.": a["SALES_A"],
        "Доля промо в кг, %": 100 * a["VOL_C"] / nz(a["VOL_A"]),
        "Доля промо в руб., %": 100 * a["VAL_C"] / nz(a["VAL_A"]),
        "Доля TPR в кг, %": 100 * a["VOL_D"] / nz(a["VOL_A"]),
        "Доля SP в кг, %": 100 * a["VOL_E"] / nz(a["VOL_A"]),
        "Скидка TPR, %": -tdisc,
        "Длина TPR, нед.": wavg("tlen"),
        "Интенсивность TPR, %": -tdisc.fillna(0) * a["VOL_D"] / nz(a["VOL_A"]),
        "Частота промо, нед. на магазин с продажами": (a["cnt4d"] / nz(st_a)).where(not_week),
        "Частота промо, нед. на магазин канала": (a["cnt4d"] / nz(a["univ"])).where(not_week),
        "ND, %": wavg("nd_a"), "ND промо, %": wavg("nd_c").fillna(0), "WD промо, %": wavg("wd_c").fillna(0),
        "Цена рег., руб./уп.": a["VAL_B"] / nz(a["SALES_B"]),
        "Цена промо, руб./уп.": a["VAL_C"] / nz(a["SALES_C"]),
        "Цена рег., руб./кг": a["VAL_B"] / nz(a["VOL_B"]),
        "Цена промо, руб./кг": a["VAL_C"] / nz(a["VOL_C"]),
        "Рег. цена на полке, руб./уп.": wavg("shelf"),
        "SKU в продаже, шт.": a["cnt1a"] / nz(a["cnt1a_n"]), "SKU в промо, шт.": a["cnt1c"] / a["rows"],
        "Магазинов с продажами": st_a, "Магазинов с TPR": a["cnt4d"] / a["rows"],
    })
    # Строк за узел-период больше, чем недель в периоде, — значит, в fact_data дубли
    # (между кусками их не поймать иначе). Сводная задвоила бы такие цифры.
    dup = int((a["rows"] > a["weeks"]).sum())
    if dup:
        log(f"!!! В {dup} строках выгрузки недель больше, чем в периоде: в fact_data дубли")
    return o, a


def main():
    OUT_DIR.mkdir(exist_ok=True)
    fct, mkt, per, prod = (read_dict(x) for x in ("FCT", "MKT", "PER", "PROD"))
    mkt["ord"] = range(len(mkt))
    log(f"Справочники: FCT {len(fct)}, MKT {len(mkt)}, PER {len(per)}, PROD {len(prod)}")

    skip = [any(s.upper() in name.upper() for s in MKT_SKIP) for name in mkt["SHORT"]]
    mkt_on = mkt[[not s for s in skip]]
    log("Рынки не берём: " + (", ".join(mkt.loc[skip, "SHORT"]) or "—"))
    twins = mkt_on["SHORT"][mkt_on["SHORT"].duplicated()]
    if len(twins):
        log("!!! Одно имя у нескольких тегов рынка (цифры задвоятся в сводной): " + ", ".join(twins))
    nodes, root, other_roots = pick_nodes(prod)
    maps, periods = period_map(per)

    agg, roots, st = read_fact(fct, maps, nodes, mkt_on, [root] + other_roots)
    checks(st, roots, root, other_roots, {"mkt": mkt, "prod": prod, "per": per})
    out, a = build(agg, nodes, mkt_on, periods, channel_universe(roots, root, maps))

    # Рынки: объём и промо тотала за последний MAT — по ним видно, что каналы
    # в сумме дают тотал, а регионы — город.
    mat = periods.loc[(periods["ptype"] == "MAT") & (periods["rel"] == "MAT"), "period"]
    is_root = (a["PROD_TAG"] == root).to_numpy()
    if len(mat):
        t = out[is_root & (out["Период"] == mat.iloc[0]).to_numpy()]
        t = pd.DataFrame({"Рынок": t["Рынок"], "Объём, т": t["Объём, т"].round(1),
                          "Промо, % кг": t["Доля промо в кг, %"].round(2), "Скидка, %": t["Скидка TPR, %"].round(2)})
        log(f"Тотал лапши по рынкам, {mat.iloc[0]}:\n" + t.to_string(index=False))

    files = {"pp_week.csv":   (out["Тип периода"] == "Неделя").to_numpy(),
             "pp_period.csv": ((out["Тип периода"] != "Неделя") & a["in_period"]).to_numpy(),
             "pp_sku.csv":    (a["in_sku"] & out["Тип периода"].isin(SKU_PERIODS)).to_numpy()}
    for fname, rows in files.items():
        out[rows].to_csv(OUT_DIR / fname, **CSV)
        log(f"{fname}: {int(rows.sum()):,} строк, {(OUT_DIR / fname).stat().st_size // 2**20} МБ")

    # Сверка с Молекулой: тотал на главном рынке (самый большой объём за последний MAT).
    if len(mat):
        main_mkt = t.loc[t["Объём, т"].idxmax(), "Рынок"]
        s = out[is_root & (out["Рынок"] == main_mkt).to_numpy()]
        s = pd.concat([s[s["Тип периода"] == "MAT"], s[s["Тип периода"] == "YTD"],
                       s[s["Тип периода"] == "Квартал"].tail(4)])
        cols = ["Период", "Недель", "Объём, т", "Доля промо в кг, %", "Доля промо в руб., %",
                "Скидка TPR, %", "Длина TPR, нед.", "Интенсивность TPR, %",
                "Частота промо, нед. на магазин с продажами", "Частота промо, нед. на магазин канала",
                "ND промо, %"]
        log(f"Для сверки с Молекулой — TOTAL NOODLES, {main_mkt}:\n" +
            s[cols].round(2).to_string(index=False))
    log("Готово.")


if __name__ == "__main__":
    main()
