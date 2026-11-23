"""Выгрузка таблиц X5 Dialog из ClickHouse в Parquet, целиком.

Запуск:  python x5_export.py
Нужно:   pip install requests pyarrow
Логин и пароль берутся из переменных среды CH_USER и CH_PASSWORD (настраиваются один раз, в скрипте их нет).

Файл скачивается во временный и заменяет старый, только если в нём ровно столько строк, сколько в таблице клика.
"""
import os
from pathlib import Path

import pyarrow.parquet as pq
import requests

# ---------------- НАСТРОЙКИ ----------------
URL = "http://ХОСТ:8123/"                                    # HTTP-интерфейс клика (HTTPS обычно на порту 8443)
TABLES = {"weekly": "БАЗА.ТАБЛИЦА_НЕДЕЛИ", "monthly": "БАЗА.ТАБЛИЦА_МЕСЯЦЫ"}
OUT = Path(r"R:\2.ДМБ\БЭД\ОПИ\02_Аналитика\03_Базы данных\01_X5 Dialog\Raw data\ClickHouse")

AUTH = (os.environ["CH_USER"], os.environ["CH_PASSWORD"])


def ask(sql, stream=False):
    """GET-запрос к клику: такие запросы клик выполняет только на чтение."""
    r = requests.get(URL, params={"query": sql}, auth=AUTH, stream=stream, timeout=600)
    if r.status_code != 200:
        raise SystemExit(f"СТОП. Клик ответил {r.status_code} на «{sql[:60]}»: {r.text[:500]}")
    return r


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    for name, table in TABLES.items():
        expected = int(ask(f"SELECT count() FROM {table}").text)
        final, tmp = OUT / f"x5_{name}.parquet", OUT / f"x5_{name}.parquet.part"
        with ask(f"SELECT * FROM {table} FORMAT Parquet", stream=True) as r, open(tmp, "wb") as f:
            for chunk in r.iter_content(chunk_size=8 << 20):
                f.write(chunk)
        try:
            got = pq.ParquetFile(tmp).metadata.num_rows
        except Exception as e:
            tmp.unlink()
            raise SystemExit(f"СТОП. {table}: файл скачался битым, клик оборвал выгрузку ({e}). Старый файл не тронут.")
        if got != expected:
            tmp.unlink()
            raise SystemExit(f"СТОП. {table}: в клике {expected} строк, скачано {got}. Старый файл не тронут.")
        tmp.replace(final)
        print(f"{table}: {got} строк -> {final.name} ({final.stat().st_size / 1e6:.0f} МБ)")


if __name__ == "__main__":
    main()
