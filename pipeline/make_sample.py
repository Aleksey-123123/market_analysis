"""Генератор синтетической нормализованной таблицы закупок.

Нужен, чтобы протестировать движок analyze.py без доступа к реальным данным
ЕИС. Данные фейковые, но структура и распределения правдоподобные: несколько
ОКПД2 x регионов с разным поведением (рост/конкуренция/несостоявшиеся).

    python make_sample.py --rows 4000 --out data/normalized.parquet
"""
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

from schema import COLUMNS

REGIONS = ["77 Москва", "54 Новосибирск", "66 Свердловская", "24 Красноярский"]
# ОКПД2 -> "профиль" ниши: (базовое НМЦК, вероятность мало участников,
#                           доля несостоявшихся, множитель роста recent/prior)
OKPD2 = {
    "24.10.6 Прокат/заготовки чёрных металлов": (4_500_000, 0.65, 0.28, 2.4),
    "28.92.2 Машины для добычи/карьеров (ЗИП)": (1_200_000, 0.55, 0.22, 1.9),
    "20.59.5 Промышленная химия/реагенты":       (800_000,  0.35, 0.10, 1.2),
    "43.99.9 Спецработы/ремонт ИССО":            (9_000_000, 0.45, 0.20, 1.6),
    "62.01.2 Разработка ПО":                     (2_000_000, 0.15, 0.05, 0.9),
    "31.01.1 Мебель офисная":                    (300_000,  0.10, 0.03, 0.7),
}


def make(rows: int, seed: int = 42) -> pd.DataFrame:
    rng = np.random.default_rng(seed)
    okpd_keys = list(OKPD2)
    end = pd.Timestamp("2026-06-30")
    start = end - pd.DateOffset(months=24)
    span_days = (end - start).days

    recs = []
    for i in range(rows):
        okpd = rng.choice(okpd_keys)
        base_nmck, p_lowcomp, p_failed, growth_mult = OKPD2[okpd]
        region = rng.choice(REGIONS)
        fz = rng.choice(["44", "223"], p=[0.4, 0.6])

        # дата: growth_mult>1 => больше закупок в последних 12 мес.
        recent = rng.random() < (growth_mult / (growth_mult + 1))
        if recent:
            day = span_days - int(rng.integers(0, 365))
        else:
            day = int(rng.integers(0, max(1, span_days - 365)))
        pub = start + pd.Timedelta(days=day)

        low_comp = rng.random() < p_lowcomp
        participants = int(rng.integers(0, 2)) if low_comp else int(rng.integers(2, 8))
        failed = rng.random() < p_failed
        canceled = rng.random() < 0.06

        # концентрация поставщиков: у «дефицитных» ниш узкий пул победителей
        pool = 3 if p_lowcomp > 0.5 else 12
        if failed or participants == 0:
            winner_inn, winner_name = "", ""
        else:
            wid = int(rng.integers(0, pool))
            winner_inn = f"77{okpd_keys.index(okpd)}{wid:04d}00"
            winner_name = f'ООО "Поставщик-{wid}"'

        nmck = float(max(50_000, rng.normal(base_nmck, base_nmck * 0.4)))
        cust = int(rng.integers(0, 8))  # 8 заказчиков на регион => повторяемость

        recs.append({
            "purchase_id": f"{fz}-{i:07d}",
            "fz": fz,
            "publish_date": pub.strftime("%Y-%m-%d"),
            "customer_inn": f"{region[:2]}CUST{cust:03d}",
            "customer_name": f"Заказчик {cust} ({region})",
            "region": region,
            "okpd2": okpd,
            "purchase_method": rng.choice(["ЭА", "запрос котировок", "ед. поставщик"]),
            "nmck": round(nmck, 2),
            "participants_count": participants,
            "winner_inn": winner_inn,
            "winner_name": winner_name,
            "is_failed": bool(failed),
            "is_canceled": bool(canceled),
            "has_advance": bool(rng.random() < 0.3),
            "execution_days": int(rng.integers(20, 200)),
            "is_smp": bool(rng.random() < 0.5),
        })

    return pd.DataFrame(recs, columns=COLUMNS)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--rows", type=int, default=4000)
    ap.add_argument("--out", default="data/normalized.parquet")
    args = ap.parse_args()
    df = make(args.rows)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.suffix.lower() in (".parquet", ".pq"):
        df.to_parquet(out, index=False)
    else:
        df.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"Сгенерировано {len(df)} строк -> {out}")


if __name__ == "__main__":
    main()
