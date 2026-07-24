"""Движок анализа закупок: поиск «узких мест» и скоринг ниш.

Вход  : нормализованная таблица (CSV/Parquet) по схеме schema.COLUMNS.
Выход : ranked-таблица по связкам (ОКПД2 x регион) с метриками узких мест
        и композитным opportunity_score, + Excel/CSV.

Логика источник-независима: неважно, откуда пришли данные (FTP ЕИС или
агрегатор) — лишь бы колонки совпадали со схемой.

Запуск:
    python analyze.py --input data/normalized.parquet --out out/niches.xlsx
    python analyze.py --input data/normalized.csv --group okpd2,region,customer_inn
"""
from __future__ import annotations

import argparse
from pathlib import Path

import numpy as np
import pandas as pd

# Веса сигналов в композитном скоре (сумма = 1.0). Настраиваются под гипотезу.
WEIGHTS = {
    "growth": 0.30,          # рост суммы закупок (recent 12m vs prior 12m)
    "low_competition": 0.25, # мало участников / единственный поставщик
    "dysfunction": 0.20,     # несостоявшиеся + отменённые (заказчик страдает)
    "concentration": 0.15,   # деньги держат 1-3 поставщика
    "repeatability": 0.10,   # повторяющиеся закупки одним заказчиком
}


def _minmax(s: pd.Series) -> pd.Series:
    """Нормировка в 0..1; константный столбец -> 0.5 (нейтрально)."""
    s = s.astype(float)
    lo, hi = s.min(), s.max()
    if not np.isfinite(lo) or not np.isfinite(hi) or hi - lo < 1e-12:
        return pd.Series(0.5, index=s.index)
    return (s - lo) / (hi - lo)


def load(path: str) -> pd.DataFrame:
    p = Path(path)
    if p.suffix.lower() in (".parquet", ".pq"):
        df = pd.read_parquet(p)
    else:
        df = pd.read_csv(p)
    df["publish_date"] = pd.to_datetime(df["publish_date"], errors="coerce")
    for col in ("is_failed", "is_canceled", "has_advance", "is_smp"):
        if col in df.columns:
            df[col] = df[col].astype("boolean").fillna(False)
    df["nmck"] = pd.to_numeric(df["nmck"], errors="coerce")
    df["participants_count"] = pd.to_numeric(df["participants_count"], errors="coerce")
    return df


def _clean_suppliers(winners: pd.Series, sums=None):
    """Возвращает (доли по сумме или по числу, серия сумм по поставщику)."""
    w = winners.astype("string")
    mask = w.notna() & (w.str.len() > 0)
    if not mask.any():
        return None
    if sums is not None:
        by = pd.Series(sums[mask].values, index=w[mask].values).groupby(level=0).sum()
    else:
        by = w[mask].value_counts()
    total = by.sum()
    if total <= 0:
        return None
    return by.sort_values(ascending=False) / total


def _hhi(winners: pd.Series, sums=None) -> float:
    """Индекс Херфиндаля по долям поставщиков (0..1). Выше = концентрированнее."""
    shares = _clean_suppliers(winners, sums)
    return float((shares.values ** 2).sum()) if shares is not None else np.nan


def _crn(winners: pd.Series, n=3, sums=None) -> float:
    """Доля топ-N поставщиков (CR-N), 0..1. Напр. CR3 > 0.7 = концентрировано."""
    shares = _clean_suppliers(winners, sums)
    return float(shares.head(n).sum()) if shares is not None else np.nan


def _n_suppliers(winners: pd.Series) -> int:
    w = winners.astype("string")
    return int(w[w.notna() & (w.str.len() > 0)].nunique())


def analyze(df: pd.DataFrame, group_cols: list[str]) -> pd.DataFrame:
    now = df["publish_date"].max()
    recent_start = now - pd.DateOffset(months=12)
    prior_start = now - pd.DateOffset(months=24)

    rows = []
    for keys, g in df.groupby(group_cols, dropna=False):
        keys = keys if isinstance(keys, tuple) else (keys,)
        n = len(g)

        recent = g[g["publish_date"] >= recent_start]
        prior = g[(g["publish_date"] >= prior_start) & (g["publish_date"] < recent_start)]
        sum_recent = recent["nmck"].sum()
        sum_prior = prior["nmck"].sum()
        # рост суммы: (recent-prior)/prior, ограничен, чтобы выбросы не рвали шкалу
        if sum_prior > 0:
            growth = np.clip((sum_recent - sum_prior) / sum_prior, -1, 5)
        else:
            growth = 5.0 if sum_recent > 0 else 0.0

        no_comp = (g["participants_count"].fillna(0) <= 1).mean()
        failed = g["is_failed"].mean()
        canceled = g["is_canceled"].mean()
        dysfunction = float(failed + canceled)
        # концентрация поставщиков — по сумме контрактов (nmck), где есть победитель
        hhi = _hhi(g["winner_inn"], g["nmck"])
        cr3 = _crn(g["winner_inn"], 3, g["nmck"])
        n_suppliers = _n_suppliers(g["winner_inn"])
        # повторяемость: доля закупок, приходящихся на заказчиков-повторников
        cust_counts = g["customer_inn"].value_counts()
        repeat_customers = int((cust_counts >= 3).sum())
        repeatability = float(g["customer_inn"].isin(cust_counts[cust_counts >= 3].index).mean())

        rows.append({
            **{c: k for c, k in zip(group_cols, keys)},
            "lots_count": n,
            "total_sum": float(g["nmck"].sum()),
            "avg_check": float(g["nmck"].mean()),
            "sum_recent_12m": float(sum_recent),
            "sum_prior_12m": float(sum_prior),
            "growth": float(growth),
            "avg_participants": float(g["participants_count"].mean()),
            "no_competition_share": float(no_comp),
            "failed_share": float(failed),
            "canceled_share": float(canceled),
            "dysfunction": dysfunction,
            "winner_hhi": hhi,
            "cr3": cr3,
            "n_suppliers": n_suppliers,
            "repeat_customers": repeat_customers,
            "repeatability": repeatability,
            "smp_share": float(g["is_smp"].mean()),
        })

    res = pd.DataFrame(rows)
    if res.empty:
        return res

    # Композитный скор из нормированных сигналов.
    res["_growth"] = _minmax(res["growth"])
    res["_low_competition"] = _minmax(res["no_competition_share"])
    res["_dysfunction"] = _minmax(res["dysfunction"])
    res["_concentration"] = _minmax(res["winner_hhi"].fillna(res["winner_hhi"].median()))
    res["_repeatability"] = _minmax(res["repeatability"])

    res["opportunity_score"] = (
        WEIGHTS["growth"] * res["_growth"]
        + WEIGHTS["low_competition"] * res["_low_competition"]
        + WEIGHTS["dysfunction"] * res["_dysfunction"]
        + WEIGHTS["concentration"] * res["_concentration"]
        + WEIGHTS["repeatability"] * res["_repeatability"]
    ).round(4)

    # barrier_manual — CAPEX/порог входа из данных не выводится: заполняется руками
    # (1 = низкий порог … 5 = высокий). Итоговый приоритет учитывает и его.
    res["barrier_manual"] = np.nan
    res = res.drop(columns=[c for c in res.columns if c.startswith("_")])
    return res.sort_values("opportunity_score", ascending=False).reset_index(drop=True)


def main() -> None:
    ap = argparse.ArgumentParser(description="Скоринг ниш госзакупок по узким местам")
    ap.add_argument("--input", required=True, help="normalized CSV/Parquet")
    ap.add_argument("--out", default="out/niches.xlsx", help="выходной файл (.xlsx/.csv)")
    ap.add_argument("--group", default="okpd2,region",
                    help="колонки группировки через запятую (напр. okpd2,region,customer_inn)")
    ap.add_argument("--min-lots", type=int, default=3,
                    help="минимум лотов в связке, чтобы попасть в отчёт")
    args = ap.parse_args()

    df = load(args.input)
    group_cols = [c.strip() for c in args.group.split(",") if c.strip()]
    res = analyze(df, group_cols)
    res = res[res["lots_count"] >= args.min_lots].reset_index(drop=True)

    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.suffix.lower() == ".csv":
        res.to_csv(out, index=False, encoding="utf-8-sig")
    else:
        res.to_excel(out, index=False)

    with pd.option_context("display.max_columns", None, "display.width", 200):
        print(f"Связок отобрано: {len(res)}  |  файл: {out}\n")
        print("ТОП-10 ниш по opportunity_score:")
        cols = group_cols + ["lots_count", "total_sum", "growth",
                             "n_suppliers", "cr3", "winner_hhi", "dysfunction",
                             "opportunity_score"]
        print(res[cols].head(10).to_string(index=False))


if __name__ == "__main__":
    main()
