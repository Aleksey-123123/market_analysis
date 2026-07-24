"""Загрузчик выгрузки из Контур.Закупки (Excel/CSV) -> нормализованная схема.

Контур.Закупки (zakupki.kontur.ru) отдаёт результаты поиска выгрузкой в
Excel/CSV. Этот загрузчик толерантно сопоставляет русские заголовки колонок
Контура с нашей схемой (schema.COLUMNS), поэтому работает даже если названия
колонок немного отличаются от версии к версии.

    python loaders/kontur_csv.py --input export.xlsx --out ../data/normalized.parquet

Если какая-то колонка не распозналась — скрипт напечатает, что не смог
сопоставить; пришлите строку заголовков, и я дополню KANDIDATES.
"""
from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from schema import COLUMNS  # noqa: E402

# Порядок важен: специфичные поля раньше общих (напр. "инн победителя" до "инн").
# Каждая колонка сопоставляется максимум одному полю. Подстроки учитывают, что
# в двухуровневой шапке Контура имена вида "Заказчик :: ИНН", "Результат :: ИНН
# победителя" — квалифицированы группой, что и снимает неоднозначность.
KANDIDATES: list[tuple[str, list[str]]] = [
    ("purchase_id",        ["реестровый номер", "номер закупки", "номер извещения", "№ закупки", "номер тендера", ":: номер", "закупка :: номер"]),
    ("fz",                 ["тип торгов", "закон", "тип закупки по фз", "44-фз", "223-фз", " фз"]),
    ("publish_date",       ["дата публикац", "дата размещен", "опубликован", "дата начала"]),
    ("okpd2",              ["окпд"]),
    ("region",             ["регион", "субъект"]),
    ("purchase_method",    ["способ отбора", "способ закупки", "способ определения", "способ размещения"]),
    ("nmck",               ["начальная цена", "начальная (макс", "нмц", "цена контракта", "цена договора", "сумма закупки"]),
    ("participants_count", ["количество заявок", "подано заявок", "число участник", "участников"]),
    ("execution_days",     ["срок исполнения", "срок поставки", "срок контракта"]),
    ("_advance",           ["аванс"]),                               # служебное -> has_advance
    ("winner_inn",         ["инн победителя", "инн поставщика"]),
    ("winner_name",        ["название победителя", "победитель", "поставщик"]),
    ("customer_inn",       ["инн заказчика", "инн организатора", "заказчик :: инн", "инн"]),
    ("customer_name",      ["заказчик :: название", "заказчик", "организатор"]),
    ("_status",            ["этап отбора", "статус", "состояние"]),  # служебное -> флаги failed/canceled
    ("_smp",               ["смп", "мсп", "субъект малого"]),        # служебное -> is_smp
]


def read_kontur(path: Path, sheet=0) -> pd.DataFrame:
    """Читает выгрузку Контура, поддерживая двухуровневую шапку.

    Верхняя строка (группы «Закупка/Заказчик/Результат») склеивается с нижней,
    что различает дубли «Название»/«ИНН». Если шапка одноуровневая — берётся как есть.
    """
    if path.suffix.lower() in (".csv", ".tsv"):
        return pd.read_csv(path, dtype=str, sep=None, engine="python")

    raw = pd.read_excel(path, sheet_name=sheet, header=None, dtype=str)
    r0, r1 = raw.iloc[0], raw.iloc[1]
    two_row = r0.notna().sum() < r1.notna().sum() and r0.notna().sum() <= len(r0) * 0.5
    if two_row:
        groups = r0.ffill()
        cols = []
        for g, s in zip(groups, r1):
            g = "" if pd.isna(g) else str(g).strip()
            s = "" if pd.isna(s) else str(s).strip()
            cols.append(f"{g} :: {s}" if g and s else (s or g))
        data = raw.iloc[2:].copy()
    else:
        cols = [("" if pd.isna(c) else str(c).strip()) for c in r0]
        data = raw.iloc[1:].copy()
    data.columns = cols
    return data.reset_index(drop=True)


def _to_float(v) -> float | None:
    if pd.isna(v):
        return None
    s = str(v).replace("\xa0", "").replace(" ", "").replace("₽", "").replace("руб", "")
    s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return None


def resolve_mapping(headers: list[str]) -> dict[str, str]:
    """schema_field -> имя колонки в файле."""
    low = {h: str(h).strip().lower() for h in headers}
    used: set[str] = set()
    mapping: dict[str, str] = {}
    for field, subs in KANDIDATES:
        for h in headers:
            if h in used:
                continue
            if any(s in low[h] for s in subs):
                mapping[field] = h
                used.add(h)
                break
    return mapping


def normalize(df_raw: pd.DataFrame, okpd2_label: str | None = None) -> pd.DataFrame:
    mapping = resolve_mapping(list(df_raw.columns))
    print("Сопоставление колонок:")
    for field, _ in KANDIDATES:
        print(f"  {field:20s} <- {mapping.get(field, '— НЕ НАЙДЕНО')}")

    out = pd.DataFrame(index=df_raw.index)
    for col in COLUMNS:
        src = mapping.get(col)
        out[col] = df_raw[src] if src else None

    # fz -> "44"/"223"
    if "fz" in mapping:
        fz = df_raw[mapping["fz"]].astype(str)
        out["fz"] = fz.str.extract(r"(44|223)")[0]

    out["nmck"] = df_raw[mapping["nmck"]].map(_to_float) if "nmck" in mapping else None
    if "participants_count" in mapping:
        out["participants_count"] = pd.to_numeric(df_raw[mapping["participants_count"]], errors="coerce").astype("Int64")
    out["publish_date"] = pd.to_datetime(out["publish_date"], errors="coerce", dayfirst=True).dt.date

    # флаги из статуса (без пробелов: ловит и "не состоялась", и "несостоявшаяся")
    status = df_raw[mapping["_status"]].astype(str).str.lower() if "_status" in mapping else pd.Series("", index=df_raw.index)
    status_ns = status.str.replace(r"\s+", "", regex=True)
    out["is_failed"] = status_ns.str.contains("несостоя", na=False)
    out["is_canceled"] = status_ns.str.contains("отмен", na=False)

    # СМП: отдельная колонка-флаг или упоминание в способе/статусе
    if "_smp" in mapping:
        smp = df_raw[mapping["_smp"]].astype(str).str.lower()
        out["is_smp"] = smp.str.contains("да|смп|мсп|1|true", na=False)
    else:
        method = out["purchase_method"].astype(str).str.lower()
        out["is_smp"] = method.str.contains("смп|мсп", na=False)

    # аванс: колонка заполнена (процент/сумма) => True
    if "_advance" in mapping:
        adv = df_raw[mapping["_advance"]].astype(str).str.strip().str.lower()
        out["has_advance"] = adv.notna() & ~adv.isin(["", "nan", "нет", "0", "0%", "без аванса", "none"])
    else:
        out["has_advance"] = False

    # ОКПД2 в выгрузке Контура часто отсутствует (поиск уже отфильтрован по нему).
    # Проставляем константу-метку, чтобы группировать/различать выгрузки.
    if okpd2_label and (out["okpd2"].isna().all() or "okpd2" not in mapping):
        out["okpd2"] = okpd2_label

    return out[COLUMNS]


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="export.xlsx или export.csv из Контур.Закупки")
    ap.add_argument("--out", default="../data/normalized.parquet")
    ap.add_argument("--sheet", default=0, help="лист Excel (имя или индекс)")
    ap.add_argument("--okpd2", default=None, help="метка ОКПД2, если её нет в файле (напр. '24.10/24.45')")
    args = ap.parse_args()

    p = Path(args.input)
    df_raw = read_kontur(p, sheet=args.sheet)
    df_raw.columns = [str(c).strip() for c in df_raw.columns]

    out = normalize(df_raw, okpd2_label=args.okpd2)
    dest = Path(args.out)
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.suffix.lower() in (".parquet", ".pq"):
        out.to_parquet(dest, index=False)
    else:
        out.to_csv(dest, index=False, encoding="utf-8-sig")
    print(f"\nНормализовано {len(out)} строк -> {dest}")


if __name__ == "__main__":
    main()
