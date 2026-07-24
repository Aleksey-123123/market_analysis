"""Загрузчик ГосПлан API (v2) -> нормализованная таблица.

ГосПлан отдаёт данные ЕИС по REST:
  * тест:  https://v2test.gosplan.info   (бесплатно, без ключа, ~10 запросов/мин)
  * прод:  https://v2.gosplan.info       (30к/год, ключ с 01.08.2026, 12000/час)
  * Swagger: https://swagger.gosplan.info (разделы 44-ФЗ / 223-ФЗ / ПП 615)

Точные пути и имена JSON-полей ГосПлана из песочницы Claude недоступны (сеть
закрыта), поэтому клиент СХЕМО-ТОЛЕРАНТНЫЙ и настраивается через config.yaml:
  1) `probe` — дёрнуть 1 страницу и распечатать реальные ключи ответа;
  2) уточнить в config.yaml пути/имена параметров, если дефолты не подошли;
  3) `pull` — постранично выкачать и сохранить нормализованный parquet.

Извлечение полей — по подстроке в имени ключа (customer+inn, winner+inn и т.п.),
что переживает мелкие отличия схемы. После `probe` пришлите ключи — докручу.

Зависимости: requests, pandas, pyarrow, pyyaml.

    python loaders/gosplan_api.py probe --config config.yaml --law 223
    python loaders/gosplan_api.py pull  --config config.yaml --law 223 --out ../data/n223.parquet
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from pathlib import Path

import pandas as pd
import requests
import yaml

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from schema import COLUMNS  # noqa: E402

# Дефолты (перекрываются секцией gosplan: в config.yaml). Пути — предположение,
# уточняются через `probe`; параметры вынесены, чтобы не хардкодить.
DEFAULTS = {
    "base_url_test": "https://v2test.gosplan.info",
    "base_url_prod": "https://v2.gosplan.info",
    "use": "test",                       # test | prod
    "api_key": "",                       # нужен только для prod
    "rpm": 10,                           # лимит теста: 10 запросов/мин
    "paths": {"44": "/api/v2/44fz/notifications",
              "223": "/api/v2/223fz/notifications"},
    "params": {                          # имена query-параметров у ГосПлана
        "page": "page",
        "size": "perPage",
        "size_value": 50,
        "date_from": "publishDateFrom",
        "date_to": "publishDateTo",
        "okpd2": "okpd2",
        "region": "region",
    },
    "start_page": 1,
    "items_path": "",                    # где в JSON лежит список (напр. "data" / "items"); "" = корень-массив
}

# Толерантный маппинг: "a+b" => ключ должен содержать И "a", И "b".
KEY_CANDIDATES: list[tuple[str, list[str]]] = [
    ("purchase_id",        ["purchasenumber", "registrationnumber", "regnumber", "notificationnumber", "noticenumber", "number"]),
    ("publish_date",       ["publishdate", "datepublished", "publisheddate", "createdate"]),
    ("okpd2",              ["okpd2", "okpd"]),
    ("region",             ["region", "subject"]),
    ("purchase_method",    ["placingway", "determination", "method", "purchasemethod"]),
    ("nmck",               ["maxprice", "startprice", "startmaxprice", "contractprice", "nmck", "price", "sum"]),
    ("participants_count", ["participantscount", "applicationscount", "bidscount", "offerscount", "participant"]),
    ("execution_days",     ["executionterm", "deliveryterm", "term"]),
    ("winner_inn",         ["winner+inn", "supplier+inn", "winnerinn"]),
    ("winner_name",        ["winner+name", "supplier+name", "winner+fullname"]),
    ("customer_inn",       ["customer+inn", "customerinn", "inn"]),
    ("customer_name",      ["customer+name", "customer+fullname", "customername"]),
    ("_status",            ["status", "stage", "state"]),
]


def cfg_get(cfg: dict) -> dict:
    m = dict(DEFAULTS)
    g = cfg.get("gosplan", {}) or {}
    m.update({k: v for k, v in g.items() if k not in ("paths", "params")})
    m["paths"] = {**DEFAULTS["paths"], **(g.get("paths") or {})}
    m["params"] = {**DEFAULTS["params"], **(g.get("params") or {})}
    return m


def flatten(obj, prefix: str = "") -> dict:
    """Разворачивает вложенный JSON в {dotted.lower.key: scalar}. Списки -> [0]."""
    out = {}
    if isinstance(obj, dict):
        for k, v in obj.items():
            out.update(flatten(v, f"{prefix}{k}".lower() if not prefix else f"{prefix}.{k}".lower()))
    elif isinstance(obj, list):
        if obj:
            out.update(flatten(obj[0], prefix))
    else:
        out[prefix] = obj
    return out


def _match(flat: dict, used: set, cands: list[str]) -> str | None:
    for cand in cands:
        parts = cand.split("+")
        for key in flat:
            if key in used:
                continue
            if all(p in key for p in parts):
                used.add(key)
                return key
    return None


def record_to_row(rec: dict) -> dict:
    flat = flatten(rec)
    used: set = set()
    row = {c: None for c in COLUMNS}
    for field, cands in KEY_CANDIDATES:
        key = _match(flat, used, cands)
        if key is not None:
            row[field if not field.startswith("_") else field] = flat[key]
    # флаги из статуса
    status = str(row.pop("_status", "") or "").lower().replace(" ", "")
    row["is_failed"] = ("несостоя" in status) or ("failed" in status) or ("notplaced" in status)
    row["is_canceled"] = ("отмен" in status) or ("cancel" in status)
    row["is_smp"] = ("смп" in status) or ("мсп" in status)
    row["has_advance"] = None
    # типы
    try:
        row["nmck"] = float(str(row["nmck"]).replace(",", ".")) if row["nmck"] not in (None, "") else None
    except (ValueError, TypeError):
        row["nmck"] = None
    pc = row.get("participants_count")
    row["participants_count"] = int(pc) if str(pc).isdigit() else None
    return {c: row.get(c) for c in COLUMNS}


def fetch_page(session, m, path, params) -> list:
    url = m[f"base_url_{m['use']}"].rstrip("/") + path
    headers = {"Authorization": f"Bearer {m['api_key']}"} if m["use"] == "prod" and m["api_key"] else {}
    for attempt in range(5):
        r = session.get(url, params=params, headers=headers, timeout=60)
        if r.status_code == 429:
            print("  429 rate limit -> ждём 60с")
            time.sleep(60)
            continue
        r.raise_for_status()
        data = r.json()
        if m["items_path"]:
            for part in m["items_path"].split("."):
                data = data.get(part, []) if isinstance(data, dict) else data
        return data if isinstance(data, list) else data.get("items", data.get("data", []))
    raise RuntimeError("не удалось получить страницу после ретраев")


def build_params(m, law, cfg) -> dict:
    p = m["params"]
    src = (cfg.get("gosplan", {}) or {}).get("query", {})
    q = {p["size"]: p["size_value"]}
    if src.get("date_from"):
        q[p["date_from"]] = src["date_from"]
    if src.get("date_to"):
        q[p["date_to"]] = src["date_to"]
    if src.get("okpd2"):
        q[p["okpd2"]] = src["okpd2"]
    if src.get("region"):
        q[p["region"]] = src["region"]
    return q


def cmd_probe(cfg, law):
    m = cfg_get(cfg)
    path = m["paths"][law]
    params = build_params(m, law, cfg)
    params[m["params"]["page"]] = m["start_page"]
    base = m["base_url_" + m["use"]]
    print("GET {}{}\n  params={}\n".format(base, path, params))
    items = fetch_page(requests.Session(), m, path, params)
    print(f"Получено записей на странице: {len(items)}")
    if items:
        print("\n--- ПЛОСКИЕ КЛЮЧИ первой записи (пришлите мне для маппинга) ---")
        for k, v in flatten(items[0]).items():
            print(f"  {k} = {str(v)[:60]}")
        print("\n--- КАК СЕЙЧАС МАПИТСЯ ---")
        print(json.dumps(record_to_row(items[0]), ensure_ascii=False, indent=2, default=str))


def cmd_pull(cfg, law, out):
    m = cfg_get(cfg)
    path = m["paths"][law]
    interval = 60.0 / max(1, m["rpm"])
    session = requests.Session()
    rows, page = [], m["start_page"]
    while True:
        params = build_params(m, law, cfg)
        params[m["params"]["page"]] = page
        t0 = time.time()
        items = fetch_page(session, m, path, params)
        if not items:
            break
        rows.extend(record_to_row(x) for x in items)
        print(f"  стр.{page}: +{len(items)}  всего {len(rows)}")
        page += 1
        if len(items) < m["params"]["size_value"]:
            break
        time.sleep(max(0, interval - (time.time() - t0)))   # соблюдаем rpm

    df = pd.DataFrame(rows, columns=COLUMNS)
    df["fz"] = law
    df["publish_date"] = pd.to_datetime(df["publish_date"], errors="coerce").dt.date
    dest = Path(out)
    dest.parent.mkdir(parents=True, exist_ok=True)
    if dest.suffix.lower() in (".parquet", ".pq"):
        df.to_parquet(dest, index=False)
    else:
        df.to_csv(dest, index=False, encoding="utf-8-sig")
    print(f"\nГотово: {len(df)} записей ({law}-ФЗ) -> {dest}")


def main():
    ap = argparse.ArgumentParser(description="ГосПлан API loader")
    ap.add_argument("cmd", choices=["probe", "pull"])
    ap.add_argument("--config", default="config.yaml")
    ap.add_argument("--law", default="223", choices=["44", "223"])
    ap.add_argument("--out", default="../data/normalized.parquet")
    args = ap.parse_args()
    cfg = yaml.safe_load(Path(args.config).read_text(encoding="utf-8")) or {}
    if args.cmd == "probe":
        cmd_probe(cfg, args.law)
    else:
        cmd_pull(cfg, args.law, args.out)


if __name__ == "__main__":
    main()
