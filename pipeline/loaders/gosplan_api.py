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
    "paths": {"44": "/fz44/purchases",
              "223": "/fz223/purchases"},
    "params": {                          # реальные query-параметры ГосПлана (openapi.json)
        "limit": "limit",                # размер страницы (макс 100)
        "skip": "skip",                  # смещение (макс 1000 -> ~1100 записей на срез!)
        "size_value": 100,
        "date_from": "published_after",
        "date_to": "published_before",
        "okpd2": "classifier",           # массив кодов ОКПД2
        "region": "region",              # массив кодов регионов (1..99)
        "stage": "stage",                # массив стадий ("1".."4")
    },
    "skip_max": 1000,                    # ограничение API на смещение
    "items_path": "",                    # где в JSON список; "" = корень-массив / авто items|data
}

# Точный маппинг под схему ГосПлана /fzNNN/purchases и /fzNNN/contracts.
# "a+b" => ключ должен содержать И "a", И "b". Порядок: специфичное раньше общего.
KEY_CANDIDATES: list[tuple[str, list[str]]] = [
    ("purchase_id",        ["reg_num", "registration_number", "purchase_number"]),
    ("publish_date",       ["published_at", "publish_date"]),
    ("okpd2",              ["okpd2"]),
    ("region",             ["region"]),
    ("nmck",               ["max_price", "contract_price", "price"]),
    ("purchase_method",    ["purchase_type"]),
    # победитель/поставщик и участники есть в /contracts и протоколах (в /purchases их нет):
    ("winner_inn",         ["supplier+inn", "winner+inn", "supplier_inn", "supplier"]),
    ("winner_name",        ["supplier+name", "winner+name"]),
    ("participants_count", ["participants_count", "applications_count", "bids_count", "offers_count"]),
    ("execution_days",     ["execution_term", "delivery_term"]),
    ("customer_inn",       ["customer"]),          # в purchases customer = ИНН заказчика
    ("_stage",             ["stage"]),             # 1..4 стадия
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
    # СМП — из кода способа закупки (напр. purchaseNoticeZKESMBO -> ...SMBO)
    method = str(row.get("purchase_method") or "").lower()
    row["is_smp"] = any(s in method for s in ("smbo", "smsp", "смп", "мсп"))
    # стадия закупки (ГосПлан): 4 -> отменена. Остальные коды уточняем по данным.
    stage = str(row.pop("_stage", "") or "").strip()
    row["is_canceled"] = stage == "4"
    row["is_failed"] = False        # «несостоявшаяся» берём из протоколов/контрактов (доп. слой)
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
    for attempt in range(7):
        try:
            r = session.get(url, params=params, headers=headers, timeout=60)
        except requests.exceptions.RequestException as e:   # обрыв связи/таймаут
            wait = min(60, 2 ** attempt)
            print(f"  сеть ({type(e).__name__}) -> ретрай через {wait}с")
            time.sleep(wait)
            continue
        if r.status_code == 429:
            print("  429 rate limit -> ждём 60с")
            time.sleep(60)
            continue
        if r.status_code >= 500:                            # временная ошибка сервера
            wait = min(60, 2 ** attempt)
            print(f"  {r.status_code} сервер -> ретрай через {wait}с")
            time.sleep(wait)
            continue
        r.raise_for_status()
        data = r.json()
        if m["items_path"]:
            for part in m["items_path"].split("."):
                data = data.get(part, []) if isinstance(data, dict) else data
        return data if isinstance(data, list) else data.get("items", data.get("data", []))
    raise RuntimeError("страница недоступна после ретраев")


def _aslist(v):
    if v in (None, "", []):
        return None
    return v if isinstance(v, list) else [v]


def month_windows(date_from: str, date_to: str):
    """Разбивает [date_from, date_to] на помесячные окна (для нарезки по времени)."""
    import datetime as dt
    start = dt.date.fromisoformat(date_from)
    end = dt.date.fromisoformat(date_to)
    cur = start.replace(day=1)
    out = []
    while cur <= end:
        nxt = cur.replace(year=cur.year + 1, month=1) if cur.month == 12 \
            else cur.replace(month=cur.month + 1)
        out.append((max(cur, start).isoformat(), min(nxt - dt.timedelta(days=1), end).isoformat()))
        cur = nxt
    return out


def build_params(m, cfg, classifier=None, region=None, date_from=None, date_to=None) -> dict:
    """Базовый фильтр (без skip). classifier/region/даты — переопределение среза."""
    p = m["params"]
    src = (cfg.get("gosplan", {}) or {}).get("query", {})
    q = {p["limit"]: p["size_value"]}
    df_ = date_from if date_from is not None else src.get("date_from")
    dt_ = date_to if date_to is not None else src.get("date_to")
    if df_:
        q[p["date_from"]] = df_
    if dt_:
        q[p["date_to"]] = dt_
    ok = _aslist(classifier if classifier is not None else src.get("okpd2"))
    if ok:
        q[p["okpd2"]] = ok
    reg = _aslist(region if region is not None else src.get("region"))
    if reg:
        q[p["region"]] = reg
    st = _aslist(src.get("stage"))
    if st:
        q[p["stage"]] = st
    return q


def cmd_discover(cfg):
    """Ищет реальные эндпоинты: сперва OpenAPI-спеку, потом перебор вариантов пути."""
    import re
    m = cfg_get(cfg)
    base = m["base_url_" + m["use"]].rstrip("/")
    s = requests.Session()

    def get(url, **kw):
        try:
            r = s.get(url, timeout=40, **kw)
            return r
        except Exception as e:  # noqa: BLE001
            print(f"  [err] {url} -> {e}")
            return None

    print("=== 1) Ищу OpenAPI-спеку (в ней перечислены все пути) ===")
    spec_urls = [
        base + "/openapi.json", base + "/swagger/v1/swagger.json",
        base + "/swagger.json", base + "/v3/api-docs", base + "/api-docs",
        base + "/api/v2/openapi.json", base + "/docs/openapi.json",
        "https://swagger.gosplan.info/swagger-config.json",
        "https://swagger.gosplan.info/swagger-config",
    ]
    found_spec = False
    for u in spec_urls:
        r = get(u)
        if r is None:
            continue
        print(f"  {r.status_code}  {u}")
        if r.status_code == 200:
            try:
                j = r.json()
            except Exception:  # noqa: BLE001
                continue
            if isinstance(j, dict) and j.get("paths"):
                found_spec = True
                print("\n  >>> НАЙДЕНА СПЕКА. Доступные пути:")
                for path, methods in j["paths"].items():
                    ms = ",".join(k.upper() for k in methods if k in ("get", "post"))
                    print(f"    {ms:8s} {path}")
            elif isinstance(j, dict) and j.get("urls"):
                print("  swagger-config ссылается на спеки:")
                for it in j["urls"]:
                    print(f"    {it.get('name')}: {it.get('url')}")
        time.sleep(1.5)
    if found_spec:
        print("\nПришлите список путей выше — я пропишу их в конфиг.")
        return

    print("\n=== 2) Спека не найдена. Пробую типовые варианты пути напрямую ===")
    q = build_params(m, cfg)
    q[m["params"]["skip"]] = 0
    variants = [
        "/api/v2/223fz/notifications", "/api/223fz/notifications",
        "/v2/223fz/notifications", "/223fz/notifications",
        "/api/v2/223fz/notification", "/api/v2/notifications",
        "/api/v2/purchases", "/api/v2/223/notifications",
        "/223-fz/notifications", "/api/v2/fz223/notifications",
        "/api/v2/notice", "/api/v2/223fz/notice",
    ]
    for path in variants:
        r = get(base + path, params=q)
        if r is None:
            continue
        body = (r.text or "")[:120].replace("\n", " ")
        print(f"  {r.status_code}  {path}   {body}")
        if r.status_code == 429:
            print("  (429 — упёрлись в лимит 10/мин, подождите минуту и повторите)")
            break
        time.sleep(6.5)   # держим <10 запросов/мин
    print("\nПришлите строки со статусом 200 (или все) — по ним найду рабочий путь.")


def cmd_probe(cfg, law, obj="purchases"):
    m = cfg_get(cfg)
    path = f"/fz{law}/{obj}"
    params = build_params(m, cfg)
    params[m["params"]["limit"]] = 5
    params[m["params"]["skip"]] = 0
    base = m["base_url_" + m["use"]]
    print("GET {}{}\n  params={}\n".format(base, path, params))

    url = base.rstrip("/") + path
    r = requests.get(url, params=params, timeout=60)
    print("HTTP", r.status_code)
    r.raise_for_status()
    raw = r.json()
    print("Тип ответа:", type(raw).__name__,
          ("| ключи-обёртки: " + ", ".join(list(raw.keys())[:10])) if isinstance(raw, dict) else "")
    items = raw if isinstance(raw, list) else raw.get("items", raw.get("data", []))
    print(f"Записей на странице: {len(items)}")
    if items:
        print("\n--- ПЛОСКИЕ КЛЮЧИ первой записи (пришлите мне это) ---")
        for k, v in flatten(items[0]).items():
            print(f"  {k} = {str(v)[:70]}")
        print("\n--- КАК СЕЙЧАС МАПИТСЯ ---")
        print(json.dumps(record_to_row(items[0]), ensure_ascii=False, indent=2, default=str))


def cmd_pull(cfg, law, out, obj="purchases"):
    m = cfg_get(cfg)
    path = f"/fz{law}/{obj}"
    p = m["params"]
    interval = 60.0 / max(1, m["rpm"])
    skip_max = m["skip_max"]
    size = p["size_value"]
    session = requests.Session()
    dest = Path(out)
    dest.parent.mkdir(parents=True, exist_ok=True)

    src = (cfg.get("gosplan", {}) or {}).get("query", {})
    classifiers = _aslist(src.get("okpd2")) or [None]
    regions_cfg = _aslist(src.get("region"))
    auto_region = regions_cfg is None          # регион не задан -> добираем срезы по регионам при упоре в лимит
    regions = regions_cfg or [None]

    by_id: dict = {}
    # Резюме: если файл уже есть — подхватываем собранное, чтобы не качать заново.
    if dest.exists():
        try:
            prev = pd.read_parquet(dest) if dest.suffix.lower() in (".parquet", ".pq") else pd.read_csv(dest)
            for rec in prev.to_dict("records"):
                pid = rec.get("purchase_id")
                if pid:
                    by_id[str(pid)] = {c: rec.get(c) for c in COLUMNS}
            print(f"Резюме: подхватил {len(by_id)} ранее собранных записей из {dest.name}")
        except Exception as e:  # noqa: BLE001
            print(f"(резюме пропущено: {e})")

    def save():
        df = pd.DataFrame(list(by_id.values()), columns=COLUMNS)
        df["fz"] = law
        df["publish_date"] = pd.to_datetime(df["publish_date"], errors="coerce").dt.date
        if dest.suffix.lower() in (".parquet", ".pq"):
            df.to_parquet(dest, index=False)
        else:
            df.to_csv(dest, index=False, encoding="utf-8-sig")
        return len(df)

    def pull_slice(cl, rg, d_from, d_to) -> str:
        """Качает один срез постранично. Возвращает 'done' | 'truncated' | 'error'."""
        skip = 0
        label = f"ОКПД={cl or 'все'} рег={rg if rg is not None else 'все'} {d_from or ''}..{d_to or ''}"
        while True:
            params = build_params(m, cfg, classifier=cl, region=rg, date_from=d_from, date_to=d_to)
            params[p["skip"]] = skip
            t0 = time.time()
            try:
                items = fetch_page(session, m, path, params)
            except Exception as e:                       # noqa: BLE001
                print(f"  ! [{label}] срез прерван ({e}) — сохраняю собранное, иду дальше")
                return "error"
            if not items:
                return "done"
            for x in items:
                row = record_to_row(x)
                if row.get("purchase_id"):
                    by_id[row["purchase_id"]] = row
            skip += size
            print(f"  [{label}] skip={skip-size}: +{len(items)}  всего {len(by_id)}")
            if len(items) < size:
                return "done"
            if skip > skip_max:
                return "truncated"
            time.sleep(max(0, interval - (time.time() - t0)))

    full_from = src.get("date_from")
    full_to = src.get("date_to")
    months = month_windows(full_from, full_to) if (full_from and full_to) else [(full_from, full_to)]

    try:
        for cl in classifiers:
            for rg in regions:
                # 1) пробуем ОКПД целиком за весь период
                status = pull_slice(cl, rg, full_from, full_to)
                save()
                # 2) упёрлись в лимит -> режем по месяцам (полнота по времени, без перекоса)
                if status == "truncated":
                    print(f"  -> ОКПД={cl} рег={rg if rg is not None else 'все'} велик: режу по месяцам")
                    for d_from, d_to in months:
                        mst = pull_slice(cl, rg, d_from, d_to)
                        # 3) месяц всё равно велик и регион не задан -> добираем по регионам
                        if mst == "truncated" and rg is None and auto_region:
                            print(f"     -> {d_from}..{d_to} велик: добираю по регионам 1..99")
                            for r in range(1, 100):
                                pull_slice(cl, r, d_from, d_to)
                        save()
    except KeyboardInterrupt:
        print("\nПрервано пользователем — сохраняю собранное.")
    finally:
        n = save()
        print(f"\nГотово/сохранено: {n} записей ({law}-ФЗ) -> {dest}")


def main():
    ap = argparse.ArgumentParser(description="ГосПлан API loader")
    ap.add_argument("cmd", choices=["discover", "probe", "pull"])
    ap.add_argument("--config", default="config.yaml")
    ap.add_argument("--law", default="223", choices=["44", "223"])
    ap.add_argument("--object", default="purchases", choices=["purchases", "contracts"],
                    help="что качать: извещения (purchases) или контракты с победителем (contracts)")
    ap.add_argument("--out", default="../data/normalized.parquet")
    args = ap.parse_args()
    cfg = yaml.safe_load(Path(args.config).read_text(encoding="utf-8")) or {}
    if args.cmd == "discover":
        cmd_discover(cfg)
    elif args.cmd == "probe":
        cmd_probe(cfg, args.law, args.object)
    else:
        cmd_pull(cfg, args.law, args.out, args.object)


if __name__ == "__main__":
    main()
