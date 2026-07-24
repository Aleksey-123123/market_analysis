"""Загрузчик открытых данных ЕИС (zakupki.gov.ru) -> нормализованная таблица.

Данные ЕИС лежат на FTP ftp://ftp.zakupki.gov.ru/ в виде zip-архивов с XML
(извещения, протоколы, сведения о договорах) по 44-ФЗ и 223-ФЗ.

ВАЖНО про этот скрипт:
  * В песочнице Claude сеть закрыта политикой — zakupki.gov.ru недоступен.
    Запускайте ЭТОТ файл там, где доступ есть (ваш ПК / VPS в РФ).
  * Структура каталогов ЕИС и имена XML-тегов со временем меняются и
    различаются по способам закупки (см. статью в /docs). Поэтому извлечение
    полей сделано ТОЛЕРАНТНЫМ: ищем узлы по подстроке в localName, а не по
    жёсткому пути. Точные пути под ваш срез данных проверьте на 2-3 файлах и
    при необходимости дополните списки-кандидаты ниже.

Зависимости: lxml, pandas, pyarrow, pyyaml.

Пример:
    python loaders/zakupki_ftp.py --config config.yaml --out ../data/normalized.parquet
"""
from __future__ import annotations

import argparse
import io
import zipfile
from datetime import datetime
from ftplib import FTP
from pathlib import Path

import pandas as pd
import yaml
from lxml import etree

import sys
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from schema import COLUMNS  # noqa: E402

FTP_HOST = "ftp.zakupki.gov.ru"
FTP_USER = "free"      # публичный доступ ЕИС (при отказе попробуйте anonymous)
FTP_PASS = "free"

# Списки подстрок-кандидатов для толерантного поиска тегов (нижний регистр).
FIELD_CANDIDATES = {
    "purchase_id":   ["purchasenoticenumber", "purchasenumber", "noticenumber", "registrationnumber"],
    "publish_date":  ["publishdate", "docpublishdate", "createdate", "plandate"],
    "customer_inn":  ["inn"],
    "customer_name": ["fullname", "shortname", "customername", "name"],
    "region":        ["region", "subject"],
    "okpd2":         ["okpd2code", "okpd2", "code"],
    "purchase_method": ["purchasecodename", "placingwayname", "method"],
    "nmck":          ["maxprice", "sum", "price", "contractsum"],
    "participants_count": ["participantquantity", "applicationcount", "bidscount"],
    "winner_inn":    ["winnerinn", "supplierinn"],
    "winner_name":   ["winnername", "suppliername", "organizationname"],
    "execution_days": ["executionterm", "term"],
}


def _localname(tag: str) -> str:
    return tag.rsplit("}", 1)[-1].lower() if isinstance(tag, str) else ""


def _find_first(el: etree._Element, substrings: list[str]) -> str | None:
    """Первый потомок, чей localName содержит любую из подстрок -> его текст."""
    for node in el.iter():
        ln = _localname(node.tag)
        if any(s in ln for s in substrings):
            txt = (node.text or "").strip()
            if txt:
                return txt
    return None


def _to_float(v: str | None) -> float | None:
    if not v:
        return None
    try:
        return float(v.replace("\xa0", "").replace(" ", "").replace(",", "."))
    except ValueError:
        return None


def parse_xml_bytes(data: bytes, fz: str) -> list[dict]:
    """Извлекает нормализованные записи из одного XML. Возвращает список dict."""
    try:
        root = etree.fromstring(data)
    except etree.XMLSyntaxError:
        return []

    # Каждый «документ» ЕИС обычно = один верхнеуровневый элемент body/item.
    # Берём сам root как одну запись (для большинства выгрузок 1 файл = 1 закупка).
    rec = {c: None for c in COLUMNS}
    rec["fz"] = fz
    for field, cands in FIELD_CANDIDATES.items():
        rec[field] = _find_first(root, cands)

    rec["nmck"] = _to_float(rec.get("nmck"))
    pc = rec.get("participants_count")
    rec["participants_count"] = int(pc) if (pc and pc.isdigit()) else None
    ed = rec.get("execution_days")
    rec["execution_days"] = int(ed) if (ed and str(ed).isdigit()) else None

    # флаги по эвристике на весь текст документа
    blob = etree.tostring(root, encoding="unicode").lower()
    rec["is_failed"] = ("несостоя" in blob) or ("failed" in blob)
    rec["is_canceled"] = ("отмен" in blob) or ("cancel" in blob)
    rec["has_advance"] = "аванс" in blob
    rec["is_smp"] = ("смп" in blob) or ("мсп" in blob) or ("subjectsmp" in blob)

    if rec.get("purchase_id"):
        return [rec]
    return []


def iter_archives(ftp: FTP, remote_dir: str):
    """Список .zip в удалённом каталоге."""
    ftp.cwd(remote_dir)
    names = ftp.nlst()
    for name in names:
        if name.lower().endswith(".zip"):
            yield name


def download_and_parse(cfg: dict) -> pd.DataFrame:
    rows: list[dict] = []
    ftp = FTP(FTP_HOST, timeout=cfg.get("ftp_timeout", 60))
    ftp.login(FTP_USER, FTP_PASS)
    try:
        for src in cfg["sources"]:
            fz = str(src["fz"])
            remote_dir = src["remote_dir"]
            limit = src.get("max_archives", 50)
            print(f"[{fz}-ФЗ] {remote_dir}")
            for i, name in enumerate(iter_archives(ftp, remote_dir)):
                if i >= limit:
                    break
                buf = io.BytesIO()
                ftp.retrbinary(f"RETR {name}", buf.write)
                buf.seek(0)
                try:
                    with zipfile.ZipFile(buf) as zf:
                        for member in zf.namelist():
                            if member.lower().endswith(".xml"):
                                rows.extend(parse_xml_bytes(zf.read(member), fz))
                except zipfile.BadZipFile:
                    print(f"  ! битый архив: {name}")
                print(f"  {name}: всего записей {len(rows)}")
    finally:
        ftp.quit()

    df = pd.DataFrame(rows, columns=COLUMNS)
    # постобработка дат/фильтр по периоду
    df["publish_date"] = pd.to_datetime(df["publish_date"], errors="coerce").dt.date
    if cfg.get("date_from"):
        df = df[df["publish_date"] >= datetime.fromisoformat(cfg["date_from"]).date()]
    return df.reset_index(drop=True)


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--config", default="config.yaml")
    ap.add_argument("--out", default="../data/normalized.parquet")
    args = ap.parse_args()
    cfg = yaml.safe_load(Path(args.config).read_text(encoding="utf-8"))
    df = download_and_parse(cfg)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    if out.suffix.lower() in (".parquet", ".pq"):
        df.to_parquet(out, index=False)
    else:
        df.to_csv(out, index=False, encoding="utf-8-sig")
    print(f"\nГотово: {len(df)} записей -> {out}")


if __name__ == "__main__":
    main()
