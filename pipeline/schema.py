"""Нормализованная схема одной закупки.

Любой загрузчик (FTP ЕИС, агрегатор, ручной CSV) должен привести данные
к этому набору колонок. Движок анализа (analyze.py) работает только с ним и
не знает, откуда пришли данные.
"""
from __future__ import annotations

# Порядок колонок нормализованной таблицы (см. раздел 4.3 стратегии).
COLUMNS = [
    "purchase_id",        # уникальный номер закупки/извещения
    "fz",                 # "44" или "223"
    "publish_date",       # дата публикации (YYYY-MM-DD)
    "customer_inn",       # ИНН заказчика
    "customer_name",      # наименование заказчика
    "region",             # регион (код или название)
    "okpd2",              # код ОКПД2 (или КТРУ)
    "purchase_method",    # способ закупки
    "nmck",               # НМЦК / цена договора, руб.
    "participants_count", # число участников (0/1 = фактически без конкуренции)
    "winner_inn",         # ИНН победителя (пусто, если не состоялась)
    "winner_name",        # наименование победителя
    "is_failed",          # закупка признана несостоявшейся (bool)
    "is_canceled",        # закупка отменена/перенесена (bool)
    "has_advance",        # предусмотрен аванс (bool)
    "execution_days",     # срок исполнения, дней
    "is_smp",             # закупка у СМП/МСП (bool)
]

DTYPES = {
    "purchase_id": "string",
    "fz": "string",
    "publish_date": "datetime64[ns]",
    "customer_inn": "string",
    "customer_name": "string",
    "region": "string",
    "okpd2": "string",
    "purchase_method": "string",
    "nmck": "float64",
    "participants_count": "Int64",
    "winner_inn": "string",
    "winner_name": "string",
    "is_failed": "boolean",
    "is_canceled": "boolean",
    "has_advance": "boolean",
    "execution_days": "Int64",
    "is_smp": "boolean",
}
