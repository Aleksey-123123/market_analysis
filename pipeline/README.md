# Пайплайн анализа госзакупок (поиск «узких мест» и скоринг ниш)

Цель: из массива закупок ЕИС получить ранжированный список ниш (связок
**ОКПД2 × регион [× заказчик]**), где сходятся **рост суммы × низкая
конкуренция × дисфункция рынка** (несостоявшиеся, отменённые, концентрация у
1-3 поставщиков, повторяемость). Это те самые «узкие места» для входа.

## Архитектура (слои)

```
[ загрузчик ] → data/normalized.(parquet|csv) → [ analyze.py ] → out/niches.xlsx
```

- **Движок `analyze.py`** — источник-независим: работает с нормализованной
  таблицей (см. `schema.py`). Протестирован на синтетике, готов к работе.
- **Загрузчики** приводят любой источник к этой схеме:
  - `loaders/kontur_csv.py` — **рекомендуемый быстрый путь**: экспорт из
    Контур.Закупки (Excel/CSV) → нормализация. Русские заголовки колонок
    сопоставляются толерантно.
  - `loaders/zakupki_ftp.py` — бесплатный бэкап: официальные открытые данные
    ЕИС (FTP), полный массив, но качать/парсить XML самому.

## Быстрый путь через Контур.Закупки (рекомендуется)

1. В `zakupki.kontur.ru` собрать поиск по гипотезе: закон (44/223 — **делать
   отдельные выгрузки**), нужные ОКПД2, регионы, период 24 мес.
   **Не отфильтровывать несостоявшиеся/отменённые — они нужны для сигналов.**
2. Экспортировать результаты в Excel (кнопка выгрузки в списке закупок).
3. Нормализовать и проанализировать:
   ```bash
   python loaders/kontur_csv.py --input export.xlsx --out data/normalized.parquet
   python analyze.py --input data/normalized.parquet --out out/niches.xlsx --group okpd2,region
   ```
   Загрузчик печатает, какие колонки распознал; если что-то «НЕ НАЙДЕНО» —
   пришлите строку заголовков, дополню сопоставление.

## Быстрый старт (демо на синтетике, без интернета)

```bash
pip install -r requirements.txt
python make_sample.py --rows 4000 --out data/normalized.parquet
python analyze.py --input data/normalized.parquet --out out/niches.xlsx --group okpd2,region
```

## Путь через ГосПлан API (автоматическая выгрузка, без ручных экспортов)

Убирает ручную выгрузку и лимит 2000: скрипт сам ходит по REST постранично.
Стартуем на **бесплатном тестовом сервере** `v2test.gosplan.info` (без ключа,
~10 запросов/мин); прод `v2.gosplan.info` (30к/год) включается позже.

```bash
pip install -r requirements.txt
# 1) настроить фильтр в config.yaml -> секция gosplan.query (date_from, okpd2, region)
# 2) посмотреть реальные поля ответа (важно: подтвердить схему):
python loaders/gosplan_api.py probe --config config.yaml --law 223
# 3) выкачать (отдельно 223 и 44) и проанализировать:
python loaders/gosplan_api.py pull --config config.yaml --law 223 --out data/n223.parquet
python loaders/gosplan_api.py pull --config config.yaml --law 44  --out data/n44.parquet
python analyze.py --input data/n223.parquet --out out/niches223.xlsx --group okpd2,region
```

> Пути эндпоинтов и имена параметров ГосПлана заданы в `DEFAULTS`/`config.yaml`
> как обоснованное предположение. **Сначала запустите `probe`** — он печатает
> плоские ключи первой записи и текущий маппинг; если что-то мапится не так,
> поправьте `gosplan.paths` / `gosplan.params` в конфиге (или пришлите вывод).

## Боевой запуск на данных ЕИС

> Запускать там, где есть доступ к `zakupki.gov.ru` (ваш ПК / VPS в РФ).
> В песочнице Claude сеть закрыта политикой — сайт недоступен.

1. Открыть FTP `ftp://ftp.zakupki.gov.ru/` (логин `free` / пароль `free`),
   найти актуальные каталоги по нужным регионам и 44/223-ФЗ.
2. Прописать пути и период в `config.yaml`.
3. Выгрузить и распарсить:
   ```bash
   python loaders/zakupki_ftp.py --config config.yaml --out data/normalized.parquet
   ```
4. Проверить извлечение полей на 2-3 файлах: теги ЕИС различаются по способам
   закупки и версиям — при необходимости дополнить списки-кандидаты в
   `FIELD_CANDIDATES` (`loaders/zakupki_ftp.py`).
5. Прогнать анализ:
   ```bash
   python analyze.py --input data/normalized.parquet --out out/niches.xlsx --group okpd2,region
   ```

## Что на выходе

Таблица связок с колонками: `lots_count`, `total_sum`, `avg_check`,
`growth` (рост суммы recent12m/prior12m), `no_competition_share` (доля лотов
с ≤1 участником), `failed_share`, `canceled_share`, `dysfunction`,
`winner_hhi` (концентрация поставщиков), `repeat_customers`, `repeatability`,
`smp_share`, **`opportunity_score`** (композит) и `barrier_manual` (порог
входа/CAPEX — заполняется руками, из данных не выводится).

Веса сигналов настраиваются в `WEIGHTS` (`analyze.py`).

## Почему Python+DuckDB/pandas, а не BaseX

BaseX + XQuery (см. статью в родительском обсуждении) отлично подходит для
**ручного исследования** пёстрого XML ЕИС. Для **повторяемого пайплайна под
ИИ-анализ** проще потоковый парсер + табличный движок: легче автоматизировать,
версионировать в Git и скармливать результат LLM. Оба подхода читают одни и те
же архивы ЕИС — выбор про удобство, а не про доступ к данным.

## Файлы

| Файл | Назначение |
| :-- | :-- |
| `schema.py` | нормализованная схема (контракт между слоями) |
| `analyze.py` | движок: метрики узких мест + скоринг ниш |
| `make_sample.py` | генератор синтетики для теста движка |
| `loaders/gosplan_api.py` | автовыгрузка через ГосПлан API (probe/pull) |
| `loaders/kontur_csv.py` | нормализация экспорта Контур.Закупки (Excel/CSV) |
| `loaders/zakupki_ftp.py` | выгрузка+парсинг открытых данных ЕИС (FTP) |
| `config.yaml` | регионы, каталоги ЕИС, период, целевые ОКПД2 |
