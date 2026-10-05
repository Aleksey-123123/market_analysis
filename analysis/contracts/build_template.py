"""Перенос рабочей таблицы тендеров в новый шаблон с выпадающими списками.

Запуск: python build_template.py source.(csv|xlsx) tenders.xlsx
"""
import re
import sys

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

import analyze_contracts as ac

FONT = 'Arial'
TODAY = pd.Timestamp.today().normalize()

STATUSES = ['идёт приём', 'заключен', 'не состоялась', 'отменена', 'итог не внесён', 'неизвестно']
REASONS = ['одна заявка', 'один допущенный', 'никого не допустили', 'нет заявок', 'другое']
METHODS = ['44-ФЗ аукцион', '44-ФЗ конкурс', '44-ФЗ котировки', '44-ФЗ (способ не указан)',
           '223-ФЗ', 'Коммерческая']
DIRECTIONS = [d for d, _ in ac.DIRECTIONS] + ['Прочее']
REGIONS = [r for r, _ in ac.REGIONS if r != 'Прочие регионы'] + ['Прочие регионы']
OUR_DECISION = ['да', 'нет', 'думаем']
OUR_RESULT = ['победа', 'проигрыш', 'заявку отклонили', 'не подавались']

# (заголовок, ширина, формат, подсказка)
COLUMNS = [
    ('Реестровый номер', 22, '@', 'Номер закупки из ЕИС / площадки. Ключ строки — не дублировать.'),
    ('Дата добавления', 12, 'DD.MM.YYYY', None),
    ('Направление', 24, None, 'Выбрать из списка'),
    ('Регион', 24, None, 'Выбрать из списка'),
    ('Наименование', 60, None, None),
    ('Заказчик', 40, None, None),
    ('Закон / способ', 20, None, 'Выбрать из списка'),
    ('НМЦК, ₽', 16, '#,##0.00', 'Числом, без «₽» и пробелов-разделителей'),
    ('Аванс, %', 9, '0%', 'Числом: 30% или 0,3'),
    ('Окончание подачи', 17, 'DD.MM.YYYY HH:MM', None),
    ('Срок исполнения', 12, 'DD.MM.YYYY', None),
    ('Ссылка ЕИС / площадка', 18, None, None),
    ('Ссылка Контур', 18, None, None),
    ('Повтор закупки №', 22, '@', 'Если это повторное объявление — реестровый номер предыдущей закупки'),
    ('Статус', 16, None, 'Выбрать из списка'),
    ('Причина (если не состоялась)', 20, None, 'Выбрать из списка'),
    ('Число участников', 10, '0', 'Сколько заявок подано'),
    ('Снижение, %', 10, '0.0%', 'Числом: 12,5% (знак минус не нужен)'),
    ('Цена контракта, ₽', 16, '#,##0.00', 'Считается автоматически: НМЦК × (1 − снижение)'),
    ('Победитель / единственный участник', 30, None, None),
    ('Подаёмся?', 10, None, 'Выбрать из списка'),
    ('Наша цена, ₽', 16, '#,##0.00', None),
    ('Наше снижение, %', 10, '0.0%', 'Считается автоматически'),
    ('Наш результат', 16, None, 'Выбрать из списка'),
    ('Комментарий / почему отклонили', 40, None, None),
    ('Адрес', 30, None, None),
    ('Исходная запись (из старой таблицы)', 30, None, 'Текст колонки «причина / сумма контракта» до переноса'),
]
COL = {name: i + 1 for i, (name, *_rest) in enumerate(COLUMNS)}
L = {name: get_column_letter(i) for name, i in COL.items()}

TAG_MAP = {
    'дороги': 'Дороги', 'мост': 'Мосты и ИССО', 'мосты': 'Мосты и ИССО',
    'вода': 'Вода, тепло, газ, канализация', 'газ': 'Вода, тепло, газ, канализация',
    'здания': 'Здания, монолит, фундаменты', 'монолит': 'Здания, монолит, фундаменты',
    'электричество': 'Электросети и освещение', 'свет': 'Электросети и освещение',
    'расчмстка': 'Электросети и освещение', 'благоустройство': 'Благоустройство и озеленение',
    'растения': 'Благоустройство и озеленение', 'барьеры': 'Барьеры / БДД / ограждения',
    'барьер': 'Барьеры / БДД / ограждения', 'защита': 'Барьеры / БДД / ограждения',
    'пиломатериалы': 'Поставка товаров / IT',
}


def reg_number(row):
    for u in (row['gov'], row['kontur']):
        m = re.search(r'(?:regNumber=|searchString=|number/|kontur\.ru/)(\d{11,19})', str(u))
        if m:
            return m.group(1)
    return None


def method(row):
    t = re.sub(r'\s+', ' ', str(row['ptype'] or '')).strip().lower()
    if t.startswith('44'):
        if 'конкурс' in t:
            return '44-ФЗ конкурс'
        if 'котиров' in t:
            return '44-ФЗ котировки'
        if 'аукцион' in t:
            return '44-ФЗ аукцион'
        return '44-ФЗ (способ не указан)'
    if t.startswith('223'):
        return '223-ФЗ'
    if t.startswith('коммерч'):
        return 'Коммерческая'
    url = str(row['gov'])
    if '/ea20/' in url or 'rts-tender.ru/auctionsearch' in url:
        return '44-ФЗ аукцион'
    if '/ok20/' in url:
        return '44-ФЗ конкурс'
    if '/zk20/' in url:
        return '44-ФЗ котировки'
    rn = reg_number(row)
    if rn and len(rn) == 19:
        return '44-ФЗ (способ не указан)'
    if rn and len(rn) == 11:
        return '223-ФЗ'
    return 'Коммерческая' if url.startswith('http') else None


def reason(L_raw):
    s = str(L_raw).lower()
    if 'нет допущ' in s:
        return 'никого не допустили'
    if 'ни одной' in s or 'нет заявок' in s:
        return 'нет заявок'
    if 'допущ' in s and re.search(r'одн|один|только', s):
        return 'один допущенный'
    if re.search(r'одна заявка|один участник|одну заявку|только одна', s):
        return 'одна заявка'
    return 'другое'


def adv_pct(v):
    m = re.search(r'(\d+(?:[.,]\d+)?)\s*%', str(v))
    return float(m.group(1).replace(',', '.')) / 100 if m else None


def to_records(src):
    df = ac.load(src)
    out = []
    for _, r in df.iterrows():
        tag = str(r['tag']).strip().lower()
        status = {'заключен': 'заключен', 'не состоялась': 'не состоялась', 'отменена': 'отменена',
                  'неизвестно': 'неизвестно'}.get(r.status)
        if status is None:
            status = 'идёт приём' if pd.notna(r.deadline) and r.deadline >= TODAY else 'итог не внесён'
        rs = reason(r.L) if status == 'не состоялась' else None
        participants = 1 if rs in ('одна заявка', 'один допущенный') else (0 if rs == 'нет заявок' else None)
        raw = None if pd.isna(r.L) else str(r.L).strip()
        out.append({
            'Реестровый номер': reg_number(r),
            'Дата добавления': pd.to_datetime(r['date'], dayfirst=True, errors='coerce'),
            'Направление': TAG_MAP.get(tag) or r.direction,
            'Регион': None if r.region == 'не определен' else r.region,
            'Наименование': re.sub(r'\s+', ' ', str(r['name'])).strip(),
            'Заказчик': re.sub(r'\s+', ' ', str(r['cust'])).strip() if pd.notna(r['cust']) else None,
            'Закон / способ': method(r),
            'НМЦК, ₽': r.nmck,
            'Аванс, %': adv_pct(r.adv),
            'Окончание подачи': r.deadline,
            'Срок исполнения': pd.to_datetime(r.term, dayfirst=True, errors='coerce'),
            'Ссылка ЕИС / площадка': r.gov if pd.notna(r.gov) else None,
            'Ссылка Контур': r.kontur if pd.notna(r.kontur) else None,
            'Статус': status,
            'Причина (если не состоялась)': rs,
            'Число участников': participants,
            'Снижение, %': None if pd.isna(r.drop_pct) or status != 'заключен' else r.drop_pct / 100,
            'Победитель / единственный участник': r.winner,
            'Адрес': re.sub(r'\s+', ' ', str(r.addr)).strip() if pd.notna(r.addr) else None,
            'Исходная запись (из старой таблицы)': raw,
        })
    rec = pd.DataFrame(out)

    # Дубли: одинаковый реестровый номер — оставляем самую заполненную строку.
    rec['_filled'] = rec.notna().sum(axis=1)
    has_rn = rec['Реестровый номер'].notna()
    dup = rec[has_rn].sort_values('_filled', ascending=False).duplicated('Реестровый номер')
    removed = rec.loc[dup[dup].index]
    rec = rec.drop(index=dup[dup].index)
    rec['Повтор закупки №'] = pd.Series(None, index=rec.index, dtype=object)

    # Повторы: та же закупка (название + НМЦК) с другим номером — ссылка на предыдущую.
    rec['_key'] = rec['Наименование'].str.lower().str[:120] + '|' + rec['НМЦК, ₽'].round(0).astype(str)
    rec = rec.sort_values(['Окончание подачи', 'Дата добавления'], na_position='first')
    prev = {}
    for i, r in rec.iterrows():
        k = r['_key']
        if k in prev and prev[k] != r['Реестровый номер']:
            rec.at[i, 'Повтор закупки №'] = prev[k]
        if isinstance(r['Реестровый номер'], str):
            prev[k] = r['Реестровый номер']
    rec = rec.sort_values(['Дата добавления', 'Окончание подачи'], na_position='last')
    return rec.drop(columns=['_filled', '_key']), removed


def style_header(ws, row, ncols, fill='1F4E78'):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = Font(name=FONT, bold=True, color='FFFFFF', size=10)
        cell.fill = PatternFill('solid', fgColor=fill)
        cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')


def build(src, dst):
    rec, removed = to_records(src)
    wb = Workbook()
    thin = Side(style='thin', color='D9D9D9')

    # --- Справочники ---
    ref = wb.active
    ref.title = 'Справочники'
    lists = [('Статус', STATUSES), ('Причина', REASONS), ('Закон / способ', METHODS),
             ('Направление', DIRECTIONS), ('Регион', REGIONS), ('Подаёмся?', OUR_DECISION),
             ('Наш результат', OUR_RESULT)]
    ranges = {}
    for j, (title, values) in enumerate(lists, start=1):
        ref.cell(row=1, column=j, value=title)
        for i, v in enumerate(values, start=2):
            ref.cell(row=i, column=j, value=v).font = Font(name=FONT, size=10)
        col = get_column_letter(j)
        # запас строк, чтобы новые значения подхватывались списком
        ranges[title] = f"Справочники!${col}$2:${col}$40"
        ref.column_dimensions[col].width = 32
    style_header(ref, 1, len(lists))
    ref.cell(row=42, column=1, value='Новые значения можно дописывать в конец столбца — они появятся в выпадающем списке.') \
        .font = Font(name=FONT, italic=True, size=9, color='595959')

    # --- Тендеры ---
    ws = wb.create_sheet('Тендеры', 0)
    for name, col in COL.items():
        ws.cell(row=1, column=col, value=name)
    style_header(ws, 1, len(COLUMNS))
    ws.row_dimensions[1].height = 45
    # колонки, которые заполняет пользователь по итогам — отдельный цвет шапки
    for name in ['Статус', 'Причина (если не состоялась)', 'Число участников', 'Снижение, %',
                 'Победитель / единственный участник']:
        ws.cell(row=1, column=COL[name]).fill = PatternFill('solid', fgColor='7F6000')
    for name in ['Подаёмся?', 'Наша цена, ₽', 'Наш результат', 'Комментарий / почему отклонили']:
        ws.cell(row=1, column=COL[name]).fill = PatternFill('solid', fgColor='375623')
    for name in ['Цена контракта, ₽', 'Наше снижение, %']:
        ws.cell(row=1, column=COL[name]).fill = PatternFill('solid', fgColor='595959')

    n = len(rec)
    last = n + 1
    max_row = int(__import__("os").environ.get("MAXROW", 1000))  # проверки и формулы растягиваем на будущие строки
    for i, (_, r) in enumerate(rec.iterrows(), start=2):
        for name, col in COL.items():
            v = r.get(name)
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                continue
            if isinstance(v, pd.Timestamp):
                v = v.to_pydatetime()
            ws.cell(row=i, column=col, value=v)
    for i in range(2, max_row + 1):
        h, s_, v = f"{L['НМЦК, ₽']}{i}", f"{L['Снижение, %']}{i}", f"{L['Наша цена, ₽']}{i}"
        ws.cell(row=i, column=COL['Цена контракта, ₽'],
                value=f'=IF(AND(ISNUMBER({h}),ISNUMBER({s_})),{h}*(1-{s_}),"")')
        ws.cell(row=i, column=COL['Наше снижение, %'],
                value=f'=IF(AND(ISNUMBER({h}),ISNUMBER({v}),{h}<>0),1-{v}/{h},"")')

    for name, col in COL.items():
        width, fmt = COLUMNS[col - 1][1], COLUMNS[col - 1][2]
        letter = get_column_letter(col)
        ws.column_dimensions[letter].width = width
        for i in range(2, max_row + 1):
            c = ws.cell(row=i, column=col)
            c.font = Font(name=FONT, size=10)
            c.border = Border(bottom=thin)
            if fmt:
                c.number_format = fmt
            if name in ('Наименование', 'Заказчик'):
                c.alignment = Alignment(wrap_text=True, vertical='top')
            else:
                c.alignment = Alignment(vertical='top')

    dv_map = {'Статус': 'Статус', 'Причина (если не состоялась)': 'Причина',
              'Закон / способ': 'Закон / способ', 'Направление': 'Направление', 'Регион': 'Регион',
              'Подаёмся?': 'Подаёмся?', 'Наш результат': 'Наш результат'}
    for col_name, list_name in dv_map.items():
        dv = DataValidation(type='list', formula1=f"={ranges[list_name]}", allow_blank=True,
                            showErrorMessage=True, errorTitle='Значение не из списка',
                            error='Выберите значение из списка (или добавьте его на лист «Справочники»).')
        dv.add(f"{L[col_name]}2:{L[col_name]}{max_row}")
        ws.add_data_validation(dv)
    for col_name, kind, mn, mx, msg in [
        ('Снижение, %', 'decimal', 0, 1, 'Снижение от 0% до 100%'),
        ('Аванс, %', 'decimal', 0, 1, 'Аванс от 0% до 100%'),
        ('Число участников', 'whole', 0, 1000, 'Целое число'),
        ('НМЦК, ₽', 'decimal', 0, 1e12, 'Числом, без «₽»'),
    ]:
        dv = DataValidation(type=kind, operator='between', formula1=str(mn), formula2=str(mx),
                            allow_blank=True, showErrorMessage=True, error=msg)
        dv.add(f"{L[col_name]}2:{L[col_name]}{max_row}")
        ws.add_data_validation(dv)
    for name, col in COL.items():
        hint = COLUMNS[col - 1][3]
        if hint:
            dv = DataValidation(allow_blank=True, showInputMessage=True, promptTitle=name[:32], prompt=hint[:250])
            dv.add(f"{get_column_letter(col)}2:{get_column_letter(col)}{max_row}")
            ws.add_data_validation(dv)

    # подсветка статусов
    st = L['Статус']
    colors = {'заключен': 'E2EFDA', 'не состоялась': 'FCE4D6', 'идёт приём': 'DDEBF7',
              'отменена': 'EDEDED', 'итог не внесён': 'FFF2CC'}
    for value, color in colors.items():
        ws.conditional_formatting.add(
            f"A2:{get_column_letter(len(COLUMNS))}{max_row}",
            FormulaRule(formula=[f'${st}2="{value}"'], fill=PatternFill('solid', fgColor=color)))
    # проверка: реестровый номер не должен повторяться
    rn = L['Реестровый номер']
    ws.conditional_formatting.add(
        f"{rn}2:{rn}{max_row}",
        FormulaRule(formula=[f'AND({rn}2<>"",COUNTIF(${rn}$2:${rn}${max_row},{rn}2)>1)'],
                    fill=PatternFill('solid', fgColor='FF9999'), font=Font(name=FONT, bold=True, color='9C0006')))
    ws.freeze_panes = 'F2'
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLUMNS))}{last}"

    # --- Сводка (формулы) ---
    sv = wb.create_sheet('Сводка', 1)
    rng = lambda name: f"Тендеры!${L[name]}$2:${L[name]}${max_row}"  # noqa: E731
    sv['A1'] = 'Сводка по таблице «Тендеры» — пересчитывается автоматически'
    sv['A1'].font = Font(name=FONT, bold=True, size=12)
    row = 3
    for i, s in enumerate(STATUSES):
        sv.cell(row=row + i, column=1, value=s)
        sv.cell(row=row + i, column=2, value=f'=COUNTIF({rng("Статус")},A{row + i})')
    sv.cell(row=row - 1, column=1, value='Статус'); sv.cell(row=row - 1, column=2, value='Закупок')
    style_header(sv, row - 1, 2)
    total_row = row + len(STATUSES)
    sv.cell(row=total_row, column=1, value='Всего').font = Font(name=FONT, bold=True)
    sv.cell(row=total_row, column=2, value=f'=SUM(B{row}:B{total_row - 1})').font = Font(name=FONT, bold=True)

    def block(start, title, values, key):
        heads = [title, 'Всего', 'Итог известен', 'Заключено', 'Не состоялось', '% не состоялось',
                 'Среднее снижение', 'Одна заявка / один допущенный', 'Никого не допустили']
        for j, h in enumerate(heads, start=1):
            sv.cell(row=start, column=j, value=h)
        style_header(sv, start, len(heads))
        k, s_, rs, d = rng(key), rng('Статус'), rng('Причина (если не состоялась)'), rng('Снижение, %')
        for i, v in enumerate(values, start=start + 1):
            sv.cell(row=i, column=1, value=v)
            sv.cell(row=i, column=2, value=f'=COUNTIF({k},A{i})')
            sv.cell(row=i, column=3, value=f'=D{i}+E{i}')
            sv.cell(row=i, column=4, value=f'=COUNTIFS({k},A{i},{s_},"заключен")')
            sv.cell(row=i, column=5, value=f'=COUNTIFS({k},A{i},{s_},"не состоялась")')
            sv.cell(row=i, column=6, value=f'=IF(C{i}>0,E{i}/C{i},"")').number_format = '0%'
            sv.cell(row=i, column=7, value=f'=IFERROR(AVERAGEIFS({d},{k},A{i},{s_},"заключен"),"")').number_format = '0.0%'
            sv.cell(row=i, column=8, value=f'=COUNTIFS({k},A{i},{rs},"одна заявка")+COUNTIFS({k},A{i},{rs},"один допущенный")')
            sv.cell(row=i, column=9, value=f'=COUNTIFS({k},A{i},{rs},"никого не допустили")')
        return start + len(values) + 2

    nxt = block(total_row + 2, 'Направление', DIRECTIONS, 'Направление')
    nxt = block(nxt, 'Регион', REGIONS, 'Регион')
    block(nxt, 'Закон / способ', METHODS, 'Закон / способ')
    for r_ in sv.iter_rows(min_row=2):
        for c in r_:
            if c.font.color is None or not c.font.bold:
                c.font = Font(name=FONT, size=10, bold=c.font.bold)
    sv.column_dimensions['A'].width = 38
    for c in 'BCDEFGHI':
        sv.column_dimensions[c].width = 15

    # --- Инструкция ---
    ins = wb.create_sheet('Инструкция', 0)
    ins.column_dimensions['A'].width = 34
    ins.column_dimensions['B'].width = 95
    lines = [
        ('Как вести таблицу', None),
        ('1. Одна строка = одна закупка', 'Ключ — «Реестровый номер». Если номер повторяется, ячейка подсветится красным.'),
        ('2. Новый тендер', 'Добавить строку внизу, «Статус» = «идёт приём». Направление, регион и закон — из списка.'),
        ('3. Тендер разыгран', 'Поменять «Статус» и заполнить жёлто-коричневые колонки: причина / число участников / снижение / победитель.'),
        ('4. Повторное объявление', 'Новая строка с новым номером, в «Повтор закупки №» — номер предыдущей. Старую строку не трогать.'),
        ('5. Наше участие', 'Зелёные колонки: «Подаёмся?», «Наша цена», «Наш результат», комментарий (почему отклонили).'),
        ('6. Числа', 'НМЦК и цены — числом без «₽». Проценты — числом: 12,5% (знак минус не нужен).'),
        ('7. Не трогать', 'Серые колонки «Цена контракта» и «Наше снижение» считаются формулами.'),
        ('8. Списки', 'Значения списков — на листе «Справочники», новые можно дописать в конец столбца.'),
        ('9. Анализ', 'Раз в неделю: Файл → Скачать → CSV (лист «Тендеры») и прислать — анализ пересчитается.'),
        (None, None),
        ('Цвета строк', 'голубой — идёт приём; зелёный — заключен; оранжевый — не состоялась; жёлтый — итог не внесён; серый — отменена.'),
        (None, None),
        ('Перенос старых данных', None),
        ('Направление и регион', 'Где не было метки — определены автоматически по названию и заказчику; стоит проверить. Пустой регион — не определился, заполнить.'),
        ('Итоги', 'Разобраны из колонки «причина / сумма контракта»; исходный текст сохранён в последней колонке.'),
        ('Пустой итог', 'Если срок подачи прошёл, а итога не было — статус «итог не внесён» (нужно заполнить).'),
        ('Дубли', f'Удалено строк-дублей с одинаковым реестровым номером: {len(removed)}. Повторные объявления связаны через «Повтор закупки №».'),
        (None, None),
        ('Пример заполненной строки', None),
    ]
    for i, (a, b) in enumerate(lines, start=1):
        ca, cb = ins.cell(row=i, column=1, value=a), ins.cell(row=i, column=2, value=b)
        ca.font = Font(name=FONT, size=10, bold=b is None and a is not None)
        cb.font = Font(name=FONT, size=10)
        cb.alignment = Alignment(wrap_text=True, vertical='top')
        ca.alignment = Alignment(vertical='top')
        if b is None and a:
            ca.font = Font(name=FONT, size=12, bold=True)
    example = [
        ('Реестровый номер', '0139200000126011914'), ('Направление', 'Мосты и ИССО'),
        ('Регион', 'Кемеровская обл. (Кузбасс)'), ('Закон / способ', '44-ФЗ аукцион'),
        ('НМЦК, ₽', '73 552 607,36'), ('Аванс, %', '30%'), ('Статус', 'заключен'),
        ('Причина (если не состоялась)', '— (пусто, т.к. заключен)'), ('Число участников', '3'),
        ('Снижение, %', '4,5%'), ('Победитель / единственный участник', 'ООО «Мостострой»'),
        ('Подаёмся?', 'да'), ('Наша цена, ₽', '71 000 000'), ('Наш результат', 'проигрыш'),
        ('Комментарий / почему отклонили', 'Проиграли 0,8 п.п.'),
    ]
    start = len(lines) + 1
    for i, (a, b) in enumerate(example):
        ins.cell(row=start + i, column=1, value=a).font = Font(name=FONT, size=10, color='595959')
        ins.cell(row=start + i, column=2, value=b).font = Font(name=FONT, size=10, color='595959')
    ins.cell(row=start + len(example), column=1,
             value='(пример — иллюстрация формата, не реальные данные)').font = Font(name=FONT, size=9, italic=True, color='808080')

    wb.save(dst)
    return rec, removed


if __name__ == '__main__':
    rec, removed = build(sys.argv[1], sys.argv[2])
    print('строк перенесено:', len(rec), '| дублей удалено:', len(removed))
    print(rec['Статус'].value_counts().to_string())
    print('повторов связано:', rec['Повтор закупки №'].notna().sum() if 'Повтор закупки №' in rec else 0)
    print('без номера:', rec['Реестровый номер'].isna().sum(), '| без региона:', rec['Регион'].isna().sum())
