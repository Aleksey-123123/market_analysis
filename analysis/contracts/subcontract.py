"""Крупные подрядчики (НМЦК от 300 млн) — к кому идти на субподряд.

Запуск: python subcontract.py таблица.csv отчёт.md компании.csv [мин_НМЦК_млн] [выводы.md]
"""
import re
import sys

import pandas as pd

import analyze_contracts as ac

REGION_PRIORITY = {'Москва и МО': 1, 'Новосибирская обл.': 2, 'Алтайский край': 3, 'Республика Алтай': 3}

# Что внутри объекта (по названию) — с чем к подрядчику можно зайти
WORKS = [
    ('мост / путепровод / эстакада', r'мост|путепровод|эстакад|пролетн|переход через|пешеходный переход в районе|искусственн\w* сооружен|ИССО',
     'монолитные ж/б работы (опоры, ригели, плиты), ремонт бетона и защитные покрытия, деформационные швы, '
     'гидроизоляция, сопряжение с насыпью'),
    ('тоннель', r'тоннел|сбойк|притоннельн', 'монолитный ж/б (обделка, порталы), гидроизоляция'),
    ('подпорные стены / оползни / откосы', r'подпорн|оползн|укрепительн|склон|откос|берегоукреп|берегов\w* укреплен|причал|волнолом|шумозащит|акустическ',
     'подпорные стены (монолит, габионы), укрепление откосов, противооползневые сооружения, экраны'),
    ('кап. ремонт / реконструкция / строительство дороги', r'капитальн\w* ремонт\w* автомобильн|реконструкц\w* автомобильн|строительств\w* (и реконструкц\w* )?(участк\w* )?автомобильн|обход|развязк',
     'ИССО в составе объекта: водопропускные трубы, малые мосты, деф. швы, укрепление откосов, подпорные стены'),
    ('ремонт / слои износа / содержание дорог', r'слоев износа|ремонт\w* автомобильн|ремонт участк|содержани',
     'деф. швы и ремонт мостов на участке, замена водопропускных труб, укрепление откосов и конусов'),
    ('защитные конструкции', r'защитн\w* конструкц|пассивн\w* защит|физ\w*\.? ?защит|беспилот|инженерно-технических средств',
     'защитные ограждающие конструкции от БПЛА: монолит/сборный ж/б, фундаменты, металлоконструкции'),
    ('ЦОД', r'ЦОД|обработки данных', 'фундаменты и монолит под здание/оборудование'),
    ('здания / общестрой', r'здани|поликлиник|школ|дошкольн|общежит|жил\w* дом|аэровокзал|аэропорт|театр|котельн|'
     r'административн|образовательн|бизнес-центр|модульн|градирн|монолит|корпус',
     'монолитные работы (фундаменты, каркас), ж/б конструкции'),
    ('очистные / сети', r'очистн|коллектор|канализац|водоснабжен|ливнев',
     'монолитные резервуары и камеры, колодцы, ж/б конструкции сооружений'),
    ('благоустройство', r'благоустр|набережн|парк', 'подпорные стенки, лестницы, монолитные элементы набережных'),
]


ADDR_RE = re.compile(r'(?:^|[\s,])(обл\.|край|респ\.|г\.)\s*([А-ЯЁ][а-яё\-]+)')
ADDR_PRIO = {'Московская': 1, 'Москва': 1, 'Новосибирская': 2, 'Алтайский': 3, 'Алтай': 3}
NAME_PLACE = [
    (1, r'московск\w* област|\bг\. ?москв|город\w* москв'),
    (2, r'новосибирск\w* област|\bг\. ?новосибирск'),
    (3, r'алтайск\w* кра|республик\w* алтай|\bг\. ?барнаул|горно-алтайск'),
]
OTHER_PLACE = r'(\w+ск\w* област|\w+ск\w* кра[йя]|республик\w* (?!казахстан|абхази|монгол|азербайдж|латви)\w+)'


def place_prio(r):
    """Приоритет по месту работ: адрес → название закупки → регион в таблице."""
    m = ADDR_RE.search(str(r.get('addr') or ''))
    if m:
        return ADDR_PRIO.get(m.group(2), 4)
    text = re.sub(r'границ\w*[^,»"]*', '', str(r['name']), flags=re.I)
    for prio, pat in NAME_PLACE:
        if re.search(pat, text, re.I):
            return prio
    if re.search(OTHER_PLACE, text, re.I):
        return 4
    return REGION_PRIORITY.get(r.region, 4)


def place_name(r):
    m = ADDR_RE.search(str(r.get('addr') or ''))
    if m:
        kind = {'обл.': 'обл.', 'край': 'край', 'респ.': 'Респ.', 'г.': 'г.'}[m.group(1)]
        return f"{m.group(2)} {kind}" if kind in ('обл.', 'край') else f"{kind} {m.group(2)}"
    return r.region


def works_of(name):
    return [(w, offer) for w, pat, offer in WORKS if re.search(pat, str(name), re.I)]


def clean_company(w):
    if not isinstance(w, str):
        return None
    w = re.sub(r'\b(Контракты|Контракт)\b', '', w).strip(' ,;')
    if len(w) < 3 or w in ('№',):
        return None
    m = re.match(r'^(ООО|АО|ПАО|ЗАО|ГУП|ИП|НО|Но)\s*(.*)$', w)
    if m:
        w = f"{m.group(1).upper() if m.group(1) != 'Но' else 'НО'} {m.group(2).strip()}"
    return w


def load_big(src, min_mln):
    df = ac.load(src)
    big = df[df.nmck >= min_mln * 1e6].copy()
    big['company'] = big.winner.map(clean_company)
    big['prio'] = big.apply(place_prio, axis=1)
    big['place'] = big.apply(place_name, axis=1)
    big['works'] = big['name'].map(works_of)
    big['role'] = big.apply(lambda r: 'победитель' if r.status == 'заключен' else
                            ('единственный участник' if r.status == 'не состоялась' else 'участник'), axis=1)
    return big


def _short(t, n):
    t = re.sub(r'\s+', ' ', str(t)).strip().replace('|', '/')
    return t if len(t) <= n else t[:n - 1] + '…'


def _link(url, label):
    return f"[{label}]({url})" if isinstance(url, str) and url.startswith('http') else ''


PRIO_NAME = {1: 'Москва и МО', 2: 'НСО', 3: 'Алтай', 4: 'другие регионы'}


WEIGHT = {'мост / путепровод / эстакада': 3, 'тоннель': 3, 'подпорные стены / оползни / откосы': 3,
          'защитные конструкции': 2, 'кап. ремонт / реконструкция / строительство дороги': 2,
          'здания / общестрой': 1, 'ЦОД': 1, 'очистные / сети': 1, 'благоустройство': 1,
          'ремонт / слои износа / содержание дорог': 1}


# Справки из открытых источников (веб-поиск 19.10.2026). Сайты не открывались — проверить перед звонком.
COMPANY_NOTES = {
    'АО «ГК «Екс»': 'Москва, генподрядчик инфраструктуры и соцобъектов по всей РФ; по данным RAEX средняя доля '
                    'субподряда ~40% — прямой кандидат. Сайт по Циклопедии: aoeks.ru '
                    '([RAEX](https://raexpert.ru/releases/2025/sep12b), [Циклопедия](https://cyclowiki.org/wiki/ЕКС_(группа_компаний)))',
    'АО «Новосибирскавтодор»': 'НСО, с 1995 г., строительство/ремонт/содержание дорог и инженерных сооружений, '
                               '15 филиалов в СФО и УФО ([Cbonds](https://cbonds.ru/company/499945/))',
    'АО «Центральное Дсу Алтайского Края»': 'Барнаул; один из учредителей ООО «Барнаульское ДСУ № 4» (25%) — '
                                           'заходить можно сразу в обе компании '
                                           '([Алтапресс](https://altapress.ru/story/torgi-na-remont-dorog-v-barnaule-zavershilis-bez-ekonomii-byudzheta-242300))',
    'ООО «Барнаульское Дсу № 4»': 'Связано с «Центральным ДСУ Алтайского края» (доля 25%) '
                                  '([Алтапресс](https://altapress.ru/story/torgi-na-remont-dorog-v-barnaule-zavershilis-bez-ekonomii-byudzheta-242300))',
    'АО «Вад»': 'Крупный федеральный генподрядчик (с 1994 г., регистрация — Вологда): дороги, подходы к Крымскому мосту, '
                '«Таврида»; под санкциями ЕС ([ДП](https://www.dp.ru/a/2025/12/19/kontrakt-na-proektirovanie))',
    'АО «Ростовавтомост»': 'Мостостроитель (Ростов-на-Дону): мост через Ростовское море, 972,6 млн '
                           '([Нормакс](https://www.normacs.info/news/71109))',
    'АО «Дсио»': 'Дорожная служба Иркутской области (госкомпания). В отчёте за 2025 г. упоминается задолженность '
                 'перед субподрядчиками — условия оплаты оговаривать заранее '
                 '([ЗС Иркутской обл.](https://eparlament.irzs.ru/Files/getFile/58990))',
}


def companies_table(big):
    rows = []
    for comp, g in big[big.company.notna()].groupby('company'):
        works = {}
        for ws in g.works:
            for w, offer in ws:
                works.setdefault(w, offer)
        supply_only = g['name'].str.contains(r'^\s*поставк|поставка программно|котел|котёл', case=False, regex=True).all()
        lighting = g['name'].str.contains(r'освещени', case=False).all()
        rel = 0 if supply_only else (1 if lighting else max(1, sum(WEIGHT.get(w, 0) for w in works)))
        rows.append({
            'Компания': comp,
            'Приоритет': int(g.prio.min()),
            'Релевантность': rel,
            'Регионы': ', '.join(sorted(set(g.place))),
            'Объектов': len(g),
            'НМЦК всего, млн': round(g.nmck.sum() / 1e6),
            'Роль': ', '.join(sorted(set(g.role))),
            'Что строят': '; '.join(works) or '—',
            'С чем зайти': ' / '.join(dict.fromkeys(works.values())) or 'уточнить по документации',
            'Заказчики': '; '.join(sorted(set(g.customer)))[:300],
            'Объекты': ' || '.join(f"{round(r.nmck / 1e6)} млн: {_short(r['name'], 140)} ({r.gov})"
                                   for _, r in g.sort_values('nmck', ascending=False).iterrows()),
            'Справка': COMPANY_NOTES.get(comp, ''),
            'Сайт': '', 'Контакт (кто)': '', 'Телефон': '', 'Email': '', 'Статус контакта': '', 'Комментарий': '',
        })
    t = pd.DataFrame(rows)
    return t.sort_values(['Приоритет', 'Релевантность', 'НМЦК всего, млн'], ascending=[True, False, False])


GC_CUSTOMER = (r'строительн\w* компани|дсу|дорстрой|автодор|мост|метрострой|стройтрансгаз|специализированный застройщик|'
               r'пик-|твэл-строй|гэсстрой|стройтрест|проектный институт')


def gc_channel(df):
    """Закупки, где заказчик — сама строительная компания (генподрядчик ищет субподрядчика). Любой суммы."""
    c = df.cust.astype(str)
    commercial = c.str.contains(r'Общество с ограниченной|Акционерное общество|^ООО|^АО', regex=True)
    is_gc = c.str.contains(GC_CUSTOMER, case=False, regex=True)
    return df[commercial & is_gc].copy()


def gc_section(df):
    g = gc_channel(df)
    if g.empty:
        return []
    out = ['## Генподрядчики, которые сами объявляют закупки на субподряд', '',
           'Самый прямой канал: заказчик здесь — строительная компания, она уже выиграла объект и ищет исполнителей. '
           'Стоит пройти аккредитацию на их площадках и следить за новыми закупками (суммы любые).', '',
           '| Генподрядчик | Закупок | Что закупали | Ссылки |', '|---|---|---|---|']
    g['gc'] = g.cust.map(lambda x: re.sub(r'\s+', ' ', str(x)).strip()
                         .replace('Общество с ограниченной ответственностью', 'ООО')
                         .replace('Акционерное общество', 'АО'))
    for name, x in sorted(g.groupby('gc'), key=lambda t: -t[1].nmck.fillna(0).sum()):
        x = x.sort_values('nmck', ascending=False)
        what = '<br>'.join(f"<small>{'' if pd.isna(r.nmck) else str(round(r.nmck / 1e6)) + ' млн: '}"
                           f"{_short(r['name'], 120)}</small>" for _, r in x.head(4).iterrows())
        if len(x) > 4:
            what += f"<br><small>…ещё {len(x) - 4}</small>"
        links = '<br>'.join(f"{_link(r.kontur, 'Контур')}" for _, r in x.head(4).iterrows())
        out.append(f"| **{_short(name, 90)}** | {len(x)} | {what} | {links} |")
    return out + ['']


def report(big, comp, min_mln, df=None, ideas=''):
    out = [f"# Крупные подрядчики для субподряда (объекты от {min_mln} млн) — {pd.Timestamp.now():%d.%m.%Y}", '',
           f"Источник — таблица тендеров: {len(big)} закупок от {min_mln} млн, у {big.company.notna().sum()} известен "
           f"победитель или единственный участник ({comp.shape[0]} компаний). Приоритет регионов: "
           "1 — Москва и МО, 2 — НСО, 3 — Алтай, 4 — остальные.", '',
           '«Единственный участник» — закупка признана несостоявшейся, но по 44-ФЗ контракт обычно заключают с ним '
           'по начальной цене: это самые «жирные» контракты, где у генподрядчика есть запас на субподряд.', '',
           '⭐ — в объектах компании есть мосты, путепроводы, тоннели, подпорные стены или откосы (наш основной профиль). '
           'Внутри региона компании отсортированы по профилю, затем по сумме.', '',
           'Заготовка для обзвона с пустыми колонками «Сайт / Контакт / Телефон / Email / Статус» — '
           '[companies.csv](companies.csv) (открыть в Google Таблицах).', '']
    if ideas:
        out += [ideas.strip(), '']
    if df is not None:
        out += gc_section(df)
    for p in (1, 2, 3, 4):
        c = comp[comp['Приоритет'] == p]
        if c.empty:
            continue
        out += [f"## {p}. {PRIO_NAME[p]} — {len(c)} компаний", '',
                '| Компания | Объектов / НМЦК, млн | Что строят | С чем зайти | Объекты |', '|---|---|---|---|---|']
        for _, r in c.iterrows():
            g = big[(big.company == r['Компания'])].sort_values('nmck', ascending=False)
            objs = '<br>'.join(
                f"<small>{round(x.nmck / 1e6)} млн, {x.role}: {_short(x['name'], 110)} "
                f"{_link(x.gov, 'ЕИС')} {_link(x.kontur, 'Контур')}</small>" for _, x in g.head(4).iterrows())
            if len(g) > 4:
                objs += f"<br><small>…ещё {len(g) - 4}</small>"
            tag = ' ⭐' if r['Релевантность'] >= 3 else (' <small>(поставка, не СМР)</small>' if r['Релевантность'] == 0 else '')
            note = f"<br><small>ℹ {r['Справка']}</small>" if r['Справка'] else ''
            out.append(f"| **{r['Компания']}**{tag}<br><small>{r['Регионы']}</small>{note} | {r['Объектов']} / {r['НМЦК всего, млн']} | "
                       f"<small>{r['Что строят']}</small> | <small>{r['С чем зайти']}</small> | {objs} |")
        out.append('')
    nw = big[big.company.isna()].sort_values(['prio', 'nmck'], ascending=[True, False])
    out += [f"## Объекты, где подрядчик ещё неизвестен ({len(nw)})", '',
            'Итог ещё не подведён или победитель не внесён в таблицу. Победителя проще всего посмотреть по ссылке на Контур '
            'и вписать в таблицу — при следующем обновлении компания попадёт в список выше.', '',
            '| Приоритет | НМЦК, млн | Статус | Объект | Заказчик | Ссылки |', '|---|---|---|---|---|---|']
    for _, x in nw.iterrows():
        st = 'идёт приём' if x.status == 'не заполнено' and pd.notna(x.deadline) and x.deadline >= pd.Timestamp.now() else x.status
        out.append(f"| {PRIO_NAME[x.prio]} | {round(x.nmck / 1e6)} | {st} | <small>{_short(x['name'], 160)}</small> | "
                   f"<small>{_short(x.customer, 60)}</small> | {_link(x.gov, 'ЕИС')} {_link(x.kontur, 'Контур')} |")
    out.append('')
    return '\n'.join(out) + '\n'


if __name__ == '__main__':
    src, md, csv = sys.argv[1:4]
    min_mln = int(sys.argv[4]) if len(sys.argv) > 4 else 300
    ideas = open(sys.argv[5], encoding='utf-8').read() if len(sys.argv) > 5 else ''
    big = load_big(src, min_mln)
    comp = companies_table(big)
    comp.to_csv(csv, index=False, encoding='utf-8-sig')
    with open(md, 'w', encoding='utf-8') as f:
        f.write(report(big, comp, min_mln, ac.load(src), ideas))
    print(len(big), 'объектов,', len(comp), 'компаний')
    print(comp[['Компания', 'Приоритет', 'Объектов', 'НМЦК всего, млн', 'Что строят']].to_string())
