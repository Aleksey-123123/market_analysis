"""Анализ итогов закупок из рабочей таблицы (колонки K — контракт, L — итог).

Запуск: python analyze_contracts.py путь/к/таблице.xlsx [out.xlsx]
"""
import re
import sys

import pandas as pd

COLS = ['date', 'tag', 'name', 'cust', 'price', 'adv', 'deadline', 'term',
        'gov', 'kontur', 'K', 'L']


def norm_status(k):
    s = str(k).strip().lower()
    if s == 'да':
        return 'заключен'
    if s == 'нет':
        return 'не состоялась'
    if s.startswith('отмен'):
        return 'отменена'
    if s == 'неизвестно':
        return 'неизвестно'
    return 'не заполнено'


PCT_RE = re.compile(r'(-?\s*\d+(?:[.,]\d+)?)\s*%')


def parse_L(v, status):
    """-> (снижение в %, победитель, причина срыва)"""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None, None, None
    if isinstance(v, (int, float)):
        x = float(v)
        return abs(x * 100 if abs(x) <= 1 else x), None, None
    s = str(v).strip()
    low = s.lower()
    pct = None
    m = PCT_RE.search(s)
    if m:
        pct = abs(float(m.group(1).replace(',', '.').replace(' ', '')))
    elif 'без снижения' in low:
        pct = 0.0
    elif re.match(r'^-\d+,!%', s):  # опечатка '-37,!%'
        pct = float(s[1:s.index(',')])
    if status == 'заключен':
        rest = PCT_RE.sub('', s)
        rest = re.sub(r'снижение|\(без снижения\)|без снижения|иск', '', rest,
                      flags=re.I).strip(' \n-.,()')
        winner = rest or None
        if winner and (winner.lower().startswith('неизвестно') or winner.upper() == 'НДС'
                       or '%' in winner or '!' in winner):
            winner = None
        return pct, winner, None
    # не состоялась
    reason = None
    if 'нет допущ' in low or 'ни одной' in low:
        reason = 'нет заявок / никого не допустили'
    elif any(w in low for w in ('одна заявка', 'один участник', 'один допущ',
                                'одна допущ', 'одну заявку', 'только одна',
                                'только один', 'допущена только', 'допущен только',
                                'допущенн только')):
        reason = 'одна заявка / один допущенный'
    winner = None
    m = re.search(r'\(([^)]+)\)', s)
    if m and 'снизил' not in m.group(1):
        winner = m.group(1)
    elif '\n' in s:
        winner = s.split('\n', 1)[1].strip() or None
    if m and 'снизил' in m.group(1):  # 'один допущен (три кита снизил на 8,5%)'
        winner = m.group(1).split('снизил')[0].strip()
    return pct, winner, reason


def parse_price(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v)
    if 'м3' in s:
        return None
    s = re.sub(r'[^\d,.]', '', s).replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


REGIONS = [
    ('Москва и МО', r'моск|моспроект|басманн|таганск|кубинк|видное|\bистр|лобн|власих|павлово-посад|рузск|шаховск|дубна|загорск|новогирее|мосводоканал|ТЭЦ № 22|сергиево'),
    ('Республика Алтай', r'республик\w* алтай|горно-алтай|усть-кокс|манжерок|белуха|чойск|каракокш'),
    ('Алтайский край', r'алтайск|барнаул|бийск|белокурих|поспелих|волчих|солонеш|быканов'),
    ('Кемеровская обл. (Кузбасс)', r'кузбас|кемеров|анжеро|новокузнец|таштагол|топки|юрг|ленинск-кузнец|евраз|колыван|кузнецк'),
    ('Томская обл.', r'томск'),
    ('Иркутская обл. / Бурятия / Забайкалье', r'иркут|бурят|байкал|нижнеудинск|култук|забайкал|вилюй|таксимо|тулун'),
    ('Красноярский край / Хакасия / Тыва', r'краснояр|хакас|крамз|норильск|хам-сыр|тыва'),
    ('Новосибирская обл.', r'новосибирс|бердск|кольцово|искитим|ордынск|татарск|черепанов|кочков|кыштов|чанов|болотн|каргат|доволен|сузун|толмачев|криводанов|бибиха|станционно|здвинск|усть-таркск|куйбышев|ГорМост|гормост|ЦОДД|СО РАН|сибирское отделение|НГУ|СГУТИ|РЭС|электрические сети|брусника|сибагро|стрелочн|сибгазмаш|дубльгис|тепличный комбинат|метрополитен|верх-мильтюш|морского сельсовета|кубовин|промышленно-логистическ|ПЛП'),
    ('Омская обл.', r'омск'),
    ('Прочие регионы', r'усть-балык|смоленщин|краснодар|нижегород|РНЦХ|хирургии|курск'),
]


def region_of(row):
    # сначала по заказчику, затем по тексту закупки
    for text in (str(row['cust']), str(row['name'])):
        for reg, pat in REGIONS:
            if re.search(pat, text, flags=re.I):
                return reg
    return 'не определен'


DIRECTIONS = [
    ('Мосты и ИССО', r'мост|путепровод|подпорн|искусственн\w* сооруж|водопропускн|откос|склон|берегоукреп|береговой|лестни|пешеходн\w* переход|тоннел|пролетн'),
    ('Дороги', r'дорог|а/д|дорожк|тротуар|аэродром|асфальт|коллейн|остановок|автобусных остановок|покрыти'),
    ('Барьеры / БДД / ограждения', r'барьерн|ограждени|указател|дорожных знак|светофор|безопасности дорожного'),
    ('Электросети и освещение', r'ЛЭП|освещ|КЛ-|ВЛ |ТП-|ПС |подстанц|электр|КРУН|РП-|кабельн|трасс\w* ВЛ'),
    ('Вода, тепло, газ, канализация', r'водо|канализ|тепло|теплотрасс|газ|КНС|ГВС|очистк\w* подземных'),
    ('Благоустройство и озеленение', r'благоустр|озелен|газон|деревьев|насажден|растени|лесовосст|каток|катка|ледов|фонтан|рубк|сквер|парк'),
    ('Здания, монолит, фундаменты', r'здани|монолит|фундамент|кровл|фасад|капитальн\w* ремонт общежит|корпус|строительств\w* объекта|свайн|шпунт|резервуар|укрыти|блок-контейнер|модульн'),
    ('Поставка товаров / IT', r'поставк|ноутбук|компьютер|сервер|панел|приобретение'),
]


def direction_of(row):
    text = str(row['name'])
    for d, pat in DIRECTIONS:
        if re.search(pat, text, flags=re.I):
            return d
    return 'Прочее'


CUSTOMER_ALIASES = [
    (r'управление автомобильных дорог «сибирь»|^ФУАД$', 'ФКУ Упрдор «Сибирь» (ФУАД)'),
    (r'красноярск\s*-\s*иркутск', 'ФКУ Упрдор Красноярск–Иркутск'),
    (r'дирекция автомобильных дорог кузбасса', 'ГКУ Дирекция автодорог Кузбасса'),
    (r'территориальное управление автомобильных дорог новосибирск', 'ГКУ НСО ТУАД'),
    (r'управление автомобильных дорог алтайского края', 'КГКУ Упрдор Алтайского края'),
    (r'горно-алтайавтодор', 'КУ РА Горно-Алтайавтодор'),
    (r'управление автомобильных дорог томской', 'ОГКУ Упрдор Томской обл.'),
    (r'гормост', 'МКУ «Гормост» (Новосибирск)'),
    (r'региональные электрические сети', 'АО «РЭС» (Новосибирск)'),
    (r'россети', 'Россети (разные филиалы)'),
    (r'автомобильные дороги .*административного округа', 'ГБУ «Автомобильные дороги» (Москва)'),
    (r'дубльгис', 'ООО «ДубльГИС»'),
    (r'манжерок', 'ООО «Всесезонный курорт «Манжерок»'),
    (r'РНЦХ|хирургии', 'РНЦХ им. Петровского'),
]


def customer_of(c):
    c = re.sub(r'\s+', ' ', str(c)).strip()
    for pat, name in CUSTOMER_ALIASES:
        if re.search(pat, c, flags=re.I):
            return name
    return c


def law_of(row):
    url = f"{row['gov']} {row['kontur']}"
    m = re.search(r'(?:regNumber=|searchString=|kontur\.ru/)(\d+)', url)
    if m:
        return '44-ФЗ' if len(m.group(1)) == 19 else '223-ФЗ / коммерческие'
    if 'rts-tender.ru/auctionsearch' in url:
        return '44-ФЗ'
    return '223-ФЗ / коммерческие'


def clean_winner(w):
    if not w:
        return None
    w = re.sub(r'^(ООО|АО|ИП|ПК|СК)\s+', '', w.strip(), flags=re.I)
    w = w.replace('«', '').replace('»', '').replace('"', '').strip()
    return w.title() if w.isupper() or w.islower() else w


def load(path):
    df = pd.read_excel(path, header=0).iloc[:, :12]
    df.columns = COLS
    df = df.dropna(subset=['name'])
    df['status'] = df.K.map(norm_status)
    parsed = [parse_L(v, s) for v, s in zip(df.L, df.status)]
    df['drop_pct'] = [p[0] for p in parsed]
    df['winner'] = [clean_winner(p[1]) for p in parsed]
    df['fail_reason'] = [p[2] for p in parsed]
    df.loc[(df.status == 'не состоялась') & df.fail_reason.isna(), 'fail_reason'] = 'другое'
    df['nmck'] = df.price.map(parse_price)
    df['region'] = df.apply(region_of, axis=1)
    df['direction'] = df.apply(direction_of, axis=1)
    df['customer'] = df.cust.map(customer_of)
    df['law'] = df.apply(law_of, axis=1)
    df['size'] = pd.cut(df.nmck, [0, 3e6, 15e6, 60e6, 300e6, 1e12],
                        labels=['до 3 млн', '3–15 млн', '15–60 млн', '60–300 млн', '300 млн+'])
    return df


def summary(df, by):
    known = df[df.status.isin(['заключен', 'не состоялась'])]
    g = df.groupby(by, observed=True)
    out = pd.DataFrame({
        'всего закупок': g.size(),
        'итог известен': known.groupby(by, observed=True).size(),
        'заключено': df[df.status == 'заключен'].groupby(by, observed=True).size(),
        'не состоялось': df[df.status == 'не состоялась'].groupby(by, observed=True).size(),
        'из них одна заявка': df[df.fail_reason == 'одна заявка / один допущенный'].groupby(by, observed=True).size(),
        'из них нет допущенных': df[df.fail_reason == 'нет заявок / никого не допустили'].groupby(by, observed=True).size(),
        'медиана снижения, %': df[df.status == 'заключен'].groupby(by, observed=True).drop_pct.median(),
        'среднее снижение, %': df[df.status == 'заключен'].groupby(by, observed=True).drop_pct.mean(),
        'НМЦК всего, млн': g.nmck.sum() / 1e6,
    }).fillna({'итог известен': 0, 'заключено': 0, 'не состоялось': 0,
               'из них одна заявка': 0, 'из них нет допущенных': 0})
    out['доля несостоявшихся, %'] = (100 * out['не состоялось'] / out['итог известен']).where(out['итог известен'] > 0)
    return out.round(1).sort_values('всего закупок', ascending=False)


def main():
    src = sys.argv[1]
    dst = sys.argv[2] if len(sys.argv) > 2 else 'contracts_analysis.xlsx'
    df = load(src)
    with pd.ExcelWriter(dst) as xw:
        cols = ['date', 'name', 'customer', 'region', 'direction', 'law', 'nmck', 'size',
                'status', 'drop_pct', 'winner', 'fail_reason', 'K', 'L', 'gov']
        df[cols].to_excel(xw, sheet_name='данные', index=False)
        summary(df, 'status').to_excel(xw, sheet_name='итоги')
        summary(df, 'direction').to_excel(xw, sheet_name='по направлениям')
        summary(df, 'region').to_excel(xw, sheet_name='по регионам')
        summary(df, 'customer').to_excel(xw, sheet_name='по заказчикам')
        summary(df, 'law').to_excel(xw, sheet_name='по закону')
        summary(df, 'size').to_excel(xw, sheet_name='по размеру НМЦК')
        w = df[df.winner.notna()]
        w.groupby('winner').agg(
            побед=('status', lambda s: (s == 'заключен').sum()),
            единственный_участник=('status', lambda s: (s == 'не состоялась').sum()),
            снижение_медиана=('drop_pct', 'median'),
            регионы=('region', lambda s: ', '.join(sorted(set(s)))),
            направления=('direction', lambda s: ', '.join(sorted(set(s)))),
            НМЦК_млн=('nmck', lambda s: round(s.sum() / 1e6, 1)),
        ).sort_values(['побед', 'единственный_участник'], ascending=False).to_excel(xw, sheet_name='победители')
        df[df.status == 'не состоялась'][['name', 'customer', 'region', 'direction', 'nmck', 'fail_reason', 'L', 'gov']] \
            .sort_values('nmck', ascending=False).to_excel(xw, sheet_name='несостоявшиеся', index=False)
    return df


if __name__ == '__main__':
    main()
