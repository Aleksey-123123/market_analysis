"""Анализ итогов закупок из рабочей таблицы (колонки K — контракт, L — итог).

Запуск: python analyze_contracts.py таблица.(csv|xlsx) отчёт.md [выводы.md] [ДД.ММ.ГГГГ ЧЧ:ММ]
"""
import re
from pathlib import Path
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
    ('Москва и МО', r'моск|моспроект|зарайск|серпухов|балаших|долгопрудн|красногорск|луховиц|люберц|талдом|егорьевск|дмитров|одинцов|басманн|таганск|кубинк|видное|\bистр|лобн|власих|павлово-посад|рузск|шаховск|дубна|загорск|новогирее|мосводоканал|ТЭЦ № 22|сергиево'),
    ('Республика Алтай', r'республик\w* алтай|горно-алтай|усть-кокс|манжерок|белуха|чойск|каракокш'),
    ('Алтайский край', r'алтайск|барнаул|бийск|белокурих|поспелих|волчих|солонеш|быканов'),
    ('Кемеровская обл. (Кузбасс)', r'кузбас|кемеров|анжеро|новокузнец|таштагол|топки|юрг|ленинск-кузнец|евраз|колыван|кузнецк'),
    ('Томская обл.', r'томск'),
    ('Иркутская обл. / Бурятия / Забайкалье', r'иркут|бурят|байкал|нижнеудинск|култук|забайкал|вилюй|таксимо|тулун'),
    ('Красноярский край / Хакасия / Тыва', r'краснояр|балахт|хакас|крамз|норильск|хам-сыр|тыва'),
    ('Новосибирская обл.', r'новосибирс|бердск|кольцово|искитим|ордынск|татарск|черепанов|кочков|кыштов|чанов|болотн|каргат|доволен|сузун|толмачев|криводанов|бибиха|станционно|здвинск|усть-таркск|куйбышев|ГорМост|гормост|ЦОДД|СО РАН|сибирское отделение|НГУ|СГУТИ|РЭС|электрические сети|брусника|сибагро|стрелочн|сибгазмаш|дубльгис|тепличный комбинат|метрополитен|верх-мильтюш|морского сельсовета|кубовин|промышленно-логистическ|ПЛП'),
    ('Омская обл.', r'омск'),
    ('Прочие регионы', r'усть-балык|смоленщин|краснодар|нижегород|РНЦХ|хирургии|курск'),
]


def region_of(row):
    # сначала по заказчику, затем по тексту закупки
    for text in (str(row['cust']), str(row.get('addr') or ''), str(row['name'])):
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
    ('Благоустройство и озеленение', r'благоустр|контейнерн|озелен|газон|деревьев|насажден|растени|лесовосст|каток|катка|ледов|фонтан|рубк|сквер|парк'),
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
    (r'москв\w*.*гормост', 'ГБУ «Гормост» (Москва)'),
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
    t = str(row.get('ptype') or '')
    if t.startswith('44'):
        return '44-ФЗ'
    if t.startswith('223') or t.lower().startswith('коммерч'):
        return '223-ФЗ / коммерческие'
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


HEADERS = {
    'date': '`', 'tag': 'метка', 'name': 'наименование', 'cust': 'заказчик',
    'price': 'цена контракта', 'adv': 'аванс', 'deadline': 'дата и время окончания подачи заявок',
    'term': 'срок исполнения контракта', 'gov': 'закупки гов', 'kontur': 'контур',
    'ptype': 'тип закупки', 'addr': 'адрес', 'K': 'контракт', 'L': 'причина / сумма контракта',
}


def load(path):
    if str(path).lower().endswith('.csv'):
        raw = pd.read_csv(path, dtype=str)
    else:
        sheets = pd.ExcelFile(path).sheet_names
        raw = pd.read_excel(path, header=0, sheet_name='Тендеры' if 'Тендеры' in sheets else 0)
    cols = {str(c).strip().lower(): c for c in raw.columns}
    if 'статус' in cols:
        return load_template(raw, cols)
    df = pd.DataFrame({k: raw[cols[h]] if h in cols else None for k, h in HEADERS.items()})
    df['deadline'] = pd.to_datetime(df.deadline, dayfirst=True, errors='coerce')
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


TEMPLATE_STATUS = {'заключен': 'заключен', 'не состоялась': 'не состоялась', 'отменена': 'отменена',
                   'неизвестно': 'неизвестно', 'идёт приём': 'не заполнено', 'итог не внесён': 'не заполнено'}
TEMPLATE_REASON = {'одна заявка': 'одна заявка / один допущенный', 'один допущенный': 'одна заявка / один допущенный',
                   'никого не допустили': 'нет заявок / никого не допустили',
                   'нет заявок': 'нет заявок / никого не допустили', 'другое': 'другое'}


def _num(v):
    if isinstance(v, str):
        v = v.replace('\xa0', '').replace(' ', '').replace('₽', '').replace('%', '').replace(',', '.')
    return pd.to_numeric(v, errors='coerce')


EXPLICIT_REGIONS = [
    ('Красноярский край / Хакасия / Тыва', r'красноярск\w* кра|хакаси|тыва'),
    ('Иркутская обл. / Бурятия / Забайкалье', r'иркутск\w* област|бурят|забайкальск'),
    ('Новосибирская обл.', r'новосибирск\w* област'),
    ('Алтайский край', r'алтайск\w* кра'),
    ('Республика Алтай', r'республик\w* алтай'),
    ('Кемеровская обл. (Кузбасс)', r'кемеровск\w* област|кузбасс'),
    ('Томская обл.', r'томск\w* област'),
    ('Омская обл.', r'омск\w* област'),
    ('Москва и МО', r'московск\w* област|г\. ?москв|город\w* москв'),
]


def explicit_region(name):
    """Регион, прямо названный в тексте закупки (если ровно один)."""
    text = re.sub(r'границ\w*[^,»"]*', '', str(name), flags=re.I)  # «граница Республики Алтай» — не место работ
    found = {reg for reg, pat in EXPLICIT_REGIONS if re.search(pat, text, re.I)}
    return found.pop() if len(found) == 1 else None


CORRECTIONS = Path(__file__).with_name('corrections.csv')
FIELD_MAP = {'статус': 'K', 'регион': 'region', 'подаёмся?': 'decision'}


def apply_corrections(df):
    """Ручные поправки из corrections.csv (пока их не внесли в саму таблицу)."""
    df['correction'] = None
    if not CORRECTIONS.exists():
        return df
    corr = pd.read_csv(CORRECTIONS, dtype=str)
    for _, c in corr.iterrows():
        col = FIELD_MAP.get(str(c['Поле']).strip().lower())
        mask = df.regnum == str(c['Реестровый номер']).strip()
        if col and mask.any():
            df.loc[mask, col] = c['Значение']
            df.loc[mask, 'correction'] = c['Комментарий']
    return df


def _pct(v):
    """«12,5%» из CSV — уже проценты; 0.125 из Excel — доля."""
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return None
    x = _num(v)
    if pd.isna(x):
        return None
    if isinstance(v, str) and '%' in v:
        return abs(x)
    return abs(x * 100 if abs(x) <= 1 else x)


def load_template(raw, cols):
    """Новый шаблон (лист «Тендеры»): статус, снижение и причина — отдельными колонками."""
    g = lambda name: raw[cols[name]] if name in cols else pd.Series(None, index=raw.index)  # noqa: E731
    df = pd.DataFrame({
        'name': g('наименование'), 'cust': g('заказчик'), 'gov': g('ссылка еис / площадка'),
        'kontur': g('ссылка контур'), 'ptype': g('закон / способ'), 'addr': g('адрес'),
        'deadline': pd.to_datetime(g('окончание подачи'), dayfirst=True, errors='coerce'),
        'tag': g('направление'), 'adv': g('аванс, %'), 'term': g('срок исполнения'),
        'date': g('дата добавления'), 'K': g('статус'), 'L': g('исходная запись (из старой таблицы)'),
        'winner': g('победитель / единственный участник'),
        'participants': g('число участников').map(_num),
        'our_result': g('наш результат'),
        'decision': g('подаёмся?'),
        'regnum': g('реестровый номер').astype(str).str.strip(),
        'region': g('регион'),
    }).dropna(subset=['name'])
    df = apply_corrections(df)
    df['status'] = df.K.map(lambda s: TEMPLATE_STATUS.get(str(s).strip().lower(), 'не заполнено'))
    df['drop_pct'] = g('снижение, %').loc[df.index].map(_pct)
    # снижение внесено, а статус забыли поменять — считаем контракт заключённым
    df['status_inferred'] = (df.status == 'не заполнено') & df.drop_pct.notna()
    df.loc[df.status_inferred, 'status'] = 'заключен'
    df['fail_reason'] = g('причина (если не состоялась)').loc[df.index].map(
        lambda s: TEMPLATE_REASON.get(str(s).strip().lower()))
    df.loc[(df.status == 'не состоялась') & df.fail_reason.isna(), 'fail_reason'] = 'другое'
    df['nmck'] = g('нмцк, ₽').loc[df.index].map(_num)
    region = df.region.copy()
    auto = df.apply(region_of, axis=1)
    # пусто или «Прочие регионы», а по тексту регион определяется точно — берём определённый
    use_auto = region.isna() | ((region == 'Прочие регионы') & ~auto.isin(['не определен', 'Прочие регионы']))
    df['region'] = region.where(~use_auto, auto)
    # регион прямо назван в названии закупки и не совпадает с таблицей — верим названию
    named = df['name'].map(explicit_region)
    wrong = named.notna() & (named != df.region) & df.correction.isna()
    df['region_fixed_from'] = None
    df.loc[wrong, 'region_fixed_from'] = df.loc[wrong, 'region']
    df.loc[wrong, 'region'] = named[wrong]
    direction = df.tag
    df['direction'] = direction.where(direction.notna(), df.apply(direction_of, axis=1))
    df['customer'] = df.cust.map(customer_of)
    law = df.ptype.astype(str)
    df['law'] = law.map(lambda t: '44-ФЗ' if str(t).startswith('44') else '223-ФЗ / коммерческие')
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
        'медиана участников': df.groupby(by, observed=True).participants.median()
        if 'participants' in df else None,
        'НМЦК всего, млн': g.nmck.sum() / 1e6,
    }).fillna({'итог известен': 0, 'заключено': 0, 'не состоялось': 0,
               'из них одна заявка': 0, 'из них нет допущенных': 0})
    out['доля несостоявшихся, %'] = (100 * out['не состоялось'] / out['итог известен']).where(out['итог известен'] > 0)
    return out.round(1).sort_values('всего закупок', ascending=False)


FLAG_RULES = [
    ('> 100 млн', lambda r: (r.nmck or 0) > 100e6),
    ('кап. ремонт', lambda r: bool(re.search(r'капитальн\w* ремонт', re.sub(
        r'не относящ\w* к капитальн\w*', '', str(r['name']), flags=re.I), re.I))),
    ('реконструкция', lambda r: bool(re.search(r'реконструкц', re.sub(
        r'после проведения работ по реконструкции', '', str(r['name']), flags=re.I), re.I))),
    ('проектное СРО', lambda r: bool(re.search(
        r'проектн|рабоч\w* документац|ПИР\b|изыскан|проектирован|разработк\w* (проектн|рабоч|документ)',
        str(r['name']), re.I))),
]


# Не наш профиль: тендер остаётся в списке, но балл снижается.
# (название правила, штраф, регулярное выражение по названию закупки)
OFF_PROFILE = [
    ('дорожное покрытие / земполотно', 25,
     r'восстановлени\w* (дорожного |асфальтобетонного )?покрыти|земляного полотна|деформаций и повреждений дорожного покрытия|'
     r'ямочн|слоев износа|асфальтиров|асфальт\w* покрыти'),
    ('работы по заявкам, объём не гарантирован', 25,
     r'дорожно-мостового хозяйства|текущ\w* ремонт\w* и содержани|неопредел\w* объ[её]м|по заявкам'),
    ('общестрой', 25, r'производственной базы|окон|двер|отделочн|кровл|фасад|текущ\w* ремонт\w* (помещени|здани)'),
    ('зимнее содержание / снег', 25, r'зимн|снег|очистк\w* (автодорог|дорог)'),
    ('видеонаблюдение / сигнализация / светофоры', 25, r'видеонаблюд|сигнализац|светофор|подсистем\w* безопасности|инженерно-технических средств|ИТСО'),
    ('детские площадки / МАФ', 25, r'детск\w* (игров\w* )?площадк|игров\w* (элемент|оборудован|комплекс)|малых архитектурных форм|\bМАФ'),
]
# «Свои» регионы — без штрафа; остальные (Омск, Иркутск, Бурятия, Красноярск, Хакасия, Тюмень, Сургут…) — минус.
CORE_REGIONS = {'Новосибирская обл.', 'Алтайский край', 'Республика Алтай', 'Кемеровская обл. (Кузбасс)',
                'Москва и МО', 'Томская обл.', 'не определен'}
REGION_PENALTY = 8
# Аванс — плюс к баллу: до 30% → +5, от 30% → +10.
ADVANCE_BONUS, ADVANCE_BONUS_HIGH = 5, 10


ROAD = r'автомобильн\w* дорог|автодорог|\bа/д\b|тротуар|дорожн\w* покрыти|проезд|улично-дорожн'
COMPLEX = (r'мост|путепровод|съезд|развязк|эстакад|тоннел|труб|подпорн|искусственн\w* сооружен|ИССО|откос|'
           r'берегоукреп|габион|оползн|склон')


def profile_penalty(r):
    name = str(r['name'])
    hits = [(n, pen) for n, pen, pat in OFF_PROFILE if re.search(pat, name, re.I)]
    # просто ремонт/содержание дорог неинтересен; с мостом, съездом, трубой и т.п. — рассматриваем
    if re.search(ROAD, name, re.I) and not re.search(COMPLEX, name, re.I):
        hits.append(('ремонт дорог без сложных объектов', 25))
    pen = max((p for _, p in hits), default=0)
    notes = [name for name, _ in hits]
    if r.region not in CORE_REGIONS:
        pen += REGION_PENALTY
        notes.append('дальний регион')
    return pen, notes


def _hist(g):
    z = g[g.status == 'заключен']
    return {'n': len(g), 'n_drop': int(z.drop_pct.notna().sum()),
            'n_part': int(g.participants.notna().sum()) if 'participants' in g else 0,
            'fail': (g.status == 'не состоялась').mean() if len(g) else None,
            'drop': z.drop_pct.median() if z.drop_pct.notna().any() else None,
            'part': g.participants.median() if 'participants' in g and g.participants.notna().any() else None}


def open_tenders(df, today):
    """Открытые тендеры: история заказчика/сегмента, оценка конкуренции, флаги ограничений."""
    hist = df[df.status.isin(['заключен', 'не состоялась'])]
    rows = []
    is_open = (df.deadline >= today) & ~df.status.isin(['отменена', 'заключен', 'не состоялась'])
    for _, r in df[is_open].iterrows():
        hc = _hist(hist[hist.customer == r.customer])
        hs = _hist(hist[(hist.direction == r.direction) & (hist.region == r.region)])
        hd = _hist(hist[hist.direction == r.direction])
        # Сглаживание: заказчик и сегмент весят по числу итогов, направление — как априорное (вес 5),
        # чтобы 2–3 итога не давали «100% несостоявшихся».
        def pick(key):
            num = den = 0.0
            wkey = {'fail': 'n', 'drop': 'n_drop', 'part': 'n_part'}[key]
            for h, w in ((hc, hc[wkey]), (hs, hs[wkey]), (hd, 5)):
                v = h[key]
                if w and v is not None and not pd.isna(v):
                    num += w * v
                    den += w
            return num / den if den else None
        fail, drop, part = pick('fail'), pick('drop'), pick('part')
        score = 0.0
        score += 40 * (fail if fail is not None else 0.4)
        score += 40 * (1 - min((drop if drop is not None else 20) / 40, 1))
        score += 20 * (1 - min(((part if part is not None else 3) - 1) / 6, 1))
        flags = [name for name, rule in FLAG_RULES if rule(r)]
        pen, notes = profile_penalty(r)
        adv = _pct(r.adv) or 0
        bonus = ADVANCE_BONUS_HIGH if adv >= 30 else (ADVANCE_BONUS if adv > 0 else 0)
        rows.append({
            'балл': max(round(score - pen + bonus), 0), 'балл по истории': round(score),
            'аванс, %': adv or None,
            'не профиль': ', '.join(notes),
            'достоверность': 'высокая' if hc['n'] >= 3 or hs['n'] >= 8 else ('средняя' if hc['n'] + hs['n'] >= 3 else 'низкая'),
            'флаги': ', '.join(flags),
            'окончание подачи': r.deadline, 'наименование': r['name'], 'заказчик': r.customer,
            'регион': r.region, 'направление': r.direction, 'НМЦК, млн': round((r.nmck or 0) / 1e6, 1) or None,
            'аванс': r.adv,
            'ожид. доля несостоявшихся, %': None if fail is None else round(100 * fail),
            'ожид. снижение, %': None if drop is None else round(drop, 1),
            'ожид. участников': None if part is None else round(part, 1),
            'заказчик: итогов': hc['n'],
            'заказчик: не сост., %': None if hc['fail'] is None else round(100 * hc['fail']),
            'заказчик: мед. снижение, %': hc['drop'],
            'сегмент: итогов': hs['n'],
            'ссылка': r.gov, 'контур': r.kontur,
            'решение': str(r.get('decision') or '').strip().lower() if pd.notna(r.get('decision')) else '',
            'поправка': r.get('correction') if pd.notna(r.get('correction')) else '',
        })
    return pd.DataFrame(rows).sort_values('балл', ascending=False)


def _fmt(v, nd=0, suf=''):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return '—'
    return f"{v:.{nd}f}{suf}".replace('.', ',') if isinstance(v, (int, float)) else str(v)


def _short(text, n=95):
    t = re.sub(r'\s+', ' ', str(text)).strip().replace('|', '/')
    return t if len(t) <= n else t[:n - 1].rstrip() + '…'


REGION_SHORT = {
    'Новосибирская обл.': 'НСО', 'Москва и МО': 'Москва/МО', 'Кемеровская обл. (Кузбасс)': 'Кузбасс',
    'Алтайский край': 'Алт. край', 'Республика Алтай': 'Респ. Алтай', 'Томская обл.': 'Томск',
    'Омская обл.': 'Омск', 'Иркутская обл. / Бурятия / Забайкалье': 'Иркутск/Бурятия',
    'Красноярский край / Хакасия / Тыва': 'Красноярск/Хакасия', 'Прочие регионы': 'прочие',
    'не определен': '?',
}


def _link(url, label):
    return f"[{label}]({url})" if isinstance(url, str) and url.startswith('http') else ''


def _tender_table(o):
    """Полное название мелким шрифтом, ссылки на ЕИС и Контур отдельной колонкой."""
    lines = ['| Балл | До | Тендер / заказчик | Регион | НМЦК, млн | Ожид. уч. / сниж. | Ссылки | Пометки |',
             '|---|---|---|---|---|---|---|---|']
    for _, r in o.iterrows():
        name = re.sub(r'\s+', ' ', str(r['наименование'])).strip().replace('|', '/')
        cust = _short(r['заказчик'], 70)
        tender = f"<small>{name}</small><br><small><i>{cust}</i></small>"
        score = f"**{r['балл']}**"
        has_adv = pd.notna(r['аванс, %']) and r['аванс, %'] > 0
        if has_adv:
            score += f"<br>💰 аванс {_fmt(r['аванс, %'], 0, '%')}"
        if r['достоверность'] != 'высокая':
            score += f"<br><small>{r['достоверность']} достов.</small>"
        if r['балл'] != r['балл по истории'] and not (has_adv and not r['не профиль']):
            score += f"<br><small>по истории {r['балл по истории']}</small>"
        links = ' · '.join(x for x in (_link(r['ссылка'], 'ЕИС'), _link(r['контур'], 'Контур')) if x)
        marks = []
        if r['флаги']:
            marks.append('⚠ ' + r['флаги'])
        if r['не профиль']:
            marks.append('↓ ' + r['не профиль'])
        if r.get('поправка'):
            marks.append('✎ ' + r['поправка'])
        lines.append(
            f"| {score} | {r['окончание подачи']:%d.%m %H:%M} | {tender} | {REGION_SHORT.get(r['регион'], r['регион'])} | "
            f"{_fmt(r['НМЦК, млн'], 1)} | {_fmt(r['ожид. участников'], 1)} / {_fmt(r['ожид. снижение, %'], 0, '%')} | "
            f"{links} | <small>{'<br>'.join(marks)}</small> |")
    return '\n'.join(lines)


def _summary_table(df, by, title, min_known=1):
    sm = summary(df, by)
    sm = sm[sm['итог известен'] >= min_known].sort_values('итог известен', ascending=False)
    lines = [f'| {title} | Итог известен | Не состоялось | Медиана снижения | Медиана участников |',
             '|---|---|---|---|---|']
    for k, r in sm.iterrows():
        lines.append(f"| {_short(k, 60)} | {int(r['итог известен'])} | {_fmt(r['доля несостоявшихся, %'], 0, '%')} | "
                     f"{_fmt(r['медиана снижения, %'], 1, '%')} | {_fmt(r['медиана участников'], 1)} |")
    return '\n'.join(lines)


def report_md(df, now, conclusions='', data_date=None):
    """Простой Markdown-отчёт: выводы, открытые тендеры по группам, статистика."""
    o = open_tenders(df, now)
    known = df[df.status.isin(['заключен', 'не состоялась'])]
    z = known[known.status == 'заключен']
    refused = o[o['решение'] == 'нет']
    o = o[o['решение'] != 'нет']
    top = o[(o['балл'] >= 65) & (o['достоверность'] != 'низкая')]
    mid = o[~o.index.isin(top.index) & (o['балл'] >= 50)]
    low = o[o['балл'] < 50]
    fixed = df[df.get('region_fixed_from', pd.Series(index=df.index, dtype=object)).notna()]
    part = z.groupby(z.participants.clip(upper=6)).drop_pct.agg(['count', 'median'])
    part_rows = '\n'.join(
        f"| {'6 и больше' if k >= 6 else int(k)} | {int(r['count'])} | {_fmt(r['median'], 1, '%')} |"
        for k, r in part.iterrows())
    single = int(((known.participants == 1) & (known.status == 'не состоялась')).sum())
    out = [
        f"# Тендеры: где подаваться — {now:%d.%m.%Y}" + (f" (таблица от {data_date})" if data_date else ''),
        '',
        f"Данные: {len(df)} закупок, итог известен по {len(known)} "
        f"(заключено {len(z)}, не состоялось {len(known) - len(z)}). "
        f"Открытых тендеров (приём заявок не закончился): **{len(o)}**.",
        '',
        '⚠ — ограничения: больше 100 млн, капитальный ремонт, реконструкция или нужна проектная СРО '
        '(определено по названию закупки — проверить по документации). Такие тендеры оставлены в выборке.',
        '',
        '↓ — балл снижен: не наш профиль (ремонт дорог без сложных объектов, дорожное покрытие, работы по заявкам, '
        'общестрой, зимнее содержание, видеонаблюдение / сигнализация / светофоры, детские площадки / МАФ) '
        'или дальний регион. «По истории» — балл до снижения. ✎ — ручная поправка.',
        '',
        '💰 — есть аванс: балл поднят (+5, при авансе от 30% — +10).',
        '',
    ]
    if conclusions:
        out += [conclusions.strip(), '']
    out += [
        f"## 1. Подавать в первую очередь ({len(top)})",
        '',
        'Высокий балл: заказчик/сегмент часто без конкуренции, снижение небольшое.',
        '',
        _tender_table(top) if len(top) else '_нет_',
        '',
        *( [f"### Вы отметили «не подаёмся» ({len(refused)})", '', _tender_table(refused), '']
           if len(refused) else []),
        f"## 2. Можно рассмотреть ({len(mid)})",
        '',
        _tender_table(mid) if len(mid) else '_нет_',
        '',
        f"## 3. Скорее не стоит ({len(low)})",
        '',
        'Много участников и сильный демпинг по истории.',
        '',
        _tender_table(low) if len(low) else '_нет_',
        '',
        *( ['## Регион исправлен по названию закупки', '',
            'В таблице стоит другой регион, а в названии прямо указан этот. В анализе взят регион из названия — '
            'стоит поправить в таблице.', '',
            '| Реестровый номер | В таблице | По названию | Закупка |', '|---|---|---|---|',
            *[f"| {r.regnum} | {r.region_fixed_from} | {r.region} | <small>{_short(r['name'], 120)}</small> |"
              for _, r in fixed.iterrows()], '']
           if len(fixed) else []),
        '## 4. На чём основан балл',
        '',
        'Балл 0–100 = ожидаемая доля несостоявшихся (40) + малое снижение (40) + мало участников (20), '
        'плюс 5 за аванс (плюс 10 при авансе от 30%), '
        'минус 25 за не наш профиль (в т.ч. просто ремонт дорог без моста/съезда/трубы/подпорной стены и работы '
        '«по заявкам» без гарантированного объёма) и минус 8 за регион вне НСО, Алтая, Кузбасса, Москвы/МО, Томска. '
        'Ручные поправки (✎) — в файле `analysis/contracts/corrections.csv`, пока их не внесли в таблицу. '
        'Ожидания берутся из истории заказчика и связки «направление × регион», при малой истории — '
        'подтягиваются к среднему по направлению. «Достоверность» показывает, сколько истории за баллом.',
        '',
        '### Снижение в зависимости от числа участников',
        '',
        '| Участников | Контрактов | Медиана снижения |',
        '|---|---|---|',
        f"| 1 | {single} | закупка «не состоялась», контракт с единственным участником ≈ по НМЦК |",
        part_rows,
        '',
        '### По направлениям',
        '',
        _summary_table(df, 'direction', 'Направление'),
        '',
        '### По регионам',
        '',
        _summary_table(df, 'region', 'Регион'),
        '',
        '### Заказчики (от 4 известных итогов)',
        '',
        _summary_table(df, 'customer', 'Заказчик', min_known=4),
        '',
        '---',
        '_Отчёт собран скриптом `analyze_contracts.py`. Колонки: «Ожид. участников / снижение» — прогноз по истории._',
    ]
    return '\n'.join(out) + '\n'


def main():
    """python analyze_contracts.py таблица.(csv|xlsx) отчёт.md [выводы.md] [ДД.ММ.ГГГГ ЧЧ:ММ]"""
    src, dst = sys.argv[1], sys.argv[2]
    conclusions = open(sys.argv[3], encoding='utf-8').read() if len(sys.argv) > 3 and sys.argv[3] != '-' else ''
    now = pd.to_datetime(sys.argv[4], dayfirst=True) if len(sys.argv) > 4 else pd.Timestamp.now()
    m = re.search(r'(\d{2}\.\d{2}\.\d{4})', src)
    data_date = m.group(1) if m and m.group(1) != f"{now:%d.%m.%Y}" else None
    df = load(src)
    with open(dst, 'w', encoding='utf-8') as f:
        f.write(report_md(df, now, conclusions, data_date))
    return df


if __name__ == '__main__':
    main()
