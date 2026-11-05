#!/usr/bin/env python3
r"""
Пересборка интерактивного отчёта из Eurasian_Area_Reports_*.xlsm.

Раскладка папок:
    2026-3\                         ← папка квартала
        Eurasian_Area_Reports_2026-3.xlsm     ← данные (берётся отсюда)
        Eurasian_Area_Report.html             ← готовый отчёт (появляется здесь)
        Генератор\                  ← build.py, template.html, logo_small.png, logo_light.png, logo_dark.png

Запуск из папки «Генератор»:   python3 build.py
    возьмёт самый свежий .xlsm из папки уровнем выше и положит отчёт туда же.
Можно явно:  python3 build.py "путь\к\файлу.xlsm" "путь\к\отчёту.html"

С паролем:   python3 build.py --password "ваш-пароль"      (или коротко: -pwd "ваш-пароль")
    данные общего отчёта по региону шифруются (AES-256), файл открывается только после ввода пароля.
    Отдельные отчёты кольев, миссий и округов из кнопки «Выгрузить» паролем не защищаются.

Требуется: pip install openpyxl   (для пароля ещё: pip install cryptography)
"""
import sys, re, json, os, glob
import tempfile
import xml.etree.ElementTree as ET
try:
    import openpyxl
except ImportError:
    sys.exit('Не установлен openpyxl. Выполните один раз:\n    python3 -m pip install openpyxl\nи запустите скрипт снова.')

# ---- флаг пароля: вынимаем его из аргументов, остальное разбирается как раньше ----
PASSWORD = None
_rest, _it = [], iter(sys.argv[1:])
for _a in _it:
    if _a in ('--password', '-pwd', '--pwd', '-p'):
        PASSWORD = next(_it, None)
        if not PASSWORD:
            sys.exit('После --password укажите пароль, например:\n    python3 build.py --password "Тихий-Кедр-47"')
    elif _a.startswith('--password='):
        PASSWORD = _a.split('=', 1)[1]
    else:
        _rest.append(_a)
sys.argv = [sys.argv[0]] + _rest
if PASSWORD is not None:
    try:
        from cryptography.hazmat.primitives.ciphers.aead import AESGCM
    except ImportError:
        sys.exit('Для пароля нужна библиотека cryptography. Выполните один раз:\n    python3 -m pip install cryptography\nи запустите скрипт снова.')

HERE = os.path.dirname(os.path.abspath(__file__))        # папка «Генератор»: скрипт, шаблон, логотипы
PARENT = os.path.dirname(HERE)                            # папка квартала: Excel и готовый отчёт


def find_xlsm(folder):
    return [f for f in glob.glob(os.path.join(folder, '*.xls[mx]')) if not os.path.basename(f).startswith('~$')]


NEED_SHEETS = {'Moscow Stake', 'St.Petersburg Stake', 'Moscow Mission', 'S&I Summary'}


def sheet_names(path):
    """Читаем названия листов из OOXML с поддержкой XML entities и namespace."""
    import zipfile
    try:
        with zipfile.ZipFile(path) as z:
            root = ET.fromstring(z.read('xl/workbook.xml'))
        namespace = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
        return {sheet.attrib['name'] for sheet in root.iter(namespace + 'sheet')
                if 'name' in sheet.attrib}
    except (OSError, KeyError, ET.ParseError, zipfile.BadZipFile):
        return set()


def pick_workbook(files):
    """Из файлов Excel в папке выбираем книгу статистики: сначала по названию, потом самую свежую;
    файлы без нужных листов (чужие таблицы) пропускаем и говорим об этом."""
    ours = lambda f: os.path.basename(f).lower().startswith('eurasian_area_report')
    skipped = []
    for f in sorted(files, key=lambda f: (not ours(f), -os.path.getmtime(f))):
        missing = NEED_SHEETS - sheet_names(f)
        if not missing:
            for sk in skipped:
                print(f'Пропускаю «{os.path.basename(sk)}» — это не книга статистики (нет нужных листов).')
            return f
        skipped.append(f)
    return None


if len(sys.argv) > 1:
    SRC = sys.argv[1]
else:
    # без аргумента: книга статистики из папки уровнем выше (а если её там нет — рядом со скриптом, как раньше)
    found = find_xlsm(PARENT) or find_xlsm(HERE)
    SRC = pick_workbook(find_xlsm(PARENT)) or pick_workbook(find_xlsm(HERE))
    if not SRC:
        names = ', '.join(f'«{os.path.basename(f)}»' for f in found) or 'нет ни одного'
        sys.exit(f'Не нашёл книгу статистики. Файлы Excel в папке: {names}.\n'
                 f'В книге должны быть листы: {", ".join(sorted(NEED_SHEETS))}.\n'
                 'Положите Eurasian_Area_Reports_….xlsm в папку квартала (уровнем выше «Генератора») или укажите путь:\n'
                 '    py build.py "..\\Eurasian_Area_Reports_2026-3.xlsm"')
    print(f'Беру файл: {SRC}')
# отчёт кладём туда же, где лежит Excel
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(os.path.dirname(os.path.abspath(SRC)), 'Eurasian_Area_Report.html')

wb = openpyxl.load_workbook(SRC, data_only=True)      # значения
wbf = openpyxl.load_workbook(SRC)                     # формулы — чтобы найти ячейки, где формулу затёрли числом

# Порядок блоков на листах подразделений (10 блоков по 7 кварталов)
KEYS = ['att', 'yw', 'ym', 'tr', 'lur', 'minb', 'mins', 'nc', 'ysa', 'kids']

# ---- Русские названия -------------------------------------------------------
RU = {
    # подразделения верхнего уровня
    'Moscow Stake': 'Московский Российский кол', 'St.Petersburg Stake': 'Санкт-Петербургский Российский кол',
    'Moscow Mission': 'Российская Московская миссия', 'Rostov District': 'Ростовский-на-Дону Российский округ',
    'Saratov District': 'Саратовский Российский округ', 'Yekaterinburg Msn': 'Российская Екатеринбургская миссия',
    'Yekaterinburg District': 'Екатеринбургский Российский округ', 'Novosibirsk Msn': 'Российская Новосибирская миссия',
    'Novosibirsk District': 'Новосибирский Российский округ', 'Armenia Mission': 'Armenia/Georgia Mission',   # русского официального названия нет — по-английски
    'Yerevan District': 'Yerevan Armenia District',
    # Московский кол
    'Kahovskiy': 'Каховский', 'Korolev': 'Королевский', 'Krasnogorskiy': 'Красногорский', 'Minsk': 'Минский',
    'Nizhniy Novgorod': 'Нижний Новгород', 'Ostankino': 'Останкинский', 'Perovo': 'Перовский',
    'Podolskiy': 'Подольский', 'Tretiyakovskiy': 'Третьяковский', 'Voronezh': 'Воронежский',
    'Zelenogradskiy': 'Зеленоградский',
    # СПб кол
    'Avtovo+Pskov': 'Автово', 'Kaliningrad': 'Калининградский', 'Kolpino+Novgorod': 'Колпино',
    'Nevsky': 'Невский', 'Petergof': 'Петергоф', 'Sestroretsk': 'Сестрорецк', 'Petrozavodsk': 'Петрозаводск',
    'Shuvalovsky': 'Шуваловский', 'Tsentralniy': 'Санкт-Петербургский центральный', 'Vyborg': 'Выборгский',
    # Ростовский округ
    'Krasnodar': 'Краснодарский центральный', 'Novorossiysk': 'Новороссийский', 'Rostov Severny': 'Ростовский северный',
    'Rostov Center': 'Ростовский центральный', 'Rostov District Branch': 'Приход Ростовского округа',
    'Sevastopol': 'Севастополь', 'Simferopol': 'Симферопольский', 'Sochi': 'Сочинский центральный', 'Taganrog': 'Таганрогский',
    'Zapadny': 'Западный',
    # Саратовский округ
    'Engels(Marks,Balakovo)': 'Энгельсский', 'Penza': 'Пензенский', 'Samara': 'Самарский',
    'Solnechniy(Dachny)': 'Солнечный', 'Tolyatti': 'Тольяттинский',
    'Volgograd Tsentralny': 'Волгоградский центральный', 'Volzshky(Zavodskoy)': 'Волжский',
    # Екатеринбург
    'Almaty': 'Алматинский', 'Astana (Nursultan)': 'Астана',
    'Yekaterinburg Mission Branch': 'Приход Российской Екатеринбургской миссии',
    'Chelyabinsk+Surgut': 'Челябинский', "Kazan'": 'Казанский', 'Perm + Izhevsk': 'Пермский',
    'Tyumen': 'Тюменский', 'Ufa + Orenburg': 'Уфимский',
    'Yekaterinburg 1st + Kurgan': 'Екатеринбургский первый', 'Yekaterinburg 2nd': 'Екатеринбургский второй',
    # Новосибирск
    'Novosibirsk Mission Branch': 'Приход Российской Новосибирской миссии', 'Ulan-Ude': 'Улан-Удэнский',
    'Vladivostok': 'Владивостокский', 'Yuzhno-Sakhalinsk': 'Южно-Сахалинский',
    'Barnaul Altaisky': 'Барнаульский алтайский', 'Krasnoyarsk': 'Красноярский', 'Novosibirsk 1st': 'Новосибирский первый',
    'Novosibirsk 2nd': 'Новосибирский второй', 'Omsk': 'Омский', 'Tomsk': 'Томский',
    # Армения и Грузия: официального русского названия нет — оставляем по-английски
    'Mission Branches': 'Armenia/Georgia Mission Branch', 'Avlabari': 'Avlabari', 'Batumi': 'Batumi', 'Temka': 'Temka',
    'Ajapnyak': 'Ajapnyak', 'Yerevan Central': 'Yerevan Central', 'Artashat': 'Artashat', 'Gyumri': 'Gyumri',
    'Vanadzor': 'Vanadzor', 'Yerevan 1st (Armenian)': 'Yerevan 1st',
    'Yerevan 2nd (English)': 'Yerevan 2nd',
    'Moscow Msn': 'Приход Российской Московской миссии',
    # короткие названия без «+ второй город»
    'Chelyabinsk': 'Челябинский', 'Perm': 'Пермский', 'Ufa': 'Уфимский', 'Yekaterinburg 1st': 'Екатеринбургский первый',
    'Avtovo': 'Автово', 'Kolpino': 'Колпино', 'Surgut': 'Сургут', 'Izhevsk': 'Ижевск', 'Orenburg': 'Оренбург',
    'Kurgan': 'Курган', 'Pskov': 'Псков', 'Novgorod': 'Новгород', 'Marks': 'Маркс', 'Balakovo': 'Балаково',
    'Moscow Virtual Branch': 'Российский Московский Виртуальный приход',
    'Moscow Virtual': 'Российский Московский Виртуальный приход',
    'Russia Moscow Virtual Branch': 'Российский Московский Виртуальный приход',
    'Virtual Branch': 'Российский Московский Виртуальный приход',
}

# Иерархия: лист -> тип, и какие столбцы на листе миссии являются округами (лист округа)
TOP = [
    ('Moscow Stake', 'кол', {}),
    ('St.Petersburg Stake', 'кол', {}),
    ('Moscow Mission', 'миссия', {'Rostov Dis': 'Rostov District', 'Saratov District': 'Saratov District'}),
    ('Yekaterinburg Msn', 'миссия', {'Yekaterinburg Dis': 'Yekaterinburg District'}),
    ('Novosibirsk Msn', 'миссия', {'Novosibirsk District': 'Novosibirsk District'}),
    ('Armenia Mission', 'миссия', {'Yerevan District': 'Yerevan District'}),
]

# Семинария / институт (лист S&I Summary): столбец начала блока для каждого подразделения
SI_COLS = {
    'Moscow Stake': 'B', 'St.Petersburg Stake': 'G', 'Saratov District': 'L', 'Rostov District': 'Q',
    'Moscow Mission': 'X', 'Yekaterinburg Msn': 'AC', 'Yekaterinburg District': 'AH',
    'Novosibirsk Msn': 'AO', 'Novosibirsk District': 'AT', 'Armenia Mission': 'BA',
    'Yerevan District': 'BF', '__area__': 'BM',
}


def isq(v):
    return isinstance(v, str) and re.fullmatch(r'\d\d-\d', v.strip()) is not None


def num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else 0


# ---- Кварталы берём из файла (последние 7 по листу Московского кола) ---------
def quarter_labels():
    ws = wb['Moscow Stake']
    labels = []
    for r in range(1, 12):
        v = ws.cell(r, 1).value
        if isq(v):
            labels.append(v.strip())
    return labels[-7:]


QS = quarter_labels()
flags = []  # проблемы с данными


def parse_sheet(name):
    ws = wb[name]
    units = [(c, str(ws.cell(1, c).value).strip()) for c in range(2, ws.max_column + 1)
             if isinstance(ws.cell(1, c).value, str) and ws.cell(1, c).value.strip()]
    blocks, r = [], 1
    while r <= ws.max_row:
        if isq(ws.cell(r, 1).value) and not isq(ws.cell(r - 1, 1).value if r > 1 else None):
            rows, rr = [], r
            while rr <= ws.max_row and isq(ws.cell(rr, 1).value):
                rows.append(rr)
                rr += 1
            blocks.append(rows)
            r = rr
        else:
            r += 1
    if len(blocks) != 10:
        raise SystemExit(f'Лист «{name}»: ожидалось 10 блоков показателей, найдено {len(blocks)}')
    out = {}
    for c, u in units:
        d = {}
        for k, rows in zip(KEYS, blocks):
            bylabel = {ws.cell(x, 1).value.strip(): x for x in rows}
            series = []
            for q in QS:
                x = bylabel.get(q)
                if x is None:
                    series.append([0, 0])
                    continue
                a, p, pct = num(ws.cell(x, c).value), num(ws.cell(x, c + 1).value), ws.cell(x, c + 2).value
                series.append([a, p])
                # сломанная формула процента в Excel
                if u != 'Total' and p > 0:
                    real = a / p
                    if isinstance(pct, (int, float)) and abs(pct - real) > 0.005:
                        cell = ws.cell(x, c + 2)
                        raw = wbf[name][cell.coordinate].value
                        hard = not (isinstance(raw, str) and raw.startswith('='))
                        flags.append(dict(kind='formula', sheet=name, unit=u, metric=k, q=q, cell=cell.coordinate, hard=hard,
                                          text=(f'в ячейке {cell.coordinate} вместо формулы вбито число {pct*100:.0f}%' if hard else
                                                f'в ячейке {cell.coordinate} стоит {pct*100:.0f}%')
                                               + f', а по цифрам {a} из {p} = {real*100:.0f}%'))
                    elif isinstance(pct, str) and pct.startswith('#'):
                        flags.append(dict(kind='formula', sheet=name, unit=u, metric=k, q=q,
                                          cell=ws.cell(x, c + 2).coordinate, hard=False,
                                          text=f'в ячейке {ws.cell(x, c + 2).coordinate} ошибка {pct}'))
            d[k] = series
        out[u] = d
    return out


SHEETS = {}
for sheet, _, dists in TOP:
    SHEETS[sheet] = parse_sheet(sheet)
    for dsheet in dists.values():
        SHEETS[dsheet] = parse_sheet(dsheet)


def _base(name):
    # «Engels(Marks,Balakovo)» и «Engels» — одно и то же: сравниваем по части до скобки, без регистра и пробелов
    # «Moscow Virtual» и «Moscow Virtual Branch» тоже одно и то же: хвост Branch / Ward не учитываем
    n = name.split('(')[0].lower().replace('ё', 'е')
    n = re.sub(r'\b(branch|ward)\b', '', n)
    return re.sub(r'[^a-zа-я0-9]', '', n)


RU_BASE = {}
for _k, _v in RU.items():
    RU_BASE.setdefault(_base(_k), _v)
NO_RU = set()


def _noparens(s):
    # на странице только название прихода: без уточнений в скобках и без «+ второй город»
    # «Энгельс (Маркс, Балаково)» → «Энгельс», «Автово + Псков» → «Автово»
    s = re.sub(r'\s*\([^)]*\)', '', s)
    return s.split('+')[0].strip()


def ru(n):
    n = n.strip()
    if n in RU:
        return _noparens(RU[n])
    if _base(n) in RU_BASE:
        return _noparens(RU_BASE[_base(n)])
    NO_RU.add(n)          # нет русского названия — покажем английское и сообщим в терминале
    return _noparens(n)


def leaf(name, sheet, typ='приход'):
    return dict(id=f'{sheet}/{name}', name=ru(name), en=name, type=typ, m=SHEETS[sheet][name], children=[])


def check_total(sheet, node):
    """Сверяем столбец Total на листе с суммой подразделений."""
    tot = SHEETS[sheet].get('Total')
    if not tot:
        return
    for k in KEYS:
        for i, q in enumerate(QS):
            if tot[k][i] != node['m'][k][i]:
                flags.append(dict(kind='total', sheet=sheet, unit='Total', metric=k, q=q,
                                  text=f'итог на листе {tot[k][i][0]} из {tot[k][i][1]}, '
                                       f'а сумма подразделений {node["m"][k][i][0]} из {node["m"][k][i][1]}'))


def aggregate(node):
    if node['children']:
        for ch in node['children']:
            aggregate(ch)
        node['m'] = {k: [[sum(ch['m'][k][i][0] for ch in node['children']),
                          sum(ch['m'][k][i][1] for ch in node['children'])] for i in range(len(QS))]
                     for k in KEYS}
    return node


def si_block(col_letter, row_start, off=1):
    # блок на листе S&I Summary: [потенциал, класс (семинария/институт), воскресные собрания]
    ws = wb['S&I Summary']
    c = openpyxl.utils.column_index_from_string(col_letter)
    res = {}
    r = row_start
    while isq(ws.cell(r, 1).value):          # читаем только подряд идущие строки кварталов этого блока
        v = ws.cell(r, 1).value
        r += 1
        if v.strip() in QS:
            res[v.strip()] = [num(ws.cell(r - 1, c + off).value), num(ws.cell(r - 1, c).value)]  # [факт, потенциал]
    return [res.get(q, [0, 0]) for q in QS]


def find_si_rows():
    ws = wb['S&I Summary']
    yrow = irow = None
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 2).value
        if isinstance(v, str) and v.strip() == 'Youth Potential' and yrow is None:
            yrow = r + 1
        if isinstance(v, str) and v.strip() == 'YSA Potential' and irow is None:
            irow = r + 1
    return yrow, irow


# ---- Строим дерево ----------------------------------------------------------
area = dict(id='area', name='Евразийский регион', en='Eurasian Area', type='регион', children=[])
for sheet, typ, dists in TOP:
    node = dict(id=sheet, name=ru(sheet), en=sheet, type=typ, children=[])
    for col in SHEETS[sheet]:
        if col == 'Total':
            continue
        if col in dists:
            dsheet = dists[col]
            dn = dict(id=dsheet, name=ru(dsheet), en=dsheet, type='округ', children=[])
            for w in SHEETS[dsheet]:
                if w != 'Total':
                    dn['children'].append(leaf(w, dsheet))
            node['children'].append(dn)
        else:
            node['children'].append(leaf(col, sheet, 'приход'))
    area['children'].append(node)
aggregate(area)

# Проверка правил иерархии
for top in area['children']:
    for ch in top['children']:
        if top['type'] == 'кол' and ch['type'] != 'приход':
            raise SystemExit(f'Ошибка иерархии: в коле «{top["name"]}» найден не приход: {ch["name"]}')
        if ch['type'] == 'округ' and any(w['type'] != 'приход' for w in ch['children']):
            raise SystemExit(f'Ошибка иерархии: в округе «{ch["name"]}» есть не только приходы')
    if top['type'] == 'миссия' and not any(c['type'] == 'приход' for c in top['children']):
        print(f'Внимание: в миссии «{top["name"]}» нет прихода миссии')

for sheet, _, dists in TOP:
    n = next(c for c in area['children'] if c['id'] == sheet)
    check_total(sheet, n)
    for dsheet in dists.values():
        check_total(dsheet, next(c for c in n['children'] if c['id'] == dsheet))

# Семинария и институт — только на уровне кольев/округов/миссий
yrow, irow = find_si_rows()


def attach_si(node):
    key = '__area__' if node['id'] == 'area' else node['id']
    if key in SI_COLS:
        node['m']['sem'] = si_block(SI_COLS[key], yrow)
        node['m']['inst'] = si_block(SI_COLS[key], irow)
        node['m']['semS'] = si_block(SI_COLS[key], yrow, 2)     # молодёжь на воскресных собраниях (лист S&I)
        node['m']['instS'] = si_block(SI_COLS[key], irow, 2)    # молодые взрослые на воскресных собраниях (лист S&I)
    for ch in node['children']:
        attach_si(ch)


attach_si(area)

# Пустые подразделения (все нули) — помечаем
def mark_empty(node):
    allzero = all(a == 0 and p == 0 for k in KEYS for a, p in node['m'][k])
    node['empty'] = allzero
    if allzero:
        flags.append(dict(kind='empty', sheet='', unit=node['en'], metric='', q='',
                          text='все показатели нулевые во всех кварталах'))
    for ch in node['children']:
        mark_empty(ch)


mark_empty(area)

# ---- Отметка в Excel о квартале
def report_quarter_label():
    try:
        return str(wb['Eurasian Area']['M2'].value or '')
    except Exception:
        return ''


# привязываем замечания к узлам дерева
IDS = {}


def index(node):
    IDS[node['id']] = node
    IDS.setdefault('en:' + node['en'], node)
    for ch in node['children']:
        index(ch)


index(area)
for f in flags:
    if f['kind'] == 'formula':
        f['node'] = f'{f["sheet"]}/{f["unit"]}' if f'{f["sheet"]}/{f["unit"]}' in IDS else f['sheet']
    elif f['kind'] == 'total':
        f['node'] = f['sheet']
    else:
        f['node'] = IDS['en:' + f['unit']]['id']

# ---- Пояснения (лист «Пояснения»): Подразделение | Квартал | Показатель | Текст | В регион (да — показывать и на странице региона) -----------------
import datetime
METRIC_NAMES = {   # как можно написать показатель в листе пояснений
    'att': ['причастие', 'посещаемость', 'посещаемость причастного собрания'],
    'yw': ['молодые женщины', 'мж'], 'ym': ['молодые мужчины', 'мм'],
    'ysa': ['молодые взрослые', 'одв'], 'kids': ['дети', 'дети 3-11', 'дети 3–11', 'первоначальное общество'],
    'nc': ['новообращённые', 'новообращенные'],
    'tr': ['рекомендация в храм', 'рекомендации', 'храмовая рекомендация'],
    'lur': ['рекомендация молодёжи', 'рекомендация молодежи', 'рекомендация ограниченного пользования'],
    'minb': ['служение, братья', 'служение братья', 'братья'], 'mins': ['служение, сёстры', 'служение сестры', 'сёстры', 'сестры'],
    'sem': ['семинария'], 'inst': ['институт'],
}
MKEY = {alias: k for k, al in METRIC_NAMES.items() for alias in al + [k]}
ROMAN = {'i': 1, 'ii': 2, 'iii': 3, 'iv': 4}


def norm(s):
    return re.sub(r'\s+', ' ', str(s).strip().lower().replace('ё', 'е'))


def parse_quarter(v):
    """26-2, 2026-2, '2 кв 2026', 'II кв. 2026', пусто = всегда. Excel любит превращать 26-2 в дату 26 февраля."""
    if v is None or str(v).strip() == '':
        return ''
    if isinstance(v, (datetime.date, datetime.datetime)):
        return f'{v.day:02d}-{v.month}' if 1 <= v.month <= 4 else None
    t = norm(v)
    m = re.fullmatch(r'(?:20)?(\d\d)\s*[-./]\s*([1-4])', t)
    if m:
        return f'{m.group(1)}-{m.group(2)}'
    m = re.fullmatch(r'([1-4]|i{1,3}|iv)\s*(?:кв\.?|квартал|q)?\s*(?:20)?(\d\d)(?:\s*г\.?)?', t)
    if m:
        k = m.group(1)
        return f'{m.group(2)}-{ROMAN.get(k, k)}'
    return None


UNIT_KEY = {}
for nid, node in IDS.items():
    if nid.startswith('en:'):
        continue
    for name in {node['en'], node['name'], nid.split('/')[-1], nid}:
        UNIT_KEY.setdefault(norm(name), node['id'])
for alias in ['регион', 'евразийский регион', 'eurasian area', 'area', 'region']:
    UNIT_KEY[alias] = 'area'
# официальные названия целиком («Уфимский приход») и прежние русские названия из старых отчётов —
# чтобы пояснения, написанные по-старому, продолжали находить своё подразделение
for nid, node in IDS.items():
    if not nid.startswith('en:') and node.get('type') == 'приход':
        UNIT_KEY.setdefault(norm(node['name'] + ' приход'), node['id'])
OLD_RU = {
    'Armenia Mission': ['Миссия Армения–Грузия', 'Миссия Армения-Грузия', 'Миссия Армения Грузия', 'Армянско-грузинская миссия', 'Армения–Грузия'],
    'Yerevan District': ['Ереванский округ', 'Ереванский армянский округ'],
    'Mission Branches': ['Приход миссии Армения–Грузия', 'Приход миссии Армения-Грузия', 'Приход Армянско-грузинской миссии'],
    'Avlabari': ['Авлабари'], 'Batumi': ['Батуми'], 'Temka': ['Темка'], 'Ajapnyak': ['Аджапняк', 'Ачапняк'],
    'Yerevan Central': ['Ереван Центральный'], 'Artashat': ['Арташат'], 'Gyumri': ['Гюмри'], 'Vanadzor': ['Ванадзор'],
    'Yerevan 1st (Armenian)': ['Ереван 1-й', 'Ереван 1-й (армянский)', 'Ереван первый'],
    'Yerevan 2nd (English)': ['Ереван 2-й', 'Ереван 2-й (английский)', 'Ереван второй'],
    'Kahovskiy': ['Каховский'], 'Korolev': ['Королёв', 'Королев'], 'Ostankino': ['Останкино'], 'Perovo': ['Перово'],
    'Voronezh': ['Воронеж'], 'Minsk': ['Минск'], 'Kaliningrad': ['Калининград'], 'Vyborg': ['Выборг'],
    'Tsentralniy': ['Центральный', 'Санкт-Петербург Центральный'], 'Krasnodar': ['Краснодар'], 'Novorossiysk': ['Новороссийск'],
    'Rostov Severny': ['Ростов Северный'], 'Rostov Center': ['Ростов Центральный'], 'Simferopol': ['Симферополь'],
    'Sochi': ['Сочи'], 'Taganrog': ['Таганрог'], 'Engels(Marks,Balakovo)': ['Энгельс'], 'Penza': ['Пенза'], 'Samara': ['Самара'],
    'Tolyatti': ['Тольятти'], 'Volgograd Tsentralny': ['Волгоград Центральный', 'Волгоград'], 'Almaty': ['Алматы'],
    'Chelyabinsk+Surgut': ['Челябинск'], "Kazan'": ['Казань'], 'Perm + Izhevsk': ['Пермь'], 'Tyumen': ['Тюмень'],
    'Ufa + Orenburg': ['Уфа'], 'Yekaterinburg 1st + Kurgan': ['Екатеринбург 1-й'], 'Yekaterinburg 2nd': ['Екатеринбург 2-й'],
    'Ulan-Ude': ['Улан-Удэ'], 'Vladivostok': ['Владивосток'], 'Yuzhno-Sakhalinsk': ['Южно-Сахалинск'],
    'Barnaul Altaisky': ['Барнаул', 'Барнаул Алтайский'], 'Krasnoyarsk': ['Красноярск'], 'Novosibirsk 1st': ['Новосибирск 1-й'],
    'Novosibirsk 2nd': ['Новосибирск 2-й'], 'Omsk': ['Омск'], 'Tomsk': ['Томск'],
}
for key, olds in OLD_RU.items():
    nid = UNIT_KEY.get(norm(key)) or UNIT_KEY.get(norm(RU.get(key, '')))
    if not nid:
        nid = next((n['id'] for i, n in IDS.items() if not i.startswith('en:') and _base(n['en']) == _base(key)), None)
    for o in olds:
        if nid:
            UNIT_KEY.setdefault(norm(o), nid)
# короткие имена верхних подразделений из сводных листов
for alias, nid in {'spb stake': 'St.Petersburg Stake', 'moscow msn': 'Moscow Mission', 'moscow mission': 'Moscow Mission',
                   'yekaterinburg mission': 'Yekaterinburg Msn', 'novosibirsk mission': 'Novosibirsk Msn',
                   'armenia/georgia msn': 'Armenia Mission', 'armenia georgia mission': 'Armenia Mission',
                   'rostov dis': 'Rostov District', 'yekaterinburg dis': 'Yekaterinburg District'}.items():
    UNIT_KEY.setdefault(alias, nid)

notes, note_problems = [], []
sheet = next((n for n in wb.sheetnames if norm(n) == 'пояснения'), None)
if sheet:
    ws = wb[sheet]
    for r in range(2, ws.max_row + 1):
        unit, q, metric, text, to_region = (ws.cell(r, c).value for c in range(1, 6))
        if not text or not str(text).strip():
            if unit:
                note_problems.append(f'строка {r}: нет текста пояснения')
            continue
        nid = UNIT_KEY.get(norm(unit or 'регион'))
        if not nid:
            note_problems.append(f'строка {r}: не знаю подразделение «{unit}» — напишите как в шапке листа Excel или как на странице отчёта')
            continue
        qq = parse_quarter(q)
        if qq is None or (qq and qq not in QS):
            note_problems.append(f'строка {r}: не понял квартал «{q}» — пишите как 26-2 (в ячейке с текстовым форматом) или «II кв. 2026»'
                                 if qq is None else f'строка {r}: квартала {qq} нет в отчёте (есть {QS[0]} … {QS[-1]})')
            continue
        mk = ''
        if metric and str(metric).strip():
            mk = MKEY.get(norm(metric).replace('ё', 'е')) or MKEY.get(norm(metric))
            if not mk:
                note_problems.append(f'строка {r}: не знаю показатель «{metric}» — пояснение будет без привязки к показателю')
                mk = ''
        region = norm(to_region or '') in ('да', 'д', 'yes', 'y', '1', '+', 'x', 'х', 'true', 'истина')
        notes.append(dict(node=nid, q=qq, metric=mk, text=str(text).strip(), row=r, region=region))

# ---- Логотип в шапке: файл logo_small.* рядом со скриптом (белый символ на прозрачном фоне) ----------
import base64
def load_logo(stem):
    for ext, mime in (('svg', 'image/svg+xml'), ('png', 'image/png'), ('jpg', 'image/jpeg'), ('jpeg', 'image/jpeg')):
        p = os.path.join(HERE, f'{stem}.{ext}')
        if os.path.exists(p):
            return f'data:{mime};base64,' + base64.b64encode(open(p, 'rb').read()).decode()
    return None
logo = {'small': load_logo('logo_small'),          # шапка
        'light': load_logo('logo_light'), 'dark': load_logo('logo_dark')}   # подвал (необязательно)

payload = dict(quarters=QS, tree=area, excelFlags=flags, excelHeader=report_quarter_label(),
               source=os.path.basename(SRC), notes=notes, logo=logo,
               built=datetime.date.today().strftime('%d.%m.%Y'))

tpl = open(os.path.join(HERE, 'template.html'), encoding='utf-8').read()
# шаблон кладём внутрь страницы (base64), чтобы кнопка «Выгрузить» могла собрать отчёты подразделений
LOCK_HTML = r'''<div id="lock" role="dialog" aria-labelledby="lock-t">
  <form id="lock-f" autocomplete="off">
    <svg width="40" height="40" viewBox="0 0 24 24" aria-hidden="true"><rect x="4.5" y="10.5" width="15" height="10" rx="2.2" fill="none" stroke="currentColor" stroke-width="1.6"/><path d="M8 10.5V7.5a4 4 0 0 1 8 0v3" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round"/><circle cx="12" cy="15.4" r="1.4" fill="currentColor"/></svg>
    <h1 id="lock-t">Отчёт защищён паролем</h1>
    <p>This report is password-protected</p>
    <input id="lock-p" type="password" placeholder="Пароль · Password" aria-label="Пароль" autofocus>
    <button id="lock-b" type="submit">Открыть · Open</button>
    <div id="lock-e" role="alert"></div>
  </form>
</div>
<style>
body.locked > *:not(#lock){display:none!important}
#lock{position:fixed;inset:0;display:grid;place-items:center;background:var(--page);padding:20px;z-index:100}
#lock form{width:min(380px,100%);background:var(--surface);border:1px solid var(--border);border-radius:16px;box-shadow:0 10px 40px rgba(0,30,55,.12);padding:28px 28px 22px;text-align:center;border-top:4px solid #007DA5}
#lock svg{color:#007DA5}
#lock h1{font-size:21px;margin:10px 0 2px}
#lock p{margin:0 0 18px;color:var(--muted);font-size:14px}
#lock input{width:100%;box-sizing:border-box;font:inherit;font-size:16px;padding:11px 14px;border:1px solid var(--axis);border-radius:10px;background:var(--surface);color:var(--ink)}
#lock input:focus{outline:2px solid #007DA5;outline-offset:1px;border-color:#007DA5}
#lock button{margin-top:12px;width:100%;font:inherit;font-weight:600;font-size:15px;padding:11px;border:0;border-radius:10px;background:#007DA5;color:#fff;cursor:pointer}
#lock button[disabled]{opacity:.65;cursor:progress}
#lock-e{min-height:20px;margin-top:10px;color:var(--bad);font-size:14px}
@keyframes lockshake{20%,60%{transform:translateX(-6px)}40%,80%{transform:translateX(6px)}}
#lock form.bad{animation:lockshake .35s}
</style>
<script>
const ENC = /*__ENC__*/null;
document.body.classList.add('locked');
(function(){
  const f = document.getElementById('lock-f'), inp = document.getElementById('lock-p'), btn = document.getElementById('lock-b'), err = document.getElementById('lock-e');
  const b = s => Uint8Array.from(atob(s), c => c.charCodeAt(0));
  if(!(window.crypto && crypto.subtle)){ err.textContent = 'Этот браузер не может расшифровать файл. Откройте его в Chrome, Edge, Firefox или Safari.'; btn.disabled = true; return; }
  f.addEventListener('submit', async e => {
    e.preventDefault(); if(!inp.value) return;
    btn.disabled = true; btn.textContent = 'Открываю… · Opening…'; err.textContent = ''; f.classList.remove('bad');
    try{
      const base = await crypto.subtle.importKey('raw', new TextEncoder().encode(inp.value), 'PBKDF2', false, ['deriveKey']);
      const key = await crypto.subtle.deriveKey({name:'PBKDF2', salt:b(ENC.s), iterations:ENC.it, hash:'SHA-256'}, base, {name:'AES-GCM', length:256}, false, ['decrypt']);
      const pt = await crypto.subtle.decrypt({name:'AES-GCM', iv:b(ENC.iv)}, key, b(ENC.ct));
      window.__DATA__ = JSON.parse(new TextDecoder().decode(pt));
      document.getElementById('lock').remove(); document.body.classList.remove('locked');
      const m = document.getElementById('main'), s = document.createElement('script'); s.textContent = m.textContent; m.remove(); document.body.append(s);
    }catch(x){
      err.textContent = 'Неверный пароль · Wrong password'; void f.offsetWidth; f.classList.add('bad'); inp.select();
      btn.disabled = false; btn.textContent = 'Открыть · Open';
    }
  });
})();
</script>
'''

tpl_b64 = base64.b64encode(tpl.encode('utf-8')).decode()


def atomic_write_text(path, content):
    """Не трогаем прежний файл, пока новый не записан и не синхронизирован."""
    folder = os.path.dirname(os.path.abspath(path))
    tmp = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', newline='',
                                         dir=folder, prefix='.report-', suffix='.tmp',
                                         delete=False) as out:
            tmp = out.name
            out.write(content)
            out.flush()
            os.fsync(out.fileno())
        os.replace(tmp, path)
        tmp = None
    finally:
        if tmp is not None:
            try:
                os.unlink(tmp)
            except FileNotFoundError:
                pass


def write_page(path, data, with_export):
    """Страница = шаблон + данные. С паролем данные шифруются целиком (см. ниже)."""
    html = tpl.replace('/*__TPL__*/""', json.dumps(tpl_b64) if with_export else '""', 1)
    data_json = json.dumps(data, ensure_ascii=False, separators=(',', ':'))
    if PASSWORD is None:
        html = html.replace('/*__DATA__*/null', data_json.replace('<', '\\u003c'), 1)
    else:
        # AES-256-GCM, ключ из пароля через PBKDF2-SHA256 (600 000 итераций, случайная соль).
        # В файле остаётся только шифротекст; браузер расшифровывает его у себя после ввода пароля, без интернета.
        import hashlib
        ITER = 600_000
        salt, iv = os.urandom(16), os.urandom(12)
        key = hashlib.pbkdf2_hmac('sha256', PASSWORD.encode('utf-8'), salt, ITER, 32)
        ct = AESGCM(key).encrypt(iv, data_json.encode('utf-8'), None)
        b64 = lambda b: base64.b64encode(b).decode()
        enc = json.dumps({'s': b64(salt), 'iv': b64(iv), 'it': ITER, 'ct': b64(ct)})
        html = html.replace('/*__DATA__*/null', 'window.__DATA__', 1)
        html = html.replace('<script id="main">', LOCK_HTML.replace('/*__ENC__*/null', enc) + '<script id="main" type="text/x-locked">', 1)
    atomic_write_text(path, html)


write_page(OUT, payload, True)
# Файл «Проверка данных» — для себя: замечания к данным (объяснены / нет) и ошибки в таблице. Не рассылается.
CHECK_OUT = os.path.join(os.path.dirname(os.path.abspath(OUT)), 'Проверка_данных.html')
write_page(CHECK_OUT, dict(payload, mode='check', noteProblems=note_problems,
                           reportFile=os.path.basename(OUT)), False)
if PASSWORD is not None:
    print('\nОтчёт зашифрован паролем (файл проверки — тем же паролем).' + ('  Внимание: пароль короче 8 символов — его легко подобрать.' if len(PASSWORD) < 8 else ''))
    print('Отдельные отчёты кольев, миссий и округов из кнопки «Выгрузить» паролем не защищаются.')
print(f'Готово: {OUT}')
print(f'Проверка данных (для себя, не рассылать): {CHECK_OUT}')
from collections import Counter
cells = Counter(f['kind'] for f in flags)
groups = Counter(f['kind'] for f in {(f['kind'], f['node'], f['metric']): f for f in flags}.values())
print(f'Кварталы: {", ".join(QS)}')
print('Замечания к самому Excel-файлу (за все кварталы):')
print(f'  неверный процент в ячейке: {cells["formula"]} ячеек, это {groups["formula"]} строк «Формула» на странице')
print(f'  итог листа не равен сумме: {cells["total"]} ячеек, это {groups["total"]} строк «Итог»')
print(f'  подразделения без данных:  {cells["empty"]}')
print('Подробно, со списком ячеек и с тем, какие замечания уже объяснены пояснениями, — в файле «Проверка_данных.html».')
if sheet:
    print(f'Пояснений прочитано: {len(notes)}' + (f', с ошибками: {len(note_problems)}' if note_problems else ''))
    for p in note_problems:
        print('  ⚠ ' + p)
else:
    print('Листа «Пояснения» нет — пояснений в отчёте не будет.')
if NO_RU:
    print('Нет русского названия (на странице будет по-английски): ' + ', '.join(sorted(NO_RU)))
    print('  Добавьте их в словарь RU в начале build.py: \'Название из Excel\': \'Русское название\',')
