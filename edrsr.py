#!/usr/bin/env python3
"""ЄДРСР: вибірка судових рішень з відкритих даних (data.gov.ua; резерв — сайт ДСА court.gov.ua), тексти, пошук за текстом, реєстр Excel і зведений PDF.

Автор: В'ячеслав Плескач. Версія 3.2 (25.11.2026): у номері справи відкидається лише літерний суфікс (-к, -ц, -а…), а не все
після останнього дефіса («10/б-711» більше не збігається з «10/б-743»); дати --from/--to — РРРР-ММ-ДД або ДД.ММ.РРРР,
інший формат — помилка (раніше мовчки 0 документів); grep завжди перезаписує <out>.docs.tsv; попередження, якщо той
самий запис ПІБ судді знайдено в кількох судах; назва суду шукається в довіднику найновішого року.
3.1 (20.11.2026): резервне джерело архівів — сайт ДСА court.gov.ua/opendata.

Команди:
  check   — перевірити, чи доступні з цього середовища портал data.gov.ua, сайт ДСА і сервер текстів od.reyestr.court.gov.ua
  courts  — знайти код суду ЄДРСР за назвою, кодом або префіксом номера справи (напр. 640 → ОАСК)
  meta    — відбір метаданих (суд, суддя в будь-якому записі ПІБ, вид, форма, інстанція, категорія, справа, дати) → TSV
  doc     — рішення за doc_id або посиланням reyestr.court.gov.ua/Review/<id> → TSV
  texts   — паралельне завантаження RTF за TSV і перетворення в TXT
  grep    — пошук за регулярним виразом у завантажених текстах → TSV збігів з фрагментами (+ відфільтрований TSV)
  report  — реєстр Excel і зведений PDF (із закладками) за TSV
  all     — meta → texts → report однією командою

Приклади:
  python3 edrsr.py check
  python3 edrsr.py courts --name "Печерський"
  python3 edrsr.py meta --years 2020-2026 --case "757/28368/24-к" --out sprava.tsv
  python3 edrsr.py meta --years 2025-2026 --court-name "Печерський районний" --judge "Вовк Сергій Володимирович" \\
      --kind 2 --form 5 --from 2025-01-01 --to 2026-10-31 --out vovk.tsv
  python3 edrsr.py doc --id https://reyestr.court.gov.ua/Review/133055529 --out doc.tsv
  python3 edrsr.py texts --in vovk.tsv --dir texts --keep-rtf
  python3 edrsr.py grep --in vovk.tsv --dir texts --pattern "скарг\\w* на постанову" --out hits.tsv
  python3 edrsr.py report --in hits.tsv.docs.tsv --dir texts --xlsx reiestr.xlsx --pdf rishennia.pdf
  python3 edrsr.py all --years 2025-2026 --case "761/7012/25" --name sprava_761_7012_25

Важливо:
  * Річні архіви групуються за роком НАДХОДЖЕННЯ до реєстру (receipt_date), а не ухвалення: --years має охоплювати
    роки від відкриття провадження до поточного (рішення кінця грудня потрапляють у наступний рік).
  * Код суду ЄДРСР (court_code, 4 знаки, напр. 2606) ≠ префікс номера справи (3 знаки, напр. 757). Команда courts показує обидва.
  * Джерело й дату даних за кожен рік скрипт пише у <out>.dzherela.txt.
  * Джерело архівів: типово data.gov.ua, а якщо портал не відповідає — сайт ДСА (ті самі набори того самого розпорядника,
    оновлюються щодня). --dzherelo dsa — одразу ДСА; --dzherelo portal — лише портал. Архіви edrsr_data_РРРР.zip
    можна також покласти в теку і передати --local <тека>.
Залежності (ставляться самі): duckdb, striprtf, openpyxl, pypdf; для PDF — LibreOffice (soffice).
"""
import argparse, csv, json, os, re, shutil, subprocess, sys, time, urllib.request, urllib.parse, ssl, zipfile, glob
from concurrent.futures import ThreadPoolExecutor

VERSION = '3.2'
AUTHOR = "В'ячеслав Плескач"
# Ідентифікатори наборів «Єдиний державний реєстр судових рішень за РРРР рік» на data.gov.ua (перевірено 05.11.2026).
PORTAL = {2006: 'a0c5d4c8-ae68-48b0-87c2-e4548b7774d4', 2007: 'dedce089-cf4b-4adf-97dc-1c5152080d05', 2008: '5d93b552-2049-43c2-b9a4-32e54653c7e6',
          2009: 'dc344166-85f7-4278-bae3-e33331c58029', 2010: '2fd047e6-35ff-4bf1-88f7-99b5261a9e7f', 2011: '7e9ea7cc-fec1-45d7-8ebe-c0f3d3e406a6',
          2012: 'e3c22d72-6827-4ad6-888d-5da637f91dd1', 2013: '3bc24b95-1f9e-41d6-bffa-c57088474e4a', 2014: 'e5c3c05a-61f5-4422-a41b-82dff1f71856',
          2015: '3951f2eb-a35b-4b35-a305-7afe471a9bca', 2016: '66169489-8877-48ae-a061-9eb91852876b', 2017: 'a7a72086-8b58-478d-b84b-eaa27d06ae9f',
          2018: 'd2c2af31-cc4e-4a85-9f35-be4dba2354c2', 2019: '1c13f2f1-8bc4-41d4-81dc-ea4527fa73de', 2020: '83f1292d-43c9-43b8-8f89-aa2f3d14fd3a',
          2021: 'bcffb3ae-8e4f-4443-8058-d525257ab69e', 2022: '42c195b3-b915-4000-934a-7741825d1812', 2023: 'e0f2c66a-f773-442e-afcd-569be6a36c82',
          2024: '1a6b5c08-0879-4387-8878-ac0027494274', 2025: '2a95ae17-ed46-42ae-ad66-dd9b7a727bb7', 2026: '16ab7f06-7414-405f-8354-0a492475272d'}
API = os.environ.get('EDRSR_PORTAL_API', 'https://data.gov.ua/api/3/action/')
# Сторінки наборів «Єдиний державний реєстр судових рішень за РРРР рік» на сайті ДСА: https://dsa.court.gov.ua/dsa/inshe/oddata/<N>/
# (перевірено 20.11.2026; для 2022–2026 ідентифікатори наборів збігаються з PORTAL). Дублікати 2018–2019 (404, 405, 407, 416)
# не оновлюються з 26.06.2026 — не використовувати. Набору 2025 (879) немає в пошуковому каталозі ДСА, тому номери вшито.
DSA = {2006: 11, 2007: 12, 2008: 13, 2009: 29, 2010: 24, 2011: 25, 2012: 26, 2013: 30, 2014: 31, 2015: 32, 2016: 33,
       2017: 34, 2018: 160, 2019: 350, 2020: 541, 2021: 680, 2022: 763, 2023: 826, 2024: 829, 2025: 879, 2026: 7636}
DSA_HOST = os.environ.get('EDRSR_DSA_HOST', 'https://dsa.court.gov.ua')
SOURCE = 'auto'          # auto | portal | dsa (--dzherelo)
PORTAL_DOWN = None       # причина, якщо портал уже не відповів у цьому запуску (щоб не чекати на кожен рік)
TEXT_HOST = 'https://od.reyestr.court.gov.ua/'
CACHE = os.environ.get('EDRSR_CACHE', os.path.expanduser('~/.cache/edrsr'))
DICTS = ('courts.csv', 'judgment_forms.csv', 'justice_kinds.csv', 'cause_categories.csv', 'regions.csv', 'instances.csv')


def log(*a):
    print(*a, file=sys.stderr, flush=True)


def need(mod, pkg=None):
    try:
        return __import__(mod)
    except ImportError:
        subprocess.run([sys.executable, '-m', 'pip', 'install', '-q', '--break-system-packages', pkg or mod], check=True)
        return __import__(mod)


def this_year():
    return time.localtime().tm_year


def parse_years(s):
    y1, _, y2 = s.partition('-')
    try:
        ys = list(range(int(y1), int(y2 or y1) + 1))
    except ValueError:
        ys = []
    if not ys:
        sys.exit(f'--years «{s}»: потрібно рік або проміжок за зростанням, напр. 2025 або 2023-2026.')
    return ys


def iso_date(s, name):
    """РРРР-ММ-ДД або ДД.ММ.РРРР → РРРР-ММ-ДД (дати в архіві порівнюються як рядки)."""
    import datetime as dt
    for f in ('%Y-%m-%d', '%d.%m.%Y'):
        try:
            return dt.datetime.strptime(s.strip(), f).date().isoformat()
        except ValueError:
            pass
    sys.exit(f'{name}: не розпізнано дату «{s}» (потрібно РРРР-ММ-ДД або ДД.ММ.РРРР).')


def ssl_ctx():
    for p in (os.environ.get('SSL_CERT_FILE'), os.environ.get('REQUESTS_CA_BUNDLE'), os.path.expanduser('~/.ccr/ca-bundle.crt'), '/root/.ccr/ca-bundle.crt'):
        if p and os.path.exists(p):
            return ssl.create_default_context(cafile=p)
    return None


def http_get(url, timeout=60):
    return urllib.request.urlopen(urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0 edrsr.py'}), timeout=timeout, context=ssl_ctx()).read()


# ---------------------------------------------------------------- check
def probe(url, timeout=25):
    t = time.time()
    try:
        http_get(url, timeout)
        return True, f'{time.time() - t:.1f} с'
    except urllib.error.HTTPError as e:
        if e.code in (403, 405, 407) and 'proxy' in str(e.headers).lower():
            return False, f'HTTP {e.code} від проксі середовища — домен не дозволено'
        # 403/429 від самого сайту (напр. захист Cloudflare під час збою data.gov.ua) — теж недоступність
        return (e.code in (400, 404)), f'HTTP {e.code}' + (' — сайт відмовляє в доступі' if e.code in (401, 403, 429) else '')
    except Exception as e:
        return False, str(e)[:120]


def cmd_check(a):
    tests = [('Портал data.gov.ua (метадані наборів)', API + 'package_show?id=' + PORTAL[max(PORTAL)]),
             ('Сайт ДСА court.gov.ua (резервне джерело архівів)', f'{DSA_HOST}/open_data_json.php?json={DSA[max(DSA)]}'),
             ('Сервер текстів od.reyestr.court.gov.ua (RTF)', TEXT_HOST)]
    ok_all, res = True, []
    for name, url in tests:
        ok, info = probe(url)
        ok_all &= ok
        res.append(ok)
        log(f'{"OK " if ok else "НЕМАЄ"}  {name}: {info}')
    if not tests or ok_all:
        log('Усе доступно: можна працювати.')
        return
    if res[1] and res[2]:
        log('\nПортал data.gov.ua недоступний, але сайт ДСА працює: архіви братимуться з ДСА автоматично — можна працювати.')
        return
    log('\nЩо робити, якщо щось недоступно:\n'
        '  * data.gov.ua і сайт ДСА обидва недоступні → вибірка неможлива без архівів. Варіанти: (1) дозволити домени data.gov.ua і\n'
        '    od.reyestr.court.gov.ua у мережевих налаштуваннях виконання коду; (2) завантажити edrsr_data_РРРР.zip з\n'
        '    data.gov.ua вручну, прикріпити в чат і запускати з --local <тека>; (3) для кількох рішень — відкрити\n'
        '    reyestr.court.gov.ua у браузері користувача (Claude in Chrome).\n'
        '  * Заблоковано лише od.reyestr → метадані працюють, а тексти — тільки через браузер або вручну.')
    sys.exit(2)


# ---------------------------------------------------------------- архіви
def portal_pkg(year):
    def get(u):
        for att in range(2):
            try:
                return json.loads(http_get(u, 60))['result']
            except Exception:
                if att == 0:
                    time.sleep(15)
        raise IOError('data.gov.ua не відповідає')
    if year in PORTAL:
        return get(API + 'package_show?id=' + PORTAL[year])
    q = urllib.parse.quote(f'Єдиний державний реєстр судових рішень за {year} рік')
    for p in sorted(get(API + f'package_search?q={q}&rows=20')['results'], key=lambda p: p['metadata_modified'], reverse=True):
        if any(r['name'] == f'edrsr_data_{year}.zip' for r in p['resources']):
            return p
    raise LookupError(f'На data.gov.ua не знайдено набору ЄДРСР за {year} рік')


def dsa_page_no(year):
    """Номер сторінки набору на сайті ДСА: вшитий або з каталогу (для нових років)."""
    if year in DSA:
        return DSA[year]
    data = urllib.parse.urlencode({'sEcho': 1, 'iColumns': 4, 'iDisplayStart': 0, 'iDisplayLength': 1000, 'iSortCol_0': 3,
                                   'sSortDir_0': 'desc', 'q_ver': 'arbitr', 'date': '~', 'sSearch': '', 'court': ''}).encode()
    req = urllib.request.Request(DSA_HOST + '/opendata_data.php', data=data, headers={
        'User-Agent': 'Mozilla/5.0 edrsr.py', 'X-Requested-With': 'XMLHttpRequest',
        'Referer': DSA_HOST + '/dsa/inshe/oddata/', 'Content-Type': 'application/x-www-form-urlencoded; charset=UTF-8'})
    rows = json.loads(urllib.request.urlopen(req, timeout=120, context=ssl_ctx()).read())['aaData']
    hits = [int(m.group(1)) for r in rows if re.match(rf'Єдиний державний реєстр судових рішень за {year} рік\.', r[0])
            for m in [re.search(r'oddata/(\d+)', r[4])] if m]
    if not hits:
        raise LookupError(f'На сайті ДСА не знайдено набору ЄДРСР за {year} рік (каталог: {DSA_HOST}/dsa/inshe/oddata/)')
    return max(hits)


def dsa_resource(year):
    """→ (URL архіву, позначка версії, дата оновлення ДД.ММ.РРРР) з паспорта набору ДСА і заголовків файлу."""
    n = dsa_page_no(year)
    pas = json.loads(http_get(f'{DSA_HOST}/open_data_json.php?json={n}', 60))
    files = {k: v for f in pas.get('Файли', []) for k, v in f.items()}
    url = files.get(f'edrsr_data_{year}.zip') or next((v for k, v in files.items() if k.endswith('.zip')), None)
    if not url:
        raise LookupError(f'У паспорті набору ДСА {n} немає архіву за {year} рік')
    h = urllib.request.urlopen(urllib.request.Request(url, method='HEAD', headers={'User-Agent': 'Mozilla/5.0 edrsr.py'}),
                               timeout=60, context=ssl_ctx()).headers
    size = int(h.get('Content-Length') or 0)
    if size < 100000:
        raise IOError(f'архів ДСА за {year} рік порожній або битий ({size} байт): {url}')
    import email.utils
    lm = h.get('Last-Modified', '')
    try:
        day = time.strftime('%d.%m.%Y', time.localtime(email.utils.parsedate_to_datetime(lm).timestamp()))
    except Exception:
        day = '?'
    return url, f'dsa:{n}:{lm}:{h.get("ETag", "")}:{size}', day, n


def download(url, z, year, where):
    log(f'[{year}] завантаження з {where}')
    t = time.time()
    subprocess.run(['curl', '-sS', '-L', '-f', '--retry', '3', '-m', '1800', '-o', z, url], check=True)
    if not zipfile.is_zipfile(z):
        raise IOError(f'завантажений файл не є ZIP ({os.path.getsize(z)} байт)')
    log(f'[{year}] завантажено за {time.time() - t:.0f} с')


def build_year(zpath, d, stamp, year):
    """Розпаковує архів, переводить documents.csv у parquet (швидкі повторні запити, менше місця) і видаляє CSV."""
    duckdb = need('duckdb')
    tmp = d + '.new'
    shutil.rmtree(tmp, ignore_errors=True)
    with zipfile.ZipFile(zpath) as zf:
        zf.extractall(tmp)
    t = time.time()
    con = duckdb.connect()
    src = os.path.join(tmp, 'documents.csv')
    con.execute(f"""COPY (SELECT * FROM read_csv('{src}', delim='\t', header=true, quote='"', all_varchar=true))
                    TO '{tmp}/documents.parquet' (FORMAT parquet, COMPRESSION zstd)""")
    os.remove(src)
    open(os.path.join(tmp, '.stamp'), 'w').write(stamp)
    shutil.rmtree(d, ignore_errors=True)
    os.replace(tmp, d)
    log(f'[{year}] підготовлено за {time.time() - t:.0f} с')


def year_dir(year, local=None):
    """→ (тека року з documents.parquet і довідниками, опис джерела)."""
    os.makedirs(CACHE, exist_ok=True)
    d = os.path.join(CACHE, f'x{year}')
    st = os.path.join(d, '.stamp')
    cached = open(st).read() if os.path.exists(st) and os.path.exists(os.path.join(d, 'documents.parquet')) else ''
    if local:
        zs = glob.glob(os.path.join(local, '**', f'edrsr_data_{year}*.zip'), recursive=True)
        if zs:
            z = max(zs, key=os.path.getmtime)
            stamp = f'local:{os.path.basename(z)}:{os.path.getsize(z)}:{int(os.path.getmtime(z))}'
            label = f'локальний архів {os.path.basename(z)} (файл від {time.strftime("%d.%m.%Y", time.localtime(os.path.getmtime(z)))})'
            if cached != stamp:
                log(f'[{year}] локальний архів {z}')
                build_year(z, d, stamp, year)
            return d, label
    global PORTAL_DOWN
    errs = []
    z = os.path.join(CACHE, f'edrsr_data_{year}.zip')
    if SOURCE != 'dsa' and not PORTAL_DOWN:
        try:
            meta = portal_pkg(year)
            zips = [r for r in meta['resources'] if r['name'].endswith('.zip')]
            url = ([r['url'] for r in zips if r['name'] == f'edrsr_data_{year}.zip'] or [r['url'] for r in zips])[0]
            stamp = meta['metadata_modified']
            label = f'data.gov.ua, набір оновлено {stamp[:10]}'
            if cached == stamp:
                return d, label
            download(url, z, year, f'data.gov.ua (набір оновлено {stamp[:16]})')
            build_year(z, d, stamp, year)
            os.remove(z)
            return d, label
        except LookupError as e:
            errs.append(str(e))
        except Exception as e:
            errs.append(f'data.gov.ua: {str(e)[:100]}')
            if 'не відповідає' in str(e) or isinstance(e, (urllib.error.URLError, OSError)):
                PORTAL_DOWN = str(e)[:100]
    elif PORTAL_DOWN:
        errs.append(f'data.gov.ua: {PORTAL_DOWN}')
    if SOURCE != 'portal':
        try:
            url, stamp, day, n = dsa_resource(year)
            label = f'сайт ДСА (court.gov.ua/opendata, набір {n}), архів оновлено {day}'
            if errs:
                log(f'[{year}] портал недоступний ({errs[-1]}) — беру архів із сайту ДСА')
            if cached == stamp:
                return d, label
            download(url, z, year, f'сайту ДСА (набір {n}, архів від {day})')
            build_year(z, d, stamp, year)
            os.remove(z)
            return d, label
        except Exception as e:
            errs.append(f'сайт ДСА: {str(e)[:100]}')
    if errs and all(e.startswith('На ') for e in errs):
        sys.exit('; '.join(errs))
    why = '; '.join(errs) or 'джерело вимкнено параметром --dzherelo'
    if cached:
        log(f'[{year}] УВАГА: джерела недоступні ({why}) — використано локальний кеш ({cached[:40]})')
        return d, f'локальний кеш ({cached[:40]}); data.gov.ua і сайт ДСА недоступні під час запуску'
    sys.exit(f'[{year}] Джерела недоступні ({why}) і кешу немає. Запустіть «check»; варіанти — у його виводі '
                 f'(зокрема --local з архівом edrsr_data_{year}.zip).')


def rd(d, f):
    return f"read_csv('{d}/{f}', delim='\t', header=true, quote='\"', all_varchar=true)"


def load_dicts(con, d):
    for t, f in (('courts', 'courts.csv'), ('forms', 'judgment_forms.csv'), ('kinds', 'justice_kinds.csv'),
                 ('cats', 'cause_categories.csv'), ('regions', 'regions.csv'), ('inst', 'instances.csv')):
        con.execute(f'CREATE OR REPLACE TABLE {t} AS SELECT * FROM {rd(d, f)}')
    con.execute("UPDATE courts SET name = trim(name)")


# ---------------------------------------------------------------- судді
APOS = str.maketrans({'’': "'", 'ʼ': "'", '`': "'", '‘': "'", 'ё': 'е'})


def jparse(s):
    """ПІБ у будь-якому записі → (множина ключів «прізвище+ініціали», множина повних ПІБ)."""
    s = (s or '').lower().translate(APOS)
    toks = [t for t in re.split(r'[\s.,]+', s) if t]
    words = [t for t in toks if len(t) > 1]
    inits = ''.join(t for t in toks if len(t) == 1)
    keys, full = set(), set()
    if len(words) == 3 and not inits:           # «Прізвище Ім'я По батькові» або «Ім'я По батькові Прізвище»
        a, b, c = words
        keys |= {(a, b[0] + c[0]), (c, a[0] + b[0])}
        full |= {(a, b, c), (c, a, b)}
    elif len(words) == 2 and not inits:         # «Прізвище Ім'я» / «Ім'я Прізвище»
        a, b = words
        keys |= {(a, b[0]), (b, a[0])}
    elif len(words) == 1:                       # «Вовк С. В.», «К.М. Недашківська», «Вовк»
        keys.add((words[0], inits))
    elif words:
        keys.add((words[0], inits))
    return keys, full


def judge_match(query, value):
    qk, qf = jparse(query)
    vk, vf = jparse(value)
    if qf and vf:
        return bool(qf & vf)
    for qs, qi in qk:
        for vs, vi in vk:
            if qs == vs and (not qi or not vi or vi.startswith(qi) or qi.startswith(vi)):
                return True
    return False


# ---------------------------------------------------------------- справи
# Ключ справи: без пробілів і регістру, без кінцевого ЛІТЕРНОГО суфікса («757/28368/24-к» → «757/28368/24»).
# Цифровий хвіст — частина номера («10/б-711», «2-98-09»); до 3.2 він відкидався, і до вибірки потрапляли чужі справи.
def case_key_sql(col):
    return f"regexp_replace(lower(replace(coalesce({col},''),' ','')), '-[^/0-9-]{{1,4}}$', '')"


def case_key(s):
    return re.sub(r'-[^/0-9-]{1,4}$', '', re.sub(r'\s+', '', s or '').lower())


# ---------------------------------------------------------------- courts
def cmd_courts(a):
    duckdb = need('duckdb')
    year = a.year or this_year()
    d, src = year_dir(year, a.local)
    con = duckdb.connect()
    load_dicts(con, d)
    idx = os.path.join(CACHE, f'prefiksy_{year}.tsv')
    if not os.path.exists(idx) or os.path.getmtime(idx) < os.path.getmtime(os.path.join(d, '.stamp')):
        log('Обчислюю префікси номерів справ за судами (один раз на рік архіву)…')
        con.execute(f"""COPY (WITH p AS (SELECT court_code, split_part(cause_num,'/',1) pre, count(*) n
                              FROM read_parquet('{d}/documents.parquet') WHERE regexp_matches(cause_num,'^\\d{{3}}/') GROUP BY 1,2),
                         t AS (SELECT court_code, sum(n) tot FROM p GROUP BY 1)
                    SELECT court_code, pre, round(100.0*n/tot,1) share FROM p JOIN t USING(court_code)
                    QUALIFY row_number() OVER (PARTITION BY court_code ORDER BY n DESC)=1) TO '{idx}' (DELIMITER '\t', HEADER)""")
    con.execute(f"CREATE TABLE pref AS SELECT * FROM read_csv('{idx}', delim='\t', header=true, all_varchar=true)")
    w = []
    if a.name:
        for part in a.name.split():
            w.append(f"lower(c.name) LIKE '%{part.lower().replace(chr(39), chr(39)*2)}%'")
    if a.code:
        w.append('c.court_code IN (' + ','.join(f"'{x.strip().zfill(4)}'" for x in a.code.split(',')) + ')')
    if a.prefix:
        w.append(f"p.pre = '{a.prefix.split('/')[0]}' AND CAST(p.share AS DOUBLE) >= 60")
    rows = con.execute(f"""SELECT c.court_code, c.name, i.name, r.name, coalesce(p.pre,''), coalesce(p.share,'')
                           FROM courts c LEFT JOIN inst i ON i.instance_code=c.instance_code
                           LEFT JOIN regions r ON r.region_code=c.region_code LEFT JOIN pref p ON p.court_code=c.court_code
                           {'WHERE ' + ' AND '.join(w) if w else ''} ORDER BY c.court_code""").fetchall()
    print('court_code\tназва\tінстанція\tрегіон\tпрефікс_справ\tчастка_%')
    for r in rows:
        print('\t'.join(str(x) for x in r))
    log(f'Знайдено {len(rows)} (довідник архіву {year}: {src}). Префікс — найчастіший перший сегмент номера справи;'
        ' для апеляційних і касаційних судів він не визначає суд (номер справи зберігається з першої інстанції).'
        ' Ліквідованих судів у довіднику поточного року може не бути — задайте --year того року, коли суд діяв.')


def resolve_court_names(con, names, soft=False):
    codes = []
    for nm in names.split(';'):
        parts = [p.lower().replace("'", "''") for p in nm.split()]
        rows = con.execute('SELECT court_code, name FROM courts WHERE ' + ' AND '.join(f"lower(name) LIKE '%{p}%'" for p in parts)).fetchall()
        if not rows and soft:
            return None
        if not rows:
            sys.exit(f'Суд «{nm}» не знайдено в довіднику. Спробуйте: edrsr.py courts --name "<частина назви>"')
        if len(rows) > 1:
            exact = [r for r in rows if r[1].lower() == nm.lower()]
            if len(exact) == 1:
                rows = exact
            else:
                sys.exit(f'Назва «{nm}» відповідає кільком судам — уточніть або задайте --court:\n' + '\n'.join(f'  {c}\t{n}' for c, n in rows))
        codes.append(rows[0][0])
        log(f'Суд: {rows[0][1]} → {rows[0][0]}')
    return codes


# ---------------------------------------------------------------- meta
SELECT_OUT = """d.doc_id, d.cause_num, substr(d.adjudication_date,1,10) AS adjudication_date, substr(d.receipt_date,1,10) AS receipt_date,
       d.judge, d.court_code, c.name AS court_name, d.judgment_code, f.name AS judgment_form,
       d.justice_kind, k.name AS justice_kind_name, d.category_code, ct.name AS category_name,
       d.status, d.doc_url, 'https://reyestr.court.gov.ua/Review/' || d.doc_id AS web_url, d.archive_year"""
JOINS = """LEFT JOIN courts c ON c.court_code = d.court_code LEFT JOIN forms f ON f.judgment_code = d.judgment_code
           LEFT JOIN kinds k ON k.justice_kind = d.justice_kind LEFT JOIN cats ct ON ct.category_code = d.category_code"""


def write_tsv(path, cols, rows):
    with open(path, 'w', encoding='utf-8', newline='') as f:
        w = csv.writer(f, delimiter='\t', lineterminator='\n')
        w.writerow(cols)
        w.writerows(rows)


def write_sources(out, used):
    from zoneinfo import ZoneInfo
    import datetime as dt
    with open(out + '.dzherela.txt', 'w', encoding='utf-8') as f:
        f.write(f'Вибірка {dt.datetime.now(ZoneInfo("Europe/Kyiv")):%d.%m.%Y %H:%M} (Київ), edrsr.py {VERSION}; джерела за роками надходження:\n')
        for y, s in used:
            f.write(f'{y}\t{s}\n')
    for y, s in used:
        log(f'  {y}: {s}')


def cmd_meta(a):
    duckdb = need('duckdb')
    years = parse_years(a.years)
    a.date_from = iso_date(a.date_from, '--from') if a.date_from else None
    a.date_to = iso_date(a.date_to, '--to') if a.date_to else None
    con = duckdb.connect()
    parts, used = [], []
    court_codes = [c.strip().zfill(4) for c in a.court.split(',') if c.strip()] if a.court else []
    if a.court_name:                              # довідник найновішого року (актуальні назви); ліквідований суд — у найстарішому
        for yy in dict.fromkeys([years[-1], years[0]]):
            load_dicts(con, year_dir(yy, a.local)[0])
            codes = resolve_court_names(con, a.court_name, soft=(yy != years[0]))
            if codes is not None:
                court_codes += codes
                break
    for y in years:
        d, src = year_dir(y, a.local)
        used.append((y, src))
        if y == years[-1] or not parts:
            load_dicts(con, d)                    # довідники найновішого року (назви судів актуальніші)
        P = f"read_parquet('{d}/documents.parquet')"
        w = []
        if court_codes: w.append('court_code IN (' + ','.join(f"'{c}'" for c in court_codes) + ')')
        if a.kind: w.append('justice_kind IN (' + ','.join(f"'{k}'" for k in a.kind.split(',')) + ')')
        if a.form: w.append('judgment_code IN (' + ','.join(f"'{k}'" for k in a.form.split(',')) + ')')
        if a.cat: w.append('category_code IN (' + ','.join(f"'{k}'" for k in a.cat.split(',')) + ')')
        if a.case:
            cs = [c.strip() for c in a.case.split(',') if c.strip()]
            like = [c for c in cs if '%' in c or '_' in c]
            exact = [case_key(c) for c in cs if c not in like]
            cw = [f"cause_num LIKE '{c}'" for c in like]
            if exact: cw.append(f"{case_key_sql('cause_num')} IN (" + ','.join(f"'{c}'" for c in exact) + ')')
            w.append('(' + ' OR '.join(cw) + ')')
        if a.date_from: w.append(f"substr(adjudication_date,1,10) >= '{a.date_from}'")
        if a.date_to: w.append(f"substr(adjudication_date,1,10) <= '{a.date_to}'")
        if a.instance:
            w.append(f"court_code IN (SELECT court_code FROM {rd(d, 'courts.csv')} WHERE instance_code IN ("
                     + ','.join(f"'{i}'" for i in a.instance.split(',')) + '))')
        if not a.include_removed: w.append("status <> '0'")
        if a.judge_like: w.append(f"judge ILIKE '%{a.judge_like}%'")
        if a.judge:
            base = ' AND '.join(w) or 'true'
            vals = [r[0] for r in con.execute(f'SELECT DISTINCT judge FROM {P} WHERE {base}').fetchall()]
            hit = [v for v in vals if any(judge_match(q, v) for q in a.judge.split(';'))]
            if not hit:
                w.append('false')
            else:
                w.append('judge IN (' + ','.join("'" + v.replace("'", "''") + "'" for v in hit) + ')')
        parts.append(f"SELECT *, '{y}' AS archive_year FROM {P}" + (' WHERE ' + ' AND '.join(w) if w else ''))
    q = f"""WITH d AS ({' UNION ALL '.join(parts)}) SELECT {SELECT_OUT} FROM d {JOINS}
            QUALIFY row_number() OVER (PARTITION BY d.doc_id ORDER BY d.archive_year DESC) = 1
            ORDER BY adjudication_date, d.doc_id"""
    t = time.time()
    rows = con.execute(q).fetchall()
    cols = [c[0] for c in con.description]
    write_tsv(a.out, cols, rows)
    log(f'Відібрано {len(rows)} документів за {time.time() - t:.1f} с → {a.out}')
    if a.judge:
        cnt = {}
        for r in rows:
            cnt[r[4]] = cnt.get(r[4], 0) + 1
        log('Записи ПІБ судді, що збіглися (перевірте, чи це одна особа):')
        for v, n in sorted(cnt.items(), key=lambda x: -x[1]):
            log(f'  {n:>6}  {v}')
        if len({tuple(sorted(jparse(v)[1])) for v in cnt if jparse(v)[1]}) > 1:
            log('  УВАГА: кілька різних повних ПІБ — можливі тезки; звузьте --court/--court-name або вкажіть повне ПІБ.')
        by = {}
        for r in rows:
            by.setdefault(r[4], {}).setdefault(r[6] or r[5], 0)
            by[r[4]][r[6] or r[5]] += 1
        for v, cs in by.items():
            if len(cs) > 1:                       # у даних переважно ініціали: тезок видно лише за судом
                log(f'  УВАГА: запис «{v}» є в {len(cs)} судах — це можуть бути різні судді (або відрядження): '
                    + '; '.join(f'{c} — {n}' for c, n in sorted(cs.items(), key=lambda x: -x[1])[:6])
                    + '. Звузьте --court / --court-name.')
    write_sources(a.out, used)
    return a.out


# ---------------------------------------------------------------- doc
def cmd_doc(a):
    duckdb = need('duckdb')
    ids = sorted({m for x in a.id.split(',') for m in re.findall(r'\d{5,}', x)}, key=int)
    if not ids:
        sys.exit('Не знайдено doc_id (очікується число або посилання …/Review/<id>)')
    years = parse_years(a.years) if a.years else list(range(this_year(), 2005, -1))
    years = sorted(years, reverse=True)
    if not a.years:
        log('Без --years архіви перебираються від поточного року назад (кожен — сотні МБ); якщо рік надходження відомий, задайте --years.')
    con = duckdb.connect()
    found, used, left = [], [], set(ids)
    for y in years:
        if not left:
            break
        d, src = year_dir(y, a.local)
        used.append((y, src))
        load_dicts(con, d)
        P = f"read_parquet('{d}/documents.parquet')"
        lst = ','.join(f"'{i}'" for i in left)
        rows = con.execute(f"SELECT {SELECT_OUT} FROM (SELECT *, '{y}' AS archive_year FROM {P} WHERE doc_id IN ({lst})) d {JOINS}").fetchall()
        cols = [c[0] for c in con.description]
        found += rows
        left -= {r[0] for r in rows}
        if rows:
            log(f'[{y}] знайдено {len(rows)}')
        mn = con.execute(f'SELECT min(CAST(doc_id AS BIGINT)) FROM {P}').fetchone()[0]
        if left and mn and all(int(i) > mn * 1.5 for i in left) and not a.years:
            log(f'  doc_id більші за діапазон {y} року — пошук у попередніх роках малоймовірний, зупиняюся')
            break
    write_tsv(a.out, cols, found)
    log(f'Знайдено {len(found)} з {len(ids)} → {a.out}' + (f'; не знайдено: {", ".join(sorted(left))}' if left else ''))
    write_sources(a.out, used)


# ---------------------------------------------------------------- texts
def read_tsv(path):
    with open(path, encoding='utf-8', newline='') as f:
        return list(csv.DictReader(f, delimiter='\t'))


def cmd_texts(a):
    need('striprtf')
    import importlib
    sr = importlib.import_module('striprtf.striprtf')
    os.makedirs(a.dir, exist_ok=True)
    rows = [r for r in read_tsv(a.inp) if r.get('doc_url')]
    def done(r):
        return os.path.exists(os.path.join(a.dir, r['doc_id'] + '.txt')) and (not a.keep_rtf or os.path.exists(os.path.join(a.dir, r['doc_id'] + '.rtf')))
    todo = [r for r in rows if not done(r)]
    log(f'З посиланням на текст: {len(rows)}, до завантаження: {len(todo)}')
    ctx = ssl_ctx()

    def get(r):
        err = ''
        for att in range(4):
            try:
                data = urllib.request.urlopen(r['doc_url'], timeout=40, context=ctx).read()
                if len(data) < 300:
                    raise IOError('порожня відповідь')
                if a.keep_rtf:
                    open(os.path.join(a.dir, r['doc_id'] + '.rtf'), 'wb').write(data)
                raw = data.decode('cp1251', 'replace')
                txt = sr.rtf_to_text(raw, errors='ignore') if raw.lstrip().startswith('{\\rtf') else re.sub(r'<[^>]+>', ' ', raw)
                open(os.path.join(a.dir, r['doc_id'] + '.txt'), 'w', encoding='utf-8').write(txt)
                return None
            except Exception as e:
                err = str(e)[:80]
                if '404' in err:
                    break
                time.sleep(2 * (att + 1))
        return (r['doc_id'], err)

    t = time.time(); fails = []; n = 0
    with ThreadPoolExecutor(a.threads) as ex:
        for res in ex.map(get, todo):
            n += 1
            if res: fails.append(res)
            if n % 500 == 0:
                log(f'  {n}/{len(todo)} за {time.time() - t:.0f} с, помилок {len(fails)}')
    dt = time.time() - t
    log(f'Готово: {n} за {dt:.0f} с ({n / max(dt, 1):.1f} док/с), помилок {len(fails)}')
    if not fails and os.path.exists(os.path.join(a.dir, '_fails.tsv')):
        os.remove(os.path.join(a.dir, '_fails.tsv'))       # перелік помилок попереднього запуску більше не чинний
    if fails:
        with open(os.path.join(a.dir, '_fails.tsv'), 'w', encoding='utf-8') as f:
            for i, e in fails: f.write(f'{i}\t{e}\n')
        log(f'  перелік помилок → {a.dir}/_fails.tsv')
        if len(fails) == n and n:
            log('  Усі завантаження невдалі: можливо, od.reyestr.court.gov.ua заблоковано — запустіть «check».')


# ---------------------------------------------------------------- grep
def cmd_grep(a):
    rx = re.compile(a.pattern, 0 if a.case_sensitive else re.I)
    rows = read_tsv(a.inp)
    out, keep = [], []
    no_text = 0
    for r in rows:
        p = os.path.join(a.dir, r['doc_id'] + '.txt')
        if not os.path.exists(p):
            no_text += 1
            continue
        t = re.sub(r'\s+', ' ', open(p, encoding='utf-8', errors='replace').read())
        ms = list(rx.finditer(t))
        if a.invert:
            if not ms:
                keep.append(r)
            continue
        if not ms:
            continue
        keep.append(dict(r, hits=str(len(ms)), snippet=' … '.join(t[max(0, m.start() - a.context):m.end() + a.context] for m in ms[:3])))
        for m in ms[:a.max_per_doc]:
            out.append([r['doc_id'], r.get('cause_num', ''), r.get('adjudication_date', ''), r.get('judgment_form', ''),
                        m.group(0), t[max(0, m.start() - a.context):m.end() + a.context], r.get('web_url', '')])
    write_tsv(a.out, ['doc_id', 'cause_num', 'adjudication_date', 'judgment_form', 'match', 'fragment', 'web_url'], out)
    # завжди перезаписуємо: інакше після пошуку без збігів лишався б файл попереднього пошуку
    cols = list(keep[0].keys()) if keep else (list(rows[0].keys()) if rows else ['doc_id']) + ([] if a.invert else ['hits', 'snippet'])
    write_tsv(a.out + '.docs.tsv', cols, [[k.get(c, '') for c in cols] for k in keep])
    if os.path.exists(a.inp + '.dzherela.txt'):
        shutil.copy(a.inp + '.dzherela.txt', a.out + '.docs.tsv.dzherela.txt')
    log(f'Документів зі збігом{" (інвертовано: без збігу)" if a.invert else ""}: {len(keep)} з {len(rows)}; '
        f'фрагментів: {len(out)} → {a.out}; документи → {a.out}.docs.tsv' + (f'; без тексту: {no_text}' if no_text else ''))


# ---------------------------------------------------------------- report
def xlsx_report(rows, path, sources=None):
    need('openpyxl')
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
    cols = [('№', None, 6), ('Дата ухвалення', 'adjudication_date', 12), ('Дата надходження', 'receipt_date', 12),
            ('Суд', 'court_name', 32), ('Суддя', 'judge', 24), ('Форма', 'judgment_form', 14),
            ('Вид судочинства', 'justice_kind_name', 16), ('Категорія', 'category_name', 40), ('Номер справи', 'cause_num', 18),
            ('Посилання', 'web_url', 14), ('Статус', 'status', 8), ('doc_id', 'doc_id', 11)]
    if rows and 'hits' in rows[0]:
        cols += [('Збігів', 'hits', 8), ('Фрагмент', 'snippet', 60)]
    wb = Workbook()
    ws = wb.active
    ws.title = 'Реєстр'
    ws.append([c[0] for c in cols])
    for i, r in enumerate(rows, 1):
        ws.append([i] + [r.get(k, '') if k != 'web_url' else 'відкрити' for _, k, _ in cols[1:]])
        cell = ws.cell(row=i + 1, column=[c[1] for c in cols].index('web_url') + 1)
        cell.hyperlink = r.get('web_url', '')
        cell.font = Font(color='0563C1', underline='single')
    for j, (_, _, wdt) in enumerate(cols, 1):
        ws.column_dimensions[get_column_letter(j)].width = wdt
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill('solid', fgColor='DDEBF7')
        c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = ws.dimensions
    if sources:
        s = wb.create_sheet('Джерела')
        for line in sources:
            s.append(line.rstrip('\n').split('\t'))
        s.column_dimensions['A'].width = 12
        s.column_dimensions['B'].width = 80
    wb.properties.creator = AUTHOR
    wb.properties.title = 'Реєстр судових рішень (ЄДРСР)'
    wb.save(path)
    log(f'Excel: {len(rows)} рядків → {path}')


def pdf_report(rows, dirpath, path):
    need('pypdf')
    from pypdf import PdfWriter, PdfReader
    soffice = shutil.which('soffice') or shutil.which('libreoffice')
    if not soffice:
        log('PDF не створено: немає LibreOffice (soffice). Встановіть його або обмежтеся Excel.')
        return
    tmp = os.path.join(dirpath, '_pdf')
    os.makedirs(tmp, exist_ok=True)
    srcs, missing = [], []
    for r in rows:
        i = r['doc_id']
        rtf, txt = os.path.join(dirpath, i + '.rtf'), os.path.join(dirpath, i + '.txt')
        src = rtf if os.path.exists(rtf) else txt if os.path.exists(txt) else None
        if not src:
            missing.append(i); continue
        if not os.path.exists(os.path.join(tmp, i + '.pdf')):
            srcs.append(src)
    t = time.time()
    for k in range(0, len(srcs), 100):
        subprocess.run([soffice, '--headless', '--norestore', '-env:UserInstallation=file:///tmp/lo_edrsr', '--convert-to', 'pdf',
                        '--outdir', tmp] + srcs[k:k + 100], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL, timeout=3600)
        log(f'  PDF: {min(k + 100, len(srcs))}/{len(srcs)} за {time.time() - t:.0f} с')
    w = PdfWriter()
    n = 0
    for r in rows:
        p = os.path.join(tmp, r['doc_id'] + '.pdf')
        if not os.path.exists(p):
            continue
        start = len(w.pages)
        for pg in PdfReader(p).pages:
            w.add_page(pg)
        w.add_outline_item(f"{r.get('adjudication_date', '')} · {r.get('judgment_form', '')} · {r.get('cause_num', '')} · {r['doc_id']}", start)
        n += 1
    w.add_metadata({'/Author': AUTHOR, '/Title': 'Тексти судових рішень (ЄДРСР)'})
    with open(path, 'wb') as f:
        w.write(f)
    lost = len(rows) - n - len(missing)
    log(f'PDF: {n} документів, {len(w.pages)} сторінок → {path}' + (f'; без тексту: {len(missing)}' if missing else '')
        + (f'; НЕ перетворено в PDF (LibreOffice): {lost}' if lost else ''))


def cmd_report(a):
    rows = read_tsv(a.inp)
    src = a.inp + '.dzherela.txt'
    sources = open(src, encoding='utf-8').readlines() if os.path.exists(src) else None
    if a.xlsx:
        xlsx_report(rows, a.xlsx, sources)
    if a.pdf:
        if len(rows) > 2000 and not a.force:
            log(f'PDF на {len(rows)} документів буде дуже великим; додайте --force або звузьте вибірку (напр. через grep).')
        else:
            pdf_report(rows, a.dir, a.pdf)


def cmd_all(a):
    a.out = a.name + '.tsv'
    cmd_meta(a)
    n = len(read_tsv(a.out))
    if not n:
        return
    a.inp, a.dir = a.out, a.name + '_texts'
    a.keep_rtf = not a.no_pdf
    cmd_texts(a)
    a.xlsx = a.name + '.xlsx'
    a.pdf = None if a.no_pdf else a.name + '.pdf'
    cmd_report(a)


# ---------------------------------------------------------------- CLI
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sp = ap.add_subparsers(dest='cmd', required=True)
    sp.add_parser('check')

    def common_src(p):
        p.add_argument('--local', help='тека з архівами edrsr_data_РРРР.zip (якщо портал і сайт ДСА недоступні)')
        p.add_argument('--dzherelo', choices=['auto', 'portal', 'dsa'], default='auto',
                       help='джерело архівів: auto — data.gov.ua, а за недоступності сайт ДСА (типово); portal; dsa')
        p.add_argument('--cache', help='тека кешу архівів (типово ~/.cache/edrsr)')
        p.add_argument('--source', help=argparse.SUPPRESS)  # сумісність із edrsr_vybirka.py 2.x; ігнорується

    c = sp.add_parser('courts'); common_src(c)
    c.add_argument('--name', help='частини назви суду (усі мають входити)')
    c.add_argument('--code', help='код(и) ЄДРСР через кому')
    c.add_argument('--prefix', help='префікс номера справи (перші 3 цифри, напр. 757 або 640)')
    c.add_argument('--year', type=int, help='рік архіву, довідник якого брати (для ліквідованих судів)')

    def filters(p):
        common_src(p)
        p.add_argument('--years', required=True, help='роки НАДХОДЖЕННЯ: 2023-2026 або 2025')
        p.add_argument('--court', help='код(и) суду ЄДРСР через кому, напр. 2606')
        p.add_argument('--court-name', help="назва суду або її частини (кілька судів — через ';')")
        p.add_argument('--judge', help="ПІБ у будь-якому записі: «Вовк С. В.», «Вовк Сергій Володимирович»; кілька — через ';'")
        p.add_argument('--judge-like', help='частина рядка judge (ILIKE)')
        p.add_argument('--kind', help='justice_kind через кому: 1 цив., 2 крим., 3 госп., 4 адм., 5 адмінправопорушення')
        p.add_argument('--form', help='judgment_code через кому: 1 вирок, 2 постанова, 3 рішення, 4 наказ, 5 ухвала, 6 окрема ухвала, 7 додаткове, 10 окрема думка')
        p.add_argument('--instance', help='інстанція суду: 1 касаційна, 2 апеляційна, 3 перша')
        p.add_argument('--cat', help='category_code через кому (ненадійно — краще grep за текстом)')
        p.add_argument('--case', help="номер(и) справи через кому (суфікс -к/-ц/-а ігнорується) або шаблон LIKE '757/%%/26-к'")
        p.add_argument('--from', dest='date_from', help='дата ухвалення від, РРРР-ММ-ДД')
        p.add_argument('--to', dest='date_to', help='дата ухвалення до, РРРР-ММ-ДД')
        p.add_argument('--include-removed', action='store_true', help='включити документи зі status=0 (вилучені, без тексту)')

    m = sp.add_parser('meta'); filters(m); m.add_argument('--out', required=True)
    d = sp.add_parser('doc'); common_src(d)
    d.add_argument('--id', required=True, help='doc_id або посилання через кому')
    d.add_argument('--years', help='роки надходження, якщо відомі (інакше — від поточного назад)')
    d.add_argument('--out', default='doc.tsv')
    t = sp.add_parser('texts')
    t.add_argument('--in', dest='inp', required=True); t.add_argument('--dir', required=True)
    t.add_argument('--threads', type=int, default=16)
    t.add_argument('--keep-rtf', action='store_true', help='зберігати RTF (потрібно для PDF у форматуванні оригіналу)')
    g = sp.add_parser('grep')
    g.add_argument('--in', dest='inp', required=True); g.add_argument('--dir', required=True)
    g.add_argument('--pattern', required=True, help='регулярний вираз Python')
    g.add_argument('--context', type=int, default=200); g.add_argument('--max-per-doc', type=int, default=5)
    g.add_argument('--case-sensitive', action='store_true'); g.add_argument('--invert', action='store_true', help='документи БЕЗ збігу')
    g.add_argument('--out', required=True)
    r = sp.add_parser('report')
    r.add_argument('--in', dest='inp', required=True); r.add_argument('--dir', default='texts')
    r.add_argument('--xlsx'); r.add_argument('--pdf'); r.add_argument('--force', action='store_true')
    al = sp.add_parser('all'); filters(al)
    al.add_argument('--name', required=True, help='основа назв файлів результату')
    al.add_argument('--threads', type=int, default=16); al.add_argument('--no-pdf', action='store_true'); al.add_argument('--force', action='store_true')
    a = ap.parse_args()
    global CACHE
    if getattr(a, 'cache', None):
        CACHE = os.path.abspath(a.cache)
    global SOURCE
    SOURCE = getattr(a, 'dzherelo', None) or 'auto'
    if getattr(a, 'source', None):
        log('УВАГА: --source застарів і ігнорується; джерело задає --dzherelo (або --local).')
    {'check': cmd_check, 'courts': cmd_courts, 'meta': cmd_meta, 'doc': cmd_doc, 'texts': cmd_texts,
     'grep': cmd_grep, 'report': cmd_report, 'all': cmd_all}[a.cmd](a)


if __name__ == '__main__':
    main()
