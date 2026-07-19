# -*- coding: utf-8 -*-
"""Кабинет руководителя ТОО «Рудный-АБАТ-2006» (полигон + сортировка).
Читает хаб Абата (пишет бот abat-bot), ничего не пишет. Стиль — как кабинет ЛТ.
Разделы переключаются как страницы (JS): Обзор / Полигон / Расходы / Касса.
env: SPREADSHEET_ID (хаб Абата), GOOGLE_CREDENTIALS (ltrading-bot@l-trading), KAB_LOGIN, KAB_PASSWORD.
"""
import os, json, time
from functools import wraps
from flask import Flask, Response, request
import gspread

app = Flask(__name__)
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID", "")
KAB_LOGIN = os.environ.get("KAB_LOGIN", "")
KAB_PASSWORD = os.environ.get("KAB_PASSWORD", "")

MSK_SHEET_ID = os.environ.get("MSK_SHEET_ID", "1qoKYiVKBQK97_wtzJNWOHfywevGOMNz8lD4r_854Tp8")

_BOOK = None
_BOOK_MSK = None
_CACHE = {}
_TTL = 60

def get_book():
    global _BOOK
    if _BOOK is None:
        creds = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        _BOOK = gspread.service_account_from_dict(creds).open_by_key(SPREADSHEET_ID)
    return _BOOK

def get_book_msk():
    """Операционная таблица весовой/сортировки Абата (Viewer у робота). Кабинет только читает."""
    global _BOOK_MSK
    if _BOOK_MSK is None:
        creds = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        _BOOK_MSK = gspread.service_account_from_dict(creds).open_by_key(MSK_SHEET_ID)
    return _BOOK_MSK

def _ws_values_of(book_fn, cache_key, title):
    now = time.time()
    hit = _CACHE.get(cache_key)
    if hit and now - hit[0] < _TTL:
        return hit[1]
    try:
        vals = book_fn().worksheet(title).get_all_values()
    except gspread.WorksheetNotFound:
        vals = []
    except Exception:
        return hit[1] if hit else []
    _CACHE[cache_key] = (now, vals)
    return vals

def ws_values(title):
    return _ws_values_of(get_book, "hub:" + title, title)

def ws_values_msk(title):
    return _ws_values_of(get_book_msk, "msk:" + title, title)

def _check(l, p): return bool(KAB_LOGIN) and l == KAB_LOGIN and p == KAB_PASSWORD
def requires_auth(f):
    @wraps(f)
    def w(*a, **k):
        if not KAB_LOGIN or not KAB_PASSWORD:
            return Response("Вход не настроен: задайте KAB_LOGIN и KAB_PASSWORD.", 503)
        au = request.authorization
        if not au or not _check(au.username, au.password):
            return Response("Требуется вход", 401, {"WWW-Authenticate": 'Basic realm="Abat 2006"'})
        return f(*a, **k)
    return w

def num(x):
    s = str(x or "").replace("\xa0", "").replace(" ", "")
    if not s: return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return 0.0

_MORD = ["", "Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
         "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
def _mkey_sort(m):
    p = m.split()
    yr = int(p[-1]) if p and p[-1].isdigit() else 0
    base = p[0].capitalize() if p else ""
    return (yr, _MORD.index(base) if base in _MORD else 99)
def iso_month(d):
    d = str(d).strip()
    if len(d) >= 7 and d[4] == "-" and d[:4].isdigit() and d[5:7].isdigit():
        y, mo = int(d[:4]), int(d[5:7])
        if 1 <= mo <= 12: return f"{_MORD[mo]} {y}"
    return d

_KASSA_INTERNAL = ("сдач", "зачислен", "перемещен", "внесен", "инкасс", "прочих денежных", "прочие денежные")
_NON_EXPENSE = ("группа компаний", "аффилир", "внутренний оборот", "внутр")
def is_expense(cat): return not any(k in cat.lower() for k in _NON_EXPENSE)

def read_reisy():
    rows = ws_values("Рейсы"); out = {}
    for r in rows[1:]:
        if len(r) < 7 or not str(r[0]).strip(): continue
        mon = iso_month(r[0]); org = str(r[3]).strip() or "—"; ves = num(r[6])
        if ves <= 0: continue
        out.setdefault(mon, {}).setdefault(org, [0.0, 0])
        out[mon][org][0] += ves; out[mon][org][1] += 1
    return out

def read_rashody():
    rows = ws_values("Расходы_1С"); by_cat = {}
    for r in rows[1:]:
        if len(r) < 3 or not str(r[0]).strip(): continue
        mon = str(r[0]).strip(); cat = str(r[2]).strip(); tot = num(r[-1])
        if not tot: continue
        by_cat.setdefault(mon, {}); by_cat[mon][cat] = by_cat[mon].get(cat, 0.0) + tot
    return by_cat

def read_kassa():
    rows = ws_values("Касса1_1С"); out = {}
    for r in rows[1:]:
        if len(r) < 4 or not str(r[0]).strip(): continue
        mon = str(r[0]).strip(); typ = str(r[1]).strip().lower(); st = str(r[2]).strip().lower(); s = num(r[3])
        if any(k in st or k in typ for k in _KASSA_INTERNAL): continue
        out.setdefault(mon, {"приход": 0.0, "расход": 0.0})
        if "приход" in typ or "пко" in typ: out[mon]["приход"] += s
        elif "расход" in typ or "рко" in typ: out[mon]["расход"] += s
    return out

def read_kassa_svod():
    """«Касса_свод» (файл Светы, вкл. Кассу-2): Месяц|Раздел|Касса|Статья|Строка|Сумма →
    {месяц: {'Касса 1': {'приход':x,'расход':y}, 'Касса 2': {...}}}"""
    rows = ws_values("Касса_свод"); out = {}
    for r in rows[1:]:
        if len(r) < 6 or not str(r[0]).strip():
            continue
        mon = str(r[0]).strip(); rz = str(r[1]).strip().lower(); ks = str(r[2]).strip() or "Касса ?"
        s = num(r[5])
        if not s:
            continue
        m = out.setdefault(mon, {})
        k = m.setdefault(ks, {"приход": 0.0, "расход": 0.0})
        if "поступ" in rz or "приход" in rz:
            k["приход"] += s
        elif "расход" in rz:
            k["расход"] += s
    return out

def _sp(n): return f"{round(n):,}".replace(",", " ")
def _spt(n): return f"{n:,.1f}".replace(",", " ")

_NAV = [("obzor", "ti-layout-grid", "Обзор"), ("polygon", "ti-building-factory-2", "Полигон"),
        ("sort", "ti-recycle", "Сортировка"), ("tech", "ti-truck", "Техника"),
        ("naprav", "ti-chart-pie", "Направления"),
        ("rashody", "ti-credit-card", "Расходы"), ("kassa", "ti-cash", "Касса")]

def read_gsm():
    """Путевые листы × Нормы_ГСМ: {месяц: {техника: [кол-во, выдано л, положено л]}}.
    «Кол-во» путевого = моточасы или рейсы (по режиму); положено = Кол-во × норма(сезон).
    Сезон: ноя–мар = зима. Норма матчится по подстроке имени техники («Т-130 р721» → «Т-130»)."""
    normy = {}   # норм.имя → {режим: (зима, лето)}
    for r in ws_values_msk("Нормы_ГСМ")[1:]:
        if len(r) < 5 or not str(r[0]).strip():
            continue
        nm = _norm_org(r[0]); rez = str(r[2]).strip().lower()
        normy.setdefault(nm, {})[rez] = (num(r[3]), num(r[4]))
    out = {}
    for r in ws_values_msk("Путевые_листы")[1:]:
        if len(r) < 8 or not str(r[0]).strip():
            continue
        mon = iso_month(r[0])
        try:
            mo = int(str(r[0]).strip()[5:7])
        except (ValueError, IndexError):
            mo = 0
        winter = mo in (11, 12, 1, 2, 3)
        teh = str(r[2]).strip(); rez = str(r[5]).strip().lower()
        qty = num(r[6]); vydano = num(r[7])
        norm = 0.0
        tn = _norm_org(teh)
        for k, rr in normy.items():
            if len(k) >= 3 and (k in tn or tn in k):
                pair = rr.get(rez) or (list(rr.values())[0] if rr else None)
                if pair:
                    norm = pair[0] if winter else pair[1]
                break
        cur = out.setdefault(mon, {}).setdefault(teh, [0.0, 0.0, 0.0])
        cur[0] += qty; cur[1] += vydano; cur[2] += qty * norm
    return out

def _norm_org(s):
    """Нормализация имени организации для матча с тарифами: без ТОО/ИП/АО, кавычек, знаков."""
    s = str(s or "").lower()
    for w in ("тоо", "ип", "ао", "«", "»", '"', "'", ".", ",", "(", ")"):
        s = s.replace(w, " ")
    return "".join(ch for ch in s if ch.isalnum())

def read_tarify():
    """«Тарифы» весовой: норм.имя → тариф строй ₸/т (колонка ТБО пустая — тарифицируется строй)."""
    out = {}
    for r in ws_values_msk("Тарифы")[1:]:
        if len(r) < 3 or not str(r[0]).strip():
            continue
        t = num(r[2]) or num(r[1])
        if t > 0:
            out[_norm_org(r[0])] = t
    return out

def tarif_of(org, tarify):
    """Подбор тарифа: нормализованные имена, совпадение подстрокой (мин. 4 символа)."""
    n = _norm_org(org)
    if not n:
        return 0.0
    for k, t in tarify.items():
        if len(k) >= 4 and (k in n or n in k):
            return t
    return 0.0

def read_sortirovka():
    """Из операционной таблицы: Сортировка_виды (выпуск кг), Отгрузка (продажи кг), Цены, Остатки.
    → ({месяц:{вид:кг}}, {месяц:{вид:(кг,₸)}}, [(вид,кг,обновлено)])"""
    prices = {}
    for r in ws_values_msk("Цены")[1:]:
        if len(r) >= 2 and str(r[0]).strip():
            prices[str(r[0]).strip().lower()] = num(r[1])
    vypusk = {}
    for r in ws_values_msk("Сортировка_виды")[1:]:
        if len(r) < 4 or not str(r[0]).strip():
            continue
        mon = iso_month(r[0]); vid = str(r[2]).strip() or "—"; kg = num(r[3])
        if kg <= 0:
            continue
        vypusk.setdefault(mon, {}); vypusk[mon][vid] = vypusk[mon].get(vid, 0.0) + kg
    otgruzka = {}
    for r in ws_values_msk("Отгрузка")[1:]:
        if len(r) < 3 or not str(r[0]).strip():
            continue
        mon = iso_month(r[0]); vid = str(r[1]).strip() or "—"; kg = num(r[2])
        if kg <= 0:
            continue
        price = prices.get(vid.lower(), 0.0)
        cur = otgruzka.setdefault(mon, {}).setdefault(vid, [0.0, 0.0])
        cur[0] += kg; cur[1] += kg * price
    ostatki = []
    for r in ws_values_msk("Остатки")[1:]:
        if len(r) >= 2 and str(r[0]).strip():
            ostatki.append((str(r[0]).strip(), num(r[1]), str(r[2]).strip() if len(r) > 2 else ""))
    return vypusk, otgruzka, ostatki

def read_naprav():
    """«Направления_1С»: Месяц|Направление|Сумма|Примечание → {месяц: [(направление, сумма, примечание)]}"""
    rows = ws_values("Направления_1С"); out = {}
    for r in rows[1:]:
        if len(r) < 3 or not str(r[0]).strip(): continue
        mon = str(r[0]).strip(); nm = str(r[1]).strip(); s = num(r[2])
        note = str(r[3]).strip() if len(r) > 3 else ""
        if not nm: continue
        out.setdefault(mon, []).append((nm, s, note))
    return out

PAGE = r"""<!DOCTYPE html><html lang=ru><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>Абат 2006 · Кабинет руководителя</title>
<link rel="icon" type="image/svg+xml" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 200 184'><defs><clipPath id='c' clipPathUnits='userSpaceOnUse'><path clip-rule='evenodd' d='M100 18 L180 166 L20 166 Z M100 142 L146 166 L54 166 Z M100 46 L134 120 L66 120 Z'/></clipPath></defs><rect x='0' y='18' width='200' height='68' fill='%2333a15c' clip-path='url(%23c)'/><rect x='0' y='86' width='200' height='42' fill='%231c87a6' clip-path='url(%23c)'/><rect x='0' y='128' width='200' height='42' fill='%23185e9e' clip-path='url(%23c)'/></svg>">
<link rel=preconnect href="https://fonts.googleapis.com"><link rel=preconnect href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Manrope:wght@500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel=stylesheet>
<link href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.7.0/dist/tabler-icons.min.css" rel=stylesheet>
<style>
:root{--bg:#f4f2ed;--panel:#ffffff;--panel2:#eaf6fa;--line:#e5e1d8;--txt:#16233a;--muted:#5a615a;--dim:#9aa09a;
--green:#33a15c;--teal:#1c87a6;--blue:#185e9e;--ink:#16233a;--red:#c0392b;--mono:'JetBrains Mono',ui-monospace,monospace;--sans:'Manrope',system-ui,sans-serif}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--txt);font-family:var(--sans);-webkit-font-smoothing:antialiased}
.layout{display:flex;min-height:100vh}
.side{width:232px;background:var(--ink);padding:24px 16px;flex-shrink:0;position:sticky;top:0;height:100vh;display:flex;flex-direction:column}
.brand{display:flex;align-items:center;gap:12px}.brand .logo{width:38px;height:35px;display:flex;align-items:center;justify-content:center}
.brand .name{font-size:18px;font-weight:800;letter-spacing:-.01em;color:#fff}.brand .name i{display:none}
.brand-sub{font-family:var(--mono);font-size:8.5px;color:#66c1dd;letter-spacing:.14em;text-transform:uppercase;margin:6px 0 28px 1px}
.nav-title{font-family:var(--mono);font-size:9px;color:#5c718c;text-transform:uppercase;letter-spacing:.14em;margin:0 0 12px 8px}
.nav-btn{display:flex;align-items:center;gap:12px;width:100%;padding:11px 13px;border-radius:10px;color:#9fb2c4;font-size:14px;font-weight:600;margin-bottom:4px;cursor:pointer;transition:.15s;border:none;background:none;font-family:var(--sans)}
.nav-btn:hover{color:#fff;background:#1e2f49}.nav-btn.on{background:var(--teal);color:#fff;font-weight:700}
.nav-btn i{font-size:19px}
.main{flex:1;padding:28px 34px;min-width:0;max-width:1160px;position:relative;z-index:1}
.wm{position:fixed;right:-90px;bottom:-80px;width:600px;height:auto;opacity:.045;pointer-events:none;z-index:0;filter:blur(1.5px)}
.top{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:24px}
.hi{font-size:24px;font-weight:800;letter-spacing:-.01em}
select{background:#fff;color:var(--txt);border:1px solid var(--line);border-radius:10px;padding:10px 14px;font-size:14px;font-weight:700;font-family:var(--sans);cursor:pointer}
select option{background:var(--panel)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:18px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:20px;position:relative;overflow:hidden}
.card::before{content:"";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--teal)}
.card .val{font-family:var(--mono);font-size:26px;font-weight:700;color:var(--ink)}.card .val.green{color:var(--green)}.card .val.red{color:var(--red)}
.card .unit{font-size:14px;color:var(--muted);margin-left:3px}.card .lbl{color:var(--muted);font-size:12.5px;margin-top:9px}
.sec{display:none}.sec.on{display:block}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:20px;margin-bottom:16px}
.panel h2{font-size:15px;font-weight:800;margin:0 0 3px;letter-spacing:-.01em}.panel h2 i{color:var(--teal);margin-right:7px}
.panel .ph{color:var(--dim);font-size:11.5px;margin-bottom:14px;font-family:var(--mono);letter-spacing:.03em}
.scroll{max-height:460px;overflow-y:auto;margin:0 -4px;padding:0 4px}
.scroll::-webkit-scrollbar{width:8px}.scroll::-webkit-scrollbar-thumb{background:var(--line);border-radius:8px}
.rw{padding:10px 2px;border-bottom:1px solid var(--line)}.rw:last-child{border-bottom:none}
.rw .t{display:flex;justify-content:space-between;font-size:13.5px}.rw .nm{font-weight:600}.rw .mn{font-family:var(--mono);color:var(--muted)}
.rw .bar{height:5px;border-radius:3px;background:#eef1ec;overflow:hidden;margin-top:6px}.rw .fill{height:100%;border-radius:3px;background:linear-gradient(90deg,var(--green),var(--teal))}
.r2{display:flex;justify-content:space-between;padding:11px 2px;border-bottom:1px solid var(--line);font-size:13.5px}.r2:last-child{border-bottom:none}
.r2 .nm{font-weight:600}.r2 .mn{font-family:var(--mono);font-weight:700;color:var(--ink)}
.muted{color:var(--muted)}.foot{color:var(--dim);font-size:11px;font-family:var(--mono);margin-top:24px}
</style></head><body>
<div class=layout>
<svg class=wm viewBox="0 0 200 184"><rect x="0" y="18" width="200" height="152" fill="#16233a" clip-path="url(#abClip)"/></svg>
<aside class=side>
  <div class=brand><div class=logo><svg width="38" height="35" viewBox="0 0 200 184"><defs><clipPath id="abClip" clipPathUnits="userSpaceOnUse"><path clip-rule="evenodd" d="M100 18 L180 166 L20 166 Z M100 142 L146 166 L54 166 Z M100 46 L134 120 L66 120 Z"/></clipPath></defs><rect x="0" y="18" width="200" height="68" fill="#33a15c" clip-path="url(#abClip)"/><rect x="0" y="86" width="200" height="42" fill="#1c87a6" clip-path="url(#abClip)"/><rect x="0" y="128" width="200" height="42" fill="#185e9e" clip-path="url(#abClip)"/></svg></div><div class=name>Абат 2006<i class="ti ti-point-filled" style="font-size:10px"></i></div></div>
  <div class=brand-sub>полигон и сортировка</div>
  <div class=nav-title>разделы</div>
  __NAV__
</aside>
<main class=main>
  <div class=top><div class=hi>Здравствуйте!</div>__MPICK__</div>
  <div class=cards>__KPIS__</div>
  __SECTIONS__
  <div class=foot>Данные из 1С и весовой полигона через хаб · обновление ~1 мин</div>
</main>
</div>
<script>
function showSec(id){
  document.querySelectorAll('.sec').forEach(function(s){s.classList.toggle('on',s.id==='sec-'+id)});
  document.querySelectorAll('.nav-btn').forEach(function(b){b.classList.toggle('on',b.dataset.sec===id)});
  try{localStorage.setItem('abat_sec',id)}catch(e){}
}
(function(){var s='obzor';try{s=localStorage.getItem('abat_sec')||'obzor'}catch(e){}
 if(!document.getElementById('sec-'+s))s='obzor';showSec(s)})();
</script>
</body></html>"""

def render(**kw):
    html = PAGE
    for t in ("NAV", "MPICK", "KPIS", "SECTIONS"):
        html = html.replace("__%s__" % t, kw.get(t.lower(), ""))
    return html

@app.route("/health")
def health(): return "ok"

@app.route("/")
@requires_auth
def home():
    reisy = read_reisy(); rashody = read_rashody(); kassa = read_kassa(); naprav = read_naprav()
    svod = read_kassa_svod()
    try:
        vypusk, otgruzka, ostatki = read_sortirovka()
    except Exception:
        vypusk, otgruzka, ostatki = {}, {}, []
    try:
        gsm = read_gsm()
    except Exception:
        gsm = {}
    months = sorted(set(list(reisy) + list(rashody) + list(kassa) + list(naprav) + list(svod) + list(vypusk)),
                    key=_mkey_sort, reverse=True)
    nav = "".join(f"<button class=nav-btn data-sec={sid} onclick=\"showSec('{sid}')\"><i class='ti {ic}'></i>{nm}</button>"
                  for sid, ic, nm in _NAV)
    if not months:
        return render(nav=nav, sections="<div class=sec on id=sec-obzor><div class=panel><div class=muted>В хабе ещё нет данных.</div></div></div>")
    sel = request.args.get("m") or months[0]
    if sel not in months: sel = months[0]

    mpick = "<select onchange=\"location.href='?m='+encodeURIComponent(this.value)\">" + \
            "".join(f"<option {'selected' if m==sel else ''}>{m}</option>" for m in months) + "</select>"

    pol = reisy.get(sel, {}); tons = sum(v[0] for v in pol.values()); trips = sum(v[1] for v in pol.values())
    cats = rashody.get(sel, {}); rashod_total = sum(v for c, v in cats.items() if is_expense(c))
    k = kassa.get(sel, {"приход": 0.0, "расход": 0.0})
    sv = svod.get(sel)
    if sv:   # слой Светы (вкл. Кассу-2) приоритетнее 1С
        k = {"приход": sum(v["приход"] for v in sv.values()),
             "расход": sum(v["расход"] for v in sv.values())}
    kassa_lbl = "Наличные приход (К1+К2)" if sv else "Касса-1 приход"

    # P&L по начислению из «Направления_1С»: выручка (Кт 6010) и операционные затраты
    # (всё разнесённое + не разнесённое, БЕЗ внутренних перекладок НЗП/Закрытия)
    _INTERNAL_NP = ("ДвижениеНЗП", "ЗакрытиеМесяца")
    rev_nach = nach_total = 0.0
    for n, s, _nt in naprav.get(sel, []):
        if "ВЫРУЧКА" in n.upper():
            rev_nach += s
            continue
        if n.startswith("(без подразд.:") and n.split(":", 1)[1].strip(" )") in _INTERNAL_NP:
            continue
        nach_total += s
    profit = rev_nach - nach_total
    sebes_t = (nach_total / tons) if tons > 0 else 0.0

    if rev_nach or nach_total:
        kpis = (
            f"<div class=card><div class='val green'>{_sp(rev_nach)}<span class=unit>₸</span></div><div class=lbl>Выручка (начисление, 1С)</div></div>"
            f"<div class=card><div class='val red'>{_sp(nach_total)}<span class=unit>₸</span></div><div class=lbl>Затраты операционные</div></div>"
            f"<div class=card><div class='val {'green' if profit >= 0 else 'red'}'>{'+' if profit >= 0 else '−'}{_sp(abs(profit))}<span class=unit>₸</span></div><div class=lbl>Прибыль за месяц</div></div>"
            f"<div class=card><div class='val green'>{_spt(tons)}<span class=unit>т</span></div><div class=lbl>Принято на полигон · {trips} рейсов</div></div>"
            f"<div class=card><div class='val'>{_sp(sebes_t)}<span class=unit>₸/т</span></div><div class=lbl>Затраты на тонну приёма</div></div>"
            f"<div class=card><div class='val'>{_sp(k['приход'])}<span class=unit>₸</span></div><div class=lbl>{kassa_lbl}</div></div>")
    else:
        kpis = (
            f"<div class=card><div class='val green'>{_spt(tons)}<span class=unit>т</span></div><div class=lbl>Принято на полигон</div></div>"
            f"<div class=card><div class='val'>{trips}</div><div class=lbl>Рейсов за месяц</div></div>"
            f"<div class=card><div class='val red'>{_sp(rashod_total)}<span class=unit>₸</span></div><div class=lbl>Расходы (чистые)</div></div>"
            f"<div class=card><div class='val'>{_sp(k['приход'])}<span class=unit>₸</span></div><div class=lbl>{kassa_lbl}</div></div>")

    pol_items = sorted(pol.items(), key=lambda x: -x[1][0])
    mx = max([v[0] for _, v in pol_items], default=1)
    try:
        tarify = read_tarify()
    except Exception:
        tarify = {}
    bill_total = 0.0
    pol_rows = ""
    for o, v in pol_items:
        t = tarif_of(o, tarify)
        calc = v[0] * t
        bill_total += calc
        extra = f" · к выставл. {_sp(calc)} ₸ ({_sp(t)}/т)" if calc else ""
        pol_rows += (
            f"<div class=rw><div class=t><span class=nm>{o}</span><span class=mn>{v[1]} рейс · {_spt(v[0])} т{extra}</span></div>"
            f"<div class=bar><span class=fill style='width:{max(3,v[0]/mx*100):.0f}%'></span></div></div>")
    pol_rows = pol_rows or "<div class=muted>за этот месяц рейсов нет (весовая синхронизируется раз в час)</div>"
    bill_note = (f"<div class=ph style='margin-top:12px'>💰 Расчётно к выставлению по строй-тарифам: "
                 f"<b>{_sp(bill_total)} ₸</b> — сверить с реализацией 1С (клиенты без тарифа не считаются)</div>"
                 if bill_total else "")

    cat_items = sorted([(c, v) for c, v in cats.items() if is_expense(c)], key=lambda x: -x[1])
    cat_rows = "".join(f"<div class=r2><span class=nm>{c}</span><span class=mn>{_sp(v)} ₸</span></div>" for c, v in cat_items) \
               or "<div class=muted>нет данных</div>"

    # КАССА: приоритет — «Касса_свод» (файл Светы, вкл. Кассу-2), иначе Касса-1 из 1С
    if sv:
        rows_k = ""
        for ks in sorted(sv):
            v = sv[ks]
            rows_k += (f"<div class=r2><span class=nm>{ks} · приход</span><span class=mn>{_sp(v['приход'])} ₸</span></div>"
                       f"<div class=r2><span class=nm>{ks} · расход</span><span class=mn>{_sp(v['расход'])} ₸</span></div>")
        rows_k += (f"<div class=r2><span class=nm style='font-weight:800'>Всего наличные</span>"
                   f"<span class=mn style='font-weight:800'>{_sp(k['приход'])} / {_sp(k['расход'])} ₸</span></div>")
        kassa_html = (f"<div class=panel><h2><i class='ti ti-cash'></i>Наличные — файл экономиста</h2>"
                      f"<div class=ph>{sel} · Касса-1 официальная + Касса-2; загружено ботом из файла Светы</div>{rows_k}</div>")
    else:
        kassa_html = (f"<div class=panel><h2><i class='ti ti-cash'></i>Касса-1 (наличные из 1С)</h2>"
                      f"<div class=ph>{sel} · внутренние перемещения исключены; Касса-2 появится из файла Светы</div>"
                      f"<div class=r2><span class=nm>Приход</span><span class=mn>{_sp(k['приход'])} ₸</span></div>"
                      f"<div class=r2><span class=nm>Расход</span><span class=mn>{_sp(k['расход'])} ₸</span></div></div>")

    # НАПРАВЛЕНИЯ — карточки по 3 направлениям (ФОТ+материалы), внутренние перекладки скрыты
    np_items = naprav.get(sel, [])
    DIR_OF = {"полигон": "Полигон", "сортировочный комплекс": "Сортировочный комплекс",
              "офис (ауп)": "Офис (АУП)", "основное подразделение": "Офис (АУП)",
              "переработка стройотходов": "Переработка стройотходов",
              "производство (прочее)": "Полигон",
              "производство": "Производство (МСК+Полигон)"}
    HUMAN = {"ОтражениеНалоговойОтчетностиВРеглУчете": "Налоги и взносы",
             "ПоступлениеТоваровУслуг": "Услуги и ТМЦ поставщиков",
             "РеализацияТоваровУслуг": "Себестоимость реализации",
             "СписаниеТоваров": "Списания материалов"}
    INTERNAL = ("ДвижениеНЗП", "ЗакрытиеМесяца")   # внутренние перекладки затрат — не операционка
    dirs = {}      # направление -> {"ФОТ": x, "Материалы": y}
    other = {}     # человеческое имя -> сумма (не разнесено)
    for n, s, _ in np_items:
        if "ВЫРУЧКА" in n.upper():
            continue
        if n.startswith("(без подразд.:"):
            raw = n.split(":", 1)[1].strip(" )")
            if raw in INTERNAL:
                continue
            hn = HUMAN.get(raw, raw)
            other[hn] = other.get(hn, 0.0) + s
            continue
        if " · " in n and not n.startswith("("):
            comp, dname = n.split(" · ", 1)          # «ФОТ · Полигон», «Налоги · Офис (АУП)» …
            d = DIR_OF.get(dname.strip().lower(), dname.strip())
            dirs.setdefault(d, {}).setdefault(comp.strip(), 0.0)
            dirs[d][comp.strip()] += s
            continue
        if n.startswith("("):
            other[n.strip("()")] = other.get(n.strip("()"), 0.0) + s
            continue
        d = DIR_OF.get(n.strip().lower(), n.strip())
        dirs.setdefault(d, {}).setdefault("Материалы", 0.0)
        dirs[d]["Материалы"] += s
    if np_items:
        order = ["Полигон", "Сортировочный комплекс", "Переработка стройотходов", "Офис (АУП)"]
        cards = []
        for d in order + [x for x in dirs if x not in order]:
            if d not in dirs:
                continue
            parts = dirs[d]; tot = sum(parts.values())
            rows = "".join(f"<div class=r2><span class=nm style='color:var(--muted)'>{k}</span>"
                           f"<span class=mn>{_sp(v)} ₸</span></div>"
                           for k, v in sorted(parts.items(), key=lambda x: -x[1]))
            cards.append(f"<div class=card><div class='val green'>{_sp(tot)}<span class=unit>₸</span></div>"
                         f"<div class=lbl style='margin-bottom:8px'>{d}</div>{rows}</div>")
        oth_total = sum(other.values())
        oth = ""
        if other:
            oth = (f"<div class=panel style='margin-top:14px'><h2><i class='ti ti-help-circle'></i>"
                   f"Не разнесено по направлениям · {_sp(oth_total)} ₸</h2>"
                   f"<div class=ph>налоги — кандидат на раскидку пропорционально ФОТ; услуги — по контрагентам</div>"
                   + "".join(f"<div class=r2><span class=nm style='color:var(--muted)'>{k}</span>"
                             f"<span class=mn style='color:var(--muted)'>{_sp(v)} ₸</span></div>"
                             for k, v in sorted(other.items(), key=lambda x: -x[1])) + "</div>")
        naprav_html = (f"<section id=sec-naprav class=sec>"
                       f"<div class=panel><h2><i class='ti ti-chart-pie'></i>Направления — что куда пошло</h2>"
                       f"<div class=ph>{sel} · операционные затраты из 1С (ФОТ по сотрудникам, материалы по "
                       f"ценам закупа); внутренние перекладки НЗП/закрытия скрыты</div>"
                       f"<div class=cards style='margin-bottom:0'>{cards and ''.join(cards) or ''}</div></div>"
                       f"{oth}</section>")
    else:
        naprav_html = ("<section id=sec-naprav class=sec><div class=panel><h2><i class='ti ti-chart-pie'></i>"
                       "Направления</h2><div class=muted>нет данных — в боте прогони /podr ММ.ГГГГ</div></div></section>")

    # СОРТИРОВКА — выпуск/отгрузка/остатки + себестоимость на кг
    vy = vypusk.get(sel, {}); og = otgruzka.get(sel, {})
    vy_kg = sum(vy.values())
    sort_zatr = sum(v for d, parts in dirs.items() if d == "Сортировочный комплекс" for v in parts.values())
    rows_v = "".join(f"<div class=r2><span class=nm>{k}</span><span class=mn>{_sp(v)} кг</span></div>"
                     for k, v in sorted(vy.items(), key=lambda x: -x[1])) or "<div class=muted>выпуска за месяц нет</div>"
    rows_o = "".join(f"<div class=r2><span class=nm>{k}</span><span class=mn>{_sp(v[0])} кг · {_sp(v[1])} ₸</span></div>"
                     for k, v in sorted(og.items(), key=lambda x: -x[1][1])) or "<div class=muted>отгрузок за месяц нет</div>"
    rows_s = "".join(f"<div class=r2><span class=nm>{k}</span><span class=mn>{_sp(v)} кг</span></div>"
                     for k, v, _u in ostatki if v > 0) or "<div class=muted>склад пуст</div>"
    sebeskg = f"{sort_zatr / vy_kg:,.0f}".replace(",", " ") + " ₸/кг" if (vy_kg > 0 and sort_zatr > 0) else "—"
    sort_html = (
        f"<section id=sec-sort class=sec>"
        f"<div class=panel><h2><i class='ti ti-recycle'></i>Выпуск вторсырья · {sel}</h2>"
        f"<div class=ph>лист «Сортировка_виды» весовой · всего {_sp(vy_kg)} кг · затраты направления ÷ выпуск = {sebeskg}</div>"
        f"<div class=scroll>{rows_v}</div></div>"
        f"<div class=panel><h2><i class='ti ti-truck-loading'></i>Отгрузка вторсырья · {sel}</h2>"
        f"<div class=ph>кг × прайс листа «Цены» (расчётно)</div>{rows_o}</div>"
        f"<div class=panel><h2><i class='ti ti-stack-2'></i>Остатки на складе</h2>"
        f"<div class=ph>лист «Остатки» весовой (текущие)</div>{rows_s}</div></section>")

    # ТЕХНИКА — ГСМ: положено vs выдано
    g = gsm.get(sel, {})
    rows_g = ""
    tot_v = tot_p = 0.0
    for teh, (qty, vyd, polozh) in sorted(g.items(), key=lambda x: -x[1][1]):
        tot_v += vyd; tot_p += polozh
        d = vyd - polozh
        dmark = (f"<span style='color:var(--red);font-weight:700'> · +{_sp(d)} л сверх</span>" if d > polozh * 0.05 + 5
                 else f"<span style='color:var(--green)'> · в норме</span>" if polozh else "")
        rows_g += (f"<div class=r2><span class=nm>{teh}</span>"
                   f"<span class=mn>{_sp(polozh)} л положено · {_sp(vyd)} л выдано{dmark}</span></div>")
    tech_html = (
        f"<section id=sec-tech class=sec><div class=panel><h2><i class='ti ti-truck'></i>Техника — ГСМ по нормам</h2>"
        f"<div class=ph>{sel} · путевые листы × нормы (сезон учтён) · итого положено {_sp(tot_p)} л / выдано {_sp(tot_v)} л</div>"
        f"{rows_g or '<div class=muted>путевых листов за месяц нет</div>'}</div></section>")

    # ОБЗОР — топ по каждому
    top_pol = "".join(f"<div class=r2><span class=nm>{o}</span><span class=mn>{_spt(v[0])} т</span></div>" for o, v in pol_items[:6]) or "<div class=muted>нет</div>"
    top_cat = "".join(f"<div class=r2><span class=nm>{c}</span><span class=mn>{_sp(v)} ₸</span></div>" for c, v in cat_items[:6]) or "<div class=muted>нет</div>"

    sections = (
        f"<section id=sec-obzor class=sec>"
        f"<div class=panel><h2><i class='ti ti-building-factory-2'></i>Топ компаний на полигоне · {sel}</h2><div class=ph>всего {_spt(tons)} т за {trips} рейсов</div>{top_pol}</div>"
        f"<div class=panel><h2><i class='ti ti-credit-card'></i>Крупнейшие расходы · {sel}</h2><div class=ph>чистые {_sp(rashod_total)} ₸</div>{top_cat}</div>"
        f"</section>"
        f"<section id=sec-polygon class=sec><div class=panel><h2><i class='ti ti-building-factory-2'></i>Полигон — приём отходов по компаниям</h2>"
        f"<div class=ph>{sel} · всего {_spt(tons)} т за {trips} рейсов</div><div class=scroll>{pol_rows}</div>{bill_note}</div></section>"
        f"{sort_html}"
        f"{tech_html}"
        f"{naprav_html}"
        f"<section id=sec-rashody class=sec><div class=panel><h2><i class='ti ti-credit-card'></i>Расходы по категориям</h2>"
        f"<div class=ph>{sel} · чистые {_sp(rashod_total)} ₸ (без переводов группе/аффилированным)</div><div class=scroll>{cat_rows}</div></div></section>"
        f"<section id=sec-kassa class=sec>{kassa_html}</section>")

    return render(nav=nav, mpick=mpick, kpis=kpis, sections=sections)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
