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

_BOOK = None
_CACHE = {}
_TTL = 60

def get_book():
    global _BOOK
    if _BOOK is None:
        creds = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        _BOOK = gspread.service_account_from_dict(creds).open_by_key(SPREADSHEET_ID)
    return _BOOK

def ws_values(title):
    now = time.time()
    hit = _CACHE.get(title)
    if hit and now - hit[0] < _TTL:
        return hit[1]
    try:
        vals = get_book().worksheet(title).get_all_values()
    except gspread.WorksheetNotFound:
        vals = []
    except Exception:
        return hit[1] if hit else []
    _CACHE[title] = (now, vals)
    return vals

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

def _sp(n): return f"{round(n):,}".replace(",", " ")
def _spt(n): return f"{n:,.1f}".replace(",", " ")

_NAV = [("obzor", "ti-layout-grid", "Обзор"), ("polygon", "ti-building-factory-2", "Полигон"),
        ("naprav", "ti-chart-pie", "Направления"), ("rashody", "ti-credit-card", "Расходы"),
        ("kassa", "ti-cash", "Касса")]

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
<link rel=preconnect href="https://fonts.googleapis.com"><link rel=preconnect href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel=stylesheet>
<link href="https://cdn.jsdelivr.net/npm/@tabler/icons-webfont@3.7.0/dist/tabler-icons.min.css" rel=stylesheet>
<style>
:root{--bg:#f4f7f5;--panel:#ffffff;--panel2:#e8f5ee;--line:#e2e9e4;--txt:#22302a;--muted:#6b7d73;--dim:#93a39a;
--green:#0f9d63;--red:#d9534a;--mono:'JetBrains Mono',ui-monospace,monospace;--sans:'Montserrat',system-ui,sans-serif}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--txt);font-family:var(--sans);-webkit-font-smoothing:antialiased}
.layout{display:flex;min-height:100vh}
.side{width:212px;background:#ffffff;padding:22px 14px;flex-shrink:0;border-right:1px solid var(--line);position:sticky;top:0;height:100vh}
.brand{display:flex;align-items:center;gap:10px}.brand .logo{width:38px;height:38px;border-radius:11px;background:var(--panel2);display:flex;align-items:center;justify-content:center;color:var(--green);font-weight:800}
.brand .name{font-size:16px;font-weight:800;letter-spacing:.03em}.brand .name i{color:var(--green)}
.brand-sub{font-size:11px;color:var(--dim);margin:3px 0 26px 3px}
.nav-title{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.12em;margin:0 0 10px 6px}
.nav-btn{display:flex;align-items:center;gap:11px;width:100%;padding:11px 13px;border-radius:10px;color:var(--muted);font-size:14px;font-weight:600;margin-bottom:4px;cursor:pointer;transition:.15s;border:none;background:none;font-family:var(--sans)}
.nav-btn:hover{color:var(--txt);background:#f0f5f1}.nav-btn.on{background:var(--panel2);color:var(--green)}
.nav-btn i{font-size:19px}
.main{flex:1;padding:24px 30px;min-width:0;max-width:1120px}
.top{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:22px}
.hi{font-size:22px;font-weight:800}
select{background:var(--panel2);color:var(--txt);border:1px solid var(--line);border-radius:10px;padding:10px 14px;font-size:14px;font-weight:700;font-family:var(--sans);cursor:pointer}
select option{background:var(--panel)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:18px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px}
.card .val{font-family:var(--mono);font-size:26px;font-weight:700}.card .val.green{color:var(--green)}.card .val.red{color:var(--red)}
.card .unit{font-size:14px;color:var(--muted);margin-left:3px}.card .lbl{color:var(--muted);font-size:12.5px;margin-top:9px}
.sec{display:none}.sec.on{display:block}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px;margin-bottom:16px}
.panel h2{font-size:15px;font-weight:700;margin:0 0 3px}.panel h2 i{color:var(--green);margin-right:6px}
.panel .ph{color:var(--dim);font-size:12px;margin-bottom:12px;font-family:var(--mono)}
.scroll{max-height:460px;overflow-y:auto;margin:0 -4px;padding:0 4px}
.scroll::-webkit-scrollbar{width:8px}.scroll::-webkit-scrollbar-thumb{background:var(--line);border-radius:8px}
.rw{padding:9px 2px;border-bottom:1px solid var(--line)}.rw:last-child{border-bottom:none}
.rw .t{display:flex;justify-content:space-between;font-size:13.5px}.rw .nm{font-weight:600}.rw .mn{font-family:var(--mono);color:var(--muted)}
.rw .bar{height:4px;border-radius:3px;background:#eef2ef;overflow:hidden;margin-top:5px}.rw .fill{height:100%;border-radius:3px;background:var(--green)}
.r2{display:flex;justify-content:space-between;padding:10px 2px;border-bottom:1px solid var(--line);font-size:13.5px}.r2:last-child{border-bottom:none}
.r2 .mn{font-family:var(--mono);font-weight:600}
.muted{color:var(--muted)}.foot{color:var(--dim);font-size:11px;font-family:var(--mono);margin-top:22px}
</style></head><body>
<div class=layout>
<aside class=side>
  <div class=brand><div class=logo><svg width="26" height="26" viewBox="0 0 48 48"><path d="M17 14 h9 l-3-4 M31 14 a9 9 0 0 1 6 9 l4-2 M33 34 h-9 l3 4 M17 34 a9 9 0 0 1 -6 -9 l-4 2" fill="none" stroke="#0f9d63" stroke-width="3.4" stroke-linecap="round" stroke-linejoin="round"/><circle cx="24" cy="24" r="4.5" fill="#0f9d63"/></svg></div><div class=name>Абат 2006<i class="ti ti-point-filled" style="font-size:10px"></i></div></div>
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
    months = sorted(set(list(reisy) + list(rashody) + list(kassa) + list(naprav)), key=_mkey_sort, reverse=True)
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

    kpis = (
        f"<div class=card><div class=val green>{_spt(tons)}<span class=unit>т</span></div><div class=lbl>Принято на полигон</div></div>"
        f"<div class=card><div class=val>{trips}</div><div class=lbl>Рейсов за месяц</div></div>"
        f"<div class=card><div class=val red>{_sp(rashod_total)}<span class=unit>₸</span></div><div class=lbl>Расходы (чистые)</div></div>"
        f"<div class=card><div class=val>{_sp(k['приход'])}<span class=unit>₸</span></div><div class=lbl>Касса-1 приход</div></div>")

    pol_items = sorted(pol.items(), key=lambda x: -x[1][0])
    mx = max([v[0] for _, v in pol_items], default=1)
    pol_rows = "".join(
        f"<div class=rw><div class=t><span class=nm>{o}</span><span class=mn>{v[1]} рейс · {_spt(v[0])} т</span></div>"
        f"<div class=bar><span class=fill style='width:{max(3,v[0]/mx*100):.0f}%'></span></div></div>"
        for o, v in pol_items) or "<div class=muted>за этот месяц рейсов нет (весовая синхронизируется раз в час)</div>"

    cat_items = sorted([(c, v) for c, v in cats.items() if is_expense(c)], key=lambda x: -x[1])
    cat_rows = "".join(f"<div class=r2><span class=nm>{c}</span><span class=mn>{_sp(v)} ₸</span></div>" for c, v in cat_items) \
               or "<div class=muted>нет данных</div>"

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

    # ОБЗОР — топ по каждому
    top_pol = "".join(f"<div class=r2><span class=nm>{o}</span><span class=mn>{_spt(v[0])} т</span></div>" for o, v in pol_items[:6]) or "<div class=muted>нет</div>"
    top_cat = "".join(f"<div class=r2><span class=nm>{c}</span><span class=mn>{_sp(v)} ₸</span></div>" for c, v in cat_items[:6]) or "<div class=muted>нет</div>"

    sections = (
        f"<section id=sec-obzor class=sec>"
        f"<div class=panel><h2><i class='ti ti-building-factory-2'></i>Топ компаний на полигоне · {sel}</h2><div class=ph>всего {_spt(tons)} т за {trips} рейсов</div>{top_pol}</div>"
        f"<div class=panel><h2><i class='ti ti-credit-card'></i>Крупнейшие расходы · {sel}</h2><div class=ph>чистые {_sp(rashod_total)} ₸</div>{top_cat}</div>"
        f"</section>"
        f"<section id=sec-polygon class=sec><div class=panel><h2><i class='ti ti-building-factory-2'></i>Полигон — приём отходов по компаниям</h2>"
        f"<div class=ph>{sel} · всего {_spt(tons)} т за {trips} рейсов</div><div class=scroll>{pol_rows}</div></div></section>"
        f"{naprav_html}"
        f"<section id=sec-rashody class=sec><div class=panel><h2><i class='ti ti-credit-card'></i>Расходы по категориям</h2>"
        f"<div class=ph>{sel} · чистые {_sp(rashod_total)} ₸ (без переводов группе/аффилированным)</div><div class=scroll>{cat_rows}</div></div></section>"
        f"<section id=sec-kassa class=sec><div class=panel><h2><i class='ti ti-cash'></i>Касса-1 (наличные из 1С)</h2>"
        f"<div class=ph>{sel} · внутренние перемещения исключены; Касса-2 — по файлу Светы</div>"
        f"<div class=r2><span class=nm>Приход</span><span class=mn>{_sp(k['приход'])} ₸</span></div>"
        f"<div class=r2><span class=nm>Расход</span><span class=mn>{_sp(k['расход'])} ₸</span></div></div></section>")

    return render(nav=nav, mpick=mpick, kpis=kpis, sections=sections)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
