# -*- coding: utf-8 -*-
"""Кабинет руководителя ТОО «Рудный-АБАТ-2006» (полигон + сортировка).
Читает хаб Абата (пишет бот abat-bot), ничего не пишет. Стиль — как кабинет ЛТ/Едиля.
Разделы: ОБЗОР (KPI) → ПОЛИГОН (приём тонн по компаниям, «Рейсы») → РАСХОДЫ (Расходы_1С) → КАССА.
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

PAGE = r"""<!DOCTYPE html><html lang=ru><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>Абат 2006 · Кабинет руководителя</title>
<link rel=preconnect href="https://fonts.googleapis.com"><link rel=preconnect href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@500;600;700;800&family=JetBrains+Mono:wght@500;700&display=swap" rel=stylesheet>
<style>
:root{--bg:#0a1410;--panel:#11201a;--panel2:#16281f;--line:#1f322a;--txt:#eaf6ef;--muted:#9db4a8;--dim:#6f8c80;
--green:#a4c91e;--green-d:#7c9617;--gold:#d9b56a;--red:#e8705a;--blu:#5ab0e0;
--mono:'JetBrains Mono',ui-monospace,Consolas,monospace;--sans:'Montserrat',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--txt);font-family:var(--sans);-webkit-font-smoothing:antialiased}
a{color:inherit;text-decoration:none}
.layout{display:flex;min-height:100vh}
.side{width:212px;background:#0d1a15;padding:22px 14px;flex-shrink:0;border-right:1px solid var(--line);position:sticky;top:0;height:100vh}
.brand{display:flex;align-items:center;gap:10px;margin-bottom:2px}
.brand .logo{width:38px;height:38px;border-radius:11px;background:var(--panel2);display:flex;align-items:center;justify-content:center;color:var(--green);font-weight:800}
.brand .name{font-size:16px;font-weight:800;letter-spacing:.03em}.brand .name .dot{color:var(--green)}
.brand-sub{font-size:11px;color:var(--dim);margin:3px 0 26px 3px}
.nav-title{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.12em;margin:0 0 10px 6px}
.nav-btn{display:flex;align-items:center;gap:11px;width:100%;padding:11px 13px;border-radius:10px;color:var(--muted);font-size:14px;font-weight:600;margin-bottom:4px;transition:.15s}
.nav-btn:hover{color:var(--txt);background:#10201a}.nav-btn.on{background:var(--panel2);color:var(--green)}
.nav-btn svg{width:19px;height:19px;flex-shrink:0}
.main{flex:1;padding:24px 30px;min-width:0;max-width:1120px}
.top{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:12px;margin-bottom:22px}
.hi{font-size:22px;font-weight:800}
select{background:var(--panel2);color:var(--txt);border:1px solid var(--line);border-radius:10px;padding:10px 14px;font-size:14px;font-weight:700;font-family:var(--sans);cursor:pointer}
select option{background:var(--panel)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(200px,1fr));gap:14px;margin-bottom:16px}
.card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px}
.card .val{font-family:var(--mono);font-size:26px;font-weight:700;letter-spacing:-.01em}
.card .val.green{color:var(--green)}.card .val.red{color:var(--red)}.card .val.blu{color:var(--blu)}
.card .unit{font-size:14px;color:var(--muted);margin-left:3px;font-family:var(--sans)}
.card .lbl{color:var(--muted);font-size:12.5px;margin-top:9px}
.panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px;margin-bottom:16px}
.panel h2{font-size:15px;font-weight:700;margin:0 0 3px}.panel .ph{color:var(--dim);font-size:12px;margin-bottom:12px;font-family:var(--mono)}
.scroll{max-height:360px;overflow-y:auto;margin:0 -6px;padding:0 6px}
.scroll::-webkit-scrollbar{width:8px}.scroll::-webkit-scrollbar-thumb{background:var(--line);border-radius:8px}
.row{display:grid;grid-template-columns:1fr 90px 128px;align-items:center;gap:10px;padding:9px 4px;border-bottom:1px solid var(--line);font-size:13.5px}
.row:last-child{border-bottom:none}.row .nm{font-weight:600}.row .mn{text-align:right;font-family:var(--mono);color:var(--muted)}
.row .bar{grid-column:1/-1;height:4px;border-radius:3px;background:#0e1b15;overflow:hidden;margin-top:-4px}
.row2{display:grid;grid-template-columns:1fr 150px;align-items:center;gap:10px;padding:10px 4px;border-bottom:1px solid var(--line);font-size:13.5px}
.row2:last-child{border-bottom:none}.row2 .mn{text-align:right;font-family:var(--mono);color:var(--txt);font-weight:600}
.fill{height:100%;border-radius:3px}
.muted{color:var(--muted)}.foot{color:var(--dim);font-size:11px;font-family:var(--mono);margin-top:22px}
.mnav{display:none}
@media(max-width:760px){.side{display:none}.main{padding:18px}.mnav{display:block}}
</style></head><body>
<div class=layout>
<aside class=side>
  <div class=brand><div class=logo>А</div><div class=name>Абат 2006<span class=dot>.</span></div></div>
  <div class=brand-sub>полигон и сортировка</div>
  <div class=nav-title>разделы</div>
  <a class="nav-btn on" href="#top"><svg viewBox="0 0 24 24" fill=none stroke=currentColor stroke-width=2><rect x=3 y=3 width=7 height=9 rx=1/><rect x=14 y=3 width=7 height=5 rx=1/><rect x=14 y=12 width=7 height=9 rx=1/><rect x=3 y=16 width=7 height=5 rx=1/></svg>Обзор</a>
  <a class="nav-btn" href="#polygon"><svg viewBox="0 0 24 24" fill=none stroke=currentColor stroke-width=2><path d="M3 21h18"/><path d="M5 21V8l7-5 7 5v13"/><path d="M9 21v-6h6v6"/></svg>Полигон</a>
  <a class="nav-btn" href="#rashody"><svg viewBox="0 0 24 24" fill=none stroke=currentColor stroke-width=2><rect x=2 y=5 width=20 height=14 rx=2/><path d="M2 10h20"/></svg>Расходы</a>
  <a class="nav-btn" href="#kassa"><svg viewBox="0 0 24 24" fill=none stroke=currentColor stroke-width=2><path d="M21 12V7H5a2 2 0 0 1 0-4h14v4"/><path d="M3 5v14a2 2 0 0 0 2 2h16v-5"/><path d="M18 12a2 2 0 0 0 0 4h4v-4z"/></svg>Касса</a>
</aside>
<main class=main id=top>
  <div class=top><div class=hi>Здравствуйте!</div>__MPICK__</div>
  <div class=cards>__KPIS__</div>
  __POLYGON__
  __RASHODY__
  __KASSA__
  <div class=foot>Данные из 1С и весовой полигона через хаб · обновление ~1 мин</div>
</main>
</div></body></html>"""

def render(**kw):
    html = PAGE
    for t in ("MPICK", "KPIS", "POLYGON", "RASHODY", "KASSA"):
        html = html.replace("__%s__" % t, kw.get(t.lower(), ""))
    return html

@app.route("/health")
def health(): return "ok"

@app.route("/")
@requires_auth
def home():
    reisy = read_reisy(); rashody = read_rashody(); kassa = read_kassa()
    months = sorted(set(list(reisy) + list(rashody) + list(kassa)), key=_mkey_sort, reverse=True)
    if not months:
        return render(polygon="<div class=panel><div class=muted>В хабе ещё нет данных. В боте: /rashody1c ММ.ГГГГ и дождись синка полигона.</div></div>")
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

    # ПОЛИГОН — компании × тонны
    pol_items = sorted(pol.items(), key=lambda x: -x[1][0])
    mx = max([v[0] for _, v in pol_items], default=1)
    if pol_items:
        rows = "".join(
            f"<div class=row><span class=nm>{o}</span><span class=mn>{v[1]} рейс</span>"
            f"<span class=mn>{_spt(v[0])} т</span>"
            f"<span class=bar><span class=fill style='width:{max(3,v[0]/mx*100):.0f}%;background:var(--green)'></span></span></div>"
            for o, v in pol_items)
        polygon = (f"<div class=panel id=polygon><h2>🏭 Полигон — приём отходов по компаниям</h2>"
                   f"<div class=ph>{sel} · всего {_spt(tons)} т за {trips} рейсов</div>"
                   f"<div class=scroll>{rows}</div></div>")
    else:
        polygon = ("<div class=panel id=polygon><h2>🏭 Полигон — приём по компаниям</h2>"
                   "<div class=muted>за этот месяц рейсов в хабе нет (весовая синхронизируется раз в час)</div></div>")

    # РАСХОДЫ — категории
    cat_items = sorted([(c, v) for c, v in cats.items() if is_expense(c)], key=lambda x: -x[1])
    if cat_items:
        rows = "".join(f"<div class=row2><span class=nm>{c}</span><span class=mn>{_sp(v)} ₸</span></div>" for c, v in cat_items)
        rashody_html = (f"<div class=panel id=rashody><h2>💸 Расходы по категориям</h2>"
                        f"<div class=ph>{sel} · чистые {_sp(rashod_total)} ₸ (без переводов группе/аффилированным)</div>"
                        f"<div class=scroll>{rows}</div></div>")
    else:
        rashody_html = "<div class=panel id=rashody><h2>💸 Расходы</h2><div class=muted>нет данных за месяц</div></div>"

    kassa_html = (f"<div class=panel id=kassa><h2>💵 Касса-1 (наличные из 1С)</h2>"
                  f"<div class=ph>{sel} · внутренние перемещения исключены; Касса-2 — по файлу Светы</div>"
                  f"<div class=row2><span class=nm>Приход</span><span class=mn>{_sp(k['приход'])} ₸</span></div>"
                  f"<div class=row2><span class=nm>Расход</span><span class=mn>{_sp(k['расход'])} ₸</span></div></div>")

    return render(mpick=mpick, kpis=kpis, polygon=polygon, rashody=rashody_html, kassa=kassa_html)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
