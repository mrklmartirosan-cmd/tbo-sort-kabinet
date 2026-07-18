# -*- coding: utf-8 -*-
"""Кабинет руководителя ТОО «Рудный-АБАТ-2006» (полигон + сортировка).
Читает хаб Абата (пишет бот abat-bot). Ничего не пишет.
Разделы: ПОЛИГОН (приём тонн по компаниям из «Рейсы») → ФИНАНСЫ (расходы из «Расходы_1С») → КАССА.
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
    base = m.split()[0].strip().capitalize()
    yr = m.split()[-1] if m.split()[-1].isdigit() else "0"
    return (int(yr), _MORD.index(base) if base in _MORD else 99)
def iso_month(d):
    """'2026-06-15' -> 'Июнь 2026'. Прочее вернём как есть."""
    d = str(d).strip()
    if len(d) >= 7 and d[4] == "-" and d[:4].isdigit() and d[5:7].isdigit():
        y, mo = int(d[:4]), int(d[5:7])
        if 1 <= mo <= 12: return f"{_MORD[mo]} {y}"
    return d

_KASSA_INTERNAL = ("сдач", "зачислен", "перемещен", "внесен", "инкасс", "прочих денежных", "прочие денежные")
_NON_EXPENSE = ("группа компаний", "аффилир", "внутренний оборот", "внутр")
def is_expense(cat): return not any(k in cat.lower() for k in _NON_EXPENSE)

def read_reisy():
    """«Рейсы»: Дата|Время|Источник|Организация|Госномер|Водитель|Вес,т|… → {месяц:{орг:[тонны,рейсов]}}."""
    rows = ws_values("Рейсы")
    out = {}
    for r in rows[1:]:
        if len(r) < 7 or not str(r[0]).strip():
            continue
        mon = iso_month(r[0]); org = str(r[3]).strip() or "—"; ves = num(r[6])
        if ves <= 0:
            continue
        out.setdefault(mon, {}).setdefault(org, [0.0, 0])
        out[mon][org][0] += ves
        out[mon][org][1] += 1
    return out

def read_rashody():
    rows = ws_values("Расходы_1С")
    by_cat = {}
    for r in rows[1:]:
        if len(r) < 3 or not str(r[0]).strip():
            continue
        mon = str(r[0]).strip(); cat = str(r[2]).strip(); tot = num(r[-1])
        if not tot: continue
        by_cat.setdefault(mon, {}); by_cat[mon][cat] = by_cat[mon].get(cat, 0.0) + tot
    return by_cat

def read_kassa():
    rows = ws_values("Касса1_1С")
    out = {}
    for r in rows[1:]:
        if len(r) < 4 or not str(r[0]).strip():
            continue
        mon = str(r[0]).strip(); typ = str(r[1]).strip().lower(); st = str(r[2]).strip().lower(); s = num(r[3])
        if any(k in st or k in typ for k in _KASSA_INTERNAL):
            continue
        out.setdefault(mon, {"приход": 0.0, "расход": 0.0})
        if "приход" in typ or "пко" in typ: out[mon]["приход"] += s
        elif "расход" in typ or "рко" in typ: out[mon]["расход"] += s
    return out

def _sp(n): return f"{round(n):,}".replace(",", " ") + " ₸"
def _spt(n): return f"{n:,.1f}".replace(",", " ") + " т"

PAGE = """<!DOCTYPE html><html lang=ru><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>Абат 2006 · Кабинет руководителя</title>
<style>
:root{--bg:#0f1720;--card:#17212b;--line:#2a3a49;--txt:#e8eef4;--mut:#8fa3b5;--grn:#33c27f;--red:#f0685f;--blu:#4aa3ff;--amb:#f0b429;--pur:#9b8cff}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--txt);font:15px/1.45 -apple-system,Segoe UI,Roboto,Arial,sans-serif}
.wrap{max-width:1040px;margin:0 auto;padding:20px 16px 60px}
h1{font-size:22px;margin:0}.sub{color:var(--mut);font-size:13px}
.tabs{display:flex;gap:6px;margin:16px 0 18px;flex-wrap:wrap}
.tab{background:var(--card);border:1px solid var(--line);color:var(--mut);padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px}
.tab.on{background:var(--blu);border-color:var(--blu);color:#04121f}
.kpis{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin-bottom:18px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.kpi .l{color:var(--mut);font-size:12px;margin-bottom:6px}.kpi .v{font-size:20px;font-weight:700}
.sec{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:16px}
.sec h2{margin:0 0 4px;font-size:16px}.sec .h{color:var(--mut);font-size:12px;margin-bottom:12px}
.bar{display:grid;grid-template-columns:220px 1fr 150px;align-items:center;gap:10px;margin:7px 0;font-size:13.5px}
.bar .track{background:#0e161e;border-radius:6px;height:16px;overflow:hidden}.bar .fill{height:100%;border-radius:6px}
.bar .amt{text-align:right;font-variant-numeric:tabular-nums;color:var(--mut)}
.muted{color:var(--mut)}.foot{color:var(--mut);font-size:12px;margin-top:24px;text-align:center}
@media(max-width:760px){.kpis{grid-template-columns:repeat(2,1fr)}.bar{grid-template-columns:120px 1fr 92px}}
</style></head><body><div class=wrap>
<h1>Абат 2006 · Полигон и сортировка</h1><div class=sub>кабинет руководителя · ТОО «Рудный-АБАТ-2006»</div>
<div class=tabs>__TABS__</div>
<div class=kpis>__KPIS__</div>
__POLYGON__
__RASHODY__
__KASSA__
<div class=foot>Абат 2006 · данные из 1С и весовой полигона через хаб (обновление ~1 мин)</div>
</div></body></html>"""

def render(**kw):
    html = PAGE
    for t in ("TABS", "KPIS", "POLYGON", "RASHODY", "KASSA"):
        html = html.replace("__%s__" % t, kw.get(t.lower(), ""))
    return html

def bars(items, color, unit="money"):
    items = [(k, v) for k, v in items if v > 0]
    items.sort(key=lambda x: -x[1])
    mx = max([v for _, v in items], default=1)
    if not items:
        return "<div class=muted>нет данных</div>"
    fmt = _spt if unit == "t" else _sp
    return "".join(
        f"<div class=bar><span>{k}</span><span class=track>"
        f"<span class=fill style='width:{max(2, v/mx*100):.0f}%;background:{color}'></span></span>"
        f"<span class=amt>{fmt(v)}</span></div>" for k, v in items)

@app.route("/health")
def health(): return "ok"

@app.route("/")
@requires_auth
def home():
    reisy = read_reisy(); rashody = read_rashody(); kassa = read_kassa()
    months = sorted(set(list(reisy) + list(rashody) + list(kassa)), key=_mkey_sort)
    if not months:
        return render(rashody="<div class=sec><div class=muted>В хабе ещё нет данных. В боте: /rashody1c ММ.ГГГГ и подожди синк полигона.</div></div>")
    sel = request.args.get("m") or months[-1]
    if sel not in months: sel = months[-1]

    pol = reisy.get(sel, {})
    tons = sum(v[0] for v in pol.values()); trips = sum(v[1] for v in pol.values())
    cats = rashody.get(sel, {})
    rashod_total = sum(v for c, v in cats.items() if is_expense(c))
    k = kassa.get(sel, {"приход": 0.0, "расход": 0.0})

    tabs = "".join(f"<a class='tab {'on' if m==sel else ''}' href='?m={m}'>{m.split()[0]}</a>" for m in months)
    kpis = (
        f"<div class=kpi><div class=l>Принято на полигон</div><div class=v>{_spt(tons)}</div></div>"
        f"<div class=kpi><div class=l>Рейсов за месяц</div><div class=v>{trips}</div></div>"
        f"<div class=kpi><div class=l>Расходы (чистые)</div><div class=v>{_sp(rashod_total)}</div></div>"
        f"<div class=kpi><div class=l>Касса-1 приход</div><div class=v>{_sp(k['приход'])}</div></div>")

    if pol:
        pol_bars = bars([(o, v[0]) for o, v in pol.items()], "var(--blu)", unit="t")
        polygon = (f"<div class=sec><h2>🏭 Полигон — приём отходов по компаниям · {sel}</h2>"
                   f"<div class=h>тонн завезено на полигон (с весов), {trips} рейсов</div>{pol_bars}</div>")
    else:
        polygon = ("<div class=sec><h2>🏭 Полигон — приём по компаниям</h2>"
                   "<div class=muted>за этот месяц рейсов в хабе нет (весовая синхронизируется ботом раз в час)</div></div>")

    rashody_html = (f"<div class=sec><h2>💸 Расходы по категориям · {sel}</h2>"
                    f"<div class=h>чистые расходы {_sp(rashod_total)} (без переводов группе и аффилированным)</div>"
                    f"{bars([(c,v) for c,v in cats.items() if is_expense(c)], 'var(--red)')}</div>")

    kassa_html = (f"<div class=sec><h2>💵 Касса-1 (наличные из 1С) · {sel}</h2>"
                  f"<div class=h>приход {_sp(k['приход'])} · расход {_sp(k['расход'])} "
                  f"(внутренние перемещения исключены; Касса-2 — по файлу Светы)</div></div>")

    return render(tabs=tabs, kpis=kpis, polygon=polygon, rashody=rashody_html, kassa=kassa_html)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
