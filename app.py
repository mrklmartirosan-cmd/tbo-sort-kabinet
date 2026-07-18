# -*- coding: utf-8 -*-
"""Кабинет руководителя ТОО «Рудный-АБАТ-2006» (полигон + сортировка).
Читает хаб Абата (тот же, что пишет бот abat-bot): «Расходы_1С», «Касса1_1С», при наличии —
«Выручка», «Направление». Ничего не пишет. Деплой: Railway abat-kabinet, env:
SPREADSHEET_ID (хаб Абата), GOOGLE_CREDENTIALS (сервис-аккаунт ltrading-bot@l-trading), KAB_LOGIN, KAB_PASSWORD.
v1: расходы по категориям + касса-1 (net). Выручка (с 6010) и разрез по 3 направлениям — по мере записи в хаб.
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
_CACHE = {}          # title -> (ts, values)
_TTL = 60

def get_book():
    global _BOOK
    if _BOOK is None:
        creds = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        gc = gspread.service_account_from_dict(creds)
        _BOOK = gc.open_by_key(SPREADSHEET_ID)
    return _BOOK

def ws_values(title):
    """get_all_values с кэшем и мягким фолбэком (не падаем, если листа нет)."""
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

def _check(login, pwd):
    return bool(KAB_LOGIN) and login == KAB_LOGIN and pwd == KAB_PASSWORD

def requires_auth(f):
    @wraps(f)
    def w(*a, **k):
        if not KAB_LOGIN or not KAB_PASSWORD:
            return Response("Вход не настроен: задайте KAB_LOGIN и KAB_PASSWORD.", 503)
        auth = request.authorization
        if not auth or not _check(auth.username, auth.password):
            return Response("Требуется вход", 401, {"WWW-Authenticate": 'Basic realm="Abat 2006"'})
        return f(*a, **k)
    return w

def num(x):
    s = str(x or "").replace("\xa0", "").replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return 0.0

# порядок месяцев для селектора
_MORD = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
         "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]
def _mkey(m):
    base = m.split()[0].strip().capitalize()
    return (_MORD.index(base) if base in _MORD else 99, m)

# внутренние движения кассы (не доход/не расход) — исключаем из net
_KASSA_INTERNAL = ("сдач", "зачислен", "перемещен", "внесен", "инкасс", "прочих денежных", "прочие денежные")

def read_rashody():
    """«Расходы_1С»: Месяц|Контрагент|Категория|…|Итого. → {месяц: {категория: сумма}}, {месяц:{контрагент:сумма}}."""
    rows = ws_values("Расходы_1С")
    by_cat = {}; by_kontr = {}
    for r in rows[1:]:
        if len(r) < 3 or not str(r[0]).strip():
            continue
        mon = str(r[0]).strip(); kontr = str(r[1]).strip(); cat = str(r[2]).strip(); tot = num(r[-1])
        if not tot:
            continue
        by_cat.setdefault(mon, {}); by_cat[mon][cat] = by_cat[mon].get(cat, 0.0) + tot
        by_kontr.setdefault(mon, {}); by_kontr[mon][kontr] = by_kontr[mon].get(kontr, 0.0) + tot
    return by_cat, by_kontr

def read_kassa():
    """«Касса1_1С»: Месяц|Тип|Статья|Сумма. → {месяц: {'приход':x,'расход':y}} без внутренних движений."""
    rows = ws_values("Касса1_1С")
    out = {}
    for r in rows[1:]:
        if len(r) < 4 or not str(r[0]).strip():
            continue
        mon = str(r[0]).strip(); typ = str(r[1]).strip().lower(); st = str(r[2]).strip().lower(); s = num(r[3])
        if any(k in st or k in typ for k in _KASSA_INTERNAL):
            continue   # внутренний перегон — не доход/расход
        out.setdefault(mon, {"приход": 0.0, "расход": 0.0})
        if "приход" in typ or "пко" in typ:
            out[mon]["приход"] += s
        elif "расход" in typ or "рко" in typ:
            out[mon]["расход"] += s
    return out

def read_vyruchka():
    """«Выручка» (если бот записал): Месяц|Направление|Сумма. → {месяц:{направление:сумма}}. Пусто — вернём {}."""
    rows = ws_values("Выручка")
    out = {}
    for r in rows[1:]:
        if len(r) < 3 or not str(r[0]).strip():
            continue
        mon = str(r[0]).strip(); direc = str(r[1]).strip() or "—"; s = num(r[2])
        out.setdefault(mon, {}); out[mon][direc] = out[mon].get(direc, 0.0) + s
    return out

# категории, которые НЕ являются «чистым расходом» (переводы группы/внутр.оборот) — прячем из итога
_NON_EXPENSE = ("группа компаний", "внутренний оборот", "внутр")

def is_expense(cat):
    c = cat.lower()
    return not any(k in c for k in _NON_EXPENSE)

PAGE = """<!DOCTYPE html><html lang=ru><head><meta charset=utf-8>
<meta name=viewport content="width=device-width,initial-scale=1">
<title>Абат 2006 · Кабинет руководителя</title>
<style>
:root{--bg:#0f1720;--card:#17212b;--line:#2a3a49;--txt:#e8eef4;--mut:#8fa3b5;--grn:#33c27f;--red:#f0685f;--blu:#4aa3ff;--pur:#9b8cff}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--txt);font:15px/1.45 -apple-system,Segoe UI,Roboto,Arial,sans-serif}
.wrap{max-width:1040px;margin:0 auto;padding:20px 16px 60px}
h1{font-size:22px;margin:0}.sub{color:var(--mut);font-size:13px}
.banner{background:#20303f;border:1px solid var(--line);border-left:3px solid var(--blu);border-radius:8px;padding:9px 12px;color:#cfe0ee;font-size:12.5px;margin:12px 0 18px}
.tabs{display:flex;gap:6px;margin:14px 0 18px;flex-wrap:wrap}
.tab{background:var(--card);border:1px solid var(--line);color:var(--mut);padding:8px 16px;border-radius:8px;text-decoration:none;font-weight:600;font-size:14px}
.tab.on{background:var(--blu);border-color:var(--blu);color:#04121f}
.kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px}
.kpi{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:14px 16px}
.kpi .l{color:var(--mut);font-size:12.5px;margin-bottom:6px}.kpi .v{font-size:21px;font-weight:700}
.sec{background:var(--card);border:1px solid var(--line);border-radius:12px;padding:16px;margin-bottom:18px}
.sec h2{margin:0 0 14px;font-size:16px}
.bar{display:grid;grid-template-columns:210px 1fr 130px;align-items:center;gap:10px;margin:7px 0;font-size:13.5px}
.bar .track{background:#0e161e;border-radius:6px;height:16px;overflow:hidden}.bar .fill{height:100%;border-radius:6px}
.bar .amt{text-align:right;font-variant-numeric:tabular-nums;color:var(--mut)}
.pos{color:var(--grn)}.neg{color:var(--red)}.muted{color:var(--mut)}
.foot{color:var(--mut);font-size:12px;margin-top:24px;text-align:center}
@media(max-width:720px){.kpis{grid-template-columns:1fr}.bar{grid-template-columns:130px 1fr 96px}}
</style></head><body><div class=wrap>
<h1>Абат 2006 · Полигон и сортировка</h1><div class=sub>кабинет руководителя · ТОО «Рудный-АБАТ-2006»</div>
{banner}
<div class=tabs>{tabs}</div>
<div class=kpis>{kpis}</div>
<div class=sec><h2>Расходы по категориям · {mon}</h2>{rashody}</div>
{vyr}
<div class=foot>Абат 2006 · кабинет читает хаб (обновление ~1 мин) · v1</div>
</div></body></html>"""

def _sp(n):
    return f"{round(n):,}".replace(",", " ") + " ₸"

def bars(items, color):
    items = [(k, v) for k, v in items if v > 0]
    items.sort(key=lambda x: -x[1])
    mx = max([v for _, v in items], default=1)
    if not items:
        return "<div class=muted>нет данных</div>"
    out = []
    for k, v in items:
        out.append(f"<div class=bar><span>{k}</span><span class=track>"
                   f"<span class=fill style='width:{max(2, v / mx * 100):.0f}%;background:{color}'></span></span>"
                   f"<span class=amt>{_sp(v)}</span></div>")
    return "".join(out)

@app.route("/")
@requires_auth
def home():
    by_cat, _ = read_rashody()
    kassa = read_kassa()
    vyr = read_vyruchka()
    months = sorted(set(list(by_cat) + list(kassa) + list(vyr)), key=_mkey)
    if not months:
        return PAGE.format(banner="<div class=banner>В хабе ещё нет данных. Запусти в боте /rashody1c ММ.ГГГГ.</div>",
                           tabs="", kpis="", mon="—", rashody="<div class=muted>нет данных</div>", vyr="")
    sel = request.args.get("m") or months[-1]
    if sel not in months:
        sel = months[-1]
    cats = by_cat.get(sel, {})
    rashod_total = sum(v for c, v in cats.items() if is_expense(c))
    k = kassa.get(sel, {"приход": 0.0, "расход": 0.0})
    vy = vyr.get(sel, {})
    vy_total = sum(vy.values())

    tabs = "".join(f"<a class='tab {'on' if m == sel else ''}' href='?m={m}'>{m.split()[0]}</a>" for m in months)
    kpis = (f"<div class=kpi><div class=l>Расходы (чистые)</div><div class=v>{_sp(rashod_total)}</div></div>"
            f"<div class=kpi><div class=l>Касса-1 приход / расход</div><div class=v>{_sp(k['приход'])} <span class=muted style='font-size:14px'>/ {_sp(k['расход'])}</span></div></div>")
    if vy_total:
        kpis += f"<div class=kpi><div class=l>Выручка (1С, сч.6010)</div><div class=v>{_sp(vy_total)}</div></div>"
    else:
        kpis += "<div class=kpi><div class=l>Выручка</div><div class=v class=muted style='font-size:15px'>подключается</div></div>"

    rashody = bars([(c, v) for c, v in cats.items() if is_expense(c)], "var(--red)")
    vyr_html = ""
    if vy_total:
        vyr_html = f"<div class=sec><h2>Выручка по направлениям · {sel}</h2>{bars(list(vy.items()), 'var(--grn)')}</div>"

    banner = ("<div class=banner>✅ Данные из 1С через хаб (расходы сверены до тенге). "
              "Разрез по 3 направлениям и выручка — по мере записи ботом.</div>")
    return PAGE.format(banner=banner, tabs=tabs, kpis=kpis, mon=sel, rashody=rashody, vyr=vyr_html)

@app.route("/health")
def health():
    return "ok"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8080)))
