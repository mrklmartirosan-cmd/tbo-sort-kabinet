import os
import json
from functools import wraps
from collections import defaultdict
from flask import Flask, jsonify, request, Response
import gspread
app = Flask(__name__)
SPREADSHEET_ID = os.environ.get("SPREADSHEET_ID")
KAB_LOGIN = os.environ.get("KAB_LOGIN", "")
KAB_PASSWORD = os.environ.get("KAB_PASSWORD", "")
def check_auth(login, password):
    return login == KAB_LOGIN and password == KAB_PASSWORD
def need_auth():
    return Response(
        "Нужен вход. Доступ только для руководства.",
        401,
        {"WWW-Authenticate": 'Basic realm="Kabinet"'},
    )
def requires_auth(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not KAB_LOGIN or not KAB_PASSWORD:
            return Response("Вход не настроен. Задайте KAB_LOGIN и KAB_PASSWORD в настройках.", 503)
        auth = request.authorization
        if not auth or not check_auth(auth.username, auth.password):
            return need_auth()
        return f(*args, **kwargs)
    return decorated
def get_book():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    creds_dict = json.loads(creds_json)
    gc = gspread.service_account_from_dict(creds_dict)
    return gc.open_by_key(SPREADSHEET_ID)
# --- Настройки клиента (лист «Настройки») — каркас «коробки» ---
# Если листа/значений нет (как у Еділ) — всё работает на значениях по умолчанию.
SETTINGS_DEFAULTS = {
    "company": "ТОО «Еділ и компания»",
    "brand": "ЕДІЛ",
    "kabinet_title": "Кабинет · Едиль",
    "subtitle": "кабинет руководителя",
}
def get_settings():
    """Читает лист «Настройки» (две колонки: параметр | значение) и сопоставляет
    человеческие подписи с ключами. Любая ошибка/отсутствие листа → значения по умолчанию."""
    s = dict(SETTINGS_DEFAULTS)
    raw = {}
    try:
        for r in get_book().worksheet("Настройки").get_all_values():
            if len(r) >= 2 and str(r[0]).strip() and str(r[1]).strip():
                raw[str(r[0]).strip().lower()] = str(r[1]).strip()
    except Exception:
        return s
    def find(*subs):
        for k, v in raw.items():
            if any(sub in k for sub in subs):
                return v
        return None
    mapping = {
        "company": ("название компании", "компани"),
        "brand": ("бренд",),
        "kabinet_title": ("заголовок кабинета", "заголовок"),
        "subtitle": ("подзаголовок",),
    }
    for key, subs in mapping.items():
        v = find(*subs)
        if v:
            s[key] = v
    return s
def to_number(value):
    if value is None:
        return 0.0
    s = str(value).strip().replace(" ", "").replace(" ", "").replace(",", ".")
    if s == "":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0
def parse_ymd(value):
    s = str(value).strip()
    if not s:
        return None
    for sep in (".", "/"):
        if sep in s:
            parts = s.split(sep)
            if len(parts) == 3:
                try:
                    d = int(parts[0]); m = int(parts[1]); y = int(parts[2])
                    if y < 100:
                        y += 2000
                    return (y, m, d)
                except ValueError:
                    pass
    if "-" in s:
        parts = s.split("-")
        if len(parts) == 3:
            try:
                y = int(parts[0]); m = int(parts[1]); d = int(parts[2])
                return (y, m, d)
            except ValueError:
                pass
    return None
@app.route("/")
def home():
    return "Кабинет жив. Вход в данные защищён паролем. Откройте /kabinet"
@app.route("/proverka")
@requires_auth
def proverka():
    try:
        book = get_book()
        ws = book.worksheet("Реализация")
        rows = ws.get_all_values()[:6]
        return jsonify({"статус": "связь есть", "лист": "Реализация", "первые_строки": rows})
    except Exception as e:
        return jsonify({"статус": "ошибка", "что_случилось": str(e)})
@app.route("/data")
@requires_auth
def data():
    """Отдаём все продажи в удобном для витрины виде (с датами, для фильтрации)."""
    try:
        book = get_book()
        ws = book.worksheet("Реализация")
        all_rows = ws.get_all_values()[1:]
        records = []
        for row in all_rows:
            if len(row) < 10:
                row = row + [""] * (10 - len(row))
            дата, покупатель, тип, фракция, кол_кг, цена, сумма_ндс, ндс, прим, оплачено = row[:10]
            if str(покупатель).strip() == "" and to_number(кол_кг) == 0:
                continue
            ymd = parse_ymd(дата)
            records.append({
                "y": ymd[0] if ymd else None,
                "m": ymd[1] if ymd else None,
                "d": ymd[2] if ymd else None,
                "buyer": str(покупатель).strip() or "(без имени)",
                "pay": str(тип).strip() or "(не указан)",
                "frac": str(фракция).strip() or "(без фракции)",
                "kg": to_number(кол_кг),
                "sum_vat": to_number(сумма_ндс),
                "vat": to_number(ндс),
                "paid": bool(str(оплачено).strip()),
                "paid_date": str(оплачено).strip(),
            })
        return jsonify({"ok": True, "records": records})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})
@app.route("/data-proizvodstvo")
@requires_auth
def data_proizvodstvo():
    """Отдаём производство: дата, фракции, всего крошки, металлокорд, вес шин.
    Лист Производство: A Дата | B Оператор | C Вес шин | D Мешки | E Нитки |
    F-J Фракции 0-1..6-8 | K Всего крошки | L Металлокорд | M Примечание.
    Строку 'Начальный остаток' в произведённое НЕ считаем — это стартовый остаток, не выработка."""
    try:
        book = get_book()
        ws = book.worksheet("Производство")
        all_rows = ws.get_all_values()[1:]
        records = []
        for row in all_rows:
            if len(row) < 13:
                row = row + [""] * (13 - len(row))
            дата, оператор, вес_шин, мешки, нитки, f01, f12, f24, f46, f68, всего, металл, прим = row[:13]
            # пропускаем стартовый остаток и пустые строки
            if "начальн" in str(оператор).strip().lower():
                continue
            всего_кг = to_number(всего)
            if всего_кг == 0 and to_number(вес_шин) == 0:
                continue
            ymd = parse_ymd(дата)
            records.append({
                "y": ymd[0] if ymd else None,
                "m": ymd[1] if ymd else None,
                "d": ymd[2] if ymd else None,
                "tyres": to_number(вес_шин),
                "f01": to_number(f01),
                "f12": to_number(f12),
                "f24": to_number(f24),
                "f46": to_number(f46),
                "f68": to_number(f68),
                "total": всего_кг,
                "metal": to_number(металл),
            })
        return jsonify({"ok": True, "records": records})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})
@app.route("/data-ostatok")
@requires_auth
def data_ostatok():
    """Текущий остаток на складе из листа-витрины 'Остаток' (считается формулами, со стартовым остатком).
    Лист Остаток: A Фракция | B Произведено | C Продано | D Остаток.
    Строки 2..6 = фракции 0-1,1-2,2-4,4-6,6-8; строка 7 = ИТОГО."""
    try:
        book = get_book()
        ws = book.worksheet("Остаток")
        all_rows = ws.get_all_values()[1:]
        fracs = []
        total = 0.0
        for row in all_rows:
            if len(row) < 4:
                row = row + [""] * (4 - len(row))
            фракция, произв, продано, остаток = row[:4]
            name = str(фракция).strip()
            if name == "":
                continue
            # строку ИТОГО не кладём в список фракций — общий считаем сами
            if "итог" in name.lower():
                continue
            ost = to_number(остаток)
            fracs.append({"frac": name, "made": to_number(произв), "sold": to_number(продано), "left": ost})
            total += ost
        return jsonify({"ok": True, "fracs": fracs, "total": total})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})
@app.route("/data-rashody")
@requires_auth
def data_rashody():
    """Расходы из листа «Расходы»: A Дата|B Сумма|C Группа|D Категория|E Нал/Безнал|F Источник|G Контрагент|H Примечание."""
    try:
        book = get_book()
        try:
            ws = book.worksheet("Расходы")
        except gspread.WorksheetNotFound:
            return jsonify({"ok": True, "records": []})
        all_rows = ws.get_all_values()[1:]
        records = []
        for row in all_rows:
            if len(row) < 8:
                row = row + [""] * (8 - len(row))
            дата, сумма, группа, категория, нб, источник, контрагент, прим = row[:8]
            s = to_number(сумма)
            if s == 0 and str(контрагент).strip() == "":
                continue
            ymd = parse_ymd(дата)
            records.append({
                "y": ymd[0] if ymd else None,
                "m": ymd[1] if ymd else None,
                "d": ymd[2] if ymd else None,
                "sum": s,
                "group": str(группа).strip() or "(без группы)",
                "cat": str(категория).strip() or "(без категории)",
                "pay": str(нб).strip() or "(не указан)",
                "src": str(источник).strip(),
                "contr": str(контрагент).strip() or "(без контрагента)",
            })
        return jsonify({"ok": True, "records": records})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})
@app.route("/data-tabel")
@requires_auth
def data_tabel():
    """Табель из листа «Табель»: A Дата | B Сотрудник | C Должность | D Часы | E Примечание."""
    try:
        book = get_book()
        try:
            ws = book.worksheet("Табель")
        except gspread.WorksheetNotFound:
            return jsonify({"ok": True, "records": []})
        records = []
        for row in ws.get_all_values()[1:]:
            if len(row) < 4:
                row = row + [""] * (4 - len(row))
            emp = str(row[1]).strip()
            if not emp:
                continue
            ymd = parse_ymd(row[0])
            records.append({
                "y": ymd[0] if ymd else None,
                "m": ymd[1] if ymd else None,
                "d": ymd[2] if ymd else None,
                "emp": emp,
                "role": str(row[2]).strip(),
                "hours": to_number(row[3]),
            })
        return jsonify({"ok": True, "records": records})
    except Exception as e:
        return jsonify({"ok": False, "error": str(e)})
# --- Финансы: файл экономиста (Калькуляция, Выводы, Анализ) ---
SEBES_FILE_ID = os.environ.get("SEBES_FILE_ID", "1ZKYCFVKrb0l-mzYQ0gtiHKOJbTNHd9TcJ9hN9i5RiFk")
KALK_SHEET = "Калькуляция себестоимости"
VYV_SHEET = "Выводы (5 мес)"
ANAL_SHEET = "Анализ и выводы"
RU_MONTHS_LIST = ["январь", "февраль", "март", "апрель", "май", "июнь",
                  "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"]
def _fin_norm(s):
    return " ".join(str(s).split()).strip().lower()
def _row_starts(grid, needle, start=0, end=None):
    needle = needle.lower()
    e = len(grid) if end is None else end
    for i in range(start, e):
        a = _fin_norm(grid[i][0] if (i < len(grid) and grid[i]) else "")
        if a.startswith(needle):
            return i
    return None
def _row_contains(grid, needle, start=0, end=None):
    needle = needle.lower()
    e = len(grid) if end is None else end
    for i in range(start, e):
        a = _fin_norm(grid[i][0] if (i < len(grid) and grid[i]) else "")
        if needle in a:
            return i
    return None
def get_sebes_book():
    creds_json = os.environ.get("GOOGLE_CREDENTIALS")
    creds_dict = json.loads(creds_json)
    gc = gspread.service_account_from_dict(creds_dict)
    return gc.open_by_key(SEBES_FILE_ID)
@app.route("/data-fin")
@requires_auth
def data_fin():
    """Финансы из файла экономиста:
    - помесячно поступления/расходы/прибыль (лист «Анализ и выводы», месяцы — строки);
    - себестоимость 1 кг по месяцам (лист «Калькуляция», месяцы — столбцы);
    - структура расходов за период (лист «Выводы», столбец «Итого»).
    Парсинг защищённый: если лист/строка не найдены — отдаём что есть, без падения."""
    out = {"ok": True, "months": [], "sebes": [], "structure": [],
           "expense_articles": [], "expense_months": []}
    try:
        book = get_sebes_book()
    except Exception as e:
        return jsonify({"ok": False, "error": "Нет доступа к файлу экономиста: " + str(e)})
    # 1) помесячно: поступления / расходы / прибыль — лист «Анализ и выводы»
    try:
        ag = book.worksheet(ANAL_SHEET).get_all_values()
        for i, row in enumerate(ag):
            if not row:
                continue
            name = str(row[0]).strip()
            if _fin_norm(name) not in RU_MONTHS_LIST:
                continue
            r = row + [""] * (5 - len(row)) if len(row) < 5 else row
            inc = to_number(r[1]); exp = to_number(r[2])
            # разницу считаем сами (колонка «разница» в исходнике по части месяцев пустая)
            out["months"].append({"month": name, "income": inc, "expense": exp, "profit": inc - exp})
    except Exception as e:
        out["months_error"] = str(e)
    # 2) себестоимость 1 кг — лист «Калькуляция», месяцы в строке 2
    try:
        kg = book.worksheet(KALK_SHEET).get_all_values()
        hdr = kg[1] if len(kg) > 1 else []
        month_cols = []
        for j, h in enumerate(hdr):
            if j == 0:
                continue
            if _fin_norm(h) in RU_MONTHS_LIST:
                month_cols.append((j, str(h).strip()))
        # себестоимость 1 кг бот пишет в фикс. строки: 29 — производственная, 30 — полная
        # (читаем по номеру строки, а не по названию — названия в шаблоне могут отличаться).
        r_made = _row_contains(kg, "произведено крошки")
        IDX_MADE = r_made if r_made is not None else 4   # стр. 5 «Произведено крошки» (запас)
        IDX_PROD1, IDX_FULL1 = 28, 29                     # стр. 29 и 30 (0-based 28/29)
        def _cell(ri, j):
            return to_number(kg[ri][j]) if (ri is not None and 0 <= ri < len(kg) and j < len(kg[ri])) else 0
        for j, mname in month_cols:
            made = _cell(IDX_MADE, j); p1 = _cell(IDX_PROD1, j); f1 = _cell(IDX_FULL1, j)
            if made == 0 and p1 == 0 and f1 == 0:
                continue
            out["sebes"].append({"month": mname, "kg": made, "prod1": p1, "full1": f1})
    except Exception as e:
        out["sebes_error"] = str(e)
    # 3) структура расходов — лист «Выводы», столбец «Итого»
    try:
        vg = book.worksheet(VYV_SHEET).get_all_values()
        hdr = vg[1] if len(vg) > 1 else []
        itog_col = None
        for j, h in enumerate(hdr):
            if "итог" in _fin_norm(h):
                itog_col = j
                break
        if itog_col is not None:
            labels = ["фот", "лизинг", "гсм", "материалы", "транспортные",
                      "основные средства", "прочие расходы", "налоги"]
            seen = set()
            for lab in labels:
                r = _row_contains(vg, lab)
                if r is None or r in seen:
                    continue
                seen.add(r)
                val = to_number(vg[r][itog_col]) if itog_col < len(vg[r]) else 0
                if val:
                    out["structure"].append({"name": str(vg[r][0]).strip(), "sum": val})
    except Exception as e:
        out["structure_error"] = str(e)
    # 4) расходы по статьям × месяцам — лист «Выводы (5 мес)» (помесячные столбцы)
    try:
        vg = book.worksheet(VYV_SHEET).get_all_values()
        hdr = vg[1] if len(vg) > 1 else []
        emonths = []
        for j, h in enumerate(hdr):
            hn = _fin_norm(h)
            mn = next((m for m in RU_MONTHS_LIST if hn == m or hn.startswith(m + " ") or hn.startswith(m)), None)
            if mn:
                emonths.append((j, mn.capitalize()))
        out["expense_months"] = [m for _, m in emonths]
        labels = ["фот", "лизинг", "гсм", "материалы", "транспортные",
                  "основные средства", "прочие расходы", "налоги"]
        seen = set()
        for lab in labels:
            r = _row_contains(vg, lab)
            if r is None or r in seen:
                continue
            seen.add(r)
            art = {"name": str(vg[r][0]).strip(), "by_month": [], "total": 0.0}
            for j, mname in emonths:
                v = to_number(vg[r][j]) if j < len(vg[r]) else 0
                art["by_month"].append(v)
                art["total"] += v
            if art["total"] or any(art["by_month"]):
                out["expense_articles"].append(art)
    except Exception as e:
        out["expense_articles_error"] = str(e)
    return jsonify(out)
@app.route("/kabinet")
@requires_auth
def kabinet():
    s = get_settings()
    html = (DASHBOARD_HTML
            .replace("{{KAB_TITLE}}", s["kabinet_title"])
            .replace("{{BRAND}}", s["brand"])
            .replace("{{SUBTITLE}}", s["subtitle"]))
    return Response(html, mimetype="text/html")
DASHBOARD_HTML = r"""<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{{KAB_TITLE}}</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link rel="preconnect" href="https://fonts.gstatic.com" crossorigin>
<link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;500;600;700;800&family=JetBrains+Mono:wght@500;600;700&display=swap" rel="stylesheet">
<script src="https://cdnjs.cloudflare.com/ajax/libs/Chart.js/4.4.1/chart.umd.min.js"></script>
<style>
  :root{
    --bg:#0a1410; --panel:#11201a; --panel2:#16281f; --line:#1f322a;
    --txt:#eaf6ef; --muted:#9db4a8; --dim:#6f8c80;
    --green:#4ee29c; --green-d:#2aa873; --gold:#d9b56a; --red:#e8705a;
    --mono:'JetBrains Mono',ui-monospace,Consolas,monospace;
    --sans:'Montserrat',system-ui,-apple-system,'Segoe UI',Roboto,sans-serif;
  }
  *{box-sizing:border-box}
  body{margin:0;background:var(--bg);color:var(--txt);font-family:var(--sans);
    -webkit-font-smoothing:antialiased}
  .layout{display:flex;min-height:100vh}
  /* боковое меню */
  .side{width:210px;background:#0d1a15;padding:22px 14px;flex-shrink:0;
    border-right:1px solid var(--line);position:sticky;top:0;height:100vh}
  .brand{display:flex;align-items:center;gap:9px;margin-bottom:2px}
  .brand .logo{width:32px;height:32px;border-radius:9px;background:var(--panel2);
    display:flex;align-items:center;justify-content:center;color:var(--green);font-size:18px}
  .brand .name{font-size:16px;font-weight:700;letter-spacing:.04em}
  .brand .name .dot{color:var(--green)}
  .brand-sub{font-size:11px;color:var(--dim);margin:2px 0 26px 3px}
  .nav-title{font-size:10px;color:var(--dim);text-transform:uppercase;letter-spacing:.12em;margin:0 0 10px 6px}
  .nav-btn{display:flex;align-items:center;gap:11px;width:100%;text-align:left;
    padding:11px 13px;border-radius:10px;border:none;background:transparent;color:var(--muted);
    font-size:14px;font-weight:500;font-family:var(--sans);cursor:pointer;margin-bottom:4px;transition:.15s}
  .nav-btn:hover{color:var(--txt);background:#10201a}
  .nav-btn.on{background:var(--panel2);color:var(--green)}
  .nav-btn svg{width:19px;height:19px;flex-shrink:0}
  /* основная область */
  .main{flex:1;padding:24px 28px;min-width:0;max-width:1180px}
  .top{display:flex;align-items:baseline;justify-content:space-between;flex-wrap:wrap;gap:10px;margin-bottom:20px}
  .hi{font-size:22px;font-weight:700;letter-spacing:.01em}
  .top-sub{font-size:12px;color:var(--dim);font-family:var(--mono)}
  .periods{display:flex;gap:8px;margin-bottom:22px;flex-wrap:wrap}
  .periods button{background:var(--panel2);color:var(--muted);border:1px solid var(--line);
    padding:10px 18px;border-radius:10px;font-size:14px;font-weight:600;cursor:pointer;
    font-family:var(--sans);transition:.15s}
  .periods button:hover{border-color:var(--green-d);color:var(--txt)}
  .periods button.on{background:var(--green);color:#06120c;border-color:var(--green);font-weight:700}
  .note-line{color:var(--muted);font-size:13px;font-family:var(--mono);margin:0 0 18px;
    background:var(--panel);border:1px solid var(--line);border-radius:10px;padding:11px 14px}
  .cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(190px,1fr));gap:14px;margin-bottom:14px}
  .card{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px 18px 16px;position:relative}
  .card .val{font-family:var(--mono);font-size:27px;font-weight:700;letter-spacing:-.01em}
  .card .val.green{color:var(--green)}
  .card .val.red{color:var(--red)}
  .card .unit{font-size:14px;font-weight:500;color:var(--muted);margin-left:4px;font-family:var(--sans)}
  .card .lbl{color:var(--muted);font-size:13px;margin-top:9px}
  .card .sub2{color:var(--dim);font-size:11px;font-family:var(--mono);margin-top:4px}
  .card .ic{position:absolute;top:16px;right:16px;width:36px;height:36px;border-radius:10px;
    background:var(--panel2);display:flex;align-items:center;justify-content:center;color:var(--green)}
  .card .ic svg{width:19px;height:19px}
  .grid2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-bottom:14px}
  @media(max-width:820px){.grid2{grid-template-columns:1fr}}
  .panel{background:var(--panel);border:1px solid var(--line);border-radius:14px;padding:18px}
  .panel h2{font-size:14px;color:var(--txt);font-weight:700;letter-spacing:.02em;margin:0 0 14px}
  canvas{max-height:240px}
  table{width:100%;border-collapse:collapse;font-size:13px}
  th{text-align:left;color:var(--muted);font-weight:600;font-size:11px;text-transform:uppercase;
    letter-spacing:.05em;padding:7px 8px;border-bottom:1px solid var(--line)}
  td{padding:10px 8px;border-bottom:1px solid var(--line);font-family:var(--mono);font-weight:500}
  td:first-child{font-family:var(--sans)}
  tr:last-child td{border-bottom:none}
  .num{text-align:right}
  .empty{color:var(--muted);text-align:center;padding:34px;font-size:14px}
  .err{color:var(--red);padding:16px;font-family:var(--mono);font-size:13px}
  .foot{color:var(--dim);font-size:11px;font-family:var(--mono);margin-top:20px}
  /* мобильное меню сверху */
  .mtabs{display:none;gap:6px;margin-bottom:18px;flex-wrap:wrap}
  @media(max-width:720px){
    .side{display:none}
    .mtabs{display:flex}
    .main{padding:18px}
  }
  .mtabs button{background:var(--panel2);color:var(--muted);border:1px solid var(--line);
    padding:9px 15px;border-radius:9px;font-size:13px;font-weight:600;cursor:pointer;font-family:var(--sans)}
  .mtabs button.on{background:var(--green);color:#06120c;border-color:var(--green)}
</style>
</head>
<body>
<div class="layout">
  <aside class="side">
    <div class="brand">
      <div class="logo"><svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M7 19a4 4 0 0 1-2-7.5"/><path d="M5 11.5A4 4 0 0 1 9 6a5 5 0 0 1 9.5 1.5A3.5 3.5 0 0 1 18 14H8"/><path d="M12 12v6"/><path d="M9 15l3-3 3 3"/></svg></div>
      <div class="name">{{BRAND}}<span class="dot">.</span></div>
    </div>
    <div class="brand-sub">{{SUBTITLE}}</div>
    <div class="nav-title">разделы</div>
    <button class="nav-btn on" data-s="sales">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 1v22"/><path d="M17 5H9.5a3.5 3.5 0 0 0 0 7h5a3.5 3.5 0 0 1 0 7H6"/></svg>
      Продажи
    </button>
    <button class="nav-btn" data-s="prod">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14.7 6.3a1 1 0 0 0 0 1.4l1.6 1.6a1 1 0 0 0 1.4 0l3.77-3.77a6 6 0 0 1-7.94 7.94l-6.91 6.91a2.12 2.12 0 0 1-3-3l6.91-6.91a6 6 0 0 1 7.94-7.94l-3.76 3.76z"/></svg>
      Производство
    </button>
    <button class="nav-btn" data-s="ost">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/></svg>
      Остаток
    </button>
    <button class="nav-btn" data-s="rashody">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="2" y="5" width="20" height="14" rx="2"/><path d="M2 10h20"/></svg>
      Расходы
    </button>
    <button class="nav-btn" data-s="tabel">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M16 21v-2a4 4 0 0 0-4-4H6a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M22 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>
      Табель
    </button>
    <button class="nav-btn" data-s="fin">
      <svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="m19 9-5 5-4-4-3 3"/></svg>
      Финансы
    </button>
  </aside>
  <main class="main">
    <div class="top">
      <div class="hi">Здравствуйте!</div>
      <div class="top-sub" id="updated">загрузка…</div>
    </div>
    <div class="mtabs" id="mtabs">
      <button data-s="sales" class="on">Продажи</button>
      <button data-s="prod">Производство</button>
      <button data-s="ost">Остаток</button>
      <button data-s="rashody">Расходы</button>
      <button data-s="tabel">Табель</button>
      <button data-s="fin">Финансы</button>
    </div>
    <div class="periods" id="periods">
      <button data-p="day">День</button>
      <button data-p="month" class="on">Месяц</button>
      <button data-p="quarter">Квартал</button>
      <button data-p="year">Год</button>
      <button data-p="all">Всё время</button>
    </div>
    <section id="sec-sales">
      <div id="content"><div class="empty">Загружаю данные…</div></div>
    </section>
    <section id="sec-prod" style="display:none">
      <div id="content-prod"><div class="empty">Загружаю данные…</div></div>
    </section>
    <section id="sec-ost" style="display:none">
      <div class="note-line">Остаток на сейчас, со стартовым остатком. Период на него не влияет.</div>
      <div id="content-ost"><div class="empty">Загружаю данные…</div></div>
    </section>
    <section id="sec-rashody" style="display:none">
      <div id="content-rashody"><div class="empty">Загружаю данные…</div></div>
    </section>
    <section id="sec-tabel" style="display:none">
      <div class="note-line">Табель заполняется через бота: пришлите фото табеля — записи появятся здесь.</div>
      <div id="content-tabel"><div class="empty">Загружаю данные…</div></div>
    </section>
    <section id="sec-fin" style="display:none">
      <div class="note-line">Финансы с начала года из файла экономиста (поступления, расходы, прибыль, себестоимость). Период на этот раздел не влияет — показывается весь год.</div>
      <div id="content-fin"><div class="empty">Загружаю данные…</div></div>
    </section>
    <div class="foot" id="foot"></div>
  </main>
</div>
<script>
let RAW = [];
let RAW_PROD = [];
let RAW_OST = null;
let RAW_RASH = [];
let RAW_TABEL = [];
let RAW_FIN = null;
let period = 'month';
let section = 'sales';
function fmt(n){return Math.round(n).toLocaleString('ru-RU');}
function fmtKg(n){return (Math.round(n*10)/10).toLocaleString('ru-RU');}
function inPeriod(r){
  if(period==='all') return true;
  const now = new Date();
  const Y = now.getFullYear(), M = now.getMonth()+1;
  if(!r.y || !r.m) return false;
  if(period==='day') return r.y===Y && r.m===M && r.d===now.getDate();
  if(period==='month') return r.y===Y && r.m===M;
  if(period==='year') return r.y===Y;
  if(period==='quarter'){
    const q = Math.floor((M-1)/3);
    const rq = Math.floor((r.m-1)/3);
    return r.y===Y && rq===q;
  }
  return true;
}
const IC = {
  coins:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="8" cy="8" r="6"/><path d="M18.09 10.37A6 6 0 1 1 10.34 18"/><path d="M7 6h1v4"/><path d="m16.71 13.88.7.71-2.82 2.82"/></svg>',
  weight:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="5" r="3"/><path d="M6.5 8h11l2.5 12H4z"/></svg>',
  tax:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M4 2v20l2-1 2 1 2-1 2 1 2-1 2 1 2-1 2 1V2l-2 1-2-1-2 1-2-1-2 1-2-1-2 1Z"/><path d="m9 15 6-6"/><circle cx="9.5" cy="9.5" r=".5"/><circle cx="14.5" cy="14.5" r=".5"/></svg>',
  deals:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M11 12H3"/><path d="m15 8 4 4-4 4"/><path d="M9 16H3"/><path d="M9 8H3"/><path d="M19 12h-8"/></svg>',
  metal:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 21h18"/><path d="M5 21V7l8-4v18"/><path d="M19 21V11l-6-4"/></svg>',
  tyre:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><circle cx="12" cy="12" r="4"/></svg>',
  days:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2"/><path d="M16 2v4M8 2v4M3 10h18"/></svg>',
  box:'<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16Z"/><path d="m3.3 7 8.7 5 8.7-5"/><path d="M12 22V12"/></svg>'
};
function card(val, unit, lbl, sub2, ic, cls){
  return '<div class="card"><div class="val '+(cls||'')+'">'+val+(unit?'<span class="unit">'+unit+'</span>':'')+'</div>'+
    '<div class="lbl">'+lbl+'</div>'+(sub2?'<div class="sub2">'+sub2+'</div>':'')+
    '<div class="ic">'+IC[ic]+'</div></div>';
}
const CHART_FONT = {weight:'700'};
function chartScales(line){
  return {x:{grid:{display:false},ticks:{font:{weight:'700'}}},
          y:{grid:{color:line},beginAtZero:true,ticks:{font:{weight:'600'}}}};
}
function render(){
  const rows = RAW.filter(inPeriod);
  const c = document.getElementById('content');
  if(rows.length===0){
    c.innerHTML = '<div class="panel"><div class="empty">За выбранный период продаж нет.<br>Попробуйте «Год» или «Всё время».</div></div>';
    return;
  }
  let kg=0,sumVat=0,vat=0,paidSum=0,unpaidSum=0;
  const byFrac={}, byMonth={}, byPay={}, byBuyer={}, byBuyerKg={}, byDebtor={}, debtCnt={};
  for(const r of rows){
    kg+=r.kg; sumVat+=r.sum_vat; vat+=r.vat;
    if(r.paid){ paidSum+=r.sum_vat; }
    else { unpaidSum+=r.sum_vat; byDebtor[r.buyer]=(byDebtor[r.buyer]||0)+r.sum_vat; debtCnt[r.buyer]=(debtCnt[r.buyer]||0)+1; }
    byFrac[r.frac]=(byFrac[r.frac]||0)+r.kg;
    byPay[r.pay]=(byPay[r.pay]||0)+r.sum_vat;
    byBuyer[r.buyer]=(byBuyer[r.buyer]||0)+r.sum_vat;
    byBuyerKg[r.buyer]=(byBuyerKg[r.buyer]||0)+r.kg;
    if(r.y&&r.m){const k=r.y+'-'+String(r.m).padStart(2,'0'); byMonth[k]=(byMonth[k]||0)+r.sum_vat;}
  }
  const debt = Object.keys(byDebtor).map(k=>({k,v:byDebtor[k],n:debtCnt[k]})).sort((a,b)=>b.v-a.v);
  let debtHtml='';
  if(debt.length){
    debtHtml = '<div class="panel" style="margin-top:14px"><h2>Должники — кто не оплатил</h2>'+
      '<table><thead><tr><th>Контрагент</th><th class="num">неопл. сделок</th><th class="num">сумма долга, ₸</th></tr></thead><tbody>'+
      debt.map(d=>'<tr><td>'+d.k+'</td><td class="num">'+d.n+'</td><td class="num" style="color:var(--red)">'+fmt(d.v)+'</td></tr>').join('')+
      '<tr><td><b>Итого долг</b></td><td class="num"><b>'+debt.reduce((s,d)=>s+d.n,0)+'</b></td><td class="num"><b style="color:var(--red)">'+fmt(unpaidSum)+'</b></td></tr>'+
      '</tbody></table></div>';
  }
  const noVat = sumVat - vat;
  c.innerHTML =
    '<div class="cards">'+
      card(fmt(sumVat),'₸','Отгружено (с НДС)','без НДС: '+fmt(noVat)+' ₸','coins','green')+
      card(fmt(paidSum),'₸','Оплачено','поступило по продажам','coins','green')+
      card(fmt(unpaidSum),'₸','Ждут оплаты (в долг)','дебиторка', 'coins', unpaidSum>0?'red':'')+
      card(fmtKg(kg),'кг','Продано крошки','','weight')+
      card(fmt(vat),'₸','НДС за период','','tax')+
      card(rows.length,'','Сделок','','deals')+
    '</div>'+
    '<div class="grid2">'+
      '<div class="panel"><h2>Продажи по фракциям, кг</h2><canvas id="cFrac"></canvas></div>'+
      '<div class="panel"><h2>Выручка по месяцам, ₸</h2><canvas id="cMonth"></canvas></div>'+
    '</div>'+
    '<div class="grid2">'+
      '<div class="panel"><h2>Наличные / Безналичные</h2><canvas id="cPay"></canvas></div>'+
      '<div class="panel"><h2>Топ контрагентов</h2>'+
        '<table><thead><tr><th>Контрагент</th><th class="num">кг</th><th class="num">сумма, ₸</th></tr></thead><tbody id="topBody"></tbody></table>'+
      '</div>'+
    '</div>'+
    debtHtml;
  const top = Object.keys(byBuyer).map(k=>({k,v:byBuyer[k],kg:byBuyerKg[k]})).sort((a,b)=>b.v-a.v).slice(0,10);
  document.getElementById('topBody').innerHTML = top.map(t=>
    '<tr><td>'+t.k+'</td><td class="num">'+fmtKg(t.kg)+'</td><td class="num">'+fmt(t.v)+'</td></tr>').join('');
  const green='#4ee29c', gold='#d9b56a', line='rgba(120,150,135,.15)', muted='#c5d6cd';
  Chart.defaults.color = muted; Chart.defaults.font.family="'Montserrat',sans-serif";
  Chart.defaults.font.size = 12; Chart.defaults.font.weight = '600';
  const fracOrder = ['0-1','1-2','2-4','4-6','6-8'];
  const fracLabels=[], fracData=[];
  for(const f of fracOrder){ if(byFrac[f]!==undefined){fracLabels.push(f);fracData.push(Math.round(byFrac[f]*10)/10);} }
  for(const f in byFrac){ if(!fracOrder.includes(f)){fracLabels.push(f);fracData.push(Math.round(byFrac[f]*10)/10);} }
  new Chart(document.getElementById('cFrac'),{type:'bar',
    data:{labels:fracLabels,datasets:[{data:fracData,backgroundColor:green,borderRadius:6,barThickness:26}]},
    options:{plugins:{legend:{display:false}},scales:chartScales(line)}});
  const mKeys=Object.keys(byMonth).sort();
  new Chart(document.getElementById('cMonth'),{type:'line',
    data:{labels:mKeys,datasets:[{data:mKeys.map(k=>Math.round(byMonth[k])),borderColor:green,backgroundColor:'rgba(78,226,156,.12)',fill:true,tension:.3,pointRadius:3,pointBackgroundColor:green,borderWidth:3}]},
    options:{plugins:{legend:{display:false}},scales:chartScales(line)}});
  const pKeys=Object.keys(byPay);
  new Chart(document.getElementById('cPay'),{type:'doughnut',
    data:{labels:pKeys,datasets:[{data:pKeys.map(k=>Math.round(byPay[k])),backgroundColor:[green,gold,muted,'#5a8fd9'],borderColor:'#11201a',borderWidth:3}]},
    options:{plugins:{legend:{position:'bottom'}},cutout:'62%'}});
}
function renderProd(){
  const rows = RAW_PROD.filter(inPeriod);
  const c = document.getElementById('content-prod');
  if(rows.length===0){
    c.innerHTML = '<div class="panel"><div class="empty">За выбранный период производства нет.<br>Попробуйте «Год» или «Всё время».</div></div>';
    return;
  }
  let total=0, metal=0, tyres=0;
  const byFrac={'0-1':0,'1-2':0,'2-4':0,'4-6':0,'6-8':0}, byMonth={};
  const days=new Set();
  for(const r of rows){
    total+=r.total; metal+=r.metal; tyres+=r.tyres;
    byFrac['0-1']+=r.f01; byFrac['1-2']+=r.f12; byFrac['2-4']+=r.f24;
    byFrac['4-6']+=r.f46; byFrac['6-8']+=r.f68;
    if(r.y&&r.m){const k=r.y+'-'+String(r.m).padStart(2,'0'); byMonth[k]=(byMonth[k]||0)+r.total;}
    if(r.y&&r.m&&r.d) days.add(r.y+'-'+r.m+'-'+r.d);
  }
  c.innerHTML =
    '<div class="cards">'+
      card(fmtKg(total),'кг','Произведено крошки','','weight','green')+
      card(fmtKg(metal),'кг','Металлокорд','','metal')+
      card(fmtKg(tyres),'кг','Переработано шин','','tyre')+
      card(days.size,'','Рабочих дней','','days')+
    '</div>'+
    '<div class="grid2">'+
      '<div class="panel"><h2>Производство крошки по месяцам, кг</h2><canvas id="cpMonth"></canvas></div>'+
      '<div class="panel"><h2>Произведено по фракциям, кг</h2><canvas id="cpFrac"></canvas></div>'+
    '</div>';
  const green='#4ee29c', line='rgba(120,150,135,.15)', muted='#c5d6cd';
  Chart.defaults.color = muted; Chart.defaults.font.size = 12; Chart.defaults.font.weight = '600';
  const mKeys=Object.keys(byMonth).sort();
  new Chart(document.getElementById('cpMonth'),{type:'line',
    data:{labels:mKeys,datasets:[{data:mKeys.map(k=>Math.round(byMonth[k]*10)/10),borderColor:green,backgroundColor:'rgba(78,226,156,.12)',fill:true,tension:.3,pointRadius:3,pointBackgroundColor:green,borderWidth:3}]},
    options:{plugins:{legend:{display:false}},scales:chartScales(line)}});
  const fracOrder=['0-1','1-2','2-4','4-6','6-8'];
  new Chart(document.getElementById('cpFrac'),{type:'bar',
    data:{labels:fracOrder,datasets:[{data:fracOrder.map(f=>Math.round(byFrac[f]*10)/10),backgroundColor:green,borderRadius:6,barThickness:28}]},
    options:{plugins:{legend:{display:false}},scales:chartScales(line)}});
}
function renderOstatok(){
  const c = document.getElementById('content-ost');
  if(!RAW_OST || !RAW_OST.fracs || RAW_OST.fracs.length===0){
    c.innerHTML = '<div class="panel"><div class="empty">Лист «Остаток» пуст или не найден.</div></div>';
    return;
  }
  const fracOrder=['0-1','1-2','2-4','4-6','6-8'];
  const sorted = RAW_OST.fracs.slice().sort((a,b)=>{
    let ia=fracOrder.indexOf(a.frac), ib=fracOrder.indexOf(b.frac);
    if(ia<0)ia=99; if(ib<0)ib=99; return ia-ib;
  });
  let cards = card(fmtKg(RAW_OST.total),'кг','Всего на складе','','box','green');
  for(const f of sorted){
    const cls = f.left<0 ? 'red' : '';
    cards += card(fmtKg(f.left),'кг','Фракция '+f.frac,'произв. '+fmtKg(f.made)+' · прод. '+fmtKg(f.sold),'box',cls);
  }
  const anyNeg = sorted.some(f=>f.left<0);
  const warn = anyNeg ? '<div class="panel" style="margin-top:14px"><div class="empty" style="color:var(--red)">Где-то остаток в минусе — продано больше, чем есть прихода. Проверьте стартовую строку «Начальный остаток» и записи реализации по этой фракции.</div></div>' : '';
  c.innerHTML = '<div class="cards">'+cards+'</div>'+warn;
}
function renderRashody(){
  const rows = RAW_RASH.filter(inPeriod);
  const c = document.getElementById('content-rashody');
  if(rows.length===0){
    c.innerHTML = '<div class="panel"><div class="empty">За выбранный период расходов нет.<br>Попробуйте «Месяц», «Год» или «Всё время».</div></div>';
    return;
  }
  let total=0, beznal=0, nal=0;
  const beznalContr={}, nalContr={}, byCat={};
  const normContr=s=>{
    s=String(s||'').toLowerCase().replace(/["'«»“”]/g,' ');
    ['индивидуальный предприниматель','товарищество с ограниченной ответственностью','акционерное общество','общество с ограниченной ответственностью','крестьянское хозяйство','частное предприятие'].forEach(f=>{s=s.split(f).join(' ');});
    s=s.replace(/\b(тоо|ао|ип|оао|зао|ооо|llp|llc|кх|чп)\b/g,' ');
    return s.replace(/[^a-zа-яё0-9і]/g,'');
  };
  const addC=(obj,name,sum)=>{const k=normContr(name)||name; if(!obj[k])obj[k]={name:name,v:0}; obj[k].v+=sum; if(name&&name.length<obj[k].name.length)obj[k].name=name;};
  for(const r of rows){
    total+=r.sum;
    byCat[r.cat]=(byCat[r.cat]||0)+r.sum;
    if(r.pay==='Наличный'){ nal+=r.sum; addC(nalContr,r.contr,r.sum); }
    else { beznal+=r.sum; addC(beznalContr,r.contr,r.sum); }
  }
  function tbl(obj, head){
    const arr=Object.values(obj).sort((a,b)=>b.v-a.v);
    if(arr.length===0) return '<div class="empty">нет</div>';
    return '<table><thead><tr><th>'+head+'</th><th class="num">сумма, ₸</th></tr></thead><tbody>'+
      arr.map(t=>'<tr><td>'+t.name+'</td><td class="num">'+fmt(t.v)+'</td></tr>').join('')+'</tbody></table>';
  }
  const catArr=Object.keys(byCat).map(k=>({k,v:byCat[k]})).sort((a,b)=>b.v-a.v);
  c.innerHTML =
    '<div class="cards">'+
      card(fmt(total),'₸','Всего расходов','','coins','red')+
      card(fmt(beznal),'₸','Безнал','','coins')+
      card(fmt(nal),'₸','Наличные','','coins')+
    '</div>'+
    '<div class="grid2">'+
      '<div class="panel"><h2>Безнал — кому платим</h2>'+tbl(beznalContr,'Контрагент')+'</div>'+
      '<div class="panel"><h2>Наличные — кому платим</h2>'+tbl(nalContr,'Контрагент')+'</div>'+
    '</div>'+
    '<div class="grid2">'+
      '<div class="panel"><h2>Расходы по категориям, ₸</h2><canvas id="cRashCat"></canvas></div>'+
      '<div class="panel"><h2>Категории — суммы</h2>'+tbl(byCat,'Категория')+'</div>'+
    '</div>';
  const line='rgba(120,150,135,.15)', muted='#c5d6cd';
  Chart.defaults.color = muted; Chart.defaults.font.size = 12; Chart.defaults.font.weight = '600';
  new Chart(document.getElementById('cRashCat'),{type:'bar',
    data:{labels:catArr.map(t=>t.k),datasets:[{data:catArr.map(t=>Math.round(t.v)),backgroundColor:'#e8705a',borderRadius:6,barThickness:20}]},
    options:{indexAxis:'y',plugins:{legend:{display:false}},scales:chartScales(line)}});
}
function renderTabel(){
  const rows = RAW_TABEL.filter(inPeriod);
  const c = document.getElementById('content-tabel');
  if(rows.length===0){
    c.innerHTML = '<div class="panel"><div class="empty">За выбранный период записей табеля нет.<br>Пришлите фото табеля боту — он заполнит сам.</div></div>';
    return;
  }
  let hours=0;
  const byEmp={}, days=new Set();
  for(const r of rows){
    hours+=r.hours;
    if(!byEmp[r.emp]) byEmp[r.emp]={role:r.role,sh:0,h:0,days:[]};
    byEmp[r.emp].sh++; byEmp[r.emp].h+=r.hours;
    if(r.role) byEmp[r.emp].role=r.role;
    if(r.d) byEmp[r.emp].days.push(r.d);
    if(r.y&&r.m&&r.d) days.add(r.y+'-'+r.m+'-'+r.d);
  }
  const emps=Object.keys(byEmp).sort((a,b)=>byEmp[b].h-byEmp[a].h);
  const showDays = (period==='month'||period==='day');
  c.innerHTML =
    '<div class="cards">'+
      card(emps.length,'','Сотрудников выходило','','days','green')+
      card(rows.length,'','Человеко-дней (смен)','','deals')+
      card(fmtKg(hours),'ч','Отработано часов','','days')+
      card(days.size,'','Рабочих дней','','days')+
    '</div>'+
    '<div class="panel"><h2>Кто когда вышел и сколько отработал</h2>'+
    '<table><thead><tr><th>Сотрудник</th><th>Должность</th><th class="num">смен</th><th class="num">часов</th>'+
    (showDays?'<th>дни месяца (числа)</th>':'')+'</tr></thead><tbody>'+
    emps.map(e=>{
      const x=byEmp[e];
      const dl = showDays ? '<td style="font-family:var(--mono);font-size:12px">'+x.days.sort((a,b)=>a-b).join(', ')+'</td>' : '';
      return '<tr><td>'+e+'</td><td>'+(x.role||'—')+'</td><td class="num">'+x.sh+'</td><td class="num">'+fmtKg(x.h)+'</td>'+dl+'</tr>';
    }).join('')+'</tbody></table></div>';
}
function renderFin(){
  const c = document.getElementById('content-fin');
  if(!RAW_FIN || RAW_FIN.ok===false){
    c.innerHTML = '<div class="panel"><div class="err">Не удалось загрузить финданные: '+((RAW_FIN&&RAW_FIN.error)||'нет данных')+'</div></div>';
    return;
  }
  const M = RAW_FIN.months||[];
  if(M.length===0){
    c.innerHTML = '<div class="panel"><div class="empty">В файле экономиста не нашёл помесячные данные (лист «Анализ и выводы»).<br>Проверим названия листов/строк — пришлите скрин этого экрана.</div></div>';
    return;
  }
  let tInc=0,tExp=0;
  for(const m of M){ tInc+=m.income; tExp+=m.expense; }
  const tProf=tInc-tExp;
  const S = RAW_FIN.sebes||[];
  const lastS = S.filter(s=>s.full1>0).slice(-1)[0];
  let html = '<div class="cards">'+
    card(fmt(tInc),'₸','Поступления с начала года','','coins','green')+
    card(fmt(tExp),'₸','Расходы с начала года','','coins','red')+
    card(fmt(tProf),'₸','Поступления − расходы','', 'coins', tProf>=0?'green':'red')+
    (lastS? card(fmt(lastS.full1),'₸/кг','Полная себест. 1 кг ('+lastS.month+')','произв.: '+fmt(lastS.prod1)+' ₸/кг','weight') : '')+
  '</div>';
  html += '<div class="grid2">'+
    '<div class="panel"><h2>Поступления и расходы по месяцам, ₸</h2><canvas id="cFinM"></canvas></div>'+
    '<div class="panel"><h2>По месяцам</h2><table><thead><tr><th>Месяц</th><th class="num">Поступл.</th><th class="num">Расход</th><th class="num">Разница</th></tr></thead><tbody>'+
      M.map(m=>'<tr><td>'+m.month+'</td><td class="num">'+fmt(m.income)+'</td><td class="num">'+fmt(m.expense)+'</td><td class="num" style="color:'+(m.profit>=0?'var(--green)':'var(--red)')+'">'+fmt(m.profit)+'</td></tr>').join('')+
      '<tr><td><b>Итого</b></td><td class="num"><b>'+fmt(tInc)+'</b></td><td class="num"><b>'+fmt(tExp)+'</b></td><td class="num"><b>'+fmt(tProf)+'</b></td></tr>'+
    '</tbody></table></div>'+
  '</div>';
  let blocks='';
  if(S.length){
    blocks += '<div class="panel"><h2>Себестоимость 1 кг по месяцам, ₸</h2><table><thead><tr><th>Месяц</th><th class="num">Крошка, кг</th><th class="num">Произв. 1 кг</th><th class="num">Полная 1 кг</th></tr></thead><tbody>'+
      S.map(s=>'<tr><td>'+s.month+'</td><td class="num">'+fmtKg(s.kg)+'</td><td class="num">'+fmt(s.prod1)+'</td><td class="num">'+fmt(s.full1)+'</td></tr>').join('')+'</tbody></table></div>';
  }
  const ST = RAW_FIN.structure||[];
  if(ST.length){
    blocks += '<div class="panel"><h2>Структура расходов (с начала года)</h2><canvas id="cFinS"></canvas></div>';
  }
  if(blocks) html += '<div class="grid2">'+blocks+'</div>';
  const EA = RAW_FIN.expense_articles||[];
  const EM = RAW_FIN.expense_months||[];
  if(EA.length && EM.length){
    const th = EM.map(m=>'<th class="num">'+m+'</th>').join('');
    const colTot = EM.map((_,i)=>EA.reduce((s,a)=>s+(a.by_month[i]||0),0));
    let body = EA.map(a=>'<tr><td>'+a.name+'</td>'+
      a.by_month.map(v=>'<td class="num">'+fmt(v)+'</td>').join('')+
      '<td class="num"><b>'+fmt(a.total)+'</b></td></tr>').join('');
    const grand = colTot.reduce((s,v)=>s+v,0);
    body += '<tr><td><b>Итого</b></td>'+colTot.map(v=>'<td class="num"><b>'+fmt(v)+'</b></td>').join('')+
      '<td class="num"><b>'+fmt(grand)+'</b></td></tr>';
    html += '<div class="panel"><h2>Расходы по статьям и месяцам, ₸ <span style="color:var(--dim);font-weight:500;font-size:12px">(из таблицы экономиста)</span></h2>'+
      '<div style="overflow-x:auto"><table><thead><tr><th>Статья</th>'+th+'<th class="num">Итого</th></tr></thead><tbody>'+body+'</tbody></table></div></div>';
    html += '<div class="grid2">'+
      '<div class="panel"><h2>Расходы по месяцам по статьям, ₸</h2><canvas id="cFinStack"></canvas></div>'+
      '<div class="panel"><h2>Производство по месяцам, кг</h2><canvas id="cFinProd"></canvas></div>'+
    '</div>';
  }
  c.innerHTML = html;
  const green='#4ee29c', red='#e8705a', gold='#d9b56a', line='rgba(120,150,135,.15)', muted='#c5d6cd';
  Chart.defaults.color = muted; Chart.defaults.font.size = 12; Chart.defaults.font.weight='600';
  new Chart(document.getElementById('cFinM'),{type:'bar',
    data:{labels:M.map(m=>m.month),datasets:[
      {label:'Поступления',data:M.map(m=>Math.round(m.income)),backgroundColor:green,borderRadius:5},
      {label:'Расходы',data:M.map(m=>Math.round(m.expense)),backgroundColor:red,borderRadius:5}]},
    options:{plugins:{legend:{position:'bottom'}},scales:chartScales(line)}});
  if(ST.length){
    new Chart(document.getElementById('cFinS'),{type:'doughnut',
      data:{labels:ST.map(s=>s.name),datasets:[{data:ST.map(s=>Math.round(s.sum)),backgroundColor:[green,gold,red,'#5a8fd9',muted,'#b48ee1','#e1b14e','#6fcf97'],borderColor:'#11201a',borderWidth:3}]},
      options:{plugins:{legend:{position:'bottom'}},cutout:'58%'}});
  }
  if(EA.length && EM.length){
    const pal=[green,gold,red,'#5a8fd9',muted,'#b48ee1','#e1b14e','#6fcf97'];
    new Chart(document.getElementById('cFinStack'),{type:'bar',
      data:{labels:EM,datasets:EA.map((a,i)=>({label:a.name,data:a.by_month.map(v=>Math.round(v)),backgroundColor:pal[i%pal.length],borderRadius:4}))},
      options:{plugins:{legend:{position:'bottom'}},scales:{x:{stacked:true,grid:{display:false},ticks:{font:{weight:'700'}}},y:{stacked:true,grid:{color:line},beginAtZero:true,ticks:{font:{weight:'600'}}}}}});
    const prodM = S.filter(s=>s.kg>0);
    new Chart(document.getElementById('cFinProd'),{type:'bar',
      data:{labels:prodM.map(s=>s.month),datasets:[{data:prodM.map(s=>Math.round(s.kg*10)/10),backgroundColor:green,borderRadius:6,barThickness:28}]},
      options:{plugins:{legend:{display:false}},scales:chartScales(line)}});
  }
}
function setSection(s){
  section = s;
  document.querySelectorAll('.nav-btn').forEach(b=>b.classList.toggle('on', b.dataset.s===s));
  document.querySelectorAll('.mtabs button').forEach(b=>b.classList.toggle('on', b.dataset.s===s));
  document.getElementById('sec-sales').style.display = s==='sales'?'block':'none';
  document.getElementById('sec-prod').style.display  = s==='prod' ?'block':'none';
  document.getElementById('sec-ost').style.display   = s==='ost'  ?'block':'none';
  document.getElementById('sec-rashody').style.display = s==='rashody'?'block':'none';
  document.getElementById('sec-tabel').style.display = s==='tabel'?'block':'none';
  document.getElementById('sec-fin').style.display = s==='fin'?'block':'none';
  document.getElementById('periods').style.display = (s==='ost'||s==='fin')?'none':'flex';
}
document.querySelectorAll('.nav-btn, .mtabs button').forEach(b=>{
  b.addEventListener('click',()=>setSection(b.dataset.s));
});
document.getElementById('periods').addEventListener('click',e=>{
  if(e.target.tagName!=='BUTTON')return;
  period=e.target.dataset.p;
  document.querySelectorAll('.periods button').forEach(b=>b.classList.remove('on'));
  e.target.classList.add('on');
  render();
  renderProd();
  renderRashody();
  renderTabel();
});
function updFoot(){
  document.getElementById('foot').textContent='Данные из таблицы склада · продаж '+RAW.length+', производства '+RAW_PROD.length;
}
fetch('/data').then(r=>r.json()).then(j=>{
  if(!j.ok){document.getElementById('content').innerHTML='<div class="err">Ошибка: '+j.error+'</div>';return;}
  RAW=j.records;
  document.getElementById('updated').textContent='обновлено '+new Date().toLocaleString('ru-RU');
  updFoot(); render();
}).catch(e=>{document.getElementById('content').innerHTML='<div class="err">Не удалось загрузить: '+e+'</div>';});
fetch('/data-proizvodstvo').then(r=>r.json()).then(j=>{
  if(!j.ok){document.getElementById('content-prod').innerHTML='<div class="err">Ошибка: '+j.error+'</div>';return;}
  RAW_PROD=j.records; updFoot(); renderProd();
}).catch(e=>{document.getElementById('content-prod').innerHTML='<div class="err">Не удалось загрузить: '+e+'</div>';});
fetch('/data-ostatok').then(r=>r.json()).then(j=>{
  if(!j.ok){document.getElementById('content-ost').innerHTML='<div class="err">Ошибка: '+j.error+'</div>';return;}
  RAW_OST=j; renderOstatok();
}).catch(e=>{document.getElementById('content-ost').innerHTML='<div class="err">Не удалось загрузить: '+e+'</div>';});
fetch('/data-rashody').then(r=>r.json()).then(j=>{
  if(!j.ok){document.getElementById('content-rashody').innerHTML='<div class="err">Ошибка: '+j.error+'</div>';return;}
  RAW_RASH=j.records; renderRashody();
}).catch(e=>{document.getElementById('content-rashody').innerHTML='<div class="err">Не удалось загрузить: '+e+'</div>';});
fetch('/data-tabel').then(r=>r.json()).then(j=>{
  if(!j.ok){document.getElementById('content-tabel').innerHTML='<div class="err">Ошибка: '+j.error+'</div>';return;}
  RAW_TABEL=j.records; renderTabel();
}).catch(e=>{document.getElementById('content-tabel').innerHTML='<div class="err">Не удалось загрузить: '+e+'</div>';});
fetch('/data-fin').then(r=>r.json()).then(j=>{
  RAW_FIN=j; renderFin();
}).catch(e=>{document.getElementById('content-fin').innerHTML='<div class="err">Не удалось загрузить: '+e+'</div>';});
</script>
</body>
</html>"""
if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))
    app.run(host="0.0.0.0", port=port)
