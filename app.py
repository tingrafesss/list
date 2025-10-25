#!/usr/bin/env python3
"""
Book Discount Checker — Flask + SQLite (single file)
- Светлая тема
- CSV/Excel (.csv/.xlsx/.xls)
- Терпимое распознавание заголовков (рус/англ, синонимы)
- Авто-определение строки заголовков в Excel (если заголовки не на первой строке)
- Авто-распознавание разделителя CSV
- ID — главный ключ для пересечений
"""
from __future__ import annotations

import csv
import io
import os
import sqlite3
from datetime import datetime, date
from typing import List, Dict, Optional

from flask import (
    Flask,
    g,
    redirect,
    render_template_string,
    request,
    send_file,
    url_for,
    flash,
)
from slugify import slugify
from jinja2 import DictLoader

APP_TITLE = "Book Discount Checker"
DB_PATH = os.environ.get("BOOKCHECK_DB", "bookcheck.sqlite3")
UPLOAD_LIMIT_MB = 12
ALLOWED_EXT = {".csv", ".xlsx", ".xls"}

# ---- tokens for fuzzy header matching ----
HEADER_TOKENS = {
    "id": ["id", "ид", "код", "артикул", "bookid", "номеркниги", "номер", "кодкниги"],
    "title": ["title", "наименование", "название", "name", "наим"],
    "fmt": ["format", "формат", "тип", "вид"],
    "author": ["author", "автор"],
    "category": ["category", "категория", "жанр", "раздел"],
    "price": ["price", "цена", "стоимость"],
    "discount": ["discount", "скидка", "percent", "процент"],
}

def key_norm(s: str) -> str:
    return "".join(ch.lower() for ch in str(s) if ch.isalnum())

def build_header_map(keys: List[str]) -> Dict[str, str]:
    norm_keys = {key_norm(k): k for k in keys}
    mapping: Dict[str, str] = {}
    for internal, tokens in HEADER_TOKENS.items():
        for nk, orig in norm_keys.items():
            for t in tokens:
                tn = key_norm(t)
                if tn and (tn in nk or nk in tn):
                    mapping[internal] = orig
                    break
            if internal in mapping:
                break
    return mapping

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("FLASK_SECRET", "dev-secret-change-me")
app.config["MAX_CONTENT_LENGTH"] = UPLOAD_LIMIT_MB * 1024 * 1024

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()

SCHEMA_SQL = """
CREATE TABLE IF NOT EXISTS books (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  owner TEXT NOT NULL,
  book_id TEXT,
  title TEXT,
  title_key TEXT,
  fmt TEXT,
  author TEXT,
  category TEXT,
  price TEXT,
  discount TEXT,
  start_date TEXT NOT NULL,
  end_date TEXT NOT NULL,
  created_at TEXT NOT NULL
);
CREATE INDEX IF NOT EXISTS idx_books_owner ON books(owner);
CREATE INDEX IF NOT EXISTS idx_books_bookid ON books(book_id);
CREATE INDEX IF NOT EXISTS idx_books_titlekey ON books(title_key);
CREATE INDEX IF NOT EXISTS idx_books_dates ON books(start_date, end_date);
"""

with app.app_context():
    db = get_db()
    db.executescript(SCHEMA_SQL)
    cols = {row[1] for row in db.execute("PRAGMA table_info(books)")}
    for new_col in ["book_id", "fmt", "category", "price", "discount"]:
        if new_col not in cols:
            try:
                db.execute(f"ALTER TABLE books ADD COLUMN {new_col} TEXT;")
            except Exception:
                pass
    db.commit()

def parse_date(s: str) -> date:
    return datetime.strptime(s.strip(), "%Y-%m-%d").date()

def norm_title(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    return slugify(s, lowercase=True).replace("-", "")

def allowed_file(filename: str) -> bool:
    ext = os.path.splitext(filename)[1].lower()
    return ext in ALLOWED_EXT

def normalize_id(val) -> str:
    if val is None or val == "":
        return ""
    try:
        f = float(val)
        if f.is_integer():
            return str(int(f))
    except Exception:
        pass
    return str(val).strip()

def sniff_delimiter(text: str) -> str:
    sample = text[:4096]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=[",",";","\\t","|"])
        return dialect.delimiter
    except Exception:
        if ";" in sample and "," not in sample.splitlines()[0]:
            return ";"
        return ","

# --- Excel header row detection ---
def detect_header_row_xlsx(content: bytes) -> Optional[int]:
    try:
        import pandas as pd  # type: ignore
    except Exception:
        return None
    df = pd.read_excel(io.BytesIO(content), header=None, nrows=20)
    best_idx, best_score = None, -1
    for i in range(min(20, len(df))):
        row_vals = [str(x) for x in list(df.iloc[i].values)]
        keys = [str(v) for v in row_vals if v and str(v) != 'nan']
        mapping = build_header_map(keys)
        score = len(mapping)
        if ("id" in mapping or "title" in mapping) and score > best_score:
            best_idx, best_score = i, score
    return best_idx

# ---- Tabular readers (CSV/Excel) ----
def read_table(file_storage, mode: str) -> List[dict]:
    filename = file_storage.filename or ""
    ext = os.path.splitext(filename)[1].lower()
    content = file_storage.read()
    file_storage.stream.seek(0)

    if ext == ".csv":
        text = content.decode("utf-8-sig", errors="ignore")
        delim = sniff_delimiter(text)
        reader = csv.DictReader(io.StringIO(text), delimiter=delim)
        raw_rows = list(reader)
        headers = list(reader.fieldnames or [])
    else:
        try:
            import pandas as pd  # type: ignore
        except Exception as e:
            raise RuntimeError("Для Excel установите зависимости: pip install pandas openpyxl") from e
        header_row = detect_header_row_xlsx(content)
        if header_row is None:
            df = pd.read_excel(io.BytesIO(content))  # default header
        else:
            df = pd.read_excel(io.BytesIO(content), header=header_row)
        headers = [str(c) for c in df.columns]
        raw_rows = df.to_dict(orient="records")

    header_map = build_header_map(headers)

    norm_rows: List[dict] = []
    for r in raw_rows:
        # make also a lowercase-key view for robustness
        lower = {str(k): (v.strip() if isinstance(v, str) else ("" if v is None else v)) for k, v in r.items()}
        row_norm = {}
        for internal, orig_key in header_map.items():
            row_norm[internal] = lower.get(orig_key, lower.get(str(orig_key).lower(), ""))
        # fallback for title if mapping failed
        if "title" not in row_norm or row_norm["title"] == "":
            for k in ("title","наименование","название","Наименование","Название"):
                if k in lower and lower[k]:
                    row_norm["title"] = lower[k]
                    break
        norm_rows.append(row_norm)

    if not norm_rows:
        return []

    header_has_id = "id" in header_map
    header_has_title = "title" in header_map or any("title" in r for r in norm_rows)

    if mode == "upload" and not header_has_id:
        raise ValueError("Нужна колонка 'id' (или её эквивалент: Ид/Код/Артикул и т.п.).")
    if mode == "check" and not (header_has_id or header_has_title):
        raise ValueError("Нужна колонка 'id' или 'title' (Наименование/Название).")

    return norm_rows

# -------------------- Templates (коротко) --------------------
BASE_HTML = """
<!doctype html><html lang="ru"><head>
<meta charset="utf-8"/><meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>{{ title }}</title>
<link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css" rel="stylesheet">
<style>body{background:#f2f4f8;color:#1f2937}.card{background:#fff;border:1px solid #d1d5db}.navbar{background:#e5e7eb;color:#1f2937}
a{color:#2563eb}.form-label{margin-top:.5rem}.badge-soft{background:#e5e7eb;color:#374151}.table thead th{color:#374151}.table{color:#1f2937}</style>
</head><body>
<nav class="navbar navbar-expand-lg navbar-light mb-4"><div class="container">
<a class="navbar-brand" href="{{ url_for('index') }}">📚 {{ app_title }}</a>
<div class="d-flex gap-2"><a class="btn btn-sm btn-outline-dark" href="{{ url_for('upload_page') }}">Загрузить список</a>
<a class="btn btn-sm btn-primary" href="{{ url_for('check_page') }}">Проверить совпадения</a></div></div></nav>
<a class="btn btn-sm btn-outline-secondary" href="{{ url_for('lists') }}">Списки</a>
<div class="container">{% with messages = get_flashed_messages(with_categories=true) %}{% if messages %}{% for c,m in messages %}
<div class="alert alert-{{ 'danger' if c=='error' else c }}">{{ m }}</div>{% endfor %}{% endif %}{% endwith %}{% block content %}{% endblock %}</div>
</body></html>
"""

INDEX_HTML = "{% extends 'base.html' %}{% block content %}<div class='card p-4'><h1 class='h4'>Главная</h1><p>Сайт понимает Excel, где заголовки не на первой строке (например, если сверху стоит строка «Таблица 1»).</p></div>{% endblock %}"

UPLOAD_HTML = """
{% extends 'base.html' %}{% block content %}
<div class="card p-4">
  <h1 class="h4">Загрузить список</h1>
  <form method="post" enctype="multipart/form-data" action="{{ url_for('upload') }}" class="mt-3">
    <div class="row g-3">
      <div class="col-md-3"><label class="form-label">Название</label><input name="owner" class="form-control" required placeholder="Ввести"/></div>
      <div class="col-md-3"><label class="form-label">Начало скидки</label><input type="date" name="start" class="form-control" required/></div>
      <div class="col-md-3"><label class="form-label">Конец скидки</label><input type="date" name="end" class="form-control" required/></div>
      <div class="col-md-3"><label class="form-label">Файл CSV/Excel</label><input type="file" name="file" class="form-control" accept=".csv,.xlsx,.xls" required/></div>
    </div>
    <button class="btn btn-primary mt-3">Загрузить</button>
  </form>
</div>
{% endblock %}
"""

CHECK_HTML = """
{% extends 'base.html' %}
{% block content %}
<div class="card p-4">
  <h1 class="h4">Проверить совпадения</h1>
  <form method="post" enctype="multipart/form-data" action="{{ url_for('check') }}" class="mt-3">
    <div class="row g-3">
      <div class="col-md-4">
        <label class="form-label">Название</label>
        <input name="owner" class="form-control" required placeholder="Ввести" />
      </div>
      <div class="col-md-4">
        <label class="form-label">Период вашей акции — начало</label>
        <input type="date" name="start" class="form-control" required />
      </div>
      <div class="col-md-4">
        <label class="form-label">Период вашей акции — конец</label>
        <input type="date" name="end" class="form-control" required />
      </div>
      <div class="col-12">
        <label class="form-label">Файл CSV/Excel (достаточно id или Наименование)</label>
        <input type="file" name="file" class="form-control" accept=".csv,.xlsx,.xls" required />
      </div>
    </div>
    <button class="btn btn-primary mt-3">Проверить</button>
  </form>
</div>

{% if results %}
<div class="card p-4 mt-4">
  <h2 class="h5">Результат</h2>
  <p class="mb-2">
    <span class="badge badge-soft">Конфликтов:</span> {{ results.conflicts|length }}
    |
    <span class="badge badge-soft">Чистых:</span> {{ results.clean|length }}
  </p>

  <div class="row g-4">
    <!-- Левая колонка: конфликты -->
    <div class="col-lg-6">
      <h3 class="h6">Пересекаются</h3>
      {% if results.conflicts %}
      <div class="table-responsive">
        <table class="table table-sm">
          <thead>
            <tr>
              <th>ID</th><th>Наименование</th><th>Автор</th><th>Тип</th>
              <th>Категория</th><th>Цена</th><th>Скидка</th><th>Совпадения</th>
            </tr>
          </thead>
          <tbody>
          {% for item in results.conflicts %}
            <tr>
              <td>{{ item['book_id'] or '' }}</td>
              <td>{{ item['title'] or '' }}</td>
              <td>{{ item['author'] or '' }}</td>
              <td>{{ item['fmt'] or '' }}</td>
              <td>{{ item['category'] or '' }}</td>
              <td>{{ item['price'] or '' }}</td>
              <td>{{ item['discount'] or '' }}</td>
              <td>
                {% for hit in item['matches'] %}
                  <div>
                    <span class="badge badge-soft">{{ hit['owner'] }}</span>
                    • {{ hit['start_date'] }}→{{ hit['end_date'] }}
                  </div>
                {% endfor %}
              </td>
            </tr>
          {% endfor %}
          </tbody>
        </table>
      </div>
      {% else %}
        <p class="text-muted mb-0">Совпадений не найдено.</p>
      {% endif %}
    </div>

    <!-- Правая колонка: чистые -->
    <div class="col-lg-6">
      <div class="d-flex justify-content-between align-items-center">
        <h3 class="h6 mb-0">Можно брать (без конфликтов)</h3>
        <a class="btn btn-outline-dark btn-sm mb-2"
           href="{{ url_for('download_clean', token=results.token) }}">⬇️ Скачать Excel</a>
      </div>
      {% if results.clean %}
      <div class="table-responsive mt-2">
        <table class="table table-sm">
          <thead>
            <tr>
              <th>ID</th><th>Наименование</th><th>Автор</th><th>Тип</th>
              <th>Категория</th><th>Цена</th><th>Скидка</th>
            </tr>
          </thead>
          <tbody>
          {% for item in results.clean %}
            <tr>
              <td>{{ item['book_id'] or '' }}</td>
              <td>{{ item['title'] or '' }}</td>
              <td>{{ item['author'] or '' }}</td>
              <td>{{ item['fmt'] or '' }}</td>
              <td>{{ item['category'] or '' }}</td>
              <td>{{ item['price'] or '' }}</td>
              <td>{{ item['discount'] or '' }}</td>
            </tr>
          {% endfor %}
          </tbody>
        </table>
      </div>
      {% else %}
        <p class="text-muted mt-2 mb-0">Все позиции пересекаются.</p>
      {% endif %}
    </div>
  </div>
</div>
{% endif %}
{% endblock %}
"""
LISTS_HTML = """
{% extends 'base.html' %}
{% block content %}
<div class="card p-4">
  <h1 class="h4">📜 Загруженные списки</h1>
  {% if rows %}
  <div class="table-responsive mt-3">
    <table class="table table-sm">
      <thead>
        <tr><th>Магазин</th><th>Начало</th><th>Конец</th><th>Позиций</th><th></th></tr>
      </thead>
      <tbody>
      {% for r in rows %}
        <tr>
          <td>{{ r['owner'] }}</td>
          <td>{{ r['start_date'] }}</td>
          <td>{{ r['end_date'] }}</td>
          <td>{{ r['total'] }}</td>
          <td>
            <a class="btn btn-sm btn-outline-primary"
               href="{{ url_for('list_detail', owner=r['owner'], start=r['start_date'], end=r['end_date']) }}">
              Просмотр
            </a>
          </td>
        </tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
  {% else %}
    <p class="text-muted mt-3">Пока нет загруженных списков.</p>
  {% endif %}
</div>
{% endblock %}
"""

LIST_DETAIL_HTML = """
{% extends 'base.html' %}
{% block content %}
<div class="card p-4">
  <h1 class="h4">📚 {{ owner }} — {{ start }} → {{ end }}</h1>
  {% if items %}
  <div class="table-responsive mt-3">
    <table class="table table-sm">
      <thead>
        <tr>
          <th>ID</th><th>Наименование</th><th>Автор</th><th>Тип</th>
          <th>Категория</th><th>Цена</th><th>Скидка</th>
        </tr>
      </thead>
      <tbody>
      {% for i in items %}
        <tr>
          <td>{{ i['book_id'] or '' }}</td>
          <td>{{ i['title'] or '' }}</td>
          <td>{{ i['author'] or '' }}</td>
          <td>{{ i['fmt'] or '' }}</td>
          <td>{{ i['category'] or '' }}</td>
          <td>{{ i['price'] or '' }}</td>
          <td>{{ i['discount'] or '' }}</td>
        </tr>
      {% endfor %}
      </tbody>
    </table>
  </div>
  {% else %}
    <p class="text-muted mt-3">Пустой список.</p>
  {% endif %}
</div>
{% endblock %}
"""


# -------------------- Routes --------------------
@app.route("/")
def index():
    return render_template_string(INDEX_HTML, title=APP_TITLE, app_title=APP_TITLE)

@app.route("/upload", methods=["GET"])
def upload_page():
    return render_template_string(UPLOAD_HTML, title=f"Загрузка — {APP_TITLE}", app_title=APP_TITLE)

@app.route("/upload", methods=["POST"])
def upload():
    owner = request.form.get("owner", "").strip()
    start = request.form.get("start", "").strip()
    end = request.form.get("end", "").strip()
    file = request.files.get("file")
    if not owner or not start or not end or not file or file.filename == "":
        flash("Заполните поля и выберите файл.", "error"); return redirect(url_for("upload_page"))
    try:
        s = parse_date(start); e = parse_date(end); 
        if e < s: raise ValueError("Конец периода раньше начала.")
    except Exception as ex:
        flash(f"Неверные даты: {ex}", "error"); return redirect(url_for("upload_page"))
    if not allowed_file(file.filename):
        flash("Разрешены только .csv/.xlsx/.xls", "error"); return redirect(url_for("upload_page"))
    try:
        rows = read_table(file, mode="upload")
        rows_to_insert = []
        for r in rows:
            book_id = normalize_id(r.get("id"))
            title   = (str(r.get("title") or "").strip()) or None
            author  = (str(r.get("author") or "").strip()) or None
            fmt     = (str(r.get("fmt") or "").strip()) or None
            category= (str(r.get("category") or "").strip()) or None
            price   = (str(r.get("price") or "").strip()) or None
            disc    = (str(r.get("discount") or "").strip()) or None
            if not book_id and not title:
                continue
            rows_to_insert.append((owner, book_id or None, title, (slugify(title, lowercase=True).replace('-','') if title else None),
                                   fmt, author, category, price, disc, s.isoformat(), e.isoformat(), datetime.utcnow().isoformat()))
        if not rows_to_insert: raise ValueError("Нет валидных строк.")
        db = get_db()
        db.executemany("""
            INSERT INTO books(owner, book_id, title, title_key, fmt, author, category, price, discount, start_date, end_date, created_at)
            VALUES(?,?,?,?,?,?,?,?,?,?,?,?)
        """, rows_to_insert)
        db.commit()
        flash(f"Успех: добавлено {len(rows_to_insert)} записей.", "success")
    except Exception as e:
        flash(f"Ошибка: {e}", "error")
    return redirect(url_for("upload_page"))

@app.route("/check", methods=["GET"])
def check_page():
    return render_template_string(CHECK_HTML, title=f"Проверка — {APP_TITLE}", app_title=APP_TITLE, results=None)

# КЭШ для выгрузки "чистого" списка
_CLEAN_CACHE: Dict[str, Dict] = {}

@app.route("/check", methods=["POST"])
def check():
    owner = request.form.get("owner", "").strip()
    start = request.form.get("start", "").strip()
    end = request.form.get("end", "").strip()
    file = request.files.get("file")

    if not owner or not start or not end:
        flash("Заполните все поля.", "error")
        return redirect(url_for("check_page"))

    # валидация дат
    try:
        s = parse_date(start)
        e = parse_date(end)
        if e < s:
            raise ValueError("Конец периода раньше начала.")
    except Exception as ex:
        flash(f"Неверные даты: {ex}", "error")
        return redirect(url_for("check_page"))

    # файл
    if not file or file.filename == "":
        flash("Файл не выбран.", "error")
        return redirect(url_for("check_page"))
    if not allowed_file(file.filename):
        flash("Разрешены только .csv/.xlsx/.xls", "error")
        return redirect(url_for("check_page"))

    # читаем кандидатов (допускается только id или только название)
    try:
        rows = read_table(file, mode="check")
        candidates = []
        for r in rows:
            book_id = normalize_id(r.get("id")) or None
            title   = (str(r.get("title") or "").strip()) or None
            author  = (str(r.get("author") or "").strip()) or None
            fmt     = (str(r.get("fmt") or "").strip()) or None
            category= (str(r.get("category") or "").strip()) or None
            price   = (str(r.get("price") or "").strip()) or None
            disc    = (str(r.get("discount") or "").strip()) or None
            if not book_id and not title:
                continue
            candidates.append({
                "book_id": book_id,
                "title": title,
                "title_key": (norm_title(title) if title else None),
                "author": author,
                "fmt": fmt,
                "category": category,
                "price": price,
                "discount": disc,
            })
        if not candidates:
            raise ValueError("Нет валидных строк.")
    except Exception as e:
        flash(f"Ошибка в файле: {e}", "error")
        return redirect(url_for("check_page"))

    # собираем потенциальные совпадения одним запросом
    ids = [c["book_id"] for c in candidates if c["book_id"]]
    title_keys = [c["title_key"] for c in candidates if (not c["book_id"]) and c["title_key"]]

    db = get_db()
    sql_parts = []
    params: List = []
    if ids:
        sql_parts.append(f"(book_id IN ({','.join(['?']*len(ids))}))")
        params.extend(ids)
    if title_keys:
        sql_parts.append(f"(book_id IS NULL AND title_key IN ({','.join(['?']*len(title_keys))}))")
        params.extend(title_keys)

    sql_where = " OR ".join(sql_parts) if sql_parts else "1=0"
    # пересечение по датам: есть overlap с [s, e]
    sql = f"""
        SELECT * FROM books
        WHERE ({sql_where})
          AND NOT (date(end_date) < date(?) OR date(?) < date(start_date))
    """
    params.extend([s.isoformat(), s.isoformat()])
    hits = db.execute(sql, params).fetchall() if sql_parts else []

    # индексируем найденные совпадения
    by_id: Dict[str, List[sqlite3.Row]] = {}
    by_tk: Dict[str, List[sqlite3.Row]] = {}
    for h in hits:
        if h["book_id"]:
            by_id.setdefault(h["book_id"], []).append(h)
        else:
            by_tk.setdefault(h["title_key"], []).append(h)

    conflicts, clean = [], []
    for c in candidates:
        matched = []
        if c["book_id"] and c["book_id"] in by_id:
            matched = by_id[c["book_id"]]
        elif (not c["book_id"]) and c["title_key"] and c["title_key"] in by_tk:
            matched = by_tk[c["title_key"]]

        if matched:
            conflicts.append({
                "book_id": c["book_id"],
                "title": c["title"],
                "author": c["author"],
                "fmt": c["fmt"],
                "category": c["category"],
                "price": c["price"],
                "discount": c["discount"],
                "matches": [
                    {"owner": m["owner"], "start_date": m["start_date"], "end_date": m["end_date"]}
                    for m in matched
                ],
            })
        else:
            clean.append({
                "book_id": c["book_id"],
                "title": c["title"],
                "author": c["author"],
                "fmt": c["fmt"],
                "category": c["category"],
                "price": c["price"],
                "discount": c["discount"],
            })

    # сохраняем «чистый» список для скачивания
    token = f"{owner}-{datetime.now().timestamp()}"
    _CLEAN_CACHE[token] = {"rows": clean, "owner": owner}

    results = {"conflicts": conflicts, "clean": clean, "token": token}
    return render_template_string(
        CHECK_HTML,
        title=f"Проверка — {APP_TITLE}",
        app_title=APP_TITLE,
        results=results,
    )

# скачать «чистый» список
@app.route("/download/clean/<token>")
def download_clean(token: str):
    payload = _CLEAN_CACHE.get(token)
    if not payload:
        flash("Ссылка устарела. Перезапустите проверку.", "error")
        return redirect(url_for("check_page"))

    # --- формируем Excel-файл (.xlsx) ---
    try:
        from openpyxl import Workbook
    except ImportError:
        flash("Нужно установить openpyxl: pip install openpyxl", "error")
        return redirect(url_for("check_page"))

    wb = Workbook()
    ws = wb.active
    ws.title = "Чистые позиции"

    # шапка
    headers = ["ID", "Наименование", "Автор", "Тип", "Категория", "Цена", "Скидка"]
    ws.append(headers)

    # строки
    for r in payload["rows"]:
        ws.append([
            r.get("book_id") or "",
            r.get("title") or "",
            r.get("author") or "",
            r.get("fmt") or "",
            r.get("category") or "",
            r.get("price") or "",
            r.get("discount") or "",
        ])

    # сохраняем в память
    from io import BytesIO
    mem = BytesIO()
    wb.save(mem)
    mem.seek(0)

    filename = f"clean_{slugify(payload['owner'])}_{datetime.now().strftime('%Y%m%d%H%M%S')}.xlsx"
    return send_file(
        mem,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )

@app.route("/ping")
def ping():
    return "ok", 200

@app.route("/lists")
def lists():
    db = get_db()
    # группируем по owner, start_date, end_date
    rows = db.execute("""
        SELECT owner, start_date, end_date, COUNT(*) as total,
               MIN(created_at) as created_at
        FROM books
        GROUP BY owner, start_date, end_date
        ORDER BY created_at DESC
    """).fetchall()

    return render_template_string(LISTS_HTML,
        title=f"Загруженные списки — {APP_TITLE}",
        app_title=APP_TITLE,
        rows=rows
    )

@app.route("/lists/<owner>/<start>/<end>")
def list_detail(owner, start, end):
    db = get_db()
    items = db.execute("""
        SELECT book_id, title, author, fmt, category, price, discount
        FROM books
        WHERE owner = ? AND start_date = ? AND end_date = ?
        ORDER BY title
    """, (owner, start, end)).fetchall()

    return render_template_string(LIST_DETAIL_HTML,
        title=f"Список {owner} {start}–{end}",
        app_title=APP_TITLE,
        owner=owner,
        start=start,
        end=end,
        items=items
    )


# ------------- Jinja loader -------------
app.jinja_loader = DictLoader({'base.html': BASE_HTML})

if __name__ == "__main__":
    with app.app_context():
        get_db()
    app.run(debug=True, host="127.0.0.1", port=8000, use_reloader=False)

