import os
from datetime import date, datetime
from io import BytesIO

import psycopg2
from psycopg2.extras import RealDictCursor
from flask import Flask, request, redirect, url_for, render_template, flash, send_file
from dotenv import load_dotenv
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

load_dotenv()
app = Flask(__name__)
app.secret_key = os.getenv("FLASK_SECRET_KEY", "change-this-development-key")

def db():
    return psycopg2.connect(
        host=os.getenv("PGHOST", "localhost"),
        port=os.getenv("PGPORT", "5432"),
        dbname=os.getenv("PGDATABASE", "postgres"),
        user=os.getenv("PGUSER", "postgres"),
        password=os.getenv("PGPASSWORD", ""),
        options="-c search_path=" + os.getenv("PGSCHEMA", "public")
    )

def rows(sql, params=()):
    with db() as conn:
        with conn.cursor(cursor_factory=RealDictCursor) as cur:
            cur.execute(sql, params)
            return cur.fetchall()

def execute(sql, params=()):
    with db() as conn:
        with conn.cursor() as cur:
            cur.execute(sql, params)

def ref_data():
    return {
        "users": rows("SELECT id AS user_id, full_name, role::text AS role FROM users ORDER BY full_name"),
        "animals": rows("""SELECT id AS animal_id, COALESCE(nickname, tag_number, 'Животное #' || id) AS label
                           FROM animals ORDER BY label"""),
        "groups": rows('SELECT  id AS group_id, name FROM "groups" ORDER BY name')
    }

@app.route("/")
def home():
    selected = request.args.get("day", date.today().isoformat())
    try:
        day = date.fromisoformat(selected)
    except ValueError:
        day = date.today()
    summary = rows("""
        SELECT
          COALESCE((SELECT SUM(volume_liters) FROM milk_yield WHERE date_yield=%s),0) AS yielded,
          COALESCE((SELECT SUM(volume_liters) FROM milk_movement
                    WHERE date_movement=%s AND movement_type='Расход'),0) AS spent,
          COALESCE((SELECT SUM(volume_liters) FROM milk_yield WHERE date_yield<=%s),0)
          - COALESCE((SELECT SUM(volume_liters) FROM milk_movement
                      WHERE date_movement<=%s AND movement_type='Расход'),0) AS balance
    """, (day, day, day, day))[0]
    recent = rows("""
        SELECT my.date_yield, my.shift::text AS shift, my.volume_liters,
               COALESCE(a.nickname, a.tag_number, g.name, 'Без указания') AS subject,
               u.full_name
        FROM milk_yield my
        LEFT JOIN animals a ON a.id=my.animal_id    
        LEFT JOIN "groups" g ON g.id=my.group_id
        LEFT JOIN users u ON u.id=my.user_id
        WHERE my.date_yield=%s
        ORDER BY my.id DESC LIMIT 8
    """, (day,))
    return render_template("index.html", day=day.isoformat(), summary=summary, recent=recent)

@app.route("/yield", methods=["GET", "POST"])
def yield_entry():
    refs = ref_data()
    if request.method == "POST":
        try:
            day = date.fromisoformat(request.form["date_yield"])
            shift = request.form["shift"]
            volume = float(request.form["volume_liters"])
            animal_id = request.form.get("animal_id") or None
            group_id = request.form.get("group_id") or None
            user_id = request.form.get("user_id") or None
            if volume <= 0:
                raise ValueError("Объём должен быть больше нуля.")
            if not animal_id and not group_id:
                raise ValueError("Выберите корову или группу.")
            if animal_id and group_id:
                raise ValueError("Выберите либо корову, либо группу.")
            execute("""INSERT INTO milk_yield
                (id, date_yield, shift, animal_id, group_id, volume_liters, user_id)
                VALUES ((SELECT COALESCE(MAX(id), 0) + 1 FROM milk_yield), %s,%s,%s,%s,%s,%s)""",
                (day, shift, animal_id, group_id, volume, user_id))
            flash("Надой сохранён.", "success")
            return redirect(url_for("home", day=day.isoformat()))
        except Exception as e:
            flash(f"Не удалось сохранить надой: {e}", "error")
    return render_template("yield.html", refs=refs, today=date.today().isoformat())

@app.route("/movement", methods=["GET", "POST"])
def movement_entry():
    refs = ref_data()
    if request.method == "POST":
        try:
            day = date.fromisoformat(request.form["date_movement"])
            movement_type = request.form["movement_type"]
            direction = request.form.get("direction") or None
            volume = float(request.form["volume_liters"])
            user_id = request.form.get("user_id") or None
            comment = request.form.get("comment", "").strip() or None
            if volume <= 0:
                raise ValueError("Объём должен быть больше нуля.")
            if movement_type == "Расход" and not direction:
                raise ValueError("Для расхода выберите направление.")
            execute("""INSERT INTO milk_movement
                (id, date_movement, movement_type, direction, volume_liters, comment, user_id)
                VALUES ((SELECT COALESCE(MAX(id), 0) + 1 FROM milk_movement), %s,%s,%s,%s,%s,%s)""",
                (day, movement_type, direction, volume, comment, user_id))
            flash("Движение молока сохранено.", "success")
            return redirect(url_for("home", day=day.isoformat()))
        except Exception as e:
            flash(f"Не удалось сохранить движение: {e}", "error")
    return render_template("movement.html", refs=refs, today=date.today().isoformat())

@app.route("/journal")
def journal():
    start = request.args.get("start", date.today().replace(day=1).isoformat())
    end = request.args.get("end", date.today().isoformat())
    yields = rows("""SELECT my.date_yield AS day, my.shift::text AS shift, my.volume_liters AS liters,
                     COALESCE(a.nickname, a.tag_number, g.name, 'Без указания') AS subject,
                     u.full_name AS person
                     FROM milk_yield my LEFT JOIN animals a ON a.id=my.animal_id
                     LEFT JOIN "groups" g ON g.id=my.group_id
                     LEFT JOIN users u ON u.id=my.user_id
                     WHERE my.date_yield BETWEEN %s AND %s ORDER BY my.date_yield DESC, my.id DESC""",
                  (start, end))
    movements = rows("""SELECT mm.date_movement AS day, mm.movement_type::text AS movement_type,
                        mm.direction::text AS direction, mm.volume_liters AS liters, mm.comment,
                        u.full_name AS person FROM milk_movement mm
                        LEFT JOIN users u ON u.id=mm.user_id
                        WHERE mm.date_movement BETWEEN %s AND %s
                        ORDER BY mm.date_movement DESC, mm.id DESC""", (start, end))
    return render_template("journal.html", yields=yields, movements=movements, start=start, end=end)

def report_data(start, end):
    return rows("""
      WITH days AS (
        SELECT generate_series(%s::date, %s::date, interval '1 day')::date AS d
      ), y AS (
        SELECT date_yield d, SUM(volume_liters) income
        FROM milk_yield WHERE date_yield BETWEEN %s AND %s GROUP BY date_yield
      ), m AS (
        SELECT date_movement d,
          SUM(CASE WHEN direction='Отгрузка на завод' THEN volume_liters ELSE 0 END) sold,
          SUM(CASE WHEN direction='Выпойка телят' THEN volume_liters ELSE 0 END) calves,
          SUM(CASE WHEN direction='Столовая' THEN volume_liters ELSE 0 END) canteen,
          SUM(CASE WHEN movement_type='Расход' THEN volume_liters ELSE 0 END) expense
        FROM milk_movement WHERE date_movement BETWEEN %s AND %s GROUP BY date_movement
      )
      SELECT days.d AS day, COALESCE(y.income,0) AS income,
        COALESCE(m.sold,0) AS sold, COALESCE(m.calves,0) AS calves,
        COALESCE(m.canteen,0) AS canteen, COALESCE(m.expense,0) AS expense,
        SUM(COALESCE(y.income,0)-COALESCE(m.expense,0)) OVER (ORDER BY days.d) AS balance
      FROM days LEFT JOIN y ON y.d=days.d LEFT JOIN m ON m.d=days.d ORDER BY days.d
    """, (start, end, start, end, start, end))

def sp21_data(day):
    return rows("""
      SELECT u.full_name AS worker,
        COUNT(DISTINCT my.animal_id) AS cows,
        SUM(CASE WHEN my.shift='Утро' THEN my.volume_liters ELSE 0 END) AS morning,
        SUM(CASE WHEN my.shift='День' THEN my.volume_liters ELSE 0 END) AS daytime,
        SUM(CASE WHEN my.shift='Вечер' THEN my.volume_liters ELSE 0 END) AS evening,
        SUM(my.volume_liters) AS total,
        ROUND(AVG(ld.fat_percent)::numeric,2) AS fat,
        ROUND(SUM(my.volume_liters*COALESCE(ld.fat_percent,0)/100)::numeric,2) AS fat_units
      FROM milk_yield my LEFT JOIN users u ON my.user_id=u.id
      LEFT JOIN laboratory_data ld ON
        ((my.animal_id IS NOT NULL AND ld.animal_id=my.animal_id) OR
         (my.animal_id IS NULL AND ld.yield_id=my.id))
        AND my.date_yield=ld.date_analyze
      WHERE my.date_yield=%s
      GROUP BY u.full_name ORDER BY u.full_name
    """, (day,))

@app.route("/reports")
def reports():
    return render_template("reports.html", today=date.today().isoformat(),
                           month_start=date.today().replace(day=1).isoformat())

@app.route("/export/sp23")
def export_sp23():
    start = request.args.get("start", date.today().replace(day=1).isoformat())
    end = request.args.get("end", date.today().isoformat())
    data = report_data(start, end)
    wb = Workbook()
    ws = wb.active
    ws.title = "СП-23"
    ws.append(["Ведомость учёта движения молока (проектный отчёт)"])
    ws.append([f"Период: {start} — {end}"])
    ws.append(["Дата", "Поступило (л)", "Реализовано (л)", "Выпойка телятам (л)",
               "Столовая (л)", "Всего расход (л)", "Остаток на конец дня (л)"])
    for r in data:
        ws.append([r["day"], float(r["income"]), float(r["sold"]), float(r["calves"]),
                   float(r["canteen"]), float(r["expense"]), float(r["balance"])])
    for cell in ws[3]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="286090")
        cell.alignment = Alignment(wrap_text=True, vertical="center")
    ws.freeze_panes = "A4"
    ws.auto_filter.ref = f"A3:G{ws.max_row}"
    for col, width in {"A":15,"B":20,"C":20,"D":24,"E":18,"F":21,"G":27}.items():
        ws.column_dimensions[col].width = width
    out = BytesIO()
    wb.save(out); out.seek(0)
    return send_file(out, as_attachment=True, download_name=f"SP-23_{start}_{end}.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.route("/export/sp21")
def export_sp21():
    day = request.args.get("day", date.today().isoformat())
    data = sp21_data(day)

    wb = Workbook()
    ws = wb.active
    ws.title = "СП-21"

    # Настройки ширины колонок
    for col, width in {"A": 8, "B": 30, "C": 10, "D": 15, "E": 15, "F": 15, "G": 15, "H": 15, "I": 15}.items():
        ws.column_dimensions[col].width = width

    # Объединяем ячейки для заголовка
    ws.merge_cells("A1:I1")
    ws["A1"] = "Типовая межотраслевая форма № СП-21"
    ws["A1"].font = Font(bold=True, size=12)
    ws["A1"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A2:I2")
    ws["A2"] = "ЖУРНАЛ УЧЕТА НАДОЯ МОЛОКА"
    ws["A2"].font = Font(bold=True, size=14)
    ws["A2"].alignment = Alignment(horizontal="center")

    ws.merge_cells("A4:I4")
    ws["A4"] = f"Дата: {day}"

    # Заголовки таблицы
    headers = ["№", "ФИО доярки (мастера машинной дойки)", "Обслуживалось коров",
               "Надоено молока (утром)", "Надоено молока (в полдень)", "Надоено молока (вечером)",
               "Всего надоено", "Жирность, %", "Жиро-единиц"]
    ws.append([])  # Пустая строка для отступа
    ws.append(headers)

    # Оформление заголовков таблицы
    header_row = ws.max_row
    border_style = Border(left=Side(style='thin'), right=Side(style='thin'),
                          top=Side(style='thin'), bottom=Side(style='thin'))

    for col_idx, cell in enumerate(ws[header_row], start=1):
        cell.font = Font(bold=True, color="000000")
        cell.fill = PatternFill("solid", fgColor="D9E1F2")
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        cell.border = border_style

    # Заполнение данными
    row_num = 1
    for r in data:
        row_data = [
            row_num,
            r["worker"] or "Не указан",
            r["cows"] or 0,
            float(r["morning"] or 0),
            float(r["daytime"] or 0),
            float(r["evening"] or 0),
            float(r["total"] or 0),
            "" if r["fat"] is None else float(r["fat"]),
            float(r["fat_units"] or 0)
        ]
        ws.append(row_data)
        row_num += 1

        # Оформление строки данных
        current_row = ws.max_row
        for cell in ws[current_row]:
            cell.border = border_style
            if cell.column_letter in ['C', 'D', 'E', 'F', 'G', 'H', 'I']:
                cell.alignment = Alignment(horizontal="center")
                # Форматируем числа
                if cell.column_letter in ['D', 'E', 'F', 'G', 'I']:
                    cell.number_format = '0.00'

    # Если данных нет, добавляем пустую строку для красоты
    if row_num == 1:
        ws.append(["", "За выбранную дату записей нет", "", "", "", "", "", "", ""])
        for cell in ws[ws.max_row]:
            cell.border = border_style

    out = BytesIO()
    wb.save(out)
    out.seek(0)

    return send_file(
        out,
        as_attachment=True,
        download_name=f"SP-21_{day}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=int(os.getenv("PORT", "5000")))
