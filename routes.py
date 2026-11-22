from datetime import date
from flask import Blueprint, request, redirect, url_for, render_template, flash
from db import rows, execute, ref_data

routes_bp = Blueprint('routes', __name__)


@routes_bp.route("/")
def home():
    selected = request.args.get("day", date.today().isoformat())
    try:
        day = date.fromisoformat(selected)
    except ValueError:
        day = date.today()

    summary = rows("""
        SELECT
            COALESCE((SELECT SUM(volume_liters) FROM milk_yield WHERE date_yield=%s),0) AS yielded,
            COALESCE((SELECT SUM(volume_liters) FROM milk_movement WHERE date_movement=%s AND movement_type='Расход'),0) AS spent,
            COALESCE((SELECT SUM(volume_liters) FROM milk_yield WHERE date_yield <=%s),0)
            - COALESCE((SELECT SUM(volume_liters) FROM milk_movement WHERE date_movement <=%s AND movement_type='Расход'),0) AS balance
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


@routes_bp.route("/yield", methods=["GET", "POST"])
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

            execute("""
                INSERT INTO milk_yield (id, date_yield, shift, animal_id, group_id, volume_liters, user_id)
                VALUES ((SELECT COALESCE(MAX(id), 0) + 1 FROM milk_yield), %s,%s,%s,%s,%s,%s)
            """, (day, shift, animal_id, group_id, volume, user_id))

            flash("Надой сохранён.", "success")
            return redirect(url_for("routes.home", day=day.isoformat()))
        except Exception as e:
            flash(f"Не удалось сохранить надой: {e}", "error")

    return render_template("yield.html", refs=refs, today=date.today().isoformat())


@routes_bp.route("/movement", methods=["GET", "POST"])
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

            execute("""
                INSERT INTO milk_movement (id, date_movement, movement_type, direction, volume_liters, comment, user_id)
                VALUES ((SELECT COALESCE(MAX(id), 0) + 1 FROM milk_movement), %s,%s,%s,%s,%s,%s)
            """, (day, movement_type, direction, volume, comment, user_id))

            flash("Движение молока сохранено.", "success")
            return redirect(url_for("routes.home", day=day.isoformat()))
        except Exception as e:
            flash(f"Не удалось сохранить движение: {e}", "error")

    return render_template("movement.html", refs=refs, today=date.today().isoformat())


@routes_bp.route("/journal")
def journal():
    start = request.args.get("start", date.today().replace(day=1).isoformat())
    end = request.args.get("end", date.today().isoformat())

    yields = rows("""
        SELECT my.date_yield AS day, my.shift::text AS shift, my.volume_liters AS liters,
               COALESCE(a.nickname, a.tag_number, g.name, 'Без указания') AS subject,
               u.full_name AS person
        FROM milk_yield my 
        LEFT JOIN animals a ON a.id=my.animal_id
        LEFT JOIN "groups" g ON g.id=my.group_id
        LEFT JOIN users u ON u.id=my.user_id
        WHERE my.date_yield BETWEEN %s AND %s 
        ORDER BY my.date_yield DESC, my.id DESC
    """, (start, end))

    movements = rows("""
        SELECT mm.date_movement AS day, mm.movement_type::text AS movement_type,
               mm.direction::text AS direction, mm.volume_liters AS liters, mm.comment,
               u.full_name AS person 
        FROM milk_movement mm
        LEFT JOIN users u ON u.id=mm.user_id
        WHERE mm.date_movement BETWEEN %s AND %s
        ORDER BY mm.date_movement DESC, mm.id DESC
    """, (start, end))

    return render_template("journal.html", yields=yields, movements=movements, start=start, end=end)


@routes_bp.route("/reports")
def reports():
    return render_template("reports.html", today=date.today().isoformat(),
                           month_start=date.today().replace(day=1).isoformat())