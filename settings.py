from flask import Blueprint, request, redirect, url_for, render_template, flash
from db import rows, execute, ref_data

settings_bp = Blueprint('settings', __name__)

def get_settings():
    """Получить все настройки как словарь"""
    result = rows('SELECT "key", "value" FROM settings')
    return {row['key']: row['value'] for row in result}

def update_settings(settings_dict):
    """Обновить настройки (upsert)"""
    for key, value in settings_dict.items():
        execute("""
            INSERT INTO settings ("key", "value", updated_at)
            VALUES (%s, %s, CURRENT_TIMESTAMP)
            ON CONFLICT ("key") DO UPDATE
            SET "value" = EXCLUDED."value", updated_at = CURRENT_TIMESTAMP
        """, (key, value))

@settings_bp.route("/settings", methods=["GET", "POST"])
def settings():
    if request.method == "POST":
        try:
            settings_data = {
                'org_name': request.form.get('org_name', '').strip(),
                'department': request.form.get('department', '').strip(),
                'farm': request.form.get('farm', '').strip(),
                'brigade': request.form.get('brigade', '').strip(),
                'responsible_person': request.form.get('responsible_person', '').strip(),
                'milk_price': request.form.get('milk_price', '0').strip(),
                'debit_account': request.form.get('debit_account', '').strip(),
                'credit_account': request.form.get('credit_account', '').strip()
            }

            if not settings_data['org_name']:
                raise ValueError("Название организации обязательно")

            try:
                float(settings_data['milk_price'])
            except ValueError:
                raise ValueError("Цена за молоко должна быть числом")

            update_settings(settings_data)
            flash("Настройки сохранены", "success")
            return redirect(url_for("settings.settings"))

        except Exception as e:
            flash(f"Ошибка сохранения: {e}", "error")

    current_settings = get_settings()
    refs = ref_data()
    return render_template("settings.html", settings=current_settings, refs=refs)

@settings_bp.route("/settings/users/add", methods=["POST"])
def add_user():
    try:
        full_name = request.form.get("full_name", "").strip()
        role = request.form.get("role", "").strip()

        if not full_name:
            raise ValueError("ФИО обязательно")
        if not role:
            raise ValueError("Роль обязательна")

        execute("""
            INSERT INTO users (id, full_name, role)
            VALUES ((SELECT COALESCE(MAX(id), 0) + 1 FROM users), %s, %s)
        """, (full_name, role))

        flash(f"Сотрудник '{full_name}' добавлен", "success")
    except Exception as e:
        flash(f"Ошибка добавления: {e}", "error")

    return redirect(url_for("settings.settings"))

@settings_bp.route("/settings/users/delete/<int:user_id>", methods=["POST"])
def delete_user(user_id):
    try:
        count = rows("SELECT COUNT(*) AS cnt FROM milk_yield WHERE user_id=%s", (user_id,))[0]['cnt']
        if count > 0:
            raise ValueError(f"Нельзя удалить: у сотрудника {count} записей надоев")

        execute("DELETE FROM users WHERE id=%s", (user_id,))
        flash("Сотрудник удалён", "success")
    except Exception as e:
        flash(f"Ошибка удаления: {e}", "error")

    return redirect(url_for("settings.settings"))

@settings_bp.route("/settings/animals/add", methods=["POST"])
def add_animal():
    try:
        tag_number = request.form.get("tag_number", "").strip() or None
        nickname = request.form.get("nickname", "").strip() or None
        group_id = request.form.get("group_id") or None
        status = request.form.get("status", "Дойная").strip()
        date_of_birth = request.form.get("date_of_birth") or None

        if not tag_number:
            raise ValueError("Номер (бирка) обязателен")

        execute("""
            INSERT INTO animals (id, tag_number, nickname, group_id, status, date_of_birth)
            VALUES ((SELECT COALESCE(MAX(id), 0) + 1 FROM animals), %s, %s, %s, %s, %s)
        """, (tag_number, nickname, group_id, status, date_of_birth))

        flash("Животное добавлено", "success")
    except Exception as e:
        flash(f"Ошибка добавления: {e}", "error")

    return redirect(url_for("settings.settings"))

@settings_bp.route("/settings/animals/delete/<int:animal_id>", methods=["POST"])
def delete_animal(animal_id):
    try:
        count = rows("SELECT COUNT(*) AS cnt FROM milk_yield WHERE animal_id=%s", (animal_id,))[0]['cnt']
        if count > 0:
            raise ValueError(f"Нельзя удалить: у животного {count} записей надоев")

        execute("DELETE FROM animals WHERE id=%s", (animal_id,))
        flash("Животное удалено", "success")
    except Exception as e:
        flash(f"Ошибка удаления: {e}", "error")

    return redirect(url_for("settings.settings"))

@settings_bp.route("/settings/groups/add", methods=["POST"])
def add_group():
    try:
        name = request.form.get("name", "").strip()

        if not name:
            raise ValueError("Название группы обязательно")

        execute("""
            INSERT INTO groups (id, name)
            VALUES ((SELECT COALESCE(MAX(id), 0) + 1 FROM groups), %s)
        """, (name,))

        flash(f"Группа '{name}' добавлена", "success")
    except Exception as e:
        flash(f"Ошибка добавления: {e}", "error")

    return redirect(url_for("settings.settings"))

@settings_bp.route("/settings/groups/delete/<int:group_id>", methods=["POST"])
def delete_group(group_id):
    try:
        count = rows("SELECT COUNT(*) AS cnt FROM animals WHERE group_id=%s", (group_id,))[0]['cnt']
        if count > 0:
            raise ValueError(f"Нельзя удалить: в группе {count} животных")

        execute('DELETE FROM groups WHERE id=%s', (group_id,))
        flash("Группа удалена", "success")
    except Exception as e:
        flash(f"Ошибка удаления: {e}", "error")

    return redirect(url_for("settings.settings"))