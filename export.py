from datetime import date
from io import BytesIO
from flask import Blueprint, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from db import rows
from settings import get_settings

export_bp = Blueprint('export', __name__)


def report_data(start, end):
    """Данные для СП-23"""
    return rows("""
        WITH days AS (
            SELECT generate_series(%s::date, %s::date, interval '1 day')::date AS d
        ), y AS (
            SELECT date_yield d, SUM(volume_liters) income
            FROM milk_yield WHERE date_yield BETWEEN %s AND %s GROUP BY date_yield
        ), m AS (
            SELECT date_movement d,
                   SUM(CASE WHEN direction='Отгрузка на завод' THEN volume_liters ELSE 0 END) sold,
                   SUM(CASE WHEN direction='Переработка' THEN volume_liters ELSE 0 END) processed,
                   SUM(CASE WHEN direction='Выпойка телят' THEN volume_liters ELSE 0 END) calves,
                   SUM(CASE WHEN direction='Столовая' THEN volume_liters ELSE 0 END) canteen,
                   SUM(CASE WHEN movement_type='Расход' THEN volume_liters ELSE 0 END) expense
            FROM milk_movement WHERE date_movement BETWEEN %s AND %s GROUP BY date_movement
        )
        SELECT days.d AS day, COALESCE(y.income,0) AS income,
               COALESCE(m.sold,0) AS sold,
               COALESCE(m.processed,0) AS processed,
               COALESCE(m.calves,0) AS calves,
               COALESCE(m.canteen,0) AS canteen,
               COALESCE(m.expense,0) AS expense,
               SUM(COALESCE(y.income,0)-COALESCE(m.expense,0)) OVER (ORDER BY days.d) AS balance
        FROM days LEFT JOIN y ON y.d=days.d LEFT JOIN m ON m.d=days.d ORDER BY days.d
    """, (start, end, start, end, start, end))


def sp21_data(day):
    """Данные для СП-21"""
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


@export_bp.route("/export/sp23")
def export_sp23():
    start = request.args.get("start", date.today().replace(day=1).isoformat())
    end = request.args.get("end", date.today().isoformat())
    data = report_data(start, end)
    settings = get_settings()  # Получаем настройки

    d_start = date.fromisoformat(start)
    d_end = date.fromisoformat(end)

    # Формируем строку периода для шапки
    if d_start.month == d_end.month and d_start.year == d_end.year:
        period_str = f"{d_start.strftime('%B %Y').lower()}"
    else:
        period_str = f"{d_start.strftime('%d.%m.%Y')} — {d_end.strftime('%d.%m.%Y')}"

    wb = Workbook()

    # --- СТИЛИ ---
    thin = Side(style='thin')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_bottom = Border(bottom=thin)

    font_title = Font(name='Times New Roman', size=14, bold=True)
    font_header = Font(name='Times New Roman', size=10, bold=True)
    font_normal = Font(name='Times New Roman', size=10)
    font_small = Font(name='Times New Roman', size=9)
    font_tiny = Font(name='Times New Roman', size=8)

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')

    def apply_border(ws, min_row, max_row, min_col, max_col):
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(row=r, column=c).border = border_all

    # =========================================================================
    # ЛИСТ 1: стр1 (ЛИЦЕВАЯ СТОРОНА, ДНИ 1-18)
    # =========================================================================
    ws1 = wb.active
    ws1.title = "стр1"
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 1
    ws1.page_setup.orientation = 'landscape'
    ws1.page_margins.left = 0.4
    ws1.page_margins.right = 0.4

    # Колонки: A(число), B(поступило), C-I(расход), J(остаток), K(жирность)
    widths1 = {'A': 7, 'B': 12, 'C': 9, 'D': 11, 'E': 13, 'F': 10, 'G': 11, 'H': 14, 'I': 11, 'J': 12, 'K': 14}
    for col, w in widths1.items():
        ws1.column_dimensions[col].width = w

    # --- ШАПКА СПРАВА ---
    ws1['E2'] = "Типовая межотраслевая форма № СП-23"
    ws1['E2'].font = font_normal
    ws1['E2'].alignment = align_right
    ws1.merge_cells('E2:K2')

    ws1['E3'] = "Утверждена постановлением Госкомстата России"
    ws1['E3'].font = font_normal
    ws1['E3'].alignment = align_right
    ws1.merge_cells('E3:K3')

    ws1['E4'] = "от 29.09.97 № 68"
    ws1['E4'].font = font_normal
    ws1['E4'].alignment = align_right
    ws1.merge_cells('E4:K4')

    # --- НАЗВАНИЕ ПО ЦЕНТРУ ---
    ws1['A6'] = "ВЕДОМОСТЬ №"
    ws1['A6'].font = font_title
    ws1['A6'].alignment = align_center
    ws1.merge_cells('A6:D6')

    ws1['A7'] = "учета движения молока"
    ws1['A7'].font = font_title
    ws1['A7'].alignment = align_center
    ws1.merge_cells('A7:D7')

    ws1['A8'] = f"за {period_str}"
    ws1['A8'].font = font_normal
    ws1['A8'].alignment = align_center
    ws1.merge_cells('A8:D8')

    # Коды справа
    ws1['H8'] = "Коды"
    ws1['H8'].font = font_normal
    ws1['H8'].alignment = align_center
    ws1.merge_cells('H8:K8')

    ws1['H9'] = "Форма по ОКУД"
    ws1['H9'].font = font_normal
    ws1['H9'].alignment = align_right
    ws1['I9'] = "0325023"
    ws1['I9'].alignment = align_center

    ws1['H10'] = "Дата составления"
    ws1['H10'].font = font_normal
    ws1['H10'].alignment = align_right
    ws1['I10'] = date.today().strftime("%d.%m.%Y")
    ws1['I10'].border = border_bottom

    # Поля организации — ИЗ НАСТРОЕК
    ws1['A10'] = "Организация"
    ws1['A10'].font = font_normal
    ws1.merge_cells('B10:D10')
    ws1['B10'] = settings.get('org_name', '')
    ws1['B10'].border = border_bottom

    ws1['H11'] = "по ОКПО"
    ws1['H11'].font = font_normal
    ws1['H11'].alignment = align_right
    ws1['I11'].border = border_bottom

    ws1['A11'] = "Отделение (участок)"
    ws1['A11'].font = font_normal
    ws1.merge_cells('B11:D11')
    ws1['B11'] = settings.get('department', '')
    ws1['B11'].border = border_bottom

    ws1['A12'] = "Ферма"
    ws1['A12'].font = font_normal
    ws1.merge_cells('B12:D12')
    ws1['B12'] = settings.get('farm', '')
    ws1['B12'].border = border_bottom

    ws1['A13'] = "Бригада"
    ws1['A13'].font = font_normal
    ws1.merge_cells('B13:D13')
    ws1['B13'] = settings.get('brigade', '')
    ws1['B13'].border = border_bottom

    ws1['A14'] = "Материально ответственное лицо"
    ws1['A14'].font = font_normal
    ws1.merge_cells('B14:D14')
    ws1['B14'] = settings.get('responsible_person', '')
    ws1['B14'].border = border_bottom

    # --- ШАПКА ТАБЛИЦЫ (строки 16-18) ---
    # Число месяца
    ws1['A16'] = "Число\nмеся-\nца"
    ws1['A16'].font = font_header
    ws1['A16'].alignment = align_center
    ws1.merge_cells('A16:A18')

    # Поступило за сутки, кг
    ws1['B16'] = "Поступило за сутки, кг"
    ws1['B16'].font = font_header
    ws1['B16'].alignment = align_center
    ws1.merge_cells('B16:B17')
    ws1['B18'] = "всего"
    ws1['B18'].font = font_header
    ws1['B18'].alignment = align_center

    # Расход, кг (объединяет C-I в строке 16)
    ws1['C16'] = "Расход, кг"
    ws1['C16'].font = font_header
    ws1['C16'].alignment = align_center
    ws1.merge_cells('C16:I16')

    # Подзаголовки расхода (строка 17)
    ws1['C17'] = "всего"
    ws1['C17'].font = font_header
    ws1['C17'].alignment = align_center
    ws1.merge_cells('C17:C18')

    ws1['D17'] = "реа-\nлизо-\nвано"
    ws1['D17'].font = font_header
    ws1['D17'].alignment = align_center
    ws1.merge_cells('D17:D18')

    ws1['E17'] = "пере-\nдано\nв пере-\nработ-\nку"
    ws1['E17'].font = font_header
    ws1['E17'].alignment = align_center
    ws1.merge_cells('E17:E18')

    ws1['F17'] = "на выпойку"
    ws1['F17'].font = font_header
    ws1['F17'].alignment = align_center
    ws1.merge_cells('F17:G17')

    ws1['H17'] = "на об-\nщест-\nвенное\nпита-\nние"
    ws1['H17'].font = font_header
    ws1['H17'].alignment = align_center
    ws1.merge_cells('H17:H18')

    ws1['I17'] = "всего\nза\nдень"
    ws1['I17'].font = font_header
    ws1['I17'].alignment = align_center
    ws1.merge_cells('I17:I18')

    # Под "на выпойку" - телятам и поросятам (строка 18)
    ws1['F18'] = "теля-\nтам"
    ws1['F18'].font = font_header
    ws1['F18'].alignment = align_center

    ws1['G18'] = "поро-\nся-\nтам"
    ws1['G18'].font = font_header
    ws1['G18'].alignment = align_center

    # Остаток на конец дня
    ws1['J16'] = "Оста-\nток на\nконец\nдня,\nкг"
    ws1['J16'].font = font_header
    ws1['J16'].alignment = align_center
    ws1.merge_cells('J16:J18')

    # Контрольное определение жирности
    ws1['K16'] = "Контрольное\nопределение\nжирности\nмолока,\n%"
    ws1['K16'].font = font_header
    ws1['K16'].alignment = align_center
    ws1.merge_cells('K16:K18')

    # Рамки шапки
    apply_border(ws1, 16, 18, 1, 11)

    # --- ЗАПОЛНЕНИЕ ДАННЫМИ (18 строк, дни 1-18) ---
    start_row = 19
    for i in range(1, 19):
        row = start_row + i - 1
        ws1[f'A{row}'] = i
        ws1[f'A{row}'].alignment = align_center

        # Ищем данные для этого дня месяца
        day_data = None
        for r in data:
            if r['day'].day == i and r['day'].month == d_start.month and r['day'].year == d_start.year:
                day_data = r
                break

        if day_data:
            ws1[f'B{row}'] = float(day_data['income'] or 0)
            ws1[f'C{row}'] = float(day_data['expense'] or 0)
            ws1[f'D{row}'] = float(day_data['sold'] or 0)
            ws1[f'E{row}'] = float(day_data.get('processed', 0) or 0)  # Переработка
            ws1[f'F{row}'] = float(day_data['calves'] or 0)
            # G - поросятам (нет данных)
            ws1[f'H{row}'] = float(day_data['canteen'] or 0)
            ws1[f'I{row}'] = float(day_data['expense'] or 0)
            ws1[f'J{row}'] = float(day_data['balance'] or 0)

        apply_border(ws1, row, row, 1, 11)

        # Форматирование чисел
        for col in 'BCDFHIJ':
            ws1[f'{col}{row}'].alignment = align_center
            ws1[f'{col}{row}'].number_format = '0.00'

    # =========================================================================
    # ЛИСТ 2: стр2 (ОБОРОТНАЯ СТОРОНА, ДНИ 19-31 + ИТОГО)
    # =========================================================================
    ws2 = wb.create_sheet("стр2")
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 1
    ws2.page_setup.orientation = 'landscape'
    ws2.page_margins.left = 0.4
    ws2.page_margins.right = 0.4

    for col, w in widths1.items():
        ws2.column_dimensions[col].width = w

    # Заголовок оборотной стороны
    ws2['A1'] = "Оборотная сторона формы № СП-23"
    ws2['A1'].font = font_header
    ws2['A1'].alignment = align_right
    ws2.merge_cells('A1:K1')

    # --- ШАПКА ТАБЛИЦЫ (такая же, как на стр1) ---
    ws2['A2'] = "Число\nмеся-\nца"
    ws2['A2'].font = font_header
    ws2['A2'].alignment = align_center
    ws2.merge_cells('A2:A4')

    ws2['B2'] = "Поступило за сутки, кг"
    ws2['B2'].font = font_header
    ws2['B2'].alignment = align_center
    ws2.merge_cells('B2:B3')
    ws2['B4'] = "всего"
    ws2['B4'].font = font_header
    ws2['B4'].alignment = align_center

    ws2['C2'] = "Расход, кг"
    ws2['C2'].font = font_header
    ws2['C2'].alignment = align_center
    ws2.merge_cells('C2:I2')

    ws2['C3'] = "всего"
    ws2['C3'].font = font_header
    ws2['C3'].alignment = align_center
    ws2.merge_cells('C3:C4')

    ws2['D3'] = "реа-\nлизо-\nвано"
    ws2['D3'].font = font_header
    ws2['D3'].alignment = align_center
    ws2.merge_cells('D3:D4')

    ws2['E3'] = "пере-\nдано\nв пере-\nработ-\nку"
    ws2['E3'].font = font_header
    ws2['E3'].alignment = align_center
    ws2.merge_cells('E3:E4')

    ws2['F3'] = "на выпойку"
    ws2['F3'].font = font_header
    ws2['F3'].alignment = align_center
    ws2.merge_cells('F3:G3')

    ws2['H3'] = "на об-\nщест-\nвенное\nпита-\nние"
    ws2['H3'].font = font_header
    ws2['H3'].alignment = align_center
    ws2.merge_cells('H3:H4')

    ws2['I3'] = "всего\nза\nдень"
    ws2['I3'].font = font_header
    ws2['I3'].alignment = align_center
    ws2.merge_cells('I3:I4')

    ws2['F4'] = "теля-\nтам"
    ws2['F4'].font = font_header
    ws2['F4'].alignment = align_center

    ws2['G4'] = "поро-\nся-\nтам"
    ws2['G4'].font = font_header
    ws2['G4'].alignment = align_center

    ws2['J2'] = "Оста-\nток на\nконец\nдня,\nкг"
    ws2['J2'].font = font_header
    ws2['J2'].alignment = align_center
    ws2.merge_cells('J2:J4')

    ws2['K2'] = "Контрольное\nопределение\nжирности\nмолока,\n%"
    ws2['K2'].font = font_header
    ws2['K2'].alignment = align_center
    ws2.merge_cells('K2:K4')

    apply_border(ws2, 2, 4, 1, 11)

    # --- ЗАПОЛНЕНИЕ ДАННЫМИ (дни 19-31) ---
    start_row2 = 5
    for i in range(19, 32):
        row = start_row2 + (i - 19)
        ws2[f'A{row}'] = i
        ws2[f'A{row}'].alignment = align_center

        day_data = None
        for r in data:
            if r['day'].day == i and r['day'].month == d_start.month and r['day'].year == d_start.year:
                day_data = r
                break

        if day_data:
            ws2[f'B{row}'] = float(day_data['income'] or 0)
            ws2[f'C{row}'] = float(day_data['expense'] or 0)
            ws2[f'D{row}'] = float(day_data['sold'] or 0)
            ws2[f'E{row}'] = float(day_data.get('processed', 0) or 0)  # Переработка
            ws2[f'F{row}'] = float(day_data['calves'] or 0)
            ws2[f'H{row}'] = float(day_data['canteen'] or 0)
            ws2[f'I{row}'] = float(day_data['expense'] or 0)
            ws2[f'J{row}'] = float(day_data['balance'] or 0)

        apply_border(ws2, row, row, 1, 11)

        for col in 'BCDFHIJ':
            ws2[f'{col}{row}'].alignment = align_center
            ws2[f'{col}{row}'].number_format = '0.00'

    # --- ИТОГО ЗА МЕСЯЦ ---
    total_row = start_row2 + 13  # после 31-го числа (строка 18)

    # "Итого за месяц" — выравнивание по левому краю
    ws2[f'A{total_row}'] = "Итого за месяц"
    ws2[f'A{total_row}'].font = Font(name='Times New Roman', size=10, bold=True)
    ws2[f'A{total_row}'].alignment = align_left
    ws2.merge_cells(f'A{total_row}:B{total_row}')

    # Крестики в колонках J, K (под жирностью и качеством)
    ws2[f'J{total_row}'] = "Х"
    ws2[f'J{total_row}'].alignment = align_center
    ws2[f'K{total_row}'] = "Х"
    ws2[f'K{total_row}'].alignment = align_center

    apply_border(ws2, total_row, total_row, 1, 11)

    # --- Цена, Сумма, Дебет, Кредит ---
    milk_price = float(settings.get('milk_price', 0) or 0)
    total_income = sum(float(r["income"]) for r in data)
    total_sum = total_income * milk_price

    # Цена (по левому краю)
    ws2[f'A{total_row + 1}'] = "Цена, руб. коп."
    ws2[f'A{total_row + 1}'].font = font_normal
    ws2[f'A{total_row + 1}'].alignment = align_left
    ws2[f'K{total_row + 1}'] = milk_price
    ws2[f'K{total_row + 1}'].number_format = '0.00'
    ws2[f'K{total_row + 1}'].border = border_all

    # Сумма (по левому краю)
    ws2[f'A{total_row + 2}'] = "Сумма, руб. коп."
    ws2[f'A{total_row + 2}'].font = font_normal
    ws2[f'A{total_row + 2}'].alignment = align_left
    ws2[f'K{total_row + 2}'] = total_sum
    ws2[f'K{total_row + 2}'].number_format = '0.00'
    ws2[f'K{total_row + 2}'].border = border_all

    # Код синтетического и аналитического учета — ОБЪЕДИНЕНИЕ ПО ВЕРТИКАЛИ (на 2 строки)
    ws2[f'A{total_row + 3}'] = "Код синтетического\nи аналитического учета:"
    ws2[f'A{total_row + 3}'].font = font_normal
    ws2[f'A{total_row + 3}'].alignment = Alignment(horizontal='left', vertical='center', wrap_text=True)
    ws2.merge_cells(f'A{total_row + 3}:B{total_row + 4}')

    # "дебет" — в отдельной ячейке (по левому краю)
    ws2[f'C{total_row + 3}'] = "дебет"
    ws2[f'C{total_row + 3}'].font = font_normal
    ws2[f'C{total_row + 3}'].alignment = align_left
    ws2[f'K{total_row + 3}'] = settings.get('debit_account', '')
    ws2[f'K{total_row + 3}'].border = border_all

    # Крестики в строке дебет
    ws2[f'I{total_row + 3}'] = "Х"
    ws2[f'I{total_row + 3}'].alignment = align_center
    ws2[f'J{total_row + 3}'] = "Х"
    ws2[f'J{total_row + 3}'].alignment = align_center
    ws2[f'K{total_row + 3}'].alignment = align_center

    # "кредит" — в отдельной ячейке (по левому краю), строка ниже
    ws2[f'C{total_row + 4}'] = "кредит"
    ws2[f'C{total_row + 4}'].font = font_normal
    ws2[f'C{total_row + 4}'].alignment = align_left
    ws2[f'K{total_row + 4}'] = settings.get('credit_account', '')
    ws2[f'K{total_row + 4}'].border = border_all

    # Крестики в строке кредит
    ws2[f'I{total_row + 4}'] = "Х"
    ws2[f'I{total_row + 4}'].alignment = align_center
    ws2[f'J{total_row + 4}'] = "Х"
    ws2[f'J{total_row + 4}'].alignment = align_center
    ws2[f'K{total_row + 4}'].alignment = align_center

    # --- Подписи ---
    sig_row = total_row + 6

    ws2[f'C{sig_row}'] = "(должность)"
    ws2[f'C{sig_row}'].font = font_tiny
    ws2[f'C{sig_row}'].alignment = align_center
    ws2[f'C{sig_row + 1}'].border = border_bottom

    ws2[f'E{sig_row}'] = "(подпись)"
    ws2[f'E{sig_row}'].font = font_tiny
    ws2[f'E{sig_row}'].alignment = align_center
    ws2[f'E{sig_row + 1}'].border = border_bottom

    ws2[f'G{sig_row}'] = "(расшифровка подписи)"
    ws2[f'G{sig_row}'].font = font_tiny
    ws2[f'G{sig_row}'].alignment = align_center
    ws2[f'G{sig_row + 1}'].border = border_bottom

    # Вторая подпись
    sig_row2 = sig_row + 3
    ws2[f'C{sig_row2}'] = "(должность)"
    ws2[f'C{sig_row2}'].font = font_tiny
    ws2[f'C{sig_row2}'].alignment = align_center
    ws2[f'C{sig_row2 + 1}'].border = border_bottom

    ws2[f'E{sig_row2}'] = "(подпись)"
    ws2[f'E{sig_row2}'].font = font_tiny
    ws2[f'E{sig_row2}'].alignment = align_center
    ws2[f'E{sig_row2 + 1}'].border = border_bottom

    ws2[f'G{sig_row2}'] = "(расшифровка подписи)"
    ws2[f'G{sig_row2}'].font = font_tiny
    ws2[f'G{sig_row2}'].alignment = align_center
    ws2[f'G{sig_row2 + 1}'].border = border_bottom

    # Дата
    date_row = sig_row2 + 3
    ws2[f'C{date_row}'] = '"'
    ws2[f'D{date_row}'] = "___"
    ws2[f'E{date_row}'] = '"'
    ws2[f'F{date_row}'] = "__________ 20__ г."

    # =========================================================================
    # СОХРАНЕНИЕ
    # =========================================================================
    out = BytesIO()
    wb.save(out)
    out.seek(0)

    return send_file(
        out,
        as_attachment=True,
        download_name=f"SP-23_{start}_{end}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@export_bp.route("/export/sp21")
def export_sp21():
    day = request.args.get("day", date.today().isoformat())
    data = sp21_data(day)
    settings = get_settings()  # Получаем настройки

    d_obj = date.fromisoformat(day)
    date_str = d_obj.strftime("%d.%m.%Y")

    wb = Workbook()

    # --- СТИЛИ ---
    thin = Side(style='thin')
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    border_bottom = Border(bottom=thin)

    font_title = Font(name='Times New Roman', size=14, bold=True)
    font_header = Font(name='Times New Roman', size=10, bold=True)
    font_normal = Font(name='Times New Roman', size=10)
    font_small = Font(name='Times New Roman', size=9)
    font_tiny = Font(name='Times New Roman', size=8)

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_left = Alignment(horizontal='left', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center')

    def apply_border(ws, min_row, max_row, min_col, max_col):
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                ws.cell(row=r, column=c).border = border_all

    # =========================================================================
    # ЛИСТ 1: стр1 (ТИТУЛЬНЫЙ)
    # =========================================================================
    ws1 = wb.active
    ws1.title = "стр1"
    ws1.sheet_properties.pageSetUpPr.fitToPage = True
    ws1.page_setup.fitToWidth = 1
    ws1.page_setup.fitToHeight = 1
    ws1.page_margins.left = 0.5
    ws1.page_margins.right = 0.5

    # Шапка справа — сначала значения, потом объединение
    ws1['E2'] = "Типовая межотраслевая форма № СП-21"
    ws1['E2'].font = font_normal
    ws1['E2'].alignment = align_right
    ws1.merge_cells('E2:J2')

    ws1['E3'] = "Утверждена постановлением Госкомстата России"
    ws1['E3'].font = font_normal
    ws1['E3'].alignment = align_right
    ws1.merge_cells('E3:J3')

    ws1['E4'] = "от 29.09.97 № 68"
    ws1['E4'].font = font_normal
    ws1['E4'].alignment = align_right
    ws1.merge_cells('E4:J4')

    # Название по центру (левая часть)
    ws1['A6'] = "ЖУРНАЛ №"
    ws1['A6'].font = font_title
    ws1['A6'].alignment = align_center
    ws1.merge_cells('A6:F6')
    ws1['B6'].border = border_bottom

    ws1['H7'] = "Коды"
    ws1['H7'].font = font_normal
    ws1['H7'].alignment = align_center
    ws1.merge_cells('H7:J7')

    ws1['A7'] = "учета надоя молока"
    ws1['A7'].font = font_title
    ws1['A7'].alignment = align_center
    ws1.merge_cells('A7:F7')

    ws1['H8'] = "Форма по ОКУД"
    ws1['H8'].font = font_normal
    ws1['H8'].alignment = align_right
    ws1['I8'] = "0325021"
    ws1['I8'].alignment = align_center

    # Поля организации — ИЗ НАСТРОЕК
    ws1['A8'] = "Организация"
    ws1['A8'].font = font_normal
    ws1.merge_cells('B8:F8')
    ws1['B8'] = settings.get('org_name', '')
    ws1['B8'].border = border_bottom

    ws1['H9'] = "по ОКПО"
    ws1['H9'].font = font_normal
    ws1['H9'].alignment = align_right
    ws1['I9'].border = border_bottom

    ws1['A10'] = "Отделение"
    ws1['A10'].font = font_normal
    ws1.merge_cells('B10:F10')
    ws1['B10'] = settings.get('department', '')
    ws1['B10'].border = border_bottom

    ws1['A11'] = "Ферма"
    ws1['A11'].font = font_normal
    ws1.merge_cells('B11:F11')
    ws1['B11'] = settings.get('farm', '')
    ws1['B11'].border = border_bottom

    ws1['A12'] = "Бригада"
    ws1['A12'].font = font_normal
    ws1.merge_cells('B12:F12')
    ws1['B12'] = settings.get('brigade', '')
    ws1['B12'].border = border_bottom

    ws1['H13'] = "Дата составления"
    ws1['H13'].font = font_normal
    ws1['H13'].alignment = align_right
    ws1['I13'] = date_str
    ws1['I13'].border = border_bottom

    # Итоги по центру — ИЗ НАСТРОЕК И ДАННЫХ
    total_milk = sum(float(r.get('total', 0) or 0) for r in data)
    milk_price = float(settings.get('milk_price', 0) or 0)
    total_sum = total_milk * milk_price

    ws1['A14'] = f"Всего по документу оприходовано молока: {total_milk:.2f} кг"
    ws1['A14'].font = font_normal
    ws1['A14'].alignment = align_center
    ws1.merge_cells('A14:F14')
    ws1['E14'].border = border_bottom

    # Столбец справа — ИЗ НАСТРОЕК
    ws1['H16'] = "Количество, кг"
    ws1['H16'].font = font_normal
    ws1['H16'].alignment = align_right
    ws1['I16'] = total_milk
    ws1['I16'].number_format = '0.00'
    ws1['I16'].border = border_bottom

    ws1['H17'] = "Сумма, руб. коп."
    ws1['H17'].font = font_normal
    ws1['H17'].alignment = align_right
    ws1['I17'] = total_sum
    ws1['I17'].number_format = '0.00'
    ws1['I17'].border = border_bottom

    ws1['H18'] = "Дебет счета"
    ws1['H18'].font = font_normal
    ws1['H18'].alignment = align_right
    ws1['I18'] = settings.get('debit_account', '')
    ws1['I18'].border = border_bottom

    ws1['H19'] = "Кредит счета"
    ws1['H19'].font = font_normal
    ws1['H19'].alignment = align_right
    ws1['I19'] = settings.get('credit_account', '')
    ws1['I19'].border = border_bottom

    widths1 = {'A': 15, 'B': 25, 'C': 15, 'D': 10, 'E': 12, 'F': 12, 'G': 5, 'H': 20, 'I': 15, 'J': 10}
    for col, w in widths1.items():
        ws1.column_dimensions[col].width = w

    # =========================================================================
    # ЛИСТ 2: стр2 (ДВЕ ОТДЕЛЬНЫЕ ТАБЛИЦЫ)
    # =========================================================================
    ws2 = wb.create_sheet("стр2")
    ws2.sheet_properties.pageSetUpPr.fitToPage = True
    ws2.page_setup.fitToWidth = 1
    ws2.page_setup.fitToHeight = 1
    ws2.page_setup.orientation = 'landscape'
    ws2.page_margins.left = 0.3
    ws2.page_margins.right = 0.3

    # Колонки: A(номер), B(пустой), C(ФИО), D(пустой между таблицами),
    # E(номер), F-G(обслуж.коров), H-K(надоено), L(жирн), M(жиро-ед), N(прочие), O(подпись),
    # P(Всего надоено), Q(Данные о качестве)
    widths2 = {
        'A': 5, 'B': 3, 'C': 30,  # Левая таблица
        'D': 3,  # Разделитель между таблицами
        'E': 5, 'F': 9, 'G': 9,  # Правая таблица: номер, обслуживалось (всего/из них)
        'H': 7, 'I': 7, 'J': 7, 'K': 7,  # Надоено (утро/полдень/вечер/всего)
        'L': 9, 'M': 9, 'N': 9, 'O': 11,  # жирн/жиро-ед/прочие/подпись
        'P': 11, 'Q': 11  # Всего надоено / Данные о качестве
    }
    for col, w in widths2.items():
        ws2.column_dimensions[col].width = w

    # Строка 1: две надписи на одной строке
    ws2['A1'] = "1-я страница журнала по форме № СП-21"
    ws2['A1'].font = font_header
    ws2['A1'].alignment = align_left
    ws2.merge_cells('A1:C1')

    ws2['E1'] = "18-я страница журнала по форме № СП-21"
    ws2['E1'].font = font_header
    ws2['E1'].alignment = align_right
    ws2.merge_cells('E1:Q1')

    # ===== ЛЕВАЯ ТАБЛИЦА (A-C, строки 2-4 заголовки) =====
    # Номер по порядку
    ws2['A2'] = "Номер\nпо\nпорядку"
    ws2['A2'].font = font_header
    ws2['A2'].alignment = align_center
    ws2.merge_cells('A2:A4')

    # ФИО
    ws2['C2'] = "Фамилия, имя, отчество\nдоярки (мастера машинной дойки)"
    ws2['C2'].font = font_header
    ws2['C2'].alignment = align_center
    ws2.merge_cells('C2:C4')

    # Рамки левой таблицы (строки 2-4, колонки A-C)
    apply_border(ws2, 2, 4, 1, 3)

    # ===== ПРАВАЯ ТАБЛИЦА (E-Q, строки 2-4 заголовки) =====
    # Номер по порядку
    ws2['E2'] = "Номер\nпо\nпорядку"
    ws2['E2'].font = font_header
    ws2['E2'].alignment = align_center
    ws2.merge_cells('E2:E4')

    # Обслуживалось коров (строка 3, объединяет F-G)
    ws2['F3'] = "обслужива-\nлось коров"
    ws2['F3'].font = font_header
    ws2['F3'].alignment = align_center
    ws2.merge_cells('F3:G3')

    # Надоено молока, кг (строка 3, объединяет H-K)
    ws2['H3'] = "надоено молока, кг"
    ws2['H3'].font = font_header
    ws2['H3'].alignment = align_center
    ws2.merge_cells('H3:K3')

    # Жирность, %
    ws2['L2'] = "жир-\nность,\n%"
    ws2['L2'].font = font_header
    ws2['L2'].alignment = align_center
    ws2.merge_cells('L2:L4')

    # Жиро-единиц
    ws2['M2'] = "жиро-\nеди-\nниц"
    ws2['M2'].font = font_header
    ws2['M2'].alignment = align_center
    ws2.merge_cells('M2:M4')

    # Прочие данные о качестве
    ws2['N2'] = "прочие\nданные\nо ка-\nчестве"
    ws2['N2'].font = font_header
    ws2['N2'].alignment = align_center
    ws2.merge_cells('N2:N4')

    # Подпись доярки
    ws2['O2'] = "подпись\nдоярки\n(мастера\nмашинной\nдойки)"
    ws2['O2'].font = font_header
    ws2['O2'].alignment = align_center
    ws2.merge_cells('O2:O4')

    # Всего надоено молока, кг
    ws2['P2'] = "Всего\nнадо-\nено\nмоло-\nка,\nкг"
    ws2['P2'].font = font_header
    ws2['P2'].alignment = align_center
    ws2.merge_cells('P2:P4')

    # Данные о качестве
    ws2['Q2'] = "Данные\nо\nкачестве"
    ws2['Q2'].font = font_header
    ws2['Q2'].alignment = align_center
    ws2.merge_cells('Q2:Q4')

    # Подзаголовки строки 4
    ws2['F4'] = "всего"
    ws2['F4'].font = font_header
    ws2['F4'].alignment = align_center

    ws2['G4'] = "из них\nдои-\nлось"
    ws2['G4'].font = font_header
    ws2['G4'].alignment = align_center

    ws2['H4'] = "ут-\nром"
    ws2['H4'].font = font_header
    ws2['H4'].alignment = align_center

    ws2['I4'] = "в\nпол-\nдень"
    ws2['I4'].font = font_header
    ws2['I4'].alignment = align_center

    ws2['J4'] = "ве-\nче-\nром"
    ws2['J4'].font = font_header
    ws2['J4'].alignment = align_center

    ws2['K4'] = "всего"
    ws2['K4'].font = font_header
    ws2['K4'].alignment = align_center

    # Рамки правой таблицы (строки 2-4, колонки E-Q)
    apply_border(ws2, 2, 4, 5, 17)

    # ===== ЗАПОЛНЕНИЕ ДАННЫМИ (18 строк) =====
    start_row = 5
    for i in range(1, 19):
        row = start_row + i - 1

        # Левая таблица: номер и ФИО
        ws2[f'A{row}'] = i
        ws2[f'A{row}'].alignment = align_center
        if i <= len(data):
            ws2[f'C{row}'] = data[i - 1]['worker'] or ""
        apply_border(ws2, row, row, 1, 3)

        # Правая таблица: номер и данные
        ws2[f'E{row}'] = i
        ws2[f'E{row}'].alignment = align_center
        if i <= len(data):
            r = data[i - 1]
            ws2[f'F{row}'] = r['cows'] or 0
            ws2[f'G{row}'] = r['cows'] or 0  # из них доилось = обслуживалось (упрощённо)
            ws2[f'H{row}'] = float(r['morning'] or 0)
            ws2[f'I{row}'] = float(r['daytime'] or 0)
            ws2[f'J{row}'] = float(r['evening'] or 0)
            ws2[f'K{row}'] = float(r['total'] or 0)
            ws2[f'L{row}'] = float(r['fat']) if r['fat'] is not None else ""
            ws2[f'M{row}'] = float(r['fat_units'] or 0)
        apply_border(ws2, row, row, 5, 17)

        # Форматирование чисел
        for col in 'FGHIKLM':
            ws2[f'{col}{row}'].alignment = align_center
            if col in 'FGHIKM':
                ws2[f'{col}{row}'].number_format = '0.00'

    # ===== ПОДВАЛ =====
    foot_row = start_row + 18  # строка после 18 данных

    # "Итого за день" под левой таблицей
    ws2[f'A{foot_row}'] = "Итого за день"
    ws2[f'A{foot_row}'].font = Font(name='Times New Roman', size=10, bold=True)
    ws2.merge_cells(f'A{foot_row}:C{foot_row}')
    apply_border(ws2, foot_row, foot_row, 1, 3)

    # Крестики в строке "Итого" под правой таблицей (под жирностью и подписью)
    ws2[f'L{foot_row}'] = "Х"
    ws2[f'L{foot_row}'].alignment = align_center
    ws2[f'O{foot_row}'] = "Х"
    ws2[f'O{foot_row}'].alignment = align_center
    apply_border(ws2, foot_row, foot_row, 5, 17)

    # Цена, Сумма, Дебет, Кредит — под столбцом "Подпись доярки" (O) и "Всего надоено" (P)
    price_row = foot_row + 2
    ws2[f'O{price_row}'] = "Цена, руб. коп."
    ws2[f'O{price_row}'].font = font_normal
    ws2[f'O{price_row}'].alignment = align_right
    ws2[f'P{price_row}'] = milk_price
    ws2[f'P{price_row}'].number_format = '0.00'

    ws2[f'O{price_row + 1}'] = "Сумма, руб. коп."
    ws2[f'O{price_row + 1}'].font = font_normal
    ws2[f'O{price_row + 1}'].alignment = align_right
    ws2[f'P{price_row + 1}'] = total_sum
    ws2[f'P{price_row + 1}'].number_format = '0.00'

    ws2[f'N{price_row + 2}'] = "Код синтетического\nи аналитического учета:"
    ws2[f'N{price_row + 2}'].font = font_normal
    ws2[f'N{price_row + 2}'].alignment = Alignment(horizontal='right', vertical='center', wrap_text=True)
    ws2.merge_cells(f'N{price_row + 2}:N{price_row + 3}')  # объединяем по вертикали

    ws2[f'O{price_row + 3}'] = "дебет"
    ws2[f'O{price_row + 3}'].font = font_normal
    ws2[f'O{price_row + 3}'].alignment = align_right
    ws2[f'P{price_row + 3}'] = settings.get('debit_account', '')

    ws2[f'O{price_row + 4}'] = "кредит"
    ws2[f'O{price_row + 4}'].font = font_normal
    ws2[f'O{price_row + 4}'].alignment = align_right
    ws2[f'P{price_row + 4}'] = settings.get('credit_account', '')

    # Рамки для ячеек значений (под P — Всего надоено)
    for r in range(price_row, price_row + 5):
        ws2[f'P{r}'].border = border_all

    # Зав. фермой и Бригадир — под левой таблицей
    sig_row = foot_row + 6
    ws2[f'A{sig_row}'] = "Зав. фермой"
    ws2[f'A{sig_row}'].font = Font(name='Times New Roman', size=10, bold=True)
    ws2[f'C{sig_row}'].border = border_bottom
    ws2[f'E{sig_row}'] = "(подпись)"
    ws2[f'E{sig_row}'].font = font_tiny
    ws2[f'G{sig_row}'].border = border_bottom
    ws2[f'I{sig_row}'] = "(расшифровка подписи)"
    ws2[f'I{sig_row}'].font = font_tiny

    ws2[f'A{sig_row + 2}'] = "Бригадир"
    ws2[f'A{sig_row + 2}'].font = Font(name='Times New Roman', size=10, bold=True)
    ws2[f'C{sig_row + 2}'].border = border_bottom
    ws2[f'E{sig_row + 2}'] = "(подпись)"
    ws2[f'E{sig_row + 2}'].font = font_tiny
    ws2[f'G{sig_row + 2}'].border = border_bottom
    ws2[f'I{sig_row + 2}'] = "(расшифровка подписи)"
    ws2[f'I{sig_row + 2}'].font = font_tiny

    # Дата
    date_row = sig_row + 4
    ws2[f'A{date_row}'] = '"'
    ws2[f'B{date_row}'] = "___"
    ws2[f'C{date_row}'] = '"'
    ws2[f'D{date_row}'] = "__________ 20__ г."

    # =========================================================================
    # ЛИСТ 3: стр3 (ШАБЛОН БЕЗ ДАТЫ, БЕЗ НАДПИСИ СВЕРХУ)
    # =========================================================================
    ws3 = wb.create_sheet("стр3")
    ws3.sheet_properties.pageSetUpPr.fitToPage = True
    ws3.page_setup.fitToWidth = 1
    ws3.page_setup.fitToHeight = 1
    ws3.page_setup.orientation = 'landscape'
    ws3.page_margins.left = 0.3
    ws3.page_margins.right = 0.3

    # Те же колонки, что и на стр2, но без ФИО (левая таблица только с номером)
    widths3 = {
        'A': 5, 'B': 3, 'C': 3,  # Левая таблица (только номер)
        'D': 3,  # Разделитель
        'E': 5, 'F': 9, 'G': 9,
        'H': 7, 'I': 7, 'J': 7, 'K': 7,
        'L': 9, 'M': 9, 'N': 9, 'O': 11,
        'P': 11, 'Q': 11
    }
    for col, w in widths3.items():
        ws3.column_dimensions[col].width = w

    # Надпись сверху убрана по требованию

    # ===== ЛЕВАЯ ТАБЛИЦА (только номер) =====
    ws3['A2'] = "Номер\nпо\nпорядку"
    ws3['A2'].font = font_header
    ws3['A2'].alignment = align_center
    ws3.merge_cells('A2:A4')
    apply_border(ws3, 2, 4, 1, 3)

    # ===== ПРАВАЯ ТАБЛИЦА (та же структура, что на стр2) =====
    ws3['E2'] = "Номер\nпо\nпорядку"
    ws3['E2'].font = font_header
    ws3['E2'].alignment = align_center
    ws3.merge_cells('E2:E4')

    ws3['F3'] = "обслужива-\nлось коров"
    ws3['F3'].font = font_header
    ws3['F3'].alignment = align_center
    ws3.merge_cells('F3:G3')

    ws3['H3'] = "надоено молока, кг"
    ws3['H3'].font = font_header
    ws3['H3'].alignment = align_center
    ws3.merge_cells('H3:K3')

    ws3['L2'] = "жир-\nность,\n%"
    ws3['L2'].font = font_header
    ws3['L2'].alignment = align_center
    ws3.merge_cells('L2:L4')

    ws3['M2'] = "жиро-\nеди-\nниц"
    ws3['M2'].font = font_header
    ws3['M2'].alignment = align_center
    ws3.merge_cells('M2:M4')

    ws3['N2'] = "прочие\nданные\nо ка-\nчестве"
    ws3['N2'].font = font_header
    ws3['N2'].alignment = align_center
    ws3.merge_cells('N2:N4')

    ws3['O2'] = "подпись\nдоярки\n(мастера\nмашинной\nдойки)"
    ws3['O2'].font = font_header
    ws3['O2'].alignment = align_center
    ws3.merge_cells('O2:O4')

    ws3['P2'] = "Всего\nнадо-\nено\nмоло-\nка,\nкг"
    ws3['P2'].font = font_header
    ws3['P2'].alignment = align_center
    ws3.merge_cells('P2:P4')

    ws3['Q2'] = "Данные\nо\nкачестве"
    ws3['Q2'].font = font_header
    ws3['Q2'].alignment = align_center
    ws3.merge_cells('Q2:Q4')

    ws3['F4'] = "всего"
    ws3['F4'].font = font_header
    ws3['F4'].alignment = align_center
    ws3['G4'] = "из них\nдои-\nлось"
    ws3['G4'].font = font_header
    ws3['G4'].alignment = align_center
    ws3['H4'] = "ут-\nром"
    ws3['H4'].font = font_header
    ws3['H4'].alignment = align_center
    ws3['I4'] = "в\nпол-\nдень"
    ws3['I4'].font = font_header
    ws3['I4'].alignment = align_center
    ws3['J4'] = "ве-\nче-\nром"
    ws3['J4'].font = font_header
    ws3['J4'].alignment = align_center
    ws3['K4'] = "всего"
    ws3['K4'].font = font_header
    ws3['K4'].alignment = align_center

    apply_border(ws3, 2, 4, 5, 17)

    # Пустая сетка на 18 строк
    start_row3 = 5
    for i in range(1, 19):
        row = start_row3 + i - 1
        ws3[f'A{row}'] = i
        ws3[f'A{row}'].alignment = align_center
        ws3[f'E{row}'] = i
        ws3[f'E{row}'].alignment = align_center
        apply_border(ws3, row, row, 1, 3)
        apply_border(ws3, row, row, 5, 17)
        for col in 'FGHIKLM':
            ws3[f'{col}{row}'].alignment = align_center
            if col in 'FGHIKM':
                ws3[f'{col}{row}'].number_format = '0.00'

    # Подвал стр3 — аналогично стр2
    foot_row3 = start_row3 + 18

    ws3[f'A{foot_row3}'] = "Итого за день"
    ws3[f'A{foot_row3}'].font = Font(name='Times New Roman', size=10, bold=True)
    ws3.merge_cells(f'A{foot_row3}:C{foot_row3}')
    apply_border(ws3, foot_row3, foot_row3, 1, 3)

    ws3[f'L{foot_row3}'] = "Х"
    ws3[f'L{foot_row3}'].alignment = align_center
    ws3[f'O{foot_row3}'] = "Х"
    ws3[f'O{foot_row3}'].alignment = align_center
    apply_border(ws3, foot_row3, foot_row3, 5, 17)

    price_row3 = foot_row3 + 2
    ws3[f'O{price_row3}'] = "Цена, руб. коп."
    ws3[f'O{price_row3}'].font = font_normal
    ws3[f'O{price_row3}'].alignment = align_right
    ws3[f'P{price_row3}'] = milk_price
    ws3[f'P{price_row3}'].number_format = '0.00'

    ws3[f'O{price_row3 + 1}'] = "Сумма, руб. коп."
    ws3[f'O{price_row3 + 1}'].font = font_normal
    ws3[f'O{price_row3 + 1}'].alignment = align_right
    ws3[f'P{price_row3 + 1}'] = total_sum
    ws3[f'P{price_row3 + 1}'].number_format = '0.00'

    ws3[f'N{price_row3 + 2}'] = "Код синтетического\nи аналитического учета:"
    ws3[f'N{price_row3 + 2}'].font = font_normal
    ws3[f'N{price_row3 + 2}'].alignment = align_right
    ws3.merge_cells(f'N{price_row3 + 2}:O{price_row3 + 2}')

    ws3[f'O{price_row3 + 3}'] = "дебет"
    ws3[f'O{price_row3 + 3}'].font = font_normal
    ws3[f'O{price_row3 + 3}'].alignment = align_right
    ws3[f'P{price_row3 + 3}'] = settings.get('debit_account', '')

    ws3[f'O{price_row3 + 4}'] = "кредит"
    ws3[f'O{price_row3 + 4}'].font = font_normal
    ws3[f'O{price_row3 + 4}'].alignment = align_right
    ws3[f'P{price_row3 + 4}'] = settings.get('credit_account', '')

    for r in range(price_row3, price_row3 + 5):
        ws3[f'P{r}'].border = border_all

    sig_row3 = foot_row3 + 6
    ws3[f'A{sig_row3}'] = "Зав. фермой"
    ws3[f'A{sig_row3}'].font = Font(name='Times New Roman', size=10, bold=True)
    ws3[f'C{sig_row3}'].border = border_bottom
    ws3[f'E{sig_row3}'] = "(подпись)"
    ws3[f'E{sig_row3}'].font = font_tiny
    ws3[f'G{sig_row3}'].border = border_bottom
    ws3[f'I{sig_row3}'] = "(расшифровка подписи)"
    ws3[f'I{sig_row3}'].font = font_tiny

    ws3[f'A{sig_row3 + 2}'] = "Бригадир"
    ws3[f'A{sig_row3 + 2}'].font = Font(name='Times New Roman', size=10, bold=True)
    ws3[f'C{sig_row3 + 2}'].border = border_bottom
    ws3[f'E{sig_row3 + 2}'] = "(подпись)"
    ws3[f'E{sig_row3 + 2}'].font = font_tiny
    ws3[f'G{sig_row3 + 2}'].border = border_bottom
    ws3[f'I{sig_row3 + 2}'] = "(расшифровка подписи)"
    ws3[f'I{sig_row3 + 2}'].font = font_tiny

    date_row3 = sig_row3 + 4
    ws3[f'A{date_row3}'] = '"'
    ws3[f'B{date_row3}'] = "___"
    ws3[f'C{date_row3}'] = '"'
    ws3[f'D{date_row3}'] = "__________ 20__ г."

    ws3[f'A{date_row3 + 2}'] = "По данному образцу печатать со 2-ой по 17-ю страницы журнала по форме № СП-21"
    ws3[f'A{date_row3 + 2}'].font = font_small
    ws3.merge_cells(f'A{date_row3 + 2}:Q{date_row3 + 2}')

    # =========================================================================
    # СОХРАНЕНИЕ
    # =========================================================================
    out = BytesIO()
    wb.save(out)
    out.seek(0)

    return send_file(
        out,
        as_attachment=True,
        download_name=f"SP-21_{day}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )