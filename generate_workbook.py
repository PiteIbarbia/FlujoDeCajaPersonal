import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference, PieChart
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import date
import os

def build_multi_year_cash_flow():
    target_dir = r"C:\Users\Pite\Documents\CashFlowHogar\Flujo de caja Personal"
    output_path = os.path.join(target_dir, "Flujo_de_Caja_Personal.xlsx")

    # 1. READ EXISTING USER DATA TO PRESERVE ALL USER TRANSACTIONS
    existing_records = []
    if os.path.exists(output_path):
        try:
            wb_old = openpyxl.load_workbook(output_path, data_only=False)
            if "Registro Diario" in wb_old.sheetnames:
                ws_old = wb_old["Registro Diario"]
                for r in range(5, 500):
                    date_val = ws_old.cell(r, 2).value
                    if date_val:
                        # Extract columns B to I
                        rec = [
                            ws_old.cell(r, 2).value, # Fecha
                            ws_old.cell(r, 4).value, # Tipo
                            ws_old.cell(r, 5).value, # Categoria
                            ws_old.cell(r, 6).value, # Concepto
                            ws_old.cell(r, 7).value, # Medio de Pago
                            ws_old.cell(r, 8).value, # Monto
                            ws_old.cell(r, 9).value  # Notas
                        ]
                        existing_records.append(rec)
            wb_old.close()
            print(f"Preserved {len(existing_records)} existing user records.")
        except Exception as e:
            print("Notice loading old workbook:", e)

    # Fallback to defaults if empty
    if not existing_records:
        existing_records = [
            [date(2026, 10, 23), "Gasto Hormiga", "Donaciones, Rifas y Colaboraciones", "Rifa CERMUN", "Transferencia / Billetera Virtual (Mercado Pago, etc.)", 0, "Colaboración a Lolo Urbina"],
            [date(2026, 10, 27), "Gasto Hormiga", "Café al paso y Desayunos fuera", "Starbucks", "Transferencia / Billetera Virtual (Mercado Pago, etc.)", 17900, "Café en Starbucks"],
            [date(2026, 10, 30), "Ingreso", "Sueldo / Salario", "Guillermo Mattioli", "Efectivo", 600000, "Sueldo mensual"],
            [date(2026, 10, 30), "Gasto Hormiga", "Café al paso y Desayunos fuera", "Starbucks", "Transferencia / Billetera Virtual (Mercado Pago, etc.)", 10400, "Café en Starbucks"],
            [date(2026, 10, 30), "Gasto Variable", "Danza, Ballet y Actividades Deportivas", "Ballet", "Efectivo", 210000, "Deuda Sandra Valentini"],
            [date(2026, 10, 30), "Gasto Variable", "Danza, Ballet y Actividades Deportivas", "Ballet", "Efectivo", 20000, "Deuda clase particular Olga"],
            [date(2026, 10, 30), "Gasto Fijo", "Suscripciones Digitales (Streaming, Apps)", "Meli +", "Transferencia / Billetera Virtual (Mercado Pago, etc.)", 20990, "Suscripción Meli +"]
        ]

    # Create new workbook
    wb = openpyxl.Workbook()
    default_sheet = wb.active

    # -------------------------------------------------------------
    # COLOR PALETTE & DESIGN SYSTEM
    # -------------------------------------------------------------
    NAVY_DARK = "1B365D"       # Main Headers
    NAVY_MID = "2C4C7C"        # Sub-headers
    BLUE_PALE = "EBF3FA"       # Subheader light fill
    BLUE_CARD = "F0F4F8"       # Card bg
    GREEN_SUCCESS = "27AE60"   # Income / Positive Flow
    GREEN_PALE = "EAFAF1"      # Income background
    RED_ACCENT = "C0392B"      # Expense / Alert
    RED_PALE = "FDEDEC"        # Expense background
    ORANGE_ALERT = "D35400"    # Hormiga / Warning
    ORANGE_PALE = "FEF5E7"     # Warning background
    GRAY_TEXT = "5D6D7E"       # Muted text
    GRAY_LIGHT = "F8F9FA"      # Alternate rows
    GRAY_BORDER = "D5D8DC"     # Subtle borders
    WHITE = "FFFFFF"

    # Fonts
    font_main_title = Font(name="Segoe UI", size=15, bold=True, color=NAVY_DARK)
    font_subtitle = Font(name="Segoe UI", size=10, italic=True, color=GRAY_TEXT)
    font_table_header = Font(name="Segoe UI", size=10, bold=True, color=WHITE)
    font_section_header = Font(name="Segoe UI", size=10, bold=True, color=NAVY_DARK)
    font_body = Font(name="Segoe UI", size=10, color="2C3E50")
    font_body_bold = Font(name="Segoe UI", size=10, bold=True, color="2C3E50")
    font_kpi_num = Font(name="Segoe UI", size=14, bold=True, color=NAVY_DARK)
    font_kpi_label = Font(name="Segoe UI", size=8, bold=True, color=GRAY_TEXT)

    # Fills
    fill_header = PatternFill(start_color=NAVY_DARK, end_color=NAVY_DARK, fill_type="solid")
    fill_sub_blue = PatternFill(start_color=BLUE_PALE, end_color=BLUE_PALE, fill_type="solid")
    fill_inc_card = PatternFill(start_color=GREEN_PALE, end_color=GREEN_PALE, fill_type="solid")
    fill_exp_card = PatternFill(start_color=RED_PALE, end_color=RED_PALE, fill_type="solid")
    fill_hor_card = PatternFill(start_color=ORANGE_PALE, end_color=ORANGE_PALE, fill_type="solid")
    fill_total = PatternFill(start_color="D6E4F0", end_color="D6E4F0", fill_type="solid")
    fill_zebra = PatternFill(start_color="FDFEFE", end_color="FDFEFE", fill_type="solid")

    # Borders
    thin_border = Side(style="thin", color=GRAY_BORDER)
    double_bottom = Side(style="double", color=NAVY_DARK)
    cell_border = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    total_border = Border(top=thin_border, bottom=double_bottom, left=thin_border, right=thin_border)

    # Alignments
    align_center = Alignment(horizontal="center", vertical="center")
    align_left = Alignment(horizontal="left", vertical="center")
    align_right = Alignment(horizontal="right", vertical="center")

    CURRENCY_FORMAT = "$ #,##0.00"
    PERCENT_FORMAT = "0.0%"
    DATE_FORMAT = "DD/MM/YYYY"

    # =============================================================
    # 1. SHEET: CONFIGURACION (Catalogs & Multi-Month Calendar)
    # =============================================================
    ws_config = wb.create_sheet(title="Configuracion")
    ws_config.views.sheetView[0].showGridLines = True

    ws_config["A1"] = "CONFIGURACIÓN FINANCIERA, CATEGORÍAS Y CALENDARIO MULTI-ANUAL"
    ws_config["A1"].font = font_main_title
    ws_config["A2"] = "Define los presupuestos base, medios de pago y el calendario de meses para 2026 y 2027."
    ws_config["A2"].font = font_subtitle

    # Initial bank balance anchor (before October 2026)
    ws_config["J2"] = "Saldo Inicial de Caja al 01/10/2026:"
    ws_config["J2"].font = font_body_bold
    ws_config["J2"].alignment = align_right
    ws_config["K2"] = 0.00
    ws_config["K2"].font = font_body_bold
    ws_config["K2"].number_format = CURRENCY_FORMAT
    ws_config["K2"].fill = fill_sub_blue
    ws_config["K2"].border = cell_border

    config_headers = [
        ("A4", "Tipo de Movimiento"),
        ("B4", "Categoría"),
        ("C4", "Grupo de Gasto"),
        ("D4", "Presupuesto Mensual Base"),
        ("F4", "Medio de Pago"),
        ("H4", "Período / Mes"),
        ("I4", "Fecha Inicio"),
        ("J4", "Fecha Fin"),
        ("K4", "Año"),
        ("L4", "Etiqueta Corta")
    ]
    for cell_ref, text in config_headers:
        ws_config[cell_ref] = text
        ws_config[cell_ref].font = font_table_header
        ws_config[cell_ref].fill = fill_header
        ws_config[cell_ref].alignment = align_center
        ws_config[cell_ref].border = cell_border

    tipos_movimiento = [
        "Ingreso",
        "Gasto Fijo",
        "Gasto Variable",
        "Gasto Hormiga",
        "Ahorro / Inversión"
    ]
    for i, tipo in enumerate(tipos_movimiento, start=5):
        ws_config[f"A{i}"] = tipo
        ws_config[f"A{i}"].font = font_body
        ws_config[f"A{i}"].border = cell_border
        ws_config[f"A{i}"].alignment = align_left

    categorias_data = [
        # INGRESOS
        ("Sueldo / Salario", "Ingreso", 600000),
        ("Honorarios / Freelance", "Ingreso", 0),
        ("Rendimientos / Inversiones", "Ingreso", 0),
        ("Otros Ingresos", "Ingreso", 0),
        # GASTOS FIJOS
        ("Vivienda (Alquiler / Hipoteca)", "Gasto Fijo", 180000),
        ("Expensas y Mantenimiento Edificio", "Gasto Fijo", 35000),
        ("Servicios Básicos (Luz, Gas, Agua)", "Gasto Fijo", 25000),
        ("Internet, Telefonía y Cable", "Gasto Fijo", 22000),
        ("Seguros (Salud / Auto / Hogar)", "Gasto Fijo", 30000),
        ("Educación / Cursos / Cuotas", "Gasto Fijo", 20000),
        ("Suscripciones Digitales (Streaming, Apps)", "Gasto Fijo", 21000),
        ("Pago de Deudas y Préstamos", "Gasto Fijo", 0),
        # GASTOS VARIABLES
        ("Supermercado y Alimentos del Hogar", "Gasto Variable", 120000),
        ("Danza, Ballet y Actividades Deportivas", "Gasto Variable", 240000),
        ("Transporte y Combustible", "Gasto Variable", 40000),
        ("Farmacia y Salud", "Gasto Variable", 15000),
        ("Cuidado Personal y Limpieza", "Gasto Variable", 15000),
        ("Ropa y Calzado", "Gasto Variable", 25000),
        ("Salidas, Bares y Restaurantes", "Gasto Variable", 35000),
        ("Delivery y Comida Rápida", "Gasto Variable", 20000),
        # GASTOS HORMIGA
        ("Café al paso y Desayunos fuera", "Gasto Hormiga", 30000),
        ("Kiosco, Snacks y Golosinas", "Gasto Hormiga", 8000),
        ("Propinas y Pequeños Gastos en Efectivo", "Gasto Hormiga", 6000),
        ("Donaciones, Rifas y Colaboraciones", "Gasto Hormiga", 5000),
        ("Compras Impulsivas / Antojos", "Gasto Hormiga", 10000),
        # AHORRO E INVERSION
        ("Fondo de Emergencia", "Ahorro / Inversión", 40000),
        ("Inversiones y Metas de Ahorro", "Ahorro / Inversión", 50000)
    ]

    for idx, (cat, grp, pres) in enumerate(categorias_data, start=5):
        ws_config[f"B{idx}"] = cat
        ws_config[f"C{idx}"] = grp
        ws_config[f"D{idx}"] = pres

        ws_config[f"B{idx}"].font = font_body
        ws_config[f"C{idx}"].font = font_body
        ws_config[f"D{idx}"].font = font_body

        ws_config[f"B{idx}"].border = cell_border
        ws_config[f"C{idx}"].border = cell_border
        ws_config[f"D{idx}"].border = cell_border

        ws_config[f"D{idx}"].number_format = CURRENCY_FORMAT
        ws_config[f"D{idx}"].alignment = align_right

    medios_pago = [
        "Efectivo",
        "Tarjeta de Débito",
        "Tarjeta de Crédito",
        "Transferencia / Billetera Virtual (Mercado Pago, etc.)",
        "Cuenta Bancaria Directa"
    ]
    for idx, mp in enumerate(medios_pago, start=5):
        ws_config[f"F{idx}"] = mp
        ws_config[f"F{idx}"].font = font_body
        ws_config[f"F{idx}"].border = cell_border
        ws_config[f"F{idx}"].alignment = align_left

    last_cat_row = 4 + len(categorias_data)
    ws_config[f"C{last_cat_row + 1}"] = "Total Presupuestado (Gastos + Ahorro):"
    ws_config[f"C{last_cat_row + 1}"].font = font_body_bold
    ws_config[f"C{last_cat_row + 1}"].alignment = align_right
    ws_config[f"D{last_cat_row + 1}"] = f'=SUMIFS(D5:D{last_cat_row}, C5:C{last_cat_row}, "<>Ingreso")'
    ws_config[f"D{last_cat_row + 1}"].font = font_body_bold
    ws_config[f"D{last_cat_row + 1}"].fill = fill_total
    ws_config[f"D{last_cat_row + 1}"].border = total_border
    ws_config[f"D{last_cat_row + 1}"].number_format = CURRENCY_FORMAT

    # CALENDARIO DE MESES: Q4 2026 + TODO 2027 (15 meses)
    calendario_meses = [
        ("Octubre 2026", date(2026, 10, 1), date(2026, 10, 31), 2026, "Oct-26"),
        ("Noviembre 2026", date(2026, 11, 1), date(2026, 11, 30), 2026, "Nov-26"),
        ("Diciembre 2026", date(2026, 12, 1), date(2026, 12, 31), 2026, "Dic-26"),
        ("Enero 2027", date(2027, 1, 1), date(2027, 1, 31), 2027, "Ene-27"),
        ("Febrero 2027", date(2027, 2, 1), date(2027, 2, 28), 2027, "Feb-27"),
        ("Marzo 2027", date(2027, 3, 1), date(2027, 3, 31), 2027, "Mar-27"),
        ("Abril 2027", date(2027, 4, 1), date(2027, 4, 30), 2027, "Abr-27"),
        ("Mayo 2027", date(2027, 5, 1), date(2027, 5, 31), 2027, "May-27"),
        ("Junio 2027", date(2027, 6, 1), date(2027, 6, 30), 2027, "Jun-27"),
        ("Julio 2027", date(2027, 7, 1), date(2027, 7, 31), 2027, "Jul-27"),
        ("Agosto 2027", date(2027, 8, 1), date(2027, 8, 31), 2027, "Ago-27"),
        ("Septiembre 2027", date(2027, 9, 1), date(2027, 9, 30), 2027, "Sep-27"),
        ("Octubre 2027", date(2027, 10, 1), date(2027, 10, 31), 2027, "Oct-27"),
        ("Noviembre 2027", date(2027, 11, 1), date(2027, 11, 30), 2027, "Nov-27"),
        ("Diciembre 2027", date(2027, 12, 1), date(2027, 12, 31), 2027, "Dic-27")
    ]

    for m_idx, (m_name, d_ini, d_fin, anio, short_lbl) in enumerate(calendario_meses, start=5):
        ws_config[f"H{m_idx}"] = m_name
        ws_config[f"I{m_idx}"] = d_ini
        ws_config[f"J{m_idx}"] = d_fin
        ws_config[f"K{m_idx}"] = anio
        ws_config[f"L{m_idx}"] = short_lbl

        ws_config[f"H{m_idx}"].font = font_body
        ws_config[f"I{m_idx}"].font = font_body
        ws_config[f"J{m_idx}"].font = font_body
        ws_config[f"K{m_idx}"].font = font_body
        ws_config[f"L{m_idx}"].font = font_body

        ws_config[f"H{m_idx}"].border = cell_border
        ws_config[f"I{m_idx}"].border = cell_border
        ws_config[f"J{m_idx}"].border = cell_border
        ws_config[f"K{m_idx}"].border = cell_border
        ws_config[f"L{m_idx}"].border = cell_border

        ws_config[f"I{m_idx}"].number_format = DATE_FORMAT
        ws_config[f"J{m_idx}"].number_format = DATE_FORMAT
        ws_config[f"I{m_idx}"].alignment = align_center
        ws_config[f"J{m_idx}"].alignment = align_center
        ws_config[f"K{m_idx}"].alignment = align_center
        ws_config[f"L{m_idx}"].alignment = align_center

    ws_config.column_dimensions["A"].width = 24
    ws_config.column_dimensions["B"].width = 42
    ws_config.column_dimensions["C"].width = 24
    ws_config.column_dimensions["D"].width = 30
    ws_config.column_dimensions["E"].width = 4
    ws_config.column_dimensions["F"].width = 45
    ws_config.column_dimensions["G"].width = 4
    ws_config.column_dimensions["H"].width = 22
    ws_config.column_dimensions["I"].width = 16
    ws_config.column_dimensions["J"].width = 16
    ws_config.column_dimensions["K"].width = 12
    ws_config.column_dimensions["L"].width = 16

    # =============================================================
    # 2. SHEET: REGISTRO DIARIO (1500 continuous rows for 2026-2027)
    # =============================================================
    ws_daily = wb.create_sheet(title="Registro Diario")
    ws_daily.views.sheetView[0].showGridLines = True

    ws_daily["A1"] = "REGISTRO DIARIO DE INGRESOS Y GASTOS (2026 - 2027)"
    ws_daily["A1"].font = font_main_title
    ws_daily["A2"] = "Anota aquí tus movimientos cotidianos. Las fórmulas calculan automáticamente el número de registro (#) y la semana del mes."
    ws_daily["A2"].font = font_subtitle

    daily_headers = [
        ("A4", "#"),
        ("B4", "Fecha"),
        ("C4", "Semana"),
        ("D4", "Tipo de Movimiento"),
        ("E4", "Categoría"),
        ("F4", "Concepto / Detalle Específico"),
        ("G4", "Medio de Pago"),
        ("H4", "Monto ($)"),
        ("I4", "Notas / Justificación")
    ]
    for cell_ref, text in daily_headers:
        ws_daily[cell_ref] = text
        ws_daily[cell_ref].font = font_table_header
        ws_daily[cell_ref].fill = fill_header
        ws_daily[cell_ref].alignment = align_center
        ws_daily[cell_ref].border = cell_border

    # Insert preserved user records
    for i, rec in enumerate(existing_records, start=5):
        fecha, tipo, cat, detalle, mp, monto, notas = rec
        ws_daily[f"A{i}"] = f'=IF(B{i}="","",ROW()-4)'
        ws_daily[f"B{i}"] = fecha
        ws_daily[f"C{i}"] = f'=IF(B{i}="","", "Semana " & MIN(5, INT((DAY(B{i})-1)/7)+1))'
        ws_daily[f"D{i}"] = tipo
        ws_daily[f"E{i}"] = cat
        ws_daily[f"F{i}"] = detalle
        ws_daily[f"G{i}"] = mp
        ws_daily[f"H{i}"] = monto
        ws_daily[f"I{i}"] = notas

        ws_daily[f"A{i}"].alignment = align_center
        ws_daily[f"B{i}"].alignment = align_center
        ws_daily[f"B{i}"].number_format = DATE_FORMAT
        ws_daily[f"C{i}"].alignment = align_center
        ws_daily[f"D{i}"].alignment = align_left
        ws_daily[f"E{i}"].alignment = align_left
        ws_daily[f"F{i}"].alignment = align_left
        ws_daily[f"G{i}"].alignment = align_left
        ws_daily[f"H{i}"].alignment = align_right
        ws_daily[f"H{i}"].number_format = CURRENCY_FORMAT
        ws_daily[f"I{i}"].alignment = align_left

        for col in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
            ws_daily[f"{col}{i}"].font = font_body
            ws_daily[f"{col}{i}"].border = cell_border

        if tipo == "Gasto Hormiga":
            ws_daily[f"D{i}"].fill = fill_hor_card
            ws_daily[f"H{i}"].fill = fill_hor_card
        elif tipo == "Ingreso":
            ws_daily[f"D{i}"].fill = fill_inc_card
            ws_daily[f"H{i}"].fill = fill_inc_card

    # Expand empty template rows up to row 1500!
    start_empty = 5 + len(existing_records)
    end_empty = 1505
    for row in range(start_empty, end_empty):
        ws_daily[f"A{row}"] = f'=IF(B{row}="","",ROW()-4)'
        ws_daily[f"A{row}"].alignment = align_center
        ws_daily[f"C{row}"] = f'=IF(B{row}="","", "Semana " & MIN(5, INT((DAY(B{row})-1)/7)+1))'
        ws_daily[f"C{row}"].alignment = align_center
        ws_daily[f"B{row}"].number_format = DATE_FORMAT
        ws_daily[f"B{row}"].alignment = align_center
        ws_daily[f"H{row}"].number_format = CURRENCY_FORMAT
        ws_daily[f"H{row}"].alignment = align_right

        for col in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
            ws_daily[f"{col}{row}"].font = font_body
            ws_daily[f"{col}{row}"].border = cell_border

    # Dropdowns for daily register
    dv_tipo = DataValidation(type="list", formula1="=Configuracion!$A$5:$A$9", allow_blank=True)
    ws_daily.add_data_validation(dv_tipo)
    dv_tipo.add(f"D5:D{end_empty}")

    dv_cat = DataValidation(type="list", formula1=f"=Configuracion!$B$5:$B${last_cat_row}", allow_blank=True)
    ws_daily.add_data_validation(dv_cat)
    dv_cat.add(f"E5:E{end_empty}")

    dv_mp = DataValidation(type="list", formula1="=Configuracion!$F$5:$F$9", allow_blank=True)
    ws_daily.add_data_validation(dv_mp)
    dv_mp.add(f"G5:G{end_empty}")

    ws_daily.column_dimensions["A"].width = 6
    ws_daily.column_dimensions["B"].width = 14
    ws_daily.column_dimensions["C"].width = 13
    ws_daily.column_dimensions["D"].width = 22
    ws_daily.column_dimensions["E"].width = 40
    ws_daily.column_dimensions["F"].width = 38
    ws_daily.column_dimensions["G"].width = 32
    ws_daily.column_dimensions["H"].width = 18
    ws_daily.column_dimensions["I"].width = 32

    # =============================================================
    # 3. SHEET: TABLERO MENSUAL (Dynamic Month Selector Oct 26 - Dec 27)
    # =============================================================
    ws_dash = wb.create_sheet(title="Tablero Mensual")
    ws_dash.views.sheetView[0].showGridLines = True

    ws_dash["B2"] = "CONTROL FINANCIERO MENSUAL (OCTUBRE 2026 - DICIEMBRE 2027)"
    ws_dash["B2"].font = font_main_title
    ws_dash["B3"] = "Elige cualquier mes en el selector: todo el tablero, semanas, gráficos y categorías se actualizan al instante."
    ws_dash["B3"].font = font_subtitle

    # Selector de Mes en F5
    ws_dash["E5"] = "Selecciona el Mes / Período:"
    ws_dash["E5"].font = font_body_bold
    ws_dash["E5"].alignment = align_right
    ws_dash["F5"] = "Noviembre 2026"  # Current/next working month
    ws_dash["F5"].font = Font(name="Segoe UI", size=11, bold=True, color=NAVY_DARK)
    ws_dash["F5"].alignment = align_center
    ws_dash["F5"].fill = fill_sub_blue
    ws_dash["F5"].border = cell_border

    # Data validation dropdown for F5
    dv_mes_dash = DataValidation(type="list", formula1="=Configuracion!$H$5:$H$19", allow_blank=False)
    ws_dash.add_data_validation(dv_mes_dash)
    dv_mes_dash.add("F5")

    # Helper cells for StartDate and EndDate (in P4 and P5)
    ws_dash["P4"] = "=VLOOKUP(F5, Configuracion!$H$5:$J$19, 2, FALSE)"
    ws_dash["P5"] = "=VLOOKUP(F5, Configuracion!$H$5:$J$19, 3, FALSE)"
    ws_dash["P4"].number_format = DATE_FORMAT
    ws_dash["P5"].number_format = DATE_FORMAT
    ws_dash["O4"] = "Fecha Inicio:"
    ws_dash["O5"] = "Fecha Fin:"
    ws_dash["O4"].font = font_subtitle
    ws_dash["O5"].font = font_subtitle
    ws_dash["P4"].font = font_subtitle
    ws_dash["P5"].font = font_subtitle

    # Saldo Inicial Dinámico (C5): Suma todos los flujos anteriores a la fecha de inicio del mes seleccionado
    ws_dash["B5"] = "Saldo Inicial Disponible del Mes:"
    ws_dash["B5"].font = font_body_bold
    ws_dash["B5"].alignment = align_left
    ws_dash["C5"] = "=Configuracion!$K$2 + SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Ingreso\", 'Registro Diario'!$B$5:$B$1500, \"<\"&$P$4) - (SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Gasto Fijo\", 'Registro Diario'!$B$5:$B$1500, \"<\"&$P$4) + SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Gasto Variable\", 'Registro Diario'!$B$5:$B$1500, \"<\"&$P$4) + SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Gasto Hormiga\", 'Registro Diario'!$B$5:$B$1500, \"<\"&$P$4) + SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Ahorro / Inversión\", 'Registro Diario'!$B$5:$B$1500, \"<\"&$P$4))"
    ws_dash["C5"].font = font_kpi_num
    ws_dash["C5"].number_format = CURRENCY_FORMAT
    ws_dash["C5"].alignment = align_right
    ws_dash["C5"].fill = fill_sub_blue
    ws_dash["C5"].border = cell_border

    # KPI 1: Ingresos Totales del Mes (B7:C8)
    ws_dash.merge_cells("B7:C7")
    ws_dash["B7"] = "TOTAL INGRESOS DEL MES"
    ws_dash["B7"].font = font_kpi_label
    ws_dash["B7"].alignment = align_center
    ws_dash["B7"].fill = fill_inc_card

    ws_dash.merge_cells("B8:C8")
    ws_dash["B8"] = "=SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Ingreso\", 'Registro Diario'!$B$5:$B$1500, \">=\"&$P$4, 'Registro Diario'!$B$5:$B$1500, \"<=\"&$P$5)"
    ws_dash["B8"].font = Font(name="Segoe UI", size=13, bold=True, color=GREEN_SUCCESS)
    ws_dash["B8"].alignment = align_center
    ws_dash["B8"].number_format = CURRENCY_FORMAT
    ws_dash["B8"].fill = fill_inc_card
    for r in [7, 8]:
        for c in ["B", "C"]:
            ws_dash[f"{c}{r}"].border = cell_border

    # KPI 2: Gastos Operativos del Mes (D7:E8)
    ws_dash.merge_cells("D7:E7")
    ws_dash["D7"] = "GASTOS TOTALES (FIJOS+VAR+HORMIGA)"
    ws_dash["D7"].font = font_kpi_label
    ws_dash["D7"].alignment = align_center
    ws_dash["D7"].fill = fill_exp_card

    ws_dash.merge_cells("D8:E8")
    ws_dash["D8"] = "=SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Gasto Fijo\", 'Registro Diario'!$B$5:$B$1500, \">=\"&$P$4, 'Registro Diario'!$B$5:$B$1500, \"<=\"&$P$5) + SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Gasto Variable\", 'Registro Diario'!$B$5:$B$1500, \">=\"&$P$4, 'Registro Diario'!$B$5:$B$1500, \"<=\"&$P$5) + SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Gasto Hormiga\", 'Registro Diario'!$B$5:$B$1500, \">=\"&$P$4, 'Registro Diario'!$B$5:$B$1500, \"<=\"&$P$5)"
    ws_dash["D8"].font = Font(name="Segoe UI", size=13, bold=True, color=RED_ACCENT)
    ws_dash["D8"].alignment = align_center
    ws_dash["D8"].number_format = CURRENCY_FORMAT
    ws_dash["D8"].fill = fill_exp_card
    for r in [7, 8]:
        for c in ["D", "E"]:
            ws_dash[f"{c}{r}"].border = cell_border

    # KPI 3: Gastos Hormiga Fugas (F7:F8)
    ws_dash["F7"] = "GASTOS HORMIGA (FUGAS)"
    ws_dash["F7"].font = font_kpi_label
    ws_dash["F7"].alignment = align_center
    ws_dash["F7"].fill = fill_hor_card

    ws_dash["F8"] = "=SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Gasto Hormiga\", 'Registro Diario'!$B$5:$B$1500, \">=\"&$P$4, 'Registro Diario'!$B$5:$B$1500, \"<=\"&$P$5)"
    ws_dash["F8"].font = Font(name="Segoe UI", size=13, bold=True, color=ORANGE_ALERT)
    ws_dash["F8"].alignment = align_center
    ws_dash["F8"].number_format = CURRENCY_FORMAT
    ws_dash["F8"].fill = fill_hor_card
    ws_dash["F7"].border = cell_border
    ws_dash["F8"].border = cell_border

    # KPI 4: Ahorro e Inversión (G7:G8)
    ws_dash["G7"] = "AHORRO / INVERSIÓN"
    ws_dash["G7"].font = font_kpi_label
    ws_dash["G7"].alignment = align_center
    ws_dash["G7"].fill = fill_sub_blue

    ws_dash["G8"] = "=SUMIFS('Registro Diario'!$H$5:$H$1500, 'Registro Diario'!$D$5:$D$1500, \"Ahorro / Inversión\", 'Registro Diario'!$B$5:$B$1500, \">=\"&$P$4, 'Registro Diario'!$B$5:$B$1500, \"<=\"&$P$5)"
    ws_dash["G8"].font = Font(name="Segoe UI", size=13, bold=True, color=NAVY_DARK)
    ws_dash["G8"].alignment = align_center
    ws_dash["G8"].number_format = CURRENCY_FORMAT
    ws_dash["G8"].fill = fill_sub_blue
    ws_dash["G7"].border = cell_border
    ws_dash["G8"].border = cell_border

    # KPI 5: Flujo Neto del Mes (H7:H8)
    ws_dash["H7"] = "FLUJO NETO DEL MES"
    ws_dash["H7"].font = font_kpi_label
    ws_dash["H7"].alignment = align_center
    ws_dash["H7"].fill = fill_total

    ws_dash["H8"] = "=B8-D8-G8"
    ws_dash["H8"].font = Font(name="Segoe UI", size=13, bold=True, color=NAVY_DARK)
    ws_dash["H8"].alignment = align_center
    ws_dash["H8"].number_format = CURRENCY_FORMAT
    ws_dash["H8"].fill = fill_total
    ws_dash["H7"].border = cell_border
    ws_dash["H8"].border = cell_border

    # KPI 6: Saldo Final Disponible (I7:I8)
    ws_dash["I7"] = "SALDO FINAL DISPONIBLE"
    ws_dash["I7"].font = font_kpi_label
    ws_dash["I7"].alignment = align_center
    ws_dash["I7"].fill = fill_total

    ws_dash["I8"] = "=C5+H8"
    ws_dash["I8"].font = Font(name="Segoe UI", size=13, bold=True, color=NAVY_DARK)
    ws_dash["I8"].alignment = align_center
    ws_dash["I8"].number_format = CURRENCY_FORMAT
    ws_dash["I8"].fill = fill_total
    ws_dash["I7"].border = cell_border
    ws_dash["I8"].border = cell_border

    # TABLA 1: FLUJO SEMANA A SEMANA DEL MES SELECCIONADO (B11:I18)
    ws_dash["B11"] = "SEGUIMIENTO SEMANAL DEL MES SELECCIONADO"
    ws_dash["B11"].font = font_table_header
    ws_dash["B11"].fill = fill_header
    ws_dash.merge_cells("B11:I11")

    sem_headers = [
        ("B12", "Semana"),
        ("C12", "Días del Mes"),
        ("D12", "Ingresos Semanales"),
        ("E12", "Gastos Operativos"),
        ("F12", "Ahorro / Inversión"),
        ("G12", "Flujo Neto Semanal"),
        ("H12", "Saldo al Cierre"),
        ("I12", "De los cuales: Hormiga")
    ]
    for cell_ref, text in sem_headers:
        ws_dash[cell_ref] = text
        ws_dash[cell_ref].font = font_section_header
        ws_dash[cell_ref].fill = fill_sub_blue
        ws_dash[cell_ref].alignment = align_center
        ws_dash[cell_ref].border = cell_border

    semanas_info = [
        ("Semana 1", "Días 01 al 07", 13),
        ("Semana 2", "Días 08 al 14", 14),
        ("Semana 3", "Días 15 al 21", 15),
        ("Semana 4", "Días 22 al 28", 16),
        ("Semana 5", "Días 29 al fin de mes", 17)
    ]

    for sem_name, dias, row in semanas_info:
        ws_dash[f"B{row}"] = sem_name
        ws_dash[f"C{row}"] = dias
        ws_dash[f"D{row}"] = f'=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$C$5:$C$1500, "{sem_name}", \'Registro Diario\'!$D$5:$D$1500, "Ingreso", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5)'
        ws_dash[f"E{row}"] = f'=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$C$5:$C$1500, "{sem_name}", \'Registro Diario\'!$D$5:$D$1500, "Gasto Fijo", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5) + SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$C$5:$C$1500, "{sem_name}", \'Registro Diario\'!$D$5:$D$1500, "Gasto Variable", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5) + SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$C$5:$C$1500, "{sem_name}", \'Registro Diario\'!$D$5:$D$1500, "Gasto Hormiga", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5)'
        ws_dash[f"F{row}"] = f'=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$C$5:$C$1500, "{sem_name}", \'Registro Diario\'!$D$5:$D$1500, "Ahorro / Inversión", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5)'
        ws_dash[f"G{row}"] = f'=D{row}-E{row}-F{row}'
        
        # Saldo Proyectado al Cierre
        if row == 13:
            ws_dash[f"H{row}"] = f'=$C$5+G{row}'
        else:
            ws_dash[f"H{row}"] = f'=H{row-1}+G{row}'

        ws_dash[f"I{row}"] = f'=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$C$5:$C$1500, "{sem_name}", \'Registro Diario\'!$D$5:$D$1500, "Gasto Hormiga", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5)'

        ws_dash[f"B{row}"].alignment = align_center
        ws_dash[f"C{row}"].alignment = align_center
        ws_dash[f"D{row}"].alignment = align_right
        ws_dash[f"E{row}"].alignment = align_right
        ws_dash[f"F{row}"].alignment = align_right
        ws_dash[f"G{row}"].alignment = align_right
        ws_dash[f"H{row}"].alignment = align_right
        ws_dash[f"I{row}"].alignment = align_right

        ws_dash[f"D{row}"].number_format = CURRENCY_FORMAT
        ws_dash[f"E{row}"].number_format = CURRENCY_FORMAT
        ws_dash[f"F{row}"].number_format = CURRENCY_FORMAT
        ws_dash[f"G{row}"].number_format = CURRENCY_FORMAT
        ws_dash[f"H{row}"].number_format = CURRENCY_FORMAT
        ws_dash[f"I{row}"].number_format = CURRENCY_FORMAT

        for c in ["B", "C", "D", "E", "F", "G", "H", "I"]:
            ws_dash[f"{c}{row}"].font = font_body
            ws_dash[f"{c}{row}"].border = cell_border

    tot_row = 18
    ws_dash[f"B{tot_row}"] = "TOTAL DEL MES"
    ws_dash.merge_cells(f"B{tot_row}:C{tot_row}")
    ws_dash[f"D{tot_row}"] = f'=SUM(D13:D17)'
    ws_dash[f"E{tot_row}"] = f'=SUM(E13:E17)'
    ws_dash[f"F{tot_row}"] = f'=SUM(F13:F17)'
    ws_dash[f"G{tot_row}"] = f'=SUM(G13:G17)'
    ws_dash[f"H{tot_row}"] = f'=H17'
    ws_dash[f"I{tot_row}"] = f'=SUM(I13:I17)'

    for c in ["B", "C", "D", "E", "F", "G", "H", "I"]:
        ws_dash[f"{c}{tot_row}"].font = font_body_bold
        ws_dash[f"{c}{tot_row}"].fill = fill_total
        ws_dash[f"{c}{tot_row}"].border = total_border
        if c in ["D", "E", "F", "G", "H", "I"]:
            ws_dash[f"{c}{tot_row}"].number_format = CURRENCY_FORMAT
            ws_dash[f"{c}{tot_row}"].alignment = align_right
        else:
            ws_dash[f"{c}{tot_row}"].alignment = align_center

    # Bar chart for weekly flow of selected month
    chart_bar = BarChart()
    chart_bar.type = "col"
    chart_bar.style = 10
    chart_bar.title = "Flujo Semanal: Ingresos vs. Gastos Operativos"
    chart_bar.y_axis.title = "Monto ($)"
    chart_bar.x_axis.title = "Semana"
    chart_bar.height = 7
    chart_bar.width = 13

    data_weekly = Reference(ws_dash, min_col=4, min_row=12, max_col=5, max_row=17)
    cats_weekly = Reference(ws_dash, min_col=2, min_row=13, max_row=17)
    chart_bar.add_data(data_weekly, titles_from_data=True)
    chart_bar.set_categories(cats_weekly)
    ws_dash.add_chart(chart_bar, "K11")

    # TABLA 2: CONTROL DE PRESUPUESTO POR CATEGORÍA PARA EL MES SELECCIONADO
    ws_dash["B20"] = "CONTROL DE PRESUPUESTO VS. GASTO REAL DEL MES"
    ws_dash["B20"].font = font_table_header
    ws_dash["B20"].fill = fill_header
    ws_dash.merge_cells("B20:H20")

    cat_table_headers = [
        ("B21", "Categoría"),
        ("C21", "Tipo de Gasto"),
        ("D21", "Presupuesto Mensual"),
        ("E21", "Gasto Real del Mes"),
        ("F21", "Diferencia / Disponible"),
        ("G21", "% Ejecutado"),
        ("H21", "Estado / Alerta")
    ]
    for cell_ref, text in cat_table_headers:
        ws_dash[cell_ref] = text
        ws_dash[cell_ref].font = font_section_header
        ws_dash[cell_ref].fill = fill_sub_blue
        ws_dash[cell_ref].alignment = align_center
        ws_dash[cell_ref].border = cell_border

    expense_categories = [item for item in categorias_data if item[1] != "Ingreso"]

    curr_cat_row = 22
    for cat_name, grp_name, pres_val in expense_categories:
        ws_dash[f"B{curr_cat_row}"] = cat_name
        ws_dash[f"C{curr_cat_row}"] = grp_name
        ws_dash[f"D{curr_cat_row}"] = pres_val
        ws_dash[f"E{curr_cat_row}"] = f'=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$E$5:$E$1500, "{cat_name}", \'Registro Diario\'!$B$5:$B$1500, \">=\"&$P$4, \'Registro Diario\'!$B$5:$B$1500, \"<=\"&$P$5)'
        ws_dash[f"F{curr_cat_row}"] = f'=D{curr_cat_row}-E{curr_cat_row}'
        ws_dash[f"G{curr_cat_row}"] = f'=IF(D{curr_cat_row}>0, E{curr_cat_row}/D{curr_cat_row}, 0)'
        ws_dash[f"H{curr_cat_row}"] = f'=IF(E{curr_cat_row}>D{curr_cat_row}, "⚠️ Excedido", IF(G{curr_cat_row}>0.85, "🟡 Alerta (+85%)", "✅ Bajo Control"))'

        ws_dash[f"B{curr_cat_row}"].alignment = align_left
        ws_dash[f"C{curr_cat_row}"].alignment = align_center
        ws_dash[f"D{curr_cat_row}"].alignment = align_right
        ws_dash[f"E{curr_cat_row}"].alignment = align_right
        ws_dash[f"F{curr_cat_row}"].alignment = align_right
        ws_dash[f"G{curr_cat_row}"].alignment = align_right
        ws_dash[f"H{curr_cat_row}"].alignment = align_center

        ws_dash[f"D{curr_cat_row}"].number_format = CURRENCY_FORMAT
        ws_dash[f"E{curr_cat_row}"].number_format = CURRENCY_FORMAT
        ws_dash[f"F{curr_cat_row}"].number_format = CURRENCY_FORMAT
        ws_dash[f"G{curr_cat_row}"].number_format = PERCENT_FORMAT

        for c in ["B", "C", "D", "E", "F", "G", "H"]:
            ws_dash[f"{c}{curr_cat_row}"].font = font_body
            ws_dash[f"{c}{curr_cat_row}"].border = cell_border

        curr_cat_row += 1

    tot_cat_row = curr_cat_row
    ws_dash[f"B{tot_cat_row}"] = "TOTAL PRESUPUESTO Y GASTOS MONITOREADOS"
    ws_dash.merge_cells(f"B{tot_cat_row}:C{tot_cat_row}")
    ws_dash[f"D{tot_cat_row}"] = f'=SUM(D22:D{tot_cat_row-1})'
    ws_dash[f"E{tot_cat_row}"] = f'=SUM(E22:E{tot_cat_row-1})'
    ws_dash[f"F{tot_cat_row}"] = f'=D{tot_cat_row}-E{tot_cat_row}'
    ws_dash[f"G{tot_cat_row}"] = f'=IF(D{tot_cat_row}>0, E{tot_cat_row}/D{tot_cat_row}, 0)'
    ws_dash[f"H{tot_cat_row}"] = f'=IF(E{tot_cat_row}>D{tot_cat_row}, "⚠️ Excedido", "✅ Global OK")'

    for c in ["B", "C", "D", "E", "F", "G", "H"]:
        ws_dash[f"{c}{tot_cat_row}"].font = font_body_bold
        ws_dash[f"{c}{tot_cat_row}"].fill = fill_total
        ws_dash[f"{c}{tot_cat_row}"].border = total_border
        if c in ["D", "E", "F"]:
            ws_dash[f"{c}{tot_cat_row}"].number_format = CURRENCY_FORMAT
            ws_dash[f"{c}{tot_cat_row}"].alignment = align_right
        elif c == "G":
            ws_dash[f"{c}{tot_cat_row}"].number_format = PERCENT_FORMAT
            ws_dash[f"{c}{tot_cat_row}"].alignment = align_right
        else:
            ws_dash[f"{c}{tot_cat_row}"].alignment = align_center

    # Mini-table for Pie Chart
    ws_dash["K26"] = "Distribución del Gasto del Mes"
    ws_dash["L26"] = "Monto ($)"
    ws_dash["K26"].font = font_table_header
    ws_dash["K26"].fill = fill_header
    ws_dash["L26"].font = font_table_header
    ws_dash["L26"].fill = fill_header

    pie_items = [
        ("Gastos Fijos", '=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$D$5:$D$1500, "Gasto Fijo", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5)'),
        ("Gastos Variables", '=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$D$5:$D$1500, "Gasto Variable", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5)'),
        ("Gastos Hormiga", '=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$D$5:$D$1500, "Gasto Hormiga", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5)'),
        ("Ahorro / Inversión", '=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$D$5:$D$1500, "Ahorro / Inversión", \'Registro Diario\'!$B$5:$B$1500, ">="&$P$4, \'Registro Diario\'!$B$5:$B$1500, "<="&$P$5)')
    ]
    for p_idx, (p_label, p_formula) in enumerate(pie_items, start=27):
        ws_dash[f"K{p_idx}"] = p_label
        ws_dash[f"L{p_idx}"] = p_formula
        ws_dash[f"K{p_idx}"].font = font_body
        ws_dash[f"L{p_idx}"].font = font_body
        ws_dash[f"K{p_idx}"].border = cell_border
        ws_dash[f"L{p_idx}"].border = cell_border
        ws_dash[f"L{p_idx}"].number_format = CURRENCY_FORMAT

    chart_pie = PieChart()
    chart_pie.title = "Composición de Salidas de Dinero del Mes"
    chart_pie.height = 7
    chart_pie.width = 12
    pie_data = Reference(ws_dash, min_col=12, min_row=26, max_row=30)
    pie_cats = Reference(ws_dash, min_col=11, min_row=27, max_row=30)
    chart_pie.add_data(pie_data, titles_from_data=True)
    chart_pie.set_categories(pie_cats)
    ws_dash.add_chart(chart_pie, "K32")

    ws_dash.column_dimensions["A"].width = 3
    ws_dash.column_dimensions["B"].width = 40
    ws_dash.column_dimensions["C"].width = 20
    ws_dash.column_dimensions["D"].width = 22
    ws_dash.column_dimensions["E"].width = 22
    ws_dash.column_dimensions["F"].width = 22
    ws_dash.column_dimensions["G"].width = 20
    ws_dash.column_dimensions["H"].width = 22
    ws_dash.column_dimensions["I"].width = 24
    ws_dash.column_dimensions["J"].width = 4
    ws_dash.column_dimensions["K"].width = 24
    ws_dash.column_dimensions["L"].width = 20

    # =============================================================
    # 4. SHEET: CONSOLIDADO ANUAL (Master 15-month matrix: Oct 26 - Dec 27)
    # =============================================================
    ws_annual = wb.create_sheet(title="Consolidado Anual")
    ws_annual.views.sheetView[0].showGridLines = True

    ws_annual["A1"] = "CONSOLIDADO MULTI-ANUAL DE FLUJO DE CAJA (OCTUBRE 2026 - DICIEMBRE 2027)"
    ws_annual["A1"].font = font_main_title
    ws_annual["A2"] = "Visión panorámica mes a mes: Resumen de liquidez, evolución de ahorros y desglose completo por categorías."
    ws_annual["A2"].font = font_subtitle

    # Define columns for the 15 months + Totals
    # Col C: Oct 26, Col D: Nov 26, Col E: Dic 26, Col F: TOTAL Q4 2026
    # Col G: Ene 27 ... Col R: Dic 27, Col S: TOTAL AÑO 2027
    month_cols = [
        ("C", "Oct-26", date(2026, 10, 1), date(2026, 10, 31), False),
        ("D", "Nov-26", date(2026, 11, 1), date(2026, 11, 30), False),
        ("E", "Dic-26", date(2026, 12, 1), date(2026, 12, 31), False),
        ("F", "TOTAL Q4 2026", None, None, True),
        ("G", "Ene-27", date(2027, 1, 1), date(2027, 1, 31), False),
        ("H", "Feb-27", date(2027, 2, 1), date(2027, 2, 28), False),
        ("I", "Mar-27", date(2027, 3, 1), date(2027, 3, 31), False),
        ("J", "Abr-27", date(2027, 4, 1), date(2027, 4, 30), False),
        ("K", "May-27", date(2027, 5, 1), date(2027, 5, 31), False),
        ("L", "Jun-27", date(2027, 6, 1), date(2027, 6, 30), False),
        ("M", "Jul-27", date(2027, 7, 1), date(2027, 7, 31), False),
        ("N", "Ago-27", date(2027, 8, 1), date(2027, 8, 31), False),
        ("O", "Sep-27", date(2027, 9, 1), date(2027, 9, 30), False),
        ("P", "Oct-27", date(2027, 10, 1), date(2027, 10, 31), False),
        ("Q", "Nov-27", date(2027, 11, 1), date(2027, 11, 30), False),
        ("R", "Dic-27", date(2027, 12, 1), date(2027, 12, 31), False),
        ("S", "TOTAL AÑO 2027", None, None, True)
    ]

    # Row 4 (StartDate) and Row 5 (EndDate) - hidden for clean formulas
    for col_letter, lbl, d_start, d_end, is_tot in month_cols:
        if not is_tot:
            ws_annual[f"{col_letter}4"] = d_start
            ws_annual[f"{col_letter}5"] = d_end
            ws_annual[f"{col_letter}4"].number_format = DATE_FORMAT
            ws_annual[f"{col_letter}5"].number_format = DATE_FORMAT
    ws_annual.row_dimensions[4].hidden = True
    ws_annual.row_dimensions[5].hidden = True

    # Headers in Row 6
    ws_annual["A6"] = "Concepto / Categoría"
    ws_annual["B6"] = "Grupo de Gasto"
    ws_annual["A6"].font = font_table_header
    ws_annual["A6"].fill = fill_header
    ws_annual["A6"].alignment = align_center
    ws_annual["A6"].border = cell_border

    ws_annual["B6"].font = font_table_header
    ws_annual["B6"].fill = fill_header
    ws_annual["B6"].alignment = align_center
    ws_annual["B6"].border = cell_border

    for col_letter, lbl, _, _, is_tot in month_cols:
        cell_ref = f"{col_letter}6"
        ws_annual[cell_ref] = lbl
        ws_annual[cell_ref].font = font_table_header
        ws_annual[cell_ref].fill = fill_header if not is_tot else PatternFill(start_color="112239", end_color="112239", fill_type="solid")
        ws_annual[cell_ref].alignment = align_center
        ws_annual[cell_ref].border = cell_border

    # SECTION 1: RESUMEN DE FLUJO DE CAJA (Rows 8 to 16)
    ws_annual["A7"] = "1. RESUMEN DE FLUJO DE CAJA (CASH FLOW)"
    ws_annual["A7"].font = font_section_header
    ws_annual["A7"].fill = fill_sub_blue
    ws_annual.merge_cells("A7:S7")

    cf_rows = [
        (8, "Saldo Inicial de Caja", "Balance"),
        (9, "(+) Total Ingresos", "Ingreso"),
        (10, "(-) Gastos Fijos", "Gasto Fijo"),
        (11, "(-) Gastos Variables", "Gasto Variable"),
        (12, "(-) Gastos Hormiga (Fugas)", "Gasto Hormiga"),
        (13, "(-) Ahorro e Inversión", "Ahorro / Inversión"),
        (14, "(=) Flujo Neto del Período", "Neto"),
        (15, "(=) Saldo Final Acumulado", "Saldo Final"),
        (16, "Tasa de Ahorro (%)", "Ratio")
    ]

    for r_num, r_label, r_type in cf_rows:
        ws_annual[f"A{r_num}"] = r_label
        ws_annual[f"B{r_num}"] = r_type
        ws_annual[f"A{r_num}"].font = font_body_bold if r_type in ["Neto", "Saldo Final", "Balance"] else font_body
        ws_annual[f"B{r_num}"].font = font_body
        ws_annual[f"A{r_num}"].border = cell_border
        ws_annual[f"B{r_num}"].border = cell_border
        ws_annual[f"B{r_num}"].alignment = align_center

        if r_type in ["Neto", "Saldo Final"]:
            ws_annual[f"A{r_num}"].fill = fill_total
            ws_annual[f"B{r_num}"].fill = fill_total

        # Populate columns C to S
        for col_letter, _, _, _, is_tot in month_cols:
            c_ref = f"{col_letter}{r_num}"

            if r_type == "Balance":
                # Saldo Inicial
                if col_letter == "C":
                    ws_annual[c_ref] = "=Configuracion!$K$2"
                elif col_letter == "D":
                    ws_annual[c_ref] = "=C15"
                elif col_letter == "E":
                    ws_annual[c_ref] = "=D15"
                elif col_letter == "F": # Total Q4 2026 starts with C8
                    ws_annual[c_ref] = "=C8"
                elif col_letter == "G": # Ene 27 starts with Dic 26 ending
                    ws_annual[c_ref] = "=E15"
                elif col_letter == "S": # Total 2027 starts with G8
                    ws_annual[c_ref] = "=G8"
                else: # Sequential months
                    prev_col = chr(ord(col_letter) - 1)
                    ws_annual[c_ref] = f"={prev_col}15"

            elif r_type in ["Ingreso", "Gasto Fijo", "Gasto Variable", "Gasto Hormiga", "Ahorro / Inversión"]:
                if is_tot:
                    if col_letter == "F":
                        ws_annual[c_ref] = f"=SUM(C{r_num}:E{r_num})"
                    elif col_letter == "S":
                        ws_annual[c_ref] = f"=SUM(G{r_num}:R{r_num})"
                else:
                    ws_annual[c_ref] = f'=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$D$5:$D$1500, "{r_type}", \'Registro Diario\'!$B$5:$B$1500, \">=\"&{col_letter}$4, \'Registro Diario\'!$B$5:$B$1500, \"<=\"&{col_letter}$5)'

            elif r_type == "Neto":
                # Flujo Neto = Ingresos - Gastos Fijos - Gastos Var - Hormiga - Ahorro
                ws_annual[c_ref] = f"={col_letter}9-{col_letter}10-{col_letter}11-{col_letter}12-{col_letter}13"

            elif r_type == "Saldo Final":
                # Saldo Final = Saldo Inicial + Flujo Neto
                ws_annual[c_ref] = f"={col_letter}8+{col_letter}14"

            elif r_type == "Ratio":
                # Tasa de Ahorro = (Ahorro + Flujo Neto) / Ingresos si Ingresos > 0
                ws_annual[c_ref] = f"=IF({col_letter}9>0, ({col_letter}13+{col_letter}14)/{col_letter}9, 0)"

            # Styling
            ws_annual[c_ref].border = cell_border
            ws_annual[c_ref].alignment = align_right
            if r_type in ["Neto", "Saldo Final"]:
                ws_annual[c_ref].font = font_body_bold
                ws_annual[c_ref].fill = fill_total
            else:
                ws_annual[c_ref].font = font_body

            if r_type == "Ratio":
                ws_annual[c_ref].number_format = PERCENT_FORMAT
            else:
                ws_annual[c_ref].number_format = CURRENCY_FORMAT

    # SECTION 2: DESGLOSE COMPLETO POR CATEGORÍA MES A MES (Row 18 onwards)
    ws_annual["A18"] = "2. EVOLUCIÓN MENSUAL POR CATEGORÍA DE GASTO E INGRESO"
    ws_annual["A18"].font = font_section_header
    ws_annual["A18"].fill = fill_sub_blue
    ws_annual.merge_cells("A18:S18")

    row_cat_start = 19
    for c_idx, (cat_name, grp_name, _) in enumerate(categorias_data, start=row_cat_start):
        ws_annual[f"A{c_idx}"] = cat_name
        ws_annual[f"B{c_idx}"] = grp_name
        ws_annual[f"A{c_idx}"].font = font_body
        ws_annual[f"B{c_idx}"].font = font_body
        ws_annual[f"A{c_idx}"].border = cell_border
        ws_annual[f"B{c_idx}"].border = cell_border
        ws_annual[f"B{c_idx}"].alignment = align_center

        for col_letter, _, _, _, is_tot in month_cols:
            c_ref = f"{col_letter}{c_idx}"
            if is_tot:
                if col_letter == "F":
                    ws_annual[c_ref] = f"=SUM(C{c_idx}:E{c_idx})"
                elif col_letter == "S":
                    ws_annual[c_ref] = f"=SUM(G{c_idx}:R{c_idx})"
            else:
                ws_annual[c_ref] = f'=SUMIFS(\'Registro Diario\'!$H$5:$H$1500, \'Registro Diario\'!$E$5:$E$1500, $A{c_idx}, \'Registro Diario\'!$B$5:$B$1500, \">=\"&{col_letter}$4, \'Registro Diario\'!$B$5:$B$1500, \"<=\"&{col_letter}$5)'

            ws_annual[c_ref].font = font_body
            ws_annual[c_ref].border = cell_border
            ws_annual[c_ref].alignment = align_right
            ws_annual[c_ref].number_format = CURRENCY_FORMAT

    # Freeze panes so category names and months stay visible when scrolling
    ws_annual.freeze_panes = "C7"

    ws_annual.column_dimensions["A"].width = 38
    ws_annual.column_dimensions["B"].width = 20
    for col_letter, _, _, _, is_tot in month_cols:
        ws_annual.column_dimensions[col_letter].width = 18 if is_tot else 14

    # =============================================================
    # 5. SHEET: PROYECCION SEMANAL (Weekly Liquidity Planner)
    # =============================================================
    ws_plan = wb.create_sheet(title="Proyeccion Semanal")
    ws_plan.views.sheetView[0].showGridLines = True

    ws_plan["A1"] = "PLANIFICACIÓN Y PROYECCIÓN DE LIQUIDEZ SEMANAL"
    ws_plan["A1"].font = font_main_title
    ws_plan["A2"] = "Anticipa cobros y vencimientos semana a semana para evitar quedar sin liquidez antes del próximo ingreso."
    ws_plan["A2"].font = font_subtitle

    # Month selector in E2
    ws_plan["D2"] = "Mes a Planificar:"
    ws_plan["D2"].font = font_body_bold
    ws_plan["D2"].alignment = align_right
    ws_plan["E2"] = "Noviembre 2026"
    ws_plan["E2"].font = font_body_bold
    ws_plan["E2"].fill = fill_sub_blue
    ws_plan["E2"].border = cell_border
    ws_plan["E2"].alignment = align_center

    dv_plan_mes = DataValidation(type="list", formula1="=Configuracion!$H$5:$H$19", allow_blank=False)
    ws_plan.add_data_validation(dv_plan_mes)
    dv_plan_mes.add("E2")

    plan_headers = [
        ("A4", "Semana"),
        ("B4", "Período"),
        ("C4", "Saldo Inicial Estimado"),
        ("D4", "(+) Ingresos Previstos"),
        ("E4", "(-) Pagos Fijos Programados"),
        ("F4", "(-) Presupuesto Gastos Diarios"),
        ("G4", "(-) Ahorro Planificado"),
        ("H4", "(=) Saldo Estimado al Cierre"),
        ("I4", "Semáforo de Liquidez")
    ]
    for cell_ref, text in plan_headers:
        ws_plan[cell_ref] = text
        ws_plan[cell_ref].font = font_table_header
        ws_plan[cell_ref].fill = fill_header
        ws_plan[cell_ref].alignment = align_center
        ws_plan[cell_ref].border = cell_border

    semanas_plan = [
        ("Semana 1", "Días 01 al 07", 5, "='Tablero Mensual'!$C$5", 0, 220000, 45000, 20000),
        ("Semana 2", "Días 08 al 14", 6, "=H5", 0, 45000, 45000, 20000),
        ("Semana 3", "Días 15 al 21", 7, "=H6", 0, 30000, 45000, 0),
        ("Semana 4", "Días 22 al 28", 8, "=H7", 0, 25000, 45000, 0),
        ("Semana 5", "Días 29 al fin de mes", 9, "=H8", 600000, 230000, 30000, 0)
    ]

    for sem_label, per_label, row_p, s_ini, inc_p, fix_p, var_p, ahr_p in semanas_plan:
        ws_plan[f"A{row_p}"] = sem_label
        ws_plan[f"B{row_p}"] = per_label
        ws_plan[f"C{row_p}"] = s_ini
        ws_plan[f"D{row_p}"] = inc_p
        ws_plan[f"E{row_p}"] = fix_p
        ws_plan[f"F{row_p}"] = var_p
        ws_plan[f"G{row_p}"] = ahr_p
        ws_plan[f"H{row_p}"] = f'=C{row_p}+D{row_p}-E{row_p}-F{row_p}-G{row_p}'
        ws_plan[f"I{row_p}"] = f'=IF(H{row_p}<0, "🚨 Déficit (Sin Liquidez)", IF(H{row_p}<30000, "⚠️ Alerta Saldo Mínimo", "✅ Liquidez Saludable"))'

        ws_plan[f"A{row_p}"].alignment = align_center
        ws_plan[f"B{row_p}"].alignment = align_center
        ws_plan[f"C{row_p}"].alignment = align_right
        ws_plan[f"D{row_p}"].alignment = align_right
        ws_plan[f"E{row_p}"].alignment = align_right
        ws_plan[f"F{row_p}"].alignment = align_right
        ws_plan[f"G{row_p}"].alignment = align_right
        ws_plan[f"H{row_p}"].alignment = align_right
        ws_plan[f"I{row_p}"].alignment = align_center

        for col_p in ["C", "D", "E", "F", "G", "H"]:
            ws_plan[f"{col_p}{row_p}"].number_format = CURRENCY_FORMAT

        for col_p in ["A", "B", "C", "D", "E", "F", "G", "H", "I"]:
            ws_plan[f"{col_p}{row_p}"].font = font_body
            ws_plan[f"{col_p}{row_p}"].border = cell_border

    ws_plan.column_dimensions["A"].width = 14
    ws_plan.column_dimensions["B"].width = 20
    ws_plan.column_dimensions["C"].width = 24
    ws_plan.column_dimensions["D"].width = 22
    ws_plan.column_dimensions["E"].width = 26
    ws_plan.column_dimensions["F"].width = 28
    ws_plan.column_dimensions["G"].width = 22
    ws_plan.column_dimensions["H"].width = 26
    ws_plan.column_dimensions["I"].width = 28

    # =============================================================
    # 6. SHEET: GUIA Y CONSEJOS (Updated with multi-year guide)
    # =============================================================
    ws_guide = wb.create_sheet(title="Guia y Consejos")
    ws_guide.views.sheetView[0].showGridLines = True

    ws_guide["A1"] = "GUÍA DE USO Y PLANIFICACIÓN MULTI-ANUAL (2026 - 2027)"
    ws_guide["A1"].font = font_main_title
    ws_guide["A2"] = "Aprende a sostener el control continuo de tus gastos y a planificar con anticipación todo el año 2027."
    ws_guide["A2"].font = font_subtitle

    guide_sections = [
        ("1. ¿Cómo funciona la planilla extendida a 2026 y 2027?", [
            "• El 'Registro Diario' es tu único libro contable continuo: aquí anotas todo lo que ocurre desde octubre de 2026 hasta diciembre de 2027.",
            "• El 'Tablero Mensual' tiene un selector desplegable en la celda F5: al cambiar de mes (ej. pasar de Octubre a Noviembre 2026 o a Enero 2027), todo el tablero, las semanas, gráficos y categorías se filtran automáticamente para ese período.",
            "• La hoja 'Consolidado Anual' muestra los 15 meses en una sola pantalla: te permite comparar meses entre sí, detectar picos de gasto estacionales y ver el ahorro total acumulado de cada año."
        ]),
        ("2. Continuidad Automática de Saldos entre Meses", [
            "• El Saldo Inicial de cada mes se calcula automáticamente sumando todos los ingresos y restando todos los gastos ocurridos con anterioridad a esa fecha.",
            "• Esto garantiza que el dinero sobrante de fin de octubre pase directamente como saldo inicial a noviembre, y lo de diciembre pase a enero de 2027 sin necesidad de copiar valores a mano."
        ]),
        ("3. El Hábito Cotidiano de los 3 Minutos", [
            "• Registra tus movimientos el mismo día: al finalizar la jornada o antes de acostarte, abre el 'Registro Diario'.",
            "• Basta con colocar la Fecha, elegir Tipo y Categoría del menú, escribir el Detalle y el Monto.",
            "• El número de registro (#) y la semana del mes se calculan automáticamente."
        ]),
        ("4. Planificación del Año 2027 en 'Consolidado Anual'", [
            "• En 'Consolidado Anual' puedes observar la columna 'TOTAL Q4 2026' para ver cómo cierras el año actual y la columna 'TOTAL AÑO 2027' para medir tu progreso a largo plazo.",
            "• Si prevés gastos extraordinarios (vacaciones en enero/febrero, matrículas escolares en marzo o pagos anuales de seguros), planifícalos en 'Proyeccion Semanal' para saber cuánta liquidez reservar."
        ]),
        ("5. Personalización de Presupuestos y Categorías", [
            "• En la hoja 'Configuracion' puedes editar el monto del 'Presupuesto Mensual Base' de cualquier categoría para adaptarlo a tus ingresos actuales o ajustarlo ante inflación."
        ])
    ]

    curr_row = 4
    for title, points in guide_sections:
        ws_guide[f"A{curr_row}"] = title
        ws_guide[f"A{curr_row}"].font = font_section_header
        ws_guide[f"A{curr_row}"].fill = fill_sub_blue
        ws_guide[f"A{curr_row}"].border = cell_border
        curr_row += 1
        for p in points:
            ws_guide[f"A{curr_row}"] = p
            ws_guide[f"A{curr_row}"].font = font_body
            curr_row += 1
        curr_row += 1

    ws_guide.column_dimensions["A"].width = 115

    # Remove default empty sheet
    wb.remove(default_sheet)

    # Order sheets logically:
    # 1. Tablero Mensual (Dynamic Dashboard)
    # 2. Consolidado Anual (Master 15-month Matrix)
    # 3. Registro Diario (Continuous Daily Ledger)
    # 4. Proyeccion Semanal (Weekly Cash Flow Forecast)
    # 5. Configuracion (Catalogs & 15-Month Calendar)
    # 6. Guia y Consejos (Handbook)
    wb._sheets = [ws_dash, ws_annual, ws_daily, ws_plan, ws_config, ws_guide]

    # Save to user's official directory
    wb.save(output_path)
    print(f"Multi-year Cash Flow successfully built and saved to: {output_path}")

if __name__ == "__main__":
    build_multi_year_cash_flow()
