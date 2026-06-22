import tkinter as tk
from tkinter import messagebox, ttk
from tkcalendar import Calendar
import pandas as pd
import os
from datetime import datetime
from fpdf import FPDF
from fpdf.enums import XPos, YPos
import customtkinter as ctk
from tkcalendar import Calendar # Sigue siendo compatible

# Configuración global de estilo
ctk.set_appearance_mode("System")  
ctk.set_default_color_theme("blue")



# --- CONFIGURACIÓN DE CARPETAS ---
ESCRITORIO = os.path.join(os.path.expanduser("~"), "Desktop")
CARPETA_RAIZ = os.path.join(ESCRITORIO, "Racon Asistencias", "2026")
ARCHIVO = "Racon_Asistencia.xlsx"
MESES = ["01 - Enero", "02 - Febrero", "03 - Marzo", "04 - Abril", "05 - Mayo", "06 - Junio", 
         "07 - Julio", "08 - Agosto", "09 - Septiembre", "10 - Octubre", "11 - Noviembre", "12 - Diciembre", "13 - Anual"]

def crear_estructura_carpetas():
    if not os.path.exists(CARPETA_RAIZ):
        os.makedirs(CARPETA_RAIZ, exist_ok=True) # Crea la carpeta 2026 primero
    
    for mes in MESES:
        ruta_mes = os.path.join(CARPETA_RAIZ, mes)
        if not os.path.exists(ruta_mes):
            os.makedirs(ruta_mes)

# --- FUNCIONES DE GESTIÓN DE DATOS ---
def guardar_archivo(df_personal, df_asistencia, df_pagos, df_bonos):
    try:
        with pd.ExcelWriter(ARCHIVO, engine='openpyxl') as writer:
            df_personal.to_excel(writer, sheet_name='Personal', index=False)
            df_asistencia.to_excel(writer, sheet_name='Asistencia', index=False)
            df_pagos.to_excel(writer, sheet_name='Pagos', index=False)
            df_bonos.to_excel(writer, sheet_name='Bonos', index=False)
    except Exception as e:
        messagebox.showerror("Error", f"Error al guardar: {e}")

def leer_datos(sheet_name):
    columnas = {
        'Personal': ['Nombre', 'Apellido'],
        'Asistencia': ['Fecha', 'Nombre', 'Apellido', 'Proyecto'],
        'Pagos': ['Nombre', 'Apellido', 'SueldoDiario'],
        'Bonos': ['Fecha', 'Nombre', 'Apellido', 'Tipo', 'Descripcion', 'Monto']
    }
    try:
        if not os.path.exists(ARCHIVO):
            return pd.DataFrame(columns=columnas.get(sheet_name, []))
        xls = pd.ExcelFile(ARCHIVO, engine='openpyxl')
        if sheet_name not in xls.sheet_names:
            return pd.DataFrame(columns=columnas.get(sheet_name, []))
        df = pd.read_excel(ARCHIVO, sheet_name=sheet_name, engine='openpyxl')
        if sheet_name in ['Asistencia', 'Bonos'] and not df.empty:
            df['Fecha'] = pd.to_datetime(df['Fecha'], errors='coerce')
        return df
    except:
        return pd.DataFrame(columns=columnas.get(sheet_name, []))

def inicializar_excel():
    if not os.path.exists(ARCHIVO):
        guardar_archivo(
            pd.DataFrame(columns=['Nombre', 'Apellido']),
            pd.DataFrame(columns=['Fecha', 'Nombre', 'Apellido', 'Proyecto']),
            pd.DataFrame(columns=['Nombre', 'Apellido', 'SueldoDiario']),
            pd.DataFrame(columns=['Fecha', 'Nombre', 'Apellido', 'Tipo', 'Descripcion', 'Monto'])
        )
    else:
        # Si el archivo ya existe pero no tiene la hoja Bonos, la agrega
        df_personal = leer_datos('Personal')
        df_asistencia = leer_datos('Asistencia')
        df_pagos = leer_datos('Pagos')
        df_bonos = leer_datos('Bonos')
        guardar_archivo(df_personal, df_asistencia, df_pagos, df_bonos)

# --- GENERADOR DE PDF ---
def generar_pdf_personal(nombre, apellido, mes_idx):
    mes_nombre = MESES[mes_idx]
    anio = datetime.now().year

    
    # 1. DETERMINAR RANGO Y DÍAS HÁBILES
    if mes_idx == 12:  # Caso ANUAL
        start_date = f"{anio}-01-01"
        end_date = f"{anio}-12-31"
        titulo_mes = f"Año {anio}"
    else:
        mes_num = mes_idx + 1
        start_date = f"{anio}-{mes_num:02d}-01"
        ultimo_dia = pd.Period(f'{anio}-{mes_num}').days_in_month
        end_date = f"{anio}-{mes_num:02d}-{ultimo_dia}"
        titulo_mes = mes_nombre
    
    dias_rango = pd.date_range(start=start_date, end=end_date)
    # Calculamos todos los días hábiles (Lunes a Viernes) del periodo
    habiles = [d.normalize() for d in dias_rango if d.weekday() < 5]

    # 2. CARGA Y FILTRADO DE DATOS
    df_a = leer_datos('Asistencia')
    df_p = leer_datos('Pagos')
    df_b = leer_datos('Bonos')
    
    df_a['Fecha'] = pd.to_datetime(df_a['Fecha'])
    if not df_b.empty: df_b['Fecha'] = pd.to_datetime(df_b['Fecha'])
    
    # Filtro para este empleado
    asistencias = df_a[(df_a['Nombre'] == nombre) & (df_a['Apellido'] == apellido) &
                       (df_a['Fecha'] >= start_date) & (df_a['Fecha'] <= end_date)]
    
    bonos = df_b[(df_b['Nombre'] == nombre) & (df_b['Apellido'] == apellido) &
                 (df_b['Fecha'] >= start_date) & (df_b['Fecha'] <= end_date)]
    
    # Cálculos
    dias_trabajados = len(asistencias)
    sueldo_row = df_p[(df_p['Nombre'] == nombre) & (df_p['Apellido'] == apellido)]
    diario = float(sueldo_row['SueldoDiario'].iloc[0]) if not sueldo_row.empty else 0.0
    
    pago_base = dias_trabajados * diario
    total_bonos = bonos['Monto'].sum() if not bonos.empty else 0.0
    total_pagar = pago_base + total_bonos
    
    # Cálculo de faltas
    fechas_asistidas = asistencias['Fecha'].dt.normalize().tolist()
    faltas = [d.strftime('%Y-%m-%d') for d in habiles if d not in fechas_asistidas]

    # 3. GENERACIÓN PDF
    pdf = FPDF()
    pdf.set_margins(10, 10, 10)
    pdf.add_page()
    
    # Encabezado
    pdf.set_font("helvetica", 'B', 16)
    pdf.cell(0, 10, text=f"Reporte: {nombre} {apellido}", new_x=XPos.LMARGIN, new_y=YPos.NEXT, align='C')
    
    # Cuerpo principal
    pdf.set_font("helvetica", size=12)
    pdf.ln(5)
    pdf.cell(0, 10, text=f"Mes: {mes_nombre}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 10, text=f"Dias trabajados: {dias_trabajados}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 10, text=f"Pago base: {pago_base:,.2f} EUR", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 10, text=f"Bonos y extras: {total_bonos:,.2f} EUR", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.cell(0, 10, text=f"Total a pagar: {total_pagar:,.2f} EUR", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    
    # Sección Bonos Detallados
    pdf.ln(5)
    pdf.set_font("helvetica", 'B', 12)
    pdf.cell(0, 10, text="Detalle de Bonos:", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("helvetica", size=10)
    if bonos.empty:
        pdf.cell(0, 8, text="No hay bonos registrados.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        for _, row in bonos.iterrows():
            f_txt = row['Fecha'].strftime('%Y-%m-%d')
            pdf.cell(0, 7, text=f"- {f_txt} | {row['Tipo']}: {row['Monto']:,.2f} EUR", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # Sección Faltas
    pdf.ln(5)
    pdf.set_font("helvetica", 'B', 12)
    pdf.cell(0, 10, text="Días hábiles no trabajados (Faltas):", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    pdf.set_font("helvetica", size=10)
    if not faltas:
        pdf.cell(0, 8, text="No hay faltas registradas.", new_x=XPos.LMARGIN, new_y=YPos.NEXT)
    else:
        for f in faltas:
            pdf.cell(0, 7, text=f"- {f}", new_x=XPos.LMARGIN, new_y=YPos.NEXT)

    # 4. GUARDADO
    ruta_carpeta = os.path.join(CARPETA_RAIZ, mes_nombre)
    if not os.path.exists(ruta_carpeta): os.makedirs(ruta_carpeta)
    
    nombre_archivo = f"{nombre}.{apellido}.{mes_nombre.lower()}.pdf"
    ruta_final = os.path.join(ruta_carpeta, nombre_archivo)
    
    # Manejo de contadores para evitar PermissionError
    contador = 1
    while os.path.exists(ruta_final):
        ruta_final = os.path.join(ruta_carpeta, f"{nombre}.{apellido}.{mes_nombre.lower()}_{contador}.pdf")
        contador += 1
    
    pdf.output(ruta_final)
    messagebox.showinfo("Éxito", f"PDF generado:\n{os.path.basename(ruta_final)}")
    
    

    
def abrir_calendario_asistencia(nombre, apellido, mes=None, anio=None):
    mes = mes or datetime.now().month
    anio = anio or datetime.now().year
    v = ctk.CTkToplevel()
    v.title(f"Asistencia: {nombre} {apellido}")
    
    centrar_ventana(v, 400, 550)
    v.attributes("-topmost", True)
    v.after(100, lambda: v.attributes("-topmost", False))
    
    # 1. Preparación de datos
    df_a = leer_datos('Asistencia')
    df_a['Fecha'] = pd.to_datetime(df_a['Fecha'])
    df_emp = df_a[(df_a['Nombre'].str.strip() == nombre.strip()) & 
                  (df_a['Apellido'].str.strip() == apellido.strip())].copy()

    # 2. Función de actualización del contador (Corregida)
    def actualizar_contador(e=None):
        # Usamos ._date, que es la propiedad interna que rastrea la VISTA actual
        # y no la fecha seleccionada.
        mes_visible = cal._date.month
        anio_visible = cal._date.year
        
        # Filtramos
        df_mes = df_emp[
            (df_emp['Fecha'].dt.month == mes_visible) & 
            (df_emp['Fecha'].dt.year == anio_visible)
        ]
        
        cantidad = len(df_mes)
        
        # 4. Actualizar interfaz
        lbl_contador.configure(text=f"Días trabajados: {cantidad}")
        

    # 3. Creación del Calendario
    cal = Calendar(v, selectmode='none', date_pattern='yyyy-mm-dd', 
                   year=anio, month=mes, 
                   background="#3498db", headersbackground="#2980b9", 
                   showothermonthdays=False, locale='es_ES',
                   font=("Segoe UI", 14), headersfont=("Segoe UI", 16, "bold"))
    cal.pack(pady=(20, 10), padx=20, fill="both", expand=True)
    

    # Marcado de días
    for _, row in df_emp.iterrows():
        cal.calevent_create(row['Fecha'].date(), 'Trabajado', 'work')
    cal.tag_config('work', background='#a9dfbf', foreground='#1e8449')

    # 4. EVENTOS (Aquí está el cambio clave)
    # Vinculamos tanto el cambio de mes como el clic en el calendario
    cal.bind("<<CalendarMonthChanged>>", actualizar_contador)
    cal.bind("<Button-1>", lambda e: v.after(50, actualizar_contador))

    # 5. Etiqueta del contador
    lbl_contador = ctk.CTkLabel(v, text="Días trabajados: 0", 
                                font=("Segoe UI", 16, "bold"), text_color="#2c3e50")
    lbl_contador.pack(pady=10)

    # 6. Forzar actualización al abrir
    v.after(300, actualizar_contador)

    ctk.CTkButton(v, text="Cerrar", command=v.destroy, fg_color="#7f8c8d").pack(pady=20)

def centrar_ventana(v, ancho, alto):
    # Obtener el ancho y alto de la pantalla del usuario
    pantalla_ancho = v.winfo_screenwidth()
    pantalla_alto = v.winfo_screenheight()
    
    # Calcular las coordenadas x e y para centrar
    x = (pantalla_ancho // 2) - (ancho // 2)
    y = (pantalla_alto // 2) - (alto // 2)
    
    # Aplicar la geometría con la posición calculada
    v.geometry(f"{ancho}x{alto}+{x}+{y}")
    
# --- VENTANAS ------------------------------------------------------------------
def ventana_bonos():
    v = ctk.CTkToplevel()
    v.title("Registrar Bonos / Extras")
    ancho, alto = 900, 700

    centrar_ventana(v, ancho, alto)
    v.configure(bg="white")
    v.attributes("-topmost", True)
    v.after(100, lambda: v.attributes("-topmost", False))

    main_frame = ctk.CTkFrame(v, corner_radius=20, fg_color="#f5f5f5")
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # --- COLUMNA IZQUIERDA: BUSCADOR Y LISTA DE PERSONAL ---
    col_izq = ctk.CTkFrame(main_frame, fg_color="transparent")
    col_izq.pack(side="left", fill="both", expand=True, padx=20, pady=20)

    ctk.CTkLabel(col_izq, text="Selección de Empleados", font=("Segoe UI", 16, "bold")).pack(pady=(0, 10))
    
    entry_busq = ctk.CTkEntry(col_izq, placeholder_text="Buscar empleado...", width=300)
    entry_busq.pack(pady=5)

    lista_frame = ctk.CTkScrollableFrame(col_izq, fg_color="white", height=400)
    lista_frame.pack(fill="both", expand=True, pady=10)

    df_pers = leer_datos('Personal')
    lista = sorted([f"{row['Nombre']} {row['Apellido']}" for _, row in df_pers.iterrows()])
    check_vars = {n: ctk.BooleanVar() for n in lista}

    def actualizar(e=None):
        for w in lista_frame.winfo_children(): w.destroy()
        filtro = entry_busq.get().lower()
        for n in [x for x in lista if filtro in x.lower()]:
            ctk.CTkCheckBox(lista_frame, text=n, variable=check_vars[n]).pack(anchor="w", pady=5, padx=10)

    entry_busq.bind("<KeyRelease>", actualizar)
    actualizar()

    # --- COLUMNA DERECHA: CALENDARIO Y DATOS ---
    col_der = ctk.CTkFrame(main_frame, fg_color="transparent")
    col_der.pack(side="right", fill="y", padx=30, pady=20)

    ctk.CTkLabel(col_der, text="Seleccione Fecha:", font=("Segoe UI", 16, "bold")).pack(pady=(0, 10))

    # Contenedor del calendario (tkcalendar)
    frame_cal = ctk.CTkFrame(col_der, fg_color="white", corner_radius=10)
    frame_cal.pack(pady=10)
    
    cal = Calendar(frame_cal, selectmode='day', date_pattern='yyyy-mm-dd', 
                   font=("Segoe UI", 11), background="#3498db", 
                   headersbackground="#2980b9", normalbackground="white", 
                   cursor="hand2", borderwidth=0, showothermonthdays=False, locale='es_ES')
    cal.pack(padx=5, pady=5)

    ctk.CTkLabel(col_der, text="Detalles del Bono", font=("Segoe UI", 16, "bold")).pack(pady=(20, 10))
    
    combo_tipo = ctk.CTkComboBox(col_der, values=["Hora extra", "Nocturna", "Bono"], state="readonly", width=250)
    combo_tipo.set("Bono")
    combo_tipo.pack(pady=5)


    ent_monto = ctk.CTkEntry(col_der, placeholder_text="Monto", width=250)
    ent_monto.pack(pady=5)

    def guardar_bono():
        sel = [n for n, var in check_vars.items() if var.get()]
        if not sel:
            messagebox.showwarning("Atención", "Seleccione al menos una persona.")
            return
        try:
            monto = float(ent_monto.get().strip())
        except:
            messagebox.showerror("Error", "Ingrese un monto válido.")
            return
        
        df_bonos = leer_datos('Bonos')
        nuevos = [{'Fecha': pd.to_datetime(cal.get_date()), 'Nombre': n.split()[0], 'Apellido': " ".join(n.split()[1:]), 'Tipo': combo_tipo.get(), 'Monto': monto} for n in sel]
        
        guardar_archivo(leer_datos('Personal'), leer_datos('Asistencia'), leer_datos('Pagos'), pd.concat([df_bonos, pd.DataFrame(nuevos)], ignore_index=True))
        messagebox.showinfo("Éxito", "Bonos registrados correctamente.")
        v.destroy()

    ctk.CTkButton(col_der, text="GUARDAR BONO", fg_color="#3498db", hover_color="#2980b9", 
                  font=("Segoe UI", 14, "bold"), height=40, width=250, command=guardar_bono).pack(pady=30)
         

        
def ventana_nomina():
    v = ctk.CTkToplevel()
    v.title("Gestión de Nómina")
    ancho, alto = 850, 600
    
    centrar_ventana(v, ancho, alto)
    v.configure(bg="white")
    v.attributes("-topmost", True)
    v.after(100, lambda: v.attributes("-topmost", False))
    main_frame = ctk.CTkFrame(v, corner_radius=20, fg_color="#f5f5f5")
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    ctk.CTkLabel(main_frame, text="CÁLCULO DE NÓMINA", font=("Segoe UI", 20, "bold"), text_color="#333333").pack(pady=10)

    controls = ctk.CTkFrame(main_frame, fg_color="transparent")
    controls.pack(pady=5)

    entry_busq = ctk.CTkEntry(controls, placeholder_text="Buscar empleado...", width=250)
    entry_busq.pack(side="left", padx=5)
    
    combo_mes = ctk.CTkComboBox(controls, values=MESES , width=150, state="readonly")
    combo_mes.set(MESES[datetime.now().month - 1])
    combo_mes.pack(side="left", padx=5)
    
    combo_anio = ctk.CTkComboBox(controls, values=[str(i) for i in range(2025, 2036)], width=100, state="readonly")
    combo_anio.set(str(datetime.now().year))
    combo_anio.pack(side="left", padx=5)

    table_container = ctk.CTkScrollableFrame(main_frame, fg_color="white", height=320, corner_radius=10)
    table_container.pack(fill="x", padx=20, pady=15)

    header_frame = ctk.CTkFrame(table_container, fg_color="#3498db", corner_radius=5)
    header_frame.pack(fill="x", pady=(0, 5))
    for col in ('Personal', 'Días', 'Base', 'Bonos', 'Total'):
        ctk.CTkLabel(header_frame, text=col, font=("Segoe UI", 12, "bold"), text_color="white", width=150).pack(side="left")

    empleado_seleccionado = {"nombre": None, "apellido": None}
    datos_calculados = []

    def seleccionar_fila(row, nombre, apellido):
        empleado_seleccionado["nombre"] = nombre
        empleado_seleccionado["apellido"] = apellido
        for child in table_container.winfo_children():
            if isinstance(child, ctk.CTkFrame) and child != header_frame:
                child.configure(fg_color="transparent")
        row.configure(fg_color="#d1e7dd")

    def renderizar_tabla(e=None):
        for widget in table_container.winfo_children():
            if widget != header_frame: widget.destroy()
        
        filtro = entry_busq.get().lower()
        for emp in datos_calculados:
            if filtro in f"{emp['nombre']} {emp['apellido']}".lower():
                row = ctk.CTkFrame(table_container, fg_color="transparent", cursor="hand2")
                row.pack(fill="x", pady=2)
                row.bind("<Button-1>", lambda e, r=row, n=emp['nombre'], a=emp['apellido']: seleccionar_fila(r, n, a))
                
                datos = [f"{emp['nombre']} {emp['apellido']}", str(emp['dias']), f"{emp['base']:,.2f}", f"{emp['bonos']:,.2f}", f"{emp['total']:,.2f}"]
                for d in datos:
                    lbl = ctk.CTkLabel(row, text=d, width=150)
                    lbl.pack(side="left")
                    lbl.bind("<Button-1>", lambda e, r=row, n=emp['nombre'], a=emp['apellido']: seleccionar_fila(r, n, a))

    def calcular(e=None):
        datos_calculados.clear()
        df_a = leer_datos('Asistencia')
        df_p = leer_datos('Pagos')
        df_b = leer_datos('Bonos')
        
        anio_sel = int(combo_anio.get())
        
        # 1. Filtramos datos por año
        df_a_year = df_a[df_a['Fecha'].dt.year == anio_sel]
        df_b_year = df_b[df_b['Fecha'].dt.year == anio_sel]
        
        # 2. Definimos si es Anual o Mes específico
        if combo_mes.get() == "13 - Anual":
            df_a_mes = df_a_year
            df_b_mes = df_b_year
        else:
            mes_sel = MESES.index(combo_mes.get()) + 1
            df_a_mes = df_a_year[df_a_year['Fecha'].dt.month == mes_sel]
            df_b_mes = df_b_year[df_b_year['Fecha'].dt.month == mes_sel]
            
        # 3. Agrupamos asistencias
        agrupado = df_a_mes.groupby(['Nombre', 'Apellido']).size().reset_index(name='Dias')
        
        # 4. Obtenemos lista única de empleados que trabajaron O tienen bonos
        empleados_asistencia = set(zip(agrupado['Nombre'], agrupado['Apellido']))
        empleados_bonos = set(zip(df_b_mes['Nombre'], df_b_mes['Apellido']))
        empleados = sorted(empleados_asistencia | empleados_bonos)
        
        for n, a in empleados:
            # Días trabajados
            f_asis = agrupado[(agrupado['Nombre'] == n) & (agrupado['Apellido'] == a)]
            d = int(f_asis['Dias'].iloc[0]) if not f_asis.empty else 0
            
            # Sueldo base
            f_pago = df_p[(df_p['Nombre'] == n) & (df_p['Apellido'] == a)]
            diario = float(f_pago['SueldoDiario'].iloc[0]) if not f_pago.empty else 0.0
            base = d * diario
            
            # Bonos
            bonos = df_b_mes[(df_b_mes['Nombre'] == n) & (df_b_mes['Apellido'] == a)]['Monto'].sum()
            
            datos_calculados.append({"nombre": n, "apellido": a, "dias": d, "base": base, "bonos": bonos, "total": base + bonos})
        
        renderizar_tabla()
        

    def abrir_previsualizacion():
        # 1. Validación de selección
        if not empleado_seleccionado["nombre"]:
            messagebox.showwarning("Atención", "Seleccione un empleado.")
            return
        
        # 2. Búsqueda segura: Refrescamos la búsqueda sobre la lista actual de datos_calculados
        emp = next((i for i in datos_calculados 
                    if i["nombre"] == empleado_seleccionado["nombre"] 
                    and i["apellido"] == empleado_seleccionado["apellido"]), None)
        
        # 3. Validación de existencia en el mes seleccionado
        if not emp:
            messagebox.showinfo("Información", f"El empleado {empleado_seleccionado['nombre']} no tiene registros en el periodo seleccionado.")
            return
        
        # 4. Creación de la ventana (todo igual que antes)
        v_prev = ctk.CTkToplevel()
        v_prev.title("Reporte de Pago")
        ancho, alto = 360, 310
        centrar_ventana(v_prev, ancho, alto)
        v_prev.attributes("-topmost", True)
        v_prev.after(100, lambda: v_prev.attributes("-topmost", False))
        
        bg_frame = ctk.CTkFrame(v_prev, fg_color="#f0f0f0", corner_radius=0)
        bg_frame.pack(fill="both", expand=True)
        
        card = ctk.CTkFrame(bg_frame, fg_color="white", corner_radius=15)
        card.pack(fill="both", expand=True, padx=20, pady=20)
        
        ctk.CTkLabel(card, text=f"Reporte: {emp['nombre'].upper()} {emp['apellido'].upper()}", 
                     font=("Arial", 14, "bold"), text_color="black").pack(fill="x", pady=(20, 15))
        
        datos = [
            ("Mes:", combo_mes.get()),
            ("Días trabajados:", str(emp['dias'])),
            ("Pago base:", f"{emp['base']:,.2f} EUR"),
            ("Bonos y extras:", f"{emp['bonos']:,.2f} EUR"),
            ("Total a pagar:", f"{emp['total']:,.2f} EUR")
        ]
        
        for lbl, val in datos:
            fila = ctk.CTkFrame(card, fg_color="transparent")
            fila.pack(fill="x", pady=2, padx=40)
            ctk.CTkLabel(fila, text=lbl, font=("Arial", 12), text_color="black").pack(side="left")
            ctk.CTkLabel(fila, text=f" {val}", font=("Arial", 12, "bold"), text_color="black").pack(side="left")
            
    # Botones
    btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
    btn_frame.pack(pady=10)
    
    ctk.CTkButton(btn_frame, text="Resumen", command=abrir_previsualizacion, width=120, fg_color="#27ae60").pack(side="left", padx=5)
    def manejar_exportacion():
        if not empleado_seleccionado["nombre"]:
            messagebox.showwarning("Atención", "Seleccione un empleado.")
            return

        mes_seleccionado = combo_mes.get()
        # MESES es la lista global que ya tiene "13 - Anual" al final
        idx_mes = MESES.index(mes_seleccionado)
        
        generar_pdf_personal(
            empleado_seleccionado["nombre"], 
            empleado_seleccionado["apellido"], 
            idx_mes
        )

    # --- BOTÓN DE EXPORTAR PDF ---
    ctk.CTkButton(
        btn_frame, 
        text="Exportar PDF", 
        command=manejar_exportacion, 
        width=120, 
        fg_color="#3498db"
    ).pack(side="left", padx=5)
    
    entry_busq.bind("<KeyRelease>", renderizar_tabla)
    combo_mes.configure(command=calcular)
    combo_anio.configure(command=calcular)
    calcular()
    
def ventana_asistencia():
    
    v = ctk.CTkToplevel()
    v.title("Registro de Asistencia")
    ancho, alto = 1000, 600
    
    # Llamas a la función de centrado ANTES de configurar otros atributos
    centrar_ventana(v, ancho, alto)
    v.attributes("-topmost", True)
    v.after(100, lambda: v.attributes("-topmost", False))
    

    # Contenedor principal que divide la pantalla en 2 columnas
    main_container = ctk.CTkFrame(v, fg_color="transparent")
    main_container.pack(fill="both", expand=True, padx=20, pady=20)
    
    # --- SECCIÓN IZQUIERDA: Calendario ---
    left_frame = ctk.CTkFrame(main_container, width=350)
    left_frame.pack(side="left", fill="y", padx=(0, 20))
    
    ctk.CTkLabel(left_frame, text="Seleccionar Fecha", font=("Segoe UI", 30, "bold")).pack(pady=10)
    cal = Calendar(left_frame, selectmode='day', date_pattern='yyyy-mm-dd', 
                   background="#3498db",headersbackground="#2980b9", foreground="white", borderwidth=0, showothermonthdays=False, locale='es_ES')
    cal.pack(pady=40, padx=20, expand=True)

    # --- SECCIÓN DERECHA: Lista y Controles ---
    right_frame = ctk.CTkFrame(main_container)
    right_frame.pack(side="right", fill="both", expand=True)
    
    ctk.CTkLabel(right_frame, text="Gestión de Asistencia", font=("Segoe UI", 16, "bold")).pack(pady=10)
    
    # Barra de búsqueda
    entry_busq = ctk.CTkEntry(right_frame, placeholder_text="Buscar empleado...", width=400)
    entry_busq.pack(pady=10, padx=20)
    
    # Checkbox "Seleccionar todos"
    check_todos = ctk.BooleanVar()
    df_pers = leer_datos('Personal')
    lista = sorted([f"{row['Nombre']} {row['Apellido']}" for _, row in df_pers.iterrows()])
    check_vars = {n: ctk.BooleanVar() for n in lista}
    
    def toggle_todos():
        for var in check_vars.values(): var.set(check_todos.get())
    
    ctk.CTkCheckBox(right_frame, text="Seleccionar todos", variable=check_todos, command=toggle_todos).pack(pady=5)
    
    # Lista con Scroll
    scroll_frame = ctk.CTkScrollableFrame(right_frame, height=300)
    scroll_frame.pack(pady=10, padx=20, fill="both", expand=True)
    
    def actualizar(e=None):
        for widget in scroll_frame.winfo_children(): widget.destroy()
        busqueda = entry_busq.get().lower()
        for n in lista:
            if busqueda in n.lower():
                ctk.CTkCheckBox(scroll_frame, text=n, variable=check_vars[n]).pack(pady=5, anchor="w")
                
    entry_busq.bind("<KeyRelease>", actualizar)
    actualizar()

    # --- Lógica de Registro (Botón al final de la derecha) ---
    def registrar():
        df_a = leer_datos('Asistencia')
        fecha_sel = pd.to_datetime(cal.get_date())
        sel = [n for n, var in check_vars.items() if var.get()]
        
        if not sel:
            messagebox.showwarning("Atención", "Seleccione al menos una persona.")
            return
        
        # Validar duplicados
        asistencias_hoy = df_a[df_a['Fecha'] == fecha_sel]
        duplicados = [n for n in sel if not asistencias_hoy[(asistencias_hoy['Nombre'] == n.split()[0]) & (asistencias_hoy['Apellido'] == n.split()[1])].empty]
        
        if duplicados:
            messagebox.showerror("Error", f"Ya tienen asistencia registrada:\n{', '.join(duplicados)}")
            return

        nuevas = [{'Fecha': fecha_sel, 'Nombre': n.split()[0], 'Apellido': n.split()[1], 'Proyecto': 'General'} for n in sel]
        guardar_archivo(leer_datos('Personal'), pd.concat([df_a, pd.DataFrame(nuevas)], ignore_index=True), leer_datos('Pagos'), leer_datos('Bonos'))
        messagebox.showinfo("Éxito", "Asistencia registrada correctamente")
        v.destroy()

    ctk.CTkButton(
        right_frame, 
        text="GUARDAR REGISTRO", 
        width=250,              # Ancho fijo en lugar de fill="x"
        height=45,              # Altura cómoda
        fg_color="#3498db", 
        hover_color="#2980b9",
        font=("Segoe UI", 14, "bold"), 
        command=registrar
    ).pack(pady=20)
    

def ventana_ver_asistencia():
    v = ctk.CTkToplevel()
    v.title("Reporte de Asistencia")
    ancho, alto = 1000, 600
    
    centrar_ventana(v, ancho, alto)
    v.attributes("-topmost", True)
    v.after(100, lambda: v.attributes("-topmost", False))

    main_container = ctk.CTkFrame(v, fg_color="transparent")
    main_container.pack(fill="both", expand=True, padx=20, pady=20)

    # --- SECCIÓN IZQUIERDA ---
    left_frame = ctk.CTkFrame(main_container, width=350)
    left_frame.pack(side="left", fill="y", padx=(0, 20))

    ctk.CTkLabel(left_frame, text="Seleccionar Fecha", font=("Segoe UI", 30, "bold")).pack(pady=10)
    cal = Calendar(left_frame, selectmode='day', date_pattern='yyyy-mm-dd', 
                   background="#3498db", headersbackground="#2980b9", foreground="white", 
                   borderwidth=0, showothermonthdays=False, locale='es_ES')
    cal.pack(pady=40, padx=20, expand=True)
    
    lbl_cantidad = ctk.CTkLabel(left_frame, text="Cantidad: 0", font=("Segoe UI", 20, "bold"), text_color="black")
    lbl_cantidad.pack(pady=(0, 20))

    # --- SECCIÓN DERECHA ---
    right_frame = ctk.CTkFrame(main_container, fg_color="transparent")
    right_frame.pack(side="right", fill="both", expand=True)

    ctk.CTkLabel(right_frame, text="Detalle de Asistencias", font=("Segoe UI", 16, "bold")).pack(pady=10)

    container_tabla = ctk.CTkFrame(right_frame, fg_color="#dbdbdb", corner_radius=20)
    container_tabla.pack(pady=10, padx=20, fill="both", expand=True)

    scroll_frame = ctk.CTkScrollableFrame(container_tabla, fg_color="transparent")
    scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)

    def consultar(e=None):
        for widget in scroll_frame.winfo_children(): widget.destroy()
        df = leer_datos('Asistencia')
        res = df[df['Fecha'] == pd.to_datetime(cal.get_date())].sort_values(by=['Nombre', 'Apellido'])
        
        for _, row in res.iterrows():
            nombre_completo = f"{row['Nombre']} {row['Apellido']}"
            
            # 1. Fila interactiva
            fila = ctk.CTkFrame(scroll_frame, fg_color="transparent", height=40, corner_radius=0)
            fila.pack(fill="x", pady=1)
            
            # 2. Etiqueta (usamos place o pack en un sub-frame)
            lbl = ctk.CTkLabel(fila, text=nombre_completo, text_color="#333333", anchor="w")
            lbl.pack(fill="x", padx=20, pady=10)
            
            # 3. Funciones de sombreado
            def on_enter(e, f=fila): f.configure(fg_color="#c0c0c0")
            def on_leave(e, f=fila): f.configure(fg_color="transparent")
            def ejecutar(e=None, n=row['Nombre'], a=row['Apellido']):
                # Aquí puedes abrir el detalle o hacer lo que necesites
                pass
            
            # 4. Asignar eventos
            for w in [fila, lbl]:
                w.bind("<Enter>", on_enter)
                w.bind("<Leave>", on_leave)
                w.bind("<Button-1>", ejecutar)
                w.configure(cursor="hand2")
                
        lbl_cantidad.configure(text=f"Cantidad: {len(res)}")

    cal.bind("<<CalendarSelected>>", lambda e: consultar())
    
    ctk.CTkButton(
        right_frame, 
        text="VER TOTALES MENSUALES", 
        width=250, height=45, 
        fg_color="#3498db", hover_color="#2980b9",
        font=("Segoe UI", 14, "bold"), 
        command=ventana_totales).pack(pady=20)

    consultar()



def ventana_totales():
    # 1. Ventana estándar
    v = ctk.CTkToplevel()
    v.title("Reporte de Totales")
    ancho, alto = 600, 600
    centrar_ventana(v, ancho, alto)
    v.attributes("-topmost", True)
    v.after(100, lambda: v.attributes("-topmost", False))

    main_container = ctk.CTkFrame(v, fg_color="transparent")
    main_container.pack(fill="both", expand=True, padx=20, pady=20)

    # 2. Encabezado
    ctk.CTkLabel(main_container, text="Totales Trabajados", font=("Segoe UI", 22, "bold")).pack(pady=10)

    # 3. CREACIÓN DEL CONTENEDOR ANTES DE USARLO
    controls_frame = ctk.CTkFrame(main_container, fg_color="transparent")
    controls_frame.pack(pady=5)

    # 4. Filtros (Usando la lista filtrada para excluir "Anual")
    fecha_actual = datetime.now()
    lista_opciones = ["Todos"] + [m for m in MESES if "Anual" not in m]
    
    combo_mes = ctk.CTkComboBox(controls_frame, values=lista_opciones, width=140)
    # Seleccionamos el mes actual si existe en la lista filtrada
    mes_actual_str = MESES[fecha_actual.month - 1]
    combo_mes.set(mes_actual_str if mes_actual_str in lista_opciones else lista_opciones[0])
    combo_mes.pack(side="left", padx=5)

    combo_anio = ctk.CTkComboBox(controls_frame, values=[str(i) for i in range(2025, 2036)], width=90)
    combo_anio.set(str(fecha_actual.year))
    combo_anio.pack(side="left", padx=5)

    ent_busq = ctk.CTkEntry(main_container, placeholder_text="Buscar por nombre...", width=350)
    ent_busq.pack(pady=10)
    

    # Tabla
    container_tabla = ctk.CTkFrame(main_container, fg_color="#dbdbdb", corner_radius=15, width=400)
    container_tabla.pack(fill="y", expand=True, pady=10)

    header_tabla = ctk.CTkFrame(container_tabla, fg_color="transparent")
    header_tabla.pack(fill="x", padx=30, pady=(15, 5))
    ctk.CTkLabel(header_tabla, text="Nombre", font=("Segoe UI", 12, "bold"), text_color="#333333").pack(side="left")
    ctk.CTkLabel(header_tabla, text="Días", font=("Segoe UI", 12, "bold"), text_color="#333333").pack(side="right")
    ctk.CTkFrame(container_tabla, height=2, fg_color="#bfbfbf").pack(fill="x", padx=20)

    scroll_frame = ctk.CTkScrollableFrame(container_tabla, fg_color="transparent", width=350)
    scroll_frame.pack(fill="both", expand=True, padx=5, pady=5)

    def actualizar(e=None):
        for widget in scroll_frame.winfo_children(): widget.destroy()
        
        df = leer_datos('Asistencia')
        df['Fecha'] = pd.to_datetime(df['Fecha'])
        df = df[df['Fecha'].dt.year == int(combo_anio.get())]
        
        seleccion = combo_mes.get()
        if seleccion != "Todos":
            try:
                # Ajusta esto si tu lista MESES no tiene el formato "01 - Nombre"
                mes_num = MESES.index(seleccion) + 1
                df = df[df['Fecha'].dt.month == mes_num]
            except: pass
            
        busqueda = ent_busq.get().lower()
        contar = df.groupby(['Nombre', 'Apellido']).size().reset_index(name='D')
        
        for _, r in contar.iterrows():
            nombre_completo = f"{r['Nombre']} {r['Apellido']}"
            if busqueda in nombre_completo.lower():
                
                fila = ctk.CTkFrame(scroll_frame, fg_color="transparent", height=40, corner_radius=0)
                fila.pack(fill="x", pady=1)
                
                inner = ctk.CTkFrame(fila, fg_color="transparent")
                inner.pack(fill="both", expand=True, padx=20)
                
                lbl_n = ctk.CTkLabel(inner, text=nombre_completo, text_color="#333333", anchor="w")
                lbl_n.pack(side="left")
                
                lbl_d = ctk.CTkLabel(inner, text=str(r['D']), text_color="#333333", font=("Segoe UI", 12, "bold"), anchor="e")
                lbl_d.pack(side="right")
                
                # Función para manejar el evento de click
                def ejecutar(e=None, n=r['Nombre'], a=r['Apellido']):
                    # 1. Obtener los valores actuales de los filtros
                    mes_sel = combo_mes.get()
                    anio_sel = int(combo_anio.get())
                    
                    # 2. Determinar el mes (si es 'Todos', abrimos el mes actual)
                    if mes_sel == "Todos":
                        mes_num = datetime.now().month
                    else:
                        # Obtenemos el índice del mes seleccionado en la lista original MESES
                        # Esto asegura que el número sea correcto incluso si filtramos opciones
                        mes_num = MESES.index(mes_sel) + 1
                    
                    # 3. Llamar al calendario pasando los valores reales obtenidos
                    abrir_calendario_asistencia(n, a, mes=mes_num, anio=anio_sel)
                
                # Eventos
                def crear_handler(f):
                    def on_enter(e): f.configure(fg_color="#c0c0c0")
                    def on_leave(e): f.configure(fg_color="transparent")
                    return on_enter, on_leave

                h_enter, h_leave = crear_handler(fila)
                
                for w in [fila, inner, lbl_n, lbl_d]:
                    w.bind("<Enter>", h_enter)
                    w.bind("<Leave>", h_leave)
                    w.bind("<Double-Button-1>", ejecutar)
                    w.configure(cursor="hand2")
                    
    combo_mes.configure(command=actualizar)
    combo_anio.configure(command=actualizar)
    ent_busq.bind("<KeyRelease>", actualizar)
    
    actualizar()


    



def ventana_agregar_personal():
    v = ctk.CTkToplevel(); v.title("Agregar Personal"); v.configure(bg="white")
    ancho, alto = 400, 350
    
    # Llamas a la función de centrado ANTES de configurar otros atributos
    centrar_ventana(v, ancho, alto)
    v.attributes("-topmost", True)
    v.after(100, lambda: v.attributes("-topmost", False))
    
    main_frame = ctk.CTkFrame(v, corner_radius=20, fg_color="#f5f5f5")
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)
    
    ctk.CTkLabel(main_frame, text="Registrar Empleado", font=("Segoe UI", 20, "bold"), text_color="#333333").pack(pady=20)
    
    # 4. Inputs modernos
    ent_n = ctk.CTkEntry(main_frame, placeholder_text="Nombre", width=260, height=40)
    ent_n.pack(pady=5)
    
    ent_a = ctk.CTkEntry(main_frame, placeholder_text="Apellido", width=260, height=40)
    ent_a.pack(pady=5)
    
    ent_s = ctk.CTkEntry(main_frame, placeholder_text="Sueldo Diario", width=260, height=40)
    ent_s.pack(pady=5)
    
    # 5. Lógica de guardado
    def guardar():
        try:
            nombre = ent_n.get().strip()
            apellido = ent_a.get().strip()
            sueldo = float(ent_s.get().strip())
            
            if not nombre or not apellido:
                messagebox.showwarning("Atención", "Todos los campos son obligatorios")
                return
                
            df_p = leer_datos('Personal')
            df_s = leer_datos('Pagos')
            
            nuevo_p = pd.DataFrame({'Nombre': [nombre], 'Apellido': [apellido]})
            nuevo_s = pd.DataFrame({'Nombre': [nombre], 'Apellido': [apellido], 'SueldoDiario': [sueldo]})
            
            guardar_archivo(
                pd.concat([df_p, nuevo_p], ignore_index=True),
                leer_datos('Asistencia'),
                pd.concat([df_s, nuevo_s], ignore_index=True),
                leer_datos('Bonos')
            )
            messagebox.showinfo("Éxito", "Personal guardado correctamente")
            v.destroy()
        except ValueError:
            messagebox.showerror("Error", "El sueldo debe ser un número válido")
        except Exception as e:
            messagebox.showerror("Error", f"Error: {e}")

    # 6. Botón de acción moderno
    btn = ctk.CTkButton(main_frame, text="GUARDAR EMPLEADO", width=260, height=45, 
                        fg_color="#3498db", hover_color="#2980b9",
                        font=("Segoe UI", 14, "bold"), command=guardar)
    btn.pack(pady=30)
def iniciar_app():
    crear_estructura_carpetas()
    inicializar_excel()
    
    root = ctk.CTk()
    root.title("Menú Principal")
    root.resizable(False, False)
    ancho, alto = 450, 450
    # Llamamos a la función de centrado antes de mostrar la ventana
    centrar_ventana(root, ancho, alto)
    
    

    # Contenedor principal
    main_frame = ctk.CTkFrame(root, corner_radius=20, fg_color="#ffffff")
    main_frame.pack(fill="both", expand=True, padx=20, pady=20)

    # Encabezado
    ctk.CTkLabel(main_frame, text="RACON", font=("Segoe UI", 24, "bold"), text_color="#333333").pack(pady=20)
    
    # Estilo de botones
    btn_args = {"width": 280, "height": 45, "font": ("Segoe UI", 14, "bold"), "corner_radius": 10}

    # Botones
    ctk.CTkButton(main_frame, text="ASISTENCIA", command=ventana_asistencia, fg_color="#34495e", **btn_args).pack(pady=10)
    ctk.CTkButton(main_frame, text="VER ASISTENCIA", command=ventana_ver_asistencia, fg_color="#34495e", **btn_args).pack(pady=5)
    ctk.CTkButton(main_frame, text="AGREGAR PERSONAL", command=ventana_agregar_personal, fg_color="#34495e", **btn_args).pack(pady=5)
    
    # Botones destacados
    ctk.CTkButton(main_frame, text="BONOS / HORAS EXTRA", command=ventana_bonos, fg_color="#3498db", text_color="black", **btn_args).pack(pady=5)
    ctk.CTkButton(main_frame, text="NÓMINA", command=ventana_nomina, fg_color="#3498db", **btn_args).pack(pady=5)

    # Pie de página
    ctk.CTkLabel(main_frame, text="© 2026 Racon Software", font=("Segoe UI", 10), text_color="#999999").pack(side="bottom", pady=20)

    root.mainloop()

if __name__ == "__main__":
    iniciar_app()
