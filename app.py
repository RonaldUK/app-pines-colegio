"""
Sistema Web para Entrega y Gestión de Pines Escolares
Desarrollado en Python con Streamlit y OpenPyXL
Trabaja directamente con el archivo Excel local: Entrega_Pines_Colegio_BP.xlsx
"""

import streamlit as st
import pandas as pd
import openpyxl
from datetime import datetime
import os

st.set_page_config(
    page_title="Sistema de Pines Escolares",
    page_icon="🎖️",
    layout="wide",
    initial_sidebar_state="expanded"
)

EXCEL_PATH = "Entrega_Pines_Colegio_BP.xlsx"

# ----------------- FUNCIONES DE ACCESO AL EXCEL -----------------
def get_workbook():
    if not os.path.exists(EXCEL_PATH):
        wb = openpyxl.Workbook()
        ws_est = wb.active
        ws_est.title = "Estudiantes"
        ws_est.append(["Nro", "DNI", "Estudiante", "Grado y Sección"])
        
        ws_ent = wb.create_sheet("ENTREGA_PIN")
        ws_ent.append(["Nro", "Fecha y Hora", "DNI", "Estudiante", "Grado y Sección", "Nombre del Pin", "Observación"])
        
        ws_pin = wb.create_sheet("PINES")
        ws_pin.append(["Nro", "Nombre del Pin", "Categoría", "Descripción"])
        pines_base = [
            ("01", "Plan Lector", "Lectura", "Comprensión lectora y lectura escolar"),
            ("02", "Crucimatic", "Matemáticas", "Habilidad en razonamiento matemático"),
            ("03", "Vóley", "Deporte", "Destacada participación deportiva"),
            ("04", "Fútbol", "Deporte", "Destacada participación deportiva"),
            ("05", "Arte y Música", "Artes", "Creatividad artística y musical"),
            ("06", "Ranking", "Rendimiento", "Puesto sobresaliente en cuadro de honor"),
            ("07", "Los Que Más Saben", "Conocimiento", "Conocimientos generales"),
            ("08", "Religión", "Valores", "Vivencia de valores y respeto"),
            ("09", "Inglés", "Idiomas", "Dominio del idioma inglés")
        ]
        for p in pines_base:
            ws_pin.append(list(p))
            
        wb.create_sheet("REPORTES_RESUMEN")
        wb.save(EXCEL_PATH)
    return openpyxl.load_workbook(EXCEL_PATH)

def cargar_datos():
    wb = get_workbook()
    
    # 1. Estudiantes
    ws_est = wb["Estudiantes"]
    data_est = list(ws_est.iter_rows(values_only=True))
    if data_est and len(data_est) > 1:
        headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(data_est[0])]
        df_est = pd.DataFrame(data_est[1:], columns=headers).dropna(how='all')
        if 'DNI' in df_est.columns:
            df_est['DNI'] = df_est['DNI'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    else:
        df_est = pd.DataFrame(columns=["Nro", "DNI", "Estudiante", "Grado y Sección"])

    # 2. Pines
    ws_pin = wb["PINES"]
    data_pin = list(ws_pin.iter_rows(values_only=True))
    if data_pin and len(data_pin) > 1:
        headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(data_pin[0])]
        df_pin = pd.DataFrame(data_pin[1:], columns=headers).dropna(how='all')
    else:
        df_pin = pd.DataFrame(columns=["Nro", "Nombre del Pin", "Categoría", "Descripción"])

    # 3. Entregas
    ws_ent = wb["ENTREGA_PIN"]
    data_ent = list(ws_ent.iter_rows(values_only=True))
    if data_ent and len(data_ent) > 1:
        headers = [str(h).strip() if h else f"col_{i}" for i, h in enumerate(data_ent[0])]
        df_ent = pd.DataFrame(data_ent[1:], columns=headers).dropna(how='all')
        if 'DNI' in df_ent.columns:
            df_ent['DNI'] = df_ent['DNI'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
    else:
        df_ent = pd.DataFrame(columns=["Nro", "Fecha y Hora", "DNI", "Estudiante", "Grado y Sección", "Nombre del Pin", "Observación"])
        
    return df_est, df_pin, df_ent

# ----------------- VISTA LATERAL (SIDEBAR) -----------------
st.sidebar.title("🎖️ Gestión de Pines")
st.sidebar.markdown("---")
opcion = st.sidebar.radio(
    "Seleccione el Módulo:",
    ["🎁 Entrega de Pines", "👥 Gestión de Alumnos", "🏷️ Gestión de Pines", "📊 Reportes y Consultas"],
    index=0
)

# --- BOTÓN DE DESCARGA DE EXCEL ---
st.sidebar.markdown("---")
st.sidebar.subheader("📥 Exportar Datos")
try:
    with open(EXCEL_PATH, "rb") as file:
        st.sidebar.download_button(
            label="Descargar Excel Completo",
            data=file,
            file_name="Reporte_Pines_Colegio.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            help="Descarga la base de datos completa para tus informes."
        )
except FileNotFoundError:
    st.sidebar.warning("El archivo Excel aún no se ha generado.")


# Cargar data fresca
df_est, df_pin, df_ent = cargar_datos()

# ==========================================================
# 1. MODULO DE ENTREGA DE PINES
# ==========================================================
if opcion == "🎁 Entrega de Pines":
    st.header("Formulario de Entrega de Pines")
    st.caption("Selecciona un alumno de la tabla para ver sus pines previos y registrar uno nuevo.")
    
    col_izq, col_der = st.columns([1.3, 1])
    
    with col_izq:
        st.subheader("1. Seleccionar Estudiante")
        filtro_txt = st.text_input("🔍 Buscar por Nombre, DNI o Grado:", key="busq_entrega_pin", placeholder="Escriba para filtrar...")
        
        df_filtrada = df_est[['DNI', 'Estudiante', 'Grado y Sección']].copy()
        if filtro_txt:
            condicion = (
                df_filtrada['Estudiante'].astype(str).str.contains(filtro_txt, case=False, na=False) |
                df_filtrada['DNI'].astype(str).str.contains(filtro_txt, case=False, na=False) |
                df_filtrada['Grado y Sección'].astype(str).str.contains(filtro_txt, case=False, na=False)
            )
            df_filtrada = df_filtrada[condicion]
        
        df_filtrada = df_filtrada.reset_index(drop=True)
        
        # Tabla interactiva con selección de fila
        st.caption("👆 Haz clic en una fila para seleccionarlo:")
        evento_seleccion = st.dataframe(
            df_filtrada,
            use_container_width=True,
            hide_index=True,
            selection_mode="single-row",
            on_select="rerun",
            key="tabla_alumnos_entrega"
        )
        
        alumno_sel = None
        filas_seleccionadas = evento_seleccion.selection.get("rows", [])
        if filas_seleccionadas:
            idx = filas_seleccionadas[0]
            if idx < len(df_filtrada):
                alumno_sel = df_filtrada.iloc[idx]
        elif not df_filtrada.empty:
            # Si no ha hecho clic aún, tomar el primero por defecto o permitir elegir
            alumno_sel = df_filtrada.iloc[0]

    with col_der:
        if alumno_sel is not None:
            dni_actual = str(alumno_sel['DNI'])
            st.subheader("2. Pines Asignados Previamente")
            st.info(f"👤 **Alumno:** {alumno_sel['Estudiante']}\n\n📌 **DNI:** {dni_actual} | **Grado:** {alumno_sel['Grado y Sección']}")
            
            pines_previos = df_ent[df_ent['DNI'] == dni_actual]
            pines_ya_entregados_nombres = []
            
            if not pines_previos.empty:
                pines_ya_entregados_nombres = pines_previos['Nombre del Pin'].dropna().str.strip().tolist()
                for _, r in pines_previos.iterrows():
                    st.success(f"🏅 **{r['Nombre del Pin']}** — *Entregado:* {r.get('Fecha y Hora', 'N/D')}")
            else:
                st.write("✨ *Este estudiante aún no tiene ningún pin registrado.*")
                
            st.markdown("---")
            st.subheader("3. Asignar Nuevo Pin")
            
            todos_pines = df_pin['Nombre del Pin'].dropna().str.strip().tolist()
            pines_disponibles = [p for p in todos_pines if p not in pines_ya_entregados_nombres]
            
            if not pines_disponibles:
                st.warning("⚠️ ¡El estudiante ya cuenta con todos los pines disponibles!")
            else:
                with st.form("form_entrega_directo"):
                    pin_seleccionado = st.selectbox("Seleccione el Pin:", pines_disponibles)
                    observacion = st.text_input("Observación / Motivo (opcional):", placeholder="Ej: Participación destacada...")
                    btn_entregar = st.form_submit_button("✅ Entregar Pin")
                    
                    if btn_entregar:
                        ahora = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        wb = get_workbook()
                        ws_ent = wb["ENTREGA_PIN"]
                        nuevo_id = str(ws_ent.max_row).zfill(3)
                        
                        nueva_fila = [
                            nuevo_id,
                            ahora,
                            dni_actual,
                            alumno_sel['Estudiante'],
                            alumno_sel['Grado y Sección'],
                            pin_seleccionado,
                            observacion
                        ]
                        ws_ent.append(nueva_fila)
                        wb.save(EXCEL_PATH)
                        st.balloons()
                        st.success(f"¡Pin '{pin_seleccionado}' registrado con éxito!")
                        st.rerun()
        else:
            st.warning("No hay estudiantes para mostrar.")

# ==========================================================
# 2. MODULO GESTIÓN DE ALUMNOS (CRUD INTERACTIVO)
# ==========================================================
elif opcion == "👥 Gestión de Alumnos":
    st.header("Gestión de Estudiantes")
    
    col_izq_a, col_der_a = st.columns([1.3, 1])
    
    with col_izq_a:
        st.subheader("Catálogo de Estudiantes")
        filtro_a = st.text_input("🔍 Buscar Alumno por Nombre, DNI o Grado:", key="busq_crud_alumnos")
        
        df_mostrar_a = df_est[['DNI', 'Estudiante', 'Grado y Sección']].copy()
        if filtro_a:
            df_mostrar_a = df_mostrar_a[
                df_mostrar_a['Estudiante'].astype(str).str.contains(filtro_a, case=False, na=False) |
                df_mostrar_a['DNI'].astype(str).str.contains(filtro_a, case=False, na=False) |
                df_mostrar_a['Grado y Sección'].astype(str).str.contains(filtro_a, case=False, na=False)
            ]
        df_mostrar_a = df_mostrar_a.reset_index(drop=True)
        
        st.caption("👆 Haz clic en una fila para Seleccionarlo:")
        sel_tabla_a = st.dataframe(
            df_mostrar_a,
            use_container_width=True,
            hide_index=True,
            selection_mode="single-row",
            on_select="rerun",
            key="tabla_crud_alumnos"
        )
        
        alumno_sel_crud = None
        filas_a = sel_tabla_a.selection.get("rows", [])
        if filas_a and filas_a[0] < len(df_mostrar_a):
            alumno_sel_crud = df_mostrar_a.iloc[filas_a[0]]
            
    with col_der_a:
        accion = st.radio("Acción a realizar:", ["✏️ Editar Seleccionado", "🗑️ Eliminar Seleccionado", "➕ Registrar Nuevo Alumno"], horizontal=True)
        
        if accion == "✏️ Editar Seleccionado":
            st.subheader("Editar Alumno")
            if alumno_sel_crud is not None:
                st.info(f"Modificando al alumno DNI: **{alumno_sel_crud['DNI']}**")
                with st.form("form_edit_alumno_directo"):
                    nuevo_nombre = st.text_input("Nombres y Apellidos:", value=alumno_sel_crud['Estudiante'])
                    nuevo_grado = st.text_input("Grado y Sección:", value=alumno_sel_crud['Grado y Sección'])
                    btn_guardar_ed = st.form_submit_button("💾 Guardar Cambios")
                    
                    if btn_guardar_ed:
                        wb = get_workbook()
                        ws = wb["Estudiantes"]
                        for row in ws.iter_rows(min_row=2):
                            val_dni = str(row[1].value).replace(".0", "").strip()
                            if val_dni == str(alumno_sel_crud['DNI']):
                                row[2].value = nuevo_nombre.strip().upper()
                                row[3].value = nuevo_grado.strip().upper()
                                break
                        wb.save(EXCEL_PATH)
                        st.success("¡Alumno actualizado correctamente!")
                        st.rerun()
            else:
                st.info("👈 Seleccione un alumno en la tabla para editarlo.")
                
        elif accion == "🗑️ Eliminar Seleccionado":
            st.subheader("Eliminar Alumno")
            if alumno_sel_crud is not None:
                st.warning(f"¿Está seguro de eliminar a **{alumno_sel_crud['Estudiante']}** (DNI: {alumno_sel_crud['DNI']})?")
                if st.button("⚠️ Confirmar Eliminación", type="primary"):
                    wb = get_workbook()
                    ws = wb["Estudiantes"]
                    fila_borrar = None
                    for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
                        val_dni = str(row[1].value).replace(".0", "").strip()
                        if val_dni == str(alumno_sel_crud['DNI']):
                            fila_borrar = idx
                            break
                    if fila_borrar:
                        ws.delete_rows(fila_borrar, 1)
                        wb.save(EXCEL_PATH)
                        st.success("Alumno eliminado correctamente.")
                        st.rerun()
            else:
                st.info("👈 Seleccione un alumno en la tabla para eliminarlo.")
                
        elif accion == "➕ Registrar Nuevo Alumno":
            st.subheader("Registrar Nuevo Estudiante")
            with st.form("form_crear_alu_directo"):
                c_dni = st.text_input("DNI del Alumno:")
                c_nom = st.text_input("Nombres y Apellidos:")
                c_gra = st.text_input("Grado y Sección:", placeholder="Ej: SEGUNDO 'B' DE PRIMARIA")
                btn_crear_a = st.form_submit_button("➕ Registrar Estudiante")
                
                if btn_crear_a:
                    if not c_dni or not c_nom:
                        st.error("DNI y Nombres son requeridos.")
                    elif c_dni.strip() in df_est['DNI'].values:
                        st.error("Ya existe un alumno con ese DNI.")
                    else:
                        wb = get_workbook()
                        ws = wb["Estudiantes"]
                        nuevo_nro = ws.max_row
                        ws.append([nuevo_nro, c_dni.strip(), c_nom.strip().upper(), c_gra.strip().upper()])
                        wb.save(EXCEL_PATH)
                        st.success(f"Alumno {c_nom} registrado con éxito.")
                        st.rerun()

# ==========================================================
# 3. MODULO GESTIÓN DE PINES (CRUD INTERACTIVO)
# ==========================================================
elif opcion == "🏷️ Gestión de Pines":
    st.header("Gestión del Catálogo de Pines")
    
    col_izq_p, col_der_p = st.columns([1.3, 1])
    
    with col_izq_p:
        st.subheader("Catálogo Actual")
        filtro_p = st.text_input("🔍 Buscar Pin por Nombre o Categoría:", key="busq_crud_pines")
        
        df_mostrar_p = df_pin[['Nro', 'Nombre del Pin', 'Categoría', 'Descripción']].copy()
        if filtro_p:
            df_mostrar_p = df_mostrar_p[
                df_mostrar_p['Nombre del Pin'].astype(str).str.contains(filtro_p, case=False, na=False) |
                df_mostrar_p['Categoría'].astype(str).str.contains(filtro_p, case=False, na=False)
            ]
        df_mostrar_p = df_mostrar_p.reset_index(drop=True)
        
        st.caption("👆 Haz clic en una fila para Seleccionarlo:")
        sel_tabla_p = st.dataframe(
            df_mostrar_p,
            use_container_width=True,
            hide_index=True,
            selection_mode="single-row",
            on_select="rerun",
            key="tabla_crud_pines"
        )
        
        pin_sel_crud = None
        filas_p = sel_tabla_p.selection.get("rows", [])
        if filas_p and filas_p[0] < len(df_mostrar_p):
            pin_sel_crud = df_mostrar_p.iloc[filas_p[0]]
            
    with col_der_p:
        accion_p = st.radio("Acción a realizar:", ["✏️ Editar Seleccionado", "🗑️ Eliminar Seleccionado", "➕ Crear Nuevo Pin"], horizontal=True)
        
        if accion_p == "✏️ Editar Seleccionado":
            st.subheader("Editar Pin")
            if pin_sel_crud is not None:
                st.info(f"Modificando el Pin: **{pin_sel_crud['Nombre del Pin']}**")
                with st.form("form_edit_pin_directo"):
                    p_nom = st.text_input("Nombre del Pin:", value=pin_sel_crud['Nombre del Pin'])
                    p_cat = st.text_input("Categoría:", value=pin_sel_crud['Categoría'] if pd.notna(pin_sel_crud['Categoría']) else "")
                    p_desc = st.text_area("Descripción:", value=pin_sel_crud['Descripción'] if pd.notna(pin_sel_crud['Descripción']) else "")
                    btn_up_pin = st.form_submit_button("💾 Guardar Cambios")
                    
                    if btn_up_pin:
                        wb = get_workbook()
                        ws = wb["PINES"]
                        for row in ws.iter_rows(min_row=2):
                            if str(row[1].value).strip().lower() == str(pin_sel_crud['Nombre del Pin']).strip().lower():
                                row[1].value = p_nom.strip()
                                row[2].value = p_cat.strip()
                                row[3].value = p_desc.strip()
                                break
                        wb.save(EXCEL_PATH)
                        st.success("¡Pin actualizado correctamente!")
                        st.rerun()
            else:
                st.info("👈 Seleccione un pin en la tabla para editarlo.")
                
        elif accion_p == "🗑️ Eliminar Seleccionado":
            st.subheader("Eliminar Pin")
            if pin_sel_crud is not None:
                st.warning(f"¿Está seguro de eliminar el pin **{pin_sel_crud['Nombre del Pin']}**?")
                if st.button("⚠️ Confirmar Eliminación de Pin", type="primary"):
                    wb = get_workbook()
                    ws = wb["PINES"]
                    fila_p_borrar = None
                    for idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
                        if str(row[1].value).strip().lower() == str(pin_sel_crud['Nombre del Pin']).strip().lower():
                            fila_p_borrar = idx
                            break
                    if fila_p_borrar:
                        ws.delete_rows(fila_p_borrar, 1)
                        wb.save(EXCEL_PATH)
                        st.success("Pin eliminado del catálogo.")
                        st.rerun()
            else:
                st.info("👈 Seleccione un pin en la tabla para eliminarlo.")
                
        elif accion_p == "➕ Crear Nuevo Pin":
            st.subheader("Crear Nuevo Pin")
            with st.form("form_add_nuevo_pin"):
                np_nom = st.text_input("Nombre del Pin:")
                np_cat = st.text_input("Categoría:")
                np_desc = st.text_area("Descripción:")
                btn_crear_pin = st.form_submit_button("➕ Crear Pin")
                
                if btn_crear_pin:
                    if not np_nom:
                        st.error("El nombre del pin es obligatorio.")
                    elif np_nom.lower().strip() in df_pin['Nombre del Pin'].str.lower().str.strip().values:
                        st.error("Ya existe un pin con ese nombre.")
                    else:
                        wb = get_workbook()
                        ws = wb["PINES"]
                        nro_pin = str(ws.max_row).zfill(2)
                        ws.append([nro_pin, np_nom.strip(), np_cat.strip(), np_desc.strip()])
                        wb.save(EXCEL_PATH)
                        st.success(f"Pin '{np_nom}' registrado correctamente.")
                        st.rerun()

# ==========================================================
# 4. MODULO REPORTES Y CONSULTAS
# ==========================================================
elif opcion == "📊 Reportes y Consultas":
    st.header("Reportes y Estadísticas de Pines Entregados")
    
    col_m1, col_m2, col_m3 = st.columns(3)
    col_m1.metric("Total Estudiantes", len(df_est))
    col_m2.metric("Total Entregas Realizadas", len(df_ent))
    alumnos_con_pin = df_ent['DNI'].nunique() if not df_ent.empty else 0
    col_m3.metric("Estudiantes con al menos 1 Pin", alumnos_con_pin)
    
    st.markdown("---")
    
    st.subheader("🔍 Consulta Historial por Estudiante")
    busq_hist = st.text_input("Ingrese Nombre o DNI del Estudiante:", placeholder="Buscar...")
    
    if busq_hist:
        df_hist = df_ent[
            df_ent['Estudiante'].astype(str).str.contains(busq_hist, case=False, na=False) |
            df_ent['DNI'].astype(str).str.contains(busq_hist, case=False, na=False)
        ]
        if not df_hist.empty:
            st.dataframe(df_hist[['Fecha y Hora', 'DNI', 'Estudiante', 'Grado y Sección', 'Nombre del Pin', 'Observación']], use_container_width=True, hide_index=True)
        else:
            st.warning("No se encontraron entregas para este estudiante.")
            
    st.markdown("---")
    st.subheader("📈 Resumen por Grado y Sección")
    if not df_ent.empty and not df_est.empty:
        resumen_grados = df_ent.groupby('Grado y Sección').agg(
            Total_Pines_Entregados=('Nombre del Pin', 'count'),
            Alumnos_Beneficiados=('DNI', 'nunique')
        ).reset_index()
        
        st.dataframe(resumen_grados, use_container_width=True, hide_index=True)
    else:
        st.info("No hay datos suficientes para generar estadísticas.")
        
    st.markdown("---")
    st.subheader("📋 Registro Completo de Entregas (Hoja ENTREGA_PIN)")
    st.dataframe(df_ent, use_container_width=True, hide_index=True)
