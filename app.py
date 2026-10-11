import streamlit as st
from docx import Document
import io
import requests
from datetime import datetime

# --- FUNCIONES DE OBTENCIÓN DE DATOS ---
@st.cache_data(ttl=3600) # Se actualiza cada 1 hora para no saturar la API
def obtener_uf_hoy():
    """Obtiene el valor de la UF al día de hoy mediante la API de mindicador.cl"""
    try:
        response = requests.get('https://mindicador.cl/api/uf', timeout=5)
        data = response.json()
        return float(data['serie'][0]['valor'])
    except:
        # Valor de respaldo en caso de que la API falle
        return 41000.0

# --- FUNCIONES MATEMÁTICAS ---
def calcular_hipoteca_con_gracia(monto_credito, tasa_anual, plazo_anios, meses_gracia):
    tasa_mensual = (tasa_anual / 100) / 12
    meses_totales = plazo_anios * 12
    
    if meses_gracia > 0:
        monto_credito = monto_credito * ((1 + tasa_mensual)**meses_gracia)
        
    meses_pago = meses_totales - meses_gracia
    
    if tasa_mensual == 0:
        return monto_credito / meses_pago
        
    cuota = monto_credito * (tasa_mensual * (1 + tasa_mensual)**meses_pago) / ((1 + tasa_mensual)**meses_pago - 1)
    return cuota

def estimar_gastos_operacionales(monto_credito_uf, estado_propiedad):
    # 0.2% solo si es Nuevo CON DFL-2. Si es Nuevo SIN DFL-2 o Usado, aplica 0.8%
    tasa_impuesto = 0.002 if estado_propiedad == "Nuevo (Con beneficios DFL-2)" else 0.008
    impuesto_mutuo = monto_credito_uf * tasa_impuesto
    
    gastos_fijos_banco_uf = 18 
    return impuesto_mutuo + gastos_fijos_banco_uf

# --- FUNCIÓN DEL DOCUMENTO WORD ---
def generar_prospecto_avanzado(proposito, ubicacion, distribucion, estado, financiamiento, 
                               banco, tipo_tasa, valor_uf, valor_clp, pie_clp, 
                               usar_vehiculo, cap_vehiculo, g_operacionales, meses_gracia, 
                               arriendo, dividendo, flujo, cap_rate, carga_actual):
    
    doc = Document()
    doc.add_heading(f"Prospecto Inmobiliario: {proposito}", 0)
    doc.add_paragraph(f"Análisis para propiedad {estado}, tipología {distribucion} en {ubicacion}.")
    
    # Sección 1: Estructura de Compra
    doc.add_heading("1. Estructura de Compra y Financiamiento", level=1)
    tabla1 = doc.add_table(rows=7, cols=2)
    tabla1.style = 'Table Grid'
    
    datos_op = [
        ("Valor de la Propiedad", f"{valor_uf} UF ($ {valor_clp:,.0f})"),
        ("Pie Requerido", f"$ {pie_clp:,.0f}"),
        ("Gastos Operacionales (Aprox)", f"$ {g_operacionales:,.0f}"),
        ("Banco Seleccionado", banco),
        ("Tipo de Tasa", tipo_tasa),
        ("Meses de Gracia", f"{meses_gracia} meses"),
        ("Beneficio/Ley Aplicado", financiamiento)
    ]
    for i, (param, val) in enumerate(datos_op):
        tabla1.rows[i].cells[0].text = param
        tabla1.rows[i].cells[1].text = val
        
    if usar_vehiculo:
        doc.add_paragraph(f"\nNota de Capital: Se inyectarán $ {cap_vehiculo:,.0f} por liquidación de un vehículo. " + 
                          ("Cubre el pie requerido." if cap_vehiculo >= pie_clp else f"Faltan $ {(pie_clp - cap_vehiculo):,.0f} para el pie."))

    # Sección 2: Flujo Mensual
    doc.add_heading("2. Análisis Financiero Mensual", level=1)
    doc.add_paragraph(f"Carga Financiera: El dividendo representa un {carga_actual:.1f}% de tus ingresos.")
    
    tabla2 = doc.add_table(rows=3 if proposito == "Inversión" else 2, cols=2)
    tabla2.style = 'Table Grid'
    
    if proposito == "Inversión":
        datos_flujo = [
            ("Arriendo Mensual Estimado", f"$ {arriendo:,.0f}"),
            ("Dividendo Mensual", f"$ {dividendo:,.0f}"),
            ("Flujo de Caja Mensual", f"$ {flujo:,.0f}")
        ]
        doc.add_paragraph(f"\nRentabilidad Bruta Anual (Cap Rate): {cap_rate:.2f}%")
    else:
        # Si es para vivir, no hay arriendo, es puro costo.
        datos_flujo = [
            ("Dividendo Mensual (Costo Vivienda)", f"$ {dividendo:,.0f}"),
            ("Gastos Operacionales Iniciales a cubrir", f"$ {g_operacionales:,.0f}")
        ]
        
    for i, (param, val) in enumerate(datos_flujo):
        tabla2.rows[i].cells[0].text = param
        tabla2.rows[i].cells[1].text = val

    return doc

# --- INTERFAZ STREAMLIT ---
st.set_page_config(page_title="Simulador Inmobiliario PRO", page_icon="🏢", layout="wide")
st.title("🏢 Simulador Inmobiliario - Nivel Bancario")

# OBTENER UF EN TIEMPO REAL
uf_hoy = obtener_uf_hoy()

# 1. PROPÓSITO DE LA COMPRA
proposito = st.radio("¿Cuál es el propósito principal de esta compra?", ["Inversión", "Vivienda para vivir (Residencia Principal)"], horizontal=True)
st.divider()

# 2. CARACTERÍSTICAS
st.subheader("1. Características del Activo")
c1, c2, c3 = st.columns(3)
with c1:
    ubicacion = st.text_input("Ubicación", "Metro Rodrigo de Araya, Santiago")
    distribucion = st.selectbox("Distribución", ["Estudio", "1D1B", "2D1B", "2D2B", "3D1B"])
with c2:
    estado_propiedad = st.radio("Estado de la Propiedad", [
        "Nuevo (Con beneficios DFL-2)", 
        "Nuevo (Sin DFL-2)", 
        "Usado (Trato Directo / Corredor)"
    ])
    # FOGAES o Bono Pie aplica a propiedades nuevas, sin importar si tienen DFL-2 o no
    if "Nuevo" in estado_propiedad:
        financiamiento = st.selectbox("Apoyo Estatal / Inmobiliaria", ["Ninguno", "FOGAES (10% Garantía)", "Bono Pie (10% - 15%)"])
    else:
        financiamiento = "No aplica (Usado)"
with c3:
    valor_uf_actual = st.number_input("Valor UF Actual (Actualizado Auto)", value=uf_hoy, step=100.0)
    valor_propiedad_uf = st.number_input("Valor Propiedad (UF)", value=2200, step=50)
    
    if proposito == "Inversión":
        arriendo_estimado = st.number_input("Arriendo Mensual (CLP)", value=350000, step=5000)
    else:
        arriendo_estimado = 0
        st.info("Al ser para vivienda, no se calcula arriendo ni Cap Rate.")

# BASE DE DATOS DE TASAS DE BANCOS
bancos_tasas = {
    "BancoEstado": 4.1,
    "Banco de Chile": 4.5,
    "Santander": 4.6,
    "BCI": 4.4,
    "Scotiabank": 4.3,
    "Itaú": 4.7,
    "Consorcio": 4.2,
    "Mutuaria / Otro": 4.5
}

# 3. CRÉDITO Y BANCO
st.subheader("2. Financiamiento y Banco")
c4, c5, c6 = st.columns(3)
with c4:
    banco_seleccionado = st.selectbox("Seleccionar Institución Financiera", list(bancos_tasas.keys()))
    porcentaje_pie = st.slider("Pie / Enganche (%)", 0, 30, 20)
    plazo_anios = st.selectbox("Plazo del Crédito", [15, 20, 25, 30, 35, 40], index=3)
with c5:
    tasa_sugerida = bancos_tasas[banco_seleccionado]
    tipo_tasa = st.selectbox("Tipo de Tasa", ["Fija", "Mixta", "Variable"])
    tasa_anual = st.number_input("Tasa de Interés (%)", value=tasa_sugerida, step=0.1, help="Tasa promedio. Puedes modificarla si tienes una pre-aprobación distinta.")
with c6:
    meses_gracia = st.selectbox("Meses de Gracia", [0, 1, 2, 3, 4, 5, 6])
    usar_vehiculo = st.checkbox("¿Integrar venta de vehículo personal para el pie?", value=False)
    if usar_vehiculo:
        cap_vehiculo = st.number_input("Valor de venta estimado del vehículo (CLP)", value=15000000, step=500000)
    else:
        cap_vehiculo = 0

st.subheader("3. Perfil de Riesgo")
ingreso_actual = st.number_input("Ingreso Líquido Actual (CLP)", value=1500000, step=100000)

if st.button("Ejecutar Simulación Avanzada", type="primary"):
    valor_propiedad_clp = valor_propiedad_uf * valor_uf_actual
    pie_clp = valor_propiedad_clp * (porcentaje_pie / 100)
    monto_credito_uf_base = valor_propiedad_uf * (1 - (porcentaje_pie / 100))
    
    gastos_uf = estimar_gastos_operacionales(monto_credito_uf_base, estado_propiedad)
    gastos_clp = gastos_uf * valor_uf_actual
    
    dividendo_uf = calcular_hipoteca_con_gracia(monto_credito_uf_base, tasa_anual, plazo_anios, meses_gracia)
    dividendo_clp = dividendo_uf * valor_uf_actual
    
    carga_actual = (dividendo_clp / ingreso_actual) * 100 if ingreso_actual > 0 else 0
    
    st.divider()
    st.markdown("### Resultados de la Simulación")
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Pie + Gastos Operacionales", f"$ {(pie_clp + gastos_clp):,.0f}")
    r2.metric("Dividendo (Post-Gracia)", f"$ {dividendo_clp:,.0f}")
    
    if proposito == "Inversión":
        flujo_caja = arriendo_estimado - dividendo_clp
        cap_rate = (arriendo_estimado * 12) / valor_propiedad_clp * 100
        r3.metric("Flujo de Caja", f"$ {flujo_caja:,.0f}", delta="Positivo" if flujo_caja > 0 else "Negativo")
    else:
        flujo_caja = 0
        cap_rate = 0
        r3.metric("Propósito", "Para Vivir", delta="Costo Mensual", delta_color="off")
        
    r4.metric("Carga Financiera", f"{carga_actual:.1f}%", delta="Apto" if carga_actual <= 25 else "Riesgo Banco", delta_color="inverse" if carga_actual > 25 else "normal")

    doc = generar_prospecto_avanzado(
        proposito, ubicacion, distribucion, estado_propiedad, financiamiento, 
        banco_seleccionado, tipo_tasa, valor_propiedad_uf, valor_propiedad_clp, 
        pie_clp, usar_vehiculo, cap_vehiculo, gastos_clp, meses_gracia, 
        arriendo_estimado, dividendo_clp, flujo_caja, cap_rate, carga_actual
    )
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.download_button(
        label="📥 Descargar Prospecto Profesional (.docx)", 
        data=buffer, 
        file_name=f"Prospecto_{banco_seleccionado}.docx", 
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )
