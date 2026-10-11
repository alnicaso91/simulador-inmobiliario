import streamlit as st
from docx import Document
import io

def calcular_hipoteca_con_gracia(monto_credito, tasa_anual, plazo_anios, meses_gracia):
    """Calcula el dividendo mensual capitalizando los intereses si hay meses de gracia."""
    tasa_mensual = (tasa_anual / 100) / 12
    meses_totales = plazo_anios * 12
    
    # Si hay meses de gracia, los intereses se suman al capital inicial
    if meses_gracia > 0:
        monto_credito = monto_credito * ((1 + tasa_mensual)**meses_gracia)
        
    meses_pago = meses_totales - meses_gracia
    
    if tasa_mensual == 0:
        return monto_credito / meses_pago
        
    cuota = monto_credito * (tasa_mensual * (1 + tasa_mensual)**meses_pago) / ((1 + tasa_mensual)**meses_pago - 1)
    return cuota

def estimar_gastos_operacionales(monto_credito_uf, estado_propiedad):
    """Calcula un aproximado de gastos operacionales según beneficios tributarios."""
    # Impuesto al mutuo (0.8% estándar, 0.2% si es DFL-2 Nuevo)
    tasa_impuesto = 0.002 if "Nuevo" in estado_propiedad else 0.008
    impuesto_mutuo = monto_credito_uf * tasa_impuesto
    
    # Gastos fijos aproximados en UF (Tasación, Notaría, Conservador, Estudio Títulos)
    gastos_fijos_banco_uf = 18 
    return impuesto_mutuo + gastos_fijos_banco_uf

def generar_prospecto_avanzado(ubicacion, distribucion, estado, financiamiento, tipo_tasa, 
                               valor_uf, valor_clp, pie_clp, usar_vehiculo, cap_vehiculo, 
                               g_operacionales, meses_gracia, arriendo, dividendo, flujo, 
                               cap_rate, carga_actual):
    """Ensambla y da formato al documento de Word final."""
    doc = Document()
    doc.add_heading(f"Prospecto de Inversión Inmobiliaria Avanzado", 0)
    doc.add_paragraph(f"Análisis para propiedad {estado}, tipología {distribucion} en {ubicacion}.")
    
    # Sección 1: Estructura de Compra
    doc.add_heading("1. Estructura de Compra y Financiamiento", level=1)
    tabla1 = doc.add_table(rows=6, cols=2)
    tabla1.style = 'Table Grid'
    
    datos_op = [
        ("Valor de la Propiedad", f"{valor_uf} UF ($ {valor_clp:,.0f})"),
        ("Pie Requerido", f"$ {pie_clp:,.0f}"),
        ("Gastos Operacionales (Aprox)", f"$ {g_operacionales:,.0f}"),
        ("Tipo de Tasa", tipo_tasa),
        ("Meses de Gracia", f"{meses_gracia} meses"),
        ("Beneficio/Ley Aplicado", financiamiento)
    ]
    for i, (param, val) in enumerate(datos_op):
        tabla1.rows[i].cells[0].text = param
        tabla1.rows[i].cells[1].text = val
        
    if usar_vehiculo:
        doc.add_paragraph(f"\nNota de Capital: Se contempla la inyección de $ {cap_vehiculo:,.0f} provenientes de la liquidación de un vehículo personal. " + 
                          ("Este monto cubre el pie requerido en su totalidad." if cap_vehiculo >= pie_clp else f"Falta un saldo de $ {(pie_clp - cap_vehiculo):,.0f} para cubrir el pie."))
                          
    if "Mixta" in tipo_tasa or "Variable" in tipo_tasa:
        doc.add_paragraph("⚠️ Advertencia de Tasa: Al no ser fija, el dividendo podría sufrir variaciones significativas según indicadores macroeconómicos al finalizar el periodo fijo.")

    # Sección 2: Flujo y Riesgo
    doc.add_heading("2. Flujo de Caja y Riesgo Bancario", level=1)
    doc.add_paragraph(f"Carga Financiera Actual: El dividendo representa un {carga_actual:.1f}% de tus ingresos líquidos.")
    
    tabla2 = doc.add_table(rows=3, cols=2)
    tabla2.style = 'Table Grid'
    
    datos_flujo = [
        ("Arriendo Mensual Estimado", f"$ {arriendo:,.0f}"),
        ("Dividendo Mensual", f"$ {dividendo:,.0f}"),
        ("Flujo de Caja Mensual", f"$ {flujo:,.0f}")
    ]
    for i, (param, val) in enumerate(datos_flujo):
        tabla2.rows[i].cells[0].text = param
        tabla2.rows[i].cells[1].text = val

    return doc

# --- INTERFAZ DE USUARIO (STREAMLIT) ---
st.set_page_config(page_title="Simulador Inmobiliario PRO", page_icon="🏢", layout="wide")
st.title("🏢 Simulador Inmobiliario - Nivel Bancario")
st.write("Analiza rentabilidad, gastos operacionales, beneficios tributarios y viabilidad de financiamiento en segundos.")

# 1. PARÁMETROS DE PROPIEDAD
st.subheader("1. Características del Activo")
c1, c2, c3 = st.columns(3)
with c1:
    ubicacion = st.text_input("Ubicación", "Metro Rodrigo de Araya, Santiago")
    distribucion = st.selectbox("Distribución", ["Estudio", "1D1B", "2D1B", "2D2B", "3D1B", "3D2B"])
with c2:
    estado_propiedad = st.radio("Estado de la Propiedad", ["Usado (Trato Directo / Corredor)", "Nuevo (Con beneficios DFL-2)"])
    if "Nuevo" in estado_propiedad:
        financiamiento = st.selectbox("Apoyo Estatal / Inmobiliaria", ["Ninguno", "FOGAES (10% Garantía)", "Bono Pie (10% - 15%)"])
    else:
        financiamiento = "No aplica (Usado)"
with c3:
    valor_uf_actual = st.number_input("Valor UF Actual (CLP)", value=41000, step=100)
    valor_propiedad_uf = st.number_input("Valor Propiedad (UF)", value=2200, step=50)
    arriendo_estimado = st.number_input("Arriendo Mensual (CLP)", value=350000, step=5000)

# 2. ESTRUCTURA DE CRÉDITO
st.subheader("2. Estructura del Crédito y Gastos")
c4, c5, c6 = st.columns(3)
with c4:
    porcentaje_pie = st.slider("Pie / Enganche (%)", 0, 30, 20)
    plazo_anios = st.selectbox("Plazo del Crédito", [15, 20, 25, 30, 35, 40], index=3) # Por defecto 30 años
with c5:
    tipo_tasa = st.selectbox("Tipo de Tasa", ["Fija (Recomendada)", "Mixta (3-5 años fija, luego variable)", "Variable"])
    tasa_anual = st.number_input("Tasa de Interés Anual (%)", value=4.5, step=0.1)
with c6:
    meses_gracia = st.selectbox("Meses de Gracia", [0, 1, 2, 3, 4, 5, 6], help="Retrasa el pago inicial, pero suma intereses al capital.")
    usar_vehiculo = st.checkbox("¿Integrar venta de vehículo personal para levantar capital?", value=False)
    if usar_vehiculo:
        cap_vehiculo = st.number_input("Valor de venta estimado del vehículo (CLP)", value=15000000, step=500000)
    else:
        cap_vehiculo = 0

# 3. PERFIL FINANCIERO
st.subheader("3. Perfil de Riesgo")
ingreso_actual = st.number_input("Ingreso Líquido Actual (CLP)", value=1500000, step=100000)

# BOTÓN DE EJECUCIÓN
if st.button("Ejecutar Simulación Avanzada", type="primary"):
    # Cálculos Matemáticos
    valor_propiedad_clp = valor_propiedad_uf * valor_uf_actual
    pie_clp = valor_propiedad_clp * (porcentaje_pie / 100)
    monto_credito_uf_base = valor_propiedad_uf * (1 - (porcentaje_pie / 100))
    
    # Gastos Operacionales
    gastos_uf = estimar_gastos_operacionales(monto_credito_uf_base, estado_propiedad)
    gastos_clp = gastos_uf * valor_uf_actual
    
    # Dividendo ajustado por gracia
    dividendo_uf = calcular_hipoteca_con_gracia(monto_credito_uf_base, tasa_anual, plazo_anios, meses_gracia)
    dividendo_clp = dividendo_uf * valor_uf_actual
    
    flujo_caja = arriendo_estimado - dividendo_clp
    cap_rate = (arriendo_estimado * 12) / valor_propiedad_clp * 100
    carga_actual = (dividendo_clp / ingreso_actual) * 100 if ingreso_actual > 0 else 0
    
    # Renderizado visual de resultados en la web
    st.divider()
    st.markdown("### Resultados de la Simulación")
    r1, r2, r3, r4 = st.columns(4)
    r1.metric("Pie + Gastos Operacionales", f"$ {(pie_clp + gastos_clp):,.0f}")
    r2.metric("Dividendo (Post-Gracia)", f"$ {dividendo_clp:,.0f}")
    r3.metric("Flujo de Caja", f"$ {flujo_caja:,.0f}", delta="Positivo" if flujo_caja > 0 else "Negativo")
    r4.metric("Carga Financiera", f"{carga_actual:.1f}%", delta="Apto" if carga_actual <= 25 else "Riesgo Banco", delta_color="inverse" if carga_actual > 25 else "normal")

    # Generación y Descarga del Documento Word
    doc = generar_prospecto_avanzado(
        ubicacion, distribucion, estado_propiedad, financiamiento, tipo_tasa, 
        valor_propiedad_uf, valor_propiedad_clp, pie_clp, usar_vehiculo, 
        cap_vehiculo, gastos_clp, meses_gracia, arriendo_estimado, 
        dividendo_clp, flujo_caja, cap_rate, carga_actual
    )
    
    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    
    st.download_button(
        label="📥 Descargar Prospecto Profesional (.docx)", 
        data=buffer, 
        file_name="Prospecto_Avanzado.docx", 
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )