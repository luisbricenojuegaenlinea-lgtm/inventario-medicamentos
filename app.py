"""
╔══════════════════════════════════════════════════════════════════╗
║   CONTROL DE INVENTARIO - MEDICAMENTOS ANTIHIPERTENSIVOS        ║
║   Aplicación con Streamlit + Google Sheets (gspread)            ║
║   Compatible con ejecución LOCAL y Streamlit Cloud              ║
╚══════════════════════════════════════════════════════════════════╝

MODOS DE EJECUCIÓN:

  • LOCAL: Usa el archivo .json de credenciales en la misma carpeta.
  • NUBE (Streamlit Cloud): Lee credenciales desde st.secrets.

Para ejecutar localmente:
  streamlit run app.py
"""

import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime, date, timedelta
from decimal import Decimal, ROUND_HALF_UP
import json
import os

# ═══════════════════════════════════════════════════════════════════
# ██  CONFIGURACIÓN  ██
# ═══════════════════════════════════════════════════════════════════
CREDENTIALS_FILE = "control-pastillas-504117-9d0038fd1f57.json"
SPREADSHEET_NAME = "Inventario_Antihipertensivos"
# ═══════════════════════════════════════════════════════════════════

# Constantes del sistema
PASTILLAS_POR_CAJA = 30
DIAS_ALERTA_MINIMO = 15
FECHAS_DE_COBRO = {4, 5, 15, 30}
COL_DESCUENTO = 5  # Columna E: Fecha Ultimo Descuento

# Alcances (scopes) de la API de Google
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


# ─── Conexión a Google Sheets ────────────────────────────────────
@st.cache_resource(ttl=300)
def conectar_google_sheets():
    """
    Crea y cachea la conexión al documento de Google Sheets.
    Detecta automáticamente si está en Streamlit Cloud o local.
    """
    try:
        # Modo 1: Streamlit Cloud (credenciales en secrets)
        if "gcp_service_account" in st.secrets:
            creds_dict = dict(st.secrets["gcp_service_account"])
            creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
        # Modo 2: Local (archivo .json)
        elif os.path.exists(CREDENTIALS_FILE):
            creds = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        else:
            st.error(
                f"❌ **No se encontraron credenciales.**\n\n"
                f"**Local:** Coloca `{CREDENTIALS_FILE}` en la carpeta de la app.\n\n"
                f"**Streamlit Cloud:** Configura los secrets con tu JSON de credenciales."
            )
            st.stop()

        cliente = gspread.authorize(creds)
        hoja = cliente.open(SPREADSHEET_NAME).sheet1
        return hoja

    except gspread.exceptions.SpreadsheetNotFound:
        st.error(
            f"❌ **No se encontró la hoja de cálculo:** `{SPREADSHEET_NAME}`\n\n"
            "Verifica que:\n"
            "1. El nombre sea exacto (mayúsculas y espacios cuentan).\n"
            "2. El bot (Service Account) tenga permisos de Editor en esa hoja."
        )
        st.stop()
    except Exception as e:
        st.error(f"❌ **Error de conexión:** {e}")
        st.stop()


def asegurar_columna_descuento(hoja):
    """
    Verifica que exista la columna E con encabezado 'Fecha Ultimo Descuento'.
    Si no existe, la crea automáticamente.
    """
    try:
        encabezado_e = hoja.cell(1, COL_DESCUENTO).value
        if encabezado_e != "Fecha Ultimo Descuento":
            hoja.update_cell(1, COL_DESCUENTO, "Fecha Ultimo Descuento")
    except Exception:
        hoja.update_cell(1, COL_DESCUENTO, "Fecha Ultimo Descuento")


def leer_inventario(hoja):
    """
    Lee todas las filas del inventario desde Google Sheets.
    Retorna una lista de diccionarios con los datos de cada medicamento.
    Usa Decimal para precisión exacta con dosis decimales (0.5).
    """
    registros = hoja.get_all_records()
    inventario = []

    for i, fila in enumerate(registros):
        nombre = str(fila.get("Medicamento", "")).strip()
        if not nombre:
            continue

        # Usar Decimal para evitar errores de punto flotante
        stock_raw = fila.get("Stock Actual", 0)
        dosis_raw = fila.get("Dosis Diaria", 0)

        stock = Decimal(str(stock_raw))
        dosis = Decimal(str(dosis_raw))

        fecha_str = str(fila.get("Fecha de Ultima Actualizacion", ""))

        # Leer fecha del último descuento (columna E)
        fecha_descuento_str = str(fila.get("Fecha Ultimo Descuento", "")).strip()

        inventario.append({
            "fila_sheets": i + 2,
            "medicamento": nombre,
            "stock": stock,
            "dosis": dosis,
            "fecha_actualizacion": fecha_str,
            "fecha_ultimo_descuento": fecha_descuento_str,
        })

    return inventario


def aplicar_descuento_automatico(hoja, inventario):
    """
    DESCUENTO INTELIGENTE CON RECUPERACIÓN:
    Calcula cuántos días han pasado desde el último descuento
    y resta todas las dosis acumuladas de una vez.

    Ejemplo: si no abriste la app en 3 días, descuenta 3 dosis.
    Usa Decimal para precisión exacta con dosis de 0.5.
    """
    hoy = date.today()
    hoy_str = hoy.strftime("%Y-%m-%d")
    resultados = []

    for med in inventario:
        fecha_desc = med["fecha_ultimo_descuento"]

        # Verificar si ya se descontó hoy
        if fecha_desc and fecha_desc.startswith(hoy_str):
            resultados.append({
                "medicamento": med["medicamento"],
                "estado": "ya_aplicado",
                "stock": med["stock"],
                "dosis": med["dosis"],
            })
            continue

        # Calcular días pendientes de descuento
        dias_pendientes = 1  # Mínimo descontar hoy

        if fecha_desc:
            try:
                # Extraer solo la parte de fecha (YYYY-MM-DD)
                fecha_ultimo = datetime.strptime(fecha_desc[:10], "%Y-%m-%d").date()
                diferencia = (hoy - fecha_ultimo).days
                if diferencia > 1:
                    dias_pendientes = diferencia
            except (ValueError, IndexError):
                dias_pendientes = 1  # Si no se puede parsear, descontar solo hoy

        # Calcular descuento total
        descuento_total = med["dosis"] * Decimal(str(dias_pendientes))
        nuevo_stock = med["stock"] - descuento_total

        if nuevo_stock < 0:
            # Descontar lo máximo posible sin quedar negativo
            if med["dosis"] > 0:
                dias_posibles = int(med["stock"] / med["dosis"])
                if dias_posibles > 0:
                    descuento_total = med["dosis"] * Decimal(str(dias_posibles))
                    nuevo_stock = med["stock"] - descuento_total
                    dias_pendientes = dias_posibles
                else:
                    resultados.append({
                        "medicamento": med["medicamento"],
                        "estado": "sin_stock",
                        "stock": med["stock"],
                        "dosis": med["dosis"],
                    })
                    continue
            else:
                resultados.append({
                    "medicamento": med["medicamento"],
                    "estado": "sin_stock",
                    "stock": med["stock"],
                    "dosis": med["dosis"],
                })
                continue

        # Actualizar stock en Google Sheets
        actualizar_celda_stock(hoja, med["fila_sheets"], nuevo_stock)

        # Marcar la fecha del descuento en columna E
        fecha_hora = datetime.now().strftime("%Y-%m-%d %H:%M")
        hoja.update_cell(med["fila_sheets"], COL_DESCUENTO, fecha_hora)

        resultados.append({
            "medicamento": med["medicamento"],
            "estado": "aplicado",
            "stock_anterior": med["stock"],
            "stock_nuevo": nuevo_stock,
            "dosis": med["dosis"],
            "dias_descontados": dias_pendientes,
        })

    return resultados


def actualizar_celda_stock(hoja, fila, nuevo_stock):
    """
    Actualiza el stock (columna B) y la fecha (columna D) en Google Sheets.
    Convierte Decimal a float para compatibilidad con la API de Sheets.
    """
    fecha_hoy = datetime.now().strftime("%Y-%m-%d %H:%M")
    valor_stock = float(nuevo_stock)

    # Si el valor es entero, enviar como int para limpieza visual
    if nuevo_stock == int(nuevo_stock):
        valor_stock = int(nuevo_stock)

    hoja.update_cell(fila, 2, valor_stock)         # Columna B: Stock Actual
    hoja.update_cell(fila, 4, fecha_hoy)            # Columna D: Fecha de Ultima Actualizacion


def calcular_dias_restantes(stock, dosis):
    """
    Calcula cuántos días durará el stock actual según la dosis diaria.
    Usa Decimal para precisión exacta.
    """
    if dosis <= 0:
        return Decimal("9999")
    return (stock / dosis).quantize(Decimal("0.1"), rounding=ROUND_HALF_UP)


def verificar_alerta_cobro(inventario):
    """
    Verifica si hoy es día de cobro y si algún medicamento no durará
    al menos 15 días más. Retorna lista de alertas.
    """
    hoy = date.today()
    dia_actual = hoy.day
    alertas = []

    if dia_actual not in FECHAS_DE_COBRO:
        return alertas

    for med in inventario:
        dias = calcular_dias_restantes(med["stock"], med["dosis"])
        if dias < DIAS_ALERTA_MINIMO:
            alertas.append({
                "medicamento": med["medicamento"],
                "stock": med["stock"],
                "dias_restantes": dias,
                "dosis": med["dosis"],
            })

    return alertas


# ─── Estilos CSS personalizados ──────────────────────────────────
def aplicar_estilos():
    st.markdown("""
    <style>
        /* Tipografía general */
        @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
        html, body, [class*="st-"] { font-family: 'Inter', sans-serif; }

        /* Tarjetas de medicamento */
        .med-card {
            background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
            border: 1px solid rgba(99, 102, 241, 0.3);
            border-radius: 16px;
            padding: 1.5rem;
            margin-bottom: 1rem;
            box-shadow: 0 4px 20px rgba(0, 0, 0, 0.3);
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }
        .med-card:hover {
            transform: translateY(-2px);
            box-shadow: 0 8px 30px rgba(99, 102, 241, 0.2);
        }
        .med-nombre {
            font-size: 1.15rem;
            font-weight: 600;
            color: #a5b4fc;
            margin-bottom: 0.5rem;
            letter-spacing: 0.02em;
        }
        .med-stock {
            font-size: 2.2rem;
            font-weight: 700;
            color: #e0e7ff;
            line-height: 1.1;
        }
        .med-stock-unidad {
            font-size: 0.85rem;
            color: #818cf8;
            font-weight: 500;
        }
        .med-detalle {
            font-size: 0.82rem;
            color: #94a3b8;
            margin-top: 0.5rem;
        }

        /* Stock bajo */
        .stock-bajo { border-color: rgba(239, 68, 68, 0.5); }
        .stock-bajo .med-stock { color: #fca5a5; }

        /* Stock medio */
        .stock-medio { border-color: rgba(251, 191, 36, 0.4); }
        .stock-medio .med-stock { color: #fde68a; }

        /* Stock alto */
        .stock-alto { border-color: rgba(52, 211, 153, 0.4); }
        .stock-alto .med-stock { color: #6ee7b7; }

        /* Alerta de cobro */
        .alerta-cobro {
            background: linear-gradient(135deg, #450a0a 0%, #7f1d1d 100%);
            border: 2px solid #ef4444;
            border-radius: 12px;
            padding: 1.2rem;
            margin: 0.8rem 0;
            animation: pulse-border 2s ease-in-out infinite;
        }
        @keyframes pulse-border {
            0%, 100% { border-color: #ef4444; box-shadow: 0 0 10px rgba(239, 68, 68, 0.3); }
            50% { border-color: #f87171; box-shadow: 0 0 25px rgba(239, 68, 68, 0.5); }
        }
        .alerta-titulo {
            color: #fca5a5;
            font-weight: 700;
            font-size: 1.05rem;
        }
        .alerta-texto {
            color: #fecaca;
            font-size: 0.9rem;
            margin-top: 0.3rem;
        }

        /* Banner día de cobro */
        .cobro-banner {
            background: linear-gradient(90deg, #312e81, #4338ca, #312e81);
            border-radius: 10px;
            padding: 0.8rem 1.2rem;
            text-align: center;
            margin-bottom: 1rem;
        }
        .cobro-banner-texto {
            color: #c7d2fe;
            font-size: 0.9rem;
            font-weight: 500;
        }

        /* Sección de éxito */
        .exito-msg {
            background: linear-gradient(135deg, #064e3b 0%, #065f46 100%);
            border: 1px solid #10b981;
            border-radius: 12px;
            padding: 1rem 1.2rem;
            margin: 0.8rem 0;
        }
        .exito-msg p {
            color: #a7f3d0;
            font-size: 0.95rem;
            margin: 0;
        }

        /* Header personalizado */
        .app-header {
            text-align: center;
            padding: 1rem 0 0.5rem 0;
        }
        .app-header h1 {
            background: linear-gradient(90deg, #818cf8, #a78bfa, #c084fc);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            font-size: 1.8rem;
            font-weight: 700;
            margin-bottom: 0.2rem;
        }
        .app-header p {
            color: #64748b;
            font-size: 0.85rem;
        }

        /* Divider sutil */
        .divider {
            border: none;
            height: 1px;
            background: linear-gradient(90deg, transparent, rgba(99, 102, 241, 0.3), transparent);
            margin: 1.5rem 0;
        }
    </style>
    """, unsafe_allow_html=True)


def renderizar_tarjeta_medicamento(med):
    """Renderiza una tarjeta visual para un medicamento."""
    dias = calcular_dias_restantes(med["stock"], med["dosis"])

    # Determinar clase CSS según nivel de stock
    if dias < 10:
        clase_stock = "stock-bajo"
        icono = "🔴"
    elif dias < 20:
        clase_stock = "stock-medio"
        icono = "🟡"
    else:
        clase_stock = "stock-alto"
        icono = "🟢"

    # Formatear stock: mostrar entero si es posible, decimal si no
    stock_display = int(med["stock"]) if med["stock"] == int(med["stock"]) else med["stock"]
    dosis_display = int(med["dosis"]) if med["dosis"] == int(med["dosis"]) else med["dosis"]

    st.markdown(f"""
    <div class="med-card {clase_stock}">
        <div class="med-nombre">{icono} {med['medicamento']}</div>
        <div class="med-stock">{stock_display} <span class="med-stock-unidad">pastillas</span></div>
        <div class="med-detalle">
            Dosis diaria: {dosis_display} · Alcanza para <b>{dias}</b> días
        </div>
        <div class="med-detalle">
            Última actualización: {med['fecha_actualizacion'] or '—'}
        </div>
    </div>
    """, unsafe_allow_html=True)


# ═══════════════════════════════════════════════════════════════════
#                     APLICACIÓN PRINCIPAL
# ═══════════════════════════════════════════════════════════════════
def main():
    st.set_page_config(
        page_title="Inventario Antihipertensivos",
        page_icon="💊",
        layout="wide",
        initial_sidebar_state="collapsed",
    )

    aplicar_estilos()

    # ─── Header ──────────────────────────────────────────────
    st.markdown("""
    <div class="app-header">
        <h1>💊 Control de Inventario</h1>
        <p>Medicamentos Antihipertensivos · Conectado a Google Sheets</p>
    </div>
    <hr class="divider">
    """, unsafe_allow_html=True)

    # ─── Conexión ────────────────────────────────────────────
    hoja = conectar_google_sheets()

    # ─── Asegurar columna E para descuentos ──────────────────
    asegurar_columna_descuento(hoja)

    # ─── Leer inventario ─────────────────────────────────────
    inventario = leer_inventario(hoja)

    if not inventario:
        st.warning("⚠️ No se encontraron medicamentos en la hoja de cálculo. "
                    "Verifica que los encabezados sean: Medicamento, Stock Actual, "
                    "Dosis Diaria, Fecha de Ultima Actualizacion")
        st.stop()

    # ─── Descuento Diario Automático con Recuperación ────────
    # Se ejecuta UNA VEZ por sesión. Calcula días pendientes y
    # descuenta todas las dosis acumuladas de golpe.
    if "descuento_aplicado" not in st.session_state:
        resultados_descuento = aplicar_descuento_automatico(hoja, inventario)
        hubo_descuentos = any(r["estado"] == "aplicado" for r in resultados_descuento)
        st.session_state["descuento_aplicado"] = True
        st.session_state["resultados_descuento"] = resultados_descuento
        if hubo_descuentos:
            # Re-leer inventario después del descuento
            st.cache_resource.clear()
            st.rerun()

    # ─── Alerta de Día de Cobro ──────────────────────────────
    hoy = date.today()
    dia_actual = hoy.day
    es_dia_cobro = dia_actual in FECHAS_DE_COBRO

    if es_dia_cobro:
        st.markdown(f"""
        <div class="cobro-banner">
            <div class="cobro-banner-texto">
                📅 Hoy es <b>día de cobro</b> (día {dia_actual} del mes) — Verificando stock proyectado a 15 días...
            </div>
        </div>
        """, unsafe_allow_html=True)

        alertas = verificar_alerta_cobro(inventario)

        if alertas:
            for alerta in alertas:
                st.markdown(f"""
                <div class="alerta-cobro">
                    <div class="alerta-titulo">
                        ⚠️ ALERTA: {alerta['medicamento']} NO alcanzará 15 días
                    </div>
                    <div class="alerta-texto">
                        Stock actual: <b>{alerta['stock']}</b> pastillas ·
                        Dosis diaria: <b>{alerta['dosis']}</b> ·
                        Duración estimada: <b>{alerta['dias_restantes']} días</b> ·
                        Se necesitan al menos <b>{float(alerta['dosis']) * DIAS_ALERTA_MINIMO}</b> pastillas para cubrir 15 días.
                    </div>
                </div>
                """, unsafe_allow_html=True)
        else:
            st.success("✅ Día de cobro: Todos los medicamentos tienen stock suficiente para al menos 15 días.")

    # ─── Visualización en Tiempo Real ────────────────────────
    st.markdown("### 📦 Stock Actual")

    cols = st.columns(len(inventario))
    for col, med in zip(cols, inventario):
        with col:
            renderizar_tarjeta_medicamento(med)

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # ─── Módulos de Acción (2 columnas) ──────────────────────
    col_compra, col_descuento = st.columns(2)

    # ─── Módulo de Reposición de Cajas ───────────────────────
    with col_compra:
        st.markdown("### 🛒 Registrar Compra")
        st.caption(f"Suma automáticamente **{PASTILLAS_POR_CAJA} pastillas** al stock del medicamento seleccionado.")

        nombres = [m["medicamento"] for m in inventario]
        seleccion_compra = st.selectbox(
            "Medicamento comprado:",
            options=nombres,
            key="select_compra",
        )

        if st.button("📥 Registrar Compra", key="btn_compra", use_container_width=True, type="primary"):
            med_sel = next(m for m in inventario if m["medicamento"] == seleccion_compra)
            stock_anterior = med_sel["stock"]
            nuevo_stock = stock_anterior + Decimal(str(PASTILLAS_POR_CAJA))

            actualizar_celda_stock(hoja, med_sel["fila_sheets"], nuevo_stock)

            st.cache_resource.clear()
            if "descuento_aplicado" in st.session_state:
                del st.session_state["descuento_aplicado"]

            st.rerun()

    # ─── Panel de Estado: Descuento Diario Automático ────────
    with col_descuento:
        st.markdown("### 💊 Descuento Diario Automático")
        st.caption("El sistema descuenta automáticamente la dosis diaria al abrir la app. Si no la abriste en varios días, recupera todos los días pendientes.")

        resultados = st.session_state.get("resultados_descuento", [])

        if resultados:
            for r in resultados:
                dosis_display = int(r["dosis"]) if r["dosis"] == int(r["dosis"]) else r["dosis"]

                if r["estado"] == "aplicado":
                    dias_txt = r.get("dias_descontados", 1)
                    total_desc = r["dosis"] * Decimal(str(dias_txt))
                    total_display = int(total_desc) if total_desc == int(total_desc) else total_desc
                    plural = "día" if dias_txt == 1 else "días"
                    st.markdown(f"""
                    <div class="exito-msg">
                        <p>✅ <b>{r['medicamento']}</b>: {r['stock_anterior']} → <b>{r['stock_nuevo']}</b> pastillas (−{total_display} por {dias_txt} {plural})</p>
                    </div>
                    """, unsafe_allow_html=True)
                elif r["estado"] == "ya_aplicado":
                    st.markdown(f"""
                    <div class="exito-msg" style="border-color: #6366f1; background: linear-gradient(135deg, #1e1b4b 0%, #312e81 100%);">
                        <p style="color: #c7d2fe;">☑️ <b>{r['medicamento']}</b>: dosis de hoy ya descontada · Stock: <b>{r['stock']}</b></p>
                    </div>
                    """, unsafe_allow_html=True)
                elif r["estado"] == "sin_stock":
                    st.markdown(f"""
                    <div class="alerta-cobro">
                        <div class="alerta-titulo">⚠️ {r['medicamento']}: stock insuficiente</div>
                        <div class="alerta-texto">Stock actual: <b>{r['stock']}</b> · Dosis requerida: <b>{dosis_display}</b></div>
                    </div>
                    """, unsafe_allow_html=True)
        else:
            st.info("ℹ️ El descuento automático se ejecuta al abrir la app cada día.")

    # ─── Footer informativo ──────────────────────────────────
    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    footer_cols = st.columns(3)
    with footer_cols[0]:
        st.caption(f"📅 Fecha actual: **{hoy.strftime('%d/%m/%Y')}**")
    with footer_cols[1]:
        st.caption(f"📆 Días de cobro: **{', '.join(str(d) for d in sorted(FECHAS_DE_COBRO))}**")
    with footer_cols[2]:
        cobro_texto = "🟢 **Sí, hoy es día de cobro**" if es_dia_cobro else "⚪ No es día de cobro"
        st.caption(f"Estado: {cobro_texto}")

    # Botón de refresco manual
    st.markdown("")
    if st.button("🔄 Refrescar datos desde Google Sheets", use_container_width=True):
        st.cache_resource.clear()
        if "descuento_aplicado" in st.session_state:
            del st.session_state["descuento_aplicado"]
        st.rerun()


if __name__ == "__main__":
    main()
