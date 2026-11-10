import os
import io
import copy
import tempfile
from motor_sec_cad import PROYECTO_DEFAULT, calcular_metricas_sec, generar_dxf_proyecto


def _corriendo_en_streamlit():
    try:
        from streamlit.runtime.scriptrunner import get_script_run_ctx
        return get_script_run_ctx() is not None
    except Exception:
        return False


# =====================================================================
# MODO 1: APLICACIÓN WEB INTERACTIVA (STREAMLIT CLOUD / LOCAL)
# =====================================================================
def ejecutar_interfaz_streamlit():
    import streamlit as st
    import pandas as pd

    st.set_page_config(page_title="Motor Planos SEC TE1 - Chile", layout="wide")
    st.title("⚡ Motor Automatizado de Planos Eléctricos SEC (TE1) - AutoCAD DXF")
    st.caption("Generación paramétrica de plantas de arquitectura, empalme con doble tierra, cuadro de cargas y clasificación D.S. N° 92/83.")

    proyecto = copy.deepcopy(PROYECTO_DEFAULT)
    cli = proyecto["cliente"]

    with st.sidebar:
        st.header("1. Antecedentes del Proyecto")
        cli["propietario"] = st.text_input("Propietario", value=cli["propietario"])
        cli["rut"] = st.text_input("RUT Propietario", value=cli["rut"])
        cli["direccion"] = st.text_input("Calle / Sector", value=cli["direccion"])
        cli["comuna"] = st.text_input("Comuna", value=cli["comuna"])
        cli["poste_red"] = st.text_input("Poste / S/E Distribuidora", value=cli["poste_red"])

        st.header("2. Parámetros del Alimentador")
        cli["alimentador_largo_m"] = st.number_input("Longitud Alimentador L (m)", min_value=1.0, value=float(cli["alimentador_largo_m"]), step=5.0)
        cli["alimentador_seccion_mm2"] = st.selectbox("Sección Alimentador S (mm²)", options=[2.5, 4.0, 6.0, 10.0, 16.0, 25.0, 35.0], index=5)
        cli["pot_enchufe_w"] = st.selectbox("Potencia Unitaria por Enchufe (W)", options=[250, 300, 420], index=0)
        cli["pot_foco_led_w"] = st.selectbox("Potencia Unitaria por Foco LED (W)", options=[18, 70, 80, 150], index=2)

        plantilla_subida = st.file_uploader("Plantilla DXF Base (Opcional)", type=["dxf"])

    calc = calcular_metricas_sec(proyecto)

    # Tarjetas de resumen normativo SEC
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Total Centros", f"{calc['tot_centros']}", f"{calc['tot_ench']} Ench. / {calc['tot_luces']} Luces")
    c2.metric("Potencia Instalada", f"{calc['pot_tot_w']:,} W", f"{calc['pot_tot_kw']:.2f} kW")
    c3.metric("Corriente Total (Fase R)", f"{calc['cor_tot_r']:.2f} A", f"Disy. Ench: {calc['disy_ench']}")
    c4.metric("Caída de Tensión (VP)", f"{calc['vp_val']:.3f} V", f"Sección Teórica: {calc['s_val']:.2f} mm²")
    c5.metric("Clasificación D.S. 92/83", f"Tipo {calc['letra_sec']}", "Paso 3 e-Declarador")

    st.info(f"**Criterio SEC D.S. N° 92/83 detectado:** Se selecciona **Tipo {calc['letra_sec']}** — {calc['motivo_sec']}.")

    # Tabla resumen de recintos
    st.subheader("Recintos y Centros del Borrador")
    filas_hab = []
    for h in proyecto["habitaciones"]:
        filas_hab.append({
            "Recinto": h["nombre"],
            "Ancho (m)": h["w"],
            "Largo (m)": h["h"],
            "Centros Luz (⊗)": len(h["luces"]),
            "Interruptores": len(h["interruptores"]),
            "Enchufes": len(h["enchufes"])
        })
    st.dataframe(pd.DataFrame(filas_hab), use_container_width=True)

    # Generar archivo DXF para descarga directa en el navegador
    ruta_temp_tpl = None
    if plantilla_subida is not None:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmp_in:
            tmp_in.write(plantilla_subida.read())
            ruta_temp_tpl = tmp_in.name

    doc, _ = generar_dxf_proyecto(proyecto, ruta_plantilla=ruta_temp_tpl)
    stream_dxf = io.StringIO()
    doc.write(stream_dxf)
    dxf_bytes = stream_dxf.getvalue().encode("utf-8")

    st.download_button(
        label="⬇️ DESCARGAR PLANO AUTOCAD COMPLETO (.DXF)",
        data=dxf_bytes,
        file_name=f"TE1_{cli['propietario'].replace(' ', '_')}.dxf",
        mime="application/dxf",
        use_container_width=True
    )


# =====================================================================
# MODO 2: EJECUTOR DIRECTO DE ESCRITORIO WINDOWS (DOBLE CLIC / PLAY)
# =====================================================================
def ejecutar_escritorio_windows():
    import tkinter as tk
    from tkinter import filedialog, messagebox

    root = tk.Tk()
    root.withdraw()

    ruta_plantilla = filedialog.askopenfilename(
        title="1. Selecciona tu plantilla DXF (ej. CHECHIN.dxf o Cancela para plano nuevo)",
        filetypes=[("Archivos DXF de AutoCAD", "*.dxf"), ("Todos los archivos", "*.*")]
    )

    doc, calc = generar_dxf_proyecto(PROYECTO_DEFAULT, ruta_plantilla=ruta_plantilla)

    carpeta_inicial = os.path.dirname(ruta_plantilla) if ruta_plantilla else os.path.expanduser("~\\Downloads")
    ruta_salida = filedialog.asksaveasfilename(
        title="2. Guarda tu plano SEC TE1 listo (.dxf)",
        initialdir=carpeta_inicial,
        initialfile="CHECHIN_SEC_TE1_COMPLETO.dxf",
        defaultextension=".dxf",
        filetypes=[("Archivo DXF de AutoCAD", "*.dxf")]
    )

    if not ruta_salida:
        return

    doc.saveas(ruta_salida)
    messagebox.showinfo(
        "¡Plano SEC TE1 Generado!",
        f"Archivo guardado en:\n{ruta_salida}\n\n"
        f"• Centros: {calc['tot_ench']} enchufes + {calc['tot_luces']} luces = {calc['tot_centros']}\n"
        f"• Potencia Total: {calc['pot_tot_w']} W ({calc['pot_tot_kw']} kW) | I = {calc['cor_tot_r']} A\n"
        f"• Caída de Tensión (L={calc['L']:.0f}m, S={calc['S_cond']:.0f}mm²): VP = {calc['vp_val']} V\n"
        f"• Letra e-Declarador (D.S. 92/83): TIPO {calc['letra_sec']}"
    )
    os.startfile(ruta_salida)


if __name__ == "__main__":
    if _corriendo_en_streamlit():
        ejecutar_interfaz_streamlit()
    else:
        ejecutar_escritorio_windows()