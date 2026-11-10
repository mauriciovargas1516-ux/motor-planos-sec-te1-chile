import os
import io
import math
import ezdxf
from ezdxf import bbox
from ezdxf.enums import TextEntityAlignment

try:
    import openpyxl
    OPENPYXL_OK = True
except ImportError:
    OPENPYXL_OK = False

ESP_MURO = 0.12

# =====================================================================
# CONFIGURACIÓN BASE POR DEFECTO (CASO REAL: PROYECTO RURAL "CHECHIN")
# =====================================================================
PROYECTO_DEFAULT = {
    "cliente": {
        "propietario": "RAUL HUMBERTO PUÑANCO ANTILEF",
        "rut": "15.914.402-K",
        "direccion": "Sector Loma de la Piedra",
        "comuna": "San Juan de la Costa",
        "poste_red": "Poste BS 4086 / S/E 46845",
        "acometida_largo": "L :15 metros",
        "acometida_cable": "concentrico 2 x 6 mm2",
        "alimentador_largo_m": 120.0,
        "alimentador_seccion_mm2": 25.0,
        "alimentador_texto_1": "L :120 mts (3 Postes 75x75x2m)",
        "alimentador_texto_2": "Preensamblado 2x25 mm2",
        "alimentador_texto_3": "bajada 15m concentrico 2x6 mm2",
        "pot_foco_led_w": 80,
        "pot_calefaccion_w": 600,
        "pot_horno_w": 250,
        "pot_enchufe_w": 250
    },
    "tda": (9.60, 7.15),
    "ubicacion_linea1": "Dormitorio, cocina, comedor",
    "ubicacion_linea2": "baño",
    "habitaciones": [
        {
            "nombre": "DORM. 2", "x": 2.0, "y": 8.0, "w": 3.5, "h": 3.0,
            "caja": (4.0, 8.4),
            "luces": [(3.8, 9.5)],
            "interruptores": [(4.4, 8.3)],
            "enchufes": [(2.7, 11.0 - ESP_MURO, "ARRIBA"), (2.0 + ESP_MURO, 8.6, "IZQ")],
            "ventanas": [(2.0, 9.2, 2.0, 10.5)]
        },
        {
            "nombre": "DORM. 1", "x": 5.5, "y": 8.0, "w": 4.0, "h": 3.0,
            "caja": (6.2, 8.4),
            "luces": [(7.2, 9.5), (6.2, 11.6)],
            "interruptores": [(5.8, 8.3)],
            "enchufes": [(8.0, 11.0 - ESP_MURO, "ARRIBA"), (6.8, 8.0 + ESP_MURO, "ABAJO")],
            "ventanas": [(9.5, 8.8, 9.5, 10.2), (5.5, 9.0, 5.5, 10.5)]
        },
        {
            "nombre": "COCINA", "x": 0.0, "y": 5.0, "w": 3.5, "h": 3.0,
            "caja": (2.2, 7.6),
            "luces": [(0.7, 6.8), (1.8, 5.6)],
            "interruptores": [(1.8, 7.75)],
            "enchufes": [(1.1, 8.0 - ESP_MURO, "ARRIBA"), (2.8, 5.0 + ESP_MURO, "ABAJO")],
            "ventanas": [(0.0, 5.8, 0.0, 7.2)]
        },
        {
            "nombre": "COMEDOR", "x": 3.5, "y": 3.5, "w": 6.5, "h": 4.5,
            "caja": (8.8, 7.4),
            "luces": [(9.2, 6.0), (6.8, 4.2)],
            "interruptores": [(8.5, 7.7), (6.0, 3.9)],
            "enchufes": [(8.1, 8.0 - ESP_MURO, "ARRIBA"), (10.0 - ESP_MURO, 6.2, "DER")],
            "ventanas": [(10.0, 4.0, 10.0, 5.2), (10.0, 5.8, 10.0, 7.0)]
        },
        {
            "nombre": "BAÑO", "x": 0.0, "y": 3.0, "w": 3.5, "h": 2.0,
            "caja": (3.1, 4.2),
            "luces": [(1.8, 4.2)],
            "interruptores": [(3.4, 4.2)],
            "enchufes": [(3.5 - ESP_MURO, 3.6, "DER")],
            "ventanas": [(0.0, 3.7, 0.0, 4.5)]
        },
        {
            "nombre": "DORM. 4", "x": 0.0, "y": 0.0, "w": 5.0, "h": 3.0,
            "caja": (3.8, 2.7),
            "luces": [(2.6, 1.4)],
            "interruptores": [(4.2, 2.75)],
            "enchufes": [(3.2, 3.0 - ESP_MURO, "ARRIBA"), (2.2, 3.0 - ESP_MURO, "ARRIBA"), (0.0 + ESP_MURO, 1.2, "IZQ")],
            "ventanas": [(0.0, 0.8, 0.0, 2.5)]
        },
        {
            "nombre": "DORM. 3", "x": 5.0, "y": 0.0, "w": 4.0, "h": 3.5,
            "caja": (6.0, 3.1),
            "luces": [(7.0, 1.8), (7.3, -0.5)],
            "interruptores": [(5.5, 3.2)],
            "enchufes": [(5.2, 3.5 - ESP_MURO, "ARRIBA"), (9.0 - ESP_MURO, 1.5, "DER")],
            "ventanas": [(9.0, 0.8, 9.0, 2.2)]
        }
    ]
}


# =====================================================================
# 1. MOTOR DE CÁLCULO ELÉCTRICO Y CLASIFICACIÓN D.S. N° 92/83 (TE1)
# =====================================================================
def calcular_metricas_sec(proyecto):
    cli = proyecto["cliente"]
    habs = proyecto["habitaciones"]

    tot_luces = sum(len(h["luces"]) for h in habs)
    tot_ench = sum(len(h["enchufes"]) for h in habs)
    tot_centros = tot_luces + tot_ench

    pot_e = tot_ench * int(cli["pot_enchufe_w"])
    pot_l = tot_luces * int(cli["pot_foco_led_w"])
    pot_tot_w = pot_e + pot_l
    pot_tot_kw = round(pot_tot_w / 1000.0, 3)

    cor_e = pot_e / 220.0
    cor_l = pot_l / 220.0
    cor_tot = cor_e + cor_l
    cor_tot_r = round(cor_tot, 2)

    L = float(cli["alimentador_largo_m"])
    S_cond = float(cli["alimentador_seccion_mm2"])

    num_vp = round(2.0 * L * cor_tot_r, 2)
    den_vp = int(56 * S_cond)
    vp_val = round(num_vp / den_vp, 3) if den_vp > 0 else 0.0

    num_s = round(2.0 * L * 0.018 * cor_tot_r, 4)
    s_val = round(num_s / 6.6, 3)

    # Clasificación automática según D.S. N° 92/83 (Paso 3 e-Declarador SEC)
    if L <= 0:
        letra_sec = "F"
        motivo_sec = "Alumbrado en BT máx. 10 kW sin alimentadores"
    elif L <= 100.0 and pot_tot_kw <= 10.0:
        letra_sec = "D"
        motivo_sec = "Alumbrado BT con alimentador <= 100 m y <= 10 kW por fase"
    else:
        letra_sec = "B"
        motivo_sec = f"Baja tensión hasta 500 kW (Alimentador de {L:.0f} m supera límite de 100 m del Tipo D)"

    return {
        "tot_luces": tot_luces,
        "tot_ench": tot_ench,
        "tot_centros": tot_centros,
        "pot_e": pot_e,
        "pot_l": pot_l,
        "pot_tot_w": pot_tot_w,
        "pot_tot_kw": pot_tot_kw,
        "cor_e": cor_e,
        "cor_l": cor_l,
        "cor_tot_r": cor_tot_r,
        "L": L,
        "S_cond": S_cond,
        "num_vp": num_vp,
        "den_vp": den_vp,
        "vp_val": vp_val,
        "num_s": num_s,
        "s_val": s_val,
        "disy_ench": "1X16A" if cor_e <= 16.0 else "1X25A",
        "letra_sec": letra_sec,
        "motivo_sec": motivo_sec
    }


# =====================================================================
# 2. GEOMETRÍA VECTORIAL Y SIMBOLOGÍA SÓLIDA EN AUTOCAD (EZDXF)
# =====================================================================
def dibujar_rectangulo(espacio, x1, y1, x2, y2, capa):
    espacio.add_lwpolyline(
        [(x1, y1), (x2, y1), (x2, y2), (x1, y2)],
        close=True,
        dxfattribs={"layer": capa}
    )


def dibujar_tda_solido(espacio, x, y, w, h, num_cto, escala):
    hatch = espacio.add_hatch(color=3, dxfattribs={"layer": "TDA_VERDE"})
    hatch.paths.add_polyline_path(
        [(x, y), (x + w, y), (x + w, y + h), (x, y + h)],
        is_closed=True
    )
    dibujar_rectangulo(espacio, x, y, x + w, y + h, "TDA_VERDE")

    cx, cy = x - 0.25 * escala, y + h / 2.0
    espacio.add_circle((cx, cy), radius=0.18 * escala, dxfattribs={"layer": "SIMBOLOS_AMARILLO"})
    espacio.add_text(
        str(num_cto),
        dxfattribs={"layer": "SIMBOLOS_AMARILLO", "height": 0.20 * escala}
    ).set_placement((cx, cy), align=TextEntityAlignment.MIDDLE_CENTER)


def dibujar_caja_derivacion_solida(espacio, pt, escala):
    cx, cy = pt
    r = 0.095 * escala
    hatch = espacio.add_hatch(color=7, dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    hatch.paths.add_polyline_path(
        [(cx - r, cy, 1.0), (cx + r, cy, 1.0)],
        is_closed=True
    )
    espacio.add_circle((cx, cy), radius=r, dxfattribs={"layer": "SIMBOLOS_BLANCO"})


def dibujar_enchufe_solido_con_patita(espacio, pt_muro, pared, escala):
    cx, cy = pt_muro
    r = 0.19 * escala
    largo_patita = 0.11 * escala

    direcciones = {
        "ARRIBA": (0.0, -1.0),
        "ABAJO":  (0.0,  1.0),
        "IZQ":    (1.0,  0.0),
        "DER":   (-1.0,  0.0),
    }
    nx, ny = direcciones.get(pared, (0.0, 1.0))
    tx, ty = ny, -nx

    p1 = (cx - r * tx, cy - r * ty)
    p2 = (cx + r * tx, cy + r * ty)

    hatch = espacio.add_hatch(color=7, dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    hatch.paths.add_polyline_path(
        [(p2[0], p2[1], 1.0), (p1[0], p1[1], 0.0)],
        is_closed=True
    )
    espacio.add_lwpolyline(
        [(p2[0], p2[1], 0, 0, 1.0), (p1[0], p1[1], 0, 0, 0.0)],
        close=True,
        dxfattribs={"layer": "SIMBOLOS_BLANCO"}
    )

    punta_domo = (cx + r * nx, cy + r * ny)
    fin_patita = (cx + (r + largo_patita) * nx, cy + (r + largo_patita) * ny)
    espacio.add_line(punta_domo, fin_patita, dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    return punta_domo


def dibujar_arco_suave(espacio, p1, p2, signo=1.0, capa="CANALIZACION_CIAN"):
    x1, y1 = p1
    x2, y2 = p2
    dx, dy = x2 - x1, y2 - y1
    d = math.hypot(dx, dy)
    if d < 1e-4:
        return

    mx, my = (x1 + x2) / 2.0, (y1 + y2) / 2.0
    nx, ny = -dy / d, dx / d
    h = d * 1.15 * signo
    cx, cy = mx + h * nx, my + h * ny
    radio = math.hypot(x1 - cx, y1 - cy)

    ang1 = math.degrees(math.atan2(y1 - cy, x1 - cx))
    ang2 = math.degrees(math.atan2(y2 - cy, x2 - cx))

    if (ang2 - ang1) % 360.0 < 180.0:
        start_a, end_a = ang1, ang2
    else:
        start_a, end_a = ang2, ang1

    espacio.add_arc((cx, cy), radius=radio, start_angle=start_a, end_angle=end_a, dxfattribs={"layer": capa})


def dibujar_simbolo_tierra_ttp(espacio, cx, y_top, escala):
    w_bar = 0.32 * escala
    espacio.add_line((cx - w_bar, y_top), (cx + w_bar, y_top), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    for i in range(5):
        x_ini = cx - w_bar + (i + 0.5) * (2 * w_bar / 5.0)
        espacio.add_line(
            (x_ini, y_top),
            (x_ini - 0.16 * escala, y_top - 0.20 * escala),
            dxfattribs={"layer": "SIMBOLOS_BLANCO"}
        )


def dibujar_simbolo_tierra_tts(espacio, cx, y_top, escala):
    for i, w_h in enumerate([0.32, 0.23, 0.15, 0.07]):
        yi = y_top - (i * 0.065) * escala
        espacio.add_line((cx - w_h * escala, yi), (cx + w_h * escala, yi), dxfattribs={"layer": "SIMBOLOS_BLANCO"})


def dibujar_conjunto_empalme_medidor_tierras(espacio, mx, my, escala, cli, p_salida_tda=None):
    def S(v):
        return v * escala

    def P(dx, dy):
        return (mx + dx * escala, my + dy * escala)

    def txt_izq(dx, dy, texto, h=0.145, capa="TEXTOS_SEC"):
        espacio.add_text(
            str(texto),
            dxfattribs={"layer": capa, "height": S(h)}
        ).set_placement(P(dx, dy), align=TextEntityAlignment.MIDDLE_LEFT)

    if p_salida_tda is not None:
        p_codo_sup = (mx + S(2.6), p_salida_tda[1])
        p_tope_medidor = P(0.0, 0.70)
        espacio.add_line(p_salida_tda, p_codo_sup, dxfattribs={"layer": "ALIMENTADOR_ROJO"})
        espacio.add_line(p_codo_sup, p_tope_medidor, dxfattribs={"layer": "ALIMENTADOR_ROJO"})

        txt_izq(1.35, 2.05, "Alimentador 01", h=0.14)
        espacio.add_line(P(1.35, 1.93), P(2.55, 1.93), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
        txt_izq(1.35, 1.76, cli["alimentador_texto_1"], h=0.13)
        txt_izq(1.35, 1.56, cli["alimentador_texto_2"], h=0.13)
        txt_izq(1.35, 1.36, cli["alimentador_texto_3"], h=0.13)

    # Cajón superior e inferior del Medidor "M"
    dibujar_rectangulo(espacio, *P(-0.35, 0.0), *P(0.35, 0.70), "SIMBOLOS_BLANCO")
    espacio.add_text("M", dxfattribs={"layer": "TEXTOS_SEC", "height": S(0.42)}).set_placement(
        P(0.0, 0.35), align=TextEntityAlignment.MIDDLE_CENTER
    )
    dibujar_rectangulo(espacio, *P(-0.35, -0.70), *P(0.35, 0.0), "SIMBOLOS_BLANCO")
    dibujar_rectangulo(espacio, *P(-0.27, -0.62), *P(0.27, -0.08), "SIMBOLOS_BLANCO")
    espacio.add_line(P(-0.27, -0.13), P(0.22, -0.62), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(-0.22, -0.08), P(0.27, -0.57), dxfattribs={"layer": "SIMBOLOS_BLANCO"})

    # Tierra izquierda (Verde)
    espacio.add_line(P(-0.35, -0.35), P(-2.30, -0.35), dxfattribs={"layer": "TDA_VERDE"})
    espacio.add_circle(P(-2.30, -0.35), radius=S(0.10), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(-2.40, -0.35), P(-3.15, -0.35), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(-3.15, -0.35), P(-3.15, -0.58), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    txt_izq(-4.85, -0.25, "Camara de registro", h=0.15)
    txt_izq(-4.55, -0.46, "PVC 160mm Ø", h=0.15)

    espacio.add_line(P(-2.30, -0.45), P(-2.30, -0.65), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(-3.15, -0.58), P(-2.30, -0.58), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    dibujar_simbolo_tierra_ttp(espacio, mx - S(3.15), my - S(0.58), escala)
    dibujar_simbolo_tierra_tts(espacio, mx - S(2.30), my - S(0.65), escala)

    txt_izq(-4.75, -0.95, "Ttp", h=0.16)
    txt_izq(-2.15, -0.95, "Tts", h=0.16)
    txt_izq(-4.75, -1.20, "conductor 5,26 mm² coviflex", h=0.15)
    txt_izq(-4.65, -1.42, "para ambas tierras", h=0.15)
    txt_izq(-4.75, -1.64, "t.p.r.e 20 mm2", h=0.15)
    txt_izq(-4.75, -1.86, "Resistencia tierra: 17 Ohms", h=0.15)
    txt_izq(-4.75, -2.08, "instrumento ETCR3000B", h=0.15)

    # Tierra derecha (Cián + Barra Copperweld)
    espacio.add_line(P(0.35, -0.38), P(2.60, -0.38), dxfattribs={"layer": "CANALIZACION_CIAN"})
    espacio.add_line(P(2.60, -0.38), P(2.60, -1.35), dxfattribs={"layer": "CANALIZACION_CIAN"})
    espacio.add_circle(P(2.60, -1.35), radius=S(0.09), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(2.51, -1.35), P(2.69, -1.35), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(2.60, -1.44), P(2.60, -1.26), dxfattribs={"layer": "SIMBOLOS_BLANCO"})

    espacio.add_line(P(2.05, -1.35), P(2.51, -1.35), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(2.69, -1.35), P(3.00, -1.35), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(2.05, -1.35), P(2.05, -1.55), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(3.00, -1.35), P(3.00, -1.55), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    dibujar_simbolo_tierra_ttp(espacio, mx + S(2.05), my - S(1.55), escala)
    dibujar_simbolo_tierra_tts(espacio, mx + S(3.00), my - S(1.55), escala)

    espacio.add_line(P(2.72, -1.28), P(2.90, -1.18), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(2.90, -1.18), P(4.20, -1.18), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    txt_izq(3.25, -0.88, "Camara de registro", h=0.125)
    txt_izq(3.40, -1.06, "PVC 160mm Ø", h=0.125)
    txt_izq(3.40, -1.32, "1 Barra Copperweld", h=0.125)
    txt_izq(3.40, -1.50, '1,5 m x 5/8"', h=0.125)

    txt_izq(1.90, -1.90, "Ttp", h=0.135)
    txt_izq(3.55, -1.90, "Tts", h=0.135)
    txt_izq(1.90, -2.10, "conductor 5,26 mm² coviflex para ambas tierras", h=0.13)
    txt_izq(1.90, -2.28, "t.a.g 20mm2", h=0.13)
    txt_izq(1.90, -2.46, "Resistencia tierra: 16 Ohms", h=0.13)
    txt_izq(1.90, -2.64, "Instrumento ECTR3000B", h=0.13)

    # Acometida y Poste Distribuidora
    espacio.add_line(P(0.0, -0.70), P(0.0, -2.35), dxfattribs={"layer": "ALIMENTADOR_ROJO"})
    dibujar_rectangulo(espacio, *P(-0.14, -3.05), *P(0.14, -2.35), "SIMBOLOS_BLANCO")
    txt_izq(-1.45, -1.25, "Acometida", h=0.15)
    txt_izq(-1.55, -1.46, cli["acometida_largo"], h=0.15)
    txt_izq(-2.15, -1.67, cli["acometida_cable"], h=0.15)
    txt_izq(-3.15, -2.82, "POSTE DISTRIBUIDORA", h=0.19)
    txt_izq(-3.15, -3.08, cli["poste_red"], h=0.14)


# =====================================================================
# 3. TABLAS VECTORIALES SEC Y ESQUEMA DE POSTACIÓN RURAL
# =====================================================================
def dibujar_tablas_y_postacion_cad(espacio, ox, oy, escala, proyecto, calc):
    cli = proyecto["cliente"]

    def S(v):
        return v * escala

    def P(x, y):
        return (ox + x * escala, oy + y * escala)

    def texto_celda(x1, y1, x2, y2, txt, h=0.17, capa="TEXTOS_SEC"):
        if txt == "" or txt is None:
            return
        cx, cy = (x1 + x2) / 2.0, (y1 + y2) / 2.0
        espacio.add_text(
            str(txt),
            dxfattribs={"layer": capa, "height": S(h)}
        ).set_placement(P(cx, cy), align=TextEntityAlignment.MIDDLE_CENTER)

    def texto_celda_doble(x1, y1, x2, y2, linea1, linea2, h=0.165, capa="TEXTOS_SEC"):
        cx, cy = (x1 + x2) / 2.0, (y1 + y2) / 2.0
        sep = 0.16
        espacio.add_text(str(linea1), dxfattribs={"layer": capa, "height": S(h)}).set_placement(
            P(cx, cy + sep), align=TextEntityAlignment.MIDDLE_CENTER
        )
        espacio.add_text(str(linea2), dxfattribs={"layer": capa, "height": S(h)}).set_placement(
            P(cx, cy - sep), align=TextEntityAlignment.MIDDLE_CENTER
        )

    anchos_col = [1.3, 1.8, 2.4, 1.6, 1.8, 1.9, 1.8, 1.8, 1.6, 1.6, 1.6, 1.9, 2.8, 2.3, 3.2, 5.8]
    x_cols = [0.0]
    for w in anchos_col:
        x_cols.append(x_cols[-1] + w)
    ancho_tabla = x_cols[-1]

    y_top = 15.0
    y_r1 = y_top - 0.75
    y_r2 = y_r1 - 1.25
    y_r3 = y_r2 - 0.90

    altos_datos = [1.05, 1.05] + [0.52] * 8
    y_filas = [y_r3]
    for h_f in altos_datos:
        y_filas.append(y_filas[-1] - h_f)
    y_bot = y_filas[-1]

    dibujar_rectangulo(espacio, *P(0, y_bot), *P(ancho_tabla, y_top), "MUROS_CIAN")
    espacio.add_line(P(0, y_r1), P(ancho_tabla, y_r1), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    texto_celda(0, y_r1, ancho_tabla, y_top, "CUADRO DE CARGAS DE ALUMBRADO", h=0.34)

    for yf in y_filas:
        espacio.add_line(P(0, yf), P(ancho_tabla, yf), dxfattribs={"layer": "SIMBOLOS_BLANCO"})

    espacio.add_line(P(x_cols[1], y_r2), P(x_cols[5], y_r2), dxfattribs={"layer": "SIMBOLOS_BLANCO"})
    espacio.add_line(P(x_cols[8], y_r2), P(x_cols[15], y_r2), dxfattribs={"layer": "SIMBOLOS_BLANCO"})

    cols_completas_r1 = {0, 1, 2, 3, 4, 5, 6, 7, 8, 11, 13, 15, 16}
    for idx_c, xc in enumerate(x_cols):
        y_ini = y_r1 if idx_c in cols_completas_r1 else y_r2
        espacio.add_line(P(xc, y_ini), P(xc, y_bot), dxfattribs={"layer": "SIMBOLOS_BLANCO"})

    texto_celda(x_cols[0], y_r3, x_cols[1], y_r1, "CTO N°", h=0.18)
    texto_celda(x_cols[1], y_r2, x_cols[2], y_r1, "FOCO LED", h=0.17)
    texto_celda_doble(x_cols[2], y_r2, x_cols[3], y_r1, "ENCHUFE", "CALEFACCION", h=0.16)
    texto_celda(x_cols[3], y_r2, x_cols[4], y_r1, "HORNO", h=0.17)
    texto_celda(x_cols[4], y_r2, x_cols[5], y_r1, "ENCHUFE", h=0.17)
    texto_celda_doble(x_cols[5], y_r3, x_cols[6], y_r1, "TOTAL", "CENTROS", h=0.17)
    texto_celda(x_cols[6], y_r3, x_cols[7], y_r1, "POT W.", h=0.18)
    texto_celda(x_cols[7], y_r3, x_cols[8], y_r1, "FASE (A)", h=0.18)
    texto_celda(x_cols[8], y_r2, x_cols[11], y_r1, "CORRIENTE POR FASE (A)", h=0.18)
    texto_celda(x_cols[11], y_r2, x_cols[13], y_r1, "PROTECCIONES", h=0.18)
    texto_celda(x_cols[13], y_r2, x_cols[15], y_r1, "CANALIZACIÓN", h=0.18)
    texto_celda(x_cols[15], y_r3, x_cols[16], y_r1, "UBICACIÓN", h=0.19)

    sub_headers = [
        (1, f"{cli['pot_foco_led_w']} W"),
        (2, f"{cli['pot_calefaccion_w']} W"),
        (3, f"{cli['pot_horno_w']} W"),
        (4, f"{cli['pot_enchufe_w']} W"),
        (8, "R"), (9, "S"), (10, "T"),
        (11, "DISY"), (12, "DIF"),
        (13, "COND. MM2"), (14, "DUCTO MM")
    ]
    for idx_c, txt_sub in sub_headers:
        texto_celda(x_cols[idx_c], y_r3, x_cols[idx_c + 1], y_r2, txt_sub, h=0.17)

    filas_valores = [
        ["1", "", "", "", str(calc["tot_ench"]), str(calc["tot_ench"]), f"{calc['pot_e']:,}", "R", f"{calc['cor_e']:.2f}", "", "", calc["disy_ench"], "2X25A 30mA", "EVA 2,5", "BANDEJA 20X10", "DOBLE_UBIC"],
        ["2", str(calc["tot_luces"]), "", "", "", str(calc["tot_luces"]), f"{calc['pot_l']:,}", "R", f"{calc['cor_l']:.2f}", "", "", "1X10A", "2X25A 30mA", "EVA 1,5", "BANDEJA 20X10", "DOBLE_UBIC"],
        ["3", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["4", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["5", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["6", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["7", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["8", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["9", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""],
        ["TOTAL", str(calc["tot_luces"]), "0", "0", str(calc["tot_ench"]), str(calc["tot_centros"]), f"{calc['pot_tot_w']:,}", f"{calc['cor_tot_r']:.2f}", f"{calc['cor_tot_r']:.2f}", "0.00", "0.00", "", "", "", "", ""]
    ]

    for idx_f, fila in enumerate(filas_valores):
        y_sup = y_filas[idx_f]
        y_inf = y_filas[idx_f + 1]
        for idx_c, val in enumerate(fila):
            if val == "DOBLE_UBIC":
                texto_celda_doble(
                    x_cols[idx_c], y_inf, x_cols[idx_c + 1], y_sup,
                    proyecto["ubicacion_linea1"],
                    proyecto["ubicacion_linea2"],
                    h=0.165
                )
            else:
                texto_celda(x_cols[idx_c], y_inf, x_cols[idx_c + 1], y_sup, val, h=0.18)

    # Recuadros de Cálculo Caída de Tensión y Alimentador
    y_calc_top = y_bot - 0.9
    y_calc_bot = y_calc_top - 4.2
    dibujar_rectangulo(espacio, *P(0.0, y_calc_bot), *P(9.5, y_calc_top), "SIMBOLOS_BLANCO")
    dibujar_rectangulo(espacio, *P(10.2, y_calc_bot), *P(20.5, y_calc_top), "SIMBOLOS_BLANCO")

    lineas_vp = [
        "CALCULO CAIDA TENSION",
        "VP = (2 x L x I) / (56 x S)",
        f"VP = (2 x {calc['L']:.0f} x {calc['cor_tot_r']:.2f}) / (56 x {calc['S_cond']:.0f})",
        f"VP = {calc['num_vp']} / {calc['den_vp']}",
        f"VP = {calc['vp_val']} V."
    ]
    for i, l_txt in enumerate(lineas_vp):
        texto_celda(0.2, y_calc_top - (i + 1) * 0.80, 9.3, y_calc_top - i * 0.80, l_txt, h=0.21)

    lineas_s = [
        "CALCULO ALIMENTADOR",
        "S = (2 x L x rho x I) / vp",
        f"S = (2 x {calc['L']:.0f} x 0,018 x {calc['cor_tot_r']:.2f}) / 6,6",
        f"S = {calc['num_s']} / 6,6",
        f"S = {calc['s_val']} mm2."
    ]
    for i, l_txt in enumerate(lineas_s):
        texto_celda(10.4, y_calc_top - (i + 1) * 0.80, 20.3, y_calc_top - i * 0.80, l_txt, h=0.21)


def generar_dxf_proyecto(proyecto=None, ruta_plantilla=None):
    """Genera el objeto doc de ezdxf y el diccionario de cálculos listo para guardar o descargar."""
    if proyecto is None:
        proyecto = PROYECTO_DEFAULT

    calc = calcular_metricas_sec(proyecto)
    cli = proyecto["cliente"]

    if ruta_plantilla and os.path.exists(ruta_plantilla):
        doc = ezdxf.readfile(ruta_plantilla)
    else:
        doc = ezdxf.new("R2018")

    msp = doc.modelspace()

    for nombre_capa, color_aci in [
        ("MUROS_CIAN", 4),
        ("CANALIZACION_CIAN", 4),
        ("SIMBOLOS_AMARILLO", 2),
        ("SIMBOLOS_BLANCO", 7),
        ("TDA_VERDE", 3),
        ("ALIMENTADOR_ROJO", 1),
        ("TEXTOS_SEC", 7)
    ]:
        if nombre_capa not in doc.layers:
            doc.layers.add(name=nombre_capa, color=color_aci)

    # Reemplazar textos en viñeta si se cargó plantilla anterior
    reemplazos = {
        "francisco bilbao 1916": cli["direccion"],
        "FRANCISCO BILBAO 1916": cli["direccion"].upper(),
        "Calle: francisco bilbao 1916": f"Sector: {cli['direccion']}",
        "Comuna:Osorno": f"Comuna: {cli['comuna']}",
        "JORGE BARRIA OBREQUE": cli["propietario"],
        "14.095.710-6": cli["rut"]
    }
    for esp in [msp] + [blk for blk in doc.blocks if not blk.name.startswith("*")]:
        for ent in esp.query("TEXT MTEXT"):
            es_mtext = ent.dxftype() == "MTEXT"
            txt = ent.text if es_mtext else ent.dxf.text
            if not txt:
                continue
            nuevo = txt
            for k, v in reemplazos.items():
                if k in nuevo:
                    nuevo = nuevo.replace(k, v)
            if nuevo != txt:
                if es_mtext:
                    ent.text = nuevo
                else:
                    ent.dxf.text = nuevo

    escala = 1.0
    ox, oy = 0.0, 0.0
    try:
        ext = bbox.extents(msp, fast=True)
        if ext.has_data:
            alto_actual = ext.size.y
            if alto_actual > 0:
                escala = max(alto_actual / 32.0, 0.5)
            ox = ext.extmax.x + (4.0 * escala)
            oy = ext.extmin.y + (2.0 * escala)
    except Exception:
        escala = 10.0

    def S(val):
        return val * escala

    def P(x, y):
        return (ox + x * escala, oy + y * escala)

    config_plantas = [
        (17.5, "PLANTA ENCHUFES", "ENCHUFES", 1),
        (0.0,  "PLANTA ILUMINACION", "LUZ", 2)
    ]

    for desp_y, titulo, modo, num_cto in config_plantas:
        msp.add_text(titulo, dxfattribs={"layer": "TEXTOS_SEC", "height": S(0.55)}).set_placement(
            P(5.0, desp_y + 12.8), align=TextEntityAlignment.MIDDLE_CENTER
        )

        tx, ty = proyecto["tda"][0], desp_y + proyecto["tda"][1]
        px_tda, py_tda = P(tx, ty)
        dibujar_tda_solido(msp, px_tda, py_tda, S(0.28), S(0.55), num_cto, escala)

        if modo == "ENCHUFES":
            p_salida_derecha_tda = P(tx + 0.28, ty + 0.27)
            dibujar_conjunto_empalme_medidor_tierras(
                msp,
                mx=ox + 16.0 * escala,
                my=oy + (desp_y + 4.2) * escala,
                escala=escala,
                cli=cli,
                p_salida_tda=p_salida_derecha_tda
            )

        caja_previa = P(tx, ty + 0.27)

        for hab in proyecto["habitaciones"]:
            hx, hy, hw, hh = hab["x"], desp_y + hab["y"], hab["w"], hab["h"]
            dibujar_rectangulo(msp, *P(hx, hy), *P(hx + hw, hy + hh), "MUROS_CIAN")
            dibujar_rectangulo(msp, *P(hx + ESP_MURO, hy + ESP_MURO), *P(hx + hw - ESP_MURO, hy + hh - ESP_MURO), "MUROS_CIAN")

            for vx1, vy1, vx2, vy2 in hab["ventanas"]:
                msp.add_line(P(vx1, desp_y + vy1), P(vx2, desp_y + vy2), dxfattribs={"layer": "SIMBOLOS_BLANCO"})

            msp.add_text(hab["nombre"], dxfattribs={"layer": "TEXTOS_SEC", "height": S(0.24)}).set_placement(
                P(hx + hw / 2.0, hy + hh - 0.55), align=TextEntityAlignment.MIDDLE_CENTER
            )

            caja_actual = P(hab["caja"][0], desp_y + hab["caja"][1])
            dibujar_caja_derivacion_solida(msp, caja_actual, escala)
            dibujar_arco_suave(msp, caja_previa, caja_actual, signo=1.0, capa="CANALIZACION_CIAN")
            caja_previa = caja_actual

            if modo == "LUZ":
                for lx_rel, ly_rel in hab["luces"]:
                    pt_luz = P(lx_rel, desp_y + ly_rel)
                    r_l, d_l = S(0.25), S(0.175)
                    msp.add_circle(pt_luz, radius=r_l, dxfattribs={"layer": "SIMBOLOS_AMARILLO"})
                    msp.add_line((pt_luz[0] - d_l, pt_luz[1] - d_l), (pt_luz[0] + d_l, pt_luz[1] + d_l), dxfattribs={"layer": "SIMBOLOS_AMARILLO"})
                    msp.add_line((pt_luz[0] - d_l, pt_luz[1] + d_l), (pt_luz[0] + d_l, pt_luz[1] - d_l), dxfattribs={"layer": "SIMBOLOS_AMARILLO"})
                    dibujar_arco_suave(msp, caja_actual, pt_luz, signo=1.0, capa="CANALIZACION_CIAN")

                for ix_rel, iy_rel in hab["interruptores"]:
                    pt_int = P(ix_rel, desp_y + iy_rel)
                    msp.add_circle(pt_int, radius=S(0.08), dxfattribs={"layer": "SIMBOLOS_AMARILLO"})
                    msp.add_line(pt_int, (pt_int[0] - S(0.25), pt_int[1]), dxfattribs={"layer": "SIMBOLOS_AMARILLO"})
                    msp.add_line((pt_int[0] - S(0.25), pt_int[1]), (pt_int[0] - S(0.25), pt_int[1] + S(0.12)), dxfattribs={"layer": "SIMBOLOS_AMARILLO"})
                    dibujar_arco_suave(msp, caja_actual, pt_int, signo=-1.0, capa="CANALIZACION_CIAN")
            else:
                for ex_rel, ey_rel, pared in hab["enchufes"]:
                    pt_muro = P(ex_rel, desp_y + ey_rel)
                    punta_domo = dibujar_enchufe_solido_con_patita(msp, pt_muro, pared, escala)
                    dibujar_arco_suave(msp, caja_actual, punta_domo, signo=-1.0, capa="CANALIZACION_CIAN")

    dibujar_tablas_y_postacion_cad(msp, ox + 23.0 * escala, oy + 10.0 * escala, escala, proyecto, calc)

    if "EMPALME_MEDIDOR_TIERRAS" not in doc.blocks:
        blk_m = doc.blocks.new(name="EMPALME_MEDIDOR_TIERRAS")
        dibujar_conjunto_empalme_medidor_tierras(blk_m, 0.0, 0.0, escala, cli, p_salida_tda=None)

    try:
        for vp in doc.viewports.get("*Active"):
            vp.dxf.center = (ox + 12.0 * escala, oy + 16.0 * escala)
            vp.dxf.height = 36.0 * escala
    except Exception:
        pass

    return doc, calc