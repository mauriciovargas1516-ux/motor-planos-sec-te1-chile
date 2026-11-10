#  Motor Paramétrico de Planos Eléctricos SEC (TE1) — Chile

![Python](https://img.shields.io/badge/Python-3.10%2B-blue?logo=python&logoColor=white)
![AutoCAD DXF](https://img.shields.io/badge/CAD-ezdxf%20R2018-red?logo=autodesk&logoColor=white)
![Streamlit](https://img.shields.io/badge/UI-Streamlit%20%2B%20Tkinter-ff4b4b?logo=streamlit&logoColor=white)
![Normativa SEC](https://img.shields.io/badge/Normativa-SEC%20RIC%20%7C%20D.S.%2092%2F83-success)

Motor de ingeniería en Python para la **generación automatizada de planos eléctricos interiores y rurales (.dxf)**, cálculo de cuadros de cargas de alumbrado, dimensionamiento de alimentadores y clasificación automática de trámites **TE1** ante la Superintendencia de Electricidad y Combustibles (**SEC - e-Declarador**).

---

##  Características Principales

* **Doble Planta Arquitectónica y Eléctrica Automatizada (`ezdxf`):** Genera simultáneamente la **Planta de Enchufes (Cto. 1)** y la **Planta de Iluminación (Cto. 2)** en capas estandarizadas de AutoCAD (`MUROS_CIAN`, `CANALIZACION_CIAN`, `SIMBOLOS_AMARILLO`, `SIMBOLOS_BLANCO`, `TDA_VERDE`, `ALIMENTADOR_ROJO`).
* **Simbología Normativa Vectorial con Relleno Sólido (`HATCH`):**
  * Tomas de corriente orientadas automáticamente hacia el interior de cada recinto según el muro de apoyo (`ARRIBA`, `ABAJO`, `IZQ`, `DER`).
  * Cajas de derivación sólidas, centros de iluminación (`⊗`), interruptores y tablero de distribución (`TDA`) unidos mediante arcos suaves de canalización.
* **Conjunto Completo de Empalme y Doble Puesta a Tierra:** Dibuja automáticamente el medidor `M`, acometida desde poste de distribución, cámaras de registro `PVC 160mm Ø`, barra Copperweld de `1,5 m x 5/8"` y tierras de protección/servicio (`Ttp` y `Tts`) con sus respectivas mediciones de resistencia.
* **Cálculo Matemático y Cuadro de Cargas Nativo en CAD:** Construye la grilla vectorial completa del **Cuadro de Cargas de Alumbrado** (sin depender de objetos OLE de Excel) y resuelve las ecuaciones normativas de caída de tensión y sección mínima del conductor:

$$VP = \frac{2 \times L \times I}{56 \times S} \qquad \text{y} \qquad S = \frac{2 \times L \times \rho \times I}{vp_{\text{máx}}}$$

* **Clasificador Automático D.S. N° 92/83 (Paso 3 e-Declarador):** Evalúa la potencia instalada (kW) y la longitud del alimentador ($L$) para determinar automáticamente si la declaración corresponde a **Tipo F**, **Tipo D** ($\le 100\text{ m}$ y $\le 10\text{ kW}$) o **Tipo B** (alimentadores rurales $> 100\text{ m}$ hasta $500\text{ kW}$).

---

##  Caso de Estudio Precargado (Empalme Rural 120 m)

El repositorio incluye precargado un levantamiento real de 7 recintos con alimentador aéreo en cable preensamblado:

| Parámetro Técnico | Valor Calculado | Criterio Normativo SEC |
| :--- | :--- | :--- |
| **Circuito 1 (Enchufes)** | 14 centros × 250 W = **3,500 W** (**15.91 A**) | Disyuntor `1X16A` / Diferencial `2X25A 30mA` / `EVA 2,5 mm²` |
| **Circuito 2 (Iluminación)** | 11 centros × 80 W = **880 W** (**4.00 A**) | Disyuntor `1X10A` / Diferencial `2X25A 30mA` / `EVA 1,5 mm²` |
| **Total Instalación** | **25 centros** — **4,380 W** (**19.91 A**) | Empalme monofásico 220 V |
| **Caída de Tensión ($VP$)** | **3.413 V** ($L = 120\text{ m}$, $S = 25\text{ mm}^2$) | Cumple $VP \le 6.6\text{ V}$ (3% en Baja Tensión) |
| **Sección Teórica ($S$)** | **13.032 mm²** | Se proyecta Preensamblado `2x25 mm²` |
| **Clasificación D.S. N° 92/83** | **Tipo B** | Supera límite de 100 m de alimentador del Tipo D |

---

##  Instalación y Ejecución (Modo Dual)

1. Clonar el repositorio e instalar dependencias:
```bash
git clone [https://github.com/mauriciovargas1516-ux/motor-planos-sec-te1-chile.git](https://github.com/mauriciovargas1516-ux/motor-planos-sec-te1-chile.git)
cd motor-planos-sec-te1-chile
pip install -r requirements.txt
