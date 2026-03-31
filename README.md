# automatizacion-planeacion-vpp

[![Python 3.9+](https://img.shields.io/badge/python-3.9%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Tests](https://img.shields.io/badge/tests-81%20passed-brightgreen.svg)]()
[![AI Assisted](https://img.shields.io/badge/AI%20Assisted-Claude%20%7C%20Gemini-blue)]()

**Automatización del seguimiento de planeación estratégica de ProColombia.**

Automatiza el ciclo completo de seguimiento trimestral de la Hoja de Ruta: genera matrices Excel para captura, convierte los datos en presentaciones PowerPoint institucionales y consolida la información de todas las unidades en un Excel maestro para análisis gerencial.

---

> **English summary:** Python package that automates ProColombia's quarterly strategic planning follow-up. Generates Excel capture matrices for organizational units, converts filled data into institutional PowerPoint presentations using template markers, and consolidates all units into a master Excel for management analysis. Supports 3 unit families (Misional, Territorial, Transversal), 81+ automated tests, Google Colab optimized.

---

## Tabla de contenidos

1. [¿Para quién es este paquete?](#1-para-quién-es-este-paquete)
2. [Instalación](#2-instalación)
3. [Inicio rápido](#3-inicio-rápido)
4. [Las tres familias de unidades](#4-las-tres-familias-de-unidades)
5. [Estructura de carpetas](#5-estructura-de-carpetas)
6. [Operaciones disponibles](#6-operaciones-disponibles)
7. [Arquitectura del paquete](#7-arquitectura-del-paquete)
8. [Configuración personalizada](#8-configuración-personalizada)
9. [Flujo completo paso a paso](#9-flujo-completo-paso-a-paso)
10. [Errores comunes y solución](#10-errores-comunes-y-solución)
11. [Tests automatizados](#11-tests-automatizados)
12. [FAQ — Preguntas frecuentes](#12-faq--preguntas-frecuentes)
13. [Cómo citar](#13-cómo-citar)
14. [Créditos y agradecimientos](#14-créditos-y-agradecimientos)
15. [Metodología de desarrollo](#15-metodología-de-desarrollo)
16. [Licencia](#16-licencia)

---

## 1. ¿Para quién es este paquete?

Este paquete es para ti si:

- Eres parte del equipo de **planeación estratégica de ProColombia** y necesitas automatizar el seguimiento trimestral de la Hoja de Ruta.
- Eres **analista o coordinador de gestión** y quieres consolidar la información de múltiples unidades sin procesar manualmente decenas de archivos Excel.
- Necesitas generar **presentaciones institucionales estandarizadas** (PPTX) a partir de datos estructurados.
- Usas **Google Colab** y no quieres instalar nada en tu computador, aunque lo puedes usar localmente.

**No necesitas:** experiencia avanzada en programación. Saber ejecutar celdas en Google Colab es suficiente para operar el sistema completo.

---

## 2. Instalación

### Opción A — Google Colab (recomendado)

```python
from google.colab import drive
drive.mount('/content/drive')
!pip install python-pptx openpyxl -q

import sys
RUTA = "/content/drive/MyDrive/ProColombia/Automatizaciones/VPP"
sys.path.insert(0, RUTA)

from procolombia import *
banner()
```

### Opción B — Desde GitHub

```bash
pip install git+https://github.com/enriqueforero/procolombia-vpp.git
```

### Requisitos del sistema

| Dependencia | Versión mínima | Uso |
|---|---|---|
| Python | 3.9+ | Tipado moderno, f-strings, pathlib |
| pandas | 1.5+ | Lectura de Excel y consolidación |
| openpyxl | 3.0+ | Creación y lectura de archivos .xlsx |
| python-pptx | 0.6+ | Creación y manipulación de archivos .pptx |

> Pandas y openpyxl ya vienen preinstalados en Google Colab. Solo es necesario instalar `python-pptx`.

---

## 3. Inicio rápido

### Procesar Excel y generar presentaciones (3 líneas)

```python
from procolombia import OrquestadorUniversal

orq = OrquestadorUniversal(base_dir=RUTA)
resultados = orq.procesar_lote()
```

El sistema detecta automáticamente la familia de cada Excel, selecciona el lector y generador PPTX correcto, reemplaza todos los marcadores `{{...}}`, elimina slides vacíos y guarda las presentaciones en `02_pptx_salida/`.

### Consolidar toda la información

```python
ruta_consolidado = orq.consolidar()
```

Genera un Excel maestro en `03_consolidado/` con hojas de resumen, DOFA, acciones propias y contribuciones de todas las unidades.

---

## 4. Las tres familias de unidades

El sistema reconoce y maneja tres familias de unidades organizacionales, cada una con estructura, hojas de Excel y plantilla PPTX diferente.

| Familia | Unidades que incluye | Plantilla PPTX |
|---|---|---|
| **MISIONAL** | VP Exportaciones, VP Inversión, VP Turismo, Marca País | `Plantilla_Misional.pptx` |
| **TERRITORIAL** | Hubs, Oficinas Comerciales (Oficom), Oficinas Regionales (OfiReg), FIDIREP | `Plantilla_Territorial.pptx` |
| **TRANSVERSAL** | Gerencias transversales (GIC, Comunicaciones, Informática, etc.) | `Plantilla_Transversal.pptx` |

La detección es automática: el sistema examina las hojas del Excel para determinar la familia sin intervención del usuario.

---

## 5. Estructura de carpetas

```
📁 VPP/  (directorio base)
├── 📁 procolombia/                  ← Paquete Python (NO modificar excepto config.py)
│   ├── __init__.py                  ← API pública (lo que se importa)
│   ├── config.py                    ← Configuración y enums
│   ├── utils.py                     ← Utilidades y estilos Excel
│   ├── excel_constructores.py       ← Generadores de matrices Excel
│   ├── excel_lectores.py            ← Lectores de Excel diligenciados
│   ├── pptx_gen.py                  ← Generadores y plantillas PPTX
│   ├── orquestador.py               ← OrquestadorUniversal (punto de entrada)
│   ├── ejemplos.py                  ← Datos de ejemplo, banner y guía
│   └── tests.py                     ← Tests automatizados (81 tests)
│
├── 📁 01_excels_entrada/            ← Excel diligenciados por las áreas
├── 📁 02_pptx_salida/               ← Presentaciones generadas
├── 📁 03_consolidado/               ← Excel maestro consolidado
├── 📁 04_plantillas/                ← Plantillas PPTX por familia
├── README.md
└── LICENSE
```

> Las subcarpetas de trabajo (`01_` a `04_`) se crean automáticamente al instanciar el orquestador. Los archivos de datos dentro de estas carpetas (`.xlsx`, `.pptx`) no se versionan en git — solo el código fuente del paquete y los archivos de documentación.

---

## 6. Operaciones disponibles

### 6.1 Construir plantillas PPTX

```python
orq = OrquestadorUniversal(base_dir=RUTA)
rutas = orq.construir_plantillas()
```

Genera las plantillas Territorial (~28 slides) y Transversal (~34 slides) programáticamente sin formato institucional. La plantilla Misional debe subirse manualmente (es el archivo institucional de ProColombia).

### 6.2 Generar Excel vacíos para las áreas

```python
# Individual
orq.generar_excel('VP Exportaciones', 'EJE', trimestre='Q1', anio='2026', num_lineas=3)

# En lote
UNIDADES = [
    ('VP Exportaciones', 'EJE', 3),
    ('VP Inversión', 'EJE', 3),
    ('VP Turismo', 'EJE', 4),
    ('Marca País', 'MARCA PAÍS', 4),
    ('Hub Norteamérica', 'HUB'),
    ('Gerencia Inteligencia Comercial', 'TRANSVERSAL', 3),
]
for u in UNIDADES:
    nombre, tipo = u[0], u[1]
    nl = u[2] if len(u) > 2 else 5
    orq.generar_excel(nombre, tipo, trimestre='Q1', anio='2026', num_lineas=nl)
```

### 6.3 Procesar lote completo

```python
orq = OrquestadorUniversal(base_dir=RUTA)
resultados = orq.procesar_lote()
```

### 6.4 Consolidar en Excel maestro

```python
ruta_consolidado = orq.consolidar()
```

---

## 7. Arquitectura del paquete

El paquete tiene 9 archivos organizados por responsabilidad única (SRP). Las dependencias fluyen en una dirección: `config` ← `utils` ← `excel`/`pptx_gen` ← `orquestador`.

| Módulo | Líneas | Responsabilidad |
|---|---|---|
| `config.py` | ~197 | Configuración centralizada, enums (`TipoUnidad`, `FamiliaUnidad`), ejes misionales |
| `utils.py` | ~120 | Logger, safe string, estilos Excel, medidor de tiempo |
| `excel_constructores.py` | ~1,458 | Generadores de matrices Excel (Misional, Territorial, Transversal) |
| `excel_lectores.py` | ~604 | Lectores de Excel diligenciados con `itertuples()` |
| `pptx_gen.py` | ~934 | Generadores PPTX + constructores de plantillas |
| `orquestador.py` | ~280 | Router inteligente, procesamiento por lotes, consolidación |
| `ejemplos.py` | ~554 | Datos de ejemplo para pruebas y demostraciones |
| `tests.py` | ~452 | Suite de 81 tests automatizados |
| `__init__.py` | ~46 | API pública (`__all__` con 11 símbolos) |

---

## 8. Configuración personalizada

La configuración se centraliza en la clase `Config`. Puede personalizarla sin modificar archivos del paquete:

```python
cfg = Config(
    max_lineas_estrategicas=4,
    max_acciones_por_linea=10,
    max_dofa_por_cuadrante=8,
    max_tendencias=5,
    password='mi_clave_2026',
)
orq = OrquestadorUniversal(config=cfg, base_dir=RUTA)
```

Todos los parámetros tienen valores por defecto razonables. Ver la documentación completa de parámetros en `procolombia/config.py`.

---

## 9. Flujo completo paso a paso

```
┌─────────────────────────────────────────────────────────────────┐
│  1. PREPARAR                                                    │
│     Subir módulos .py a Google Drive                            │
│     Subir Plantilla_Misional.pptx a 04_plantillas/              │
│     Ejecutar orq.construir_plantillas() para las otras 2        │
├─────────────────────────────────────────────────────────────────┤
│  2. GENERAR EXCEL                                               │
│     Ejecutar orq.generar_excel() por cada unidad                │
│     Los Excel se guardan en 01_excels_entrada/                  │
├─────────────────────────────────────────────────────────────────┤
│  3. DISTRIBUIR                                                  │
│     Enviar los Excel a las áreas para que los diligencien       │
├─────────────────────────────────────────────────────────────────┤
│  4. RECOPILAR                                                   │
│     Recibir los Excel diligenciados                             │
│     Colocarlos en 01_excels_entrada/                            │
├─────────────────────────────────────────────────────────────────┤
│  5. PROCESAR                                                    │
│     Ejecutar orq.procesar_lote()                                │
│     Las presentaciones se generan en 02_pptx_salida/            │
├─────────────────────────────────────────────────────────────────┤
│  6. CONSOLIDAR                                                  │
│     Ejecutar orq.consolidar()                                   │
│     El Excel maestro se genera en 03_consolidado/               │
├─────────────────────────────────────────────────────────────────┤
│  7. ENTREGAR                                                    │
│     Descargar PPTX de 02_pptx_salida/ y distribuir              │
└─────────────────────────────────────────────────────────────────┘
```

---

## 10. Errores comunes y solución

| Error | Causa | Solución |
|---|---|---|
| `Plantilla XXXX no encontrada` | La plantilla PPTX no existe en `04_plantillas/` | Misional: subir manualmente. Territorial/Transversal: ejecutar `orq.construir_plantillas()` |
| `Tipo de unidad 'XXX' no reconocido` | El campo "Tipo de unidad" en la hoja CONFIGURACIÓN no es válido | Usar uno de: `EJE`, `MARCA PAÍS`, `HUB`, `OFICOM`, `OFIREG`, `FIDIREP`, `TRANSVERSAL` |
| `No hay archivos .xlsx en 01_excels_entrada/` | Carpeta vacía o ruta incorrecta | Verificar que `base_dir` apunte al directorio correcto |
| Marcadores `{{...}}` sin reemplazar | Formato mixto dentro del marcador en la plantilla | Abrir la plantilla y aplicar formato uniforme al texto del marcador |

---

## 11. Tests automatizados

El sistema incluye una suite de 81 tests que cubren configuración, enums, generación de Excel, lectura roundtrip, detección de familia, consolidación, y verificación de que no se use `iterrows()`.

```python
%run procolombia/tests.py
```

---

## 12. FAQ — Preguntas frecuentes

**¿Puedo procesar archivos de las tres familias en una sola corrida?**
Sí. `procesar_lote()` detecta la familia de cada archivo automáticamente y aplica el lector, generador y plantilla correctos.

**¿Qué pasa si un archivo falla durante el procesamiento?**
Los errores se manejan individualmente. Si un archivo falla, los demás se procesan normalmente y el reporte final muestra cuáles tuvieron error.

**¿Puedo generar la plantilla Misional programáticamente?**
No. La plantilla Misional es el archivo institucional de ProColombia con diseño oficial. Debe subirse manualmente a `04_plantillas/`. Las plantillas Territorial y Transversal sí se generan programáticamente.

**¿Qué pasa con los slides de líneas estratégicas vacías?**
Se eliminan automáticamente. Si una unidad tiene 3 líneas diligenciadas de un máximo de 5, los slides de las líneas 4 y 5 se remueven de la presentación final.

---

## 13. Cómo citar

Si usas este paquete en un trabajo, reporte o publicación, puedes citarlo así:

**Formato BibTeX:**

```bibtex
@software{automatizacion-planeacion-vpp,
  author  = {Forero Herrera, Néstor Enrique},
  title   = {procolombia-vpp: Automatización de planeación estratégica de ProColombia},
  year    = {2026},
  version = {0.1.0},
  url     = {https://github.com/EnriqueForero/automatizacion-planeacion-vpp}
}
```

**Formato texto (APA):**

> Forero Herrera, N. E. (2026). *automatizacion-planeacion-vpp* (v0.1.0) [Software]. GitHub. https://github.com/EnriqueForero/automatizacion-planeacion-vpp

---

## 14. Créditos y agradecimientos

Este proyecto es resultado de un trabajo colaborativo entre el equipo de la Vicepresidencia de Planeación (VPP) de ProColombia y la Gerencia de Inteligencia Comercial (GIC).

**Andrea Molano** y **Walter Castaño**, del equipo de la VPP, fueron fundamentales en la concepción de este sistema. Su profundo conocimiento del proceso de seguimiento de la Hoja de Ruta, la ideación de los requerimientos funcionales, la definición de las reglas de negocio y el entendimiento detallado de las necesidades de cada familia de unidades hicieron posible traducir un proceso institucional complejo en una solución automatizada coherente y útil. Su acompañamiento durante todo el ciclo de desarrollo — desde la especificación inicial hasta la validación de los entregables — garantizó que el sistema respondiera fielmente a la realidad operativa de ProColombia.

**Sol** (practicante) contribuyó con una versión preliminar en Python que sirvió como punto de partida y referencia para explorar la viabilidad técnica de la automatización. Su trabajo inicial permitió identificar los retos clave del problema y orientar las decisiones de arquitectura de la versión actual.

**Néstor Enrique Forero Herrera** (GIC - Coordinación Analítica) fue responsable de la arquitectura del sistema, el diseño técnico, la implementación del paquete `procolombia/`, la suite de pruebas y la documentación.

---

## 15. Metodología de desarrollo

La arquitectura, la lógica de negocio y los requerimientos son de autoría humana. El autor asume responsabilidad total sobre la integridad del código publicado.

Se utilizaron modelos de IA generativa (Claude y Gemini) como asistencia técnica para generación de código boilerplate, optimización de consultas y sugerencias de refactorización.

Validación: Ningún bloque asistido por IA fue integrado sin revisión crítica, ajuste al contexto del dominio y validación mediante pruebas funcionales.

---

## 16. Licencia

MIT — Néstor Enrique Forero Herrera · Colombia · 2026