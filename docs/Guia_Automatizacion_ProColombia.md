# Guía de Uso — Automatización Planeación Estratégica ProColombia v5.1

---

## 1. ¿Qué es esta automatización?

Es un sistema de tres módulos en Python que automatiza el ciclo completo de seguimiento trimestral de la Hoja de Ruta de ProColombia. El sistema cubre tres operaciones principales:

1. **Generar matrices Excel** vacías (o pre-llenadas) para que las áreas reporten sus avances trimestrales.
2. **Procesar los Excel diligenciados** y convertirlos automáticamente en presentaciones PowerPoint (.pptx) usando plantillas institucionales con marcadores.
3. **Consolidar toda la información** de múltiples unidades en un Excel maestro para análisis gerencial.

El sistema reconoce y maneja tres familias de unidades organizacionales, cada una con estructura, hojas de Excel y plantilla PPTX diferente.

---

## 2. Las tres familias de unidades

| Familia | Unidades que incluye | Estructura del Excel | Plantilla PPTX |
|---|---|---|---|
| **MISIONAL** | VP Exportaciones, VP Inversión, VP Turismo, Marca País | Líneas estratégicas propias, DOFA, Tendencias, Casos de Éxito, Metas | `Plantilla_Misional.pptx` |
| **TERRITORIAL** | Hubs, Oficinas Comerciales (Oficom), Oficinas Regionales (OfiReg), FIDIREP | Contribuciones a los 4 ejes misionales, DOFA, Tendencias por Eje, Metas | `Plantilla_Territorial.pptx` |
| **TRANSVERSAL** | Gerencias transversales (GIC, Comunicaciones, Informática, etc.) | Líneas estratégicas propias + Contribuciones a los 4 ejes | `Plantilla_Transversal.pptx` |

### ¿Cómo se distingue cada familia?

El sistema detecta la familia automáticamente examinando las hojas del Excel:

- Si tiene hojas `MP_LE1` (contribuciones) **y** `LÍNEA ESTRATÉGICA 1` (líneas propias) → **TRANSVERSAL**
- Si tiene `MP_LE1` pero **no** `LÍNEA ESTRATÉGICA 1` → **TERRITORIAL**
- Si no tiene `MP_LE1` → **MISIONAL**

También se puede identificar por el campo "Tipo de unidad" en la hoja CONFIGURACIÓN:

| Valor del campo | Familia asignada |
|---|---|
| `EJE`, `MARCA PAÍS`, `VICEPRESIDENCIA` | MISIONAL |
| `HUB`, `OFICOM`, `OFIREG`, `FIDIREP` | TERRITORIAL |
| `TRANSVERSAL` | TRANSVERSAL |

---

## 3. Estructura de carpetas

El sistema trabaja con un paquete Python (`procolombia/`) y cuatro subcarpetas de trabajo. Las subcarpetas se crean automáticamente al instanciar el orquestador.

```
📁 VPP/  (directorio base)
├── 📁 procolombia/                  ← Paquete Python (NO modificar excepto config.py)
│   ├── __init__.py                  ← API pública (lo que se importa)
│   ├── config.py                    ← Configuración y enums (LO QUE USTED TOCA)
│   ├── utils.py                     ← Utilidades y estilos Excel
│   ├── excel.py                     ← Constructores y lectores de Excel
│   ├── pptx_gen.py                  ← Generadores y plantillas PPTX
│   ├── orquestador.py               ← OrquestadorUniversal (punto de entrada)
│   └── ejemplos.py                  ← Datos de ejemplo, banner y guía
│
├── 📁 01_excels_entrada/            ← Aquí se colocan los Excel diligenciados
├── 📁 02_pptx_salida/               ← Aquí se generan las presentaciones
├── 📁 03_consolidado/               ← Aquí se genera el Excel maestro
└── 📁 04_plantillas/                ← Plantillas PPTX por familia
    ├── Plantilla_Misional.pptx      ← Se sube manualmente (institucional)
    ├── Plantilla_Territorial.pptx   ← Se genera con el sistema
    └── Plantilla_Transversal.pptx   ← Se genera con el sistema
```

> **Importante:** La configuración (límites, carpetas, contraseña, nombres de plantillas) se encuentra en `procolombia/config.py`, clase `Config`. Es el único archivo del paquete que debería necesitar editar. También puede personalizar los parámetros sin tocar el archivo, pasando un objeto `Config` al orquestador desde la celda del notebook.

---

## 4. Arquitectura del paquete `procolombia/`

El paquete tiene 7 archivos organizados por responsabilidad única (SRP). Cada archivo hace una sola cosa y las dependencias fluyen en una dirección: `config` ← `utils` ← `excel`/`pptx_gen` ← `orquestador`.

### 4.1 `config.py` — Configuración y datos de referencia (197 líneas)

Es el archivo que el usuario puede necesitar editar. No depende de ningún otro módulo.

| Clase / Constante | Responsabilidad |
|---|---|
| `TipoUnidad` | Enum con los tipos válidos: EJE, HUB, OFICOM, OFIREG, TRANSVERSAL, MARCA PAÍS |
| `FamiliaUnidad` | Enum con las 3 familias. Incluye método `desde_tipo()` para mapear tipo → familia |
| `Config` | Dataclass con toda la configuración centralizada (límites, carpetas, contraseñas, nombres de plantillas) |
| `EjeMisional` | Dataclass que define un eje misional (prefijo, nombre, líneas estratégicas, color) |
| `EJES_REFERENCIA` | Diccionario con los 4 ejes: Marca País, VP Turismo, VP Inversión, VP Exportaciones |
| `ORDEN_EJES` | Lista con el orden de presentación: `['MP', 'TUR', 'INV', 'EXP']` |

### 4.2 `utils.py` — Utilidades compartidas (120 líneas)

Funciones helper, estilos Excel y configuración de logging. No depende de ningún otro módulo del paquete.

| Clase / Función | Responsabilidad |
|---|---|
| `log` | Logger centralizado (`ProColombia`) |
| `_ss()` | Safe string: convierte cualquier valor a string limpio |
| `_trunc()` | Trunca texto a un máximo de caracteres |
| `_safe_filename()` | Convierte nombre a formato seguro para archivos |
| `medir_tiempo()` | Crea un medidor de tiempo reutilizable |
| `EstilosExcel` | Fuentes, rellenos, alineaciones y bordes reutilizables para Excel |

### 4.3 `excel.py` — Constructores y lectores de Excel (1,895 líneas)

Contiene las 6 clases de Excel (3 constructores + 3 lectores):

| Clase | Responsabilidad |
|---|---|
| `ConstructorExcelMisional` | Genera la matriz Excel de captura para unidades misionales (8 hojas) |
| `ConstructorExcelTerritorial` | Genera la matriz Excel para unidades territoriales (22 hojas) |
| `ConstructorExcelTransversal` | Genera la matriz Excel para unidades transversales (~27 hojas) |
| `LectorExcel` | Lee un Excel misional diligenciado y retorna un diccionario estructurado |
| `LectorExcelTerritorial` | Lee un Excel territorial (tendencias por eje + contribuciones) |
| `LectorExcelTransversal` | Lee un Excel transversal (líneas propias + contribuciones) |

### 4.4 `pptx_gen.py` — Generadores y plantillas PPTX (929 líneas)

Contiene las 5 clases de PPTX (2 constructores de plantilla + 3 generadores):

| Clase | Responsabilidad |
|---|---|
| `GeneradorPPTXMisional` | Reemplaza marcadores `{{...}}` en la plantilla PPTX misional |
| `GeneradorPPTXTerritorial` | Reemplaza marcadores en la plantilla territorial |
| `GeneradorPPTXTransversal` | Combina reemplazos misional + territorial |
| `ConstructorPlantillaTerritorial` | Construye la plantilla PPTX territorial programáticamente (28 slides) |
| `ConstructorPlantillaTransversal` | Construye la plantilla PPTX transversal (~34 slides) |

### 4.5 `orquestador.py` — Punto de entrada principal (278 líneas)

| Clase | Responsabilidad |
|---|---|
| `OrquestadorUniversal` | Router inteligente: detecta familia, selecciona lector/generador/plantilla, procesa lotes, consolida |

### 4.6 `ejemplos.py` — Datos de ejemplo y ayuda (554 líneas)

| Función | Responsabilidad |
|---|---|
| `datos_ejemplo_turismo()` | Datos de ejemplo de VP Turismo (misional) |
| `datos_ejemplo_hub_norteamerica()` | Datos de ejemplo de Hub Norteamérica (territorial) |
| `datos_ejemplo_gic()` | Datos de ejemplo de la Gerencia de Inteligencia Comercial (transversal) |
| `guia_colab()` | Imprime guía de uso rápida en consola |
| `banner()` | Imprime el banner informativo del sistema |

### 4.7 `__init__.py` — API pública (46 líneas)

Define `__all__` con exactamente 11 símbolos que se exponen al hacer `from procolombia import *`. Esto evita contaminar el namespace del notebook con clases internas como `EstilosExcel`, `LectorExcel`, etc.

---

## 5. Configuración del sistema en Google Colab

### 5.1 Preparación inicial (una sola vez)

**Celda 1 — Montar Google Drive e instalar dependencias:**

```python
from google.colab import drive
drive.mount('/content/drive')
!pip install python-pptx openpyxl -q
```

**Celda 2 — Cargar los módulos:**

```python
import sys

RUTA = "/content/drive/MyDrive/ProColombia/Automatizaciones/VPP"
sys.path.insert(0, RUTA)

from procolombia import *

banner()
```

> **Nota:** Al hacer `from procolombia import *` solo se importan 11 símbolos controlados por `__all__`: `Config`, `OrquestadorUniversal`, `FamiliaUnidad`, `TipoUnidad`, `EJES_REFERENCIA`, `ORDEN_EJES`, las 3 funciones de datos de ejemplo, `guia_colab` y `banner`. No se contamina el namespace con clases internas.

### 5.2 Personalizar la configuración (opcional)

Para modificar los valores por defecto, pase un objeto `Config` personalizado:

```python
cfg = Config(
    max_lineas_estrategicas=4,    # Máximo de líneas por unidad (1-6)
    max_acciones_por_linea=10,    # Máximo de acciones por línea (≥1)
    max_indicadores_por_linea=8,  # Máximo de indicadores por línea
    max_tendencias=5,             # Máximo de tendencias
    max_dofa_por_cuadrante=8,     # Máximo de ítems DOFA por cuadrante
    max_casos_exito=6,            # Máximo de casos de éxito
    max_metas=12,                 # Máximo de metas generales
    max_chars_campo=500,          # Caracteres máximos por celda Excel
    max_chars_slide=450,          # Caracteres máximos por campo en slide
    password='mi_clave_2026',     # Contraseña de protección de hojas
)
orq = OrquestadorUniversal(config=cfg, base_dir=RUTA)
```

Si no necesita personalizar nada, use los valores por defecto:

```python
orq = OrquestadorUniversal(base_dir=RUTA)
```

---

## 6. Operaciones disponibles

### 6.1 Construir plantillas PPTX

Genera las plantillas Territorial y Transversal programáticamente. La plantilla Misional debe subirse manualmente (es el archivo institucional de ProColombia).

```python
orq = OrquestadorUniversal(base_dir=RUTA)
rutas = orq.construir_plantillas()
```

**Resultado:** Crea dos archivos en `04_plantillas/`:

| Archivo | Slides | Contenido |
|---|---|---|
| `Plantilla_Territorial.pptx` | ~28 | Portada, DOFA, Tendencias por eje, 15 hojas de contribución, Metas |
| `Plantilla_Transversal.pptx` | ~34 | Todo lo territorial + slides de líneas estratégicas propias |

> **Cuándo ejecutar:** Solo necesita hacerse una vez, o cuando se quiera regenerar las plantillas (por ejemplo, si cambió la configuración de límites).

### 6.2 Generar Excel vacíos para las áreas

Crea las matrices de captura que las áreas deben diligenciar.

**Individual:**

```python
orq.generar_excel(
    'VP Exportaciones',    # Nombre de la unidad
    'EJE',                 # Tipo de unidad
    trimestre='Q1',        # Trimestre
    anio='2026',           # Año
    num_lineas=3           # Número de líneas estratégicas
)
```

**En lote (múltiples unidades):**

```python
UNIDADES = [
    # (nombre, tipo, num_lineas)  — num_lineas es opcional (default: 5)
    ('VP Exportaciones', 'EJE', 3),
    ('VP Inversión', 'EJE', 3),
    ('VP Turismo', 'EJE', 4),
    ('Marca País', 'MARCA PAÍS', 4),
    ('Hub Norteamérica', 'HUB'),         # Sin num_lineas → usa default
    ('Hub Europa', 'HUB'),
    ('Oficom EE.UU.', 'OFICOM'),
    ('OfiReg Medellín', 'OFIREG'),
    ('Gerencia Inteligencia Comercial', 'TRANSVERSAL', 3),
    ('Gerencia Comunicaciones', 'TRANSVERSAL', 3),
]

for u in UNIDADES:
    nombre, tipo = u[0], u[1]
    nl = u[2] if len(u) > 2 else 5
    orq.generar_excel(nombre, tipo, trimestre='Q1', anio='2026', num_lineas=nl)
```

**Con datos pre-llenados** (para incluir información base de la Hoja de Ruta):

```python
datos_gic = datos_ejemplo_gic()
orq.generar_excel(
    'Gerencia Inteligencia Comercial', 'TRANSVERSAL',
    trimestre='Q1', anio='2026', num_lineas=3,
    datos_base=datos_gic
)
```

**Resultado:** Archivos `.xlsx` en `01_excels_entrada/` con nombre formato: `Q1_2026_EJE_VP_Exportaciones.xlsx`

### 6.3 Procesar los Excel diligenciados y generar presentaciones

Una vez las áreas devuelvan los Excel diligenciados, colóquelos en `01_excels_entrada/` y ejecute:

```python
%%time
import gc
if 'orq' in globals():
    del orq; gc.collect()

orq = OrquestadorUniversal(base_dir=RUTA)
resultados = orq.procesar_lote()
```

**Lo que hace internamente para cada archivo:**

1. Detecta la familia examinando las hojas del Excel.
2. Selecciona el lector correcto (`LectorExcel`, `LectorExcelTerritorial` o `LectorExcelTransversal`).
3. Extrae todos los datos a un diccionario estructurado.
4. Busca la plantilla PPTX correspondiente en `04_plantillas/`.
5. Reemplaza todos los marcadores `{{...}}` con los datos extraídos.
6. Elimina los slides de líneas estratégicas que no tengan contenido diligenciado.
7. Limpia marcadores residuales (quedan como texto vacío).
8. Guarda la presentación en `02_pptx_salida/`.

**Resultado:** Archivos `.pptx` en `02_pptx_salida/` con nombre formato: `Q1_2026_EJE_VP_Exportaciones_Seguimiento.pptx`

**Reporte de salida:**

```
══════════════════════════════════════════════════════════════════════
  PROCESAMIENTO UNIVERSAL — 7 archivos
══════════════════════════════════════════════════════════════════════
  🔍 Q1_2026_EJE_VP_Exportaciones.xlsx → familia detectada: MISIONAL
  ✅ VP Exportaciones                (MISIONAL    ) → Q1_2026_EJE_VP_Exportaciones_Seguimiento.pptx
  🔍 Q1_2026_HUB_Hub_Norteamérica.xlsx → familia detectada: TERRITORIAL
  ✅ Hub Norteamérica                (TERRITORIAL ) → Q1_2026_HUB_Hub_Norteamérica_Seguimiento.pptx
  ...

  ✅ 7 exitosos | ❌ 0 errores
```

### 6.4 Consolidar información en Excel maestro

Genera un Excel analítico con toda la información de todas las unidades:

```python
ruta_consolidado = orq.consolidar()
```

**Resultado:** Archivo en `03_consolidado/` con nombre `Consolidado_Universal_20260328_1430.xlsx` (con timestamp).

**Hojas del consolidado:**

| Hoja | Contenido |
|---|---|
| `RESUMEN` | Una fila por unidad con métricas: ítems DOFA, líneas propias, acciones, contribuciones, metas |
| `DOFA` | Todos los ítems DOFA de todas las unidades (cuadrante, base, estado, actualización) |
| `ACCIONES_PROPIAS` | Acciones de líneas estratégicas propias (misionales y transversales) |
| `CONTRIBUCIONES` | Acciones de contribución a ejes misionales (territoriales y transversales) |

---

## 7. Estructura de los Excel generados

### 7.1 Excel Misional (8 hojas)

| Hoja | Propósito | Color pestaña |
|---|---|---|
| `INSTRUCCIONES` | Guía de uso para el usuario | Azul oscuro |
| `CONFIGURACIÓN` | Datos del área: trimestre, año, tipo, nombre, fechas | Azul medio |
| `DOFA` | Seguimiento de Debilidades, Oportunidades, Fortalezas, Amenazas | Rojo |
| `TENDENCIAS` | Tendencias del sector con estado y actualización | Púrpura |
| `LÍNEA ESTRATÉGICA 1..N` | Acciones, actividades, avances e indicadores por línea | Colores variados |
| `CASOS DE ÉXITO` | Casos destacados del trimestre | Verde |
| `METAS GENERALES` | Indicadores, metas y avances acumulados | Azul |
| `MONITOREO` | Completitud automática con fórmulas (no editable) | Gris |

### 7.2 Excel Territorial (22 hojas)

| Hoja | Propósito |
|---|---|
| `INSTRUCCIONES` | Guía de uso |
| `CONFIGURACIÓN` | Datos del hub/oficina |
| `TENDENCIAS POR EJE` | Tendencias, foco y aporte por cada eje misional (TUR, INV, EXP) |
| `DOFA` | Seguimiento DOFA |
| `MP_LE1` a `MP_LE4` | Contribuciones a Marca País (4 líneas) |
| `TUR_LE1` a `TUR_LE4` | Contribuciones a VP Turismo (4 líneas) |
| `INV_LE1` a `INV_LE3` | Contribuciones a VP Inversión (3 líneas) |
| `EXP_LE1` a `EXP_LE4` | Contribuciones a VP Exportaciones (4 líneas) |
| `METAS GENERALES` | Indicadores y avances |
| `PRESUPUESTO` | Ejecución presupuestal |
| `MONITOREO` | Completitud automática |

### 7.3 Excel Transversal (~27 hojas)

Combina la estructura misional y territorial:

- Hojas de líneas estratégicas propias (`LÍNEA ESTRATÉGICA 1..N`) como las misionales.
- Hojas de contribuciones a ejes (`MP_LE1`, `TUR_LE1`, etc.) como las territoriales.
- DOFA, Metas y Monitoreo.
- No incluye TENDENCIAS ni CASOS DE ÉXITO.

---

## 8. Sistema de marcadores PPTX

Las plantillas usan marcadores con la sintaxis `{{NOMBRE}}` que el sistema reemplaza automáticamente con los datos del Excel.

### 8.1 Marcadores generales

| Marcador | Valor |
|---|---|
| `{{TRIMESTRE}}` | Q1, Q2, Q3 o Q4 |
| `{{UNIDAD}}` | Nombre de la unidad |
| `{{AÑO}}` | Año del seguimiento |
| `{{TIPO}}` | Tipo de unidad |

### 8.2 Marcadores DOFA

| Marcador | Ejemplo |
|---|---|
| `{{DEB_BASE_1}}` a `{{DEB_BASE_10}}` | Debilidades (actualizadas si aplica) |
| `{{OPO_BASE_1}}` a `{{OPO_BASE_10}}` | Oportunidades |
| `{{FOR_BASE_1}}` a `{{FOR_BASE_10}}` | Fortalezas |
| `{{AME_BASE_1}}` a `{{AME_BASE_10}}` | Amenazas |

**Lógica de reemplazo DOFA:** Si el estado es "Se elimina", el marcador queda vacío. Si es "Se actualiza" y tiene texto de actualización, se usa el texto nuevo. En cualquier otro caso, se usa el texto base original.

### 8.3 Marcadores de Tendencias (misional)

| Marcador | Valor |
|---|---|
| `{{TEND_1}}` a `{{TEND_7}}` | Texto de la tendencia (actualizado si aplica) |

### 8.4 Marcadores de Líneas Estratégicas (misional + transversal)

| Marcador | Valor |
|---|---|
| `{{LE1_NOMBRE}}` | Nombre de la línea estratégica 1 |
| `{{LE1_ACC_1}}` | Acción 1 de la línea 1 |
| `{{LE1_ACT_1}}` | Actividad 1 de la línea 1 |
| `{{LE1_AVA_1}}` | Avance de la acción 1 de la línea 1 |
| `{{LE1_IND_1}}` | Indicador 1 de la línea 1 |
| `{{LE1_META_1}}` | Meta del indicador 1 |
| `{{LE1_RES_1}}` | Resultado/avance del indicador 1 |

Los índices van de `LE1` a `LE5` y las acciones de `_ACC_1` a `_ACC_12`.

### 8.5 Marcadores de Casos de Éxito y Metas

| Marcador | Valor |
|---|---|
| `{{CASO_1_TIT}}` | Título del caso de éxito 1 |
| `{{CASO_1_DESC}}` | Descripción del caso 1 |
| `{{META_1_IND}}` | Indicador de meta 1 |
| `{{META_1_META}}` | Valor meta 1 |
| `{{META_1_AVA}}` | Avance de meta 1 |

### 8.6 Marcadores territoriales (contribuciones)

| Marcador | Valor |
|---|---|
| `{{MP_LE1_ACC_1}}` | Acción 1, Línea 1 de Marca País |
| `{{TUR_LE2_AVA_3}}` | Avance de acción 3, Línea 2 de Turismo |
| `{{INV_LE1_IND_1}}` | Indicador 1, Línea 1 de Inversión |

### 8.7 Eliminación inteligente de slides

Los slides que contienen marcadores de líneas estratégicas sin contenido se eliminan automáticamente. Por ejemplo, si una unidad solo tiene 3 líneas estratégicas diligenciadas, los slides de LE4 y LE5 se eliminan de la presentación final. Los marcadores residuales que no fueron reemplazados se limpian a texto vacío.

---

## 9. Datos de ejemplo incluidos

El sistema incluye funciones con datos reales extraídos de presentaciones existentes, útiles para pruebas y demostraciones:

| Función | Familia | Descripción |
|---|---|---|
| `datos_ejemplo_turismo()` | MISIONAL | VP Turismo con DOFA, tendencias, líneas y metas completas |
| `datos_ejemplo_hub_norteamerica()` | TERRITORIAL | Hub Norteamérica con contribuciones a los 4 ejes |
| `datos_ejemplo_gic()` | TRANSVERSAL | Gerencia de Inteligencia Comercial con líneas propias + contribuciones |

**Uso:**

```python
# Generar Excel pre-llenado con datos de ejemplo
datos = datos_ejemplo_turismo()
orq.generar_excel('VP Turismo', 'EJE', trimestre='Q1', anio='2026',
                  num_lineas=4, datos_base=datos)
```

---

## 10. Lo que el sistema puede hacer

- Generar Excel de captura para cualquier combinación de familias en una sola corrida.
- Procesar lotes mixtos de archivos (misionales + territoriales + transversales simultáneamente).
- Detectar la familia automáticamente sin intervención del usuario.
- Pre-llenar Excel con datos base de la Hoja de Ruta anterior.
- Generar presentaciones PPTX con marcadores reemplazados y slides dinámicos.
- Eliminar slides de líneas estratégicas vacías automáticamente.
- Limpiar cualquier marcador residual `{{...}}` que no haya sido reemplazado.
- Consolidar información de todas las unidades en un Excel analítico.
- Proteger hojas del Excel con contraseña (configurable).
- Aplicar formato condicional (celdas pendientes se marcan en rojo).
- Validar datos con listas desplegables (estados, trimestres, tipos).
- Manejar errores individualmente (si un archivo falla, los demás se procesan normalmente).

---

## 11. Lo que el sistema NO puede hacer

- **No edita presentaciones existentes:** Siempre genera una nueva desde la plantilla. Si necesita modificar una PPTX ya generada, debe hacerlo manualmente.
- **No sube ni baja archivos de Drive automáticamente:** El usuario debe colocar los Excel en `01_excels_entrada/` y recoger las PPTX de `02_pptx_salida/` manualmente.
- **No envía correos ni notificaciones:** El flujo de distribución a las áreas es manual.
- **No valida la calidad del contenido:** Solo verifica que las hojas y celdas existan y tengan datos. No evalúa si el texto es coherente o completo.
- **No fusiona presentaciones:** Genera un PPTX por unidad. Si necesita una presentación consolidada, debe unirlas manualmente o con otra herramienta.
- **No soporta formatos que no sean `.xlsx`:** Archivos `.xls` (formato legacy), `.csv` o Google Sheets no son compatibles.
- **No procesa imágenes ni gráficos desde el Excel:** Solo extrae texto. Si el Excel contiene gráficos o imágenes, estos se ignoran.
- **No hace rollback automático:** Si un archivo sale mal, debe corregir el Excel y volver a ejecutar `procesar_lote()`.
- **No soporta más de 6 líneas estratégicas por unidad** (límite configurable en `Config`, máximo 6).
- **No genera la plantilla Misional:** Esta plantilla es el archivo institucional de ProColombia y debe subirse manualmente a `04_plantillas/`. Solo las plantillas Territorial y Transversal se generan programáticamente.

---

## 12. Errores comunes y solución

### Error: `Worksheet named 'TENDENCIAS' not found`

**Causa:** Se está usando el código antiguo (archivos sueltos `procolombia_planeacion.py` con `Orquestador()`) en lugar del paquete refactorizado.

**Solución:** Migre al paquete `procolombia/` y use:
```python
from procolombia import OrquestadorUniversal
orq = OrquestadorUniversal(base_dir=RUTA)
resultados = orq.procesar_lote()
```

### Error: `Plantilla XXXX no encontrada`

**Causa:** La plantilla PPTX no existe en `04_plantillas/`.

**Solución:**
- Para la Misional: suba `Plantilla_Misional.pptx` manualmente.
- Para Territorial/Transversal: ejecute `orq.construir_plantillas()`.

### Error: `Tipo de unidad 'XXX' no reconocido`

**Causa:** La hoja CONFIGURACIÓN del Excel tiene un tipo de unidad que no está en el mapeo del sistema.

**Solución:** Asegúrese de que el campo "Tipo de unidad" sea uno de: `EJE`, `MARCA PAÍS`, `HUB`, `OFICOM`, `OFIREG`, `FIDIREP` o `TRANSVERSAL`.

### Error: `No hay archivos .xlsx en 01_excels_entrada/`

**Causa:** La carpeta de entrada está vacía o la ruta base es incorrecta.

**Solución:** Verifique que los Excel estén en la carpeta correcta y que `base_dir` apunte al directorio raíz donde están las 4 subcarpetas.

### Los marcadores `{{...}}` aparecen en el PPTX sin reemplazar

**Causa posible 1:** El marcador en la plantilla tiene un nombre que no coincide con el que genera el sistema (por ejemplo, un espacio extra o un carácter diferente).

**Causa posible 2:** Los runs de PowerPoint dividen el marcador en fragmentos. El sistema concatena los runs de cada párrafo para reconstruir el marcador completo, pero si hay formatos mixtos dentro del marcador, puede fallar.

**Solución:** Abra la plantilla, seleccione todo el texto del marcador, y aplique un formato uniforme (misma fuente, tamaño y color en todo el `{{NOMBRE}}`).

### La presentación sale con slides vacíos

**Causa:** Las líneas estratégicas no tenían datos (nombre, acciones o indicadores), pero la plantilla no tenía marcadores estándar `LE1_`, `LE2_`, etc. que permitan al sistema identificar qué slides eliminar.

**Solución:** Verifique que la plantilla use los marcadores con el formato exacto `{{LEn_...}}` donde `n` es el número de línea.

---

## 13. Referencia rápida de celdas para Colab

### Celda de preparación (ejecutar una vez por sesión)

```python
# ── Montar Drive ─────────────────────────────────────────────
from google.colab import drive
drive.mount('/content/drive')
!pip install python-pptx openpyxl -q

# ── Cargar módulos ───────────────────────────────────────────
import sys
RUTA = "/content/drive/MyDrive/ProColombia/Automatizaciones/VPP"
sys.path.insert(0, RUTA)

from procolombia import *
banner()
```

### Celda para construir plantillas

```python
orq = OrquestadorUniversal(base_dir=RUTA)
orq.construir_plantillas()
```

### Celda para generar Excel vacíos

```python
UNIDADES = [
    ('VP Exportaciones', 'EJE', 3),
    ('VP Inversión', 'EJE', 3),
    ('VP Turismo', 'EJE', 4),
    ('Marca País', 'MARCA PAÍS', 4),
    ('Hub Norteamérica', 'HUB'),
    ('Oficom EE.UU.', 'OFICOM'),
    ('Gerencia Inteligencia Comercial', 'TRANSVERSAL', 3),
]
for u in UNIDADES:
    nombre, tipo = u[0], u[1]
    nl = u[2] if len(u) > 2 else 5
    orq.generar_excel(nombre, tipo, trimestre='Q1', anio='2026', num_lineas=nl)
```

### Celda para procesar lote

```python
%%time
import gc
if 'orq' in globals():
    del orq; gc.collect()

orq = OrquestadorUniversal(base_dir=RUTA)
resultados = orq.procesar_lote()
```

### Celda para consolidar

```python
ruta_consolidado = orq.consolidar()
```

### Celda para ver la guía rápida en consola

```python
guia_colab()
```

---

## 14. Parámetros configurables de `Config`

| Parámetro | Default | Descripción |
|---|---|---|
| `max_lineas_estrategicas` | 5 | Máximo de líneas por unidad (1–6) |
| `max_acciones_por_linea` | 12 | Máximo de acciones por línea estratégica |
| `max_indicadores_por_linea` | 10 | Máximo de indicadores por línea |
| `max_tendencias` | 7 | Máximo de tendencias (misional) |
| `max_dofa_por_cuadrante` | 10 | Máximo de ítems DOFA por cuadrante |
| `max_casos_exito` | 10 | Máximo de casos de éxito (misional) |
| `max_metas` | 15 | Máximo de metas generales |
| `max_chars_campo` | 500 | Caracteres máximos por celda Excel |
| `max_chars_slide` | 450 | Caracteres máximos por campo en slide |
| `password` | `planeacion2026` | Contraseña de protección de hojas |
| `dir_entrada` | `01_excels_entrada` | Carpeta de Excel de entrada |
| `dir_salida` | `02_pptx_salida` | Carpeta de PPTX de salida |
| `dir_consolidado` | `03_consolidado` | Carpeta del consolidado |
| `dir_plantillas` | `04_plantillas` | Carpeta de plantillas |
| `plantilla_misional` | `Plantilla_Misional.pptx` | Nombre del archivo de plantilla misional |
| `plantilla_territorial` | `Plantilla_Territorial.pptx` | Nombre del archivo de plantilla territorial |
| `plantilla_transversal` | `Plantilla_Transversal.pptx` | Nombre del archivo de plantilla transversal |

---

## 15. Flujo completo paso a paso

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
│     (proceso manual fuera del sistema)                          │
├─────────────────────────────────────────────────────────────────┤
│  4. RECOPILAR                                                   │
│     Recibir los Excel diligenciados de las áreas                │
│     Colocarlos en 01_excels_entrada/                            │
│     (pueden convivir con los vacíos: se sobrescriben)           │
├─────────────────────────────────────────────────────────────────┤
│  5. PROCESAR                                                    │
│     Ejecutar orq.procesar_lote()                                │
│     Las presentaciones se generan en 02_pptx_salida/            │
├─────────────────────────────────────────────────────────────────┤
│  6. CONSOLIDAR (opcional)                                       │
│     Ejecutar orq.consolidar()                                   │
│     El Excel maestro se genera en 03_consolidado/               │
├─────────────────────────────────────────────────────────────────┤
│  7. ENTREGAR                                                    │
│     Descargar PPTX de 02_pptx_salida/                           │
│     Distribuir a las áreas y a la dirección                     │
│     (proceso manual fuera del sistema)                          │
└─────────────────────────────────────────────────────────────────┘
```

---

## 16. Ejes misionales de referencia

Estas son las líneas estratégicas predefinidas de cada eje misional. Las unidades territoriales y transversales reportan sus contribuciones a estas líneas:

### Marca País (MP) — 4 líneas

1. Fomentar la imagen positiva del país a nivel internacional.
2. Trabajo en conjunto con aliados del sector público y privado.
3. Comercializar la Marca País mediante participación en actividades comerciales.
4. Apoyar áreas transversales y oficinas en solicitudes institucionales.

### VP Turismo (TUR) — 4 líneas

1. Liderar el dinamismo en la conectividad aérea, marítima y transfronteriza.
2. Desarrollar campañas y acciones segmentadas (B2B / B2C).
3. Promover a Colombia como destino de turismo de reuniones de alto impacto.
4. Fomentar la promoción a través de las seis regiones turísticas.

### VP Inversión (INV) — 3 líneas

1. Apoyar a inversionistas instalados en proyectos de reinversión.
2. Promover la atracción de IED de empresas nuevas.
3. Promover la atracción de IED a las diferentes regiones del país.

### VP Exportaciones (EXP) — 4 líneas

1. Promover la canasta exportable No Minero Energética en mercados internacionales.
2. Apoyar la diversificación de las exportaciones desde la demanda y la oferta.
3. Adecuar la oferta exportable mediante el cierre de brechas.
4. Capacitar la oferta exportable para la generación de cultura exportadora.

---

## 17. Código de colores en los Excel

| Color | Significado | Acción del usuario |
|---|---|---|
| Azul claro | Campo editable | Diligenciar |
| Amarillo claro | Información base de la Hoja de Ruta | No editar (solo referencia) |
| Gris | Campo calculado o bloqueado | No editar |
| Rojo claro | Campo pendiente (formato condicional) | Requiere atención |
| Blanco | Campo vacío editable | Diligenciar si aplica |

---

## 18. Dependencias del sistema

| Paquete | Versión mínima | Uso |
|---|---|---|
| `pandas` | ≥ 1.5 | Lectura de Excel y consolidación |
| `openpyxl` | ≥ 3.0 | Creación y lectura de archivos .xlsx |
| `python-pptx` | ≥ 0.6 | Creación y manipulación de archivos .pptx |
| Python | ≥ 3.9 | Tipado moderno, f-strings, pathlib |

Instalación en Colab:

```python
!pip install python-pptx openpyxl -q
```

> Pandas y openpyxl ya vienen preinstalados en Google Colab. Solo es necesario instalar `python-pptx`.
