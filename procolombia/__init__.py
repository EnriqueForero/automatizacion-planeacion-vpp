# -*- coding: utf-8 -*-
"""
Automatización — Seguimiento Planeación Estratégica ProColombia v5.1

Uso desde Google Colab:

    import sys
    sys.path.insert(0, "/content/drive/MyDrive/.../VPP")
    from procolombia import *

    orq = OrquestadorUniversal(base_dir=RUTA)
    resultados = orq.procesar_lote()
"""

# --- API pública (lo que el usuario necesita) ---
from .config import Config, FamiliaUnidad, TipoUnidad, EJES_REFERENCIA, ORDEN_EJES
from .orquestador import OrquestadorUniversal
from .ejemplos import (
    datos_ejemplo_turismo,
    datos_ejemplo_hub_norteamerica,
    datos_ejemplo_gic,
    guia_colab,
    banner,
)

# --- Control explícito de 'from procolombia import *' ---
__all__ = [
    # Configuración (lo que el usuario toca)
    'Config',
    'FamiliaUnidad',
    'TipoUnidad',
    'EJES_REFERENCIA',
    'ORDEN_EJES',

    # Punto de entrada principal
    'OrquestadorUniversal',

    # Datos de ejemplo y ayuda
    'datos_ejemplo_turismo',
    'datos_ejemplo_hub_norteamerica',
    'datos_ejemplo_gic',
    'guia_colab',
    'banner',
]

__version__ = "0.1.0"
