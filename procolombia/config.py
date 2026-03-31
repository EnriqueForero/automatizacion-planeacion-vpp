# -*- coding: utf-8 -*-
"""
Configuración centralizada, enums y datos de referencia.

Este módulo NO depende de ningún otro módulo del paquete.
Todo lo que aquí se define es "lo que el usuario puede tocar"
o "constantes de referencia compartidas por todos los módulos".
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path
from typing import Dict, List, Optional


# ═══════════════════════════════════════════════════════════════════════
# ENUMS
# ═══════════════════════════════════════════════════════════════════════

class TipoUnidad(str, Enum):
    """Tipos de unidad reconocidos por el sistema."""
    EJE = "EJE"
    TRANSVERSAL = "TRANSVERSAL"
    HUB = "HUB"
    OFICOM = "OFICOM"
    OFIREG = "OFIREG"
    MARCA_PAIS = "MARCA PAÍS"

    @classmethod
    def valores_validos(cls) -> List[str]:
        return [e.value for e in cls]


class FamiliaUnidad(str, Enum):
    """Familia de procesamiento (determina plantilla y estructura)."""
    MISIONAL = "MISIONAL"
    TERRITORIAL = "TERRITORIAL"
    TRANSVERSAL = "TRANSVERSAL"

    @classmethod
    def desde_tipo(cls, tipo: str) -> 'FamiliaUnidad':
        """Determina la familia a partir del tipo de unidad."""
        tipo_upper = tipo.upper().strip()
        mapa = {
            "EJE": cls.MISIONAL,
            "MARCA PAÍS": cls.MISIONAL,
            "MARCA PAIS": cls.MISIONAL,
            "VICEPRESIDENCIA": cls.MISIONAL,
            "TRANSVERSAL": cls.TRANSVERSAL,
            "HUB": cls.TERRITORIAL,
            "OFICOM": cls.TERRITORIAL,
            "OFIREG": cls.TERRITORIAL,
            "FIDIREP": cls.TERRITORIAL,
        }
        familia = mapa.get(tipo_upper)
        if familia is None:
            raise ValueError(
                f"Tipo de unidad '{tipo}' no reconocido. "
                f"Válidos: {list(mapa.keys())}"
            )
        return familia


# ═══════════════════════════════════════════════════════════════════════
# CONFIGURACIÓN CENTRALIZADA
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class Config:
    """
    Configuración centralizada del sistema. Sin números mágicos.

    Para personalizar, pase un Config modificado al OrquestadorUniversal:

        cfg = Config(max_lineas_estrategicas=4, password='mi_clave')
        orq = OrquestadorUniversal(config=cfg)
    """

    # --- Límites estructurales ---
    max_lineas_estrategicas: int = 5
    max_acciones_por_linea: int = 12
    max_indicadores_por_linea: int = 10
    max_tendencias: int = 7
    max_dofa_por_cuadrante: int = 10
    max_casos_exito: int = 10
    max_metas: int = 15

    # --- Límites de texto ---
    max_chars_campo: int = 500
    max_chars_slide: int = 450

    # --- Protección ---
    password: str = 'planeacion2026'

    # --- Carpetas ---
    dir_entrada: str = '01_excels_entrada'
    dir_salida: str = '02_pptx_salida'
    dir_consolidado: str = '03_consolidado'
    dir_plantillas: str = '04_plantillas'

    # --- Nombres de plantillas por familia ---
    plantilla_misional: str = 'Plantilla_Misional.pptx'
    plantilla_territorial: str = 'Plantilla_Territorial.pptx'
    plantilla_transversal: str = 'Plantilla_Transversal.pptx'

    def __post_init__(self):
        if self.max_acciones_por_linea < 1:
            raise ValueError("max_acciones_por_linea debe ser >= 1")
        if self.max_lineas_estrategicas < 1 or self.max_lineas_estrategicas > 6:
            raise ValueError("max_lineas_estrategicas debe estar entre 1 y 6")

    def crear_carpetas(self, base_dir: str = '.') -> None:
        """Crea la estructura de carpetas si no existe."""
        for d in [self.dir_entrada, self.dir_salida,
                  self.dir_consolidado, self.dir_plantillas]:
            Path(base_dir, d).mkdir(parents=True, exist_ok=True)

    def ruta_plantilla(self, familia: FamiliaUnidad,
                       base_dir: str = '.') -> Optional[Path]:
        """Retorna la ruta de la plantilla según la familia."""
        nombres = {
            FamiliaUnidad.MISIONAL: self.plantilla_misional,
            FamiliaUnidad.TERRITORIAL: self.plantilla_territorial,
            FamiliaUnidad.TRANSVERSAL: self.plantilla_transversal,
        }
        ruta = Path(base_dir, self.dir_plantillas, nombres[familia])
        return ruta if ruta.exists() else None


# ═══════════════════════════════════════════════════════════════════════
# EJES MISIONALES DE REFERENCIA
# ═══════════════════════════════════════════════════════════════════════

@dataclass
class EjeMisional:
    """Definición de un eje misional con sus líneas estratégicas."""
    prefijo: str          # MP, TUR, INV, EXP
    nombre: str           # Nombre completo
    nombre_corto: str     # Para encabezados
    lineas: List[str]     # Nombres de las líneas
    color: str = '1B3A5C' # Color de pestaña


EJES_REFERENCIA: Dict[str, EjeMisional] = {
    'MP': EjeMisional(
        prefijo='MP',
        nombre='MARCA PAÍS',
        nombre_corto='Marca País',
        color='E67E22',
        lineas=[
            'Fomentar la imagen positiva del país: Desarrollar acciones estratégicas para fomentar el sentido de pertenencia y dar a conocer la marca país a nivel internacional.',
            'Trabajo en conjunto con aliados: Sumar esfuerzos con empresas del sector público y privado para desarrollar proyectos que den visibilidad a la Marca País.',
            'Comercializar la Marca País: Participación en actividades comerciales para generar ingresos a través de productos que representen la Marca.',
            'Apoyar áreas transversales y oficinas: Apoyar en solicitudes institucionales de oficinas comerciales, regionales y Presidencia.',
        ],
    ),
    'TUR': EjeMisional(
        prefijo='TUR',
        nombre='VP TURISMO',
        nombre_corto='VP Turismo',
        color='2980B9',
        lineas=[
            'Liderar el dinamismo en la conectividad aérea, marítima y transfronteriza.',
            'Desarrollar campañas y acciones segmentadas (B2B – Público profesional / B2C – Público final).',
            'Promover a Colombia como destino de turismo de reuniones de alto impacto.',
            'Fomentar la promoción a través de las seis regiones turísticas.',
        ],
    ),
    'INV': EjeMisional(
        prefijo='INV',
        nombre='VP INVERSIÓN',
        nombre_corto='VP Inversión',
        color='27AE60',
        lineas=[
            'Apoyar a los inversionistas instalados en Colombia para que puedan desarrollar proyectos de reinversión.',
            'Promover la atracción de IED de empresas nuevas a Colombia, que quieran usar el país como plataforma exportadora.',
            'Promover la atracción de IED a las diferentes regiones del país a través de propuestas de valor diferenciales.',
        ],
    ),
    'EXP': EjeMisional(
        prefijo='EXP',
        nombre='VP EXPORTACIONES',
        nombre_corto='VP Exportaciones',
        color='C0392B',
        lineas=[
            'Promover la canasta exportable de bienes y servicios No Minero Energéticos en los mercados internacionales.',
            'Apoyar la diversificación de las exportaciones No Minero Energéticas desde la demanda y la oferta.',
            'Adecuar la oferta exportable mediante el cierre de brechas.',
            'Capacitar la oferta exportable para la generación de cultura exportadora.',
        ],
    ),
}

# Orden de presentación de ejes en hojas de contribución
ORDEN_EJES: List[str] = ['MP', 'TUR', 'INV', 'EXP']
