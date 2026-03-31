# -*- coding: utf-8 -*-
"""
Lectores de Excel para las 3 familias.

Leen los Excel diligenciados por las áreas y retornan
diccionarios estructurados listos para generar PPTX.

Clases:
    LectorExcel              — Lee archivos MISIONAL
    LectorExcelTerritorial   — Lee archivos TERRITORIAL (Hub, Oficom, OfiReg)
    LectorExcelTransversal   — Lee archivos TRANSVERSAL (Gerencias)

Notas de rendimiento:
    - Se usa itertuples() en vez de iterrows() en todas las lecturas.
      itertuples() retorna namedtuples (10x más rápido que iterrows
      que retorna Series).
    - Los DataFrames de lectura son pequeños (<100 filas), pero se
      sigue la regla de oro de evitar iterrows sobre DataFrames.
"""

from __future__ import annotations

import os
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import openpyxl

from .config import (
    Config, FamiliaUnidad,
    EJES_REFERENCIA, ORDEN_EJES,
)
from .utils import _ss, log


# ═══════════════════════════════════════════════════════════════════════
# HELPERS INTERNOS
# ═══════════════════════════════════════════════════════════════════════

def _safe_col(row: tuple, idx: int, default: str = '') -> str:
    """Extrae columna de un tuple de forma segura.

    Args:
        row: Fila como tuple (de itertuples o values).
        idx: Índice de la columna.
        default: Valor si el índice no existe.

    Returns:
        String limpio del valor, o default si fuera de rango.
    """
    if idx < len(row):
        return _ss(row[idx], default)
    return default


def _es_numero(valor: str) -> bool:
    """Verifica si un string representa un número (ID de fila).

    Args:
        valor: Texto a evaluar.

    Returns:
        True si el valor es parseable como int.
    """
    try:
        int(float(valor))
        return True
    except (ValueError, TypeError):
        return False


# ═══════════════════════════════════════════════════════════════════════
# LECTOR EXCEL — MISIONAL
# ═══════════════════════════════════════════════════════════════════════

class LectorExcel:
    """
    Lee un Excel misional diligenciado y retorna un dict estructurado.

    Tolerante a variantes de formato (v1 y v2 de CONFIGURACIÓN).
    Hojas que lee: CONFIGURACIÓN, DOFA, TENDENCIAS, LÍNEA ESTRATÉGICA 1..N,
                   CASOS DE ÉXITO, METAS GENERALES.

    Returns (de .leer()):
        dict con keys: config, familia, dofa, tendencias, lineas,
                       casos_exito, metas
    """

    def __init__(self, config: Config) -> None:
        self.cfg = config

    def leer(self, ruta: str) -> Dict[str, Any]:
        """Lee un Excel MISIONAL y retorna dict con todos los datos.

        Para archivos territoriales/transversales, usar el lector
        correspondiente (LectorExcelTerritorial / LectorExcelTransversal)
        o bien OrquestadorUniversal que selecciona automáticamente.

        Args:
            ruta: Ruta al archivo .xlsx.

        Returns:
            Dict con keys: config, familia, dofa, tendencias, lineas,
                           casos_exito, metas.

        Raises:
            ValueError: Si faltan hojas obligatorias (CONFIGURACIÓN, DOFA).
        """
        nombre_archivo = os.path.basename(ruta)
        log.info(f"📖 Leyendo: {nombre_archivo}")

        # Pre-flight: verificar hojas disponibles
        wb = openpyxl.load_workbook(ruta, read_only=True)
        hojas = set(wb.sheetnames)
        wb.close()

        hojas_misional = {'CONFIGURACIÓN', 'DOFA'}
        faltantes = hojas_misional - hojas
        if faltantes:
            raise ValueError(
                f"Hojas obligatorias faltantes en {nombre_archivo}: "
                f"{faltantes}. ¿El archivo es válido?"
            )

        # Advertir si parece territorial/transversal
        if 'MP_LE1' in hojas and 'TENDENCIAS' not in hojas:
            log.warning(
                f"  ⚠️ {nombre_archivo} parece TERRITORIAL/TRANSVERSAL "
                f"(tiene MP_LE1 pero no TENDENCIAS). "
                f"Considere usar OrquestadorUniversal."
            )

        config = self._leer_configuracion(ruta)
        dofa = self._leer_dofa(ruta)
        tendencias = self._leer_tendencias(ruta)
        lineas = self._leer_lineas(ruta)
        casos = self._leer_casos(ruta)
        metas = self._leer_metas(ruta)

        # Validar tipo de unidad
        tipo = config.get('Tipo de unidad', '')
        try:
            familia = FamiliaUnidad.desde_tipo(tipo)
        except ValueError:
            log.warning(f"⚠️  Tipo '{tipo}' no estándar. Se procesará como EJE.")
            familia = FamiliaUnidad.MISIONAL

        data: Dict[str, Any] = {
            'config': config,
            'familia': familia.value,
            'dofa': dofa,
            'tendencias': tendencias,
            'lineas': lineas,
            'casos_exito': casos,
            'metas': metas,
        }

        n_dofa = sum(len(v) for v in dofa.values())
        log.info(f"  → {config.get('Nombre de la unidad', '?')} | "
                 f"{n_dofa} DOFA | {len(lineas)} líneas | "
                 f"{len(casos)} casos | {len(metas)} metas")
        return data

    def _leer_configuracion(self, ruta: str) -> Dict[str, str]:
        """Lee hoja CONFIGURACIÓN y retorna dict clave→valor.

        Tolerante a variantes: salta filas de título/encabezado.
        Agrega '_archivo' con el nombre del archivo fuente.
        """
        df = pd.read_excel(ruta, sheet_name='CONFIGURACIÓN', header=None)
        config: Dict[str, str] = {}
        for row in df.itertuples(index=False):
            k = _ss(row[0])
            v = _safe_col(row, 1)
            if k and k not in ('CONFIGURACIÓN GENERAL', 'Parámetro'):
                config[k] = v
        config['_archivo'] = os.path.basename(ruta)
        return config

    def _leer_dofa(self, ruta: str) -> Dict[str, List[Dict[str, str]]]:
        """Lee hoja DOFA con sus 4 cuadrantes.

        Parseo secuencial: detecta encabezados de cuadrante
        (DEBILIDADES, OPORTUNIDADES, FORTALEZAS, AMENAZAS)
        y acumula los ítems debajo de cada uno.

        Returns:
            Dict con 4 keys (cuadrantes), cada una con lista de
            dicts {base, estado, actualizacion}.
        """
        df = pd.read_excel(ruta, sheet_name='DOFA', header=None, skiprows=3)
        dofa: Dict[str, List[Dict[str, str]]] = {
            'DEBILIDADES': [], 'OPORTUNIDADES': [],
            'FORTALEZAS': [], 'AMENAZAS': [],
        }
        cuad: Optional[str] = None
        for row in df.itertuples(index=False):
            va = _ss(row[0])
            if va in dofa:
                cuad = va
                continue
            if cuad is None:
                continue
            b = _safe_col(row, 1)
            e = _safe_col(row, 2)
            a = _safe_col(row, 3)
            if b or e or a:
                dofa[cuad].append({
                    'base': b, 'estado': e, 'actualizacion': a,
                })
        return dofa

    def _leer_tendencias(self, ruta: str) -> List[Dict[str, str]]:
        """Lee hoja TENDENCIAS. Retorna lista vacía si no existe.

        Returns:
            Lista de dicts {base, estado, actualizacion}.
        """
        try:
            df = pd.read_excel(ruta, sheet_name='TENDENCIAS',
                              header=None, skiprows=2)
        except (ValueError, KeyError):
            log.warning(f"  ⚠️ Hoja 'TENDENCIAS' no encontrada en "
                        f"{os.path.basename(ruta)} — "
                        f"¿es archivo territorial/transversal?")
            return []
        tendencias: List[Dict[str, str]] = []
        for row in df.itertuples(index=False):
            b = _safe_col(row, 1)
            e = _safe_col(row, 2)
            a = _safe_col(row, 3)
            if b or a:
                tendencias.append({
                    'base': b, 'estado': e, 'actualizacion': a,
                })
        return tendencias

    def _leer_lineas(self, ruta: str) -> List[Dict[str, Any]]:
        """Lee hojas LÍNEA ESTRATÉGICA 1..N.

        Parseo secuencial con estado: detecta secciones ACCIONES
        e INDICADORES dentro de cada hoja, y acumula los datos.
        Salta hojas que no existen.

        Returns:
            Lista de dicts, uno por línea encontrada, con keys:
            nombre, acciones (list), indicadores (list).
        """
        lineas: List[Dict[str, Any]] = []
        for n in range(1, self.cfg.max_lineas_estrategicas + 1):
            sn = f'LÍNEA ESTRATÉGICA {n}'
            try:
                df = pd.read_excel(ruta, sheet_name=sn, header=None)
            except (ValueError, KeyError):
                continue

            nombre = _ss(df.iloc[1, 1]) if len(df) > 1 and \
                len(df.columns) > 1 else ''
            acciones: List[Dict[str, str]] = []
            indicadores: List[Dict[str, str]] = []
            sec: Optional[str] = None

            for row in df.itertuples(index=False):
                va = _ss(row[0])
                if 'ACCIONES' in va.upper() and 'ACTIVIDADES' in va.upper():
                    sec = 'a'
                    continue
                if 'INDICADORES' in va.upper():
                    sec = 'i'
                    continue
                if va == '#' or not _es_numero(va):
                    continue

                if sec == 'a':
                    ac = _safe_col(row, 1)
                    at = _safe_col(row, 2)
                    av = _safe_col(row, 3)
                    es = _safe_col(row, 4)
                    if ac or av:
                        acciones.append({
                            'accion': ac, 'actividad': at,
                            'avance': av, 'estado': es,
                        })
                elif sec == 'i':
                    ind = _safe_col(row, 1)
                    mt = _safe_col(row, 2)
                    av = _safe_col(row, 3)
                    ob = _safe_col(row, 4)
                    if ind:
                        indicadores.append({
                            'indicador': ind, 'meta': mt,
                            'avance': av, 'observaciones': ob,
                        })

            lineas.append({
                'nombre': nombre,
                'acciones': acciones,
                'indicadores': indicadores,
            })
        return lineas

    def _leer_casos(self, ruta: str) -> List[Dict[str, str]]:
        """Lee hoja CASOS DE ÉXITO. Retorna lista vacía si no existe.

        Filtra filas cuya primera columna es un número (ID).

        Returns:
            Lista de dicts {titulo, descripcion, eje}.
        """
        try:
            df = pd.read_excel(ruta, sheet_name='CASOS DE ÉXITO',
                              header=None, skiprows=2)
        except (ValueError, KeyError):
            log.warning(f"  ⚠️ Hoja 'CASOS DE ÉXITO' no encontrada en "
                        f"{os.path.basename(ruta)}")
            return []
        casos: List[Dict[str, str]] = []
        for row in df.itertuples(index=False):
            if not _es_numero(_ss(row[0])):
                continue
            t = _safe_col(row, 1)
            d = _safe_col(row, 2)
            e = _safe_col(row, 3)
            if t:
                casos.append({
                    'titulo': t, 'descripcion': d, 'eje': e,
                })
        return casos

    def _leer_metas(self, ruta: str) -> List[Dict[str, str]]:
        """Lee hoja METAS GENERALES. Retorna lista vacía si no existe.

        Filtra filas cuya primera columna es un número (ID).

        Returns:
            Lista de dicts {indicador, meta, avance}.
        """
        try:
            df = pd.read_excel(ruta, sheet_name='METAS GENERALES',
                              header=None, skiprows=2)
        except (ValueError, KeyError):
            log.warning(f"  ⚠️ Hoja 'METAS GENERALES' no encontrada en "
                        f"{os.path.basename(ruta)}")
            return []
        metas: List[Dict[str, str]] = []
        for row in df.itertuples(index=False):
            if not _es_numero(_ss(row[0])):
                continue
            ind = _safe_col(row, 1)
            mt = _safe_col(row, 2)
            av = _safe_col(row, 3)
            if ind:
                metas.append({
                    'indicador': ind, 'meta': mt, 'avance': av,
                })
        return metas


# ═══════════════════════════════════════════════════════════════════════
# LECTOR EXCEL — TERRITORIAL
# ═══════════════════════════════════════════════════════════════════════

class LectorExcelTerritorial:
    """Lee un Excel territorial y retorna dict estructurado.

    Hojas que lee: CONFIGURACIÓN, DOFA (via LectorExcel),
                   TENDENCIAS POR EJE, MP_LE1..EXP_LE4, METAS GENERALES.

    Returns (de .leer()):
        dict con keys: config, familia, dofa, tendencias_por_eje,
                       contribuciones, metas, lineas (vacía).
    """

    def __init__(self, config: Config) -> None:
        self.cfg = config

    def leer(self, ruta: str) -> Dict[str, Any]:
        """Lee un Excel territorial completo.

        Reutiliza LectorExcel base para config, dofa y metas.
        Lee tendencias por eje y contribuciones con métodos propios.

        Args:
            ruta: Ruta al archivo .xlsx.

        Returns:
            Dict con datos del territorio, incluyendo contribuciones
            a los 4 ejes misionales.
        """
        log.info(f"📖 Leyendo territorial: {os.path.basename(ruta)}")

        # Reutilizar lector base para config y dofa
        lector_base = LectorExcel(self.cfg)
        config = lector_base._leer_configuracion(ruta)
        dofa = lector_base._leer_dofa(ruta)
        metas = lector_base._leer_metas(ruta)

        # Tendencias por eje
        tend_por_eje = self._leer_tendencias_por_eje(ruta)

        # Contribuciones
        contribuciones: Dict[str, List[Dict]] = {}
        for eje_key in ORDEN_EJES:
            eje = EJES_REFERENCIA[eje_key]
            contrib_list: List[Dict] = []
            for le_idx in range(1, len(eje.lineas) + 1):
                sn = f'{eje.prefijo}_LE{le_idx}'
                try:
                    le_data = self._leer_hoja_contribucion(ruta, sn)
                    contrib_list.append(le_data)
                except (ValueError, KeyError):
                    contrib_list.append({'acciones': [], 'indicadores': []})
            contribuciones[eje_key] = contrib_list

        data: Dict[str, Any] = {
            'config': config,
            'familia': FamiliaUnidad.TERRITORIAL.value,
            'dofa': dofa,
            'tendencias_por_eje': tend_por_eje,
            'contribuciones': contribuciones,
            'metas': metas,
            'lineas': [],  # Territoriales no tienen líneas propias
        }

        n_contrib = sum(
            len(le.get('acciones', []))
            for les in contribuciones.values()
            for le in les
        )
        log.info(f"  → {config.get('Nombre de la unidad', '?')} | "
                 f"{n_contrib} acciones de contribución | "
                 f"{len(metas)} metas")
        return data

    def _leer_tendencias_por_eje(self, ruta: str) -> Dict[str, Dict[str, List[str]]]:
        """Lee la hoja TENDENCIAS POR EJE.

        Parseo secuencial con estado: detecta encabezados de eje
        (VP TURISMO, VP INVERSIÓN, VP EXPORTACIONES) y acumula
        las 3 secciones (tendencias, foco, aporte) debajo.

        Returns:
            Dict eje→{tendencias, foco, aporte}, donde cada
            sección es una lista de strings.
        """
        try:
            df = pd.read_excel(ruta, sheet_name='TENDENCIAS POR EJE',
                              header=None, skiprows=2)
        except (ValueError, KeyError):
            return {}

        result: Dict[str, Dict[str, List[str]]] = {}
        current_eje: Optional[str] = None
        sec_order = ['tendencias', 'foco', 'aporte']
        sec_idx = 0

        for row in df.itertuples(index=False):
            val0 = _ss(row[0])
            val1 = _safe_col(row, 1)

            # Detectar encabezado de eje
            for eje_key in ['TUR', 'INV', 'EXP']:
                eje_nombre = EJES_REFERENCIA[eje_key].nombre
                if eje_nombre in val0.upper():
                    current_eje = eje_key
                    result[eje_key] = {}
                    sec_idx = 0
                    break
            else:
                if current_eje and val1:
                    sec = sec_order[min(sec_idx, 2)]
                    items = [x.strip().lstrip('•').strip()
                             for x in val1.split('\n') if x.strip()]
                    result[current_eje][sec] = items
                    sec_idx += 1

        return result

    def _leer_hoja_contribucion(
        self, ruta: str, sheet: str,
    ) -> Dict[str, List[Dict[str, str]]]:
        """Lee una hoja de contribución a eje (MP_LE1, TUR_LE2, etc.).

        Tiene la misma estructura interna que una línea estratégica
        misional: secciones ACCIONES e INDICADORES.

        Args:
            ruta: Ruta al archivo .xlsx.
            sheet: Nombre de la hoja (ej: 'MP_LE1').

        Returns:
            Dict con keys 'acciones' e 'indicadores', cada una
            con lista de dicts.
        """
        df = pd.read_excel(ruta, sheet_name=sheet, header=None)
        acciones: List[Dict[str, str]] = []
        indicadores: List[Dict[str, str]] = []
        sec: Optional[str] = None

        for row in df.itertuples(index=False):
            va = _ss(row[0])
            if 'ACCIONES' in va.upper():
                sec = 'a'
                continue
            if 'INDICADORES' in va.upper():
                sec = 'i'
                continue
            if va == '#' or not _es_numero(va):
                continue

            if sec == 'a':
                ac = _safe_col(row, 1)
                at = _safe_col(row, 2)
                av = _safe_col(row, 3)
                es = _safe_col(row, 4)
                if ac or av:
                    acciones.append({
                        'accion': ac, 'actividad': at,
                        'avance': av, 'estado': es,
                    })
            elif sec == 'i':
                ind = _safe_col(row, 1)
                mt = _safe_col(row, 2)
                av = _safe_col(row, 3)
                ob = _safe_col(row, 4)
                if ind:
                    indicadores.append({
                        'indicador': ind, 'meta': mt,
                        'avance': av, 'observaciones': ob,
                    })

        return {'acciones': acciones, 'indicadores': indicadores}


# ═══════════════════════════════════════════════════════════════════════
# LECTOR EXCEL — TRANSVERSAL
# ═══════════════════════════════════════════════════════════════════════

class LectorExcelTransversal:
    """Lee Excel transversal: líneas propias + contribuciones a ejes.

    Combina funcionalidad de LectorExcel (líneas propias)
    y LectorExcelTerritorial (contribuciones a ejes).

    Returns (de .leer()):
        dict con keys: config, familia, dofa, lineas, contribuciones,
                       metas, tendencias (vacía), casos_exito (vacía).
    """

    def __init__(self, config: Config) -> None:
        self.cfg = config
        self._base = LectorExcel(config)
        self._terr = LectorExcelTerritorial(config)

    def leer(self, ruta: str) -> Dict[str, Any]:
        """Lee un Excel transversal completo.

        Args:
            ruta: Ruta al archivo .xlsx.

        Returns:
            Dict con líneas propias + contribuciones a ejes.
        """
        log.info(f"📖 Leyendo transversal: {os.path.basename(ruta)}")

        config = self._base._leer_configuracion(ruta)
        dofa = self._base._leer_dofa(ruta)
        lineas = self._base._leer_lineas(ruta)
        metas = self._base._leer_metas(ruta)

        # Contribuciones (reutiliza lector territorial)
        contribuciones: Dict[str, List[Dict]] = {}
        for eje_key in ORDEN_EJES:
            eje = EJES_REFERENCIA[eje_key]
            contrib_list: List[Dict] = []
            for le_idx in range(1, len(eje.lineas) + 1):
                sn = f'{eje.prefijo}_LE{le_idx}'
                try:
                    le_data = self._terr._leer_hoja_contribucion(ruta, sn)
                    contrib_list.append(le_data)
                except (ValueError, KeyError):
                    contrib_list.append({'acciones': [], 'indicadores': []})
            contribuciones[eje_key] = contrib_list

        data: Dict[str, Any] = {
            'config': config,
            'familia': FamiliaUnidad.TRANSVERSAL.value,
            'dofa': dofa,
            'lineas': lineas,
            'contribuciones': contribuciones,
            'metas': metas,
            'tendencias': [],
            'casos_exito': [],
        }

        n_propias = sum(len(le.get('acciones', [])) for le in lineas)
        n_contrib = sum(
            len(le.get('acciones', []))
            for les in contribuciones.values() for le in les
        )
        log.info(f"  → {config.get('Nombre de la unidad', '?')} | "
                 f"{len(lineas)} líneas propias ({n_propias} acc) | "
                 f"{n_contrib} acciones contribución | {len(metas)} metas")
        return data
