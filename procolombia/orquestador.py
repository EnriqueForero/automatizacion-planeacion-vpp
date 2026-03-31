# -*- coding: utf-8 -*-
"""
OrquestadorUniversal — Punto de entrada principal del sistema.

Router inteligente que detecta la familia de cada Excel
y selecciona automáticamente el lector, generador y plantilla correctos.
"""

from __future__ import annotations

import os
import glob
from datetime import datetime
from typing import Dict, List, Optional

import pandas as pd
import openpyxl

from .config import Config, FamiliaUnidad, EJES_REFERENCIA, ORDEN_EJES
from .utils import _safe_filename, log
from .excel_constructores import (
    ConstructorExcelMisional,
    ConstructorExcelTerritorial,
    ConstructorExcelTransversal,
)
from .excel_lectores import (
    LectorExcel,
    LectorExcelTerritorial,
    LectorExcelTransversal,
)
from .pptx_gen import (
    GeneradorPPTXMisional,
    GeneradorPPTXTerritorial,
    GeneradorPPTXTransversal,
    ConstructorPlantillaTerritorial,
    ConstructorPlantillaTransversal,
)

class OrquestadorUniversal:
    """
    Router inteligente que:
    1. Lee CONFIGURACIÓN de cada Excel
    2. Detecta familia automáticamente
    3. Selecciona plantilla y generador correctos
    4. Procesa por lotes cualquier mezcla de familias
    """

    def __init__(self, config: Config = None, base_dir: str = '.'):
        self.cfg = config or Config()
        self.base_dir = base_dir
        self.cfg.crear_carpetas(base_dir)

        self._constructores = {
            FamiliaUnidad.MISIONAL: ConstructorExcelMisional(self.cfg),
            FamiliaUnidad.TERRITORIAL: ConstructorExcelTerritorial(self.cfg),
            FamiliaUnidad.TRANSVERSAL: ConstructorExcelTransversal(self.cfg),
        }
        self._generadores = {
            FamiliaUnidad.MISIONAL: GeneradorPPTXMisional(self.cfg),
            FamiliaUnidad.TERRITORIAL: GeneradorPPTXTerritorial(self.cfg),
            FamiliaUnidad.TRANSVERSAL: GeneradorPPTXTransversal(self.cfg),
        }
        self._lectores = {
            FamiliaUnidad.MISIONAL: LectorExcel(self.cfg),
            FamiliaUnidad.TERRITORIAL: LectorExcelTerritorial(self.cfg),
            FamiliaUnidad.TRANSVERSAL: LectorExcelTransversal(self.cfg),
        }

    def _detectar_familia(self, ruta: str) -> FamiliaUnidad:
        """Detecta familia examinando las hojas del Excel."""
        wb = openpyxl.load_workbook(ruta, read_only=True)
        sheets = set(wb.sheetnames)
        wb.close()

        tiene_contrib = 'MP_LE1' in sheets
        tiene_lineas_propias = 'LÍNEA ESTRATÉGICA 1' in sheets

        if tiene_contrib and tiene_lineas_propias:
            return FamiliaUnidad.TRANSVERSAL
        elif tiene_contrib:
            return FamiliaUnidad.TERRITORIAL
        else:
            return FamiliaUnidad.MISIONAL

    def _leer_excel(self, ruta: str,
                    familia: FamiliaUnidad) -> Dict:
        """Selecciona el lector correcto según familia."""
        return self._lectores[familia].leer(ruta)

    def generar_excel(self, nombre: str, tipo: str, **kwargs) -> str:
        """Genera Excel detectando familia desde el tipo."""
        familia = FamiliaUnidad.desde_tipo(tipo)
        constructor = self._constructores[familia]
        output_dir = os.path.join(self.base_dir, self.cfg.dir_entrada)
        return constructor.generar(nombre, tipo, output_dir=output_dir,
                                  **kwargs)

    def construir_plantillas(self) -> Dict[str, str]:
        """Construye las 3 plantillas PPTX en 04_plantillas/."""
        dir_p = os.path.join(self.base_dir, self.cfg.dir_plantillas)
        rutas = {}

        # Misional: solo verificar que exista (se copia de la institucional)
        r_m = os.path.join(dir_p, self.cfg.plantilla_misional)
        if os.path.exists(r_m):
            rutas['MISIONAL'] = r_m
            log.info(f"✅ Plantilla misional existente: {r_m}")
        else:
            log.warning(f"⚠️  Plantilla misional no encontrada: {r_m}")

        # Territorial
        r_t = os.path.join(dir_p, self.cfg.plantilla_territorial)
        ConstructorPlantillaTerritorial(self.cfg).construir(r_t)
        rutas['TERRITORIAL'] = r_t

        # Transversal
        r_tr = os.path.join(dir_p, self.cfg.plantilla_transversal)
        ConstructorPlantillaTransversal(self.cfg).construir(r_tr)
        rutas['TRANSVERSAL'] = r_tr

        return rutas

    def procesar_lote(self) -> List[Dict]:
        """Procesa todos los Excel en entrada, detectando familia."""
        dir_in = os.path.join(self.base_dir, self.cfg.dir_entrada)
        archivos = sorted(glob.glob(os.path.join(dir_in, '*.xlsx')))
        if not archivos:
            log.warning(f"⚠️  Sin archivos .xlsx en {dir_in}/")
            return []

        print(f"\n{'='*70}")
        print(f"  PROCESAMIENTO UNIVERSAL — {len(archivos)} archivos")
        print(f"{'='*70}\n")

        resultados = []
        for archivo in archivos:
            try:
                familia = self._detectar_familia(archivo)
                log.info(f"🔍 {os.path.basename(archivo)} → "
                         f"familia detectada: {familia.value}")
                data = self._leer_excel(archivo, familia)
                cfg_d = data['config']
                nombre = cfg_d.get('Nombre de la unidad', 'Sin_Nombre')
                q = cfg_d.get('Trimestre en seguimiento', 'Q1')
                año = cfg_d.get('Año', '2026')
                tipo = cfg_d.get('Tipo de unidad', 'EJE')

                ruta_plantilla = self.cfg.ruta_plantilla(
                    familia, self.base_dir)
                if not ruta_plantilla:
                    raise FileNotFoundError(
                        f"Plantilla {familia.value} no encontrada")

                safe = _safe_filename(nombre)
                nombre_pptx = (f'{q}_{año}_{tipo}_{safe}_Seguimiento.pptx')
                dir_out = os.path.join(self.base_dir, self.cfg.dir_salida)
                ruta_pptx = os.path.join(dir_out, nombre_pptx)

                generador = self._generadores[familia]
                generador.generar(data, str(ruta_plantilla), ruta_pptx)

                n_le = len([le for le in data.get('lineas', [])
                           if GeneradorPPTXMisional._linea_tiene_contenido(le)])
                resultados.append({
                    'archivo': os.path.basename(archivo),
                    'unidad': nombre, 'tipo': tipo,
                    'familia': familia.value,
                    'lineas': n_le, 'pptx': nombre_pptx,
                    'status': '✅'
                })
            except Exception as e:
                log.error(f"❌ {os.path.basename(archivo)}: {e}")
                import traceback; traceback.print_exc()
                resultados.append({
                    'archivo': os.path.basename(archivo),
                    'unidad': '?', 'tipo': '?', 'familia': '?',
                    'lineas': 0, 'pptx': '-',
                    'status': f'❌ {e}'
                })

        print(f"\n{'='*70}")
        print(f"  RESULTADOS")
        print(f"{'='*70}")
        for r in resultados:
            print(f"  {r['status']} {r['unidad']:30s} "
                  f"({r['familia']:12s}) → {r['pptx']}")
        err = sum(1 for r in resultados if '❌' in r['status'])
        ok = len(resultados) - err
        print(f"\n  ✅ {ok} exitosos | ❌ {err} errores\n")
        return resultados

    def consolidar(self) -> Optional[str]:
        """Consolida todos los Excel de entrada en un maestro."""
        dir_in = os.path.join(self.base_dir, self.cfg.dir_entrada)
        archivos = sorted(glob.glob(os.path.join(dir_in, '*.xlsx')))
        if not archivos:
            return None

        rows_resumen, rows_dofa = [], []
        rows_acciones_propias, rows_contribuciones = [], []

        for archivo in archivos:
            try:
                familia = self._detectar_familia(archivo)
                data = self._leer_excel(archivo, familia)
            except Exception as e:
                log.error(f"Error: {os.path.basename(archivo)}: {e}")
                continue

            cfg_d = data['config']
            base = {
                'Unidad': cfg_d.get('Nombre de la unidad', '?'),
                'Tipo': cfg_d.get('Tipo de unidad', '?'),
                'Familia': familia.value,
                'Trimestre': cfg_d.get('Trimestre en seguimiento', '?'),
                'Año': cfg_d.get('Año', '?'),
            }

            # DOFA
            for cuad, items in data.get('dofa', {}).items():
                for i, it in enumerate(items, 1):
                    rows_dofa.append({**base, 'Cuadrante': cuad,
                                     'Item': i,
                                     'Base': it.get('base', ''),
                                     'Estado': it.get('estado', ''),
                                     'Actualización': it.get('actualizacion', '')})

            # Líneas propias
            for li, le in enumerate(data.get('lineas', []), 1):
                for ai, ac in enumerate(le.get('acciones', []), 1):
                    rows_acciones_propias.append({
                        **base, 'Línea': li, 'Nombre_LE': le.get('nombre', ''),
                        'Acción': ai, 'Texto': ac.get('accion', ''),
                        'Avance': ac.get('avance', ''),
                        'Estado': ac.get('estado', '')})

            # Contribuciones
            for eje_key, contrib_list in data.get('contribuciones', {}).items():
                for le_idx, le_data in enumerate(contrib_list, 1):
                    for ai, ac in enumerate(le_data.get('acciones', []), 1):
                        rows_contribuciones.append({
                            **base, 'Eje': eje_key,
                            'Línea_Eje': le_idx,
                            'Acción': ai,
                            'Texto': ac.get('accion', ''),
                            'Avance': ac.get('avance', ''),
                            'Estado': ac.get('estado', '')})

            td = sum(len(v) for v in data.get('dofa', {}).values())
            rows_resumen.append({
                **base,
                'DOFA_Items': td,
                'Líneas_Propias': len(data.get('lineas', [])),
                'Acciones_Propias': sum(len(le.get('acciones', []))
                                       for le in data.get('lineas', [])),
                'Contribuciones': sum(len(le.get('acciones', []))
                                     for les in data.get('contribuciones', {}).values()
                                     for le in les),
                'Metas': len(data.get('metas', [])),
            })

        ts = datetime.now().strftime('%Y%m%d_%H%M')
        dir_out = os.path.join(self.base_dir, self.cfg.dir_consolidado)
        ruta = os.path.join(dir_out, f'Consolidado_Universal_{ts}.xlsx')

        with pd.ExcelWriter(ruta, engine='openpyxl') as writer:
            for nm, rows in [('RESUMEN', rows_resumen),
                             ('DOFA', rows_dofa),
                             ('ACCIONES_PROPIAS', rows_acciones_propias),
                             ('CONTRIBUCIONES', rows_contribuciones)]:
                if rows:
                    pd.DataFrame(rows).to_excel(writer, sheet_name=nm,
                                                index=False)

        log.info(f"📊 Consolidado: {ruta}")
        print(f"\n  📊 {os.path.basename(ruta)}")
        print(f"  Unidades: {len(rows_resumen)} | DOFA: {len(rows_dofa)}")
        print(f"  Acciones propias: {len(rows_acciones_propias)} | "
              f"Contribuciones: {len(rows_contribuciones)}")
        return ruta
