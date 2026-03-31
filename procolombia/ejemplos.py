# -*- coding: utf-8 -*-
"""
Datos de ejemplo, guía de uso y banner.

Estos datos son para pruebas y demostraciones.
Se extraen de presentaciones reales de ProColombia.
"""

from __future__ import annotations
from typing import Dict


# ═══════════════════════════════════════════════════════════════════════
# DATOS DE EJEMPLO — VP TURISMO (MISIONAL)
# ═══════════════════════════════════════════════════════════════════════

def datos_ejemplo_turismo() -> Dict:
    """Retorna datos base de VP Turismo extraídos de la presentación real."""
    return {
        'dofa': {
            'DEBILIDADES': [
                'Demoras en procesos administrativos y jurídicos.',
                'Insuficiencia de material audiovisual para campañas de turismo arqueológico, gastronómico, musical, naturaleza, buceo, fiestas, pueblos patrimonio, cruceros y venues no tradicionales.',
                'Equipo insuficiente para aumento de presupuesto. Vacantes pendientes: Profesional Hotelería, Gerente Innovación, Asesor Cartagena, Asesor Reuniones EE.UU., Profesional Conectividad, Director Turismo EE.UU., Asesor Cruceros.',
                'Falta de estructuración de proyectos turísticos con VP Inversiones para atracción de IED en territorios.',
            ],
            'OPORTUNIDADES': [
                'Crecimiento en tendencias de turismo sostenible y comunitario.',
                'Oportunidades en mercados lejanos (Asia, Europa Oriental) ligadas a nuevas rutas aéreas (Emirates, Turkish, Qatar).',
                'Promoción a través de plataformas digitales y Netflix (Cien Años de Soledad II) y Copa Mundial 2026.',
                'Turismo de reuniones y bodas destino como segmentos estratégicos.',
                'Oportunidades en turismo de salud y turismo deportivo.',
                'Inteligencia artificial como aliada para optimizar procesos y analizar tendencias.',
                'Lanzamiento del Colombia Meetings Travel Mart.',
                'Diversificar y fortalecer nuevos mercados, reduciendo dependencia de mercados tradicionales.',
            ],
            'FORTALEZAS': [
                'ProColombia se consolida como referente en agencias de promoción turística.',
                'Modelo de Potencialidad que optimiza la estrategia de internacionalización.',
                'Equipos con experiencia y conocimiento de oferta y demanda.',
                'Regiones Turísticas con oferta diversa.',
                'Crecimiento en conectividad aérea con nuevos vuelos y rutas.',
                'Estrategias de promoción digital complementadas con métodos tradicionales.',
                'Posicionamiento en ferias internacionales y eventos especializados.',
                'Planes de trabajo exitosos con aliados nacionales e internacionales.',
            ],
            'AMENAZAS': [
                'Elevada carga tributaria sobre tiquetes aéreos disminuye competitividad.',
                'Percepción de seguridad negativa en destinos emergentes.',
                'Recuperación de mercados competidores post-pandemia (Perú, Asia).',
                'Conectividad interna insuficiente y escasez de aerolíneas internacionales hacia destinos emergentes.',
                'Incertidumbre geopolítica como riesgo para la estabilidad del sector.',
            ],
        },
        'tendencias': [
            'Turismo multi destino en auge: viajes a varios destinos combinando experiencias de vida silvestre, rutas gastronómicas y turismo deportivo.',
            'Hoteles activando nuevas fuentes de ingreso monetizando espacios subutilizados (pases diarios a no huéspedes para piscinas, spas, gimnasios).',
            'IA y analítica predictiva para ofertas inteligentes: incremento del 20% en ingreso por habitación.',
            'Modo IA de Google reconfigura la búsqueda a reserva directa, reduciendo intermediación de OTAs.',
            'Industria de reuniones innova con formatos en cruceros: talleres, sesiones de trabajo, logística simplificada.',
        ],
        'lineas': [
            {
                'nombre': 'Liderar el dinamismo en la conectividad aérea, marítima y transfronteriza.',
                'acciones': [
                    {'accion': 'Mantener y expandir la conectividad aérea internacional', 'actividad': 'Identificación y consolidación de nuevas rutas internacionales', 'avance': 'Se confirmaron 3 nuevas rutas Emirates, Turkish y Qatar hacia Colombia para 2026.', 'estado': 'En progreso'},
                    {'accion': 'Estrategias de marketing con aerolíneas', 'actividad': 'Promoción con Air Europa, KLM, Lufthansa, Edelweiss, IBERIA, Air Canada, WestJet', 'avance': 'Campañas conjuntas con 6 aerolíneas ejecutadas en Q1.', 'estado': 'En progreso'},
                    {'accion': 'Posicionar Colombia como destino de cruceros', 'actividad': 'Participación en FCCA, Expedition Cruise Conference, Seatrade Global', 'avance': 'Participación en Seatrade Miami con stand propio. 12 reuniones con navieras.', 'estado': 'Completada'},
                ],
                'indicadores': [
                    {'indicador': 'Visitantes no residentes', 'meta': '7.000.000', 'avance': '1.850.000', 'observaciones': 'Avance Q1'},
                    {'indicador': 'Frecuencias aéreas anuales', 'meta': '83.305', 'avance': '41.653', 'observaciones': 'Estimación GIC'},
                    {'indicador': 'Sillas aéreas anuales', 'meta': '15.986.459', 'avance': '7.993.230', 'observaciones': 'Estimación GIC'},
                ],
            },
            {
                'nombre': 'Desarrollar campañas y acciones segmentadas (B2B/B2C).',
                'acciones': [
                    {'accion': 'Ejecutar acciones digitales y tradicionales en mercados estratégicos', 'actividad': 'Campañas en mercados priorizados por modelo de potencialidad', 'avance': 'Campañas B2C lanzadas en 8 mercados. Alcance estimado 45M impresiones.', 'estado': 'En progreso'},
                    {'accion': 'Participación en ferias y eventos', 'actividad': 'BTL, FITUR, ITB, WTM, TTG, ATM, ITB Asia', 'avance': 'Participación en FITUR e ITB con delegaciones de 20+ empresarios.', 'estado': 'Completada'},
                    {'accion': 'Escenarios de encadenamiento turístico', 'actividad': 'Colombia Travel Mart, Nature Travel Mart, Ruedas de Encadenamiento', 'avance': 'CTM 2026 en planeación para septiembre.', 'estado': 'En progreso'},
                    {'accion': 'Plan de Medios con acciones de awareness', 'actividad': 'Estrategias de alto impacto y recordación de marca', 'avance': 'Plan de medios Q1 ejecutado con Netflix (Cien Años II).', 'estado': 'En progreso'},
                ],
                'indicadores': [
                    {'indicador': 'Visitantes no residentes', 'meta': '7.000.000', 'avance': '1.850.000', 'observaciones': ''},
                    {'indicador': 'Empresas nacionales con servicios', 'meta': '1.030', 'avance': '285', 'observaciones': ''},
                    {'indicador': 'Empresas exterior con servicios', 'meta': '3.352', 'avance': '890', 'observaciones': ''},
                ],
            },
            {
                'nombre': 'Promover a Colombia como destino de turismo de reuniones de alto impacto.',
                'acciones': [
                    {'accion': 'Participar en ferias internacionales de Industria de Reuniones', 'actividad': 'IMEX America, IMEX Frankfurt, IBTM World, Fiexpo Latinoamérica', 'avance': 'Stand en IMEX Frankfurt con 35 citas de negocio.', 'estado': 'Completada'},
                    {'accion': 'Sinergias con Bureaus y entidades regionales', 'actividad': 'Fortalecer red de Bureaux e impulsar industria de Reuniones', 'avance': 'Reunión Red Nacional de Bureaux en febrero. 8 Bureaux participantes.', 'estado': 'En progreso'},
                    {'accion': 'Aprovechamiento de membresías (ICCA, PCMA, SITE)', 'actividad': 'Participación en SITE y Destination Alliance', 'avance': 'Renovación membresía ICCA y SITE para 2026.', 'estado': 'Completada'},
                ],
                'indicadores': [
                    {'indicador': 'Visitantes no residentes', 'meta': '7.000.000', 'avance': '1.850.000', 'observaciones': ''},
                    {'indicador': 'Captación de congresos y eventos', 'meta': '565', 'avance': '120', 'observaciones': ''},
                ],
            },
            {
                'nombre': 'Fomentar la promoción a través de las seis regiones turísticas.',
                'acciones': [
                    {'accion': 'Ruta Exportadora de Turismo', 'actividad': 'PFE, Preparación y Adecuación, Club del Producto, Ruedas de Encadenamiento', 'avance': 'PFE con 6.400 participantes en Q1 (meta anual 15.600).', 'estado': 'En progreso'},
                    {'accion': 'Presentaciones de Destino B2B y acciones B2C', 'actividad': 'Presentaciones en Alemania, Argentina, Canadá, Corea, Ecuador, Francia, México', 'avance': 'Presentaciones de destino en 4 mercados completadas en Q1.', 'estado': 'En progreso'},
                    {'accion': 'Alianzas con operadores internacionales', 'actividad': 'Campañas con GADVENTURES, Intrepid, Kensington Tours, TUI, NUBA', 'avance': 'Alianza con Intrepid para 3 nuevas rutas en regiones emergentes.', 'estado': 'En progreso'},
                ],
                'indicadores': [
                    {'indicador': 'Visitantes no residentes', 'meta': '7.000.000', 'avance': '1.850.000', 'observaciones': ''},
                    {'indicador': 'Empresas nacionales con servicios', 'meta': '1.030', 'avance': '285', 'observaciones': ''},
                    {'indicador': 'Empresas exterior con servicios', 'meta': '3.352', 'avance': '890', 'observaciones': ''},
                    {'indicador': 'N° asistentes PFE', 'meta': '15.600', 'avance': '6.400', 'observaciones': ''},
                ],
            },
        ],
        'metas': [
            {'indicador': 'Eventos captados (Meta PES)', 'meta': '565', 'avance': '120'},
            {'indicador': 'Frecuencias internacionales anuales', 'meta': '83305', 'avance': '41653'},
            {'indicador': 'Sillas internacionales anuales', 'meta': '15986459', 'avance': '7993230'},
            {'indicador': 'N° Participantes PFE', 'meta': '15600', 'avance': '6400'},
            {'indicador': 'Visitantes no residentes (Meta País)', 'meta': '7000000', 'avance': '1850000'},
            {'indicador': 'Empresas nacionales con servicios', 'meta': '1030', 'avance': '285'},
            {'indicador': 'Empresas exterior con servicios', 'meta': '3352', 'avance': '890'},
        ],
    }


# ═══════════════════════════════════════════════════════════════════════
# BANNER DE INICIO
# ═══════════════════════════════════════════════════════════════════════


# ═══════════════════════════════════════════════════════════════════════
# DATOS DE EJEMPLO — HUB NORTEAMÉRICA (TERRITORIAL)
# ═══════════════════════════════════════════════════════════════════════

def datos_ejemplo_hub_norteamerica() -> Dict:
    """Datos reales extraídos de la presentación Hub Norteamérica."""
    return {
        'dofa': {
            'DEBILIDADES': [
                {'base': 'Carencia de proyectos de inversión estructurados a nivel nacional.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Presupuesto para actividades de promoción.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Recurso humano insuficiente en oficinas comerciales y en Colombia.', 'estado': 'Se actualiza', 'actualizacion': 'Se incorporó 1 asesor adicional en Miami para turismo de reuniones.'},
                {'base': 'Falta de herramientas de información comercial.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Percepción de inseguridad e inestabilidad política y social en Colombia.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Necesidad de fortalecer activaciones de Marca País y Plan de Medios segmentado.', 'estado': 'Se mantiene', 'actualizacion': ''},
            ],
            'OPORTUNIDADES': [
                {'base': 'Participación de Colombia en escenarios/foros geopolíticos.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Mejor aprovechamiento de los tratados comerciales vigentes.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Potencializar estrategia de nearshoring y fondos de inversión.', 'estado': 'Se actualiza', 'actualizacion': 'Nuevos aranceles de EE.UU. generan oportunidades adicionales de nearshoring.'},
                {'base': 'Seguimiento detallado a eventos y agendas de otras entidades.', 'estado': 'Se mantiene', 'actualizacion': ''},
            ],
            'FORTALEZAS': [
                {'base': 'Red de oficinas comerciales en 3 mercados estratégicos (EE.UU., Canadá, Caribe).', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Experiencia y reconocimiento internacional como agencia de promoción.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Relacionamiento sólido con aerolíneas, navieras y operadores turísticos.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Liderazgo en estrategia de conectividad aérea y cruceros.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Compromiso de colaboradores con la misionalidad de ProColombia.', 'estado': 'Se mantiene', 'actualizacion': ''},
            ],
            'AMENAZAS': [
                {'base': 'Cambios macroeconómicos, regulatorios y de seguridad que afecten IED.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Inestabilidad económica y política, percepción de inseguridad.', 'estado': 'Se actualiza', 'actualizacion': 'Descertificación por parte de EE.UU. agrava percepción.'},
                {'base': 'Baja competitividad en esquemas de incentivos para inversión.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Tensiones diplomáticas Colombia-EE.UU. impactan promoción.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Imposición de aranceles a países de la región por EE.UU.', 'estado': 'Se mantiene', 'actualizacion': ''},
            ],
        },
        'tendencias_por_eje': {
            'TUR': {
                'tendencias': [
                    'Resistencia al turismo por temas migratorios y de seguridad desde EE.UU.',
                    'Estrategias digitales y automatización de procesos.',
                    'Consolidación de rutas aéreas y operaciones de tour operadores.',
                    'Soft Travel: bienestar, simplicidad, off the grid experiences.',
                ],
                'foco': [
                    'Consolidación de conectividad aérea y marítima.',
                    'Estructuración de estrategias Black Travel, Turismo Indígena, Turismo Diverso.',
                    'Estrategia digital segmentada B2B y B2C.',
                ],
                'aporte': [
                    'Consolidación de Colombia como destino #1 del viajero de Norteamérica.',
                    'Liderazgo de la estrategia global de conectividad aérea y cruceros.',
                    'Estrategia digital fortalecida en turismo vacacional y reuniones (Cvent).',
                ],
            },
            'INV': {
                'tendencias': [
                    'Descertificación y tensiones diplomáticas elevan riesgo reputacional.',
                    'Oportunidades en proyectos ESG, joint ventures y nearshoring.',
                    'Colombia mantiene flujos de IED en energías renovables e infraestructura.',
                ],
                'foco': [
                    'Potenciar inversión en MRO, telecomunicaciones, agroindustria, TI y energías renovables.',
                    'Mantenimiento de relaciones con inversionistas actuales.',
                    'Fomentar proyectos con impacto ambiental y social positivo.',
                ],
                'aporte': [
                    'Apoyo continuo a inversionistas extranjeros establecidos.',
                    'Detección de leads de alto valor y colaboración entre áreas.',
                    'Focalizar esfuerzos en inversionistas alineados con proyectos disponibles.',
                ],
            },
            'EXP': {
                'tendencias': [
                    'Oportunidades en nuevos canales de compra.',
                    'Empresas con prácticas éticas, sostenibilidad, economía circular.',
                    'Preferencia de consumidor por productos Hechos no en China.',
                    'Imposición de aranceles de EE.UU. afecta comercio del HUB.',
                ],
                'foco': [
                    'Priorización de cadenas y sectores estratégicos por mercado.',
                    'Diversificación de exportaciones a nivel de producto, canal y territorios.',
                    'Aprovechamiento nearshoring y TLC.',
                ],
                'aporte': [
                    'Compradores de valor agregado.',
                    'Participación en ferias especializadas.',
                    'Información con fuentes primarias de los mercados.',
                    'Misiones clientes VIP a Colombia.',
                ],
            },
        },
        'contribuciones': {
            'MP': [
                {'acciones': [
                    {'accion': 'Identificar eventos y venues para visibilidad Marca País', 'actividad': 'Ferias 3 ejes, eventos embajadas, Mundial Fútbol, Super Bowl, Miami Open', 'avance': 'Activación Marca País en Super Bowl LIX con 15.000 impactos directos.', 'estado': 'Completada'},
                    {'accion': 'Apoyar estrategia digital de sostenimiento narrativa', 'actividad': 'Always on y plan de medios segmentado por mercado', 'avance': 'Campaña digital Q1 con 8M impresiones en EE.UU. y Canadá.', 'estado': 'En progreso'},
                    {'accion': 'Identificar influenciadores relevantes', 'actividad': 'Posicionamiento Marca País en mercados del Hub', 'avance': '5 influencers de viajes activados en Q1.', 'estado': 'En progreso'},
                ], 'indicadores': [
                    {'indicador': 'Activaciones Marca País', 'meta': '15', 'avance': '4', 'observaciones': ''},
                ]},
                {'acciones': [
                    {'accion': 'Identificar nuevos Embajadores de Marca País', 'actividad': 'Embajadores relevantes por mercado del HUB', 'avance': '2 candidatos identificados para EE.UU.', 'estado': 'En progreso'},
                    {'accion': 'Capitalizar serie Cien Años de Soledad', 'actividad': 'Alianzas Netflix y vinculación Marca País', 'avance': 'Evento de lanzamiento temporada 2 en planeación.', 'estado': 'En progreso'},
                    {'accion': 'Acercar Marca País a empresarios colombianos instalados', 'actividad': 'Desayunos de empresarios, eventos cámaras de comercio', 'avance': 'Desayuno empresarios en Miami con 45 asistentes.', 'estado': 'Completada'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Apoyar comercialización de productos Marca País', 'actividad': 'Tienda digital, ferias comerciales', 'avance': 'Presencia en 2 ferias con productos Marca País.', 'estado': 'En progreso'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Apoyar solicitudes institucionales de oficinas y Presidencia', 'actividad': 'Coordinación logística y promocional', 'avance': '3 solicitudes institucionales atendidas en Q1.', 'estado': 'Completada'},
                ], 'indicadores': []},
            ],
            'TUR': [
                {'acciones': [
                    {'accion': 'Mantener incremento de capacidad aérea desde Norteamérica', 'actividad': 'Nuevas rutas y frecuencias, west coast EE.UU. y Canadá', 'avance': '3 nuevas frecuencias confirmadas con Air Canada y JetBlue.', 'estado': 'En progreso'},
                    {'accion': 'Planes de promoción con aerolíneas co-financiados', 'actividad': 'Campañas conjuntas Air Europa, KLM, American Airlines', 'avance': 'Campaña con American Airlines ejecutada en Q1.', 'estado': 'Completada'},
                    {'accion': 'Fortalecer alianzas marítimas', 'actividad': 'AMAWaterways, FCCA, Royal Caribbean, cruceros expedición', 'avance': 'Alianza con AMAWaterways para Río Magdalena 2026.', 'estado': 'En progreso'},
                    {'accion': 'Consolidar cruceros fluviales Río Magdalena', 'actividad': 'Lanzamiento y promoción rutas fluviales', 'avance': 'Primer crucero fluvial operando desde enero 2026.', 'estado': 'Completada'},
                ], 'indicadores': [
                    {'indicador': 'Incremento frecuencias aéreas', 'meta': '5%', 'avance': '3%', 'observaciones': 'Estimación Q1'},
                    {'indicador': 'Recaladas de cruceros', 'meta': '350', 'avance': '95', 'observaciones': ''},
                ]},
                {'acciones': [
                    {'accion': 'Articulación plan de medios con estrategia turismo', 'actividad': 'Sinergias con colombia.travel, portal transaccional, OTAs', 'avance': 'Portal transaccional con 45.000 visitas en Q1.', 'estado': 'En progreso'},
                    {'accion': 'Consolidar estrategia digital y transaccional', 'actividad': 'Campañas con aerolíneas, OTA y metabuscadores', 'avance': 'Campañas con Expedia y Kayak activas.', 'estado': 'En progreso'},
                ], 'indicadores': [
                    {'indicador': 'Empresas con servicios (Exterior)', 'meta': '3.352', 'avance': '850', 'observaciones': ''},
                ]},
                {'acciones': [
                    {'accion': 'Participación en ferias de reuniones', 'actividad': 'IMEX America, IBTM, Fiexpo Latinoamérica', 'avance': 'Stand en IMEX America confirmado para octubre 2026.', 'estado': 'En progreso'},
                ], 'indicadores': [
                    {'indicador': 'Eventos captados desde Hub', 'meta': '80', 'avance': '18', 'observaciones': ''},
                ]},
                {'acciones': [
                    {'accion': 'Promoción de las seis regiones turísticas en Norteamérica', 'actividad': 'Presentaciones destino, FAM trips, press trips', 'avance': '2 FAM trips ejecutados (Andes y Pacífico).', 'estado': 'En progreso'},
                ], 'indicadores': []},
            ],
            'INV': [
                {'acciones': [
                    {'accion': 'Estrategia marketing inversión para aftercare', 'actividad': 'Comunicación digital con inversionistas instalados', 'avance': 'Newsletter trimestral enviado a 200+ inversionistas.', 'estado': 'Completada'},
                    {'accion': 'Podcast WHY COLOMBIA', 'actividad': 'Casos de éxito y testimonios de empresas instaladas', 'avance': '3 episodios grabados con empresas de Texas y Florida.', 'estado': 'En progreso'},
                ], 'indicadores': [
                    {'indicador': 'Empresas con servicios aftercare', 'meta': '120', 'avance': '35', 'observaciones': ''},
                ]},
                {'acciones': [
                    {'accion': 'Portafolio estratégico de oportunidades sostenibles', 'actividad': 'Maletín digital sectores clave: renovables, MRO, VC, TI', 'avance': 'Portafolio digital con 45 proyectos actualizado.', 'estado': 'Completada'},
                    {'accion': 'Promoción Vitrina de Oportunidades', 'actividad': 'Plataformas digitales para visibilizar oportunidades', 'avance': '12 proyectos destacados en LinkedIn con 50K impresiones.', 'estado': 'En progreso'},
                    {'accion': 'Agendas de inversión con tomadores de decisiones', 'actividad': 'Reuniones individuales con líderes y gerentes', 'avance': '8 reuniones ejecutadas con fondos de inversión en NY.', 'estado': 'En progreso'},
                ], 'indicadores': [
                    {'indicador': 'Leads de inversión generados', 'meta': '50', 'avance': '14', 'observaciones': ''},
                ]},
                {'acciones': [
                    {'accion': 'Trabajar con APRIS para giras internacionales', 'actividad': 'Visitas a potenciales inversionistas en regiones', 'avance': 'Gira a Carolina del Norte ejecutada (manufactura).', 'estado': 'Completada'},
                ], 'indicadores': []},
            ],
            'EXP': [
                {'acciones': [
                    {'accion': 'Participar en escenarios comerciales claves', 'actividad': 'Ferias y eventos en EE.UU. y mercados del Hub', 'avance': 'Participación en Winter Fancy Food Show (enero) con 15 empresas.', 'estado': 'Completada'},
                    {'accion': 'Captura de nuevos compradores de valor agregado', 'actividad': 'Misiones clientes VIP, desayunos instalados', 'avance': 'Desayuno compradores en Houston con 8 importadores.', 'estado': 'En progreso'},
                    {'accion': 'Estrategia de nearshoring/Why Colombia', 'actividad': 'Promoción ventajas Colombia en nearshoring', 'avance': 'Presentación Why Colombia en conferencia CEAL Miami.', 'estado': 'Completada'},
                ], 'indicadores': [
                    {'indicador': 'Oportunidades de exportación generadas', 'meta': '2.000', 'avance': '520', 'observaciones': ''},
                    {'indicador': 'Compradores con negocios', 'meta': '800', 'avance': '195', 'observaciones': ''},
                ]},
                {'acciones': [
                    {'accion': 'Identificar nuevos mercados y nichos en EE.UU.', 'actividad': 'Explorar estados poco explorados, canales alternativos', 'avance': 'Mapeo de oportunidades en 5 nuevos estados completado.', 'estado': 'Completada'},
                    {'accion': 'Estrategia sinergia con VP Turismo para canal cruceros', 'actividad': 'Cruce de oferta exportable con canal cruceros/hoteles', 'avance': 'Piloto de showcase de productos colombianos en crucero AMAWaterways.', 'estado': 'En progreso'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Acompañamiento sofisticación oferta exportable', 'actividad': 'Cierre de brechas, certificaciones, empaques', 'avance': '12 empresas en programa de adecuación para mercado EE.UU.', 'estado': 'En progreso'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Capacitación cultura exportadora', 'actividad': 'Webinars, talleres, formación en comercio exterior', 'avance': '3 webinars realizados con 280 participantes.', 'estado': 'En progreso'},
                ], 'indicadores': [
                    {'indicador': 'Participantes en formación', 'meta': '1.500', 'avance': '280', 'observaciones': ''},
                ]},
            ],
        },
        'metas': [
            {'indicador': 'Visitantes no residentes desde Hub', 'meta': '1.200.000', 'avance': '310.000'},
            {'indicador': 'Oportunidades exportación', 'meta': '2.000', 'avance': '520'},
            {'indicador': 'Compradores con negocios', 'meta': '800', 'avance': '195'},
            {'indicador': 'Proyectos de inversión atraídos', 'meta': '25', 'avance': '6'},
            {'indicador': 'Eventos captados (reuniones)', 'meta': '80', 'avance': '18'},
            {'indicador': 'Empresas con servicios aftercare', 'meta': '120', 'avance': '35'},
        ],
    }


# ═══════════════════════════════════════════════════════════════════════
# DATOS DE EJEMPLO — GIC (TRANSVERSAL)
# ═══════════════════════════════════════════════════════════════════════

def datos_ejemplo_gic() -> Dict:
    """Datos reales de GIC extraídos de la presentación."""
    return {
        'dofa': {
            'DEBILIDADES': [
                {'base': 'Vacantes en el equipo (licencia maternidad, Senior Logística, Junior Proyectos).', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Continuar mejorando capacidades técnicas con nuevas tecnologías emergentes.', 'estado': 'Se actualiza', 'actualizacion': 'Se inició programa de formación en IA generativa para 8 analistas.'},
                {'base': 'Falta de posicionamiento en la entidad - divulgación portafolio de servicios.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Falta de canales de comunicación claros.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Exceso de trabajo reactivo y cortos tiempos de respuesta.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Alta dependencia del área de tecnología.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Duplicidad de plantillas y procesos que genera desgaste.', 'estado': 'Se actualiza', 'actualizacion': 'Se consolidaron 3 plantillas en 1 formato unificado.'},
            ],
            'OPORTUNIDADES': [
                {'base': 'Alinear y capacitar equipos nuevos OFICOM y OFIREG.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Proporcionar capacitación en IA, Python, gestión de proyectos.', 'estado': 'Se actualiza', 'actualizacion': 'Bootcamp interno de Python con 15 participantes en Q1.'},
                {'base': 'Propender por toma de decisiones basadas en datos.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Automatización de procesos.', 'estado': 'Se actualiza', 'actualizacion': 'Automatización de seguimiento de planeación en desarrollo.'},
                {'base': 'Reconectar desde lo sectorial con las OFICOM.', 'estado': 'Se mantiene', 'actualizacion': ''},
            ],
            'FORTALEZAS': [
                {'base': 'Equipo actualizándose en nuevas tecnologías y en constante capacitación.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Conocimiento especializado a nivel de sectores y mercados.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Habilidades analíticas, investigación de mercados, recursividad.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Equipo diverso en formación profesional.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Red de contactos y expertos para obtención de información.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Trabajos altamente técnicos bajo metodologías de investigación.', 'estado': 'Se mantiene', 'actualizacion': ''},
            ],
            'AMENAZAS': [
                {'base': 'Infraestructura tecnológica insuficiente para herramientas de inteligencia.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'GIC es la única área transversal con salario variable.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Falta de articulación con MINCIT y entidades del Gobierno.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'Desarticulación entre equipo directivo VPs.', 'estado': 'Se mantiene', 'actualizacion': ''},
                {'base': 'No hay acceso eficiente a bases de datos de TI.', 'estado': 'Se actualiza', 'actualizacion': 'Se logró acceso directo a 2 bases previamente restringidas.'},
            ],
        },
        'lineas': [
            {'nombre': 'Generar y transferir conocimiento (análisis e información) para apoyar la estrategia de ProColombia.',
             'acciones': [
                 {'accion': 'Generar información de valor agregado y especializada', 'actividad': 'Identificación de oportunidades y entendimiento de mercados', 'avance': 'Publicados 12 informes sectoriales y 8 perfiles de mercado en Q1.', 'estado': 'En progreso'},
                 {'accion': 'Mantener actualizadas cajas de herramientas sectoriales', 'actividad': 'Actualización para Exportaciones, IED y Turismo', 'avance': 'Caja de herramientas Turismo actualizada. Exportaciones en proceso.', 'estado': 'En progreso'},
                 {'accion': 'Levantamiento de información primaria de tendencias', 'actividad': 'Ferias, eventos, entrevistas con expertos', 'avance': 'Participación en FITUR e ITB para levantamiento de tendencias 2026.', 'estado': 'Completada'},
             ], 'indicadores': [
                 {'indicador': 'Informes sectoriales publicados', 'meta': '48', 'avance': '12', 'observaciones': 'Q1'},
                 {'indicador': 'Perfiles de mercado actualizados', 'meta': '30', 'avance': '8', 'observaciones': ''},
             ]},
            {'nombre': 'Desarrollar procesos de gestión del conocimiento y acompañamiento metodológico.',
             'acciones': [
                 {'accion': 'Actualizar modelos de potencialidad', 'actividad': 'Turismo, turismo corporativo, asociativo', 'avance': 'Modelo de potencialidad turismo recalibrado con datos 2025.', 'estado': 'Completada'},
                 {'accion': 'Actualizar y centralizar fuentes en Biblioteca ProColombia', 'actividad': 'Facilitar acceso, autogestión y análisis de información', 'avance': '15 nuevas fuentes integradas a la Biblioteca ProColombia.', 'estado': 'En progreso'},
                 {'accion': 'Transformar información con diseño y visualizaciones', 'actividad': 'Mejorar comprensión y usabilidad de insumos estratégicos', 'avance': '4 dashboards interactivos creados para VPT y VPE.', 'estado': 'En progreso'},
             ], 'indicadores': [
                 {'indicador': 'Modelos de potencialidad actualizados', 'meta': '4', 'avance': '1', 'observaciones': ''},
                 {'indicador': 'Fuentes integradas a Biblioteca', 'meta': '40', 'avance': '15', 'observaciones': ''},
             ]},
            {'nombre': 'Promover sinergias internas y externas para generación de información de calidad.',
             'acciones': [
                 {'accion': 'Intercambiar información entre equipos sectoriales', 'actividad': 'Reuniones periódicas con VPT, VPI, VPE', 'avance': 'Comités mensuales establecidos con las 3 VPs.', 'estado': 'En progreso'},
                 {'accion': 'Fortalecer relacionamiento con proveedores logísticos', 'actividad': 'Incrementar oferta del Directorio Logístico', 'avance': '8 nuevos proveedores incorporados al directorio.', 'estado': 'En progreso'},
             ], 'indicadores': [
                 {'indicador': 'Reuniones de sinergia realizadas', 'meta': '36', 'avance': '9', 'observaciones': 'Trimestral'},
             ]},
        ],
        'contribuciones': {
            'MP': [
                {'acciones': [{'accion': 'Apoyar con datos e información la estrategia de Marca País', 'actividad': 'Informes de percepción y posicionamiento', 'avance': 'Informe de percepción Marca País 2025 entregado.', 'estado': 'Completada'}], 'indicadores': []},
                {'acciones': [], 'indicadores': []},
                {'acciones': [], 'indicadores': []},
                {'acciones': [{'accion': 'Soporte analítico a solicitudes institucionales', 'actividad': 'Datos para presentaciones de Presidencia', 'avance': '5 solicitudes atendidas con datos en Q1.', 'estado': 'En progreso'}], 'indicadores': []},
            ],
            'TUR': [
                {'acciones': [
                    {'accion': 'Actualizar modelos de potencialidad turismo y conectividad', 'actividad': 'Enfocar actividades de promoción en mercados clave', 'avance': 'Modelo recalibrado con datos consolidados 2025.', 'estado': 'Completada'},
                    {'accion': 'Generar información para mejorar conectividad aérea', 'actividad': 'Conjunto con VPT, aerolíneas con cupos de carga', 'avance': 'Estudio de conectividad para 3 nuevas rutas entregado.', 'estado': 'Completada'},
                ], 'indicadores': [{'indicador': 'Estudios de conectividad', 'meta': '6', 'avance': '2', 'observaciones': ''}]},
                {'acciones': [
                    {'accion': 'Actualizar caja de herramientas turismo y Why Colombia', 'actividad': 'Perfiles de turista y mercados priorizados', 'avance': '4 perfiles de turista actualizados.', 'estado': 'En progreso'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Soporte analítico para industria de reuniones', 'actividad': 'Datos para captación de eventos', 'avance': 'Base de datos de eventos actualizada con 200+ registros.', 'estado': 'En progreso'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Actualizar medición de impacto actividades turismo', 'actividad': 'Evaluación de actividades de promoción', 'avance': 'Metodología de medición de impacto Q4 2025 revisada.', 'estado': 'Completada'},
                ], 'indicadores': []},
            ],
            'INV': [
                {'acciones': [
                    {'accion': 'Actualizar caja de herramientas para promoción de IED', 'actividad': 'Herramientas sectoriales de inversión', 'avance': 'Actualización de fichas sectoriales para 5 sectores prioritarios.', 'estado': 'En progreso'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Construir estudios de mercado para inversión', 'actividad': 'Cartillas y estudios para mercados priorizados', 'avance': '2 estudios de mercado completados (MRO y TI).', 'estado': 'Completada'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Soporte a identificación de proyectos regionales', 'actividad': 'Información de valor de regiones para IED', 'avance': 'Mapeo de 15 proyectos regionales con datos GIC.', 'estado': 'En progreso'},
                ], 'indicadores': []},
            ],
            'EXP': [
                {'acciones': [
                    {'accion': 'Actualizar caja de herramientas exportaciones NME', 'actividad': 'Información para bienes NME y servicios', 'avance': 'Caja de herramientas servicios actualizada Q1.', 'estado': 'Completada'},
                    {'accion': 'Acompañar procesos de formación exportadora', 'actividad': 'Capacitaciones y metodologías de internacionalización', 'avance': '3 sesiones de capacitación a OFICOM sobre uso de herramientas.', 'estado': 'En progreso'},
                ], 'indicadores': [{'indicador': 'Capacitaciones realizadas', 'meta': '12', 'avance': '3', 'observaciones': ''}]},
                {'acciones': [
                    {'accion': 'Construir estudios de mercado para diversificación', 'actividad': 'Cartillas de mercados con potencial exportador', 'avance': '2 cartillas de mercados emergentes (Vietnam y Polonia).', 'estado': 'Completada'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Fortalecer Directorio Logístico', 'actividad': 'Ampliar red de proveedores logísticos', 'avance': '8 nuevos proveedores incorporados.', 'estado': 'En progreso'},
                ], 'indicadores': []},
                {'acciones': [
                    {'accion': 'Actualizar herramientas estadísticas: PADEX y segmentación', 'actividad': 'Modelos de potencialidad y segmentación de empresas', 'avance': 'PADEX recalibrado con datos exportación 2025.', 'estado': 'Completada'},
                ], 'indicadores': []},
            ],
        },
        'metas': [
            {'indicador': 'Informes sectoriales publicados', 'meta': '48', 'avance': '12'},
            {'indicador': 'Perfiles de mercado actualizados', 'meta': '30', 'avance': '8'},
            {'indicador': 'Modelos de potencialidad actualizados', 'meta': '4', 'avance': '1'},
            {'indicador': 'Fuentes integradas a Biblioteca ProColombia', 'meta': '40', 'avance': '15'},
            {'indicador': 'Reuniones de sinergia realizadas', 'meta': '36', 'avance': '9'},
            {'indicador': 'Capacitaciones a OFICOM/OFIREG', 'meta': '12', 'avance': '3'},
            {'indicador': 'Estudios de mercado producidos', 'meta': '20', 'avance': '6'},
        ],
    }


# ═══════════════════════════════════════════════════════════════════════
# GUÍA DE USO EN GOOGLE COLAB
# ═══════════════════════════════════════════════════════════════════════



# ═══════════════════════════════════════════════════════════════════════
# BANNER Y GUÍA
# ═══════════════════════════════════════════════════════════════════════

def banner():
    print("""
╔══════════════════════════════════════════════════════════════════╗
║  AUTOMATIZACIÓN v5 — Planeación Estratégica ProColombia         ║
║                                                                  ║
║  📁 01_excels_entrada/  ← Excel diligenciados                   ║
║  📁 02_pptx_salida/     ← Presentaciones generadas              ║
║  📁 03_consolidado/     ← Excel maestro                         ║
║  📁 04_plantillas/      ← Plantillas PPTX por familia           ║
║                                                                  ║
║  Familias: MISIONAL | TERRITORIAL | TRANSVERSAL                 ║
║  Soporte: hasta 5 líneas estratégicas + eliminación dinámica    ║
╚══════════════════════════════════════════════════════════════════╝
    """)


def guia_colab():
    """Imprime la guía de uso rápida en Google Colab."""
    print("""
╔══════════════════════════════════════════════════════════════════╗
║        GUÍA DE USO EN GOOGLE COLAB — ProColombia v5.1           ║
╚══════════════════════════════════════════════════════════════════╝

═══════════════════════════════════════════════════════════════════
  PASO 1: PREPARACIÓN (una sola vez)
═══════════════════════════════════════════════════════════════════

  1.1 Suba la carpeta 'procolombia/' completa a Google Drive:

      📁 VPP/
      ├── 📁 procolombia/           ← Paquete Python (7 archivos)
      │   ├── __init__.py
      │   ├── config.py             ← Configuración (lo que usted toca)
      │   ├── utils.py
      │   ├── excel.py
      │   ├── pptx_gen.py
      │   ├── orquestador.py
      │   └── ejemplos.py
      ├── 📁 01_excels_entrada/
      ├── 📁 02_pptx_salida/
      ├── 📁 03_consolidado/
      └── 📁 04_plantillas/
          └── Plantilla_Misional.pptx

═══════════════════════════════════════════════════════════════════
  PASO 2: CELDA DE CONFIGURACIÓN (una vez por sesión)
═══════════════════════════════════════════════════════════════════

  from google.colab import drive
  drive.mount('/content/drive')
  !pip install python-pptx openpyxl -q

  import sys
  RUTA = "/content/drive/MyDrive/ProColombia/Automatizaciones/VPP"
  sys.path.insert(0, RUTA)

  from procolombia import *
  banner()

═══════════════════════════════════════════════════════════════════
  PASO 3: CONSTRUIR PLANTILLAS (una vez)
═══════════════════════════════════════════════════════════════════

  orq = OrquestadorUniversal(base_dir=RUTA)
  orq.construir_plantillas()

═══════════════════════════════════════════════════════════════════
  PASO 4: GENERAR EXCEL PARA LAS ÁREAS
═══════════════════════════════════════════════════════════════════

  orq.generar_excel('VP Turismo', 'EJE', trimestre='Q1',
                    anio='2026', num_lineas=4)

═══════════════════════════════════════════════════════════════════
  PASO 5: PROCESAR (después del diligenciamiento)
═══════════════════════════════════════════════════════════════════

  resultados = orq.procesar_lote()

═══════════════════════════════════════════════════════════════════
  PASO 6: CONSOLIDAR (opcional)
═══════════════════════════════════════════════════════════════════

  ruta_consolidado = orq.consolidar()

═══════════════════════════════════════════════════════════════════
  PERSONALIZACIÓN
═══════════════════════════════════════════════════════════════════

  cfg = Config(max_lineas_estrategicas=4, password='mi_clave')
  orq = OrquestadorUniversal(config=cfg, base_dir=RUTA)
""")
