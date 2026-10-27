# -*- coding: utf-8 -*-
"""
ralumin/config.py — Todo lo que es propio de RALUMIN E.I.R.L. en un solo lugar.

Aquí se cambian grupos (archivadores), cuentas de pago por defecto, reglas contables de la empresa,
pares de reparto ADM/VTAS y umbrales. El motor no tiene nada de Ralumin escrito a fuego.
"""
from __future__ import annotations

from dataclasses import dataclass

from motor.reglas import Regla, REGLAS_SISTEMA

EMPRESA = 'RALUMIN E.I.R.L.'
RUC = '20602768440'


# ----------------------------------------------------------------------------- grupos (archivadores)
@dataclass(frozen=True)
class Grupo:
    id: str
    nombre: str
    moneda: str            # 'PEN' | 'USD' | '' (cualquiera)
    condicion: str         # 'CON' | 'CRE' | ''
    cuenta_pago: str       # cuenta de pago por defecto (Haber, columna AH)
    orden: int             # orden de importación / de pestañas


GRUPOS = [
    Grupo('caja_chica', 'Caja chica', 'PEN', 'CON', '1011', 1),
    Grupo('caja_chica_c23', 'Caja chica Calle 23', 'PEN', 'CON', '1011', 2),
    Grupo('soles_contado', 'Soles contado', 'PEN', 'CON', '104103', 3),
    Grupo('soles_contado_c23', 'Soles contado Calle 23', 'PEN', 'CON', '104103', 4),
    Grupo('usd_contado', 'Dólares contado', 'USD', 'CON', '104102', 5),
    Grupo('usd_credito', 'Dólares crédito', 'USD', 'CRE', '4212', 6),
    Grupo('usd_60_120', 'Dólares 60–120 días', 'USD', 'CRE', '4212', 7),
    Grupo('soles_credito', 'Soles crédito', 'PEN', 'CRE', '4212', 8),
    Grupo('sin_grupo', 'Sin grupo', '', '', '', 9),
]
GRUPOS_POR_ID = {g.id: g for g in GRUPOS}
GRUPO_NOMBRE = {g.id: g.nombre for g in GRUPOS}

# Cuenta de pago: solo estas raíces son válidas en AH (la 4212 únicamente con condición CRE).
RAICES_PAGO = ('101', '104')
CUENTA_CREDITO = '4212'

# Contado en soles va a caja chica si el total no pasa de este monto (y no hay otra pista).
UMBRAL_CAJA_CHICA = 700.0
# Plazo (días entre emisión y última cuota) desde el que una compra en dólares va a "60–120 días".
DIAS_60_120 = 60
# Palabras que indican un rubro típico de caja chica (peajes, grifos, estacionamientos, ferreterías).
KW_CAJA_CHICA = ('PEAJE', 'ESTACIONAMIENTO', 'PARQUEO', 'GASOHOL', 'GASOLINA', 'DIESEL', 'PETROLEO', 'GLP', 'GNV',
                 'COMBUSTIBLE', 'FERRETERIA', 'FERRETERÍA')

# ----------------------------------------------------------------------------- retenciones y comisiones
TASA_RETENCION = 3.0               # % del importe total
UMBRAL_RETENCION = 700.0           # suma del mismo proveedor pagada en la misma fecha (S/)
CUENTA_RETENCION = '401141'        # IGV - Retenciones por pagar
CUENTA_COMISION = '6391094'        # Gastos bancarios - ADM (la que usa Ralumin)
LEYENDAS_NO_RETENER = ('NO RETENER', 'BUEN CONTRIBUYENTE', 'AGENTE DE RETENCION', 'AGENTE DE RETENCIÓN',
                       'AGENTE DE PERCEPCION', 'AGENTE DE PERCEPCIÓN', 'NO SUJETO A RETENCION', 'NO SUJETO A RETENCIÓN')

# ----------------------------------------------------------------------------- reparto 60/40 ADM–VTAS
PORCENTAJE_ADM = 60.0
# Raíces de cuenta que Ralumin reparte entre la divisionaria ADM (…094 / …09401) y VTAS (…095 / …09501).
# Los peajes (6315) NO se reparten: van solo a 6315095 (decisión de Jiguel, 26/10/2026).
RAICES_REPARTO = ('6560', '656009', '6364', '6361', '6363', '6365', '6510', '6312', '6530', '6540', '6343', '6323')


def par_vtas(cuenta_adm: str) -> str:
    """6560094 → 6560095 ; 656009401 → 656009501. '' si la cuenta no es ADM."""
    c = (cuenta_adm or '').strip()
    if c.endswith('094'):
        return c[:-3] + '095'
    if c.endswith('09401'):
        return c[:-5] + '09501'
    return ''


def es_repartible(cuenta: str) -> bool:
    c = (cuenta or '').strip()
    return bool(par_vtas(c)) and c.startswith(RAICES_REPARTO)


# ----------------------------------------------------------------------------- reglas contables de Ralumin
# Mandan sobre la memoria y el plan (capa 3 del motor). Se pueden ampliar aquí.
REGLAS_RALUMIN = [
    Regla(id=-101, tipo='cuenta', nombre='Peajes → 6315095 (una sola cuenta)', palabras='PEAJE,LIMA EXPRESA,RUTAS DE LIMA,VIA EXPRESA,VÍA EXPRESA',
          cuenta='6315095', prioridad=10, sistema=1, nota='Decisión de Jiguel: los peajes van solo a 6315095, sin reparto 60/40.'),
    Regla(id=-102, tipo='cuenta', nombre='Lima Expresa → 6315095', ruc_proveedor='20523621212', cuenta='6315095', prioridad=9, sistema=1),
    Regla(id=-103, tipo='cuenta', nombre='Estacionamiento → 6315095', palabras='ESTACIONAMIENTO,PARQUEO,PLAYA DE ESTAC',
          cuenta='6315095', prioridad=12, sistema=1),
    Regla(id=-104, tipo='cuenta', nombre='Combustible → 656009401', palabras='GASOHOL,GASOLINA,DIESEL,PETROLEO,PETRÓLEO,GLP,GNV,COMBUSTIBLE',
          cuenta='656009401', prioridad=20, sistema=1, nota='Ralumin reparte combustible 60 % ADM (656009401) / 40 % VTAS (656009501).'),
    Regla(id=-105, tipo='cuenta', nombre='ITF → 6412094', palabras='IMPUESTO A LAS TRANSACCIONES FINANCIERAS,ITF ',
          cuenta='6412094', prioridad=20, sistema=1),
    Regla(id=-106, tipo='cuenta', nombre='Nacionalización → 609192012', palabras='NACIONALIZACION,NACIONALIZACIÓN,DESADUANAJE,AGENCIAMIENTO DE ADUANA,AGENTE DE ADUANA',
          cuenta='609192012', prioridad=25, sistema=1),
    Regla(id=-107, tipo='cuenta', nombre='Telefonía → 6364094', palabras='TELEFONIA,TELEFONÍA,PLAN POSTPAGO,SERVICIO MOVIL,SERVICIO MÓVIL,RECIBO TELEFONICO',
          cuenta='6364094', prioridad=30, sistema=1),
]

REGLAS = list(REGLAS_SISTEMA) + REGLAS_RALUMIN

# Compra de bienes sin pista (proveedor nuevo y sin palabra clave): en Ralumin lo más común es material auxiliar.
CUENTA_BIENES_DEFECTO = '6031025'
# Suministros genéricos (no combustible)
CUENTA_SUMINISTROS = '6560094'
KW_COMBUSTIBLE = ('GASOHOL', 'GASOLINA', 'DIESEL', 'PETROLEO', 'PETRÓLEO', 'GLP', 'GNV', 'COMBUSTIBLE', 'DB5', 'BIODIESEL')

# ----------------------------------------------------------------------------- glosas con el estilo de Ralumin
# raíz de la cuenta de gasto → concepto (sin la forma de pago). La más larga que coincida gana.
GLOSAS = {
    '6315': 'GASTOS DE VIAJE - PEAJE - ESTACIONAMIENTO, ETC.',
    '656009': 'SUMINISTROS - GASOLINA',
    '6560': 'COMPRA SUMINISTRO',
    '6031': 'COMPRA MATERIALES Y SUMINISTROS',
    '6011': 'COMPRA MERCADERIAS',
    '609192': 'NACIONALIZACION',
    '6364': 'SERVICIO TELEFONIA',
    '6361': 'SERVICIO ELECTRICIDAD',
    '6363': 'SERVICIO AGUA',
    '6365': 'SERVICIO INTERNET',
    '6311': 'SERVICIO TRANSPORTE DE CARGA',
    '6312': 'SERVICIOS DE CORREOS',
    '6510': 'SEGURO SCTR DE HONORARIOS, SOAT Y OTROS',
    '6273': 'SEGURO SCTR DE TRABAJADORES',
    '6373': 'CONSUMO CON RECARGO AL CONSUMO (GAST. REPRE.)',
    '6412': 'ITF',
    '6329': 'OTROS SERVICIOS',
    '6323': 'ASESOR Y CONSUL Y CONTABLE',
    '6322': 'LEGAL Y TRIBUTARIO',
    '6530': 'SUSCRIPCIONES Y COTIZACIONES',
    '6540': 'LICENCIAS, SOFT DERECHO DE USO',
    '6343': 'SERV MANT REPARACION INM MAQ Y EQUIPO',
}
# Conceptos que en el Diario de Ralumin no llevan forma de pago en la glosa.
GLOSAS_SIN_FORMA_PAGO = ('6315', '656009', '6412')
