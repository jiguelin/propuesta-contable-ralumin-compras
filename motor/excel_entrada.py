# -*- coding: utf-8 -*-
"""
motor/excel_entrada.py — Comprobantes de compra desde un Excel (cuando no se pueden descargar los XML).

El Excel tiene el mismo orden de columnas que la plantilla de importación de NewContaSis (A–AX), con las cuentas vacías:
    A emisión · B/AD vencimiento · C tipo · D serie · F número · G tipo doc. proveedor · H RUC · I razón social
    J base gravada · K IGV · P no gravadas · Q ISC · R otros tributos y cargos · S total      ← SIEMPRE en soles
    U/V constancia de detracción · W tipo de cambio · X–AA documento de referencia (NC/ND) · AC importe en US$
Reglas de lectura (indicadas por Jiguel):
  · Si AC tiene un valor, el comprobante es en DÓLARES (y W trae su tipo de cambio). En soles AC viene vacío y W = 1.
  · El Excel no trae ítems ni glosa ni forma de pago: la cuenta sale del historial del proveedor (Diario) y de su nombre;
    la condición CON/CRE, del vencimiento (AD > emisión → crédito) o del historial (si se le pagaba con 4212).
Las filas pueden empezar en cualquier fila (con o sin cabecera): se toma toda fila con fecha en A y tipo en C.

Cada fila se convierte en un `Comprobante` igual al de un XML, así el resto del motor no cambia. Los importes originales
en soles y el TC se guardan aparte para exportarlos tal cual.
"""
from __future__ import annotations

import io
import re
from datetime import date, datetime
from decimal import ROUND_HALF_UP, Decimal

import openpyxl

from .xml_parser import TIPOS_DOC, Comprobante

COL = {k: i for i, k in enumerate(
    ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z',
     'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV',
     'AW', 'AX'])}
TIPOS_EXTRA = {
    '14': ('servicio_publico', 'RECIBO DE SERVICIOS PÚBLICOS'),
    '50': ('dua', 'DECLARACIÓN ÚNICA DE ADUANAS (DUA)'),
    '52': ('despacho_simplificado', 'DESPACHO SIMPLIFICADO DE IMPORTACIÓN'),
    '53': ('courier', 'DECLARACIÓN DE MENSAJERÍA O COURIER'),
    '54': ('liquidacion_aduanas', 'LIQUIDACIÓN DE COBRANZA (ADUANAS)'),
}
TIPOS_IMPORTACION = ('50', '52', '53', '54')
_C = Decimal('0.01')


def _txt(v) -> str:
    if v is None:
        return ''
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return ' '.join(str(v).split())


def _num(v) -> float:
    if v is None or v == '':
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).replace(',', '').strip() or 0)
    except ValueError:
        return 0.0


def _fecha(v) -> str:
    """→ 'YYYY-MM-DD' o ''."""
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = _txt(v)
    for fmt in ('%d/%m/%Y', '%Y-%m-%d', '%d-%m-%Y'):
        try:
            return datetime.strptime(s[:10], fmt).date().isoformat()
        except ValueError:
            pass
    return ''


def _vacio(v) -> bool:
    return v is None or _txt(v) == ''


def leer_excel_compras(data: bytes) -> tuple[list[dict], list[str]]:
    """Devuelve (filas, avisos). Cada fila es un dict serializable en JSON con las columnas útiles."""
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    filas, avisos = [], []
    for ws in wb.worksheets:
        for n, r in enumerate(ws.iter_rows(values_only=True), start=1):
            r = list(r) + [None] * (50 - len(r))
            fe, tipo = _fecha(r[COL['A']]), _txt(r[COL['C']])
            if not fe or not re.fullmatch(r'\d{1,2}', tipo):
                continue
            serie, numero = _txt(r[COL['D']]), _txt(r[COL['F']])
            if not numero:
                avisos.append(f'{ws.title} fila {n}: sin número de comprobante; se omitió.')
                continue
            ac = r[COL['AC']]
            filas.append({
                'hoja': ws.title, 'fila': n, 'emision': fe, 'vencimiento': _fecha(r[COL['AD']]) or _fecha(r[COL['B']]),
                'tipo': tipo.zfill(2), 'serie': serie, 'anio_dua': _txt(r[COL['E']]), 'numero': numero,
                'tipo_doc_prov': _txt(r[COL['G']]), 'ruc': _txt(r[COL['H']]), 'nombre': _txt(r[COL['I']]),
                'J': _num(r[COL['J']]), 'K': _num(r[COL['K']]), 'L': _num(r[COL['L']]), 'M': _num(r[COL['M']]),
                'N': _num(r[COL['N']]), 'O': _num(r[COL['O']]), 'P': _num(r[COL['P']]), 'Q': _num(r[COL['Q']]),
                'R': _num(r[COL['R']]), 'S': _num(r[COL['S']]),
                'det_constancia': _txt(r[COL['U']]), 'det_fecha': _fecha(r[COL['V']]),
                'tc': _num(r[COL['W']]),
                'ref_fecha': _fecha(r[COL['X']]), 'ref_tipo': _txt(r[COL['Y']]), 'ref_serie': _txt(r[COL['Z']]), 'ref_numero': _txt(r[COL['AA']]),
                'usd': not _vacio(ac), 'AC': _num(ac),
            })
    return filas, avisos


def pid_fila(f: dict) -> str:
    return f'{f["ruc"]}|{f["tipo"]}|{f["serie"]}-{f["numero"]}'


def _d(x: float) -> Decimal:
    return Decimal(str(x)).quantize(_C, ROUND_HALF_UP)


def comprobante_desde_fila(f: dict, nombre_archivo: str = 'Excel') -> Comprobante:
    tipo = f['tipo']
    c = Comprobante(archivo=f'{nombre_archivo} · fila {f["fila"]}')
    c.tipo_codigo = tipo
    c.tipo_documento, c.tipo_nombre = TIPOS_DOC.get(tipo) or TIPOS_EXTRA.get(tipo) or ('otro', f'DOCUMENTO TIPO {tipo}')
    c.serie_numero = f'{f["serie"]}-{f["numero"]}'
    c.fecha_emision = f['emision']
    c.fecha_vencimiento = f['vencimiento'] if f['vencimiento'] and f['vencimiento'] > f['emision'] else ''
    c.periodo = f['emision'][:7]
    c.ruc_emisor = f['ruc'] if f['ruc'] not in ('0',) else ''
    c.nombre_emisor = f['nombre'] or (f'PROVEEDOR {c.ruc_emisor}' if c.ruc_emisor else 'PROVEEDOR DEL EXTERIOR')
    c.tipo_doc_cliente = '6'
    usd = bool(f['usd'])
    c.moneda_codigo, c.moneda, c.simbolo = ('USD', 'dolares', 'US$') if usd else ('PEN', 'soles', 'S/')

    # importes: el Excel los trae en soles; un comprobante en dólares se lleva a dólares con AC / S (o 1 / TC)
    S = f['S']
    if usd:
        k = (f['AC'] / S) if S else (1 / f['tc'] if f['tc'] else 1.0)
    else:
        k = 1.0
    c.gravadas = _d((f['J'] + f['L'] + f['N']) * k)
    c.inafectas = _d(f['P'] * k)
    c.isc = _d(f['Q'] * k)
    c.otros_cargos = _d(f['R'] * k)
    c.importe_total = _d(f['AC']) if usd and S else _d(S * k)
    c.igv = c.importe_total - c.gravadas - c.inafectas - c.isc - c.otros_cargos if (f['K'] or f['M'] or f['O']) else Decimal(0)
    c.valor_venta = c.gravadas + c.inafectas
    c.tiene_igv = c.igv > 0
    base = f['J'] + f['L'] + f['N']
    if base > 0 and (f['K'] or f['M'] or f['O']):
        tasa = (f['K'] + f['M'] + f['O']) / base * 100
        c.igv_tasa = Decimal('10.5') if abs(tasa - 10.5) < 0.6 else (Decimal('10') if abs(tasa - 10) < 0.3 else Decimal('18'))

    # forma de pago: el Excel no la trae; si vence después de la emisión, es crédito
    if c.fecha_vencimiento:
        c.forma_pago = 'Credito'
        c.cuotas = [{'numero': '1', 'monto': c.importe_total, 'vence': c.fecha_vencimiento}]

    # NC / ND: documento de referencia
    if f['ref_serie'] or f['ref_numero']:
        c.doc_referencia = f'{f["ref_serie"]}-{f["ref_numero"]}'

    # detracción (si el Excel trae la constancia)
    if f['det_constancia']:
        c.tiene_detraccion = True

    # datos propios del Excel (no son campos del XML)
    c.origen = 'excel'
    c.excel = f
    return c


def es_de_excel(c) -> bool:
    return getattr(c, 'origen', '') == 'excel'
