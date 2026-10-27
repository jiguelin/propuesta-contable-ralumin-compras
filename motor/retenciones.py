# -*- coding: utf-8 -*-
"""
motor/retenciones.py — Régimen de retenciones del IGV (Ralumin es agente de retención).

Reglas (orientación SUNAT "Aplicación del régimen y operaciones excluidas"):
  · Solo compras gravadas con IGV sustentadas con factura (01) o nota de débito (08).
  · Se retiene si la suma de los comprobantes del MISMO proveedor pagados en la MISMA fecha supera S/ 700.
  · Tasa 3 % del importe total.
  · Excluidas: exoneradas/inafectas, proveedor buen contribuyente / agente de retención / agente de percepción
    (padrones SUNAT, a la fecha de pago), operaciones con detracción, boletas/tickets sin crédito fiscal,
    recibos de luz/agua y boletos de aviación, liquidaciones de compra.

También lee el XML del Comprobante de Retención electrónico (tipo 20) que emite Ralumin y los padrones SUNAT.
"""
from __future__ import annotations

import io
import re
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime

from lxml import etree

_PARSER = etree.XMLParser(recover=True, huge_tree=True, remove_blank_text=True)


# ----------------------------------------------------------------------------- comprobante de retención (XML)
@dataclass
class ReferenciaRetencion:
    tipo: str            # 01 factura, 08 ND…
    serie_numero: str
    fecha: str
    total: float
    pagado: float
    retencion: float
    fecha_pago: str = ''
    moneda: str = 'PEN'


@dataclass
class ComprobanteRetencion:
    serie_numero: str
    fecha: str
    ruc_proveedor: str
    nombre_proveedor: str
    tasa: float
    total_retenido: float
    referencias: list[ReferenciaRetencion] = field(default_factory=list)
    archivo: str = ''

    @property
    def serie(self) -> str:
        return self.serie_numero.split('-')[0] if '-' in self.serie_numero else self.serie_numero

    @property
    def numero(self) -> str:
        return self.serie_numero.split('-', 1)[1] if '-' in self.serie_numero else ''


def _ln(el) -> str:
    return etree.QName(el).localname


def _hijo(el, nombre: str):
    for h in el.iter():
        if h is not el and _ln(h) == nombre:
            return h
    return None


def _hijo_directo(el, nombre: str):
    for h in el:
        if isinstance(h.tag, str) and _ln(h) == nombre:
            return h
    return None


def _txt(el, *ruta) -> str:
    """Busca por nombres locales en cadena (sin depender de namespaces)."""
    cur = el
    for nombre in ruta:
        if cur is None:
            return ''
        cur = _hijo(cur, nombre)
    return ' '.join((cur.text or '').split()) if cur is not None else ''


def _f(v) -> float:
    try:
        return float(v)
    except (TypeError, ValueError):
        return 0.0


def leer_xml_retencion(data: bytes, nombre: str = '') -> ComprobanteRetencion | None:
    try:
        root = etree.fromstring(data, _PARSER)
    except Exception:
        return None
    if root is None or _ln(root) != 'Retention':
        return None
    receptor = _hijo(root, 'ReceiverParty')
    cr = ComprobanteRetencion(
        serie_numero=_txt(root, 'ID') if _hijo_directo(root, 'ID') is None else ' '.join((_hijo_directo(root, 'ID').text or '').split()),
        fecha=' '.join((_hijo_directo(root, 'IssueDate').text or '').split()) if _hijo_directo(root, 'IssueDate') is not None else '',
        ruc_proveedor=_txt(receptor, 'PartyIdentification', 'ID') if receptor is not None else '',
        nombre_proveedor=_txt(receptor, 'RegistrationName') if receptor is not None else '',
        tasa=_f(_txt(root, 'SUNATRetentionPercent')),
        total_retenido=_f(' '.join((_hijo_directo(root, 'TotalAmount').text or '').split())) if _hijo_directo(root, 'TotalAmount') is not None else 0.0,
        archivo=nombre,
    )
    for ref in [e for e in root.iter() if isinstance(e.tag, str) and _ln(e) == 'SUNATRetentionDocumentReference']:
        idn = _hijo_directo(ref, 'ID')
        r = ReferenciaRetencion(
            tipo=(idn.get('schemeID') if idn is not None else '') or '01',
            serie_numero=' '.join((idn.text or '').split()) if idn is not None else '',
            fecha=_txt(ref, 'IssueDate'),
            total=_f(_txt(ref, 'TotalInvoiceAmount')),
            pagado=_f(_txt(ref, 'Payment', 'PaidAmount')),
            retencion=_f(_txt(ref, 'SUNATRetentionAmount')),
            fecha_pago=_txt(ref, 'Payment', 'PaidDate') or _txt(ref, 'SUNATRetentionDate'),
        )
        tot = _hijo(ref, 'TotalInvoiceAmount')
        if tot is not None and tot.get('currencyID'):
            r.moneda = tot.get('currencyID')
        cr.referencias.append(r)
    if not cr.total_retenido:
        cr.total_retenido = round(sum(r.retencion for r in cr.referencias), 2)
    return cr


def leer_retenciones(archivos: list[tuple[str, bytes]]) -> tuple[list[ComprobanteRetencion], list[str]]:
    """XML sueltos o ZIP. Devuelve (comprobantes, nombres que no se pudieron leer)."""
    out, malos = [], []
    for nombre, data in archivos:
        if nombre.lower().endswith('.zip'):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for n in z.namelist():
                    if n.lower().endswith('.xml') and not n.startswith('__MACOSX'):
                        cr = leer_xml_retencion(z.read(n), n.split('/')[-1])
                        (out.append(cr) if cr else malos.append(n))
        elif nombre.lower().endswith('.xml'):
            cr = leer_xml_retencion(data, nombre)
            (out.append(cr) if cr else malos.append(nombre))
    return out, malos


# ----------------------------------------------------------------------------- padrones SUNAT
def leer_padron(archivos: list[tuple[str, bytes]]) -> set[str]:
    """Padrón de buenos contribuyentes / agentes de retención / agentes de percepción (TXT con '|', CSV, Excel o ZIP).
    Se toma todo RUC de 11 dígitos al inicio de cada fila."""
    rucs: set[str] = set()

    def de_texto(t: str):
        for linea in t.splitlines():
            m = re.match(r'\s*"?(\d{11})', linea)
            if m:
                rucs.add(m.group(1))

    for nombre, data in archivos:
        low = nombre.lower()
        if low.endswith('.zip'):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for n in z.namelist():
                    if not n.endswith('/'):
                        rucs |= leer_padron([(n, z.read(n))])
        elif low.endswith(('.xlsx', '.xlsm')):
            import openpyxl
            wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
            for ws in wb.worksheets:
                for r in ws.iter_rows(values_only=True):
                    for v in r[:3]:
                        s = str(v or '').strip()
                        if re.fullmatch(r'\d{11}', s):
                            rucs.add(s)
                            break
        else:
            for enc in ('utf-8', 'latin-1'):
                try:
                    de_texto(data.decode(enc))
                    break
                except UnicodeDecodeError:
                    continue
    return rucs


# ----------------------------------------------------------------------------- reglas
TIPOS_CON_RETENCION = ('01', '08')


def motivo_exclusion(c, padron_excluidos: set[str] | None, leyendas_no_retener: tuple[str, ...]) -> str:
    """'' si la operación PUEDE estar sujeta a retención; si no, el motivo de exclusión."""
    if c.tipo_codigo not in TIPOS_CON_RETENCION:
        if c.tipo_codigo == '03':
            return 'boleta sin derecho a crédito fiscal'
        if c.tipo_codigo == '14':
            return 'recibo de servicios públicos'
        if c.tipo_codigo == '04':
            return 'liquidación de compra'
        return f'tipo de comprobante {c.tipo_codigo}'
    if not c.igv:
        return 'operación exonerada o inafecta'
    if c.tiene_detraccion:
        return 'operación sujeta a detracción'
    if c.tiene_percepcion:
        return 'el proveedor cobra percepción (agente de percepción)'
    texto = ' '.join(c.leyendas + c.observaciones).upper()
    for l in leyendas_no_retener:
        if l in texto:
            return f'el comprobante dice "{l.title()}"'
    if padron_excluidos and c.ruc_emisor in padron_excluidos:
        return 'proveedor en padrón SUNAT (buen contribuyente o agente)'
    return ''


def importe_retencion(total_soles: float, tasa: float) -> float:
    return round(total_soles * tasa / 100 + 1e-9, 2)
