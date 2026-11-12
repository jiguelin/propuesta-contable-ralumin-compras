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

También lee los comprobantes de retención que emite Ralumin, de dos maneras:
  · el XML del Comprobante de Retención electrónico (tipo 20), suelto o en ZIP;
  · el Excel de SUNAT "Consulta de Comprobantes Emitidos/Recibidos" (SOL → comprobantes de retención), .xls o .xlsx,
    cuando no se tiene el XML: un bloque por comprobante de retención (Emisor, Receptor, N° de comprobante, fecha,
    importe retenido, tasa, estado) y debajo la tabla de facturas (tipo, serie, número, fecha, total, N° de pago,
    importe pagado, retención S/, monto neto S/).
Y los padrones SUNAT.
"""
from __future__ import annotations

import io
import re
import unicodedata
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
    neto: float = 0.0          # monto neto pagado en soles (Excel de SUNAT)
    tc_pago: float = 0.0       # en dólares: TC del día de pago que usó el comprobante = (neto + retención) / pagado


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
    estado: str = ''           # EMITIDO / ANULADO / REVERTIDO (Excel de SUNAT); '' en el XML
    ruc_agente: str = ''       # quien emite (Ralumin)

    @property
    def anulado(self) -> bool:
        e = (self.estado or '').upper()
        return any(p in e for p in ('ANULAD', 'REVERTID', 'BAJA', 'RECHAZ'))

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


EXT_EXCEL = ('.xls', '.xlsx', '.xlsm')


def leer_retenciones(archivos: list[tuple[str, bytes]]) -> tuple[list[ComprobanteRetencion], list[str]]:
    """XML sueltos, ZIP o el Excel de SUNAT. Devuelve (comprobantes, nombres que no se pudieron leer).
    Un mismo comprobante de retención subido dos veces (o por XML y por Excel) cuenta una sola vez: manda el primero."""
    out, malos = [], []
    for nombre, data in archivos:
        low = nombre.lower()
        if low.endswith('.zip'):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for n in z.namelist():
                    if n.startswith('__MACOSX') or n.endswith('/'):
                        continue
                    if n.lower().endswith('.xml'):
                        cr = leer_xml_retencion(z.read(n), n.split('/')[-1])
                        (out.append(cr) if cr else malos.append(n))
                    elif n.lower().endswith(EXT_EXCEL):
                        crs, _ = leer_excel_retenciones(z.read(n), n.split('/')[-1])
                        (out.extend(crs) if crs else malos.append(n))
        elif low.endswith('.xml'):
            cr = leer_xml_retencion(data, nombre)
            (out.append(cr) if cr else malos.append(nombre))
        elif low.endswith(EXT_EXCEL):
            crs, _ = leer_excel_retenciones(data, nombre)
            (out.extend(crs) if crs else malos.append(nombre))
    unicos, vistos = [], set()
    for cr in out:
        k = (cr.ruc_proveedor, re.sub(r'\s', '', cr.serie_numero).upper())
        if k not in vistos:
            vistos.add(k)
            unicos.append(cr)
    return unicos, malos


# ----------------------------------------------------------------------------- comprobante de retención (Excel de SUNAT)
def _clave_texto(v) -> str:
    """Etiqueta comparable: sin tildes, sin el mojibake del reporte ('Nï¿½mero'), solo A–Z y 0–9."""
    s = unicodedata.normalize('NFKD', str(v or '')).encode('ascii', 'ignore').decode().upper()
    return re.sub(r'[^A-Z0-9]', '', s)


def _celdas_excel(data: bytes) -> list[list]:
    """Filas de la primera hoja con datos: .xls (xlrd), .xlsx (openpyxl) o el HTML que algunos portales guardan como .xls."""
    cab = data[:8]
    if cab.startswith(b'\xd0\xcf\x11\xe0'):                       # Excel 97-2003 (BIFF)
        import xlrd
        wb = xlrd.open_workbook(file_contents=data)
        hojas = []
        for sh in wb.sheets():
            filas = []
            for r in range(sh.nrows):
                fila = []
                for c in range(sh.ncols):
                    cel = sh.cell(r, c)
                    if cel.ctype == xlrd.XL_CELL_DATE:
                        try:
                            fila.append(xlrd.xldate_as_datetime(cel.value, wb.datemode))
                            continue
                        except Exception:
                            pass
                    fila.append(cel.value)
                filas.append(fila)
            hojas.append(filas)
    elif cab.startswith(b'PK'):                                     # .xlsx
        import openpyxl
        wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
        hojas = [[list(r) for r in ws.iter_rows(values_only=True)] for ws in wb.worksheets]
    else:                                                           # HTML con extensión .xls
        root = etree.fromstring(data, etree.HTMLParser(recover=True))
        hojas = [[[' '.join(''.join(td.itertext()).split()) for td in tr if isinstance(td.tag, str) and td.tag.lower() in ('td', 'th')]
                   for tr in root.iter('tr')]] if root is not None else []
    for filas in hojas:
        if any(any(str(v or '').strip() for v in f) for f in filas):
            return filas
    return []


def _monto(v) -> tuple[float, str]:
    """'$ 12551.85' → (12551.85, 'USD'); 'S/ 800.00' → (800.0, 'PEN'); 1.294,24 o 1,294.24 → número."""
    if isinstance(v, (int, float)):
        return float(v), ''
    s = str(v or '').strip()
    if not s:
        return 0.0, ''
    up = s.upper()
    moneda = 'USD' if ('$' in s and 'S/' not in up) or 'US' in up else ('PEN' if 'S/' in up or 'PEN' in up else '')
    num = re.sub(r'[^0-9,.\-]', '', s)
    if ',' in num and '.' in num:
        num = num.replace(',', '') if num.rfind('.') > num.rfind(',') else num.replace('.', '').replace(',', '.')
    elif ',' in num:
        num = num.replace(',', '.') if len(num.split(',')[-1]) == 2 else num.replace(',', '')
    try:
        return float(num), moneda
    except ValueError:
        return 0.0, moneda


def _fecha(v) -> str:
    if isinstance(v, datetime):
        return v.date().isoformat()
    if isinstance(v, date):
        return v.isoformat()
    s = str(v or '').strip()
    m = re.search(r'(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{4})', s)
    if m:
        try:
            return date(int(m.group(3)), int(m.group(2)), int(m.group(1))).isoformat()
        except ValueError:
            return ''
    m = re.search(r'(\d{4})-(\d{2})-(\d{2})', s)
    return m.group(0) if m else ''


TIPO_POR_NOMBRE = (('NOTADEDEBITO', '08'), ('NOTADECREDITO', '07'), ('FACTURA', '01'), ('BOLETA', '03'), ('RECIBO', '14'),
                   ('LIQUIDACION', '04'), ('TICKET', '12'))


def _tipo(v) -> str:
    s = _clave_texto(v)
    if re.fullmatch(r'\d{1,2}', s):
        return s.zfill(2)
    for clave, cod in TIPO_POR_NOMBRE:
        if s.startswith(clave):
            return cod
    return '01'


def _columnas_detalle(fila: list) -> dict[str, int]:
    col = {}
    for j, v in enumerate(fila):
        k = _clave_texto(v)
        if not k:
            continue
        if k.startswith('TIPODE') or k == 'TIPO':
            col.setdefault('tipo', j)
        elif k.startswith('SERIE'):
            col.setdefault('serie', j)
        elif 'MERO' in k and 'PAGO' not in k:                      # Número / Nï¿½mero
            col.setdefault('numero', j)
        elif k.startswith('FECHA'):
            col.setdefault('fecha', j)
        elif k.startswith('TOTAL'):
            col.setdefault('total', j)
        elif 'PAGO' in k and not k.startswith('IMPORTE') and not k.startswith('MONTO'):
            col.setdefault('nro_pago', j)
        elif k.startswith('IMPORTEPAGADO'):
            col.setdefault('pagado', j)
        elif k.startswith('RETENC'):
            col.setdefault('retencion', j)
        elif k.startswith('MONTONETO') or k.startswith('NETO'):
            col.setdefault('neto', j)
    return col


def _valor_a_la_derecha(fila: list, j: int):
    for v in fila[j + 1:]:
        if str(v if v is not None else '').strip():
            return v
    return ''


def leer_excel_retenciones(data: bytes, nombre: str = '') -> tuple[list[ComprobanteRetencion], list[str]]:
    """Excel de SUNAT "Consulta de Comprobantes Emitidos" de retenciones. Devuelve (comprobantes, avisos).
    Cada referencia lleva la fecha de pago = fecha del comprobante de retención (se emite al pagar)."""
    try:
        filas = _celdas_excel(data)
    except Exception as ex:                                          # archivo dañado o de otro tipo
        return [], [f'No se pudo abrir: {ex}']
    out: list[ComprobanteRetencion] = []
    avisos: list[str] = []
    cur: ComprobanteRetencion | None = None
    col: dict[str, int] = {}

    def nuevo() -> ComprobanteRetencion:
        return ComprobanteRetencion(serie_numero='', fecha='', ruc_proveedor='', nombre_proveedor='', tasa=0.0, total_retenido=0.0, archivo=nombre)

    def cerrar():
        nonlocal cur, col
        if cur and cur.serie_numero and cur.referencias:
            for r in cur.referencias:
                r.fecha_pago = r.fecha_pago or cur.fecha
            if not cur.total_retenido:
                cur.total_retenido = round(sum(r.retencion for r in cur.referencias), 2)
            out.append(cur)
        elif cur and cur.serie_numero:
            avisos.append(f'{cur.serie_numero}: sin facturas en el detalle.')
        cur, col = None, {}

    for fila in filas:
        fila = list(fila)
        claves = [_clave_texto(v) for v in fila]
        # ---- tabla de detalle ----
        if any(k.startswith('TIPODEDOCUMENTO') for k in claves) and any(k.startswith('SERIE') for k in claves):
            col = _columnas_detalle(fila)
            continue
        es_cabecera = any(k.startswith(('EMISOR', 'RECEPTOR', 'IMPORTETOTALRETENIDO', 'NODECOMPROBANTE')) for k in claves)
        if cur is not None and col and not es_cabecera and {'serie', 'numero'} <= col.keys():
            serie = str(fila[col['serie']] if col['serie'] < len(fila) else '').strip()
            numero = fila[col['numero']] if col['numero'] < len(fila) else ''
            numero = str(int(numero)) if isinstance(numero, float) and numero.is_integer() else str(numero or '').strip()
            if serie and numero:
                g = lambda k: fila[col[k]] if k in col and col[k] < len(fila) else ''
                total, mon_t = _monto(g('total'))
                pagado, mon_p = _monto(g('pagado'))
                ret, _ = _monto(g('retencion'))
                neto, _ = _monto(g('neto'))
                moneda = mon_t or mon_p or 'PEN'
                r = ReferenciaRetencion(tipo=_tipo(g('tipo')), serie_numero=f'{serie}-{numero}', fecha=_fecha(g('fecha')), total=total,
                                        pagado=pagado or total, retencion=round(ret, 2), moneda=moneda, neto=round(neto, 2))
                if moneda == 'USD' and r.pagado and neto:
                    r.tc_pago = round((neto + r.retencion) / r.pagado, 3)
                cur.referencias.append(r)
                continue
        # ---- cabecera del comprobante ----
        for j, k in enumerate(claves):
            if not k:
                continue
            if k.startswith('EMISOR'):
                cerrar()
                cur = nuevo()
                m = re.search(r'\d{11}', str(_valor_a_la_derecha(fila, j)))
                cur.ruc_agente = m.group(0) if m else ''
            elif k.startswith('NODECOMPROBANTE') or k.startswith('NUMERODECOMPROBANTE') or k.startswith('NDECOMPROBANTE'):
                if cur is None or cur.serie_numero:
                    if cur is not None and cur.serie_numero:
                        cerrar()
                    cur = cur or nuevo()
                cur.serie_numero = re.sub(r'\s', '', str(_valor_a_la_derecha(fila, j))).upper()
            elif cur is None:
                continue
            elif k.startswith('RECEPTOR') or k.startswith('PROVEEDOR'):
                v = str(_valor_a_la_derecha(fila, j))
                m = re.search(r'\d{11}', v)
                cur.ruc_proveedor = m.group(0) if m else ''
                cur.nombre_proveedor = ' '.join(v[m.end():].lstrip(' -–:').split()) if m else v.strip()
            elif k.startswith('FECHADEEMISI') or k == 'FECHA':
                cur.fecha = _fecha(_valor_a_la_derecha(fila, j))
            elif k.startswith('IMPORTETOTALRETENIDO') or k.startswith('TOTALRETENIDO'):
                cur.total_retenido = round(_monto(_valor_a_la_derecha(fila, j))[0], 2)
            elif k.startswith('TASA'):
                cur.tasa = _monto(_valor_a_la_derecha(fila, j))[0]
            elif k.startswith('ESTADO'):
                cur.estado = str(_valor_a_la_derecha(fila, j)).strip().upper()
    cerrar()
    for cr in out:
        suma = round(sum(r.retencion for r in cr.referencias), 2)
        if cr.total_retenido and abs(suma - cr.total_retenido) > 0.011:
            avisos.append(f'{cr.serie_numero}: el importe retenido (S/ {cr.total_retenido:,.2f}) no cuadra con la suma del detalle (S/ {suma:,.2f}).')
        if not cr.ruc_proveedor:
            avisos.append(f'{cr.serie_numero}: no se encontró el RUC del proveedor (Receptor).')
    return out, avisos


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
