# -*- coding: utf-8 -*-
"""
motor/diario.py — Lector del "Libro Diario Detallado" (Diario Analítico) de Contasis, subdiario de compras.

Formato (export de NewContaSis):
    cabecera: MES | FECHA | CUENTA | DESCRIPCION | DEBITO S/. | CREDITO S/. | M. | T/C | DEBITO US$. | CREDITO US$. |
              GLOSA | R.U.C. | R. | DOC. | NÚMERO | FEC DOC. | FEC VENC. | C. COSTO | PRESUP. | GLOSA 2
    cada asiento empieza con una fila "SUB DIARIO 01 ASIENTO No N" y termina con "TOTAL ASIENTO:".

Sirve para tres cosas:
  1. Memoria por proveedor: qué cuenta de gasto y qué cuenta de pago se usaron (lo reciente pesa más).
  2. "Ya contabilizado": RUC + tipo + serie + número de todo lo que ya está en Contasis.
  3. Confirmar bloques: después de importar un bloque, se sube el Diario y se marca lo que entró.
"""
from __future__ import annotations

import io
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime

import openpyxl

from . import pcge

RAICES_PAGO = ('101', '104')


@dataclass
class Linea:
    cuenta: str
    descripcion: str
    debe: float
    haber: float
    moneda: str = ''
    tc: float = 0.0
    glosa: str = ''
    ruc: str = ''
    doc: str = ''
    numero: str = ''
    fecha_doc: date | None = None


@dataclass
class Asiento:
    mes: str
    numero: int
    fecha: date | None
    lineas: list[Linea] = field(default_factory=list)

    # ---------- datos del comprobante ----------
    def _primera(self, attr):
        return next((getattr(l, attr) for l in self.lineas if getattr(l, attr)), '')

    @property
    def ruc(self) -> str:
        return self._primera('ruc')

    @property
    def doc(self) -> str:
        return self._primera('doc')

    @property
    def numero_doc(self) -> str:
        return self._primera('numero')

    @property
    def fecha_doc(self):
        return self._primera('fecha_doc') or self.fecha

    @property
    def glosa(self) -> str:
        return self._primera('glosa')

    @property
    def moneda(self) -> str:
        return self._primera('moneda')

    @property
    def clave(self) -> tuple | None:
        return clave_comprobante(self.ruc, self.doc, self.numero_doc)

    # ---------- interpretación contable ----------
    @property
    def lineas_gasto(self) -> list[Linea]:
        """Cargos de gasto/costo/activo (sin IGV, destinos 9x/20/25/61/79 ni la comisión bancaria)."""
        out = []
        for l in self.lineas:
            c = l.cuenta
            if l.debe <= 0:
                continue
            if c.startswith(('40', '42', '10', '9', '79', '61', '20', '21', '25', '26', '28')):
                continue
            if c.startswith('6391') and any(x.cuenta[0] == '6' and not x.cuenta.startswith(('61', '6391')) and x.debe > 0 for x in self.lineas):
                continue
            out.append(l)
        return out

    @property
    def cuenta_gasto(self) -> str:
        g = self.lineas_gasto
        return max(g, key=lambda l: l.debe).cuenta if g else ''

    @property
    def repartido(self) -> bool:
        """True si el gasto se partió entre la divisionaria ADM (094) y la VTAS (095)."""
        cs = [l.cuenta for l in self.lineas_gasto]
        return any(c.endswith(('094', '09401')) for c in cs) and any(c.endswith(('095', '09501')) for c in cs)

    @property
    def cuenta_pago(self) -> str:
        """Cuenta 10 del Haber; si no hay, 4212 (crédito o pago aparte)."""
        pago = [l for l in self.lineas if l.haber > 0 and l.cuenta.startswith(RAICES_PAGO)]
        if pago:
            return max(pago, key=lambda l: l.haber).cuenta
        if any(l.cuenta.startswith('4212') and l.haber > 0 for l in self.lineas):
            return '4212'
        return ''

    @property
    def es_compra(self) -> bool:
        return bool(self.cuenta_gasto) and any(l.cuenta.startswith('4212') and l.haber > 0 for l in self.lineas) or \
            (bool(self.cuenta_gasto) and any(l.cuenta.startswith(RAICES_PAGO) and l.haber > 0 for l in self.lineas))

    @property
    def retencion(self) -> float:
        return round(sum(l.haber for l in self.lineas if l.cuenta.startswith('40114')), 2)

    @property
    def comision(self) -> float:
        return round(sum(l.debe for l in self.lineas if l.cuenta.startswith('6391')), 2) if self.lineas_gasto and \
            not self.cuenta_gasto.startswith('6391') else 0.0


def _s(v) -> str:
    return ' '.join(str(v).split()) if v is not None else ''


def _n(v) -> float:
    try:
        return float(v or 0)
    except (TypeError, ValueError):
        try:
            return float(str(v).replace(',', ''))
        except ValueError:
            return 0.0


def _fecha(v):
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    s = _s(v)
    for fmt in ('%d/%m/%Y', '%Y-%m-%d'):
        try:
            return datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    return None


def normalizar_serie_numero(texto: str) -> tuple[str, str]:
    """'00F008-0000000973081' → ('F008', '973081'); 'F008-973081' → igual."""
    t = (texto or '').strip().upper()
    if '-' in t:
        s, n = t.split('-', 1)
    else:
        s, n = t[:-8] if len(t) > 8 else '', t[-8:]
    s = s.strip().lstrip('0') or '0'
    n = re.sub(r'\D', '', n).lstrip('0') or '0'
    return s, n


def clave_comprobante(ruc: str, tipo: str, serie_numero: str) -> tuple | None:
    ruc = (ruc or '').strip()
    if not ruc or not serie_numero:
        return None
    s, n = normalizar_serie_numero(serie_numero)
    return (ruc, (tipo or '').strip().zfill(2), s, n)


def leer_diario(origen) -> list[Asiento]:
    """origen: ruta, bytes o file-like del Excel del Diario."""
    if isinstance(origen, (bytes, bytearray)):
        origen = io.BytesIO(origen)
    wb = openpyxl.load_workbook(origen, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    filas = ws.iter_rows(values_only=True)

    col = {}
    asientos: list[Asiento] = []
    actual: Asiento | None = None
    for r in filas:
        vals = [_s(v) for v in r]
        if not col:
            up = [v.upper() for v in vals]
            if 'CUENTA' in up and any(v.startswith('DEBITO') or v.startswith('DEBE') for v in up):
                def idx(*nombres):
                    for nombre in nombres:
                        for j, v in enumerate(up):
                            if v == nombre or v.startswith(nombre):
                                return j
                    return None
                col = dict(mes=idx('MES'), fecha=idx('FECHA'), cuenta=idx('CUENTA'), desc=idx('DESCRIPCION', 'DESCRIPCIÓN'),
                           debe=idx('DEBITO S', 'DEBE'), haber=idx('CREDITO S', 'HABER'), m=idx('M.'), tc=idx('T/C'),
                           glosa=idx('GLOSA'), ruc=idx('R.U.C', 'RUC'), doc=idx('DOC.'), numero=idx('NÚMERO', 'NUMERO'),
                           fdoc=idx('FEC DOC'))
            continue
        texto = ' '.join(v for v in vals if v)
        m = re.search(r'(?i)sub\s*diario\s*\d+\s*asiento\s*no\.?:?\s*(\d+)', texto)
        if m:
            actual = Asiento(mes='', numero=int(m.group(1)), fecha=None)
            asientos.append(actual)
            continue
        if actual is None or texto.upper().startswith('TOTAL ASIENTO'):
            continue

        def g(k):
            j = col.get(k)
            return r[j] if j is not None and j < len(r) else None
        cuenta = _s(g('cuenta'))
        if not re.fullmatch(r'\d{2,12}', cuenta):
            continue
        l = Linea(cuenta=cuenta, descripcion=_s(g('desc')), debe=_n(g('debe')), haber=_n(g('haber')), moneda=_s(g('m')),
                  tc=_n(g('tc')), glosa=_s(g('glosa')), ruc=_s(g('ruc')), doc=_s(g('doc')), numero=_s(g('numero')),
                  fecha_doc=_fecha(g('fdoc')))
        if not actual.mes:
            actual.mes = _s(g('mes')).zfill(2)
            actual.fecha = _fecha(g('fecha'))
        actual.lineas.append(l)
    return [a for a in asientos if a.lineas]


# ----------------------------------------------------------------------------- memoria
@dataclass
class MemoriaDiario:
    """Lo que el Diario enseña de cada proveedor."""
    gasto: dict = field(default_factory=lambda: defaultdict(Counter))      # ruc → Counter(cuenta → peso, lo reciente pesa más)
    gasto_n: dict = field(default_factory=lambda: defaultdict(Counter))    # ruc → Counter(cuenta → veces)
    pago: dict = field(default_factory=lambda: defaultdict(lambda: defaultdict(Counter)))  # ruc → mes → Counter(cuenta pago)
    reparto: dict = field(default_factory=lambda: defaultdict(Counter))    # ruc → Counter(True/False)
    ultimas: dict = field(default_factory=lambda: defaultdict(list))       # ruc → [(fecha, doc, cuenta gasto, cuenta pago, total)]
    frecuencia: Counter = field(default_factory=Counter)                   # cuenta → veces
    contabilizados: set = field(default_factory=set)                       # claves (ruc, tipo, serie, número)
    asientos_por_clave: dict = field(default_factory=dict)                 # clave → (mes, número de asiento)
    n_asientos: int = 0
    meses: list = field(default_factory=list)

    def memoria_proveedor(self, ruc: str) -> list[dict]:
        """Formato que espera motor.clasificador: [{concepto, cuenta, veces}] ordenado por veces."""
        out = []
        veces = self.gasto_n.get(ruc, Counter())
        for cuenta, _peso in self.gasto.get(ruc, Counter()).most_common():
            fam = pcge.familia_por_prefijo(cuenta)
            out.append({'concepto': fam['prefijo'] if fam else cuenta[:3], 'cuenta': cuenta, 'veces': veces.get(cuenta, 1)})
        return out

    def pago_reciente(self, ruc: str) -> str:
        """Cuenta de pago más usada con este proveedor en el último mes en que aparece en el Diario."""
        por_mes = self.pago.get(ruc)
        if not por_mes:
            return ''
        ultimo = max(por_mes)
        return por_mes[ultimo].most_common(1)[0][0]

    def suele_repartir(self, ruc: str) -> bool | None:
        c = self.reparto.get(ruc)
        if not c:
            return None
        return c[True] >= c[False]


def construir_memoria(asientos: list[Asiento], meses_recientes: int = 3) -> MemoriaDiario:
    mem = MemoriaDiario()
    meses = sorted({a.mes for a in asientos if a.mes})
    mem.meses = meses
    recientes = set(meses[-meses_recientes:])
    for a in asientos:
        mem.n_asientos += 1
        for l in a.lineas:
            mem.frecuencia[l.cuenta] += 1
        k = a.clave
        if k:
            mem.contabilizados.add(k)
            mem.asientos_por_clave[k] = (a.mes, a.numero)
        ruc = a.ruc
        cg = a.cuenta_gasto
        if not ruc or not cg or a.glosa.upper().startswith(('PAGO', 'CANCELACION')):
            continue
        # lo reciente pesa más: el último mes vale 1, el anterior 0.6, el otro 0.36… (Ralumin cambia de criterio con el tiempo)
        peso = 0.6 ** (len(meses) - 1 - meses.index(a.mes)) if a.mes in meses else 0.1
        mem.gasto[ruc][cg] += peso
        mem.gasto_n[ruc][cg] += 1
        mem.reparto[ruc][a.repartido] += 1
        if a.cuenta_pago:
            mem.pago[ruc][a.mes][a.cuenta_pago] += 1
        total = round(sum(l.haber for l in a.lineas if l.cuenta.startswith('4212')) or
                      sum(l.haber for l in a.lineas if l.cuenta.startswith(RAICES_PAGO)), 2)
        mem.ultimas[ruc].append((a.fecha_doc, f'{a.doc} {a.numero_doc}', cg, a.cuenta_pago, total))
    for ruc in mem.ultimas:
        mem.ultimas[ruc] = sorted(mem.ultimas[ruc], key=lambda x: x[0] or date.min, reverse=True)[:5]
    return mem
