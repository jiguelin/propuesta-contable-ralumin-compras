# -*- coding: utf-8 -*-
"""
motor/sesion.py — El trabajo de un mes: comprobantes, decisiones de la persona y bloques.

Ciclo de un comprobante:
    pendiente ──(crear bloque)──▶ exportado ──(confirmar / Diario)──▶ contabilizado
        ▲                              │
        └────────(devolver)────────────┘
    excluido: fuera del subdiario 01 (bancos, otro mes, detracción sin constancia…) o descartado con motivo.

Lo contabilizado y lo exportado desaparecen de la propuesta; así se trabaja por bloques y un error
queda limitado a su bloque. Todo se guarda como JSON (XML incluidos, comprimidos) en el almacén.
"""
from __future__ import annotations

import base64
import zlib
from collections import defaultdict
from dataclasses import asdict, dataclass, field, fields
from datetime import datetime

from ralumin import config as cfg

from . import pcge
from . import reglas as kb
from .clasificador import Contexto, Propuesta, procesar_lote
from .detracciones import cruzar
from .diario import MemoriaDiario, clave_comprobante, normalizar_serie_numero
from .glosas import glosa_ralumin
from .grupos import asignar_grupo, cuenta_pago_defecto, es_credito, es_cuenta_me
from .retenciones import importe_retencion, motivo_exclusion
from .xml_parser import parse_xml_bytes, texto_items

ESTADOS = ('pendiente', 'exportado', 'contabilizado', 'excluido')


# ----------------------------------------------------------------------------- decisiones (lo que se guarda)
@dataclass
class Decision:
    grupo: str = ''
    grupo_manual: bool = False
    cuenta_gasto: str = ''
    gasto_manual: bool = False
    reparto: bool = False
    reparto_manual: bool = False
    cuenta_pago: str = ''
    pago_manual: bool = False
    condicion: str = ''               # CON | CRE
    condicion_manual: bool = False
    pago_agrupado: bool = False
    fecha_pago: str = ''              # ISO; por defecto la de emisión
    retencion: bool = False
    retencion_manual: bool = False    # la persona marcó o desmarcó la casilla
    retencion_monto: float = 0.0
    retencion_serie: str = ''
    retencion_numero: str = ''
    retencion_fecha: str = ''
    motivo_no_retencion: str = ''
    comision: float = 0.0
    glosa: str = ''
    glosa_manual: bool = False
    revisado: bool = False
    estado: str = 'pendiente'
    motivo: str = ''                  # motivo de exclusión o nota de estado
    bloque: int = 0
    asiento: str = ''                 # número de asiento en Contasis (si se conoce)
    nota: str = ''

    @classmethod
    def desde_dict(cls, d: dict) -> 'Decision':
        nombres = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in (d or {}).items() if k in nombres})


@dataclass
class Bloque:
    numero: int
    nombre: str
    ids: list[str]
    estado: str = 'exportado'          # exportado | contabilizado | anulado
    creado: str = ''
    confirmado: str = ''
    asiento_inicial: int = 0
    nota: str = ''

    @property
    def codigo(self) -> str:
        return f'B{self.numero:02d}'


# ----------------------------------------------------------------------------- maestros (no se guardan en la sesión)
@dataclass
class Maestros:
    plan: object = None                         # PlanCuentas
    tc: object = None                           # HistorialTC
    memoria: MemoriaDiario | None = None
    entidades: dict = field(default_factory=dict)     # ruc → nombre
    padron: set = field(default_factory=set)          # RUC exceptuados de retención
    retenciones: list = field(default_factory=list)   # ComprobanteRetencion
    constancias: list = field(default_factory=list)   # Constancia (detracciones)
    ia: object = None


# ----------------------------------------------------------------------------- fila = propuesta del motor + decisión
@dataclass
class Fila:
    p: Propuesta
    d: Decision
    avisos: list[str] = field(default_factory=list)
    nivel: str = 'ok'                  # ok | amarillo | rojo
    retencion_sugerida: bool = False
    retencion_calculada: float = 0.0
    retencion_exclusion: str = ''
    ya_en_diario: bool = False
    proveedor_nuevo: bool = False

    @property
    def id(self) -> str:
        return self.p.id

    @property
    def c(self):
        return self.p.c

    @property
    def tc(self) -> float:
        return float(self.p.tc or 0)

    @property
    def total(self) -> float:
        return float(self.c.importe_total)

    @property
    def total_soles(self) -> float:
        return round(self.total * (self.tc if self.c.moneda_codigo == 'USD' and self.tc else 1.0), 2)

    @property
    def tiene_ajuste(self) -> bool:
        return bool(self.d.retencion) or bool(self.d.comision)

    @property
    def semaforo(self) -> str:
        return {'ok': '🟢', 'amarillo': '🟡', 'rojo': '🔴'}[self.nivel]


def _b64(data: bytes) -> str:
    return base64.b64encode(zlib.compress(data, 9)).decode()


def _unb64(s: str) -> bytes:
    return zlib.decompress(base64.b64decode(s))


# ----------------------------------------------------------------------------- sesión
class Sesion:
    def __init__(self, periodo: str = ''):
        self.periodo = periodo                     # 'YYYY-MM'
        self.xml: dict[str, tuple[str, bytes]] = {}     # id → (nombre de archivo, bytes)
        self.decisiones: dict[str, Decision] = {}
        self.bloques: list[Bloque] = []
        self.calle23: set[str] = set()
        self.filas: dict[str, Fila] = {}
        self.no_leidos: list[tuple[str, str]] = []      # (archivo, motivo)
        self.constancias_sin_factura: list = []
        self.retenciones_sin_factura: list = []
        self.actualizado = ''

    # ------------------------------------------------------------------ carga
    def agregar_xml(self, archivos: list[tuple[str, bytes]]) -> tuple[int, int]:
        """Agrega XML (ya extraídos de ZIP). Devuelve (nuevos, repetidos). Los que no son comprobantes se informan."""
        nuevos = repetidos = 0
        for nombre, data in archivos:
            c = parse_xml_bytes(data, nombre)
            if c.error:
                self.no_leidos.append((nombre, c.error))
                continue
            pid = f'{c.ruc_emisor}|{c.tipo_codigo}|{c.serie_numero}'
            if pid in self.xml:
                repetidos += 1
                continue
            self.xml[pid] = (nombre, data)
            nuevos += 1
        return nuevos, repetidos

    def meses(self) -> dict[str, int]:
        from collections import Counter
        cnt = Counter()
        for nombre, data in self.xml.values():
            c = parse_xml_bytes(data, nombre)
            if c.fecha_emision:
                cnt[c.fecha_emision[:7]] += 1
        return dict(sorted(cnt.items()))

    # ------------------------------------------------------------------ proceso
    def procesar(self, m: Maestros, umbral_ia: int = 90):
        mem = m.memoria
        ctx = Contexto(plan=m.plan, tc=m.tc, memoria_proveedor=(mem.memoria_proveedor if mem else None),
                       historial_stats=dict(mem.frecuencia) if mem else {}, cuenta_haber='4212',
                       alerta_monto=20000, excluir_bancos=True, ia=m.ia, umbral_ia=umbral_ia, reglas=cfg.REGLAS)
        ctx.periodo = self.periodo
        comps = [parse_xml_bytes(b, n) for n, b in self.xml.values()]
        comps.sort(key=lambda c: (c.fecha_emision, c.serie_numero))
        props = procesar_lote(comps, ctx)
        if m.constancias:
            _, self.constancias_sin_factura = cruzar(props, m.constancias)

        filas: dict[str, Fila] = {}
        for p in props:
            d = self.decisiones.get(p.id) or Decision()
            f = Fila(p=p, d=d)
            self._defaults(f, m)
            self.decisiones[p.id] = d
            filas[p.id] = f
        self.filas = filas
        self._cruzar_retenciones(m)
        self.evaluar(m)
        self.actualizado = datetime.now().isoformat(timespec='seconds')

    def _defaults(self, f: Fila, m: Maestros):
        p, c, d = f.p, f.p.c, f.d
        mem = m.memoria
        pago_mem = mem.pago_reciente(c.ruc_emisor) if mem else ''
        # ---- estado ----
        clave = clave_comprobante(c.ruc_emisor, c.tipo_codigo, c.serie_numero)
        f.ya_en_diario = bool(mem and clave in mem.contabilizados)
        if d.estado in ('exportado', 'contabilizado'):
            pass
        elif f.ya_en_diario:
            d.estado, d.motivo = 'contabilizado', 'Ya está en el Diario de Contasis'
            if mem and clave in mem.asientos_por_clave:
                mes, n = mem.asientos_por_clave[clave]
                d.asiento = f'{mes}-{n}'
        elif p.estado in ('excluido', 'error', 'duplicado'):
            d.estado, d.motivo = 'excluido', p.motivo
        elif p.c.tiene_detraccion and not p.det_constancia:
            d.estado = 'excluido'
            d.motivo = (f'Detracción {c.detraccion_porcentaje}% sin constancia de depósito: sin depósito no hay crédito fiscal. '
                        'Cargue la constancia y vuelva a procesar.')
        elif d.estado == 'excluido' and d.motivo and not d.motivo.startswith('Manual'):
            d.estado, d.motivo = 'pendiente', ''      # ya no aplica la exclusión automática (p. ej. llegó la constancia)
        # ---- grupo y condición ----
        if not d.condicion_manual:
            d.condicion = 'CRE' if es_credito(c) else 'CON'
        if not d.grupo_manual:
            d.grupo = asignar_grupo(c, pago_mem, self.calle23)
        # ---- cuentas ----
        if not d.gasto_manual:
            d.cuenta_gasto = _ajuste_gasto(p, mem, m.plan)
        if not d.reparto_manual:
            suele = mem.suele_repartir(c.ruc_emisor) if mem else None
            d.reparto = cfg.es_repartible(d.cuenta_gasto) and suele is not False
        if not d.pago_manual:
            if d.condicion == 'CRE':
                d.cuenta_pago = cfg.CUENTA_CREDITO
            else:
                d.cuenta_pago = cuenta_pago_defecto(c, d.grupo, pago_mem, m.plan) if not es_credito(c) else \
                    ('104102' if c.moneda_codigo == 'USD' else '104103')
        if not d.fecha_pago:
            d.fecha_pago = c.fecha_emision
        self._glosa(f)

    def _glosa(self, f: Fila):
        d, c = f.d, f.c
        if d.glosa_manual:
            return
        d.glosa, _ = glosa_ralumin(d.cuenta_gasto, d.condicion == 'CRE', retencion=d.retencion, comision=bool(d.comision),
                                   detraccion=c.tiene_detraccion, respaldo=f.p.glosa, tipo_doc=c.tipo_codigo)

    def _cruzar_retenciones(self, m: Maestros):
        """Llena la retención con los XML de comprobantes de retención emitidos por Ralumin."""
        indice = {}
        for f in self.filas.values():
            s, n = normalizar_serie_numero(f.c.serie_numero)
            indice[(f.c.ruc_emisor, s, n)] = f
        sin = []
        for cr in m.retenciones:
            for ref in cr.referencias:
                s, n = normalizar_serie_numero(ref.serie_numero)
                f = indice.get((cr.ruc_proveedor, s, n))
                if not f:
                    sin.append((cr.serie_numero, cr.ruc_proveedor, ref.serie_numero, ref.retencion))
                    continue
                d = f.d
                if d.estado in ('exportado', 'contabilizado'):
                    continue
                d.retencion, d.retencion_manual = True, True
                d.retencion_monto = round(ref.retencion, 2)
                d.retencion_serie, d.retencion_numero, d.retencion_fecha = cr.serie, cr.numero, cr.fecha
                if ref.fecha_pago:
                    d.fecha_pago = ref.fecha_pago[:10]
        self.retenciones_sin_factura = sin

    # ------------------------------------------------------------------ evaluación (avisos, retenciones, semáforo)
    def evaluar(self, m: Maestros):
        plan = m.plan
        # suma por proveedor y fecha de pago, para la regla de los S/ 700
        sumas: dict[tuple, float] = defaultdict(float)
        for f in self.filas.values():
            if f.d.estado == 'excluido':
                continue
            f.retencion_exclusion = motivo_exclusion(f.c, m.padron, cfg.LEYENDAS_NO_RETENER)
            if not f.retencion_exclusion:
                sumas[(f.c.ruc_emisor, f.d.fecha_pago)] += f.total_soles
        for f in self.filas.values():
            d, c = f.d, f.c
            f.avisos = []
            rojo = amarillo = False
            if d.estado == 'excluido':
                f.nivel = 'ok'
                continue
            # ---- retención ----
            f.retencion_sugerida = False
            f.retencion_calculada = importe_retencion(f.total_soles, cfg.TASA_RETENCION)
            supera = sumas.get((c.ruc_emisor, d.fecha_pago), 0) > cfg.UMBRAL_RETENCION
            if not f.retencion_exclusion and supera and d.condicion == 'CON':
                f.retencion_sugerida = True
            if d.estado == 'pendiente' and not d.retencion_manual:
                d.retencion = f.retencion_sugerida
                d.retencion_monto = f.retencion_calculada if d.retencion else 0.0
            if f.retencion_sugerida and not d.retencion and not d.motivo_no_retencion:
                rojo = True
                f.avisos.append(f'Cumple las reglas de retención (S/ {f.retencion_calculada:,.2f}) pero no está marcada: márquela o indique el motivo.')
            if d.retencion:
                if d.condicion == 'CRE' or d.cuenta_pago.startswith('4212'):
                    rojo = True
                    f.avisos.append('Retención en una compra al crédito: se aplica al pagar (fuera de esta versión). Desmarque o cambie la cuenta de pago.')
                if f.retencion_exclusion:
                    amarillo = True
                    f.avisos.append(f'Retención marcada, pero la regla la excluye: {f.retencion_exclusion}.')
                if d.retencion_serie and abs(d.retencion_monto - f.retencion_calculada) > 0.011:
                    rojo = True
                    f.avisos.append(f'El comprobante de retención dice S/ {d.retencion_monto:,.2f} y el 3 % calculado es S/ {f.retencion_calculada:,.2f}.')
                if not d.retencion_serie:
                    amarillo = True
                    f.avisos.append('Retención sin comprobante de retención (serie-número): cárguelo o digítelo.')
                if f.retencion_sugerida and not m.padron:
                    amarillo = True
                    f.avisos.append('Verificar padrón: sin padrones SUNAT no se puede descartar buen contribuyente o agente.')
                if c.moneda_codigo == 'USD':
                    amarillo = True
                    f.avisos.append('Retención en dólares: se calcula en soles al TC de la fecha de pago; revisar.')
            # ---- comisión ----
            if d.comision and not d.cuenta_pago.startswith('104'):
                rojo = True
                f.avisos.append('La comisión bancaria solo va con una cuenta de banco (104).')
            # ---- cuenta de gasto ----
            if not d.cuenta_gasto:
                rojo = True
                f.avisos.append('Falta la cuenta de gasto.')
            elif plan:
                cu = plan.get(d.cuenta_gasto)
                if not cu:
                    rojo = True
                    f.avisos.append(f'La cuenta de gasto {d.cuenta_gasto} no existe en el plan.')
                elif not cu.es_hoja:
                    rojo = True
                    f.avisos.append(f'La cuenta {d.cuenta_gasto} tiene subcuentas: use una de último nivel.')
                if d.cuenta_gasto.startswith(('10', '40', '42')):
                    rojo = True
                    f.avisos.append('La cuenta de gasto no puede ser una 10, 40 o 42.')
            if kb.cuenta_bloqueada(d.cuenta_gasto, cfg.REGLAS):
                rojo = True
                f.avisos.append(f'La cuenta {d.cuenta_gasto} está bloqueada (SUNAT no reconoce la 6399).')
            # ---- cuenta de pago ----
            cp = d.cuenta_pago
            if not cp:
                rojo = True
                f.avisos.append('Falta la cuenta de pago.')
            else:
                if not (cp.startswith(cfg.RAICES_PAGO) or cp == cfg.CUENTA_CREDITO):
                    rojo = True
                    f.avisos.append(f'La cuenta de pago {cp} no es válida: use 101x, 104xxx o 4212.')
                if plan and not plan.existe(cp):
                    rojo = True
                    f.avisos.append(f'La cuenta de pago {cp} no existe en el plan.')
                if cp == cfg.CUENTA_CREDITO and d.condicion == 'CON':
                    rojo = True
                    f.avisos.append('Compra al contado con 4212: al contado va directo a la cuenta 10 (1011, 1012 o la 104 del banco).')
                if cp.startswith(cfg.RAICES_PAGO) and d.condicion == 'CRE':
                    amarillo = True
                    f.avisos.append('El XML dice crédito pero la cuenta de pago es una 10: confirme que se pagó al registrarla.')
                if cp.startswith(cfg.RAICES_PAGO) and es_cuenta_me(cp, plan) != (c.moneda_codigo == 'USD'):
                    amarillo = True
                    f.avisos.append(f'La cuenta {cp} es de otra moneda que la factura ({c.moneda_codigo}).')
            # ---- tipo de cambio ----
            if not f.p.tc:
                rojo = True
                f.avisos.append(f'Sin tipo de cambio SUNAT para {c.fecha_emision}: cargue el PDF del mes.')
            # ---- proveedor nuevo ----
            f.proveedor_nuevo = bool(m.entidades) and c.ruc_emisor not in m.entidades
            if f.proveedor_nuevo:
                amarillo = True
                f.avisos.append('Proveedor nuevo: no está en Entidades de Contasis; créelo antes de importar.')
            # ---- confianza de la cuenta propuesta ----
            if not d.gasto_manual:
                if f.p.confianza < 75:
                    rojo = True
                    f.avisos.append(f'Cuenta de gasto con baja confianza ({f.p.confianza} %): revisar.')
                elif f.p.confianza < 95:
                    amarillo = True
            if d.reparto:
                f.avisos.append(f'Reparto {cfg.PORCENTAJE_ADM:.0f}/{100 - cfg.PORCENTAJE_ADM:.0f}: va 100 % a {d.cuenta_gasto} y se ajusta con la hoja REPARTO 60-40.')
            # alertas del motor (detracción, IGV 10.5 %, activo fijo…)
            f.avisos.extend(a for a in f.p.alertas if 'tipo de cambio' not in a.lower())
            self._glosa(f)
            f.nivel = 'rojo' if rojo else ('amarillo' if amarillo else 'ok')

    # ------------------------------------------------------------------ edición
    def editar(self, pid: str, m: Maestros, _evaluar: bool = True, **cambios):
        f = self.filas[pid]
        d = f.d
        if d.estado in ('exportado', 'contabilizado'):
            raise ValueError('Ese comprobante ya está en un bloque: devuélvalo a pendiente para editarlo.')
        for k, v in cambios.items():
            if not hasattr(d, k):
                continue
            if isinstance(v, str):
                v = v.strip()
            if getattr(d, k) == v:
                continue
            setattr(d, k, v)
            marca = {'grupo': 'grupo_manual', 'cuenta_gasto': 'gasto_manual', 'cuenta_pago': 'pago_manual', 'reparto': 'reparto_manual',
                     'condicion': 'condicion_manual', 'retencion': 'retencion_manual', 'glosa': 'glosa_manual'}.get(k)
            if marca:
                setattr(d, marca, True)
            if k == 'grupo' and not d.pago_manual:
                g = cfg.GRUPOS_POR_ID.get(v)
                if g and g.cuenta_pago:
                    d.cuenta_pago = g.cuenta_pago if d.condicion == 'CON' or g.cuenta_pago == cfg.CUENTA_CREDITO else d.cuenta_pago
            if k == 'cuenta_gasto' and not d.reparto_manual:
                d.reparto = cfg.es_repartible(v)
            if k == 'retencion' and v and not d.retencion_monto:
                d.retencion_monto = f.retencion_calculada
            if k == 'retencion' and not v:
                d.retencion_monto = 0.0
            if k == 'estado' and v == 'excluido' and not d.motivo:
                d.motivo = 'Manual'
        if _evaluar:
            self.evaluar(m)

    def aplicar_a(self, ids: list[str], m: Maestros, **cambios):
        for pid in ids:
            if pid in self.filas and self.filas[pid].d.estado == 'pendiente':
                self.editar(pid, m, _evaluar=False, **cambios)
        self.evaluar(m)

    def marcar_calle23(self, ruc: str, si: bool, m: Maestros):
        (self.calle23.add if si else self.calle23.discard)(ruc)
        for f in self.filas.values():
            if f.c.ruc_emisor == ruc and f.d.estado == 'pendiente' and not f.d.grupo_manual:
                f.d.grupo = asignar_grupo(f.c, m.memoria.pago_reciente(ruc) if m.memoria else '', self.calle23)
        self.evaluar(m)

    # ------------------------------------------------------------------ consultas
    def lista(self, estado: str | None = 'pendiente') -> list[Fila]:
        fs = [f for f in self.filas.values() if estado is None or f.d.estado == estado]
        orden = {g.id: g.orden for g in cfg.GRUPOS}
        return sorted(fs, key=lambda f: (orden.get(f.d.grupo, 99), _orden_archivador(f)))

    def avance(self) -> dict:
        r = {e: 0 for e in ESTADOS}
        monto = {e: {'PEN': 0.0, 'USD': 0.0} for e in ESTADOS}
        for f in self.filas.values():
            r[f.d.estado] += 1
            monto[f.d.estado]['USD' if f.c.moneda_codigo == 'USD' else 'PEN'] += f.total
        r['total'] = sum(r[e] for e in ESTADOS if e != 'excluido')
        r['montos'] = monto
        return r

    def bloque(self, numero: int) -> Bloque | None:
        return next((b for b in self.bloques if b.numero == numero), None)

    # ------------------------------------------------------------------ bloques
    def crear_bloque(self, ids: list[str], nombre: str, m: Maestros) -> Bloque:
        self.evaluar(m)
        sel = [self.filas[i] for i in ids if i in self.filas]
        if not sel:
            raise ValueError('No hay comprobantes seleccionados.')
        no_pend = [f for f in sel if f.d.estado != 'pendiente']
        if no_pend:
            raise ValueError(f'{len(no_pend)} comprobante(s) no están pendientes (ya exportados, contabilizados o excluidos).')
        rojos = [f for f in sel if f.nivel == 'rojo']
        if rojos:
            raise ValueError(f'{len(rojos)} comprobante(s) en rojo: corríjalos antes de crear el bloque '
                             f'(por ejemplo {rojos[0].c.serie_numero}: {rojos[0].avisos[0] if rojos[0].avisos else ""}).')
        normales = sorted([f for f in sel if not f.tiene_ajuste], key=_orden_archivador)
        ajustes = sorted([f for f in sel if f.tiene_ajuste], key=_orden_archivador)
        numero = max([b.numero for b in self.bloques] + [0]) + 1
        b = Bloque(numero=numero, nombre=(nombre or '').strip() or f'Bloque {numero}', ids=[f.id for f in normales + ajustes],
                   creado=datetime.now().isoformat(timespec='seconds'))
        for f in normales + ajustes:
            f.d.estado, f.d.bloque = 'exportado', numero
        self.bloques.append(b)
        return b

    def filas_bloque(self, b: Bloque) -> list[Fila]:
        return [self.filas[i] for i in b.ids if i in self.filas]

    def devolver(self, ids: list[str]):
        """Devuelve a pendiente (la importación falló o hay que corregir)."""
        for pid in ids:
            f = self.filas.get(pid)
            if not f or f.d.estado not in ('exportado', 'contabilizado'):
                continue
            b = self.bloque(f.d.bloque)
            f.d.estado, f.d.bloque, f.d.asiento = 'pendiente', 0, ''
            if b:
                b.ids = [i for i in b.ids if i != pid]
                if not b.ids:
                    b.estado = 'anulado'

    def confirmar_bloque(self, numero: int, asiento_inicial: int = 0, mes: str = ''):
        b = self.bloque(numero)
        if not b:
            raise ValueError('Bloque no encontrado.')
        b.estado, b.confirmado = 'contabilizado', datetime.now().isoformat(timespec='seconds')
        b.asiento_inicial = int(asiento_inicial or 0)
        for k, f in enumerate(self.filas_bloque(b)):
            if f.d.estado == 'exportado':
                f.d.estado = 'contabilizado'
                f.d.motivo = f'Confirmado en el bloque {b.codigo}'
                if b.asiento_inicial:
                    f.d.asiento = f'{mes or self.periodo[5:7]}-{b.asiento_inicial + k}'

    def confirmar_con_diario(self, mem: MemoriaDiario) -> dict:
        """Marca como contabilizado todo lo que ya aparece en el Diario. Devuelve el detalle por bloque."""
        res = {'confirmados': 0, 'faltan': [], 'pendientes_en_diario': 0}
        for f in self.filas.values():
            k = clave_comprobante(f.c.ruc_emisor, f.c.tipo_codigo, f.c.serie_numero)
            esta = k in mem.contabilizados
            if esta and f.d.estado in ('exportado', 'pendiente'):
                if f.d.estado == 'pendiente':
                    res['pendientes_en_diario'] += 1
                else:
                    res['confirmados'] += 1
                f.d.estado = 'contabilizado'
                f.d.motivo = 'Confirmado con el Diario de Contasis'
                mes, n = mem.asientos_por_clave.get(k, ('', ''))
                f.d.asiento = f'{mes}-{n}' if n else ''
            elif not esta and f.d.estado == 'exportado':
                res['faltan'].append(f)
        for b in self.bloques:
            if b.estado == 'exportado' and all(self.filas[i].d.estado == 'contabilizado' for i in b.ids if i in self.filas):
                b.estado, b.confirmado = 'contabilizado', datetime.now().isoformat(timespec='seconds')
        return res

    # ------------------------------------------------------------------ persistencia
    def a_dict(self) -> dict:
        return {
            'version': 1, 'periodo': self.periodo, 'actualizado': self.actualizado,
            'xml': {pid: [n, _b64(b)] for pid, (n, b) in self.xml.items()},
            'decisiones': {pid: asdict(d) for pid, d in self.decisiones.items() if pid in self.xml},
            'bloques': [asdict(b) for b in self.bloques],
            'calle23': sorted(self.calle23),
        }

    @classmethod
    def desde_dict(cls, data: dict) -> 'Sesion':
        s = cls(data.get('periodo', ''))
        s.actualizado = data.get('actualizado', '')
        s.xml = {pid: (n, _unb64(b)) for pid, (n, b) in (data.get('xml') or {}).items()}
        s.decisiones = {pid: Decision.desde_dict(d) for pid, d in (data.get('decisiones') or {}).items()}
        s.bloques = [Bloque(**b) for b in data.get('bloques') or []]
        s.calle23 = set(data.get('calle23') or [])
        return s


def _ajuste_gasto(p: Propuesta, mem: MemoriaDiario | None, plan) -> str:
    """Correcciones propias de Ralumin sobre la cuenta que propone el motor general."""
    cuenta = p.cuenta
    if p.fuente in ('regla', 'memoria', 'manual'):
        return cuenta
    texto = f'{p.c.nombre_emisor} {texto_items(p.c)}'.upper()
    sin_memoria = not (mem and mem.gasto.get(p.c.ruc_emisor))
    # combustible solo si de verdad es combustible
    if cuenta.startswith('656009') and not any(k in texto for k in cfg.KW_COMBUSTIBLE):
        cuenta = cfg.CUENTA_SUMINISTROS
    # bienes sin ninguna pista → material auxiliar (lo más usado por Ralumin); queda en rojo para revisar
    if sin_memoria and 'sin palabra clave' in (p.explicacion or '') and (not plan or plan.existe(cfg.CUENTA_BIENES_DEFECTO)):
        cuenta = cfg.CUENTA_BIENES_DEFECTO
    return cuenta


def _orden_archivador(f: Fila) -> tuple:
    """Del más reciente al más antiguo, como están en el archivador físico."""
    c = f.c
    s, n = normalizar_serie_numero(c.serie_numero)
    return (_neg(c.fecha_emision), s, -int(n) if n.isdigit() else 0)


def _neg(fecha: str) -> int:
    try:
        return -int((fecha or '')[:10].replace('-', ''))
    except ValueError:
        return 0
