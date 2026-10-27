# -*- coding: utf-8 -*-
"""
motor/grupos.py — Asigna cada comprobante a un grupo (archivador) y propone su cuenta de pago.

El XML dice la moneda y la forma de pago, pero no cómo se pagó. Reglas, en orden:
  1. Crédito (forma de pago "Credito" o cuotas): dólares → "Dólares crédito" o "60–120 días" según el plazo;
     soles → "Soles crédito". Cuenta de pago 4212.
  2. Contado en dólares → "Dólares contado" (104102).
  3. Contado en soles → "Caja chica" (1011) si la última cuenta de pago del proveedor en el Diario fue 1011,
     si es un rubro de caja chica (peajes, grifos, estacionamientos, ferreterías) o si el total no pasa del umbral;
     si la última cuenta de pago fue una 104, o nada de lo anterior aplica → "Soles contado" (104103).
  4. Calle 23: proveedores marcados por el usuario → la variante "Calle 23" del grupo.
Las compras al contado nunca van a 4212 (decisión de Jiguel, 26/10/2026).
"""
from __future__ import annotations

from datetime import datetime

from ralumin import config as cfg

from .xml_parser import RUCS_COMBUSTIBLE, texto_items


def es_credito(c) -> bool:
    return 'CRED' in (c.forma_pago or '').upper() or bool(c.cuotas)


def plazo_dias(c) -> int:
    """Días entre la emisión y el último vencimiento (cuotas o fecha de vencimiento)."""
    fechas = [q.get('vence') for q in c.cuotas if q.get('vence')] + ([c.fecha_vencimiento] if c.fecha_vencimiento else [])
    try:
        emi = datetime.strptime(c.fecha_emision[:10], '%Y-%m-%d')
        ult = max(datetime.strptime(f[:10], '%Y-%m-%d') for f in fechas)
        return max(0, (ult - emi).days)
    except (ValueError, TypeError):
        return 0


def es_cuenta_me(cuenta: str, plan=None) -> bool:
    """¿La cuenta de pago es en moneda extranjera? (1012, 104102, 104104… o descripción con ME / US$)."""
    c = (cuenta or '').strip()
    if c in ('1012', '104102', '104104'):
        return True
    cu = plan.get(c) if plan else None
    if cu:
        d = f' {cu.descripcion.upper()} '
        return any(k in d for k in (' ME ', ' ME-', '-ME', 'US$', 'DOLAR', 'DÓLAR', ' USD'))
    return False


def rubro_caja_chica(c) -> bool:
    if c.ruc_emisor in RUCS_COMBUSTIBLE:
        return True
    t = f' {c.nombre_emisor} {texto_items(c)} '.upper()
    return any(k in t for k in cfg.KW_CAJA_CHICA)


def asignar_grupo(c, pago_reciente: str = '', calle23: set[str] | None = None) -> str:
    calle23 = calle23 or set()
    usd = c.moneda_codigo == 'USD'
    if es_credito(c):
        if usd:
            g = 'usd_60_120' if plazo_dias(c) >= cfg.DIAS_60_120 else 'usd_credito'
        else:
            g = 'soles_credito'
        return g
    if usd:
        return 'usd_contado'
    if pago_reciente.startswith('101'):
        g = 'caja_chica'
    elif pago_reciente.startswith('104'):
        g = 'soles_contado'
    elif rubro_caja_chica(c) or float(c.importe_total) <= cfg.UMBRAL_CAJA_CHICA:
        g = 'caja_chica'
    else:
        g = 'soles_contado'
    if c.ruc_emisor in calle23:
        g += '_c23'
    return g


def cuenta_pago_defecto(c, grupo: str, pago_reciente: str = '', plan=None) -> str:
    """Cuenta de pago propuesta. Crédito → 4212. Contado → la del grupo, o la 10 que el proveedor usó
    la última vez si es de la misma moneda (por ejemplo otro banco)."""
    if es_credito(c):
        return cfg.CUENTA_CREDITO
    g = cfg.GRUPOS_POR_ID.get(grupo)
    base = g.cuenta_pago if g and g.cuenta_pago and g.cuenta_pago != cfg.CUENTA_CREDITO else ('104102' if c.moneda_codigo == 'USD' else '1011')
    if pago_reciente.startswith(cfg.RAICES_PAGO) and es_cuenta_me(pago_reciente, plan) == (c.moneda_codigo == 'USD'):
        # el grupo manda sobre efectivo vs banco; la memoria solo elige el banco concreto
        if base.startswith('104') and pago_reciente.startswith('104'):
            return pago_reciente
        if base.startswith('101') and pago_reciente.startswith('101'):
            return pago_reciente
    return base
