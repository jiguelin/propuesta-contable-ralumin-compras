# -*- coding: utf-8 -*-
"""
motor/glosas.py — Glosa con el estilo del Diario de Ralumin:
    <concepto de la cuenta> - EFECTIVO | - CREDITO  [- RETENCION] [+ GASTO BANCARIO] [- DETRA]
Ejemplos reales: "COMPRA MATERIALES Y SUMINISTROS - EFECTIVO - RETENCION", "NACIONALIZACION CREDITO - COSTO",
"GASTOS DE VIAJE - PEAJE - ESTACIONAMIENTO, ETC.". Siempre pasa por la limpieza para el PLE (60 caracteres).
"""
from __future__ import annotations

from ralumin import config as cfg

from .texto import normalizar


def concepto(cuenta: str) -> str:
    c = (cuenta or '').strip()
    mejor = ''
    for raiz in cfg.GLOSAS:
        if c.startswith(raiz) and len(raiz) > len(mejor):
            mejor = raiz
    return cfg.GLOSAS.get(mejor, '')


def glosa_ralumin(cuenta: str, credito: bool, retencion: bool = False, comision: bool = False, detraccion: bool = False,
                  respaldo: str = '', tipo_doc: str = '01') -> tuple[str, bool]:
    """Devuelve (glosa, hubo_que_reconstruir_texto)."""
    base = concepto(cuenta)
    c = (cuenta or '')
    if not base:
        base = (respaldo or 'COMPRA').upper()
        return normalizar(base, True, 60)
    if c.startswith('609192'):
        g = f'{base}{" CREDITO" if credito else ""} - COSTO'
    elif c.startswith(cfg.GLOSAS_SIN_FORMA_PAGO):
        g = base
    else:
        g = f'{base} - {"CREDITO" if credito else "EFECTIVO"}'
    if detraccion:
        g += ' - DETRA'
    if retencion:
        g += ' - RETENCION'
    if comision:
        g += ' + GASTO BANCARIO'
    if tipo_doc == '07':
        g = 'NC ' + g
    elif tipo_doc == '08':
        g = 'ND ' + g
    for largo, corto in ABREVIATURAS:          # que entre en 60 caracteres sin cortar la marca del final
        if len(g) <= 60:
            break
        g = g.replace(largo, corto)
    return normalizar(g, True, 60)


ABREVIATURAS = [('+ GASTO BANCARIO', '+ G.BANCARIO'), ('MATERIALES Y SUMINISTROS', 'MAT. Y SUMINISTROS'), (' - EFECTIVO', ' - EFECT.'),
                ('COMPRA ', ''), (' - RETENCION', ' - RETENC.')]
