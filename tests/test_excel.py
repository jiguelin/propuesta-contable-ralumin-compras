# -*- coding: utf-8 -*-
"""Pruebas del modo Excel (cuando no hay XML). Datos ficticios."""
from __future__ import annotations

import io
from datetime import date, datetime

import openpyxl
from openpyxl import Workbook

from motor.diario import construir_memoria, leer_diario
from motor.excel_entrada import comprobante_desde_fila, leer_excel_compras
from motor.exportar import excel_bloque
from motor.sesion import Sesion
from motor.tipo_cambio import HistorialTC
from tests.fabrica import diario_excel, xml_factura
from tests.test_ralumin import PROV_CAJA, PROV_MAT, PROV_USD, fila, maestros


def excel_comprobantes(filas: list[dict]) -> bytes:
    """Excel con el orden de la plantilla de Contasis (A–AX), 3 filas vacías arriba, como el que entrega SUNAT/SIRE."""
    wb = Workbook()
    ws = wb.active
    ws.title = 'Comprobantes'
    for _ in range(3):
        ws.append([None] * 50)
    for f in filas:
        r = [None] * 50
        r[0] = f['fecha']
        r[1] = f.get('vence')
        r[2], r[3], r[5] = f.get('tipo', '01'), f['serie'], f['numero']
        r[6], r[7], r[8] = '6', f['ruc'], f['nombre']
        r[9], r[10] = f['base'], f['igv']
        for j in (11, 12, 13, 14, 15, 16, 17):
            r[j] = f.get({15: 'P', 17: 'R'}.get(j, ''), 0)
        r[18] = round(f['base'] + f['igv'] + f.get('P', 0) + f.get('R', 0), 2)
        r[22] = f.get('tc', 1)
        r[28] = f.get('usd')            # AC: solo en dólares (texto, como viene)
        r[29] = f.get('vence')
        r[48] = '0.00'
        ws.append(r)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


FILAS = [
    {'fecha': datetime(2026, 7, 20), 'serie': 'F010', 'numero': '100', 'ruc': PROV_CAJA, 'nombre': 'FERRE SAC', 'base': 100.0, 'igv': 18.0},
    # en dólares: AC = 294.88, W = 3.439, importes J–S en soles
    {'fecha': datetime(2026, 7, 14), 'serie': 'E001', 'numero': '1117', 'ruc': PROV_USD, 'nombre': 'PARTS EIRL', 'base': 859.41, 'igv': 154.69,
     'tc': 3.439, 'usd': '294.88'},
    # crédito por vencimiento posterior
    {'fecha': datetime(2026, 7, 3), 'vence': datetime(2026, 7, 15), 'serie': 'E001', 'numero': '1099', 'ruc': '20999999907', 'nombre': 'CONTA SAC',
     'base': 3155.0, 'igv': 567.9},
    # importación (courier) y periodo anterior
    {'fecha': datetime(2026, 7, 2), 'tipo': '53', 'serie': '235', 'numero': '01197741', 'ruc': '0', 'nombre': None, 'base': 8083.2, 'igv': 1454.98,
     'tc': 3.368, 'usd': '2832.0'},
    {'fecha': datetime(2026, 4, 22), 'serie': 'F001', 'numero': '77', 'ruc': PROV_MAT, 'nombre': 'MATERIALES SAC', 'base': 500.0, 'igv': 90.0},
]


def test_lee_excel_y_moneda_por_columna_AC():
    filas, avisos = leer_excel_compras(excel_comprobantes(FILAS))
    assert len(filas) == 5 and not avisos
    usd = comprobante_desde_fila(filas[1])
    assert usd.moneda_codigo == 'USD' and float(usd.importe_total) == 294.88
    assert abs(float(usd.gravadas) - 249.90) < 0.02                 # 859.41 × 294.88 / 1014.10
    pen = comprobante_desde_fila(filas[0])
    assert pen.moneda_codigo == 'PEN' and float(pen.importe_total) == 118.0 and pen.igv_tasa == 18
    assert comprobante_desde_fila(filas[2]).forma_pago == 'Credito'


def test_sesion_con_excel():
    s = Sesion('2026-07')
    n, r, _ = s.agregar_excel('octubre.xlsx', excel_comprobantes(FILAS))
    assert (n, r) == (5, 0)
    m = maestros()
    m.tc = HistorialTC([])                                          # sin PDF: el TC del Excel cubre su fecha
    s.procesar(m)
    caja = fila(s, 'F010-100')
    assert caja.origen == 'Excel' and caja.d.cuenta_gasto == '6560094' and caja.d.cuenta_pago == '1011'
    usd = fila(s, 'E001-1117')
    assert usd.p.tc == 3.439 and usd.d.grupo == 'usd_contado' and usd.total_soles == 1014.09
    assert fila(s, 'E001-1099').d.condicion == 'CRE' and fila(s, 'E001-1099').d.cuenta_pago == '4212'
    assert fila(s, '235-01197741').d.estado == 'excluido' and 'Importación' in fila(s, '235-01197741').d.motivo
    ant = fila(s, 'F001-77')
    assert ant.d.estado == 'pendiente' and any('periodo anterior' in a for a in ant.avisos)
    # el Excel del bloque lleva los importes en soles y el US$ tal cual vinieron
    s.editar(usd.id, m, revisado=True, retencion=False, motivo_no_retencion='buen contribuyente (prueba)')
    assert usd.nivel != 'rojo', usd.avisos
    b = s.crear_bloque([usd.id], 'USD', m)
    r = list(openpyxl.load_workbook(io.BytesIO(excel_bloque(s, b))).active.iter_rows(values_only=True))[0]
    assert (r[9], r[10], r[18], r[22], r[27], r[28]) == (249.9, 44.98, 294.88, 3.439, 'D', 294.88)


def test_credito_por_historial():
    mem = construir_memoria(leer_diario(diario_excel([
        {'mes': '06', 'n': 3, 'ruc': PROV_USD, 'doc': '01', 'numero': '00E001-0000000000050', 'fecha': date(2026, 6, 1),
         'lineas': [('6031025', 1000, 0), ('401111', 180, 0), ('4212', 0, 1180)]}])))
    m = maestros()
    m.memoria = mem
    s = Sesion('2026-07')
    s.agregar_excel('x.xlsx', excel_comprobantes([FILAS[1]]))
    s.procesar(m)
    f = fila(s, 'E001-1117')
    assert f.d.condicion == 'CRE' and f.d.cuenta_pago == '4212' and f.d.grupo == 'usd_credito'
    assert f.d.cuenta_gasto == '6031025'                                # cuenta del historial
    assert any('historial' in a for a in f.avisos)


def test_xml_y_excel_juntos_manda_el_xml():
    s = Sesion('2026-07')
    s.agregar_excel('x.xlsx', excel_comprobantes([FILAS[0]]))
    s.agregar_xml([('a.xml', xml_factura(PROV_CAJA, 'FERRE SAC', 'F010-00000100', '2026-07-20', 100, items=[('BROCAS', 100)]))])
    s.procesar(maestros())
    assert s.excel_cubiertos == 1
    activos = [f for f in s.filas.values() if f.c.ruc_emisor == PROV_CAJA]
    assert len(activos) == 1 and activos[0].origen == 'XML'


def test_sesion_json_con_excel():
    import json
    s = Sesion('2026-07')
    s.agregar_excel('x.xlsx', excel_comprobantes(FILAS[:2]))
    m = maestros()
    s.procesar(m)
    s.editar(fila(s, 'F010-100').id, m, cuenta_gasto='6560095')
    s2 = Sesion.desde_dict(json.loads(json.dumps(s.a_dict())))
    s2.procesar(m)
    assert fila(s2, 'F010-100').d.cuenta_gasto == '6560095' and len(s2.excel) == 2
