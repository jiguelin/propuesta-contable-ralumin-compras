# -*- coding: utf-8 -*-
"""Comprobantes de retención desde el Excel de SUNAT (cuando no hay XML). Datos ficticios."""
from __future__ import annotations

from motor.retenciones import leer_excel_retenciones, leer_retenciones
from motor.sesion import Sesion
from tests.fabrica import excel_retenciones_sunat, xml_factura
from tests.test_excel import FILAS, excel_comprobantes
from tests.test_ralumin import PROV_CAJA, PROV_MAT, PROV_USD, fila, maestros

REPORTE = [
    # factura en soles del mes, pagada al contado → cruza
    {'serie': 'E001 - 1052', 'fecha': '31/07/2026', 'ruc': PROV_MAT, 'nombre': 'MATERIALES SAC',
     'facturas': [('E001', '185', '31/07/2026', 'S/ 7600.01', 228.00, 7372.01)]},
    # factura en dólares (TC del pago distinto al de la emisión) → cruza; número con ceros a la izquierda
    {'serie': 'E001 - 1054', 'fecha': '14/07/2026', 'ruc': PROV_USD, 'nombre': 'PARTS EIRL',
     'facturas': [('E001', '00001117', '14/07/2026', '$ 294.88', 30.44, 983.66)]},
    # dos facturas de un año anterior (ya registradas al crédito) → no están en las compras del mes
    {'serie': 'E001 - 1047', 'fecha': '01/07/2026', 'ruc': '20999999909', 'nombre': 'HP DEMO SAC',
     'facturas': [('E001', '35', '17/11/2025', '$ 12551.85', 1294.24, 41846.47), ('E001', '44', '17/12/2025', '$ 10207.64', 1052.51, 34031.15)]},
    # anulado: no se aplica
    {'serie': 'E001 - 1060', 'fecha': '20/07/2026', 'ruc': PROV_CAJA, 'nombre': 'FERRE SAC', 'estado': 'ANULADO',
     'facturas': [('F010', '100', '20/07/2026', 'S/ 118.00', 3.54, 114.46)]},
]


def test_lee_reporte_sunat():
    crs, avisos = leer_excel_retenciones(excel_retenciones_sunat(REPORTE), 'JULIO.xlsx')
    assert not avisos and [c.serie_numero for c in crs] == ['E001-1052', 'E001-1054', 'E001-1047', 'E001-1060']
    a, b, c, d = crs
    assert (a.ruc_proveedor, a.nombre_proveedor, a.fecha, a.tasa, a.total_retenido) == (PROV_MAT, 'MATERIALES SAC', '2026-07-31', 3.0, 228.0)
    r = a.referencias[0]
    assert (r.tipo, r.serie_numero, r.fecha, r.total, r.retencion, r.moneda, r.fecha_pago) == ('01', 'E001-185', '2026-07-31', 7600.01, 228.0, 'PEN', '2026-07-31')
    u = b.referencias[0]
    assert u.moneda == 'USD' and u.total == 294.88 and u.tc_pago == 3.439
    assert len(c.referencias) == 2 and d.anulado and not a.anulado


def test_mismo_reporte_dos_veces_cuenta_una():
    data = excel_retenciones_sunat(REPORTE)
    crs, malos = leer_retenciones([('a.xlsx', data), ('a (1).xlsx', data)])
    assert len(crs) == 4 and not malos
    crs, malos = leer_retenciones([('otra.xlsx', b'no es excel')])
    assert crs == [] and malos == ['otra.xlsx']


def test_cruce_con_las_compras_del_mes():
    crs, _ = leer_retenciones([('JULIO.xlsx', excel_retenciones_sunat(REPORTE))])
    m = maestros(retenciones=crs)
    s = Sesion('2026-07')
    s.agregar_xml([('f.xml', xml_factura(PROV_MAT, 'MATERIALES SAC', 'E001-185', '2026-07-31', 6440.69, 1159.32))])
    s.agregar_excel('compras.xlsx', excel_comprobantes(FILAS[:2]))          # F010-100 (S/) y E001-1117 (US$)
    s.procesar(m)
    soles = fila(s, 'E001-185').d
    assert soles.retencion and soles.retencion_monto == 228.0
    assert (soles.retencion_serie, soles.retencion_numero, soles.retencion_fecha) == ('E001', '1052', '2026-07-31')
    usd = fila(s, 'E001-1117')
    assert usd.d.retencion and usd.d.retencion_monto == 30.44 and usd.d.fecha_pago == '2026-07-14'
    assert any('se usa el comprobante' in a for a in usd.avisos)
    assert not any('El comprobante de retención dice' in a for a in usd.avisos)    # diferencia por TC del pago: no es error
    # el anulado no marca nada; las facturas de 2025 quedan como "sin factura en el lote"
    assert not fila(s, 'F010-100').d.retencion_serie
    assert s.retenciones_anuladas == ['E001-1060']
    assert {x[2] for x in s.retenciones_sin_factura} == {'E001-35', 'E001-44'}
    assert {x[1] for x in s.retenciones_cruzadas} == {'E001-185', 'E001-1117'}


def test_con_reporte_manda_el_reporte_no_la_regla():
    """Dos facturas del mismo proveedor el mismo día suman más de S/ 700: la regla sugiere retención,
    pero si el reporte del mes no tiene comprobante para ellas, no se marcan (solo aviso amarillo)."""
    from motor.retenciones import ComprobanteRetencion, ReferenciaRetencion
    otro = ComprobanteRetencion('E001-1', '2026-07-31', '20999999909', 'OTRO SAC', 3.0, 3.0,
                                [ReferenciaRetencion('01', 'F001-9', '2026-07-31', 100.0, 100.0, 3.0)])
    xmls = [xml_factura(PROV_CAJA, 'FERRE SAC', 'F010-70', '2026-07-20', 400, items=[('BROCAS', 400)]),
            xml_factura(PROV_CAJA, 'FERRE SAC', 'F010-71', '2026-07-20', 300, items=[('BROCAS', 300)])]
    s = Sesion('2026-07')
    s.agregar_xml([(f'x{k}.xml', x) for k, x in enumerate(xmls)])
    s.procesar(maestros())                                                  # sin reporte: regla de los S/ 700
    assert fila(s, 'F010-70').d.retencion
    s2 = Sesion('2026-07')
    s2.agregar_xml([(f'x{k}.xml', x) for k, x in enumerate(xmls)])
    s2.procesar(maestros(retenciones=[otro]))                               # con reporte: manda el reporte
    f = fila(s2, 'F010-70')
    assert f.retencion_sugerida and not f.d.retencion and f.d.retencion_monto == 0
    assert f.nivel != 'rojo' or not any('reglas de retención' in a for a in f.avisos)
    assert any('no aparece en los comprobantes de retención' in a for a in f.avisos)
