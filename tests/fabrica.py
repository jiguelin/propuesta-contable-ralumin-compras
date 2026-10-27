# -*- coding: utf-8 -*-
"""Arma XML UBL y Diarios de prueba con datos ficticios (el repositorio no guarda datos reales de clientes)."""
from __future__ import annotations

import io

from openpyxl import Workbook

RUC_EMPRESA = '20602768440'


def xml_factura(ruc: str, nombre: str, serie_numero: str, fecha: str, base: float, igv: float | None = None, moneda: str = 'PEN',
                forma_pago: str = 'Contado', cuotas: list[tuple[str, float]] | None = None, items: list[tuple[str, float]] | None = None,
                leyendas: list[tuple[str, str]] | None = None, detraccion: tuple[str, float, float] | None = None, tipo: str = '01') -> bytes:
    igv = round(base * 0.18, 2) if igv is None else igv
    total = round(base + igv, 2)
    items = items or [('PRODUCTO DE PRUEBA', base)]
    pagos = f'<cac:PaymentTerms><cbc:ID>FormaPago</cbc:ID><cbc:PaymentMeansID>{forma_pago}</cbc:PaymentMeansID></cac:PaymentTerms>'
    for k, (vence, monto) in enumerate(cuotas or [], start=1):
        pagos += (f'<cac:PaymentTerms><cbc:ID>FormaPago</cbc:ID><cbc:PaymentMeansID>Cuota{k:03d}</cbc:PaymentMeansID>'
                  f'<cbc:Amount currencyID="{moneda}">{monto}</cbc:Amount><cbc:PaymentDueDate>{vence}</cbc:PaymentDueDate></cac:PaymentTerms>')
    if detraccion:
        cod, pct, monto = detraccion
        pagos += (f'<cac:PaymentTerms><cbc:ID>Detraccion</cbc:ID><cbc:PaymentMeansID>{cod}</cbc:PaymentMeansID>'
                  f'<cbc:PaymentPercent>{pct}</cbc:PaymentPercent><cbc:Amount currencyID="PEN">{monto}</cbc:Amount></cac:PaymentTerms>')
    notas = ''.join(f'<cbc:Note languageLocaleID="{c}">{t}</cbc:Note>' for c, t in (leyendas or []))
    lineas = ''.join(
        f'<cac:InvoiceLine><cbc:ID>{k}</cbc:ID><cbc:InvoicedQuantity unitCode="NIU">1</cbc:InvoicedQuantity>'
        f'<cbc:LineExtensionAmount currencyID="{moneda}">{v}</cbc:LineExtensionAmount>'
        f'<cac:TaxTotal><cbc:TaxAmount currencyID="{moneda}">{round(v * 0.18, 2)}</cbc:TaxAmount><cac:TaxSubtotal>'
        f'<cbc:TaxAmount currencyID="{moneda}">{round(v * 0.18, 2)}</cbc:TaxAmount><cac:TaxCategory><cbc:Percent>18</cbc:Percent>'
        f'<cbc:TaxExemptionReasonCode>10</cbc:TaxExemptionReasonCode><cac:TaxScheme><cbc:ID>1000</cbc:ID><cbc:Name>IGV</cbc:Name></cac:TaxScheme>'
        f'</cac:TaxCategory></cac:TaxSubtotal></cac:TaxTotal><cac:Item><cbc:Description>{d}</cbc:Description></cac:Item>'
        f'<cac:Price><cbc:PriceAmount currencyID="{moneda}">{v}</cbc:PriceAmount></cac:Price></cac:InvoiceLine>'
        for k, (d, v) in enumerate(items, start=1))
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<Invoice xmlns="urn:oasis:names:specification:ubl:schema:xsd:Invoice-2"
 xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
 xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2">
<cbc:ID>{serie_numero}</cbc:ID><cbc:IssueDate>{fecha}</cbc:IssueDate>
<cbc:InvoiceTypeCode listID="0101">{tipo}</cbc:InvoiceTypeCode>{notas}
<cbc:DocumentCurrencyCode>{moneda}</cbc:DocumentCurrencyCode>
<cac:AccountingSupplierParty><cac:Party><cac:PartyIdentification><cbc:ID schemeID="6">{ruc}</cbc:ID></cac:PartyIdentification>
<cac:PartyLegalEntity><cbc:RegistrationName>{nombre}</cbc:RegistrationName></cac:PartyLegalEntity></cac:Party></cac:AccountingSupplierParty>
<cac:AccountingCustomerParty><cac:Party><cac:PartyIdentification><cbc:ID schemeID="6">{RUC_EMPRESA}</cbc:ID></cac:PartyIdentification>
<cac:PartyLegalEntity><cbc:RegistrationName>RALUMIN E.I.R.L.</cbc:RegistrationName></cac:PartyLegalEntity></cac:Party></cac:AccountingCustomerParty>
{pagos}
<cac:TaxTotal><cbc:TaxAmount currencyID="{moneda}">{igv}</cbc:TaxAmount><cac:TaxSubtotal><cbc:TaxableAmount currencyID="{moneda}">{base}</cbc:TaxableAmount>
<cbc:TaxAmount currencyID="{moneda}">{igv}</cbc:TaxAmount><cac:TaxCategory><cbc:Percent>18</cbc:Percent><cac:TaxScheme><cbc:ID>1000</cbc:ID><cbc:Name>IGV</cbc:Name></cac:TaxScheme></cac:TaxCategory></cac:TaxSubtotal></cac:TaxTotal>
<cac:LegalMonetaryTotal><cbc:LineExtensionAmount currencyID="{moneda}">{base}</cbc:LineExtensionAmount><cbc:TaxInclusiveAmount currencyID="{moneda}">{total}</cbc:TaxInclusiveAmount>
<cbc:PayableAmount currencyID="{moneda}">{total}</cbc:PayableAmount></cac:LegalMonetaryTotal>
{lineas}
</Invoice>'''.encode()


def xml_retencion(serie_numero: str, fecha: str, ruc_prov: str, nombre_prov: str, refs: list[tuple[str, str, float, float]]) -> bytes:
    """refs: [(serie-número factura, fecha, total, retención)]"""
    docs = ''.join(
        f'<sac:SUNATRetentionDocumentReference><cbc:ID schemeID="01">{sn}</cbc:ID><cbc:IssueDate>{f}</cbc:IssueDate>'
        f'<cbc:TotalInvoiceAmount currencyID="PEN">{tot}</cbc:TotalInvoiceAmount><cac:Payment><cbc:ID>1</cbc:ID>'
        f'<cbc:PaidAmount currencyID="PEN">{tot}</cbc:PaidAmount><cbc:PaidDate>{fecha}</cbc:PaidDate></cac:Payment>'
        f'<sac:SUNATRetentionInformation><sac:SUNATRetentionAmount currencyID="PEN">{ret}</sac:SUNATRetentionAmount>'
        f'<sac:SUNATRetentionDate>{fecha}</sac:SUNATRetentionDate><sac:SUNATNetTotalPaid currencyID="PEN">{round(tot - ret, 2)}</sac:SUNATNetTotalPaid>'
        f'</sac:SUNATRetentionInformation></sac:SUNATRetentionDocumentReference>' for sn, f, tot, ret in refs)
    total = round(sum(r[3] for r in refs), 2)
    return f'''<?xml version="1.0" encoding="UTF-8"?>
<Retention xmlns="urn:sunat:names:specification:ubl:peru:schema:xsd:Retention-1"
 xmlns:cac="urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2"
 xmlns:cbc="urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2"
 xmlns:sac="urn:sunat:names:specification:ubl:peru:schema:xsd:SunatAggregateComponents-1">
<cbc:ID>{serie_numero}</cbc:ID><cbc:IssueDate>{fecha}</cbc:IssueDate>
<cac:AgentParty><cac:PartyIdentification><cbc:ID schemeID="6">{RUC_EMPRESA}</cbc:ID></cac:PartyIdentification></cac:AgentParty>
<cac:ReceiverParty><cac:PartyIdentification><cbc:ID schemeID="6">{ruc_prov}</cbc:ID></cac:PartyIdentification>
<cac:PartyLegalEntity><cbc:RegistrationName>{nombre_prov}</cbc:RegistrationName></cac:PartyLegalEntity></cac:ReceiverParty>
<sac:SUNATRetentionSystemCode>01</sac:SUNATRetentionSystemCode><sac:SUNATRetentionPercent>3.00</sac:SUNATRetentionPercent>
<cbc:TotalAmount currencyID="PEN">{total}</cbc:TotalAmount>
{docs}
</Retention>'''.encode()


def plan_excel(cuentas: list[tuple[str, str]]) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.append([None, 'RALUMIN E.I.R.L.'])
    ws.append([None, 'R.U.C.: ' + RUC_EMPRESA])
    ws.append([])
    ws.append([None, 'CUENTA', None, None, 'DESCRIPCION', 'C.BAL', 'A.DEBE', 'A.HABER', 'TIPO', 'ANALISIS', 'CENTRO DE COSTOS'])
    for c, d in cuentas:
        ws.append([None, c, None, None, d, '', '', '', '', 'Ninguno', 'Sin centro de Costos'])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def diario_excel(asientos: list[dict]) -> bytes:
    """asientos: [{'mes': '07', 'n': 1, 'ruc':..., 'doc': '01', 'numero': '00F001-0000000000001', 'fecha': date, 'lineas': [(cta, debe, haber)], 'glosa': ''}]"""
    wb = Workbook()
    ws = wb.active
    ws.append([None, 'RALUMIN E.I.R.L.'])
    ws.append([None, 'LIBRO DIARIO DETALLADO'])
    ws.append([None, 'MES', 'FECHA', 'CUENTA', 'DESCRIPCION', 'DEBITO S/.', 'CREDITO S/.', 'M.', 'T/C', 'DEBITO US$.', 'CREDITO US$.',
               'GLOSA', 'R.U.C.', 'R.', 'DOC.', 'NÚMERO', 'FEC DOC.'])
    for a in asientos:
        ws.append([None, None, None, None, f'SUB DIARIO 01 ASIENTO No {a["n"]}'])
        for cta, debe, haber in a['lineas']:
            ws.append([None, a['mes'], a['fecha'], cta, 'X', debe, haber, 'S', 3.4, 0, 0, a.get('glosa', 'COMPRA'), a['ruc'], '', a['doc'],
                       a['numero'], a['fecha']])
        ws.append([None, None, None, None, 'TOTAL ASIENTO:'])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
