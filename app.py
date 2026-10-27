# -*- coding: utf-8 -*-
"""
Compras Ralumin — interfaz Streamlit.

    streamlit run app.py

Flujo: cargar XML → revisar por grupo con el buscador → seleccionar → crear bloque → descargar su Excel →
importar en Contasis → confirmar (botón o Diario). Lo contabilizado sale de la propuesta.
Toda la lógica vive en `motor/` y `ralumin/`; este archivo solo dibuja.
"""
from __future__ import annotations

import base64
import hashlib
import json
import os
from datetime import date, datetime

import pandas as pd
import streamlit as st

from motor import __version__
from motor.almacen import Almacen
from motor.detracciones import leer_constancias
from motor.diario import construir_memoria, leer_diario
from motor.exportar import excel_bloque, excel_mes, excel_revision_bloque
from motor.ia import crear_ia
from motor.maestros import ARCHIVOS, guardar_archivo, leer_archivo, leer_entidades, repo_es_privado, subir_a_github
from motor.plan_cuentas import leer_plan_excel
from motor.retenciones import leer_padron, leer_retenciones
from motor.sesion import Maestros, Sesion
from motor.tipo_cambio import TC, HistorialTC, leer_pdf_sunat, leer_tabla
from motor.xml_parser import parse_xml_bytes
from ralumin import config as cfg

st.set_page_config(page_title='Compras Ralumin', page_icon='🧾', layout='wide')
st.markdown("""
<style>
  .block-container {padding-top: 1.4rem; max-width: 1500px;}
  h1 {font-weight: 700; letter-spacing: -0.02em; margin-bottom: 0;}
  .stMetric {background: #f7f7f9; border-radius: 12px; padding: 8px 14px;}
  .pequeno {font-size: 0.82rem; color: #6b7280;}
</style>
""", unsafe_allow_html=True)


def secreto(clave: str, defecto: str = '') -> str:
    try:
        v = st.secrets.get(clave)
    except Exception:
        v = None
    return v or os.environ.get(clave) or defecto


# ----------------------------------------------------------------------------- clave de acceso
CLAVE = secreto('APP_PASSWORD')
if CLAVE and not st.session_state.get('acceso'):
    st.title('Compras Ralumin')
    with st.form('acceso_form'):
        intento = st.text_input('Clave de acceso', type='password')
        if st.form_submit_button('Ingresar', type='primary'):
            if intento == CLAVE:
                st.session_state['acceso'] = True
                st.rerun()
            else:
                st.error('Clave incorrecta.')
    st.stop()


# ----------------------------------------------------------------------------- almacén y maestros
@st.cache_resource
def almacen() -> Almacen:
    return Almacen(url=secreto('DATABASE_URL'))


ALM = almacen()
REPO = secreto('GITHUB_REPO', 'jiguelin/propuesta-contable-ralumin-compras')
TOKEN = secreto('GITHUB_TOKEN')


def _hash(b: bytes) -> str:
    return hashlib.sha1(b).hexdigest()


def cargar_maestros() -> tuple[Maestros, dict]:
    """Lee plan, entidades, Diario, TC, padrón y archivos del periodo; cachea lo pesado por contenido."""
    cache = st.session_state.setdefault('_cache_maestros', {})
    info = {}
    m = Maestros()
    data, origen = leer_archivo(ALM, 'plan')
    if data:
        k = 'plan:' + _hash(data)
        if k not in cache:
            cache[k] = leer_plan_excel(data)
        m.plan = cache[k]
        info['plan'] = f'{len(m.plan.cuentas):,} cuentas · {origen}'
    data, origen = leer_archivo(ALM, 'entidades')
    if data:
        k = 'ent:' + _hash(data)
        if k not in cache:
            cache[k] = leer_entidades(data)
        m.entidades = cache[k]
        info['entidades'] = f'{len(m.entidades):,} RUC · {origen}'
    data, origen = leer_archivo(ALM, 'diario')
    if data:
        k = 'dia:' + _hash(data)
        if k not in cache:
            cache[k] = construir_memoria(leer_diario(data))
        m.memoria = cache[k]
        meses = m.memoria.meses
        info['diario'] = f'{m.memoria.n_asientos:,} asientos ({meses[0]}–{meses[-1]}) · {origen}' if meses else origen
    m.tc = HistorialTC([TC(date.fromisoformat(f), c, v) for f, c, v in ALM.leer('tc', defecto=[])])
    m.padron = set(ALM.leer('padron', defecto=[]))
    per = st.session_state.get('periodo', '')
    if per:
        archivos = [(n, base64.b64decode(b)) for n, b in ALM.leer(f'retenciones:{per}', defecto=[])]
        m.retenciones = leer_retenciones(archivos)[0] if archivos else []
        archivos = [(n, base64.b64decode(b)) for n, b in ALM.leer(f'constancias:{per}', defecto=[])]
        m.constancias = leer_constancias(archivos) if archivos else []
    if st.session_state.get('usar_ia') and secreto('ANTHROPIC_API_KEY'):
        m.ia = crear_ia(secreto('ANTHROPIC_API_KEY'))
    return m, info


def guardar_sesion(s: Sesion):
    ALM.guardar(f'sesion:{s.periodo}', s.a_dict())


def abrir_sesion(periodo: str) -> Sesion:
    data = ALM.leer(f'sesion:{periodo}')
    return Sesion.desde_dict(data) if data else Sesion(periodo)


def reprocesar():
    s: Sesion = st.session_state['sesion']
    m, _ = cargar_maestros()
    s.procesar(m)
    guardar_sesion(s)
    st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1


def extraer_xml(archivos):
    import io
    import zipfile
    out = []
    for nombre, data in archivos:
        if nombre.lower().endswith('.zip'):
            with zipfile.ZipFile(io.BytesIO(data)) as z:
                for n in z.namelist():
                    if n.lower().endswith('.xml') and not n.startswith('__MACOSX'):
                        out.append((n.split('/')[-1], z.read(n)))
        elif nombre.lower().endswith('.xml'):
            out.append((nombre, data))
    return out


# ----------------------------------------------------------------------------- periodo y sesión
def meses_opciones() -> list[str]:
    hoy = date.today()
    ms = []
    y, mth = hoy.year, hoy.month
    for _ in range(14):
        ms.append(f'{y}-{mth:02d}')
        mth -= 1
        if mth == 0:
            y, mth = y - 1, 12
    guardados = [k.split(':', 1)[1] for k in ALM.claves('sesion:')]
    return sorted(set(ms) | set(guardados), reverse=True)


with st.sidebar:
    st.markdown(f'### 🧾 Compras Ralumin\n<span class="pequeno">v{__version__} · {cfg.EMPRESA} · RUC {cfg.RUC}</span>', unsafe_allow_html=True)
    opciones = meses_opciones()
    guardadas = sorted(ALM.claves('sesion:'), key=ALM.fecha, reverse=True)
    # por defecto: el último mes en que se trabajó; si no hay ninguno, el mes anterior
    previo = st.session_state.get('periodo') or (guardadas[0].split(':', 1)[1] if guardadas else (opciones[1] if len(opciones) > 1 else opciones[0]))
    periodo = st.selectbox('Mes a trabajar', opciones, index=opciones.index(previo) if previo in opciones else 0)
    if st.session_state.get('periodo') != periodo or 'sesion' not in st.session_state:
        st.session_state['periodo'] = periodo
        st.session_state['sesion'] = abrir_sesion(periodo)
        st.session_state['sel'] = set()
        if st.session_state['sesion'].xml:
            reprocesar()

S: Sesion = st.session_state['sesion']
M, INFO = cargar_maestros()
if S.xml and not S.filas:
    S.procesar(M)

# ----------------------------------------------------------------------------- barra lateral: maestros
with st.sidebar:
    st.markdown('**Maestros**')
    st.markdown(f"<span class='pequeno'>Plan: {INFO.get('plan', '⚠️ falta')}<br>Entidades: {INFO.get('entidades', '⚠️ falta')}<br>"
                f"Diario: {INFO.get('diario', '⚠️ falta')}<br>Tipo de cambio: {', '.join(M.tc.meses_cargados()) or '⚠️ falta'}<br>"
                f"Padrones SUNAT: {len(M.padron):,} RUC</span>", unsafe_allow_html=True)
    if M.tc and periodo not in M.tc.meses_cargados():
        st.error(f'Falta el PDF de tipo de cambio SUNAT de {periodo}: sin él no se puede descargar.')

    with st.expander('Actualizar maestros'):
        st.caption('Sube la versión nueva exportada de Contasis. El app la usa desde ese momento.')
        for clave, (archivo, titulo) in ARCHIVOS.items():
            up = st.file_uploader(titulo, type=['xlsx', 'xls'], key=f'up_{clave}')
            if up is not None and st.button(f'Guardar {titulo.split(" (")[0].lower()}', key=f'btn_{clave}'):
                data = up.getvalue()
                try:
                    if clave == 'plan':
                        n = len(leer_plan_excel(data).cuentas)
                        assert n, 'No encontré cuentas en el archivo.'
                    elif clave == 'entidades':
                        n = len(leer_entidades(data))
                        assert n, 'No encontré RUC en el archivo.'
                    else:
                        n = len(leer_diario(data))
                        assert n, 'No encontré asientos en el archivo.'
                    guardar_archivo(ALM, clave, data)
                    msg = f'{titulo}: guardado ({n:,}).'
                    if TOKEN and clave in ('plan', 'entidades'):
                        privado = repo_es_privado(TOKEN, REPO)
                        sha = subir_a_github(TOKEN, REPO, f'datos/{archivo}', data, f'Actualiza {archivo} desde el app ({datetime.now():%Y-%m-%d %H:%M})')
                        msg += f' También se actualizó en GitHub (commit {sha[:7]}).' + ('' if privado else ' Ojo: el repositorio es público.')
                    st.success(msg)
                    if clave == 'diario' and S.xml:
                        m2, _ = cargar_maestros()
                        r = S.confirmar_con_diario(m2.memoria)
                        st.info(f'Con este Diario: {r["confirmados"]} comprobante(s) de bloques confirmados, '
                                f'{r["pendientes_en_diario"]} pendientes que ya estaban en Contasis, {len(r["faltan"])} exportados que aún no aparecen.')
                    reprocesar()
                except Exception as ex:
                    st.error(f'No se pudo leer: {ex}')
        tcs = st.file_uploader('Tipo de cambio SUNAT (PDF del mes, o Excel fecha|compra|venta)', type=['pdf', 'xlsx', 'csv'], accept_multiple_files=True, key='up_tc')
        if tcs and st.button('Guardar tipo de cambio'):
            actual = {f: (c, v) for f, c, v in ALM.leer('tc', defecto=[])}
            n = 0
            for u in tcs:
                try:
                    items = leer_pdf_sunat(u.getvalue()) if u.name.lower().endswith('.pdf') else leer_tabla(u.getvalue())
                    for t in items:
                        actual[t.fecha.isoformat()] = (t.compra, t.venta)
                        n += 1
                except Exception as ex:
                    st.error(f'{u.name}: {ex}')
            ALM.guardar('tc', [[f, c, v] for f, (c, v) in sorted(actual.items())])
            st.success(f'{n} días de tipo de cambio guardados.')
            reprocesar()
        pads = st.file_uploader('Padrones SUNAT (buenos contribuyentes, agentes de retención y de percepción)', type=['txt', 'zip', 'csv', 'xlsx'],
                                accept_multiple_files=True, key='up_padron')
        if pads and st.button('Guardar padrones'):
            rucs = leer_padron([(u.name, u.getvalue()) for u in pads])
            ALM.guardar('padron', sorted(rucs))
            st.success(f'{len(rucs):,} RUC en padrones.')
            reprocesar()

    with st.expander('Calle 23'):
        st.caption('Proveedores cuyas compras van a los grupos de Calle 23.')
        provs = sorted({(f.c.ruc_emisor, f.c.nombre_emisor) for f in S.filas.values()}, key=lambda x: x[1])
        etiquetas = {r: f'{n[:38]} ({r})' for r, n in provs}
        elegidos = st.multiselect('Proveedores', list(etiquetas), placeholder='Elegir proveedores', default=[r for r in S.calle23 if r in etiquetas], format_func=lambda r: etiquetas.get(r, r))
        if set(elegidos) != (S.calle23 & set(etiquetas)):
            for r in set(etiquetas):
                S.marcar_calle23(r, r in elegidos, M)
            guardar_sesion(S)
            st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
            st.rerun()

    with st.expander('Opciones y respaldo'):
        if secreto('ANTHROPIC_API_KEY'):
            st.checkbox('Usar IA en los casos dudosos', key='usar_ia', value=st.session_state.get('usar_ia', False))
        st.download_button('Descargar respaldo de la sesión (JSON)', json.dumps(S.a_dict()).encode(), f'SESION_{periodo}.json', 'application/json')
        rest = st.file_uploader('Restaurar respaldo', type=['json'], key='up_resp')
        if rest is not None and st.button('Restaurar'):
            st.session_state['sesion'] = Sesion.desde_dict(json.loads(rest.getvalue()))
            st.session_state['periodo'] = st.session_state['sesion'].periodo
            reprocesar()
            st.rerun()
        if not ALM.persistente:
            st.warning('Almacén local (SQLite): en Streamlit Cloud se borra al reiniciar. Configure DATABASE_URL (Postgres) '
                       'o descargue el respaldo al terminar.')

# ----------------------------------------------------------------------------- cabecera y avance
st.title('Compras Ralumin')
av = S.avance()
c1, c2, c3, c4, c5 = st.columns(5)
c1.metric('Comprobantes del mes', av['total'])
c2.metric('Pendientes', av['pendiente'])
c3.metric('En bloques exportados', av['exportado'])
c4.metric('Contabilizados', av['contabilizado'])
c5.metric('Pendiente S/ · US$', f"{av['montos']['pendiente']['PEN']:,.0f} · {av['montos']['pendiente']['USD']:,.0f}")
if av['total']:
    st.progress(av['contabilizado'] / av['total'], text=f"{av['contabilizado'] / av['total']:.0%} contabilizado")

tab_carga, tab_prop, tab_bloq, tab_otros = st.tabs(['1 · Cargar', '2 · Propuesta', '3 · Bloques', 'Excluidos y contabilizados'])

# ----------------------------------------------------------------------------- 1. carga
with tab_carga:
    st.markdown('Sube los XML de compras del mes (sueltos o en ZIP). Se pueden subir en varias tandas: los repetidos se ignoran.')
    ups = st.file_uploader('XML o ZIP de compras', type=['xml', 'zip'], accept_multiple_files=True, key=f'up_xml_{st.session_state.get("up_v", 0)}')
    if ups and st.button('Agregar y procesar', type='primary'):
        xmls = extraer_xml([(u.name, u.getvalue()) for u in ups])
        nuevos, rep = S.agregar_xml(xmls)
        reprocesar()
        st.session_state['up_v'] = st.session_state.get('up_v', 0) + 1
        st.success(f'{nuevos} comprobante(s) nuevos, {rep} repetidos.')
        meses = S.meses()
        otros = {k: v for k, v in meses.items() if k != periodo}
        if otros:
            st.warning(f'Hay comprobantes de otros meses ({", ".join(f"{k}: {v}" for k, v in otros.items())}); quedan en Excluidos.')
    col1, col2 = st.columns(2)
    with col1:
        cre = st.file_uploader('Comprobantes de retención emitidos por Ralumin (XML o ZIP)', type=['xml', 'zip'], accept_multiple_files=True, key='up_cre')
        if cre and st.button('Guardar comprobantes de retención'):
            prev = ALM.leer(f'retenciones:{periodo}', defecto=[])
            nuevos = [(u.name, base64.b64encode(u.getvalue()).decode()) for u in cre]
            ok, malos = leer_retenciones([(u.name, u.getvalue()) for u in cre])
            ALM.guardar(f'retenciones:{periodo}', prev + nuevos)
            st.success(f'{len(ok)} comprobante(s) de retención leídos.' + (f' No se pudieron leer: {", ".join(malos)}' if malos else ''))
            reprocesar()
    with col2:
        cons = st.file_uploader('Constancias de detracción (TXT/CSV/Excel de SUNAT, PDF o ZIP)', type=['txt', 'csv', 'xlsx', 'pdf', 'zip'],
                                accept_multiple_files=True, key='up_cons')
        if cons and st.button('Guardar constancias'):
            prev = ALM.leer(f'constancias:{periodo}', defecto=[])
            ALM.guardar(f'constancias:{periodo}', prev + [(u.name, base64.b64encode(u.getvalue()).decode()) for u in cons])
            st.success('Constancias guardadas.')
            reprocesar()
    if S.no_leidos:
        with st.expander(f'{len(S.no_leidos)} archivo(s) que no son comprobantes de compra'):
            st.dataframe(pd.DataFrame(S.no_leidos, columns=['Archivo', 'Motivo']), hide_index=True, width='stretch')
    if S.retenciones_sin_factura:
        st.warning(f'{len(S.retenciones_sin_factura)} referencia(s) de comprobantes de retención sin factura en el lote: '
                   + ', '.join(f'{a} → {c}' for a, _, c, _ in S.retenciones_sin_factura[:8]))
    if S.xml and st.button('Volver a procesar todo'):
        reprocesar()
        st.rerun()

# ----------------------------------------------------------------------------- 2. propuesta
GRUPO_ID = {g.nombre: g.id for g in cfg.GRUPOS}


def opciones_pago() -> list[str]:
    base = ['1011', '1012', '104101', '104102', '104103', '104104', '104105', '4212']
    if M.plan:
        base = sorted({c.codigo for c in M.plan.hojas() if c.codigo.startswith(cfg.RAICES_PAGO)} | {'4212'})
    return base


def tabla(filas) -> pd.DataFrame:
    sel = st.session_state.setdefault('sel', set())
    reg = []
    for f in filas:
        d, c = f.d, f.c
        reg.append({
            'id': f.id, 'Sel': f.id in sel, '': f.semaforo, 'Grupo': cfg.GRUPO_NOMBRE.get(d.grupo, d.grupo), 'Fecha': c.fecha_emision,
            'Comprobante': c.serie_numero, 'RUC': c.ruc_emisor, 'Proveedor': c.nombre_emisor[:40], 'Mon': 'US$' if c.moneda_codigo == 'USD' else 'S/',
            'Total': round(f.total, 2), 'Cuenta gasto': d.cuenta_gasto, '60/40': d.reparto, 'Cuenta pago': d.cuenta_pago, 'Cond': d.condicion,
            'Retención': d.retencion, 'Ret. S/': round(d.retencion_monto or 0, 2), 'Comp. retención': f'{d.retencion_serie}-{d.retencion_numero}' if d.retencion_serie else '',
            'Comisión': round(d.comision or 0, 2), 'Pago agrupado': d.pago_agrupado, 'Revisado': d.revisado, 'Glosa': d.glosa,
            'Avisos': ' | '.join(f.avisos),
        })
    return pd.DataFrame(reg)


def aplicar_edicion(original: pd.DataFrame, editado: pd.DataFrame) -> int:
    n = 0
    sel = st.session_state.setdefault('sel', set())
    mapa = {'Cuenta gasto': 'cuenta_gasto', '60/40': 'reparto', 'Cuenta pago': 'cuenta_pago', 'Cond': 'condicion', 'Retención': 'retencion',
            'Ret. S/': 'retencion_monto', 'Comisión': 'comision', 'Pago agrupado': 'pago_agrupado', 'Revisado': 'revisado', 'Glosa': 'glosa'}
    for (_, a), (_, b) in zip(original.iterrows(), editado.iterrows()):
        pid = a['id']
        if bool(b['Sel']) != bool(a['Sel']):
            (sel.add if b['Sel'] else sel.discard)(pid)
        cambios = {}
        for col, campo in mapa.items():
            va, vb = a[col], b[col]
            if (pd.isna(vb) and pd.isna(va)) or va == vb:
                continue
            cambios[campo] = (float(vb) if campo in ('retencion_monto', 'comision') else (bool(vb) if isinstance(va, (bool,)) or campo in
                              ('reparto', 'retencion', 'pago_agrupado', 'revisado') else str(vb or '')))
        if a['Grupo'] != b['Grupo']:
            cambios['grupo'] = GRUPO_ID.get(b['Grupo'], b['Grupo'])
        if b['Comp. retención'] != a['Comp. retención']:
            sn = str(b['Comp. retención'] or '').strip().upper()
            serie, num = (sn.split('-', 1) + [''])[:2] if sn else ('', '')
            cambios.update(retencion_serie=serie, retencion_numero=num.lstrip('0'))
        if cambios:
            try:
                S.editar(pid, M, _evaluar=False, **cambios)
                n += 1
            except ValueError as ex:
                st.error(str(ex))
    if n:
        S.evaluar(M)
    return n


with tab_prop:
    if not S.filas:
        st.info('Primero sube los XML en la pestaña 1 · Cargar, y verifica en la barra lateral que estén el plan, el Diario y el tipo de cambio.')
    else:
        pend = S.lista('pendiente')
        conteo = {g.id: sum(1 for f in pend if f.d.grupo == g.id) for g in cfg.GRUPOS}
        op_grupo = ['Todos'] + [g.id for g in cfg.GRUPOS if conteo[g.id]]
        etiqueta = {'Todos': f'Todos ({len(pend)})', **{g.id: f'{g.nombre} ({conteo[g.id]})' for g in cfg.GRUPOS}}
        grupo_sel = st.radio('Grupo', op_grupo, horizontal=True, format_func=lambda x: etiqueta[x], key='grupo_sel')
        fc1, fc2, fc3, fc4 = st.columns([3, 1.2, 1.2, 1.2])
        buscar = fc1.text_input('🔎 Buscar en todos los grupos (RUC, proveedor, serie-número, importe, ítem o cuenta)', key='buscar')
        mon = fc2.selectbox('Moneda', ['Todas', 'S/', 'US$'])
        nivel = fc3.selectbox('Semáforo', ['Todos', '🔴 rojo', '🟡 amarillo', '🟢 verde'])
        solo_nr = fc4.checkbox('Solo no revisados')

        def pasa(f) -> bool:
            if buscar.strip():
                t = buscar.strip().upper()
                campos = ' '.join([f.c.ruc_emisor, f.c.nombre_emisor, f.c.serie_numero, f'{f.total:.2f}', f.d.cuenta_gasto, f.d.cuenta_pago,
                                   ' '.join(l.descripcion for l in f.c.lineas)]).upper()
                if not all(p in campos for p in t.split()):
                    return False
            elif grupo_sel != 'Todos' and f.d.grupo != grupo_sel:
                return False
            if mon != 'Todas' and (('US$' if f.c.moneda_codigo == 'USD' else 'S/') != mon):
                return False
            if nivel != 'Todos' and f.nivel != {'🔴 rojo': 'rojo', '🟡 amarillo': 'amarillo', '🟢 verde': 'ok'}[nivel]:
                return False
            if solo_nr and f.d.revisado:
                return False
            return True

        vista = [f for f in pend if pasa(f)]
        sel = st.session_state.setdefault('sel', set())
        sel &= {f.id for f in pend}
        b1, b2, b3, _ = st.columns([1.6, 1.4, 1.2, 3])
        if b1.button(f'☑️ Seleccionar visibles ({len(vista)})'):
            sel |= {f.id for f in vista}
            st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
            st.rerun()
        if b2.button('🟢 Solo las verdes'):
            sel |= {f.id for f in vista if f.nivel == 'ok'}
            st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
            st.rerun()
        if b3.button('Quitar selección'):
            sel.clear()
            st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
            st.rerun()

        df = tabla(vista)
        if df.empty:
            st.info('Nada pendiente con estos filtros.')
        else:
            editado = st.data_editor(
                df, key=f'editor_{st.session_state.get("editor_v", 0)}', hide_index=True, width='stretch', height=min(640, 38 + 35 * len(df)),
                column_order=['Sel', '', 'Grupo', 'Fecha', 'Comprobante', 'RUC', 'Proveedor', 'Mon', 'Total', 'Cuenta gasto', '60/40', 'Cuenta pago',
                              'Cond', 'Retención', 'Ret. S/', 'Comp. retención', 'Comisión', 'Pago agrupado', 'Revisado', 'Glosa', 'Avisos'],
                disabled=['', 'Fecha', 'Comprobante', 'RUC', 'Proveedor', 'Mon', 'Total', 'Avisos'],
                column_config={
                    'Sel': st.column_config.CheckboxColumn('Sel', width='small'),
                    'Grupo': st.column_config.SelectboxColumn('Grupo', options=[g.nombre for g in cfg.GRUPOS], width='medium'),
                    'Total': st.column_config.NumberColumn('Total', format='%.2f'),
                    'Cuenta gasto': st.column_config.TextColumn('Cuenta gasto', help='Cuenta de último nivel del plan (Debe, columna AF).'),
                    'Cuenta pago': st.column_config.SelectboxColumn('Cuenta pago', options=opciones_pago(), help='Haber, columna AH.'),
                    'Cond': st.column_config.SelectboxColumn('Cond', options=['CON', 'CRE'], width='small'),
                    'Ret. S/': st.column_config.NumberColumn('Ret. S/', format='%.2f', min_value=0.0),
                    'Comp. retención': st.column_config.TextColumn('Comp. retención', help='Serie-número del comprobante de retención, p. ej. E001-989'),
                    'Comisión': st.column_config.NumberColumn('Comisión', format='%.2f', min_value=0.0, help='Comisión bancaria de la transferencia.'),
                    'Glosa': st.column_config.TextColumn('Glosa', max_chars=60, width='large'),
                    'Avisos': st.column_config.TextColumn('Avisos', width='large'),
                })
            if aplicar_edicion(df, editado):
                guardar_sesion(S)
                st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
                st.rerun()

        # ---- acciones sobre la selección ----
        sel_filas = [S.filas[i] for i in sel if i in S.filas]
        st.markdown(f'**Seleccionadas: {len(sel_filas)}**' + (f" · S/ {sum(f.total for f in sel_filas if f.c.moneda_codigo != 'USD'):,.2f}"
                                                               f" · US$ {sum(f.total for f in sel_filas if f.c.moneda_codigo == 'USD'):,.2f}"
                                                               f" · 🔴 {sum(1 for f in sel_filas if f.nivel == 'rojo')}" if sel_filas else ''))
        if sel_filas:
            a1, a2, a3, a4 = st.columns(4)
            with a1:
                cp = st.selectbox('Cuenta de pago', opciones_pago(), key='acc_cp')
                if st.button('Poner cuenta de pago'):
                    S.aplicar_a(list(sel), M, cuenta_pago=cp)
                    guardar_sesion(S)
                    st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
                    st.rerun()
            with a2:
                cg = st.text_input('Cuenta de gasto', key='acc_cg')
                if st.button('Poner cuenta de gasto') and cg.strip():
                    S.aplicar_a(list(sel), M, cuenta_gasto=cg.strip())
                    guardar_sesion(S)
                    st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
                    st.rerun()
            with a3:
                gr = st.selectbox('Mover a grupo', [g.nombre for g in cfg.GRUPOS if g.id != 'sin_grupo'], key='acc_gr')
                if st.button('Mover'):
                    S.aplicar_a(list(sel), M, grupo=GRUPO_ID[gr])
                    guardar_sesion(S)
                    st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
                    st.rerun()
            with a4:
                com = st.number_input('Comisión bancaria', min_value=0.0, value=4.30, step=0.10, format='%.2f', key='acc_com')
                if st.button('Poner comisión'):
                    S.aplicar_a(list(sel), M, comision=float(com))
                    guardar_sesion(S)
                    st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
                    st.rerun()
            b1, b2, b3 = st.columns([1, 1, 2])
            with b1:
                if st.button('Marcar revisadas'):
                    S.aplicar_a(list(sel), M, revisado=True)
                    guardar_sesion(S)
                    st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
                    st.rerun()
            with b2:
                motivo = st.text_input('Motivo para excluir', key='acc_mot')
                if st.button('Excluir') and motivo.strip():
                    S.aplicar_a(list(sel), M, estado='excluido', motivo=f'Manual: {motivo.strip()}')
                    sel.clear()
                    guardar_sesion(S)
                    st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
                    st.rerun()
            with b3:
                sugerido = ''
                grupos_sel = {f.d.grupo for f in sel_filas}
                if len(grupos_sel) == 1:
                    sugerido = cfg.GRUPO_NOMBRE.get(next(iter(grupos_sel)), '')
                nombre = st.text_input('Nombre del bloque', value=sugerido, key='acc_nom')
                if st.button(f'📦 Crear bloque con {len(sel_filas)} comprobante(s)', type='primary'):
                    try:
                        b = S.crear_bloque(list(sel), nombre, M)
                        sel.clear()
                        guardar_sesion(S)
                        st.session_state['ultimo_bloque'] = b.numero
                        st.session_state['editor_v'] = st.session_state.get('editor_v', 0) + 1
                        st.rerun()
                    except ValueError as ex:
                        st.error(str(ex))
        if st.session_state.get('ultimo_bloque'):
            b = S.bloque(st.session_state['ultimo_bloque'])
            if b:
                st.success(f'Bloque {b.codigo} «{b.nombre}» creado con {len(b.ids)} comprobantes. Descárgalo en la pestaña 3 · Bloques.')

        # ---- detalle ----
        with st.expander('Detalle de un comprobante e historial del proveedor'):
            if vista:
                elegido = st.selectbox('Comprobante', [f.id for f in vista],
                                       format_func=lambda i: f'{S.filas[i].c.serie_numero} · {S.filas[i].c.nombre_emisor[:40]} · {S.filas[i].total:,.2f}')
                f = S.filas[elegido]
                c = f.c
                d1, d2 = st.columns(2)
                with d1:
                    st.markdown(f'**{c.nombre_emisor}** · RUC {c.ruc_emisor}  \n{c.tipo_nombre} {c.serie_numero} del {c.fecha_emision}  \n'
                                f'Forma de pago: {c.forma_pago or "-"}{" · " + str(len(c.cuotas)) + " cuota(s)" if c.cuotas else ""} · '
                                f'Moneda {c.moneda_codigo} · TC {f.p.tc or "-"}')
                    st.dataframe(pd.DataFrame([{'Ítem': l.descripcion, 'Cant.': float(l.cantidad), 'Valor': float(l.valor_venta)} for l in c.lineas]),
                                 hide_index=True, width='stretch')
                    if c.leyendas or c.observaciones:
                        st.caption('Leyendas: ' + ' | '.join(c.leyendas + c.observaciones))
                with d2:
                    st.markdown(f'**Propuesta:** {f.p.explicacion}')
                    if f.retencion_exclusion:
                        st.caption(f'Retención: no aplica ({f.retencion_exclusion}).')
                    elif f.retencion_sugerida:
                        st.caption(f'Retención sugerida: S/ {f.retencion_calculada:,.2f}.')
                    if M.memoria and M.memoria.ultimas.get(c.ruc_emisor):
                        st.markdown('**Últimas compras de este proveedor en el Diario**')
                        st.dataframe(pd.DataFrame(M.memoria.ultimas[c.ruc_emisor], columns=['Fecha', 'Documento', 'Cuenta gasto', 'Cuenta pago', 'Total']),
                                     hide_index=True, width='stretch')
                    else:
                        st.caption('Sin compras anteriores en el Diario cargado.')
                    for a in f.avisos:
                        st.markdown(f'- {a}')

        with st.expander('Buscar una cuenta en el plan'):
            q = st.text_input('Código o palabras de la descripción', key='buscar_cta')
            if q and M.plan:
                st.dataframe(pd.DataFrame([{'Cuenta': c.codigo, 'Descripción': c.descripcion, 'Movimiento': 'sí' if c.es_hoja else 'no (título)'}
                                           for c in M.plan.buscar(q, 25)]), hide_index=True, width='stretch')

# ----------------------------------------------------------------------------- 3. bloques
with tab_bloq:
    if not S.bloques:
        st.info('Todavía no hay bloques. En 2 · Propuesta selecciona comprobantes y pulsa «Crear bloque».')
    else:
        st.download_button('Resumen del mes (Excel)', excel_mes(S, M.plan), f'RESUMEN_{periodo}.xlsx')
        st.caption('Después de importar un bloque en Contasis, confírmalo con el botón o subiendo el Diario actualizado '
                   '(barra lateral → Actualizar maestros → Diario): el app marca solo lo que ya entró.')
    for b in sorted(S.bloques, key=lambda x: -x.numero):
        fs = S.filas_bloque(b)
        if b.estado == 'anulado' and not fs:
            continue
        icono = {'exportado': '📤', 'contabilizado': '✅', 'anulado': '⚪'}[b.estado]
        tot_s = sum(f.total for f in fs if f.c.moneda_codigo != 'USD')
        tot_d = sum(f.total for f in fs if f.c.moneda_codigo == 'USD')
        with st.expander(f'{icono} {b.codigo} · {b.nombre} · {len(fs)} comprobantes · S/ {tot_s:,.2f} · US$ {tot_d:,.2f} · {b.estado}',
                         expanded=(b.numero == st.session_state.get('ultimo_bloque'))):
            k1, k2, k3 = st.columns(3)
            nombre_arch = f'CONTASIS_{periodo.replace("-", "")}_{b.codigo}_{"".join(ch for ch in b.nombre.upper() if ch.isalnum())[:20]}.xlsx'
            k1.download_button('⬇️ Excel para Contasis', excel_bloque(S, b), nombre_arch, type='primary', key=f'dx_{b.numero}')
            k2.download_button('Excel de revisión del bloque', excel_revision_bloque(S, b, M.plan), f'REVISION_{periodo.replace("-", "")}_{b.codigo}.xlsx',
                               key=f'dr_{b.numero}')
            aj = sum(1 for f in fs if f.tiene_ajuste)
            if aj:
                k3.warning(f'{aj} con comisión o retención: completar con la hoja AJUSTES DE PAGO.')
            st.dataframe(pd.DataFrame([{'Comprobante': f.c.serie_numero, 'Proveedor': f.c.nombre_emisor[:40], 'Total': f.total, 'Mon': f.c.moneda_codigo,
                                        'Gasto': f.d.cuenta_gasto, 'Pago': f.d.cuenta_pago, 'Ret.': f'{f.d.retencion_monto:,.2f}' if f.d.retencion else '',
                                        'Comisión': f'{f.d.comision:,.2f}' if f.d.comision else '', 'Estado': f.d.estado, 'Asiento': f.d.asiento} for f in fs]),
                         hide_index=True, width='stretch')
            if b.estado == 'exportado':
                c1, c2, c3 = st.columns(3)
                with c1:
                    ai = st.number_input('Primer número de asiento en Contasis (opcional)', min_value=0, step=1, key=f'ai_{b.numero}')
                    if st.button('✅ Entró bien: confirmar contabilizado', key=f'ok_{b.numero}'):
                        S.confirmar_bloque(b.numero, int(ai), periodo[5:7])
                        guardar_sesion(S)
                        st.rerun()
                with c2:
                    dev = st.multiselect('Devolver a pendiente', [f.id for f in fs if f.d.estado == 'exportado'], key=f'dv_{b.numero}', placeholder='Elegir comprobantes',
                                         format_func=lambda i: S.filas[i].c.serie_numero)
                    if st.button('Devolver seleccionados', key=f'dvb_{b.numero}') and dev:
                        S.devolver(dev)
                        guardar_sesion(S)
                        st.rerun()
                with c3:
                    if st.button('Devolver todo el bloque', key=f'dvt_{b.numero}'):
                        S.devolver([f.id for f in fs])
                        guardar_sesion(S)
                        st.rerun()

# ----------------------------------------------------------------------------- otros
with tab_otros:
    exc = S.lista('excluido')
    st.markdown(f'**Excluidos ({len(exc)})** — bancos (subdiario 22), otro mes, detracción sin constancia, no comprobantes, excluidos a mano.')
    if exc:
        st.dataframe(pd.DataFrame([{'Comprobante': f.c.serie_numero, 'RUC': f.c.ruc_emisor, 'Proveedor': f.c.nombre_emisor[:40], 'Fecha': f.c.fecha_emision,
                                    'Total': f.total, 'Mon': f.c.moneda_codigo, 'Motivo': f.d.motivo} for f in exc]), hide_index=True, width='stretch')
        react = st.multiselect('Reactivar (vuelven a pendiente)', [f.id for f in exc if f.d.motivo.startswith('Manual')], placeholder='Elegir',
                               format_func=lambda i: S.filas[i].c.serie_numero)
        if react and st.button('Reactivar'):
            for i in react:
                S.editar(i, M, _evaluar=False, estado='pendiente', motivo='')
            S.evaluar(M)
            guardar_sesion(S)
            st.rerun()
    con = S.lista('contabilizado')
    st.markdown(f'**Contabilizados ({len(con)})**')
    if con:
        st.dataframe(pd.DataFrame([{'Comprobante': f.c.serie_numero, 'Proveedor': f.c.nombre_emisor[:40], 'Total': f.total, 'Mon': f.c.moneda_codigo,
                                    'Bloque': f'B{f.d.bloque:02d}' if f.d.bloque else '', 'Asiento': f.d.asiento, 'Cómo': f.d.motivo} for f in con]),
                     hide_index=True, width='stretch')

st.caption(f'Almacén: {"Postgres" if ALM.persistente else "SQLite local"} · Sesión guardada {S.actualizado.replace("T", " ") or "-"}')
