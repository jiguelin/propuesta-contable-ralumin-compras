# -*- coding: utf-8 -*-
"""
motor/maestros.py — Plan de cuentas, Diario y Entidades de Ralumin: de dónde se leen y cómo se actualizan.

Orden de búsqueda de cada archivo:
  1. El almacén (la última versión subida desde el app).
  2. La carpeta `datos/` del repositorio (plan_cuentas.xlsx, diario.xlsx, entidades.xlsx).

"Actualizar maestros" en el app guarda el archivo nuevo en el almacén y, si hay GITHUB_TOKEN en los secrets,
también lo sube a `datos/` del repositorio con un commit (así queda versionado). Solo tiene sentido subirlos
al repositorio si es PRIVADO: en un repositorio público cualquiera podría descargarlos.
"""
from __future__ import annotations

import base64
import io
import json
import re
import urllib.error
import urllib.request
from pathlib import Path

import openpyxl

from .almacen import Almacen

DATOS = Path(__file__).resolve().parent.parent / 'datos'

ARCHIVOS = {
    'plan': ('plan_cuentas.xlsx', 'Plan de cuentas'),
    'diario': ('diario.xlsx', 'Diario Analítico (subdiario de compras)'),
    'entidades': ('entidades.xlsx', 'Entidades (clientes y proveedores)'),
}


def leer_archivo(alm: Almacen, nombre: str) -> tuple[bytes | None, str]:
    """Devuelve (bytes, origen) del maestro, o (None, '')."""
    b64 = alm.leer(f'maestro:{nombre}', json_=False)
    if b64:
        return base64.b64decode(b64), f'subido el {alm.fecha("maestro:" + nombre)[:16].replace("T", " ")}'
    ruta = DATOS / ARCHIVOS[nombre][0]
    if ruta.exists():
        return ruta.read_bytes(), f'repositorio (datos/{ruta.name})'
    return None, ''


def guardar_archivo(alm: Almacen, nombre: str, data: bytes):
    alm.guardar(f'maestro:{nombre}', base64.b64encode(data).decode())


def leer_entidades(data: bytes) -> dict[str, str]:
    """Excel 'CLIENTES - PROVEEDORES' de Contasis → {RUC: nombre}."""
    wb = openpyxl.load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    out = {}
    for ws in wb.worksheets:
        for r in ws.iter_rows(values_only=True):
            vals = [' '.join(str(v).split()) for v in r if v is not None and str(v).strip()]
            if len(vals) >= 1 and re.fullmatch(r'\d{8,11}', vals[0]):
                out[vals[0]] = vals[1] if len(vals) > 1 else ''
    return out


# ----------------------------------------------------------------------------- GitHub
def subir_a_github(token: str, repo: str, ruta: str, data: bytes, mensaje: str, rama: str = 'main') -> str:
    """Crea o reemplaza un archivo del repositorio con la API de GitHub. Devuelve el SHA del commit."""
    url = f'https://api.github.com/repos/{repo}/contents/{ruta}'
    cab = {'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28',
           'User-Agent': 'compras-ralumin'}
    sha = None
    try:
        with urllib.request.urlopen(urllib.request.Request(f'{url}?ref={rama}', headers=cab), timeout=30) as r:
            sha = json.loads(r.read())['sha']
    except urllib.error.HTTPError as ex:
        if ex.code != 404:
            raise
    cuerpo = {'message': mensaje, 'content': base64.b64encode(data).decode(), 'branch': rama}
    if sha:
        cuerpo['sha'] = sha
    req = urllib.request.Request(url, data=json.dumps(cuerpo).encode(), method='PUT', headers={**cab, 'Content-Type': 'application/json'})
    with urllib.request.urlopen(req, timeout=60) as r:
        return json.loads(r.read())['commit']['sha']


def repo_es_privado(token: str, repo: str) -> bool | None:
    try:
        req = urllib.request.Request(f'https://api.github.com/repos/{repo}',
                                     headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github+json', 'User-Agent': 'compras-ralumin'})
        with urllib.request.urlopen(req, timeout=20) as r:
            return bool(json.loads(r.read()).get('private'))
    except Exception:
        return None
