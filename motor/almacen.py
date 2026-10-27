# -*- coding: utf-8 -*-
"""
motor/almacen.py — Dónde se guarda el trabajo (clave → texto).

  · Postgres (recomendado en Streamlit Cloud, por ejemplo Supabase gratis): poner DATABASE_URL en los secrets.
    Sobrevive a los reinicios de Streamlit.
  · SQLite local (por defecto, `datos/ralumin.db`): sirve en local; en Streamlit Cloud se borra al reiniciar,
    por eso cada descarga trae además el JSON de respaldo de la sesión.

Claves usadas: 'sesion:YYYY-MM', 'maestro:<nombre>' (archivos en base64), 'tc' (tipo de cambio), 'padron'.
"""
from __future__ import annotations

import json
import os
import sqlite3
from datetime import datetime
from pathlib import Path


class Almacen:
    def __init__(self, url: str | None = None, ruta_sqlite: str | None = None):
        self.url = url or os.environ.get('DATABASE_URL') or ''
        self.tipo = 'sqlite'
        if self.url.startswith(('postgres://', 'postgresql://')):
            try:
                import psycopg2  # noqa: F401
                self.tipo = 'postgres'
            except ImportError:
                self.tipo = 'sqlite'
        if self.tipo == 'sqlite':
            ruta = ruta_sqlite or os.environ.get('RALUMIN_DB') or str(Path(__file__).resolve().parent.parent / 'datos' / 'ralumin.db')
            Path(ruta).parent.mkdir(parents=True, exist_ok=True)
            self.ruta = ruta
        self._crear()

    @property
    def persistente(self) -> bool:
        return self.tipo == 'postgres'

    # ------------------------------------------------------------------ conexión
    def _con(self):
        if self.tipo == 'postgres':
            import psycopg2
            return psycopg2.connect(self.url)
        return sqlite3.connect(self.ruta)

    def _crear(self):
        con = self._con()
        try:
            cur = con.cursor()
            cur.execute('CREATE TABLE IF NOT EXISTS ralumin_kv (clave TEXT PRIMARY KEY, valor TEXT NOT NULL, actualizado TEXT NOT NULL)')
            con.commit()
        finally:
            con.close()

    # ------------------------------------------------------------------ operaciones
    def guardar(self, clave: str, valor) -> None:
        texto = valor if isinstance(valor, str) else json.dumps(valor, ensure_ascii=False)
        ahora = datetime.now().isoformat(timespec='seconds')
        con = self._con()
        try:
            cur = con.cursor()
            if self.tipo == 'postgres':
                cur.execute('INSERT INTO ralumin_kv (clave, valor, actualizado) VALUES (%s, %s, %s) '
                            'ON CONFLICT (clave) DO UPDATE SET valor = EXCLUDED.valor, actualizado = EXCLUDED.actualizado',
                            (clave, texto, ahora))
            else:
                cur.execute('INSERT INTO ralumin_kv (clave, valor, actualizado) VALUES (?, ?, ?) '
                            'ON CONFLICT(clave) DO UPDATE SET valor = excluded.valor, actualizado = excluded.actualizado',
                            (clave, texto, ahora))
            con.commit()
        finally:
            con.close()

    def leer(self, clave: str, json_: bool = True, defecto=None):
        con = self._con()
        try:
            cur = con.cursor()
            cur.execute('SELECT valor FROM ralumin_kv WHERE clave = ' + ('%s' if self.tipo == 'postgres' else '?'), (clave,))
            r = cur.fetchone()
        finally:
            con.close()
        if not r:
            return defecto
        return json.loads(r[0]) if json_ else r[0]

    def fecha(self, clave: str) -> str:
        con = self._con()
        try:
            cur = con.cursor()
            cur.execute('SELECT actualizado FROM ralumin_kv WHERE clave = ' + ('%s' if self.tipo == 'postgres' else '?'), (clave,))
            r = cur.fetchone()
        finally:
            con.close()
        return r[0] if r else ''

    def claves(self, prefijo: str = '') -> list[str]:
        con = self._con()
        try:
            cur = con.cursor()
            cur.execute('SELECT clave FROM ralumin_kv WHERE clave LIKE ' + ('%s' if self.tipo == 'postgres' else '?') + ' ORDER BY clave',
                        (prefijo + '%',))
            return [r[0] for r in cur.fetchall()]
        finally:
            con.close()

    def borrar(self, clave: str):
        con = self._con()
        try:
            cur = con.cursor()
            cur.execute('DELETE FROM ralumin_kv WHERE clave = ' + ('%s' if self.tipo == 'postgres' else '?'), (clave,))
            con.commit()
        finally:
            con.close()
