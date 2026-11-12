# -*- coding: utf-8 -*-
"""
motor/almacen.py — Dónde se guarda el trabajo (clave → texto).

  · Postgres (recomendado en Streamlit Cloud, por ejemplo Neon gratis): poner DATABASE_URL en los secrets.
    Sobrevive a los reinicios y actualizaciones de Streamlit.
  · SQLite local (por defecto, `datos/ralumin.db`): sirve en local; en Streamlit Cloud se borra al reiniciar,
    por eso existe el respaldo JSON de la sesión.

Claves usadas: 'sesion:YYYY-MM', 'maestro:<nombre>' (archivos en base64), 'tc' (tipo de cambio), 'padron',
'retenciones:YYYY-MM', 'constancias:YYYY-MM'.

Rendimiento: el app vuelve a leer los maestros (plan, Diario…) en cada clic. Con una base en internet eso sería una
conexión nueva y varios MB por clic, así que:
  · cada valor leído queda en memoria; guardar() y borrar() lo actualizan (el app es el único que escribe);
  · en Postgres se reutiliza una sola conexión y, si se cortó (Neon apaga la base tras 5 minutos sin uso), se reconecta
    y se reintenta una vez.
"""
from __future__ import annotations

import json
import os
import sqlite3
import threading
from datetime import datetime
from pathlib import Path

_NADA = object()


class ErrorAlmacen(RuntimeError):
    """No se pudo conectar con la base de datos (DATABASE_URL mal escrita, clave errada, base apagada…)."""


class Almacen:
    def __init__(self, url: str | None = None, ruta_sqlite: str | None = None):
        self.url = (url or os.environ.get('DATABASE_URL') or '').strip()
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
        self._lock = threading.RLock()
        self._conexion = None
        self._valores: dict[str, object] = {}       # clave → texto (o _NADA si no existe)
        self._fechas: dict[str, str] = {}
        self._ejecutar('CREATE TABLE IF NOT EXISTS ralumin_kv (clave TEXT PRIMARY KEY, valor TEXT NOT NULL, actualizado TEXT NOT NULL)',
                       (), escribir=True)

    @property
    def persistente(self) -> bool:
        return self.tipo == 'postgres'

    @property
    def descripcion(self) -> str:
        if self.tipo != 'postgres':
            return 'SQLite local'
        try:
            from urllib.parse import urlparse
            return f'Postgres ({urlparse(self.url).hostname})'
        except Exception:
            return 'Postgres'

    # ------------------------------------------------------------------ conexión
    def _conectar(self):
        if self.tipo == 'postgres':
            import psycopg2
            try:
                return psycopg2.connect(self.url, connect_timeout=15)
            except Exception as ex:
                raise ErrorAlmacen(str(ex).strip().splitlines()[0] if str(ex).strip() else type(ex).__name__) from ex
        return sqlite3.connect(self.ruta, check_same_thread=False)

    def _ejecutar(self, sql: str, params: tuple, escribir: bool = False, todos: bool = False):
        """Ejecuta con la conexión reutilizada; si se cortó, reconecta y reintenta una vez."""
        if self.tipo != 'postgres':
            sql = sql.replace('%s', '?')
        with self._lock:
            for intento in (1, 2):
                try:
                    if self._conexion is None:
                        self._conexion = self._conectar()
                    cur = self._conexion.cursor()
                    cur.execute(sql, params)
                    r = cur.fetchall() if todos else (None if escribir else cur.fetchone())
                    self._conexion.commit()
                    return r
                except ErrorAlmacen:
                    raise
                except Exception:
                    try:
                        self._conexion.close()
                    except Exception:
                        pass
                    self._conexion = None
                    if intento == 2 or self.tipo != 'postgres':
                        raise

    # ------------------------------------------------------------------ operaciones
    def guardar(self, clave: str, valor) -> None:
        texto = valor if isinstance(valor, str) else json.dumps(valor, ensure_ascii=False)
        ahora = datetime.now().isoformat(timespec='seconds')
        self._ejecutar('INSERT INTO ralumin_kv (clave, valor, actualizado) VALUES (%s, %s, %s) '
                       'ON CONFLICT (clave) DO UPDATE SET valor = EXCLUDED.valor, actualizado = EXCLUDED.actualizado',
                       (clave, texto, ahora), escribir=True)
        with self._lock:
            self._valores[clave] = texto
            self._fechas[clave] = ahora

    def _texto(self, clave: str):
        with self._lock:
            if clave in self._valores:
                return self._valores[clave]
        r = self._ejecutar('SELECT valor, actualizado FROM ralumin_kv WHERE clave = %s', (clave,))
        with self._lock:
            self._valores[clave] = r[0] if r else _NADA
            if r:
                self._fechas[clave] = r[1]
            return self._valores[clave]

    def leer(self, clave: str, json_: bool = True, defecto=None):
        t = self._texto(clave)
        if t is _NADA:
            return defecto
        return json.loads(t) if json_ else t

    def fecha(self, clave: str) -> str:
        self._texto(clave)
        return self._fechas.get(clave, '')

    def claves(self, prefijo: str = '') -> list[str]:
        filas = self._ejecutar('SELECT clave FROM ralumin_kv WHERE clave LIKE %s ORDER BY clave', (prefijo + '%',), todos=True)
        return [r[0] for r in filas]

    def borrar(self, clave: str):
        self._ejecutar('DELETE FROM ralumin_kv WHERE clave = %s', (clave,), escribir=True)
        with self._lock:
            self._valores[clave] = _NADA
            self._fechas.pop(clave, None)

    def olvidar_cache(self):
        """Por si alguien cambió la base por fuera del app."""
        with self._lock:
            self._valores.clear()
            self._fechas.clear()
