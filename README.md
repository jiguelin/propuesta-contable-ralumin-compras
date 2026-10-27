# Compras Ralumin

App de compras **solo para RALUMIN E.I.R.L.** (RUC 20602768440). Recibe los XML de compras del mes, los reparte en los
mismos grupos que los archivadores físicos y entrega **un Excel por bloque**, listo para importar en Contasis
(plantilla "Formato Registro de Compras – NewContaSis", 50 columnas A–AX).

Se trabaja **por bloques**: buscar → seleccionar → crear bloque → descargar → importar → confirmar.
Lo que ya se contabilizó desaparece de la propuesta, así cada bloque se revisa solo y un error queda acotado a ese bloque.

```
XML ─▶ lector UBL ─▶ excluir (bancos, otro mes, detracción sin constancia, ya en el Diario)
    ─▶ grupo (moneda × forma de pago × costumbre del proveedor) ─▶ cuenta de gasto (reglas Ralumin → memoria del Diario → plan → IA)
    ─▶ cuenta de pago (1011 / 104xxx / 4212) ─▶ retención (regla de los S/ 700 y exclusiones SUNAT) ─▶ glosa estilo Ralumin
    ─▶ bloque ─▶ Excel Contasis + Excel de revisión (AJUSTES DE PAGO, REPARTO 60-40, PAGOS AGRUPADOS, PROVEEDORES NUEVOS…)
```

Basado en el motor de [Propuesta Contable de Compras](https://github.com/jiguelin/propuesta-contable-compras) v0.3.4
(copiado, no compartido: un cambio en un app no rompe el otro).

## Cómo se usa (cada mes)

1. **Barra lateral → Actualizar maestros:** subir el **Diario** de compras exportado de Contasis (acumulado) y el
   **PDF de tipo de cambio SUNAT** del mes. Si alguien creó o cambió cuentas, subir también el **plan de cuentas**;
   si hay proveedores nuevos, las **entidades**. Opcional: padrones SUNAT (buenos contribuyentes, agentes de retención y de percepción).
2. **1 · Cargar:** subir los XML del mes (sueltos o ZIP; se puede en varias tandas). Opcional: los XML de los comprobantes
   de retención que emite Ralumin y las constancias de detracción.
3. **2 · Propuesta:** elegir un grupo o **buscar** (RUC, proveedor, serie-número, importe, ítem). Corregir en la tabla
   la cuenta de gasto, la cuenta de pago, la retención o la **comisión bancaria**. Seleccionar y **Crear bloque**.
4. **3 · Bloques:** descargar el Excel del bloque, importarlo en Contasis y confirmar: con el botón o subiendo el Diario
   actualizado (el app marca solo lo que ya entró). Si algo falló: **Devolver a pendiente**.

Semáforo: 🟢 listo · 🟡 revisar · 🔴 no deja crear el bloque hasta corregirlo.

## Reglas de Ralumin (en `ralumin/config.py`)

| Tema | Regla |
|---|---|
| Grupos | Caja chica (1011) · Caja chica Calle 23 (1011) · Soles contado (104103) · Soles contado Calle 23 (104103) · Dólares contado (104102) · Dólares crédito (4212) · Dólares 60–120 días (4212) · Soles crédito (4212) |
| Contado | Siempre directo a la cuenta 10 con su divisionaria; **nunca 4212** (la 4212 solo para crédito). |
| Calle 23 | Mismas cuentas: 1011 en efectivo o la 104 del banco. Los grupos de Calle 23 solo siguen el orden del archivador. |
| Peajes | Una sola cuenta: **6315095**, sin reparto. |
| Reparto 60/40 | Suministros, combustible, telefonía y demás cuentas 094/095: se importa 100 % en la 094 y la hoja REPARTO 60-40 dice cómo partirla (hasta validar la prueba T3). |
| Retención | Ralumin es agente de retención: se marca sola (3 % del total) si la suma del mismo proveedor pagada en la misma fecha supera S/ 700, salvo exclusiones SUNAT (exonerada/inafecta, buen contribuyente, agente de retención o de percepción, detracción, boletas, recibos de luz/agua, liquidaciones de compra). |
| Comisión bancaria | Se escribe el monto en la fila (solo con cuenta 104). |
| Ajustes de pago | Las filas con retención o comisión van al final del Excel del bloque; la hoja AJUSTES DE PAGO dice qué agregar en cada asiento: 6391094 al Debe, 401141 al Haber y el monto final de la cuenta 10 (= total − retención + comisión). |
| Tipo de cambio | TC venta SUNAT de la fecha de emisión (PDF del mes); si falta, no se puede crear el bloque. |
| 6399 | Nunca se propone (SUNAT solo reconoce 6391 y 6392 en la 639). |
| Detracción | Sin constancia de depósito no entra al Excel. |

## Datos

- `datos/plan_cuentas.xlsx` y `datos/entidades.xlsx`: maestros de Ralumin. Se actualizan desde el app.
- El **Diario no se guarda en el repositorio**: se sube cada mes desde el app y queda en el almacén.
- Almacén: **Postgres** si hay `DATABASE_URL` en los secrets (recomendado en Streamlit Cloud: el disco se borra al reiniciar);
  si no, SQLite local (`datos/ralumin.db`). En la barra lateral se puede descargar y restaurar un respaldo JSON de la sesión.

## Secrets (Streamlit Cloud → Settings → Secrets)

```toml
APP_PASSWORD = "clave-para-entrar"
DATABASE_URL = "postgresql://usuario:clave@host:5432/postgres"   # recomendado (Supabase/Neon gratis)
GITHUB_TOKEN = "github_pat_..."     # opcional: para que "Actualizar maestros" guarde plan y entidades en datos/ del repositorio
ANTHROPIC_API_KEY = "sk-ant-..."    # opcional: IA solo en casos dudosos
```

`GITHUB_TOKEN`: token *fine-grained* con permiso **Contents: Read and write** solo sobre este repositorio.

## Instalación local

```bash
pip install -r requirements.txt
streamlit run app.py
python -m pytest -q          # pruebas con datos ficticios
```

## Estructura

```
app.py                 ← interfaz Streamlit (solo dibuja)
ralumin/config.py      ← todo lo propio de Ralumin: grupos, cuentas, reglas, glosas, umbrales
motor/
  sesion.py            ← el mes: comprobantes, decisiones, bloques y estados
  grupos.py            ← grupo y cuenta de pago
  retenciones.py       ← reglas de retención, XML del comprobante de retención, padrones
  diario.py            ← lector del Diario de Contasis, memoria por proveedor, "ya contabilizado"
  exportar.py          ← Excel del bloque, Excel de revisión, resumen del mes
  glosas.py            ← glosa con el estilo de Ralumin
  maestros.py          ← plan, entidades, Diario: lectura y actualización (almacén / GitHub)
  almacen.py           ← Postgres o SQLite
  (copiados del app de compras v0.3.4) xml_parser, clasificador, pcge, reglas, plan_cuentas,
  tipo_cambio, detracciones, excel_newcontasis, texto, ia
tests/                 ← pruebas con XML y Diarios ficticios
```

## Pruebas pendientes en Contasis

| Prueba | Qué se importa | Qué decide |
|---|---|---|
| T1 | Factura con AK–AP de retención y AH = 104103 | Si la retención puede ir por plantilla |
| T3 | Un comprobante en dos filas (60 % / 40 %) | Si el reparto 60/40 puede ir por plantilla |
| T4 | Tres facturas en un orden conocido | Si Contasis numera los asientos en el orden de las filas |
| T5 | Factura en dólares al crédito | Confirma importes en soles + TC + equivalente en US$ |
