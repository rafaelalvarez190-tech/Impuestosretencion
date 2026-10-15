# Autorretención sobre ingresos (Python + Streamlit)

Calcula la **liquidación preliminar de autorretención** a partir del auxiliar contable.

- **Base** = suma neta (con signo) de la columna `valor` de las cuentas cuyo código empieza por **4**, por empresa y período. No usa `debito`, `credito`, `saldoanterior` ni `saldoactual`.
- **Autorretención** = Base × Porcentaje / 100. El porcentaje se configura por empresa en el **Maestro de empresas** (no está en el código); sin tarifa no se calcula.
- **Retenciones 2365 y 2368** del período: por cuenta, con base, tarifa, retención causada (créditos), débitos (cierres/pagos/devoluciones), neto y desglose jurídica/natural. Las cuentas de autorretención y de pago se muestran pero no suman en el total de retenciones practicadas.
- Importes con `Decimal` / centavos enteros; persistencia en SQLite; sin duplicados al reprocesar (clave = empresa + `id`).
- Informe por cuenta con total, rastreo a movimientos originales, revisión de cuentas negativas, controles y exportación a Excel.

## Flujo
1. **Cargar auxiliar**: al procesar el archivo la app detecta las empresas que trae y pasa sola al paso 2.
2. **Tarifas por empresa**: pide la tarifa de autorretención (%) de cada empresa detectada (queda guardada en el *maestro de empresas* para las próximas cargas) y los **NIT excluidos del cálculo** (por defecto `800197268`). Guardar → pasa a la liquidación.
3. **Liquidación**: base, % y autorretención, retenciones 2365/2368, controles y exportación a Excel.

**NIT excluidos:** los movimientos cuyo tercero (`identificacion`) tenga un NIT excluido no entran a la base ni a las retenciones; se listan aparte (pantalla y hoja «Excluidos por NIT» del Excel) para trazabilidad.

## Ejecutar en local

```bash
python -m venv .venv && source .venv/bin/activate   # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
python -m pytest tests          # pruebas del cálculo
```

## Publicar en Streamlit Community Cloud

1. Suba esta carpeta a un repositorio de GitHub (`git init`, `git add .`, `git commit`, `git push`).
2. En <https://share.streamlit.io> → **Create app** → elija el repositorio, rama `main`, archivo principal `app.py` → **Deploy**.

### Advertencias importantes de la publicación

- **Persistencia:** en Streamlit Community Cloud el disco es efímero; la base SQLite (`data/autorretencion.db`) **se pierde cuando la app se reinicia o se duerme**. Para uso real, despliegue en un servidor con disco persistente (Docker/Render/Railway/VM) o cambie `db.py` a una base externa (p. ej. Postgres/Supabase). Mientras tanto, exporte a Excel cada liquidación y vuelva a cargar el maestro de empresas tras cada reinicio.
- **Privacidad:** el auxiliar contable se sube al servidor de la app. Use un repositorio **privado**, restrinja el acceso a la app (Settings → Sharing) y **quite `tests/auxiliar-ejemplo.csv`** si el repositorio fuera público.
- **Sin clave de edición:** cualquiera que abra la app puede cambiar el maestro de empresas; restrinja el acceso desde *Settings → Sharing*.

## Convención de signo
En el auxiliar el ingreso (crédito) viene negativo. Se suma `valor` tal cual y se **presenta invertido** (ingresos positivos); la casilla de la pestaña Liquidación permite ver el signo original.

> Liquidación preliminar sujeta a revisión. La parametrización de la tarifa no sustituye la validación de las reglas tributarias aplicables a la empresa y al período.
