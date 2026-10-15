# Autorretención sobre ingresos (Python + Streamlit)

Calcula la **liquidación preliminar de autorretención** a partir del auxiliar contable.

- **Base** = suma neta (con signo) de la columna `valor` de las cuentas cuyo código empieza por **4**, por empresa y período. No usa `debito`, `credito`, `saldoanterior` ni `saldoactual`.
- **Autorretención** = Base × Porcentaje / 100. El porcentaje se parametriza por empresa y vigencia (no está en el código).
- Importes con `Decimal` / centavos enteros; persistencia en SQLite; sin duplicados al reprocesar (clave = empresa + `id`).
- Informe por cuenta con total, rastreo a movimientos originales, revisión de cuentas negativas, controles y exportación a Excel.

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
3. (Recomendado) En *Settings → Secrets / Environment* defina `AUTORRETENCION_PIN` con la clave de edición de tarifas; si no, la app pide crearla la primera vez.

### Advertencias importantes de la publicación

- **Persistencia:** en Streamlit Community Cloud el disco es efímero; la base SQLite (`data/autorretencion.db`) **se pierde cuando la app se reinicia o se duerme**. Para uso real, despliegue en un servidor con disco persistente (Docker/Render/Railway/VM) o cambie `db.py` a una base externa (p. ej. Postgres/Supabase). Mientras tanto, exporte a Excel cada liquidación y guarde las tarifas.
- **Privacidad:** el auxiliar contable se sube al servidor de la app. Use un repositorio **privado**, restrinja el acceso a la app (Settings → Sharing) y **quite `tests/auxiliar-ejemplo.csv`** si el repositorio fuera público.
- **Clave de edición:** protege la pantalla de configuración; no sustituye el control de acceso a la app.

## Convención de signo
En el auxiliar el ingreso (crédito) viene negativo. Se suma `valor` tal cual y se **presenta invertido** (ingresos positivos); la casilla de la pestaña Liquidación permite ver el signo original.

> Liquidación preliminar sujeta a revisión. La parametrización de la tarifa no sustituye la validación de las reglas tributarias aplicables a la empresa y al período.
