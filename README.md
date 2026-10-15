# Autorretención sobre ingresos

Aplicación web (sitio estático, sin servidor) que calcula la **liquidación preliminar de autorretención** a partir del auxiliar contable.

- Base = suma neta de la columna `valor` de las cuentas cuyo código empieza por **4** (empresa y período elegidos). No usa `debito`, `credito`, `saldoanterior` ni `saldoactual`.
- Autorretención = Base × Porcentaje / 100. El porcentaje **no está en el código**: se parametriza por empresa y vigencia.
- Todo se procesa en el navegador (IndexedDB); los datos contables no salen del equipo.

## Uso

1. **Auxiliar**: cargue el CSV/Excel (columnas obligatorias: `empresa, fecha, codigocuenta, nombrecuenta, valor`). Recargar el mismo archivo no duplica movimientos (clave = empresa + `id`).
2. **Configuración**: desbloquee la edición (la primera vez crea una clave) y registre tarifa, vigencia desde/hasta, estado, referencia normativa y observaciones. Se advierte si dos tarifas activas de una empresa se superponen; en ese caso (o si ninguna tarifa cubre todo el período) **no se calcula**.
3. **Liquidación**: elija empresa y período. Verá base, % y autorretención por separado, el detalle por cuenta, el rastreo a los movimientos originales, la revisión de cuentas negativas, los controles y el botón **Exportar a Excel** (Resumen, Detalle por cuenta, Movimientos, Cuentas negativas, Validaciones).

Convención de signo: en el auxiliar el ingreso (crédito) viene negativo. Por defecto se suma `valor` tal cual y se **presenta invertido** para que el ingreso aparezca positivo; la casilla de la pantalla de liquidación permite ver el signo original.

> Liquidación preliminar sujeta a revisión. La parametrización de la tarifa no sustituye la validación de las reglas tributarias aplicables.
> La clave de edición solo protege la pantalla en el navegador; no es seguridad de servidor. Los datos y tarifas viven en el navegador de cada usuario: use *Exportar copia* para respaldar las tarifas.

## Publicar

```bash
git init && git add . && git commit -m "Autorretención: versión inicial"
git branch -M main
git remote add origin https://github.com/<usuario>/<repositorio>.git
git push -u origin main
```

**Netlify**: *Add new site → Import an existing project → GitHub →* elija el repositorio. Build command: vacío; Publish directory: `.` (ya está en `netlify.toml`). Cada `git push` redespliega.

## Pruebas del cálculo

```bash
node tests/core.test.mjs
```
`tests/auxiliar-ejemplo.csv` es el auxiliar de ejemplo (ConfianzaGlobal, sept. 2026). Contiene datos contables: use un repositorio **privado** o quite ese archivo antes de publicarlo.

Librería de terceros incluida: SheetJS (`vendor/xlsx.full.min.js`, Apache-2.0).
