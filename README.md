# RIPS-json-excel

Conversión en **ambos sentidos** de archivos **RIPS** (facturación de servicios de salud, Colombia) entre **JSON** y **Excel (.xlsx)**, con una hoja por cada sección del archivo. Ambos scripts procesan en lote todos los archivos que estén en la misma carpeta del script o del ejecutable.

| Script | Entrada | Salida |
|---|---|---|
| `json_a_excel.py` | `X.json` | `X_convertido.xlsx` |
| `excel_a_json.py` | `X.xlsx` | `X_convertido.json` |

> ⚠️ **Privacidad:** los RIPS contienen datos personales y de salud. Este repositorio solo incluye datos **sintéticos** (`ejemplo_EJ000001.json`). No suba archivos reales; el `.gitignore` bloquea `*.json` y `*.xlsx` salvo los de ejemplo.

## Requisitos

- Python 3.8 o superior
- [openpyxl](https://pypi.org/project/openpyxl/)

```bash
pip install -r requirements.txt
```

## Uso

Copie el script que necesite en la carpeta con los archivos y ejecútelo.

**JSON → Excel**

```bash
python json_a_excel.py
```

Por cada `X.json` genera `X_convertido.xlsx` y muestra el número de filas por hoja.

**Excel → JSON**

```bash
python excel_a_json.py
```

Por cada `X.xlsx` genera `X_convertido.json`. Se ignoran los archivos temporales de Excel (`~$...`).

> Un Excel generado por `json_a_excel.py` se llama `X_convertido.xlsx`; al volver a JSON daría `X_convertido_convertido.json`. Renómbrelo a `X.xlsx` antes de convertir si quiere evitarlo.

### Probar con el ejemplo

Con `ejemplo_EJ000001.json` en la carpeta:

```bash
python json_a_excel.py          # crea ejemplo_EJ000001_convertido.xlsx
```

Para probar el sentido contrario, renombre ese Excel a `ejemplo.xlsx` y ejecute `python excel_a_json.py`: obtendrá `ejemplo_convertido.json`, idéntico al JSON de ejemplo.

### Crear ejecutables (.exe)

```bash
pip install pyinstaller
pyinstaller --onefile json_a_excel.py
pyinstaller --onefile excel_a_json.py
```

Los ejecutables quedan en `dist/`. Cópielos a la carpeta con los archivos y haga doble clic. Al correr como `.exe` buscan los archivos en la carpeta del propio `.exe` y esperan un Enter al terminar para que se pueda leer el resumen.

## Estructura del Excel generado

| Hoja | Contenido |
|---|---|
| `Transaccion` | Encabezado de la factura (`numDocumentoIdObligado`, `numFactura`, `tipoNota`, `numNota`) y documento del usuario (`UstipoDocumentoIdentificacion`, `numDocumentoIdentificacion`). Una fila por usuario. |
| `Usuarios` | Datos de cada usuario, con `numFactura` como primera columna. No incluye el bloque `servicios`. |
| `Consultas`, `Medicamentos`, `Procedimientos`, `Urgencias`, `Hospitalizacion`, `RecienNacidos`, `OtrosServicios` | Una hoja por cada tipo de servicio presente en el JSON. Se omiten los tipos que no vienen. |

Cada hoja de servicios agrega al final tres columnas que enlazan la fila con su usuario y factura:

- `numDocumentoIdentificacionUsuario`
- `numFactura`
- `UstipoDocumentoIdentificacion`

### Reglas de conversión (JSON → Excel)

- Si un JSON trae varios usuarios, los servicios de todos se juntan en la misma hoja, identificados por las columnas de enlace.
- Los códigos con ceros a la izquierda (`"04"`, `"05"`, `"11001"`) se conservan como **texto**; los valores numéricos del JSON se conservan como números.
- Los valores `null` quedan como celdas vacías.
- Si aparece un campo no previsto dentro de un servicio, se agrega como columna; si aparece un tipo de servicio nuevo, se crea su hoja con el nombre de la clave.
- Si un valor es una lista u objeto anidado, se guarda como texto JSON en la celda.

### Reglas de conversión (Excel → JSON)

- Los datos de `Transaccion` forman el encabezado y los de `Usuarios` la lista `usuarios` (sin la columna `numFactura`).
- Cada fila de una hoja de servicios se asigna a su usuario por `UstipoDocumentoIdentificacion` + `numDocumentoIdentificacionUsuario`. Esas tres columnas de enlace no se escriben en el JSON.
- Las celdas vacías pasan a `null`; los números guardados como decimales enteros (`27200.0`) vuelven a ser enteros.
- Si Excel convirtió una fecha a formato de fecha, se escribe de nuevo como texto (`2026-09-30 11:37`).
- Las hojas ausentes o vacías no generan su clave en el JSON.
- Si una fila de servicio apunta a un usuario que no existe en `Usuarios`, el script se detiene con un error.
- **Códigos con ceros a la izquierda** (`04`, `11001`): si edita el Excel a mano, deben seguir como texto. Si Excel los convierte en número, el JSON los tendrá sin ceros y dejará de ser válido.

## Archivos del repositorio

| Archivo | Descripción |
|---|---|
| `json_a_excel.py` | Conversión JSON → Excel |
| `excel_a_json.py` | Conversión Excel → JSON |
| `ejemplo_EJ000001.json` | JSON de entrada sintético (2 usuarios, datos ficticios) |
| `requirements.txt` | Dependencias |
| `.gitignore` | Evita subir datos reales por accidente |

## Datos sintéticos

Todos los identificadores del ejemplo son ficticios: NIT `900000001`, prestador `990000000001`, documentos `1000000001`…, factura `EJ000001`, códigos de tecnología `XX0001`/`00000000-01` y nombres genéricos ("MEDICAMENTO EJEMPLO A"). Los códigos VIDA son UUID aleatorios. Los códigos de catálogo (CUPS, CIE-10, modalidades, etc.) son valores válidos de forma pero no corresponden a ninguna atención real.

## Limitaciones

- No valida el contenido contra la normativa de RIPS (Resolución 2275 de 2023 y anexos técnicos); solo transforma la estructura.
- Probado con facturas que incluyen consultas, medicamentos, procedimientos, hospitalización, recién nacidos y otros servicios. La ida y vuelta (JSON → Excel → JSON) reproduce el archivo original con los mismos valores y orden de claves.
- `excel_a_json.py` está pensado para libros generados por `json_a_excel.py`; un Excel con otra estructura no se interpreta.

## Licencia

MIT, ver `LICENSE`.