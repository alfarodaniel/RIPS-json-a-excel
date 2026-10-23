# rips-json-a-excel

Convierte archivos **RIPS en formato JSON** (facturación de servicios de salud, Colombia) a un libro de **Excel (.xlsx)** con una hoja por cada sección del archivo. Procesa en lote todos los `.json` que estén en la misma carpeta del script o del ejecutable.

> ⚠️ **Privacidad:** los RIPS contienen datos personales y de salud. Este repositorio solo incluye datos **sintéticos** (`ejemplo_EJ000001.json`). No suba archivos reales; el `.gitignore` bloquea `*.json` y `*.xlsx` salvo los de ejemplo.

## Requisitos

- Python 3.8 o superior
- [openpyxl](https://pypi.org/project/openpyxl/)

```bash
pip install -r requirements.txt
```

## Uso

1. Copie `json_a_excel.py` en la carpeta que contiene los `.json`.
2. Ejecútelo:

   ```bash
   python json_a_excel.py
   ```

3. Por cada `X.json` se genera `X_convertido.xlsx` en la misma carpeta, y se imprime un resumen con el número de filas por hoja.

### Probar con el ejemplo

```bash
python json_a_excel.py
```

Con `ejemplo_EJ000001.json` en la carpeta se genera `ejemplo_EJ000001_convertido.xlsx` (incluido en el repo como referencia del resultado).

### Crear un ejecutable (.exe)

```bash
pip install pyinstaller
pyinstaller --onefile json_a_excel.py
```

El ejecutable queda en `dist/json_a_excel.exe`. Cópielo a la carpeta con los JSON y haga doble clic. Al correr como `.exe` busca los JSON en la carpeta del propio `.exe` y espera un Enter al terminar para que se pueda leer el resumen.

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

### Reglas de conversión

- Si un JSON trae varios usuarios, los servicios de todos se juntan en la misma hoja, identificados por las columnas de enlace.
- Los códigos con ceros a la izquierda (`"04"`, `"05"`, `"11001"`) se conservan como **texto**; los valores numéricos del JSON se conservan como números.
- Los valores `null` quedan como celdas vacías.
- Si aparece un campo no previsto dentro de un servicio, se agrega como columna; si aparece un tipo de servicio nuevo, se crea su hoja con el nombre de la clave.
- Si un valor es una lista u objeto anidado, se guarda como texto JSON en la celda.

## Archivos del repositorio

| Archivo | Descripción |
|---|---|
| `json_a_excel.py` | Script de conversión |
| `ejemplo_EJ000001.json` | JSON de entrada sintético (2 usuarios, datos ficticios) |
| `ejemplo_EJ000001_convertido.xlsx` | Resultado esperado para el ejemplo |
| `requirements.txt` | Dependencias |
| `.gitignore` | Evita subir datos reales por accidente |

## Datos sintéticos

Todos los identificadores del ejemplo son ficticios: NIT `900000001`, prestador `990000000001`, documentos `1000000001`…, factura `EJ000001`, códigos de tecnología `XX0001`/`00000000-01` y nombres genéricos ("MEDICAMENTO EJEMPLO A"). Los códigos VIDA son UUID aleatorios. Los códigos de catálogo (CUPS, CIE-10, modalidades, etc.) son valores válidos de forma pero no corresponden a ninguna atención real.

## Limitaciones

- No valida el contenido contra la normativa de RIPS (Resolución 2275 de 2023 y anexos técnicos); solo transforma la estructura.
- Probado con facturas que incluyen consultas, medicamentos, procedimientos, hospitalización, recién nacidos y otros servicios.

## Licencia

Elija la licencia que prefiera (por ejemplo MIT) y agregue un archivo `LICENSE`.
