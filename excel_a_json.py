# -*- coding: utf-8 -*-
"""
Convierte libros de Excel (.xlsx) generados con json_a_excel.py de vuelta a
archivos RIPS en formato JSON. Es el proceso inverso de json_a_excel.py.

Uso: dejar este script (o el .exe generado con PyInstaller) en la carpeta que
contiene los .xlsx y ejecutarlo. Por cada "X.xlsx" se crea "X_convertido.json".
(Si el Excel se llama "X_convertido.xlsx", el JSON se llama "X_convertido.json".)

Hojas esperadas: Transaccion, Usuarios y una por servicio (Consultas,
Medicamentos, Procedimientos, Urgencias, Hospitalizacion, RecienNacidos,
OtrosServicios). Cada fila de servicio se asigna a su usuario mediante
(UstipoDocumentoIdentificacion, numDocumentoIdentificacionUsuario).

Requisito: pip install openpyxl
"""
import json
import os
import sys
import traceback
from datetime import datetime, date, time

from openpyxl import load_workbook

# Hoja -> clave dentro de "servicios" (el orden de las hojas define el orden en el JSON)
CLAVES_SERVICIO = {
    "Consultas": "consultas",
    "Medicamentos": "medicamentos",
    "Procedimientos": "procedimientos",
    "Urgencias": "urgencias",
    "Hospitalizacion": "hospitalizacion",
    "RecienNacidos": "recienNacidos",
    "OtrosServicios": "otrosServicios",
}

CAMPOS_ENCABEZADO = ["numDocumentoIdObligado", "numFactura", "tipoNota", "numNota"]
# Columnas de enlace agregadas por json_a_excel.py (no existen en el JSON original)
ENLACE = ("numDocumentoIdentificacionUsuario", "numFactura", "UstipoDocumentoIdentificacion")


def carpeta_base():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def valor(v):
    """Normaliza el valor de una celda al tipo que tendría en el JSON."""
    if v is None:
        return None
    if isinstance(v, bool):
        return v
    if isinstance(v, float) and v.is_integer():
        return int(v)
    if isinstance(v, datetime):  # por si Excel convirtió una fecha
        return v.strftime("%Y-%m-%d %H:%M") if (v.hour or v.minute) else v.strftime("%Y-%m-%d")
    if isinstance(v, date):
        return v.strftime("%Y-%m-%d")
    if isinstance(v, time):
        return v.strftime("%H:%M")
    if isinstance(v, str):
        if v == "":
            return None
        s = v.strip()
        if s[:1] in "[{" and s[-1:] in "]}":  # listas/objetos guardados como texto
            try:
                return json.loads(s)
            except ValueError:
                pass
    return v


def filas_de(ws):
    """Devuelve (encabezados, [dict por fila]) ignorando filas vacías."""
    it = ws.iter_rows(values_only=True)
    try:
        enc = [c for c in next(it)]
    except StopIteration:
        return [], []
    # recortar columnas sin encabezado
    while enc and enc[-1] is None:
        enc.pop()
    filas = []
    for fila in it:
        fila = fila[: len(enc)]
        if all(c is None or (isinstance(c, str) and c.strip() == "") for c in fila):
            continue
        filas.append({h: valor(c) for h, c in zip(enc, fila) if h is not None})
    return [h for h in enc if h is not None], filas


def convertir(ruta_xlsx, ruta_json):
    wb = load_workbook(ruta_xlsx, data_only=True, read_only=True)
    hojas = {ws.title: ws for ws in wb.worksheets}
    if "Transaccion" not in hojas or "Usuarios" not in hojas:
        raise ValueError("El libro debe tener las hojas 'Transaccion' y 'Usuarios'.")

    # --- encabezado ---
    _, filas_tx = filas_de(hojas["Transaccion"])
    tx = filas_tx[0] if filas_tx else {}
    data = {k: tx.get(k) for k in CAMPOS_ENCABEZADO}

    # --- usuarios ---
    _, filas_us = filas_de(hojas["Usuarios"])
    usuarios = []
    indice = {}  # (tipo, numero) -> usuario
    for f in filas_us:
        u = {k: v for k, v in f.items() if k != "numFactura"}
        u["servicios"] = {}
        usuarios.append(u)
        indice[(u.get("tipoDocumentoIdentificacion"), u.get("numDocumentoIdentificacion"))] = u

    # --- servicios ---
    resumen = {"Usuarios": len(usuarios)}
    for ws in wb.worksheets:
        clave = CLAVES_SERVICIO.get(ws.title)
        if clave is None:
            if ws.title not in ("Transaccion", "Usuarios"):
                print(f"        (hoja ignorada: '{ws.title}')")
            continue
        _, filas = filas_de(ws)
        for f in filas:
            llave = (f.get("UstipoDocumentoIdentificacion"), f.get("numDocumentoIdentificacionUsuario"))
            u = indice.get(llave)
            if u is None:
                raise ValueError(
                    f"Hoja '{ws.title}': no existe en 'Usuarios' el usuario {llave[0]} {llave[1]}."
                )
            u["servicios"].setdefault(clave, []).append({k: v for k, v in f.items() if k not in ENLACE})
        resumen[ws.title] = len(filas)

    data["usuarios"] = usuarios
    with open(ruta_json, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)
    wb.close()
    return resumen


def main():
    carpeta = carpeta_base()
    archivos = sorted(
        f for f in os.listdir(carpeta)
        if f.lower().endswith(".xlsx") and not f.startswith("~$")
    )
    if not archivos:
        print(f"No se encontraron archivos .xlsx en: {carpeta}")
    ok = err = 0
    for nombre in archivos:
        ruta = os.path.join(carpeta, nombre)
        salida = os.path.join(carpeta, os.path.splitext(nombre)[0] + "_convertido.json")
        try:
            resumen = convertir(ruta, salida)
            ok += 1
            print(f"[OK]    {nombre} -> {os.path.basename(salida)}")
            print("        " + ", ".join(f"{h}: {n}" for h, n in resumen.items()))
        except Exception as e:
            err += 1
            print(f"[ERROR] {nombre}: {e}")
            traceback.print_exc()
    print(f"\nTerminado. Convertidos: {ok}  Con error: {err}")
    if getattr(sys, "frozen", False):
        input("Presione Enter para cerrar...")


if __name__ == "__main__":
    main()
