# -*- coding: utf-8 -*-
"""
Web service del monitoreo de noticias.

GET  /noticias?analista=ANDREA,LUISA&desde=2026-06-01&hasta=2026-06-10&tipo=excel|json
     - Sin fechas: ultimos 2 dias. Sin tipo: excel (descarga directa).
     - formato=consolidado -> una hoja por fuente (formato de la Llamada a la API).
GET  /            formulario web
GET  /salud       estado y total de noticias
POST /cargar      carga de noticias (JSON, requiere header X-API-Key)
"""
import io
import os
import sqlite3
from datetime import datetime, timedelta

from flask import Flask, request, jsonify, send_file, Response
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

app = Flask(__name__)

RUTA_BD = os.environ.get("RUTA_BD", os.path.join(os.path.dirname(__file__), "monitoreo.db"))
API_KEY = os.environ.get("API_KEY", "")   # definir en Render

ESQUEMA = """
CREATE TABLE IF NOT EXISTS noticias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    analista TEXT NOT NULL,
    fuente   TEXT NOT NULL,
    titular  TEXT,
    link     TEXT NOT NULL UNIQUE,
    fecha    TEXT,
    cuerpo   TEXT,
    cargado_en TEXT DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_fecha ON noticias(fecha);
CREATE INDEX IF NOT EXISTS idx_analista ON noticias(analista);
"""


def conectar():
    con = sqlite3.connect(RUTA_BD, timeout=30)
    con.executescript(ESQUEMA)
    return con


# ----------------------------- consultas ------------------------------------
def filtrar(analista, desde, hasta):
    cond, params = [], []
    if analista:
        lista = [a.strip().upper() for a in analista.split(",") if a.strip()]
        cond.append("UPPER(analista) IN (%s)" % ",".join("?" * len(lista)))
        params += lista
    cond.append("fecha >= ?"); params.append(desde + " 00:00:00")
    cond.append("fecha <= ?"); params.append(hasta + " 23:59:59")
    sql = ("SELECT analista, fuente, titular, link, fecha, cuerpo FROM noticias "
           "WHERE " + " AND ".join(cond) + " ORDER BY analista, fecha DESC")
    return sql, params


def encabezado(ws, columnas, anchos):
    ws.append(columnas)
    azul = PatternFill("solid", start_color="1F3864")
    for c in range(1, len(columnas) + 1):
        cel = ws.cell(row=1, column=c)
        cel.font = Font(bold=True, color="FFFFFF", name="Arial")
        cel.fill = azul
    for col, ancho in zip("ABCDEFG", anchos):
        ws.column_dimensions[col].width = ancho
    ws.freeze_panes = "A2"


def excel_plano(filas):
    wb = Workbook(); ws = wb.active; ws.title = "Noticias"
    encabezado(ws, ["ANALISTA", "FUENTE", "TITULAR", "LINK", "FECHA", "CUERPO DE LA NOTICIA"],
               (12, 18, 60, 50, 20, 120))
    for f in filas:
        ws.append(list(f))
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf


def excel_consolidado(filas):
    wb = Workbook(); wb.remove(wb.active)
    for analista, fuente, titular, link, fecha, cuerpo in filas:
        hoja = str(fuente or "sin_fuente")[:31]
        if hoja not in [w.title for w in wb.worksheets]:
            ws = wb.create_sheet(hoja)
            encabezado(ws, ["TITULAR", "LINK", "FECHA", "CUERPO DE LA NOTICIA"],
                       (60, 50, 22, 120))
        wb[hoja].append([titular, link, fecha, cuerpo])
    if not wb.worksheets:
        ws = wb.create_sheet("Noticias")
        encabezado(ws, ["TITULAR", "LINK", "FECHA", "CUERPO DE LA NOTICIA"], (60, 50, 22, 120))
    buf = io.BytesIO(); wb.save(buf); buf.seek(0)
    return buf


@app.get("/noticias")
def noticias():
    hoy = datetime.utcnow().date()
    desde = request.args.get("desde") or str(hoy - timedelta(days=2))
    hasta = request.args.get("hasta") or str(hoy)
    tipo = (request.args.get("tipo") or "excel").lower()
    formato = (request.args.get("formato") or "plano").lower()

    sql, params = filtrar(request.args.get("analista", ""), desde, hasta)
    con = conectar()
    filas = con.execute(sql, params).fetchall()
    con.close()

    if tipo == "json":
        claves = ["analista", "fuente", "titular", "link", "fecha", "cuerpo"]
        return jsonify({"total": len(filas), "desde": desde, "hasta": hasta,
                        "noticias": [dict(zip(claves, f)) for f in filas]})

    buf = excel_consolidado(filas) if formato == "consolidado" else excel_plano(filas)
    nombre = f"Noticias_{desde}_a_{hasta}.xlsx"
    return send_file(buf, as_attachment=True, download_name=nombre,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


# ----------------------------- carga remota ---------------------------------
@app.post("/cargar")
def cargar():
    if not API_KEY or request.headers.get("X-API-Key") != API_KEY:
        return jsonify({"error": "API key invalida"}), 401
    datos = request.get_json(silent=True)
    if not isinstance(datos, list):
        return jsonify({"error": "se espera una lista JSON de noticias"}), 400

    con = conectar(); cur = con.cursor()
    nuevas = mejoradas = 0
    for n in datos:
        link = str(n.get("link", "")).strip()
        if not link.lower().startswith("http"):
            continue
        cuerpo = str(n.get("cuerpo") or "")
        existe = cur.execute("SELECT LENGTH(cuerpo) FROM noticias WHERE link=?", (link,)).fetchone()
        if existe is None:
            cur.execute("INSERT INTO noticias (analista, fuente, titular, link, fecha, cuerpo) "
                        "VALUES (?,?,?,?,?,?)",
                        (str(n.get("analista") or ""), str(n.get("fuente") or ""),
                         str(n.get("titular") or ""), link, n.get("fecha"), cuerpo))
            nuevas += 1
        elif len(cuerpo) > (existe[0] or 0):
            cur.execute("UPDATE noticias SET cuerpo=?, fecha=COALESCE(fecha,?) WHERE link=?",
                        (cuerpo, n.get("fecha"), link))
            mejoradas += 1
    con.commit()
    total = cur.execute("SELECT COUNT(*) FROM noticias").fetchone()[0]
    con.close()
    return jsonify({"recibidas": len(datos), "nuevas": nuevas,
                    "mejoradas": mejoradas, "total_en_base": total})


# ----------------------------- estado y portada -----------------------------
@app.get("/salud")
def salud():
    con = conectar()
    total, ultima = con.execute("SELECT COUNT(*), MAX(cargado_en) FROM noticias").fetchone()
    con.close()
    return jsonify({"ok": True, "total_noticias": total, "ultima_carga": ultima})


@app.get("/")
def portada():
    con = conectar()
    total, = con.execute("SELECT COUNT(*) FROM noticias").fetchone()
    con.close()
    return Response(f"""<!doctype html><html lang="es"><head><meta charset="utf-8">
<title>Monitoreo de noticias</title>
<style>body{{font-family:Segoe UI,Arial;background:#f4f6fb;padding:40px}}
.caja{{max-width:540px;margin:auto;background:#fff;border-radius:12px;padding:28px 32px;
box-shadow:0 2px 10px rgba(31,56,100,.15)}}h1{{color:#1F3864;font-size:21px;margin-top:0}}
label{{display:block;margin:10px 0 4px;font-weight:600;color:#2E4057}}
input,select{{width:100%;padding:8px;border:1px solid #c5cede;border-radius:6px;box-sizing:border-box}}
button{{margin-top:16px;background:#1F3864;color:#fff;border:0;border-radius:6px;
padding:10px 22px;cursor:pointer}}.nota{{color:#667;font-size:12px;margin-top:12px}}</style>
</head><body><div class="caja"><h1>Monitoreo de noticias</h1>
<form action="/noticias" method="get">
<label>Analistas (vacío = todos)</label><input name="analista" placeholder="ANDREA,LUISA">
<label>Desde (vacío = hace 2 días)</label><input type="date" name="desde">
<label>Hasta (vacío = hoy)</label><input type="date" name="hasta">
<label>Tipo</label><select name="tipo"><option value="excel">Excel</option><option value="json">JSON</option></select>
<label>Formato del Excel</label><select name="formato"><option value="plano">Plano</option>
<option value="consolidado">Consolidado para la API (hoja por fuente)</option></select>
<button type="submit">Descargar</button></form>
<p class="nota">{total:,} noticias en la base.</p></div></body></html>""",
        mimetype="text/html")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)))
