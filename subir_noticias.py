# -*- coding: utf-8 -*-
"""
Sube las noticias de tu base LOCAL (Base de datos/monitoreo.db) al web service.
Solo libreria estandar: no hay que instalar nada y no toca GitHub.

Uso (desde esta carpeta):
  python subir_noticias.py --url https://TU-SERVICIO.onrender.com --key TU_API_KEY
  python subir_noticias.py --url ... --key ... --dias 30     (por defecto: 7)
  python subir_noticias.py --url ... --key ... --todo        (toda la base)
"""
import argparse
import json
import sqlite3
import sys
import urllib.request
from datetime import datetime, timedelta
from pathlib import Path

BD_LOCAL = Path(__file__).parent.parent / "Base de datos" / "monitoreo.db"
LOTE = 300


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--url", required=True, help="https://tu-servicio.onrender.com")
    ap.add_argument("--key", required=True, help="API key (la misma de Render)")
    ap.add_argument("--dias", type=int, default=7, help="subir noticias de los ultimos N dias")
    ap.add_argument("--todo", action="store_true", help="subir toda la base")
    ap.add_argument("--bd", default=str(BD_LOCAL), help="ruta de monitoreo.db")
    args = ap.parse_args()

    con = sqlite3.connect(args.bd)
    sql = "SELECT analista, fuente, titular, link, fecha, cuerpo FROM noticias"
    params = ()
    if not args.todo:
        corte = (datetime.now() - timedelta(days=args.dias)).strftime("%Y-%m-%d")
        sql += " WHERE fecha >= ? OR cargado_en >= ?"
        params = (corte, corte)
    filas = con.execute(sql, params).fetchall()
    con.close()
    print(f"{len(filas)} noticias para subir a {args.url}")

    claves = ["analista", "fuente", "titular", "link", "fecha", "cuerpo"]
    total_nuevas = total_mejoradas = 0
    for i in range(0, len(filas), LOTE):
        lote = [dict(zip(claves, f)) for f in filas[i:i + LOTE]]
        datos = json.dumps(lote).encode("utf-8")
        req = urllib.request.Request(
            args.url.rstrip("/") + "/cargar", data=datos, method="POST",
            headers={"Content-Type": "application/json", "X-API-Key": args.key})
        try:
            with urllib.request.urlopen(req, timeout=120) as r:
                resp = json.loads(r.read())
            total_nuevas += resp.get("nuevas", 0)
            total_mejoradas += resp.get("mejoradas", 0)
            print(f"  lote {i//LOTE + 1}: +{resp.get('nuevas',0)} nuevas, "
                  f"~{resp.get('mejoradas',0)} mejoradas (total alla: {resp.get('total_en_base','?')})")
        except Exception as e:
            print(f"  ERROR en lote {i//LOTE + 1}: {e}")
            sys.exit(1)
    print(f"Listo: {total_nuevas} nuevas, {total_mejoradas} mejoradas.")


if __name__ == "__main__":
    main()
