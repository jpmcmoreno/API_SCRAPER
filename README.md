# Web service del monitoreo de noticias

Servicio web que entrega las noticias acumuladas por analista y rango de fechas, en Excel (descarga directa) o JSON. La información se carga desde tu computador por HTTP — **nunca hay que tocar GitHub para actualizar datos**.

## Endpoints

| Endpoint | Qué hace |
|---|---|
| `GET /noticias` | Descarga las noticias. Parámetros: `analista` (uno o varios con coma; vacío = todos), `desde`, `hasta` (YYYY-MM-DD; sin fechas = **últimos 2 días**), `tipo` (`excel` o `json`; sin tipo = **excel**), `formato` (`plano` o `consolidado` = una hoja por fuente, el de la Llamada a la API) |
| `GET /` | Formulario web para descargar sin saberse los parámetros |
| `GET /salud` | Estado y total de noticias |
| `POST /cargar` | Recibe noticias (lo usa `subir_noticias.py`; requiere header `X-API-Key`) |

Ejemplos:
```
https://TU-SERVICIO.onrender.com/noticias                          ← Excel, todos, últimos 2 días
https://TU-SERVICIO.onrender.com/noticias?analista=ANDREA,LUISA&desde=2026-06-01&hasta=2026-06-10
https://TU-SERVICIO.onrender.com/noticias?analista=SANTIAGO&tipo=json
https://TU-SERVICIO.onrender.com/noticias?formato=consolidado
```

## Subir el servicio a Render (una sola vez)

1. **GitHub**: crea un repositorio nuevo (puede ser privado) y sube SOLO el contenido de esta carpeta (`app.py`, `requirements.txt`, `render.yaml`, `subir_noticias.py`, `.gitignore`, `README.md`). Desde esta carpeta:
   ```
   git init
   git add .
   git commit -m "Web service monitoreo"
   git branch -M main
   git remote add origin https://github.com/TU_USUARIO/monitoreo-noticias.git
   git push -u origin main
   ```
2. **Render**: entra a render.com → **New → Web Service** → conecta tu cuenta de GitHub y elige el repositorio.
3. Render detecta `render.yaml` y rellena todo solo. Si lo pide a mano:
   - Runtime: **Python 3** · Build: `pip install -r requirements.txt` · Start: `gunicorn app:app` · Plan: **Free**
4. En **Environment** agrega la variable `API_KEY` con una clave que tú inventes (larga y secreta). Esa misma clave la usarás para subir datos.
5. Deploy. Tu servicio queda en `https://monitoreo-noticias.onrender.com` (o el nombre que elijas).

## Cargar la información desde tu computador

Cada vez que quieras actualizar el servicio (después de una corrida de scrapers + `Cargar_BD.bat`):

```
cd webservice_monitoreo
python subir_noticias.py --url https://TU-SERVICIO.onrender.com --key TU_API_KEY
```

- Por defecto sube los últimos 7 días (`--dias 30` para más, `--todo` para toda la base).
- Es idempotente: no crea duplicados (link único) y si llega un cuerpo más completo, actualiza.
- Lee tu base local `..\Base de datos\monitoreo.db` automáticamente.
- Solo usa librería estándar de Python: no instala nada.

Lo puedes automatizar añadiendo esa línea al final de `EJECUTAR_Y_ACUMULAR.bat` si quieres que cada corrida de 3 horas actualice también el servicio.

## Ojo con el plan Free de Render

- El servicio **se duerme** tras ~15 min sin uso: la primera petición después tarda ~1 minuto en despertar. Las siguientes son instantáneas.
- El disco es **efímero**: en cada deploy o reinicio la base del servicio puede quedar vacía. No pasa nada — tu base local es la fuente de verdad: ejecuta `subir_noticias.py --todo` y en un minuto está repoblada. (Si algún día quieres persistencia real, en Render se le agrega un disco pagado o se migra a su Postgres.)

## Probar en local antes de subir

```
pip install flask openpyxl
set API_KEY=loquesea
python app.py
```
Abre http://localhost:8000 — y en otra terminal: `python subir_noticias.py --url http://localhost:8000 --key loquesea --todo`
