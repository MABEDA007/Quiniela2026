# Dashboard Quiniela Mundial 2026

Aplicacion web para leer quinielas convertidas a CSV, consultar resultados desde el scoreboard publico de ESPN y calcular la tabla de posiciones.

## Flujo recomendado para Render gratis

Render gratis no permite disco persistente. Por eso las quinielas se publican como CSV dentro del paquete:

```powershell
& "C:\Users\eduar\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\quiniela_dashboard\export_excel_to_csv.py
```

Ese comando lee los `.xlsx` de la carpeta `Quin2026` y genera:

```text
quiniela_dashboard\data\quinielas.csv
```

Despues sube los cambios a GitHub y Render desplegara el dashboard leyendo ese CSV desde el arranque.

## Uso local

Desde PowerShell:

```powershell
$env:ADMIN_TOKEN="UNA_CLAVE_PRIVADA_PARA_SUBIR_EXCEL"
.\quiniela_dashboard\start_dashboard.ps1
```

Abre:

```text
http://127.0.0.1:8765
```

Tambien puedes abrir `quiniela_dashboard\start_dashboard.bat` con doble clic. Deja la ventana abierta mientras uses el dashboard.

## Publicar en internet con Render

1. Crea un repositorio GitHub con esta carpeta/proyecto.
2. En Render, crea un nuevo `Blueprint` o `Web Service` usando `render.yaml`.
3. No necesitas configurar llave de API para resultados: la app usa el endpoint publico de ESPN.
4. Al terminar el deploy, Render dara una URL publica. Esa URL se puede abrir desde celular o cualquier computadora.

Si usas `Blueprint`, Render lee `render.yaml`. En `Blueprint Path` escribe exactamente:

```text
render.yaml
```

Si creas el servicio manualmente:

```text
Build command: pip install -r quiniela_dashboard/requirements.txt
Start command: python quiniela_dashboard/app.py --host 0.0.0.0 --port $PORT --quiniela-dir quiniela_dashboard/data
Health check path: /health
```

## Subir quinielas

En Render gratis, no uses el panel `Admin` como almacenamiento permanente porque los archivos subidos se pueden perder al reiniciar el servicio. Para actualizar quinielas, agrega o cambia los Excel localmente, ejecuta `export_excel_to_csv.py`, sube el CSV a GitHub y deja que Render redepliegue.

El panel `Admin` sigue existiendo para pruebas o para un plan con disco persistente, pero el flujo gratis estable es CSV empaquetado.

## Formato esperado del Excel

El script detecta archivos `.xlsx`. En la hoja de quiniela lee:

- Columna E: numero de partido.
- Columna F: fecha.
- Columna G: hora.
- Columna I: equipo local.
- Columna J: goles pronosticados del local.
- Columna L: goles pronosticados del visitante.
- Columna M: equipo visitante.
- Columna O: sede.

Puntuacion:

- 3 puntos por marcador exacto.
- 1 punto por acertar ganador o empate.
- 0 puntos por fallar.

La fuente de resultados es ESPN:

```text
https://site.api.espn.com/apis/site/v2/sports/soccer/fifa.world/scoreboard
```
