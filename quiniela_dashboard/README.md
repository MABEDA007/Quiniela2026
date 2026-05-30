# Dashboard Quiniela Mundial 2026

Aplicacion web para leer quinielas `.xlsx`, consultar resultados de API-Football y calcular la tabla de posiciones.

## Uso local

Desde PowerShell:

```powershell
$env:API_FOOTBALL_KEY="TU_LLAVE_DE_API_FOOTBALL"
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
3. Configura estas variables secretas:
   - `API_FOOTBALL_KEY`: tu llave de API-Football.
   - `ADMIN_TOKEN`: una clave privada que solo tu conozcas para subir quinielas.
4. Render creara un disco persistente en `/var/data`, donde se guardaran los Excel y el cache de resultados.
5. Al terminar el deploy, Render dara una URL publica. Esa URL se puede abrir desde celular o cualquier computadora.

Si usas `Blueprint`, Render lee `render.yaml`. Si creas el servicio manualmente:

```text
Build command: pip install -r quiniela_dashboard/requirements.txt
Start command: python quiniela_dashboard/app.py --host 0.0.0.0 --port $PORT --quiniela-dir /var/data/quinielas
Health check path: /health
Disk mount path: /var/data
```

## Subir quinielas

Cuando `ADMIN_TOKEN` esta configurado, el dashboard muestra un panel `Admin`. Ingresa el token, selecciona el archivo `.xlsx` y subelo. Cualquier visitante puede ver la tabla, pero solo quien tenga el token puede cargar archivos.

## Formato esperado del Excel

La app detecta archivos `.xlsx`. En la hoja de quiniela lee:

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

La API oficial del Mundial 2026 usa `league=1` y `season=2026`.
