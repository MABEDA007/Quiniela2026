# Publicar Dashboard Quiniela Mundial 2026 en Render

## Archivos listos

- `render.yaml`: configuracion de Render Blueprint.
- `quiniela_dashboard/requirements.txt`: dependencia Python.
- `quiniela_dashboard/app.py`: servidor web.
- `quiniela_dashboard/static/`: interfaz responsiva para desktop y celular.

## Pasos

1. Sube este proyecto a GitHub.
2. En Render, elige `New` -> `Blueprint`.
3. Conecta el repositorio.
4. Render detectara `render.yaml`.
5. Ingresa los secretos solicitados:
   - `API_FOOTBALL_KEY`: llave de API-Football.
   - `ADMIN_TOKEN`: clave privada para subir archivos Excel.
6. Espera el deploy.
7. Abre la URL publica `https://...onrender.com`.

## Despues del deploy

1. Entra a la URL publica.
2. En el panel `Admin`, escribe `ADMIN_TOKEN`.
3. Sube cada quiniela `.xlsx`.
4. Comparte la URL con los participantes.

## Seguridad

Usa un `ADMIN_TOKEN` largo, por ejemplo una frase o valor aleatorio de 24+ caracteres. La URL publica es para ver el dashboard; el token es solo para administrar archivos.
