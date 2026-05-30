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
   - `ADMIN_TOKEN`: opcional en Render gratis. Sirve para pruebas de subida, pero no como almacenamiento permanente.
6. Espera el deploy.
7. Abre la URL publica `https://...onrender.com`.

## Despues del deploy

1. Entra a la URL publica.
2. Verifica que aparezcan los participantes cargados desde `quiniela_dashboard/data/quinielas.csv`.
3. Comparte la URL con los participantes.

## Actualizar quinielas en Render gratis

1. Guarda los `.xlsx` nuevos o corregidos en `Quin2026`.
2. Ejecuta:

```powershell
& "C:\Users\eduar\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe" .\quiniela_dashboard\export_excel_to_csv.py
```

3. Sube a GitHub el archivo actualizado `quiniela_dashboard/data/quinielas.csv`.
4. Render redepliega y carga el CSV desde el inicio.

## Seguridad

Si usas `ADMIN_TOKEN`, que sea largo, por ejemplo una frase o valor aleatorio de 24+ caracteres. La URL publica es para ver el dashboard.
