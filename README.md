# Contabilidad Chancay

Sistema local de conciliación contable. Permite descargar reportes de facturación de Xafiro y Neo por rango de fechas.

## Requisitos

- Windows con Python 3.12.3.
- Credenciales de Xafiro y Neo en `.env`.

## Ejecutar

El entorno `venv_ofi` ya queda creado y tiene las dependencias instaladas. Desde la carpeta del proyecto ejecuta:

```powershell
py app.py
```

El archivo `app.py` dirige el lanzador de Windows al Python 3.12.3 del entorno. Como alternativa directa:

```powershell
.\venv_ofi\Scripts\python.exe app.py
```

Luego abre `http://127.0.0.1:5050`.

En desarrollo, los cambios en `templates/index.html` se reflejan al recargar el navegador. Si cambias `frontend/app.ts` o `static/styles.css`, recarga la página; para TypeScript conviene dejar corriendo:

```powershell
cmd /c npm run watch
```

## Solución de problemas

Si al exportar aparece un mensaje de conexión con Xafiro, revisa que la PC pueda abrir `https://hotel.xafiro.net` y que firewall, antivirus, VPN o proxy permitan a Python conectarse a internet. La app muestra un mensaje específico cuando Windows bloquea el socket de Python.

## Frontend TypeScript

El navegador recibe `static/app.js`, compilado desde `frontend/app.ts`:

```powershell
npm install
npm run build
```

## Variables de entorno

Consulta `.env.example`. El archivo `.env` real está excluido de Git y nunca se envía al navegador.
