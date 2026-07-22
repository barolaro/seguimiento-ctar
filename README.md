# Seguimiento CTAR — SSMOC

Plataforma Streamlit para registrar y consultar el avance de solicitudes enviadas por el Hospital al CTAR.

## Perfiles

- **Hospital:** acceso de solo lectura.
- **Administrador CTAR:** registra solicitudes y actualiza estados mediante una clave privada configurada en Streamlit.

## Ejecutar localmente

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Subir a GitHub

1. Crea un repositorio nuevo en GitHub, por ejemplo `seguimiento-ctar`.
2. Descomprime este proyecto y abre una terminal dentro de la carpeta.
3. Ejecuta:

```bash
git init
git add .
git commit -m "Primera versión Seguimiento CTAR"
git branch -M main
git remote add origin https://github.com/TU-USUARIO/seguimiento-ctar.git
git push -u origin main
```

## Publicar en Streamlit Community Cloud

1. Ingresa a https://share.streamlit.io con tu cuenta de GitHub.
2. Selecciona **Create app**.
3. Elige el repositorio, la rama `main` y el archivo `app.py`.
4. En **Advanced settings > Secrets**, agrega:

```toml
ADMIN_PASSWORD = "CAMBIAR-POR-UNA-CLAVE-SEGURA"

[connections.gsheets]
spreadsheet = "URL-DE-TU-GOOGLE-SHEET"
worksheet = "solicitudes"
type = "service_account"
# Completar debajo con los datos del archivo JSON de Google Cloud.
```

5. Presiona **Deploy**.

## Almacenamiento permanente con Google Sheets

1. Crea una planilla privada de Google Sheets.
2. Cambia el nombre de la primera pestaña a `solicitudes`.
3. En Google Cloud habilita **Google Sheets API** y **Google Drive API**.
4. Crea una cuenta de servicio y descarga su archivo de credenciales JSON.
5. Comparte la planilla con el correo `client_email` de esa cuenta como **Editor**.
6. Crea una carpeta privada en Google Drive para los adjuntos y compártela como **Editor** con el mismo correo `client_email`.
7. Copia el identificador de la carpeta (la parte de la URL ubicada después de `folders/`) y agrégalo a los Secrets:

```toml
DRIVE_FOLDER_ID = "IDENTIFICADOR-DE-LA-CARPETA"
```

8. Copia en los Secrets de Streamlit la configuración indicada en `.streamlit/secrets.toml.example`, reemplazando los valores de ejemplo por los del JSON.

Las credenciales deben guardarse solamente en los Secrets privados de Streamlit y nunca en GitHub.

Con esta configuración, las solicitudes permanecen en Google Sheets y los adjuntos en Google Drive aunque Streamlit se duerma, reinicie o vuelva a desplegarse. Si Google Sheets no está configurado, la aplicación utiliza temporalmente `data/solicitudes.json`.

El Administrador puede avanzar cada solicitud a la etapa siguiente, realizar un cambio manual de estado y eliminar registros con confirmación.

## Colores institucionales

- Azul SSMOC: `#006FB3`
- Azul oscuro: `#074C77`
- Rojo institucional: `#E72F3B`
- Blanco: `#FFFFFF`
