# Seguimiento CTAR — SSMOC

Plataforma Streamlit para registrar y consultar el avance de solicitudes enviadas por el Hospital al CTAR.

## Perfiles

- **Hospital:** acceso de solo lectura.
- **Administrador CTAR:** registra solicitudes y actualiza estados. Clave inicial: `265727`.

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
ADMIN_PASSWORD = "265727"
```

5. Presiona **Deploy**.

## Importante sobre los datos

La versión incluida guarda la información en `data/solicitudes.json`. En Streamlit Community Cloud, los archivos escritos durante la ejecución pueden perderse cuando la aplicación se reinicia. Para uso institucional permanente, se recomienda conectar una base de datos externa como Supabase, PostgreSQL o Google Sheets.

## Colores institucionales

- Azul SSMOC: `#006FB3`
- Azul oscuro: `#074C77`
- Rojo institucional: `#E72F3B`
- Blanco: `#FFFFFF`
