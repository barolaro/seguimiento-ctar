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
SUPABASE_URL = "https://TU-PROYECTO.supabase.co"
SUPABASE_KEY = "TU-CLAVE-SERVICE-ROLE-PRIVADA"
```

5. Presiona **Deploy**.

## Base de datos permanente con Supabase

1. Crea un proyecto gratuito en https://supabase.com.
2. Abre **SQL Editor** y ejecuta el contenido de `supabase_setup.sql`.
3. En **Project Settings > API**, copia la URL del proyecto y la clave privada `service_role`.
4. Agrégalas a los Secrets de Streamlit junto con la clave del Administrador.

La clave `service_role` debe guardarse solamente en los Secrets privados de Streamlit y nunca en GitHub.

Con esta configuración, las solicitudes permanecen guardadas aunque Streamlit se duerma, reinicie o vuelva a desplegarse. Si no se configuran esas dos variables, la aplicación utiliza temporalmente `data/solicitudes.json`.

El Administrador puede avanzar cada solicitud a la etapa siguiente, realizar un cambio manual de estado y eliminar registros con confirmación.

## Colores institucionales

- Azul SSMOC: `#006FB3`
- Azul oscuro: `#074C77`
- Rojo institucional: `#E72F3B`
- Blanco: `#FFFFFF`
