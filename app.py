import hashlib
import io
import json
import base64
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload
from docx import Document
from streamlit_gsheets import GSheetsConnection

BASE = Path(__file__).parent
DATA_FILE = BASE / "data" / "solicitudes.json"
UPLOAD_DIR = BASE / "data" / "documentos"
LOGO = BASE / "assets" / "logo_ssmoc.jpg"
ETAPAS = ["Hospital envía", "CTAR revisa", "CTAR acuerda", "Acta se firma", "Proceso finaliza"]
TIPOS_MATERIA = ["Baja", "Reposición", "Adquisición / gradualidad", "Especificaciones técnicas", "Evaluación", "Informativo", "Administrativo", "Tema general CTAR"]
FORMALIZACIONES = ["Borrador interno", "En revisión CTAR", "Acuerdo adoptado", "Acta en firma", "Formalizado"]
COLORES = {"Hospital envía": "#0B6DAA", "CTAR revisa": "#F2A900", "CTAR acuerda": "#7B4FA3", "Acta se firma": "#008C95", "Proceso finaliza": "#198754"}
PROXIMOS = {
    "Hospital envía": "CTAR revisará los antecedentes recibidos",
    "CTAR revisa": "Adoptar acuerdo en sesión CTAR",
    "CTAR acuerda": "Preparar el acta para revisión y firma",
    "Acta se firma": "Completar las firmas pendientes",
    "Proceso finaliza": "Sin acciones pendientes",
}


class ArchivoMemoria:
    def __init__(self, nombre, mime_type, contenido):
        self.name = nombre
        self.type = mime_type
        self._contenido = contenido

    def getvalue(self):
        return self._contenido

st.set_page_config(page_title="Seguimiento CTAR | SSMOC", page_icon="🏥", layout="wide", initial_sidebar_state="expanded")


def css():
    st.markdown("""
    <style>
    :root{--azul:#006FB3;--azul-oscuro:#074C77;--rojo:#E72F3B;--gris:#F3F6F8;--texto:#183247}
    .stApp{background:linear-gradient(180deg,#F5F9FC 0,#FFFFFF 440px)}
    [data-testid="stSidebar"]{background:#F7FAFC;border-right:1px solid #DCE6EC}
    [data-testid="stSidebar"] img{border-radius:8px}
    h1,h2,h3{color:var(--texto);letter-spacing:-.02em}
    .hero{padding:1.4rem 1.6rem;background:linear-gradient(120deg,#074C77,#0875B5);border-radius:16px;color:white;margin-bottom:1rem;box-shadow:0 10px 28px #074C7726}
    .hero h1{color:white;margin:0;font-size:2rem}.hero p{margin:.45rem 0 0;color:#E8F5FC}
    .perfil{display:inline-block;background:#FFFFFF25;border:1px solid #FFFFFF45;border-radius:999px;padding:.3rem .65rem;margin-top:.8rem;font-size:.8rem}
    .pasos{display:grid;grid-template-columns:repeat(5,1fr);gap:.55rem;margin:.8rem 0 1.2rem}.paso{background:white;border:1px solid #DDE7ED;border-radius:11px;padding:.75rem;min-height:88px}.paso b{display:block;color:#123D59;font-size:.86rem}.paso small{color:#718495}.numero{display:inline-grid;place-items:center;width:25px;height:25px;border-radius:50%;background:#E5F2FA;color:#006FB3;font-weight:800;margin-bottom:.45rem}
    .ayuda{background:#EDF7FD;border-left:4px solid #006FB3;padding:.8rem 1rem;border-radius:8px;color:#34576C;margin:.5rem 0 1rem;font-size:.9rem}
    .estado{display:inline-block;color:white;border-radius:999px;padding:.28rem .65rem;font-weight:700;font-size:.78rem}
    .ficha{background:white;border:1px solid #DEE7EC;border-radius:13px;padding:1rem;margin:.55rem 0;box-shadow:0 4px 14px #153D5410}.ficha h4{margin:0;color:#183A50}.meta{color:#718595;font-size:.83rem;margin:.25rem 0}.proximo{background:#F0F8FC;border-radius:8px;padding:.55rem .7rem;color:#164D6B;font-size:.84rem;margin-top:.55rem}
    .solo-lectura{background:#EAF7EF;border:1px solid #BFE2CB;color:#27683D;padding:.65rem .8rem;border-radius:8px;font-size:.85rem}
    div[data-testid="stMetric"]{background:white;border:1px solid #DEE7EC;padding:.8rem 1rem;border-radius:11px;box-shadow:0 3px 10px #153D540D}
    .stButton>button{border-radius:8px;font-weight:700}.stButton>button[kind="primary"]{background:#006FB3;border-color:#006FB3}
    div[data-testid="stDialog"] div[role="dialog"]{position:fixed!important;right:0!important;top:0!important;height:100vh!important;max-height:100vh!important;width:680px!important;max-width:96vw!important;border-radius:0!important;margin:0!important;padding:22px 30px!important;overflow-y:auto!important}
    .admin-eye{font-size:11px;font-weight:800;letter-spacing:.14em;color:#0872BC}.admin-title{font-size:24px;font-weight:700;color:#142B40;line-height:1.25;margin:10px 0 8px}.admin-meta{color:#718595;font-size:12px;margin-bottom:18px}.admin-section{font-size:14px;font-weight:800;color:#183A50;border-top:1px solid #E1E8EE;padding-top:18px;margin:20px 0 10px}.admin-summary{background:#EEF6FB;border-left:3px solid #0872BC;border-radius:6px;padding:13px 15px;color:#294A62;font-size:12px;line-height:1.55;margin:14px 0}
    @media(max-width:900px){.pasos{grid-template-columns:1fr}.paso{min-height:auto}}
    </style>""", unsafe_allow_html=True)


def datos_iniciales():
    return [
        {"id":"CTAR-149-01","tema":"Hervidores industriales (3)","hospital":"Hospital Dr. Félix Bulnes","servicio":"SEDILE","sic":"60.045","inventario":"HFB-SED-118 / 120","motivo":"Baja y reposición por término de vida útil","fecha_ingreso":"2026-07-16","estado":"Acta se firma","ultima_actualizacion":"2026-07-21 - Acta CTAR N.° 149 actualizada","proximo_paso":"Obtener firmas de los integrantes CTAR","observaciones":"Pronunciamientos favorables recibidos por correo.","documentos":[]},
        {"id":"CTAR-140-06","tema":"Ventiladores neonatales (10)","hospital":"Hospital Dr. Félix Bulnes","servicio":"Neonatología","sic":"52.881","inventario":"6 AVEA / 4 VN500","motivo":"Reposición integral por obsolescencia","fecha_ingreso":"2026-06-10","estado":"CTAR revisa","ultima_actualizacion":"2026-07-18 - EETT clínicas recibidas","proximo_paso":"Revisar especificaciones técnicas en sesión CTAR","observaciones":"Preferencia clínica informada: Dräger VN800.","documentos":[]},
        {"id":"CTAR-139-04","tema":"Cámaras de videolaparoscopía","hospital":"Hospital Dr. Félix Bulnes","servicio":"Pabellón","sic":"60.729 / 7.990","inventario":"Por confirmar","motivo":"Equipos no reparables; solicitud de baja","fecha_ingreso":"2026-06-30","estado":"Hospital envía","ultima_actualizacion":"2026-07-17 - Se solicitó completar antecedentes","proximo_paso":"Hospital debe remitir inventarios y EETT","observaciones":"Vinculada al Ord. HFB N.° 763.","documentos":[]},
    ]


def cliente_gsheets():
    try:
        configuracion = st.secrets["connections"]["gsheets"]
        if not configuracion.get("spreadsheet"):
            return None
        return st.connection("gsheets", type=GSheetsConnection)
    except (KeyError, FileNotFoundError):
        return None


def almacenamiento_permanente():
    return cliente_gsheets() is not None


def cliente_drive():
    """Crea un cliente privado de Drive usando la misma cuenta de servicio."""
    try:
        config = dict(st.secrets["connections"]["gsheets"])
        info = {k: config[k] for k in (
            "type", "project_id", "private_key_id", "private_key",
            "client_email", "client_id", "auth_uri", "token_uri",
            "auth_provider_x509_cert_url", "client_x509_cert_url",
        ) if k in config}
        credenciales = service_account.Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/drive"]
        )
        return build("drive", "v3", credentials=credenciales, cache_discovery=False)
    except Exception:
        return None


def carpeta_drive_id():
    try:
        return str(st.secrets["DRIVE_FOLDER_ID"]).strip()
    except (KeyError, FileNotFoundError):
        return ""


def configuracion_apps_script():
    try:
        return (
            str(st.secrets["APPS_SCRIPT_URL"]).strip(),
            str(st.secrets["APPS_SCRIPT_TOKEN"]).strip(),
        )
    except (KeyError, FileNotFoundError):
        return "", ""


def llamar_apps_script(payload):
    url, token = configuracion_apps_script()
    if not url or not token:
        raise RuntimeError("Falta configurar APPS_SCRIPT_URL y APPS_SCRIPT_TOKEN.")
    payload["token"] = token
    respuesta = requests.post(url, json=payload, timeout=90)
    respuesta.raise_for_status()
    resultado = respuesta.json()
    if not resultado.get("ok"):
        raise RuntimeError(resultado.get("error", "Google Apps Script rechazó la operación."))
    return resultado


def subir_a_drive(archivo):
    script_url, _ = configuracion_apps_script()
    if script_url:
        resultado = llamar_apps_script({
            "accion": "subir",
            "nombre": Path(archivo.name).name,
            "mime_type": archivo.type or "application/octet-stream",
            "contenido": base64.b64encode(archivo.getvalue()).decode("ascii"),
        })
        return {
            "id": resultado["id"],
            "nombre": resultado.get("nombre", Path(archivo.name).name),
            "mime_type": resultado.get("mime_type", archivo.type or "application/octet-stream"),
            "url": resultado.get("url", ""),
            "fuente": "apps_script",
        }
    drive = cliente_drive()
    carpeta = carpeta_drive_id()
    if not drive or not carpeta:
        raise RuntimeError("Falta configurar DRIVE_FOLDER_ID o las credenciales de Google Drive.")
    nombre = Path(archivo.name).name
    medio = MediaIoBaseUpload(
        io.BytesIO(archivo.getvalue()),
        mimetype=archivo.type or "application/octet-stream",
        resumable=False,
    )
    creado = drive.files().create(
        body={"name": nombre, "parents": [carpeta]},
        media_body=medio,
        fields="id,name,mimeType,webViewLink",
        supportsAllDrives=True,
    ).execute()
    return {
        "id": creado["id"],
        "nombre": creado.get("name", nombre),
        "mime_type": creado.get("mimeType", archivo.type or "application/octet-stream"),
        "url": creado.get("webViewLink", ""),
    }


def descargar_de_drive(file_id, fuente="drive"):
    if fuente == "apps_script":
        resultado = llamar_apps_script({"accion": "descargar", "id": file_id})
        return base64.b64decode(resultado["contenido"])
    drive = cliente_drive()
    if not drive:
        raise RuntimeError("No fue posible conectar con Google Drive.")
    salida = io.BytesIO()
    descarga = MediaIoBaseDownload(
        salida, drive.files().get_media(fileId=file_id, supportsAllDrives=True)
    )
    terminado = False
    while not terminado:
        _, terminado = descarga.next_chunk()
    return salida.getvalue()


def extraer_temas_acta(archivo):
    documento = Document(io.BytesIO(archivo.getvalue()))
    textos = [p.text.strip() for p in documento.paragraphs if p.text.strip()]
    inicio = next((i for i, t in enumerate(textos) if "temas a tratar" in t.lower()), None)
    fin = next((i for i, t in enumerate(textos) if "desarrollo de temas" in t.lower()), None)
    if inicio is None:
        raise ValueError("No se encontró la sección 'Temas a Tratar'.")
    candidatos = textos[inicio + 1:fin if fin is not None else len(textos)]
    candidatos = [t for t in candidatos if not t.lower().startswith("el inspector fiscal cita")]
    return candidatos


def valor_booleano(valor):
    if isinstance(valor, bool):
        return valor
    return str(valor).strip().lower() in {"true", "verdadero", "1", "sí", "si"}


def cargar():
    conexion = cliente_gsheets()
    if conexion:
        try:
            tabla = conexion.read(worksheet="solicitudes", ttl=0).dropna(how="all").fillna("")
            if tabla.empty:
                return []
            registros = tabla.to_dict("records")
            for registro in registros:
                documentos = registro.get("documentos", "[]")
                try:
                    registro["documentos"] = json.loads(documentos) if isinstance(documentos, str) else []
                except json.JSONDecodeError:
                    registro["documentos"] = []
                registro.setdefault("clase_registro", "Solicitud / equipo")
                registro.setdefault("ctar_numero", "Sin asignar")
                registro.setdefault("tipo_materia", "Tema general CTAR")
                registro.setdefault("formalizacion", "En revisión CTAR")
                registro.setdefault("publicar_hospital", True)
            return registros
        except Exception as error:
            st.error(f"No fue posible leer la planilla de Google Sheets: {error}")
            return []
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    if not DATA_FILE.exists():
        guardar(datos_iniciales())
    try:
        return json.loads(DATA_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return datos_iniciales()


def guardar(datos):
    conexion = cliente_gsheets()
    if conexion:
        columnas = ["id", "tema", "hospital", "servicio", "sic", "inventario", "motivo", "fecha_ingreso", "estado", "ultima_actualizacion", "proximo_paso", "observaciones", "documentos", "clase_registro", "ctar_numero", "tipo_materia", "formalizacion", "publicar_hospital"]
        filas = []
        for item in datos:
            fila = item.copy()
            fila["documentos"] = json.dumps(fila.get("documentos", []), ensure_ascii=False)
            filas.append(fila)
        tabla = pd.DataFrame(filas, columns=columnas)
        conexion.update(worksheet="solicitudes", data=tabla)
        return
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")


def eliminar_solicitud(solicitud_id):
    datos = [item for item in cargar() if item["id"] != solicitud_id]
    guardar(datos)


def clave_admin():
    try:
        return str(st.secrets["ADMIN_PASSWORD"])
    except (KeyError, FileNotFoundError):
        return ""


def login():
    if "perfil" not in st.session_state:
        st.session_state.perfil = None
    if st.session_state.perfil:
        return
    col1, col2, col3 = st.columns([1, 2.1, 1])
    with col2:
        if LOGO.exists(): st.image(str(LOGO), width=210)
        st.title("Seguimiento de solicitudes CTAR")
        st.caption("Servicio de Salud Metropolitano Occidente")
        st.info("El Hospital puede consultar el avance. La administración CTAR requiere una clave.")
        c1, c2 = st.columns(2)
        with c1:
            st.subheader("🏥 Hospital")
            st.write("Consulta estados, fechas, próximos pasos y documentos. No permite modificar información.")
            if st.button("Ingresar como Hospital", type="primary", use_container_width=True):
                st.session_state.perfil = "Hospital"; st.rerun()
        with c2:
            st.subheader("🔐 Administrador CTAR")
            st.write("Registra solicitudes y actualiza el avance del proceso.")
            with st.form("login_admin"):
                clave = st.text_input("Clave de acceso", type="password")
                if st.form_submit_button("Ingresar a administración", use_container_width=True):
                    if not clave_admin():
                        st.error("La clave administrativa no está configurada en los Secrets de Streamlit.")
                    elif hashlib.sha256(clave.encode()).hexdigest() == hashlib.sha256(clave_admin().encode()).hexdigest():
                        st.session_state.perfil = "Administrador CTAR"; st.rerun()
                    else: st.error("La clave ingresada no es correcta.")
    st.stop()


def ficha_detalle(s):
    color = COLORES[s["estado"]]
    st.markdown(f"<span class='estado' style='background:{color}'>{s['estado']}</span>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns(3)
    c1.text_input("Hospital", s["hospital"], disabled=True, key=f"h_{s['id']}")
    c2.text_input("Servicio solicitante", s["servicio"], disabled=True, key=f"sv_{s['id']}")
    c3.text_input("Fecha de ingreso", s["fecha_ingreso"], disabled=True, key=f"f_{s['id']}")
    c4, c5 = st.columns(2)
    c4.text_input("Número SIC", s["sic"], disabled=True, key=f"sic_{s['id']}")
    c5.text_input("Número de inventario", s["inventario"], disabled=True, key=f"inv_{s['id']}")
    st.markdown(f"**Motivo de la solicitud:** {s['motivo']}")
    st.info(f"**Última actualización:** {s['ultima_actualizacion']}\n\n**Próximo paso:** {s['proximo_paso']}")
    st.write("**Observaciones:**", s["observaciones"] or "Sin observaciones")
    if s.get("documentos"):
        st.write("**Documentos asociados:**")
        for i, documento in enumerate(s["documentos"]):
            if isinstance(documento, dict) and documento.get("id"):
                nombre = documento.get("nombre", "Documento")
                try:
                    contenido = descargar_de_drive(documento["id"], documento.get("fuente", "drive"))
                    st.download_button(
                        f"📎 {nombre}", contenido, file_name=nombre,
                        mime=documento.get("mime_type"), key=f"d_{s['id']}_{i}"
                    )
                except Exception as error:
                    st.warning(f"No fue posible abrir {nombre}: {error}")
            else:
                st.caption("⚠️ Adjunto antiguo no disponible. Debe cargarse nuevamente para guardarlo en Drive.")


def fecha_legible(valor):
    try:
        return datetime.strptime(str(valor)[:10], "%Y-%m-%d").strftime("%d-%m-%Y")
    except ValueError:
        return str(valor)


@st.dialog("Ficha de seguimiento", width="large")
def ficha_hospital_lateral(item):
    estado_actual = item.get("estado", ETAPAS[0])
    indice = ETAPAS.index(estado_actual) if estado_actual in ETAPAS else 0
    color = COLORES.get(estado_actual, "#0B6DAA")
    st.markdown("<div class='drawer-eye'>FICHA DE SEGUIMIENTO</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='drawer-title'>{item.get('tema','')}</div><span class='h-pill' style='background:{color}18;color:{color}'>● &nbsp;{estado_actual}</span>", unsafe_allow_html=True)
    formalizacion = item.get("formalizacion", "En revisión CTAR")
    st.caption(f"CTAR N.° {item.get('ctar_numero','Sin asignar')} · {item.get('tipo_materia','Tema general CTAR')} · {formalizacion}")
    if formalizacion != "Formalizado":
        st.warning("Antecedente en revisión. Su contenido puede modificarse y no constituye todavía un acuerdo formal del CTAR.")
    progreso = "".join(
        f"<div class='drawer-stage {'done' if i <= indice else ''}'><span>{i+1}</span></div>"
        for i in range(len(ETAPAS))
    )
    st.markdown(f"<div class='drawer-progress'>{progreso}</div>", unsafe_allow_html=True)
    st.markdown(f"""
    <div class='drawer-grid'>
      <div><small>HOSPITAL</small><b>{item.get('hospital','')}</b></div>
      <div><small>SERVICIO SOLICITANTE</small><b>{item.get('servicio','')}</b></div>
      <div><small>NÚMERO SIC</small><b>{item.get('sic','')}</b></div>
      <div><small>NÚMERO DE INVENTARIO</small><b>{item.get('inventario','')}</b></div>
      <div><small>FECHA DE INGRESO</small><b>{fecha_legible(item.get('fecha_ingreso',''))}</b></div>
      <div><small>MOTIVO DE LA SOLICITUD</small><b>{item.get('motivo','')}</b></div>
    </div>
    <div class='drawer-update'><small>ÚLTIMA ACTUALIZACIÓN</small><b>{item.get('ultima_actualizacion','')}</b><small>PRÓXIMO PASO</small><b>{item.get('proximo_paso','')}</b></div>
    <div class='drawer-section'><small>OBSERVACIONES</small><p>{item.get('observaciones') or 'Sin observaciones'}</p></div>
    <div class='drawer-section'><small>DOCUMENTOS Y CORREOS ASOCIADOS</small></div>
    """, unsafe_allow_html=True)
    if not item.get("documentos"):
        st.caption("No hay documentos asociados.")
    for i, documento in enumerate(item.get("documentos", [])):
        if isinstance(documento, dict) and documento.get("id"):
            nombre = documento.get("nombre", "Documento")
            try:
                contenido = descargar_de_drive(documento["id"], documento.get("fuente", "drive"))
                st.download_button(f"▧  {nombre}", contenido, file_name=nombre,
                    mime=documento.get("mime_type"), key=f"drawer_doc_{item.get('id')}_{i}",
                    use_container_width=True)
            except Exception as error:
                st.warning(f"No fue posible abrir {nombre}: {error}")
        else:
            st.caption("Adjunto antiguo no disponible; debe cargarse nuevamente.")


@st.dialog("Gestión de solicitud", width="large")
def ficha_admin_lateral(s, datos):
    estado = s.get("estado", ETAPAS[0])
    idx = ETAPAS.index(estado) if estado in ETAPAS else 0
    color = COLORES.get(estado, "#0B6DAA")
    st.markdown("<div class='admin-eye'>FICHA ADMINISTRATIVA</div>", unsafe_allow_html=True)
    st.markdown(f"<div class='admin-title'>{s.get('tema','')}</div>", unsafe_allow_html=True)
    st.markdown(
        f"<span class='estado' style='background:{color}'>{estado}</span> "
        f"<span class='estado' style='background:#60758A'>{s.get('formalizacion','En revisión CTAR')}</span>"
        f"<div class='admin-meta'>CTAR N.° {s.get('ctar_numero','Sin asignar')} · {s.get('tipo_materia','Tema general CTAR')} · {s.get('servicio','')}</div>",
        unsafe_allow_html=True,
    )
    st.markdown(
        f"<div class='admin-summary'><b>Última actualización:</b> {s.get('ultima_actualizacion','')}<br>"
        f"<b>Próximo paso:</b> {s.get('proximo_paso','')}</div>",
        unsafe_allow_html=True,
    )

    st.markdown("<div class='admin-section'>Avance del proceso</div>", unsafe_allow_html=True)
    if idx < len(ETAPAS) - 1:
        siguiente = ETAPAS[idx + 1]
        if st.button(f"Avanzar a: {siguiente} →", key=f"drawer_next_{s['id']}", type="primary", use_container_width=True):
            s.update({
                "estado": siguiente,
                "ultima_actualizacion": f"{date.today().isoformat()} - Estado actualizado a {siguiente}",
                "proximo_paso": PROXIMOS[siguiente],
            })
            guardar(datos)
            st.rerun()
    else:
        st.success("Esta solicitud ya se encuentra finalizada.")

    st.markdown("<div class='admin-section'>Editar información</div>", unsafe_allow_html=True)
    with st.form(f"drawer_edit_{s['id']}"):
        tema_editado = st.text_input("Tema o equipo", s.get("tema", ""))
        servicio_editado = st.text_input("Servicio solicitante", s.get("servicio", ""))
        ed1, ed2 = st.columns(2)
        sic_editado = ed1.text_input("Número SIC", s.get("sic", ""))
        inventario_editado = ed2.text_input("Número de inventario", s.get("inventario", ""))
        fecha_editada = st.text_input("Fecha de ingreso", s.get("fecha_ingreso", ""))
        motivo_editado = st.text_area("Motivo de la solicitud", s.get("motivo", ""))
        observaciones_editadas = st.text_area("Observaciones", s.get("observaciones", ""))
        mg1, mg2 = st.columns(2)
        ctar_editado = mg1.text_input("Sesión CTAR", s.get("ctar_numero", "Sin asignar"))
        tipo_actual = s.get("tipo_materia", "Tema general CTAR")
        tipo_editado = mg2.selectbox("Tipo de materia", TIPOS_MATERIA, index=TIPOS_MATERIA.index(tipo_actual) if tipo_actual in TIPOS_MATERIA else len(TIPOS_MATERIA)-1)
        form_actual = s.get("formalizacion", "En revisión CTAR")
        formalizacion_editada = st.selectbox("Nivel de formalización", FORMALIZACIONES, index=FORMALIZACIONES.index(form_actual) if form_actual in FORMALIZACIONES else 1)
        publicar_editado = st.checkbox("Visible para el Hospital", value=valor_booleano(s.get("publicar_hospital", True)))
        nuevo = st.selectbox("Estado del proceso", ETAPAS, index=idx)
        actual = st.text_input("Última actualización", s.get("ultima_actualizacion", ""))
        prox = st.text_input("Próximo paso", s.get("proximo_paso", ""))

        documentos_actuales = [d.get("nombre", "Documento") if isinstance(d, dict) else str(d) for d in s.get("documentos", [])]
        quitar_documentos = st.multiselect("Quitar documentos", documentos_actuales) if documentos_actuales else []
        docs_nuevos = st.file_uploader("Agregar documentos", accept_multiple_files=True)
        guardar_cambios = st.form_submit_button("Guardar todos los cambios", type="primary", use_container_width=True)

    if guardar_cambios:
        documentos = [d for d in s.get("documentos", []) if (d.get("nombre", "Documento") if isinstance(d, dict) else str(d)) not in quitar_documentos]
        try:
            for doc in docs_nuevos:
                documentos.append(subir_a_drive(doc))
        except Exception as error:
            st.error(f"No fue posible guardar los nuevos adjuntos: {error}")
            st.stop()
        s.update({
            "tema": tema_editado, "servicio": servicio_editado, "sic": sic_editado,
            "inventario": inventario_editado, "fecha_ingreso": fecha_editada,
            "motivo": motivo_editado, "observaciones": observaciones_editadas,
            "estado": nuevo, "ultima_actualizacion": actual, "proximo_paso": prox,
            "documentos": documentos, "ctar_numero": ctar_editado,
            "tipo_materia": tipo_editado, "formalizacion": formalizacion_editada,
            "publicar_hospital": publicar_editado,
        })
        guardar(datos)
        st.rerun()

    st.markdown("<div class='admin-section'>Documentos asociados</div>", unsafe_allow_html=True)
    if not s.get("documentos"):
        st.caption("No hay documentos asociados.")
    for i, documento in enumerate(s.get("documentos", [])):
        if isinstance(documento, dict) and documento.get("id"):
            nombre = documento.get("nombre", "Documento")
            try:
                contenido = descargar_de_drive(documento["id"], documento.get("fuente", "drive"))
                st.download_button(f"📎 {nombre}", contenido, file_name=nombre, mime=documento.get("mime_type"), key=f"admin_doc_{s['id']}_{i}", use_container_width=True)
            except Exception as error:
                st.warning(f"No fue posible abrir {nombre}: {error}")

    st.markdown("<div class='admin-section'>Eliminar registro</div>", unsafe_allow_html=True)
    confirmar = st.checkbox("Confirmo que deseo eliminar esta solicitud", key=f"drawer_confirm_{s['id']}")
    if st.button("🗑 Eliminar solicitud", key=f"drawer_delete_{s['id']}", disabled=not confirmar, use_container_width=True):
        eliminar_solicitud(s["id"])
        st.rerun()


def vista_hospital(datos):
    datos = [d for d in datos if valor_booleano(d.get("publicar_hospital", True))]
    st.markdown("""
    <style>
    [data-testid="stSidebar"], [data-testid="stHeader"]{display:none!important}
    .block-container{max-width:1480px;padding:1rem 2rem 2.5rem!important}
    .h-top{display:flex;align-items:center;gap:13px;border-bottom:1px solid #E1E8EF;padding:0 0 14px}
    .h-logo{width:42px;height:42px;border-radius:11px;background:#0C3157;color:white;display:grid;place-items:center;font:800 22px Archivo,sans-serif}
    .h-brand b{display:block;color:#13283C;font-size:16px}.h-brand span{color:#6A7C90;font-size:12px}
    .h-live{margin-left:auto;color:#607286;font-size:12px}.h-dot{display:inline-block;width:8px;height:8px;border-radius:50%;background:#14A46F;margin-right:7px}
    .h-intro{display:flex;justify-content:space-between;align-items:end;padding:28px 4px 18px}
    .h-eye{font-size:11px;font-weight:800;letter-spacing:.14em;color:#0872BC}.h-title{font-size:31px;color:#102A43;margin:7px 0 4px}.h-sub{color:#62768B;font-size:14px}
    .h-count{min-width:150px;text-align:center;color:#17334F;line-height:1.35}.h-count b{display:block;font-size:27px;color:#0B3157;margin-bottom:2px}.h-count span{display:block;font-size:12px}.h-count small{display:block;color:#8293A6;font-size:10px;margin-top:5px}
    .h-flow{display:grid;grid-template-columns:repeat(5,1fr);background:#fff;border:1px solid #D9E4ED;border-radius:14px;padding:20px 22px;margin-bottom:14px;box-shadow:0 5px 18px #123A5A0A}
    .h-step{position:relative;padding-left:40px}.h-step:not(:last-child):after{content:'→';position:absolute;right:10px;top:9px;color:#9AAFC1}.h-n{position:absolute;left:0;top:0;width:30px;height:30px;border-radius:50%;background:#EAF3F9;color:#0872BC;display:grid;place-items:center;font-weight:800}.h-step b{display:block;font-size:13px;color:#142B40}.h-step small{color:#8293A6;font-size:10px}
    .h-guides{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:18px}.h-guide{background:#fff;border:1px solid #DCE6EE;border-radius:11px;padding:13px 16px;color:#607488;font-size:11px}.h-guide b{display:block;color:#18344D;font-size:13px;margin-bottom:2px}
    .h-box{background:#fff;border:1px solid #D8E3EC;border-radius:14px;padding:20px 22px;box-shadow:0 5px 18px #123A5A0A}.h-box-title{font-size:17px;color:#142C43;margin-bottom:2px}.h-muted{color:#8293A6;font-size:11px}
    .h-head,.h-row{display:grid;grid-template-columns:2.25fr 1.25fr 1.05fr 1.35fr 2fr 2.2fr .35fr;gap:12px;align-items:center}.h-head{background:#F1F5F8;color:#60758B;font-size:9px;font-weight:800;letter-spacing:.05em;padding:10px 12px;margin:12px -22px 0}.h-row{padding:13px 0;border-bottom:1px solid #E5EBF0;font-size:11px;color:#1B344B}.h-row:last-child{border-bottom:0}.h-topic b{display:block;font-size:12px;color:#102A43}.h-topic small,.h-service small{display:block;color:#8394A7;margin-top:2px}.h-pill{display:inline-block;padding:5px 9px;border-radius:999px;font-weight:700;font-size:10px;white-space:nowrap}.h-arrow{width:28px;height:28px;border-radius:50%;background:#EDF5FA;display:grid;place-items:center;color:#0B6DAA;font-size:17px}
    div[data-testid="stDialog"] div[role="dialog"]{position:fixed!important;right:0!important;top:0!important;height:100vh!important;max-height:100vh!important;width:590px!important;max-width:94vw!important;border-radius:0!important;margin:0!important;padding:22px 30px!important;overflow-y:auto!important}
    .drawer-eye{font-size:11px;font-weight:800;letter-spacing:.15em;color:#0872BC}.drawer-title{font-size:25px;color:#142B40;margin:12px 0 10px}.drawer-progress{display:grid;grid-template-columns:repeat(5,1fr);margin:28px 0 24px;border-bottom:1px solid #DCE5EC;padding-bottom:22px}.drawer-stage{position:relative;text-align:center}.drawer-stage:before{content:'';position:absolute;top:14px;left:-50%;width:100%;height:2px;background:#E0E7ED}.drawer-stage:first-child:before{display:none}.drawer-stage span{position:relative;z-index:1;display:inline-grid;place-items:center;width:28px;height:28px;border-radius:50%;background:#E8EDF1;color:#8090A0;font-weight:800}.drawer-stage.done span,.drawer-stage.done:before{background:#0D355D;color:white}.drawer-grid{display:grid;grid-template-columns:1fr 1fr;gap:20px 28px;margin:8px 0 22px}.drawer-grid small,.drawer-update small,.drawer-section small{display:block;color:#8092A7;font-size:9px;letter-spacing:.04em;margin-bottom:6px}.drawer-grid b,.drawer-update b{display:block;color:#20364B;font-size:12px;line-height:1.45}.drawer-update{background:#EEF5FA;border-left:3px solid #0B78C2;border-radius:5px;padding:15px 17px;margin-bottom:24px}.drawer-update b{margin-bottom:13px}.drawer-update b:last-child{margin-bottom:0}.drawer-section{margin:18px 0;color:#40586D;font-size:12px}.drawer-section p{margin:0;line-height:1.5}
    .h-docstate{display:block;color:#8A6A16;font-size:9px;margin-top:4px}.h-footer{display:grid;grid-template-columns:1fr 1fr 1fr;align-items:center;border-top:1px solid #DFE7EE;margin-top:24px;padding:15px 4px 2px;color:#8495A7;font-size:10px}.h-footer div:nth-child(2){text-align:center;color:#60758A}.h-footer div:last-child{text-align:right}
    @media(max-width:900px){.h-flow,.h-guides{grid-template-columns:1fr}.h-step{margin:5px 0}.h-step:after{display:none}.h-head{display:none}.h-row{grid-template-columns:1fr;padding:16px}.h-row>div{margin:2px 0}.h-count{display:none}.h-footer{grid-template-columns:1fr;text-align:center;gap:5px}.h-footer div:last-child{text-align:center}}
    </style>
    """, unsafe_allow_html=True)

    top1, top2 = st.columns([8, 1])
    with top1:
        st.markdown("<div class='h-top'><div class='h-logo'>C</div><div class='h-brand'><b>Seguimiento CTAR</b><span>Hospital Dr. Félix Bulnes</span></div><div class='h-live'><span class='h-dot'></span>Información actualizada</div></div>", unsafe_allow_html=True)
    with top2:
        if st.button("Salir", use_container_width=True):
            st.session_state.perfil = None
            st.rerun()

    activas = sum(x.get("estado") != "Proceso finaliza" for x in datos)
    st.markdown(f"<div class='h-intro'><div><div class='h-eye'>CONTROL Y TRAZABILIDAD</div><div class='h-title'>Solicitudes al CTAR</div><div class='h-sub'>Consulta el avance de cada solicitud enviada por el Hospital, desde su ingreso hasta el cierre del proceso.</div></div><div class='h-count'><b>{activas}</b><span>solicitudes activas</span><small>{len(datos)} registradas en total</small></div></div>", unsafe_allow_html=True)

    descripciones = ["Solicitud recibida", "Antecedentes en análisis", "Acuerdo adoptado", "Documento en firmas", "Seguimiento cerrado"]
    flujo = "".join(f"<div class='h-step'><span class='h-n'>{i}</span><b>{etapa}</b><small>{descripciones[i-1]}</small></div>" for i, etapa in enumerate(ETAPAS, 1))
    st.markdown(f"<div class='h-flow'>{flujo}</div>", unsafe_allow_html=True)
    st.markdown("<div class='h-guides'><div class='h-guide'><b>⌕ ¿Cómo consultar?</b>Busca el equipo o número SIC y abre su ficha para revisar los antecedentes.</div><div class='h-guide'><b>→ Revisa el próximo paso</b>Te indica claramente qué acción sigue y qué falta para continuar.</div><div class='h-guide'><b>✓ Información centralizada</b>El estado, las fechas y los documentos se encuentran en un solo lugar.</div></div>", unsafe_allow_html=True)

    with st.container(border=True):
        st.markdown("<div class='h-box-title'>Listado de solicitudes</div><div class='h-muted'>Busca y filtra la información registrada</div>", unsafe_allow_html=True)
        f1, f2, f3, f4, f5 = st.columns([2, 1, 1, 1, 1])
        buscar = f1.text_input("Buscar", placeholder="Equipo, SIC o inventario...", label_visibility="collapsed")
        estado = f2.selectbox("Estado", ["Todos los estados"] + ETAPAS, label_visibility="collapsed")
        servicios = sorted({str(x.get("servicio", "")) for x in datos})
        servicio = f3.selectbox("Servicio", ["Todos los servicios"] + servicios, label_visibility="collapsed")
        sesiones = sorted({str(x.get("ctar_numero", "Sin asignar")) for x in datos})
        sesion = f4.selectbox("Sesión CTAR", ["Todas las sesiones"] + sesiones, label_visibility="collapsed")
        desde = f5.date_input("Desde", value=None, label_visibility="collapsed")

        filtrados = []
        for item in datos:
            texto = f"{item.get('tema','')} {item.get('sic','')} {item.get('inventario','')} {item.get('servicio','')}".lower()
            fecha_ok = True
            if desde:
                try: fecha_ok = date.fromisoformat(str(item.get("fecha_ingreso", ""))[:10]) >= desde
                except ValueError: fecha_ok = True
            if (buscar.lower() in texto and
                (estado == "Todos los estados" or item.get("estado") == estado) and
                (servicio == "Todos los servicios" or item.get("servicio") == servicio) and fecha_ok and
                (sesion == "Todas las sesiones" or str(item.get("ctar_numero", "Sin asignar")) == sesion)):
                filtrados.append(item)

        st.caption(f"{len(filtrados)} resultado(s) encontrado(s)")
        st.markdown("<div class='h-head'><div>TEMA O EQUIPO</div><div>SERVICIO</div><div>FECHA DE INGRESO</div><div>ESTADO ACTUAL</div><div>ÚLTIMA ACTUALIZACIÓN</div><div>PRÓXIMO PASO</div><div></div></div>", unsafe_allow_html=True)
        for item in filtrados:
            color = COLORES.get(item.get("estado"), "#0B6DAA")
            fila, accion = st.columns([20, 1])
            with fila:
                html = f"<div class='h-row' style='grid-template-columns:2.25fr 1.25fr 1.05fr 1.35fr 2fr 2.2fr'><div class='h-topic'><b>{item.get('tema','')}</b><small>CTAR {item.get('ctar_numero','S/A')} · {item.get('tipo_materia','')}</small></div><div class='h-service'>{item.get('servicio','')}<small>SIC {item.get('sic','')} · Inv. {item.get('inventario','')}</small></div><div>{fecha_legible(item.get('fecha_ingreso',''))}</div><div><span class='h-pill' style='background:{color}18;color:{color}'>● &nbsp;{item.get('estado','')}</span><span class='h-docstate'>{item.get('formalizacion','')}</span></div><div>{item.get('ultima_actualizacion','')}</div><div>{item.get('proximo_paso','')}</div></div>"
                st.markdown(html, unsafe_allow_html=True)
            with accion:
                if st.button("›", key=f"open_hospital_{item.get('id')}", help="Abrir ficha"):
                    ficha_hospital_lateral(item)
    st.markdown(f"<div class='h-footer'><div>Seguimiento CTAR · Servicio de Salud Metropolitano Occidente</div><div>Desarrollado por <b>Bayron Retamal González</b></div><div>Última sincronización: {fecha_legible(date.today().isoformat())}</div></div>", unsafe_allow_html=True)




def main():
    css(); login(); datos = cargar(); perfil = st.session_state.perfil
    if perfil == "Hospital":
        vista_hospital(datos)
        return
    with st.sidebar:
        if LOGO.exists(): st.image(str(LOGO), use_container_width=True)
        st.markdown(f"**Perfil activo:** {perfil}")
        st.caption("Solo el Administrador CTAR puede modificar información.")
        if st.button("Cerrar sesión", use_container_width=True):
            st.session_state.perfil = None; st.rerun()
        st.divider(); st.markdown("**¿Qué significa cada etapa?**")
        for i, etapa in enumerate(ETAPAS, 1): st.caption(f"{i}. {etapa}")

    st.markdown(f"<div class='hero'><h1>Seguimiento de solicitudes CTAR</h1><p>Consulta clara y centralizada del avance de las solicitudes del Hospital.</p><span class='perfil'>{perfil}</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='pasos'>" + "".join([f"<div class='paso'><span class='numero'>{i}</span><b>{e}</b><small>{d}</small></div>" for i,(e,d) in enumerate(zip(ETAPAS,["Solicitud recibida","Antecedentes en revisión","Acuerdo adoptado","Acta en proceso de firma","Seguimiento cerrado"]),1)]) + "</div>", unsafe_allow_html=True)
    if perfil == "Hospital": st.markdown("<div class='solo-lectura'>✓ Estás en modo consulta. Puedes revisar toda la información, pero no modificarla.</div>", unsafe_allow_html=True)
    if perfil == "Administrador CTAR" and not almacenamiento_permanente():
        st.warning("Google Sheets todavía no está conectado. Mientras tanto, los datos se guardan solo temporalmente en Streamlit.")

    activas = sum(x["estado"] != "Proceso finaliza" for x in datos)
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Solicitudes totales", len(datos)); m2.metric("En revisión", sum(x["estado"]=="CTAR revisa" for x in datos)); m3.metric("En firma", sum(x["estado"]=="Acta se firma" for x in datos)); m4.metric("Activas", activas)

    if perfil == "Administrador CTAR":
        with st.expander("📄 Importar temas desde un borrador de acta", expanded=False):
            st.info("El archivo se analiza como borrador. Los temas no se publican al Hospital hasta que los revises y autorices su visibilidad.")
            ia1, ia2 = st.columns([1, 2])
            numero_acta = ia1.text_input("Número de sesión CTAR", placeholder="Ej.: 149", key="import_ctar_num")
            archivo_acta = ia2.file_uploader("Borrador de acta Word", type=["docx"], key="import_acta")
            if st.button("Analizar temas del borrador", disabled=not (numero_acta and archivo_acta), use_container_width=True):
                try:
                    temas = extraer_temas_acta(archivo_acta)
                    if not temas:
                        st.error("No se encontraron temas para importar.")
                    else:
                        st.session_state["importacion_ctar"] = {
                            "numero": numero_acta,
                            "nombre": archivo_acta.name,
                            "mime_type": archivo_acta.type,
                            "contenido": archivo_acta.getvalue(),
                            "temas": temas,
                        }
                        st.rerun()
                except Exception as error:
                    st.error(f"No fue posible leer el borrador: {error}")

            importacion = st.session_state.get("importacion_ctar")
            if importacion:
                st.success(f"Se detectaron {len(importacion['temas'])} temas en el CTAR N.° {importacion['numero']}. Revisa la tabla antes de guardar.")
                tabla_temas = pd.DataFrame([
                    {"N.°": n, "Tema": tema, "Tipo de materia": "Tema general CTAR", "Responsable inicial": "CTAR", "Visible Hospital": False}
                    for n, tema in enumerate(importacion["temas"], 1)
                ])
                temas_revisados = st.data_editor(
                    tabla_temas, use_container_width=True, hide_index=True,
                    column_config={
                        "Tipo de materia": st.column_config.SelectboxColumn(options=TIPOS_MATERIA, required=True),
                        "Responsable inicial": st.column_config.SelectboxColumn(options=["CTAR", "Hospital", "SSMOC", "MINSAL", "Inspección Fiscal", "Sociedad Concesionaria"], required=True),
                        "Visible Hospital": st.column_config.CheckboxColumn(),
                    },
                    key=f"editor_import_{importacion['numero']}",
                )
                ic1, ic2 = st.columns(2)
                if ic1.button("Guardar temas revisados", type="primary", use_container_width=True):
                    try:
                        acta_memoria = ArchivoMemoria(importacion["nombre"], importacion["mime_type"], importacion["contenido"])
                        documento_acta = subir_a_drive(acta_memoria)
                        existentes = {(str(x.get("ctar_numero")), str(x.get("tema")).strip().lower()) for x in datos}
                        creados = 0
                        for _, fila in temas_revisados.iterrows():
                            clave = (str(importacion["numero"]), str(fila["Tema"]).strip().lower())
                            if clave in existentes:
                                continue
                            datos.insert(0, {
                                "id": f"CTAR-{importacion['numero']}-T{int(fila['N.°']):02d}-{datetime.now().strftime('%H%M%S')}",
                                "tema": str(fila["Tema"]).strip(), "hospital": "Hospital Dr. Félix Bulnes",
                                "servicio": str(fila["Responsable inicial"]), "sic": "No aplica", "inventario": "No aplica",
                                "motivo": f"Tema N.° {int(fila['N.°'])} incorporado en el borrador del Acta CTAR N.° {importacion['numero']}.",
                                "fecha_ingreso": date.today().isoformat(), "estado": "CTAR revisa",
                                "ultima_actualizacion": f"{date.today().isoformat()} - Tema importado desde borrador",
                                "proximo_paso": "Revisar y validar el contenido en el CTAR",
                                "observaciones": "Antecedente sujeto a modificación hasta la formalización del acta.",
                                "documentos": [documento_acta], "clase_registro": "Tema de sesión",
                                "ctar_numero": str(importacion["numero"]), "tipo_materia": str(fila["Tipo de materia"]),
                                "formalizacion": "Borrador interno", "publicar_hospital": bool(fila["Visible Hospital"]),
                            })
                            existentes.add(clave); creados += 1
                        guardar(datos)
                        del st.session_state["importacion_ctar"]
                        st.success(f"Se guardaron {creados} temas sin duplicar registros.")
                        st.rerun()
                    except Exception as error:
                        st.error(f"No fue posible guardar la importación: {error}")
                if ic2.button("Cancelar importación", use_container_width=True):
                    del st.session_state["importacion_ctar"]
                    st.rerun()

        with st.expander("➕ Registrar una nueva solicitud", expanded=False):
            with st.form("nueva"):
                a,b = st.columns(2); tema=a.text_input("Tema o equipo *"); servicio=b.text_input("Servicio solicitante *")
                c,d,e = st.columns(3); sic=c.text_input("Número SIC"); inventario=d.text_input("Número de inventario"); ingreso=e.date_input("Fecha de ingreso", date.today())
                g,h,i = st.columns(3)
                ctar_numero = g.text_input("Sesión CTAR", placeholder="Ej.: 149")
                tipo_materia = h.selectbox("Tipo de materia", TIPOS_MATERIA)
                formalizacion = i.selectbox("Nivel de formalización", FORMALIZACIONES, index=1)
                publicar = st.checkbox("Visible para el Hospital", value=False, help="Actívalo solo cuando corresponda compartir el antecedente. Los borradores se mostrarán con una advertencia.")
                motivo=st.text_area("Motivo de la solicitud *"); obs=st.text_area("Observaciones")
                docs=st.file_uploader("Documentos o correos asociados", accept_multiple_files=True)
                if st.form_submit_button("Guardar solicitud", type="primary"):
                    if not tema or not servicio or not motivo: st.error("Completa los campos obligatorios.")
                    else:
                        doc_paths=[]
                        try:
                            for doc in docs:
                                doc_paths.append(subir_a_drive(doc))
                        except Exception as error:
                            st.error(f"No fue posible guardar los adjuntos en Google Drive: {error}")
                            st.stop()
                        datos.insert(0,{"id":f"CTAR-{datetime.now().strftime('%Y%m%d%H%M%S')}","tema":tema,"hospital":"Hospital Dr. Félix Bulnes","servicio":servicio,"sic":sic or "Por asignar","inventario":inventario or "Por confirmar","motivo":motivo,"fecha_ingreso":ingreso.isoformat(),"estado":ETAPAS[0],"ultima_actualizacion":f"{date.today().isoformat()} - Solicitud registrada","proximo_paso":"CTAR revisará los antecedentes recibidos","observaciones":obs,"documentos":doc_paths,"clase_registro":"Solicitud / equipo","ctar_numero":ctar_numero or "Sin asignar","tipo_materia":tipo_materia,"formalizacion":formalizacion,"publicar_hospital":publicar})
                        guardar(datos); st.success("Solicitud registrada correctamente."); st.rerun()

    st.subheader("Solicitudes registradas")
    st.markdown("<div class='ayuda'><b>Cómo consultar:</b> usa los filtros y luego abre la ficha del equipo. Revisa especialmente el estado actual y el próximo paso.</div>", unsafe_allow_html=True)
    f1,f2,f3 = st.columns([2,1,1]); buscar=f1.text_input("Buscar", placeholder="Equipo, SIC, inventario o servicio..."); estado=f2.selectbox("Estado",["Todos"]+ETAPAS); servicios=sorted({x["servicio"] for x in datos}); serv=f3.selectbox("Servicio",["Todos"]+servicios)
    filtrados=[x for x in datos if buscar.lower() in f"{x['tema']} {x['sic']} {x['inventario']} {x['servicio']}".lower() and (estado=="Todos" or x["estado"]==estado) and (serv=="Todos" or x["servicio"]==serv)]
    st.caption(f"{len(filtrados)} resultado(s) encontrado(s)")
    for s in filtrados:
        color=COLORES[s["estado"]]
        tarjeta, abrir = st.columns([12, 1])
        with tarjeta:
            st.markdown(f"<div class='ficha'><h4>{s['tema']}</h4><div class='meta'>CTAR {s.get('ctar_numero','Sin asignar')} · {s.get('tipo_materia','Tema general CTAR')} · {s['servicio']} · SIC {s['sic']} · Ingreso {s['fecha_ingreso']}</div><span class='estado' style='background:{color}'>{s['estado']}</span> <span class='estado' style='background:#60758A'>{s.get('formalizacion','En revisión CTAR')}</span><div class='proximo'><b>Próximo paso:</b> {s['proximo_paso']}</div></div>",unsafe_allow_html=True)
        with abrir:
            st.write("")
            if st.button("›", key=f"open_admin_{s['id']}", help="Abrir gestión lateral", use_container_width=True):
                ficha_admin_lateral(s, datos)
    if perfil == "Administrador CTAR" and datos:
        st.divider(); df=pd.DataFrame(datos); st.download_button("⬇ Descargar seguimiento en CSV",df.to_csv(index=False).encode("utf-8-sig"),"seguimiento_ctar.csv","text/csv")


if __name__ == "__main__": main()
