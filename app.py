import hashlib
import json
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st
from supabase import Client, create_client

BASE = Path(__file__).parent
DATA_FILE = BASE / "data" / "solicitudes.json"
UPLOAD_DIR = BASE / "data" / "documentos"
LOGO = BASE / "assets" / "logo_ssmoc.jpg"
ETAPAS = ["Hospital envía", "CTAR revisa", "CTAR acuerda", "Acta se firma", "Proceso finaliza"]
COLORES = {"Hospital envía": "#0B6DAA", "CTAR revisa": "#F2A900", "CTAR acuerda": "#7B4FA3", "Acta se firma": "#008C95", "Proceso finaliza": "#198754"}
PROXIMOS = {
    "Hospital envía": "CTAR revisará los antecedentes recibidos",
    "CTAR revisa": "Adoptar acuerdo en sesión CTAR",
    "CTAR acuerda": "Preparar el acta para revisión y firma",
    "Acta se firma": "Completar las firmas pendientes",
    "Proceso finaliza": "Sin acciones pendientes",
}

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
    @media(max-width:900px){.pasos{grid-template-columns:1fr}.paso{min-height:auto}}
    </style>""", unsafe_allow_html=True)


def datos_iniciales():
    return [
        {"id":"CTAR-149-01","tema":"Hervidores industriales (3)","hospital":"Hospital Dr. Félix Bulnes","servicio":"SEDILE","sic":"60.045","inventario":"HFB-SED-118 / 120","motivo":"Baja y reposición por término de vida útil","fecha_ingreso":"2026-07-16","estado":"Acta se firma","ultima_actualizacion":"2026-07-21 - Acta CTAR N.° 149 actualizada","proximo_paso":"Obtener firmas de los integrantes CTAR","observaciones":"Pronunciamientos favorables recibidos por correo.","documentos":[]},
        {"id":"CTAR-140-06","tema":"Ventiladores neonatales (10)","hospital":"Hospital Dr. Félix Bulnes","servicio":"Neonatología","sic":"52.881","inventario":"6 AVEA / 4 VN500","motivo":"Reposición integral por obsolescencia","fecha_ingreso":"2026-06-10","estado":"CTAR revisa","ultima_actualizacion":"2026-07-18 - EETT clínicas recibidas","proximo_paso":"Revisar especificaciones técnicas en sesión CTAR","observaciones":"Preferencia clínica informada: Dräger VN800.","documentos":[]},
        {"id":"CTAR-139-04","tema":"Cámaras de videolaparoscopía","hospital":"Hospital Dr. Félix Bulnes","servicio":"Pabellón","sic":"60.729 / 7.990","inventario":"Por confirmar","motivo":"Equipos no reparables; solicitud de baja","fecha_ingreso":"2026-06-30","estado":"Hospital envía","ultima_actualizacion":"2026-07-17 - Se solicitó completar antecedentes","proximo_paso":"Hospital debe remitir inventarios y EETT","observaciones":"Vinculada al Ord. HFB N.° 763.","documentos":[]},
    ]


@st.cache_resource
def cliente_supabase():
    try:
        url = st.secrets["SUPABASE_URL"]
        key = st.secrets["SUPABASE_KEY"]
        return create_client(url, key)
    except (KeyError, FileNotFoundError):
        return None


def almacenamiento_permanente():
    return cliente_supabase() is not None


def cargar():
    db = cliente_supabase()
    if db:
        try:
            respuesta = db.table("solicitudes").select("data").order("updated_at", desc=True).execute()
            return [fila["data"] for fila in respuesta.data]
        except Exception as error:
            st.error(f"No fue posible conectar con la base de datos: {error}")
            return []
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    if not DATA_FILE.exists():
        guardar(datos_iniciales())
    try:
        return json.loads(DATA_FILE.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return datos_iniciales()


def guardar(datos):
    db = cliente_supabase()
    if db:
        if datos:
            filas = [{"id": item["id"], "data": item, "updated_at": datetime.now().isoformat()} for item in datos]
            db.table("solicitudes").upsert(filas).execute()
        return
    DATA_FILE.parent.mkdir(parents=True, exist_ok=True)
    DATA_FILE.write_text(json.dumps(datos, ensure_ascii=False, indent=2), encoding="utf-8")


def eliminar_solicitud(solicitud_id):
    db = cliente_supabase()
    if db:
        db.table("solicitudes").delete().eq("id", solicitud_id).execute()
        return
    datos = [item for item in cargar() if item["id"] != solicitud_id]
    guardar(datos)


def clave_admin():
    try:
        return str(st.secrets["ADMIN_PASSWORD"])
    except (KeyError, FileNotFoundError):
        return "265727"


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
                    if hashlib.sha256(clave.encode()).hexdigest() == hashlib.sha256(clave_admin().encode()).hexdigest():
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
        for ruta in s["documentos"]:
            archivo = BASE / ruta
            if archivo.exists(): st.download_button(f"📎 {archivo.name}", archivo.read_bytes(), file_name=archivo.name, key=f"d_{s['id']}_{archivo.name}")


def main():
    css(); login(); datos = cargar(); perfil = st.session_state.perfil
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
        st.warning("La base de datos permanente todavía no está configurada. Agrega SUPABASE_URL y SUPABASE_KEY en los Secrets de Streamlit.")

    activas = sum(x["estado"] != "Proceso finaliza" for x in datos)
    m1,m2,m3,m4 = st.columns(4)
    m1.metric("Solicitudes totales", len(datos)); m2.metric("En revisión", sum(x["estado"]=="CTAR revisa" for x in datos)); m3.metric("En firma", sum(x["estado"]=="Acta se firma" for x in datos)); m4.metric("Activas", activas)

    if perfil == "Administrador CTAR":
        with st.expander("➕ Registrar una nueva solicitud", expanded=False):
            with st.form("nueva"):
                a,b = st.columns(2); tema=a.text_input("Tema o equipo *"); servicio=b.text_input("Servicio solicitante *")
                c,d,e = st.columns(3); sic=c.text_input("Número SIC"); inventario=d.text_input("Número de inventario"); ingreso=e.date_input("Fecha de ingreso", date.today())
                motivo=st.text_area("Motivo de la solicitud *"); obs=st.text_area("Observaciones")
                docs=st.file_uploader("Documentos o correos asociados", accept_multiple_files=True)
                if st.form_submit_button("Guardar solicitud", type="primary"):
                    if not tema or not servicio or not motivo: st.error("Completa los campos obligatorios.")
                    else:
                        doc_paths=[]; UPLOAD_DIR.mkdir(parents=True,exist_ok=True)
                        for doc in docs:
                            safe=f"{datetime.now().strftime('%Y%m%d%H%M%S')}_{Path(doc.name).name}"; path=UPLOAD_DIR/safe; path.write_bytes(doc.getbuffer()); doc_paths.append(str(path.relative_to(BASE)))
                        datos.insert(0,{"id":f"CTAR-{datetime.now().strftime('%Y%m%d%H%M%S')}","tema":tema,"hospital":"Hospital Dr. Félix Bulnes","servicio":servicio,"sic":sic or "Por asignar","inventario":inventario or "Por confirmar","motivo":motivo,"fecha_ingreso":ingreso.isoformat(),"estado":ETAPAS[0],"ultima_actualizacion":f"{date.today().isoformat()} - Solicitud registrada","proximo_paso":"CTAR revisará los antecedentes recibidos","observaciones":obs,"documentos":doc_paths})
                        guardar(datos); st.success("Solicitud registrada correctamente."); st.rerun()

    st.subheader("Solicitudes registradas")
    st.markdown("<div class='ayuda'><b>Cómo consultar:</b> usa los filtros y luego abre la ficha del equipo. Revisa especialmente el estado actual y el próximo paso.</div>", unsafe_allow_html=True)
    f1,f2,f3 = st.columns([2,1,1]); buscar=f1.text_input("Buscar", placeholder="Equipo, SIC, inventario o servicio..."); estado=f2.selectbox("Estado",["Todos"]+ETAPAS); servicios=sorted({x["servicio"] for x in datos}); serv=f3.selectbox("Servicio",["Todos"]+servicios)
    filtrados=[x for x in datos if buscar.lower() in f"{x['tema']} {x['sic']} {x['inventario']} {x['servicio']}".lower() and (estado=="Todos" or x["estado"]==estado) and (serv=="Todos" or x["servicio"]==serv)]
    st.caption(f"{len(filtrados)} resultado(s) encontrado(s)")
    for s in filtrados:
        color=COLORES[s["estado"]]
        st.markdown(f"<div class='ficha'><h4>{s['tema']}</h4><div class='meta'>{s['servicio']} · SIC {s['sic']} · Ingreso {s['fecha_ingreso']}</div><span class='estado' style='background:{color}'>{s['estado']}</span><div class='proximo'><b>Próximo paso:</b> {s['proximo_paso']}</div></div>",unsafe_allow_html=True)
        with st.expander(f"Ver ficha completa — {s['tema']}"):
            ficha_detalle(s)
            if perfil == "Administrador CTAR":
                idx=ETAPAS.index(s["estado"])
                st.markdown("#### Gestión de la solicitud")
                if idx < len(ETAPAS)-1:
                    siguiente = ETAPAS[idx+1]
                    if st.button(f"Avanzar a: {siguiente} →", key=f"next_{s['id']}", type="primary", use_container_width=True):
                        s.update({"estado": siguiente, "ultima_actualizacion": f"{date.today().isoformat()} - Estado actualizado a {siguiente}", "proximo_paso": PROXIMOS[siguiente]})
                        guardar(datos); st.success(f"La solicitud avanzó a {siguiente}."); st.rerun()
                else:
                    st.success("Esta solicitud ya se encuentra finalizada.")
                st.markdown("##### Editar información ingresada")
                ed1, ed2 = st.columns(2)
                tema_editado = ed1.text_input("Tema o equipo", s["tema"], key=f"tema_edit_{s['id']}")
                servicio_editado = ed2.text_input("Servicio solicitante", s["servicio"], key=f"serv_edit_{s['id']}")
                ed3, ed4, ed5 = st.columns(3)
                sic_editado = ed3.text_input("Número SIC", s["sic"], key=f"sic_edit_{s['id']}")
                inventario_editado = ed4.text_input("Número de inventario", s["inventario"], key=f"inv_edit_{s['id']}")
                fecha_editada = ed5.text_input("Fecha de ingreso", s["fecha_ingreso"], key=f"fecha_edit_{s['id']}")
                motivo_editado = st.text_area("Motivo de la solicitud", s["motivo"], key=f"motivo_edit_{s['id']}")
                observaciones_editadas = st.text_area("Observaciones", s.get("observaciones", ""), key=f"obs_edit_{s['id']}")
                nuevo=st.selectbox("Cambiar estado manualmente",ETAPAS,index=idx,key=f"est_{s['id']}")
                actual=st.text_input("Última actualización",s["ultima_actualizacion"],key=f"act_{s['id']}")
                prox=st.text_input("Próximo paso",s["proximo_paso"],key=f"prox_{s['id']}")
                if st.button("Guardar todos los cambios",key=f"save_{s['id']}"):
                    s.update({"tema":tema_editado,"servicio":servicio_editado,"sic":sic_editado,"inventario":inventario_editado,"fecha_ingreso":fecha_editada,"motivo":motivo_editado,"observaciones":observaciones_editadas,"estado":nuevo,"ultima_actualizacion":actual,"proximo_paso":prox})
                    guardar(datos); st.success("Cambios guardados correctamente."); st.rerun()
                st.divider()
                confirmar = st.checkbox("Confirmo que deseo eliminar esta solicitud", key=f"confirm_{s['id']}")
                if st.button("🗑 Eliminar solicitud", key=f"delete_{s['id']}", disabled=not confirmar):
                    eliminar_solicitud(s["id"]); st.success("Solicitud eliminada correctamente."); st.rerun()
    if perfil == "Administrador CTAR" and datos:
        st.divider(); df=pd.DataFrame(datos); st.download_button("⬇ Descargar seguimiento en CSV",df.to_csv(index=False).encode("utf-8-sig"),"seguimiento_ctar.csv","text/csv")


if __name__ == "__main__": main()
