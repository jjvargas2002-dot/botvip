from functools import wraps
from io import BytesIO
import csv
import io
import requests
from datetime import datetime

from flask import Flask, render_template, request, send_file, redirect, url_for, session, flash
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from sqlalchemy import text
from werkzeug.security import generate_password_hash, check_password_hash
from config import Config
from models import db, Cliente, Caso, Conversacion, EstadoConversacion, Operador, Averia
from whatsapp import enviar_mensaje

app = Flask(__name__)
app.config.from_object(Config)
db.init_app(app)


def login_requerido(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "operador_id" not in session:
            flash("Por favor, inicia sesión para acceder al sistema.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


def obtener_estadisticas(branch=None, es_admin=False):
    try:
        query_total = Averia.query
        query_pendientes = Averia.query.filter_by(estado="PENDIENTE")
        query_reparados = Averia.query.filter_by(estado="REPARADO")

        if not es_admin and branch and branch != "ALL":
            query_total = query_total.filter_by(branch=branch)
            query_pendientes = query_pendientes.filter_by(branch=branch)
            query_reparados = query_reparados.filter_by(branch=branch)

        return {
            "totales": query_total.count(),
            "pendientes": query_pendientes.count(),
            "reparados": query_reparados.count(),
        }
    except Exception as e:
        print("Error obteniendo estadísticas:", e)
        return {
            "totales": 0,
            "pendientes": 0,
            "reparados": 0,
        }


def registrar_conversacion(averia_id, remitente, mensaje):
    conversacion = Conversacion(
        averia_id=averia_id,
        remitente=remitente,
        mensaje=mensaje,
    )
    db.session.add(conversacion)


def asegurar_esquema():
    # Asegurar que las columnas existan dinámicamente si las tablas ya existen en el DB
    columnas = [
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS dni VARCHAR(20) UNIQUE",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS rol VARCHAR(20) DEFAULT 'operador'",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS branch VARCHAR(30) DEFAULT 'ALL'",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS telefono VARCHAR(20)",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS activo BOOLEAN DEFAULT TRUE",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        
        "ALTER TABLE conversaciones ADD COLUMN IF NOT EXISTS averia_id INTEGER REFERENCES averias(id)",
        "ALTER TABLE estados_conversacion ADD COLUMN IF NOT EXISTS averia_id INTEGER REFERENCES averias(id)",
        
        # Columnas de averias
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS branch VARCHAR(50)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS codigo_wo VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS cuenta VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS detalles TEXT",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS dias_pendientes DOUBLE PRECISION",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS estado VARCHAR(50) DEFAULT 'PENDIENTE'",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS status_caja VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS contrata VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS periodo_pendiente VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS site VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS caja VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS coordenadas VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS origen VARCHAR(20) DEFAULT 'SHEETS'",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS fecha_resolucion TIMESTAMP",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS tecnico_id INTEGER REFERENCES operadores(id)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS material_cable_m INTEGER DEFAULT 0",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS material_conectores INTEGER DEFAULT 0",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS material_rosetas INTEGER DEFAULT 0",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS material_mangas INTEGER DEFAULT 0",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS material_acopladores INTEGER DEFAULT 0",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS material_comentarios TEXT",
        
        "ALTER TABLE averias DROP CONSTRAINT IF EXISTS averias_cuenta_key",
        "ALTER TABLE averias ALTER COLUMN cuenta DROP NOT NULL"
    ]

    for consulta in columnas:
        try:
            db.session.execute(text(consulta))
        except Exception as e:
            print(f"Error ejecutando consulta de esquema {consulta}: {e}")
            db.session.rollback()
    db.session.commit()


def crear_operador_defecto():
    try:
        # 1. Crear/restablecer Administrador general FBB
        admin = Operador.query.filter(db.func.lower(Operador.dni) == "fbb").first()
        if admin:
            admin.dni = "FBB"
            admin.password_hash = generate_password_hash("Bitel@123")
            admin.rol = "admin"
            admin.branch = "ALL"
            admin.correo = "admin@botvip.com"
        else:
            default_pwd = generate_password_hash("Bitel@123")
            admin = Operador(
                nombre="Administrador FBB",
                dni="FBB",
                correo="admin@botvip.com",
                password_hash=default_pwd,
                rol="admin",
                branch="ALL",
                activo=True
            )
            db.session.add(admin)
        
        # 2. Crear operadores para cada Branch
        sedes = ["LI1", "LI2", "LI3", "LI4", "LI7", "ARE", "PIU", "SAN", "CAJ", "LAL", "HUN", "CUS", "JUN"]
        for sede in sedes:
            op_sede = Operador.query.filter(db.func.lower(Operador.dni) == sede.lower()).first()
            if not op_sede:
                pwd_sede = generate_password_hash("Vip@123")
                nuevo_op = Operador(
                    nombre=f"Operador Sede {sede}",
                    dni=sede.upper(),
                    correo=f"{sede.lower()}@botvip.com",
                    password_hash=pwd_sede,
                    rol="operador",
                    branch=sede.upper(),
                    activo=True
                )
                db.session.add(nuevo_op)
                print(f"Semilla creada para sede {sede}")
        
        db.session.commit()
        print("Sedes y Administrador sembrados correctamente.")
    except Exception as e:
        print("Error al crear operadores por defecto:", e)
        db.session.rollback()


def sincronizar_drive():
    url = "https://docs.google.com/spreadsheets/d/1eaNxCpm8JF1JcZS3_ldwMRINGYFaW6RsQQWybvRi_P8/export?format=csv&gid=1775459558"
    try:
        response = requests.get(url, timeout=15)
        if response.status_code != 200:
            return False, f"Error de conexión con Google Sheets (Status: {response.status_code})"
        
        content = response.content.decode('utf-8')
        f = io.StringIO(content)
        reader = csv.reader(f)
        
        # Leer cabecera
        header = next(reader, None)
        if not header:
            return False, "El archivo de Google Sheets está vacío."
        
        # Mapeo de columnas por índice
        indices = {}
        for i, col in enumerate(header):
            col_name = col.strip().upper().replace("Ó", "O").replace("Í", "I")
            indices[col_name] = i
        
        # Verificar columnas críticas
        columnas_requeridas = ["BRANCH", "CUENTA", "COORDENADAS"]
        for col in columnas_requeridas:
            if col not in indices:
                return False, f"Falta la columna requerida: {col}"
        
        cuentas_drive = set()
        nuevos = 0
        actualizados = 0
        nuevos_por_branch = {}
        
        for row in reader:
            if not row or len(row) <= max(indices.values()):
                continue
            
            branch = row[indices["BRANCH"]].strip()
            cuenta = row[indices["CUENTA"]].strip()
            
            # Saltar si no hay cuenta o branch
            if not cuenta or not branch:
                continue
            
            cuentas_drive.add(cuenta)
            
            # Extraer campos
            codigo_wo = row[indices.get("CODIGO DE WO")].strip() if "CODIGO DE WO" in indices else ""
            detalles = row[indices.get("DETALLES")].strip() if "DETALLES" in indices else ""
            
            dias_str = row[indices.get("DIAS PENDIENTES")].strip() if "DIAS PENDIENTES" in indices else ""
            try:
                dias_pendientes = float(dias_str) if dias_str else 0.0
            except ValueError:
                dias_pendientes = 0.0
                
            status_ont = row[indices.get("ESTADO")].strip() if "ESTADO" in indices else ""
            status_caja = row[indices.get("STATUS DE LA CAJA")].strip() if "STATUS DE LA CAJA" in indices else ""
            contrata = row[indices.get("CONTRATA")].strip() if "CONTRATA" in indices else ""
            periodo_pendiente = row[indices.get("PERIODO PENDIENTE")].strip() if "PERIODO PENDIENTE" in indices else ""
            site = row[indices.get("SITE")].strip() if "SITE" in indices else ""
            caja = row[indices.get("CAJA")].strip() if "CAJA" in indices else ""
            coordenadas = row[indices.get("COORDENADAS")].strip() if "COORDENADAS" in indices else ""
            
            # Buscar en BD
            averia = Averia.query.filter_by(cuenta=cuenta).first()
            if averia:
                # Si ya existe, actualizar datos del drive solo si está pendiente localmente
                if averia.estado == "PENDIENTE":
                    averia.branch = branch
                    averia.codigo_wo = codigo_wo
                    averia.detalles = detalles
                    averia.dias_pendientes = dias_pendientes
                    averia.status_caja = status_caja if status_caja else status_ont
                    averia.contrata = contrata
                    averia.periodo_pendiente = periodo_pendiente
                    averia.site = site
                    averia.caja = caja
                    averia.coordenadas = coordenadas
                    actualizados += 1
            else:
                # Crear nueva avería
                nueva = Averia(
                    branch=branch,
                    codigo_wo=codigo_wo,
                    cuenta=cuenta,
                    detalles=detalles,
                    dias_pendientes=dias_pendientes,
                    estado="PENDIENTE",
                    status_caja=status_caja if status_caja else status_ont,
                    contrata=contrata,
                    periodo_pendiente=periodo_pendiente,
                    site=site,
                    caja=caja,
                    coordenadas=coordenadas,
                    origen="SHEETS"
                )
                db.session.add(nueva)
                nuevos += 1
                if branch not in nuevos_por_branch:
                    nuevos_por_branch[branch] = []
                nuevos_por_branch[branch].append(nueva)
                
        # Cerrar averías de tipo 'SHEETS' que ya no están en el Drive (fueron solucionadas externamente)
        cerrados_ext = Averia.query.filter(
            Averia.origen == "SHEETS",
            Averia.estado == "PENDIENTE",
            ~Averia.cuenta.in_(list(cuentas_drive))
        ).all()
        
        for av in cerrados_ext:
            av.estado = "REPARADO"
            av.fecha_resolucion = datetime.now()
            av.material_comentarios = "Cerrado automáticamente al no figurar en la lista del Drive."
        
        db.session.commit()
        
        # Notificar por WhatsApp de las nuevas averías importadas
        for br, lista_av in nuevos_por_branch.items():
            total_nuevos = len(lista_av)
            if total_nuevos == 0:
                continue
            
            # Buscar operadores activos de esta sede con teléfono
            operadores = Operador.query.filter_by(branch=br, activo=True).all()
            for op in operadores:
                if op.telefono:
                    num = op.telefono.strip()
                    if num:
                        try:
                            if total_nuevos <= 3:
                                for av in lista_av:
                                    msg = (
                                        f"🔔 *Nueva avería asignada ({br})*:\n\n"
                                        f"• *Cuenta*: {av.cuenta}\n"
                                        f"• *Caja*: {av.caja}\n"
                                        f"• *Obs*: {av.detalles or 'Sin detalles'}\n\n"
                                        f"💡 Escribe `reparar {av.cuenta}` cuando la soluciones."
                                    )
                                    enviar_mensaje(num, msg)
                            else:
                                msg = (
                                    f"🔔 *Nuevas averías asignadas ({br})*:\n\n"
                                    f"Se han registrado *{total_nuevos}* nuevas averías para tu sede en el sistema.\n\n"
                                    f"📋 Escribe `todos` para ver la lista completa de pendientes."
                                )
                                enviar_mensaje(num, msg)
                        except Exception as e_notif:
                            print(f"Error enviando notificación WhatsApp a {num}: {e_notif}")
                            
        total_cerrados = len(cerrados_ext)
        return True, f"Sincronización exitosa: {nuevos} creados, {actualizados} actualizados y {total_cerrados} cerrados externamente."
    except Exception as e:
        db.session.rollback()
        print("Error en sincronización de drive:", e)
        return False, f"Error en sincronización: {str(e)}"


@app.route("/diagnostico-operadores")
def diagnostico_operadores():
    try:
        crear_operador_defecto()
        ops = Operador.query.all()
        resultado = f"<h3>Diagnóstico de Operadores y Sedes</h3>"
        resultado += "<table border='1' cellpadding='5' style='border-collapse:collapse;'>"
        resultado += "<tr><th>ID</th><th>Nombre</th><th>DNI</th><th>Sede</th><th>Teléfono</th><th>Rol</th><th>Activo</th><th>Verificación Password (Bitel@123 / Vip@123)</th></tr>"
        
        for op in ops:
            matches_pwd = check_password_hash(op.password_hash, "Bitel@123") or check_password_hash(op.password_hash, "Vip@123")
            resultado += f"<tr>"
            resultado += f"<td>{op.id}</td>"
            resultado += f"<td>{op.nombre}</td>"
            resultado += f"<td>{op.dni}</td>"
            resultado += f"<td>{op.branch}</td>"
            resultado += f"<td>{op.telefono or 'No registrado'}</td>"
            resultado += f"<td>{op.rol}</td>"
            resultado += f"<td>{op.activo}</td>"
            resultado += f"<td>{'CORRECTA' if matches_pwd else 'OTRA CONTRASEÑA'}</td>"
            resultado += f"</tr>"
            
        resultado += "</table>"
        resultado += "<br><a href='/login'>Ir al Login</a>"
        return resultado
    except Exception as e:
        db.session.rollback()
        return f"Error en diagnóstico: {str(e)}"


@app.route("/login", methods=["GET", "POST"])
def login():
    if "operador_id" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        dni = request.form["dni"].strip()
        password = request.form["password"].strip()

        operador = Operador.query.filter(db.func.lower(Operador.dni) == db.func.lower(dni), Operador.activo == True).first()
        if operador and check_password_hash(operador.password_hash, password):
            session["operador_id"] = operador.id
            session["operador_nombre"] = operador.nombre
            session["operador_dni"] = operador.dni
            session["operador_rol"] = operador.rol
            session["operador_branch"] = operador.branch
            
            flash(f"¡Bienvenido de nuevo, {operador.nombre} ({operador.branch})!", "success")
            return redirect(url_for("dashboard"))
        else:
            flash("DNI/Usuario o contraseña incorrectos.", "danger")

    return render_template("login.html")


@app.route("/register", methods=["GET", "POST"])
def register():
    if "operador_id" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        dni = request.form["dni"].strip()
        nombre = request.form["nombre"].strip()
        correo = request.form["correo"].strip()
        password = request.form["password"].strip()
        branch = request.form["branch"].strip().upper()
        telefono = request.form.get("telefono", "").strip()

        existente = Operador.query.filter((db.func.lower(Operador.dni) == db.func.lower(dni)) | (db.func.lower(Operador.correo) == db.func.lower(correo))).first()
        if existente:
            flash("Ya existe un operador con ese DNI o correo.", "danger")
            return render_template("register.html")

        pwd_hash = generate_password_hash(password or dni)

        nuevo = Operador(
            dni=dni,
            nombre=nombre,
            correo=correo,
            password_hash=pwd_hash,
            rol="operador",
            branch=branch,
            telefono=telefono or None,
            activo=True
        )
        db.session.add(nuevo)
        db.session.commit()
        flash("Cuenta registrada con éxito. Ya puedes iniciar sesión.", "success")
        return redirect(url_for("login"))

    return render_template("register.html")


@app.route("/operadores")
@login_requerido
def listar_operadores():
    if session.get("operador_rol") != "admin":
        flash("Acceso denegado: Se requieren permisos de administrador.", "danger")
        return redirect(url_for("dashboard"))

    operadores = Operador.query.order_by(Operador.id.desc()).all()
    stats = obtener_estadisticas(branch=session.get("operador_branch"), es_admin=True)
    return render_template("operadores.html", operadores=operadores, stats=stats)


@app.route("/operadores/editar-telefono", methods=["POST"])
@login_requerido
def editar_telefono():
    if session.get("operador_rol") != "admin":
        flash("Acceso denegado.", "danger")
        return redirect(url_for("dashboard"))

    op_id = request.form.get("operador_id")
    telefono = request.form.get("telefono", "").strip()
    
    operador = db.session.get(Operador, op_id)
    if operador:
        operador.telefono = telefono or None
        db.session.commit()
        flash(f"Teléfono de {operador.nombre} actualizado a {telefono or 'Vacío'}.", "success")
    else:
        flash("Operador no encontrado.", "danger")
        
    return redirect(url_for("listar_operadores"))


@app.route("/logout")
def logout():
    session.clear()
    flash("Has cerrado sesión correctamente.", "info")
    return redirect(url_for("login"))


@app.route("/cambiar-password", methods=["POST"])
@login_requerido
def cambiar_password():
    password_actual = request.form["password_actual"].strip()
    nueva_password = request.form["nueva_password"].strip()
    confirmar_password = request.form["confirmar_password"].strip()

    if nueva_password != confirmar_password:
        flash("La nueva contraseña y la confirmación no coinciden.", "danger")
        return redirect(request.referrer or url_for("dashboard"))

    operador = db.session.get(Operador, session["operador_id"])
    if operador and check_password_hash(operador.password_hash, password_actual):
        operador.password_hash = generate_password_hash(nueva_password)
        db.session.commit()
        flash("Tu contraseña ha sido cambiada exitosamente.", "success")
    else:
        flash("La contraseña actual es incorrecta.", "danger")

    return redirect(request.referrer or url_for("dashboard"))


@app.route("/", methods=["GET"])
@login_requerido
def dashboard():
    es_admin = session.get("operador_rol") == "admin"
    branch = session.get("operador_branch")
    
    stats = obtener_estadisticas(branch=branch, es_admin=es_admin)
    
    # Obtener las averías pendientes correspondientes a la sede
    query = Averia.query.filter_by(estado="PENDIENTE")
    if not es_admin and branch != "ALL":
        query = query.filter_by(branch=branch)
        
    averias_pendientes = query.order_by(Averia.dias_pendientes.desc().nullslast()).all()
    
    # Serializar datos para Leaflet map
    map_data = []
    for av in averias_pendientes:
        if av.coordenadas:
            coords = av.coordenadas.split(",")
            if len(coords) == 2:
                try:
                    lat = float(coords[0].strip())
                    lng = float(coords[1].strip())
                    map_data.append({
                        "id": av.id,
                        "branch": av.branch,
                        "cuenta": av.cuenta,
                        "codigo_wo": av.codigo_wo or "N/A",
                        "detalles": av.detalles or "Sin descripción",
                        "contrata": av.contrata or "Sin contrata",
                        "site": av.site or "N/A",
                        "caja": av.caja or "N/A",
                        "dias": av.dias_pendientes or 0,
                        "lat": lat,
                        "lng": lng
                    })
                except ValueError:
                    continue
                    
    # Verificar si es una sede de provincias que requiere añadir averías manualmente
    es_provincia = branch not in ["LI1", "LI2", "LI3", "LI4", "LI7", "ALL"]
    
    import json
    return render_template(
        "dashboard.html", 
        stats=stats, 
        map_json=json.dumps(map_data),
        averias_list=averias_pendientes,
        es_provincia=es_provincia
    )


@app.route("/averias/sync", methods=["POST", "GET"])
@login_requerido
def sync_sheets():
    # Solo permitir sincronizar a Lima y Administradores
    branch = session.get("operador_branch")
    if branch not in ["LI1", "LI2", "LI3", "LI4", "LI7", "ALL"] and session.get("operador_rol") != "admin":
        flash("Tu sede no tiene habilitada la sincronización automática.", "warning")
        return redirect(url_for("dashboard"))
        
    success, message = sincronizar_drive()
    if success:
        flash(message, "success")
    else:
        flash(message, "danger")
    return redirect(url_for("dashboard"))


@app.route("/averias/resolver/<int:id>", methods=["POST"])
@login_requerido
def resolver_averia(id):
    averia = db.session.get(Averia, id)
    if not averia:
        flash("La avería no existe.", "danger")
        return redirect(url_for("dashboard"))
        
    # Verificar pertenencia a la sede
    branch = session.get("operador_branch")
    if session.get("operador_rol") != "admin" and branch != "ALL" and averia.branch != branch:
        flash("No tienes permisos para resolver averías de otra sede.", "danger")
        return redirect(url_for("dashboard"))
        
    try:
        cable_m = int(request.form.get("cable_m", 0) or 0)
        conectores = int(request.form.get("conectores", 0) or 0)
        rosetas = int(request.form.get("rosetas", 0) or 0)
        mangas = int(request.form.get("mangas", 0) or 0)
        acopladores = int(request.form.get("acopladores", 0) or 0)
        comentarios = request.form.get("comentarios", "").strip()
        
        averia.estado = "REPARADO"
        averia.fecha_resolucion = datetime.now()
        averia.tecnico_id = session.get("operador_id")
        averia.material_cable_m = cable_m
        averia.material_conectores = conectores
        averia.material_rosetas = rosetas
        averia.material_mangas = mangas
        averia.material_acopladores = acopladores
        averia.material_comentarios = comentarios or "Reparado desde el portal"
        
        db.session.commit()
        flash(f"Avería de cuenta {averia.cuenta} resuelta y materiales registrados.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error al resolver avería: {str(e)}", "danger")
        
    return redirect(request.referrer or url_for("dashboard"))


def notificar_nueva_averia_individual(averia):
    # Buscar técnicos de la sede que tengan teléfono y estén activos
    operadores = Operador.query.filter_by(branch=averia.branch, activo=True).all()
    for op in operadores:
        if op.telefono:
            num = op.telefono.strip()
            if num:
                try:
                    msg = (
                        f"🔔 *Nueva avería asignada a tu sede ({averia.branch})*:\n\n"
                        f"• *Cuenta*: {averia.cuenta}\n"
                        f"• *Caja*: {averia.caja}\n"
                        f"• *Obs*: {averia.detalles or 'Sin detalles'}\n\n"
                        f"💡 Escribe `reparar {averia.cuenta}` cuando la soluciones."
                    )
                    enviar_mensaje(num, msg)
                except Exception as e:
                    print(f"Error notificando nueva avería individual a {num}: {e}")


@app.route("/averias/crear", methods=["POST"])
@login_requerido
def crear_averia_manual():
    branch = session.get("operador_branch")
    es_admin = session.get("operador_rol") == "admin"
    
    # Provincias y Admin pueden crear manual
    if not es_admin and branch in ["LI1", "LI2", "LI3", "LI4", "LI7"]:
        flash("Tu sede no permite agregar averías manualmente (se sincronizan desde el drive).", "danger")
        return redirect(url_for("dashboard"))
        
    try:
        cuenta = request.form.get("cuenta", "").strip()
        if not cuenta:
            import uuid
            cuenta = f"SIN_CUENTA_{uuid.uuid4().hex[:8].upper()}"
            
        site = request.form["site"].strip().upper()
        xbox = request.form["xbox"].strip().upper()
        caja_input = request.form["caja"].strip().upper()
        coordenadas = request.form["coordenadas"].strip()
        detalles = request.form["detalles"].strip()
        contrata = request.form.get("contrata", "").strip()
        
        # Sede de registro
        target_branch = branch if not es_admin else request.form.get("branch", "ARE").strip().upper()
        
        # Validar XBOX
        if xbox not in ["XB01", "XB02"]:
            flash("La XBOX debe ser XB01 o XB02.", "danger")
            return redirect(url_for("dashboard"))
            
        # Componer código de caja: SITE-XBOX-caja
        caja_compuesta = f"{site}-{xbox}-{caja_input}"
        
        # Validar si ya existe la cuenta (solo si no es un placeholder autogenerado)
        existente = None
        if not cuenta.startswith("SIN_CUENTA_"):
            existente = Averia.query.filter_by(cuenta=cuenta).first()
        if existente:
            if existente.estado == "PENDIENTE":
                flash(f"La cuenta {cuenta} ya tiene una avería pendiente activa.", "warning")
                return redirect(url_for("dashboard"))
            else:
                # Si ya existía pero estaba reparada, la reactivamos o creamos una nueva
                # Para evitar duplicados de cuenta unique, eliminamos la anterior o la editamos.
                # Lo mejor es reactivar la existente poniéndola PENDIENTE
                existente.estado = "PENDIENTE"
                existente.branch = target_branch
                existente.caja = caja_compuesta
                existente.site = site
                existente.coordenadas = coordenadas
                existente.detalles = detalles
                existente.contrata = contrata
                existente.dias_pendientes = 0.0
                existente.origen = "MANUAL"
                db.session.commit()
                
                # Notificar a los técnicos de la reactivación
                try:
                    notificar_nueva_averia_individual(existente)
                except Exception as e_notif:
                    print("Error notificando reactivación:", e_notif)
                    
                flash(f"Se reactivó la avería para la cuenta {cuenta}.", "success")
                return redirect(url_for("dashboard"))
        
        nueva = Averia(
            branch=target_branch,
            cuenta=cuenta,
            site=site,
            caja=caja_compuesta,
            coordenadas=coordenadas,
            detalles=detalles,
            contrata=contrata or "Propia",
            estado="PENDIENTE",
            dias_pendientes=0.0,
            origen="MANUAL"
        )
        db.session.add(nueva)
        db.session.commit()
        
        # Notificar a los técnicos por WhatsApp
        try:
            notificar_nueva_averia_individual(nueva)
        except Exception as e_notif:
            print("Error notificado nueva avería manual:", e_notif)
            
        flash(f"Avería manual para cuenta {cuenta} creada con éxito.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error al crear avería: {str(e)}", "danger")
        
    return redirect(url_for("dashboard"))


@app.route("/averias", methods=["GET"])
@login_requerido
def listar_averias():
    es_admin = session.get("operador_rol") == "admin"
    branch = session.get("operador_branch")
    
    query = Averia.query
    if not es_admin and branch != "ALL":
        query = query.filter_by(branch=branch)
        
    averias = query.order_by(Averia.id.desc()).all()
    stats = obtener_estadisticas(branch=branch, es_admin=es_admin)
    return render_template("averias.html", averias=averias, stats=stats)


@app.route("/averias/exportar", methods=["GET"])
@login_requerido
def exportar_averias():
    es_admin = session.get("operador_rol") == "admin"
    branch = session.get("operador_branch")
    
    query = Averia.query
    if not es_admin and branch != "ALL":
        query = query.filter_by(branch=branch)
        
    averias = query.order_by(Averia.id.desc()).all()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Reporte Averias ODN"

    headers = [
        "ID",
        "Sede (Branch)",
        "Código de WO",
        "Cuenta",
        "Detalles",
        "Días Pendientes",
        "Estado del Ticket",
        "Contrata",
        "Site",
        "Caja",
        "Coordenadas",
        "Origen",
        "Fecha Creación",
        "Fecha Resolución",
        "Técnico que Resolvió",
        "Cable Drop (m)",
        "Conectores Mecánicos",
        "Rosetas Ópticas",
        "Mangas/Bandejas",
        "Acopladores",
        "Comentarios Solución"
    ]
    sheet.append(headers)

    for av in averias:
        tecnico_nombre = av.tecnico.nombre if av.tecnico else ""
        sheet.append([
            av.id,
            av.branch,
            av.codigo_wo or "",
            av.cuenta,
            av.detalles or "",
            av.dias_pendientes or 0.0,
            av.estado,
            av.contrata or "",
            av.site or "",
            av.caja or "",
            av.coordenadas or "",
            av.origen,
            av.fecha_creacion.strftime("%Y-%m-%d %H:%M") if av.fecha_creacion else "",
            av.fecha_resolucion.strftime("%Y-%m-%d %H:%M") if av.fecha_resolucion else "Pendiente",
            tecnico_nombre,
            av.material_cable_m,
            av.material_conectores,
            av.material_rosetas,
            av.material_mangas,
            av.material_acopladores,
            av.material_comentarios or ""
        ])

    header_fill = PatternFill("solid", fgColor="1D4ED8")
    header_font = Font(color="FFFFFF", bold=True)

    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")

    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions

    for column_cells in sheet.columns:
        max_length = max(len(str(cell.value or "")) for cell in column_cells)
        column_letter = get_column_letter(column_cells[0].column)
        sheet.column_dimensions[column_letter].width = min(max(max_length + 2, 12), 45)

    for row in sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    return send_file(
        output,
        as_attachment=True,
        download_name=f"reporte_averias_{branch.lower() if branch != 'ALL' else 'global'}.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/webhook", methods=["GET", "POST"])
def webhook():
    if request.method == "GET":
        verify_token = "botvip_token"
        if (
            request.args.get("hub.mode") == "subscribe"
            and request.args.get("hub.verify_token") == verify_token
        ):
            return request.args.get("hub.challenge"), 200
        return "Verification failed", 403

    data = request.get_json()
    try:
        value = data["entry"][0]["changes"][0]["value"]
        messages = value.get("messages")

        if not messages:
            return "OK", 200

        msg = messages[0]
        telefono = msg.get("from")
        texto = msg.get("text", {}).get("body", "")

        if not telefono or not texto:
            return "OK", 200

        procesar_mensaje_tecnico(telefono, texto)

    except Exception as e:
        print("Webhook error:", e)

    return "OK", 200


def procesar_mensaje_tecnico(telefono, texto):
    texto_original = texto.strip()
    texto_normalizado = texto_original.lower()

    # Buscar operador por su número de teléfono (removiendo el prefijo 51 de país si existe)
    tel_normalizado = telefono
    if tel_normalizado.startswith("51") and len(tel_normalizado) > 9:
        tel_normalizado = tel_normalizado[2:]
        
    operador = Operador.query.filter(
        (Operador.telefono == telefono) | 
        (Operador.telefono == tel_normalizado) |
        (db.func.right(Operador.telefono, 9) == db.func.right(telefono, 9)),
        Operador.activo == True
    ).first()

    if not operador:
        # Enviar mensaje indicando número no autorizado
        msg_no_auth = (
            f"❌ *Número no autorizado*\n\n"
            f"Tu número de WhatsApp ({telefono}) no se encuentra registrado en el sistema de técnicos de BotVip.\n\n"
            f"Por favor solicita al Administrador que registre tu número en el panel."
        )
        enviar_mensaje(telefono, msg_no_auth)
        return

    # 0. Cancelación global de flujos guiados
    if texto_normalizado in ["/cancelar", "cancelar"]:
        estado = EstadoConversacion.query.filter_by(telefono=telefono).first()
        if estado:
            # Si estábamos creando una avería manual, eliminar el registro temporal incompleto
            if estado.paso_actual.startswith("crear_"):
                averia_temp = db.session.get(Averia, estado.averia_id)
                if averia_temp:
                    db.session.delete(averia_temp)
            db.session.delete(estado)
            db.session.commit()
            enviar_mensaje(telefono, "🔄 *Operación cancelada*. Volviendo al menú principal.")
        else:
            enviar_mensaje(telefono, "💡 No tienes ningún flujo activo para cancelar.")
        return

    # Si el operador existe, procesamos su conversación
    estado = EstadoConversacion.query.filter_by(telefono=telefono).first()
    
    if estado:
        # Flujo guiado en progreso
        averia = db.session.get(Averia, estado.averia_id)
        if not averia or (averia.estado == "REPARADO" and not estado.paso_actual.startswith("crear_")):
            # Cancelar flujo si la avería no existe o ya está reparada
            db.session.delete(estado)
            db.session.commit()
            enviar_mensaje(telefono, "❌ La avería que estabas editando ya no está disponible o ya fue solucionada.")
            return

        registrar_conversacion(averia.id, "tecnico", texto_original)

        # ----------------------------------------------------
        # FLUJO GUIADO 1: RESOLVER AVERÍA (Reportar materiales)
        # ----------------------------------------------------
        if estado.paso_actual == "esperando_cable":
            try:
                cable_m = int(texto_original)
                if cable_m < 0: raise ValueError
            except ValueError:
                enviar_mensaje(telefono, "❌ *Valor inválido*. Por favor responde solo con un número entero positivo (ej: `120` o `0`):")
                return
            
            averia.material_cable_m = cable_m
            estado.paso_actual = "esperando_conectores"
            db.session.commit()
            
            msg = "🔧 *Paso 2/5*: ¿Cuántos conectores mecánicos utilizaste? (Responde con número entero, ej. `2`, o `0`)."
            registrar_conversacion(averia.id, "bot", msg)
            enviar_mensaje(telefono, msg)
            
        elif estado.paso_actual == "esperando_conectores":
            try:
                conectores = int(texto_original)
                if conectores < 0: raise ValueError
            except ValueError:
                enviar_mensaje(telefono, "❌ *Valor inválido*. Por favor responde solo con un número entero positivo (ej: `2` o `0`):")
                return
            
            averia.material_conectores = conectores
            estado.paso_actual = "esperando_rosetas"
            db.session.commit()
            
            msg = "🏠 *Paso 3/5*: ¿Cuántas rosetas ópticas utilizaste? (Responde con número entero, ej. `1`, o `0`)."
            registrar_conversacion(averia.id, "bot", msg)
            enviar_mensaje(telefono, msg)
            
        elif estado.paso_actual == "esperando_rosetas":
            try:
                rosetas = int(texto_original)
                if rosetas < 0: raise ValueError
            except ValueError:
                enviar_mensaje(telefono, "❌ *Valor inválido*. Por favor responde solo con un número entero positivo (ej: `1` o `0`):")
                return
            
            averia.material_rosetas = rosetas
            estado.paso_actual = "esperando_mangas"
            db.session.commit()
            
            msg = "📦 *Paso 4/5*: ¿Cuántas mangas/bandejas de empalme utilizaste? (Responde con número entero, ej. `1`, o `0`)."
            registrar_conversacion(averia.id, "bot", msg)
            enviar_mensaje(telefono, msg)
            
        elif estado.paso_actual == "esperando_mangas":
            try:
                mangas = int(texto_original)
                if mangas < 0: raise ValueError
            except ValueError:
                enviar_mensaje(telefono, "❌ *Valor inválido*. Por favor responde solo con un número entero positivo (ej: `1` o `0`):")
                return
            
            averia.material_mangas = mangas
            estado.paso_actual = "esperando_comentarios"
            db.session.commit()
            
            msg = "💬 *Paso 5/5*: Escribe un breve comentario u observación sobre la solución (o responde `ninguno`):"
            registrar_conversacion(averia.id, "bot", msg)
            enviar_mensaje(telefono, msg)
            
        elif estado.paso_actual == "esperando_comentarios":
            averia.material_comentarios = texto_original if texto_normalizado != "ninguno" else "Reparado sin comentarios."
            averia.estado = "REPARADO"
            averia.fecha_resolucion = datetime.now()
            averia.tecnico_id = operador.id
            
            # Borrar estado
            db.session.delete(estado)
            db.session.commit()
            
            msg = f"✅ *¡Reparación registrada con éxito!*\n\nLa avería de la cuenta *{averia.cuenta}* ha sido cerrada y los materiales se registraron para la sede *{operador.branch}*."
            registrar_conversacion(averia.id, "bot", msg)
            enviar_mensaje(telefono, msg)

        # ----------------------------------------------------
        # FLUJO GUIADO 2: CREAR AVERÍA MANUALMENTE
        # ----------------------------------------------------
        elif estado.paso_actual == "crear_esperando_cuenta":
            cuenta_val = texto_original
            if texto_normalizado == "ninguna":
                import uuid
                cuenta_val = f"SIN_CUENTA_{uuid.uuid4().hex[:8].upper()}"
            
            averia.cuenta = cuenta_val
            estado.paso_actual = "crear_esperando_site"
            db.session.commit()
            
            enviar_mensaje(telefono, "📍 *Paso 2/7 (Site)*: Escribe el Site (ej: `ARE0071` o `CAL0072`):")
            
        elif estado.paso_actual == "crear_esperando_site":
            averia.site = texto_original.upper()
            estado.paso_actual = "crear_esperando_xbox"
            db.session.commit()
            
            enviar_mensaje(telefono, "📦 *Paso 3/7 (XBOX)*: Responde con `XB01` o `XB02`:")
            
        elif estado.paso_actual == "crear_esperando_xbox":
            xbox_val = texto_original.upper()
            if xbox_val not in ["XB01", "XB02"]:
                enviar_mensaje(telefono, "❌ *Opción inválida*. Por favor responde solo con `XB01` o `XB02`:")
                return
            
            averia.caja = xbox_val  # Guardamos temporalmente el xbox en caja
            estado.paso_actual = "crear_esperando_caja"
            db.session.commit()
            
            enviar_mensaje(telefono, "📥 *Paso 4/7 (Caja)*: Escribe la caja/splitter (ej: `SB111` o `EB214`):")
            
        elif estado.paso_actual == "crear_esperando_caja":
            xbox_temp = averia.caja
            averia.caja = f"{averia.site}-{xbox_temp}-{texto_original.upper()}"
            estado.paso_actual = "crear_esperando_coordenadas"
            db.session.commit()
            
            enviar_mensaje(telefono, "🗺️ *Paso 5/7 (Coordenadas)*: Envía las coordenadas en formato lat,lng (ej: `-16.3988,-71.5369`):")
            
        elif estado.paso_actual == "crear_esperando_coordenadas":
            coords = texto_original.split(",")
            if len(coords) != 2:
                enviar_mensaje(telefono, "❌ *Formato inválido*. Por favor envía las coordenadas en formato lat,lng (ej: `-16.3988,-71.5369`):")
                return
            try:
                float(coords[0].strip())
                float(coords[1].strip())
            except ValueError:
                enviar_mensaje(telefono, "❌ *Formato inválido*. Asegúrate de ingresar números decimales separados por una coma (ej: `-16.3988,-71.5369`):")
                return
                
            averia.coordenadas = texto_original
            estado.paso_actual = "crear_esperando_detalles"
            db.session.commit()
            
            enviar_mensaje(telefono, "💬 *Paso 6/7 (Detalles)*: Escribe una breve descripción de la avería:")
            
        elif estado.paso_actual == "crear_esperando_detalles":
            averia.detalles = texto_original
            estado.paso_actual = "crear_esperando_contrata"
            db.session.commit()
            
            enviar_mensaje(telefono, "🚛 *Paso 7/7 (Contrata)*: Escribe el nombre de la contrata o responde `ninguna` para omitir:")
            
        elif estado.paso_actual == "crear_esperando_contrata":
            averia.contrata = texto_original if texto_normalizado != "ninguna" else "Propia"
            averia.estado = "PENDIENTE"
            averia.dias_pendientes = 0.0
            
            # Limpiar estado
            db.session.delete(estado)
            db.session.commit()
            
            # Notificar a los otros técnicos de la sede
            try:
                notificar_nueva_averia_individual(averia)
            except Exception as e_notif:
                print("Error notificando nueva avería desde WhatsApp:", e_notif)
                
            msg = (
                f"✅ *¡Avería creada con éxito!*\n\n"
                f"• *Cuenta*: {averia.cuenta}\n"
                f"• *Caja*: {averia.caja}\n"
                f"• *Detalles*: {averia.detalles}\n"
                f"• *Sede*: {averia.branch}\n\n"
                f"El ticket ha sido publicado en el mapa de control."
            )
            enviar_mensaje(telefono, msg)
    else:
        # Separar en palabras para extraer comando y argumentos
        palabras = texto_original.split()
        first_word = palabras[0].lower().strip() if palabras else ""
        # Quitar '/' y acentos del primer término
        cmd_clean = first_word.replace("/", "").replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u")
        rest_text = " ".join(palabras[1:]).strip() if len(palabras) > 1 else ""
        
        # Normalizar el texto completo sin acentos ni slashes
        texto_norm = texto_normalizado.replace("/", "").replace("á", "a").replace("é", "e").replace("í", "i").replace("ó", "o").replace("ú", "u").strip()

        # 1. BUSCAR
        if cmd_clean in ["buscar", "bus", "find", "filtrar", "b"] or texto_norm.startswith("buscar ") or texto_norm.startswith("filtrar "):
            q = rest_text if rest_text else (texto_original[7:].strip() if texto_normalizado.startswith("buscar ") else "")
            if not q:
                enviar_mensaje(telefono, "💡 *Uso*: escribe `buscar <caja o cuenta>` (ej: `buscar SB111` o `buscar albertj10`)")
                return
                
            query = Averia.query.filter_by(estado="PENDIENTE")
            if operador.branch != "ALL":
                query = query.filter_by(branch=operador.branch)
                
            query = query.filter(
                (db.func.lower(Averia.caja).like(f"%{q.lower()}%")) |
                (db.func.lower(Averia.cuenta).like(f"%{q.lower()}%")) |
                (db.func.lower(Averia.codigo_wo).like(f"%{q.lower()}%"))
            )
            
            resultados = query.limit(5).all()
            if not resultados:
                enviar_mensaje(telefono, f"🔍 No se encontraron averías pendientes para '{q}' en tu sede ({operador.branch}).")
                return
                
            msg = f"🔍 *Averías encontradas en '{q}':*\n\n"
            for av in resultados:
                maps_link = f"https://www.google.com/maps/search/?api=1&query={av.coordenadas}" if av.coordenadas else "Sin coordenadas"
                msg += f"• *Cuenta*: {av.cuenta}\n"
                msg += f"  *Caja*: {av.caja}\n"
                msg += f"  *Detalles*: {av.detalles or 'Sin descripción'}\n"
                msg += f"  *Mapa*: {maps_link}\n\n"
            enviar_mensaje(telefono, msg.strip())

        # 2. VER TODOS LOS PENDIENTES
        elif cmd_clean in ["todos", "todo", "pendientes", "lista", "listar", "ver"] or texto_norm in ["ver todos", "ver pendientes", "todos los pendientes", "todas", "ver todas", "lista de pendientes"]:
            query = Averia.query.filter_by(estado="PENDIENTE")
            if operador.branch != "ALL":
                query = query.filter_by(branch=operador.branch)
                
            resultados = query.order_by(Averia.dias_pendientes.desc().nullslast(), Averia.id.desc()).limit(20).all()
            total_pendientes = query.count()
            
            if not resultados:
                enviar_mensaje(telefono, f"📋 No hay averías pendientes registradas para la sede {operador.branch}.")
                return
                
            msg = f"📋 *Todos los pendientes ({operador.branch})* - Total: {total_pendientes}\n\n"
            for i, av in enumerate(resultados, 1):
                maps_link = f"https://www.google.com/maps/search/?api=1&query={av.coordenadas}" if av.coordenadas else "Sin coordenadas"
                msg += f"*{i}. Cuenta*: {av.cuenta}\n"
                msg += f"   *Caja*: {av.caja}\n"
                msg += f"   *Obs*: {av.detalles or 'Sin detalles'}\n"
                msg += f"   *Mapa*: {maps_link}\n\n"
                
            if total_pendientes > 20:
                msg += f"⚠️ *Mostrando las 20 averías más antiguas de un total de {total_pendientes} pendientes.*\n\n"
            msg += "💡 Para cerrar una avería, envía:\n`reparar <cuenta>`"
            enviar_mensaje(telefono, msg.strip())

        # 3. MIS AVERIAS
        elif cmd_clean == "mis" or texto_norm in ["mis averias", "mis_averias", "mis cases", "mis tickets", "mis averia"]:
            query = Averia.query.filter_by(estado="PENDIENTE")
            if operador.branch != "ALL":
                query = query.filter_by(branch=operador.branch)
                
            resultados = query.order_by(Averia.dias_pendientes.desc().nullslast()).limit(5).all()
            if not resultados:
                enviar_mensaje(telefono, f"📋 No tienes averías pendientes registradas en tu sede ({operador.branch}).")
                return
                
            msg = f"📋 *Averías pendientes más críticas ({operador.branch}):*\n\n"
            for av in resultados:
                maps_link = f"https://www.google.com/maps/search/?api=1&query={av.coordenadas}" if av.coordenadas else "Sin coordenadas"
                msg += f"• *Cuenta*: {av.cuenta} | *Caja*: {av.caja}\n"
                msg += f"  *Detalles*: {av.detalles or 'Sin detalles'}\n"
                msg += f"  *Mapa*: {maps_link}\n\n"
            msg += "💡 Para cerrar una avería, envía:\n`reparar <cuenta>`"
            enviar_mensaje(telefono, msg.strip())

        # 4. REPARAR / CERRAR AVERIA
        elif cmd_clean in ["reparar", "rep", "solucionar", "sol", "arreglar", "cerrar", "r"] or texto_norm.startswith("reparar ") or texto_norm.startswith("solucionar ") or texto_norm.startswith("rep "):
            cuenta_req = rest_text if rest_text else (texto_original[8:].strip() if texto_normalizado.startswith("reparar ") else "")
            if not cuenta_req:
                enviar_mensaje(telefono, "💡 *Uso*: escribe `reparar <cuenta>` (ej: `reparar 15_gftth_albertj10`)")
                return
                
            query = Averia.query.filter_by(cuenta=cuenta_req, estado="PENDIENTE")
            if operador.branch != "ALL":
                query = query.filter_by(branch=operador.branch)
                
            averia = query.first()
            if not averia:
                enviar_mensaje(telefono, f"❌ No se encontró ninguna avería pendiente para la cuenta '{cuenta_req}' en tu sede ({operador.branch}).")
                return
                
            # Crear estado conversacion
            nuevo_estado = EstadoConversacion(
                telefono=telefono,
                averia_id=averia.id,
                paso_actual="esperando_cable"
            )
            db.session.add(nuevo_estado)
            db.session.commit()
            
            registrar_conversacion(averia.id, "bot", "Iniciando flujo de reparación")
            
            msg = f"🛠️ *Flujo de Reparación iniciado* para la cuenta `{averia.cuenta}`.\n\n🔌 *Paso 1/5*: ¿Cuántos metros de cable drop utilizaste? (Responde con número entero, ej. `100`, o `0` si ninguno)."
            enviar_mensaje(telefono, msg)

        # 5. CREAR AVERIA MANUAL
        elif cmd_clean in ["crear", "nuevo", "nueva", "n"] or texto_norm in ["crear averia", "crear caso", "crear_averia"]:
            # Verificar que sea de una sede manual/provincias, o admin
            if operador.branch in ["LI1", "LI2", "LI3", "LI4", "LI7"] and operador.rol != "admin":
                enviar_mensaje(telefono, "❌ Tu sede no tiene habilitado el registro manual de averías.")
                return
                
            # Inicializar avería temporal vacía
            import uuid
            nueva = Averia(
                branch=operador.branch if operador.branch != "ALL" else "ARE",
                cuenta=f"SIN_CUENTA_TEMP_{uuid.uuid4().hex[:6].upper()}", # Cuenta temporal
                estado="PENDIENTE",
                origen="MANUAL",
                detalles="En proceso de registro..."
            )
            db.session.add(nueva)
            db.session.commit()
            
            # Crear estado
            nuevo_estado = EstadoConversacion(
                telefono=telefono,
                averia_id=nueva.id,
                paso_actual="crear_esperando_cuenta"
            )
            db.session.add(nuevo_estado)
            db.session.commit()
            
            enviar_mensaje(
                telefono,
                "📝 *Crear Nueva Avería Manual*\n"
                "*(Escribe 'cancelar' en cualquier momento para abortar)*\n\n"
                "🔌 *Paso 1/7 (Cuenta)*: Escribe la cuenta del cliente (ej: `15_gftth_juan`) o responde `ninguna` si no aplica:"
            )

        # 6. RESUMEN / METRICAS
        elif cmd_clean in ["resumen", "res", "estado", "status", "info"]:
            stats = obtener_estadisticas(branch=operador.branch, es_admin=(operador.branch == "ALL"))
            msg = (
                f"📊 *Resumen de Averías ({operador.branch})*:\n\n"
                f"• *Pendientes*: {stats['pendientes']}\n"
                f"• *Reparadas*: {stats['reparados']}\n"
                f"• *Total*: {stats['totales']}\n\n"
                f"🌐 Ver mapa interactivo:\nhttps://botvip-iz55.onrender.com"
            )
            enviar_mensaje(telefono, msg)

        # 7. AYUDA / MENU POR DEFECTO
        else:
            # Menu de ayuda
            menu_help = (
                f"👋 ¡Hola, *{operador.nombre}*!\n"
                f"Sede: *{operador.branch}*\n\n"
                f"⚙️ *Comandos rápidos disponibles* (sin '/'):\n"
                f"🔍 `buscar <texto>` - Buscar averías por caja o cuenta\n"
                f"📋 `todos` - Ver todos los pendientes de tu sede\n"
                f"📋 `mis` - Ver las 5 averías más antiguas/urgentes\n"
                f"🔧 `reparar <cuenta>` - Registrar solución y materiales\n"
                f"📊 `resumen` - Ver estadísticas de tu sede\n"
            )
            if operador.branch not in ["LI1", "LI2", "LI3", "LI4", "LI7"] or operador.rol == "admin":
                menu_help += f"📝 `crear` - Registrar nueva avería manual\n"
            
            menu_help += f"\n💡 *En cualquier flujo*, escribe `cancelar` para volver aquí.\n"
            menu_help += f"🌐 Portal Web: https://botvip-iz55.onrender.com"
            enviar_mensaje(telefono, menu_help)


with app.app_context():
    try:
        db.create_all()
        asegurar_esquema()
        crear_operador_defecto()
        print("Base de datos y esquemas inicializados correctamente.")
    except Exception as e:
        print("Error inicializando base de datos en arranque:", e)


if __name__ == "__main__":
    app.run(port=5000)
