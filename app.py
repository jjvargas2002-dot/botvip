from functools import wraps
from io import BytesIO

from flask import Flask, render_template, request, send_file, redirect, url_for, session, flash
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from sqlalchemy import text
from werkzeug.security import generate_password_hash, check_password_hash
from config import Config
from models import db, Cliente, Caso, Conversacion, EstadoConversacion, Operador
from whatsapp import enviar_mensaje

app = Flask(__name__)
app.config.from_object(Config)
db.init_app(app)

with app.app_context():
    try:
        db.create_all()
        asegurar_esquema()
        crear_operador_defecto()
        print("Base de datos y esquemas inicializados correctamente.")
    except Exception as e:
        print("Error inicializando base de datos en arranque:", e)


def login_requerido(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if "operador_id" not in session:
            flash("Por favor, inicia sesión para acceder al sistema.", "warning")
            return redirect(url_for("login"))
        return f(*args, **kwargs)
    return decorated_function


def obtener_estadisticas(operador_id=None, es_admin=False):
    try:
        query_total = Caso.query
        query_pendientes = Caso.query.filter_by(estado="PENDIENTE")
        query_asignados = Caso.query.filter_by(estado="PENDIENTE_ASIGNACION")
        query_cerrados = Caso.query.filter_by(estado="CERRADO")

        if not es_admin and operador_id:
            query_total = query_total.filter_by(operador_id=operador_id)
            query_pendientes = query_pendientes.filter_by(operador_id=operador_id)
            query_asignados = query_asignados.filter_by(operador_id=operador_id)
            query_cerrados = query_cerrados.filter_by(operador_id=operador_id)

        return {
            "totales": query_total.count(),
            "pendientes": query_pendientes.count(),
            "asignados": query_asignados.count(),
            "cerrados": query_cerrados.count(),
        }
    except Exception as e:
        print("Error obteniendo estadísticas:", e)
        return {
            "totales": 0,
            "pendientes": 0,
            "asignados": 0,
            "cerrados": 0,
        }


def registrar_conversacion(caso_id, remitente, mensaje):
    conversacion = Conversacion(
        caso_id=caso_id,
        remitente=remitente,
        mensaje=mensaje,
    )
    db.session.add(conversacion)


def mensaje_inicio(caso):
    if caso.tipo == "INSTALACION":
        return (
            "👋 ¡Hola! Te saludamos de *FTTH VIP*.\n\n"
            "Hemos recibido tu solicitud de instalación. Para ayudarte más rápido, por favor indícanos tu nombre completo."
        )

    return (
        "👋 ¡Hola! Te saludamos de *FTTH VIP*.\n\n"
        "Hemos recibido tu reporte de avería. Lamentamos el inconveniente. Para ayudarte más rápido, por favor indícanos tu nombre completo."
    )


def asegurar_esquema():
    ajustes = [
        "ALTER TABLE clientes DROP CONSTRAINT IF EXISTS clientes_telefono_key",
    ]

    columnas = [
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS dni VARCHAR(20) UNIQUE",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS rol VARCHAR(20) DEFAULT 'operador'",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS activo BOOLEAN DEFAULT TRUE",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        "ALTER TABLE clientes ADD COLUMN IF NOT EXISTS numero_referencia VARCHAR(100)",
        "ALTER TABLE clientes ADD COLUMN IF NOT EXISTS fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        "ALTER TABLE casos ADD COLUMN IF NOT EXISTS problema TEXT",
        "ALTER TABLE casos ADD COLUMN IF NOT EXISTS fecha_disponible DATE",
        "ALTER TABLE casos ADD COLUMN IF NOT EXISTS hora_disponible VARCHAR(100)",
        "ALTER TABLE casos ADD COLUMN IF NOT EXISTS tecnico_nombre VARCHAR(100)",
        "ALTER TABLE casos ADD COLUMN IF NOT EXISTS tecnico_telefono VARCHAR(20)",
        "ALTER TABLE casos ADD COLUMN IF NOT EXISTS tecnico_estado VARCHAR(30) DEFAULT 'PENDIENTE'",
        "ALTER TABLE casos ADD COLUMN IF NOT EXISTS fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        "ALTER TABLE casos ADD COLUMN IF NOT EXISTS fecha_cierre TIMESTAMP",
        "ALTER TABLE estados_conversacion ADD COLUMN IF NOT EXISTS fecha_actualizacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
    ]

    for consulta in ajustes:
        db.session.execute(text(consulta))

    for consulta in columnas:
        try:
            db.session.execute(text(consulta))
        except Exception as e:
            print(f"Error ejecutando columna {consulta}: {e}")
            db.session.rollback()

    db.session.commit()


def crear_operador_defecto():
    try:
        # Buscar si ya existe por DNI
        admin = Operador.query.filter_by(dni="FBB").first()
        if admin:
            # Forzar la contraseña a "Bitel@123" y rol a admin para asegurar el acceso
            admin.password_hash = generate_password_hash("Bitel@123")
            admin.rol = "admin"
            admin.correo = "admin@botvip.com"
            db.session.commit()
            print("Operador administrador FBB existente actualizado con contraseña Bitel@123.")
            return

        # Si no existe por DNI FBB, buscar si existe por el correo único (sea el viejo bitel o el nuevo vip)
        admin_correo = Operador.query.filter(Operador.correo.in_(["admin@botbitel.com", "admin@botvip.com"])).first()
        if admin_correo:
            # Actualizar el operador existente
            admin_correo.dni = "FBB"
            admin_correo.nombre = "Administrador FBB"
            admin_correo.correo = "admin@botvip.com"
            admin_correo.password_hash = generate_password_hash("Bitel@123")
            admin_correo.rol = "admin"
            db.session.commit()
            print("Operador administrador actualizado a DNI FBB con contraseña Bitel@123.")
            return

        # Si no existe ninguno de los dos, crear uno nuevo
        default_pwd = generate_password_hash("Bitel@123")
        admin = Operador(
            nombre="Administrador FBB",
            dni="FBB",
            correo="admin@botvip.com",
            password_hash=default_pwd,
            rol="admin",
            activo=True
        )
        db.session.add(admin)
        db.session.commit()
        print("Operador administrador por defecto (FBB) creado con contraseña Bitel@123.")
    except Exception as e:
        print("Error al crear operador por defecto:", e)
        db.session.rollback()


@app.route("/login", methods=["GET", "POST"])
def login():
    if "operador_id" in session:
        return redirect(url_for("dashboard"))

    if request.method == "POST":
        dni = request.form["dni"].strip()
        password = request.form["password"].strip()

        operador = Operador.query.filter_by(dni=dni, activo=True).first()
        if operador and check_password_hash(operador.password_hash, password):
            session["operador_id"] = operador.id
            session["operador_nombre"] = operador.nombre
            session["operador_dni"] = operador.dni
            session["operador_rol"] = operador.rol
            
            if password == dni:
                flash("Por seguridad, te sugerimos cambiar tu contraseña actual (que es igual a tu DNI).", "warning")
            else:
                flash(f"¡Bienvenido de nuevo, {operador.nombre}!", "success")
            return redirect(url_for("dashboard"))
        else:
            flash("DNI o contraseña incorrectos.", "danger")

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

        # Validar si ya existe por DNI o correo
        existente = Operador.query.filter((Operador.dni == dni) | (Operador.correo == correo)).first()
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
        flash("Acceso denegado: Se requieren permisos de administrador para ver esta sección.", "danger")
        return redirect(url_for("dashboard"))

    operadores = Operador.query.order_by(Operador.id.desc()).all()
    stats = obtener_estadisticas(operador_id=session.get("operador_id"), es_admin=True)
    return render_template("operadores.html", operadores=operadores, stats=stats)


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


@app.route("/", methods=["GET", "POST"])
@login_requerido
def dashboard():
    es_admin = session.get("operador_rol") == "admin"
    stats = obtener_estadisticas(operador_id=session.get("operador_id"), es_admin=es_admin)
    if request.method == "POST":
        telefono = request.form["telefono"].strip()
        codigo = request.form["codigo_cliente"].strip()
        tipo = request.form["tipo"].strip()
        tecnico_telefono = request.form.get("tecnico_telefono", "").strip()

        cliente = Cliente.query.filter_by(codigo_cliente=codigo).first()

        if not cliente:
            cliente = Cliente(
                codigo_cliente=codigo,
                telefono=telefono,
            )
            db.session.add(cliente)
            db.session.commit()
        else:
            cliente.telefono = telefono

        caso = Caso(
            cliente_id=cliente.id,
            operador_id=session["operador_id"],
            tipo=tipo,
            estado="PENDIENTE",
            tecnico_telefono=tecnico_telefono or None,
        )

        db.session.add(caso)
        db.session.commit()

        estado = EstadoConversacion.query.filter_by(telefono=telefono).first()

        if estado:
            estado.caso_id = caso.id
            estado.paso_actual = "esperando_nombre"
        else:
            estado = EstadoConversacion(
                telefono=telefono,
                caso_id=caso.id,
                paso_actual="esperando_nombre",
            )
            db.session.add(estado)

        msg = mensaje_inicio(caso)
        registrar_conversacion(caso.id, "bot", msg)
        db.session.commit()

        enviar_mensaje(telefono, msg)

        stats = obtener_estadisticas(operador_id=session.get("operador_id"), es_admin=es_admin)
        return render_template("dashboard.html", mensaje="Caso creado y notificado", stats=stats)

    return render_template("dashboard.html", stats=stats)


@app.route("/casos")
@login_requerido
def listar_casos():
    es_admin = session.get("operador_rol") == "admin"
    query = db.session.query(Caso, Cliente).join(Cliente, Caso.cliente_id == Cliente.id)
    
    if not es_admin:
        query = query.filter(Caso.operador_id == session["operador_id"])
        
    casos = query.order_by(Caso.id.desc()).all()
    stats = obtener_estadisticas(operador_id=session.get("operador_id"), es_admin=es_admin)
    return render_template("casos.html", casos=casos, stats=stats)


@app.route("/casos/exportar")
@login_requerido
def exportar_casos():
    es_admin = session.get("operador_rol") == "admin"
    query = db.session.query(Caso, Cliente).join(Cliente, Caso.cliente_id == Cliente.id)
    
    if not es_admin:
        query = query.filter(Caso.operador_id == session["operador_id"])
        
    casos = query.order_by(Caso.id.desc()).all()

    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Casos FTTH"

    headers = [
        "ID",
        "Codigo cliente",
        "Telefono cliente",
        "Nombre",
        "Numero de referencia",
        "Tipo",
        "Estado",
        "Telefono tecnico",
        "Disponibilidad",
        "Direccion",
        "Referencia",
        "Fecha creacion",
    ]
    sheet.append(headers)

    for caso, cliente in casos:
        sheet.append([
            caso.id,
            cliente.codigo_cliente,
            cliente.telefono,
            cliente.nombre,
            cliente.numero_referencia,
            caso.tipo,
            caso.estado,
            caso.tecnico_telefono,
            caso.hora_disponible,
            cliente.direccion,
            cliente.referencia,
            caso.fecha_creacion.strftime("%Y-%m-%d %H:%M") if caso.fecha_creacion else "",
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
        download_name="casos_ftth.xlsx",
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

        procesar_mensaje(telefono, texto)

    except Exception as e:
        print("Webhook error:", e)

    return "OK", 200


def procesar_mensaje(telefono, texto):
    texto_original = texto.strip()
    texto_normalizado = texto_original.lower()

    estado = EstadoConversacion.query.filter_by(telefono=telefono).first()
    caso = db.session.get(Caso, estado.caso_id) if estado else None
    cliente = db.session.get(Cliente, caso.cliente_id) if caso else None

    # COMANDOS GLOBALES
    if estado and caso:
        if texto_normalizado in ["menu", "reiniciar", "/menu", "/reiniciar"]:
            estado.paso_actual = "esperando_nombre"
            respuesta = "🔄 Entendido, vamos a reiniciar los datos de tu ticket.\n\n" + mensaje_inicio(caso)
            registrar_conversacion(caso.id, "cliente", texto_original)
            registrar_conversacion(caso.id, "bot", respuesta)
            db.session.commit()
            enviar_mensaje(telefono, respuesta)
            return

        if texto_normalizado in ["asesor", "humano", "operador"]:
            estado.paso_actual = "operador"
            respuesta = "📞 De acuerdo, he pausado el bot. En breve un asesor se pondrá en contacto contigo."
            registrar_conversacion(caso.id, "cliente", texto_original)
            registrar_conversacion(caso.id, "bot", respuesta)
            db.session.commit()
            enviar_mensaje(telefono, respuesta)
            return

        if estado.paso_actual == "operador":
            # Si está pausado para operador humano, solo registramos el mensaje y no respondemos automáticamente
            registrar_conversacion(caso.id, "cliente", texto_original)
            db.session.commit()
            return

    if not estado:
        cliente = (
            Cliente.query.filter_by(telefono=telefono)
            .order_by(Cliente.id.desc())
            .first()
        )

        if not cliente:
            cliente = Cliente(codigo_cliente=f"AUTO_{telefono}", telefono=telefono)
            db.session.add(cliente)
            db.session.commit()

        tipo = (
            "AVERIA"
            if any(x in texto_normalizado for x in ["internet", "sin", "caido", "lento"])
            else "INSTALACION"
        )

        caso = Caso(
            cliente_id=cliente.id,
            tipo=tipo,
            estado="PENDIENTE",
        )

        db.session.add(caso)
        db.session.commit()

        estado = EstadoConversacion(
            telefono=telefono,
            caso_id=caso.id,
            paso_actual="esperando_nombre",
        )

        respuesta = mensaje_inicio(caso)
        registrar_conversacion(caso.id, "cliente", texto_original)
        registrar_conversacion(caso.id, "bot", respuesta)
        db.session.add(estado)
        db.session.commit()

        enviar_mensaje(telefono, respuesta)
        return

    if not caso or not cliente:
        respuesta = "No encontramos un caso activo de *FTTH VIP* para este número. Por favor espera a que un operador registre tu solicitud."
        enviar_mensaje(telefono, respuesta)
        return

    registrar_conversacion(caso.id, "cliente", texto_original)

    if estado.paso_actual == "esperando_nombre":
        cliente.nombre = texto_original
        estado.paso_actual = "esperando_numero_referencia"
        respuesta = "✅ Gracias. Ahora indícanos tu número de referencia (código de cliente de *FTTH VIP*)."

    elif estado.paso_actual == "esperando_numero_referencia":
        cliente.numero_referencia = texto_original
        estado.paso_actual = "esperando_direccion"
        respuesta = "📍 Perfecto. Indícanos tu dirección completa."

    elif estado.paso_actual == "esperando_direccion":
        cliente.direccion = texto_original
        estado.paso_actual = "esperando_referencia"
        respuesta = "📌 Gracias. Ahora indícanos una referencia de ubicación."

    elif estado.paso_actual == "esperando_referencia":
        cliente.referencia = texto_original
        estado.paso_actual = "esperando_disponibilidad"
        respuesta = "🕒 ¿Qué día y hora tienes disponibilidad para la atención presencial?"

    elif estado.paso_actual == "esperando_disponibilidad":
        caso.hora_disponible = texto_original
        caso.estado = "PENDIENTE_ASIGNACION"
        estado.paso_actual = "cerrado"
        respuesta = "✅ Listo. Tu ticket de *FTTH VIP* fue registrado correctamente y un técnico será asignado."

        if caso.tecnico_telefono:
            mensaje_tecnico = f"""
*NUEVO TICKET FTTH VIP*

Tipo: {caso.tipo}
Cliente: {cliente.codigo_cliente}
Numero de referencia: {cliente.numero_referencia}
Tel: {cliente.telefono}

Direccion: {cliente.direccion}
Referencia: {cliente.referencia}
Disponibilidad: {caso.hora_disponible}

Estado: {caso.estado}
""".strip()

            registrar_conversacion(caso.id, "bot", mensaje_tecnico)
            enviar_mensaje(caso.tecnico_telefono, mensaje_tecnico)

    else:
        respuesta = "Tu ticket de *FTTH VIP* ya se encuentra registrado y en proceso. Estamos dando seguimiento a tu caso."

    registrar_conversacion(caso.id, "bot", respuesta)
    db.session.commit()

    enviar_mensaje(telefono, respuesta)


if __name__ == "__main__":
    app.run(port=5000)
