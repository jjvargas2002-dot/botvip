from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Operador(db.Model):
    __tablename__ = "operadores"

    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    dni = db.Column(db.String(20), unique=True, nullable=False)
    correo = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    rol = db.Column(db.String(20), default="operador")
    activo = db.Column(db.Boolean, default=True)
    fecha_creacion = db.Column(db.DateTime, server_default=db.func.current_timestamp())


class Cliente(db.Model):
    __tablename__ = "clientes"

    id = db.Column(db.Integer, primary_key=True)
    codigo_cliente = db.Column(db.String(100), unique=True, nullable=False)
    telefono = db.Column(db.String(20), nullable=False)
    nombre = db.Column(db.String(200))
    numero_referencia = db.Column(db.String(100))
    direccion = db.Column(db.Text)
    referencia = db.Column(db.Text)
    fecha_creacion = db.Column(db.DateTime, server_default=db.func.current_timestamp())


class Caso(db.Model):
    __tablename__ = "casos"

    id = db.Column(db.Integer, primary_key=True)
    cliente_id = db.Column(db.Integer, db.ForeignKey("clientes.id"), nullable=False)
    operador_id = db.Column(db.Integer, db.ForeignKey("operadores.id"))
    tipo = db.Column(db.String(20), nullable=False)  # INSTALACION / AVERIA
    estado = db.Column(db.String(30), default="PENDIENTE")
    problema = db.Column(db.Text)
    fecha_disponible = db.Column(db.Date)
    hora_disponible = db.Column(db.String(100))
    tecnico_nombre = db.Column(db.String(100))
    tecnico_telefono = db.Column(db.String(20))
    tecnico_estado = db.Column(db.String(30), server_default="PENDIENTE")
    fecha_creacion = db.Column(db.DateTime, server_default=db.func.current_timestamp())
    fecha_cierre = db.Column(db.DateTime)


class Conversacion(db.Model):
    __tablename__ = "conversaciones"

    id = db.Column(db.Integer, primary_key=True)
    caso_id = db.Column(db.Integer, db.ForeignKey("casos.id"))
    remitente = db.Column(db.String(20))
    mensaje = db.Column(db.Text)
    fecha = db.Column(db.DateTime, server_default=db.func.current_timestamp())


class EstadoConversacion(db.Model):
    __tablename__ = "estados_conversacion"

    telefono = db.Column(db.String(20), primary_key=True)
    caso_id = db.Column(db.Integer, db.ForeignKey("casos.id"))
    paso_actual = db.Column(db.String(50))
    fecha_actualizacion = db.Column(
        db.DateTime,
        server_default=db.func.current_timestamp(),
        onupdate=db.func.current_timestamp(),
    )
