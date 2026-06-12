from flask_sqlalchemy import SQLAlchemy

db = SQLAlchemy()


class Operador(db.Model):
    __tablename__ = "operadores"

    id = db.Column(db.Integer, primary_key=True)
    nombre = db.Column(db.String(100), nullable=False)
    dni = db.Column(db.String(20), unique=True, nullable=False)
    correo = db.Column(db.String(150), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    rol = db.Column(db.String(20), default="operador")  # admin / operador
    branch = db.Column(db.String(30), default="ALL")    # LI1, ARE, ALL, etc.
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


class Averia(db.Model):
    __tablename__ = "averias"

    id = db.Column(db.Integer, primary_key=True)
    branch = db.Column(db.String(50), nullable=False)
    codigo_wo = db.Column(db.String(100))
    cuenta = db.Column(db.String(100), nullable=True)
    detalles = db.Column(db.Text)
    dias_pendientes = db.Column(db.Float)
    estado = db.Column(db.String(50), default="PENDIENTE")  # PENDIENTE / REPARADO
    status_caja = db.Column(db.String(100))
    contrata = db.Column(db.String(100))
    periodo_pendiente = db.Column(db.String(100))
    site = db.Column(db.String(100))
    caja = db.Column(db.String(100))
    coordenadas = db.Column(db.String(100))
    origen = db.Column(db.String(20), default="SHEETS")  # SHEETS / MANUAL
    fecha_creacion = db.Column(db.DateTime, server_default=db.func.current_timestamp())
    fecha_resolucion = db.Column(db.DateTime)
    tecnico_id = db.Column(db.Integer, db.ForeignKey("operadores.id"))

    # Materiales usados
    material_cable_m = db.Column(db.Integer, default=0)
    material_conectores = db.Column(db.Integer, default=0)
    material_rosetas = db.Column(db.Integer, default=0)
    material_mangas = db.Column(db.Integer, default=0)
    material_acopladores = db.Column(db.Integer, default=0)
    material_comentarios = db.Column(db.Text)
    materiales_json = db.Column(db.Text)

    # Relación
    tecnico = db.relationship("Operador", backref="averias_resueltas", foreign_keys=[tecnico_id])

    @property
    def materiales_dict(self):
        import json
        if self.materiales_json:
            try:
                return json.loads(self.materiales_json)
            except Exception:
                pass
        
        # Fallback para compatibilidad con registros antiguos
        res = {}
        if self.material_cable_m:
            res["299381|Cable de fibra óptica 1 hilo - Drop"] = self.material_cable_m
        if self.material_conectores:
            res["299378|Conector rápido SC/APC"] = self.material_conectores
        if self.material_rosetas:
            res["Sin Código|Roseta Óptica"] = self.material_rosetas
        if self.material_mangas:
            res["299799|Caja de empalme tipo domo 24Fo, 8 puertos"] = self.material_mangas
        if self.material_acopladores:
            res["Sin Código|Acoplador Óptico"] = self.material_acopladores
        return res



