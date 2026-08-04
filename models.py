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
    tipificacion = db.Column(db.String(100), nullable=True)
    cuentas_asociadas = db.Column(db.Text, nullable=True)
    resolucion_fuente = db.Column(db.String(10), nullable=True)  # APP / SHEET

    # Relación
    tecnico = db.relationship("Operador", backref="averias_resueltas", foreign_keys=[tecnico_id])

    @property
    def tipo_reparacion(self):
        """Clasifica la reparación en AGRUPADO / PRINCIPAL_APP / PRINCIPAL_SHEET."""
        if self.estado != "REPARADO":
            return ""
        if self.material_comentarios and "Agrupado en la avería principal" in self.material_comentarios:
            return "AGRUPADO"
        if self.resolucion_fuente == "APP":
            return "PRINCIPAL_APP"
        if self.resolucion_fuente == "SHEET":
            return "PRINCIPAL_SHEET"
        # Registros antiguos sin resolucion_fuente: inferir por el comentario
        comentario = (self.material_comentarios or "").lower()
        if "marcado como resuelto en el drive" in comentario or "cerrado automáticamente" in comentario:
            return "PRINCIPAL_SHEET"
        return "PRINCIPAL_APP"

    @property
    def tipo_reparacion_label(self):
        return {
            "AGRUPADO": "Agrupado",
            "PRINCIPAL_APP": "Principal: Fuente app",
            "PRINCIPAL_SHEET": "Principal: Fuente sheet",
        }.get(self.tipo_reparacion, "")

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
            res["283866|Cable Drop"] = self.material_cable_m
        if self.material_conectores:
            res["299378|FAC"] = self.material_conectores
        if self.material_rosetas:
            res["299379|Waterproof"] = self.material_rosetas
        if self.material_mangas:
            res["299799|Mufas"] = self.material_mangas
        if self.material_acopladores:
            res["294790|Preconectorizado"] = self.material_acopladores
        return res

    @property
    def wos_detalle(self):
        """
        Parsea los códigos de WO (uno o múltiples separados por saltos de línea, comas o espacios).
        Extrae la fecha de creación de cada código (formato YYYYMMDD ej: 20260707 -> 07/07/2026)
        y calcula los días transcurridos para cada WO.
        Retorna la lista ordenada desde la WO más antigua a la más reciente.
        """
        import re
        from datetime import datetime, date
        
        if not self.codigo_wo or self.codigo_wo.strip() in ["", "N/A", "Sin WO", "None"]:
            return []
            
        raw_parts = [p.strip() for p in re.split(r'[\r\n,;]+', self.codigo_wo) if p.strip()]
        if not raw_parts:
            return []
            
        # Determinar fecha base para cálculo de días
        try:
            from app import obtener_hora_peru
            hoy = obtener_hora_peru().date()
        except Exception:
            hoy = datetime.now().date()
            
        items = []
        for code in raw_parts:
            # Buscar fecha YYYYMMDD en el código de WO (ej. WO_SPM_20260707_170732631)
            match = re.search(r'(?:^|[^0-9])(20\d{2})(\d{2})(\d{2})(?:[^0-9]|$)', code)
            fecha_dt = None
            fecha_str = ""
            dias = 0
            
            if match:
                try:
                    y, m, d = int(match.group(1)), int(match.group(2)), int(match.group(3))
                    fecha_dt = date(y, m, d)
                    fecha_str = f"{d:02d}/{m:02d}/{y}"
                    dias = max(0, (hoy - fecha_dt).days)
                except Exception:
                    pass
                    
            if not fecha_dt and self.fecha_creacion:
                fecha_dt = self.fecha_creacion.date()
                fecha_str = fecha_dt.strftime("%d/%m/%Y")
                dias = max(0, (hoy - fecha_dt).days)
                
            items.append({
                "codigo": code,
                "fecha_str": fecha_str,
                "fecha_dt": fecha_dt,
                "dias": dias
            })
            
        # Ordenar de la WO más antigua a la más reciente (fecha más temprana primero)
        items.sort(key=lambda x: (x["fecha_dt"] if x["fecha_dt"] else date.max, -x["dias"]))
        return items

    @property
    def codigo_wo_ordenado(self):
        """Retorna los códigos de WO ordenados de la más antigua a la más reciente."""
        detalles = self.wos_detalle
        if not detalles:
            return self.codigo_wo or ""
        return "\n".join([d["codigo"] for d in detalles])


class StockBranch(db.Model):
    __tablename__ = "stock_branch"

    id = db.Column(db.Integer, primary_key=True)
    branch = db.Column(db.String(50), nullable=False)
    material_codigo = db.Column(db.String(50), nullable=False)
    material_nombre = db.Column(db.String(255), nullable=False)
    stock_actual = db.Column(db.Integer, default=0)
    stock_enviado_noc = db.Column(db.Integer, default=0)
    fecha_envio_noc = db.Column(db.Date, nullable=True)
    fecha_actualizacion = db.Column(db.DateTime, server_default=db.func.current_timestamp(), onupdate=db.func.current_timestamp())

    __table_args__ = (db.UniqueConstraint("branch", "material_nombre", name="uq_branch_material_nombre"),)


class Box(db.Model):
    __tablename__ = "boxes"

    id = db.Column(db.Integer, primary_key=True)
    caja = db.Column(db.String(100), unique=True, nullable=False, index=True)
    latitud = db.Column(db.String(50))
    longitud = db.Column(db.String(50))




