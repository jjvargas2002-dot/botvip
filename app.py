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
from models import db, Cliente, Caso, Operador, Averia

app = Flask(__name__)
app.config.from_object(Config)
db.init_app(app)

MATERIALES_MASTER = [
    # Sección OLT
    {"codigo": "291368", "nombre": "OLT GPON ZTE C610 - 16 puertos DC", "seccion": "OLT"},
    {"codigo": "291368", "nombre": "Cable de alimentación OC para chasis C610 (15m)", "seccion": "OLT"},
    {"codigo": "291368", "nombre": "Cable de puesta a tierra de protección (10m)", "seccion": "OLT"},
    {"codigo": "294701", "nombre": "Transceptor óptico bidireccional monofibra SFP GPON-OLT clase C+", "seccion": "OLT"},
    {"codigo": "283866", "nombre": "Módulo óptico bidireccional de doble fibra y canal único ZTE 10Km/1.25Gb", "seccion": "OLT"},
    # Sección ODF
    {"codigo": "294702", "nombre": "Patch Cord SC/UPC-SC/UPC 3m", "seccion": "ODF"},
    {"codigo": "89", "nombre": "Patch Cord LC/UPC-LC/UPC 3m", "seccion": "ODF"},
    {"codigo": "87", "nombre": "Patch Cord LC/UPC-LC/UPC 15m", "seccion": "ODF"},
    {"codigo": "86", "nombre": "Patch Cord LC/UPC-LC/UPC 10m", "seccion": "ODF"},
    {"codigo": "267990", "nombre": "ODF 48 puertos SC/UPC", "seccion": "ODF"},
    {"codigo": "2925", "nombre": "ODF 24 puertos SC/UPC", "seccion": "ODF"},
    {"codigo": "265191", "nombre": "Gabinete exterior para OLT y ODF", "seccion": "ODF"},
    # Sección CAJAS
    {"codigo": "299798", "nombre": "Caja de empalme tipo domo 48Fo, 8 puertos", "seccion": "CAJAS"},
    {"codigo": "299799", "nombre": "Caja de empalme tipo domo 24Fo, 8 puertos", "seccion": "CAJAS"},
    {"codigo": "424", "nombre": "Caja de empalme plana 24Fo, 4 puertos", "seccion": "CAJAS"},
    {"codigo": "423", "nombre": "Caja de empalme plana 12Fo, 4 puertos", "seccion": "CAJAS"},
    {"codigo": "8810", "nombre": "Caja de empalme plana 4Fo, 3 puertos", "seccion": "CAJAS"},
    {"codigo": "262079", "nombre": "Caja de empalme plana 1Fo, 2 puertos, empalme mecánico", "seccion": "CAJAS"},
    {"codigo": "300476", "nombre": "ODB, 11 puertos para adaptadores (HUB-BOX 70/30)", "seccion": "CAJAS"},
    {"codigo": "300477", "nombre": "ODB, 11 puertos para adaptadores (SUB-BOX 70/30)", "seccion": "CAJAS"},
    {"codigo": "300478", "nombre": "ODB, 11 puertos para adaptadores (SUB-BOX 50/50)", "seccion": "CAJAS"},
    {"codigo": "300479", "nombre": "ODB, 10 puertos para adaptadores (END-BOX 50/50)", "seccion": "CAJAS"},
    {"codigo": "298476", "nombre": "ODB, 9 puertos para adaptadores (EXP-BOX 50/50)", "seccion": "CAJAS"},
    {"codigo": "298912", "nombre": "ODB, 9 puertos para adaptadores (Edificio)", "seccion": "CAJAS"},
    {"codigo": "296344", "nombre": "ZTE_ODB, 2 puertos de fibra para adaptadores (Caja de unión)", "seccion": "CAJAS"},
    {"codigo": "294765", "nombre": "ZTE_ODB, 8 puertos MPO y 4 puertos de fibra para adaptadores (X-BOX)", "seccion": "CAJAS"},
    {"codigo": "294764", "nombre": "ZTE_ODB, 5 puertos para adaptadores (HUB-BOX)", "seccion": "CAJAS"},
    {"codigo": "294767", "nombre": "ZTE_ODB, 11 puertos para adaptadores (SUB-BOX 70/30)", "seccion": "CAJAS"},
    {"codigo": "294768", "nombre": "ZTE_ODB, 11 puertos para adaptadores (SUB-BOX 50/50)", "seccion": "CAJAS"},
    {"codigo": "294766", "nombre": "ZTE_ODB, 10 puertos para adaptadores (END-BOX 50/50)", "seccion": "CAJAS"},
    {"codigo": "296345", "nombre": "ZTE_ODB, 9 puertos para adaptadores (EXP-BOX 50/50)", "seccion": "CAJAS"},
    # Sección de Conectores
    {"codigo": "299379", "nombre": "Conector de campo impermeable para caja ZTE", "seccion": "CONECTORES"},
    {"codigo": "299378", "nombre": "Conector rápido SC/APC", "seccion": "CONECTORES"},
    {"codigo": "Sin Código", "nombre": "Splitter PLC 1x8 (Pigtail PLC 1x8)", "seccion": "CONECTORES"},
    # Sección de Cables
    {"codigo": "43978", "nombre": "Cable de fibra óptica ADSS 48 fibras, vano 100m", "seccion": "CABLES"},
    {"codigo": "63004", "nombre": "Cable de fibra óptica ADSS 24 fibras, vano 200m", "seccion": "CABLES"},
    {"codigo": "67", "nombre": "Cable de fibra óptica ADSS 24 fibras, vano 100m", "seccion": "CABLES"},
    {"codigo": "299344", "nombre": "Cable de fibra óptica ASU 12Fo, vano 100m", "seccion": "CABLES"},
    {"codigo": "300378", "nombre": "Cable de fibra óptica ASU 4Fo, vano 100m", "seccion": "CABLES"},
    {"codigo": "9266", "nombre": "Cable de fibra óptica ASU 4Fo, vano 100m - Flexible", "seccion": "CABLES"},
    {"codigo": "300379", "nombre": "Cable de fibra óptica ASU 1Fo, vano 100m", "seccion": "CABLES"},
    {"codigo": "294795", "nombre": "Fibra óptica preconectorizada de 12 núcleos 100m MPO/APC", "seccion": "CABLES"},
    {"codigo": "294796", "nombre": "Fibra óptica preconectorizada de 12 núcleos 200m MPO/APC", "seccion": "CABLES"},
    {"codigo": "296347", "nombre": "Fibra óptica preconectorizada de 12 núcleos 300m MPO/APC", "seccion": "CABLES"},
    {"codigo": "294797", "nombre": "Fibra óptica preconectorizada de 12 núcleos 500m MPO/APC", "seccion": "CABLES"},
    {"codigo": "294798", "nombre": "Fibra óptica preconectorizada de 12 núcleos 600m MPO/APC", "seccion": "CABLES"},
    {"codigo": "294799", "nombre": "Fibra óptica preconectorizada de 12 núcleos 700m MPO/APC", "seccion": "CABLES"},
    {"codigo": "294790", "nombre": "Fibra óptica preconectorizada de 1 núcleo 50m SC/APC", "seccion": "CABLES"},
    {"codigo": "294791", "nombre": "Fibra óptica preconectorizada de 1 núcleo 100m SC/APC", "seccion": "CABLES"},
    {"codigo": "294792", "nombre": "Fibra óptica preconectorizada de 1 núcleo 150m SC/APC", "seccion": "CABLES"},
    {"codigo": "294793", "nombre": "Fibra óptica preconectorizada de 1 núcleo 200m SC/APC", "seccion": "CABLES"},
    {"codigo": "294794", "nombre": "Fibra óptica preconectorizada de 1 núcleo 250m SC/APC", "seccion": "CABLES"},
    {"codigo": "295995", "nombre": "Fibra óptica preconectorizada de 1 núcleo 300m SC/APC", "seccion": "CABLES"},
    {"codigo": "299381", "nombre": "Cable de fibra óptica 1 hilo - Drop", "seccion": "CABLES"},
    # Sección de Accesorios
    {"codigo": "296348", "nombre": "ZTE_Patch Cord SC/APC 7m", "seccion": "ACCESORIOS"},
    {"codigo": "306", "nombre": "Grapa de tensión para vano de 200m", "seccion": "ACCESORIOS"},
    {"codigo": "305", "nombre": "Grapa de tensión para vano de 100m", "seccion": "ACCESORIOS"},
    {"codigo": "313", "nombre": "Grapa de suspensión para vano de 100m", "seccion": "ACCESORIOS"},
    {"codigo": "350", "nombre": "Abrazaderas para cable OPGW", "seccion": "ACCESORIOS"},
    {"codigo": "298983", "nombre": "Retención preformada plástica para cable de 5mm-8mm", "seccion": "ACCESORIOS"},
    {"codigo": "294929", "nombre": "Retención preformada para cable de 6.8mm (Cable MPO)", "seccion": "ACCESORIOS"},
    {"codigo": "294928", "nombre": "Retención preformada para cable de 5mm (Cable preconectorizado)", "seccion": "ACCESORIOS"},
    {"codigo": "299380", "nombre": "Templador para cable Drop", "seccion": "ACCESORIOS"},
    {"codigo": "28497", "nombre": "Grillete tipo D", "seccion": "ACCESORIOS"},
    {"codigo": "294898", "nombre": "Grillete tipo trébol", "seccion": "ACCESORIOS"},
    {"codigo": "285043", "nombre": "Brazo de soporte 1.0m", "seccion": "ACCESORIOS"},
    {"codigo": "295136", "nombre": "Brazo de soporte 0.6m", "seccion": "ACCESORIOS"},
    {"codigo": "24941", "nombre": "Cruceta 60cm", "seccion": "ACCESORIOS"},
    {"codigo": "295280", "nombre": "Cable de acero 4mm, con recubrimiento PVC", "seccion": "ACCESORIOS"},
    {"codigo": "295360", "nombre": "Abrazadera colgante dieléctrica", "seccion": "ACCESORIOS"},
    {"codigo": "20910", "nombre": "Candado de acero 3/8\" para cable de acero", "seccion": "ACCESORIOS"},
    {"codigo": "8332", "nombre": "Cinta Bandit + juego de hebillas", "seccion": "ACCESORIOS"},
    {"codigo": "294899", "nombre": "Fleje de acero para postes con diámetros de 100-200mm", "seccion": "ACCESORIOS"},
    {"codigo": "295660", "nombre": "Cinta Bandit", "seccion": "ACCESORIOS"},
    {"codigo": "295661", "nombre": "Juego de hebillas", "seccion": "ACCESORIOS"},
    {"codigo": "Sin Código", "nombre": "Etiqueta para cable de fibra óptica", "seccion": "ACCESORIOS"},
    {"codigo": "9610", "nombre": "Tubo corrugado de PVC ignífugo 25mm", "seccion": "ACCESORIOS"},
    {"codigo": "Sin Código", "nombre": "Abrazadera tipo oreja para cable", "seccion": "ACCESORIOS"},
    {"codigo": "Sin Código", "nombre": "Cintillo para cable 4x200mm", "seccion": "ACCESORIOS"}
]

@app.context_processor
def utility_processor():
    materiales_por_seccion = {}
    for mat in MATERIALES_MASTER:
        sec = mat["seccion"]
        if sec not in materiales_por_seccion:
            materiales_por_seccion[sec] = []
        materiales_por_seccion[sec].append(mat)
    return dict(materiales_por_seccion=materiales_por_seccion, MATERIALES_MASTER=MATERIALES_MASTER)

def sincronizar_sites():
    url_sites = "https://docs.google.com/spreadsheets/d/1eaNxCpm8JF1JcZS3_ldwMRINGYFaW6RsQQWybvRi_P8/export?format=csv&gid=894046404"
    try:
        response = requests.get(url_sites, timeout=15)
        if response.status_code != 200:
            return False, f"Error de conexión con pestaña SITES (Status: {response.status_code})"
        
        content = response.content.decode('utf-8')
        f = io.StringIO(content)
        reader = csv.reader(f)
        
        header = None
        indices = {}
        for row in reader:
            row_upper = [col.strip().upper() for col in row]
            if "BRANCH" in row_upper and "SITE LOGICAL" in row_upper:
                header = row_upper
                indices = {col.strip().upper(): i for i, col in enumerate(row)}
                break
                
        if not header:
            return False, "Faltan columnas Branch o Site Logical en la pestaña SITES"
            
        branch_idx = indices.get("BRANCH")
        site_idx = indices.get("SITE LOGICAL")
            
        sites_list = []
        for row in reader:
            if not row or len(row) <= max(branch_idx, site_idx):
                continue
            branch = row[branch_idx].strip().upper()
            site = row[site_idx].strip().upper()
            if branch and site and site != "SITE LOGICAL":
                sites_list.append({
                    "branch": branch,
                    "site": site
                })
        
        import os
        import json
        static_dir = os.path.join(app.root_path, "static")
        os.makedirs(static_dir, exist_ok=True)
        with open(os.path.join(static_dir, "sites.json"), "w", encoding="utf-8") as out:
            json.dump(sites_list, out, ensure_ascii=False, indent=2)
            
        return True, f"Sincronizados {len(sites_list)} sites."
    except Exception as e:
        print("Error en sincronización de sites:", e)
        return False, f"Error en sites: {str(e)}"


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



def asegurar_esquema():
    # Asegurar que las columnas existan dinámicamente si las tablas ya existen en el DB
    columnas = [
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS dni VARCHAR(20) UNIQUE",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS rol VARCHAR(20) DEFAULT 'operador'",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS branch VARCHAR(30) DEFAULT 'ALL'",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS activo BOOLEAN DEFAULT TRUE",
        "ALTER TABLE operadores ADD COLUMN IF NOT EXISTS fecha_creacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP",
        
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
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS materiales_json TEXT",
        
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
        
        # 3. Crear/restablecer usuario NOC
        noc = Operador.query.filter(db.func.lower(Operador.dni) == "noc").first()
        if noc:
            noc.dni = "NOC"
            noc.password_hash = generate_password_hash("Bitel@123")
            noc.rol = "noc"
            noc.branch = "ALL"
            noc.correo = "noc@botvip.com"
        else:
            default_pwd = generate_password_hash("Bitel@123")
            noc = Operador(
                nombre="Usuario NOC",
                dni="NOC",
                correo="noc@botvip.com",
                password_hash=default_pwd,
                rol="noc",
                branch="ALL",
                activo=True
            )
            db.session.add(noc)
        
        db.session.commit()
        print("Sedes, Administrador y NOC sembrados correctamente.")
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
        total_cerrados = len(cerrados_ext)
        
        # Sincronizar sites
        success_sites, msg_sites = sincronizar_sites()
        if not success_sites:
            print("Advertencia en sincronización de sites:", msg_sites)
            
        return True, f"Sincronización exitosa: {nuevos} creados, {actualizados} actualizados y {total_cerrados} cerrados externamente."
    except Exception as e:
        db.session.rollback()
        print("Error en sincronización de drive:", e)
        return False, f"Error en sincronización: {str(e)}"


@app.route("/manifest.json")
def serve_manifest():
    return app.send_static_file("manifest.json")


@app.route("/sw.js")
def serve_sw():
    response = app.send_static_file("sw.js")
    response.headers["Service-Worker-Allowed"] = "/"
    return response


@app.route("/api/averias/alertas", methods=["GET"])
@login_requerido
def api_alertas_averias():
    last_id = request.args.get("last_id", type=int)
    if last_id is None:
        return {"alertas": []}, 200
        
    branch = session.get("operador_branch")
    es_admin = session.get("operador_rol") == "admin"
    
    query = Averia.query.filter(Averia.estado == "PENDIENTE", Averia.id > last_id)
    if not es_admin and branch != "ALL":
        query = query.filter_by(branch=branch)
        
    nuevas = query.order_by(Averia.id.asc()).all()
    
    alertas = []
    for av in nuevas:
        alertas.append({
            "id": av.id,
            "cuenta": av.cuenta,
            "branch": av.branch,
            "caja": av.caja or "N/A",
            "detalles": av.detalles or "Sin descripción",
            "coordenadas": av.coordenadas or ""
        })
        
    return {"alertas": alertas}, 200


@app.route("/diagnostico-operadores")
def diagnostico_operadores():
    try:
        crear_operador_defecto()
        ops = Operador.query.all()
        resultado = f"<h3>Diagnóstico de Operadores y Sedes</h3>"
        resultado += "<table border='1' cellpadding='5' style='border-collapse:collapse;'>"
        resultado += "<tr><th>ID</th><th>Nombre</th><th>DNI</th><th>Sede</th><th>Rol</th><th>Activo</th><th>Verificación Password (Bitel@123 / Vip@123)</th></tr>"
        
        for op in ops:
            matches_pwd = check_password_hash(op.password_hash, "Bitel@123") or check_password_hash(op.password_hash, "Vip@123")
            resultado += f"<tr>"
            resultado += f"<td>{op.id}</td>"
            resultado += f"<td>{op.nombre}</td>"
            resultado += f"<td>{op.dni}</td>"
            resultado += f"<td>{op.branch}</td>"
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
    if session.get("operador_rol") == "noc":
        return redirect(url_for("noc_dashboard"))
        
    es_admin = session.get("operador_rol") == "admin"
    branch = session.get("operador_branch")
    
    stats = obtener_estadisticas(branch=branch, es_admin=es_admin)
    
    # Obtener las averías pendientes y reparadas (últimas 100) correspondientes a la sede
    query_pendientes = Averia.query.filter_by(estado="PENDIENTE")
    query_reparadas = Averia.query.filter_by(estado="REPARADO")
    
    if not es_admin and branch != "ALL":
        query_pendientes = query_pendientes.filter_by(branch=branch)
        query_reparadas = query_reparadas.filter_by(branch=branch)
        
    averias_pendientes = query_pendientes.order_by(Averia.dias_pendientes.desc().nullslast()).all()
    averias_reparadas = query_reparadas.order_by(Averia.fecha_resolucion.desc()).limit(100).all()
    averias_totales = averias_pendientes + averias_reparadas
    
    # Serializar datos para Leaflet map
    map_data = []
    for av in averias_totales:
        if av.coordenadas:
            coords = av.coordenadas.split(",")
            if len(coords) == 2:
                try:
                    lat = float(coords[0].strip())
                    lng = float(coords[1].strip())
                    map_data.append({
                        "id": av.id,
                        "branch": av.branch,
                        "cuenta": av.cuenta or "Sin cuenta",
                        "codigo_wo": av.codigo_wo or "N/A",
                        "detalles": av.detalles or "Sin descripción",
                        "contrata": av.contrata or "Sin contrata",
                        "site": av.site or "N/A",
                        "caja": av.caja or "N/A",
                        "dias": av.dias_pendientes or 0,
                        "estado": av.estado,
                        "lat": lat,
                        "lng": lng
                    })
                except ValueError:
                    continue
                    
    # Habilitar formulario manual para todos
    es_provincia = True
    
    from sqlalchemy import func
    # Calculate material totals for this branch/sede
    query_sums = db.session.query(
        func.sum(Averia.material_cable_m).label("cable"),
        func.sum(Averia.material_conectores).label("conectores"),
        func.sum(Averia.material_rosetas).label("rosetas"),
        func.sum(Averia.material_mangas).label("mangas"),
        func.sum(Averia.material_acopladores).label("acopladores")
    ).filter_by(estado="REPARADO")
    
    if not es_admin and branch != "ALL":
        query_sums = query_sums.filter_by(branch=branch)
        
    sums = query_sums.first()
    
    materials_totals = {
        "cable": (sums.cable if sums and sums.cable is not None else 0),
        "conectores": (sums.conectores if sums and sums.conectores is not None else 0),
        "rosetas": (sums.rosetas if sums and sums.rosetas is not None else 0),
        "mangas": (sums.mangas if sums and sums.mangas is not None else 0),
        "acopladores": (sums.acopladores if sums and sums.acopladores is not None else 0)
    }
    
    import json
    return render_template(
        "dashboard.html", 
        stats=stats, 
        map_json=json.dumps(map_data),
        averias_list=averias_pendientes,
        es_provincia=es_provincia,
        materials=materials_totals
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
        import json
        materiales_json_str = request.form.get("materiales_json", "{}")
        comentarios = request.form.get("comentarios", "").strip()
        
        averia.estado = "REPARADO"
        averia.fecha_resolucion = datetime.now()
        averia.tecnico_id = session.get("operador_id")
        averia.materiales_json = materiales_json_str
        averia.material_comentarios = comentarios or "Reparado desde el portal"
        
        # Calcular valores compatibles para las 5 columnas básicas
        try:
            mats = json.loads(materiales_json_str)
        except Exception:
            mats = {}
            
        cable_m = 0
        conectores = 0
        rosetas = 0
        mangas = 0
        acopladores = 0
        
        for key, cant in mats.items():
            parts = key.split("|")
            codigo = parts[0]
            nombre = parts[1]
            # Mapear cable drop
            if "Drop" in nombre or codigo == "299381":
                cable_m += cant
            # Mapear conectores
            elif "Conector" in nombre or codigo in ["299378", "299379"]:
                conectores += cant
            # Mapear rosetas
            elif "Roseta" in nombre:
                rosetas += cant
            # Mapear mangas/cajas de empalme
            elif "empalme" in nombre.lower() or "Caja de empalme" in nombre or codigo in ["299798", "299799", "424", "423", "8810", "262079"]:
                mangas += cant
            # Mapear acopladores
            elif "Acoplador" in nombre or "Patch Cord" in nombre or "transceptor" in nombre.lower() or "módulo" in nombre.lower():
                acopladores += cant
                
        averia.material_cable_m = cable_m
        averia.material_conectores = conectores
        averia.material_rosetas = rosetas
        averia.material_mangas = mangas
        averia.material_acopladores = acopladores
        
        db.session.commit()
        flash(f"Avería de cuenta {averia.cuenta} resuelta y materiales registrados.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error al resolver avería: {str(e)}", "danger")
        
    return redirect(request.referrer or url_for("dashboard"))


@app.route("/averias/crear", methods=["POST"])
@login_requerido
def crear_averia_manual():
    branch = session.get("operador_branch")
    es_admin = session.get("operador_rol") == "admin"
    
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
        
        # Validar XBOX o HUBOX
        if not (xbox.startswith("XB") or xbox.startswith("HB")):
            flash("El tipo de caja debe ser XBOX (XB01, XB02) o HUBOX (HB01, HB02...).", "danger")
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
                existente.contrata = contrata or "TGI"
                existente.dias_pendientes = 0.0
                existente.origen = "MANUAL"
                db.session.commit()
                flash(f"Se reactivó la avería para la cuenta {cuenta}.", "success")
                return redirect(url_for("dashboard"))
        
        nueva = Averia(
            branch=target_branch,
            cuenta=cuenta,
            site=site,
            caja=caja_compuesta,
            coordenadas=coordenadas,
            detalles=detalles,
            contrata=contrata or "TGI",
            estado="PENDIENTE",
            dias_pendientes=0.0,
            origen="MANUAL"
        )
        db.session.add(nueva)
        db.session.commit()
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
        "Técnico que Resolvió"
    ]

    for m in MATERIALES_MASTER:
        headers.append(f"[{m['codigo']}] {m['nombre']}")

    headers.append("Comentarios Solución")
    sheet.append(headers)

    for av in averias:
        tecnico_nombre = av.tecnico.nombre if av.tecnico else ""
        row_data = [
            av.id,
            av.branch,
            av.codigo_wo or "",
            av.cuenta or "",
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
            tecnico_nombre
        ]

        mats_dict = av.materiales_dict
        for m in MATERIALES_MASTER:
            key = f"{m['codigo']}|{m['nombre']}"
            cant = mats_dict.get(key)
            if cant is None:
                if m['codigo'] != "Sin Código":
                    matching_cants = [c for k, c in mats_dict.items() if k.startswith(m['codigo'] + "|")]
                    if len(matching_cants) == 1:
                        cant = matching_cants[0]
            row_data.append(cant if cant is not None else "")

        row_data.append(av.material_comentarios or "")
        sheet.append(row_data)

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


@app.route("/noc")
@login_requerido
def noc_dashboard():
    if session.get("operador_rol") not in ["noc", "admin"]:
        flash("Acceso denegado: Se requieren permisos de NOC o Administrador.", "danger")
        return redirect(url_for("dashboard"))
        
    reparadas = Averia.query.filter_by(estado="REPARADO").order_by(Averia.fecha_resolucion.desc()).all()
    
    total_cable = 0
    total_conectores = 0
    total_rosetas = 0
    total_mangas = 0
    total_acopladores = 0
    
    MESES_ES = {
        1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
        5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
        9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
    }
    
    por_mes = {}
    por_branch = {}
    
    for av in reparadas:
        fecha = av.fecha_resolucion
        if fecha:
            mes_key = fecha.strftime("%Y-%m")
            mes_label = f"{MESES_ES.get(fecha.month, 'Desconocido')} {fecha.year}"
        else:
            mes_key = "Sin Fecha"
            mes_label = "Sin Fecha"
            
        br = av.branch or "Sin Sede"
        
        cable = av.material_cable_m or 0
        conectores = av.material_conectores or 0
        rosetas = av.material_rosetas or 0
        mangas = av.material_mangas or 0
        acopladores = av.material_acopladores or 0
        
        total_cable += cable
        total_conectores += conectores
        total_rosetas += rosetas
        total_mangas += mangas
        total_acopladores += acopladores
        
        # Agrupación por Mes
        if mes_key not in por_mes:
            por_mes[mes_key] = {
                "key": mes_key,
                "label": mes_label,
                "cable": 0,
                "conectores": 0,
                "rosetas": 0,
                "mangas": 0,
                "acopladores": 0,
                "total_casos": 0
            }
        por_mes[mes_key]["cable"] += cable
        por_mes[mes_key]["conectores"] += conectores
        por_mes[mes_key]["rosetas"] += rosetas
        por_mes[mes_key]["mangas"] += mangas
        por_mes[mes_key]["acopladores"] += acopladores
        por_mes[mes_key]["total_casos"] += 1
        
        # Agrupación por Branch (Sede)
        if br not in por_branch:
            por_branch[br] = {
                "branch": br,
                "cable": 0,
                "conectores": 0,
                "rosetas": 0,
                "mangas": 0,
                "acopladores": 0,
                "total_casos": 0
            }
        por_branch[br]["cable"] += cable
        por_branch[br]["conectores"] += conectores
        por_branch[br]["rosetas"] += rosetas
        por_branch[br]["mangas"] += mangas
        por_branch[br]["acopladores"] += acopladores
        por_branch[br]["total_casos"] += 1

    # Ordenar agrupaciones
    por_mes_lista = sorted(por_mes.values(), key=lambda x: x["key"], reverse=True)
    por_branch_lista = sorted(por_branch.values(), key=lambda x: x["branch"])
    
    stats_noc = {
        "total_reparadas": len(reparadas),
        "total_cable": total_cable,
        "total_conectores": total_conectores,
        "total_rosetas": total_rosetas,
        "total_mangas": total_mangas,
        "total_acopladores": total_acopladores
    }
    
    return render_template(
        "noc_dashboard.html",
        reparadas=reparadas,
        por_mes=por_mes_lista,
        por_branch=por_branch_lista,
        stats=stats_noc
    )


with app.app_context():
    try:
        db.create_all()
        asegurar_esquema()
        crear_operador_defecto()
        # Verify if static/sites.json exists, otherwise fetch it
        import os
        sites_path = os.path.join(app.root_path, "static", "sites.json")
        if not os.path.exists(sites_path):
            print("SITES cache not found, fetching...")
            sincronizar_sites()
            
        print("Base de datos y esquemas inicializados correctamente.")
    except Exception as e:
        print("Error inicializando base de datos en arranque:", e)


if __name__ == "__main__":
    app.run(port=5000)
