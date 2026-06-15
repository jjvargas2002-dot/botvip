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
from models import db, Cliente, Caso, Operador, Averia, StockBranch

app = Flask(__name__)
app.config.from_object(Config)
db.init_app(app)

MATERIALES_MASTER = [
    # Sección OLT
    {"codigo": "291368", "nombre": "OLT GPON ZTE C610 - 16 puertos DC", "seccion": "OLT"},
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
    {"codigo": "299379", "nombre": "Conector Waterproof", "seccion": "CONECTORES"},
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

def obtener_material_mapeado(codigo, nombre):
    nombre_lower = nombre.lower()
    
    # 1. Buscar en catálogo completo para obtener la sección si es necesario
    master_item = next((m for m in MATERIALES_MASTER if m["codigo"] == codigo or m["nombre"].lower() == nombre_lower), None)
    seccion = master_item["seccion"] if master_item else ""
    
    # 2. Cable Drop (código 283866)
    if codigo == "283866" or "drop" in nombre_lower:
        return "283866", "Cable Drop", "CABLES"
        
    # 3. FAC (código 299378)
    elif codigo == "299378" or "fac" in nombre_lower or "conector rápido" in nombre_lower or "conector rapido" in nombre_lower:
        return "299378", "FAC", "CONECTORES"
        
    # 4. Waterproof (código 299379)
    elif codigo == "299379" or "waterproof" in nombre_lower:
        return "299379", "Waterproof", "CONECTORES"
        
    # 5. Mufas (código 299799): agrupar todos los códigos de la sección "CAJAS"
    elif seccion == "CAJAS" or "empalme" in nombre_lower or "caja de empalme" in nombre_lower or codigo in ["299798", "299799", "424", "423", "8810", "262079"] or "odb" in nombre_lower:
        return "299799", "Mufas", "CAJAS"
        
    # 6. Preconectorizado (código 294790): agrupar todas las fibras que digan "Preconectorizado" de la "Sección CABLES"
    elif (seccion == "CABLES" and "preconectoriza" in nombre_lower) or "preconectoriza" in nombre_lower:
        return "294790", "Preconectorizado", "CABLES"
        
    # Fallback to original
    return codigo, nombre, seccion

@app.context_processor
def utility_processor():
    materiales_por_seccion = {}
    
    # Materiales comunes
    materiales_comunes_nombres = [
        "Fibra óptica preconectorizada de 12 núcleos 200m MPO/APC",
        "Caja de empalme plana 12Fo, 4 puertos",
        "Tubo corrugado de PVC ignífugo 25mm",
        "Conector Waterproof",
        "Grapa de tensión para vano de 200m",
        "Grapa de tensión para vano de 100m"
    ]
    
    comunes = []
    for nombre_comun in materiales_comunes_nombres:
        found = next((m for m in MATERIALES_MASTER if m["nombre"] == nombre_comun), None)
        if found:
            comunes.append(found)
            
    if comunes:
        materiales_por_seccion["MATERIALES COMUNES"] = comunes
        
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
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS tipificacion VARCHAR(100)",
        "ALTER TABLE averias ADD COLUMN IF NOT EXISTS cuentas_asociadas TEXT",
        """
        CREATE TABLE IF NOT EXISTS stock_branch (
            id SERIAL PRIMARY KEY,
            branch VARCHAR(50) NOT NULL,
            material_codigo VARCHAR(50) NOT NULL,
            material_nombre VARCHAR(255) NOT NULL,
            stock_actual INTEGER DEFAULT 0,
            stock_enviado_noc INTEGER DEFAULT 0,
            fecha_envio_noc DATE,
            fecha_actualizacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(branch, material_nombre)
        )
        """,
        "ALTER TABLE stock_branch DROP CONSTRAINT IF EXISTS uq_branch_material",
        "ALTER TABLE stock_branch DROP CONSTRAINT IF EXISTS stock_branch_branch_material_codigo_key",
        "ALTER TABLE stock_branch ADD CONSTRAINT uq_branch_material_nombre UNIQUE (branch, material_nombre)",
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
            
            # Verificar si en el Drive figura como resuelto
            status_upper = status_ont.upper().strip()
            status_caja_upper = status_caja.upper().strip()
            resuelto_keywords = ["REPARADO", "SOLUCIONADO", "CERRADO", "OK", "ATENDIDO", "RESUELTO"]
            es_resuelto_drive = any(k in status_upper or k in status_caja_upper for k in resuelto_keywords)

            # Buscar en BD
            averia = Averia.query.filter_by(cuenta=cuenta).first()
            if averia:
                # Si ya existe, actualizar datos del drive solo si está pendiente localmente
                if averia.estado == "PENDIENTE":
                    if es_resuelto_drive:
                        averia.estado = "REPARADO"
                        averia.fecha_resolucion = datetime.now()
                        averia.material_comentarios = f"Marcado como resuelto en el Drive (Estado: {status_ont or status_caja})"
                    else:
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
                    estado="REPARADO" if es_resuelto_drive else "PENDIENTE",
                    fecha_resolucion=datetime.now() if es_resuelto_drive else None,
                    material_comentarios=f"Marcado como resuelto en el Drive al importar" if es_resuelto_drive else None,
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
    
    # Obtener las averías pendientes y las reparadas en los últimos 12 meses
    from datetime import datetime, timedelta
    hace_12_meses = datetime.now() - timedelta(days=365)
    query_pendientes = Averia.query.filter_by(estado="PENDIENTE")
    query_reparadas = Averia.query.filter(Averia.estado == "REPARADO", Averia.fecha_resolucion >= hace_12_meses)
    
    if not es_admin and branch != "ALL":
        query_pendientes = query_pendientes.filter_by(branch=branch)
        query_reparadas = query_reparadas.filter_by(branch=branch)
        
    averias_pendientes = query_pendientes.order_by(Averia.dias_pendientes.desc().nullslast()).all()
    averias_reparadas = query_reparadas.order_by(Averia.fecha_resolucion.desc()).all()
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
                        "lng": lng,
                        "materiales_json": av.materiales_json or "{}",
                        "comentarios": av.material_comentarios or "",
                        "tipificacion": av.tipificacion or "",
                        "fecha_creacion": av.fecha_creacion.strftime("%Y-%m-%d") if av.fecha_creacion else "",
                        "fecha_resolucion": av.fecha_resolucion.strftime("%Y-%m-%d") if av.fecha_resolucion else "",
                        "cable": av.material_cable_m or 0,
                        "conectores": av.material_conectores or 0,
                        "rosetas": av.material_rosetas or 0,
                        "mangas": av.material_mangas or 0,
                        "acopladores": av.material_acopladores or 0
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
        averias_list=averias_totales,
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


def ajustar_stock_por_consumo(branch, old_mats_dict, new_mats_dict):
    """
    Compara old_mats_dict y new_mats_dict, calcula la diferencia (new_qty - old_qty)
    y descuenta esa diferencia de StockBranch.stock_actual para la sede dada.
    """
    all_keys = set(list(old_mats_dict.keys()) + list(new_mats_dict.keys()))
    for key in all_keys:
        parts = key.split("|")
        if len(parts) < 2:
            continue
        codigo = parts[0]
        nombre = parts[1]
        
        old_qty = old_mats_dict.get(key, 0)
        new_qty = new_mats_dict.get(key, 0)
        
        try:
            old_qty = int(old_qty)
        except (ValueError, TypeError):
            old_qty = 0
            
        try:
            new_qty = int(new_qty)
        except (ValueError, TypeError):
            new_qty = 0
            
        diff = new_qty - old_qty
        if diff == 0:
            continue
            
        reg = StockBranch.query.filter_by(branch=branch, material_nombre=nombre).first()
        if not reg:
            reg = StockBranch(
                branch=branch,
                material_codigo=codigo,
                material_nombre=nombre,
                stock_actual=0
            )
            db.session.add(reg)
            
        reg.stock_actual = max(0, reg.stock_actual - diff)


@app.route("/averias/resolver/<int:id>", methods=["POST"])
@login_requerido
def resolver_averia(id):
    if session.get("operador_rol") == "noc":
        flash("No tienes permisos para realizar esta acción.", "danger")
        return redirect(url_for("dashboard"))
        
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
        
        # Obtener materiales viejos antes de actualizar
        old_mats = averia.materiales_dict
        
        try:
            mats = json.loads(materiales_json_str)
        except Exception:
            mats = {}
            
        # Ajustar el stock según el consumo
        ajustar_stock_por_consumo(averia.branch, old_mats, mats)
        
        averia.estado = "REPARADO"
        averia.fecha_resolucion = datetime.now()
        averia.tecnico_id = session.get("operador_id")
        averia.materiales_json = materiales_json_str
        averia.material_comentarios = comentarios or "Reparado desde el portal"
        averia.tipificacion = request.form.get("tipificacion", "").strip()
        
        # Calcular valores compatibles para las 5 columnas básicas
        cable_m = 0
        conectores = 0
        rosetas = 0
        mangas = 0
        acopladores = 0
        
        for key, cant in mats.items():
            parts = key.split("|")
            codigo = parts[0]
            nombre = parts[1] if len(parts) > 1 else ""
            
            m_cod, m_nom, m_sec = obtener_material_mapeado(codigo, nombre)
            
            if m_nom == "Cable Drop":
                cable_m += cant
            elif m_nom == "FAC":
                conectores += cant
            elif m_nom == "Waterproof":
                rosetas += cant
            elif m_nom == "Mufas":
                mangas += cant
            elif m_nom == "Preconectorizado":
                acopladores += cant
                
        averia.material_cable_m = cable_m
        averia.material_conectores = conectores
        averia.material_rosetas = rosetas
        averia.material_mangas = mangas
        averia.material_acopladores = acopladores
        
        db.session.commit()
        flash(f"Avería de cuenta {averia.cuenta} resuelta y materiales registrados.", "success")
        return redirect(url_for("agrupar_clientes_averia", id=averia.id))
    except Exception as e:
        db.session.rollback()
        flash(f"Error al resolver avería: {str(e)}", "danger")
        return redirect(url_for("dashboard"))



@app.route("/averias/agrupar/<int:id>", methods=["GET", "POST"])
@login_requerido
def agrupar_clientes_averia(id):
    if session.get("operador_rol") == "noc":
        flash("No tienes permisos para realizar esta acción.", "danger")
        return redirect(url_for("dashboard"))
        
    averia = db.session.get(Averia, id)
    if not averia:
        flash("La avería no existe.", "danger")
        return redirect(url_for("dashboard"))
        
    # parse caja values
    def parse_caja(caja_str, site_val):
        if not caja_str:
            return site_val or "", "", "", ""
        parts = [p.strip().upper() for p in caja_str.split("-") if p.strip()]
        site = site_val or ""
        xbox = ""
        hubox = ""
        subox = ""
        
        if len(parts) > 0:
            site = parts[0]
            
        for p in parts[1:]:
            if p.startswith("XB"):
                xbox = p
            elif p.startswith("HB"):
                hubox = p
            else:
                subox = p
                
        return site, xbox, hubox, subox
        
    site, xbox, hubox, subox = parse_caja(averia.caja, averia.site)
    
    # 1. Si la avería (principal) está REPARADA y pasaron más de 7 días desde su resolución, bloquear la agrupación
    if averia.estado == "REPARADO" and averia.fecha_resolucion:
        import datetime
        now = datetime.datetime.now()
        ref_date = averia.fecha_resolucion.replace(tzinfo=None) if averia.fecha_resolucion.tzinfo else averia.fecha_resolucion
        if (now - ref_date).days > 7:
            flash(f"No hay ninguna avería reparada (principal) en el SITE '{site}' para agrupar. Por favor resuelve el ticket principal primero.", "warning")
            return redirect(url_for("dashboard"))
            
    # 2. Si la avería local está PENDIENTE, buscar el principal reparado de este site para redirigir
    if averia.estado == "PENDIENTE":
        import datetime
        from sqlalchemy import or_, func
        site_clean = site.strip().upper() if site else ""
        principales = Averia.query.filter(
            func.upper(func.trim(Averia.site)) == site_clean,
            Averia.estado == "REPARADO",
            or_(
                Averia.material_comentarios.is_(None),
                Averia.material_comentarios == "",
                ~Averia.material_comentarios.like("%Agrupado en la avería principal%")
            )
        ).order_by(Averia.id.desc()).all()
        
        principal_valido = None
        now = datetime.datetime.now()
        for p in principales:
            if p.fecha_resolucion:
                ref_date = p.fecha_resolucion.replace(tzinfo=None) if p.fecha_resolucion.tzinfo else p.fecha_resolucion
                if (now - ref_date).days <= 7:
                    principal_valido = p
                    break
                    
        if principal_valido:
            flash(f"Redirigido a la avería principal del SITE {site} (ID {principal_valido.id}) para realizar la agrupación.", "info")
            return redirect(url_for("agrupar_clientes_averia", id=principal_valido.id))
        else:
            flash(f"No hay ninguna avería reparada (principal) en el SITE '{site}' para agrupar. Por favor resuelve el ticket principal primero.", "warning")
            return redirect(url_for("dashboard"))

    if request.method == "POST":
        new_xbox = request.form.get("xbox", "").strip().upper()
        new_hubox = request.form.get("hubox", "").strip().upper()
        new_subox = request.form.get("subox", "").strip().upper()
        
        caja_parts = [averia.site or site]
        if new_xbox:
            caja_parts.append(new_xbox)
        if new_hubox:
            caja_parts.append(new_hubox)
        if new_subox:
            caja_parts.append(new_subox)
            
        caja_compuesta = "-".join(caja_parts)
        averia.caja = caja_compuesta
        
        selected_accounts = request.form.getlist("selected_clientes")
        new_set = set(selected_accounts)
        
        # Parse old associated list
        old_list = [c.strip() for c in (averia.cuentas_asociadas or "").split(",") if c.strip()]
        old_set = set(old_list)
        
        # Accounts to add to group (new - old)
        to_add = new_set - old_set
        # Accounts to remove from group (old - new)
        to_remove = old_set - new_set
        
        from sqlalchemy import func
        site_clean = site.strip().upper() if site else ""
        
        # For new ones, mark as REPARADO
        if to_add:
            averias_to_add = Averia.query.filter(
                func.upper(func.trim(Averia.site)) == site_clean,
                Averia.cuenta.in_(to_add),
                Averia.id != averia.id
            ).all()
            for av_g in averias_to_add:
                av_g.estado = "REPARADO"
                av_g.fecha_resolucion = datetime.now()
                av_g.tecnico_id = session.get("operador_id")
                av_g.material_comentarios = f"Agrupado en la avería principal (ID {averia.id})"
                av_g.tipificacion = averia.tipificacion
                av_g.material_cable_m = 0
                av_g.material_conectores = 0
                av_g.material_rosetas = 0
                av_g.material_mangas = 0
                av_g.material_acopladores = 0
                av_g.materiales_json = "{}"
                
        # For removed ones, revert back to PENDING
        if to_remove:
            averias_to_remove = Averia.query.filter(
                func.upper(func.trim(Averia.site)) == site_clean,
                Averia.cuenta.in_(to_remove),
                Averia.id != averia.id
            ).all()
            for av_g in averias_to_remove:
                av_g.estado = "PENDIENTE"
                av_g.fecha_resolucion = None
                av_g.tecnico_id = None
                av_g.material_comentarios = None
                av_g.tipificacion = None
                
        averia.cuentas_asociadas = ", ".join(selected_accounts) if selected_accounts else None
        
        try:
            db.session.commit()
            flash("Asociación de clientes y caja guardada con éxito.", "success")
        except Exception as e:
            db.session.rollback()
            flash(f"Error al guardar agrupación: {str(e)}", "danger")
            
        return redirect(url_for("dashboard"))
        
    # GET: query accounts on the same site
    from sqlalchemy import func
    site_clean = site.strip().upper() if site else ""
    associated_list = [c.strip() for c in (averia.cuentas_asociadas or "").split(",") if c.strip()]
    
    # Query all other averias on the same site
    cuentas_query = Averia.query.filter(
        func.upper(func.trim(Averia.site)) == site_clean,
        Averia.id != averia.id
    ).all()
    
    clientes_dict = {c.codigo_cliente: c for c in Cliente.query.all()}
    
    clientes_del_site = []
    vistas = set()
    for row in cuentas_query:
        if row.cuenta and row.cuenta not in vistas:
            is_associated = row.cuenta in associated_list
            
            # Check if this repaired account is already grouped to another main ticket
            is_already_grouped = False
            if row.estado == "REPARADO" and not is_associated:
                if row.cuentas_asociadas:
                    # It is a main ticket of another group
                    is_already_grouped = True
                elif row.material_comentarios and "Agrupado en la avería principal" in row.material_comentarios:
                    if f"ID {averia.id}" not in row.material_comentarios:
                        # It belongs to another group
                        is_already_grouped = True
            
            # Show if: PENDING, or already associated with this group,
            # or REPARADO but not associated with another group yet
            if row.estado == "PENDIENTE" or is_associated or (row.estado == "REPARADO" and not is_already_grouped):
                vistas.add(row.cuenta)
                cl = clientes_dict.get(row.cuenta)
                clientes_del_site.append({
                    "cuenta": row.cuenta,
                    "caja": row.caja or "",
                    "estado": row.estado,
                    "id": row.id,
                    "nombre": cl.nombre if cl else "Cliente de Sheet",
                    "seleccionado": is_associated
                })
            
    return render_template(
        "agrupar.html",
        averia=averia,
        site=site,
        xbox=xbox,
        hubox=hubox,
        clientes_del_site=clientes_del_site
    )


@app.route("/averias/eliminar/<int:id>", methods=["POST"])
@login_requerido
def eliminar_averia(id):
    es_fbb = (session.get("operador_nombre", "").upper() == "FBB" or 
              session.get("operador_dni", "").upper() == "FBB")
    if session.get("operador_rol") != "admin" or not es_fbb:
        flash("No tienes permisos para eliminar registros. Solo la cuenta de FBB (Admin) puede realizar esta acción.", "danger")
        return redirect(url_for("dashboard"))
        
    averia = Averia.query.get_or_404(id)
    if averia.estado == "REPARADO" and averia.materiales_json:
        try:
            import json
            mats = json.loads(averia.materiales_json)
            # Devolvemos el stock: pasar mats como old_mats y un dict vacio como new_mats
            ajustar_stock_por_consumo(averia.branch, mats, {})
        except Exception as e:
            print("Error devolviendo stock al eliminar averia:", e)
            
    try:
        db.session.delete(averia)
        db.session.commit()
        flash(f"Avería ID {id} ({averia.cuenta or 'Sin Cuenta'}) eliminada correctamente y stock devuelto.", "success")
    except Exception as e:
        db.session.rollback()
        flash(f"Error al eliminar la avería: {str(e)}", "danger")
        
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
        xbox = request.form.get("xbox", "").strip().upper()
        hubox = request.form.get("hubox", "").strip().upper()
        caja_input = request.form.get("caja", "").strip().upper()
        coordenadas = request.form["coordenadas"].strip()
        detalles = request.form["detalles"].strip()
        contrata = request.form.get("contrata", "").strip()
        
        # Sede de registro
        target_branch = branch if not es_admin else request.form.get("branch", "ARE").strip().upper()
        
        # Validar XBOX y HUBOX si se proveen
        if xbox and not xbox.startswith("XB"):
            flash("El tipo de caja XBOX debe comenzar con XB.", "danger")
            return redirect(url_for("dashboard"))
        if hubox and not hubox.startswith("HB"):
            flash("El tipo de caja HUBOX debe comenzar con HB.", "danger")
            return redirect(url_for("dashboard"))
            
        # Componer código de caja
        parts = [site]
        if xbox:
            parts.append(xbox)
        if hubox:
            parts.append(hubox)
        if caja_input:
            parts.append(caja_input)
        caja_compuesta = "-".join(parts)
        
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


@app.route("/stock", methods=["GET", "POST"])
@login_requerido
def registrar_stock():
    branch = session.get("operador_branch")
    es_admin = session.get("operador_rol") == "admin"
    es_noc = session.get("operador_rol") == "noc"
    
    if es_admin or es_noc:
        sede_actual = request.args.get("branch", "ARE").strip().upper()
    else:
        sede_actual = branch.strip().upper()
        
    if request.method == "POST":
        if es_noc:
            flash("No tienes permisos para modificar el inventario.", "danger")
            return redirect(url_for("registrar_stock", branch=sede_actual))
        codigos = request.form.getlist("material_codigo")
        nombres = request.form.getlist("material_nombre")
        actuales = request.form.getlist("stock_actual")
        enviados = request.form.getlist("stock_enviado_noc")
        fechas = request.form.getlist("fecha_envio_noc")
        
        for i, nombre in enumerate(nombres):
            cod = codigos[i] if len(codigos) > i else ""
            if not nombre:
                continue
                
            reg = StockBranch.query.filter_by(branch=sede_actual, material_nombre=nombre).first()
            if not reg:
                reg = StockBranch(
                    branch=sede_actual,
                    material_codigo=cod,
                    material_nombre=nombre
                )
                db.session.add(reg)
                
            try:
                submitted_actual = int(actuales[i]) if actuales[i] else 0
            except ValueError:
                submitted_actual = 0
                
            try:
                submitted_enviado = int(enviados[i]) if enviados[i] else 0
            except ValueError:
                submitted_enviado = 0
                
            reg.stock_actual = submitted_actual
            
            # Solo actualizar los detalles del último envío NOC si se ingresó un valor > 0
            if submitted_enviado > 0:
                reg.stock_enviado_noc = submitted_enviado
                fecha_str = fechas[i] if len(fechas) > i else ""
                if fecha_str:
                    try:
                        reg.fecha_envio_noc = datetime.strptime(fecha_str, "%Y-%m-%d").date()
                    except ValueError:
                        reg.fecha_envio_noc = datetime.today().date()
                else:
                    reg.fecha_envio_noc = datetime.today().date()

                
        try:
            db.session.commit()
            flash("Stock actualizado correctamente.", "success")
        except Exception as e:
            db.session.rollback()
            flash(f"Error al actualizar stock: {str(e)}", "danger")
            
        return redirect(url_for("registrar_stock", branch=sede_actual))
        
    stock_regs = StockBranch.query.filter_by(branch=sede_actual).all()
    stock_dict = {r.material_nombre: r for r in stock_regs}
    
    materiales_stock = []
    for m in MATERIALES_MASTER:
        reg = stock_dict.get(m["nombre"])
        materiales_stock.append({
            "codigo": m["codigo"],
            "nombre": m["nombre"],
            "seccion": m["seccion"],
            "stock_actual": reg.stock_actual if reg else 0,
            "stock_enviado_noc": reg.stock_enviado_noc if reg else 0,
            "fecha_envio_noc": reg.fecha_envio_noc.strftime("%Y-%m-%d") if reg and reg.fecha_envio_noc else ""
        })
        
    sedes = ["LI1", "LI2", "LI3", "LI4", "LI7", "ARE", "PIU", "SAN", "CAJ", "LAL", "HUN", "CUS", "JUN"]
    
    materiales_por_seccion_stock = {}
    for m in materiales_stock:
        sec = m["seccion"]
        if sec not in materiales_por_seccion_stock:
            materiales_por_seccion_stock[sec] = []
        materiales_por_seccion_stock[sec].append(m)
        
    # Calcular consumos de esta sede
    reparadas_sede = Averia.query.filter_by(branch=sede_actual, estado="REPARADO").all()
    consumo_dict = {f"{m['codigo']}|{m['nombre']}": 0 for m in MATERIALES_MASTER}
    
    for av in reparadas_sede:
        mats_dict = av.materiales_dict
        for key, cant in mats_dict.items():
            if cant:
                if key in consumo_dict:
                    consumo_dict[key] += cant
                else:
                    parts = key.split("|")
                    if len(parts) > 0:
                        code = parts[0]
                        found = next((m for m in MATERIALES_MASTER if m["codigo"] == code), None)
                        if found:
                            master_key = f"{found['codigo']}|{found['nombre']}"
                            consumo_dict[master_key] += cant

    consumo_materiales = []
    for m in MATERIALES_MASTER:
        key = f"{m['codigo']}|{m['nombre']}"
        cant = consumo_dict.get(key, 0)
        if cant > 0:
            consumo_materiales.append({
                "codigo": m["codigo"],
                "nombre": m["nombre"],
                "seccion": m["seccion"],
                "cantidad": cant
            })
            
    # Calcular tipificaciones de esta sede
    typifications = [
        "Personas externa cortó",
        "Camión grande rompió fibra",
        "Puerto sucio",
        "Refusión de hilo / cambio de módulo",
        "Cambio de caja",
        "Caja robada",
        "Reemplazo de poste eléctrico",
        "Conector roto",
        "OLT caída"
    ]
    tipificaciones_sede = []
    for typ in typifications:
        count = sum(1 for av in reparadas_sede if av.tipificacion == typ)
        tipificaciones_sede.append({"tipificacion": typ, "cantidad": count})
        
    return render_template(
        "stock.html",
        materiales_por_seccion=materiales_por_seccion_stock,
        sede_actual=sede_actual,
        sedes=sedes,
        es_admin_or_noc=(es_admin or es_noc),
        es_noc=es_noc,
        consumo_materiales=consumo_materiales,
        tipificaciones_data=tipificaciones_sede,
        total_reparadas_sede=len(reparadas_sede)
    )


@app.route("/stock/exportar/<branch>", methods=["GET"])
@login_requerido
def exportar_inventario_sede(branch):
    user_branch = session.get("operador_branch")
    es_admin = session.get("operador_rol") == "admin"
    es_noc = session.get("operador_rol") == "noc"
    
    branch = branch.strip().upper()
    if not es_admin and not es_noc and user_branch != branch:
        flash("No tienes permisos para descargar el inventario de esta sede.", "danger")
        return redirect(url_for("registrar_stock", branch=user_branch))
        
    stock_regs = StockBranch.query.filter_by(branch=branch).all()
    stock_dict = {r.material_nombre: r for r in stock_regs}
    
    wb = Workbook()
    ws = wb.active
    ws.title = f"Inventario - {branch}"
    
    header_fill = PatternFill(start_color="0EA5E9", end_color="0EA5E9", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=11)
    
    headers = [
        "Sección",
        "Código",
        "Material",
        "Stock Actual",
        "Último Stock Enviado NOC",
        "Último Fecha Envío NOC"
    ]
    
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
    row_idx = 2
    for m in MATERIALES_MASTER:
        reg = stock_dict.get(m["nombre"])
        stock_actual = reg.stock_actual if reg else 0
        enviado_noc = reg.stock_enviado_noc if reg else 0
        fecha_noc = reg.fecha_envio_noc.strftime("%Y-%m-%d") if reg and reg.fecha_envio_noc else ""
        
        row_values = [
            m["seccion"],
            m["codigo"],
            m["nombre"],
            stock_actual,
            enviado_noc,
            fecha_noc
        ]
        
        for col_idx, val in enumerate(row_values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            cell.font = regular_font
            if col_idx in [1, 2, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")
                
        row_idx += 1
        
    for col in ws.columns:
        max_len = 0
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        col_letter = get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)
        
    ws.row_dimensions[1].height = 28
    
    out = BytesIO()
    wb.save(out)
    out.seek(0)
    
    filename = f"inventario_{branch.lower()}_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        out,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )


@app.route("/averias", methods=["GET"])
@login_requerido
def listar_averias():
    es_admin = session.get("operador_rol") == "admin"
    es_noc = session.get("operador_rol") == "noc"
    branch = session.get("operador_branch")
    
    query = Averia.query
    if not es_admin and not es_noc and branch != "ALL":
        query = query.filter_by(branch=branch)
        
    averias = query.order_by(Averia.id.desc()).all()
    stats = obtener_estadisticas(branch=branch, es_admin=(es_admin or es_noc))
    return render_template("averias.html", averias=averias, stats=stats)


@app.route("/averias/exportar", methods=["GET"])
@login_requerido
def exportar_averias():
    import collections
    es_admin = session.get("operador_rol") == "admin"
    es_noc = session.get("operador_rol") == "noc"
    user_branch = session.get("operador_branch")
    
    branch_arg = request.args.get("branch")
    month_arg = request.args.get("month")
    
    if not es_admin and not es_noc and user_branch != "ALL":
        target_branch = user_branch
    else:
        target_branch = branch_arg if branch_arg else "ALL"
        
    query = Averia.query
    if target_branch != "ALL":
        query = query.filter_by(branch=target_branch)
        
    averias = query.order_by(Averia.id.desc()).all()
    
    # Filter by month python-side (consistent with frontend data-month)
    if month_arg:
        filtered_averias = []
        for av in averias:
            av_date = av.fecha_resolucion if av.fecha_resolucion else av.fecha_creacion
            if av_date and av_date.strftime("%Y-%m") == month_arg:
                filtered_averias.append(av)
        averias = filtered_averias
        
    reparadas = [av for av in averias if av.estado == "REPARADO" and av.fecha_resolucion]
    
    # Group reparadas by month (YYYY-MM)
    reparadas_por_mes = collections.defaultdict(list)
    for av in reparadas:
        mes_str = av.fecha_resolucion.strftime("%Y-%m")
        reparadas_por_mes[mes_str].append(av)
        
    # Group reparadas by month and branch
    reparadas_por_mes_y_branch = collections.defaultdict(list)
    for av in reparadas:
        mes_str = av.fecha_resolucion.strftime("%Y-%m")
        br = av.branch or "Sin Sede"
        reparadas_por_mes_y_branch[(mes_str, br)].append(av)
        
    # Build dictionary of stock by branch and material name
    stock_by_branch = collections.defaultdict(dict)
    all_stock = StockBranch.query.all()
    for r in all_stock:
        stock_by_branch[r.branch][r.material_nombre] = {
            "stock_actual": r.stock_actual or 0,
            "stock_enviado_noc": r.stock_enviado_noc or 0,
            "fecha_envio_noc": r.fecha_envio_noc.strftime("%Y-%m-%d") if r.fecha_envio_noc else ""
        }
        
    workbook = Workbook()
    
    # Styling definitions
    purple_fill = PatternFill("solid", fgColor="5B21B6") # Deep purple header
    purple_font = Font(color="FFFFFF", bold=True)
    
    cyan_fill = PatternFill("solid", fgColor="ECFEFF") # Light cyan accent
    cyan_font = Font(color="0891B2", bold=True)
    
    green_fill = PatternFill("solid", fgColor="D1FAE5") # Light green total
    green_font = Font(color="065F46", bold=True)
    
    bold_font = Font(bold=True)
    center_align = Alignment(horizontal="center", vertical="center")
    left_align = Alignment(horizontal="left", vertical="center")
    right_align = Alignment(horizontal="right", vertical="center")
    
    def get_material_quantity(mats_dict, material_obj):
        key = f"{material_obj['codigo']}|{material_obj['nombre']}"
        cant = mats_dict.get(key)
        if cant is not None:
            return cant
        if material_obj['codigo'] and material_obj['codigo'] != "Sin Código":
            matching_cants = [c for k, c in mats_dict.items() if k.startswith(material_obj['codigo'] + "|")]
            if len(matching_cants) > 0:
                try:
                    return sum(int(c) for c in matching_cants if c is not None)
                except Exception:
                    pass
        return 0

    # 1. SUMMARY SHEET
    summary_sheet = workbook.active
    summary_sheet.title = "SUMMARY"
    
    # Title
    summary_sheet.merge_cells("A1:D1")
    summary_sheet["A1"] = "DASHBOARD DE CONSUMO E INCIDENCIAS"
    summary_sheet["A1"].font = Font(size=16, bold=True, color="5B21B6")
    summary_sheet["A1"].alignment = left_align
    summary_sheet.row_dimensions[1].height = 30
    
    summary_sheet["A2"] = f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    summary_sheet["A2"].font = Font(italic=True, color="64748B")
    
    curr_row = 4
    months_keys = sorted(list(reparadas_por_mes.keys()), reverse=True)
    
    # Table 1: Material Consumption
    summary_sheet.cell(row=curr_row, column=1, value="RESUMEN DE CONSUMO DE MATERIALES").font = Font(size=12, bold=True, color="0891B2")
    curr_row += 1
    
    headers_t1 = ["Material", "Código"] + months_keys + ["Total Consumido"]
    for col_idx, h in enumerate(headers_t1, start=1):
        cell = summary_sheet.cell(row=curr_row, column=col_idx, value=h)
        cell.fill = purple_fill
        cell.font = purple_font
        cell.alignment = center_align
    
    summary_sheet.row_dimensions[curr_row].height = 25
    curr_row += 1
    
    for m in MATERIALES_MASTER:
        has_consumption = False
        row_values = []
        for mes in months_keys:
            cant = sum(get_material_quantity(av.materiales_dict, m) for av in reparadas_por_mes[mes])
            row_values.append(cant)
            if cant > 0:
                has_consumption = True
                
        if not has_consumption:
            continue
            
        summary_sheet.cell(row=curr_row, column=1, value=m["nombre"]).alignment = left_align
        summary_sheet.cell(row=curr_row, column=2, value=m["codigo"]).alignment = center_align
        
        for idx, cant in enumerate(row_values):
            c_cell = summary_sheet.cell(row=curr_row, column=3 + idx, value=cant)
            c_cell.alignment = right_align
            
        total_cell = summary_sheet.cell(row=curr_row, column=3 + len(months_keys))
        total_cell.alignment = right_align
        total_cell.font = bold_font
        total_cell.fill = green_fill
        if months_keys:
            start_letter = get_column_letter(3)
            end_letter = get_column_letter(2 + len(months_keys))
            total_cell.value = f"=SUM({start_letter}{curr_row}:{end_letter}{curr_row})"
        else:
            total_cell.value = 0
        
        curr_row += 1
        
    curr_row += 2
    
    # Table 2: Typifications
    summary_sheet.cell(row=curr_row, column=1, value="RESUMEN DE TIPIFICACIONES DE AVERÍAS").font = Font(size=12, bold=True, color="0891B2")
    curr_row += 1
    
    typifications = [
        "Personas externa cortó",
        "Camión grande rompió fibra",
        "Puerto sucio",
        "Refusión de hilo / cambio de módulo",
        "Cambio de caja",
        "Caja robada",
        "Reemplazo de poste eléctrico",
        "Conector roto",
        "OLT caída"
    ]
    
    headers_t2 = ["Tipificación"] + months_keys + ["Total Incidencias"]
    for col_idx, h in enumerate(headers_t2, start=1):
        cell = summary_sheet.cell(row=curr_row, column=col_idx, value=h)
        cell.fill = purple_fill
        cell.font = purple_font
        cell.alignment = center_align
        
    summary_sheet.row_dimensions[curr_row].height = 25
    curr_row += 1
    
    for typ in typifications:
        summary_sheet.cell(row=curr_row, column=1, value=typ).alignment = left_align
        
        row_counts = []
        for mes in months_keys:
            count = sum(1 for av in reparadas_por_mes[mes] if av.tipificacion == typ)
            row_counts.append(count)
            
        for idx, count in enumerate(row_counts):
            c_cell = summary_sheet.cell(row=curr_row, column=2 + idx, value=count)
            c_cell.alignment = right_align
            
        total_cell = summary_sheet.cell(row=curr_row, column=2 + len(months_keys))
        total_cell.alignment = right_align
        total_cell.font = bold_font
        total_cell.fill = green_fill
        if months_keys:
            start_letter = get_column_letter(2)
            end_letter = get_column_letter(1 + len(months_keys))
            total_cell.value = f"=SUM({start_letter}{curr_row}:{end_letter}{curr_row})"
        else:
            total_cell.value = 0
        
        curr_row += 1

    # 2. RAW DATA SHEET ("Reporte Averias ODN")
    raw_sheet = workbook.create_sheet(title="Reporte Averias ODN")
    
    headers_raw = [
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
        "Tipificación",
        "Cuentas Agrupadas"
    ]
    
    for m in MATERIALES_MASTER:
        headers_raw.append(f"[{m['codigo']}] {m['nombre']}")
        
    headers_raw.append("Comentarios Solución")
    raw_sheet.append(headers_raw)
    
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
            tecnico_nombre,
            av.tipificacion or "",
            av.cuentas_asociadas or ""
        ]
        
        mats_dict = av.materiales_dict
        for m in MATERIALES_MASTER:
            cant = get_material_quantity(mats_dict, m)
            row_data.append(cant if cant > 0 else "")
            
        row_data.append(av.material_comentarios or "")
        raw_sheet.append(row_data)
        
    raw_header_fill = PatternFill("solid", fgColor="1D4ED8")
    raw_header_font = Font(color="FFFFFF", bold=True)
    
    for cell in raw_sheet[1]:
        cell.fill = raw_header_fill
        cell.font = raw_header_font
        cell.alignment = center_align
        
    raw_sheet.freeze_panes = "A2"
    raw_sheet.auto_filter.ref = raw_sheet.dimensions
    
    for row in raw_sheet.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    # 3. MONTHLY BALANCE SHEETS
    if target_branch != "ALL":
        keys_to_generate = [(mes_str, target_branch) for mes_str in sorted(list(reparadas_por_mes.keys()), reverse=True)]
    else:
        # Group by month and sort branches alphabetically
        grouped_by_month = collections.defaultdict(list)
        for (mes_str, br) in reparadas_por_mes_y_branch.keys():
            grouped_by_month[mes_str].append(br)
            
        keys_to_generate = []
        for mes_str in sorted(grouped_by_month.keys(), reverse=True):
            for br in sorted(grouped_by_month[mes_str]):
                keys_to_generate.append((mes_str, br))
                
    for mes_str, br in keys_to_generate:
        sheet_title = f"{mes_str} ({br})" if target_branch == "ALL" else mes_str
        month_sheet = workbook.create_sheet(title=sheet_title)
        
        # Title row
        month_sheet.merge_cells("A1:F1")
        month_sheet["A1"] = f"BALANCE DE MATERIALES - Sede: {br} - Mes: {mes_str}"
        month_sheet["A1"].font = Font(size=14, bold=True, color="5B21B6")
        month_sheet["A1"].alignment = left_align
        month_sheet.row_dimensions[1].height = 30
        
        headers_m = [
            "№",
            "MATERIAL",
            "CODIGO",
            "STOCK ACTUAL",
            "STOCK ENVIADO NOC",
            "FECHA ENVIO NOC",
            "BITEL",
            "TOTAL"
        ]
        
        for col_idx, h in enumerate(headers_m, start=1):
            cell = month_sheet.cell(row=3, column=col_idx, value=h)
            cell.fill = purple_fill
            cell.font = purple_font
            cell.alignment = center_align
            
        month_sheet.row_dimensions[3].height = 25
        
        st_dict_br = stock_by_branch.get(br, {})
        reparadas_m_b = reparadas_por_mes_y_branch.get((mes_str, br), []) if target_branch == "ALL" else reparadas_por_mes.get(mes_str, [])
        
        for idx, m in enumerate(MATERIALES_MASTER, start=1):
            row_num = 3 + idx
            
            st_data = st_dict_br.get(m["nombre"], {"stock_actual": 0, "stock_enviado_noc": 0, "fecha_envio_noc": ""})
            
            month_sheet.cell(row=row_num, column=1, value=idx).alignment = center_align
            month_sheet.cell(row=row_num, column=2, value=m["nombre"]).alignment = left_align
            month_sheet.cell(row=row_num, column=3, value=m["codigo"]).alignment = center_align
            
            month_sheet.cell(row=row_num, column=4, value=st_data["stock_actual"]).alignment = right_align
            month_sheet.cell(row=row_num, column=5, value=st_data["stock_enviado_noc"]).alignment = right_align
            month_sheet.cell(row=row_num, column=6, value=st_data["fecha_envio_noc"]).alignment = center_align
            
            # Bitel consumption (total sum of material consumed in this month/branch combo)
            cant = sum(get_material_quantity(av.materiales_dict, m) for av in reparadas_m_b)
            c_cell = month_sheet.cell(row=row_num, column=7, value=cant)
            c_cell.alignment = right_align
                
            total_cell = month_sheet.cell(row=row_num, column=8)
            total_cell.value = f"=(D{row_num}+E{row_num})-G{row_num}"
            total_cell.alignment = right_align
            total_cell.font = bold_font
            total_cell.fill = green_fill

    # Global column auto-fit
    for ws in workbook.worksheets:
        for col in ws.columns:
            max_len = 0
            for cell in col:
                val = str(cell.value or "")
                if val.startswith("="):
                    val = "0.00"
                if len(val) > max_len:
                    max_len = len(val)
            col_letter = get_column_letter(col[0].column)
            ws.column_dimensions[col_letter].width = min(max(max_len + 3, 10), 50)

    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    
    # Construct descriptive filename
    filename_parts = ["reporte_averias"]
    if target_branch != "ALL":
        filename_parts.append(target_branch.lower())
    else:
        filename_parts.append("global")
    if month_arg:
        filename_parts.append(month_arg)
    download_name = "_".join(filename_parts) + ".xlsx"
    
    return send_file(
        output,
        as_attachment=True,
        download_name=download_name,
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
    
    typifications = [
        "Personas externa cortó",
        "Camión grande rompió fibra",
        "Puerto sucio",
        "Refusión de hilo / cambio de módulo",
        "Cambio de caja",
        "Caja robada",
        "Reemplazo de poste eléctrico",
        "Conector roto",
        "OLT caída"
    ]
    tipificaciones_data = []
    for typ in typifications:
        count = sum(1 for av in reparadas if av.tipificacion == typ)
        tipificaciones_data.append({"tipificacion": typ, "cantidad": count})
        
    stock_regs = StockBranch.query.order_by(StockBranch.branch, StockBranch.material_codigo).all()
    
    return render_template(
        "noc_dashboard.html",
        reparadas=reparadas,
        por_mes=por_mes_lista,
        por_branch=por_branch_lista,
        stats=stats_noc,
        tipificaciones_data=tipificaciones_data,
        stock_regs=stock_regs
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
