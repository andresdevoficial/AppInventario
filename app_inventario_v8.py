# -*- coding: utf-8 -*-
"""
===============================================================================
SISTEMA DE CONTROL DE INVENTARIOS Y KARDEX - VERSIÓN 8.0 OPTIMIZADA
===============================================================================
- Motor de base de datos: MySQL / MariaDB (Connection Pooling).
- Módulos CRUD integrados: Productos (Consulta, Edición y Eliminación Segura),
  Responsables y Categorías.
- Manejo seguro de transacciones, escaping HTML e integridad referencial.
- Exportación Dual (PDF / Excel XML) de Stock General.
- Pie de página institucional "DESARROLLADO POR ANDRES".
===============================================================================
"""

import os
import sys
import json
import io
import html
import urllib.parse
import webbrowser
import threading
import time
from contextlib import contextmanager
from http.server import ThreadingHTTPServer, BaseHTTPRequestHandler
from datetime import datetime
import mysql.connector
from mysql.connector import pooling

try:
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib import colors
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False

# Configuración de Conexión MySQL

# Configuración dinámica para Aiven / Render


# Configuración dinámica leyendo únicamente las variables de entorno de Render
MYSQL_CONFIG = {
    "host": os.environ.get("DB_HOST"),
    "user": os.environ.get("DB_USER", "avnadmin"),
    "password": os.environ.get("DB_PASSWORD"),
    "database": os.environ.get("DB_NAME", "defaultdb"),
    "port": int(os.environ.get("DB_PORT", 28693)),
    "ssl_disabled": False  # Requerido por Aiven
}

PORT = 8000
db_pool = None

# =============================================================================
# 1. GESTIÓN DE CONEXIONES Y BASE DE DATOS
# =============================================================================

def init_db_pool():
    global db_pool
    try:
        # Inicializa directamente el pool conectándose a la BD existente en Aiven
        db_pool = pooling.MySQLConnectionPool(
            pool_name="inventario_pool_v8",
            pool_size=10,
            pool_reset_session=True,
            **MYSQL_CONFIG
        )
        print("Pool de conexiones a la base de datos creado exitosamente.")
    except Exception as e:
        print(f"Error creando el pool de conexiones MySQL: {e}")
        # No usamos sys.exit(1) para evitar tumbar la aplicación

@contextmanager
def get_db_context():
    """Context Manager para asegurar el cierre automático de conexiones y cursores."""
    conn = db_pool.get_connection()
    try:
        yield conn
    finally:
        conn.close()

def init_db():
    with get_db_context() as conn:
        cursor = conn.cursor()
        try:
            conn.start_transaction()

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS responsables (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nombre VARCHAR(100) NOT NULL UNIQUE,
                cargo VARCHAR(100) DEFAULT 'Operativo',
                fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP
            ) ENGINE=InnoDB;
            """)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS categorias (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nombre VARCHAR(100) UNIQUE NOT NULL,
                prefijo VARCHAR(10) UNIQUE NOT NULL
            ) ENGINE=InnoDB;
            """)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS ubicaciones (
                id INT AUTO_INCREMENT PRIMARY KEY,
                nombre VARCHAR(100) UNIQUE NOT NULL
            ) ENGINE=InnoDB;
            """)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS productos (
                id INT AUTO_INCREMENT PRIMARY KEY,
                codigo VARCHAR(20) UNIQUE NOT NULL,
                nombre VARCHAR(150) NOT NULL,
                categoria_id INT NOT NULL,
                stock_minimo DECIMAL(10,2) DEFAULT 15.00 CHECK(stock_minimo >= 0),
                fecha_registro DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (categoria_id) REFERENCES categorias(id) ON DELETE RESTRICT
            ) ENGINE=InnoDB;
            """)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS stock (
                producto_id INT PRIMARY KEY,
                cantidad DECIMAL(10,2) NOT NULL DEFAULT 0.00 CHECK(cantidad >= 0),
                fecha_actualizacion DATETIME DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
                FOREIGN KEY (producto_id) REFERENCES productos(id) ON DELETE RESTRICT
            ) ENGINE=InnoDB;
            """)

            cursor.execute("""
            CREATE TABLE IF NOT EXISTS historial_kardex (
                id INT AUTO_INCREMENT PRIMARY KEY,
                producto_id INT NOT NULL,
                tipo_movimiento ENUM('ENTRADA', 'SALIDA', 'AJUSTE') NOT NULL,
                cantidad DECIMAL(10,2) NOT NULL,
                stock_resultante DECIMAL(10,2) NOT NULL,
                presentacion VARCHAR(100),
                unidad_medida VARCHAR(50),
                ubicacion_id INT,
                motivo VARCHAR(150),
                responsable VARCHAR(100) NOT NULL,
                fecha DATETIME DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (producto_id) REFERENCES productos(id) ON DELETE RESTRICT,
                FOREIGN KEY (ubicacion_id) REFERENCES ubicaciones(id) ON DELETE SET NULL
            ) ENGINE=InnoDB;
            """)

            conn.commit()
            cargar_catalogos_iniciales(conn)
        except Exception as e:
            conn.rollback()
            raise e
        finally:
            cursor.close()

def cargar_catalogos_iniciales(conn):
    cursor = conn.cursor(dictionary=True)

    resp_iniciales = [
        ("Andrés Pérez", "Administrador"),
        ("María Rodríguez", "Supervisora"),
        ("Carlos Gómez", "Auxiliar de Almacén")
    ]
    for nom, car in resp_iniciales:
        cursor.execute("INSERT IGNORE INTO responsables (nombre, cargo) VALUES (%s, %s);", (nom, car))

    cats = [
        ("Cafetería", "CAF"),
        ("Lácteos", "LACT"),
        ("Bebidas", "BEB"),
        ("Insumos", "INS"),
        ("Alimentos", "ALIM")
    ]
    for nom, pref in cats:
        cursor.execute("INSERT IGNORE INTO categorias (nombre, prefijo) VALUES (%s, %s);", (nom, pref))

    ubis = [
        "Refrigerador 1 Cocina",
        "Refrigerador 2 Cocina",
        "Refrigerador 1 Barra",
        "Refrigerador 2 Barra"
    ]
    for u in ubis:
        cursor.execute("INSERT IGNORE INTO ubicaciones (nombre) VALUES (%s);", (u,))

    cursor.execute("SELECT COUNT(*) as total FROM productos;")
    if cursor.fetchone()['total'] == 0:
        cursor.execute("SELECT id FROM categorias WHERE prefijo = 'CAF';")
        cat_caf_row = cursor.fetchone()
        cursor.execute("SELECT id FROM categorias WHERE prefijo = 'LACT';")
        cat_lact_row = cursor.fetchone()

        if cat_caf_row and cat_lact_row:
            cat_caf = cat_caf_row['id']
            cat_lact = cat_lact_row['id']

            items_semilla = [
                ("CAF-001", "Café Grano Espresso 1kg", cat_caf, 50.0),
                ("CAF-002", "Café Descafeinado 500g", cat_caf, 12.0),
                ("LACT-001", "Leche Entera 1L", cat_lact, 100.0)
            ]

            for cod, nom, c_id, cant in items_semilla:
                cursor.execute("INSERT INTO productos (codigo, nombre, categoria_id) VALUES (%s, %s, %s);", (cod, nom, c_id))
                p_id = cursor.lastrowid
                cursor.execute("INSERT INTO stock (producto_id, cantidad) VALUES (%s, %s);", (p_id, cant))
                cursor.execute("""
                    INSERT INTO historial_kardex (producto_id, tipo_movimiento, cantidad, stock_resultante, presentacion, unidad_medida, responsable)
                    VALUES (%s, 'ENTRADA', %s, %s, 'Unidad Individual', 'Unidades', 'Andrés Pérez');
                """, (p_id, cant, cant))

    conn.commit()
    cursor.close()

# =============================================================================
# 2. GENERADOR DE REPORTES PDF Y EXCEL
# =============================================================================

def generar_excel_stock_general():
    with get_db_context() as conn:
        cursor = conn.cursor(dictionary=True)
        cursor.execute("""
            SELECT p.codigo, p.nombre, c.nombre as categoria, COALESCE(s.cantidad, 0) as cantidad, p.stock_minimo
            FROM productos p
            JOIN categorias c ON p.categoria_id = c.id
            LEFT JOIN stock s ON p.id = s.producto_id
            ORDER BY c.nombre ASC, p.nombre ASC;
        """)
        filas = cursor.fetchall()
        cursor.close()

    xml_data = """<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet">
 <Styles>
  <Style ss:ID="Header">
   <Font ss:Bold="1" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#0F172A" ss:Pattern="Solid"/>
   <Alignment ss:Horizontal="Center"/>
  </Style>
  <Style ss:ID="Critico">
   <Font ss:Color="#991B1B" ss:Bold="1"/>
  </Style>
  <Style ss:ID="OK">
   <Font ss:Color="#166534" ss:Bold="1"/>
  </Style>
 </Styles>
 <Worksheet ss:Name="Stock General">
  <Table>
   <Column ss:Width="100"/>
   <Column ss:Width="200"/>
   <Column ss:Width="120"/>
   <Column ss:Width="100"/>
   <Column ss:Width="100"/>
   <Column ss:Width="80"/>
   <Row ss:StyleID="Header">
    <Cell><Data ss:Type="String">Código</Data></Cell>
    <Cell><Data ss:Type="String">Producto</Data></Cell>
    <Cell><Data ss:Type="String">Categoría</Data></Cell>
    <Cell><Data ss:Type="String">Stock Disponible</Data></Cell>
    <Cell><Data ss:Type="String">Stock Mínimo</Data></Cell>
    <Cell><Data ss:Type="String">Estado</Data></Cell>
   </Row>
"""
    for f in filas:
        cant = float(f['cantidad'])
        s_min = float(f['stock_minimo'])
        estado = "CRÍTICO" if cant < s_min else "OK"
        style_estado = "Critico" if estado == "CRÍTICO" else "OK"
        
        xml_data += f"""   <Row>
    <Cell><Data ss:Type="String">{html.escape(f['codigo'])}</Data></Cell>
    <Cell><Data ss:Type="String">{html.escape(f['nombre'])}</Data></Cell>
    <Cell><Data ss:Type="String">{html.escape(f['categoria'])}</Data></Cell>
    <Cell><Data ss:Type="Number">{cant:.2f}</Data></Cell>
    <Cell><Data ss:Type="Number">{s_min:.2f}</Data></Cell>
    <Cell ss:StyleID="{style_estado}"><Data ss:Type="String">{estado}</Data></Cell>
   </Row>
"""
    xml_data += """  </Table>
 </Worksheet>
</Workbook>"""

    return xml_data.encode('utf-8')

def generar_pdf_filtrado(tipo_reporte, param=None):
    if not REPORTLAB_AVAILABLE:
        return None

    with get_db_context() as conn:
        cursor = conn.cursor(dictionary=True)
        styles = getSampleStyleSheet()
        buffer = io.BytesIO()
        doc = SimpleDocTemplate(buffer, pagesize=letter, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
        elementos = []

        try:
            if tipo_reporte == "stock_general":
                titulo = "REPORTE GENERAL DE STOCK DISPONIBLE EN INVENTARIO"
                cursor.execute("""
                    SELECT p.codigo, p.nombre, c.nombre as categoria, COALESCE(s.cantidad, 0) as cantidad, p.stock_minimo
                    FROM productos p
                    JOIN categorias c ON p.categoria_id = c.id
                    LEFT JOIN stock s ON p.id = s.producto_id
                    ORDER BY c.nombre ASC, p.nombre ASC;
                """)
                filas = cursor.fetchall()

                elementos.append(Paragraph(f"<b>{titulo}</b>", styles['Title']))
                elementos.append(Spacer(1, 15))

                tabla_datos = [["Código", "Producto", "Categoría", "Stock Disponible", "Estado"]]
                for f in filas:
                    cant = float(f['cantidad'])
                    s_min = float(f['stock_minimo'])
                    estado = "CRÍTICO" if cant < s_min else "OK"
                    tabla_datos.append([
                        f['codigo'],
                        Paragraph(f['nombre'], styles['Normal']),
                        f['categoria'],
                        f"{cant:.2f}",
                        estado
                    ])

                t = Table(tabla_datos, colWidths=[80, 210, 110, 90, 60])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0f172a")),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('ALIGN', (1,0), (1,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
                ]))
                elementos.append(t)

            elif tipo_reporte == "stock_ubicacion":
                ubi_id = param
                cursor.execute("SELECT nombre FROM ubicaciones WHERE id = %s;", (ubi_id,))
                ubi_row = cursor.fetchone()
                nombre_ubi = ubi_row['nombre'] if ubi_row else "Desconocida"

                titulo = f"REPORTE DE STOCK EN UBICACIÓN: {nombre_ubi.upper()}"
                cursor.execute("""
                    SELECT p.codigo, p.nombre, c.nombre as categoria, SUM(k.cantidad) as total_enviado, MAX(k.fecha) as ultima_salida
                    FROM historial_kardex k
                    JOIN productos p ON k.producto_id = p.id
                    JOIN categorias c ON p.categoria_id = c.id
                    WHERE k.tipo_movimiento = 'SALIDA' AND k.ubicacion_id = %s
                    GROUP BY p.id, p.codigo, p.nombre, c.nombre
                    ORDER BY p.nombre ASC;
                """, (ubi_id,))
                filas = cursor.fetchall()

                elementos.append(Paragraph(f"<b>{titulo}</b>", styles['Title']))
                elementos.append(Spacer(1, 15))

                tabla_datos = [["Código", "Producto", "Categoría", "Cant. Despachada", "Último Movimiento"]]
                for f in filas:
                    tabla_datos.append([
                        f['codigo'],
                        Paragraph(f['nombre'], styles['Normal']),
                        f['categoria'],
                        f"{float(f['total_enviado']):.2f}",
                        str(f['ultima_salida'])[:19]
                    ])

                t = Table(tabla_datos, colWidths=[80, 200, 110, 80, 100])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#0369a1")),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('ALIGN', (1,0), (1,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#bae6fd")),
                ]))
                elementos.append(t)

            elif tipo_reporte == "bajo_stock":
                titulo = "REPORTE DE PRODUCTOS EN BAJO STOCK"
                cursor.execute("""
                    SELECT p.codigo, p.nombre, c.nombre as categoria, s.cantidad, p.stock_minimo
                    FROM productos p
                    JOIN categorias c ON p.categoria_id = c.id
                    JOIN stock s ON p.id = s.producto_id
                    WHERE s.cantidad < p.stock_minimo
                    ORDER BY s.cantidad ASC;
                """)
                filas = cursor.fetchall()
                
                elementos.append(Paragraph(f"<b>{titulo}</b>", styles['Title']))
                elementos.append(Spacer(1, 15))
                
                tabla_datos = [["Código", "Producto", "Categoría", "Stock Actual", "Stock Mínimo"]]
                for f in filas:
                    tabla_datos.append([
                        f['codigo'],
                        Paragraph(f['nombre'], styles['Normal']),
                        f['categoria'],
                        f"{float(f['cantidad']):.2f}",
                        f"{float(f['stock_minimo']):.2f}"
                    ])
                
                t = Table(tabla_datos, colWidths=[80, 200, 100, 90, 90])
                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#991b1b")),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('ALIGN', (1,0), (1,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#fca5a5")),
                ]))
                elementos.append(t)

            elif tipo_reporte in ("ENTRADA", "SALIDA", "AJUSTE"):
                titulos = {
                    "ENTRADA": "REPORTE DE ENTRADAS AL INVENTARIO",
                    "SALIDA": "REPORTE DE SALIDAS DE INVENTARIO Y UBICACIONES",
                    "AJUSTE": "REPORTE DE AJUSTES (MERMAS Y CONSUMO)"
                }
                elementos.append(Paragraph(f"<b>{titulos[tipo_reporte]}</b>", styles['Title']))
                elementos.append(Spacer(1, 15))

                cursor.execute("""
                    SELECT k.fecha, p.codigo, p.nombre, k.cantidad, k.stock_resultante, k.presentacion, k.unidad_medida,
                           COALESCE(u.nombre, 'N/A') as ubicacion, COALESCE(k.motivo, 'N/A') as motivo, k.responsable
                    FROM historial_kardex k
                    JOIN productos p ON k.producto_id = p.id
                    LEFT JOIN ubicaciones u ON k.ubicacion_id = u.id
                    WHERE k.tipo_movimiento = %s
                    ORDER BY k.fecha DESC;
                """, (tipo_reporte,))
                filas = cursor.fetchall()

                if tipo_reporte == "SALIDA":
                    tabla_datos = [["Fecha", "Código", "Producto", "Cant.", "Ubicación Origen", "Responsable"]]
                    for f in filas:
                        tabla_datos.append([
                            str(f['fecha'])[:19], f['codigo'], Paragraph(f['nombre'], styles['Normal']),
                            f"{float(f['cantidad']):.2f}", f['ubicacion'], f['responsable']
                        ])
                    t = Table(tabla_datos, colWidths=[95, 65, 150, 50, 120, 80])
                elif tipo_reporte == "AJUSTE":
                    tabla_datos = [["Fecha", "Código", "Producto", "Nuevo Stock", "Motivo Ajuste", "Responsable"]]
                    for f in filas:
                        tabla_datos.append([
                            str(f['fecha'])[:19], f['codigo'], Paragraph(f['nombre'], styles['Normal']),
                            f"{float(f['stock_resultante']):.2f}", f['motivo'], f['responsable']
                        ])
                    t = Table(tabla_datos, colWidths=[95, 65, 150, 70, 100, 80])
                else:
                    tabla_datos = [["Fecha", "Código", "Producto", "Presentación / Unidad", "Cant. Total", "Responsable"]]
                    for f in filas:
                        pres_str = f"{f['presentacion'] or 'S/D'} ({f['unidad_medida'] or 'Unid.'})"
                        tabla_datos.append([
                            str(f['fecha'])[:19], f['codigo'], Paragraph(f['nombre'], styles['Normal']),
                            pres_str, f"{float(f['cantidad']):.2f}", f['responsable']
                        ])
                    t = Table(tabla_datos, colWidths=[90, 60, 140, 120, 60, 70])

                t.setStyle(TableStyle([
                    ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1e293b")),
                    ('TEXTCOLOR', (0,0), (-1,0), colors.white),
                    ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                    ('ALIGN', (2,0), (2,-1), 'LEFT'),
                    ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#cbd5e1")),
                ]))
                elementos.append(t)

            doc.build(elementos)
            buffer.seek(0)
            return buffer
        finally:
            cursor.close()

# =============================================================================
# 3. INTERFAZ WEB SPA CON CRUD COMPLETO
# =============================================================================

NUEVA_INTERFAZ_HTML = """<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Panel Inventario v8.0 MySQL</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
    <style>
        :root {
            --bg-body: #f1f5f9;
            --sidebar-bg: #0f172a;
            --card-bg: #ffffff;
            --accent: #2563eb;
            --accent-hover: #1d4ed8;
            --text-main: #334155;
            --text-muted: #64748b;
            --danger: #dc2626;
            --success: #16a34a;
            --warning: #d97706;
        }

        * { margin:0; padding:0; box-sizing:border-box; font-family:'Inter', sans-serif; }
        body { background: var(--bg-body); color: var(--text-main); display: flex; flex-direction: column; min-height: 100vh; }

        .app-container { display: flex; flex: 1; }

        aside { width: 310px; background: var(--sidebar-bg); color: white; padding: 24px; display: flex; flex-direction: column; gap: 16px; overflow-y: auto; }
        aside h2 { font-size: 1.1rem; font-weight: 700; color: #38bdf8; }
        
        nav button {
            width: 100%; text-align: left; padding: 12px; background: transparent; border: none;
            color: #94a3b8; font-weight: 600; cursor: pointer; border-radius: 8px; margin-bottom: 4px;
            transition: all 0.2s;
        }
        nav button.active, nav button:hover { background: #1e293b; color: #ffffff; }

        .pdf-group { border-top: 1px solid #334155; padding-top: 12px; display: flex; flex-direction: column; gap: 6px; }
        .pdf-group span { font-size: 0.75rem; color: #64748b; text-transform: uppercase; font-weight: 700; margin-top: 4px; }
        .btn-pdf {
            background: #334155; color: #f8fafc; padding: 8px 10px; border: none;
            border-radius: 6px; font-size: 0.78rem; text-decoration: none; font-weight: 500;
            display: flex; align-items: center; justify-content: space-between; cursor: pointer;
        }
        .btn-pdf:hover { background: #475569; }
        .btn-pdf-accent { background: #0284c7; }
        .btn-pdf-accent:hover { background: #0369a1; }

        main { flex: 1; padding: 32px; overflow-y: auto; display: flex; flex-direction: column; gap: 20px; }

        .page-header {
            background: var(--card-bg);
            padding: 18px 24px;
            border-radius: 12px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            display: flex;
            align-items: center;
            justify-content: space-between;
            border-left: 5px solid var(--accent);
        }
        .page-header h1 { font-size: 1.35rem; font-weight: 700; color: var(--sidebar-bg); }
        .page-header span { font-size: 0.85rem; color: var(--text-muted); font-weight: 500; }

        .kpi-grid { display: grid; grid-template-columns: repeat(3, 1fr); gap: 20px; }
        .kpi-card { background: var(--card-bg); padding: 20px; border-radius: 12px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }
        .kpi-card span { font-size: 0.85rem; color: var(--text-muted); font-weight: 600; }
        .kpi-card h3 { font-size: 1.6rem; margin-top: 8px; color: var(--sidebar-bg); }

        .content-card { background: var(--card-bg); border-radius: 12px; padding: 24px; box-shadow: 0 1px 3px rgba(0,0,0,0.05); }

        .form-grid { display: grid; grid-template-columns: repeat(2, 1fr); gap: 16px; margin-bottom: 20px; }
        .form-group { display: flex; flex-direction: column; gap: 6px; }
        label { font-size: 0.85rem; font-weight: 600; color: var(--text-main); }
        input, select { padding: 10px 14px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 0.9rem; }
        
        .calc-box {
            grid-column: span 2;
            background: #eff6ff;
            border: 1px dashed #93c5fd;
            padding: 12px 16px;
            border-radius: 8px;
            font-size: 0.88rem;
            color: #1e40af;
            font-weight: 600;
        }

        .btn-action {
            background: var(--accent); color: white; padding: 12px 20px; border: none;
            border-radius: 8px; font-weight: 600; cursor: pointer; transition: background 0.2s;
        }
        .btn-action:hover { background: var(--accent-hover); }

        .btn-sm { padding: 6px 12px; font-size: 0.8rem; border-radius: 6px; border: none; cursor: pointer; font-weight: 600; }
        .btn-edit { background: #f59e0b; color: white; margin-right: 4px; }
        .btn-edit:hover { background: #d97706; }
        .btn-delete { background: #ef4444; color: white; }
        .btn-delete:hover { background: #dc2626; }

        table { width: 100%; border-collapse: collapse; margin-top: 12px; }
        th, td { padding: 12px 16px; text-align: left; border-bottom: 1px solid #e2e8f0; font-size: 0.9rem; }
        th { background: #f8fafc; color: var(--text-muted); font-weight: 600; }

        .badge { padding: 4px 10px; border-radius: 20px; font-size: 0.75rem; font-weight: 700; }
        .badge-entrada { background: #dcfce7; color: #15803d; }
        .badge-salida { background: #fee2e2; color: #b91c1c; }
        .badge-ajuste { background: #fef3c7; color: #b45309; }

        #alert { padding: 14px; border-radius: 8px; display: none; font-weight: 500; }
        .alert-err { background: #fef2f2; border: 1px solid #fca5a5; color: #991b1b; }
        .alert-ok { background: #f0fdf4; border: 1px solid #86efac; color: #166534; }

        .modal { display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(15,23,42,0.6); align-items:center; justify-content:center; z-index:100; }
        .modal-content { background:white; padding:24px; border-radius:12px; max-width:450px; width:90%; box-shadow:0 10px 25px -5px rgba(0,0,0,0.1); }
        .modal-buttons { display:flex; gap:12px; margin-top:20px; justify-content:flex-end; }
        .btn-modal { padding:10px 16px; border-radius:6px; border:none; font-weight:600; cursor:pointer; }
        .btn-pdf-m { background:#0284c7; color:white; text-decoration:none; display:inline-block; }
        .btn-excel-m { background:#16a34a; color:white; text-decoration:none; display:inline-block; }
        .btn-close { background:#94a3b8; color:white; }

        .search-box { margin-bottom: 16px; display: flex; gap: 10px; }
        .search-box input { flex: 1; }

        footer {
            background: #0f172a;
            color: #94a3b8;
            text-align: center;
            padding: 12px;
            font-size: 0.85rem;
            font-weight: 600;
            border-top: 1px solid #1e293b;
        }
    </style>
</head>
<body>

<div class="app-container">
    <aside>
        <h2>📦 Gestor Inventario v8.0</h2>
        <nav>
            <button class="active" onclick="verSeccion('movimientos', '🔄 Registro de Movimientos de Inventario', event)">🔄 Movimientos</button>
            <button onclick="verSeccion('productos', '➕ Alta y Registro de Nuevos Productos', event)">➕ Nuevo Producto</button>
            <!-- OPCIÓN AÑADIDA JUSTO DEBAJO DE NUEVO PRODUCTO -->
            <button onclick="verSeccion('editar-productos', '✏️ Gestión y Edición de Productos', event)">✏️ Editar Producto</button>
            <button onclick="verSeccion('categorias', '🏷️ Gestión de Categorías', event)">🏷️ Categorías</button>
            <button onclick="verSeccion('responsables', '👤 Gestión de Responsables', event)">👤 Responsables</button>
            <button onclick="verSeccion('kardex', '📋 Historial General de Kardex', event)">📋 Kardex General</button>
        </nav>

        <div class="pdf-group">
            <span>Reportes de Inventario</span>
            <button onclick="abrirModalExportar()" class="btn-pdf btn-pdf-accent">📊 Stock General <span>EXPORTAR</span></button>
            <a href="/api/pdf/bajo_stock" target="_blank" class="btn-pdf">⚠️ Bajo Stock <span>PDF</span></a>
            
            <span>Stock por Ubicación</span>
            <div id="pdf-ubicaciones-list" style="display:flex; flex-direction:column; gap:4px;"></div>

            <span>Historial Kardex</span>
            <a href="/api/pdf/ENTRADA" target="_blank" class="btn-pdf">📥 Solo Entradas <span>PDF</span></a>
            <a href="/api/pdf/SALIDA" target="_blank" class="btn-pdf">📤 Solo Salidas <span>PDF</span></a>
            <a href="/api/pdf/AJUSTE" target="_blank" class="btn-pdf">🛠️ Solo Ajustes <span>PDF</span></a>
        </div>
    </aside>

    <main>
        <div class="page-header">
            <h1 id="page-title">🔄 Registro de Movimientos de Inventario</h1>
            <span>Módulo Principal System v8.0</span>
        </div>

        <div id="alert"></div>

        <div class="kpi-grid">
            <div class="kpi-card">
                <span>Total SKU Activos</span>
                <h3 id="kpi-total">0</h3>
            </div>
            <div class="kpi-card">
                <span>Alerta Bajo Stock</span>
                <h3 id="kpi-bajo" style="color: var(--danger);">0</h3>
            </div>
            <div class="kpi-card">
                <span>Movimientos Totales</span>
                <h3 id="kpi-movs">0</h3>
            </div>
        </div>

        <!-- REGISTRO MOVIMIENTO -->
        <div id="sec-movimientos" class="content-card">
            <h3 style="margin-bottom: 16px;">Operación de Stock</h3>
            <form onsubmit="procesarMovimiento(event)">
                <div class="form-grid">
                    <div class="form-group">
                        <label>Producto</label>
                        <select id="mov-prod" required></select>
                    </div>
                    <div class="form-group">
                        <label>Tipo de Operación</label>
                        <select id="mov-tipo" onchange="evaluarCamposTipo()" required>
                            <option value="ENTRADA">ENTRADA (+)</option>
                            <option value="SALIDA">SALIDA (-)</option>
                            <option value="AJUSTE">AJUSTE (Fijar directo)</option>
                        </select>
                    </div>

                    <div class="form-group grp-entrada">
                        <label>Presentación de Entrada</label>
                        <input type="text" id="mov-pres" placeholder="Ej: Caja, Bulto, Paquete, Saco">
                    </div>
                    <div class="form-group grp-entrada">
                        <label>Unidad de Medida / Cant. por Presentación</label>
                        <select id="mov-um" onchange="calcularUnidadesEntrada()">
                            <option value="Unidades">Unidades Individuales (1:1)</option>
                            <option value="Caja x 6">Caja por 6 Unidades</option>
                            <option value="Caja x 12">Caja por 12 Unidades</option>
                            <option value="Caja x 24">Caja por 24 Unidades</option>
                            <option value="Paquete x 10">Paquete por 10 Unidades</option>
                            <option value="Kilogramos">Kilogramos (kg)</option>
                            <option value="Litros">Litros (L)</option>
                        </select>
                    </div>

                    <div class="form-group">
                        <label id="lbl-cant">Cantidad de Presentaciones / Bultos</label>
                        <input type="number" step="any" id="mov-cant" placeholder="Ej: 10" oninput="calcularUnidadesEntrada()" required>
                    </div>

                    <div class="form-group">
                        <label>Responsable</label>
                        <select id="mov-resp" class="select-responsables" required></select>
                    </div>

                    <div class="calc-box grp-entrada" id="calc-result">
                        💡 Total a ingresar al Stock: <span id="cant-calculada">0.00</span> Unidades.
                    </div>

                    <div class="form-group" id="grp-ubicacion" style="display:none;">
                        <label>Ubicación de Destino</label>
                        <select id="mov-ubi"></select>
                    </div>

                    <div class="form-group" id="grp-motivo" style="display:none;">
                        <label>Motivo del Ajuste</label>
                        <select id="mov-motivo">
                            <option value="Merma">Merma / Insumo Perdido</option>
                            <option value="Consumo de Empleado">Consumo de Empleado</option>
                        </select>
                    </div>
                </div>
                <button type="submit" class="btn-action">Procesar Operación</button>
            </form>

            <h3 style="margin-top: 32px; margin-bottom: 12px;">Existencias en Tiempo Real</h3>
            <table>
                <thead>
                    <tr>
                        <th>Código</th>
                        <th>Producto</th>
                        <th>Categoría</th>
                        <th>Stock Actual</th>
                        <th>Mínimo</th>
                    </tr>
                </thead>
                <tbody id="tbl-stock"></tbody>
            </table>
        </div>

        <!-- NUEVO PRODUCTO -->
        <div id="sec-productos" class="content-card" style="display:none;">
            <h3 style="margin-bottom: 16px;">Alta de Producto (Código Autogenerado)</h3>
            <form onsubmit="crearProducto(event)">
                <div class="form-grid">
                    <div class="form-group">
                        <label>Categoría (Define el Prefijo)</label>
                        <select id="prod-cat" required></select>
                    </div>
                    <div class="form-group">
                        <label>Nombre del Producto</label>
                        <input type="text" id="prod-nom" placeholder="Ej: Té Matcha 250g" required>
                    </div>

                    <div class="form-group">
                        <label>Presentación Inicial</label>
                        <input type="text" id="prod-pres" placeholder="Ej: Caja, Bulto, Unidades Individuales">
                    </div>
                    <div class="form-group">
                        <label>Unidad de Medida / Cant. por Presentación</label>
                        <select id="prod-um" onchange="calcularUnidadesNuevoProducto()">
                            <option value="Unidades">Unidades Individuales (1:1)</option>
                            <option value="Caja x 6">Caja por 6 Unidades</option>
                            <option value="Caja x 12">Caja por 12 Unidades</option>
                            <option value="Caja x 24">Caja por 24 Unidades</option>
                            <option value="Paquete x 10">Paquete por 10 Unidades</option>
                            <option value="Kilogramos">Kilogramos (kg)</option>
                            <option value="Litros">Litros (L)</option>
                        </select>
                    </div>

                    <div class="form-group">
                        <label>Stock Mínimo (Alerta)</label>
                        <input type="number" step="any" id="prod-min" value="15.0" required>
                    </div>
                    <div class="form-group">
                        <label>Stock Inicial (Presentaciones / Bultos)</label>
                        <input type="number" step="any" id="prod-init" value="0" oninput="calcularUnidadesNuevoProducto()" required>
                    </div>

                    <div class="form-group">
                        <label>Responsable de Alta</label>
                        <select id="prod-resp" class="select-responsables" required></select>
                    </div>

                    <div class="calc-box" id="calc-result-prod">
                        💡 Total Inicial a Ingresar al Stock: <span id="cant-calculada-prod">0.00</span> Unidades.
                    </div>
                </div>
                <button type="submit" class="btn-action">Guardar Producto</button>
            </form>
        </div>

        <!-- MÓDULO EDITAR PRODUCTO (CRUD COMPLETO) -->
        <div id="sec-editar-productos" class="content-card" style="display:none;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 16px;">
                <h3>Administración de Productos en Inventario</h3>
            </div>
            <div class="search-box">
                <input type="text" id="busqueda-prod" placeholder="🔍 Buscar producto por nombre o código SKU..." oninput="filtrarProductosCRUD()">
            </div>
            <table>
                <thead>
                    <tr>
                        <th>Código</th>
                        <th>Nombre del Producto</th>
                        <th>Categoría</th>
                        <th>Stock Actual</th>
                        <th>Stock Mínimo</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody id="tbl-crud-productos"></tbody>
            </table>
        </div>

        <!-- CRUD CATEGORÍAS -->
        <div id="sec-categorias" class="content-card" style="display:none;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 16px;">
                <h3>Administración de Categorías</h3>
                <button class="btn-action" onclick="abrirModalCategoria()">+ Nueva Categoría</button>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Nombre de Categoría</th>
                        <th>Prefijo</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody id="tbl-categorias"></tbody>
            </table>
        </div>

        <!-- CRUD RESPONSABLES -->
        <div id="sec-responsables" class="content-card" style="display:none;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom: 16px;">
                <h3>Administración de Personal Responsable</h3>
                <button class="btn-action" onclick="abrirModalResponsable()">+ Nuevo Responsable</button>
            </div>
            <table>
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Nombre Completo</th>
                        <th>Cargo / Función</th>
                        <th>Fecha Registro</th>
                        <th>Acciones</th>
                    </tr>
                </thead>
                <tbody id="tbl-responsables"></tbody>
            </table>
        </div>

        <!-- KARDEX -->
        <div id="sec-kardex" class="content-card" style="display:none;">
            <h3>Historial Completo (Kardex)</h3>
            <table>
                <thead>
                    <tr>
                        <th>Fecha</th>
                        <th>Código</th>
                        <th>Producto</th>
                        <th>Tipo</th>
                        <th>Presentación / U.M</th>
                        <th>Cantidad Netas</th>
                        <th>Stock Final</th>
                        <th>Ubicación / Motivo</th>
                        <th>Responsable</th>
                    </tr>
                </thead>
                <tbody id="tbl-kardex"></tbody>
            </table>
        </div>
    </main>
</div>

<!-- MODAL EDITAR PRODUCTO -->
<div id="modal-edit-prod" class="modal">
    <div class="modal-content">
        <h3>Editar Producto</h3>
        <form onsubmit="guardarEdicionProducto(event)" style="margin-top:16px; text-align:left;">
            <input type="hidden" id="edit-prod-id">
            <div class="form-group" style="margin-bottom:12px;">
                <label>Código SKU (No Modificable)</label>
                <input type="text" id="edit-prod-cod" disabled style="background:#e2e8f0;">
            </div>
            <div class="form-group" style="margin-bottom:12px;">
                <label>Nombre del Producto</label>
                <input type="text" id="edit-prod-nom" required>
            </div>
            <div class="form-group" style="margin-bottom:12px;">
                <label>Stock Mínimo</label>
                <input type="number" step="any" id="edit-prod-min" required>
            </div>
            <div class="modal-buttons">
                <button type="submit" class="btn-modal btn-pdf-m">Guardar Cambios</button>
                <button type="button" onclick="cerrarModalEditProducto()" class="btn-modal btn-close">Cancelar</button>
            </div>
        </form>
    </div>
</div>

<!-- MODAL CONFIRMACIÓN ELIMINAR PRODUCTO -->
<div id="modal-confirm-delete" class="modal">
    <div class="modal-content" style="text-align:center;">
        <h3 style="color:var(--danger);">⚠️ Confirmar Eliminación</h3>
        <p style="margin-top:12px; font-size:0.9rem; color:var(--text-main);" id="txt-confirm-delete">
            ¿Estás seguro de que deseas eliminar este producto del inventario?
        </p>
        <p style="margin-top:6px; font-size:0.8rem; color:var(--text-muted);">
            Se eliminará el producto, sus existencias y su historial kardex de forma permanente.
        </p>
        <div class="modal-buttons" style="justify-content:center; margin-top:20px;">
            <button id="btn-confirm-delete-action" class="btn-modal btn-delete">Sí, Eliminar</button>
            <button onclick="cerrarModalConfirmDelete()" class="btn-modal btn-close">Cancelar</button>
        </div>
    </div>
</div>

<!-- MODAL CATEGORÍA (CREAR/EDITAR) -->
<div id="modal-cat" class="modal">
    <div class="modal-content">
        <h3 id="modal-cat-title">Registrar Categoría</h3>
        <form onsubmit="guardarCategoria(event)" style="margin-top:16px; text-align:left;">
            <input type="hidden" id="cat-id">
            <div class="form-group" style="margin-bottom:12px;">
                <label>Nombre de la Categoría</label>
                <input type="text" id="cat-nom" required placeholder="Ej: Repostería">
            </div>
            <div class="form-group" style="margin-bottom:12px;">
                <label>Prefijo (Código SKU)</label>
                <input type="text" id="cat-pref" placeholder="Ej: REP" required maxlength="10">
            </div>
            <div class="modal-buttons">
                <button type="submit" class="btn-modal btn-pdf-m">Guardar</button>
                <button type="button" onclick="cerrarModalCategoria()" class="btn-modal btn-close">Cancelar</button>
            </div>
        </form>
    </div>
</div>

<!-- MODAL RESPONSABLE (CREAR/EDITAR) -->
<div id="modal-resp" class="modal">
    <div class="modal-content">
        <h3 id="modal-resp-title">Registrar Responsable</h3>
        <form onsubmit="guardarResponsable(event)" style="margin-top:16px; text-align:left;">
            <input type="hidden" id="resp-id">
            <div class="form-group" style="margin-bottom:12px;">
                <label>Nombre Completo</label>
                <input type="text" id="resp-nom" required>
            </div>
            <div class="form-group" style="margin-bottom:12px;">
                <label>Cargo / Rol</label>
                <input type="text" id="resp-cargo" placeholder="Ej: Supervisor de Almacén" required>
            </div>
            <div class="modal-buttons">
                <button type="submit" class="btn-modal btn-pdf-m">Guardar</button>
                <button type="button" onclick="cerrarModalResponsable()" class="btn-modal btn-close">Cancelar</button>
            </div>
        </form>
    </div>
</div>

<!-- MODAL SELECCIÓN EXPORTACIÓN -->
<div id="modal-export" class="modal">
    <div class="modal-content" style="text-align:center;">
        <h3>Exportar Stock General</h3>
        <p style="margin-top:8px; font-size:0.9rem; color:var(--text-muted);">¿En qué formato deseas generar el reporte de inventario?</p>
        <div class="modal-buttons" style="justify-content:center;">
            <a href="/api/pdf/stock_general" target="_blank" onclick="cerrarModalExportar()" class="btn-modal btn-pdf-m">📄 Descargar PDF</a>
            <a href="/api/excel/stock_general" download="Stock_General.xls" onclick="cerrarModalExportar()" class="btn-modal btn-excel-m">📊 Descargar Excel</a>
            <button onclick="cerrarModalExportar()" class="btn-modal btn-close">Cancelar</button>
        </div>
    </div>
</div>

<footer>
    DESARROLLADO POR ANDRES
</footer>

<script>
    let productosCache = [];

    function escapeHtml(text) {
        if (!text) return '';
        return String(text)
            .replace(/&/g, "&amp;")
            .replace(/</g, "&lt;")
            .replace(/>/g, "&gt;")
            .replace(/"/g, "&quot;")
            .replace(/'/g, "&#039;");
    }

    document.addEventListener('DOMContentLoaded', () => {
        cargarCatalogos();
        cargarResponsables();
        cargarDatos();
        evaluarCamposTipo();
    });

    function verSeccion(sec, titulo, ev) {
        document.getElementById('sec-movimientos').style.display = sec === 'movimientos' ? 'block' : 'none';
        document.getElementById('sec-productos').style.display = sec === 'productos' ? 'block' : 'none';
        document.getElementById('sec-editar-productos').style.display = sec === 'editar-productos' ? 'block' : 'none';
        document.getElementById('sec-categorias').style.display = sec === 'categorias' ? 'block' : 'none';
        document.getElementById('sec-responsables').style.display = sec === 'responsables' ? 'block' : 'none';
        document.getElementById('sec-kardex').style.display = sec === 'kardex' ? 'block' : 'none';
        
        document.getElementById('page-title').innerText = titulo;

        document.querySelectorAll('nav button').forEach(b => b.classList.remove('active'));
        if (ev) ev.target.classList.add('active');
    }

    function evaluarCamposTipo() {
        const tipo = document.getElementById('mov-tipo').value;
        const elementosEntrada = document.querySelectorAll('.grp-entrada');
        
        if (tipo === 'ENTRADA') {
            elementosEntrada.forEach(el => el.style.display = 'flex');
            document.getElementById('lbl-cant').innerText = 'Cantidad de Presentaciones / Bultos';
        } else {
            elementosEntrada.forEach(el => el.style.display = 'none');
            document.getElementById('lbl-cant').innerText = 'Cantidad Total';
        }

        document.getElementById('grp-ubicacion').style.display = tipo === 'SALIDA' ? 'flex' : 'none';
        document.getElementById('grp-motivo').style.display = tipo === 'AJUSTE' ? 'flex' : 'none';
        
        calcularUnidadesEntrada();
    }

    function obtenerFactorConversion(um) {
        if (um === 'Caja x 6') return 6;
        if (um === 'Caja x 12') return 12;
        if (um === 'Caja x 24') return 24;
        if (um === 'Paquete x 10') return 10;
        return 1;
    }

    function calcularUnidadesEntrada() {
        const tipo = document.getElementById('mov-tipo').value;
        if (tipo !== 'ENTRADA') return;

        const cant = parseFloat(document.getElementById('mov-cant').value) || 0;
        const um = document.getElementById('mov-um').value;
        const totalNeto = cant * obtenerFactorConversion(um);
        document.getElementById('cant-calculada').innerText = totalNeto.toFixed(2);
    }

    function calcularUnidadesNuevoProducto() {
        const cant = parseFloat(document.getElementById('prod-init').value) || 0;
        const um = document.getElementById('prod-um').value;
        const totalNeto = cant * obtenerFactorConversion(um);
        document.getElementById('cant-calculada-prod').innerText = totalNeto.toFixed(2);
    }

    function abrirModalExportar() { document.getElementById('modal-export').style.display = 'flex'; }
    function cerrarModalExportar() { document.getElementById('modal-export').style.display = 'none'; }

    function mostrarAlerta(txt, esError = false) {
        const el = document.getElementById('alert');
        el.className = esError ? 'alert-err' : 'alert-ok';
        el.innerText = txt;
        el.style.display = 'block';
        setTimeout(() => { el.style.display = 'none'; }, 4000);
    }

    async function cargarCatalogos() {
        const resCat = await fetch('/api/categorias');
        const cats = await resCat.json();
        
        document.getElementById('prod-cat').innerHTML = cats.map(c => 
            `<option value="${c.id}">${escapeHtml(c.nombre)} (${escapeHtml(c.prefijo)})</option>`
        ).join('');

        document.getElementById('tbl-categorias').innerHTML = cats.map(c => `
            <tr>
                <td><b>${c.id}</b></td>
                <td>${escapeHtml(c.nombre)}</td>
                <td><b>${escapeHtml(c.prefijo)}</b></td>
                <td>
                    <button class="btn-sm btn-edit" onclick='editarCategoria(${JSON.stringify(c)})'>Editar</button>
                    <button class="btn-sm btn-delete" onclick="eliminarCategoria(${c.id})">Eliminar</button>
                </td>
            </tr>
        `).join('');

        const resUbi = await fetch('/api/ubicaciones');
        const ubis = await resUbi.json();
        document.getElementById('mov-ubi').innerHTML = ubis.map(u => `<option value="${u.id}">${escapeHtml(u.nombre)}</option>`).join('');

        document.getElementById('pdf-ubicaciones-list').innerHTML = ubis.map(u => 
            `<a href="/api/pdf/stock_ubicacion?ubi_id=${u.id}" target="_blank" class="btn-pdf">❄️ ${escapeHtml(u.nombre)} <span>PDF</span></a>`
        ).join('');
    }

    function abrirModalCategoria() {
        document.getElementById('cat-id').value = '';
        document.getElementById('cat-nom').value = '';
        document.getElementById('cat-pref').value = '';
        document.getElementById('modal-cat-title').innerText = 'Registrar Categoría';
        document.getElementById('modal-cat').style.display = 'flex';
    }

    function cerrarModalCategoria() {
        document.getElementById('modal-cat').style.display = 'none';
    }

    function editarCategoria(c) {
        document.getElementById('cat-id').value = c.id;
        document.getElementById('cat-nom').value = c.nombre;
        document.getElementById('cat-pref').value = c.prefijo;
        document.getElementById('modal-cat-title').innerText = 'Editar Categoría';
        document.getElementById('modal-cat').style.display = 'flex';
    }

    async function guardarCategoria(e) {
        e.preventDefault();
        const id = document.getElementById('cat-id').value;
        const payload = {
            nombre: document.getElementById('cat-nom').value,
            prefijo: document.getElementById('cat-pref').value.toUpperCase()
        };

        const method = id ? 'PUT' : 'POST';
        if (id) payload.id = parseInt(id);

        const res = await fetch('/api/categorias', {
            method: method,
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(payload)
        });

        const resp = await res.json();
        if (res.ok) {
            mostrarAlerta(id ? 'Categoría actualizada' : 'Categoría creada');
            cerrarModalCategoria();
            cargarCatalogos();
        } else {
            mostrarAlerta(resp.error, true);
        }
    }

    async function eliminarCategoria(id) {
        if (!confirm('¿Estás seguro de eliminar esta categoría?')) return;

        const res = await fetch(`/api/categorias?id=${id}`, { method: 'DELETE' });
        const resp = await res.json();
        if (res.ok) {
            mostrarAlerta('Categoría eliminada con éxito');
            cargarCatalogos();
        } else {
            mostrarAlerta(resp.error, true);
        }
    }

    async function cargarResponsables() {
        const res = await fetch('/api/responsables');
        const lista = await res.json();
        
        const selects = document.querySelectorAll('.select-responsables');
        selects.forEach(s => {
            s.innerHTML = lista.map(r => `<option value="${escapeHtml(r.nombre)}">${escapeHtml(r.nombre)} (${escapeHtml(r.cargo)})</option>`).join('');
        });

        document.getElementById('tbl-responsables').innerHTML = lista.map(r => `
            <tr>
                <td><b>${r.id}</b></td>
                <td>${escapeHtml(r.nombre)}</td>
                <td>${escapeHtml(r.cargo)}</td>
                <td>${String(r.fecha_registro).substring(0, 10)}</td>
                <td>
                    <button class="btn-sm btn-edit" onclick='editarResponsable(${JSON.stringify(r)})'>Editar</button>
                    <button class="btn-sm btn-delete" onclick="eliminarResponsable(${r.id})">Eliminar</button>
                </td>
            </tr>
        `).join('');
    }

    function abrirModalResponsable() {
        document.getElementById('resp-id').value = '';
        document.getElementById('resp-nom').value = '';
        document.getElementById('resp-cargo').value = '';
        document.getElementById('modal-resp-title').innerText = 'Registrar Responsable';
        document.getElementById('modal-resp').style.display = 'flex';
    }

    function cerrarModalResponsable() {
        document.getElementById('modal-resp').style.display = 'none';
    }

    function editarResponsable(r) {
        document.getElementById('resp-id').value = r.id;
        document.getElementById('resp-nom').value = r.nombre;
        document.getElementById('resp-cargo').value = r.cargo;
        document.getElementById('modal-resp-title').innerText = 'Editar Responsable';
        document.getElementById('modal-resp').style.display = 'flex';
    }

    async function guardarResponsable(e) {
        e.preventDefault();
        const id = document.getElementById('resp-id').value;
        const payload = {
            nombre: document.getElementById('resp-nom').value,
            cargo: document.getElementById('resp-cargo').value
        };

        const method = id ? 'PUT' : 'POST';
        if (id) payload.id = parseInt(id);

        const res = await fetch('/api/responsables', {
            method: method,
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(payload)
        });

        const resp = await res.json();
        if (res.ok) {
            mostrarAlerta(id ? 'Responsable actualizado' : 'Responsable creado');
            cerrarModalResponsable();
            cargarResponsables();
        } else {
            mostrarAlerta(resp.error, true);
        }
    }

    async function eliminarResponsable(id) {
        if (!confirm('¿Estás seguro de eliminar este responsable?')) return;

        const res = await fetch(`/api/responsables?id=${id}`, { method: 'DELETE' });
        const resp = await res.json();
        if (res.ok) {
            mostrarAlerta('Responsable eliminado');
            cargarResponsables();
        } else {
            mostrarAlerta(resp.error, true);
        }
    }

    async function cargarDatos() {
        try {
            const resStock = await fetch('/api/stock');
            productosCache = await resStock.json();
            
            const resKardex = await fetch('/api/kardex');
            const dataKardex = await resKardex.json();

            document.getElementById('kpi-total').innerText = productosCache.length;
            document.getElementById('kpi-bajo').innerText = productosCache.filter(x => x.cantidad < x.stock_minimo).length;
            document.getElementById('kpi-movs').innerText = dataKardex.length;

            const sel = document.getElementById('mov-prod');
            sel.innerHTML = productosCache.map(p => `<option value="${p.id}">${escapeHtml(p.codigo)} - ${escapeHtml(p.nombre)} (Stock: ${p.cantidad})</option>`).join('');

            document.getElementById('tbl-stock').innerHTML = productosCache.map(p => `
                <tr>
                    <td><b>${escapeHtml(p.codigo)}</b></td>
                    <td>${escapeHtml(p.nombre)}</td>
                    <td>${escapeHtml(p.categoria)}</td>
                    <td><b style="color:${p.cantidad < p.stock_minimo ? '#dc2626' : '#16a34a'}">${p.cantidad.toFixed(2)}</b></td>
                    <td>${p.stock_minimo.toFixed(2)}</td>
                </tr>
            `).join('');

            renderizarTablaCRUDProductos(productosCache);

            document.getElementById('tbl-kardex').innerHTML = dataKardex.map(k => {
                let detalle = '-';
                if (k.tipo_movimiento === 'SALIDA') detalle = k.ubicacion || 'N/A';
                if (k.tipo_movimiento === 'AJUSTE') detalle = k.motivo || 'N/A';

                const pres = k.presentacion ? `${escapeHtml(k.presentacion)} (${escapeHtml(k.unidad_medida || 'Unid')})` : '-';

                return `
                    <tr>
                        <td>${String(k.fecha).substring(0, 19)}</td>
                        <td><b>${escapeHtml(k.codigo)}</b></td>
                        <td>${escapeHtml(k.nombre)}</td>
                        <td><span class="badge badge-${k.tipo_movimiento.toLowerCase()}">${k.tipo_movimiento}</span></td>
                        <td>${pres}</td>
                        <td><b>${k.cantidad.toFixed(2)}</b></td>
                        <td><b>${k.stock_resultante.toFixed(2)}</b></td>
                        <td>${escapeHtml(detalle)}</td>
                        <td>${escapeHtml(k.responsable)}</td>
                    </tr>
                `;
            }).join('');

        } catch (e) {
            mostrarAlerta('Error cargando información', true);
        }
    }

    /* === FUNCIONES DEL CRUD DE PRODUCTOS (EDITAR / ELIMINAR) === */
    function renderizarTablaCRUDProductos(lista) {
        document.getElementById('tbl-crud-productos').innerHTML = lista.map(p => `
            <tr>
                <td><b>${escapeHtml(p.codigo)}</b></td>
                <td>${escapeHtml(p.nombre)}</td>
                <td>${escapeHtml(p.categoria)}</td>
                <td><b style="color:${p.cantidad < p.stock_minimo ? '#dc2626' : '#16a34a'}">${p.cantidad.toFixed(2)}</b></td>
                <td>${p.stock_minimo.toFixed(2)}</td>
                <td>
                    <button class="btn-sm btn-edit" onclick='abrirModalEditProducto(${JSON.stringify(p)})'>Editar</button>
                    <button class="btn-sm btn-delete" onclick="solicitarEliminarProducto(${p.id}, '${escapeHtml(p.nombre)}')">Eliminar</button>
                </td>
            </tr>
        `).join('');
    }

    function filtrarProductosCRUD() {
        const query = document.getElementById('busqueda-prod').value.toLowerCase();
        const filtrados = productosCache.filter(p => 
            p.nombre.toLowerCase().includes(query) || p.codigo.toLowerCase().includes(query)
        );
        renderizarTablaCRUDProductos(filtrados);
    }

    function abrirModalEditProducto(p) {
        document.getElementById('edit-prod-id').value = p.id;
        document.getElementById('edit-prod-cod').value = p.codigo;
        document.getElementById('edit-prod-nom').value = p.nombre;
        document.getElementById('edit-prod-min').value = p.stock_minimo;
        document.getElementById('modal-edit-prod').style.display = 'flex';
    }

    function cerrarModalEditProducto() {
        document.getElementById('modal-edit-prod').style.display = 'none';
    }

    async function guardarEdicionProducto(e) {
        e.preventDefault();
        const id = parseInt(document.getElementById('edit-prod-id').value);
        const payload = {
            id: id,
            nombre: document.getElementById('edit-prod-nom').value,
            stock_minimo: parseFloat(document.getElementById('edit-prod-min').value)
        };

        const res = await fetch('/api/productos', {
            method: 'PUT',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(payload)
        });

        const resp = await res.json();
        if (res.ok) {
            mostrarAlerta('Producto editado exitosamente con confirmación');
            cerrarModalEditProducto();
            cargarDatos();
        } else {
            mostrarAlerta(resp.error, true);
        }
    }

    function solicitarEliminarProducto(id, nombre) {
        document.getElementById('txt-confirm-delete').innerText = `¿Estás seguro de eliminar el producto "${nombre}" del inventario?`;
        const btn = document.getElementById('btn-confirm-delete-action');
        btn.onclick = () => ejecutarEliminacionProducto(id);
        document.getElementById('modal-confirm-delete').style.display = 'flex';
    }

    function cerrarModalConfirmDelete() {
        document.getElementById('modal-confirm-delete').style.display = 'none';
    }

    async function ejecutarEliminacionProducto(id) {
        const res = await fetch(`/api/productos?id=${id}`, { method: 'DELETE' });
        const resp = await res.json();
        cerrarModalConfirmDelete();

        if (res.ok) {
            mostrarAlerta('Producto y sus datos asociados eliminados del inventario');
            cargarDatos();
        } else {
            mostrarAlerta(resp.error, true);
        }
    }

    async function procesarMovimiento(e) {
        e.preventDefault();
        const tipo = document.getElementById('mov-tipo').value;
        const cantIngresada = parseFloat(document.getElementById('mov-cant').value);
        
        let cantFinal = cantIngresada;
        if (tipo === 'ENTRADA') {
            const um = document.getElementById('mov-um').value;
            cantFinal = cantIngresada * obtenerFactorConversion(um);
        }

        const payload = {
            producto_id: parseInt(document.getElementById('mov-prod').value),
            tipo: tipo,
            cantidad: cantFinal,
            presentacion: tipo === 'ENTRADA' ? document.getElementById('mov-pres').value : null,
            unidad_medida: tipo === 'ENTRADA' ? document.getElementById('mov-um').value : null,
            responsable: document.getElementById('mov-resp').value,
            ubicacion_id: tipo === 'SALIDA' ? parseInt(document.getElementById('mov-ubi').value) : null,
            motivo: tipo === 'AJUSTE' ? document.getElementById('mov-motivo').value : null
        };

        const res = await fetch('/api/movimientos', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(payload)
        });

        const resp = await res.json();
        if (res.ok) {
            mostrarAlerta('Movimiento registrado con éxito');
            document.getElementById('mov-cant').value = '';
            document.getElementById('mov-pres').value = '';
            cargarDatos();
        } else {
            mostrarAlerta(resp.error, true);
        }
    }

    async function crearProducto(e) {
        e.preventDefault();
        const cantInit = parseFloat(document.getElementById('prod-init').value) || 0;
        const um = document.getElementById('prod-um').value;
        const cantFinal = cantInit * obtenerFactorConversion(um);

        const payload = {
            categoria_id: parseInt(document.getElementById('prod-cat').value),
            nombre: document.getElementById('prod-nom').value,
            stock_minimo: parseFloat(document.getElementById('prod-min').value),
            stock_inicial: cantFinal,
            presentacion: document.getElementById('prod-pres').value || 'Unidad Individual',
            unidad_medida: um,
            responsable: document.getElementById('prod-resp').value
        };

        const res = await fetch('/api/productos', {
            method: 'POST',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify(payload)
        });

        const resp = await res.json();
        if (res.ok) {
            mostrarAlerta('Producto creado correctamente');
            document.getElementById('prod-nom').value = '';
            document.getElementById('prod-pres').value = '';
            document.getElementById('prod-init').value = '0';
            cargarDatos();
            verSeccion('movimientos', '🔄 Registro de Movimientos de Inventario');
        } else {
            mostrarAlerta(resp.error, true);
        }
    }
</script>
</body>
</html>
"""

# =============================================================================
# 4. MANEJADOR HTTP Y ENDPOINTS API
# =============================================================================

class AppRequestHandler(BaseHTTPRequestHandler):

    def _json(self, data, status=200):
        body = json.dumps(data, default=str).encode('utf-8')
        self.send_response(status)
        self.send_header('Content-Type', 'application/json')
        self.send_header('Content-Length', str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        query = urllib.parse.parse_qs(parsed_url.query)

        if path in ("/", "/index.html"):
            body = NUEVA_INTERFAZ_HTML.encode('utf-8')
            self.send_response(200)
            self.send_header('Content-Type', 'text/html; charset=utf-8')
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        elif path == "/api/responsables":
            with get_db_context() as conn:
                cur = conn.cursor(dictionary=True)
                cur.execute("SELECT id, nombre, cargo, fecha_registro FROM responsables ORDER BY nombre ASC;")
                res = cur.fetchall()
                cur.close()
            self._json(res)

        elif path == "/api/categorias":
            with get_db_context() as conn:
                cur = conn.cursor(dictionary=True)
                cur.execute("SELECT id, nombre, prefijo FROM categorias ORDER BY nombre ASC;")
                res = cur.fetchall()
                cur.close()
            self._json(res)

        elif path == "/api/ubicaciones":
            with get_db_context() as conn:
                cur = conn.cursor(dictionary=True)
                cur.execute("SELECT id, nombre FROM ubicaciones ORDER BY nombre ASC;")
                res = cur.fetchall()
                cur.close()
            self._json(res)

        elif path == "/api/stock":
            with get_db_context() as conn:
                cur = conn.cursor(dictionary=True)
                cur.execute("""
                    SELECT p.id, p.codigo, p.nombre, c.nombre as categoria, COALESCE(s.cantidad, 0) as cantidad, p.stock_minimo
                    FROM productos p
                    JOIN categorias c ON p.categoria_id = c.id
                    LEFT JOIN stock s ON p.id = s.producto_id
                    ORDER BY p.id DESC;
                """)
                res = cur.fetchall()
                for row in res:
                    row['cantidad'] = float(row['cantidad'])
                    row['stock_minimo'] = float(row['stock_minimo'])
                cur.close()
            self._json(res)

        elif path == "/api/kardex":
            with get_db_context() as conn:
                cur = conn.cursor(dictionary=True)
                cur.execute("""
                    SELECT k.id, k.fecha, p.codigo, p.nombre, k.tipo_movimiento, k.cantidad, k.stock_resultante,
                           k.presentacion, k.unidad_medida, u.nombre as ubicacion, k.motivo, k.responsable
                    FROM historial_kardex k
                    JOIN productos p ON k.producto_id = p.id
                    LEFT JOIN ubicaciones u ON k.ubicacion_id = u.id
                    ORDER BY k.id DESC;
                """)
                res = cur.fetchall()
                for row in res:
                    row['cantidad'] = float(row['cantidad'])
                    row['stock_resultante'] = float(row['stock_resultante'])
                cur.close()
            self._json(res)

        elif path == "/api/excel/stock_general":
            excel_bytes = generar_excel_stock_general()
            self.send_response(200)
            self.send_header('Content-Type', 'application/vnd.ms-excel')
            self.send_header('Content-Disposition', 'attachment; filename="Stock_General.xls"')
            self.send_header('Content-Length', str(len(excel_bytes)))
            self.end_headers()
            self.wfile.write(excel_bytes)

        elif path.startswith("/api/pdf/"):
            reporte = path.replace("/api/pdf/", "")
            param = query.get('ubi_id', [None])[0]

            pdf_buffer = generar_pdf_filtrado(reporte, param)
            if pdf_buffer:
                pdf_bytes = pdf_buffer.getvalue()
                self.send_response(200)
                self.send_header('Content-Type', 'application/pdf')
                self.send_header('Content-Length', str(len(pdf_bytes)))
                self.end_headers()
                self.wfile.write(pdf_bytes)
            else:
                self._json({"error": "Error generando PDF"}, 500)
        else:
            self._json({"error": "Ruta no encontrada"}, 404)

    def do_POST(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        data = self._read_json()

        if path == "/api/categorias":
            with get_db_context() as conn:
                cur = conn.cursor()
                try:
                    cur.execute("INSERT INTO categorias (nombre, prefijo) VALUES (%s, %s);", 
                                (data['nombre'], data['prefijo'].upper()))
                    conn.commit()
                    self._json({"status": "ok", "id": cur.lastrowid})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

        elif path == "/api/responsables":
            with get_db_context() as conn:
                cur = conn.cursor()
                try:
                    cur.execute("INSERT INTO responsables (nombre, cargo) VALUES (%s, %s);", (data['nombre'], data['cargo']))
                    conn.commit()
                    self._json({"status": "ok", "id": cur.lastrowid})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

        elif path == "/api/productos":
            with get_db_context() as conn:
                cur = conn.cursor(dictionary=True)
                try:
                    conn.start_transaction()
                    cur.execute("SELECT prefijo FROM categorias WHERE id = %s;", (data['categoria_id'],))
                    cat = cur.fetchone()
                    if not cat:
                        raise Exception("Categoría no encontrada.")

                    prefijo = cat['prefijo']
                    cur.execute("SELECT COALESCE(MAX(id), 0) + 1 as cnt FROM productos WHERE categoria_id = %s;", (data['categoria_id'],))
                    count = cur.fetchone()['cnt']
                    codigo = f"{prefijo}-{count:03d}"

                    cur.execute("INSERT INTO productos (codigo, nombre, categoria_id, stock_minimo) VALUES (%s, %s, %s, %s);",
                                (codigo, data['nombre'], data['categoria_id'], data.get('stock_minimo', 15.0)))
                    p_id = cur.lastrowid

                    stock_initial = float(data.get('stock_inicial', 0.0))
                    cur.execute("INSERT INTO stock (producto_id, cantidad) VALUES (%s, %s);", (p_id, stock_initial))

                    if stock_initial > 0:
                        cur.execute("""
                            INSERT INTO historial_kardex (producto_id, tipo_movimiento, cantidad, stock_resultante, presentacion, unidad_medida, responsable)
                            VALUES (%s, 'ENTRADA', %s, %s, %s, %s, %s);
                        """, (p_id, stock_initial, stock_initial, data.get('presentacion'), data.get('unidad_medida'), data['responsable']))

                    conn.commit()
                    self._json({"status": "ok", "codigo": codigo})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

        elif path == "/api/movimientos":
            with get_db_context() as conn:
                cur = conn.cursor(dictionary=True)
                try:
                    conn.start_transaction()
                    p_id = data['producto_id']
                    tipo = data['tipo']
                    cant = float(data['cantidad'])

                    cur.execute("SELECT cantidad FROM stock WHERE producto_id = %s FOR UPDATE;", (p_id,))
                    row = cur.fetchone()
                    curr_stock = float(row['cantidad']) if row else 0.0

                    if tipo == 'ENTRADA':
                        new_stock = curr_stock + cant
                    elif tipo == 'SALIDA':
                        if cant > curr_stock:
                            raise Exception(f"Stock insuficiente. Stock actual: {curr_stock}")
                        new_stock = curr_stock - cant
                    elif tipo == 'AJUSTE':
                        new_stock = cant

                    cur.execute("UPDATE stock SET cantidad = %s WHERE producto_id = %s;", (new_stock, p_id))

                    cur.execute("""
                        INSERT INTO historial_kardex (producto_id, tipo_movimiento, cantidad, stock_resultante, presentacion, unidad_medida, ubicacion_id, motivo, responsable)
                        VALUES (%s, %s, %s, %s, %s, %s, %s, %s, %s);
                    """, (p_id, tipo, cant, new_stock, data.get('presentacion'), data.get('unidad_medida'), data.get('ubicacion_id'), data.get('motivo'), data['responsable']))

                    conn.commit()
                    self._json({"status": "ok", "nuevo_stock": new_stock})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

    def do_PUT(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        data = self._read_json()

        if path == "/api/productos":
            with get_db_context() as conn:
                cur = conn.cursor()
                try:
                    cur.execute("UPDATE productos SET nombre = %s, stock_minimo = %s WHERE id = %s;", 
                                (data['nombre'], data['stock_minimo'], data['id']))
                    conn.commit()
                    self._json({"status": "ok"})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

        elif path == "/api/categorias":
            with get_db_context() as conn:
                cur = conn.cursor()
                try:
                    cur.execute("UPDATE categorias SET nombre = %s, prefijo = %s WHERE id = %s;", 
                                (data['nombre'], data['prefijo'].upper(), data['id']))
                    conn.commit()
                    self._json({"status": "ok"})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

        elif path == "/api/responsables":
            with get_db_context() as conn:
                cur = conn.cursor()
                try:
                    cur.execute("UPDATE responsables SET nombre = %s, cargo = %s WHERE id = %s;", (data['nombre'], data['cargo'], data['id']))
                    conn.commit()
                    self._json({"status": "ok"})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

    def do_DELETE(self):
        parsed_url = urllib.parse.urlparse(self.path)
        path = parsed_url.path
        query = urllib.parse.parse_qs(parsed_url.query)

        if path == "/api/productos":
            p_id = query.get('id', [None])[0]
            if not p_id:
                self._json({"error": "ID de producto requerido"}, 400)
                return

            with get_db_context() as conn:
                cur = conn.cursor()
                try:
                    conn.start_transaction()
                    # Eliminación en cascada manual respetando claves foráneas
                    cur.execute("DELETE FROM historial_kardex WHERE producto_id = %s;", (p_id,))
                    cur.execute("DELETE FROM stock WHERE producto_id = %s;", (p_id,))
                    cur.execute("DELETE FROM productos WHERE id = %s;", (p_id,))
                    conn.commit()
                    self._json({"status": "ok"})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

        elif path == "/api/categorias":
            c_id = query.get('id', [None])[0]
            if not c_id:
                self._json({"error": "ID de categoría requerido"}, 400)
                return

            with get_db_context() as conn:
                cur = conn.cursor(dictionary=True)
                try:
                    cur.execute("SELECT COUNT(*) as total FROM productos WHERE categoria_id = %s;", (c_id,))
                    if cur.fetchone()['total'] > 0:
                        raise Exception("No se puede eliminar la categoría porque tiene productos registrados asimilados.")

                    cur.execute("DELETE FROM categorias WHERE id = %s;", (c_id,))
                    conn.commit()
                    self._json({"status": "ok"})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

        elif path == "/api/responsables":
            r_id = query.get('id', [None])[0]
            if not r_id:
                self._json({"error": "ID de responsable requerido"}, 400)
                return

            with get_db_context() as conn:
                cur = conn.cursor()
                try:
                    cur.execute("DELETE FROM responsables WHERE id = %s;", (r_id,))
                    conn.commit()
                    self._json({"status": "ok"})
                except Exception as e:
                    conn.rollback()
                    self._json({"error": str(e)}, 400)
                finally:
                    cur.close()

    def _read_json(self):
        content_length = int(self.headers.get('Content-Length', 0))
        post_data = self.rfile.read(content_length) if content_length > 0 else b'{}'
        return json.loads(post_data.decode('utf-8'))

# =============================================================================
# 5. INICIALIZACIÓN
# =============================================================================

def abrir_navegador():
    time.sleep(1.2)
    webbrowser.open(f"http://localhost:{PORT}")

    # Adaptador WSGI para Gunicorn / Render
def application(environ, start_response):
    # Inicializa las tablas en Aiven si aún no existen
    try:
        init_db_pool()
        init_db()
    except Exception:
        pass

    # Maneja la petición
    import io
    from wsgiref.handlers import SimpleHandler

    class DummyServer:
        def __init__(self, environ):
            self.base_environ = environ

    stdout = io.BytesIO()
    stderr = io.BytesIO()

    handler = SimpleHandler(
        environ['wsgi.input'],
        stdout,
        stderr,
        environ
    )
    handler.run(AppRequestHandler)

    # Extrae el código de respuesta HTTP y encabezados
    output = stdout.getvalue()
    parts = output.split(b'\r\n\r\n', 1)
    header_bytes = parts[0]
    body = parts[1] if len(parts) > 1 else b''

    lines = header_bytes.split(b'\r\n')
    status_line = lines[0].decode('iso-8859-1')
    status = status_line.split(' ', 1)[1] if ' ' in status_line else '200 OK'

    headers = []
    for line in lines[1:]:
        if b':' in line:
            k, v = line.split(b':', 1)
            headers.append((k.decode('iso-8859-1').strip(), v.decode('iso-8859-1').strip()))

    start_response(status, headers)
    return [body]

if __name__ == '__main__':
    print("Iniciando Sistema de Inventarios v8.0 (MySQL + CRUD Productos)...")
    init_db_pool()
    init_db()

    server = ThreadingHTTPServer(('0.0.0.0', PORT), AppRequestHandler)
    print(f"Servidor HTTP corriendo exitosamente en http://localhost:{PORT}")

    threading.Thread(target=abrir_navegador, daemon=True).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServidor detenido.")
        server.shutdown()
        sys.exit(0)
