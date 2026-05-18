from flask import Flask, render_template, request, jsonify
import sqlite3
import os
import pandas as pd
import requests
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font

app = Flask(__name__)

DB='database/inspecciones.db'
EXCEL='excel/preoperacional.xlsx'

GOOGLE_SCRIPT_URL="https://script.google.com/macros/s/AKfycbwdM3p5xM1fkHxba7cy-nY-KMq9Grz1W1-gUPyLz3nJ_iZNmUdYgZfDQf2Ud1jL1rgJ/exec"


def crear_bd():
    os.makedirs('database', exist_ok=True)
    os.makedirs('excel', exist_ok=True)
    os.makedirs('static/uploads', exist_ok=True)

    conexion = sqlite3.connect(DB)
    cursor = conexion.cursor()

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS inspecciones(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        fecha TEXT,
        encargado TEXT,
        documento TEXT,
        placa TEXT,
        kilometraje TEXT,
        moto TEXT,
        parte TEXT,
        estado TEXT,
        observacion TEXT,
        foto TEXT
    )
    """)

    cursor.execute("""
    CREATE TABLE IF NOT EXISTS vehiculos(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        placa TEXT UNIQUE,
        tipo TEXT,
        soat TEXT,
        tecnomecanica TEXT,
        foto TEXT
    )
    """)

    conexion.commit()
    conexion.close()


@app.route('/')
def inicio():
    return render_template('index.html')


@app.route('/guardar', methods=['POST'])
def guardar():

    datos=request.form
    archivo=request.files.get('foto')

    nombre_foto='Sin foto'

    if archivo and archivo.filename:
        extension=os.path.splitext(archivo.filename)[1]
        nombre_foto=f"{datos['placa']}_{datos['parte']}{extension}"

        ruta=os.path.join('static/uploads',nombre_foto)
        archivo.save(ruta)

    conexion=sqlite3.connect(DB)
    cursor=conexion.cursor()

    cursor.execute("""
    INSERT INTO inspecciones(
    fecha,encargado,documento,placa,
    kilometraje,moto,parte,
    estado,observacion,foto)
    VALUES(?,?,?,?,?,?,?,?,?,?)
    """,(
    datos['fecha'],
    datos['encargado'],
    datos['documento'],
    datos['placa'],
    datos['kilometraje'],
    datos['moto'],
    datos['parte'],
    datos['estado'],
    datos['observacion'],
    nombre_foto
    ))

    conexion.commit()
    conexion.close()

    try:
        fecha=datos['fecha'].split("T")[0]
        hora=datos['fecha'].split("T")[1][:5]

        requests.post(
            GOOGLE_SCRIPT_URL,
            json={
                "fecha":fecha,
                "hora":hora,
                "encargado":datos['encargado'],
                "documento":datos['documento'],
                "placa":datos['placa'],
                "kilometraje":datos['kilometraje'],
                "moto":datos['moto'],
                "parte":datos['parte'],
                "estado":datos['estado'],
                "observacion":datos['observacion'],
                "foto":nombre_foto
            }
        )
    except Exception as e:
        print(e)

    fecha_excel=datos['fecha'].split('T')[0]

    conexion=sqlite3.connect(DB)
    consulta=f"""
    SELECT
    substr(fecha,1,10) as Fecha,
    substr(fecha,12,5) as Hora,
    encargado,documento,placa,
    kilometraje,moto,parte,
    estado,observacion,foto
    FROM inspecciones
    WHERE fecha LIKE '{fecha_excel}%'
    """

    df=pd.read_sql(consulta,conexion)
    conexion.close()

    with pd.ExcelWriter(EXCEL,engine='openpyxl') as writer:
        df.to_excel(writer,sheet_name=fecha_excel,index=False)

    libro=load_workbook(EXCEL)
    hoja=libro[fecha_excel]

    for celda in hoja[1]:
        celda.font=Font(bold=True)

    libro.save(EXCEL)

    return jsonify({"mensaje":"guardado"})


@app.route('/agregar_vehiculo')
def agregar_vehiculo():

    conexion = sqlite3.connect(DB)
    cursor = conexion.cursor()

    vehiculos=[

    ('TLJ39F','Automática','2026-05-30','2026-05-30','TLJ39F.jpg'),

    ('NWJ76F','Cambios','2026-05-30','2026-05-30','NWJ76F.jpg'),

    ('NWJ06F','Cambios','2026-05-30','2026-05-30','NWJ76F.jpg'),

    ('XXK43G','Cambios','2026-05-30','2026-05-30','XXK43G.jpg')

    ]


    cursor.execute(
        "DELETE FROM vehiculos"
    )

    cursor.executemany("""

    INSERT INTO vehiculos(

    placa,
    tipo,
    soat,
    tecnomecanica,
    foto

    )

    VALUES(
    ?,?,?,?,?

    )

    """,vehiculos)

    conexion.commit()
    conexion.close()

    return 'Vehículos cargados'

@app.route('/vehiculo/<placa>')
def obtener_vehiculo(placa):

    conexion=sqlite3.connect(DB)
    conexion.row_factory=sqlite3.Row
    cursor=conexion.cursor()

    cursor.execute(
    'SELECT * FROM vehiculos WHERE placa=?',
    (placa,)
    )

    vehiculo=cursor.fetchone()
    conexion.close()

    if not vehiculo:
        return jsonify({'error':'No encontrado'})

    vehiculo=dict(vehiculo)

    hoy=datetime.now().date()

    soat=vehiculo['soat'] or '2026-05-30'
    tecno=vehiculo['tecnomecanica'] or '2026-05-30'

    fecha_soat=datetime.strptime(
    soat,
    '%Y-%m-%d'
    ).date()

    fecha_tecno=datetime.strptime(
    tecno,
    '%Y-%m-%d'
    ).date()

    vehiculo['dias_soat']=(fecha_soat-hoy).days
    vehiculo['dias_tecno']=(fecha_tecno-hoy).days

    return jsonify(vehiculo)


@app.route('/actualizar_documentos',methods=['POST'])
def actualizar_documentos():

    datos=request.form

    conexion=sqlite3.connect(DB)
    cursor=conexion.cursor()

    cursor.execute("""
    UPDATE vehiculos
    SET soat=?,tecnomecanica=?
    WHERE placa=?
    """,(
    datos['soat'],
    datos['tecnomecanica'],
    datos['placa']
    ))

    conexion.commit()
    conexion.close()

    return jsonify({
    'mensaje':'Documentos actualizados'
    })


if __name__=='__main__':
    crear_bd()
    app.run(debug=True)