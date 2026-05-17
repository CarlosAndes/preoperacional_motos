from flask import Flask, render_template, request, jsonify
import sqlite3
import os
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.styles import Font
from openpyxl.drawing.image import Image

app = Flask(__name__)

DB = 'database/inspecciones.db'
EXCEL = 'excel/preoperacional.xlsx'

# URL DE TU GOOGLE APPS SCRIPT
GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycbwdM3p5xM1fkHxba7cy-nY-KMq9Grz1W1-gUPyLz3nJ_iZNmUdYgZfDQf2Ud1jL1rgJ/exec"


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

    conexion.commit()
    conexion.close()


@app.route('/')
def inicio():
    return render_template('index.html')


@app.route('/guardar', methods=['POST'])
def guardar():

    datos = request.form
    archivo = request.files.get('foto')

    nombre_foto = 'Sin foto'

    if archivo and archivo.filename:

        nombre_foto = archivo.filename

        ruta = os.path.join(
            'static/uploads',
            nombre_foto
        )

        archivo.save(ruta)

    # GUARDAR EN SQLITE
    conexion = sqlite3.connect(DB)
    cursor = conexion.cursor()

    cursor.execute("""

    INSERT INTO inspecciones(

    fecha,
    encargado,
    documento,
    placa,
    kilometraje,

    moto,
    parte,
    estado,

    observacion,
    foto

    )

    VALUES(
    ?,?,?,?,?,?,?,?,?,?
    )

    """,

    (

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


    # ENVIAR A GOOGLE SHEETS
    try:

        fecha = datos['fecha'].split("T")[0]
        hora = datos['fecha'].split("T")[1][:5]

        requests.post(

            GOOGLE_SCRIPT_URL,

            json={

                "fecha": fecha,
                "hora": hora,

                "encargado": datos['encargado'],
                "documento": datos['documento'],
                "placa": datos['placa'],
                "kilometraje": datos['kilometraje'],

                "moto": datos['moto'],
                "parte": datos['parte'],
                "estado": datos['estado'],

                "observacion": datos['observacion']

            }

        )

        print("Datos enviados a Google Sheets")

    except Exception as e:

        print("Error Google:", e)


    # GENERAR EXCEL
    fecha_excel = datos['fecha'].split("T")[0]

    conexion = sqlite3.connect(DB)

    consulta = f"""

    SELECT

    substr(fecha,1,10) as Fecha,
    substr(fecha,12,5) as Hora,

    encargado,
    documento,
    placa,
    kilometraje,

    moto,
    parte,
    estado,
    observacion,
    foto

    FROM inspecciones

    WHERE fecha LIKE
    '{fecha_excel}%'

    """

    df = pd.read_sql(
        consulta,
        conexion
    )

    conexion.close()


    with pd.ExcelWriter(
        EXCEL,
        engine='openpyxl'
    ) as writer:

        df.to_excel(
            writer,
            sheet_name=fecha_excel,
            index=False
        )


    libro = load_workbook(EXCEL)

    hoja = libro[fecha_excel]


    for celda in hoja[1]:

        celda.font = Font(
            bold=True
        )


    hoja['L1'] = 'EVIDENCIA'


    fila = hoja.max_row


    if nombre_foto != 'Sin foto':

        ruta_imagen = os.path.join(
            'static/uploads',
            nombre_foto
        )

        if os.path.exists(
            ruta_imagen
        ):

            try:

                img = Image(
                    ruta_imagen
                )

                img.width = 80
                img.height = 60

                hoja.add_image(
                    img,
                    f'L{fila}'
                )

                hoja.row_dimensions[
                    fila
                ].height = 50

                hoja.column_dimensions[
                    'L'
                ].width = 20

            except:

                print(
                    'No se pudo insertar imagen'
                )


    libro.save(EXCEL)

    print("\nGUARDADO OK\n")


    return jsonify({

        "mensaje":"guardado"

    })


if __name__ == '__main__':

    crear_bd()

    app.run(debug=True)