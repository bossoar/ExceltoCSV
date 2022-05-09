from logging import debug
from os import P_DETACH
from warnings import catch_warnings
from flask import Flask
from flask import render_template,request,redirect
from logging import FileHandler,WARNING
import logging
from numpy import int64
import pandas as pd
import pandas as pd
import time
import logging
from werkzeug.utils import secure_filename
from werkzeug.datastructures import  FileStorage
import os
from flask import Flask, render_template, request, redirect, url_for, send_from_directory,send_file,current_app,Response
# from werkzeug import secure_filename
from werkzeug.utils import secure_filename
from werkzeug.datastructures import  FileStorage
import zipfile
import shutil
import glob
import pyodbc
import numpy as np


logging.basicConfig(filename='error4.log',level=logging.DEBUG)


app = Flask(__name__)


# Conectar base de datos
# conexion = pyodbc.connect('DRIVER={SQL Server Native Client 11.0};'
#                             'SERVER=SRV13-VMSQL\\NORPATAGONICA;'
#                             'DATABASE=prueba;'
#                            ' UID=UsuarioOSDE;'
#                             'PWD=EDSOoirausU') 


# This is the path to the upload directory
app.config['UPLOAD_FOLDER'] = 'Archivos/'

app.config['CSV_FOLDER'] = 'CSV/'
# These are the extension that we are accepting to be uploaded
app.config['ALLOWED_EXTENSIONS'] = set(['xls','xlsx','XLS','XLSX'])

def remove_file():
        try:
            py_files = glob.glob(app.config['UPLOAD_FOLDER']+'*.xls')

            for py_file in py_files:
                try:
                    os.remove(py_file)
                except OSError as e:
                    print(f"Error:{ e.strerror}")
        except Exception as e:
                        print(e)



# Creo el zip con los csv
def csv_to_zip():

    try:

        print(os.path.join(app.config['CSV_FOLDER']))
        print(os.listdir(os.path.join(app.config['CSV_FOLDER'])))
        
        list_files = os.listdir(os.path.join(app.config['CSV_FOLDER']))
        with zipfile.ZipFile('csv.zip','w') as zipF:

            # Add multiple files to the zip
                    zipF.write(os.path.join(app.config['CSV_FOLDER'])+'CartaAnalisisManual.csv')
                    zipF.write(os.path.join(app.config['CSV_FOLDER'])+'CartaEmpresa.csv')
                    zipF.write(os.path.join(app.config['CSV_FOLDER'])+'CartaSocioDeudor.csv')
                    zipF.write(os.path.join(app.config['CSV_FOLDER'])+'CartaSocioSinDeuda.csv')
                    zipF.write(os.path.join(app.config['CSV_FOLDER'])+'Afiliaciones.csv')

                    # Elimino los csv                    
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'CartaAnalisisManual.csv')
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'CartaEmpresa.csv')
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'CartaSocioDeudor.csv')
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'CartaSocioSinDeuda.csv')
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'Afiliaciones.csv')



    except Exception as e:
                    print(e) 



# For a given file, return whether it's an allowed type or not
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']

@app.route('/')
def index():
    return render_template('index.html')


# Route that will process the file upload
@app.route('/upload', methods=['POST'])
def upload():
    # Get the name of the uploaded files
    try:
        uploaded_files = request.files.getlist("file[]")
        # print(uploaded_files)
        filenames = []
        cantidadArchivos = 0
        for file in uploaded_files:
            # Check if the file is one of the allowed types/extensions
            if file and allowed_file(file.filename):
                # Make the filename safe, remove unsupported chars
                filename = secure_filename(file.filename)
                # Move the file form the temporal folder to the upload
                # folder we setup
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
                # Save the filename into a list, we'll use it later
                filenames.append(filename)
                cantidadArchivos = cantidadArchivos + 1
        # uno los archivos y filtro .


        print(cantidadArchivos)

        ExportCSV()

        # genero el zip
        csv_to_zip()

        # Elimino los archivos importados    
        remove_file()

        return render_template('download.html')
    except Exception as e:
                    print(e)


# Para descargar los CSV
@app.route('/download')
def download_file():
	path = "csv.zip"
    # print(send_from_directory(app.config['CSV_FOLDER'])
	return send_file(path, as_attachment=True)    

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['CSV_FOLDER'],
                               filename)

        

# @app.route('/ExportCSV/<filenames>')
def ExportCSV():
        try:
            ABT = {}
            MPP = {}
            i = 1

            # recorro la carpeta de los archivos
            # Verifico si es ABT o MPP
            with os.scandir(os.path.join(app.config['UPLOAD_FOLDER'])) as ficheros:
                for fichero in ficheros: 
                    if 'ABT' in fichero.name:                  
                        ABT[1] = fichero.name
                    else:
                        MPP[2] = fichero.name   
                    i += 1

            if not ABT[1]:

               print('ABT[1] no tiene datos')

            else:

               print('ABT[1] tiene datos')            

            
            ABT1 = pd.read_excel(os.path.join(os.path.join(app.config['UPLOAD_FOLDER']), ABT[1] ))
            # ventana = tkinter.Tk()

            MPP2 = pd.read_excel(os.path.join(os.path.join(app.config['UPLOAD_FOLDER']), MPP[2] ))

            ABT1.columns = ABT1.columns.str.replace(' ','_')
            MPP2.columns = MPP2.columns.str.replace(' ','_')

            ABT1.columns = ABT1.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
            MPP2.columns = MPP2.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

            # Combino las columnas de nombre y apellido 
            MPP2["APELLIDO__NOMBRE"] = MPP2["APELLIDO"] +" "+ MPP2["NOMBRE"]

            # Renombro columnas
            MPP2 = MPP2.rename(columns={"FILGES":"Filial"})    
            MPP2 = MPP2.rename(columns={"CUILSOC":"CUIL"})
            MPP2 = MPP2.rename(columns={"RAZON_SOCIAL":"Razon"})
            MPP2 = MPP2.rename(columns={"MOD":"Mod"})
            MPP2 = MPP2.rename(columns={"IC_S":"ICSoc"})
            MPP2 = MPP2.rename(columns={"CUITEMP":"CUIT"})
            ABT1 = ABT1.rename(columns={"Situacion_Informada":"DIAGNOSTICO"})
            ABT1 = ABT1.rename(columns={"IOSoc":"CONTR"})
            MPP2 = MPP2.rename(columns={"MAILSOCIO":"MAILSOC"})
            ABT1 = ABT1.rename(columns={"CuitAT":"Cambio_Empleador"})


           

              
            # Defino las columnas que se van a exportar
            header = ["Categoria", "Mod","DIAGNOSTICO","CANT_EST","PCO","CUIL","CONTR","APELLIDO__NOMBRE","CUIT","Razon","MaxDeFIN_REL_LAB","MAILSOC","Cambio_Empleador"]

            frames = [ABT1, MPP2]

            result = pd.concat(frames)

            try:

                # saco los Nan
                # cambio tipo de columna a int por el problema de 1.0
                result["CANT_EST"] = result["CANT_EST"].fillna(0).apply(np.int64) # Convert nan/null to 0
                result["Cambio_Empleador"] = result["Cambio_Empleador"].fillna(0).apply(np.int64) # Convert nan/null to 0
                
               

            except Exception as e:
                        print(e)

            # print(result)

            # Filtro por categoria y filial 
            newrCartaSocioDeudor = result.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("deudor",case=False)')

            newrCartaEmpresa = result.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("Empresa",case=False)')

            newrCartaSocioSinDeuda = result.query('Filial in ('"13"', '"9"', '"30"') & ~Categoria.str.contains("deudor",case=False) & Categoria.str.contains("CartaSocio|Carta Socio",case=False)')

            newAfiliaciones = result.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("Cambio Empleador|Sumatoria|Pluriempleo",case=False) ')
              
            newrAnalisisManual = result.query('Filial in ('"13"', '"9"', '"30"') & ~Categoria.str.contains("deudor|CartaSocio|Carta Socio|Empresa|Empleador|Sumatoria|Pluriempleo",case=False) ')

            # exporto
            newrCartaSocioDeudor.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaSocioDeudor.csv'), columns = header , index=False)
            newrCartaEmpresa.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaEmpresa.csv'), columns = header , index=False)
            newrCartaSocioSinDeuda.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaSocioSinDeuda.csv'), columns = header , index=False)
            newAfiliaciones.to_csv(os.path.join(app.config['CSV_FOLDER'],'Afiliaciones.csv'), columns = header , index=False)
            newrAnalisisManual.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaAnalisisManual.csv'), columns = header , index=False)

            

            return render_template('upload.html')
        except Exception as e:
                    print(e) 



if __name__ == '__main__':
    app.run(debug=True) 