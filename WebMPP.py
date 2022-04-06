from logging import debug
from os import P_DETACH
from warnings import catch_warnings
from flask import Flask
from flask import render_template,request,redirect
from logging import FileHandler,WARNING
import logging
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

logging.basicConfig(filename='error4.log',level=logging.DEBUG)


app = Flask(__name__)


# This is the path to the upload directory
app.config['UPLOAD_FOLDER'] = 'Archivos/'

app.config['CSV_FOLDER'] = 'CSV/'
# These are the extension that we are accepting to be uploaded
app.config['ALLOWED_EXTENSIONS'] = set(['xls','xlsx','XLS','XLSX'])

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

                    # Elimino los csv                    
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'CartaAnalisisManual.csv')
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'CartaEmpresa.csv')
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'CartaSocioDeudor.csv')
                    os.remove(os.path.join(app.config['CSV_FOLDER'])+'CartaSocioSinDeuda.csv')
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
    
        # uno los archivos y filtro .
        ExportCSV()

        csv_to_zip()
        # print(filenames)
        # return render_template('download.html',filename=filename)
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
            Bajas = {}
            Patagonia = {}
            i = 1

            # recorro la carpeta de los archivos
            # Verifico si es bajas o patagonia
            with os.scandir(os.path.join(app.config['UPLOAD_FOLDER'])) as ficheros:
                for fichero in ficheros: 
                    if 'Bajas' in fichero.name  :                  
                        Bajas[1] = fichero.name
                    else:
                        Patagonia[2] = fichero.name   
                    i += 1



            
            Bajas1 = pd.read_excel(os.path.join(os.path.join(app.config['UPLOAD_FOLDER']), Bajas[1] ))
            # ventana = tkinter.Tk()

            Patagonia2 = pd.read_excel(os.path.join(os.path.join(app.config['UPLOAD_FOLDER']), Patagonia[2] ))

            Bajas1.columns = Bajas1.columns.str.replace(' ','_')
            Patagonia2.columns = Patagonia2.columns.str.replace(' ','_')

            Bajas1.columns = Bajas1.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
            Patagonia2.columns = Patagonia2.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

  
            # Defino las columnas que se van a exportar
            header = ["Categoria", "Mod" ,"CUIL","ICSoc","APELLIDO__NOMBRE","CUIT","Razon","MaxDeFIN_REL_LAB","Situacion_Informada"]

            frames = [Bajas1, Patagonia2]

            result = pd.concat(frames)

            # Filtro por categoria y filial 
            newrCartaSocioDeudor = result.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("deudor",case=False)')

            newrCartaEmpresa = result.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("Empresa",case=False)')

            newrCartaSocioSinDeuda = result.query('Filial in ('"13"', '"9"', '"30"') & ~Categoria.str.contains("deudor",case=False) & Categoria.str.contains("CartaSocio",case=False)')

            newrAnalisisManual = result.query('Filial in ('"13"', '"9"', '"30"') & ~Categoria.str.contains("deudor",case=False) &  ~Categoria.str.contains("CartaSocio",case=False) & ~Categoria.str.contains("Empresa",case=False)')

            # exporto
            newrCartaSocioDeudor.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaSocioDeudor.csv'), columns = header , index=False)
            newrCartaEmpresa.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaEmpresa.csv'), columns = header , index=False)
            newrCartaSocioSinDeuda.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaSocioSinDeuda.csv'), columns = header , index=False)
            newrAnalisisManual.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaAnalisisManual.csv'), columns = header , index=False)

            return render_template('upload.html')
        except Exception as e:
                    print(e) 



if __name__ == '__main__':
    app.run(debug=True) 