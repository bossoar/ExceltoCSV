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


logging.basicConfig(filename='error4.log',level=logging.DEBUG)


app = Flask(__name__)


# This is the path to the upload directory
app.config['UPLOAD_FOLDER'] = 'Archivos/'

app.config['CSV_FOLDER'] = 'CSV/'
# These are the extension that we are accepting to be uploaded
app.config['ALLOWED_EXTENSIONS'] = set(['xls','xlsx','XLS','XLSX'])

# For a given file, return whether it's an allowed type or not
def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1] in app.config['ALLOWED_EXTENSIONS']

# This route will show a form to perform an AJAX request
# jQuery is loaded to execute the request and update the
# value of the operation
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
                # Redirect the user to the uploaded_file route, which
                # will basicaly show on the browser the uploaded file
        # Load an html page with a link to each uploaded file
    
        # uno los archivos y filtro .
        ExportCSV()
        # print(filenames)
        # return render_template('download.html',filename=filename)
        return render_template('download.html')
    except Exception as e:
                    print(e)


# Para descargar los CSV
@app.route('/download')
def download_file():
	path = app.config['CSV_FOLDER']+"CartaSocioDeudor.csv"
    # print(send_from_directory(app.config['CSV_FOLDER'])
	return send_file(path, as_attachment=True)    

    # return send_from_directory(app.config['CSV_FOLDER'],
    #                            filename,as_attachment=True)


# This route is expecting a parameter containing the name
# of a file. Then it will locate that file on the upload
# directory and show it on the browser, so if the user uploads
# an image, that image is going to be show after the upload
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
            with os.scandir(os.path.join(app.config['UPLOAD_FOLDER'])) as ficheros:
                for fichero in ficheros: 
                    if 'Bajas' in fichero.name  :                  
                        Bajas[1] = fichero.name
                    else:
                        Patagonia[2] = fichero.name   
                    i += 1


            # Verifico si es bajas o patagonia
            

            print(Bajas[1])    
            print(Patagonia[2])  
            
            Bajas1 = pd.read_excel(os.path.join(os.path.join(app.config['UPLOAD_FOLDER']), Bajas[1] ))
            # ventana = tkinter.Tk()

            Patagonia2 = pd.read_excel(os.path.join(os.path.join(app.config['UPLOAD_FOLDER']), Patagonia[2] ))

            

            Bajas1.columns = Bajas1.columns.str.replace(' ','_')
            Patagonia2.columns = Patagonia2.columns.str.replace(' ','_')

            

            Bajas1.columns = Bajas1.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
            Patagonia2.columns = Patagonia2.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

            # Filtro por categoria y filial 
            # newBajas1 = Bajas1.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("deudor",case=False)')
            print (Bajas1.columns)
            print (Patagonia2.columns)


            # Filtro por categoria y filial 
            # newPatagonia2 = Patagonia2.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("deudor",case=False)')
            # print (newPatagonia2)


            # renombro la columna
            # newBajas1 = newBajas1.rename(columns={"MaxDeFIN_REL_LAB":"FECHA DE BAJA ABT"})

            # newPatagonia2 = newPatagonia2.rename(columns={"Filial":"Filial"})
            # newPatagonia2 = newPatagonia2.rename(columns={"CUILSOC":"CUIL"})
            # newPatagonia2 = newPatagonia2.rename(columns={"RAZON SOCIAL":"Razon"})

            # print(newPatagonia2) 

            # Defino las columnas que se van a exportar
            header = ["Categoria", "Mod" ,"CUIL","ICSoc","APELLIDO__NOMBRE","CUIT","Razon","MaxDeFIN_REL_LAB",]
            # print (newBajas1.columns)


            # Combino las columnas de nombre y apellido 
            # newPatagonia2["APELLIDO__NOMBRE"] = newPatagonia2["APELLIDO"] +" "+ newPatagonia2["NOMBRE"]
            # print(newPatagonia2)

            # Defino las columnas que se van a exportar
            # header2 = ["CUIL", "APELLIDO__NOMBRE" , "Razon"]

            frames = [Bajas1, Patagonia2]

            result = pd.concat(frames)

            # print(frames)

            # Filtro por categoria y filial 
            newrCartaSocioDeudor = result.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("deudor",case=False)')

            newrCartaEmpresa = result.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("Empresa",case=False)')

            newrCartaSocioSinDeuda = result.query('Filial in ('"13"', '"9"', '"30"') & ~Categoria.str.contains("deudor",case=False) & Categoria.str.contains("CartaSocio",case=False)')

            newrAnalisisManual = result.query('Filial in ('"13"', '"9"', '"30"') & ~Categoria.str.contains("deudor",case=False) &  ~Categoria.str.contains("CartaSocio",case=False) & ~Categoria.str.contains("Empresa",case=False)')
            
            # print(newrCartaEmpresa)
            

            # exporto
            newrCartaSocioDeudor.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaSocioDeudor.csv'), columns = header , index=False)
            newrCartaEmpresa.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaEmpresa.csv'), columns = header , index=False)
            newrCartaSocioSinDeuda.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaSocioSinDeuda.csv'), columns = header , index=False)
            newrAnalisisManual.to_csv(os.path.join(app.config['CSV_FOLDER'],'CartaAnalisisManual .csv'), columns = header , index=False)


            # print(newPatagonia2)    
           

            return render_template('upload.html')
        except Exception as e:
                    print(e) 





if __name__ == '__main__':
    app.run(debug=True) 