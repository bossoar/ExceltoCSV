from logging import debug
from os import P_DETACH
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


logging.basicConfig(filename='error4.log',level=logging.DEBUG)


app = Flask(__name__)

# Carpeta de subida
app.config['UPLOAD_FOLDER'] = './Archivos'



@app.route("/")
def upload_file():
     # renderiamos la plantilla "formulario.html"
 return render_template('form.html')

@app.route("/upload", methods=['POST'])
def uploader():
 if request.method == 'POST':
  # obtenemos el archivo del input "archivo"
  f = request.files['archivo']
  filename = secure_filename(f.filename)
  # Guardamos el archivo en el directorio "Archivos PDF"
  f.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
  # Retornamos una respuesta satisfactoria
  return "<h1>Archivo subido exitosamente</h1>"




@app.route('/')
def index():
                
            df = pd.read_excel (r'Bajas Tempranas Noviembre 2021.xlsx')
            # ventana = tkinter.Tk()

            df2 = pd.read_excel (r'PATAGONIA NORTE.XLS')

            # etiqueta = tkinter.Label(ventana, text = "Hola Mundo" , bg = "blue")

            # ventana.mainloop()  

            # remove spaces in columns name
            df.columns = df.columns.str.replace(' ','_')

            # Creo dataframe vacio por error 
            # df2 = pd.DataFrame(pd.np.empty((0, 29)))   

            # print(df2)


            # df2.columns = df.columns.str.replace(' ','_')


            df.columns = df.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')
            df2.columns = df2.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')

            # Filtro por categoria y filial 
            newdf = df.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("deudor",case=False)')
            # print (newdf.columns)


            # Filtro por categoria y filial 
            newdf2 = df2.query('FILGES in ('"13"', '"9"', '"30"') & Categoria.str.contains("deudor",case=False)')
            # print (newdf2.columns)


            # renombro la columna
            newdf = newdf.rename(columns={"MaxDeFIN_REL_LAB":"FECHA DE BAJA ABT"})

            newdf2 = newdf2.rename(columns={"FILGES":"Filial"})
            newdf2 = newdf2.rename(columns={"CUILSOC":"CUIL"})
            newdf2 = newdf2.rename(columns={"RAZON SOCIAL":"Razon"})

            # print(newdf2) 

            # Defino las columnas que se van a exportar
            header = ["CUIL", "APELLIDO__NOMBRE" , "Razon","FECHA DE BAJA ABT"]



            # Combino las columnas de nombre y apellido 
            newdf2["APELLIDO__NOMBRE"] = newdf2["APELLIDO"] +" "+ newdf2["NOMBRE"]

            # Defino las columnas que se van a exportar
            header2 = ["CUIL", "APELLIDO__NOMBRE" , "Razon"]

            frames = [newdf, newdf2]

            result = pd.concat(frames)

            print(result)
            

            # exporto
            result.to_csv('CartaSocioDeudor.csv', columns = header , index=False)


            print(newdf2)    
            

            return render_template('AutomatMPP.html')






if __name__ == '__main__':
    app.run(debug=True)            