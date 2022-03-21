import pandas as pd

df = pd.read_excel (r'Bajas Tempranas Noviembre 2021 (1).xlsx')

# remove spaces in columns name
df.columns = df.columns.str.replace(' ','_')

# str.decode('utf-8').replace(u'\xf1', 'n')

df.columns = df.columns.str.normalize('NFKD').str.encode('ascii', errors='ignore').str.decode('utf-8')



# seleccion de columnas 
# col = df[["Categoría","OS","Mod","Filial"]]
print (df.columns)




# Filtro por categoria y filial 

newdf = df.query('Filial in ('"13"', '"9"', '"30"') & Categoria.str.contains("deudor",case=False)')
print (newdf.columns)

# "Courses in ('Spark','PySpark')"

# Defino las columnas que se van a exportar


# # renombro la columna
newdf = newdf.rename(columns={"MaxDeFIN_REL_LAB":"FECHA"})

header = ["CUIL", "APELLIDO__NOMBRE" , "Razon","FECHA"]

# exporto

newdf.to_csv('CartaSocioDeudor.csv', columns = header , index=False)
# newdf.to_csv('CartaSocioDeudor.csv', index=False)
# 

print(newdf)