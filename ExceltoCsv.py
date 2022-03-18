import pandas as pd

df = pd.read_excel (r'Bajas Tempranas Noviembre 2021 (1).xlsx')

# remove spaces in columns name
df.columns = df.columns.str.replace(' ','_')


col = df[["Categoría","OS","Mod"]]
print (col.head())


header = ["Categoría", "OS" , "Mod"]
df.to_csv('output.csv', columns = header , index=False)
