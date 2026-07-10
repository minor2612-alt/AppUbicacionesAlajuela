import pandas as pd
import sqlite3

EXCEL = "data/Ubicaciones Alajuela Glide.xlsx"
DB = "ubicaciones.db"

df = pd.read_excel(EXCEL)

conexion = sqlite3.connect(DB)

df.to_sql("productos", conexion, if_exists="replace", index=False)

conexion.close()

print("Base de datos creada correctamente.") 
