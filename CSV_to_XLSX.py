#!/usr/bin/python
# -*- coding: utf-8 -*-
import os
import pandas as pd
import configparser

# Leer las rutas configuradas en config.ini
config = configparser.ConfigParser()
config.read('config.ini')

# Obtener las rutas desde el archivo de configuración
path_csv = config['rutas']['path_csv']
path_xlsx = config['rutas']['path_xlsx']

contenido = os.listdir(path_csv)
for fichero in contenido:
    try:
        if os.path.isfile(os.path.join(path_csv, fichero)) and fichero.endswith('.csv'):
            df = pd.read_csv(f'{path_csv}/{fichero}', header=None, encoding='utf-8')

            # Filtra las filas con información relevante
            df = df[df[0].str.startswith(('CATEGORIES', 'DTSTART', 'SUMMARY', 'DESCRIPTION', 'TRIGGER'))]

            # Pivota el dataframe para todas las columnas
            df_pivot = df.transpose()

            # Guarda el dataframe modificado en un archivo XLSX
            df_pivot.to_excel(os.path.join(path_xlsx, f"{fichero.replace('.csv', '.xlsx')}"), index=False, header=None)

        else:
            print(f"El archivo {fichero} no es un archivo CSV o no tiene la extensión .csv")

    except Exception as e:
        print(f"Se produjo un error al procesar el archivo {fichero}: {str(e)}")

print("Conversion CSV to XLSX finished!")
