#!/usr/bin/python
# -*- coding: utf-8 -*-
import icalendar
import csv
import os
import configparser

# Leer las rutas configuradas en config.ini
config = configparser.ConfigParser()
config.read('config.ini')

def extraer_categories(component):
    categories = component.get("categories")
    if categories:
        # Convertirlo en lista de cadenas
        return [str(cat) for cat in categories.cats]
    else:
        # o en cadena vacia si no hay categorías
        return []


def extraer_dtstart(component):
    dtstart = component.decoded("dtstart")
    # Convertir deltatime en cadena
    if dtstart:
        return str(dtstart)
    else:
        return []


def extraer_summary(component):
    summary = component.decoded("summary")
    if summary:
        return summary
    else:
        return []


def extraer_description(component):
    description = component.get("description")
    # Convertir descripcion en cadena, suponiendo utf8 encoding
    if description:
        return description
    else:
        return str(description)


def extraer_trigger(component):
    # Convertir el deltatime en cadena
    trigger = component.get("trigger")
    if trigger:
        return str(trigger.dt)
    else:
        return []


mapa_extractores_vevent = {
    "CATEGORIES": extraer_categories,
    "DTSTART": extraer_dtstart,
    "SUMMARY": extraer_summary,
    "DESCRIPTION": extraer_description,
    "TRIGGER": extraer_trigger,
}
mapa_extractores_alarm = {
    "DESCRIPTION": extraer_description,
    "TRIGGER": extraer_trigger,
}


def extraer_campos(component):
    dic = {}
    if component.name == "VEVENT":
        mapa = mapa_extractores_vevent
    elif component.name == "VALARM":
        mapa = mapa_extractores_alarm
    else:
        return None
    for campo, funcion in mapa.items():
        dic[campo] = funcion(component)
    return dic


datos = {
         "CATEGORIES": [],
         "DTSTART": [],
         "SUMMARY": [],
         "DESCRIPTION": [],
         "TRIGGER": [],
         }

# Obtener las rutas desde el archivo de configuración
path_ics = config['rutas']['path_ics']
path_csv = config['rutas']['path_csv']

contenido = os.listdir(path_ics)
for fichero in contenido:
    try:
        if os.path.isfile(os.path.join(path_ics, fichero)) and fichero.endswith('.ics'):
            e = open(f'{path_ics}/{fichero}', 'r', encoding='utf-8',
                     errors='strict')
            ecal = icalendar.Calendar.from_ical(e.read())
            for component in ecal.walk():
                dic = extraer_campos(component)
                if not dic:  # Nos saltamos los componentes no reconocidos
                    continue
                for clave, valor in dic.items():
                    datos[clave].append(valor)
            fichero = fichero.replace('.ics', '.csv')

            fichero_csv = open(f'{path_csv}/{fichero}', 'w')
            final = csv.writer(fichero_csv)
            final.writerows(datos.items())
            fichero_csv.close()
            e.close()

        else:
            print(f"El archivo {fichero} no es un archivo ICS o no tiene la extensión .ics")

    except Exception as e:
        print(f"Se produjo un error al procesar el archivo {fichero}: {str(e)}")

print("Conversion ICS to CSV finished!")
