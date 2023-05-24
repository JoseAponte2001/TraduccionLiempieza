import re
import pandas as pd
from googletrans import Translator
global translator
translator = Translator()

## FROMATO Y TRADUCCIÓN
def formato(frame):
    # Recorrer por filas
    for indice, fila in frame.iterrows():
        titulo = frame.iat[indice, 1]
        contenido = frame.iat[indice, 2]
        titulo_l = limpieza(titulo)
        contenido_l = limpieza(contenido)
        frame.iat[indice, 1] = titulo_l
        frame.iat[indice, 2] = contenido_l
    return frame

def limpieza(texto):
    try:
        texto_sin_iconos = re.sub(r'[^\w\s]', '', str(texto))
        texto_sin_saltos = texto_sin_iconos.replace('\n', '. ')
        return texto_sin_saltos
    except Exception as e:
        print(f"Se produjo una excepción: {e}")

def traduccion(frame):
    # Recorrer por filas
    for indice, fila in frame.iterrows():
        titulo = frame.iat[indice, 1]
        contenido = frame.iat[indice, 2]
        titulo_t = translate(titulo)
        contenido_t = translate(contenido)
        frame.iat[indice, 1] = titulo_t
        frame.iat[indice, 2] = contenido_t
    return frame

def translate(texto):
    global translator
    try:
        translation = translator.translate(texto, dest='en')
        return translation.text
    except Exception as e:
        print(f"Se produjo una excepción en traducción: {e}")


# Leer el archivo Excel
data_frame_el_espectador = pd.read_excel('el_espectador.xlsx')
data_frame_el_tiempo = pd.read_excel('eltiempo.xlsx')
data_frame_semana = pd.read_excel('semana.xlsx')

# fomato
frame_el_espectador = formato(data_frame_el_espectador)
frame_el_tiempo = formato(data_frame_el_tiempo)
frame_semana = formato(data_frame_semana)

#traducción
frame_el_espectador_traducido = traduccion(frame_el_espectador)
frame_el_tiempo_traducido = traduccion(frame_el_tiempo)
frame_semana_traducido = traduccion(frame_semana)

#generación
df_el_espectador = pd.DataFrame(frame_el_espectador_traducido)
df_semana = pd.DataFrame(frame_semana_traducido)
df_el_tiempo = pd.DataFrame(frame_el_tiempo_traducido)

# Escribir el DataFrame en un archivo de Excel
df_el_espectador.to_excel('Output/el_espectador.xlsx', index=False)
df_semana.to_excel('Output/eltiempo.xlsx', index=False)
df_el_tiempo.to_excel('Output/semana.xlsx', index=False)