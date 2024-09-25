import pandas as pd
import numpy as np
import os
import glob
import zipfile

os.chdir('C:/Users/Mayra Herrera/OneDrive - Laboratorios Vanquish SA de CV/Documentos/T551')
extension = 'txt'
archivos = [i for i in glob.glob('*.{}'.format(extension))]
print("Archivos ZIP encontrados:", archivos)

nombresColumnas = ["PATENTE ADUANAL","INDICE","CLAVE DE SECCION ADUANERA DE DESPACHO","FRACCION ARANCELARIA",
                   "SECUENCIA DE LA FRACCION ARANCELARIA","SUBDIVISION DE LA FRACCION","DESCRIPCION DE LA MERCANIA",
                   "PRECIO UNITARIO (MXP)","VALOR ADUANA (MXP)","VALOR COMERCIAL (MXP)","VALOR COMERCIAL (DOLARES)",
                   "Cantidades","CLAVE DE UNIDAD DE MEDIDA COMERCIAL","CANTIDAD MERCANIAS EN UNIDADES DE MEDIDA DE LA TARIFA",
                   "CLAVE DE UNIDAD DE MEDIDA DE LA TARIFA","VALOR AGREGADO","CLAVE DE VINCULACION","CLAVE DE METODO DE VALORIZACION",
                   "CÓDIGO DE LA MERCANCIA O PRODUCTO","MARCA","MODELO","CLAVE DE PAIS ORIGEN","CLAVE DE PAIS COMPRADOR / VENDEDOR",
                   "PRODUCTO","EMPRESA","CONCENTRACION","PRESENTACION","FECHA IMPORTACION"]

df = pd.concat([pd.read_csv(f,header=None,delimiter='|') for f in archivos ], ignore_index=True)
df.shape


