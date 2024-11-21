import os
import math
import pandas as pd
import numpy as np
from openpyxl import load_workbook

filePath = '/content/PIAM_UNICAUCA3.xlsx'
outputPathXlsx = '/content/IntegradorPIAM.xlsx'

mapeo = {
    "Valor Factura": "MTR NETA"}

def cargarArchivosDataframes(filePath):
    if not os.path.isfile(filePath):
        raise FileNotFoundError(f"{filePath} no encontrado.")
    print(f"Archivo {filePath} encontrado.")
    try:
        dicInsumos = pd.read_excel(filePath, sheet_name=None, engine='openpyxl')
        for nombre, df in dicInsumos.items():
            df.columns = df.columns.str.strip()
        return dicInsumos
    except Exception as e:
        raise Exception(f"Error al cargar los DataFrames: {e}")

def integradorPiam2020(df):
    PIAM2020_1_GOB = df['PIAM2020_1_GOB']
    PIAM2020_1_GOB_R21 = df['PIAM2020_1_GOB_R21']
    SQ20202024 = df['SQ20202024']

    piam2020Gob = pd.merge(PIAM2020_1_GOB, SQ20202024, left_on='BOLETA', right_on='Documento', how='left')
    print(piam2020Gob.columns)
    piam2020GobFinal = piam2020Gob[['CUENTA', 'BOLETA', 'Id  factura', 'IDENTIFICACION', 'TERCERO',
                                    'PROGRAMA', 'Estado Actual', 'MTR NETA', 'APORTES GOBERNCION',
                                    'Periodico Academico', 'Tipo de Financiacion']]

    piam2020GobR = pd.merge(PIAM2020_1_GOB_R21, SQ20202024, left_on='BOLETA', right_on='Documento', how='left')
    piam2020GobRFinal = piam2020GobR[['CUENTA', 'BOLETA', 'Id  factura', 'IDENTIFICACION', 'TERCERO',
                                      'PROGRAMA', 'Estado Actual', 'Valor Factura', 'APORTES GOBERNCION',
                                      'Periodico Academico', 'Tipo de Financiacion']]

    piam2020GobRFinal = piam2020GobRFinal.rename(columns=mapeo)
    piam2020GobF = pd.concat([piam2020GobFinal, piam2020GobRFinal], ignore_index=True)

    return piam2020GobF




## ETL INTEGRADOR PIAM
# EXTRACCION
dicInsumos = cargarArchivosDataframes(filePath)

# MANIPULACION
piam2020GobF = integradorPiam2020(dicInsumos)

#CARGA
with pd.ExcelWriter(outputPathXlsx, engine='xlsxwriter') as writer:
    if piam2020GobF is not None:
      piam2020GobF.to_excel(writer, sheet_name='piam2020Gob', index=False)
print("Los resultados han sido guardados en el documento y archivo Excel.")
