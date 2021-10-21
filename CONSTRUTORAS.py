import numpy as np
import pandas as pd
import os

depara = pd.read_excel("deparaTODOS.xlsx", header=6, usecols='A,F').fillna('')
sienge = pd.read_excel("siengeTODOS.xlsx", header=5, usecols='A,J')


try:
    delete = pd.read_excel("siengeTODOS", usecols="H", header=None, sheet_name='ROTTA ELY')
    delete = delete.iloc[0,0]
except:
    delete = 0
if delete == 1:
    sienge = pd.read_excel("siengeTODOS.xlsx", usecols="A,J", sheet_name='ROTTA ELY')
else:
    sienge = pd.read_excel("siengeTODOS.xlsx", usecols="A,I", sheet_name='ROTTA ELY')


# TREATMENT DEPARA

depara.columns = ['Cta. Sienge', 'Cta. CG']

# TREATMENT SIENGE

sienge.columns = ['Fornecedores', 'Valores']

sienge['Fornecedor'] = np.nan
sienge['Saldo Sienge'] = np.nan
sienge['Fornecedor'] = np.where(sienge['Fornecedores'] == 'Fornecedor', sienge['Fornecedores'].shift(-1), sienge['Fornecedor'])
sienge['Saldo Sienge'] = np.where(sienge['Fornecedores'] == 'Total do Fornecedor', sienge['Valores'], sienge['Saldo Sienge'])
sienge['Saldo Sienge'] = sienge['Saldo Sienge'].bfill()

sienge = sienge.fillna('')
mascara = sienge['Fornecedor'] != ''
sienge = sienge[mascara]
sienge = sienge.drop(columns=['Fornecedores', 'Valores'])

sienge['Cta. Sienge'] = sienge['Fornecedor'].str.split('-').str[0]
sienge['Cta. Sienge'] = pd.to_numeric(sienge['Cta. Sienge'])

# MERGE FROM DEPARA IN SIENGE

sienge = pd.merge(sienge, depara, on=['Cta. Sienge'], how='left')

# USING MERGE AS A DIFFERENCE WORKSHEET

divergencias = sienge.copy().fillna('')
del divergencias['Saldo Sienge']
divergencias['Fornecedor'] = divergencias['Fornecedor'].str.split('-').str[1]
mascara = divergencias['Cta. CG'] == ''
divergencias = divergencias[mascara]







writer = pd.ExcelWriter('CONSTRUTORAS - FORNECEDORES.xlsx', engine='xlsxwriter')
sienge.to_excel(writer, sheet_name="sienge", index=False, header=True, startrow=0)
divergencias.to_excel(writer, sheet_name="Divergências", index=False, header=True, startrow=0)

workbook = writer.book
worksheet = writer.sheets['Divergências']

format1 = workbook.add_format({'bg_color': 'white', 'pattern': 1, 'valign': 'center'})
format2 = workbook.add_format({'bg_color': 'white', 'pattern': 1, 'valign': 'center', 'num_format': '[$R$] #,##0.00'})

worksheet.set_column('A:A', 88, format1)
worksheet.set_column('B:C', 12, format1)
worksheet.set_column('D:XFD', None, None, {'hidden': True})




writer.save()
os.startfile('CONSTRUTORAS - FORNECEDORES.xlsx')




´LPDAÁDL´PDA´LPDA
KPODAKPOADDA

PKODAKOPDAOKDA


DAOPKADKOPODAAD


KOPDAKOPDAODA