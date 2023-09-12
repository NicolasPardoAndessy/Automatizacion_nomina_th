import pandas as pd
import xlwings as xw
from datetime import datetime
import os
import shutil
import warnings

input('''
        Bienvenid@ a la automatizacion de resumen de nomina actual.
      
        tenga en cuenta tener los archivos "Reporte Novedades Nómina.xlsx" y "Reporte resumen siigo.xlsx" cargados en la carpeta
      
        presione enter para continuar''')

date = datetime.now()
date_month= date.strftime('%m')
month=int(date_month)-1

# Rutas de los archivos de Excel
archivo_origen = 'Reporte Novedades Nómina.xlsx'
archivo_destino = 'Plantilla nomina mensual.xlsx'
reporte_resumen = 'Reporte resumen siigo.xlsx'

# Leer el archivo Excel en un DataFrame
warnings.simplefilter("ignore")
df1 = pd.read_excel(reporte_resumen, header=5)
df2 = pd.read_excel(archivo_destino, header=5)
df3 = pd.read_excel(archivo_origen, header=4)

df1 = df1.sort_values(by='Nombre', ascending=True, ignore_index=True)
df2 = df2.sort_values(by='Nombre', ascending=True, ignore_index=True)
columnasdf1=['Nombre','Identificación','Sueldo','Auxilio de transporte','Horas extras diurnas 125%','Hora extra diurna dominical o festiva','Vacaciones disfrutadas','Bonificación Por Asignación','Total Ingresos','Fondo de salud','Fondo de pensión','Fondo de solidaridad pensional','Retefuente','Fondo Ahorro Mutuo Protección','Plan Premium Sanitas (Beneficiario)','Total deducciones','Neto a Pagar']
form_num = ['Identificación','Sueldo','Auxilio de transporte','Horas extras diurnas 125%','Hora extra diurna dominical o festiva','Vacaciones disfrutadas','Bonificación Por Asignación','Total Ingresos','Fondo de salud','Fondo de pensión','Fondo de solidaridad pensional','Retefuente','Fondo Ahorro Mutuo Protección','Plan Premium Sanitas (Beneficiario)','Total deducciones','Neto a Pagar']

df1 =df1[columnasdf1]
for columna in form_num:
    df1[columna] = pd.to_numeric(df1[columna], errors='coerce')



df1.insert(2,'Area', df2['Area'])
df1.insert(5,'# Horas extras diurnas 125%', 0)
df1.insert(7,'# Hora extra diurna dominical o festiva', 0)
df1.insert(9,'# días Vacaciones disfrutadas', 0)
df1.insert(13,'Total ingresos2', '')
df1.insert(21,'Total deducciones2', '')
df1.insert(22,'Neto a pagar2', '')
df1.fillna(0, inplace=True)

#df1['Total ingresos2']=df1['Sueldo'] + df1['Auxilio de transporte'] +  df1['Horas extras diurnas 125%'] + df1['Hora extra diurna dominical o festiva'] + df1['Vacaciones disfrutadas'] + df1['Bonificación Por Asignación']


#df1['Total deducciones2']=df1['Fondo de salud'] + df1['Fondo de pensión'] +  df1['Fondo de solidaridad pensional']+ df1['Retefuente'] + df1['Fondo Ahorro Mutuo Protección'] + df1['Plan Premium Sanitas (Beneficiario)'] 

#df1['Neto a pagar2']= df1['Total ingresos2'] - df1['Total deducciones2']

#print(df1['Total ingresos2'])
#print(df1)

columnasdf3=['COLABORADOR','Identificación del empleado','¿Qué novedad le vas a cargar?','Asigna la cantidad de Días/Horas o el valor de la novedad']
df3 =df3[columnasdf3]
#print(df3['Identificación del empleado'])
pivote = df3.pivot_table(index=['COLABORADOR','Identificación del empleado'], columns='¿Qué novedad le vas a cargar?', values='Asigna la cantidad de Días/Horas o el valor de la novedad',aggfunc='sum',fill_value=0)
pivote.to_excel("novedades1.xlsx")

novedades = pd.read_excel("novedades1.xlsx",engine="openpyxl")
#print(novedades['Identificación del empleado'])

for i in range(len(df1)):
    for j in range(len(novedades)):
        if (df1.iloc[i]["Identificación"]==novedades.iloc[j]["Identificación del empleado"]):
            df1["# Horas extras diurnas 125%"][i]=novedades.iloc[j]["10- Horas extras diurnas 125%- Ingreso"]
            df1["# Hora extra diurna dominical o festiva"][i]=novedades.iloc[j]["07- Hora extra diurna dominical o festiva- Ingreso"]
            df1["# días Vacaciones disfrutadas"][i]=novedades.iloc[j]["31- Vacaciones disfrutadas- Ingreso"]
            #print(df1["# Horas extras diurnas 125%"][i])


#print(df1["# Horas extras diurnas 125%"],df1["# Hora extra diurna dominical o festiva"],df1["# días Vacaciones disfrutadas"])

df1.to_excel("Reporte nomina temporal.xlsx", index=False)

if not os.path.exists('resumen de nomina mes '+str(month)):
    os.makedirs('resumen de nomina mes '+str(month))

shutil.move("novedades1.xlsx",'resumen de nomina mes ' +str(month)+'/')

archivo_original = 'Plantilla nomina mensual.xlsx'
nuevo_nombre ='Nomina mensual mes '+str(month)+'.xlsx'
shutil.copy2('Plantilla nomina mensual.xlsx', os.path.join('resumen de nomina mes ' +str(month)+'/', nuevo_nombre))

archivo_origen1="Reporte nomina temporal.xlsx"
archivo_destino1 = 'resumen de nomina mes ' +str(month)+'/'+nuevo_nombre


libro_origen=xw.Book(archivo_origen1)
archivo_destino1= xw.Book(archivo_destino1)

hoja_origen=libro_origen.sheets.active
hoja_destino=archivo_destino1.sheets.active

fila_origen = 2
columna_origen = 1
fila_destino = 7
columna_destino = 1
columnas_a_excluir=['Total ingresos2','Total deducciones2','Neto a pagar2']

for i in range(hoja_origen.api.UsedRange.Rows.Count):
    for j in range(hoja_origen.api.UsedRange.Columns.Count):
        valor = hoja_origen.range(fila_origen + i, columna_origen + j).value
        columna_actual = hoja_origen.range(1, j + 1).value
                #print(columna_actual)
        if columna_actual not in columnas_a_excluir:
            hoja_destino.range(fila_destino + i, columna_destino + j).value = valor
archivo_destino1.save()
archivo_destino1.close()
libro_origen.close()
os.remove(archivo_origen1)

areas = ['Admon','TI','Andessy', 'Efpac','Stream','TH']
archivo_original = 'Plantilla nomina mensual area.xlsx'

for i in areas:
    origen = df1[df1['Area'] == i]
    origen.to_excel("Area"+i+".xlsx", index=False)
    #df1.to_excel("Reporte nomina temporal.xlsx", index=False)

    nuevo_nombre ='Resumen nomina mes '+str(month)+ ' '+ i +'.xlsx'
    #print (nuevo_nombre)

    shutil.copy2('Plantilla nomina mensual area.xlsx', os.path.join('resumen de nomina mes ' +str(month)+'/', nuevo_nombre))

    archivo_origen1="Area"+i+".xlsx"
    archivo_destino1 = 'resumen de nomina mes ' +str(month)+'/'+nuevo_nombre

    libro_origen=xw.Book(archivo_origen1)
    archivo_destino1= xw.Book(archivo_destino1)

    hoja_origen=libro_origen.sheets.active
    hoja_destino=archivo_destino1.sheets.active

    fila_origen = 2
    columna_origen = 1
    fila_destino = 7
    columna_destino = 1
    columnas_a_excluir=['Total ingresos2','Total deducciones2','Neto a pagar2']
    for i in range(hoja_origen.api.UsedRange.Rows.Count):
        for j in range(hoja_origen.api.UsedRange.Columns.Count):
            valor = hoja_origen.range(fila_origen + i, columna_origen + j).value
            columna_actual = hoja_origen.range(1, j + 1).value
                #print(columna_actual)
            if columna_actual not in columnas_a_excluir:
                hoja_destino.range(fila_destino + i, columna_destino + j).value = valor
    archivo_destino1.save()
    archivo_destino1.close()
    libro_origen.close()
    os.remove(archivo_origen1)


''' for filtro in areas:
        columna_filtro = filtro['C']
        criterio_filtro = filtro[i]
        
        hoja_origen.range(f'{columna_filtro}1').autofilter()
        hoja_origen.range(f'{columna_filtro}1').current_region.auto_filter(1, criterio_filtro)

        hoja_destino=archivo_destino1.sheets.active

        fila_origen = 2
        columna_origen = 1
        fila_destino = 7
        columna_destino = 1
        columnas_a_excluir=['Total ingresos2','Total deducciones2','Neto a pagar2']
        for i in range(hoja_origen.api.UsedRange.Rows.Count):
            for j in range(hoja_origen.api.UsedRange.Columns.Count):
                valor = hoja_origen.range(fila_origen + i, columna_origen + j).value
                columna_actual = hoja_origen.range(1, j + 1).value
                #print(columna_actual)
                if columna_actual not in columnas_a_excluir:
                    hoja_destino.range(fila_destino + i, columna_destino + j).value = valor
        archivo_destino1.save()'''

print('finalizado')
print('''
                                    |
                                    |
                                    |
                                  .-'-.
                                 ' ___ '
                       ---------'  .-.  '---------
       _________________________'  '-'  '_________________________
        -------|---|--/    \==][^',_m_,'^][==/    \--|---|-------
                      \    /  ||/   H   \||  \    /
                       '--'   OO   O|O   OO   '--'
                      AUTOMATIZACION DE NOMINA ADC
                              EQUIPO DE TH
                        DESIGNED BY: NICOLAS PARDO
     
     ''')
input("Press enter to continue")

