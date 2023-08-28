from RPA.Browser.Selenium import Selenium;
from RPA.Excel.Application import Application
from RPA.Windows import Windows
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files;
import time
import random
import os
import shutil
from datetime import date
from datetime import datetime
lib = Files()
fecha_actual = datetime.now()

def masterlibros():
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    DtableFinal=lib.read_worksheet_as_table(name="Export",header=True, start=1).data
    return DtableFinal

def Asignaconsultafecha():
    
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    master=lib.read_worksheet_as_table(name="Export",header=True, start=1).data

    lib.set_cell_value(1,"Q","Consultar")

    celda=1
    try:
        for s in master:
            celda=int(celda+1)
            fecha=str(s[14])
            mes=fecha
            strmes=mes[5:7]
            Disponibles=str(s[10])

            if str(fecha) == "None":
             lib.set_cell_value(int(celda),"Q","SI")
                        

        

                #valida si las celdas la fecha actual 
            if fecha_actual.strftime('%Y') in fecha:
                    mes_string=fecha_actual.strftime('%m')
                    mesConsulta=str(int(mes_string)-1)
                #valida si la fecha actual tiene un largo de 1 o 2 
                    if len(mesConsulta) == 1:
                        mesConsultacero=str("0"+mesConsulta)
                    else:
                        mesConsultacero=str(mesConsulta)
                #valida si el mes consulta es el anterior     
                    if mesConsultacero in strmes: 
                        if str(s[8]) == "None":
                            break
                        else:   
                            lib.set_cell_value(int(celda),"Q","SI")
        

    except:
            pass
        
    lib.save_workbook()
    lib.close_workbook()

def AsignaCodigoComuna():
    lib = Files()
    lib.open_workbook('Data\Base de Existencias Unidades.xlsx')       
    lib.read_worksheet("Export")       
    master=lib.read_worksheet_as_table(name="Export",header=True, start=1).data

    lib.set_cell_value(1,"P","Codigo comuna")

    celda=1
    
    for s in master:

        try:
            if str(s[8]) == "None":
             break
            else:
                celda=int(celda+1)
                region=str(s[8])
                mregion=region.upper()
                comuna=str(s[9])
                fecha=str(s[14])
                mes=fecha
                codigo=CodigoComuna(region,comuna)
                lib.set_cell_value(int(celda),"P",codigo)
                lib.set_cell_value(int(celda),"I",mregion)

        except:
            pass




    lib.save_workbook()
    lib.close_workbook()

def mesconsultar():
 fecha_actual = datetime.now()
 #fecha_formateada = fecha_actual.strftime('%d/%m/%Y')
 fecha_formateada = fecha_actual.strftime('%Y-%m')
 print(fecha_formateada)

def CodigoComuna(region,comuna):
    regionOUT=region.replace("Region ","")
    lib.open_workbook('Data\Codigos Comunas.xlsx')       
    lib.read_worksheet(regionOUT)       
    listacomuna=lib.read_worksheet_as_table(name=regionOUT,header=True, start=1).data

    for x in listacomuna :
        #print(str(x[0])+"="+str(comuna))
        if str(x[0]) == str(comuna):
            consulta=x[2]
            
            return consulta
            break

def task_Modelos():
        

        tiempoInicio=time.time()
        dt=masterlibros()
        for resumen in dt:
            region=str(resumen[8])
            comuna=str(resumen[9])
            RolMatriz=str(resumen[7])
            print(str(datetime.now())+"   :Consultando "+ RolMatriz +" "+region+" - "+comuna)
            AsignaCodigoComuna()
            
        tiempoFinal=time.time() 
        TiempoTotal=tiempoFinal-tiempoInicio
        print("Tiempo total de ejecucion es "+str(TiempoTotal) + " seg")

def logscraping(carpeta,Rolmatriz):
    f=open('Log Scraping/'+carpeta+".txt","r")
    CapturaSCRAPIADO= ([{ }])
    

    for x in f:
         CUOTA=x[0:7]
         VALOR_CUOTA=x[8:15]
         NRO_FOLIO=x[16:25]
         VENCIMIENTO=x[26:36]
         TOTAL_A_PAGAR=x[37:48]
         EMAIL=x[48:55]

        
         if str(x[48:55]) in " ":
            print("------------------------")
         else: 
             CapturaSCRAPIADO.append({
               'CUOTA':CUOTA,
               'VALOR CUOTA':int(VALOR_CUOTA), 
               'NRO FOLIO':NRO_FOLIO,
               'VENCIMIENTO':VENCIMIENTO,
               'TOTAL A PAGAR':TOTAL_A_PAGAR,
               'EMAIL':EMAIL
                 })
                
             
    lib.create_workbook()
    lib.create_worksheet(Rolmatriz)
    lib.append_rows_to_worksheet(CapturaSCRAPIADO, header=True)
    lib.save_workbook('Excel/'+carpeta+".xlsx")
    lib.close_workbook()

def salida(carpeta,Rolmatriz,rut,inmobiliaria,Región,Comuna):

    """los datos necesarios son :
    carpeta: str
    Rolmatriz: str
    rut: str
    inmobiliaria: str
    Región: str
    Comuna: str

    """
   
    lib.open_workbook("Excel/"+carpeta+'.xlsx')        #ubicacion del libro
    lib.read_worksheet(str(Rolmatriz))       #nombre de la hoja
    outlista=lib.read_worksheet_as_table(name=str(Rolmatriz),header=True, start=1).data
    lib.close_workbook()
    recol=0
#encabezados
    tabla=([{ }])
    
    for x in outlista:

     if str(x[1])=="None":
        total=0

     else:

        try:
            tabla.append({
               'RUT':rut,
               'INMOBILIARIA':inmobiliaria, 
               'Región':Región,
               'cuota':str(x[0]),
               'Comuna':Comuna,
               'Rolmatriz':Rolmatriz,
               'Informacion Tesoreria':"Informacion_Tesoreria",
               'Monto':str(x[1])
                
                 })
         
            total=total+int(x[1])
            print(total)

        except:
            pass
#Diligenciamos los totales        
    tabla.append({

               'RUT':"",
               'INMOBILIARIA':"", 
               'Región':"",
               'cuota':"",
               'Comuna':"",
               'Rolmatriz':"",
               'Informacion Tesoreria':"total",
               'Monto':str(total)
                
                 })
        
    try:   
            #open("Salida\SalidaUnidadesVendidas.xlsx")
            lib.open_workbook("Salida\SalidaUnidadesVendidas.xlsx")
            print("el libro existe")
            Existe=lib.worksheet_exists(rut)
            print(Existe)
            
            if Existe == True:
                 
                lib.read_worksheet("Salida\SalidaUnidadesVendidas.xlsx")
                DtableFinal=lib.read_worksheet_as_table(name=rut,header=True, start=1).data
                lib.append_rows_to_worksheet(tabla, header=True)
    except:
            print("el libro no existe")
            Existe=False
            lib.create_workbook("Salida\SalidaUnidadesVendidas.xlsx")
            lib.create_worksheet(rut)
            lib.append_rows_to_worksheet(tabla, header=True)
            
        

       
    lib.save_workbook()
    lib.close_workbook()
    tabla=([{ }]) 


"""
carpeta="San Bernardo 4505-54"
Rolmatriz="4505-54"
rut="11111111"
inmobiliaria="ejemplo"
Región="metropolitana "
Comuna="nunoa"


#logscraping(carpeta)
salida(carpeta,Rolmatriz,rut,inmobiliaria,Región,Comuna)"""








