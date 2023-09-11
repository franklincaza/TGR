import models
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
from shutil import rmtree
import openpyxl
from openpyxl import Workbook

wb = Workbook()
lib = Files()
fecha_actual = datetime.now()
app = Application()
import logging



Dt=models.master()

logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s | %(name)s | %(levelname)s | %(message)s',
                    filename= 'log procesos' )

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
         TOTAL_A_PAGAR=str(x[37:48]).replace(".","")
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

def salidaout(carpeta,Rolmatriz,rut,inmobiliaria,Región,Comuna):

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
     totalapagar = x[4].replace(".", "").replace(",", "")
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
               'Informacion Tesoreria':str(x[0]),
               'Monto':totalapagar
                
                 })
         
            total=total+totalapagar
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
               'Monto':totalapagar
                
                 })
        
    try:   
            #open("Salida\SalidaUnidadesVendidas.xlsx")
            lib.open_workbook("Out Hojas Scraping/"+carpeta+".xlsx")
            print("el libro existe")
            Existe=lib.worksheet_exists(inmobiliaria)
            print(Existe)
            
            if Existe == True:
                print("prueba")
                 
                lib.read_worksheet("Out Hojas Scraping/"+carpeta+".xlsx")
                DtableFinal=lib.read_worksheet_as_table(name=inmobiliaria,header=True, start=1).data
                lib.append_rows_to_worksheet(tabla, header=True)


    except:
            print("el libro no existe")
            Existe=False
            lib.create_workbook("Out Hojas Scraping/"+carpeta+".xlsx")
            lib.create_worksheet(inmobiliaria)
            lib.append_rows_to_worksheet(tabla, header=True)
            
        

       
    lib.save_workbook("Out Hojas Scraping/"+carpeta+".xlsx")
    lib.close_workbook()

    f = open("Out Hojas Scraping/"+carpeta+".txt", "a")
    f.write(str(tabla))
    f.close()

    tabla=([{ }]) 

def total(carpeta):
    contenido = carpeta +".txt"

    f=open('Log Scraping/'+ contenido,"r")
    for wq in f:
        if str(wq[1:4]) != "    ":
            w = open("Log Scraping/total.txt", "a")
            w.write(wq.replace("Enviar",carpeta))
      
def lectura(carpeta,Rolmatriz,rut,inmobiliaria,Región,Comuna):
  contenido= carpeta
  ubicacion=contenido.replace(".txt","")
  total(carpeta=ubicacion)
  f=open('Log Scraping/total.txt',"r")
  CapturaSCRAPIADO= ([{ }])
  for x in f:
         RUT=rut
         INMOBILIARIA=inmobiliaria
         Region=Región
         Comuna=Comuna
         Rol_Matriz=Rolmatriz
         Informacion_Tesoreria=x[0:7]
         monto=str(x[37:48]).replace(".","")
         

        
         if str(x[48:55]) in " ":
            print("------------------------")
         else: 
             CapturaSCRAPIADO.append({
               'RUT':rut,
               'INMOBILIARIA':inmobiliaria, 
                'Región':Region,
               'Comuna':Comuna,
               'Rol_Matriz':Rolmatriz,
               'Informacion_Tesoreria':Informacion_Tesoreria,
               'monto':int(monto),
                 })
  
  return CapturaSCRAPIADO
 
def creacionExcelResumen():
     lib.create_workbook()
     for i in range(120):  
        lib.create_worksheet(str(i))
     for X in range(120):
        lib.set_cell_format(2+X,"G",fmt=0.00)  
     lib.save_workbook("Log Scraping\Resumen.xlsx")

def datosexceltotal(h,tabla):

    lib.open_workbook('Log Scraping\Resumen.xlsx')        #ubicacion del libro
    lib.read_worksheet(str(h))  

    lib.read_worksheet(str(h))       #activamos las cabezeras
    listaTotales=lib.read_worksheet_as_table(name=str(h),header=True, start=1).data

    lib.append_rows_to_worksheet(tabla,True)      
    lib.save_workbook()
    lib.close_workbook()
    
def subtotal(h):
    
    lib.open_workbook('Log Scraping\Resumen.xlsx')        #ubicacion del libro
    lib.read_worksheet(str(h))      #nombre de la hoja
    lib.read_worksheet(str(h))       #activamos las cabezeras
    listaTotales=lib.read_worksheet_as_table(name=str(h),header=True, start=1).data


    a=1
    subtotal=0
    while a<100:
        lib.set_cell_format(1+a,"G",fmt=0.00)
        Monto=lib.get_cell_value(1+a,"G")       
#cuando encontramos los valores fina        
        if a == 10 or a == 20 or a == 30 or a == 40 or a == 50 or a == 60 or a == 70 or a == 80 or a == 90 or a == 100:
           lib.insert_rows_before(row=3+a)

           try:     
            subtotal = Monto + subtotal
           except:
               pass

           lib.set_cell_value(3+a,"G",subtotal,fmt=0.00)
          
           lib.set_cell_value(3+a,"F","total")
           subtotal=0
           Monto=0
           a=2+a
           lib.save_workbook()
        else:
           a=1+a 

           try:     
            subtotal = Monto + subtotal
           except:

            pass  
            
           lib.save_workbook()
           #Cuando no encontramos valores
        if Monto is None:
               lib.insert_rows_before(row=1+a)
               lib.set_cell_value(2+a,"E","total")
               try:     
                    subtotal = Monto + Monto
               except:
                    pass
               
               lib.set_cell_value(2+a,"G",subtotal,fmt=0.00)         
              
               subtotal=0
               a=1+a
               lib.save_workbook()
               
               break
        
      

   
    print("termino ")
    lib.close_workbook()

def copiamosformatos(h):  
    tabla=([{ }])
    lib.open_workbook('Log Scraping\Resumen.xlsx')        #ubicacion del libro
    lib.read_worksheet(h)      #nombre de la hoja
    lib.read_worksheet(h)      #activamos las cabezeras
    listaTotales=lib.read_worksheet_as_table(name=h,header=True, start=1).data
  
    return listaTotales

def reporteHojas(h,tabla):
    lib.open_workbook("Data\Resumen_Contribuciones_Terreno_2023.xlsx")      
    lib.read_worksheet(str(h)) 
    lib.clear_cell_range("B7:Z1000")
    celda=0
    for  x in tabla:
            celda=1+celda
            if x[5]=="total":
                lib.set_cell_value(celda+6,"G",x[5])
                lib.set_cell_value(celda+6,"H",x[6])
            else:
                lib.set_cell_value(celda+6,"B",x[0])
                lib.set_cell_value(celda+6,"C",x[1])
                lib.set_cell_value(celda+6,"D",x[2])
                lib.set_cell_value(celda+6,"E",x[3])
                lib.set_cell_value(celda+6,"F",x[4])
                lib.set_cell_value(celda+6,"G",x[5])
                lib.set_cell_value(celda+6,"H",x[6])

    
    lib.save_workbook ()                                             
    lib.close_workbook()
    print("proceso terminado")

def lecturaALL(carpeta,rut,Región,Comuna,Rolmatriz,inmobiliaria,h): 
 lib.open_workbook("Data\Resumen_Contribuciones_Terreno_2023.xlsx")      
 lib.read_worksheet(str(h))
 lib.set_active_worksheet(str(h)) 
 lib.clear_cell_range("B7:Z1000")
 try:#capturando datos txt
  f=open('Log Scraping/'+carpeta+'.txt',"r")
  
  CapturaSCRAPIADO= ([{ }])
  for x in f:
        linea=x.replace("Enviar","")
        print(linea)
        RUT=rut
        INMOBILIARIA=inmobiliaria
        Region=Región
        Comuna=Comuna
        Rol_Matriz=Rolmatriz
        Informacion_Tesoreria=linea[0:7]
        monto=str(linea[37:47]).replace(".","")
        
        
        print(monto)
        if str(x[48:55]) in " ":
            print("------------------------")
        else: 
            CapturaSCRAPIADO.append({
            'RUT':rut,
            'INMOBILIARIA':inmobiliaria, 
             'Región':Region,
            'Comuna':Comuna,
            'Rol_Matriz':Rolmatriz,
            'Informacion_Tesoreria':Informacion_Tesoreria,
            'monto':int(monto),
              })
 except:#Si es error
    print("No pudo leer informacion" + carpeta)
    logging.error("No pudo leer infor   macion" + carpeta)
    pass
 celda=0
 for x in CapturaSCRAPIADO: #Ingresamos la tabla 
    celda=celda+1
    lib.set_cell_value(6+celda,"B",x.get('RUT'))
    lib.set_cell_value(6+celda,"C",inmobiliaria)
    lib.set_cell_value(6+celda,"D",x.get('Región'))
    lib.set_cell_value(6+celda,"E",x.get('Comuna'))
    lib.set_cell_value(6+celda,"F",x.get('Rol_Matriz'))
    lib.set_cell_value(6+celda,"G",x.get('Informacion_Tesoreria'))
    lib.set_cell_value(6+celda,"H",x.get('monto'))
 a=0
 subtotal=0
 while a<200: # agregamos los sub totales
     lib.set_cell_format(8+a,"H",fmt=0.00)
     Monto=lib.get_cell_value(8+a,"H")            
     if a == 10 or a == 22 or a == 34 or a == 46 or a == 58 or a == 70 or a == 82 or a == 94 or a == 106 or a == 118:
        lib.insert_rows_before(row=9+a)
        #try:subtotal = Monto + subtotal
        #except:pass
            
        lib.set_cell_value(8+a,"H",subtotal,fmt=0.00)         
        lib.set_cell_value(8+a,"G","total")  
        lib.set_cell_value(8+a,"B","-")
        lib.set_cell_value(8+a,"C","-")
        lib.set_cell_value(8+a,"D","-")
        lib.set_cell_value(8+a,"E","-")
        lib.set_cell_value(8+a,"F","-")      
        subtotal=0
        Monto=0

        a=2+a
        lib.save_workbook()
     else:
        a=1+a 
        try:subtotal = Monto + subtotal
        except:pass  
     lib.save_workbook()
           #Cuando no encontramos valores
     if Monto is None:
        print(a)
        print(subtotal)
        lib.insert_rows_before(row=9+a)
        lib.set_cell_value(8+a,"G","total")
        lib.set_cell_value(8+a,"B","-")
        lib.set_cell_value(8+a,"C","-")
        lib.set_cell_value(8+a,"D","-")
        lib.set_cell_value(8+a,"E","-")
        lib.set_cell_value(8+a,"F","-")
        try:     
            subtotal = Monto + Monto
        except:
            pass
        lib.set_cell_value(8+a,"H",subtotal,fmt=0.00)                   
        subtotal=0
        a=1+a
        lib.save_workbook()
        break
     

         
     
     
 lib.delete_rows(7)
 lib.save_workbook()             
 lib.close_workbook()

def sumaResumen(h):
    lib.open_workbook("Data\Resumen_Contribuciones_Terreno_2023.xlsx") 
    lib.read_worksheet(str(h))
    lib.set_active_worksheet(str(h))
    cantidad=lib.find_empty_row()
    total=0
    # Calculado los totales para el resumen 
    for resumen in range(int(cantidad)):
        captura = lib.get_cell_value(6+resumen,"G")
        PRUEBA=lib.get_cell_value(7,"G")
        if captura == "total":
                if captura is None: valor = 0
                valor=lib.get_cell_value(6+resumen,"H")
                total= float(valor) + total             
    # Ingresando los totales en los resumen 
    lib.read_worksheet("Resumen")
    lib.set_active_worksheet("Resumen")
    cantidad=lib.find_empty_row()
    for resumen in range(int(cantidad)):
        numero=lib.get_cell_value(resumen+1,"A")
        if h == numero:
            lib.set_cell_value(1+resumen,"E",total)
    
    # Ingresando los totales en los resumen 
    lib.read_worksheet(str(h))
    lib.set_active_worksheet(str(h))
    cantidad=lib.find_empty_row()
    fecha_actual = datetime.now()
    fecha_formateada = fecha_actual.strftime('%Y')
    AÑO=[]
    CUOTA=[]
    for resumen in range(int(cantidad)):
        captura = lib.get_cell_value(6+resumen,"G")
        if str(captura).__contains__(fecha_formateada):
            CUOTA.append(captura[0:1])
            AÑO.append(captura[2:9])

    print(min(CUOTA))    
    strcuota = min(CUOTA)
    straño = max(AÑO)  
    pago= "pago contribucciones "+  str(strcuota) + " - " + str(straño)

    lib.read_worksheet("Resumen")
    lib.set_active_worksheet("Resumen")
    for resumen in range(int(cantidad)):
        numero=lib.get_cell_value(resumen+1,"A")
        if h == numero:
            
            lib.set_cell_value(1+resumen,"F",pago)

    lib.save_workbook()
    lib.close_workbook()

def pagocontribucciones(h) :
    lib.open_workbook("Data\Resumen_Contribuciones_Terreno_2023.xlsx") 
    lib.read_worksheet(str(h))
    lib.set_active_worksheet(str(h))
    cantidad=lib.find_empty_row()
    fecha_actual = datetime.now()
    fecha_formateada = fecha_actual.strftime('%Y')
    AÑO=[]
    CUOTA=[]
    for resumen in range(int(cantidad)):
        captura = lib.get_cell_value(6+resumen,"G")
        if str(captura).__contains__(fecha_formateada):
            CUOTA.append(captura[0:1])
            AÑO.append(captura[2:9])

    print(min(CUOTA)) 
    print(min(AÑO))    
    strcuota = min(CUOTA)
    straño = max(AÑO)  
    pago= "pago contribucciones "+  str(strcuota) + " - " + str(straño)
    print(pago)
    linea=lib.find_empty_row()
    lib.read_worksheet("Resumen")
    lib.set_active_worksheet("Resumen")
    try:
        for resumen in range(int(500)):
            numero=lib.get_cell_value(resumen+1,"A")
            
            if h == numero:          
                lib.set_cell_value(1+resumen,"F",pago)


    except:pass
    lib.save_workbook()
    lib.close_workbook()


def sumatorias():
    Dt=models.master()
    
    for dtable in Dt:
            try:
                    if dtable[5] == "SI":
                        strcomuna="{} [{}]"
                        strrolmatriz="{}- {}"
                        Rut=str(dtable[0])
                        Inmobiliaria=dtable[1]
                        Asset=dtable[2]
                        Carpeta=dtable[3]
                        Hoja=dtable[4]
                        Activo=dtable[5]
                        region=dtable[6]
                        comuna=strcomuna.format(dtable[7],dtable[10])
                        rolmatriz=strrolmatriz.format(dtable[8],dtable[9])
                        rol1=dtable[8]                               
                        rol2=dtable[9]
                        Codigo=dtable[10]
                        print(Carpeta + " -inicio")
                        print(Carpeta + " -Finalizado")
                        lecturaALL(Carpeta,Rut,region,comuna,rolmatriz,Inmobiliaria,Hoja)                      
                        
                        sumaResumen(Hoja)
                        pagocontribucciones(Hoja)
                        logscraping(Carpeta,rolmatriz)
            except PermissionError:
                logging.error("Error ´por tener el libro de excel abierto")
                print("Error ´por tener el libro de excel abierto")
                pass
            except :
                logging.error("error en los calculos de los sub totales")
                print("error en los calculos de los sub totales")
                pass
  


