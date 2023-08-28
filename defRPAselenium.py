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


listSCRAPIADO= ([{
               'CUOTA':"",
               'NRO FOLIO':"", 
               'VALOR':"",
               'VENCIMIENTO':"",
                'TOTA A PAGAR':"",
                 }])

listSFormato= ([{
               'pathubicacion':"",
               'Nombre Solicitante':"", 
               'fecha':"",
               'gerente':"",
                'Rut':"",
                'Monto':"",
                'RUTtesoria':"",
                'Direccio':"",
                'Glosagasto':"",
                'Detallegasto':"",
                'CentroGestion':"",
                'Contribuciones':"",
                 }])

browser = Selenium()
library = Windows() 
lib = Files()
app = Application()
año="2018"

def Pyasset(asset):
    lib.open_workbook("PyAsset\Config.xlsx")      #ubicacion del libro
    lib.read_worksheet("Variables")       #nombre de la hoja
    config=lib.read_worksheet_as_table(name='Variables',header=True, start=1).data
    for x in config:
        if x[0]==asset:
            exitdato= str(x[1])
        
            return exitdato

def openweb(u):

<<<<<<< HEAD
    browser.open_available_browser(u,browser_selection="firefox",use_profile=True)
=======
    browser.open_available_browser(u,browser_selection="firefox")
>>>>>>> ef8b97fea3e6ec85e691ae382088ebe9938b7ef4
   
   
    #browser.open_available_browser(url=u,browser_selection="Chrome",use_profile=True,profile_name="franklin ramirez", profile_path=tpath)
    #browser.open_available_browser(url=u)
    browser.maximize_browser_window() 

    validacion= browser.get_text("//DIV[@class='dentro_letra'][text()='Contribuciones']")
    if validacion == 'Contribuciones': print("ingresando a "+validacion) 
    state_tgc_Inicio=True
 
    time.sleep(random.uniform(5,7))

def clickweb(elemento):
    time.sleep(random.uniform(1,2))
    browser.click_element(elemento)
    time.sleep(random.uniform(1,2))

def typeinputText(elemento,texto):
    time.sleep(random.uniform(1,2))
    browser.input_text(elemento,texto)
    time.sleep(random.uniform(1,2))

def obtenertabla(elemento,columna,celdas):
    time.sleep(random.uniform(1,2))
    browser.get_table_cell(locator=elemento,column=columna,row=celdas)
    time.sleep(random.uniform(1,2))

def obtenerTexto(elemento):
    time.sleep(random.uniform(1,2))
    browser.get_text(elemento)
    time.sleep(random.uniform(1,2))

def tiempoespera():
    time.sleep(random.uniform(11,15))

def cerraNavegador():
    browser.close_browser()
    print("----------------------proceso terminado----------------------")

def destacar(elemento):
    browser.highlight_elements(elemento)
    time.sleep(random.uniform(3,7))

def LOGconsulta(Región,Comuna,RolMatriz,Rol):
    print('----------------------Consultado-----------------------------')
    print('region = '+str(Región))
    print('Comuna = '+str(Comuna))
    print('Rol Matriz = '+str(RolMatriz))
    print('Rol = '+str(Rol))

def extraertablita():

    
    print(browser.get_text("//DIV[@id='example_info']/self::DIV"))
    
    scraping=browser.get_text("//TABLE[@id='example']")
    #recorrerFilasDescargas()
    print(scraping)
    return scraping

def recorrerFilasDescargas(carpeta,scraping,rol,hoja):
   
    row=0
    tabledata=txtscraping(carpeta)
    for celda in tabledata:
        row=row+1  
        consecutivo=str(row)     
        try:
                CUOTA = celda.get('CUOTA')
                VALOR=  celda.get('VALOR')
                
                si=str(CUOTA).find("-")
    
                if si == -1:
                    print("la cuota no es visible ")
                else:
                    row=int(row-1  )
                    consecutivo=str(row)
                    obtenerTexto("//TABLE[@id='example']//tr["+str(row)+"]//td[3]")
                    FOLIO=obtenerTexto("//TABLE[@id='example']//tr["+str(row)+"]//td[3]")
                    print("El consecutivo es " + str(consecutivo ))
                    clickweb("//TABLE[@id='example']//tr["+str(row)+"]//td[3]")

                    try:
                        creacioncarpetas (carpeta)
                    except:
                         pass
                    
                    savepdf(carpeta,str(consecutivo ),CUOTA,str(rol))
                    row=row+1 
        except:
             pass
        finally:
            pass

def recorriendoFormatoSolicitud(carpeta,hoja):
    row=0
    
    tabledata=txtscraping(carpeta)
    try:
        for celda in tabledata:
            row=row+1  
            consecutivo=str(row)  
            CUOTA = celda.get('CUOTA')            
            VALOR=  celda.get('VALOR')
            si=str(CUOTA).find("-")

            if si == -1:                
                print("-----------------------------------------------------------------------")
            else:
                print("consultado hoja : "+hoja)
                print("consultado Cuota : "+str(CUOTA))
                print("consultado Monto : "+str(VALOR))
                row=int(row-1  )
                
               
                 
                row=row+1      
    except:
        pass
    finally:
            row=0
        
            tabledata=txtscraping(carpeta)
   
            for celda in tabledata:
                row=row+1  
                consecutivo=str(row)  
                CUOTA = celda.get('CUOTA')            
                VALOR=  celda.get('VALOR')
                si=str(CUOTA).find("-")

                if si == -1:                
                    print("-----------------------------------------------------------------------")
                else:
   
                    print("consultado hoja : "+hoja)
                    print("consultado Cuota : "+str(CUOTA))
                    print("consultado Monto : "+str(VALOR))
                    row=int(row-1  )
                
                    
                    row=row+1      
        
def validacion():
    validacion= browser.get_text("//DIV[@class='dentro_letra'][text()='Contribuciones']")
    if validacion == 'Contribuciones': print("ingresando a "+validacion) 
    return validacion
       
def navegacion(region,comuna,rol1,rol2,ruta,hoja):
    def interacion():
        openweb("https://www.tesoreria.cl/ContribucionesPorRolWEB/muestraBusqueda?tipoPago=PortalContribPresencial")                             
        clickweb("//SELECT[@id='region']/self::SELECT")
        clickweb("//option[text()='"+region+"']")
        clickweb("//SELECT[@id='comunas']")
        clickweb("//option[text()='"+comuna+"']")
        typeinputText("//INPUT[@id='rol']",rol1)
        typeinputText("//INPUT[@id='subRol']",rol2)
        clickweb("//INPUT[@id='btnRecaptchaV3Envio']/self::INPUT")
        tiempoespera()
    interacion()
    
    try: # Validando si la tabla funciona
        valida=obtenerTexto("//TD[@class='celdaContenido2  sorting_1'][text()='No se encontraron Deudas']/self::TD")
        textovalidacion='No se encontraron Deudas'
        if valida == textovalidacion:
                    tabla =extraertablita()
                    export(ruta,tabla)
                    print(tabla) 
                    destacar("//TABLE[@id='example']//tbody//tr//td")
                    
                   
    except:
        try:# proceso de consulta
                    tabla =extraertablita()
                    export(ruta,tabla)
                    destacar("//TABLE[@id='example']//tbody//tr//td") 
                    pdfrol=str(rol1)+"-"+str(rol2)
                    
                    recorrerFilasDescargas(ruta,tabla,str(pdfrol),hoja)
                            

        except:# proceso de consulta reintento #1
            tabla ="""Recatcha no me permitio hacer la consulta"""
            cerraNavegador()
            if """Recatcha no me permitio hacer la consulta"""==tabla:
                print("Reintamos hacer la consulta")
                
                
                try: # Validando si la tabla funciona
                    valida=obtenerTexto("//TD[@class='celdaContenido2  sorting_1'][text()='No se encontraron Deudas']/self::TD")
                    textovalidacion='No se encontraron Deudas'
                    if valida == textovalidacion:
                        tabla =extraertablita()
                        export(ruta,tabla)
                        print(tabla) 
                        destacar("//TABLE[@id='example']//tbody//tr//td") 
                    else:
                         pass    
                       
                        
                except:             
                    # proceso de consulta reintento #2
                        tabla =extraertablita()
                        export(ruta,tabla)
                        destacar("//TABLE[@id='example']//tbody//tr//td")
                        print(tabla) 
                        pdfrol=str(rol1)+"-"+str(rol2)
                       
                        recorrerFilasDescargas(ruta,tabla,str(pdfrol),hoja)
                        
                        
                        
                        
                        
                        pass
                            
                        if validacion()=='Contribuciones':
                            tabla='Contribuciones' 
                            pass

                        elif  tabla == 'No se encontraron Deudas':
                                pass

    finally:
         
         pass
         cerraNavegador()                    

def savepdf(carpeta,consecutivo,cuota,rol):
 base=Pyasset(asset="base")
 txt=base+carpeta
 salida="Cupon de pago "+str(consecutivo)
 if str(consecutivo)=="1":
     consecutivo="1"
 
 
 try:
        file = open(txt+"\\"+salida)
        print(file) # File handler
        file.close()
 except:
    
    library.click("name:imprimirAr")
    time.sleep(4.5)    
    library.send_keys(keys="{CTRL}S")    
    time.sleep(4)

    if str(consecutivo)==str("1"):
        library.send_keys(keys=txt)
        time.sleep(5)
        library.send_keys(keys="{Enter}")
        time.sleep(2)
        library.send_keys(keys="{Alt}N")
        time.sleep(2)
        library.send_keys(keys=str(salida))
        time.sleep(3)
        library.send_keys(keys="{Enter}")
        print("PDF gurdado con exito " + salida)
        library.click("name:imprimirAr")
        time.sleep(1)
        library.send_keys(keys="{Ctrl}W")
        
        origen=txt+"\\"+salida+".pdf"
        destino=txt+"\\"+"Cupon de pago "+str(rol)+" "+str(cuota)+".pdf"

        cambionombre(origen, destino)         

    if str(consecutivo)!=str("1"):
        library.send_keys(keys="{Alt}N")
        time.sleep(2)
        library.send_keys(keys=str(salida))
        time.sleep(3)
        library.send_keys(keys="{Enter}")
        print("PDF gurdado con exito " + salida)
        library.click("name:imprimirAr")
        time.sleep(1)
        library.send_keys(keys="{Ctrl}W")
        
        origen=txt+"\\"+salida+".pdf"
        destino=txt+"\\"+"Cupon de pago "+str(rol)+" "+str(cuota)+".pdf"

        cambionombre(origen, destino)
   
def txtscraping(carpeta):
  f=open('Log Scraping/'+carpeta+".txt","r")

  scrp=[]
  for x in f:
       if x.find(" ")!= 0:
           scrp.append(x)
  liscon=[]

  print(scrp.index)
  
  for u in scrp:
      final=u.find(" ")
      largo=len(u)

      Sumatoria=0

      
      CUOTA=str(u)[0:final]
      Sumatoria=Sumatoria+len(CUOTA)+1

      dato=(str(u)[(Sumatoria):(largo-final)]).find(" ") 
      VALOR=(str(u)[(Sumatoria):Sumatoria+dato]).replace(","," ")
      Sumatoria=Sumatoria+len(VALOR)+1
      

      dato=(str(u)[(Sumatoria):(largo-final)]).find(" ")
      NRO_FOLIO=(str(u)[(Sumatoria):Sumatoria+dato]).replace(","," ")
      Sumatoria=Sumatoria+len(NRO_FOLIO)+1
      


      dato=(str(u)[(Sumatoria):(largo-final)]).find(" ")
      VENCIMIENTO=(str(u)[(Sumatoria):Sumatoria+dato]).replace(","," ")
      Sumatoria=Sumatoria+len(VENCIMIENTO)+1
    

      dato=(str(u)[(Sumatoria):(largo-final)]).find(" ")
      TOTAPAGAR=(str(u)[(Sumatoria):Sumatoria+dato]).replace(","," ")
      Sumatoria=Sumatoria+len(TOTAPAGAR)+1

              
      listSCRAPIADO.append({
               'CUOTA':CUOTA,
               'NRO FOLIO':NRO_FOLIO, 
               'VALOR':VALOR,
               'VENCIMIENTO':VENCIMIENTO,
               'TOTA A PAGAR':TOTAPAGAR,
                 },
      )
         
      
  return listSCRAPIADO
    
def export(Carpeta,tabla):
     
     datosscrap=str(tabla) 
     outmensaje=datosscrap
     outmensaje=outmensaje.replace("VALOR"," ")
     outmensaje=outmensaje.replace("CUOTA"," ")
     outmensaje=outmensaje.replace("VALOR CUOTA"," " )
     outmensaje=outmensaje.replace("NRO FOLIO"," " )
     outmensaje=outmensaje.replace("VENCIMIENTO"," " )
     outmensaje=outmensaje.replace("TOTAL A PAGAR"," " )
     outmensaje=outmensaje.replace("EMAIL"," " )
     outmensaje=outmensaje.replace("DESCARGAR"," " )
     outmensaje=outmensaje.replace("""CUOTA
VALOR CUOTA
NRO FOLIO
VENCIMIENTO
TOTAL A PAGAR
EMAIL
DESCARGAR"""," " )

     try:
        file = open("Log Scraping/"+Carpeta+".txt")
        print(file) # File handler
        file.close()
       
     except:
        print("Archivo no existe se genera uno nuevo  "+ "Log Scraping/"+Carpeta+".txt")
        nom="Log Scraping/"+Carpeta+".txt"     
        f = open(nom, "a")
        f.write(outmensaje)
        f.close() 
                      
def cambionombre(origen, destino):
    archivo = origen
    nombre_nuevo = destino

    

    print("archivo → "+ archivo )
    print("Destino → "+ nombre_nuevo )

    os.rename(archivo, nombre_nuevo)

def Resumen():
    lib.open_workbook("Data\\Resumen_Contribuciones_Terreno_2023.xlsx")      #ubicacion del libro
    lib.read_worksheet("Resumen")                                              #nombre de la hoja
    dtresumen=lib.read_worksheet_as_table(name='Resumen',header=True, start=1).data
    return dtresumen

def master():
   lib.open_workbook("Data\Master.xlsx")      #ubicacion del libro
   lib.read_worksheet("Listado")       #nombre de la hoja
   DtMaster=lib.read_worksheet_as_table(name='Listado',header=True, start=1).data

   return DtMaster

def diligenciarResumen(h,carpeta):
    dtcon=txtscraping(carpeta)
   
       #ahora = datetime.now()
       #consulta=str(ahora.year)
    consulta="2018"
    
     
    for txt in dtcon:
            CUOTA = txt.get('CUOTA') 
                       
            VALOR=  txt.get('VALOR')
            if str(CUOTA)[2:]==consulta : 
                cu = CUOTA             
                v = VALOR
        
                lib.open_workbook("Data\\Resumen_Contribuciones_Terreno_2023.xlsx")      #ubicacion del libro
                lib.read_worksheet("Resumen")                                              #nombre de la hoja
                libroresumen=lib.read_worksheet_as_table(name='Resumen',header=True, start=1).data    

                cantidad=lib.find_empty_row()

                #Ingresamos los valores 
                for celda in range(cantidad):
                
                    Numero=lib.get_cell_value(2+celda,"A")
                    if Numero==h:
                            lib.set_cell_value(2+celda,"E",str(v))
                            lib.set_cell_value(2+celda,"f","pago contribucciones "+str(cu))
                            lib.save_workbook() 
                                
def formatosolicitusd(h,carpeta):

    dtcon=txtscraping(carpeta)
    total=totalMacro(h)

    fecha_actual = datetime.now()

    fecha_formateada = fecha_actual.strftime('%d/%m/%Y')

    #ahora = datetime.now()
    #consulta=str(ahora.year)
    consulta="2018"
    
      
    for txt in dtcon:
            CUOTA = txt.get('CUOTA') 
                       
            VALOR=  txt.get('VALOR')
            if str(CUOTA)[2:]==consulta : 
                cu = CUOTA             
                v = VALOR 
                origen='Data\\Formato Solicitud Pago.xlsx'         
                destino="Formato Solicitud\\"+carpeta +" " +" Cuota " + str(cu) + " Formato Solicitud Pago.xlsx"
                #shutil.copy(origen,destino )

                    
                    
                datac=Resumen()

                for x in datac:
                    if x[0]==h:
                        
                        lib.open_workbook(origen)      
                        lib.read_worksheet("Solicitud")                                              
                        libroresumen=lib.read_worksheet_as_table(name='Solicitud',header=True, start=1).data
                        
                        lib.set_cell_value(8,"D",str(x[8]))

                        lib.set_cell_value(6,"H",str(fecha_formateada))
                        lib.set_cell_value(10,"D","Enrique Carrasco")
                        lib.set_cell_value(12,"D",str(x[3]))
                        lib.set_cell_value(12,"H",str(x[2]))  
                        lib.set_cell_value(14,"C",int(total), fmt="0.00")
                        lib.set_cell_value(20,"C","Teatinos 28, Santiago")
                        lib.set_cell_value(22,"D","pago contribucciones "+str(CUOTA))
                        lib.set_cell_value(24,"D","pago contribucciones "+str(CUOTA))
                        lib.set_cell_value(26,"D",str(x[12]))
                        lib.set_cell_value(28,"D",str(x[12]))
                        lib.set_cell_value(30,"D",str("Contribucciones"))
                        lib.save_workbook(destino)
                        lib.close_workbook()
                        
def diligenciarhojas(h,carpeta,REGION,COMUNA,ROLMATRIZ,RUT,INMOBILIARIA,rol1,rol2):
    dtcon=txtscraping(carpeta)
    R=0
    celda=0

    for txt in dtcon:
         celda=1+celda

    print("el total de celdas es → "+str(celda))
    
    for txt in dtcon:
            CUOTA = txt.get('CUOTA') 
            print(CUOTA)           
            VALOR=  txt.get('VALOR')
        
            lib.open_workbook("Data\\Resumen_Contribuciones_Terreno_2023.xlsx")      #ubicacion del libro
            lib.read_worksheet(str(h))                                                  #nombre de la hoja
            libroresumen=lib.read_worksheet_as_table(name=str(h),header=True, start=1).data    
            
            R=1+R 
                       
            lib.set_cell_value(6+R,"B",RUT) 
            lib.set_cell_value(6+R,"C",INMOBILIARIA)
            lib.set_cell_value(6+R,"D",REGION)
            lib.set_cell_value(6+R,"E",COMUNA)
            lib.set_cell_value(6,"H","Monto")
            lib.set_cell_value(5+R,"H",VALOR,fmt="0.00")
            lib.set_cell_value(6+R,"D",REGION)
            lib.set_cell_value(6+R,"E",COMUNA)
            lib.set_cell_value(6+R,"F",ROLMATRIZ)                   
            lib.save_workbook()
            
    print("el total de R es → "+str(R))
    R=0       
    lib.clear_cell_range("B16:H77")        
    for txt in dtcon:
            CUOTA = txt.get('CUOTA')                        
            VALOR=  txt.get('VALOR')
            R=1+R
            VO=lib.get_cell_value(5+R,"H")
            if VO is None:
                print(VO)
                lib.set_cell_value(5+R,"G"," ")
                lib.set_cell_value(5+R,"F"," ")
                lib.set_cell_value(5+R,"E"," ")
                lib.set_cell_value(5+R,"D"," ")
                lib.set_cell_value(5+R,"C"," ")
                lib.set_cell_value(5+R,"B"," ")
                break
            else:
                lib.set_cell_value(5+R,"G",CUOTA,fmt="0")

        
       
            
    #lib.set_cell_value(7+(R+2),"G","Total") 
    #lib.set_cell_formula("H17","=SUMA(H7:H16)",True)
   
    lib.set_cell_value(6,"H","Monto") 
    lib.save_workbook("Salida\\Resumen_Contribuciones_Terreno_2023.xlsx")#"Salida\\Resumen_Contribuciones_Terreno_2023.xlsx"
    lib.close_workbook ()      

def bakup():
     
     print("Realizamos el bakup")
     origen='Data\\BACKUP\\Resumen_Contribuciones_Terreno_2023.xlsx'         
     destino="Data\\Resumen_Contribuciones_Terreno_2023.xlsx"
     shutil.copy(origen,destino )

def creacioncarpetas (carpeta):

    os.mkdir('PDF/'+carpeta)    
    print("creacion de carpetas  PDF/"+carpeta) 

def Macros (h):
    lib.open_workbook("Data\Macro TGR.xlsm")      
    lib.read_worksheet("MACRO")                                                                     
    libroresumen=lib.read_worksheet_as_table(name="MACRO",header=True, start=1).data 

    lib.set_cell_value(3,"B",str(h))
    lib.save_workbook()
    lib.close_workbook()
    time.sleep(10)

    app.open_application(visible=True)
    try:
         library.click("name:Cerrar")
    except:
         pass

    app.open_workbook('Data\Macro TGR.xlsm')
    app.set_active_worksheet(sheetname="MACRO")
    time.sleep(5)
    app.run_macro("Main")
    time.sleep(5)
    app.save_excel()
    app.quit_application()

def totalMacro(h):
    lib.open_workbook("Data\Resumen_Contribuciones_Terreno_2023.xlsx")      
    lib.read_worksheet(h)                                                                     
    libroresumen=lib.read_worksheet_as_table(name=str(h),header=True, start=1).data 

    TOTAL =lib.get_cell_value(20,"H")

    lib.save_workbook()
    lib.close_workbook()
    return TOTAL

def formatoTotal(h,carpeta):


    totalv=int(totalMacro(h))
    print(str(totalv))
    
    dtcon=txtscraping(carpeta)

    fecha_actual = datetime.now()

    fecha_formateada = fecha_actual.strftime('%d/%m/%Y')
    

    #ahora = datetime.now()
    #consulta=str(ahora.year)
    consulta="2018"
    
      
    for txt in dtcon:
            CUOTA = txt.get('CUOTA') 
                       
            VALOR=  txt.get('VALOR')
            if str(CUOTA)[2:]==consulta : 
                cu = CUOTA             
                v = VALOR 
       
                destino="Formato Solicitud\\"+carpeta +" " +" Cuota " + str(cu) + " Formato Solicitud Pago.xlsx"
                #shutil.copy(origen,destino )

                datac=Resumen()

                for x in datac:
                    if x[0]==h:

                        lib.open_workbook(destino)      
                        lib.read_worksheet("Solicitud")                                                                     
                        libroresumen=lib.read_worksheet_as_table(name='Solicitud',header=True, start=1).data 

                    
                        lib.set_cell_value(14,"C",int(totalv), fmt="0.00")
                        

            

    lib.save_workbook()
    lib.close_workbook()

def fGuardar(h,carpeta):

    dtcon=txtscraping(carpeta)

    fecha_actual = datetime.now()

    fecha_formateada = fecha_actual.strftime('%d/%m/%Y')

    #ahora = datetime.now()
    #consulta=str(ahora.year)
    consulta="2018"
    
      
    for txt in dtcon:
            CUOTA = txt.get('CUOTA') 
                       
            VALOR=  txt.get('VALOR')
            if str(CUOTA)[2:]==consulta : 
                cu = CUOTA             
                v = VALOR 
     
    destino="Formato Solicitud\\"+carpeta +" " +" Cuota " + str(cu) + ".xlsm"

    lib.open_workbook("Data\Resumen_Contribuciones_Terreno_2023.xlsm")      
    lib.read_worksheet("Solicitud")                                                                     
    libroresumen=lib.read_worksheet_as_table(name='Solicitud',header=True, start=1).data 

    lib.set_cell_value(1,"k",int(h))

    lib.save_workbook(destino)
    lib.close_workbook()

def ResumenFinal ():
     lib.open_workbook('Data\Resumen_Contribuciones_Terreno_2023.xlsx')        #ubicacion del libro
     lib.read_worksheet('Resumen')       #nombre de la hoja
     lista=lib.read_worksheet_as_table(name='Resumen',header=True, start=1).data

     ultimaFila= lib.find_empty_row()

     for celda in range(ultimaFila):
         TOTAL= lib.get_cell_value(2+int(celda),"E")
         if TOTAL == "=+'1'!$H$9":
              print("True "+str(TOTAL))
         else:
              print("false "+str(TOTAL))
              HOJA=lib.get_cell_value(2+int(celda),"A")
              
              lib.read_worksheet(str(HOJA))
              tablaTotal=lib.get_cell_value(20,"H")
             

              lib.read_worksheet('Resumen')
              lib.set_cell_value(2+int(celda),"E",int(tablaTotal))



              lib.save_workbook()
              lib.close_workbook()

def limpiarResumen():
     
     
     lib.open_workbook("Data\\Resumen_Contribuciones_Terreno_2023.xlsx")      #ubicacion del libro
     lib.read_worksheet("94")                                                 #nombre de la hoja
     libroresumen=lib.read_worksheet_as_table(name="94",header=True, start=1).data  
      

     lib.clear_cell_range("G7:G1000")
     rango="B{}:H{}"
     
     #Comparaciones 
     item1=lib.get_cell_value(7,"H")
     item2=lib.get_cell_value(8,"H")
     item3=lib.get_cell_value(9,"H")
     item4=lib.get_cell_value(10,"H")
     item5=lib.get_cell_value(11,"H")
     item6=lib.get_cell_value(12,"H")
     item7=lib.get_cell_value(13,"H")
     item8=lib.get_cell_value(14,"H")
     item9=lib.get_cell_value(16,"H")
     item10=lib.get_cell_value(17,"H")
      
     busquedad=0
     for x in range(1000):           
            Cels=str(x+8)
                   
            if item1==lib.get_cell_value(7,"H"):
               busquedad=1+busquedad   
            elif busquedad>1:      
               lib.clear_cell_range(rango.format(Cels,Cels))

     busquedad=0
     for x in range(1000):           
            Cels=str(x+9)
                   
            if item1==lib.get_cell_value(8,"H"):
               busquedad=1+busquedad   
            elif busquedad>1:      
               lib.clear_cell_range(rango.format(Cels,Cels))




     lib.save_workbook()
     lib.close_workbook()

def salida():
     print("Realizamos la salida ")
     origen='Data\Resumen_Contribuciones_Terreno_2023.xlsx'         
     destino="Salida\Resumen_Contribuciones_Terreno_2023.xlsx"
     shutil.copy(origen,destino )

<<<<<<< HEAD


=======
>>>>>>> ef8b97fea3e6ec85e691ae382088ebe9938b7ef4
