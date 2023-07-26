from RPA.Browser.Selenium import Selenium;
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

def Pyasset(asset):
    lib.open_workbook("PyAsset\Config.xlsx")      #ubicacion del libro
    lib.read_worksheet("Variables")       #nombre de la hoja
    config=lib.read_worksheet_as_table(name='Variables',header=True, start=1).data
    for x in config:
        if x[0]==asset:
            exitdato= str(x[1])
        
            return exitdato

def openweb(url):
    browser.open_available_browser(url,browser_selection="firefox")
    browser.maximize_browser_window() 
    validacion= browser.get_text("//DIV[@class='dentro_letra'][text()='Contribuciones']")
    if validacion == 'Contribuciones': print("ingresando a "+validacion) 
    state_tgc_Inicio=True
    time.sleep(random.uniform(1,3))



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
                
                FormatoSolicitud(hoja,CUOTA, VALOR) 
                 
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
                    recorrerFilasDescargas(ruta,tabla,rol2,hoja)
                    
                   
                    
                    
                   
                    
                            

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
                        recorrerFilasDescargas(ruta,tabla,rol2,hoja)
                        
                        
                        
                        
                        
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
     

def FormatoSolicitud(h,CUOTA, valor):
   tba=Resumen()
 #formato de resumen 
   for rem in tba:
        N=str(rem[0])
        
        
        if N == str(h):
         
         
         print("Diligenciamos formato de Solicitud de la hoja  "+str(h) +"  Con la cuota  "+str(CUOTA)+"  con el monto a pagar  "+str(valor))    
         Rut=str(rem[2])
         
         origen='Data\\Formato Solicitud Pago.xlsx'         
         destino="Formato Solicitud\\"+CUOTA +" " +" Monto " + str(valor) + " Formato Solicitud Pago.xlsx"
         
         print("Destinos → "+destino)
    #Copias el libro de formato de solicitud
         shutil.copy(origen,destino )
         
        
         lib.open_workbook(destino)        #ubicacion del libro
         lib.read_worksheet("Solicitud")                                                                   #nombre de la hoja                                          
         formato=lib.read_worksheet_as_table(name='Solicitud',header=True, start=1).data
        
         NombreSolicitante=rem[8]
         print(NombreSolicitante)
         now = str(datetime.now())
         print(now)
         gerente=Pyasset(asset="Gerente")
         print(gerente)
         InmobiliariaGiradora=str(rem[3])
         print(InmobiliariaGiradora)
         Monto=str(valor)
         print(Monto)
         Rut=str(rem[2])
         print(Rut)
         RUTtesoria=Pyasset(asset="RutTesoreria")
         print(RUTtesoria)
         Dirección=Pyasset(asset="Dirección")
         print(Dirección)
         Glosagasto="Pago de sobretasa cuota "+str(CUOTA)
         print(Glosagasto)
         Detallegasto= Glosagasto
         print(Detallegasto)
         CentroGestion=str(rem[12])
         print(CentroGestion)


        """   listSFormato.append({
                    'pathubicacion':"Formato Solicitud\\"+CUOTA +" " +" Monto " + str(valor) + " Formato Solicitud Pago.xlsx",
                    'Nombre Solicitante':rem[8], 
                    'fecha':str(datetime.now()),
                    'gerente':str(Pyasset(asset="Gerente")),
                        'Rut':rem[2],
                        'Monto':str(valor),
                        'RUTtesoria':str(Pyasset(asset="RutTesoreria")),
                        'Direccio':str(Pyasset(asset="Dirección")),
                        'Glosagasto':"Pago de sobretasa cuota "+str(CUOTA),
                        'Detallegasto':"Pago de sobretasa cuota "+str(CUOTA),
                        'CentroGestion':str(rem[12]),
                        'Contribuciones':str(rem[12]),
                        },
            )





        except:
            
            pass
         
  
        print(listSFormato) 

        
        return listSFormato"""
       
def diligenciamiento_formatos():
 dtform = listSFormato   


 for dt in dtform:
            pathubicacion=  dt.get("pathubicacion")
            NombreSolicitante= dt.get("NombreSolicitante")
            now=dt.get("fecha")
            gerente=dt.get("gerente")
            InmobiliariaGiradora=dt.get("pathubicacion")
            Rut=dt.get("Rut")
            Monto=dt.get("Monto")
            RUTtesoria=dt.get("RUTtesoria")
            Dirección=dt.get("Direccio")
            Glosagasto=dt.get("Glosagasto")
            Detallegasto=dt.get("Detallegasto")
            CentroGestion=dt.get("CentroGestion")
            Contribuciones=dt.get("Contribuciones")

            celdaexcel (pathubicacion,
                        NombreSolicitante,
                        now,gerente,
                        InmobiliariaGiradora,
                        Rut,
                        Monto,
                        RUTtesoria,
                        Dirección,
                        Glosagasto,
                        Detallegasto,
                        CentroGestion,
                        Contribuciones)
            
  


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

def celdaexcel (pathubicacion,NombreSolicitante,now,gerente,InmobiliariaGiradora,Rut,Monto,RUTtesoria,Dirección,Glosagasto,Detallegasto,CentroGestion,Contribuciones):
   lib.open_workbook(pathubicacion)        #ubicacion del libro
   lib.read_worksheet('Solicitud')       #nombre de la hoja
   lib.read_worksheet('Solicitud')       #activamos las cabezeras
   lista=lib.read_worksheet_as_table(name='Solicitud',header=True, start=1).data 
   lib.set_cell_value(8,"D",NombreSolicitante)
   lib.set_cell_value(6,"H",now)
   lib.set_cell_value(10,"D",gerente)
   lib.set_cell_value(12,"D",InmobiliariaGiradora)
   lib.set_cell_value(12,"H",Rut)  
   lib.set_cell_value(14,"C",Monto)
   lib.set_cell_value(14,"C",RUTtesoria)
   lib.set_cell_value(20,"C",Dirección)
   lib.set_cell_value(22,"D",Glosagasto)
   lib.set_cell_value(24,"D",Detallegasto)
   lib.set_cell_value(28,"D",CentroGestion)
   lib.set_cell_value(30,"D",Contribuciones)
   lib.save_workbook()
"""
pathubicacion="Formato Solicitud//9-2017  Monto 107918995 Formato Solicitud Pago.xlsx"
NombreSolicitante="NombreSolicitante"
now="now"
gerente="gerente"
InmobiliariaGiradora="InmobiliariaGiradora"
Rut="Rut" 
Monto="Monto"
RUTtesoria="RUTtesoria"
Dirección="gkks"
Glosagasto="Glosagasto"
Detallegasto="Detallegasto"
CentroGestion="CentroGestion"
Contribuciones="Contribuciones"


celdaexcel (pathubicacion,
            NombreSolicitante,
            now,
            gerente,
            InmobiliariaGiradora,
            Rut,
            Monto,
            RUTtesoria,
            Dirección,
            Glosagasto,
            Detallegasto,
            CentroGestion,
            Contribuciones,)

"""


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

                #Limpiando celdas
                #for celda in range(cantidad):
                #   lib.set_cell_value(2+celda,"E","")
                #   lib.set_cell_value(2+celda,"F","Sin rol vigente ")

                #Ingresamos los valores 
                for celda in range(cantidad):
                
                    Numero=lib.get_cell_value(2+celda,"A")
                    if Numero==h:
                            lib.set_cell_value(2+celda,"E",str(v))
                            lib.set_cell_value(2+celda,"f","Contribucion Cuota "+str(cu))
                            lib.save_workbook()        


         

def formatosolicitusd(h,carpeta):
    dtcon=txtscraping(carpeta)

    #ahora = datetime.now()
    #consulta=str(ahora.year)
    consulta="2018"
    
      
    for txt in dtcon:
            CUOTA = txt.get('CUOTA') 
                       
            VALOR=  txt.get('VALOR')
           # if str(CUOTA)[2:]==consulta : 
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
                      
                      lib.set_cell_value(6,"H",str(datetime.now()))
                      lib.set_cell_value(10,"D","franklin")
                      lib.set_cell_value(12,"D",str(x[3]))
                      lib.set_cell_value(12,"H",str(x[2]))  
                      lib.set_cell_value(14,"C",str(v))
                      lib.set_cell_value(14,"C","60.805.000-0")
                      lib.set_cell_value(20,"C","Teatinos 28, Santiago")
                      lib.set_cell_value(22,"D","Pago de sobretasa cuota "+str(CUOTA))
                      lib.set_cell_value(24,"D","Pago de sobretasa cuota "+str(CUOTA))
                      lib.set_cell_value(28,"D",str(x[12]))
                      lib.set_cell_value(30,"D",str(x[12]))
                      lib.save_workbook(destino)
                      lib.close_workbook()
                          

def diligenciarhojas(h,carpeta,REGION,COMUNA,ROLMATRIZ):
    dtcon=txtscraping(carpeta)
    R=0

    for txt in dtcon:
            CUOTA = txt.get('CUOTA') 
            print(CUOTA)           
            VALOR=  txt.get('VALOR')
        
            lib.open_workbook("Data\\Resumen_Contribuciones_Terreno_2023.xlsx")      #ubicacion del libro
            lib.read_worksheet(h)                                              #nombre de la hoja
            libroresumen=lib.read_worksheet_as_table(name=h,header=True, start=1).data    
            
            RUT=lib.get_cell_value(7,"B")
            INMOBILIARIA=lib.get_cell_value(7,"C")
            Reg=lib.get_cell_value(7,"D")
            Com=lib.get_cell_value(7,"E")
            RMatriz=lib.get_cell_value(7,"F")
            InfoTesoreria=lib.get_cell_value(7,"G")

            R=1+R
            if R==1:
            
                lib.set_cell_value(6+R,"B",RUT) 
                lib.set_cell_value(6+R,"C",INMOBILIARIA)
                lib.set_cell_value(6+R,"D",Reg)
                lib.set_cell_value(6+R,"E",Com)
                lib.set_cell_value(6+R,"F",RMatriz)
                lib.set_cell_value(6+R,"G",RMatriz)
                
                lib.set_cell_value(6,"H","Monto")
                lib.set_cell_value(5+R,"H",VALOR)
                #lib.set_cell_value(6+R,"D",REGION)
                #lib.set_cell_value(6+R,"E",COMUNA)
                #lib.set_cell_value(6+R,"F",COMUNA)
                
                lib.save_workbook()
            else:
                 print("------------------------------------------")
            
            
    lib.set_cell_value(7+(R+2),"G","Total") 
    lib.set_cell_formula("H18","=SUBTOTALES(9;H7:H17)")
    lib.save_workbook("Salida\\Resumen_Contribuciones_Terreno_2023.xlsx") 
    lib.close_workbook ()      


def bakup():
     
     print("Realizamos el bakup")

    

     origen='Data\\BACKUP\\Resumen_Contribuciones_Terreno_2023.xlsx'         
     destino="Data\\Resumen_Contribuciones_Terreno_2023.xlsx"
     shutil.copy(origen,destino )


         


      
            



    



