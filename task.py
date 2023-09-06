import defRPAselenium
import moldesTerrenos
import models
import moldesTerrenos
from RPA.Browser.Selenium import Selenium;
import os
from shutil import rmtree
import time
import logging





logging.basicConfig(level=logging.INFO,
                    format='%(asctime)s | %(name)s | %(levelname)s | %(message)s',
                    filename= 'log procesos' )

#Calculamos tiempo de ejecucion
tiempoInicio=time.time()


browser = Selenium()
Dt=models.master()
urlbase=str(defRPAselenium.Pyasset(asset="base"))
UrlMacro=defRPAselenium.Pyasset(asset="Ruta ")
libro=defRPAselenium.Pyasset(asset="LIBRO ")

def eliminarcarpetas():
    try:
        rmtree("PDF")
        rmtree("CSV")
        rmtree("Log Scraping")
        rmtree("Formato Solicitud")
        rmtree("Salida")  
        rmtree("Excel")    
        rmtree("Out Hojas Scraping ")
        logging.info("Eliminamos carpetas")


    except:
        pass

def Creacionescarpetas():
    logging.info("Creado las carpetas para PDF's")

    try:
        os.mkdir('PDF')
        os.mkdir('CSV')
        os.mkdir('Excel')
        os.mkdir("Formato Solicitud")  
        os.mkdir("Log Scraping")
        os.mkdir("Salida") 
        os.mkdir("Excel") 
        os.mkdir("Out Hojas Scraping ") 

    except:
        pass

def task():
     
            for dtable in Dt:
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
    

                    
                    logging.info("creacion de hoja Resumen")
                    logging.info(defRPAselenium.LOGconsulta(region,comuna,rol1,rol2))
                    try:
                     os.remove("Log Scraping\total.txt")
                    except:
                        pass
                    
                    cantidad=0
                    consulta=True
                    while consulta==True:
                        cantidad=1+cantidad
                        estadoConsulta="consulta de terrenos de la region {0} , comuna {1} y rolmatriz {2}-{3} --- consulta # {4} ".format(region,comuna,rol1,rol2,cantidad)
                        logging.info(estadoConsulta)

                        try:
                            tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                            stado=False
                            consulta=False 
 
                        except:                      
                        #Tercer reintento para garantizar continuidad si encuentra Recatchat
                            try: 
                                logging.warning("segundo reintento ") 
                                time.sleep(3)
                                tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                                stado=False
                                consulta=False 

                            except:
                                
                                logging.warning("tercer reintento ") 
                                try:
                                    time.sleep(3)
                                    tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                                    stado=False
                                    consulta=False 
                                    pass
                                except:
                                    pass
                            
                            finally:
                                pass   
                                defRPAselenium.cerraNavegador() 

                                
                    if consulta==stado :
                       consulta=False 
                       

                    try:
                        try:
                            logging.info("----------------Diligenciando resumen--------------------------------------- ")
                            defRPAselenium.diligenciarResumen(Hoja,Carpeta)
                            stado=False
                        except:
                            logging.error("Fallo la funcion diligenciarResumen()")

                        try:
                            logging.info("----------------Diligenciando hojas resumen por sheets de excel--------------- ")
                            defRPAselenium.diligenciarhojas(Hoja,Carpeta,region,comuna,str(rolmatriz),str(Rut),str(Inmobiliaria),str(rol1),str(rol2)) 
                        except:
                            logging.error("Fallo la funcion diligenciarhojas()")

                        logging.info("----------------Ejecutando Macros     -------------------------------------------- ")
                        try:
                            defRPAselenium.Macros(str(Hoja))
                          
                        except:
                            logging.error("Fallo la funcio Macros()")

                        logging.info("----------------Salidas de excel -------------------------------------------- ")
                       

                        
                        logging.info("----------------Diligenciando Formato de solicitud--------------------------- ")
                        try:
                            defRPAselenium.formatosolicitusd(Hoja,Carpeta)
                        except:
                            logging.error("Fallo la funcio formatosolicitusd().")

                        logging.info("----------------TEST TOTALES ------------------------------------------------ ")
                        
                        try:
                                moldesTerrenos.logscraping(Carpeta,str(rolmatriz))
                        except:
                                logging.error("Fallo la funcio logscraping()  .")
                        
                        try:tabla=moldesTerrenos.lectura(Carpeta,rolmatriz,Rut,Inmobiliaria,region,comuna)
                        except:logging.error("Fallo la funcio lectura().")
                            
                        try: moldesTerrenos.datosexceltotal(Hoja,tabla)
                        except:logging.error("Fallo la funcio datosexceltotal().")
                        try:moldesTerrenos.subtotal(Hoja)
                        except:logging.error("Fallo la funcio subtotal().")
                        try:moldesTerrenos.reporteHojas(Hoja,tabla,region)
                        except:logging.error("Fallo la funcio reporteHojas().")

                    except:
                            pass
                
def tgc():
   task()  
                 
if __name__ == "__main__":
   
   eliminarcarpetas()
   Creacionescarpetas()
   moldesTerrenos.creacionExcelResumen()
   defRPAselenium.bakup()
   tgc()
   logging.info('Ejecucion finalizada')
   tiempoFinal=time.time() 
   TiempoTotal=tiempoFinal-tiempoInicio
   print("Tiempo total de ejecucion es "+str(TiempoTotal) + " seg")
   logging.info("Tiempo total de ejecucion es "+str(TiempoTotal) + " seg")
   
 
 





