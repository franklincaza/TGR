import defRPAselenium
import moldesTerrenos
import models
from RPA.Browser.Selenium import Selenium;
import os
from shutil import rmtree
import time
import logging

logging.basicConfig(level=logging.DEBUG,
                    format='%(asctime)s - %(threadName)s - %(LevelName)s - %(message)s',
                    filename="logProcess.txt",
                    filemode="a" )

browser = Selenium()
Dt=moldesTerrenos.masterlibros()
urlbase=str(defRPAselenium.Pyasset(asset="base"))
UrlMacro=defRPAselenium.Pyasset(asset="Ruta ")
libro=defRPAselenium.Pyasset(asset="LIBRO ")




#moldesTerrenos.task_Modelos()
#moldesTerrenos.Asignaconsultafecha()






def eliminarcarpetas():
    try:
        rmtree("PDF")
        rmtree("CSV")
        rmtree("Log Scraping")
        rmtree("Formato Solicitud")
        rmtree("Salida") 
        rmtree('Excel')     
        print("Eliminamos carpetas")

    except:
        pass

def Creacionescarpetas():
    print("Creado las carpetas para PDF's")

    try:
        os.mkdir('PDF')
        os.mkdir('CSV')
        os.mkdir('Excel')
        os.mkdir("Formato Solicitud")  
        os.mkdir("Log Scraping")
        os.mkdir("Salida") 

    except:
        pass

def task():
    
        for dtable in Dt:
         try:
            if dtable[16] == "SI":

                Rut=str(dtable[3])
                Inmobiliaria=dtable[1]
                Asset=dtable[2]
                
                Carpeta=str(dtable[9]+" "+dtable[7])
                Hoja=dtable[8]
                Activo=dtable[5]
                region=dtable[8]
                rolmatriz=dtable[7]
                rol1=dtable[5]                               
                rol2=dtable[6]
                Codigo=dtable[15]
                comuna=dtable[15]

                log=defRPAselenium.LOGconsulta(region,comuna,rol1,rol2)
                logging.info(log)
         except: 
                pass 
                 
                try:
                 tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                 
                except:
                   
                   #Tercer reintento para garantizar continuidad si encuentra Recatchat
                   try: 
                    logging.warning("segundo reintento ") 
                    time.sleep(3)
                    tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                   except:
                      logging.warning("tercer reintento ") 
                   try:
                         time.sleep(3)
                         tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                         pass
                   except:
                       pass
                       logging.error("Error en la consulta validar , si fue por tema de retcachat")
                   
                finally:
                    pass   
                    defRPAselenium.cerraNavegador() 
                    logging.info("cerrando navegador")

                try:

                    logging.info("----------------Diligenciando resumen--------------------------------------- ")
                    # defRPAselenium.diligenciarResumen(Hoja,Carpeta)
                    moldesTerrenos.logscraping(Carpeta,rolmatriz)
                    

                    logging.info("----------------Diligenciando hojas resumen por sheets de excel--------------- ")
                    #defRPAselenium.diligenciarhojas(Hoja,Carpeta,region,comuna,str(rolmatriz),str(Rut),str(Inmobiliaria),str(rol1),str(rol2)) 
                    moldesTerrenos.salida(Carpeta,rolmatriz,Rut,Inmobiliaria,region,comuna)
                    logging.info("----------------Ejecutando Macros-------------------------------------------- ")
                   # defRPAselenium.Macros(str(Hoja))
                   # defRPAselenium.Macros(str(Hoja))
                    
                    logging.info("----------------Diligenciando Formato de solicitud--------------------------- ")
                   # defRPAselenium.formatosolicitusd(Hoja,Carpeta)

                    

                except:
                        logging.error("Error en diligenciamiento de exceles de salida")
                        pass

def tgc():
 task()                   
    
         
if __name__ == "__main__":
   
   eliminarcarpetas()
   Creacionescarpetas()
   defRPAselenium.bakup()
  
   tgc()
   logging.info('Ejecucion finalizada')
   
 
 





