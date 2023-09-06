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
                    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
                    filename= 'log procesos' )




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
    


                    logging.info(defRPAselenium.LOGconsulta(region,comuna,rol1,rol2))
                    
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
                        logging.info("----------------Diligenciando resumen--------------------------------------- ")
                        defRPAselenium.diligenciarResumen(Hoja,Carpeta)
                        stado=False


                        logging.info("----------------Diligenciando hojas resumen por sheets de excel--------------- ")
                        defRPAselenium.diligenciarhojas(Hoja,Carpeta,region,comuna,str(rolmatriz),str(Rut),str(Inmobiliaria),str(rol1),str(rol2)) 

                        logging.info("----------------Ejecutando Macros-------------------------------------------- ")

                        
                        defRPAselenium.Macros(str(Hoja))
                        defRPAselenium.Macros(str(Hoja))
                        moldesTerrenos.logscraping(Carpeta,str(rolmatriz))
                        moldesTerrenos.salida(Carpeta,rolmatriz,Rut,Inmobiliaria,region,comuna)
                        defRPAselenium.salida(Carpeta,rolmatriz,Rut,Inmobiliaria,region,comuna)
                        
                        logging.info("----------------Diligenciando Formato de solicitud--------------------------- ")
                        defRPAselenium.formatosolicitusd(Hoja,Carpeta)
                    except:
                            pass
                
          

def tgc():
 
   task()  

                       
    
         
if __name__ == "__main__":
   
  # eliminarcarpetas()
  # Creacionescarpetas()
  #defRPAselenium.bakup()
   tgc()
   logging.info('Ejecucion finalizada')
   
 
 





