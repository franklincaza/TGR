import defRPAselenium
import models
from RPA.Browser.Selenium import Selenium;
import os
from shutil import rmtree
import time

browser = Selenium()
Dt=models.master()


urlbase=defRPAselenium.Pyasset(asset="base")
UrlMacro=defRPAselenium.Pyasset(asset="Ruta ")
libro=defRPAselenium.Pyasset(asset="LIBRO ")

def eliminarcarpetas():
    try:
        rmtree("PDF")
        rmtree("CSV")
        rmtree("Log Scraping")
        rmtree("Formato Solicitud")
        rmtree("Salida")      
        print("Eliminamos carpetas")

    except:
        pass


def Creacionescarpetas():
    print("Creado las carpetas para PDF's")

    try:
        os.mkdir('PDF')
        os.mkdir('CSV')
        os.mkdir("Formato Solicitud")  
        os.mkdir("Log Scraping")
        os.mkdir("Salida") 

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
 


                defRPAselenium.LOGconsulta(region,comuna,rol1,rol2)
                try:
                 tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                 
                except:
                   
                   #Tercer reintento para garantizar continuidad si encuentra Recatchat
                   try: 
                    print("segundo reintento ") 
                    time.sleep(3)
                    tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                   except:
                    
                    print("tercer reintento ") 
                    try:
                        time.sleep(3)
                        tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta,Hoja)
                        pass
                    except:
                       pass
                   
                finally:
                    pass   
                    defRPAselenium.cerraNavegador() 

                try:

                    print("----------------Diligenciando resumen--------------------------------------- ")
                    defRPAselenium.diligenciarResumen(Hoja,Carpeta)

                    print("----------------Diligenciando hojas resumen por sheets de excel--------------- ")
                    defRPAselenium.diligenciarhojas(Hoja,Carpeta,region,comuna,str(rolmatriz),str(Rut),str(Inmobiliaria),str(rol1),str(rol2)) 

                    print("----------------Ejecutando Macros-------------------------------------------- ")
                    defRPAselenium.Macros(str(Hoja))
                    defRPAselenium.Macros(str(Hoja))
                    
                    print("----------------Diligenciando Formato de solicitud--------------------------- ")
                    defRPAselenium.formatosolicitusd(Hoja,Carpeta)

                    

                except:
                        pass
                


    
           

def tgc():
 task()                   
    
         
if __name__ == "__main__":
   eliminarcarpetas()
   Creacionescarpetas()
   defRPAselenium.bakup()
   tgc()
   models.txttocsv()
   defRPAselenium.salida()
   print('Ejecucion finalizada')
   
 
 





