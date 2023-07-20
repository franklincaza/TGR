import defRPAselenium
import models
from RPA.Browser.Selenium import Selenium;
import os
from shutil import rmtree
browser = Selenium()
Dt=models.master()


urlbase=defRPAselenium.Pyasset(asset="base")


def eliminarcarpetas():
    try:
        rmtree("PDF")
        rmtree("CSV")
        rmtree("Log Scraping")
        print("Eliminamos carpetas")
    except:
        pass


def Creacionescarpetas():
    print("Creado las carpetas para PDF's")
    try:
        os.mkdir('PDF')
        os.mkdir('CSV')
        os.mkdir("Log Scraping")
    except:
        pass

    for car in Dt:     
     try:       
      os.mkdir('PDF/'+str(car[3]))    
      print(" PDF/"+str(car[3]))     
     except:
      pass
        


    




def task():
    
        for dtable in Dt:
            if dtable[5] == "SI":
                strcomuna="{} [{}]"
                Rut=str(dtable[0])
                Inmobiliaria=dtable[1]
                Asset=dtable[2]
                Carpeta=dtable[3]
                Hoja=dtable[4]
                Activo=dtable[5]
                region=dtable[6]
                comuna=strcomuna.format(dtable[7],dtable[10])
                rol1=dtable[8]                               
                rol2=dtable[9]
                Codigo=dtable[10]
            
            
                

                defRPAselenium.LOGconsulta(region,comuna,rol1,rol2)
                try:
                    tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta)                
                except:
                    defRPAselenium.cerraNavegador()
                    print("----------Reintento de trasaccion----------------------------")
                    tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta)             
                finally:
                    if tabla == """Recatcha no me permitio hacer la consulta""":
                        defRPAselenium.cerraNavegador()
                        print("----------Reintento de trasaccion----------------------------")
                        tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta)
                        
                
                    if tabla is None:
                        print("tx fallida regitros txt del RUT: "+ str(Rut)+"-"+str(region)+"-"+str(comuna)+"-"+str(rol1)+"-"+str(rol2))
                        nom="Log Scraping/"+Carpeta+".txt"  
                        f = open(nom, "a")
                        f.write("tx fallida regitros txt")
                        f.close()
                        defRPAselenium.cerraNavegador() 
                        tabla = defRPAselenium.navegacion(region,comuna,rol1,rol2,Carpeta) 
                    else:
                        datosscrap=str(tabla)  
                        nom="Log Scraping/"+Carpeta+".txt"     
                        f = open(nom, "a")
                        f.write(datosscrap)
                        f.close()
                        #link 
                        defRPAselenium.cerraNavegador()
                     

def tgc():
 task()                   
    
         
if __name__ == "__main__":
   eliminarcarpetas()
   Creacionescarpetas()
   tgc()
   models.txttocsv()
   print('Ejecucion finalizada')
 
 





