
from RPA.Windows import Windows
import time

library = Windows() 

def savepdf(carpeta,consecutivo):
 base="C:\\Users\\FRANKLIN\\Downloads\\Desarrollos\\Tareas automaticas\\tgr\\PDF\\"
 txt=base+carpeta
 salida=carpeta+"-"+str(consecutivo)

 try:
        file = open(txt+"\\"+salida)
        print(file) # File handler
        file.close()
 except:
    
    library.click("name:imprimirAr")
    time.sleep(4.5)    
    library.send_keys(keys="{CTRL}S")    
    time.sleep(2)
    library.send_keys(keys=txt)
    time.sleep(5)
    library.send_keys(keys="{Enter}")
    time.sleep(2)
    library.send_keys(keys=salida)
    time.sleep(3)
    library.send_keys(keys="{Enter}")
    print("PDF gurdado con exito" + salida)
    time.sleep(1)
    library.send_keys(keys="{Ctrl}W")
   




car="94-76182178-4-Inversiones World Logistic"
con=1

savepdf(car,con)