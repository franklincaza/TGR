from RPA.Browser.Selenium import Selenium;
from RPA.Windows import Windows
from RPA.HTTP import HTTP
from RPA.Excel.Files import Files;
import time
import random
import csv
import os
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



    

def recorrerFilasDescargas(carpeta):
    
    rango=10
    for celda in range(rango) :
        consecutivo=celda
        
        try:
                obtenerTexto("//TABLE[@id='example']//tr["+str(celda)+"]//td[3]")
                FOLIO=obtenerTexto("//TABLE[@id='example']//tr["+str(celda)+"]//td[3]")
                clickweb("//TABLE[@id='example']//tr["+str(celda)+"]//td[3]")
                savepdf(carpeta,consecutivo)

        except:
            pass

            



def validacion():
    validacion= browser.get_text("//DIV[@class='dentro_letra'][text()='Contribuciones']")
    if validacion == 'Contribuciones': print("ingresando a "+validacion) 
    return validacion
    
    
def navegacion(region,comuna,rol1,rol2,ruta):

    openweb("https://www.tesoreria.cl/ContribucionesPorRolWEB/muestraBusqueda?tipoPago=PortalContribPresencial")                             
    clickweb("//SELECT[@id='region']/self::SELECT")
    clickweb("//option[text()='"+region+"']")
    clickweb("//SELECT[@id='comunas']")
    clickweb("//option[text()='"+comuna+"']")
    typeinputText("//INPUT[@id='rol']",rol1)
    typeinputText("//INPUT[@id='subRol']",rol2)
    clickweb("//INPUT[@id='btnRecaptchaV3Envio']/self::INPUT")
    tiempoespera()
    try:
        destacar("//TABLE[@id='example']//tbody//tr//td") 
        tabla =extraertablita()
        recorrerFilasDescargas(ruta)
        
        

    except:
        tabla ="""Recatcha no me permitio hacer la consulta"""
        cerraNavegador()
        if """Recatcha no me permitio hacer la consulta"""==tabla:
            print("Reintamos hacer la consulta")
            openweb("https://www.tesoreria.cl/ContribucionesPorRolWEB/muestraBusqueda?tipoPago=PortalContribPresencial")                             
            clickweb("//SELECT[@id='region']/self::SELECT")
            clickweb("//option[text()='"+region+"']")
            clickweb("//SELECT[@id='comunas']")
            clickweb("//option[text()='"+comuna+"']")
            typeinputText("//INPUT[@id='rol']",rol1)
            typeinputText("//INPUT[@id='subRol']",rol2)
            clickweb("//INPUT[@id='btnRecaptchaV3Envio']/self::INPUT")
            tiempoespera()
            try:
                destacar("//TABLE[@id='example']//tbody//tr//td") 
                tabla =extraertablita() 
                
               
                        
            except:
                 tabla ="""Recatcha no me permitio hacer la consulta"""
                
    
    return (tabla)

    



def savepdf(carpeta,consecutivo):
 base=Pyasset(asset="base")
 txt=base+carpeta
 salida=carpeta+"-"+str(consecutivo)
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
    time.sleep(2)
    if str(consecutivo)==str("1"):
        library.send_keys(keys=txt)
        time.sleep(5)
        library.send_keys(keys="{Enter}")
        time.sleep(2)
    elif str(consecutivo)==str(""):
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
   
     
   
    



    
        
    







