from RPA.Excel.Files import Files;
import os
from shutil import rmtree
import csv
import openpyxl

paths="Data\Base Contribuciones Terrenos.xlsx"
sheet="Base Terreno"
sheet2="Base Terreno2"
lib = Files()

listMster= [{'RUT':'',
                   'Inmobiliaria':'',
                   'Asset':'',
                   'Carpeta':'',
                   'Hoja':'',
                   'Activo':'',
                   'Region':'',
                   'Comuna':'',
                   'RolMatriz':'',
                   'Rol':'',
                   'Codigo':'',
                   'status':''

             }]


def Dt_BaseTerreno():
    lib.open_workbook(paths)      #ubicacion del libro
    lib.read_worksheet(sheet)       #nombre de la hoja
    lista=lib.read_worksheet_as_table(name='Base Terreno',header=True, start=1).data
    return lista

    
def LogSCraping():
    
    contenido = os.listdir('Log Scraping/')
    for dataTxt in contenido:
        mensaje = open('Log Scraping/'+dataTxt, "r")
        outmensaje=mensaje.read()
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
        return(contenido)
    
        



def master():
   lib.open_workbook("Data\Master.xlsx")      #ubicacion del libro
   lib.read_worksheet("Listado")       #nombre de la hoja
   DtMaster=lib.read_worksheet_as_table(name='Listado',header=True, start=1).data

   return DtMaster


def txttocsv():

   contenido = os.listdir('Log Scraping/')
   for dataTxt in contenido:
        txt=dataTxt.replace(".txt",".csv")

        txt_file =r"Log Scraping/" + dataTxt
        csv_file =r"CSV/" + txt 

        in_txt = csv.reader(open(txt_file, "r"), delimiter = " ")
        out_csv = csv.writer(open(csv_file, 'w'))

        out_csv.writerows(in_txt)

        del out_csv

        
  
txttocsv()







