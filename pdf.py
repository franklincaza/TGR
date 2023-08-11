from RPA.PDF import PDF
from robot.libraries.String import String
import re

pdf = PDF()
string = String()

def extract_data_from_first_page():
    text = pdf.get_text_from_pdf("PDF\94-76182178-4-Inversiones World Logistic\Cupon de pago 4505-54 0-2019.pdf")

    x = re.search("Emision0154\d\d.\d\d\d.\d\d\d", str(text))
    y= re.search("\d\d\d.\d\d\dFECHA VIG", str(text))
    avalafecta=str(y).replace("FECHA VIG"," ")
    contribuccionMunicipal=str(x).replace("Emision0154"," ")

    f = open("demofile2.txt", "a")
    f.write(str(text))
    f.close()

    print(contribuccionMunicipal)

    return avalafecta,contribuccionMunicipal

    


extract_data_from_first_page()