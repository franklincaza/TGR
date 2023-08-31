from RPA.PDF import PDF
from robot.libraries.String import String
import re

pdf = PDF()
string = String()

def extract_data_from_first_page():
    text = pdf.get_text_from_pdf("PDF\94-76182178-4-Inversiones World Logistic\Cupon de pago 4505-54 0-2017.pdf")

    print(text)

   
    


extract_data_from_first_page()