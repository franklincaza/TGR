
from RPA.Browser.Selenium import Selenium;
from RPA.Windows import Windows
from RPA.Desktop import Desktop
import os
import time
import subprocess

    
desktop = Desktop()
browser = Selenium()
library = Windows()


browser.open_available_browser("https://www.tesoreria.cl/ContribucionesPorRolWEB/muestraBusqueda?tipoPago=PortalContribPresencial",alias="a1")
try: 
    os.system("main.exe") 
    browser.open_available_browser("https://addons.mozilla.org/firefox/downloads/file/4044701/buster_captcha_solver-2.0.1.xpi",alias="a2")
except: 
    pass
browser.switch_browser(index_or_alias="a1")
 


