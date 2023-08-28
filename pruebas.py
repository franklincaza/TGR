
from RPA.Browser.Selenium import Selenium;
from RPA.Windows import Windows
from RPA.Desktop import Desktop
import time

    
desktop = Desktop()
browser = Selenium()
library = Windows()


browser.open_user_browser("https://www.tesoreria.cl/ContribucionesPorRolWEB/muestraBusqueda?tipoPago=PortalContribPresencial")


desktop.move_mouse("point:1166,723")
desktop.click("point:1440,723")
desktop.logger()