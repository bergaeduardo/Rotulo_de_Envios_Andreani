# -*- coding: utf-8 -*-
"""
Created on Wed May 13 12:04:11 2020

@author: Black
"""

from selenium import webdriver
#from selenium.webdriver.support.ui import Select
#import pandas as pd
#from datetime import date
#from datetime import datetime
import time
# import Input
#import os
#from selenium.webdriver.common.by import By
#from selenium.webdriver.support.ui import WebDriverWait
#from selenium.webdriver.support import expected_conditions as EC
# from bs4 import BeautifulSoup4
import autoit
# Instalar autoit >>  pip install -U pyautoit

# Abrimos Chome o Firefox y le indicamoa a que pagina dirigirse
# browser = webdriver.Firefox(executable_path= r'C:\Users\eduardo.berga\AppData\Local\Continuum\anaconda3\envs\Entorno de Pruebas\Scripts\geckodriver.exe')
# browser = webdriver.Chrome(executable_path= r'H:\Programas\Anaconda\envs\Pruebas Python\Library\bin\chromedriver.exe')
browser = webdriver.Chrome(executable_path= r'C:\Users\eduardo.berga\Documents\chromedriver.exe')
browser.get('https://andreanionline.com/login')
ventana_actual= browser.window_handles[0]
browser.maximize_window()
time.sleep(5)
search_button = browser.find_element_by_xpath ('/html/body/div[2]/div/div[2]/form/div[3]/div[3]/a')
search_button.click()
time.sleep(5)
search_textinput = browser.find_element_by_id ('username')
search_textinput.send_keys('xlshop@xl.com.ar')

search_textinput = browser.find_element_by_id ('password')
search_textinput.send_keys('123456')
time.sleep(1)
search_button = browser.find_element_by_xpath ('/html/body/div[2]/div/div[2]/form/div[4]/div[3]/button')
search_button.click()
time.sleep(10)

# Segundo bloque de automatizacion
# que se utilizara para la impresion de etiquetas
browser.get('https://andreanionline.com/envio/')
opnum= '1379908'
search_textinput = browser.find_element_by_id ('filtroGeneral')
search_textinput.send_keys(opnum)
time.sleep(1)
# Presiona el boton Buscar
search_textinput = browser.find_element_by_xpath('/html/body/div[1]/section/div/div/div/form/div[1]/div[1]/button[1]')
search_textinput.click()
# Mostrar 100 registros
search_button = browser.find_element_by_xpath ('/html/body/div[1]/section/div/div/div/form/div[4]/div[1]/select').click()
search_button = browser.find_element_by_xpath ('/html/body/div[1]/section/div/div/div/form/div[4]/div[1]/select/option[4]').click()
time.sleep(5)



# Presiona el check seleccionar todos
time.sleep(10)
search_textinput = browser.find_element_by_xpath('/html/body/div[1]/section/div/div/div/form/table/thead/tr/th[1]/input')
search_textinput.click()
# Presiona el boton Imprimir
time.sleep(1)
search_textinput = browser.find_element_by_xpath('/html/body/div[1]/section/div/div/div/form/div[2]/div[1]/a[1]')
search_textinput.click()
time.sleep(10)
# Guardar el valor de la nueva ventana
ventana_nueva=browser.window_handles[1]
print(ventana_actual)
print(ventana_nueva)
time.sleep(5)
browser.switch_to_window(ventana_nueva)
browser.maximize_window()
time.sleep(5)
autoit.send("^p")
time.sleep(10)
autoit.send("{ENTER}")
time.sleep(20)


print('>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<')
print('>>>>  El proceso finalizo  <<<<')
print('>>>>>>>>>>>>>>>><<<<<<<<<<<<<<<')
browser.quit()
# ventana_nueva3=browser.window_handles[2]
# print(ventana_actual)
# print(ventana_nueva)
# print(ventana_nueva3)
# browser.switch_to_window(ventana_nueva3)
# search_textinput = browser.find_element_by_xpath('/html/body/print-preview-app//print-preview-sidebar//print-preview-button-strip//cr-button[1]')
# search_textinput.click()

# # search_textinput = browser.find_element_by_xpath('/html/body/div[2]/embed')
# # search_textinput.click()



