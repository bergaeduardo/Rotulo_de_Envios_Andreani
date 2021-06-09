# -*- coding: utf-8 -*-
"""
Created on Sat May  9 19:05:32 2020

@author: Black
"""

from selenium import webdriver
#from selenium.webdriver.support.ui import Select
#import pandas as pd
#from datetime import date
#from datetime import datetime
import time
# import Input
# import os
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC
# from bs4 import BeautifulSoup4

# Abrimos Chome y le indicamoa a que pagina dirigirse
# browser = webdriver.Chrome(executable_path= r'H:\Programas\Anaconda\envs\Pruebas Python\Library\bin\chromedriver.exe')
options = webdriver.ChromeOptions()
# options.add_experimental_option('excludeSwitches', ['enable-logging'])
options.add_argument('--log-level=3')
browser = webdriver.Chrome(executable_path= r'C:\Users\daniel.amaya\Anaconda3\chromedriver-Windows', options=options)
browser.get('https://andreanionline.com/login')
time.sleep(5)
#search_button = browser.find_element_by_xpath ('/html/body/div[2]/div/div[2]/form/div[3]/div[3]/a')
#search_button.click()
#time.sleep(5)
#search_textinput = browser.find_element_by_id ('username')
#search_textinput.send_keys('xlshop@xl.com.ar')

#search_textinput = browser.find_element_by_id ('password')
#search_textinput.send_keys('123456')
#time.sleep(1)
#search_button = browser.find_element_by_xpath ('/html/body/div[2]/div/div[2]/form/div[4]/div[3]/button')
#search_button.click()
#time.sleep(10)
#################################################
#################################################
#################################################
#browser.get('https://andreanionline.com/checkout/new')
#time.sleep(10)
#search_button = browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/ul/li[2]/a')
#search_button.click()
#time.sleep(5)
##seleccionar lista desplegable de que enviar
#search_button = browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[1]/div[1]/div/a/span').click()
#search_button = browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[1]/div[1]/div/div/ul/li[2]').click()
##seleccionar lista desplegable de tipo de envio
#search_button = browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[2]/div[1]/div/a/span').click()
#search_button = browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[1]/form/div[1]/div[2]/div[1]/div/div/ul/li[3]').click()
##seleccionar siguiente
#search_button=browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[1]/div[2]/div[2]/button')
#search_button.click()
## Cargar archivo
#elem = browser.find_element_by_xpath("//input[@type='file']")
#elem.send_keys(r'X:\Logistica\1. Logistica Interno\Historiales_de_Logistica\ImportacionMasivaEncomiendaDomicilio-copia.xlsx')
##seleccionar siguiente
#search_button=browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[2]/div[2]/div[2]/button[1]')
#search_button.click()
## Selecciona sucursal de origen
#search_button = browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/div[2]/div[1]/form/div/div/div/a/span').click()
#search_button = browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/div[2]/div[1]/form/div/div/div/div/ul/li[2]').click()
## Presiona finalizar
#search_button=browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[1]/div/div[2]/div[3]/div[2]/div[2]/button[1]')
#search_button.click()
## selecciona el check de responsabilidad
#search_button=browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[2]/div[2]/div/div[6]/div[1]/div/input')
#search_button.click()
## selecciona el boton de confirmar
#search_button=browser.find_element_by_xpath ('/html/body/div[1]/div[1]/div[2]/div[2]/div/div[6]/div[2]/a')
#search_button.click()

#elementos = browser.find_elements_by_xpath("//div[@class='message']/h4")
#lista=[]
#for p in elementos:
#    valor = p.text
#    lista.append(valor)

#separador=" "
#cadena=lista[0]
#separado = cadena.split(separador)
## La variable SEPARADO[] contiene el numero de operacion
#print(separado[3])

