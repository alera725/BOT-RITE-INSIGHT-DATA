# -*- coding: utf-8 -*-
"""
Created on Tue Feb  2 14:45:51 2021

@author: alejandro.gutierrez
"""

#Importar paqueterias
import os
import shutil
os.chdir('C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\ALEJANDRO RAMOS GTZ\\GIT\\RITE INSIGHT BOT') # relative path: scripts dir is under Lab

import unittest
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait # Guardar en las paginas
#from selenium.webdriver.chrome.options import Options
import time 
import datetime
import pandas as pd
from datetime import date, timedelta, datetime
import json

from LOGIN_PAGE import main_login
from PROCESS_PAGE import process_page


class Download_RITE_INSIGHT_DATA(unittest.TestCase):
    
    def setUp(self):
        #option = Options()
        chrome_options = webdriver.ChromeOptions()
        prefs = {'download.default_directory' : 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\DOWNLOADS'} #'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\ALEJANDRO RAMOS GTZ\\TEST'} #CAMBIAR ESTO PARA CADA TEST
        chrome_options.add_experimental_option('prefs', prefs)
        chrome_options.add_argument("--incognito")
        #chrome_options.add_argument('--headless')
        chromedriver_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\ALEJANDRO RAMOS GTZ\\PYTHON\\chromedriver'
        url = 'https://raportal.riteaid.com/ra/raportal/ramn0001.aspx'
        self.driver = webdriver.Chrome(executable_path=chromedriver_path, chrome_options=chrome_options)
        self.driver.get(url)
        self.WebDriverWait = WebDriverWait
        #self.driver.implicity_wait(7)
        self.PageInitial = main_login(self.driver)
        self.PageProcess = process_page(self.driver)
        
        #paths to save each kind of report 
        self.dir_download = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\DOWNLOADS'
        self.inventory_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        self.periods_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\PERIODS'
        self.ytd_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\YTD'
        self.item_master_listing_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\ITEM MASTER LISTING COMPETITIVE'
        
        with open('C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\ALEJANDRO RAMOS GTZ\\PYTHON\\BOT RITE INSIGHT\\info.json') as json_file:
            data = json.load(json_file)
            self.email = data["user"]
            self.pswd = data["pass"]

    #@unittest.skip('Not need now')
    def test_INVENTORY_KIND(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'KIND RITE INSIGHT '

        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1, brand='KIND') #Agregar el parametro brand='KIND'
        self.PageProcess.check_non48_49()
        self.PageProcess.excel()
        
        time.sleep(60)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue        

        # !!! IMPORTANTE !!! ENCONTRAR EL ULTIMO ARCHIVO EN LA CARPETA DE DESCARGAS Y CAMBIARLE EL NOMBRE POR EL DESEADO
        #Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.inventory_path #'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        #COPIAR ESTE INVENTORY COMO NUEVA HOJA A SU RUTA EN EL EXCEL DE LA MARCA 
        
        print("%s is READY!!"%name) 
        time.sleep(3)
        
    @unittest.skip('Not need now')
    def test_INVENTORY_KRAVE(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'KRAVE RITE INSIGHT '

        #eL BOT 1 LLEGA HASTA PRODUCT FILER Y DESPUES EL BOTON DE EXCEL 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1, brand='KRAVE') #Agregar el parametro brand='KIND'
        self.PageProcess.check_non48_49()
        self.PageProcess.excel()
        
        time.sleep(60)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue        

        # !!! IMPORTANTE !!! ENCONTRAR EL ULTIMO ARCHIVO EN LA CARPETA DE DESCARGAS Y CAMBIARLE EL NOMBRE POR EL DESEADO
        #Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.inventory_path #'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        #COPIAR ESTE INVENTORY COMO NUEVA HOJA A SU RUTA EN EL EXCEL DE LA MARCA 
        
        print("%s is READY!!"%name) 
        time.sleep(3)
        
    @unittest.skip('Not need now')
    def test_INVENTORY_EOS(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'EOS RITE INSIGHT '

        #eL BOT 1 LLEGA HASTA PRODUCT FILER Y DESPUES EL BOTON DE EXCEL 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1, brand='EOS') #Agregar el parametro brand='KIND'
        self.PageProcess.check_non48_49()
        self.PageProcess.excel()
        
        time.sleep(60)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue        

        # !!! IMPORTANTE !!! ENCONTRAR EL ULTIMO ARCHIVO EN LA CARPETA DE DESCARGAS Y CAMBIARLE EL NOMBRE POR EL DESEADO
        #Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week '+ str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.inventory_path #'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        #COPIAR ESTE INVENTORY COMO NUEVA HOJA A SU RUTA EN EL EXCEL DE LA MARCA 
        
        print("%s is READY!!"%name) 
        time.sleep(3)
        
    @unittest.skip('Not need now')
    def test_INVENTORY_STERNO(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'STERNO RITE INSIGHT '

        #eL BOT 1 LLEGA HASTA PRODUCT FILER Y DESPUES EL BOTON DE EXCEL 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1, brand='STERNO') #Agregar el parametro brand='KIND'
        self.PageProcess.check_non48_49()
        self.PageProcess.excel()
        
        time.sleep(60)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue        

        # !!! IMPORTANTE !!! ENCONTRAR EL ULTIMO ARCHIVO EN LA CARPETA DE DESCARGAS Y CAMBIARLE EL NOMBRE POR EL DESEADO
        #Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.inventory_path #'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        #COPIAR ESTE INVENTORY COMO NUEVA HOJA A SU RUTA EN EL EXCEL DE LA MARCA 
        
        print("%s is READY!!"%name) 
        time.sleep(3)

    @unittest.skip('Not need now')
    def test_INVENTORY_GOLDEN_EYE(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'GOLDEN_EYE RITE INSIGHT '

        #eL BOT 1 LLEGA HASTA PRODUCT FILER Y DESPUES EL BOTON DE EXCEL 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1, brand='GOLDEN_EYE') #Agregar el parametro brand='KIND'
        self.PageProcess.check_non48_49()
        self.PageProcess.excel()
        
        time.sleep(60)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue        

        # !!! IMPORTANTE !!! ENCONTRAR EL ULTIMO ARCHIVO EN LA CARPETA DE DESCARGAS Y CAMBIARLE EL NOMBRE POR EL DESEADO
        #Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.inventory_path #'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        #COPIAR ESTE INVENTORY COMO NUEVA HOJA A SU RUTA EN EL EXCEL DE LA MARCA 
        
        print("%s is READY!!"%name) 
        time.sleep(3)

  
    @unittest.skip('Not need now')
    def test_INVENTORY_VIRMAX(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'VIRMAX RITE INSIGHT'

        #eL BOT 1 LLEGA HASTA PRODUCT FILER Y DESPUES EL BOTON DE EXCEL 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1, brand='VIRMAX') #Agregar el parametro brand='KIND'
        self.PageProcess.check_non48_49()
        self.PageProcess.excel()
        
        time.sleep(60)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue        

        # !!! IMPORTANTE !!! ENCONTRAR EL ULTIMO ARCHIVO EN LA CARPETA DE DESCARGAS Y CAMBIARLE EL NOMBRE POR EL DESEADO
        #Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.inventory_path #'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        #COPIAR ESTE INVENTORY COMO NUEVA HOJA A SU RUTA EN EL EXCEL DE LA MARCA 
        
        print("%s is READY!!"%name) 
        time.sleep(3)
        
    @unittest.skip('Not need now')
    def test_INVENTORY_EVOLVE(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'EVOLVE RITE INSIGHT '

        #eL BOT 1 LLEGA HASTA PRODUCT FILER Y DESPUES EL BOTON DE EXCEL 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1, brand='EVOLVE') #Agregar el parametro brand='KIND'
        self.PageProcess.check_non48_49()
        self.PageProcess.excel()
        
        time.sleep(60)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue        

        # !!! IMPORTANTE !!! ENCONTRAR EL ULTIMO ARCHIVO EN LA CARPETA DE DESCARGAS Y CAMBIARLE EL NOMBRE POR EL DESEADO
        #Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.inventory_path #'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        #COPIAR ESTE INVENTORY COMO NUEVA HOJA A SU RUTA EN EL EXCEL DE LA MARCA 
        
        print("%s is READY!!"%name) 
        time.sleep(3)
        
        
    
    @unittest.skip('Not need now')
    def test_INVENTORY_SOYLENT(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'SOYLENT RITE INSIGHT '

        #eL BOT 1 LLEGA HASTA PRODUCT FILER Y DESPUES EL BOTON DE EXCEL 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1, brand='SOYLENT') #Agregar el parametro brand='KIND'
        self.PageProcess.check_non48_49()
        self.PageProcess.excel()
        
        time.sleep(60)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue        

        # !!! IMPORTANTE !!! Aqui Checar el ultimo archivo cargado de ese retailer y vemos los renglones que tiene damos un rango de +- 100 o 50 FALTAAAAAA RUTA DEL RETAILER EN EL QUE ESTAMOS Y REVISAR ULTIMO ARCHIVO Y CONTAR LOS ROWS
        previous_file_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        try:
            previous_file = max([previous_file_path + "\\" + f for f in os.listdir(previous_file_path)],key=os.path.getctime) #Ultimo archivo cargado en la ruta del retailer 
            #print(previous_file)
            previous_file_n = pd.read_excel(previous_file)#, skiprows=14)
            n_rows_prev = len(previous_file_n.index) #Count rows 
            #print(n_rows_prev)
            
            
            #CHECAR LA ULTIMA DESCARGA (PARA CAMBIAR EL NOMBRE A LA ULTIMA DESCARGA)
            Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin") 
            Initial_path = self.dir_download 
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            
            current_file = pd.read_excel(filename) #, skiprows=14)
            n_rows_current = len(current_file.index) #Count rows 
            
            #Aqui despues se hace si el nuevo archivo pasa el limite se guarda con otro nombre advirtiendo que puede tener error el archivo 
            #Sino se guarda el archivo todo normal
            limite = n_rows_prev/3
            
            if(n_rows_current>n_rows_prev+limite or n_rows_current<n_rows_prev-limite):              
                new_name = '%s'%name + str(Current_Date) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
            else:
                new_name = '%s'%name + str(Current_Date) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
        except:
            Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
            Initial_path = self.dir_download
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            new_name = '%s'%name + str(Current_Date) + '.xlsx'
            shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\INVENTORY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("%s is READY!!"%name) 
        time.sleep(3)
        
        
    @unittest.skip('Not need now')
    def test_SBIC_YTD_FROM_DECEMBER_2020(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'SBIC YTD 27DEC2020 TO LAST SATURDAY '
                        
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(2)
        self.PageProcess.set_dates(2)
        self.PageProcess.Prod_filter(2)
        self.PageProcess.check_non48_49()
        self.PageProcess.store_filter_new()
        self.PageProcess.group_filter()
        self.PageProcess.POG_filter()
        self.PageProcess.metrics_filter()
        self.PageProcess.excel()
        
        time.sleep(200)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue
            
        # !!! IMPORTANTE !!! Aqui Checar el ultimo archivo cargado de ese retailer y vemos los renglones que tiene damos un rango de +- 100 o 50 FALTAAAAAA RUTA DEL RETAILER EN EL QUE ESTAMOS Y REVISAR ULTIMO ARCHIVO Y CONTAR LOS ROWS
        previous_file_path = self.ytd_path
        #Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin") 
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        try: 
            previous_file = max([previous_file_path + "\\" + f for f in os.listdir(previous_file_path)],key=os.path.getctime) #Ultimo archivo cargado en la ruta del retailer 
            previous_file_n = pd.read_excel(previous_file)#,# skiprows=19)
            n_rows_prev = len(previous_file_n.index) #Count rows         
            
            #CHECAR LA ULTIMA DESCARGA (PARA CAMBIAR EL NOMBRE A LA ULTIMA DESCARGA)
            Initial_path = self.dir_download 
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime) #Te da todo el path
            
            current_file = pd.read_excel(filename)#, skiprows=19)
            n_rows_current = len(current_file.index) #Count rows 
            
            #Aqui despues se hace si el nuevo archivo pasa el limite se guarda con otro nombre advirtiendo que puede tener error el archivo 
            #Sino se guarda el archivo todo normal
            limite = n_rows_prev/3
            
            if(n_rows_current>n_rows_prev+limite or n_rows_current<n_rows_prev-limite):              
                new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
            else:
                new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
        except:
            Initial_path = self.dir_download
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
            shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.ytd_path
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("%s is READY!!"%name) 
        time.sleep(3)  
        

    @unittest.skip('Not need now')   
    def test_SBIC_WEEKLY(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'SBIC 1 WEEK '
        
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(3)
        self.PageProcess.set_dates(3)
        self.PageProcess.Prod_filter(3)
        self.PageProcess.check_non48_49()
        self.PageProcess.store_filter_new()
        self.PageProcess.group_filter()
        self.PageProcess.POG_filter()
        self.PageProcess.metrics_filter()
        self.PageProcess.excel()
        
        time.sleep(200)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue
            
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")

        #Encontrar el ultimo archvio descargado y cambiarle el nombre 
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.periods_path
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("%s is READY!!"%name) 
        time.sleep(3)   
        
        
    @unittest.skip('Not need now')   
    def test_SBIC_52_WEEKS(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'SBIC 52 WEEKS '
                 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(3)
        self.PageProcess.set_dates(4,'52')
        self.PageProcess.Prod_filter(3)
        self.PageProcess.check_non48_49()
        self.PageProcess.store_filter_new()
        self.PageProcess.group_filter()
        self.PageProcess.POG_filter()
        self.PageProcess.metrics_filter()
        self.PageProcess.excel()
        
        time.sleep(320)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue
            
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")

        #Encontrar el ultimo archvio descargado y cambiarle el nombre 
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.periods_path
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("%s is READY!!"%name) 
        time.sleep(3)    
        
    @unittest.skip('Not need now')   
    def test_SBIC_26_WEEKS(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'SBIC 26 WEEKS '
                 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(3)
        self.PageProcess.set_dates(4,'26')
        self.PageProcess.Prod_filter(3)
        self.PageProcess.check_non48_49()
        self.PageProcess.store_filter_new()
        self.PageProcess.group_filter()
        self.PageProcess.POG_filter()
        self.PageProcess.metrics_filter()
        self.PageProcess.excel()
        
        time.sleep(320)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue
            
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")

        #Encontrar el ultimo archvio descargado y cambiarle el nombre 
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.periods_path
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("%s is READY!!"%name) 
        time.sleep(3)         
        
    @unittest.skip('Not need now')   
    def test_SBIC_13_WEEKS(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'SBIC 13 WEEKS '
                
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(3)
        self.PageProcess.set_dates(4,'13')
        self.PageProcess.Prod_filter(3)
        self.PageProcess.check_non48_49()
        self.PageProcess.store_filter_new()
        self.PageProcess.group_filter()
        self.PageProcess.POG_filter()
        self.PageProcess.metrics_filter()
        self.PageProcess.excel()
        
        time.sleep(320)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue
            
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")

        #Encontrar el ultimo archvio descargado y cambiarle el nombre 
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.periods_path
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("%s is READY!!"%name) 
        time.sleep(3)   
    
    @unittest.skip('Not need now')   
    def test_SBIC_4_WEEKS(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'SBIC 4 WEEKS '
                
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(3)
        self.PageProcess.set_dates(4,'4')
        self.PageProcess.Prod_filter(3)
        self.PageProcess.check_non48_49()
        self.PageProcess.store_filter_new()
        self.PageProcess.group_filter()
        self.PageProcess.POG_filter()
        self.PageProcess.metrics_filter()
        self.PageProcess.excel()
        
        time.sleep(320)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue
            
        #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")

        #Encontrar el ultimo archvio descargado y cambiarle el nombre 
        Initial_path = self.dir_download
        filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
        new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
        shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.periods_path
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("%s is READY!!"%name) 
        time.sleep(3)   
    
    @unittest.skip('Not need now')   
    def test_ITEM_MASTER_LISTING(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        name = 'ITEM MASTER LISTING 52 WEEKS '
                 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(4)
        self.PageProcess.check_non48_49()
        time.sleep(3)
        self.PageProcess.excel()
        
        time.sleep(200)
        
        #Esperar a que la descarga se complete
        after = os.listdir(self.dir_download) 
        change = set(after) - set(before)
        while len(change) != 1:
            after = os.listdir(self.dir_download) 
            change = set(after) - set(before)
            if len(change) == 1:
                file_name = change.pop()
                break
            else:
                continue
            
        # !!! IMPORTANTE !!! Aqui Checar el ultimo archivo cargado de ese retailer y vemos los renglones que tiene damos un rango de +- 100 o 50 FALTAAAAAA RUTA DEL RETAILER EN EL QUE ESTAMOS Y REVISAR ULTIMO ARCHIVO Y CONTAR LOS ROWS
        previous_file_path = self.item_master_listing_path
                #Week number 
        t1 = datetime.now()
        week_number = t1.strftime("%U")
        
        try:
            previous_file = max([previous_file_path + "\\" + f for f in os.listdir(previous_file_path)],key=os.path.getctime) #Ultimo archivo cargado en la ruta del retailer 
            previous_file_n = pd.read_excel(previous_file)#, skiprows=26)
            n_rows_prev = len(previous_file_n.index) #Count rows         
            
            #CHECAR LA ULTIMA DESCARGA (PARA CAMBIAR EL NOMBRE A LA ULTIMA DESCARGA)
            Initial_path = self.dir_download 
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            
            current_file = pd.read_excel(filename)#, skiprows=26)
            n_rows_current = len(current_file.index) #Count rows 
            
            #Aqui despues se hace si el nuevo archivo pasa el limite se guarda con otro nombre advirtiendo que puede tener error el archivo 
            #Sino se guarda el archivo todo normal
            limite = n_rows_prev/3

            if(n_rows_current>n_rows_prev+limite or n_rows_current<n_rows_prev-limite):              
                new_name = 'ERROR PLEASE CHECK IF THIS FILE IS OK ' + 'week ' + str(week_number) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
            else:
                new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
        except:
            Initial_path = self.dir_download
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            new_name = '%s'%name + 'week ' + str(week_number) + '.xlsx'
            shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = self.item_master_listing_path
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("%s is READY!!"%name) 
        time.sleep(3)   

        
    def tearDown(self):
        self.driver.close()
        self.driver.quit()
        
        
if __name__ == '__main__':
    unittest.main()
       