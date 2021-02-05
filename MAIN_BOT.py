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
        self.dir_download = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\DOWNLOADS'
        
        with open('C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\ALEJANDRO RAMOS GTZ\\PYTHON\\BOT RITE INSIGHT\\info.json') as json_file:
            data = json.load(json_file)
            self.email = data["user"]
            self.pswd = data["pass"]

    @unittest.skip('Not need now')
    def test_BOT_1(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
        
        #eL BOT 1 LLEGA HASTA PRODUCT FILER Y DESPUES EL BOTON DE EXCEL 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(1)
        self.PageProcess.set_dates(1)
        self.PageProcess.Prod_filter(1)
        self.PageProcess.excel()
        
        time.sleep(120)
        
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
        previous_file_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\KIND'
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
                new_name = 'ERROR PLEASE CHECK IF THIS FILE IS OK ' + str(Current_Date) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
            else:
                new_name = 'KIND RITE INSIGHT ' + str(Current_Date) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
        except:
            Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
            Initial_path = self.dir_download
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            new_name = 'KIND RITE INSIGHT ' + str(Current_Date) + '.xlsx'
            shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
        
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\KIND'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)
       
        #Listo
        print("KIND RITE INSIGHT is READY!!") 
        time.sleep(3)
        
        
    @unittest.skip('Not need now')
    def test_BOT_2(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
                        
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(2)
        self.PageProcess.set_dates(2)
        self.PageProcess.Prod_filter(2)
        self.PageProcess.store_filter()
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
        previous_file_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\FROM DECEMBER 2020'
        try: 
            previous_file = max([previous_file_path + "\\" + f for f in os.listdir(previous_file_path)],key=os.path.getctime) #Ultimo archivo cargado en la ruta del retailer 
            previous_file_n = pd.read_excel(previous_file)#,# skiprows=19)
            n_rows_prev = len(previous_file_n.index) #Count rows         
            
            #CHECAR LA ULTIMA DESCARGA (PARA CAMBIAR EL NOMBRE A LA ULTIMA DESCARGA)
            Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin") 
            Initial_path = self.dir_download 
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime) #Te da todo el path
            
            current_file = pd.read_excel(filename)#, skiprows=19)
            n_rows_current = len(current_file.index) #Count rows 
            
            #Aqui despues se hace si el nuevo archivo pasa el limite se guarda con otro nombre advirtiendo que puede tener error el archivo 
            #Sino se guarda el archivo todo normal
            limite = n_rows_prev/3
            
            if(n_rows_current > n_rows_prev + limite or n_rows_current < n_rows_prev - limite):              
                new_name = 'ERROR PLEASE CHECK IF THIS FILE IS OK ' + str(Current_Date) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
            else:
                new_name = 'ALL PRODUCTS RITE INSIGHT 27DEC2020 TO LAST SATURDAY ' + str(Current_Date) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
        except:
            Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
            Initial_path = self.dir_download
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            new_name = 'ALL PRODUCTS RITE INSIGHT 27DEC2020 TO LAST SATURDAY ' + str(Current_Date) + '.xlsx'
            shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
        
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\FROM DECEMBER 2020'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)   
        
        print("BOT 2 ALL PRODUCTS FROM 27 DEC 2020 is READY!!") 
        time.sleep(3)
        

    @unittest.skip('Not need now')   
    def test_BOT_3(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
                 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(3)
        self.PageProcess.set_dates(3)
        self.PageProcess.Prod_filter(3)
        self.PageProcess.store_filter()
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
        previous_file_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\WEEKLY'
        try:
            previous_file = max([previous_file_path + "\\" + f for f in os.listdir(previous_file_path)],key=os.path.getctime) #Ultimo archivo cargado en la ruta del retailer 
            previous_file_n = pd.read_excel(previous_file)#, skiprows=26)
            n_rows_prev = len(previous_file_n.index) #Count rows         
            
            #CHECAR LA ULTIMA DESCARGA (PARA CAMBIAR EL NOMBRE A LA ULTIMA DESCARGA)
            Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin") 
            Initial_path = self.dir_download 
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            
            current_file = pd.read_excel(filename)#, skiprows=26)
            n_rows_current = len(current_file.index) #Count rows 
            
            #Aqui despues se hace si el nuevo archivo pasa el limite se guarda con otro nombre advirtiendo que puede tener error el archivo 
            #Sino se guarda el archivo todo normal
            limite = n_rows_prev/3
            
            if(n_rows_current>n_rows_prev+limite or n_rows_current<n_rows_prev-limite):              
                new_name = 'ERROR PLEASE CHECK IF THIS FILE IS OK ' + str(Current_Date) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
            else:
                new_name = 'ALL PRODUCTS RITE INSIGHT LAST WEEK ' + str(Current_Date) + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
        except:
            Current_Date = datetime.now().strftime("%d-%b-%Y %HHr %MMin")
            Initial_path = self.dir_download
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            new_name = 'ALL PRODUCTS RITE INSIGHT LAST WEEK ' + str(Current_Date) + '.xlsx'
            shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\WEEKLY'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("BOT 3 ALL PRODUCTS WEEKLY is READY!!") 
        time.sleep(3)
        
        
        
    #@unittest.skip('Not need now')   
    def test_ITEM_MASTER_LISTING(self):
        # Obtener la descargas antes de los test
        before = os.listdir(self.dir_download) 
                 
        self.PageInitial.step_1(self.email,self.pswd)
        self.PageProcess.POS_tab()
        self.PageProcess.select_report(4)
        #self.PageProcess.Prod_filter(4)
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
        previous_file_path = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\ITEM MASTER LISTING COMPETITIVE'
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
                new_name = 'ERROR PLEASE CHECK IF THIS FILE IS OK' + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
            else:
                new_name = 'Item Master RITE INSIGHT LAST WEEK' + '.xlsx'
                shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))
        except:
            Initial_path = self.dir_download
            filename = max([Initial_path + "\\" + f for f in os.listdir(Initial_path)],key=os.path.getctime)
            new_name = 'Item Master RITE INSIGHT LAST WEEK ' + '.xlsx'
            shutil.move(filename,os.path.join(Initial_path,r'%s' %new_name))            
            
        #MOVER EL ARCHIVO A LA UBICACION DESEADA
        new_download = 'C:\\Users\\alejandro.gutierrez\\OneDrive - Carlin Group - CA Fortune\\Documents\\KROGER SELENIUM\\RITE INSIGHT\\ITEM MASTER LISTING COMPETITIVE'
        shutil.move('%s'%self.dir_download+'\\%s'%new_name, '%s'%new_download+'\\%s'%new_name)     
        
        print("BOT 4 Item Master listing is READY!!") 
        time.sleep(3)

        
    #def tearDown(self):
     #   self.driver.close()
     #   self.driver.quit()
        
        
if __name__ == '__main__':
    unittest.main()
       