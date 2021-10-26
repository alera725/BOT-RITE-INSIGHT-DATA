# -*- coding: utf-8 -*-
"""
Created on Tue Feb  2 14:47:21 2021

@author: alejandro.gutierrez
"""

#PAGINA PROCESO

#from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait # Guardar en las paginas
from selenium.webdriver.support import expected_conditions as EC #Guardar en las paginas import unittest
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.support.ui import Select
import datetime
import time
from datetime import date, timedelta


class process_page():
    
    def __init__(self,my_driver):
        self.driver = my_driver
        self.POS_REPORT_ID = (By.XPATH, '//*[@id="ext-comp-1006__ext-comp-1002"]')
        self.Item_Inventory_Rate_of_Sale = (By.XPATH, '//*[@id="microsite_ravend.pbrpos.iteminventoryrateofsale_gold"]/a')
        self.SALES_BY_ITEM_competitive = (By.XPATH, '//*[@id="microsite_ravend.pbrpos.salesbyitemcomp_gold"]/a')
        self.Item_masterlisting_competitive = (By.XPATH, '//*[@id="microsite_ravend.pbrpos.itmmastercomp_gold"]/a')
        self.Check_48 = (By.ID,'tprodSelselCpnsNonFund')
        self.Check_49 = (By.ID,'tprodSelselCpnsFund')
        self.dates_filter_button = (By.ID, 'tinp_we')
        self.start_date = (By.ID, 'tdateSelInpStartDt')
        self.end_date = (By.ID, 'tdateSelInpEndDt')
        self.product_filter = 'ttabPanel__ext-comp-1012'
        self.vendor_dropdown = 'tprodSelSelType'
        self.flecha_vendorstartwithK = (By.ID, 'tprodSelvendorBrowser856455061_triangle')
        
        #Clientes para el Item Master Listing
        self.kind = '//*[@id="tprodSelvendorBrowser856455061.-1624650157"]'
        self.krave = '//*[@id="tprodSelvendorBrowser856455061.840630248"]'
        self.eos = '//*[@id="tprodSelvendorBrowser-726377838.1607956470"]'
        self.sterno = '//*[@id="tprodSelvendorBrowser543223747.-1886672165"]'
        self.golden_eye = '//*[@id="tprodSelvendorBrowser985283518^children"]'
        self.virmax = '//*[@id="tprodSelvendorBrowser1342839628.-1214552059"]'
        self.evolve = '//*[@id="tprodSelvendorBrowser-726377838.748606242"]'
        self.soylent = '//*[@id="tprodSelvendorBrowser543223747.1484348191"]'
        
        self.sto_filter = (By.ID, 'ttabPanel__ext-comp-1013') 
        self.store_dropdown = (By.ID, 'tstoreSelSelType')
        self.check1_store_filter = (By.ID, 'tstoreSelCompCB')
        self.check2_store_filter = (By.ID, 'tstoreSelxonlineStore')
        self.check3_store_filter = (By.ID, 'tstoreSelxwbaStore')
        self.check4_store_filter = (By.ID, 'tstoreSelxbartellsStore')
        self.set_date_range_dropdown = (By.ID, 'tdateSellastXweeks')
        
        self.store_000 = (By.ID, 'tstrHier1255953653')
        self.store_001 = (By.ID, 'tstrHier1037788259')
        self.store_0099 = (By.ID, 'tstrHier-490580200') 
        self.store_BLANK = (By.ID, 'tstrHier-880867627') 
        self.group_store_filter = (By.ID, 'ttabPanel__ext-comp-1014')  
        self.gs_dropdown = (By.ID, 'tinp_groupby')
        self.POG = (By.ID, 'ttabPanel__ext-comp-1015')
        self.check1_pog = (By.ID, 'tinp_includepog')
        self.check2_pog = (By.ID, 'tinp_inlinepog')
        
        self.metrics = (By.ID, 'ttabPanel__ext-comp-1016') 
        self.promo_sales_checkbox = (By.ID, 'tout_promosales')
        self.promo_sales_checkbox2 = (By.ID, 'tpromosales_ty')
        self.promo_sales_checkbox3 = (By.ID, 'tpromosales_ly')
        self.promo_sales_checkbox4 = (By.ID, 'tpromosales_var')
        self.promo_sales_checkbox5 = (By.ID, 'tpromosales_pct')
        self.promo_units_checkbox = (By.ID, 'tout_promounits')
        self.promo_units_checkbox2 = (By.ID, 'tpromounits_ty')
        self.promo_units_checkbox3 = (By.ID, 'tpromounits_ly')
        self.promo_units_checkbox4 = (By.ID, 'tpromounits_var')
        self.promo_units_checkbox5 = (By.ID, 'tpromounits_pct')
        self.promo_AVGRETAILPRICE_checkbox = (By.ID, 'tout_averageretail')
        self.promo_AVGRETAILPRICE_checkbox2 = (By.ID, 'taverageretail_ty')
        self.promo_AVGRETAILPRICE_checkbox3 = (By.ID, 'taverageretail_ly')
        self.promo_AVGRETAILPRICE_checkbox4 = (By.ID, 'taverageretail_var')
        self.promo_AVGRETAILPRICE_checkbox5 = (By.ID, 'taverageretail_pct') 
        
        self.promo_allty = (By.ID, 'tallty')
        self.promo_allly = (By.ID, 'tallly')
        self.promo_allvar = (By.ID, 'tallvar')
        self.promo_allpctvar = (By.ID, 'tallpct')
        
        self.excel_button = 'tbtnExcel'
        self.excel_buttonIML = 'ext-gen47'
        
        
    def POS_tab(self):
        try:
            
            #Switch to the last opened tab
            self.driver.switch_to.window(self.driver.window_handles[-1])
            
            #PAGINA PRINCIPAL SUPER VALUE RITE
            POS_button = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.POS_REPORT_ID))
            POS_button.click()
            
        except TimeoutException:
            print ("Loading -POS_BUTTON- took too much time!")



    def select_report(self, numbot):
        try:
            
            if numbot == 1:
                #Select report  
                Report = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.Item_Inventory_Rate_of_Sale))
                Report.click()
            elif numbot in (2, 3):
                #Select report  
                Report = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.SALES_BY_ITEM_competitive))
                Report.click()
            elif numbot == 4:
                #Select report  
                Report = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.Item_masterlisting_competitive))
                Report.click()                
            else:
                print('Error, please select 1 or 2 for num bot')
            
        except TimeoutException:
            print ("Loading -select_report- took too much time!")



    def check_non48_49(self):
        try:
            
            #Select report  
            element=WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.Check_48))
            if element.is_selected():
                print('Check box 48 is already selected')
                #pass
            else:
                element.click()
                
            element_49=WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.Check_49))
            if element_49.is_selected():
                print('Check box 49 is already selected')
                #pass
            else:
                element_49.click()

        except TimeoutException:
            print ("Loading -Checks 48 and 49- took too much time!")


#!!!ESTE MODIFICAR SI QUEREMOS CORRER PARA LA SEMANA PASADA!!!
    def set_dates(self, numbot, *args):
        try:
            if numbot == 1:
                
                #Select dates OF INVENTORY RATE OF SALES REPORT (LAST WEEK)
                time.sleep(15)
                select_first = Select(self.driver.find_element(*self.dates_filter_button))
                #Aqui seleccionamos la ultima semana, si queremos modificar a la pasada poner un 1
                select_first.select_by_index(0)
                #select_first.select_by_index(1)
                
            elif numbot == 2:

                #--Set Dates YTD
                
                #Start date
                start_date = '12/27/2020'  # Ultimo Domingo del ano anterior 
                Start_date = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.start_date))
                Start_date.clear()
                Start_date.send_keys(start_date)
                
                #End date
                #Encontrar el ultimo sabado #Si queremos modificar a la semana pasada ponerlo manual como el comentario de abajo sat_format
                today = date.today()
                idx = (today.weekday() + 1) % 7
                sat = today - datetime.timedelta(7+idx-6)  
                sat_format = sat.strftime("%m/%d/%Y")
                #sat_format = '09/25/2021' #ULTIMO SABADO DE LA SEMANA QUE QUEREMOS BAJAR
                end_date = sat_format # Ultimo sabado disponible
                End_date = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.end_date))
                End_date.clear()
                End_date.send_keys(end_date)


            elif numbot == 3: 
                
                #--Set Dates WEEKLY 

                #End date #Si se quieren modificar fechas ponerlo manual sat_format y sun_format
                #Encontrar el ultimo sabado
                today = date.today()
                idx = (today.weekday() + 1) % 7
                sat = today - datetime.timedelta(7+idx-6)  
                sat_format = sat.strftime("%m/%d/%Y")
                #sat_format = '09/25/2021' #Modificar en caso de cambiar fechas
                end_date = sat_format # Ultimo sabado disponible
                End_date = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.end_date))
                End_date.clear()
                End_date.send_keys(end_date)
                
                #Start date
                #Encontrar el ultimo domingo despues del ultimo sabado
                sun = sat - datetime.timedelta(days=6)
                sun_format = sun.strftime("%m/%d/%Y") #Previous Sunday before that Saturday 
                #sun_format = '09/19/2021' #Modificar en caso de cambiar fechas
                start_date = sun_format # 12/27/2020 
                Start_date = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.start_date))
                Start_date.clear()
                Start_date.send_keys(start_date)
                time.sleep(3)
                
            elif numbot == 4:
                
                #--Set Dates for period Reports (4, 13, 26, 52), si este reporte se quiere para la semana pasada se debe sacar de manera manual
                period = args[0]
                                
                #Encontrar el dropdown
                dropdown = WebDriverWait(self.driver,40).until(EC.visibility_of_element_located(self.set_date_range_dropdown))
                dropdown.click()
                
                time.sleep(2)

                select_date_period = Select(self.driver.find_element(*self.set_date_range_dropdown)) 
                select_date_period.select_by_value(period)
                #select_date_period.click()
                
                time.sleep(2)
                
                self.driver.find_element_by_id('tdateSelInpStartDt').click()
                
                time.sleep(5)

                dropdown.click()
                time.sleep(2)
                dropdown.click()
                
                time.sleep(10)


            else:
                print('Error, please select 1,2 or 3 for num bot YOU ARE IN SET_DATES')

        
        except TimeoutException:
            print ("Loading -SET_DATES- took too much time!")



    def Prod_filter(self, numbot, **kwargs): #**kwargs
        try:
            if numbot == 1: #Cambiar por cada brand dentro de este if habra mas ifs
                #Vendor Button
                #WebDriverWait(self.driver,30).until(EC.visibility_of_element_located((By.ID, self.product_filter)))
                prod_filter = self.driver.find_element_by_id(self.product_filter)
                prod_filter.click()
                
                #Select vendor  
                #WebDriverWait(self.driver, 50).until(EC.visibility_of_element_located(self.vendor_dropdown))
                VENDOR = Select(self.driver.find_element_by_id(self.vendor_dropdown))
                VENDOR.select_by_visible_text("Vendor")
                
                #Select
                vendor_start_with = {
                    'Vendors_starting_with_B' : 'tprodSelvendorBrowser1255198513_triangle',
                    'Vendors_starting_with_E' : 'tprodSelvendorBrowser-726377838_triangle',
                    'Vendors_starting_with_G' : 'tprodSelvendorBrowser985283518_triangle',
                    'Vendors_starting_with_H' : 'tprodSelvendorBrowser-1442503121_triangle',
                    'Vendors_starting_with_K' : 'tprodSelvendorBrowser856455061_triangle',
                    'Vendors_starting_with_P' : 'tprodSelvendorBrowser-1184252295_triangle',
                    'Vendors_starting_with_S' : 'tprodSelvendorBrowser543223747_triangle',
                    'Vendors_starting_with_T' : 'tprodSelvendorBrowser-1107002784_triangle',
                    'Vendors_starting_with_V' : 'tprodSelvendorBrowser1342839628_triangle'
                }
                                
               # vendor_list = {
               #     'kind' : '//*[@id="tprodSelvendorBrowser856455061.-1624650157"]',
               #     'krave' : '//*[@id="tprodSelvendorBrowser856455061.840630248"]',
               #     'eos' : '//*[@id="tprodSelvendorBrowser-726377838.1607956470"]',
               #     'sterno' : '//*[@id="tprodSelvendorBrowser543223747.-1886672165"]',
               #     'golden_eye' : '//*[@id="tprodSelvendorBrowser985283518^children"]',
               #     'virmax' : '//*[@id="tprodSelvendorBrowser1342839628.-1214552059"]',
               #     'evolve' : '//*[@id="tprodSelvendorBrowser-726377838.748606242"]',
               #     'soylent' : '//*[@id="tprodSelvendorBrowser543223747.1484348191"]'
               # }
                
                brand_name = kwargs.get('brand', None)
                
                brand_initial_word = brand_name[0]
                
                brand_name_lower = brand_name.lower()
                
                start_with = 'Vendors_starting_with_' + brand_initial_word
                
                #Seleccionar Vendors que inician con "K" A PARTIR DE AQUI INCIIAN LOS IFS Y EL PROCESO DE CADA BRAND ejemplo:  Vendors starting with E
                flecha1 = WebDriverWait(self.driver,10).until(EC.visibility_of_element_located((By.ID, vendor_start_with[start_with])))
                flecha1.click()
                
                #Select VENDOR KIND  
                #WebDriverWait(self.driver,10).until(EC.visibility_of_element_located((By.XPATH, vendor_list[brand_name.lower()])))
                exec('WebDriverWait(self.driver,10).until(EC.visibility_of_element_located((By.XPATH, self.%s)))'%(brand_name_lower)) #use %s for strings and %d for numbers 
                #vendor = self.driver.find_element_by_xpath(vendor_list[brand_name_lower])
                exec('vendor = self.driver.find_element_by_xpath(self.%s).click()'%(brand_name_lower))
                #vendor.click()
                    
                
            elif numbot in (2, 3):
                #PRODUCT FILTER 
                pf = self.driver.find_element_by_id(self.product_filter) 
                pf.click()
            
                #set products dropdown filter
                #WebDriverWait(self.driver, 50).until(EC.visibility_of_element_located(self.vendor_dropdown))
                select_vdr = Select(self.driver.find_element_by_id(self.vendor_dropdown))
                select_vdr.select_by_visible_text("All products in Carlin Group categories") #select_by_index(1)          

            elif numbot == 4:
                #PRODUCT FILTER 
                product_filter = 'ttabPanel__ext-comp-1011'
                pf = self.driver.find_element_by_id(product_filter) 
                pf.click()
                time.sleep(3)
                #set products dropdown filter
                #WebDriverWait(self.driver, 50).until(EC.visibility_of_element_located(self.vendor_dropdown))
                select_vdr = Select(self.driver.find_element_by_id(self.vendor_dropdown))
                select_vdr.select_by_visible_text("All products in Carlin Group categories") #select_by_index(1)          

            else:
                print('Error, please select 1,2 or 3 for num bot YOU ARE IN SET_DATES')
            
        except TimeoutException:
            print ("Loading -PRODUCT_FILTER- took too much time!")



    def store_filter(self):
        try:
            #Vendor Button
            stor_filter = WebDriverWait(self.driver,30).until(EC.visibility_of_element_located(self.sto_filter))
            stor_filter.click()
            
            #Set store dropdown filter  
            WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located(self.store_dropdown))
            select_store = Select(self.driver.find_element(*self.store_dropdown))
            select_store.select_by_visible_text("Store Hierarchy")
            
            #Select Stores
            store_000 = WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.store_000))
            store_000.click()
            store_001 = WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.store_001))
            store_001.click() 
            store_099 = WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.store_0099))
            store_099.click()
            store_blank = WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.store_BLANK))
            store_blank.click()
          
        except TimeoutException:
            print ("Loading -STORE_FILTER- took too much time!")
            
            
    def store_filter_new(self): #MODIFICAR 
        try:
            #Vendor Button
            stor_filter = WebDriverWait(self.driver,30).until(EC.visibility_of_element_located(self.sto_filter))
            stor_filter.click()
            
            #Set store dropdown filter  
            WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located(self.store_dropdown))
            select_store = Select(self.driver.find_element(*self.store_dropdown))
            select_store.select_by_visible_text("All Stores")
            
            #Select Stores checkboxes 
            element=WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.check1_store_filter))
            if element.is_selected():
                #print('Check box 48 is already selected')
                #pass
                element.click()
            else:
                pass
            
            element1=WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.check2_store_filter))
            if element1.is_selected():
                #print('Check box 48 is already selected')
                pass
            else:
                element1.click()

            element2=WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.check3_store_filter))
            if element2.is_selected():
                #print('Check box 48 is already selected')
                pass
            else:
                element2.click()

            element3=WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.check4_store_filter))
            if element3.is_selected():
                #print('Check box 48 is already selected')
                pass
            else:
                element3.click()                
                
        except TimeoutException:
            print ("Loading -STORE_FILTER- took too much time!")



    def group_filter(self):
        try:
            #GROUP STORE FILTER button
            group_filter = WebDriverWait(self.driver,30).until(EC.visibility_of_element_located(self.group_store_filter))
            group_filter.click()
            
            time.sleep(2)
            #set group store dropdown filter
            #WebDriverWait(self.driver, 15).until(EC.visibility_of_element_located(self.gs_dropdown))
            self.driver.find_element(*self.gs_dropdown).click()
            select_vdr = Select(self.driver.find_element(*self.gs_dropdown))
            select_vdr.select_by_visible_text("UPC")
            
        except TimeoutException:
            print ("Loading -GROUP_FILTER- took too much time!")



    def POG_filter (self):
        try:
            #POG FILTER button
            pog_button = WebDriverWait(self.driver,30).until(EC.visibility_of_element_located(self.POG))
            pog_button.click()
            
            #Select POG  checkboxes 
            time.sleep(2)
            element=WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.check1_pog))
            if element.is_selected():
                #print('Check box 48 is already selected')
                pass
                #element.click()
            else:
                element.click()

            element2=WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.check2_pog))
            if element2.is_selected():
                #print('Check box 48 is already selected')
                element2.click()
            else:
                pass
            
            time.sleep(2)
            
        except TimeoutException:
            print ("Loading -POG_FILTER- took too much time!")



    def metrics_filter(self):
        try:
            #METRICS
            metrics = WebDriverWait(self.driver,30).until(EC.visibility_of_element_located(self.metrics))
            metrics.click()
            
            #check metrics 
            #Sales
            if self.driver.find_element(*self.promo_sales_checkbox).is_selected()==False:
                self.driver.find_element(*self.promo_sales_checkbox).click()
            else:
                pass
            
            WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.promo_sales_checkbox2))
                        
            #if self.driver.find_element(*self.promo_sales_checkbox2).is_selected()==False:
            #    self.driver.find_element(*self.promo_sales_checkbox2).click()
            #else:
            #    pass
            
            #Units
            if self.driver.find_element(*self.promo_units_checkbox).is_selected()==False:
                self.driver.find_element(*self.promo_units_checkbox).click()
            else:
                pass
            
            WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.promo_units_checkbox2))
            
            #AVG PRICE
            if self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox).is_selected()==False:
                self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox).click()
            else:
                pass
            
            WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.promo_AVGRETAILPRICE_checkbox2))
        
        
            #Click en seleccionar todo de las 4 columnas 
            self.driver.find_element(*self.promo_allty).click()
            self.driver.find_element(*self.promo_allly).click()
            self.driver.find_element(*self.promo_allvar).click()
            self.driver.find_element(*self.promo_allpctvar).click()


        except TimeoutException:
            print ("Loading -metrics_filter- took too much time!")
            
            
            
    def excel(self):
        try:
            #Click excel button
            time.sleep(4)
            button = self.driver.find_element_by_id(self.excel_button)
            button.click()
            
        except TimeoutException:
            print ("Loading -Excel BUTTON- took too much time!")
            
            
    def excel_iml(self):
        try:
            #Click excel button
            time.sleep(4)
            button = self.driver.find_element_by_id(self.excel_buttonIML)
            button.click()
            
        except TimeoutException:
            print ("Loading -Excel BUTTON IML- took too much time!")           

