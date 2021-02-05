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
        self.dates_filter_button = (By.ID, 'tinp_we')
        self.start_date = (By.ID, 'tdateSelInpStartDt')
        self.end_date = (By.ID, 'tdateSelInpEndDt')
        self.product_filter = 'ttabPanel__ext-comp-1012'
        self.vendor_dropdown = 'tprodSelSelType'
        self.flecha_vendorstartwithK = (By.ID, 'tprodSelvendorBrowser856455061_triangle')
        self.kind = '//*[@id="tprodSelvendorBrowser856455061.-1624650157"]'
                                      
        self.sto_filter = (By.ID, 'ttabPanel__ext-comp-1013') 
        self.store_dropdown = (By.ID, 'tstoreSelSelType')
        self.store_000 = (By.ID, 'tstrHier1255953653')
        self.store_006 = (By.ID, 'tstrHier-1547701824')
        self.store_0099 = (By.ID, 'tstrHier-490580200') 
        self.store_BLANK = (By.ID, 'tstrHier-880867627') 
        self.group_store_filter = (By.ID, 'ttabPanel__ext-comp-1014')  
        self.gs_dropdown = (By.ID, 'tinp_groupby')
        self.POG = (By.ID, 'ttabPanel__ext-comp-1015')
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
        self.excel_button = 'tbtnExcel'
        

        
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



    def set_dates(self, numbot):
        try:
            if numbot == 1:
                #Esperar a que se cargue bien la pagina y aparezca el boton de spreadsheet
                #WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.excel_button))
                
                #Dates Button
                #WebDriverWait(self.driver,30).until(EC.visibility_of_element_located(self.dates_filter_button))
                #dates_button.click()
                
                #Select dates  
                time.sleep(15)
                select_first = Select(self.driver.find_element(*self.dates_filter_button)) 
                select_first.select_by_index(0)

            elif numbot == 2:

                #--Set Dates
                
                #Start date
                start_date = '12/27/2020'  
                Start_date = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.start_date))
                Start_date.clear()
                Start_date.send_keys(start_date)
                
                #End date
                #Encontrar el ultimo sabado
                today = date.today()
                idx = (today.weekday() + 1) % 7
                sat = today - datetime.timedelta(7+idx-6)  
                sat_format = sat.strftime("%m/%d/%Y")
                end_date = sat_format # Ultimo sabado disponible
                End_date = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.end_date))
                End_date.clear()
                End_date.send_keys(end_date)


            elif numbot == 3:
                
                #--Set Dates

                #End date
                #Encontrar el ultimo sabado
                today = date.today()
                idx = (today.weekday() + 1) % 7
                sat = today - datetime.timedelta(7+idx-6)  
                sat_format = sat.strftime("%m/%d/%Y")
                end_date = sat_format # Ultimo sabado disponible
                End_date = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.end_date))
                End_date.clear()
                End_date.send_keys(end_date)
                
                #Start date
                #Encontrar el ultimo domingo despues del ultimo sabado
                sun = sat - datetime.timedelta(days=6)
                sun_format = sun.strftime("%m/%d/%Y") #Previous Sunday before that Saturday 
                start_date = sun_format # 12/27/2020 
                Start_date = WebDriverWait(self.driver,60).until(EC.visibility_of_element_located(self.start_date))
                Start_date.clear()
                Start_date.send_keys(start_date)
                time.sleep(3)

            else:
                print('Error, please select 1,2 or 3 for num bot YOU ARE IN SET_DATES')

        
        except TimeoutException:
            print ("Loading -SET_DATES- took too much time!")



    def Prod_filter(self, numbot):
        try:
            if numbot == 1:
                #Vendor Button
                #WebDriverWait(self.driver,30).until(EC.visibility_of_element_located((By.ID, self.product_filter)))
                prod_filter = self.driver.find_element_by_id(self.product_filter)
                prod_filter.click()
                
                #Select vendor  
                #WebDriverWait(self.driver, 50).until(EC.visibility_of_element_located(self.vendor_dropdown))
                VENDOR = Select(self.driver.find_element_by_id(self.vendor_dropdown))
                VENDOR.select_by_visible_text("Vendor")
                
                flecha1 = WebDriverWait(self.driver,10).until(EC.visibility_of_element_located(self.flecha_vendorstartwithK))
                flecha1.click()
                
                #Select KIND HEALTHY SNACKS
                #time.sleep(1)
                WebDriverWait(self.driver,10).until(EC.visibility_of_element_located((By.XPATH, self.kind)))
                KINDD = self.driver.find_element_by_xpath(self.kind)
                KINDD.click()
                    
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
            store_006 = WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.store_006))
            store_006.click() 
            store_099 = WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.store_0099))
            store_099.click()
            store_blank = WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.store_BLANK))
            store_blank.click()
          
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
            print ("Loading -PRODUCT_FILTER- took too much time!")



    def POG_filter (self):
        try:
            #POG FILTER button
            pog_button = WebDriverWait(self.driver,30).until(EC.visibility_of_element_located(self.POG))
            pog_button.click()
            
            time.sleep(2)
            
        except TimeoutException:
            print ("Loading -PRODUCT_FILTER- took too much time!")



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
            
            if self.driver.find_element(*self.promo_sales_checkbox2).is_selected()==False:
                self.driver.find_element(*self.promo_sales_checkbox2).click()
            else:
                pass
            
            if self.driver.find_element(*self.promo_sales_checkbox3).is_selected()==False:
                self.driver.find_element(*self.promo_sales_checkbox3).click()
            else:
                pass
            
            if self.driver.find_element(*self.promo_sales_checkbox4).is_selected()==False:
                self.driver.find_element(*self.promo_sales_checkbox4).click()
            else:
                pass
            
            if self.driver.find_element(*self.promo_sales_checkbox5).is_selected()==False:
                self.driver.find_element(*self.promo_sales_checkbox5).click()
            else:
                pass
            
            
            #Units
            if self.driver.find_element(*self.promo_units_checkbox).is_selected()==False:
                self.driver.find_element(*self.promo_units_checkbox).click()
            else:
                pass
            
            WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.promo_units_checkbox2))
            
            if self.driver.find_element(*self.promo_units_checkbox2).is_selected()==False:
                self.driver.find_element(*self.promo_units_checkbox2).click()
            else:
                pass
 
            if self.driver.find_element(*self.promo_units_checkbox3).is_selected()==False:
                self.driver.find_element(*self.promo_units_checkbox3).click()
            else:
                pass
                        
            if self.driver.find_element(*self.promo_units_checkbox3).is_selected()==False:
                self.driver.find_element(*self.promo_units_checkbox3).click()
            else:
                pass
                        
            if self.driver.find_element(*self.promo_units_checkbox4).is_selected()==False:
                self.driver.find_element(*self.promo_units_checkbox4).click()
            else:
                pass
            
            if self.driver.find_element(*self.promo_units_checkbox5).is_selected()==False:
                self.driver.find_element(*self.promo_units_checkbox5).click()
            else:
                pass
            
            
            #AVG PRICE
            if self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox).is_selected()==False:
                self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox).click()
            else:
                pass
            
            WebDriverWait(self.driver,25).until(EC.visibility_of_element_located(self.promo_AVGRETAILPRICE_checkbox2))
            
            if self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox2).is_selected()==False:
                self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox2).click()
            else:
                pass            
            
            if self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox3).is_selected()==False:
                self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox3).click()
            else:
                pass            
            
            if self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox4).is_selected()==False:
                self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox4).click()
            else:
                pass            
     
            if self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox5).is_selected()==False:
                self.driver.find_element(*self.promo_AVGRETAILPRICE_checkbox5).click()
            else:
                pass            

        except TimeoutException:
            print ("Loading -metrics_filter- took too much time!")
            
            
            
    def excel(self):
        try:
            #Click excel button
            time.sleep(3)
            button = self.driver.find_element_by_id(self.excel_button)
            button.click()
            
        except TimeoutException:
            print ("Loading -Excel BUTTON- took too much time!")

