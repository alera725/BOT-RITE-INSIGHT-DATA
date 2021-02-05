# -*- coding: utf-8 -*-
"""
Created on Tue Feb  2 14:47:14 2021

@author: alejandro.gutierrez
"""

#PAGINA LOGIN

#from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait # Guardar en las paginas
from selenium.webdriver.support import expected_conditions as EC #Guardar en las paginas import unittest
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException
import time

class main_login():
    def __init__(self,my_driver):
        self.driver = my_driver
        self.iframe = "//iframe[@name='frameMenu']"
        self.id_login = (By.XPATH, '//*[@id="hlLogin"]/img')
        self.logout = (By.ID, 'hlLogin') #//*[@id="hlLogin"]/img imagen img src=Images/RABtnLogin.png
        
        self.iframe2 = '//*[@id="IFRAME1"]'
        self.iframe3 = "//iframe[@name='frameMenu']"
        self.id_username = (By.ID, 'textBoxUserId')
        self.id_pass = (By.ID, 'textBoxPassword')
        self.id_button = (By.ID, 'btnLogInUser')
        self.rite_insight_loyalty_xpath = (By.XPATH, '//*[@id="navBarPortal_navBarPortal_p3_tv1_item_2_cell"]/nobr')
        #navBarPortal_item_2_cell   navBarPortal_navBarPortal_p3_tv1_item_2_cell
        
    def step_1(self, email, pswd):
        try:
            WebDriverWait(self.driver,50).until(EC.visibility_of_element_located((By.XPATH, self.iframe)))
            iframe = self.driver.find_element_by_xpath(self.iframe)
            self.driver.switch_to.frame(iframe)
            #self.driver.find_element(self.id_login).click()
            time.sleep(10)
            button_login = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.id_login))
            button_login.click()
            
            #out to iframe to default content
            self.driver.switch_to.default_content()
            #time.sleep(10)
            
            #Locate iframe
            #WebDriverWait(self.driver,50).until(EC.visibility_of_element_located((By.XPATH, self.iframe2)))
            iframe2 = self.driver.find_element_by_xpath(self.iframe2)
            #switch into the iframe
            self.driver.switch_to.frame(iframe2)
            
            #Esperar a que se cargue el Login
            username = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.id_username))
            username.send_keys(email)
            passd = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.id_pass)) 
            passd.send_keys(pswd)
            click_login = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.id_button))
            click_login.click() 
            
            #out to iframe to default content
            self.driver.switch_to.default_content()

            WebDriverWait(self.driver,50).until(EC.visibility_of_element_located((By.XPATH, self.iframe3)))
            iframe = self.driver.find_element_by_xpath(self.iframe3)
            #cambiamos de frame 
            self.driver.switch_to.frame(iframe)
            
            time.sleep(2)
            
            button_rite_insight = WebDriverWait(self.driver,50).until(EC.visibility_of_element_located(self.rite_insight_loyalty_xpath))
            button_rite_insight.click()
            
            #out to iframe to default content
            self.driver.switch_to.default_content()
            
        except TimeoutException:
            print ("Loading -login- took too much time!")

