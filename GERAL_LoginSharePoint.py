from cgitb import text
from distutils.dist import DistributionMetadata
import email
from multiprocessing.sharedctypes import Value
from pickle import TRUE
import re 
import time
import subprocess
from tkinter.tix import Tree
from typing import Any
from unicodedata import name
import pyodbc
from re import search
from selenium import webdriver #import webdriver module
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By

def MontaStringComEmailsO(EmailsSemDom, browser): #OWNER
    
    number= 0

    while number < len(EmailsSemDom) : 
        EmailSemD = EmailsSemDom[number]
        indice = EmailSemD.find('@')
        EmailSemD = str(EmailSemD)[0:indice]
        
        #Preenchendo os quadros de “Owners”
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[4]/div/div/div/div[1]/div/div[1]/a').click()
        time.sleep(1)

        # Muda para a janela de e-mail
        window_after = browser.window_handles[2] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela

        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/table/tbody/tr/td/table/tbody/tr/td[1]/input')
        userEmail.click()
        userEmail.send_keys(EmailSemD)
        
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/table/tbody/tr/td/table/tbody/tr/td[2]').click()
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[2]/div/div[1]/div[2]/div[3]/table/tbody/tr/td[2]').click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/button[1]').click()
        
        #Voltando para janela anterior
        window_before = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_before) #ele muda o foco do script para a nova janela
 
        number = number + 1

    return 

def MontaStringComEmailsM(EmailsSemDom, browser): #MEMBER
    
    number= 0

    while number < len(EmailsSemDom) : 
        EmailSemD = EmailsSemDom[number]
        indice = EmailSemD.find('@')
        EmailSemD = str(EmailSemD)[0:indice]
        
        #Preenchendo os quadros de “Members”
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[5]/div[3]/div/div/div[1]/div/div[1]/a').click()
        time.sleep(1)
        
        # Muda para a janela de e-mail
        window_after = browser.window_handles[2] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela

        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/a')
        userEmail.click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/div/table/tbody/tr/td[1]/input')
        userEmail.click()
        userEmail.send_keys(EmailSemD)
        
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[1]/div/div[1]/div/table/tbody/tr/td[2]/a').click()
        time.sleep(1)
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[1]/div/div[1]/div/div[2]/div[2]/div[3]/table/tbody/tr/td[2]').click()
        userEmail = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/button[1]').click() #Salvar
        
        #Voltando para janela anterior
        window_before = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_before) #ele muda o foco do script para a nova janela
 
        number = number + 1

    return 

def Get_ListaDistribuicao(): 
    
    try:
        #con = pyodbc.connect(Trusted_Connection='yes', driver = '{SQL Server}',server = 'localhost\SQLEXPRESS' , database = 'Petro')
        con = pyodbc.connect(r"Driver={SQL Server};server=avareport.westeurope.cloudapp.azure.com;database=Petro;uid=PetroUser;pwd=nV:gR[4O2dvL")
        cur = con.cursor()
        
        cur.execute("Select * from ListaDistribuicao") 
        
        dbData = cur.fetchall()
        #lista = dbData[0][0]
         
        cur.close()
        con.close()

        return dbData

    except Exception as e:
        print('****ERRO****:', e)

def LoginSharePoint_Lista():
    
    try:
        browser = profile = webdriver.FirefoxProfile()
        browser = webdriver.Firefox(firefox_profile=profile, executable_path=(r'C:\Petro_py\geckodriver.exe'))
        browser.get("https://outlook.office365.com/ecp")
    
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '//*[@id="i0116"]')))
        
        #LoginSharePoint
        userLogin = "automacao@m365labsuporte.onmicrosoft.com" #Conta
        userPass = "Fub92902" #Senha
        
        time.sleep(3)
        userEmail = browser.find_element(By.XPATH,'//*[@id="i0116"]')
        userEmail.send_keys(userLogin)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()

        time.sleep(3)
        userPsw = browser.find_element(By.XPATH,'//*[@id="i0118"]')
        userPsw.send_keys(userPass)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
        
        WebDriverWait(browser,20).until(EC.presence_of_element_located((By.XPATH, '/html/body/form/div[10]/div[1]/ul/li[2]/span/a')))
        
        #Abrindo Grupo
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[1]/ul/li[2]/span/a').click()
        time.sleep(1)
        browser.find_element(By.XPATH,'/html/body/form/div[10]/div[3]/div[2]/a').click()
        time.sleep(2)
        browser.switch_to.frame(browser.find_element(By.XPATH,'/html/body/form/div[10]/div[23]/iframe[1]'))
        
        #TAB + ENTER na adição
        browser.find_element(By.XPATH,'//*[@id="ResultPanePlaceHolder_mbxSlbCt_ctl03_distributionGroups_DistributionGroupsResultPane_ToolBar_NewDistributionGroupSplitButton"]/table/tbody/tr/td/div/a/img').click()
        time.sleep(2)
        action = ActionChains(browser)
        action.send_keys(Keys.TAB)
        action.send_keys(Keys.ENTER)
        action.perform()
        
        # Muda para a nova Janela
        window_before = browser.window_handles[0]
        window_after = browser.window_handles[1] #ele salva o endereço dessa nova janela.
        browser.switch_to.window(window_after) #ele muda o foco do script para a nova janela
        
        time.sleep(2)
        arrayName = Get_ListaDistribuicao()
        opcaoName = arrayName[0][3]
        opcaoName = opcaoName.strip()
        userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[1]/input') 
        userName.click()
        userName.send_keys(opcaoName) #campo DisplayName
        
        time.sleep(2)            
        userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[2]/input')
        userName.click()
        userName.send_keys(opcaoName) #campo Alias
        
        time.sleep(2)
        userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[2]/div/table/tbody/tr/td[1]/input').click() #Campo Email

        #Procurando email sem @dominio em Owner
        time.sleep(1)
        campo2 = arrayName[0][9] #Pegando info do campo proprietarios
        campo2 = campo2.rstrip()
        ownerName = campo2.split(' ')
        MontaStringComEmailsO(ownerName, browser)
        
        #Procurando email sem @dominio em Membros
        time.sleep(1)
        campo1 = arrayName[0][7] #Pegando info do campo membros
        campo1 = campo1.rstrip()
        memberName = campo1.split(' ')
        MontaStringComEmailsM(memberName, browser)
               
        #Opção entrada Grupo
        dbData = Get_ListaDistribuicao()       
        opnBtn = dbData
        opcao = str(opnBtn[0][10])       
        opcao = opcao.strip()
              
        if (opcao == 'Aberta'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[1]/table/tbody/tr[1]/td/input').click()

        if (opcao == 'Fechada'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[2]/table/tbody/tr[1]/td/input').click()
        
        if (opcao == 'Owner Approval'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[1]/table/tbody/tr[3]/td/input').click()   
            
        #Opção saída Grupo
        opcao2 = str(opnBtn[0][11])
        opcao2 = opcao2.strip()
        
        if (opcao2 == 'Aberta'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[2]/table/tbody/tr[1]/td/input').click()
            
        if (opcao2 == 'Fechada'):
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[6]/div[2]/table/tbody/tr[2]/td/input').click()      
            
        #Salvando
        userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/button[1]').click()
        time.sleep(5)
        
        ############################################ LOOPING PARA CASO DE ERRO - EMAIL CADASTRADO 2 ############################################
        
        #Mensagem de erro
        txtMensage = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[1]/div/div/div/span[1]')#Texto do Erro
        texto = txtMensage.text
        
        #txtMensage = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[2]/button').click() - xpath botão OK
        #time.sleep(2)    
    
        if ('There are multiple recipients matching identity' in texto):
            txtName = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[2]/button')
            txtName.click()
            time.sleep(1)
            
            arrayName = Get_ListaDistribuicao()
            opcaoName2 = arrayName[0][4]
            opcaoName2 = opcaoName2.strip()
            
            #Campo Display Name
            txtName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[1]/input') 
            txtName.click()
            txtName.clear()
            txtName.send_keys(opcaoName2)
            time.sleep(1)

            #Campo Alias
            txtAlias = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[2]/input')
            txtAlias.click()
            txtAlias.clear()
            txtAlias.send_keys(opcaoName2)
            time.sleep(1)
            
            #Campo Email
            txtAlias = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[2]/div/table/tbody/tr/td[1]/input')
            txtAlias.click()
            txtAlias.clear()
            txtAlias.send_keys(opcaoName2)
            time.sleep(1)
            
            #Salvando
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/button[1]').click()
            time.sleep(5)
            
        ############################################ LOOPING PARA CASO DE ERRO - EMAIL CADASTRADO 3 ############################################
        
        #Mensagem de erro
        txtMensage = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[1]/div/div/div/span[1]')#Texto do Erro
        texto = txtMensage.text
        
        #txtMensage = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[2]/button').click() - xpath botão OK
        #time.sleep(2)    
    
        if ('There are multiple recipients matching identity' in texto):
            txtName = browser.find_element(By.XPATH,'/html/body/form/div[2]/div[1]/div/div[2]/div/div/div[2]/button')
            txtName.click()
            time.sleep(1)
            
            arrayName = Get_ListaDistribuicao()
            opcaoName3 = arrayName[0][5]
            opcaoName3 = opcaoName3.strip()
            
            #Campo Display Name
            txtName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[1]/input') 
            txtName.click()
            txtName.clear()
            txtName.send_keys(opcaoName3)
            time.sleep(1)

            #Campo Alias
            txtAlias = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[1]/div[2]/input')
            txtAlias.click()
            txtAlias.clear()
            txtAlias.send_keys(opcaoName3)
            time.sleep(1)
            
            #Campo Email
            txtAlias = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[2]/div/div/div/div/div[2]/div/table/tbody/tr/td[1]/input')
            txtAlias.click()
            txtAlias.clear()
            txtAlias.send_keys(opcaoName3)
            time.sleep(1)
            
            #Salvando
            userName = browser.find_element(By.XPATH,'/html/body/form/div[3]/div/div[3]/button[1]').click()
            time.sleep(5)
            
        browser.close()
        browser.switch_to.window(window_before)
        browser.close()
        
    except Exception as e:
        print('****ERRO NO FILTRO****:', e)

LoginSharePoint_Lista()