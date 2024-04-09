import os
from time import sleep
from selenium import webdriver
from selenium.webdriver.chrome.webdriver import WebDriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.remote.webelement import WebElement
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from typing import List, Literal

class Imobme:
    def __init__(self, *, user:str, password:str, ambiente:str="qas") -> None:
        self.__navegador = webdriver.Chrome()
        self.__user:str = user
        self.__password:str = password
        
        self.__url_base:str
        if (ambiente == "qas") or (ambiente == "dev"):
            self.__url_base = "http://qas.patrimarengenharia.imobme.com/"
        elif ambiente == "prd":
            self.__url_base = "https://patrimarengenharia.imobme.com/"
        else:
            raise ValueError(f"{ambiente=} não é valido")
        
        self.navegador.get(f"{self.url_base}/Autenticacao/Login")
        self.navegador.get(f"{self.url_base}/Autenticacao/Login")
        
        self._conectar()         
                    
    @property
    def navegador(self):
        return self.__navegador
    
    @property
    def url_base(self):
        return self.__url_base
    @url_base.setter
    def url_base(self, value:str):
        if not isinstance(value, str):
            raise TypeError("apenas strings")
        self.__url_base = value

        
    
    def _conectar(self):
        while True:
            try:
                self._find_element(self.navegador, By.ID, 'login', timeout=2).send_keys(self.__user)
                
                password_camp = self._find_element(self.navegador, By.ID, 'password', timeout=2)
                password_camp.clear()
                password_camp.send_keys(self.__password)
                password_camp.send_keys(Keys.RETURN)

                if ((msg_error:=self._find_element(self.navegador, By.CLASS_NAME, 'validation-summary-errors', ignore=True, timeout=2).text) != "") and ("Imobme - Autenticação" in self._find_element(self.navegador, By.TAG_NAME, 'html').text):
                    raise ConnectionRefusedError(msg_error)

                self._find_element(self.navegador, By.XPATH, '/html/body/div[2]/div[3]/div/button[1]', timeout=1, ignore=True).click()

                return
            
            except ConnectionRefusedError as error:
                raise ConnectionRefusedError(error)
            except:
                self._wait_imobme_page(delay=2)
    
    def executar(self, *, empreendimento:str, bloco:str, unidade:str) -> str:
        empreendimento = str(empreendimento)
        bloco = str(bloco)
        unidade = str(unidade)
        
        self.navegador.get(f"{self.url_base}/Contrato/Quitacao")
        
        #empreendimento
        self._wait_imobme_page(delay=0.1)
        bt_empreendimento:WebElement = self._find_element(self.navegador, By.ID, 'EmpreendimentoId_chzn')
        bt_empreendimento.click()

        lista_empreendimento:List[WebElement] = bt_empreendimento.find_elements(By.TAG_NAME, 'li')
        find_element:bool = False
        for items in lista_empreendimento:
            if items.get_attribute("innerHTML") == empreendimento:
                self._wait_imobme_page(delay=0.05)
                items.click()
                find_element = True
        if not find_element:
            bt_empreendimento.click()
            raise Exception(f"Não foi possivel encontrar o Empreendimento '{empreendimento}' no site do Imobme")

        #bloco
        self._wait_imobme_page(delay=0.1)
        bt_bloco = self._find_element(self.navegador, By.ID, 'BlocoId_chzn')
        bt_bloco.click()

        lista_bloco = bt_bloco.find_elements(By.TAG_NAME, 'li')
        find_bloco:bool = False
        for items in lista_bloco:
            if items.get_attribute("innerHTML") == bloco:
                self._wait_imobme_page(delay=0.05)
                items.click()
                find_bloco = True
        if not find_bloco:
            bt_bloco.click()
            raise Exception(f"Não foi possivel encontrar o Bloco '{bloco}' do empreendimento '{empreendimento}' no site do Imobme")        
 
        #Unidade
        self._wait_imobme_page(delay=0.1)
        bt_unidade = self._find_element(self.navegador, By.XPATH, '//*[@id="Content"]/section/div[2]/div[3]/div[3]/div')
        bt_unidade.click()

        lista_unidade = bt_unidade.find_elements(By.TAG_NAME, 'li')
        find_unidade:bool = False
        for items in lista_unidade:
            if items.text == unidade:
                self._wait_imobme_page(delay=0.05)
                self._find_element(items, By.TAG_NAME, 'label').click()
                find_unidade = True
        if not find_unidade:
            bt_unidade.click()
            raise Exception(f"Não foi possivel encontrar a Unidade '{unidade}' do Bloco '{bloco}' do empreendimento '{empreendimento}' no site do Imobme")
                
        bt_unidade.click()    
        
        #tipo de contratos
        bt_tipo_contratos = self._find_element(self.navegador, By.XPATH, '//*[@id="Content"]/section/div[2]/div[3]/div[4]/div')
        bt_tipo_contratos.click()
        self._find_element(self.navegador, By.XPATH, '//*[@id="Content"]/section/div[2]/div[3]/div[4]/div/ul/li[2]/a/label/input').click()
        bt_tipo_contratos.click()   
        
        #botão pesquisar
        self._find_element(self.navegador, By.ID, 'btnPesquisar').click() 
        
        self._wait_imobme_page()
        
        #verificar se tem contratos
        tabela_contratos = self._find_element(self.navegador, By.ID, 'tableContratoQuitacaoTeste')
        tbody_contratos = self._find_element(tabela_contratos, By.TAG_NAME, 'tbody')


        if len(tbody_contratos.find_elements(By.TAG_NAME, 'tr')) <= 0:
            raise Exception("contrato não encontrado")
        
        #clicar em quitar
        self._find_element(self.navegador, By.XPATH, '//*[@id="Footer"]/div/input').click()
        
        try:
            retorno = self._find_element(self.navegador, By.XPATH, '//*[@id="divAlert"]/div/div').text
            return f"{retorno} -> {empreendimento} -> {bloco} -> {unidade}"
        except:
            return f"-0... Contratos quitados com sucesso. -> {empreendimento} -> {bloco} -> {unidade}"
        
        
    
    def _find_element(self, browser:WebDriver|WebElement, by:str, target:str, timeout:int=15, ignore:bool=False, speak:bool=False):
        for _ in range(timeout*2):
            try:
                element = browser.find_element(by, target)
                print(f"find: {target=}") if speak else None
                return element
            except:
                sleep(0.5)
        
        if ignore:
            print(f"ignore: {target=}") if speak else None
            return browser.find_element(By.TAG_NAME, 'html')
        
        raise Exception(f"Not Found: {target=}")
        
    def _wait_imobme_page(self, *, delay:float|int=1, after:float|int=0):
        sleep(delay)
        while self.navegador.find_element(By.ID, 'feedback-loader').get_attribute("style") == 'display: block;':
            sleep(0.005) 
        sleep(after)               
                

if __name__ == "__main__":
    from credenciais import Credential
    crd = Credential("credencial_imobme_qas.json").load()
    
    
    from tkinter import filedialog
    import pandas as pd
    
    df = pd.read_excel(filedialog.askopenfilename())
    
    bot = Imobme(user=crd['user'], password=crd['password'])
    for row,value in df.iterrows():
    #bot = Imobme(user=crd['user'], password=crd['password'])
        try:
            print(bot.executar(empreendimento=value['Empreendimento'], bloco=value['Bloco'], unidade=value['Unidade']))
        except Exception as error:
            print(error)
    
    
    bot.navegador.close()
    input("fim: ")
    