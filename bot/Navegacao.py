import os
import random
import smtplib
import time
from email import encoders
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import pyautogui
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from scripts.Formatacao import Formatacao
from scripts.Login import Login
from scripts.Verificador import VerificarNotaBaixada


class Bot:
    def __init__(self):
        self.formatacao = Formatacao()
        self.verificador = VerificarNotaBaixada()
        self.login = Login()

    def ler_planilha(self):
        lerPlanilha = pd.read_excel(
            self.formatacao.get_caminhoNotas(),
            self.formatacao.get_sheetnameNotas(),
            dtype={'Documento': str, 'Protocolo': str}
        )

        lerPlanilha['Nome'] = lerPlanilha['Nome'].astype(str)
        lerPlanilha['Produto'] = lerPlanilha['Produto'].astype(str)
        lerPlanilha['Nome do AVP'] = lerPlanilha['Nome do AVP'].astype(str)
        lerPlanilha['Protocolo'] = lerPlanilha['Protocolo'].astype(str)
        lerPlanilha['Documento'] = lerPlanilha['Documento'].astype(str)
        lerPlanilha['Valor do Boleto'] = lerPlanilha['Valor do Boleto'].astype(str)
        lerPlanilha['E-mail do Titular'] = lerPlanilha['E-mail do Titular'].astype(str)

        lerPlanilha = VerificarNotaBaixada.verificadorNotaBaixada(lerPlanilha)
        lerPlanilha = lerPlanilha.reset_index(drop=True)
    
    def abrir_sites(self, browser):
        browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/login.jsf")
        browser.execute_script("window.open('', '_blank');")
        browser.switch_to.window(browser.window_handles[-1])
        browser.get("https://sgc-pss.safewebpss.com.br/gerenciamentoac/#/pages/relatorios/relatorio-emissao")
        browser.switch_to.window(browser.window_handles[0])

    def fazer_login(self,browser):
        login = Login()
        while True:
            try:
                self.loginCpf = browser.find_element(By.ID, "j_username")
                break
            except:    
                continue
        while True:
            try:
                self.password = browser.find_element(By.ID, "j_password")
                break
            except:    
                continue
        while True:
            try:
                self.botaoEntrar = browser.find_element(By.ID, "commandButton_entrar")
                break
            except:    
                continue

        cpf = str(login.get_cpf())
        senha = str(login.get_senha())
        while True:
            try:
                self.loginCpf.send_keys(cpf)
                self.password.send_keys(senha)
                self.botaoEntrar.click()
                break
            except:
                continue

        browser.get("https://sispmjp.joaopessoa.pb.gov.br:8080/nfse/paginas/index.jsf")




        

