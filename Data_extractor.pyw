import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from datetime import datetime as dt
import win32com.client as win32
import os
import pandas as pd

class DataExtractor:
    data = str(dt.now().day)+"/"+str(dt.now().month)+"/"+str(dt.now().year)

    def __init__(self) -> None:
        self.gettingData()

    def gettingData(self) -> None:
        #COMANDOS UTILIZADOS PARA ACESSAR O CHROME E FAZER TODA A PESQUISA PELO SITE
        driver = webdriver.Chrome(executable_path=r'./chromedriver.exe')
        driver.maximize_window()
        url = "https://datacatalog.worldbank.org/home"
        driver.get(url)
        time.sleep(1)
        driver.find_element(By.XPATH,'/html/body/app-root/header/nav/div/div/div[2]/div/div[2]/ul/li[2]').click()
        time.sleep(4)

        driver.find_element(By.XPATH,'/html/body/app-root/div/app-search-page/div[2]/div[2]/div/div/div/section/div/div[1]/div[3]/div/div/div/div/div/div[1]/div/div/div/h5/div/div[2]/a/div').click()
        time.sleep(4)

        driver.find_element(By.XPATH,'//*[@id="tab1"]/div/div[1]/div/div/h5/div/a[1]').click()

        time.sleep(10)

        iterator = True
        while iterator:
            archive = r"INSIRA_CAMINHO_ONDE_O_ARQUIVO_SERA_SALVO"
            if os.path.isfile(archive) : iterator = False

        driver.close()

        df = pd.read_csv(r"INSIRA_CAMINHO_ARQUIVO_ORIGINAL", sep=",")

        df = df[(df["code"] == "ABW")]
        
        print(df)

        df.to_csv(r"INSIRA_CAMINHO_ARQUIVO_ATUALIZADO", sep=",")

        input("PRESSIONE ENTER PARA ENVIAR ARQUIVO")
        self.sendEmail()

    #COMANDOS RESPONSÁVEIS POR REALIZAR O ENVIO DAS MENSAGENS POR EMAIL

    def sendEmail(self)  -> None:
        outlook = win32.Dispatch("outlook.application")
        email = outlook.CreateItem(0)
        email.To = ("INSIRA_EMAIL_DESTINATÁRIO")
        email.Subject = "Base Extraida na data "+self.data
        email.HTMLBody = f"""
        <P> Boa noite !</P>
        <P></P>
        <P>Segue em anexo a base extraida do dia {self.data}.</P>
        <P></P>
        <P>Att,</P>
        <P> <Strong> INSIRA_ASSINATURA_DESTINATÁRIO </Strong>
        """
        email.Attachments.Add(r"INSIRA_CAMINHO_ARQUIVO_ATUALIZADO")
        email.Send()
        print("Email enviado !")


if __name__ == "__main__":
    obj = DataExtractor
    obj()
    # © Renan Oliveira 2020 - 2023