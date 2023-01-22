import pandas as pd #manipula e limpa 
from playwright.sync_api import sync_playwright #para automatizar
import time #tempo
import asyncio #usado para estruturas assíncronas

nome_do_arquivo = "Contatos.xlsx"
url_do_forms = "https://docs.google.com/forms/d/e/1FAIpQLScRVru_GfOebl5FUKBFNWObWhykNV1p72M4Z4p54BgMx9DlYA/viewform"
# Carregando os dados do arquivo Excel para um DataFrame do pandas
df = pd.read_excel(nome_do_arquivo)

#funçaõ caso ocorra um erro ao fechar o navegador
async def close_browser():
    try:
        await browser.close()
    except Exception as e:
        print("Ocorreu um erro ao fechar o navegador: ", e)
        
# Iterando sobre as linhas do DataFrame
for index,row in df.iterrows():
    print("Index: " + str(index) + " - O nome da pessoa é  "  + row["nome"])
    
    with sync_playwright() as p:
         # Abrindo o navegador
        browser = p.chromium.launch(headless=False)
        # Abrindo a página do formulário
        page = browser.new_page()
        page.goto(url_do_forms)
        time.sleep(2)
        
       #preenche os campos do excel pegando os seletores do formulario atraves do 
        page.fill("#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(1) > div > div > div.AgroKb > div > div.aCsJod.oJeWuf > div > div.Xb9hP > input",row["nome"])
        page.fill("#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(2) > div > div > div.AgroKb > div > div.RpC4Ne.oJeWuf > div.Pc9Gce.Wic03c > textarea", str(row["telefone"]))
        page.fill("#mG61Hd > div.RH5hzf.RLS9Fe > div > div.o3Dpx > div:nth-child(3) > div > div > div.AgroKb > div > div.aCsJod.oJeWuf > div > div.Xb9hP > input", str(row["nota"]))
        page.click("#mG61Hd > div.RH5hzf.RLS9Fe > div > div.ThHDze > div.DE3NNc.CekdCb > div.lRwqcd > div > span > span")

