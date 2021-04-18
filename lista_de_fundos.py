# Primeiro importamos os pacotes necessários
import os

import openpyxl

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager

path = "C:\\Users\\Usuário\\Desktop"
os.chdir(path)

# Abriremos o Chrome, o site onde coletaremos os dados dos fundos e maximizamos a tela
fund = webdriver.Chrome(ChromeDriverManager().install())
fund.get("https://fiis.com.br/lista-de-fundos-imobiliarios/")
fund.maximize_window()

fileFii = openpyxl.load_workbook('fiis - Copia.xlsx')
planilha = fileFii['Planilha1']

# Agora passamos um looping for para capiturar os nomes de todos os fundos e armazenaremos cada nome na nossa lista
for i in range(2, 280, 1):
    fii = fund.find_element_by_xpath('//*[@id="items-wrapper"]/div[' + str(i) + ']/a').text
    nomeFii = str(fii[0:7])
    nomeFii = ''.join(nomeFii)
    print(nomeFii)
    planilha.cell(row=i, column=1).value = nomeFii

fileFii.save('fiis - Copia.xlsx')

os.rename('fiis - Copia.xlsx', 'fii.xlsx')






