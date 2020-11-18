# Para Capiturar a Lista de ações que compõem o indice bovespa

from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
import time

page = webdriver.Chrome(ChromeDriverManager().install())
page.get('http://bvmf.bmfbovespa.com.br/indices/ResumoCarteiraQuadrimestre.aspx?Indice=IBOV&idioma=pt-br')
time.sleep(10)
download = page.find_element_by_xpath('//*[@id="ctl00_contentPlaceHolderConteudo_Export_Test"]')
download.click()
time.sleep(10)