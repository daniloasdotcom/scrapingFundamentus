# Para o adequado funcionamento deste código é necessária a instalação dos pacotes abaixo

import openpyxl
import os
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from datetime import datetime
import time


# usamos o modulo chdir para alterar o endereço do diretório de trabalho
os.chdir("C:\\Users\\Usuário\\Desktop")
diaNegocios = "20/05/2020"

# Se seguir são criadas três listas vazias que usaremos mais adiante para:
# listAtuais - armazenar o nome das ações que estão com dados atualizados do ultimo dia de negociações
# listAntigo - armazenar o nome das ações que não estão com dados atualizados do ultimo dia de negociações
# listStocks - armazenará a lista de ações listada na planilha diariamente.xlsx

listAtuais = []
listAntigo = []
listStocks = []

# A seguir:
# O modulo load_workbook() da função openpyxl abre a planilha diariamente.xlsx
# A planilha "todas" é acessada pelo wb['Todas'] e armazenada na variável sheet

wb = openpyxl.load_workbook('diariamente.xlsx')
sheet = wb['Atualizadas']

# A partir de então a planilha está sendo acessada
# O laço for irá trabalhar nela
# 1- substituirá o i com o número 2 até o 322
# 2- sheet.cell irá acessar linha por linha, da 2 a 322 e a cada linha capiturará o valor dela e aramazenará na variável nameStock
# 3- o modulo append() colocará o valor cada valor assumido por nameStock, a cada loop for, na lista listStocks, que foi inicalmente criada como vazia

for i in range(2, 350, 1):
    nameStock = sheet.cell(row=i, column=1).value
    listStocks.append(nameStock)

# Abaixo criamos dois contadores

conta = 0  # Representa o total de ações que foram verificadas no site fundamentus a cada loop
J = 1  # Representa o total de ações verificadas no site fundamentus, com datas do ultimo dia de negociações e listadas na planilha
linAntigas = 1

# As linhas a seguir
# 1 - Amazenam a função que abre o Chome na variável fund
# 2 - Chama o modulo get() para abrir o navegador e acessar o endereço indicado
# 3 - Chama a função maximize_window() para maximizar a tela

fund = webdriver.Chrome(ChromeDriverManager().install())
fund.get("http://www.fundamentus.com.br/detalhes.php?papel=")
fund.maximize_window()

# No laço for a seguir:
# Irá passar por cada um das siglas da ações inicialmente armazenadas na listStocks
# find_element_by_id('completar') encontra a caixa no site fundamentus para inserir a sigla das ações
# send_keys(i) insere a sigla da açao i
# time.sleep(1.5) é um intervalo inserido no processamento do código para que dê tempo de um completo carregamento do código entre alguma ações exexultadas por ele
# login_attempt encontra o botão para acessar os dados da ação de interesse
# o botão é clicado ao aplicada o modulo .submit()
for i in listStocks:
    comentar = fund.find_element_by_id('completar')
    comentar.send_keys(i)
    time.sleep(1.5)
    login_attempt = fund.find_element_by_xpath('/html/body/div[1]/div[1]/form/fieldset/input[2]')
    login_attempt.submit()

    time.sleep(1.5)

    # Dado que algumas das ações listadas podem não apresentar dados e consequentemente levar a erros do nosso código, então:
    # Nos inserimos um "try" e dentro dele tentamos encontrar o dado básico que é a data dos dados do ultimo dia de mercado
    try:
        data = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[1]/tbody/tr[2]/td[4]').text

    # Caso a ação não apresente dados então o "except" imprimirá no console o código abaixo
    except:
        print('Opa tivemos um erro aqui com a ação ' + i)

    # O else execultará o código caso a data seja encontrada
    else:
        # Condição "if"
        # Se a data for encontrada com o valor determinado abaixo então a ação será armazenada na lista "listAtuais"
        # e a condição "if" continuará sendo execultada capiturando os dados presentes na página
        if data == diaNegocios:
            listAtuais.append(i)

            # Os códigos de capitura realizam dois tipos de tratamento.
            # 1- usando ".replace" para substituir "," (virgulas) por "." (pontos)
            # 2- um laço "for" que também faz o mesmo tipo de substituição, contudo ocorre da seguinte forma:
            # 2.1 - Tranformamos o conteúdo da variavem em uma lista
            # 2.1 - O laço então primeiro elimina todos os pontos, inclusive dos agrupadores de digitos
            # 2.2 - depois ele substitui a virgula "," por ".", neste caso do separador decimal
            # Obs: ambos tipos de substituição são realizadas para que o excel reconheça corretamente os dados inseridos neles
            #      já que no meu sistema operacional as configuração estão para reconhecer "." como separadore decimal

            vard = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[2]/td[2]').text
            vard_n = vard.replace(',', '.')

            varm = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[3]/td[2]/span/font').text
            varm_n = varm.replace(',', '.')

            var30d = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[4]/td[2]/span/font').text
            var30d_n = var30d.replace(',', '.')

            var12m = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[5]/td[2]/span/font').text
            var12m_n = list(var12m)
            for x, w in zip(var12m_n, range(0, len(var12m_n), 1)):
                if x == '.':
                    del var12m_n[w]
                elif x == ',':
                    var12m_n[w] = '.'
            var12m_n = ''.join(var12m_n)
            var12m_n = str(var12m_n)

            var2020 = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[6]/td[2]/span/font').text
            var2020_n = list(var2020)
            for x, w in zip(var2020_n, range(0, len(var2020_n), 1)):
                if x == '.':
                    del var2020_n[w]
                elif x == ',':
                    var2020_n[w] = '.'
            var2020_n = ''.join(var2020_n)
            var2020_n = str(var2020_n)

            var2019 = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[7]/td[2]/span/font').text
            var2019_n = list(var2019)
            for x, w in zip(var2019_n, range(0, len(var2019_n), 1)):
                if x == '.':
                    del var2019_n[w]
                elif x == ',':
                    var2019_n[w] = '.'
            var2019_n = ''.join(var2019_n)
            var2019_n = str(var2019_n)

            var2018 = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[8]/td[2]/span/font').text
            var2018_n = list(var2018)
            for x, w in zip(var2018_n, range(0, len(var2018_n), 1)):
                if x == '.':
                    del var2018_n[w]
                elif x == ',':
                    var2018_n[w] = '.'
            var2018_n = ''.join(var2018_n)
            var2018_n = str(var2018_n)

            var2017 = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[9]/td[2]/span/font').text
            var2017_n = list(var2017)
            for x, w in zip(var2017_n, range(0, len(var2017_n), 1)):
                if x == '.':
                    del var2017_n[w]
                elif x == ',':
                    var2017_n[w] = '.'
            var2017_n = ''.join(var2017_n)
            var2017_n = str(var2017_n)

            var2016 = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[10]/td[2]/span/font').text
            var2016_n = list(var2016)
            for x, w in zip(var2016_n, range(0, len(var2016_n), 1)):
                if x == '.':
                    del var2016_n[w]
                elif x == ',':
                    var2016_n[w] = '.'
            var2016_n = ''.join(var2016_n)
            var2016_n = str(var2016_n)

            var2015 = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[11]/td[2]/span/font').text
            var2015_n = list(var2015)
            for x, w in zip(var2015_n, range(0, len(var2015_n), 1)):
                if x == '.':
                    del var2015_n[w]
                elif x == ',':
                    var2015_n[w] = '.'
            var2015_n = ''.join(var2015_n)
            var2015_n = str(var2015_n)

            vpa = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[3]/td[6]/span').text
            vpa_n = vpa.replace(',', '.')

            roe = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[9]/td[6]/span').text
            roe_n = roe.replace(',', '.')

            pa = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[1]/tbody/tr[1]/td[4]').text
            pa_n = pa.replace(',', '.')

            p_vpa = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[3]/td[4]').text
            p_vpa_n = p_vpa.replace(',', '.')

            pl = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[2]/td[4]').text
            pl_n = pl.replace(',', '.')

            lpa = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[2]/td[6]').text
            lpa_n = lpa.replace(',', '.')

            divbppa = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[3]/tbody/tr[11]/td[6]').text
            divbppa_n = divbppa.replace(',', '.')

            setor = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[1]/tbody/tr[4]/td[2]').text

            empresa = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[1]/tbody/tr[3]/td[2]').text

            balanco = fund.find_element_by_xpath('/html/body/div[1]/div[2]/table[2]/tbody/tr[1]/td[4]').text

            # A seguir:
            # A planilha "Atualizadas" é acessada pelo wb['Todas'] e armazenada na variável sheet
            sheet = wb['Atualizadas']

            # Os dados são então armazenados numa lista chamada "listDados"
            listDados = [i, empresa, setor, roe_n, vpa_n, pa_n, p_vpa_n, pl_n, lpa_n,
                         divbppa_n, vard_n, varm_n, var30d_n, var12m_n, var2020_n, var2019_n,
                         var2018_n, var2017_n, var2016_n, var2015_n, balanco, data]

            # Antes de entrar no "for" que irá levar os dados para planilha o contador J recebe mais uma unidade
            # Para que ele começe na linha J + 1
            J = J + 1

            # No laço "for":
            # Cada indice da "listDados" será acessado através de "coluna"
            # J representa a linha que receberá os dados durante o "for"
            # para auxiliar no monitoramento usamos as funções de "datatime" e armazenamos o horário de coleta na ultima coluna de dados
            for coluna in range(0, len(listDados), 1):
                sheet.cell(row=J, column=coluna + 1).value = listDados[coluna]
                data_e_hora_atuais = datetime.now()
                data_e_hora_em_texto = data_e_hora_atuais.strftime('%d/%m/%Y %H:%M')
                sheet.cell(row=J, column=len(listDados) + 1).value = data_e_hora_em_texto

            # assim que completa a condição if a planilha é salva
            wb.save('diariamente.xlsx')

        # Contudo caso a data seja diferente da data determinada
        # a ação i é armazanada na lista "listAntigo"

        elif data != 'diaNegocios':
            listAntigo.append(i)

            # A seguir:
            # A planilha "Desatualizadas" é aberta pelo wb['Desatualizadas']
            sheet = wb['Desatualizadas']

            # Os dados são então armazenados numa lista chamada "listDados"
            listDados = [i, data]

            # Antes de entrar no "for" que irá levar os dados para planilha o contador linAntigas recebe mais uma unidade
            # Para que ele começe na linha linAntigas + 1
            linAntigas = linAntigas + 1

            # No laço "for":
            # Cada indice da "listDados" será acessado através de "coluna"
            # linAntigas representa a linha que receberá os dados durante o "for"
            # para auxiliar no monitoramento usamos as funções de "datatime" e armazenamos o horário de coleta na ultima coluna de dados
            for coluna in range(0, len(listDados), 1):
                sheet.cell(row=linAntigas, column=coluna + 1).value = listDados[coluna]
                data_e_hora_atuais = datetime.now()
                data_e_hora_em_texto = data_e_hora_atuais.strftime('%d/%m/%Y %H:%M')
                sheet.cell(row=linAntigas, column=3).value = data_e_hora_em_texto

            # a planilha é salva assim que completa a condição elif
            wb.save('diariamente.xlsx')

        # Para fins de monitoramento as linhas abaixo foram criadas
        conta = conta + 1

        print('Tempo total, rodando os dados, estimado é de: ', conta * 4.1, 'segundos')
        print('A ultima ação verificada foi a ', i)
        print('O total de ações verificada é de: ', conta)

    print('Lista com datas atuais possuem um total de: ', len(listAtuais))
    print('Lista com datas antigas possuem um total de: ', len(listAntigo))
