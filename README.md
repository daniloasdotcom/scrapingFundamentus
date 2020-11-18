![alt text](https://github.com/daniloaspk/scrapingFundamentus/blob/master/images/B3_web_scraping.png)


# Web Scraping com dados fundamentalista de ações da B3

## Para que servem os scripts deste repositório?

Ambos códigos capturam dados do site [fundamentus.com.br](http://fundamentus.com.br)

1. **diariamente.py**: Indicado para capturar uma lista de dados fundamentalistas de ações. Fonte dos dados: fundamentus.com.br

2. **nocomments.py**: Mesmo scritps, mas sem comentários

3. **allStocks.py**: Realiza um checkup das ações que possuem dados atualizados, dados antigos e das que apesar de listadas no site não apresentam tabela de dados. Fonte dos dados: fundamentus.com.br 

4. **geral.py**: Indicado para capturar uma lista de dados fundamentalistas de ações previamente listadas no arquivo "base.xlsx". Fonte dos dados: fundamentus.com.br


**Conteudo extra**

 - **diariamente.xlsx**: Planilha já formatada para receber os dados capturados do script **"diariamente.py"**

 - **base.xlsx**: Planilha já formatada para receber os dados capturados do script **"geral.py"**
