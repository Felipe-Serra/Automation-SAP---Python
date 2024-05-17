# Automation-SAP---Python
Processo FBL5N no sistema SAP
Automatização de baixas manuais por meio da linguagem python

Processo:
Alimenta-se a planilha "Projeto de automacao SAP.xlsx" com as informações dos clientes
Por meio da biblioteca "pandas" le-se essa planilha e a salva na variavel "tabela" em formato de dataFrame
Por meio da biblioteca "pyautogui" é possivel acessar o sistema do computador como se estivesse utilizando o mouse

Acessa-se então a FBL5N para obter as todas as faturas em aberto dos clientes de modo individual, sendo que, é na variavel "linha" que há a incrimentação do laço for para avançar para o próximo cliente da planilha

A variavel "coluna" funciona de suporte para a variável "coluna1", já que em "coluna" é copiado todos os valores de faturas do cliente e a variavel "coluna1" faz a contagem dessas faturas, que servirá como referencia para a finalização do laço de repetição while.
  
Nesse laço de repetição é onde ocorrerá a comparação dos valores de "referencia" e "data de vencimento" da planilha com os valores linha a linha da tabela SAP

Após a comparação, fora do laço while, se tiver sido encontrado a fatura correta, ocorrerá o gatilho para ativar o Script SAP da baixa.
