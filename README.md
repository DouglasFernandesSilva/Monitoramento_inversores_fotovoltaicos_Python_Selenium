# Monitoramento de inversores fotovoltaicos com Python e Selenium
O projeto consiste no monitoramento automatizado do status de funcionamento de diversos inversores de diferentes fabricantes.

Para a elaboração do código foi utilizado Python e a biblioteca Selenium, que realiza de forma autonoma o login nas contas das plataformas de monitoramento e "web scraping" das informações relevantes para o projeto.

A rotina é simples, o software faz login nas contas dos clientes nas plataformas de monitoramento, coletas as informações referentes ao status de funcionamento do inversor e caso apresente status de falha, envia um e-mail autómatico informando o problema. Após verificar o status de todos os inversores o software entra em "Stand-by" através do comando "time.sleep(3600)" por 3600 segundos. Transcorrido este periodo, o software entra em mais um ciclo de monitoramento.
