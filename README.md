# Projeto Euro6

- Objetivo: Extrair, através do python, dados do sistema Mocha (Mainframe), da empresa Mercedes Benz, para preencher a planilha de controle do projeto Euro6.

- Raciocínio:

    1. Conectar o sistema da empresa à um emulador, através do WC3270
    2. Ler os dados da planilha excel, através do Pandas, para conseguir os números AFAB
    3. Fazer o login e acessar "pro-pmenu", um código de serviço que permite a pesquisa de veículos através do número de AFAB
    4. Extrair todos os dados necessários e armazená-los em listas
    5. Transformar as listas em DataFrame
    6. Passar os DataFrames para o excel
    
- Bibliotecas usadas:

    1. Py3270
    2. Pandas
    
- Observações: 

    1. Os números AFAB já estão na planilha de controle, os dados que procuramos são:
    
           Variante, número de produção e FZ do veículo
           Variante, baumuster  e FZ do eixo
    2. Exemplo de como fica:
    
    ![exemplo 1](https://user-images.githubusercontent.com/77248245/219045117-680d34b4-dc15-4161-885c-9fdffe7da2d0.PNG)
![exemplo 2](https://user-images.githubusercontent.com/77248245/219045124-07e37d89-6b86-4f59-b935-fd20603e81ea.PNG)
