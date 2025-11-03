# Guia de Configuração e Execução do Projeto

Este documento detalha os pré-requisitos, a configuração do ambiente e as funcionalidades dos parâmetros que controlam o comportamento do processamento de dados desenvolvido para este projeto.

O objetivo é garantir que todas as dependências e estruturas obrigatórias estejam corretamente configuradas, assegurando a execução sem falhas.

## Requisitos do Projeto

Antes de iniciar qualquer processamento, instale todas as dependências necessárias descritas no arquivo requirements.txt.

No terminal, execute:

**pip install -r requirements.txt**

A biblioteca pip deverá estar atualizada e funcionando corretamente em seu ambiente Python.

## Configuração Inicial

### Definição dos Caminhos de Trabalho (path)

No arquivo principal do projeto existem variáveis path responsaveis por indicar o local onde serão gerados e armazenados os inputs e outputs.

Atualize as variáveis com o caminho absoluto conforme a tabela e exemplos abaixo.

|NOME VARIÁVEL|DESCRIÇÃO|EXEMPLO|
|-------------|---------|-------|
|pathSMS_RCS|Local onde a "PLANILHA AÇÃO SMS_RCS.xlsx" se encontra|"C:/Users/OT/Documents/Projetos/SeparadorDeRCS_SMS/input/"|
|pathWhatsApp|Local onde a "PLANILHA AÇÃO WHATSAPP.xlsx" se encontra|"C:/Users/OT/Documents/Projetos/SeparadorDeWhatsApp/input/"|
|pathCodigo|Local onde o código e estrutura de planilha obrigatória estão|"C:/Users/OT/Documents/Projetos/DocumentadorAcoesInnovare/"|

Certifique-se de que o diretório configurado já exista no sistema.

### Estrutura de Pastas Obrigatória

No diretório configurado em path devem existir obrigatoriamente as seguintes pastas:

``` bash
DocumentadorDeAcoes/
│
└── output/
```

Esse diretório será utilizado para armazenar os arquivos de saída.

A pasta ARQUIVOS AUXILIARES **deve sempre conter os arquivos utilizados para cruzamentos e validações durante o processamento.**

### Forma de utilizar

1. Verifique se as planilhas de ações estão devidamente atualizadas e os paths estão corretos

2. Inicie o código e aguarde a mensagem "Finalizado"

3. Retire a planilha para importação de eventos do COBMAIS no diretório "output"