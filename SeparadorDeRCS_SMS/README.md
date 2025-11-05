# Automação de Tratamento de Base CRM COBMAIS

Este script automatiza a limpeza, padronização e formatação da base de dados do CRM COBMAIS, preparando os dados para importação nos templates de disparo das plataformas **POINTER** e **KOLMEIA**.

## 🛠️ Requisitos

- Python 3.8+
- Pandas
- openpyxl (caso use arquivos .xlsx)

Instale os pacotes com:

**pip install -r requirements.txt**

## Template arquivo "Cronograma.csv"

{NOME_CREDOR ; QUANTIDADE_DE_ENVIOS ; NOME_ARQUIVO ; TIPO_ENVIO ; CRUZAR}

|CAMPO|VALOR|
|---------|------|
|NOME_CREDOR|Nome do credor que ficará salvo no historico de disparo|
|QUANTIDADE_DE_ENVIOS|Quantidade de envios realizada para esse credor|
|NOME_ARQUIVO|Apelido do envio para ser utilizado no nome do arquivo (Não inserir caracteres proíbidos para nome de arquivo, como "/")|
|TIPO_ENVIO|Se o envio é de SMS escrever "SMS" se for de RCS escrever "RCS"|
|CRUZAR|Caso valor seja igual a 1, cruzará com a planilha de log, caso seja diferente de um (recomento valor 0) não irá cruzar|

## Arquivos input

|NOME DO ARQUIVO|EXTENÇÃO|COLUNAS NECESSÁRIAS|DESCRIÇÃO|
|---------------|--------|-------------------|---------|
|PLANILHA AÇÃO SMS_RCS|.xlsx|CPF/CNPJ, NOME, TELEFONE, DATA ENVIO, PROJETO, TIPO DE ENVIO|Local onde o script verifica os antigos envios para não ocorrer repetição e salva os envios automáticamente (todos os clientes que estiverem nesse arquivo com data de D-7 até D0 **NÃO** receberão novas ações)| 
|PLANILHA AÇÃO SMS_RCS HISTORICO|.xlsx|CPF/CNPJ, NOME, TELEFONE, DATA ENVIO, PROJETO, TIPO DE ENVIO|Local de armazenamento dos envios (deve ser atualizado manualmente a partir da PLANILHA AÇÃO SMS_RCS)|
|Base eConsignado|.xlsx|CPF/CNPJ, Marcadores|Base de e-consignado, utilizada para separar|
|Base Consignado|.xlsx|CPF/CNPJ, Marcadores|Base de consignado e consigamais|
|Cronograma|.xlsx|NOME_CREDOR, QUANTIDADE_DE_ENVIOS, NOME_ARQUIVO, TIPO_ENVIO, CRUZAR|Input de quais ações serão separadas pelo script|
|SCORE TIER|.csv|CPF/CNPJ Numerico, SCORE TIER|Score tier de pagamento|


## Parâmetros de estratégia

No código há apenas um gatilho de estratégia chamado **primeiroNome**, responsável por definir se as ações utilizarão apenas o primeiro nome ou o nome completo do cliente (exceto para clientes **CNPJ**). Quando o gatilho está ativado (True), são usadas apenas as ações com o primeiro nome; quando está desativado (False), o nome completo é utilizado.

## 📂 Estrutura
``` bash
.
├── input/           # Pasta com arquivos da base CRM original e os inputs de número
├── output/          # Arquivos prontos para POINTER e KOLMEIA
├── separadorDeRCS_SMS.py     # Script principal de limpeza e tratativa
├── .gitignore
├── requirements.txt
└── README.md
```

## ▶️ Como usar

1. Abra o código e altere a variável "path" conforme o local do arquivo

2. Coloque o(s) arquivo(s) de base do CRM na pasta input/Bases

3. Atualize o arquivo do Cronograma na pasta input/

4. Atualize os arquivos na pasta input/ caso necessário

5. Execute o script: **python separadorDeRCS_SMS.py** (Não finalize o processo! Espere a mensagem de finalizado)

6. Os arquivos tratados estarão disponíveis na pasta output/

7. Os envios ficarão salvos no arquivo **input/PLANILHA AÇÃO SMS_RCS.xlsx**

## 📩 Plataformas suportadas

**POINTER**

**KOLMEIA**