# Automação de Tratamento de Base CRM COBMAIS

Este script automatiza a limpeza, padronização e formatação da base de dados do CRM COBMAIS, preparando os dados para importação nos templates de disparo das plataformas **PlugLead** e **CDA**.

## 🛠️ Requisitos

- Python 3.8+
- Pandas
- openpyxl (caso use arquivos .xlsx)

Instale os pacotes com:

**pip install -r requirements.txt**

## Template arquivo "Cronograma.csv"

{NUMERO_DO_DISPARO ; QUANTIDADE_DE_ENVIOS ; NOME_NUMERO ; NOME_CREDOR; CRUZAR}

|CAMPO|VALOR|
|-----|-----|
|NUMERO_DO_DISPARO| Número de telefone que efetuará o disparo|
|QUANTIDADE_DE_ENVIOS|Quantidade de envios realizada por esse número (Pode ser escrito "FULL" ao invés de um número, para enviar todos os contatos válidos do credor)|
|NOME_NUMERO|Apelido do número para ser utilizado no nome do arquivo (Não inserir caracteres proíbidos para nome de arquivo, como "/")|
|NOME_CREDOR|Nome do credor que ficará salvo no historico de disparo (utilize a aba "Lista de credores" para preencher esse campo)|
|CRUZAR|Caso valor seja igual a 1, cruzará com a planilha de log, caso seja diferente de um (recomendo valor 0) não irá cruzar|
|SEPARAR_LOTES|Caso o valor seja igual a 1, separa a campanha de CDA em lotes e e salva em XLSX, caso o valor seja diferente de 1, não separa em lotes e salva em CSV (Aplicavel somente ao CDA)| 

## Arquivos input

|NOME DO ARQUIVO|EXTENÇÃO|COLUNAS NECESSÁRIAS|DESCRIÇÃO|
|---------------|--------|-------------------|---------|
|PLANILHA AÇÃO WHATSAPP|.xlsx|NOME, TELEFONE, CPF/CNPJ, PROJETO, DATA ENVIO, TELEFONE UTILIZADO|Local onde o script verifica os antigos envios para não ocorrer repetição e salva os envios automáticamente (todos os clientes que estiverem nesse arquivo NÃO receberão novas ações)| 
|PLANILHA AÇÃO WHATSAPP HISTORICO|.xlsx|NOME, TELEFONE, CPF/CNPJ, PROJETO, DATA ENVIO, TELEFONE UTILIZADO|Local de armazenamento dos envios (deve ser atualizado manualmente a partir da PLANILHA AÇÃO WHATSAPP)|
|Base eConsignado|.xlsx|CPF/CNPJ, Marcadores|Base de e-consignado, utilizada para separar|
|Base Consignado|.xlsx|CPF/CNPJ, Marcadores|Base de consignado e consigamais|
|Cronograma|.xlsx|NUMERO_DO_DISPARO, QUANTIDADE_DE_ENVIOS, NOME_ARQUIVO, NOME_CREDOR, CRUZAR|Input de quais ações serão separadas pelo script|
|SCORE TIER|.csv|CPF/CNPJ Numerico, SCORE TIER|Score tier de pagamento|

## Parâmetros de estratégia

No código temos variáveis booleanas que quando alteradas produzem efeitos estratégicos no resultado final, abaixo segue uma tabela informando cada uma delas e suas funções.

|NOME DO GATILHO(VARIAVEL)|FUNÇÃO|
|-------------------------|------|
|**primeiroNome**|Quando ativo (= True), as ações usam apenas o primeiro nome do cliente; Quando desativado (= False), o nome completo do cliente é utilizado.|
|**comFrase**|Quando ativo (= True), são usadas frases aleatórias (diferentes abordagens para CPF e CNPJ) para as ações de plugLead; Quando desativado (= False), apenas o nome é utilizado.|
|**CDA_em_Lotes**|Quando ativo (= True), o arquivo do CDA é salvo em lotes em **CSV** dividos de 250 em 250 clientes (perfeito para ações que utilizam apenas a variável nome); Quando desativado o arquivo do CDA é salvo em apenas um lote em **XLSX** com todos os clientes (perfeito para ações mais complexas que exigem o uso de mais variáveis)| 

## 📂 Estrutura
``` bash
.
├── input/           # Pasta com arquivos da base CRM original e os inputs de número
├── output/          # Arquivos prontos para PlugLead e CDA
├── separadorDeWhatsApp.py     # Script principal de limpeza e tratativa
├── .gitignore
├── requirements.txt
└── README.md
```

## ▶️ Como usar

## ▶️ Como usar

1. Abra o código e altere a variável "path" conforme o local do arquivo

2. Coloque o(s) arquivo(s) de base do CRM na pasta input/Bases

3. Atualize o arquivo do Cronograma na pasta input/

4. Atualize os arquivos na pasta input/ caso necessário

5. Execute o script: **python separadorDeWhatsApp.py** (Não finalize o processo! Espere a mensagem de finalizado)

6. Os arquivos tratados estarão disponíveis na pasta output/

7. Os envios ficarão salvos no arquivo **input/PLANILHA AÇÃO WHATSAPP.xlsx**

## 📩 Plataformas suportadas

**PlugLead**

**CDA**