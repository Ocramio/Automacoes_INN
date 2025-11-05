# Guia de Configuração e Execução do Projeto

Este documento detalha os pré-requisitos, a configuração do ambiente e as funcionalidades dos parâmetros que controlam o comportamento do processamento de dados desenvolvido para este projeto.

O objetivo é garantir que todas as dependências e estruturas obrigatórias estejam corretamente configuradas, assegurando a execução sem falhas.

## Requisitos do Projeto

Antes de iniciar qualquer processamento, instale todas as dependências necessárias descritas no arquivo requirements.txt.

No terminal, execute:

**pip install -r requirements.txt**


A biblioteca pip deverá estar atualizada e funcionando corretamente em seu ambiente Python.

## Configuração Inicial

### Definição do Caminho de Trabalho (path)

No arquivo principal do projeto existe uma variável chamada path responsável por indicar o local onde serão gerados e armazenados os resultados.

Atualize essa variável com o caminho absoluto para a pasta de resultados em seu computador.

Exemplos:

### Windows:

**path = r"C:\Users\SeuUsuario\Documents\Projeto\RESULTADOS"**


### Linux / MacOS:

**path = "/home/seuusuario/Documentos/Projeto/RESULTADOS"**

Certifique-se de que o diretório configurado já exista no sistema.

### Estrutura de Pastas Obrigatória

No diretório configurado em path devem existir obrigatoriamente as seguintes pastas:

``` bash
RESULTADOS/
│
├── CARTÃO DE TODOS/
├── NEON/
└── GRENKE/
```

Esses diretórios serão utilizados para armazenar os arquivos de saída, organizados conforme cada projeto.

Além disso, é necessário manter uma pasta destinada aos arquivos auxiliares no padrão indicado abaixo:

``` bash
ARQUIVOS AUXILIARES/
│
└── Bases/
```

A pasta ARQUIVOS AUXILIARES **deve sempre conter os arquivos utilizados para cruzamentos e validações durante o processamento.**

## Arquivos Auxiliares

A nomenclatura, extensão e nomes das colunas dos arquivos devem seguir exatamente o padrão especificado para evitar erros.

|Nome do Arquivo|Extensão|Colunas necessárias|
|---------------|--------|-------------------|
|ACIONADOS AGENTE VIRTUAL|.csv| CPF/CNPJ|
|Base Consignado|.xlsx|CPF/CNPJ, Empresa, Marcadores|
|Base E-Consignado|.xlsx|CPF/CNPJ, Empresa, Marcadores|
|BASE TODOS 7+|.xlsx|CPF, Valor divida, Valor Minimo|
|BASE TODOS PA FIXA|.xlsx|CPF, Valor Divida, Valor Minimo|
|SCORE TIER|.csv|CPF/CNPJ Numerico, SCORE TIER|

Esses arquivos devem estar localizados obrigatoriamente em:

**/ARQUIVOS AUXILIARES**

Recomenda-se validar previamente:

• A acentuação dos nomes das colunas
• A existência de espaços extras
• A consistência entre maiúsculas e minúsculas

Qualquer divergência poderá resultar em falhas na execução ou inconsistências nos resultados.

## Parâmetros de Estratégia e Gatilhos

As variáveis controlam comportamentos específicos da lógica de tratamento. Todas são booleanas:

|Parâmetro|Função|
|---------|---------------------------|
|cruzarComODB|Variável númerica que quando possui valor = 1 Remove registros encontrados na planilha ACIONADOS AGENTE VIRTUAL, quando possui valor diferente de 0 puxa somente os registros encontrados na planilha ACIONADOS AGENTE VIRTUAL e quando possui valor = 0 segue com a base completa|
|cruzarComOBI|Realiza cruzamento com BASE TODOS 7+ e BASE TODOS PA FIXA, aplicável ao projeto Cartão de Todos|
|valorMinimo|Substitui o valor total em aberto pelo valor mínimo (depende de cruzarComOBI)|
|salvarCsv|Se True salva arquivos em CSV, caso contrário salva em XLSX|
|separarHot|Gera planilhas independentes para contatos HOT e Não HOT|
|separarHotPuro|Remove telefones Não HOT quando existir ao menos um HOT (depende de separarHot)|
|scoreTier|Habilita inclusão da classificação SCORE TIER no resultado final|
|dataNoMes|Quando ativa, mantem a data de vencimento no mês atual, quando desativada, permite que a data de vencimento passe para o próximo mês|

No código existe um mapeamento chamado dicionarioMailings que define as colunas finais de saída para cada projeto.

## Avisos Importantes

• Todos os arquivos auxiliares devem estar atualizados antes da execução.
• Estruturas incorretas podem comprometer o resultado dos cruzamentos.
• O caminho configurado em path deve existir previamente.
• Recomenda-se sempre revisar a integridade dos dados antes de qualquer rodagem.

## Suporte

Caso tenha dúvidas sobre as configurações ou apresente erros durante o processamento, recomenda-se revisar:

• O caminho configurado no path
• A estrutura e o posicionamento das pastas
• A validade dos arquivos auxiliares (nomenclatura e colunas)
• A configuração dos parâmetros booleanos