import pandas as pd
import numpy as np
import datetime as dt
import glob as gb

# Variavel para o local de armazenamento dos arquivos processados
path = "C:/Users/OT/Documents/Projetos/PosicaoCarteira/"

# Variavel para definir entre primeiro nome ou nome completo (Não aplicável a empresas)
primeiroNome = True

# Puxando data de hoje e ontem
hoje = dt.date.today().strftime("%d_%m_%Y")
ontem = (dt.date.today() - dt.timedelta(days=1)).strftime("%d_%m_%Y")

# Lendo base e-consignado
dfEConsignado = pd.read_excel(f"{path}input/Base eConsignado.xlsx", dtype={"CPF/CNPJ": str})

mapaSafras = {
    ""
}

listaColunasContratos = [
    "CPF/CNPJ",
    "CLIENTE",
    "CONTRATO",
    "CREDOR",
    "ESTAGIO",
    "TOTAL ABERTO"
]

colunasRelatorio = [
    "CARTEIRAS",
    "QTD_CLIENTES_ATIVOS",
    "QTD_CLIENTES_NOVOS",	
    "QTD_CONTRATOS_ATIVOS",
    "QTD_CONTRATOS_NOVOS",
    "VALOR_ACUMULADO",
    "QTD_TELEFONES_ATIVOS",
    "QTD_ENTRADAS",
    "QTD_SAIDAS"
]

def separarCredor(df):
    # Salvando informações em variáveis para aumentar a performance
    credor = df["CREDOR"].str.strip()
    tipo_contrato = df["TIPO DE CONTRATO"]
    descricao = df["Descrição"]
    atraso = df["ATRASO"]
    cpf_cnpj = df["CPF/CNPJ"]
    sexto_digito = cpf_cnpj.str[5].astype(int)
    sextoESetimo_digito = cpf_cnpj.str[5:7].astype(int)

    # Separando as carteiras por numpy select
    condicoesCarteiras = [
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(1, 30)) & (sexto_digito.isin([0,1,2])),
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(1, 30)) & (sexto_digito.isin(range(3,10))),
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(31, 94)) & (sexto_digito.isin([1,2])),
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(31, 94)) & (sexto_digito.isin([6,8,9])),
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(31, 94)) & (sexto_digito.isin([0,3,4,5])),
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso >= 95) & ((sextoESetimo_digito.isin(range(15,20))) 
        | (sexto_digito == 3)),
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso >= 95) & (sexto_digito == 0),
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso >= 95) & (sexto_digito == 8),
        (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso >= 95) & ((sextoESetimo_digito.isin(range(10,15))) 
        | (sexto_digito.isin([2,4,5,6,7,9]))),
        (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "PF") & (sexto_digito.isin([0,2,3,4])),
        (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "PF") & (sexto_digito.isin([1,5,6,7,8,9])),
        (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "Consignado") & (~cpf_cnpj.isin(dfEConsignado["CPF/CNPJ"])),
        (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "Consignado") & (cpf_cnpj.isin(dfEConsignado["CPF/CNPJ"])),
        (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "MEI") & (sexto_digito.isin([1,3,5,7,9])),
        (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "MEI") & (sexto_digito.isin([0,2,4,6,8])),
    ]
    resultadosCarteiras = [
        "NEON AMIGAVEL 1 - SAFRA 1",
        "NEON AMIGAVEL 1 - DG",
        "NEON AMIGAVEL 2 - SAFRA 1",
        "NEON AMIGAVEL 2 - SAFRA 2",
        "NEON AMIGAVEL 2 - DG",
        "NEON CRELIQ - SAFRA 1",
        "NEON CRELIQ - SAFRA 2",
        "NEON CRELIQ - BASE ESPECIAL",
        "NEON CRELIQ - DG",
        "NEON EMPRESTIMO - SAFRA 1",
        "NEON EMPRESTIMO - DG",
        "NEON CONSIGNADO",
        "NEON E-CONSIGNADO",
        "NEON EMPRESTIMO MEI - SAFRA 1",
        "NEON EMPRESTIMO MEI - DG"
    ]

    # Separando os projetos seguindo o deXpara 
    df["CREDOR"] = np.select(condicoesCarteiras, resultadosCarteiras, default=credor)
    print("Projetos separados e renomeados")
    return df

# Lendo relatório do dia anterior 
try:
    dfRelatorioOntem = pd.read_excel(f"{path}output/Relatórios/POSIÇÃO CARTEIRA {ontem}.xlsx",usecols="L:T")
except Exception as e:
    dfRelatorioOntem = pd.DataFrame(columns=colunasRelatorio)
    print("Não foi encontrado um arquivo com a posição carteira do dia anterior.\n" \
    f"Se deseja ver o comparativo adicione na pasta output/Relatórios seguindo o modelo \"POSIÇÃO_CARTEIRA {ontem}\"")

try:    
    dfBaseOntem = pd.read_excel(f"{path}output/Histórico/BASE {ontem}.xlsx",dtype={"CPF/CNPJ": str})
except Exception as e:
    dfBaseOntem = pd.DataFrame(columns=listaColunasContratos)
    print("Não foi encontrado um arquivo com a base do dia anterior.\n" \
    f"Se deseja ver o comparativo adicione na pasta output/Relatórios seguindo o modelo \"POSIÇÃO_CARTEIRA {ontem}\"")

# Lê uma coletanea de arquivos excel a partir de uma planilha e os unifica
def leitorDeExcel(sheetName):
    dataframes = []

    for arquivo in gb.glob(f"{path}/input/Bases/*.xlsx"):
        try:
            df = pd.read_excel(arquivo,sheet_name=sheetName,dtype={"CPF/CNPJ" : str})
            dataframes.append(df)
            print(f"Arquivo {arquivo}, planilha {sheetName} processado!")
        except Exception as e:
            print(f"Arquivo {arquivo} não foi processado e retornou o erro {e}")

    return pd.concat(dataframes, ignore_index=True)

# Carregando o arquivo
dfTelefones = leitorDeExcel("Telefones")
dfContratos = leitorDeExcel("Contratos")
dfProdutos = leitorDeExcel("Produtos")
print("Bases Lidas com sucesso")

# Adicionando coluna descrição caso não tenha
if not "Descrição" in dfProdutos.columns:
    dfProdutos.insert(5,"Descrição", None)  

# Limpando Inativos/BlockList/Números com menos de 10 digitos
dfTelefones = dfTelefones[
    (dfTelefones["BLOCKLIST"] != "SIM") & 
    (dfTelefones["ATIVO"] != "NÃO") & 
    (dfTelefones["NUMERO"].astype(str).str.len() >= 11)
]

# Retirando clientes com atraso negativo
dfContratos = dfContratos[dfContratos["ATRASO"] > 0]

# Buscando as informações da aba produtos
dfContratos = pd.merge(dfContratos,dfProdutos[["CONTRATO","TIPO DE CONTRATO", "Descrição"]],how="left", on="CONTRATO")

# Retirando as duplicadas
dfContratos = dfContratos.drop_duplicates(subset=["CPF/CNPJ", "CONTRATO"])

# Separando credores
dfContratos = separarCredor(dfContratos)

# Retirando colunas que não serão usadas
dfContratos = dfContratos[listaColunasContratos]

MapaTelefonesAtivos = dfTelefones[dfTelefones["ATIVO"] == "SIM"].value_counts(subset="CPF/CNPJ")

dfContratos["QTD_TELEFONES_ATIVOS"] = dfContratos["CPF/CNPJ"].map(MapaTelefonesAtivos)
dfContratos["QTD_TELEFONES_ATIVOS"] = dfContratos["QTD_TELEFONES_ATIVOS"].fillna(0)

dfContratos.to_excel(f"{path}output/Histórico/BASE {hoje}.xlsx", index=False)
print(f"Base de {hoje} gerada")

dfRelatorio = pd.DataFrame(columns=colunasRelatorio)

dfRelatorio["CARTEIRAS"] = dfContratos["CREDOR"].unique()
qtdClientesAtivos = dfContratos.groupby("CREDOR")["CPF/CNPJ"].nunique()
qtdClientesNovos = dfContratos[dfContratos["ESTAGIO"] == "NOVO"].drop_duplicates("CPF/CNPJ").groupby("CREDOR").size()
qtdContratosAtivos = dfContratos.groupby("CREDOR").size()
qtdContratosNovos = dfContratos[dfContratos["ESTAGIO"] == "NOVO"].groupby("CREDOR").size()
valorAcumulado = dfContratos.groupby("CREDOR")["TOTAL ABERTO"].sum()
qtdTelefonesAtivos = dfContratos.groupby("CREDOR")["QTD_TELEFONES_ATIVOS"].sum()
qtdEntradas = dfContratos[~dfContratos["CPF/CNPJ"].isin(dfBaseOntem["CPF/CNPJ"])].drop_duplicates("CPF/CNPJ").groupby("CREDOR").size()
qtdSaidas = dfBaseOntem[~dfBaseOntem["CPF/CNPJ"].isin(dfContratos["CPF/CNPJ"])].drop_duplicates("CPF/CNPJ").groupby("CREDOR").size()

dfRelatorio["QTD_CLIENTES_ATIVOS"] = dfRelatorio["CARTEIRAS"].map(qtdClientesAtivos).fillna(0).astype(int)
dfRelatorio["QTD_CLIENTES_NOVOS"] = dfRelatorio["CARTEIRAS"].map(qtdClientesNovos).fillna(0).astype(int)
dfRelatorio["QTD_CONTRATOS_ATIVOS"] = dfRelatorio["CARTEIRAS"].map(qtdContratosAtivos).fillna(0).astype(int)
dfRelatorio["QTD_CONTRATOS_NOVOS"] = dfRelatorio["CARTEIRAS"].map(qtdContratosNovos).fillna(0).astype(int)
dfRelatorio["VALOR_ACUMULADO"] = dfRelatorio["CARTEIRAS"].map(valorAcumulado).fillna(0).astype(float)
dfRelatorio["QTD_TELEFONES_ATIVOS"] = dfRelatorio["CARTEIRAS"].map(qtdTelefonesAtivos).fillna(0).astype(int)
dfRelatorio["QTD_ENTRADAS"] = dfRelatorio["CARTEIRAS"].map(qtdEntradas).fillna(0).astype(int)
dfRelatorio["QTD_SAIDAS"] = dfRelatorio["CARTEIRAS"].map(qtdSaidas).fillna(0).astype(int)

with pd.ExcelWriter(f"{path}output/Relatórios/POSIÇÃO CARTEIRA {hoje}.xlsx",engine="openpyxl") as writter:
    dfRelatorioOntem.to_excel(writter,sheet_name="POSIÇÃO CARTEIRA",index=False, startcol=0)
    dfRelatorio.to_excel(writter,sheet_name="POSIÇÃO CARTEIRA",startcol=11,index=False)

print(f"Posição carteira de {hoje} gerada")
print("Finalizado")

        
