import pandas as pd
import numpy as np
import datetime as dt
import glob as gb

# Variavel para o local de armazenamento dos arquivos processados
path = "C:/Users/OT/Documents/Projetos/Base_Innovare/"

hoje = dt.date.today().strftime("%d_%m_%Y")

# Lê uma coletanea de arquivos excel a partir de uma planilha e os unifica
def leitorDeExcel(sheetName):
    dataframes = []

    for arquivo in gb.glob(f"{path}input/bases/*.xlsx"):
        try:
            df = pd.read_excel(arquivo,sheet_name=sheetName,dtype={"CPF/CNPJ" : str})
            dataframes.append(df)
            print(f"Arquivo {arquivo}, planilha {sheetName} processado!")
        except Exception as e:
            print(f"Arquivo {arquivo} não foi processado e retornou o erro {e}")

    return pd.concat(dataframes, ignore_index=True)
    
# Carregando os arquivos
dfContratos = leitorDeExcel("Contratos")
dfProdutos = leitorDeExcel("Produtos")
dfEConsignado = pd.read_excel(f"{path}input/Base eConsignado.xlsx", dtype={"CPF/CNPJ": str})
print("Bases Lidas com sucesso")

# Adicionando coluna descrição caso não tenha
if not "Descrição" in dfProdutos.columns:
    dfProdutos.insert(5,"Descrição", None)  

# Juntando a base de contratos com a de produtos e organizando o nome das colunas
dfFinal = pd.merge(dfContratos, dfProdutos, how="inner", on="CONTRATO")[["CPF/CNPJ_x","CONTRATO","TOTAL ABERTO", "ATRASO", "DEFASAGEM","ESTAGIO","CREDOR_x","TIPO DE CONTRATO", "Descrição"]]
dfFinal.rename(columns={"CPF/CNPJ_x":"CPF/CNPJ", "CREDOR_x": "CREDOR"},inplace=True)

# Retirando CPF/CNPJ e Contratos duplicados
dfFinal = dfFinal.drop_duplicates(subset=["CPF/CNPJ","CONTRATO"])

# Salvando informações em variáveis para aumentar a performance
credor = dfFinal["CREDOR"].str.strip()
tipo_contrato = dfFinal["TIPO DE CONTRATO"]
descricao = dfFinal["Descrição"]
atraso = dfFinal["ATRASO"]
cpf_cnpj = dfFinal["CPF/CNPJ"]

# Separando as carteiras por numpy select
condicoesCarteiras = [
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(1, 30)),
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(31, 94)),
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso >= 95),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "PF"),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "Consignado") & (~cpf_cnpj.isin(dfEConsignado["CPF/CNPJ"])),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "Consignado") & (cpf_cnpj.isin(dfEConsignado["CPF/CNPJ"])),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "MEI"),
    (credor == "NEON - CONSIGA+"),
    (credor == "TODOS EMPREENDIMENTOS - 1 A 3"),
    (credor == "TODOS EMPREENDIMENTOS - 4 a 6"),
    (credor == "TODOS EMPREENDIMENTOS - 7+ - CENTRO SUL") | (credor == "TODOS EMPREENDIMENTOS - 7+ - COSTA LESTE") |
    (credor == "TODOS EMPREENDIMENTOS - 7+ - EQUATORIAL") | (credor == "TODOS EMPREENDIMENTOS - 7+ - INTERIOR") |
    (credor == "TODOS EMPREENDIMENTOS - 7+ - SAO PAULO MINAS"),
    (credor == "TODOS EMPREENDIMENTOS - NR"),
    (credor == "CARTÃO DE TODOS - VMK - 1 A 3") | (credor == "CARTÃO DE TODOS - VMK - 4+"),
    (credor == "SOCIEDADE BIBLICA DO BRASIL") | (credor == "EDUCBANK") | (credor == "ESTÍMULO") | (credor == "KONSIGAPAY") |
    (credor == "KON SECURITIZADORA") | (credor == "ODONTOCOMPANY") | (credor == "REFILIAÇÃO - VMK") | (credor == "SAMI SAÚDE"),
    (credor == "GRENKE - NEW DEBTORS"),
    (credor == "STONE - EMPRÉSTIMO")
]
resultadosCarteiras = [
    "NEON - AMIGAVEL 1",
    "NEON - AMIGAVEL 2",
    "NEON - CRELIQ",
    "NEON - EPPF",
    "NEON - CONSIGNADO",
    "NEON - E-CONSIGNADO",
    "NEON - EP MEI",
    "NEON - CONSIGA+",
    "TODOS OS EMPREENDIMENTOS 1 A 3",
    "TODOS OS EMPREENDIMENTOS 4 A 6",
    "TODOS OS EMPREENDIMENTOS 7+",
    "TODOS OS EMPREENDIMENTOS NR",
    "CARTAO DE TODOS - VMK 3 A 6",
    "CRP",
    "GRENKE",
    "STONE"
]

dfFinal["PROJETO"] = np.select(condicoesCarteiras, resultadosCarteiras, default=credor)

# Retirando colunas a mais
dfFinal.drop(columns=["CREDOR","TIPO DE CONTRATO","Descrição"],inplace=True)

# Tranformando a base em Excel
dfFinal.to_excel(f"{path}output/BASE {hoje}.xlsx", index=False)
print("Base organizada!")


        
