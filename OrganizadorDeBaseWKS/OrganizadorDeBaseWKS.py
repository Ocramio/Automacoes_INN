import pandas as pd
import numpy as np
import datetime as dt
import glob as gb

# Variavel para o local de armazenamento dos arquivos processados
path = "C:/Users/OT/Documents/Projetos/OrganizadorDeBaseWKS/"

# Variavel para definir entre primeiro nome ou nome completo (Não aplicável a empresas)
primeiroNome = True

# Puxando data do dia
hoje = dt.date.today().strftime("%d_%m_%Y")

listaCartaoDeTodos = [
    "TODOS EMPREENDIMENTOS - 1 A 3",
    "TODOS EMPREENDIMENTOS - 4 a 6",
    "TODOS EMPREENDIMENTOS - 7+ - CENTRO SUL",
    "TODOS EMPREENDIMENTOS - 7+ - COSTA LESTE",
    "TODOS EMPREENDIMENTOS - 7+ - EQUATORIAL",
    "TODOS EMPREENDIMENTOS - 7+ - INTERIOR",
    "TODOS EMPREENDIMENTOS - 7+ - SAO PAULO MINAS",
    "TODOS EMPREENDIMENTOS - NR",
    "CARTÃO DE TODOS - VMK - 1 A 3",
    "CARTÃO DE TODOS - VMK - 4+",
    "REFILIAÇÃO - VMK"
]

listaCRP = [   
    "BIZCAPITAL",
    "EDUCBANK",
    "EFIBANK",
    "ESTÍMULO",
    "EURO SECURITIZADORA S/A",
    "GRENKE - NEW DEBTORS",
    "KON SECURITIZADORA",
    "ODONTOCOMPANY",
    "PAGOL",
    "STONE - EMPRÉSTIMO",
    "PROSEGUR",
    "RECOVERY - API",
    "SAMI SAÚDE",
    "SOCIEDADE BIBLICA DO BRASIL"
    "GRUPO VITTA"
]

listaColunasContratos = [
    "CPF/CNPJ",
    "CLIENTE",
    "CREDOR",
    "INCLUSAO",
    "ARQUIVO",
    "CONTRATO",
    "ESTAGIO",
    "PRODUTO",
    "REGIAO",
    "FILIAL",
    "PLANO",
    "OBSERVAÇÃO",
    "DATA",
    "EXPIRAÇÃO",
    "ATRASO",
    "DEFASAGEM",
    "PARCELAS",
    "MENOR VCTO",
    "TOTAL ABERTO"
]

listaColunasTelefones = [
    "CPF/CNPJ",
    "CLIENTE",
    "NUMERO",
    "TIPO",
    "CONTATO",
    "WHATSAPP",
    "OBSERVAÇÃO",	
    "RAMAL",
    "ATIVO",
    "HIGIENIZADO",
    "SCORE",
    "BLOCKLIST"
]

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
dfContratos = pd.merge(dfContratos,dfProdutos[["CONTRATO","TIPO DE CONTRATO", "Descrição"]],how="left", on="CONTRATO")\
[listaColunasContratos + ["TIPO DE CONTRATO", "Descrição"]]

# Retirando as duplicadas
dfContratos = dfContratos.drop_duplicates(subset=["CPF/CNPJ", "CONTRATO"])

# Salvando informações em variáveis para aumentar a performance
credor = dfContratos["CREDOR"].str.strip()
tipo_contrato = dfContratos["TIPO DE CONTRATO"]
descricao = dfContratos["Descrição"]
atraso = dfContratos["ATRASO"]
cpf_cnpj = dfContratos["CPF/CNPJ"]

# Separando as carteiras por numpy select
condicoesCarteiras = [
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(1, 30)),
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(31, 94)),
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso >= 95),
    (credor == "NEON") & (tipo_contrato != "Fatura") & ((descricao == "PF") | (descricao == "MEI")),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "Consignado"),
    credor.isin(listaCartaoDeTodos),
    credor.isin(listaCRP)
]
resultadosCarteiras = [
    "NEON AMIGAVEL 1",
    "NEON AMIGAVEL 2",
    "NEON CRELIQ",
    "NEON EMPRESTIMO",
    "NEON CONSIGNADO",
    "CARTAO SAUDE",
    "CRP"
]

# Separando os projetos seguindo o deXpara 
dfContratos["CREDOR"] = np.select(condicoesCarteiras, resultadosCarteiras, default=credor)
print("Projetos separados e renomeados")

# Atualizando a coluna CREDOR da aba de telefones
dfTelefones = pd.merge(dfTelefones,dfContratos[["CPF/CNPJ","CREDOR"]],how="left", on="CPF/CNPJ")

# Retirando as duplicadas 
dfTelefones = dfTelefones.drop_duplicates(subset=["CPF/CNPJ", "NUMERO","CREDOR"])

# Criando uma lista de credores encontrados
credores = dfContratos["CREDOR"].unique().tolist()

# Separando e salvando as planilhas conforme os credores
for credor in credores:
    dfContratosCredor = dfContratos[dfContratos["CREDOR"] == credor]
    dfTelefonesCredor = dfTelefones[dfTelefones["CREDOR"] == credor]

    with pd.ExcelWriter(f"{path}/output/{credor} {hoje}.xlsx", engine="openpyxl") as writer:
        dfTelefonesCredor[listaColunasTelefones].to_excel(writer, sheet_name="Telefones", index=False)
        dfContratosCredor[listaColunasContratos].to_excel(writer, sheet_name="Contratos", index=False)
        

print("Finalizado")

        
