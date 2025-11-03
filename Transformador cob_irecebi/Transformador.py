import pandas as pd
import numpy as np
import glob as gb
import datetime as dt

pathInput = "C:/Users/OT/Documents/Projetos/Transformador cob_irecebi/Bases/*.xlsx"
pathOutput = "C:/Users/OT/Documents/Projetos/Transformador cob_irecebi/Resultado/"

cabecalho = pd.read_excel("Layout.xlsx")

def salvar_em_lotes(df, tamanho_lote, nome_Lote):
    for i in range(0, len(df), tamanho_lote):
        df.columns = range(df.shape[1])
        df.columns = cabecalho.columns.tolist()
        lote = df.iloc[i:i + tamanho_lote].copy() 
        df_final = pd.concat([cabecalho, lote], ignore_index=True,axis=0)
        df_final.columns = ["" if col.startswith("Unnamed:") else col for col in df_final.columns]
        df_final.to_excel(f"{pathOutput}{nome_Lote}_{i // tamanho_lote + 1}.xlsx", index=False,sheet_name="BASE_IMPORTAÇÂO") 

def leitorDeExcel(sheetName):
    dataframes = []

    for arquivo in gb.glob(pathInput):
        try:
            df = pd.read_excel(arquivo,sheet_name=sheetName,dtype={"CPF/CNPJ" : str})
            dataframes.append(df)
            print(f"Arquivo {arquivo}, planilha {sheetName} processado!")
        except Exception as e:
            print(f"Arquivo {arquivo} não foi processado e retornou o erro {e}")

    return pd.concat(dataframes, ignore_index=True)

hoje = dt.datetime.now().date()
hoje = hoje.strftime("%d/%m/%Y")

baseTelefones = leitorDeExcel("Telefones")
baseEmails = leitorDeExcel("Emails")
baseContratos = leitorDeExcel("Contratos")
baseProdutos = leitorDeExcel("Produtos")

# Adicionando coluna descrição caso não tenha
if not "Descrição" in baseProdutos.columns:
    baseProdutos.insert(5,"Descrição", None)

# Limpando Inativos/BlockList/Números com menos de 10 digitos
baseTelefones = baseTelefones[
    (baseTelefones["BLOCKLIST"] != "SIM") &
    (baseTelefones["ATIVO"] != "NÃO") &
    (baseTelefones["NUMERO"].astype(str).str.len() >= 11)
]
# Arrumando os scores
scores = {9999: 6, 999: 6, 99: 6, 9: 6}
baseTelefones["SCORE"] = baseTelefones["SCORE"].replace(scores).fillna(3)
# Ordenando telefones
baseTelefones.sort_values(
    by=["CONTATO", "SCORE"],
    ascending=[False, True],
    inplace=True,
    kind='mergesort'
)
baseTelefones.drop_duplicates(subset=["CPF/CNPJ", "NUMERO"],inplace=True)
# Colocando os numeros na horizontal
baseTelefones = baseTelefones.groupby("CPF/CNPJ")["NUMERO"].apply(list).reset_index()
maxNumeros = baseTelefones["NUMERO"].str.len().max()
listaDeColunas = [f"NUMERO{i+1}" for i in range(maxNumeros)]
numerosHorizontal = pd.DataFrame(baseTelefones["NUMERO"].tolist(), columns=listaDeColunas)
baseTelefones = pd.concat([baseTelefones["CPF/CNPJ"], numerosHorizontal], axis=1)

# Colocando os emails na horizontal
baseEmails = baseEmails[baseEmails["ATIVO"] == "SIM"]
baseEmails = baseEmails.drop_duplicates(subset=["CPF/CNPJ","EMAIL"])
baseEmails = baseEmails.groupby("CPF/CNPJ")["EMAIL"].apply(list).reset_index()
maxEmails = baseEmails["EMAIL"].str.len().max()
listaDeColunas = [f"EMAIL{i+1}" for i in range(maxEmails)]
emailsHorizontal = pd.DataFrame(baseEmails["EMAIL"].tolist(), columns=listaDeColunas)
baseEmails = pd.concat([baseEmails["CPF/CNPJ"], emailsHorizontal], axis=1)

colunasBaseIrecebi = [
    "TIPO PF/PJ", "CNPJ /CNPJ", "NOME", "NÚMERO DO CONTRATO", "NÚMERO DA PARCELA",
    "DATA DE EMISSÃO - RECIBO", "VENCIMENTO", "COMPETÊNCIA", "VALOR", "DETALHE DA PARCELA",
    "CARTEIRA CONTRATO", "CANAL DE PREFERÊNCIA", "EQUIPAMENTO / PRODUTO",
    "END. COBRANÇA (1 OU 2)", "PERÍODO COBRANÇA", "BLACKLIST", "TELEFONE 1", "E-MAIL 1",
    "ENDEREÇO 1", "BAIRRO 1", "CIDADE 1", "UF 1", "CEP 1", "TELEFONE 2", "E-MAIL 2",
    "ENDEREÇO 2", "BAIRRO 2", "CIDADE 2", "UF 2", "CEP  2", "TELEFONE 3", "E-MAIL 3",
    "ENDEREÇO 3", "BAIRRO 3", "CIDADE 3", "UF 3", "CEP  3", "TELEFONE 4", "E-MAIL 4",
    "ENDEREÇO 4", "BAIRRO 4", "CIDADE 4", "UF 4", "CEP 4"
]

baseIrecebi = pd.DataFrame(columns=colunasBaseIrecebi)

baseIrecebi.rename(columns={"CNPJ /CNPJ": "CPF/CNPJ"},inplace=True)

baseIrecebi["CPF/CNPJ"] = baseContratos["CPF/CNPJ"]
baseIrecebi = pd.merge(baseIrecebi,baseContratos[["CPF/CNPJ","CLIENTE","ATRASO","TOTAL ABERTO","CONTRATO","CREDOR"]],on="CPF/CNPJ",how="inner")
baseIrecebi = pd.merge(baseIrecebi,baseProdutos[["CONTRATO","TIPO DE CONTRATO","Descrição"]],on="CONTRATO",how="inner")
baseIrecebi = pd.merge(baseIrecebi,baseEmails,on="CPF/CNPJ",how="inner")
baseIrecebi = pd.merge(baseIrecebi,baseTelefones,on="CPF/CNPJ",how="inner")

baseIrecebi["NÚMERO DO CONTRATO"] = baseIrecebi["CONTRATO"]
baseIrecebi["NÚMERO DA PARCELA"] = 1
baseIrecebi["DATA DE EMISSÃO - RECIBO"] = hoje
baseIrecebi["VENCIMENTO"] = hoje
baseIrecebi["VALOR"] = baseIrecebi["TOTAL ABERTO"]
baseIrecebi["DETALHE DA PARCELA"] = baseIrecebi["ATRASO"]

baseIrecebi["TELEFONE 1"] = baseIrecebi["NUMERO1"]
baseIrecebi["TELEFONE 2"] = baseIrecebi["NUMERO2"]
baseIrecebi["TELEFONE 3"] = baseIrecebi["NUMERO3"]
baseIrecebi["TELEFONE 4"] = baseIrecebi["NUMERO4"]

baseIrecebi["E-MAIL 1"] = baseIrecebi["EMAIL1"]
baseIrecebi["E-MAIL 2"] = baseIrecebi["EMAIL2"]
baseIrecebi["E-MAIL 3"] = baseIrecebi["EMAIL3"]
baseIrecebi["E-MAIL 4"] = baseIrecebi["EMAIL4"]

baseIrecebi["NOME"] = baseIrecebi["CLIENTE"]

credor = baseIrecebi["CREDOR"].str.strip()
tipo_contrato = baseIrecebi["TIPO DE CONTRATO"]
descricao = baseIrecebi["Descrição"]
atraso = baseIrecebi["ATRASO"]

condicoesCarteiras = [
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(1, 30)),
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(31, 94)),
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso >= 95),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "PF"),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "Consignado"),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "MEI")
]
resultadosCarteiras = [
    "NEON AMIGAVEL 1",
    "NEON AMIGAVEL 2",
    "NEON CRELIQ",
    "NEON EMPRESTIMO",
    "NEON CONSIGNADO",
    "NEON EMPRESTIMO MEI"
]
baseIrecebi["CARTEIRA CONTRATO"] = np.select(condicoesCarteiras, resultadosCarteiras, default=credor)
print("Carteiras separadas")

condicoesPF_PJ = [
    baseIrecebi["CPF/CNPJ"].str.len() == 14,
    baseIrecebi["CPF/CNPJ"].str.len() == 11
]
resultadosPF_PJ = [
    "PJ",
    "PF"
]
print("PF/PJ separadas")
baseIrecebi["TIPO PF/PJ"] = np.select(condicoesPF_PJ,resultadosPF_PJ,default="PF")

baseIrecebi.drop(columns=baseIrecebi.columns[(baseIrecebi.columns.get_loc("CEP 4") + 1) : baseIrecebi.shape[1]],inplace=True)

# Limpando nomes
baseIrecebi["NOME"] = baseIrecebi["NOME"].fillna("").astype(str).str.replace(r'\.', '', regex=True)
baseIrecebi["NOME"] = baseIrecebi["NOME"].str.split().apply(lambda x: " ".join([palavra for palavra in x if not palavra.isnumeric()]))

salvar_em_lotes(baseIrecebi,20000,"BASE CONSOLIDADA")

print("Finalizado")