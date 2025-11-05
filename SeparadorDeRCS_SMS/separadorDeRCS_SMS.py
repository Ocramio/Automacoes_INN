import pandas as pd
import numpy as np
import datetime as dt
import glob as gb

# Variavel para o local de armazenamento dos arquivos processados
path = "C:/Users/OT/Documents/Projetos/SeparadorDeRCS_SMS/"

# Variavel para definir entre primeiro nome ou nome completo (Não aplicável a empresas)
primeiroNome = True

# Lendo o cronograma
cronograma = pd.read_excel(f"{path}input/Cronograma.xlsx",dtype={"CRUZAR":int})
cronograma = cronograma.values

# Lendo base e-consignado
dfEConsignado = pd.read_excel(f"{path}input/Base eConsignado.xlsx", dtype={"CPF/CNPJ": str})
dfConsignado = pd.read_excel(f"{path}input/Base Consignado.xlsx",dtype={"CPF/CNPJ":str})

# Puxando data do dia
hoje = dt.date.today()

try:
    envios = pd.read_excel(f"{path}input/PLANILHA AÇÃO SMS_RCS.xlsx",sheet_name="ENVIADOS",dtype={"CPF/CNPJ": str},parse_dates=["DATA ENVIO"])
    envios['DATA ENVIO'] = envios['DATA ENVIO'].dt.date
except Exception:
    print("Planilha de log não existente, será criada uma vazia")
    enviosColunas = [
        "CPF/CNPJ",
        "NOME",
        "TELEFONE",
        "DATA ENVIO",
        "PROJETO",
        "TIPO DE ENVIO"
    ]
    envios = pd.DataFrame(columns=enviosColunas)

# Organiza a planilha de "melhores" a "piores" números/clientes
def ordenarNumeros(wsFull):
    #Ordenando números
    wsFull.sort_values(
        by=["CONTATO","SCORE","SCORE TIER", "ATRASO"],
        ascending=[False, True, True, True],
        inplace=True,
        kind='mergesort'
    )
    return wsFull

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

# Salva os disparos no df envios para salvar na Planilha de envios
def salvarLogEnvio(df,tipoDisparo):
    global envios

    hoje = dt.date.today()
    dfColunas = [
        "CPF/CNPJ",
        "NOME",
        "TELEFONE",
        "DATA ENVIO",
        "PROJETO",
        "TIPO DE ENVIO"
    ]
    dfLog = pd.DataFrame(columns=dfColunas)
    dfLog["NOME"] = df["CLIENTE"]
    dfLog["TELEFONE"] = df["NUMERO"]
    dfLog["CPF/CNPJ"] = df["CPF/CNPJ"]
    dfLog["PROJETO"] = df["CARTEIRA CONTRATO"]
    dfLog["DATA ENVIO"] = hoje
    dfLog["TIPO DE ENVIO"] = tipoDisparo
    envios = pd.concat([envios,dfLog])

# Separa os lotes de RCS
def separarLotesRCS(df, nomeArquivo):
    dfLayoutLote = [
        "numero",
        "var1",
        "var2",
        "var3",
        "var4",
        "var5"
    ]

    dfLote = pd.DataFrame(columns= dfLayoutLote)
    dfLote["numero"] = df["NUMERO"].astype(int)
    dfLote["var1"] = df["CLIENTE"]
    dfLote.to_csv(f"{path}/output/{nomeArquivo} RCS.csv",index=False,encoding="utf-8",sep=";")
    print(f"{nomeArquivo} separado!")

# Separa os lotes de SMS
def separarLotesSMS(df, nomeArquivo):
    dfLayoutLote = [
        "TELEFONE",
        "NOME",
        "LOJA"
    ]

    dfLote = pd.DataFrame(columns= dfLayoutLote)
    dfLote["TELEFONE"] = df["NUMERO"].astype(int)
    dfLote["NOME"] = df["CLIENTE"]
    dfLote.to_csv(f"{path}/output/{nomeArquivo} SMS.csv",index=False,encoding="utf-8",sep=";")
    print(f"{nomeArquivo} separado!")

# Retira os clientes que já foram disparados no periodo de 7 dias
def retirarDisparadosPeriodo(dfLeft, dfRight, usarPeriodo = True, colunaNumerosLeft = "NUMERO", colunaCPF_CNPJLeft = "CPF/CNPJ", 
                             colunaNumerosRight = "TELEFONE", colunaCPF_CNPJRight = "CPF/CNPJ"):
    if(usarPeriodo):
        enviosPeriodo = dfRight[dfRight["DATA ENVIO"].between(hoje-dt.timedelta(days=7),hoje)]
    else:
        enviosPeriodo = dfRight
    return dfLeft[(~dfLeft[colunaCPF_CNPJLeft].isin(enviosPeriodo[colunaCPF_CNPJRight])) & 
                  (~dfLeft[colunaNumerosLeft].isin(enviosPeriodo[colunaNumerosRight]))].copy()

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

# Arrumando os scores
scores = {9999: 6, 999: 6, 99: 6, 9: 6}
dfTelefones["SCORE"] = dfTelefones["SCORE"].replace(scores).fillna(3)

# Ordenando telefones
dfTelefones.sort_values(
    by=["CONTATO", "SCORE"],
    ascending=[False, True],
    inplace=True,
    kind='mergesort'
)

# Limpando nomes
dfTelefones["CLIENTE"] = dfTelefones["CLIENTE"].fillna("").astype(str).str.replace(r'\.', '', regex=True)
dfTelefones["CLIENTE"] = dfTelefones["CLIENTE"].str.split().apply(lambda x: " ".join([palavra for palavra in x if not palavra.isnumeric()]))

# Coletando informações da aba Contratos
dfTelefones = pd.merge(dfTelefones[["CPF/CNPJ", "CLIENTE", "NUMERO", "SCORE", "CONTATO"]],dfContratos[["CPF/CNPJ", "CREDOR", "ATRASO", "CONTRATO"]], on="CPF/CNPJ", how="left")

# Coletando informações da aba Produtos
dfTelefones = pd.merge(dfTelefones,dfProdutos[["CONTRATO", "TIPO DE CONTRATO", "Descrição"]], on="CONTRATO", how="left")

# Coletando informação de Score tier de pagamento e realizando a conta do ranking
dfScoreTier = pd.read_csv(f"{path}input/SCORE TIER.csv", sep= ";",dtype={"CPF/CNPJ Numerico": str})
dfTelefones["CPF/CNPJ Numerico"] = dfTelefones["CPF/CNPJ"].astype(int)
dfTelefones["CPF/CNPJ Numerico"] = dfTelefones["CPF/CNPJ Numerico"].astype(str)
dfTelefones = pd.merge(dfTelefones,dfScoreTier, on="CPF/CNPJ Numerico",how="left")
dfTelefones["SCORE TIER"] = dfTelefones["SCORE TIER"].fillna(3)
dfTelefones["RANKING"] = (dfTelefones["SCORE"] + dfTelefones["SCORE TIER"])//2
dfTelefones = dfTelefones.drop(columns="CPF/CNPJ Numerico")

# Retirando clientes com atraso negativo
dfTelefones = dfTelefones[dfTelefones["ATRASO"] > 0]

# Salvando informações em variáveis para aumentar a performance
credor = dfTelefones["CREDOR"].str.strip()
tipo_contrato = dfTelefones["TIPO DE CONTRATO"]
descricao = dfTelefones["Descrição"]
atraso = dfTelefones["ATRASO"]
cpf_cnpj = dfTelefones["CPF/CNPJ"]

# Separando as carteiras por numpy select
condicoesCarteiras = [
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(1, 30)),
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso.between(31, 94)),
    (credor == "NEON") & (tipo_contrato == "Fatura") & (atraso >= 95),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "PF"),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "Consignado") & (cpf_cnpj.isin(dfConsignado["CPF/CNPJ"])),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "Consignado") & (cpf_cnpj.isin(dfEConsignado["CPF/CNPJ"])),
    (credor == "NEON") & (tipo_contrato != "Fatura") & (descricao == "MEI")
]
resultadosCarteiras = [
    "NEON AMIGAVEL 1",
    "NEON AMIGAVEL 2",
    "NEON CRELIQ",
    "NEON EMPRESTIMO",
    "NEON CONSIGNADO",
    "NEON E-CONSIGNADO",
    "NEON EMPRESTIMO MEI"
]
dfTelefones["CARTEIRA CONTRATO"] = np.select(condicoesCarteiras, resultadosCarteiras, default=credor)
print("Carteiras separadas")

ordenarNumeros(dfTelefones)

dfTelefones.dropna(subset="CPF/CNPJ",inplace=True)
dfTelefones.dropna(subset="NUMERO",inplace=True)

dfTelefones = dfTelefones.drop_duplicates(subset=["CPF/CNPJ"])

if(primeiroNome):
    dfTelefones.loc[dfTelefones["CPF/CNPJ"].str.len() == 11, "CLIENTE"] = dfTelefones.loc[dfTelefones["CPF/CNPJ"].str.len() == 11, "CLIENTE"].str.split().str[0]

# Separando os lotes
for numero in range(0, len(cronograma)):
    nomeCredor = cronograma[numero][0].strip()
    if(isinstance(cronograma[numero][1], str)):
        qtdEnvios = cronograma[numero][1].strip()
    else:
        qtdEnvios = cronograma[numero][1]
    nomeArquivo = cronograma[numero][2]
    tipoEnvio = cronograma[numero][3].strip()
    cruzar = cronograma[numero][4]

    # Retirando clientes que já constam envios caso variavel seja 1
    if(cruzar == 1):
        dfLimpa = retirarDisparadosPeriodo(dfLeft=dfTelefones,dfRight=envios,usarPeriodo=True)
    else:
        enviosDoDia = envios[envios["DATA ENVIO"] == hoje]
        dfLimpa = retirarDisparadosPeriodo(dfLeft=dfTelefones,dfRight=enviosDoDia,usarPeriodo=False)

    dfBaseHot = dfLimpa[(dfLimpa["CARTEIRA CONTRATO"].str.strip() == nomeCredor) & (dfLimpa["CONTATO"] == "SIM")].copy()
    dfBaseNaoHot = dfLimpa[(dfLimpa["CARTEIRA CONTRATO"].str.strip() == nomeCredor) & (dfLimpa["CONTATO"] == "NÃO")].copy()

    ordenarNumeros(dfBaseHot)
    ordenarNumeros(dfBaseNaoHot)

    if(qtdEnvios == "FULL"):
        if((dfBaseHot.shape[0] + dfBaseNaoHot.shape[0]) <= 0):
            print(f"Não foi encontrado nenhum cliente para o numero {nomeArquivo} do projeto {nomeCredor}!\nAdicione uma base nova")
            continue
        qtdEnvios = dfBaseHot.shape[0] + dfBaseNaoHot.shape[0]
        print(f"Foi encontrado {qtdEnvios} cliente(s) para o numero {nomeArquivo} do projeto {nomeCredor}!")
    else:
        if ((dfBaseHot.shape[0] + dfBaseNaoHot.shape[0]) < qtdEnvios): 
            print(f"O numero {nomeArquivo} do projeto {nomeCredor} não possui uma base suficiente para {qtdEnvios} envios!\nAdicione uma base nova")
            continue

    if(cruzar == 0):
        df = retirarDisparadosPeriodo(dfLeft=dfBaseHot,dfRight=envios,usarPeriodo=True).iloc[:qtdEnvios].copy()
        resto = qtdEnvios - df.shape[0]
        if(resto > 0):
            df = pd.concat([df,retirarDisparadosPeriodo(dfLeft=dfBaseNaoHot,dfRight=envios,usarPeriodo=True).iloc[:resto]]).copy()
            resto = qtdEnvios - df.shape[0]
            if(resto > 0):
                df = pd.concat([df,retirarDisparadosPeriodo(dfLeft=dfBaseHot,dfRight=df,usarPeriodo=False,colunaNumerosRight="NUMERO").iloc[:resto]]).copy()
                resto = qtdEnvios - df.shape[0]
                if(resto > 0):
                    df = pd.concat([df,retirarDisparadosPeriodo(dfLeft=dfBaseNaoHot,dfRight=df,usarPeriodo=False,colunaNumerosRight="NUMERO").iloc[:resto]]).copy()
    else:
        if(dfBaseHot.shape[0] >= qtdEnvios):
            df = dfBaseHot.iloc[:qtdEnvios,:].copy()
        else:
            df = dfBaseHot.copy()
            df = pd.concat([df,dfBaseNaoHot.iloc[:(qtdEnvios - df.shape[0]),:]]).copy()

    salvarLogEnvio(df,tipoEnvio)
    if(tipoEnvio == "RCS"):
        separarLotesRCS(df,nomeArquivo)
    else:
        separarLotesSMS(df,nomeArquivo)
        
envios.to_excel(f"{path}input/PLANILHA AÇÃO SMS_RCS.xlsx",sheet_name="ENVIADOS",index=False)
print("Finalizado!")

        
