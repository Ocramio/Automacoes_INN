import pandas as pd
import numpy as np
import datetime as dt
import glob as gb
import calendar

#Gatilho de estrategias
cruzarComODB = 1  # =0 mantem todos os clientes; 
                  # =1 puxa todos que NÃO estão em ACIONADOS AGENTE VIRTUAL;
                  # !=1 e !=0 puxa todos que ESTÃO em ACIONADOS AGENTE VIRTUAL
cruzarComOBI = True
valorMinimo = True
salvarCsv = False
separarHot = False
separarHotPuro = False
scoreTier = True
dataNoMes = False

# Variavel para o local de armazenamento dos arquivos processados
path = "C:/Users/OT/Documents/Projetos/Automação/"

# Variavel com a data atual
hoje = dt.datetime.now()

# Função que verifica se a data está dentro do mês atual e caso não retorna o ultimo dia do mês atual
def estaNoMes(data):
    if((data.month == hoje.month and data.year == hoje.year) or not dataNoMes):
        return data.strftime("%d/%m/%Y")
    else:
        ultimoDiaUtil = dt.date(year=hoje.year, month=hoje.month, day=calendar.monthrange(hoje.year,hoje.month)[1])
        while ultimoDiaUtil.weekday() >= 5:
            ultimoDiaUtil -= dt.timedelta(days=1)
        return ultimoDiaUtil.strftime("%d/%m/%Y")
 
# Função que separa e salva em duas planilhas os clientes com contato SIM (HOT) e com contato NÃO (Não HOT)
def SepararHot(wsFull, nomeArquivo,colunasParaManter = []):
    wsHot = wsFull[wsFull["CONTATO"] == "SIM"].copy()
    if(separarHotPuro):
        wsNaoHot = wsFull[(wsFull["CONTATO"] == "NÃO") & (~wsFull.iloc[:, 0].isin(wsHot["CPF/CNPJ"]))].copy()
    else:
        wsNaoHot = wsFull[wsFull["CONTATO"] == "NÃO"].copy()

    wsHot = wsHot[colunasParaManter]
    wsNaoHot = wsNaoHot[colunasParaManter]

    if(salvarCsv):
        wsHot.to_csv(nomeArquivo + " HOT.csv", index=False,sep=";",encoding="utf-8", decimal=",")
        wsNaoHot.to_csv(nomeArquivo + ".csv", index=False,sep=";",encoding="utf-8",decimal=",")
    else:
        wsHot.to_excel(nomeArquivo + " HOT.xlsx", index=False)
        wsNaoHot.to_excel(nomeArquivo + ".xlsx", index=False)

# Função que ordena os números conforme estratégia pré-definida
def ordenarNumeros(wsFull):
    #Ordenando números
    if("ATRASO" in wsFull.columns):
        wsFull.sort_values(
            by=["CONTATO","SCORE","SCORE TIER","ATRASO"],
            ascending=[False, True, True,True],
            inplace=True,
            kind='mergesort'
        )
    else:
        wsFull.sort_values(
            by=["CONTATO","SCORE","SCORE TIER"],
            ascending=[False, True,True],
            inplace=True,
            kind='mergesort'
        )
        
    return wsFull

# Lê uma coletanea de arquivos excel a partir de uma planilha e os unifica
def leitorDeExcel(sheetName):
    dataframes = []

    for arquivo in gb.glob(f"{path}ARQUIVOS AUXILIARES/Bases/*.xlsx"):
        try:
            df = pd.read_excel(arquivo,sheet_name=sheetName,dtype={"CPF/CNPJ" : str})
            dataframes.append(df)
            print(f"Arquivo {arquivo}, planilha {sheetName} processado!")
        except Exception as e:
            print(f"Arquivo {arquivo} não foi processado e retornou o erro {e}")

    return pd.concat(dataframes, ignore_index=True)

# Função que salva os arquivos conforme layout de dicionarioMailings
def salvarArquivos(df, path, projeto):
    if(separarHot):
        if(salvarCsv):
            SepararHot(df,path,dicionarioMailings[f"{projeto} CSV"])
        else:
            SepararHot(df,path,dicionarioMailings[projeto])
            print(f"{projeto} separado")  
    else:
        if(salvarCsv):
            df = df[dicionarioMailings[projeto]]
            df.to_excel(f"{path}.csv", sep=";" , index=False, encoding="utf-8")
            print(f"{projeto} separado")
        else:
            df = df[dicionarioMailings[projeto]]
            df.to_excel(f"{path}.xlsx", index=False)
            print(f"{projeto} separado")

# Criando layout de cada projeto
dicionarioMailings = {
    "Amigavel 1":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "ATRASO", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "Amigavel 1 CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "ATRASO", "DATA2"],
    "Amigavel 2":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "ATRASO", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "Amigavel 2 CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "ATRASO", "DATA2"],
    "Creliq":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "ATRASO","SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "Creliq CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL"],
    "Emprestimo PF":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL","ATRASO","SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "Emprestimo PF CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL"],
    "Emprestimo MEI":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL","ATRASO","SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "Emprestimo MEI CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL"],
    "Consignado":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL","Marcadores", "Nome do Empregador","ATRASO","SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "Consignado CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL","Marcadores", "Nome do Empregador"],
    "E-Consignado":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL","Marcadores", "Nome do Empregador","ATRASO","SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "E-Consignado CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL","Marcadores", "Nome do Empregador"],
    "TODOS 1 A 3":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Divida", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "TODOS 1 A 3 CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Divida", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2"],
    "TODOS 4 A 6":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Divida", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "TODOS 4 A 6 MINIMO":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Minimo", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "TODOS 4 A 6 CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Divida", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2"],
    "TODOS 4 A 6 MINIMO CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Minimo", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2"],
    "TODOS NR":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Divida", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "TODOS NR CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Divida", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2"],
    "TODOS 7+":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor divida", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "TODOS 7+ MINIMO":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Minimo", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "TODOS 7+ CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor divida", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2"],
    "TODOS 7+ MINIMO CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "Valor Minimo", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2"],
    "VMK":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "VMK CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO", "EMAIL", "DATA2"],
    "GRENKE":["CPF/CNPJ", "CLIENTE", "NUMERO", "CREDOR", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO","EMAIL","PARCELADO", "SCORE", "CONTATO", "SCORE TIER", "ESTAGIO", "RANKING"],
    "GRENKE CSV":["CPF/CNPJ", "CLIENTE", "NUMERO", "FILIAL", "VALOR PRINCIPAL", "VALOR ATUALIZADO", "VALOR A FECHAR", "VALOR PARCELADO", "DATA DE VENCIMENTO","EMAIL", "PARCELADO"]
}

# Carregando o arquivo
wsTelefones = leitorDeExcel("Telefones")
wsContratos = leitorDeExcel("Contratos")
wsEmails = leitorDeExcel("Emails")
wsProdutos = leitorDeExcel("Produtos")
print("Bases lidas e concatenadas")

# Adicionando coluna descrição caso não tenha
if not "Descrição" in wsProdutos.columns:
    wsProdutos.insert(5,"Descrição", None)  

# Definindo data de vencimento padrão
diaDaSemana = hoje.weekday()
if(diaDaSemana == 4):
    diaDeVencimentoPadrao = (hoje + dt.timedelta(days=4))
elif(diaDaSemana == 5):
    diaDeVencimentoPadrao = (hoje + dt.timedelta(days=3))
elif(diaDaSemana == 3):
    diaDeVencimentoPadrao = (hoje + dt.timedelta(days=4))
else:
    diaDeVencimentoPadrao = (hoje + dt.timedelta(days=2))

diaDeVencimentoPadrao = estaNoMes(diaDeVencimentoPadrao)

# Limpando Inativos/BlockList/Números com menos de 10 digitos
wsTelefones = wsTelefones[
    (wsTelefones["BLOCKLIST"] != "SIM") & 
    (wsTelefones["ATIVO"] != "NÃO") & 
    (wsTelefones["NUMERO"].astype(str).str.len() >= 11)
]

# Retirando emails inativados
wsEmails = wsEmails[
    (wsEmails["ATIVO"] != "NÃO")
]

# Mantendo apenas um email por cliente
wsEmails.drop_duplicates(subset=wsEmails.columns[0], keep="first", inplace=True)

# Arrumando os scores
scores = {9999: 6, 999: 6, 99: 6, 9: 6}
wsTelefones["SCORE"] = wsTelefones["SCORE"].replace(scores).fillna(3)

# Ordenando telefones
wsTelefones.sort_values(
    by=["CONTATO", "SCORE"],
    ascending=[False, True],
    inplace=True,
    kind='mergesort'
)

# Limpando nomes
wsTelefones["CLIENTE"] = wsTelefones["CLIENTE"].fillna("").astype(str).str.replace(r'\.', '', regex=True)
wsTelefones["CLIENTE"] = wsTelefones["CLIENTE"].str.split().apply(lambda x: " ".join([palavra for palavra in x if not palavra.isnumeric()]))

# Coletando informações da aba Contratos
wsTelefones = pd.merge(wsTelefones[["CPF/CNPJ","CLIENTE","NUMERO","SCORE","CONTATO"]],wsContratos[["CPF/CNPJ","CREDOR","TOTAL ABERTO","ATRASO","CONTRATO","ESTAGIO","FILIAL"]], on="CPF/CNPJ", how="left")
 
# Coletando informações da aba Produtos
wsTelefones = pd.merge(wsTelefones,wsProdutos[["CONTRATO","TIPO DE CONTRATO","Descrição"]], on="CONTRATO", how="left")

# Coletando informações da aba Emails
wsTelefones = pd.merge(wsTelefones,wsEmails[["CPF/CNPJ", "EMAIL"]], on="CPF/CNPJ" ,how="left")

# Coletando informação de Score tier de pagamento e realizando a conta do ranking
if(scoreTier):
    wsScoreTier = pd.read_csv(f"{path}/ARQUIVOS AUXILIARES/SCORE TIER.csv", sep= ";",dtype={"CPF/CNPJ Numerico": str})
    wsTelefones["CPF/CNPJ Numerico"] = wsTelefones["CPF/CNPJ"].astype(int)
    wsTelefones["CPF/CNPJ Numerico"] = wsTelefones["CPF/CNPJ Numerico"].astype(str)
    wsTelefones = pd.merge(wsTelefones,wsScoreTier[["CPF/CNPJ Numerico", "SCORE TIER"]], on="CPF/CNPJ Numerico",how="left")
    wsTelefones["SCORE TIER"] = wsTelefones["SCORE TIER"].fillna(3)
    wsTelefones["RANKING"] = (wsTelefones["SCORE"] + wsTelefones["SCORE TIER"])//2
else:
    wsTelefones.insert(loc=wsTelefones.columns.get_loc("ESTAGIO"),column="SCORE TIER", value="")
    wsTelefones.insert(loc=wsTelefones.columns.get_loc("ESTAGIO") + 1,column="RANKING", value=wsTelefones["SCORE"])

# Preenchendo quem não tem email por Null
wsTelefones["EMAIL"] = wsTelefones["EMAIL"].fillna("NULL")

# Retirando clientes com atraso negativo
wsTelefones = wsTelefones[wsTelefones["ATRASO"] > 0]

wsTelefones = wsTelefones.rename(columns={"TOTAL ABERTO":"VALOR PRINCIPAL", "Descrição":"DESCRIÇÃO"})

# Adicionando as colunas base
wsTelefones.insert(5, "VALOR ATUALIZADO",None)
wsTelefones.insert(6, "VALOR A FECHAR",None)
wsTelefones.insert(7, "VALOR PARCELADO",None)
wsTelefones.insert(8, "DATA DE VENCIMENTO",None)

# Estratégia limpando com os acionados 
if(cruzarComODB != 0):
    wsDB = pd.read_csv(f"{path}/ARQUIVOS AUXILIARES/ACIONADOS AGENTE VIRTUAL.csv",dtype={0: str}, header=None)
    if(cruzarComODB == 1):
        #Para pegar os que não foram acionados
        wsTelefones = wsTelefones[~(wsTelefones["CPF/CNPJ"].isin(wsDB.iloc[:,0]))]
    else:
        #Para pegar os que já foram acionados
        wsTelefones = wsTelefones[(wsTelefones["CPF/CNPJ"].isin(wsDB.iloc[:,0]))]

# Realizando contagem de CPF/CNPJ de cada carteira
qtdAmigavel1 = wsTelefones[(wsTelefones["ATRASO"] < 31) & (wsTelefones["TIPO DE CONTRATO"] == "Fatura")].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtdAmigavel2 = wsTelefones[(wsTelefones["ATRASO"] < 94) & (wsTelefones["ATRASO"] > 31) &(wsTelefones["TIPO DE CONTRATO"] == "Fatura")].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtdCreliq = wsTelefones[(wsTelefones["ATRASO"] > 94) & (wsTelefones["TIPO DE CONTRATO"] == "Fatura")].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtdEmprestimo = wsTelefones[wsTelefones["DESCRIÇÃO"] == "PF"].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtdEmprestimoMei = wsTelefones[wsTelefones["DESCRIÇÃO"]== "MEI"].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtdConsignado = wsTelefones[wsTelefones["DESCRIÇÃO"] == "Consignado"].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtdPaFixa = wsTelefones[wsTelefones["CREDOR"].str.contains("EMPREENDIMENTOS - 1 A 3|EMPREENDIMENTOS - 4 a 6|EMPREENDIMENTOS - NR", na=False)].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtdVMK = wsTelefones[wsTelefones["CREDOR"].str.contains("VMK", na=False)].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtd7Mais = wsTelefones[wsTelefones["CREDOR"].str.contains("7+", na=False)].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]
qtdGrenke = wsTelefones[wsTelefones["CREDOR"] == "GRENKE - NEW DEBTORS"].drop_duplicates(subset=wsTelefones.columns[0]).shape[0]

print(f"Quantidades de CPF/CNPJ: \n"
      f"Amigavel 1: {qtdAmigavel1}, "
      f"Amigavel 2: {qtdAmigavel2}, "
      f"Creliq: {qtdCreliq}, "
      f"Emprestimo: {qtdEmprestimo}, "
      f"Emprestimo MEI: {qtdEmprestimoMei}, "
      f"Consignado: {qtdConsignado}, "
      f"VMK: {qtdVMK}, "
      f"PA FIXA: {qtdPaFixa}, "
      f"7+: {qtd7Mais},"
      f"GRENKE: {qtdGrenke}")

# Realizando contagem de volume de cada carteira
qtdAmigavel1 = wsTelefones[(wsTelefones["ATRASO"] < 31) & (wsTelefones["TIPO DE CONTRATO"] == "Fatura")].shape[0]
qtdAmigavel2 = wsTelefones[(wsTelefones["ATRASO"] < 94) & (wsTelefones["ATRASO"] > 31) &(wsTelefones["TIPO DE CONTRATO"] == "Fatura")].shape[0]
qtdCreliq = wsTelefones[(wsTelefones["ATRASO"] > 94) & (wsTelefones["TIPO DE CONTRATO"] == "Fatura")].shape[0]
qtdEmprestimo = wsTelefones[wsTelefones["DESCRIÇÃO"] == "PF"].shape[0]
qtdEmprestimoMei = wsTelefones[wsTelefones["DESCRIÇÃO"] == "MEI"].shape[0]
qtdConsignado = wsTelefones[wsTelefones["DESCRIÇÃO"] == "Consignado"].shape[0]
qtdPaFixa = wsTelefones[wsTelefones["CREDOR"].str.contains("EMPREENDIMENTOS - 1 A 3|EMPREENDIMENTOS - 4 a 6|EMPREENDIMENTOS - NR", na=False)].shape[0]
qtdVMK = wsTelefones[wsTelefones["CREDOR"].str.contains("VMK", na=False)].shape[0]
qtd7Mais = wsTelefones[wsTelefones["CREDOR"].str.contains("7+", na=False)].shape[0]
qtdGrenke = wsTelefones[wsTelefones["CREDOR"] == "GRENKE - NEW DEBTORS"].shape[0]

print(f"Quantidades de Telefones: \n"
      f"Amigavel 1: {qtdAmigavel1}, "
      f"Amigavel 2: {qtdAmigavel2}, "
      f"Creliq: {qtdCreliq}, "
      f"Emprestimo: {qtdEmprestimo}, "
      f"Emprestimo MEI: {qtdEmprestimoMei}, "
      f"Consignado: {qtdConsignado}, "
      f"VMK: {qtdVMK}, "
      f"PA FIXA: {qtdPaFixa}, "
      f"7+: {qtd7Mais},"
      f"GRENKE: {qtdGrenke}")

# ===============================================================
# SEPARANDO AMIGAVEL 1
# ===============================================================
if(qtdAmigavel1 > 0):
    wsFiltrada = wsTelefones[(wsTelefones["ATRASO"] < 31) & (wsTelefones["TIPO DE CONTRATO"] == "Fatura")].copy()
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
    wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]

    # Retirando valores zerados
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]
    
    # Definindo vencimento personalizado
    if(diaDaSemana == 4):
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=3))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=4))
    elif(diaDaSemana == 5):
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=2))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=3))
    elif(diaDaSemana == 3):
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=4))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=5))
    elif(diaDaSemana == 2):
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=1))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=2))
    else:
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=2))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=3))
    
    diaDeVencimentoAmigavel = estaNoMes(diaDeVencimentoAmigavel)
    diaDeVencimento2Amigavel = estaNoMes(diaDeVencimento2Amigavel)

    # Aplicando Vencimento
    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoAmigavel
    wsFiltrada.insert(11, "DATA2", diaDeVencimento2Amigavel)

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)

    salvarArquivos(wsFiltrada,f"{path}RESULTADOS/NEON/AMIGAVEL 1/NEON AMIGAVEL 1","Amigavel 1")

# ===============================================================
# SEPARANDO AMIGAVEL 2
# ===============================================================
if(qtdAmigavel2 > 0):
    wsFiltrada = wsTelefones[(wsTelefones["ATRASO"] < 94) & (wsTelefones["ATRASO"] > 31) &(wsTelefones["TIPO DE CONTRATO"] == "Fatura")].copy()
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
    wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]

    # Retirando valores zerados
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

    if(diaDaSemana == 4):
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=3))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=4))
    elif(diaDaSemana == 5):
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=2))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=3))
    elif(diaDaSemana == 3):
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=4))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=5))
    elif(diaDaSemana == 2):
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=1))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=2))
    else:
        diaDeVencimentoAmigavel = (hoje + dt.timedelta(days=2))
        diaDeVencimento2Amigavel = (hoje + dt.timedelta(days=3))

    diaDeVencimentoAmigavel = estaNoMes(diaDeVencimentoAmigavel)
    diaDeVencimento2Amigavel = estaNoMes(diaDeVencimento2Amigavel)

    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoAmigavel
    wsFiltrada.insert(11, "DATA2", diaDeVencimento2Amigavel)

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)
    
    salvarArquivos(wsFiltrada,f"{path}RESULTADOS/NEON/AMIGAVEL 2/NEON AMIGAVEL 2","Amigavel 2")

# ===============================================================
# SEPARANDO CRELIQ
# ===============================================================
if (qtdCreliq > 0):
    # Filtrando Creliq
    wsFiltrada = wsTelefones[(wsTelefones["ATRASO"] > 94) & (wsTelefones["TIPO DE CONTRATO"] == "Fatura")].copy()

    # Definindo as condições para os dias de atraso
    faixaDeDescontoCreliq = [
    (wsFiltrada["ATRASO"] < 101),
    (wsFiltrada["ATRASO"] < 151),
    (wsFiltrada["ATRASO"] < 201),
    (wsFiltrada["ATRASO"] >= 201)
    ]

    # Definindo os descontos correspondentes
    descontosCreliq = [0.95, 0.9, 0.8, 0.65]
    # Calculando o valor com desconto diretamente
    wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"] * np.select(faixaDeDescontoCreliq, descontosCreliq, default=0)
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR A FECHAR"]/2
    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoPadrao

    # Retirando valores zerados
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)

    salvarArquivos(wsFiltrada,f"{path}RESULTADOS/NEON/CRELIQ/NEON CRELIQ","Creliq")

# ===============================================================
# SEPARANDO EMPRESTIMO PF
# ===============================================================
if (qtdEmprestimo > 0):
 
    wsFiltrada = wsTelefones[wsTelefones["DESCRIÇÃO"] == "PF"].copy()

    # Calculando o valor com desconto diretamente
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
    wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoPadrao

    # Retirando valores zerados
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)

    salvarArquivos(wsFiltrada,f"{path}RESULTADOS/NEON/EMPRESTIMO/NEON EMPRESTIMO PF","Emprestimo PF")

# ===============================================================
# SEPARANDO EMPRESTIMO MEI
# ===============================================================
if (qtdEmprestimoMei > 0):
 
    wsFiltrada = wsTelefones[wsTelefones["DESCRIÇÃO"] == "MEI"].copy()

    # Calculando o valor com desconto diretamente
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
    wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoPadrao

    # Retirando valores vazios
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)

    salvarArquivos(wsFiltrada,f"{path}RESULTADOS/NEON/EMPRESTIMO/NEON EMPRESTIMO MEI","Emprestimo MEI")

# ===============================================================
# SEPARANDO CONSIGNADO
# ===============================================================
if (qtdConsignado > 0):

    wsFiltrada = wsTelefones[wsTelefones["DESCRIÇÃO"] == "Consignado"].copy()

    wsFundosEmpresas = pd.read_excel(f"{path}/ARQUIVOS AUXILIARES/Base Consignado.xlsx",dtype={"CPF/CNPJ": str})

    wsFiltrada = pd.merge(wsFiltrada, wsFundosEmpresas[["CPF/CNPJ","Marcadores","Nome do Empregador"]],on=["CPF/CNPJ"], how="left")

    wsFiltrada = wsFiltrada.dropna(subset=["Marcadores","Nome do Empregador"])

    # Verificando se tem consignado
    if(wsFiltrada.shape[0] > 0):

        #Arrumando os nomes dos fundos e retirando não encontrados
        wsFiltrada = wsFiltrada.dropna(subset= wsFiltrada.columns[10])

        wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
        wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
        wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]
        wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoPadrao

        # Retirando valores vazios
        wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

        #Ordenando números
        ordenarNumeros(wsFiltrada)

        # Retirando linhas com CPF e Número duplicados
        wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)    

        salvarArquivos(wsFiltrada,f"{path}RESULTADOS/NEON/CONSIGNADO/NEON CONSIGNADO","Consignado")

# ===============================================================
# SEPARANDO E-CONSIGNADO
# ===============================================================
if (qtdConsignado > 0):

    wsFiltrada = wsTelefones[wsTelefones["DESCRIÇÃO"] == "Consignado"].copy()
    
    wsFundosEmpresas = pd.read_excel(f"{path}/ARQUIVOS AUXILIARES/Base eConsignado.xlsx",dtype={"CPF/CNPJ": str})
    wsFiltrada = pd.merge(wsFiltrada, wsFundosEmpresas[["CPF/CNPJ","Marcadores","Nome do Empregador"]],on=["CPF/CNPJ"], how="left")
    
    wsFiltrada = wsFiltrada.dropna(subset=["Marcadores","Nome do Empregador"])

    # Verificando se tem e-consignado
    if(wsFiltrada.shape[0] > 0):

        wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
        wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
        wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]
        wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoPadrao

        # Retirando valores vazios
        wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

        #Ordenando números
        ordenarNumeros(wsFiltrada)

        # Retirando linhas com CPF e Número duplicados
        wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)    

        salvarArquivos(wsFiltrada,f"{path}RESULTADOS/NEON/CONSIGNADO/NEON E-CONSIGNADO","E-Consignado")


# ===============================================================
# SEPARANDO VMK
# ===============================================================
if (qtdVMK > 0):

    wsFiltrada = wsTelefones[wsTelefones["CREDOR"].str.strip().str.contains("VMK")].copy()

    # Arrumando dia de vencimento
    if(diaDaSemana == 5):
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=2))
    elif(diaDaSemana == 4):
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=3))
    else:
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=1))
    
    diaDeVencimentoCartaoDeTodos = estaNoMes(diaDeVencimentoCartaoDeTodos)

    # Arrumando os valores
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
    wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoCartaoDeTodos
    wsFiltrada.insert(10,"DATA2",diaDeVencimentoPadrao)

    # Retirando valores vazios
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)

    salvarArquivos(wsFiltrada,f"{path}RESULTADOS/CARTÃO DE TODOS/VMK","VMK")

# ===============================================================
# SEPARANDO TODOS PA FIXA (1 A 3 | 4 A 6 | NR)
# ===============================================================
if (qtdPaFixa > 0):

    wsFiltrada = wsTelefones[wsTelefones["CREDOR"].str.strip().str.contains("EMPREENDIMENTOS - 1 A 3|EMPREENDIMENTOS - 4 a 6|EMPREENDIMENTOS - NR", na=False)].copy()

    if(cruzarComOBI):
        wsPaFixa = pd.read_excel(f"{path}/ARQUIVOS AUXILIARES/BASE TODOS PA FIXA.xlsx",dtype={"CPF": str})

        wsPaFixa.rename(columns={"CPF": "CPF/CNPJ"},inplace=True)
        wsPaFixa["CPF/CNPJ"] = wsPaFixa["CPF/CNPJ"].replace({"-": "",r"\.":""},regex=True)
        wsPaFixa["Valor Divida"] = wsPaFixa["Valor Divida"].replace({r"R\$": "", r"\.":"", ",":"."}, regex=True)
        wsPaFixa["Valor Divida"] = wsPaFixa["Valor Divida"].astype(float)
        wsPaFixa["Valor Minimo"] = wsPaFixa["Valor Minimo"].replace({r"R\$": "", r"\.":"", ",":"."}, regex=True)
        wsPaFixa["Valor Minimo"] = wsPaFixa["Valor Minimo"].astype(float)

        wsFiltrada = pd.merge(wsFiltrada, wsPaFixa[["CPF/CNPJ","Valor Divida", "Valor Minimo"]],on="CPF/CNPJ", how="left")
        wsFiltrada = wsFiltrada.dropna(subset="Valor Divida")
        wsFiltrada["Valor Minimo"] = wsFiltrada["Valor Minimo"].fillna(wsFiltrada["Valor Divida"])

    # Arrumando dia de vencimento
    if(diaDaSemana == 5):
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=2))
    elif(diaDaSemana == 4):
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=3))
    else:
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=1))

    diaDeVencimentoCartaoDeTodos = estaNoMes(diaDeVencimentoCartaoDeTodos)

    # Arrumando os valores
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
    wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoCartaoDeTodos
    wsFiltrada.insert(10,"DATA2",diaDeVencimentoPadrao)

    # Retirando valores vazios
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)
    
    # Separando 1 A 3
    wsTodos1A3 = wsFiltrada[wsFiltrada["CREDOR"].str.strip().str.contains("EMPREENDIMENTOS - 1 A 3", na=False)].copy()
    if(wsTodos1A3.shape[0] > 0):
        salvarArquivos(wsTodos1A3,f"{path}RESULTADOS/CARTÃO DE TODOS/TODOS OS EMPREENDIMENTOS 1 A 3","TODOS 1 A 3")
    
    # Separando NR
    wsTodosNr = wsFiltrada[wsFiltrada["CREDOR"].str.strip().str.contains("EMPREENDIMENTOS - NR", na=False)].copy()
    if(wsTodosNr.shape[0] > 0):
        salvarArquivos(wsTodosNr,f"{path}RESULTADOS/CARTÃO DE TODOS/TODOS OS EMPREENDIMENTOS NR","TODOS NR")

    # Separando 4 A 6
    wsTodos4A6 = wsFiltrada[wsFiltrada["CREDOR"].str.strip().str.contains("EMPREENDIMENTOS - 4 a 6", na=False)].copy()
    if(wsTodos4A6.shape[0] > 0):
        if(valorMinimo):
            salvarArquivos(wsTodos4A6,f"{path}RESULTADOS/CARTÃO DE TODOS/TODOS OS EMPREENDIMENTOS 4 A 6","TODOS 4 A 6 MINIMO")
        else:
            salvarArquivos(wsTodos4A6,f"{path}RESULTADOS/CARTÃO DE TODOS/TODOS OS EMPREENDIMENTOS 4 A 6","TODOS 4 A 6")

# ===============================================================
# SEPARANDO TODOS 7+
# ===============================================================
if (qtd7Mais > 0):

    wsFiltrada = wsTelefones[wsTelefones["CREDOR"].str.strip().str.contains("7+")].copy()

    # Puxando os valores do BI
    if(cruzarComOBI):
        ws7Mais = pd.read_excel(f"{path}/ARQUIVOS AUXILIARES/BASE TODOS 7+.xlsx",dtype={"CPF": str})

        ws7Mais.rename(columns={"CPF": "CPF/CNPJ"},inplace=True)
        ws7Mais["Valor divida"] = ws7Mais["Valor divida"].replace({r"R\$": "", r"\.":"", ",":"."}, regex=True)
        ws7Mais["Valor divida"] = ws7Mais["Valor divida"].astype(float)
        ws7Mais["Valor Minimo"] = ws7Mais["Valor Minimo"].replace({r"R\$": "", r"\.":"", ",":"."}, regex=True)
        ws7Mais["Valor Minimo"] = ws7Mais["Valor Minimo"].astype(float)

        wsFiltrada = pd.merge(wsFiltrada, ws7Mais[["CPF/CNPJ","Valor divida", "Valor Minimo"]],on="CPF/CNPJ", how="left")
        wsFiltrada = wsFiltrada.dropna(subset="Valor divida")
        wsFiltrada["Valor Minimo"] = wsFiltrada["Valor Minimo"].fillna(wsFiltrada["Valor divida"])
    
    # Arrumando dia de vencimento
    if(diaDaSemana == 5):
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=2))
    elif(diaDaSemana == 4):
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=3))
    else:
        diaDeVencimentoCartaoDeTodos = (hoje + dt.timedelta(days=1))

    diaDeVencimentoCartaoDeTodos = estaNoMes(diaDeVencimentoCartaoDeTodos)

    # Arrumando os valores
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR PRINCIPAL"] * 1.1
    wsFiltrada["VALOR A FECHAR"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR PRINCIPAL"]
    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoCartaoDeTodos 
    wsFiltrada.insert(10,"DATA2",diaDeVencimentoPadrao)

    # Retirando valores vazios
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)

    if(valorMinimo):
        salvarArquivos(wsFiltrada,f"{path}RESULTADOS/CARTÃO DE TODOS/TODOS OS EMPREENDIMENTOS 7+","TODOS 7+ MINIMO")
    else:
        salvarArquivos(wsFiltrada,f"{path}RESULTADOS/CARTÃO DE TODOS/TODOS OS EMPREENDIMENTOS 7+","TODOS 7+")

# ===============================================================
# SEPARANDO GRENKE
# ===============================================================
if (qtdGrenke > 0):
    
    wsFiltrada = wsTelefones[wsTelefones["CREDOR"].str.strip() == "GRENKE - NEW DEBTORS"].copy()

    # Arrumando os valores
    totalAbertoPorCNPJ = wsContratos[wsContratos["CREDOR"].str.strip() == "GRENKE - NEW DEBTORS"].groupby("CPF/CNPJ")["TOTAL ABERTO"].sum()
    maiorAtrasoPorCNPJ = wsContratos[wsContratos["CREDOR"].str.strip() == "GRENKE - NEW DEBTORS"].sort_values(by="ATRASO",ascending=False).drop_duplicates(subset="CPF/CNPJ",keep="first")
    wsFiltrada = pd.merge(wsFiltrada,totalAbertoPorCNPJ,how="left",on="CPF/CNPJ")
    wsFiltrada = pd.merge(wsFiltrada,maiorAtrasoPorCNPJ[["CPF/CNPJ","ATRASO"]],how="left",on="CPF/CNPJ",suffixes=("_x",""))
    wsFiltrada["VALOR PRINCIPAL"] = wsFiltrada["TOTAL ABERTO"]
    wsFiltrada["VALOR A FECHAR"] = (wsFiltrada["VALOR PRINCIPAL"] * 1.1) + (wsFiltrada["VALOR PRINCIPAL"] * (0.01/30)) * wsFiltrada["ATRASO"]
    wsFiltrada["VALOR ATUALIZADO"] = wsFiltrada["VALOR A FECHAR"]
    wsFiltrada["VALOR PARCELADO"] = wsFiltrada["VALOR A FECHAR"]
    wsFiltrada["DATA DE VENCIMENTO"] = diaDeVencimentoPadrao
    wsFiltrada["PARCELADO"] = 0
    wsFiltrada["CREDOR"] = "GC LOCACAO DE EQUIPAMENTOS GRENKE"

    # Retirando valores vazios
    wsFiltrada = wsFiltrada[wsFiltrada["VALOR PRINCIPAL"] > 0]

    #Ordenando números
    ordenarNumeros(wsFiltrada)

    # Retirando linhas com CPF e Número duplicados
    wsFiltrada.drop_duplicates(subset=["CPF/CNPJ","NUMERO"], keep="first", inplace=True)

    salvarArquivos(wsFiltrada,f"{path}RESULTADOS/GRENKE/GRENKE","GRENKE")

print("Finalizado")
