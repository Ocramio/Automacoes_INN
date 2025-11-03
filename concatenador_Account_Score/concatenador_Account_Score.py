import pandas as pd
import glob as gb
import datetime as dt

path = "C:/Users/OT/Documents/Projetos/concatenador_Account_Score/"

hoje = dt.date.today().strftime("%d_%m_%Y")

# Lê uma coletanea de arquivos excel a partir de uma planilha e os unifica
def leitorDeExcel():
    dataframes = []

    for arquivo in gb.glob(f"{path}input/*.csv"):
        try:
            df = pd.read_csv(arquivo,dtype={"CPF/CNPJ" : str},sep=";",encoding="utf-8")
            dataframes.append(df)
            print(f"Arquivo {arquivo} processado")
        except Exception as e:
            print(f"Arquivo {arquivo} não foi processado e retornou o erro {e}")

    return pd.concat(dataframes, ignore_index=True)

dfAccount_Score = leitorDeExcel()

dfAccount_Score.sort_values(by=["ScoreTier"], ascending=[True],inplace=True,kind="mergesort")

dfAccount_Score.drop_duplicates(subset=["CPF_CNPJ"],inplace=True)

dfAccount_Score.rename(columns={"CPF_CNPJ": "CPF/CNPJ Numerico", "ScoreTier": "SCORE TIER"},inplace=True)

dfAccount_Score.to_csv(f"{path}output/AccountScore_{hoje}.csv",sep=";",encoding="utf-8",index=False)

print("Finalizado")