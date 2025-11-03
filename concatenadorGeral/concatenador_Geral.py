import pandas as pd
import glob as gb
import datetime as dt

path = "C:/Users/OT/Documents/Projetos/concatenadorGeral/"

inputCSV = True

# Lê uma coletanea de arquivos excel a partir de uma planilha e os unifica
def leitorDeExcel():
    dataframes = []

    if(inputCSV):
        globPath = f"{path}input/*.csv"
    else:
        globPath = f"{path}input/*.xlsx"

    for arquivo in gb.glob(globPath):
        try:
            df = pd.read_csv(arquivo,dtype={"CPF/CNPJ" : str},sep=";",encoding="utf-8")
            dataframes.append(df)
            print(f"Arquivo {arquivo} processado")
        except Exception as e:
            print(f"Arquivo {arquivo} não foi processado e retornou o erro {e}")

    return pd.concat(dataframes, ignore_index=True)

df = leitorDeExcel()

if(inputCSV):
    df.to_csv(f"{path}output/baseConcatenada.csv",sep=";",encoding="utf-8",index=False)
else:
    df.to_excel(f"{path}output/baseConcatenada.xlsx",index=False)
print("Finalizado")