import pandas as pd
import datetime as dt

pathSMS_RCS = "C:/Users/OT/Documents/Projetos/SeparadorDeRCS_SMS/input/"
pathWhatsApp = "C:/Users/OT/Documents/Projetos/SeparadorDeWhatsApp/input/"
pathCodigo = "C:/Users/OT/Documents/Projetos/DocumentadorAcoesInnovare/"

loginOperador = "MARCO.PORTILHO"

print("Lendo Bases ...",end="", flush=True)
baseWhatsApp = pd.read_excel(f"{pathWhatsApp}PLANILHA AÇÃO WHATSAPP.xlsx")
baseSMS_RCS = pd.read_excel(f"{pathSMS_RCS}PLANILHA AÇÃO SMS_RCS.xlsx")
print(" ✅", flush=True)

layoutImportacao = [
    "CPF/CNPJ",
    "CONTRATO",
    "DATA",
    "EVENTO",
    "DESCRIÇÃO",
    "OPERADOR"
]

print("Processando arquivo de saída ...",end="", flush=True)
hoje = dt.datetime.today()
baseWhatsApp = baseWhatsApp[baseWhatsApp["DATA ENVIO"].between(hoje - dt.timedelta(days=7), hoje)]
baseSMS_RCS = baseSMS_RCS[baseSMS_RCS["DATA ENVIO"].between(hoje - dt.timedelta(days=7), hoje)]

dfEventosWhatsApp = pd.DataFrame(columns=layoutImportacao)
dfEventosSMS_RCS = pd.DataFrame(columns=layoutImportacao)

dfEventosWhatsApp["CPF/CNPJ"] = baseWhatsApp["CPF/CNPJ"]
dfEventosSMS_RCS["CPF/CNPJ"] = baseSMS_RCS["CPF/CNPJ"]

dfEventosWhatsApp = pd.merge(dfEventosWhatsApp,baseWhatsApp,how="left",on="CPF/CNPJ")
dfEventosSMS_RCS = pd.merge(dfEventosSMS_RCS,baseSMS_RCS,how="left",on="CPF/CNPJ")

dfEventosWhatsApp["DATA"] = dfEventosWhatsApp["DATA ENVIO"].apply(lambda x: x.replace(hour=10, minute=30))
dfEventosSMS_RCS["DATA"] = dfEventosSMS_RCS["DATA ENVIO"].apply(lambda x: x.replace(hour=10, minute=30))

dfEventosWhatsApp["EVENTO"] = "Envio de WhatsApp"
dfEventosSMS_RCS["EVENTO"] = "Envio de RCS - fornecedor externo"

dfEventosWhatsApp["DESCRIÇÃO"] = dfEventosWhatsApp["TELEFONE"]
dfEventosSMS_RCS["DESCRIÇÃO"] = dfEventosSMS_RCS["TELEFONE"]

dfEventosWhatsApp["OPERADOR"] = loginOperador
dfEventosSMS_RCS["OPERADOR"] = loginOperador
print(" ✅", flush=True)

print("Concatenando e salvando a base de importação ...",end="", flush=True)
dfEventosGeral = pd.concat([dfEventosSMS_RCS,dfEventosWhatsApp])

dfEventosGeral = dfEventosGeral.drop(columns=dfEventosGeral.columns[dfEventosGeral.columns.get_loc("OPERADOR") + 1:])

dfEventosGeral.to_excel(f"{pathCodigo}output/Ações Planejamento {hoje.strftime("%d_%m_%Y")}.xlsx", index=False)
print(" ✅", flush=True)

print("Finalizado", flush=True)