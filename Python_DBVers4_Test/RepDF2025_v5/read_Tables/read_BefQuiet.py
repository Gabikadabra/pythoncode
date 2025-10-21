#%%
#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
import pandas as pd
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task

#from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_rilasci import df_Rilasci_task
import os
import sys
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

from Util_logging import logger
wPythProc="read_BefQuiet"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()


#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

df_quiet=pd.read_excel(wrk_dir_inp+"\\"+"Da QuietanzePortale.xlsx",sheet_name=1, dtype={ })#"NUMERO BDO": str,"POSIZIONE": str,"NUMERO QUIETANZA": str,"NUMERO FATTURA": str,"NUMERO QUIETANZA": str,"DATA STATO":str,"PERIODO COMPETENZA SAL":str 
#Get ODAG Macro nbr
# ***Filter*** in testfilter is on and specific selection is in place
if (wrk_environm.lower()=="testfilter") & ("*all" not in listfilter_BDO):
    df_quiet=df_quiet.loc[(df_quiet["NUMERO BDO"].isin(listfilter_BDO))]

df_quiet["Mese competenza"]=df_quiet["PERIODO COMPETENZA SAL"].str.slice(5,7)
df_quiet["Anno competenza"]=df_quiet["PERIODO COMPETENZA SAL"].str.slice(0,4)

"""#begin format columns - reduce size"""
wLi=["NUMERO BDO","POSIZIONE","ID QUIETANZA PADRE",	"ID QUIETANZA FIGLIA","Anno competenza","Mese competenza"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["IMPORTO",	"COSTO SUBAPPALTO"];wDf={value:"float" for value in wLf}
wLc=["NOME STATO","MILESTONE","PERIODO COMPETENZA SAL","FORNITORE ROFI","SUBFORNITORE","ROFI","FORNITORE UTENTE CARICAMENTO","UTENTE CARICAMENTO","NOME STATO"," OWNER CAMBIO STATO","PROGR. STATO","ULTIMO STATO",	"ALLEGATO FORNITORE PRESENTE"];wDc={value:"category" for value in wLc}
wLnbr=wLi+wLf
df_quiet[wLnbr]=df_quiet[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDf,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_quiet= df_quiet.astype(convert_dict)
"""#end format columns - reduce size"""
wList_date=["DATA FATTURA"	,"DATA STATO","DATA LIQUIDAZIONE"]
#Date format
for d in wList_date:
    df_quiet[d]=pd.to_datetime(df_quiet[d],errors='coerce')
#df_quiet["BDOLin"]=df_quiet["NUMERO BDO"].multiply(1000)+df_quiet["POSIZIONE"] 

#olddf_quiet=pd.merge(df_quiet,df_BDOLin[["Forn_cod","Forn_nome","Subf_nome","Subf_cod", "BDOLin","RTI_cod"]],left_on="BDOLin", right_on="BDOLin",how="left",suffixes=('', '_rofi'))
#df_quiet=pd.merge(df_quiet,df_forn_byName[["Forn_cod","Forn_nome"]],left_on="FORNITORE ROFI", right_on="Forn_nome",how="left",suffixes=('', '_rofi'))
#df_quiet=pd.merge(df_quiet,df_forn_byName[["Forn_cod","Forn_nome"]],left_on="SUBFORNITORE", right_on="Forn_nome",how="left",suffixes=('', '_subf'))
#df_quiet=pd.merge(df_quiet,df_forn_byName[["Forn_cod","Forn_nome"]],left_on="OWNER CAMBIO STATO", right_on="Forn_nome",how="left",suffixes=('', '_own'))
#df_quiet["DATA STATO"]=df_quiet["DATA STATO"].str[0:10]
df_quiet=df_quiet.rename(columns={"POSIZIONE":"Linea","FORNITORE ROFI":"Fornitore","NUMERO BDO":"BDO","SUBFORNITORE":"Subfornitore"," OWNER CAMBIO STATO": "Ref_var_stato",'PERIODO COMPETENZA SAL':"Periodo_SAL","DATA STATO":"Data_var_stato","IMPORTO":"Tot_Quiet","NOME STATO":"Stato_Quiet","NUMERO FATTURA":"Numero Fattura","DATA FATTURA":"Data Fattura","DATA LIQUIDAZIONE":"Data Liquidazione" })
df_quiet['Periodo_SAL'].replace('/','-',inplace=True)

#df_quiet["Periodo"]= df_quiet["Periodo_SAL"].str.slice(0,4)+df_quiet["Periodo_SAL"].str[5:7]

#df_quiet["Periodo"]=df_quiet["Periodo"].astype(int)
#
df_quiet=df_quiet[df_quiet["ULTIMO STATO"]=="SI"]
df_quiet=df_quiet[df_quiet["BEF"].notna()]


logger.info('end reading quiet ' + str(df_quiet.shape[0]))

logger.info('start reading BEF')
df_BEF=pd.read_excel(wrk_dir_inp+"\\"+"Da BEF_PortaleN.xlsx", sheet_name=1, dtype={"Descrizione":str, "Nome file":str, "Data di emissione":str, "Mese competenza Verbale":str, "Descrizione Linea Ordine":str, "Subfornitura":str, "Divisione":str, "Centro di Costo":str, "Fornitore RTI":str, "Fornitore Reale":str,  "Piano Fatturazione":str, "Quota Fatturazione":str,"UDM":str, "Numero Fattura":str, "Cup":str, "Flag Ritenuta":str,"Periodo Competenza":str,"Linea":str,"Numero Linea Ordine":str})

if (wrk_environm.lower()=="testfilter") & ("*all" not in listfilter_BDO):
    df_BEF=df_BEF.loc[(df_BEF["Numero BDO"].isin(listfilter_BDO))|df_BEF["Numero Ricezione/Attestazione"].isin(listfilter_MAP)]


#df_BEF["BDOLin"]=df_BEF["Numero BDO"].multiply(1000)+df_BEF["Numero Linea Ordine","Numero della linea di ricezione"] 
#get substring periodo
df_BEF["Anno competenza"]= df_BEF["Periodo Competenza"].str.slice(3,7)
df_BEF["Mese competenza"]= df_BEF["Periodo Competenza"].str.slice(0,2)

"""#begin format columns - reduce size"""
wLi=["Numero BDO","Numero Linea Ordine","Numero Ricezione/Attestazione","Anno competenza","Mese competenza"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["Quantità","Importo Unitario",	"Importo Ricezione"];wDf={value:"float" for value in wLf}
wLc=["Subfornitura","Benestare alla Fatturazione",	"Divisione",	"Centro di Costo",	"Fornitore RTI",	"Fornitore Reale","Periodo Competenza","Quota Fatturazione","UDM","Cup",	"Flag Ritenuta","% Anticipo"];wDc={value:"category" for value in wLc}
wLnbr=wLi+wLf
df_BEF[wLnbr]=df_BEF[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDf,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_BEF= df_BEF.astype(convert_dict)
"""#end format columns - reduce size"""
wList_date=["Data Fattura",	"Data Pagamento","Data di emissione"]
#Date format
for d in wList_date:
    df_BEF[d].fillna("01/01/1900", inplace=True) #set default values
    df_BEF[d]=pd.to_datetime(df_BEF[d],errors='coerce')

#df_BEF= df_BEF.astype({"Data Fattura":float,	"Data Pagamento":float})

#old df_BEF=pd.merge(df_BEF,df_BDOLin[["Forn_cod","Forn_nome","Subf_nome","Subf_cod", "BDOLin","RTI_cod","Codice CUP"]],left_on="BDOLin", right_on="BDOLin",how="left",suffixes=('', '_rofi'))

#df_BEF=pd.merge(df_BEF,df_forn_byName[["Forn_cod","Forn_nome"]],left_on="Fornitore Reale", right_on="Forn_nome",how="left",suffixes=('', '_forn'))
#df_BEF=pd.merge(df_BEF,df_forn_byName[["Forn_cod","Forn_nome"]],left_on="Fornitore RTI", right_on="Forn_nome",how="left",suffixes=('', '_RTI'))
#df_BEF=pd.merge(df_BEF,df_forn_byName[["Forn_cod","Forn_nome"]],left_on="Subfornitura", right_on="Forn_nome",how="left",suffixes=('', '_subf'))

df_BEF=df_BEF.rename(columns={"Numero BDO":"DocAcq","Numero Ricezione/Attestazione": "MAP","Benestare alla Fatturazione": "BEF", "Numero Linea Ordine": "Linea","Mese competenza Verbale":"Mese_Verbale","Forn_cod_RTI":"FornRTI_cod","Forn_cod_subf":"Subf_cod" })
#df_BEF[["Numero Fattura","Data Fattura","Data Pagamento","Cup"]]= df_BEF[["Numero Fattura","Data Fattura","Data Pagamento","Cup"]].astype(str)

#df_BEF["Periodo"]=df_BEF["Periodo"].astype(int)
wList_date=["Data Fattura", "Data Pagamento"]
#Date format
for d in wList_date:
    df_BEF[d]=pd.to_datetime(df_BEF[d],errors='coerce')

df_BEF[["Numero Fattura","Data Fattura","Data Pagamento"]] = df_BEF[["Numero Fattura","Data Fattura","Data Pagamento"]].fillna(0)

df_BEF_Map=df_BEF.groupby(["DocAcq", "BEF", "Anno competenza","Mese competenza", "Numero Fattura", "Data Fattura", "Data Pagamento", "MAP", "Linea"], observed=True, dropna=False).agg(
TotBEF= ("Importo Ricezione","sum")
     ).reset_index()



df_BEF_BDOLin=df_BEF_Map.groupby(["DocAcq","Anno competenza","Mese competenza","Linea"], observed=True).agg(
TotBEF= ("TotBEF","sum")
     ).reset_index()
#compone la chiave Bdo + linea + periodo
#df_BEF_BDOLin["BDOLin_periodo"]=df_BEF_BDOLin["Numero BDO"].multiply(1000000000)+df_BEF_BDOLin["Linea"].multiply(10000000)+df_BEF_BDOLin["Periodo"]

df_BEF_BDO=df_BEF_Map.groupby(["DocAcq"]).agg(
TotBEF= ("TotBEF","sum")
     ).reset_index()

logger.info('end reading BEF ' + str(df_BEF_Map.shape[0]))
#Get Fatture
wpath_fpass = wrk_dir_inp+"\\MAPCons\\MAPFatture\\"
#wpath_fpass=r"\Users\fcaneri\OneDrive - ARIA S.p.A\PMO Shared\ImportDati\MAPCons\MAPFatture"+"\\"

#get Situazione Ordini for Protocollo IVA
filenames = [file for file in os.listdir(wpath_fpass) if file.endswith('.XLSX')]
df_SituazOrdini = pd.concat([pd.read_excel(wpath_fpass + file, dtype=str) for file in filenames], ignore_index=True)

"""#begin format columns - reduce size"""
wLi=["Documento d'acquisto","Posizione","MAP","Mese MAP",	"Anno MAP","Protocollo IVA","Protocollo Pag. Fattura","Num.Preacquisito di rif."];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["Valore Linea oda","Quantità",	"Valore MAP",	"Imponibile Fattura"];wDf={value:"float" for value in wLf}
wLc=["Perc.Quota MAP","Ritenute a Garanzia"];wDc={value:"category" for value in wLc}
wLnbr=wLi+wLf
df_SituazOrdini[wLnbr]=df_SituazOrdini[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDf,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_SituazOrdini= df_SituazOrdini.astype(convert_dict)
"""#end format columns - reduce size"""
wList_date=["Data doc. Fattura",	"Data Regist. Fattura","Inizio validità OdA",	"Fine validità OdA"]
#Date format
for d in wList_date:
    df_SituazOrdini[d]=pd.to_datetime(df_SituazOrdini[d],errors='coerce')


#convert_dict = {"Protocollo IVA":str,"Num.Fatt.Fornitore":str,"Protocollo Pag. Fattura":str}   # {'A': int, 'C': float}
# Convert columns using the dictionary
#df_SituazOrdini = df_SituazOrdini.astype(convert_dict)

wrenMAPColumns={"Documento d'acquisto":"DocAcq","Posizione":"Linea","Mese MAP":"Mese competenza","Anno MAP":"Anno competenza","Valore MAP":"ValoreMAP" }
df_SituazOrdini=df_SituazOrdini.rename(columns=wrenMAPColumns)

# Creating a constant value for column Count
constant_values = {"Num.Fatt.Fornitore":"", "Data doc. Fattura":"","Data doc. Fattura":"", "Desc.Fornitore Fattura":"","Protocollo IVA":"","Data di pagamento":""}
df_SituazOrdini = df_SituazOrdini.fillna(value = constant_values)
#wColselMAP=["MAP","Num.Fatt.Fornitore","Data doc. Fattura","Protocollo IVA","Data di pagamento"]
#df_SituazOrdini=df_SituazOrdini[wColselMAP] 

df_FattPass_IVA=df_SituazOrdini.groupby(
["DocAcq", "Linea", "MAP", "Mese competenza","Anno competenza"], observed=True).agg(**{"Num.Fatt.Fornitore":("Num.Fatt.Fornitore","max"), "Data doc. Fattura":("Data doc. Fattura","max"), "Desc.Fornitore Fattura":("Desc.Fornitore Fattura","max"),"Protocollo IVA":("Protocollo IVA","max"),"Data di pagamento":("Data di pagamento","max")}).reset_index()

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_BEF_BDOLin.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_BEF_BDOLin.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BEF_BDOLin.memory_usage(deep=True)) )
pass


from read_Bdo_Oda import df_forn_byName
#Read dummy tables
#from read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
from read_BdoLin import df_BDOLin

import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
# %%
