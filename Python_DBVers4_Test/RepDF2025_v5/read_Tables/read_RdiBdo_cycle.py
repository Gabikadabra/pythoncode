#%%
#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
from read_ODAG import df_ODAGMacr
#from read_rilasci import df_Rilasci_task
import os
import sys

from Util_logging import logger
wPythProc="read_RdiBdo_cycle"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]

df_RDI_old=pd.read_excel(wrk_dir_inp+"\\"+"Da RDIobsolete.xlsx",skiprows=3, dtype={})
df_RDI=pd.read_excel(wrk_dir_inp+"\\"+"Da DailyRepBdoRdi.xlsx", dtype={})
#df_RDI["ODAGMacr_txt"]=df_RDI["Contratto Gara"]+"_"+df_RDI["Contratto Macroclasse"]
df_RDI = df_RDI[~df_RDI["Richiesta d'acquisto"].isin(df_RDI_old['Numero RDI'])]
df_RDI.rename({"Contratto Gara":"ODAG","Numero Contratto Gara":"ODAGnbr","Contratto Macroclasse":"Macrocl_txt","Richiesta d'acquisto":"RDI","Richiedente":"ROI","Stato Ciclo Approvativo":"StatoRDI","Importo Richiesta":"RDI_val","Quantità":"RDI_qta","Data Creazione RDI":"DataCreaz","Data Approvazione":"DataApprov", "Documento d'acquisto":"DocAcq","Inizio Fornitura":"DataIniz","Fine Fornitura":"DataFine"}, axis=1, inplace=True)

df_RDI=pd.merge(df_RDI,df_ODAGMacr[["ODAG","Macrocl_txt","Macrocl_nbr"]],on=["ODAG","Macrocl_txt"],how="left",suffixes=('', '_verb'))

df_RDI=pd.merge(df_RDI,df_RisInt_byName[["RisInt_nome","CdC","Int_Dip"]],left_on="ROI", right_on="RisInt_nome",how="left",suffixes=('', '_MP'))

df_RDI = df_RDI.drop_duplicates(subset=['RDI'],keep='last')

#%%
df_RDI_port=pd.read_excel(wrk_dir_inp+"\\"+"Da RdI_PortaleN.xlsx",sheet_name=1, dtype={"Numero RDI":str, "Descrizione RDI":str, "Nome file PIF/IF":str, "Codifica numerica documento":str, "Stato del documento PIF/IF":str, "Divisione":str, "Centro di Costo":str, "Ultima PIF/IF Approvata":str, "Descrizione PIF/IF":str, "Utente caricamento doc IF":str, "Fornitore":str, "ROI":str})
df_RDI_port=pd.merge(df_RDI_port,df_RisInt_byName[["RisInt_nome","CdC","Int_Dip"]],left_on="ROI", right_on="RisInt_nome",how="left",suffixes=('', '_MP'))

df_RDI_port.rename({"Int_Dip":"ROI_dip","Numero RDI":"RDI","Stato del documento PIF/IF":"StatoIF","Stato Ciclo Approvativo":"StatoRDI", "CdC":"ROI_CdC"}, axis=1, inplace=True)
df_RDI_port = df_RDI_port.drop_duplicates(subset=['RDI'],keep='last')

df_BDOVar=pd.read_excel(wrk_dir_inp+"\\"+"Da BDO_PortaleN.xlsx",sheet_name=1, dtype={"Numero BDO":str, "Descrizione BDO":str, "Nome file PIF/IF":str, "Descrizione PIF/IF":str, "Codifica numerica documento":str, "Stato del documento PIF/IF":str, "Divisione":str, "Centro di Costo":str, "Ultima PIF/IF Approvata":str,  "Utente caricamento doc BDO":str, "Fornitore":str, "ROI":str, "PMO":str, "Data rifiuto PMO":str, "CTRM":str,  "Data rifiuto CTRM":str})
df_BDOVar=pd.merge(df_BDOVar,df_RisInt_byName[["RisInt_nome","CdC","Int_Dip"]],left_on="ROI", right_on="RisInt_nome",how="left",suffixes=('', '_MP'))

df_RDI_park=pd.read_excel(wrk_dir_inp+"\\"+"Da ZMMPARK.xlsx", dtype={"Richiesta d'acquisto incompleta":str, "Descrizione richiedente":str,  "Testo breve":str, "Contratto Gara":str, "Macroclasse":str, "Testo Macroclasse":str, "Fornitore":str, "Numero conto del fornitore":str })
#df_RDI_park["ODAGMacr_nbr"]=df_RDI_park["Numero Contratto Gara"]+"_"+df_RDI_park["Macroclasse"]
df_RDI_park = df_RDI_park[~df_RDI_park['Numero RDI'].isin(df_RDI_old['Numero RDI'])]
df_RDI_park=pd.merge(df_RDI_park,df_RisInt_byName[["RisInt_nome","CdC","Int_Dip"]],left_on="Descrizione richiedente", right_on="RisInt_nome",how="left",suffixes=('', '_MP'))
df_RDI_park["StatoRDI"]="Parcheggiata"
df_RDI_park.rename({"Int_Dip":"ROI_dip","Numero RDI":"RDI", "CdC":"ROI_CdC", "Descrizione richiedente":"ROI","Numero Contratto Gara":"ODAGnbr","Contratto Gara":"ODAG",  "Macroclasse":"Macrocl_nbr",  "Testo Macroclasse":"Macrocl_txt","Fornitore":"Forn_RTI_cod", "Numero conto del fornitore":"Forn_RTI"}, axis=1, inplace=True)
df_RDI_park = df_RDI_park.drop_duplicates(subset=['RDI'],keep='last')


"""#begin format columns - reduce size"""
wLi=["RDI"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=[];wDf={value:"float" for value in wLf}
wLc=["StatoIF","Divisione","Centro di Costo","Ultima PIF/IF Approvata","Utente caricamento doc IF","Fornitore","RisInt_nome","ROI","Codifica numerica documento"];wDc={value:"category" for value in wLc}
wLnbr=wLf+wLi
df_RDI_port[wLnbr]=df_RDI_port[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDf,wDc,wDi]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_RDI_port= df_RDI_port.astype(convert_dict)

#wList_date=["InizVal."	,"FineVal.",	"Data cr.",	"Data del documento acquisto","Data Fine Milestone",	"Data Effettiva Attivazione"]
"""#end format columns - reduce size"""

wList_date=["Data invio ROI","Data rifiuto ROI","Data approvazione ROI","Data caricamento"]
#Date format
for d in wList_date:
    df_RDI_port[d]=pd.to_datetime(df_RDI_port[d],errors='coerce')

"""#begin format columns - reduce size"""
wLi=["Macrocl_nbr","Forn_RTI_cod"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=[];wDf={value:"float" for value in wLf}
wLc=["Macrocl_txt","Forn_RTI","StatoRDI","ROI_CdC","ROI_dip"];wDc={value:"category" for value in wLc}
wLnbr=wLf+wLi
df_RDI_park[wLnbr]=df_RDI_park[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDf,wDc,wDi]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_RDI_park= df_RDI_park.astype(convert_dict)

#wList_date=["InizVal."	,"FineVal.",	"Data cr.",	"Data del documento acquisto","Data Fine Milestone",	"Data Effettiva Attivazione"]
"""#end format columns - reduce size"""


tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_RDI_port.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_RDI_port.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_RDI_port.memory_usage(deep=True)) )
pass
