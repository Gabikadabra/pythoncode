#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task

import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
from read_ODAG import df_ODAGMacr
from read_Bdo_Oda import df_forn_byName
#Read dummy tables
#from read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 

import os
import sys
#***  begin log - prima parte
from Util_logging import logger
wPythProc="read_BdoLin"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
#***  end log - prima parte
wDtFormat = "%d/%m/%Y" 
wActFormat = "%d.%m.%Y" 
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

#Last days differences
wrk_lastBDO_days=int(parm_dict["parm_lastBDO_days"])
w_today=datetime.now()
w_deltadays=timedelta(days=wrk_lastBDO_days)
#new limit dates
w_days_limit=w_today-w_deltadays
#
logger.info('start reading df_BDO')
#,"Valore Ordine":float,"Importo Attestato":float,"Importo Residuo":float,"Importo Annullato":float,"Fatturato":str
df_BDO=pd.read_excel(wrk_dir_inp+"\\"+"Da DailyRepBdo.xlsx", dtype={}, decimal=',')

# ***Filter*** in testfilter is on and specific selection is in place
if (wrk_environm.lower()=="testfilter") & ("*all" not in listfilter_BDO):
    df_BDO=df_BDO.loc[(df_BDO["Documento d'acquisto"].isin(listfilter_BDO))|(df_BDO["Numero Contratto Gara"].isin(listfilter_ODAG))]

#Format date
wList_date=["Data Creazione Bdo"	,"Data Approvazione",	"Data decorrenza",	"Data scadenza"]
#Date format
for d in wList_date:
    df_BDO[d]=pd.to_datetime(df_BDO[d],errors='coerce')
""" precedenti formattazioni date
df_BDO.loc[df_BDO["Data Creazione Bdo"] == "00.00.0000", "Data Creazione Bdo"] = "01.01.1980"
df_BDO.loc[df_BDO["Data Approvazione"] == "00.00.0000", "Data Approvazione"] = "01.01.1980"
df_BDO.loc[df_BDO["Data decorrenza"] == "00.00.0000", "Data decorrenza"] = "01.01.1980"
df_BDO.loc[df_BDO["Data scadenza"] == "00.00.0000", "Data scadenza"] = "31.12.2099"

df_BDO["Data Creazione Bdo"]=pd.to_datetime(df_BDO["Data Creazione Bdo"], format='%d.%m.%Y')
df_BDO["Data Approvazione"]=pd.to_datetime(df_BDO["Data Approvazione"], format='%d.%m.%Y')
df_BDO["Data decorrenza"]=pd.to_datetime(df_BDO["Data decorrenza"], format='%d.%m.%Y')
df_BDO["Data scadenza"]=pd.to_datetime(df_BDO["Data scadenza"], format='%d.%m.%Y')
"""
#df_BDO["durata"]=df_BDO["Data scadenza"]-df_BDO["Data decorrenza"]
#Per trasformare le date in formato gg-mm-aaa  è possibile utilizzare l'espressione df_BDO["Data scadenza"].dt.strftime('%d-%m-%Y')
#df_BDO["ODAGMacr_txt"] = df_BDO["Contratto Gara"].str.cat(df_BDO["Contratto Macroclasse"], sep = "_")
#Get ODAG Macro nbr

"""#begin format columns - reduce size"""
wLi=["Numero Contratto Gara","Richiesta d'acquisto",	"Documento d'acquisto", "Numero mesi"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["Check Valore Rdi Bdo","Valore Ordine",	"Quantità Ordine",	"Importo Attestato",	"Quantità Attestata",	"Importo Residuo",	"Importo Annullato","Fatturato"];wDf={value:"float" for value in wLf}
wLc=["Tipologia Fornitura","Stato Testata Ordine","Definizione","Numero Cicli in Lavorazione","Referente operativo","Unità Organizzativa",	"Nome Responsabile",	"Area Responsabile","Fornitore RTI","Flag di Ultima Consuntivazione","Codice Struttura","Chius.Amm.va"];wDc={value:"category" for value in wLc}
wLnbr=wLi+wLf
df_BDO[wLnbr]=df_BDO[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDf,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_BDO= df_BDO.astype(convert_dict)


#wList_date=["InizVal."	,"FineVal.",	"Data cr.",	"Data del documento acquisto","Data Fine Milestone",	"Data Effettiva Attivazione"]
"""#end format columns - reduce size"""
#Rename columns 
wrename_dict={"Numero Contratto Gara":"ODAGnbr","Contratto Gara":"ODAG","Contratto Macroclasse":"Macrocl_txt","Documento d'acquisto":"DocAcq","Richiesta d'acquisto":"RDI","Descrizione Bdo":"Descrizione DocAcq","Definizione":"StatoDocAcq","Data Creazione Bdo":"DataCreaz","Data Approvazione":"StartDate","Data decorrenza":"DataIniz","Data scadenza":"DataFine","Fornitore RTI":"Forn_RTI","Valore Ordine":"Forn_val","Referente operativo":"ROI","Chius.Amm.va":"DocAcq_FlagCh","Importo Residuo":"Val_resid"}

df_BDO.rename(columns=wrename_dict, inplace=True)
df_BDO=pd.merge(df_BDO,df_ODAGMacr[["ODAG","ODAGnbr","Macrocl_txt","Macrocl_nbr","DEC","RUP"]],on=["ODAG","Macrocl_txt"],how="left",suffixes=('', '_verb'))
df_BDO=pd.merge(df_BDO,df_RisInt_byName[["RisInt_nome","Int_Dip","CID","CdC"]],left_on="ROI", right_on="RisInt_nome",how="left",suffixes=('', '_int'))
df_BDO=pd.merge(df_BDO,df_forn_byName[["Forn_nome","Forn_cod"]],left_on="Forn_RTI", right_on="Forn_nome",how="left",suffixes=('', '_fRTI'))

wrename_dict={"CdC":"ROI_CdC","Int_Dip":"ROI_dip","CID":"ROI_CID","Forn_cod":"Forn_RTI_cod"}
df_BDO.rename(columns=wrename_dict, inplace=True)
#Filter if Develop environment
#if wrk_environm =="Develop":
#	df_BDO= df_BDO.loc[df_BDO["BDO"].isin(df_BDO_dummy["BDO"])]
#df_BDO["Val_Resid"]=df_BDO["Forn_val"]-df_BDO["Importo Attestato"]




df_BDOLin=pd.read_excel(wrk_dir_inp+"\\"+"Da DailyRepBdoLinee.xlsx", dtype={"Subappaltatore": str,"Tipologia Fornitura":str, }, decimal=',')
#df_BDOLin["ODAGMacr_txt"] = df_BDOLin["Contratto Gara"].str.cat(df_BDOLin["Contratto Macroclasse"], sep = "_")
#Get ODAG Macro nbr
# ***Filter*** in testfilter is on and specific selection is in place
if (wrk_environm.lower()=="testfilter") & ("*all" not in listfilter_BDO):
    df_BDOLin=df_BDOLin.loc[(df_BDOLin["Documento d'acquisto"].isin(listfilter_BDO))]

df_BDOLin=df_BDOLin.rename(columns={"Data Fine Milestone":"Data_MLS","Subappaltatore":"Subf_cod","Descrizione Subappaltatore":"Subf_nome", "Documento d'acquisto":"BDO","Fornitore di posizione":"Forn_nome","Fornitore Reale":"Fornitore RTI"})

#df_BDOLin=pd.merge(df_BDOLin,df_ODAGMacr[["ODAGMacr_txt","ODAGMacr_nbr"]],left_on="ODAGMacr_txt", right_on="ODAGMacr_txt",how="left",suffixes=('', '_verb'))
#df_BDOLin["BDOLin"]=df_BDOLin["BDO"].multiply(1000)+df_BDOLin["Posizione"]
df_BDOLin=pd.merge(df_BDOLin,df_forn_byName[["Forn_nome","Forn_cod"]],left_on="Forn_nome", right_on="Forn_nome",how="left",suffixes=('', '_fpos'))
#df_BDOLin=pd.merge(df_BDOLin,df_BDO[["BDO","ODAGMacr_nbr","ROI_CID","ROI_CdC","ROI_dip","ROI","RTI_cod","StatoBDO","Numero Contratto Gara","Contratto Macroclasse","Chius_Amm"]],left_on="BDO", right_on="BDO",how="left",suffixes=('', '__bdo'))


#df_BDOLin["Data decorrenza"]=pd.to_datetime(df_BDOLin["Data decorrenza"], format='%d.%m.%Y')
#df_BDOLin["Data scadenza"]=pd.to_datetime(df_BDOLin["Data scadenza"], format='%d.%m.%Y')
df_BDOLin.loc[df_BDOLin["Data_MLS"] == "", "Data_MLS"] = df_BDOLin["Data scadenza"]
#df_BDOLin["Data_MLS"]=pd.to_datetime(df_BDOLin["Data_MLS"], format='%d.%m.%Y',errors="coerce" )
#Format date
wList_date=["Data decorrenza"	,"Data scadenza",	"Data_MLS"]
#Date format
for d in wList_date:
    df_BDOLin[d]=pd.to_datetime(df_BDOLin[d],errors='coerce')

"""#begin format columns - reduce size"""
wLi=["BDO","Posizione","Subf_cod"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["Fatturato","Valore Ordine","Importo Attestato","Importo Subappalto","Costo Subappalto"];wDf={value:"float" for value in wLf}
wLc=["Tipologia Fornitura","Codice CUP"];wDc={value:"category" for value in wLc}
wLnbr=wLi+wLf
df_BDOLin[wLnbr]=df_BDOLin[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDf,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_BDOLin= df_BDOLin.astype(convert_dict)

#wList_date=["InizVal."	,"FineVal.",	"Data cr.",	"Data del documento acquisto","Data Fine Milestone",	"Data Effettiva Attivazione"]
"""#end format columns - reduce size"""


#df_BDOLin["Val_Resid"]=df_BDOLin["Valore Ordine"]-df_BDOLin["Importo Attestato"]

wrename_dict={"Contratto Gara":"ODAG","Fornitore RTI":"Forn_RTI","Fatturato.1":"Fatt_val","BDO":"DocAcq","Posizione":"Linea","Descrizione Bdo":"Descrizione DocAcq","Data decorrenza":"DataIniz","Data scadenza":"DataFine","Valore Ordine":"Forn_val","Importo Subappalto":"Subf_val","Codice CUP":"CUP","Descrizione Milestone":"Lin_descrizione","Tipologia Fornitura":"TipLin_Descr"}

df_BDOLin.rename(columns=wrename_dict,inplace=True)

df_BDO_forn=df_BDOLin.groupby(["DocAcq"]).agg(Fornitori= ("Forn_nome","unique"),
Subfornitori=("Subf_nome","unique")
     ).reset_index()
df_BDO_forn['Fornitori']=[','.join(map(str, l)) for l in df_BDO_forn['Fornitori']]
df_BDO_forn['Subfornitori']=[','.join(map(str, l)) for l in df_BDO_forn['Subfornitori']]

wColselPDC=["ODAG","Forn_RTI","Forn_nome","Fatt_val","DocAcq","Linea","Descrizione DocAcq","TipLin_Descr","DataIniz","DataFine","Forn_val","Importo Attestato","Subf_val","Subf_cod","Subf_nome","CUP","Lin_descrizione","Data_MLS","Forn_cod"]
#Filter Relevant columns
df_BDOLin=df_BDOLin[wColselPDC]

df_BDO=pd.merge(df_BDO,df_BDO_forn,left_on="DocAcq", right_on="DocAcq",how="left",suffixes=('', '__bdo'))

wColselPDC=["ODAG","ODAGnbr","Macrocl_txt","RDI","DocAcq","Descrizione DocAcq","StatoDocAcq","ROI","DataCreaz","StartDate","DataIniz","DataFine","Forn_RTI","Forn_val","Importo Attestato","Val_resid","Importo Annullato","Fatturato","DocAcq_FlagCh","Macrocl_nbr","DEC","RUP","ROI_dip","ROI_CID","ROI_CdC","Forn_nome","Forn_RTI_cod","Fornitori","Subfornitori"]
#Filter Relevant columns
df_BDO=df_BDO[wColselPDC]

df_BDOLin=pd.merge(df_BDOLin,df_BDO[["DocAcq","ODAGnbr","Macrocl_txt","RDI","StatoDocAcq","ROI","DataCreaz","StartDate","DocAcq_FlagCh","Macrocl_nbr","DEC","RUP","ROI_dip","ROI_CID","ROI_CdC"]],left_on="DocAcq", right_on="DocAcq",how="left",suffixes=('', '__bdo'))
df_BDOLin["Val_resid"]=df_BDOLin["Forn_val"]-df_BDOLin["Importo Attestato"]

df_BDOLin_recent= df_BDOLin[df_BDOLin["DataFine"]>w_days_limit]
df_BDO_recent= df_BDO[df_BDO["DataFine"]>w_days_limit]

#***  begin log - seconda parte
tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_BDO.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_BDO.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDO.memory_usage(deep=True)) )
#***  end log - seconda parte
pass
