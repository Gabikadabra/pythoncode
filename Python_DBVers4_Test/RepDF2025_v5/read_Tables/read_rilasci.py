#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task,listfilter_Incarico

#Read dummy tables

from read_Interni import df_RisInt_byName

from Util_logging import logger
wPythProc="read_rilasci"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
#from read_Interni import df_RisInt_byCID,df_RisInt_byName
import os
import sys

##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]

#Inizio rilasci
logger.info('start reading Rilasci')


df_Rilasci_resp=pd.read_excel(wrk_dir_inp+"\\"+"Da ZRilasciDemand.XLSX")
df_Rilasci_resp=pd.merge(df_Rilasci_resp,df_RisInt_byName[["RisInt_nome","Int_Dip"]],left_on="Nome Referente", right_on="RisInt_nome",how="left",suffixes=('', '_dmr'))





#Get demand 
df_RilasciDmd= df_Rilasci_resp.loc[df_Rilasci_resp["Flag"]=="SI"]
df_RilasciDmd= df_RilasciDmd.drop_duplicates(subset='Cod Rilascio', keep="last")
df_RilasciDmd=df_RilasciDmd.rename(columns={"Int_Dip":"resp_dmd_dip","Nome Referente":"resp_dmd"  })
#Add aggregated Column with value based on condition
#df_Rilasci_resp['Resp_dmd'] = df_Rilasci_resp.assign(Nome = df_Rilasci_resp['Nome Referente'].where(df_Rilasci_resp['Flag'] == 'SI')).groupby('Cod Rilascio')['Nome Referente'].transform('max')

#df_Rilasci_resp=pd.merge(df_Rilasci_resp,df_RisInt_byName[["Cognome Nome","Int_Dip"]],left_on="Resp_dmd", right_on="Cognome Nome",how="left",suffixes=('', '_dmr'))
#df_Rilasci_resp=pd.merge(df_Rilasci_resp,df_RisInt_byName[["Cognome Nome","Int_Dip"]],left_on="Nome Referente", right_on="Cognome Nome",how="left",suffixes=('', '_resp'))
#df_Rilasci_resp=df_Rilasci_resp.rename(columns={"Int_Dip_resp":"resp_Dip"  })
#
df_Rilasci_resp=df_Rilasci_resp.groupby(['Cod Rilascio']).agg(resp_ril = ("Nome Referente","unique"),resp_ril_dip=("Int_Dip","unique")
     ).reset_index()
#Convert list column in a string 
df_Rilasci_resp['resp_ril_dip']=[','.join(map(str, l)) for l in df_Rilasci_resp['resp_ril_dip']]
df_Rilasci_resp['resp_ril']=[','.join(map(str, l)) for l in df_Rilasci_resp['resp_ril']]
df_Rilasci_resp=pd.merge(df_Rilasci_resp,df_RilasciDmd[["Cod Rilascio","resp_dmd","resp_dmd_dip"]],left_on="Cod Rilascio", right_on="Cod Rilascio",how="left",indicator=True)

#Slect the column based on str contains
#df_Rilasci_resp['resp_ril_dip'].str.contains("DWEL", na=False)

df_RilasciPPA=pd.read_excel(wrk_dir_inp+"\\"+"Da ZRilasci_PPA_UE.xlsx", dtype={"Codice Contratto": str, "Codice CUP":str, "Numero Capitolo":str, "Anno Capitolo":str})
wCharColumns=["Codice Rilascio","Codice Attività PPA","Codice CUP","Numero Capitolo","Anno Capitolo"]
wNumColumns=["Quota rilascio a valere capitolo e PPA"]
df_RilasciPPA[wCharColumns]=df_RilasciPPA[wCharColumns].fillna("")
df_RilasciPPA[wNumColumns]=df_RilasciPPA[wNumColumns].fillna(0)

df_RilasciPPA_sub=df_RilasciPPA.groupby(["Codice Rilascio","Codice Attività PPA","Codice CUP","Numero Capitolo","Anno Capitolo"])[["Quota rilascio a valere capitolo e PPA"]].sum().reset_index()
df_RilasciPPA_subPPA=df_RilasciPPA_sub.groupby(["Codice Rilascio","Codice Attività PPA"]).count().reset_index()
df_RilasciPPA_subPPA=df_RilasciPPA_subPPA.groupby(["Codice Rilascio"], as_index=False)["Codice Attività PPA"].agg(" § ".join)



df_RilasciPPA_subCUP=df_RilasciPPA_sub.groupby(["Codice Rilascio","Codice CUP"]).count().reset_index()
df_RilasciPPA_subCUP=df_RilasciPPA_subCUP.groupby(["Codice Rilascio"], as_index=False)["Codice CUP"].agg(" § ".join)

df_Rilasci=pd.read_excel(wrk_dir_inp+"\\"+"da ZRilasci.xlsx",dtype={"INCARICO": str, "Codice CUPCodice CUP": str,"Codice Cliente": str})
df_Rilasci.rename(columns={ df_Rilasci.columns[21]: "Numero Task" }, inplace = True)
#Use legacy table
df_Rilasci_legacy=pd.read_excel(wrk_dir_inp+"\\ImportLegacy\\"+"da_ZRilasci_Legacy.xlsx",dtype={"INCARICO": str, "Codice CUPCodice CUP": str,"Codice Cliente": str})
df_Rilasci_legacy.rename(columns={ df_Rilasci_legacy.columns[21]: "Numero Task" }, inplace = True)

df_Rilasci_legacy=pd.merge(df_Rilasci_legacy,df_Rilasci,left_on="Numero Task", right_on="Numero Task",how="outer",indicator=True)
df_Rilasci_legacy=df_Rilasci_legacy[df_Rilasci_legacy['_merge']=='left_only']

#Add tasks not included in actual file
df_Rilasci = pd.concat([df_Rilasci, df_Rilasci_legacy], ignore_index=True, sort=False)

"""#end format columns - reduce size"""
if (wrk_environm.lower()=="testfilter") & ("*all" not in listfilter_Incarico):
    df_Rilasci=df_Rilasci.loc[(df_Rilasci["INCARICO"].isin(listfilter_Incarico))|(df_Rilasci["Numero Task"].isin(listfilter_Task))]



wList_date=["Data inizio Attività Ver. Contrattuale","Data di Fine Attività Ver. Contrattuale","Data inizio Attività Ver. Corrente","Data di Fine Attività Ver. Corrente","Data Validazione RC","Data Verbalizzazione","Data Rilascio RCData Rilascio RC"]
#Date format
for d in wList_date:
    df_Rilasci[d]=pd.to_datetime(df_Rilasci[d],errors='coerce')

"""#begin format columns - reduce size"""
#wLi=["Documento Acquisto",	"Posizione","Numero MAP","Cod.Risorsa","Referente Operativo","Mese competenza",	"Anno competenza","Cod. fornitore pos."];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["BDG dei costi Contrattuale","BDG dei costi Corrente","Costi consuntivati Es.Prec.","Costi consuntivati al periodo/Es.Corr.","Consuntivi Totali","BDG Ricavi Rilascio Operativo PA","Ric. Maturato al periodo /Es.Corr.","Ricavo Maturato Es.Prec.","Ricavi Totali Maturati","Importo Verbalizzato (IVA)","Importo Verbalizzato","BDG costi versione WIP","BDG ricavi versione WIP"];wDf={value:"float" for value in wLf}
wLc=["Tipo Incarico","Codice Cliente","Stato Incarico","Codice Rilascio Contr.","Master Program","Profit Center MP","Stato Progetto","Profit Center Progetto","Tipo Rilascio","Tipo Task","Stato Verbale","Linea di business","CdC Task","CdC supervisore","Codice CUPCodice CUP","Rilevante PNRRRilevante PNRR"];wDc={value:"category" for value in wLc}
wLnbr=wLf
df_Rilasci[wLnbr]=df_Rilasci[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDf,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_Rilasci= df_Rilasci.astype(convert_dict)



#Filter if Develop environment
#if wrk_environm =="Develop":
#	df_Rilasci= df_Rilasci.loc[df_Rilasci["Codice Rilascio Contr."].isin(df_Task_dummy["Cod_Rilascio"])]

df_Rilasci_subt=df_Rilasci.groupby(["Tipo Incarico", "Codice Rilascio Contr.", "Titolo Rilascio Contr.", "Tipo Rilascio", "Linea di business", "INCARICO","Stato Incarico"],observed=True)[["BDG dei costi Corrente","BDG dei costi Contrattuale", "Consuntivi Totali"]].sum().reset_index()
#Rilasci Demand
#df_RilasciDmd=pd.read_excel(wrk_dir_inp+"\\"+"Da ZRilasciDemand.XLSX",dtype={"INCARICO": str, "cod incarico": str,"Cod Rilascio": str})
#Rilasci Demand filtering
wColsel=["INCARICO","Tipo Incarico","Codice Cliente","Desc.Cliente","Titolo Incarico","Stato Incarico","Codice Rilascio Contr.","Titolo Rilascio Contr.","Master Program","Demand","Profit Center MP","Descrizione Area MP","Numero Progetto","Nome Progetto","Stato Progetto","Responsabile Progetto","Profit Center Progetto","Descrizione Area Progetto","Tipologia Progetto","Tipo Rilascio","Numero Task","Nome Task","Tipo Task","BDG dei costi Contrattuale","BDG dei costi Corrente","Data inizio Attività Ver. Corrente","Data di Fine Attività Ver. Corrente","Data Validazione RC","Consuntivi Totali","Responsabile MP","Stato Verbale","Linea di business","CdC Task","CdC Task Descrizione","Data Rilascio RCData Rilascio RC","Codice CUPCodice CUP","Rilevante PNRRRilevante PNRR"]
#Filter Relevant columns
df_Rilasci_sel=df_Rilasci[wColsel]

#wCharColumns=["Tipo Incarico", "Codice Rilascio Contr.", "Titolo Rilascio Contr.", "Tipo Rilascio", "Linea di business", "INCARICO","Master Program",	"Demand",	"Profit Center MP",	"Descrizione Area MP","Numero Progetto",	"Nome Progetto",	"Stato Progetto",	"Responsabile Progetto",	"Profit Center Progetto",	"Descrizione Area Progetto",	"Tipologia Progetto",	"Numero Task",	"Nome Task",	"Tipo Task"]
#wNumColumns=["BDG dei costi Corrente","BDG dei costi Contrattuale"]

#df_Rilasci[wCharColumns]=df_Rilasci[wCharColumns].fillna("")
#df_Rilasci[wNumColumns]=df_Rilasci[wNumColumns].fillna(0)
#Riepilogo rilasci per task
df_Rilasci_task=df_Rilasci_sel.groupby(["Tipo Incarico", "Codice Rilascio Contr.", "Titolo Rilascio Contr.", "Tipo Rilascio", "Linea di business", "INCARICO","Master Program",	"Demand",	"Profit Center MP",	"Descrizione Area MP","Numero Progetto",	"Nome Progetto",	"Stato Progetto",	"Responsabile Progetto",	"Profit Center Progetto",	"Descrizione Area Progetto",	"Tipologia Progetto",	"Numero Task",	"Nome Task",	"Tipo Task", "Stato Incarico"],observed=True).agg(BDGCostiCorrenti=("BDG dei costi Corrente",np.sum),BDGCostiContratt=("BDG dei costi Contrattuale",np.sum),ConsTotali=("Consuntivi Totali",np.sum),DataIniz=("Data inizio Attività Ver. Corrente",np.min),DataFine=("Data di Fine Attività Ver. Corrente",np.max),DataValidRC=("Data Validazione RC",np.max)).reset_index()
#[["BDG dei costi Corrente","BDG dei costi Contrattuale", "Consuntivi Totali"]].sum().reset_index()
#df_Rilasci_task=df_Rilasci_task.agg(BDGCostiCorrenti=("BDG dei costi Corrente",np.sum),BDGCostiContratt=("BDG dei costi Contrattuale",np.sum),ConsTotali=("Consuntivi Totali",np.sum),DataIniz=("Data inizio Attività Ver. Corrente",np.min),DataFine=("Data di Fine Attività Ver. Corrente",np.max),DataValidRC=("Data Validazione RC",np.max))
df_Rilasci_task=pd.merge(df_Rilasci_task,df_Rilasci_resp[["Cod Rilascio","resp_ril","resp_dmd","resp_ril_dip","resp_dmd_dip"]],left_on="Codice Rilascio Contr.", right_on="Cod Rilascio",how="left",suffixes=('', '_pos'))
df_Rilasci_task=pd.merge(df_Rilasci_task,df_RilasciPPA_subPPA[["Codice Rilascio","Codice Attività PPA"]],left_on="Codice Rilascio Contr.", right_on="Codice Rilascio",how="left",suffixes=('', '_ppa'))
df_Rilasci_task=pd.merge(df_Rilasci_task,df_RilasciPPA_subCUP[["Codice Rilascio","Codice CUP"]],left_on="Codice Rilascio Contr.", right_on="Codice Rilascio",how="left",suffixes=('', '_cup'))



#collect informatio from dummy tables for yniq rilascio
df_Rilasci_subt=pd.merge(df_Rilasci_subt,df_Rilasci_resp[["Cod Rilascio","resp_ril","resp_dmd","resp_ril_dip","resp_dmd_dip"]],left_on="Codice Rilascio Contr.", right_on="Cod Rilascio",how="left",suffixes=('', '_pos'))
df_Rilasci_subt=pd.merge(df_Rilasci_subt,df_RilasciPPA_subPPA[["Codice Rilascio","Codice Attività PPA"]],left_on="Codice Rilascio Contr.", right_on="Codice Rilascio",how="left",suffixes=('', '_ppa'))
df_Rilasci_subt=pd.merge(df_Rilasci_subt,df_RilasciPPA_subCUP[["Codice Rilascio","Codice CUP"]],left_on="Codice Rilascio Contr.", right_on="Codice Rilascio",how="left",suffixes=('', '_cup'))

df_Rilasci_task.drop(['Codice Rilascio_cup','Codice Rilascio', 'Cod Rilascio','Descrizione Area Progetto' ], axis=1, inplace=True)
df_Rilasci_task.rename({'Codice Rilascio Contr.': 'Cod_Rilascio','Titolo Rilascio Contr.':'Tit_Rilascio','Tipo Rilascio':'Tipo_Rilascio','Linea di business':'LOB','Codice Attività PPA':'Cod_PPA','Codice CUP':'Cod_CUP'}, axis=1, inplace=True)

df_Rilasci_subt.drop(['Codice Rilascio','Cod Rilascio'], axis=1, inplace=True)
df_Rilasci_subt.rename({'Codice Rilascio Contr.': 'Cod_Rilascio','Titolo Rilascio Contr.':'Tit_Rilascio','Tipo Rilascio':'Tipo_Rilascio','Linea di business':'LOB','Codice Attività PPA':'Cod_PPA','Codice CUP':'Cod_CUP', "BDG dei costi Corrente": "BDGCostiCorrente", "BDG dei costi Contrattuale": "BDGCostiContratt", "Consuntivi Totali":  "ConsTotali"     }, axis=1, inplace=True)

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_Rilasci_subt.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_Rilasci_subt.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_Rilasci_subt.memory_usage(deep=True)) )