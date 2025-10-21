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
#from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_rilasci import df_Rilasci_task
import os
import sys
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task
from Util_logging import logger
wPythProc="read_verbPass"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()

##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
df_StatoVerb=pd.read_excel(wrk_dir_inp+"\\"+"Tab_StatoVerb.xlsx", dtype={"ID_Verbale": int,"Stato Verbale":str})
df_StatoVerb_byCod= df_StatoVerb.drop_duplicates(subset='ID_Verbale', keep="first")

wDict_Verb=df_StatoVerb_byCod.groupby('ID_Verbale')['Stato Verbale'].apply(list).to_dict()

df_VerbPassSAL_daRep=pd.read_excel(wrk_dir_inp+"\\"+"Da VerbaliSAL SAP.xlsx",sheet_name=1, dtype={"Periodo competenza":str})

df_Verb_ROFI=df_VerbPassSAL_daRep[["Numero BDO","Utente Caricamento Fornitore"]]
df_Verb_ROFI=df_Verb_ROFI.rename(columns={"Numero BDO":"DocAcq"})


#Remove table rows with "Periodo competenza" not meaningful
df_VerbPassSAL_daRep=df_VerbPassSAL_daRep.loc[df_VerbPassSAL_daRep["Periodo competenza"]>"0"]

df_VerbPassSAL_daRep["Anno competenza"]=df_VerbPassSAL_daRep["Periodo competenza"].str.slice(6,10)
df_VerbPassSAL_daRep["Mese competenza"]=df_VerbPassSAL_daRep["Periodo competenza"].str.slice(3,5)

df_VerbPassSAL_daRep= df_VerbPassSAL_daRep.astype({"Anno competenza":int,"Mese competenza":int})
#df_VerbPassSAL_daRep["Periodo"]=np.where(df_VerbPassSAL_daRep["Periodo competenza"].notnull(),
#                                         df_VerbPassSAL_daRep["Periodo competenza"].str.slice(6,10)+"-"+df_VerbPassSAL_daRep["Periodo competenza"].str.slice(3,5),
#                                         ""
#                                         )
df_VerbPassSAL_daRep=pd.merge(df_VerbPassSAL_daRep,df_StatoVerb[["ID_Verbale","Stato Verbale"]],left_on=df_VerbPassSAL_daRep["Stato Verbale"].str.lower(), right_on=df_StatoVerb["Stato Verbale"].str.lower(),how="left",suffixes=('', '_MP'))

#df_VerbPassSAL_daRep["BDO_periodo"]=df_VerbPassSAL_daRep["Numero BDO"]+"_"+df_VerbPassSAL_daRep["Periodo"]

df_VerbPassSAL0=df_VerbPassSAL_daRep[["Numero BDO","Stato Verbale","ID_Verbale","Anno competenza","Mese competenza"]]
df_VerbPassSAL0=df_VerbPassSAL0.rename(columns={"Stato Verbale":"Stat_VerbSAL","ID_Verbale":"ID_VerbSAL","Numero BDO":"DocAcq"})
#sostituisce valori specifici e li imposta a valori nulli
w_find=df_VerbPassSAL0["Stat_VerbSAL"].str.lower().isin(["verbale assente",0])
df_VerbPassSAL0.loc[w_find,"Stat_VerbSAL"] = ""
#Remove duplicate from SAL
df_VerbPassSAL=df_VerbPassSAL0.sort_values("ID_VerbSAL", ascending=False).drop_duplicates(["DocAcq", "Anno competenza","Mese competenza"]).sort_index()

df_VerbPass_daRep=pd.read_excel(wrk_dir_inp+"\\"+"Da DailyRepVerbali.xlsx", dtype={})
df_VerbPass_daRep['StatoVerb_Att']= np.where(df_VerbPass_daRep["Verbale Inizio Attività"].notnull(),9,0) 
df_VerbPass_daRep['StatoVerb_Chius']= np.where(df_VerbPass_daRep["Verbale di Chiusura"].notnull(),9,0) 
#df_VerbPass_daRep['StatoVerb_Chius']=df_VerbPass_daRep["Verbale di Chiusura"].apply(lambda x: '9' if x>"" else "")
wVerbCol= ["Documento d'acquisto","Data decorrenza","Data scadenza", "Richiedente","Descrizione Bdo","CdC",'StatoVerb_Att','StatoVerb_Chius']
df_VerbPass_daRep=df_VerbPass_daRep[wVerbCol]
df_VerbPassAtt=df_VerbPass_daRep[["Documento d'acquisto","StatoVerb_Att"]]
df_VerbPassAtt=df_VerbPassAtt.rename(columns={"Documento d'acquisto":"DocAcq","StatoVerb_Att":"ID_Verbale" })
df_VerbPassChius=df_VerbPass_daRep[["Documento d'acquisto","StatoVerb_Chius"]]
df_VerbPassChius=df_VerbPassChius.rename(columns={"Documento d'acquisto":"DocAcq","StatoVerb_Chius":"ID_Verbale" })


df_VerbPassAtt_daRep=pd.read_excel(wrk_dir_inp+"\\"+"Da VerbaliAtt SAP.xlsx",sheet_name=1, dtype={})
df_VerbPassAtt_daRep=df_VerbPassAtt_daRep.rename(columns={"Numero BDO":"DocAcq"})

#get info about supplier ref.
df_Verb_ROFI_Att=df_VerbPassAtt_daRep[["DocAcq","Utente Caricamento Fornitore"]]
#Add SAL & Start doc
#df_Verb_ROFI_Att= df_Verb_ROFI_Att.drop_duplicates(subset='Numero BDO', keep="last")

#remove duplicates
df_VerbPassAtt_daRep['Stato Verbale'] = df_VerbPassAtt_daRep['Stato Verbale'].str.upper()
df_VerbPassAtt_daRep=pd.merge(df_VerbPassAtt_daRep,df_StatoVerb[["ID_Verbale","Stato Verbale"]],left_on=df_VerbPassAtt_daRep["Stato Verbale"].str.lower(), right_on=df_StatoVerb["Stato Verbale"].str.lower(),how="left",suffixes=('', '_MP'))
df_VerbPassAtt_0=df_VerbPassAtt_daRep[["DocAcq","ID_Verbale"]]
#Add Verbali Att
df_VerbPassAtt = pd.concat([df_VerbPassAtt, df_VerbPassAtt_0], ignore_index=True, sort=False)
#Keep single occurrence Verbali Att
df_VerbPassAtt["ID_Verbale"] = df_VerbPassAtt["ID_Verbale"].fillna(0)
df_VerbPassAtt=df_VerbPassAtt.sort_values(by=["ID_Verbale"]).drop_duplicates("DocAcq", keep="last")
df_VerbPassAtt=pd.merge(df_VerbPassAtt,df_StatoVerb_byCod[["ID_Verbale","Stato Verbale"]],left_on=df_VerbPassAtt["ID_Verbale"], right_on=df_StatoVerb_byCod["ID_Verbale"],how="left",suffixes=('', '_aper'))
df_VerbPassAtt=df_VerbPassAtt.rename(columns={"Stato Verbale":"Stat_VerbAt","ID_Verbale":"ID_VerbAt"})
#sostituisce valori specifici e li imposta a valori nulli
w_find=df_VerbPassAtt["Stat_VerbAt"].str.lower().isin(["verbale assente",0])
df_VerbPassAtt.loc[w_find,"Stat_VerbAt"] = ""

#df_VerbPassAtt=df_VerbPassAtt.loc[df_VerbPassAtt["Stat_VerbAt"] == "VERBALE ASSENTE", "Stat_VerbAt"] = ""

df_VerbPassChius_daRep=pd.read_excel(wrk_dir_inp+"\\"+"Da VerbaliChiusura SAP.xlsx",sheet_name=1, dtype={"Numero BDO": str})
df_VerbPassChius_daRep=df_VerbPassChius_daRep.rename(columns={"Numero BDO":"DocAcq"})
#get info about supplier ref.
df_Verb_ROFI_Ch=df_VerbPassChius_daRep[["DocAcq","Utente Caricamento Fornitore"]]

#Add SAL & Start doc
df_Verb_ROFI=pd.concat([df_Verb_ROFI_Att,df_Verb_ROFI,df_Verb_ROFI_Ch])
df_Verb_ROFI=df_Verb_ROFI.rename(columns={"Numero BDO":"DocAcq"})

df_Verb_ROFI= df_Verb_ROFI.drop_duplicates(subset='DocAcq', keep="last")



df_VerbPassChius_daRep['Stato Verbale'] = df_VerbPassChius_daRep['Stato Verbale'].str.upper()
df_VerbPassChius_daRep=pd.merge(df_VerbPassChius_daRep,df_StatoVerb[["ID_Verbale","Stato Verbale"]],left_on=df_VerbPassChius_daRep["Stato Verbale"].str.lower(), right_on=df_StatoVerb["Stato Verbale"].str.lower(),how="left",suffixes=('', '_MP'))
df_VerbPassChius_0=df_VerbPassChius_daRep[["DocAcq","ID_Verbale"]]
#Add Verbali Chius
df_VerbPassChius = pd.concat([df_VerbPassChius, df_VerbPassChius_0], ignore_index=True, sort=False)
#Keep single occurrence Verbali Att
df_VerbPassChius=df_VerbPassChius.sort_values(by=["ID_Verbale"]).drop_duplicates("DocAcq", keep="last")
df_VerbPassChius=pd.merge(df_VerbPassChius,df_StatoVerb_byCod[["ID_Verbale","Stato Verbale"]],left_on=df_VerbPassChius["ID_Verbale"], right_on=df_StatoVerb_byCod["ID_Verbale"],how="left",suffixes=('', '_chius'))
df_VerbPassChius=df_VerbPassChius.rename(columns={"Stato Verbale":"Stat_VerbCh","ID_Verbale":"ID_VerbCh"})

#sostituisce valori specifici e li imposta a valori nulli
w_find=df_VerbPassChius["Stat_VerbCh"].str.lower().isin(["verbale assente",0])
df_VerbPassChius.loc[w_find,"Stat_VerbCh"] = ""

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_VerbPassSAL.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_VerbPassSAL.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_VerbPassSAL.memory_usage(deep=True)) )


#df_VerbPassChius=df_VerbPassChius.loc[df_VerbPassChius["Stat_VerbCh"] == "VERBALE ASSENTE", "Stat_VerbCh"] = ""