#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
#Read dummy tables
#rom read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 


#Read dummy tables
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task
from read_Interni import df_RisInt_byName
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
#from read_Interni import df_RisInt_byCID,df_RisInt_byName
import os
import sys
from read_rilasci import df_RilasciPPA_sub

from Util_logging import logger
wPythProc="elab_monitVerb"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

#
from read_verbAtt import df_elab_RilPPA, df_elab_Imp_Spalm, df_elab_Verb_Spalm, df_elab_Verb_Fatt, df_elab_RefRil,df_RilInc, df_elab_Rilasci,df_elab_Incarichi, df_RCA_Verb, df_elab_Verbali, df_Incarichi,df_RefInc,df_RCA_Verb_FirmatoG,df_RCA_Verb_ChiusoG, df_RilasciPPA,df_RCA_Azioni
wrk_RptDir=dir_dict["rptDir"]
import time
wDtFormat = "%d/%m/%Y" 
wTimestr = time.strftime("%Y%m%d")
df_elab_Rilasci=df_elab_Rilasci.rename(columns={"Cod Rilascio":"cod rilascio" })


#Aggancia il CUP 
df_Rilasci_byCUP=df_RilasciPPA.groupby(['Codice Rilascio']).agg(CodiceCup = ("Codice CUP","unique")
     ).reset_index()
df_Rilasci_byCUP=df_Rilasci_byCUP.rename(columns={"Codice Rilascio":"cod rilascio" })
#Convert list column in a string 
df_Rilasci_byCUP['CodiceCup']=[','.join(map(str, l)) for l in df_Rilasci_byCUP['CodiceCup']]

#Get Incarichi by Rilascio
df_RilIncarichi=pd.merge(df_elab_Rilasci,df_elab_Incarichi,left_on="cod incarico",right_on="cod incarico", how="left",suffixes=('', '_sal'))
#Get CUP by Rilascio
df_RilIncarichi=pd.merge(df_RilIncarichi,df_Rilasci_byCUP,left_on="cod rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))
#Get Valore by Verbali Firmati
df_RilIncarichi=pd.merge(df_RilIncarichi,df_RCA_Verb_FirmatoG,left_on="cod rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))
#Get Valore by Verbali Chiusi
df_RilIncarichi=pd.merge(df_RilIncarichi,df_RCA_Verb_ChiusoG,left_on="cod rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))

#get list of Verbali
df_elab_Verbali_bycodVerb=df_elab_Verbali.groupby(["cod verbale",	"cod rilascio",	"tipo verbale",	"stato verbale",	"dt creazione",	"dt firma",	"dt chiusura"
], dropna=False).agg(
ImportoVerbale= ("importo verbale","sum")).reset_index()
#Get Info about Rilascio
df_elab_Verbali_bycodVerb=pd.merge(df_elab_Verbali_bycodVerb,df_RilIncarichi,left_on="cod rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))
df_elab_Verbali_bycodVerb=pd.merge(df_elab_Verbali_bycodVerb,df_RCA_Azioni,left_on="cod verbale",right_on="codice sap", how="left",suffixes=('', '_sal'))

df_elab_Verb_Fatt=pd.merge(df_elab_Verb_Fatt,df_elab_Verbali_bycodVerb,left_on="Cod Rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))
wDtFormat = "%d/%m/%Y" 
wTimestr = time.strftime("%Y%m%d")

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_elab_Verbali_bycodVerb.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_elab_Verbali_bycodVerb.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_elab_Verbali_bycodVerb.memory_usage(deep=True)) )

pass

