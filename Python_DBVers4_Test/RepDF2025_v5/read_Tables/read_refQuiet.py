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
wPythProc="read_refQuiet"
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

df_quiet_ref=pd.read_excel(wrk_dir_inp+"\\"+"Tab_ReferentiQuietanze.xlsx",dtype={'ODAG': str, 'CodFornitore': str})  

df_quiet_ref = df_quiet_ref.iloc[:,:6]
#df_Rilasci_resp=pd.merge(df_Rilasci_resp,df_RisInt_byName[["Cognome Nome","Int_Dip"]],left_on="Nome Referente", right_on="Cognome Nome",how="left",suffixes=('', '_dmr'))
df_quiet_ref_grp=df_quiet_ref.groupby(["Contratto","ODAG","CodFornitore"])[["Referente","Mail di contatto"]].agg("; ".join).reset_index()
df_quiet_ref_grp["ODAG_forn"]=df_quiet_ref_grp["ODAG"]+"_"+df_quiet_ref_grp["CodFornitore"]
#
tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_quiet_ref_grp.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_quiet_ref_grp.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_quiet_ref_grp.memory_usage(deep=True)) )
pass