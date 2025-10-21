#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
#rom Util_leggeParams import parm_dict
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  

import os
import sys
from Util_logging import logger
#Read dummy tables
#rom read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 

wPythProc="read_cdc"
import timeit  #funzione timeit per valutare durata estrazione
tempoInzInt=timeit.default_timer()

print(os.getcwd())

##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]

logger.info('start reading CDC')
#df_RisInt_prec["C.I.D."]=df_RisInt_prec["C.I.D."].str.replace(r'^(0+)', '').fillna('0')
df_CDC_Dip=pd.read_excel(wrk_dir_inp+"\\"+"Tab_CdCDip.xlsx", dtype={ "CdC":str,"Sigla":"category"})
df_CDC_Pmo=pd.read_excel(wrk_dir_inp+"\\"+"Tab_CdCPMO.xlsx", dtype={"PMO": int, "Centro di Costo":str})
#f_CDC_Pmo["PMO"]=df_CDC_Pmo["PMO"].astype(str).str.zfill(3)
df_CDC_Pmo.rename({'Centro di Costo': 'CdC'}, axis=1, inplace=True)
df_CDC_Dip = df_CDC_Dip.drop(["PMO","DIP"], axis=1)
df_CDC=pd.merge(df_CDC_Dip,df_CDC_Pmo[["CdC","PMO"]],left_on="CdC", right_on="CdC",how="left",suffixes=('', '__cid'))
df_CDC.rename({'Sigla': 'Dip'}, axis=1, inplace=True)

tempoFinInt = timeit.default_timer()

print("Durata Attivita *** ",tempoFinInt-tempoInzInt)
logger.info('end reading ' + str(df_CDC.shape[0])+" Durata processo: " + wPythProc + " "+ str(round(tempoFinInt-tempoInzInt,3)))

pass