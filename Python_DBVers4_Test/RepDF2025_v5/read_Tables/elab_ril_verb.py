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
from read_rilasci import df_Rilasci_resp, df_Rilasci_subt
from read_verbAtt import df_Elab_RCA_VerbSpalm, df_RCA_Verb_Fatt, df_RCA_Verb,df_RCA_Verb_riep
import os
import sys

##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]


df_Elab_RCA_VerbSpalm=pd.merge(df_Elab_RCA_VerbSpalm,df_Rilasci_resp[["Cod Rilascio","resp_ril","resp_dmd","resp_ril_dip","resp_dmd_dip"]],left_on="cod rilascio", right_on="Cod Rilascio",how="left",suffixes=('', '_cup'))
df_RCA_Verb_Fatt=pd.merge(df_RCA_Verb_Fatt,df_Rilasci_resp[["Cod Rilascio","resp_ril","resp_dmd","resp_ril_dip","resp_dmd_dip"]],left_on="Cod Rilascio", right_on="Cod Rilascio",how="left",suffixes=('', '_cup'))
df_RCA_Verb_riep=pd.merge(df_RCA_Verb_riep,df_Rilasci_subt[["Cod_Rilascio","resp_ril","resp_dmd","resp_ril_dip","resp_dmd_dip","Tit_Rilascio","LOB"]],left_on="cod rilascio", right_on="Cod_Rilascio",how="left",suffixes=('', '_cup'))
df_Rilasci_subt=pd.merge(df_Rilasci_subt,df_RCA_Verb_riep[["cod rilascio","tot_InPrep","tot_notApprov","statusApprov","statusPrep"]],left_on="Cod_Rilascio", right_on="cod rilascio",how="left",suffixes=('', '_cup'))
pass