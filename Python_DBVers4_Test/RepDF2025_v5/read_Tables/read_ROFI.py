#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
#Read dummy tables
#wrk_environm=value = parm_dict.get("parm_environment", "Develop")
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
#from read_Interni import df_RisInt_byCID,df_RisInt_byName
import os
import sys
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_home=dir_dict["homeDir"]

df_ROFI=pd.read_excel(wrk_dir_home+"\\Documenti - Share PMO Data\\ImportDati\\Staging\\ROFI\\"+"Tabella_ROFI.xlsx",dtype={'BDO': str})  

df_ROFI["ROFI"]=df_ROFI["COGNOME"]+" "+df_ROFI["NOME"]
df_ROFI['RAGIONE_SOCIALE'] = df_ROFI['RAGIONE_SOCIALE'].str.upper()
pass