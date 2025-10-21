#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from read_cdc import df_CDC
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from Util_logging import logger


import os
import sys
from Util_leggeParams import parm_dict
#Read dummy tables
#rom read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 

wPythProc="read_Interni"
import timeit  #funzione timeit per valutare durata estrazione
tempoInzInt=timeit.default_timer()
#logging.config.fileConfig('logging2.conf')

# create logger
#logger = logging.getLogger('simpleExample')

wrk_environm=value = parm_dict.get("parm_environment", "Develop")
#2)	Gestione DIRECTORY CORRENTE (contestualizzazione cartella python sulla base dell’utente)
#da inserire -> selezione directory sulla base dell'utente che utilizza Python
current_dir=os.getcwd()
wDtFormat = "%d/%m/%Y" 
wYrFormat="%Y"
	#condizione if
#%%
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]

#print(df_CDC)
#df_CDC=pd.read_excel(wrk_dir_inp+"\\"+"Tab_CdCDip.xlsx", dtype={"DIP": str, "CdC":str})
logger.debug('start reading df_RisInt')
#Get Internal resources
df_RisInt=pd.read_excel(wrk_dir_inp+"\\"+"Tab_RisInt.xlsx", dtype={"C.I.D.": int, "Tipo di Attività":str})
df_RisInt_prec=pd.read_excel(wrk_dir_inp+"\\"+"Tab_RisInt_prec.xlsx", dtype={"C.I.D.": int, "Tipo di Attività":str})

df_RisInt_var=pd.read_excel(wrk_dir_inp+"\\"+"Tab_RisInt_Var.xlsx", dtype={"CID.": str}, skiprows=2)
#df_RisInt_var["CID"]=df_RisInt_var["CID"].astype(str).str.zfill(3)
df_RisInt_var = df_RisInt_var[(df_RisInt_var['CodVar'] != "I")]

#df_RisInt["C.I.D."]=df_RisInt["C.I.D."].astype(str).str.zfill(3)

df_RisInt = pd.concat([df_RisInt, df_RisInt_prec], ignore_index=True, sort=False)
df_RisInt= df_RisInt.loc[:,~df_RisInt.columns.str.contains('^unnam', case=False)]

#Set Cid with 3 characters 0 aligned
#df_RisInt['C.I.D.']=df_RisInt['C.I.D.'].str.zfill(3) 

df_RisInt=pd.merge(df_RisInt,df_CDC[["CdC","Dip","PMO"]],left_on="CdC", right_on="CdC",how="left",suffixes=('', '_Dip'))
df_RisInt=pd.merge(df_RisInt,df_RisInt_var[["CID","CodVar"]],left_on="C.I.D.", right_on="CID",how="left",suffixes=('', '_Dip'))
# Drop columns based on index positions
df_RisInt=df_RisInt[["C.I.D.","Cognome Nome","Status dipendente","CdC","Pianif. orario di lavoro","Utenza","Matricola","E-mail","Tipo di Attività","Desc. Estesa  Tipo Attività","Desc. Breve  Tipo Attività","Tariffa Costo","Tariffa Ricavo","Dip","PMO","CID","CodVar"]]
#df_RisInt = df_RisInt.drop(df_RisInt.columns[[3,4,5,7,8,9,11,12,13,15,17,19,20,25,26,27,28,29]], axis=1)


df_RisInt.drop('CID', axis=1, inplace=True)
df_RisInt.rename({'C.I.D.': 'CID','Dip':'Int_Dip',"Cognome Nome":"RisInt_nome"}, axis=1, inplace=True)
#Dataframe column format conversion
convert_dict = {"Status dipendente": "category", "Pianif. orario di lavoro": "category","CodVar":"category"}   
df_RisInt= df_RisInt.astype(convert_dict)


#Remove duplicate by CID
df_RisInt_byCID  =  df_RisInt.drop_duplicates( subset = ["CID"],  keep = "first" )
#Remove duplicate by Name
df_RisInt_byName  =  df_RisInt.drop_duplicates( subset = ["RisInt_nome"],  keep = "first" )

tempoFinInt = timeit.default_timer()

print("Durata Attivita *** ",tempoFinInt-tempoInzInt)
#logger.info('end reading ' + str(df_RisInt.shape[0])+" Durata processo: " + wPythProc + " "+ str(round(tempoFinInt-tempoInzInt,3)))
logger.info('end reading ' + str(df_RisInt.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_RisInt.memory_usage().sum())+" Durata processo: " +  str(round(tempoFinInt-tempoInzInt,3))+ " Dataframe dimensione " + str(df_RisInt.memory_usage(deep=True)) )

pass
if __name__ == '__main__':
    logger.info("This is an info message")