#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from Util_leggeParm import parm_dict
#from read_forn import df_forn_byName
#Read dummy tables
#from read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
#from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_rilasci import df_Rilasci_task
import os
import sys

wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

from Util_logging import logger
wPythProc="read_BefQuiet"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()


##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]

#%%
logger.info('start reading quiet')

#wrk_dir_inp=wrk_home_dir+r"\Documenti - Share PMO Data\ImportDati"
#wrk_dir_inp2=wrk_home_dir+r"\ImportDati"
#Prepare write file 
wf = open(wrk_dir_inp+"\\"+'Da Quiet_parteA.txt',"w")
txt0=""
txtHdr0=""
nbr=0
#read file da copiare
with open(wrk_dir_inp+"\\"+'Da Quiet.txt',"r") as f:
	for fline in f:
		if fline=="\n":
			pass
		elif "Pagabile" in fline:
			txtHdr0= fline
			print("db",fline)
		elif "Data reg." in fline:
			print("data",fline)
			txtHdr1= fline
			txtHdr=txtHdr0.replace("\n","")+"\t"+txtHdr1
			wf.write(txtHdr)
		elif "SI" in fline[0:20] or "NO" in fline[0:20]:
			print("data",fline)
			txt0=fline
			#2410202433
#		elif "2410202433" in fline:
		else:
			txt1=fline
			txt=txt0.replace("\n","")+"\t"+txt1
			#if nbr <10 : 
			wf.write(txt)
			#nbr=nbr+1
wf.close()

dfCruscQuiet = pd.read_csv(wrk_dir_inp+"\\"+'Da Quiet_parteA.txt', sep='\t', encoding='latin-1',dtype=str)
# Example 3: Rename multiple columns
dfCruscQuiet.rename(columns={dfCruscQuiet.columns[2]: 'Pagata', dfCruscQuiet.columns[3]: 'Pagabile' },inplace=True)
dfCrQuiet_sel= dfCruscQuiet[["BEF","Pagata","Pagabile","MAP","Per.Comp"]]
dfCrQuiet_grp=dfCrQuiet_sel[dfCrQuiet_sel.duplicated(subset=["BEF","Pagata","Pagabile"],keep=False)]
pass