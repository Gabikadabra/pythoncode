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
import xlrd
#from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_rilasci import df_Rilasci_task
import os
#import os.path
import shutil
import sys
##Read working directories 
#new_dir = os.getcwd()
#get dir & intermediate dir for use of Excel workbooks with problems using read_excel
#wrk_dir_intermed=dir_dict["xlsxDir"]
wrk_dir_inp=dir_dict["inpDir"]

#shutil.copy2(wrk_dir_inp+"\\"+"Da PDC_PortaleN.xlsx", wrk_dir_intermed+"\\"+"Da PDC_PortaleN.xlsx")
#wrk_dir_inpA=r"\Users\fcaneri\OneDrive - ARIA S.p.A\PMO Condiviso\Report test\Excel_dacaricareSP"
#Check if file exist
#os.path.isfile(wrk_dir_inpA+"\\"+"Da PDC_PortaleN.xlsx")
df_RptHeader=pd.read_excel(wrk_dir_inp+"\\"+"Da PDC_PortaleN.xlsx", sheet_name="Richiesta")
wRptDate= df_RptHeader["Data Estrazione"].iloc[0]
print(wRptDate)