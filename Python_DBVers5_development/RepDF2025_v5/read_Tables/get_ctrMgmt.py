#Get file from Contract Management
import glob
import os
from Util_leggeDir import dir_dict
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  

#from read_rilasci import df_Rilasci_task
import sys

##Read working directories 
#new_dir = os.getcwd()
wrk_ctrMgmt_inp=dir_dict["ctrMgmt"]
wrk_stagDir=dir_dict["stagDir"]

df_ODAG_ctrMgmt=pd.read_excel(wrk_stagDir+"\\ctrmgmt\\"+"ODAG_ctrMgmt.xlsx", sheet_name=0,  decimal=',')
df_ODA_ctrMgmt=pd.read_excel(wrk_stagDir+"\\ctrmgmt\\"+"ODA_ctrMgmt.xlsx", sheet_name=0,  decimal=',')
df_ODA_ctrMgmt.rename({'ODA':"DocAcq"}, axis=1, inplace=True)
#df_ODA_ctrMgmt=df_ODA_ctrMgmt["ODA"].astype(int)
#df_ODAG_ctrMgmt=df_ODAG_ctrMgmt["ODAGnbr"].astype(int)
pass