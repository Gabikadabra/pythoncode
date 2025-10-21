#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
#from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_rilasci import df_Rilasci_resp, df_Rilasci_subt
import os
import sys
import re
from read_cdc import df_CDC

from Util_logging import logger
wPythProc="read_verbAtt"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()

def rename_col_by_index(dataframe, index_mapping):
    dataframe.columns = [index_mapping.get(i, col) for i, col in enumerate(dataframe.columns)]
    return dataframe
#Remove chars not numeric form column
def remove_chars(s):
#2025.04.08 begin
#   return re.sub("[^0-9,]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
    a= re.sub("[^0-9,.]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
    a=a.strip()
    return a.replace(",",".")
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
f_reptname="RptPPA" #"PMRep117-test"
winp_path = wrk_dir_inp+"\\Staging\\cicloattivo"
wRept_output=winp_path+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
workbk_Out=writer.book
wcolnames={0:"PPA_descr",1:"PPA_cod",2:"CdC_ref",3:"UO",4:"Incarico_cod",5:"Incarico_vers",6:"Incarico_vers_inz",7:"Incarico_cont",8:"Incarico_titolo",9:"Incarico_stato",10:"Incarico_riacc",11:"AzRichiesta",12:"Pian_stato",13:"Incarico_tip",14:"CdC_refInc",15:"ROA",16:"Capitolo",17:"Capitolo_tipo",18:"Costo2025_IVA",19:"Costo2025Est_IVA",20:"Costo2025Int_IVA",21:"CostoContr2025",22:"DeltaContr2025",23:"Costo2026_IVA",24:"Costo2026Est_IVA",25:"Costo2026Int_IVA",26:"CostoContr2026",27:"DeltaContr2026",28:"Costo2027_IVA",29:"Costo2027Est_IVA",30:"Costo2027Int_IVA",31:"CostoContr2027",32:"DeltaContr2027",33:"IncaricoTriennio",34:"Missione",35:"Programma",36:"DirezCompetente"}
df_MonitorInc=pd.read_excel(winp_path+"\\"+"Da_Monitoraggio_incarichi.xlsx",sheet_name="PdR per strutture") #, heeusecols="A:AK"
df_MonitorInc = rename_col_by_index(df_MonitorInc,wcolnames)
# Renaming columns using the function
df_MonitorInc=pd.merge(df_MonitorInc,df_CDC[["CdC","Dip"]],left_on="CdC_ref", right_on="CdC",how="left",suffixes=('', '_Dip'))
df_MonitorInc=pd.merge(df_MonitorInc,df_CDC[["CdC","Dip"]],left_on="CdC_refInc", right_on="CdC",how="left",suffixes=('', '_Dip'))
#rename columns
wrenMAPColumns={"CdC_Dip":"CdC_ref_dip","Dip_Dip":"CdC_refInc_dip" }
df_MonitorInc=df_MonitorInc.rename(columns=wrenMAPColumns)


tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_MonitorInc.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_MonitorInc.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_MonitorInc.memory_usage(deep=True)) )


pass