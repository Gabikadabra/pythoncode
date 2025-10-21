#%%
#New Reportistica ---
###Codice da completare
#Read IF in progress

from Util_leggeDir import dir_dict
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_ODAG import df_ODAGMacr
#from read_rilasci import df_Rilasci_task
import os
import sys
##Read working directories 
from read_cdc import df_CDC
#new_dir = os.getcwd()
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task,listfilter_Incarico

#Read dummy tables
from Util_logging import logger
wPythProc="read_IFproc"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items


wrk_dir_inp=dir_dict["inpDir"]
wrk_dir_IF=dir_dict["IFprcDir"]

#IF list
df_IFprc=pd.read_excel(wrk_dir_IF+"\\"+"RdI-IF da verificare.xlsx",sheet_name=0)
# Convert entire DataFrame to string
df_IFprc=df_IFprc.astype(str)
#Rename dframe columns
df_IFprc.rename(columns={ df_IFprc.columns[0]:"RDI",df_IFprc.columns[1]:"ROI",df_IFprc.columns[2]:"ROI_CdC",df_IFprc.columns[3]:"Descrizione RDI",df_IFprc.columns[4]:"PMO",df_IFprc.columns[5]:"StatoIF_byPMO",df_IFprc.columns[6]:"chkIF_byPMO",df_IFprc.columns[7]:"chkIFsub_byPMO",df_IFprc.columns[8]:"Note_PMO",df_IFprc.columns[9]:"BdO_byPMO-BO",df_IFprc.columns[10]:"BdONote_byPMO-BO",df_IFprc.columns[11]:"Stato_elab",df_IFprc.columns[12]:"Data_lavorazione_byPMO-BO" }, inplace = True)
df_IFprc["BdO_byPMO-BO"]=df_IFprc["BdO_byPMO-BO"].str[0:10]
#Var BDO
df_BDOvarprc=pd.read_excel(wrk_dir_IF+"\\"+"RdI-IF da verificare.xlsx",sheet_name=1,dtype={"RDI":str,"BDO":str})
df_BDOvarprc.rename(columns={df_BDOvarprc.columns[0]:"DocAcq",df_BDOvarprc.columns[2]:"ROI",df_BDOvarprc.columns[3]:"ROI_CdC",df_BDOvarprc.columns[4]:"Descrizione DocAcq",df_BDOvarprc.columns[5]:"PMO",df_BDOvarprc.columns[6]:"StatoIF_portale",df_BDOvarprc.columns[7]:"chk_byPMO",df_BDOvarprc.columns[8]:"chksub_byPMO",df_BDOvarprc.columns[9]:"Note_PMO",df_BDOvarprc.columns[10]:"BdO_byPMO-BO",df_BDOvarprc.columns[11]:"BdONote_byPMO-BO",df_BDOvarprc.columns[12]:"StatoIF",df_BDOvarprc.columns[13]:"Data_lavorazione_byPMO-BO",df_BDOvarprc.columns[15]:"Data_approv_port_byPMO-BO"} , inplace = True)

df_IFprc["Data_lavorazione_byPMO-BO"]=pd.to_datetime(df_IFprc["Data_lavorazione_byPMO-BO"], format='%Y-%m-%d',errors="coerce")
df_BDOvarprc["Data_lavorazione_byPMO-BO"]=pd.to_datetime(df_BDOvarprc["Data_lavorazione_byPMO-BO"], format='%Y-%m-%d',errors="coerce")

df_BDOvarprc["Data_approv_port_byPMO-BO"]=pd.to_datetime(df_BDOvarprc["Data_approv_port_byPMO-BO"], format='%Y-%m-%d',errors="coerce")
df_IFprc = df_IFprc.drop_duplicates(subset=['RDI'],keep='last')
df_IFprc['ROI'] = df_IFprc['ROI'].str.upper()
df_IFprc=pd.merge(df_IFprc,df_RisInt_byName[["RisInt_nome","Int_Dip","CID","CdC"]],left_on="ROI", right_on="RisInt_nome",how="left",suffixes=('', '_int'))

df_IFprc["ROI_CdC"]=np.where(df_IFprc["CdC"].notnull(),df_IFprc["CdC"],
                    np.where(df_IFprc["ROI_CdC"].notnull(),df_IFprc["ROI_CdC"], 
                             "n/a"))

df_IFprc=pd.merge(df_IFprc,df_CDC[["CdC","Dip"]],left_on="ROI_CdC", right_on="CdC",how="left",suffixes=('', '_cdc'))



df_IFprc=df_IFprc.rename(columns={"Int_Dip_int":"ROI_dip" })
#remove all unnamed columns
df_IFprc = df_IFprc.loc[:, ~df_IFprc.columns.str.contains('^Unnamed')]

"""#begin format columns - reduce size"""
wLi=["RDI"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
#wLf=["Quantità","Prezzo lordo",	"Valore netto","Importo Subappalto"];wDf={value:"float" for value in wLf}
wLc=["ROI","ROI_CdC","PMO"];wDc={value:"category" for value in wLc}
wLnbr=wLi
df_IFprc[wLnbr]=df_IFprc[wLnbr].fillna(0)
#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_IFprc= df_IFprc.astype(convert_dict)
"""#end format columns - reduce size"""




df_BDOvarprc = df_BDOvarprc.drop_duplicates(subset=['DocAcq'],keep='last')
df_BDOvarprc = df_BDOvarprc.astype({"StatoIF":"category","DocAcq":int})


#get list of modify BDO
df_BDOvar=pd.read_excel(wrk_dir_inp+"\\"+"Da BDO_PortaleN.xlsx",sheet_name=(1), dtype={})
df_BDOvar.rename(columns={df_BDOvar.columns[0]:"DocAcq",df_BDOvar.columns[5]:"StatoIF",df_BDOvar.columns[8]:"UltimaIF_appr","Fornitore":"Forn_RTI","CID":"ROI_CID","CdC":"ROI_CdC","Descrizione BDO":"Descrizione DocAcq"}, inplace = True)


wList_date=["Data caricamento","Data invio ROI","Data approvazione ROI", "Data rifiuto ROI","Data invio PMO","Data approvazione ROI", "Data rifiuto ROI","Data invio PMO","Data approvazione PMO", "Data rifiuto PMO","Versione corrente BDO","Data Versione corrente BDO", "Data effettiva decorrenza BDO"]
#Date format
for d in wList_date:
    df_BDOvar[d]=pd.to_datetime(df_BDOvar[d],errors='coerce')

"""#begin format columns - reduce size"""
wLi=["DocAcq"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
#wLf=["Quantità","Prezzo lordo",	"Valore netto","Importo Subappalto"];wDf={value:"float" for value in wLf}
wLc=["StatoIF","Divisione","Centro di Costo","UltimaIF_appr","CTRM","Codifica numerica documento"];wDc={value:"category" for value in wLc}
wLnbr=wLi
df_BDOvar[wLnbr]=df_BDOvar[wLnbr].fillna(0)
#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_BDOvar= df_BDOvar.astype(convert_dict)
"""#end format columns - reduce size"""



df_BDOvar=pd.merge(df_BDOvar,df_RisInt_byName[["RisInt_nome","Int_Dip","CID","CdC"]],left_on="ROI", right_on="RisInt_nome",how="left",suffixes=('', '_int'))
df_BDOvar=df_BDOvar.rename(columns={"Int_Dip":"ROI_dip","CID":"ROI_CID","CdC":"ROI_CdC" })
df_BDOvar=pd.merge(df_BDOvar,df_BDOvarprc[["BdO_byPMO-BO","BdONote_byPMO-BO","StatoIF","Data_lavorazione_byPMO-BO","Data_approv_port_byPMO-BO","DocAcq"]],left_on="DocAcq", right_on="DocAcq",how="left",suffixes=('', '_BDOvar'))

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_IFprc.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_IFprc.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_IFprc.memory_usage(deep=True)) )


pass

