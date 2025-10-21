#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task

import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_ODAG import df_ODAGMacr
#from read_forn import df_forn_byName,df_forn_byName
from read_Bdo_Oda import df_BDOODALin,df_forn_byName
from read_map import df_PDC, df_MAP_noverb, df_MAP_byOda_per, df_MAP_byOdaLin,df_MAP_byMAP,df_MAP_byOdaLin_per,df_MAP_noverb_byLin
#from read_verbPass import  df_VerbPassSAL, df_VerbPassAtt,df_VerbPassChius
import os
import sys
from Util_logging import logger
wPythProc="elab_Oda_Periodo"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

wDtFormat = "%d/%m/%Y" 
wActFormat = "%d.%m.%Y" 
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
wrk_frmperiod_dt=parm_dict["parm_frmperiod_dt"]
wrk_toperiod_dt=parm_dict["parm_toperiod_dt"]
wrk_recentdays=parm_dict["parm_lastBDO_days"]

df_periodi=pd.read_excel(wrk_dir_inp+"\\"+"Tab_Periodi.xlsx", dtype={"Periodo": str,"PeriodoMese": str}, decimal=',')#,"PeriodoFine": str,"PeriodoIniz": str
df_periodi["index"]="1"
df_periodi=df_periodi[(df_periodi["Periodo"]>=wrk_frmperiod_dt)&(df_periodi["Periodo"]<=wrk_toperiod_dt)]

wList_date=["PeriodoIniz"	,"PeriodoFine"]
#Date format
for d in wList_date:
    df_periodi[d]=pd.to_datetime(df_periodi[d],errors='coerce')

#df_periodi["PeriodoIniz"]=pd.to_datetime(df_periodi["PeriodoIniz"], format='%d/%m/%Y')
#df_periodi["PeriodoIniz"]=pd.to_datetime(df_periodi["PeriodoIniz"]);df_periodi["PeriodoIniz"]=df_periodi["PeriodoIniz"].dt.strftime(wDtFormat)
#df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"]);df_periodi["PeriodoFine"]=df_periodi["PeriodoFine"].dt.strftime(wDtFormat)
#df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"], format='%d/%m/%Y')

#df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"], format='%d.%m.%Y')
w_Iniz=df_periodi["PeriodoIniz"].min() 
w_Fine=df_periodi["PeriodoFine"].max() 

#df_ODALin=df_BDOODALin[(df_BDOODALin['Tipo']=='Z3')|(df_BDOODALin['Tipo']=='ZI')]

df_BDOODALin=pd.merge(df_BDOODALin,df_MAP_byOdaLin[["DocAcq","Linea","MAP_Cons","Qta_Cons"]
                                                   ],on=["DocAcq","Linea"], how="left",suffixes=('', '_sal'))
df_BDOODALin["Qta_resid"]=df_BDOODALin["Forn_qta"]-df_BDOODALin["Qta_Cons"]
df_BDOODALin["Val_resid"]=df_BDOODALin["Forn_val"]-df_BDOODALin["MAP_Cons"]

df_BDOODALin=df_BDOODALin.replace(np.nan,0)
#df_ODALin["DataFine"]=pd.to_datetime(df_ODALin["DataFine"])

df_ODALin_recent=df_BDOODALin[((df_BDOODALin['TipoDocAcq']=='Z3')|(df_BDOODALin['TipoDocAcq']=='ZI'))&(df_BDOODALin["DataFine"]>w_Iniz)]

#get DEC & RUP if available
from get_ctrMgmt import df_ODAG_ctrMgmt,df_ODA_ctrMgmt

df_ODALin_recent=pd.merge(df_ODALin_recent,df_ODA_ctrMgmt[["DocAcq","DEC","RUP"]],on=["DocAcq"], how="left")

#df_ODA = df_ODA.loc[:,~df_ODA.T.duplicated(keep='last')]
#df_ODA=df_ODA.rename(columns={"Doc. acq.":"ODA"})
#df_ODA["Val_Resid"]=df_ODA["Tot_ODA"]-df_ODA["Tot_MAP"]

#df_ODA[df_ODA_per["DataFine"] == 0, "Datafine"] = "2099-12-31 00:00:00"
#df_ODA_recent=

df_ODALin_recent["index"]="1"
df_ODALin_recent_per=pd.merge(df_ODALin_recent,df_periodi,left_on="index",right_on="index", how="left")
#Convert from datetime to string
#df_ODALin_recent_per['Anno competenza'] = df_ODALin_recent_per['PeriodoFine'].dt.year
df_ODALin_recent_per=df_ODALin_recent_per.loc[(df_ODALin_recent_per["DataIniz"]<=df_ODALin_recent_per["PeriodoFine"]) & (df_ODALin_recent_per["DataFine"]>=df_ODALin_recent_per["PeriodoIniz"])]

df_ODALin_recent_per['Anno competenza'] = df_ODALin_recent_per['PeriodoFine'].dt.year
df_ODALin_recent_per['Mese competenza'] = df_ODALin_recent_per['PeriodoFine'].dt.month
#df_ODALin_recent_per["DataFine"]=pd.to_datetime(df_ODALin_recent_per["DataFine"]); df_ODALin_recent_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
df_ODALin_recent_per=pd.merge(df_ODALin_recent_per,df_MAP_byOdaLin_per[["DocAcq","Anno competenza","Mese competenza","Linea","MAPZero"]],on=["DocAcq","Anno competenza","Mese competenza","Linea"], how="left")

df_ODA_recent_per=df_ODALin_recent_per.groupby(["DocAcq","CIG", "ROI","ROI_dip", "ROI_cod", "ROI_CdC","DataIniz", "DataFine",  "Forn_RTI","Forn_nome", "RDI", "Descrizione DocAcq", "StartDate","TipoDocAcq","StatoDocAcq","StatoDocAcq_cod","DEC","RUP","MAPZero","Anno competenza","Mese competenza"],observed=True)[["Forn_val","MAP_Cons","Val_resid"]].sum().reset_index()

#Create Flag to check if Consuntivi todo or not required
df_ODA_recent_per["Flag_daCons"]=np.where((df_ODA_recent_per["MAPZero"]!=1)|(~((df_ODA_recent_per["MAP_Cons"]>0)|(df_ODA_recent_per["Val_resid"]<100))),
True,False)
#legge lo stato del Flag da consuntivare a livello di ODA Periodo
df_ODALin_recent_per=pd.merge(df_ODALin_recent_per,df_ODA_recent_per[["DocAcq","Anno competenza","Mese competenza","Flag_daCons"]],on=["DocAcq","Anno competenza","Mese competenza"], how="left")


#df_BDO=pd.merge(df_BDO,df_periodi,left_on="index",right_on="index", how="left")

#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#df_BDO=df_BDO[(df_BDO["Data decorrenza"]<=df_BDO["PeriodoFine"]) & (df_BDO["Data scadenza"]>=df_BDO["PeriodoIniz"])]

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_ODALin_recent_per.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_ODALin_recent_per.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_ODALin_recent_per.memory_usage(deep=True)) )



#Group by BDO
#



#df_BDO = df_BDO.loc[(df_BDO['Data decorrenza']<w_Fine) & (df_BDO['Data scadenza']>w_Iniz)]
pass