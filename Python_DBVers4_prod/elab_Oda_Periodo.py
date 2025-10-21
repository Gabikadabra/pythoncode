#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from Util_leggeParm import parm_dict
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
from read_ODAG import df_ODAGMacr
from read_forn import df_forn_byName
from read_BdoOda import df_BDOODALin
from read_map import df_PDC, df_MAP_noverb, df_MAP_noCons, df_MAP_zero,df_MAP_byOda_per, df_MAP_byOdaLin,df_MAP_byMAP,df_MAP_byOdaLin_per,df_MAP_noverb_byLin
from read_verbPass import  df_VerbPassSAL, df_VerbPassAtt,df_VerbPassChius
import os
import sys
wDtFormat = "%d/%m/%Y" 
wActFormat = "%d.%m.%Y" 
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
wrk_frmperiod_dt=parm_dict["parm_frmperiod_dt"]
wrk_toperiod_dt=parm_dict["parm_toperiod_dt"]
wrk_recentdays=parm_dict["parm_lastBDO_days"]

df_periodi=pd.read_excel(wrk_dir_inp+"\\"+"Tab_Periodi.xlsx", dtype={"Periodo": str,"PeriodoMese": str,"PeriodoFine": str,"PeriodoIniz": str}, decimal=',')
df_periodi["index"]="1"
df_periodi=df_periodi[(df_periodi["Periodo"]>=wrk_frmperiod_dt)&(df_periodi["Periodo"]<=wrk_toperiod_dt)]

df_periodi["PeriodoIniz"]=pd.to_datetime(df_periodi["PeriodoIniz"], format='%d/%m/%Y')
#df_periodi["PeriodoIniz"]=pd.to_datetime(df_periodi["PeriodoIniz"]);df_periodi["PeriodoIniz"]=df_periodi["PeriodoIniz"].dt.strftime(wDtFormat)
df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"]);df_periodi["PeriodoFine"]=df_periodi["PeriodoFine"].dt.strftime(wDtFormat)
df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"], format='%d/%m/%Y')

#df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"], format='%d.%m.%Y')
w_Iniz=df_periodi["PeriodoIniz"].min() 
w_Fine=df_periodi["PeriodoFine"].max() 

df_BDOODALin=pd.merge(df_BDOODALin,df_MAP_byOdaLin[["BDOLin","ValoreCons"]],left_on="BDOLin",right_on="BDOLin", how="left",suffixes=('', '_sal'))

df_ODALin=df_BDOODALin[(df_BDOODALin['Tipo']=='Z3')|(df_BDOODALin['Tipo']=='ZI')]
df_ODALin=df_ODALin.replace(np.nan,0)
#df_ODALin["DataFine"]=pd.to_datetime(df_ODALin["DataFine"])

df_ODA=df_ODALin.groupby(["Doc. acq.","Numero CIG", "ROI","ROI_dip", "ROI_cod", "ROI_CdC","InizVal.", "FineVal.",  "Forn_main","Forn_nome", "RDI", "Testo testata BDO", "StartDate","Tipo","Stato","Stato_cod","DEC","Descrizione_DEC"])[["Valore netto","ValoreCons"]].sum().reset_index().rename(columns={'Valore netto':'Tot_ODA','Val_Resid':'Tot_Resid','ValoreCons':'Tot_MAP'})
df_ODA = df_ODA.loc[:,~df_ODA.T.duplicated(keep='last')]
df_ODA=df_ODA.rename(columns={"Doc. acq.":"ODA"})
df_ODA["Val_Resid"]=df_ODA["Tot_ODA"]-df_ODA["Tot_MAP"]

df_ODA["index"]="1"
df_ODA_per=pd.merge(df_ODA,df_periodi,left_on="index",right_on="index", how="left")
#Convert from datetime to string
df_ODA_per["PeriodoIniz"]=df_ODA_per["PeriodoIniz"].dt.strftime("%Y-%m-%d")
df_ODA_per["PeriodoFine"]=df_ODA_per["PeriodoFine"].dt.strftime("%Y-%m-%d")
#Convert from datetime to string
#pd.to_datetime(df_ODA_per["FineVal."])
#pd.to_datetime(df_ODA_per["FineVal."])

df_ODA_per["InizVal."]=pd.to_datetime(df_ODA_per["InizVal."])
df_ODA_per.loc[df_ODA_per["FineVal."] == 0, "FineVal."] = "2099-12-31 00:00:00"
df_ODA_per["FineVal."]=pd.to_datetime(df_ODA_per["FineVal."])
df_ODA_per["PeriodoIniz"]=pd.to_datetime(df_ODA_per["PeriodoIniz"])
df_ODA_per["PeriodoFine"]=pd.to_datetime(df_ODA_per["PeriodoFine"])
#df_ODA_per["FineVal."]=df_ODA_per["FineVal."].str[0:10]
#df_ODA_per["InizVal."]=df_ODA_per["InizVal."].str[0:10]
#df_BDOLin_per["BDOLin"]=df_ODA_per["BDO"]+"-"+df_ODA_per["Posizione"].str.zfill(3)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
df_ODA_per=df_ODA_per[(df_ODA_per["InizVal."]<=df_ODA_per["PeriodoFine"]) & (df_ODA_per["FineVal."]>=(df_ODA_per["PeriodoIniz"]))]
#Join BDO Lin & YYYY-MM
df_ODA_per["ODA_periodo"]=df_ODA_per["ODA"]+"_"+df_ODA_per["Periodo"]
#df_ODA_per=pd.merge(df_ODA_per,df_ODA_per[["BDO_periodo","Periodo"]],left_on="BDO_periodo",right_on="BDO_periodo", how="left",suffixes=('', '_mapzero'))

df_ODA_per=pd.merge(df_ODA_per,df_MAP_byOda_per[["ODA_periodo","ValoreMAP"]],left_on="ODA_periodo",right_on="ODA_periodo", how="left",suffixes=('', '_oda'))
df_ODA_per=pd.merge(df_ODA_per,df_MAP_zero[["BDO_periodo","Periodo"]],left_on="ODA_periodo",right_on="BDO_periodo", how="left",suffixes=('', '_mapzero'))

#df_BDOODA_full["InizVal."]=pd.to_datetime(df_BDOODA_full["InizVal."]); df_BDOODA_full["InizVal."]=df_BDOODA_full["InizVal."].dt.strftime('%d.%m.%Y')
#df_BDOODA_full["FineVal."]=pd.to_datetime(df_BDOODA_full["FineVal."]); df_BDOODA_full["FineVal."]=df_BDOODA_full["FineVal."].dt.strftime('%d.%m.%Y')
#df_BDOODA_full["Data Fine Milestone"]=pd.to_datetime(df_BDOODA_full["Data Fine Milestone"]); df_BDOODA_full["Data Fine Milestone"]=df_BDOODA_full["Data Fine Milestone"].dt.strftime('%d.%m.%Y')
#df_BDOODA_full["Data Effettiva Attivazione"]=pd.to_datetime(df_BDOODA_full["Data Effettiva Attivazione"]); df_BDOODA_full["Data Effettiva Attivazione"]=df_BDOODA_full["Data Effettiva Attivazione"].dt.strftime('%d.%m.%Y')

df_ODA_per["InizVal."]=pd.to_datetime(df_ODA_per["InizVal."]); df_ODA_per["InizVal."]=df_ODA_per["InizVal."].dt.strftime('%d.%m.%Y')
df_ODA_per["FineVal."]=pd.to_datetime(df_ODA_per["FineVal."]); df_ODA_per["FineVal."]=df_ODA_per["FineVal."].dt.strftime('%d.%m.%Y')

#Create Flag to check if Consuntivi todo or not required
df_ODA_per["Flag_daCons"]=np.where((df_ODA_per["Periodo_mapzero"]>"0")|(~((df_ODA_per["ValoreMAP"]>0)|(df_ODA_per["Val_Resid"]<100))),
"S","N")



#df_BDO=pd.merge(df_BDO,df_periodi,left_on="index",right_on="index", how="left")

#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#df_BDO=df_BDO[(df_BDO["Data decorrenza"]<=df_BDO["PeriodoFine"]) & (df_BDO["Data scadenza"]>=df_BDO["PeriodoIniz"])]




#Group by BDO
#



#df_BDO = df_BDO.loc[(df_BDO['Data decorrenza']<w_Fine) & (df_BDO['Data scadenza']>w_Iniz)]
pass