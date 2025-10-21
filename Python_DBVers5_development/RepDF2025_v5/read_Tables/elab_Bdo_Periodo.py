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
from read_Interni import df_RisInt_byCID, df_RisInt_byName
from read_ODAG import df_ODAGMacr
from elab_Oda_Periodo import df_BDOODALin
from read_Bdo_Oda import df_BDOODA

from read_BdoLin import df_BDOLin, df_BDO
from read_map import df_PDC_byLin, df_MAP_noverb, df_MAP_byOdaLin,df_MAP_byMAP,df_MAP_byOdaLin_per,df_MAP_noverb_byLin
from read_verbPass import  df_VerbPassSAL, df_VerbPassAtt,df_VerbPassChius, wDict_Verb
from read_BefQuiet import df_BEF_BDOLin
import os
import sys
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task
#Read dummy tables
#rom read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
from Util_logging import logger
wPythProc="elab_Bdo_Periodo"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()


wDtFormat = "%d/%m/%Y" 
wActFormat = "%d.%m.%Y" 
##Read working directories 


import timeit  #funzione timeit per valutare durata estrazione
tempoInz0 = timeit.default_timer()


#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
wrk_frmperiod_dt=parm_dict["parm_frmperiod_dt"]
wrk_toperiod_dt=parm_dict["parm_toperiod_dt"]
wrk_recentdays=parm_dict["parm_lastBDO_days"]

df_periodi=pd.read_excel(wrk_dir_inp+"\\"+"Tab_Periodi.xlsx", dtype={"Periodo": str,"PeriodoMese": str,"PeriodoFine": str,"PeriodoIniz": str}, decimal=',')
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

#Clean Verbali
df_VerbPassChius = df_VerbPassChius.loc[df_VerbPassChius["ID_VerbCh"]>0]
df_VerbPassAtt = df_VerbPassAtt.loc[df_VerbPassAtt["ID_VerbAt"]>0]
df_VerbPassSAL = df_VerbPassSAL.loc[df_VerbPassSAL["ID_VerbSAL"]>0]

df_BDOLin=pd.merge(df_BDOLin,df_VerbPassChius[["DocAcq","Stat_VerbCh","ID_VerbCh"]],on="DocAcq", how="left",suffixes=('', '_sal'))
df_BDOLin=pd.merge(df_BDOLin,df_VerbPassAtt[["DocAcq","Stat_VerbAt","ID_VerbAt"]],on="DocAcq", how="left",suffixes=('', '_att'))


#df_BDO_DataEff=df_BDOODALin.groupby(
#["Doc. acq."]).agg(DataEffAtt= ("StartDate","max")
#     ).reset_index()
#df_BDO_DataEff=df_BDO_DataEff.rename(columns={"Doc. acq.":"BDO"})
#Add
#df_BDOODALin=pd.merge(df_BDOODALin,df_MAP_byOdaLin[["DocAcq","Linea","MAP_ConsT","Qta_ConsT"]],left_on="BDOLin",right_on="BDOLin", how="left",suffixes=('', '_att'))
#df_BDOODALin.fillna({"ValoreCons":0,"QtaCons":0,"Quantità":0},inplace=True)
#df_BDOODALin["ValResiduo"]=df_BDOODALin["Valore netto"]-df_BDOODALin["MAP_ConsT"]

#df_BDOODALin["QtaRes"]=np.where((df_BDOODALin["ValResiduo"]<1),0,(df_BDOODALin["Forn_qta"]-df_BDOODALin["Qta_ConsT"]))



# To create df_MAP_byOda_per without 

#df_BDO=pd.merge(df_BDO,df_VerbPassChius[["DocAcq","Stat_VerbCh","ID_VerbCh"]],left_on="BDO",right_on="Numero BDO", how="left",suffixes=('', '_sal'))
#df_BDO=pd.merge(df_BDO,df_VerbPassAtt[["DocAcq","Stat_VerbAt","ID_VerbAt"]],left_on="BDO",right_on="Numero BDO", how="left",suffixes=('', '_att'))
#df_BDO=pd.merge(df_BDO,df_BDOODA[["DocAcq","StartDate"]],on="DocAcq", how="left",suffixes=('', '_str'))

#df_BDO.loc[df_BDO["DataEffAtt"].isnull(), "DataEffAtt"] = "Data decorrenza"
#df_BDO["DataEffAtt"]=pd.to_datetime(df_BDO["DataEffAtt"],format="%Y-%m-%d %H:%M:%S", errors='coerce')
df_BDOLin_recent= df_BDOLin.loc[df_BDOLin["DataFine"]>w_Iniz]


df_BDOLin_recent["index"]="1"
#wdatamin=df_periodi["PeriodoIniz"].min()
#reduce number of lines to the recent ones


df_BDOLin_recent_per=pd.merge(df_BDOLin_recent,df_periodi,left_on="index",right_on="index", how="left")
#df_BDOLin_per["BDOLin"]=df_BDOLin_per["BDO"]+"-"+df_BDOLin_per["Posizione"].str.zfill(3)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
df_BDOLin_recent_per=df_BDOLin_recent_per[(df_BDOLin_recent_per["DataFine"]>=df_BDOLin_recent_per["PeriodoIniz"])]
#Join BDO Lin & YYYY-MM
#df_BDOLin_per["BDOLin_periodo"]=df_BDOLin_per["BDOLin"]+"_"+df_BDOLin_per["Periodo"]
#df_BDOLin_per["BDO_periodo"]=df_BDOLin_per["BDO"]+"_"+df_BDOLin_per["Periodo"]
df_BDOLin_recent_per["Anno competenza"]= df_BDOLin_recent_per["Periodo"].str.slice(0,4)
df_BDOLin_recent_per["Mese competenza"]= df_BDOLin_recent_per["Periodo"].str.slice(5,7)
df_BDOLin_recent_per=df_BDOLin_recent_per.astype({"Anno competenza":int,"Mese competenza":int})

df_BDOLin_recent_per=pd.merge(df_BDOLin_recent_per,df_BDOODALin[["DocAcq", "Linea","Forn_qta","Qta_Cons","Qta_resid","Rilevante PNRR"]],on=["DocAcq", "Linea"], how="left",suffixes=('', '_pdc'))
df_BDOLin_recent_per=pd.merge(df_BDOLin_recent_per,df_MAP_byOdaLin_per,on=["DocAcq", "Linea","Mese competenza", "Anno competenza"], how="left",suffixes=('', '_noCons'))


#df_BDOLin_per=pd.merge(df_BDOLin_per,df_MAPCons_byOdaLin_per[["BDOLin_periodo","ValoreCons","Qta"]],left_on="BDOLin_periodo",right_on="BDOLin_periodo", how="left",suffixes=('', '_bdoper'))
#df_BDOLin_per=pd.merge(df_BDOLin_per,df_MAP_zero[["BDO_periodo","Periodo"]],left_on="BDO_periodo",right_on="BDO_periodo", how="left",suffixes=('', '_mapzero'))
df_BDOLin_recent_per=pd.merge(df_BDOLin_recent_per,df_MAP_noverb_byLin,on=["DocAcq", "Linea", "Mese competenza", "Anno competenza"], how="left",suffixes=('', '_noverb'))
df_BDOLin_recent_per=pd.merge(df_BDOLin_recent_per,df_VerbPassSAL,on=["DocAcq", "Mese competenza", "Anno competenza"], how="left",suffixes=('', '_sal'))
df_BDOLin_recent_per=pd.merge(df_BDOLin_recent_per,df_BEF_BDOLin[["DocAcq", "Mese competenza", "Anno competenza","TotBEF","Linea"]],on=["DocAcq","Linea", "Mese competenza", "Anno competenza"], how="left",suffixes=('', '_noverb'))
#df_BDOLin_per=pd.merge(df_BDOLin_per,df_BDOODALin[["BDOLin","Quantità","QtaCons","QtaRes"]],left_on="BDOLin",right_on="BDOLin", how="left",suffixes=('', '_noverb'))
#Replace na values for each column
df_BDOLin_recent_per.fillna({"Stat_VerbCh":"","ID_VerbCh":0,"Stat_VerbAt":"","ID_VerbAt":0,"Stat_VerbSAL":"","ID_VerbSAL":0,"ValoreCons":0,"Periodo_mapzero":"","Quantità":0,"Qta":0,"MAP_noverb":0,"ValorePDC":0,"PNRR":""},inplace=True)

df_BDOLin_recent_per.fillna({'MAP_Cons':0,'MAP_NoConsT':0,"ValorePDC":0,"TotBEF":0,"Qta_Cons":0},inplace=True)

#df_BDOLin_per["QtaRes"]=np.where(df_BDOLin_per["Lin_Resid"]<1,0,(df_BDOLin_per["Quantità"]-df_BDOLin_per["QtaCons"]))
#calculate remaining months 
df_BDOLin_recent_per["ResMonths"]=(df_BDOLin_recent_per["Data_MLS"].dt.year - df_BDOLin_recent_per["PeriodoFine"].dt.year) * 12 + (df_BDOLin_recent_per["Data_MLS"].dt.month - df_BDOLin_recent_per["PeriodoFine"].dt.month)
#df_BDOLin_per["ResMonths"]=df_BDOLin_per["ResMonths"].astype(int)
df_BDOLin_recent_per.fillna({"ResMonths":0,"Qta_resid":0},inplace=True)

df_BDOLin_recent_per = df_BDOLin_recent_per.astype({"Stat_VerbCh":"category","Stat_VerbAt":"category","Stat_VerbSAL":"category","ID_VerbCh":"int","ID_VerbAt":"int","ID_VerbSAL":"int"})

df_BDO_recent_per=df_BDOLin_recent_per.groupby(
["DocAcq","Anno competenza","Mese competenza"]).agg(
                                            ID_VerbCh=("ID_VerbCh","max"), 
                                            ID_VerbAt=("ID_VerbAt","max"), 
                                            ID_VerbSAL=("ID_VerbSAL","max"),
                                            MAP_NoCons=("MAP_NoCons","sum"),
                                            MAP_Cons=("MAP_Cons","sum"),
                                            MAPZero=("MAPZero","max"),
                                            MAP_noverb=("MAP_noverb","sum"),
                                            ValorePDC=("ValorePDC","sum"),
                                            Val_resid=("Val_resid","sum"),
                                            TotBEF=("TotBEF","sum"),
                                            PNRR=("Rilevante PNRR","max")
     ).reset_index()

#                                            Stat_VerbCh=("Stat_VerbCh","max"),
#                                            Stat_VerbAt=("Stat_VerbAt","max"),
#                                            Stat_VerbSAL=("Stat_VerbSAL","max"),
#### Verificare USO di BDO_riepMAP

df_BDO_riepMAP=df_BDOLin_recent_per.groupby(
["DocAcq"]).agg(
                                            ID_VerbCh=("ID_VerbCh","max"), 
                                            ID_VerbAt=("ID_VerbAt","max"), 
                                            ID_VerbSAL=("ID_VerbSAL","max"),
                                            MAP_NoCons=("MAP_NoCons","sum"),
                                            MAP_Cons=("MAP_Cons","sum"),
                                            MAP_noverb=("MAP_noverb","sum"),
                                            ValorePDC=("ValorePDC","sum"),
                                            TotBEF=("TotBEF","sum")
     ).reset_index()
#                                            Stat_VerbCh=("Stat_VerbCh","max"),
#                                            Stat_VerbAt=("Stat_VerbAt","max"),
   #                                         Stat_VerbSAL=("Stat_VerbSAL","max"),
    #                                        Mapzero=("Mapzero","max"),

df_BDO_recent_per["Flag_daCons"]=np.where(((df_BDO_recent_per["ValorePDC"]>0)|(df_BDO_recent_per["MAP_NoCons"]>0))|(~((df_BDO_recent_per["MAP_Cons"]>0)|(df_BDO_recent_per["MAPZero"]==1)|(df_BDO_recent_per["Val_resid"]<100))&(df_BDO_recent_per["ID_VerbCh"]<7)),True,False)
df_BDOLin_recent_per=pd.merge(df_BDOLin_recent_per,df_BDO_recent_per[["DocAcq","Anno competenza","Mese competenza","Flag_daCons"]],on=["DocAcq","Anno competenza","Mese competenza"], how="left")

df_BDO_recent_per=pd.merge(df_BDO_recent_per,df_BDO[["Macrocl_nbr","Descrizione DocAcq","DocAcq","DataIniz","DataFine","Forn_RTI","ROI_dip","ROI_CID","ROI_CdC","Fornitori","Subfornitori","Macrocl_txt","Forn_val","Importo Attestato","Val_resid","ROI","StatoDocAcq","DocAcq_FlagCh","ODAGnbr","DEC","RUP"]],on="DocAcq", how="left",suffixes=('', '_bdo'))



### da controllare le seguenti istruzioni se siano ancora operative
df_BDOLin_recent_per=pd.merge(df_BDOLin_recent_per,df_BDO[["ODAGnbr","ODAG","Macrocl_txt","Descrizione DocAcq","StatoDocAcq","DataCreaz","StartDate","DataIniz","DataFine","Forn_RTI","Forn_RTI_cod","ROI_dip","ROI_CID","ROI_CdC","Macrocl_nbr","DocAcq" ]],on="DocAcq", how="left",suffixes=('', '_bdo'))
df_BDO_recent_per=pd.merge(df_BDO_recent_per,df_BDOODA,left_on="DocAcq",right_on="DocAcq", how="left",suffixes=('', '_bdo'))
#Create Flag to check if Consuntivi todo or not required

df_BDO_recent_per["Stat_VerbCh"]=df_BDO_recent_per["ID_VerbCh"].replace(wDict_Verb)
df_BDO_recent_per["Stat_VerbAt"]=df_BDO_recent_per["ID_VerbAt"].replace(wDict_Verb)
df_BDO_recent_per["Stat_VerbSAL"]=df_BDO_recent_per["ID_VerbSAL"].replace(wDict_Verb)


###
#w_inclBDO=["2024332010", "2025330717"]
#mask1=(df_BDO_per['BDO'].isin(w_inclBDO)) 
#dfBDO_per_filtA=df_BDO_per[mask1]


#df_BDO_per["Flag_daCons"]=np.where(((df_BDO_per["ValorePDC"]>0)|(df_BDO_per["MAP_NoCons"]>0))|(~((df_BDO_per["MAP_Cons"]>0)|(df_BDO_per["Mapzero"]>"0")|(df_BDO_per["Importo Residuo"]<100))&(df_BDO_per["ID_VerbCh"]<"7")),"S","N")
#check if canone and consuntiv not done

#df_BDOLin_recent_per["FlagCan_daCons"]=np.where((~((df_BDOLin_recent_per["ValoreCons"]>0)|(df_BDOLin_recent_per["ValorePDC"]>0)|(df_BDOLin_recent_per["MAP_NoCons"]>0))&(df_BDOLin_recent_per["ID_VerbCh"]<"7")&(df_BDOLin_recent_per["Val_Resid"]>100)&(df_BDOLin_recent_per["Tipologia Fornitura"]=="a Canone")&(df_BDOLin_recent_per["ResMonths"]<df_BDOLin_recent_per["QtaRes"])),"S","N")
#df_BDO_per["Flag_daCons"]=np.where(((df_BDO_per["ValorePDC"]>0)|(df_BDO_per["MAP_NoCons"]>0))|(~((df_BDO_per["MAP_Cons"]>0)|(df_BDO_per["Mapzero"]>"0")|(df_BDO_per["ID_VerbAt"]=="9")|(df_BDO_per["Importo Residuo"]>100))),
#"S","N")
#df_BDOLin_recent_per=pd.merge(df_BDOLin_recent_per,df_BDO_per[["BDO_periodo","Flag_daCons","ROI","StatoBDO"]],left_on="BDO_periodo",right_on="BDO_periodo", how="left",suffixes=('', '_bdo'))

df_BDO=pd.merge(df_BDO,df_BDO_riepMAP[["DocAcq","MAP_Cons","MAP_NoCons","MAP_noverb","ValorePDC","TotBEF"]],on="DocAcq", how="left",suffixes=('', '_bdo'))

#df_BDO=pd.merge(df_BDO,df_periodi,left_on="index",right_on="index", how="left")

#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#df_BDO=df_BDO[(df_BDO["Data decorrenza"]<=df_BDO["PeriodoFine"]) & (df_BDO["Data scadenza"]>=df_BDO["PeriodoIniz"])]

#df_BDOLin=pd.merge(df_BDOLin,df_BDO[["ODAGnbr","ODAG","Macrocl_txt","StatoDocAcq","DataCreaz","StartDate","Forn_RTI_cod","ROI_dip","ROI_CID","ROI_CdC","Macrocl_nbr","DocAcq" ]],on="DocAcq", how="left",suffixes=('', '_bdo'))

df_BDOLin=pd.merge(df_BDOLin,df_BDOODALin,on=["DocAcq", "Linea"], how="left",suffixes=('', '_pdc'))

#Group by BDO
#





tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_BDOLin_recent_per.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_BDOLin_recent_per.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDOLin_recent_per.memory_usage(deep=True)) )

#df_BDO = df_BDO.loc[(df_BDO['Data decorrenza']<w_Fine) & (df_BDO['Data scadenza']>w_Iniz)]
pass