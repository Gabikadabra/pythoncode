#New Reportistica --- Braca od koce
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
from read_BdoOda import df_BDOODALin

from read_BdoLin import df_BDOLin, df_BDO
from read_map import df_PDC_byLin, df_MAPCons_byOdaLin_per, df_MAP_noverb, df_MAP_noCons_byLin, df_MAP_zero, df_MAP_byOdaLin,df_MAP_byMAP,df_MAP_byOdaLin_per,df_MAP_noverb_byLin
from read_verbPass import  df_VerbPassSAL, df_VerbPassAtt,df_VerbPassChius
from read_BefQuiet import df_BEF_BDOLin
from read_remapDip import w_eccez_Rilasci, w_eccez_Task,w_eccez_Bdo
import os
import sys
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

df_periodi["PeriodoIniz"]=pd.to_datetime(df_periodi["PeriodoIniz"], format='%d/%m/%Y')
#df_periodi["PeriodoIniz"]=pd.to_datetime(df_periodi["PeriodoIniz"]);df_periodi["PeriodoIniz"]=df_periodi["PeriodoIniz"].dt.strftime(wDtFormat)
df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"]);df_periodi["PeriodoFine"]=df_periodi["PeriodoFine"].dt.strftime(wDtFormat)
df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"], format='%d/%m/%Y')

#df_periodi["PeriodoFine"]=pd.to_datetime(df_periodi["PeriodoFine"], format='%d.%m.%Y')
w_Iniz=df_periodi["PeriodoIniz"].min() 
w_Fine=df_periodi["PeriodoFine"].max() 

#Clean Verbali
df_VerbPassChius = df_VerbPassChius.loc[df_VerbPassChius["ID_VerbCh"]>"0"]
df_VerbPassAtt = df_VerbPassAtt.loc[df_VerbPassAtt["ID_VerbAt"]>"0"]
df_VerbPassSAL = df_VerbPassSAL.loc[df_VerbPassSAL["ID_VerbSAL"]>"0"]

df_BDOLin=pd.merge(df_BDOLin,df_VerbPassChius[["Numero BDO","Stat_VerbCh","ID_VerbCh"]],left_on="BDO",right_on="Numero BDO", how="left",suffixes=('', '_sal'))
df_BDOLin=pd.merge(df_BDOLin,df_VerbPassAtt[["Numero BDO","Stat_VerbAt","ID_VerbAt"]],left_on="BDO",right_on="Numero BDO", how="left",suffixes=('', '_att'))
#inserito flag PNRR



df_BDO_DataEff=df_BDOODALin.groupby(
["Doc. acq."]).agg(DataEffAtt= ("StartDate","max")
     ).reset_index()
df_BDO_DataEff=df_BDO_DataEff.rename(columns={"Doc. acq.":"BDO"})
#Add
df_BDOODALin=pd.merge(df_BDOODALin,df_MAP_byOdaLin[["BDOLin","ValoreCons","QtaCons"]],left_on="BDOLin",right_on="BDOLin", how="left",suffixes=('', '_att'))
df_BDOODALin.fillna({"ValoreCons":0,"QtaCons":0,"Quantità":0},inplace=True)
df_BDOODALin["ValResiduo"]=df_BDOODALin["Valore netto"]-df_BDOODALin["ValoreCons"]

df_BDOODALin["QtaRes"]=np.where((df_BDOODALin["ValResiduo"]<1),0,(df_BDOODALin["Quantità"]-df_BDOODALin["QtaCons"]))

df_BDOLin=pd.merge(df_BDOLin,df_BDOODALin[["BDOLin","PNRR","Quantità","QtaRes","QtaCons"]],on="BDOLin", how="left",suffixes=('', '_att'))

# To create df_MAP_byOda_per without 

df_BDO=pd.merge(df_BDO,df_VerbPassChius[["Numero BDO","Stat_VerbCh","ID_VerbCh"]],left_on="BDO",right_on="Numero BDO", how="left",suffixes=('', '_sal'))
df_BDO=pd.merge(df_BDO,df_VerbPassAtt[["Numero BDO","Stat_VerbAt","ID_VerbAt"]],left_on="BDO",right_on="Numero BDO", how="left",suffixes=('', '_att'))
df_BDO=pd.merge(df_BDO,df_BDO_DataEff[["BDO","DataEffAtt"]],left_on="BDO",right_on="BDO", how="left",suffixes=('', '_str'))

df_BDO.loc[df_BDO["DataEffAtt"].isnull(), "DataEffAtt"] = "Data decorrenza"
df_BDO["DataEffAtt"]=pd.to_datetime(df_BDO["DataEffAtt"],format="%Y-%m-%d %H:%M:%S", errors='coerce')


df_BDOLin["index"]="1"
wdatamin=df_periodi["PeriodoIniz"].min()
#reduce number of lines to the recent ones
df_BDOLin_red= df_BDOLin.loc[df_BDOLin["Data scadenza"]>=wdatamin]

df_BDOLin_per=pd.merge(df_BDOLin_red,df_periodi,left_on="index",right_on="index", how="left")
df_BDOLin_per["BDOLin"]=df_BDOLin_per["BDO"]+"-"+df_BDOLin_per["Posizione"].str.zfill(3)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
df_BDOLin_per=df_BDOLin_per[(df_BDOLin_per["Data decorrenza"]<=df_BDOLin_per["PeriodoFine"]) & (df_BDOLin_per["Data scadenza"]>=df_BDOLin_per["PeriodoIniz"])]
#Join BDO Lin & YYYY-MM
df_BDOLin_per["BDOLin_periodo"]=df_BDOLin_per["BDOLin"]+"_"+df_BDOLin_per["Periodo"]
df_BDOLin_per["BDO_periodo"]=df_BDOLin_per["BDO"]+"_"+df_BDOLin_per["Periodo"]

df_BDOLin_per=pd.merge(df_BDOLin_per,df_PDC_byLin[["ValorePDC","BDOLin_periodo"]],left_on="BDOLin_periodo",right_on="BDOLin_periodo", how="left",suffixes=('', '_pdc'))
df_BDOLin_per=pd.merge(df_BDOLin_per,df_MAP_noCons_byLin[["BDOLin_periodo","MAP_NoCons"]],left_on="BDOLin_periodo",right_on="BDOLin_periodo", how="left",suffixes=('', '_noCons'))
df_BDOLin_per=pd.merge(df_BDOLin_per,df_MAPCons_byOdaLin_per[["BDOLin_periodo","ValoreCons","Qta"]],left_on="BDOLin_periodo",right_on="BDOLin_periodo", how="left",suffixes=('', '_bdoper'))
df_BDOLin_per=pd.merge(df_BDOLin_per,df_MAP_zero[["BDO_periodo","Periodo"]],left_on="BDO_periodo",right_on="BDO_periodo", how="left",suffixes=('', '_mapzero'))
df_BDOLin_per=pd.merge(df_BDOLin_per,df_MAP_noverb_byLin[["BDOLin_periodo","MAP_noverb"]],left_on="BDOLin_periodo",right_on="BDOLin_periodo", how="left",suffixes=('', '_noverb'))
df_BDOLin_per=pd.merge(df_BDOLin_per,df_VerbPassSAL[["BDO_periodo","ID_VerbSAL","Stat_VerbSAL"]],left_on="BDO_periodo",right_on="BDO_periodo", how="left",suffixes=('', '_sal'))
df_BDOLin_per=pd.merge(df_BDOLin_per,df_BEF_BDOLin[["BDOLin_periodo","TotBEF"]],left_on="BDOLin_periodo",right_on="BDOLin_periodo", how="left",suffixes=('', '_noverb'))
df_BDOLin_per=pd.merge(df_BDOLin_per,df_BDOODALin[["BDOLin","Quantità","QtaCons","QtaRes"]],left_on="BDOLin",right_on="BDOLin", how="left",suffixes=('', '_noverb'))
#Replace na values for each column
df_BDOLin_per.fillna({"Stat_VerbCh":"","ID_VerbCh":"","Stat_VerbAt":"","ID_VerbAt":"","Stat_VerbSAL":"","ID_VerbSAL":"","ValoreCons":0,"Periodo_mapzero":"","Quantità":0,"Qta":0,"MAP_noverb":0,"ValorePDC":0,"PNRR":""},inplace=True)

df_BDOLin_per.fillna({'MAP_Cons':0,'MAP_NoCons':0,"ValorePDC":0,"TotBEF":0,"QtaCons":0},inplace=True)

#df_BDOLin_per["QtaRes"]=np.where(df_BDOLin_per["Lin_Resid"]<1,0,(df_BDOLin_per["Quantità"]-df_BDOLin_per["QtaCons"]))
#calculate remaining months 
df_BDOLin_per["ResMonths"]=(df_BDOLin_per["Data_MLS"].dt.year - df_BDOLin_per["PeriodoFine"].dt.year) * 12 + (df_BDOLin_per["Data_MLS"].dt.month - df_BDOLin_per["PeriodoFine"].dt.month)
#df_BDOLin_per["ResMonths"]=df_BDOLin_per["ResMonths"].astype(int)
df_BDOLin_per.fillna({"ResMonths":0,"QtaRes":0},inplace=True)

df_BDO_per=df_BDOLin_per.groupby(
["BDO_periodo","BDO","Periodo"], observed=True).agg(
                                            Stat_VerbCh=("Stat_VerbCh","max"),
                                            ID_VerbCh=("ID_VerbCh","max"), 
                                            Stat_VerbAt=("Stat_VerbAt","max"),
                                            ID_VerbAt=("ID_VerbAt","max"), 
                                            Stat_VerbSAL=("Stat_VerbSAL","max"),
                                            ID_VerbSAL=("ID_VerbSAL","max"),
                                            MAP_NoCons=("MAP_NoCons","sum"),
                                            MAP_Cons=("ValoreCons","sum"),
                                            Mapzero=("Periodo_mapzero","max"),
                                            MAP_noverb=("MAP_noverb","sum"),
                                            ValorePDC=("ValorePDC","sum"),
                                            TotBEF=("TotBEF","sum"), 
                                            PNRR=("PNRR","max")
     ).reset_index()

df_BDO_riepMAP=df_BDOLin_per.groupby(
["BDO"]).agg(
                                            Stat_VerbCh=("Stat_VerbCh","max"),
                                            ID_VerbCh=("ID_VerbCh","max"), 
                                            Stat_VerbAt=("Stat_VerbAt","max"),
                                            ID_VerbAt=("ID_VerbAt","max"), 
                                            Stat_VerbSAL=("Stat_VerbSAL","max"),
                                            ID_VerbSAL=("ID_VerbSAL","max"),
                                            MAP_NoCons=("MAP_NoCons","sum"),
                                            MAP_Cons=("ValoreCons","sum"),
                                            Mapzero=("Periodo_mapzero","max"),
                                            MAP_noverb=("MAP_noverb","sum"),
                                            ValorePDC=("ValorePDC","sum"),
                                            TotBEF=("TotBEF","sum")
     ).reset_index()

df_BDOLin_per=pd.merge(df_BDOLin_per,df_BDO[["Contratto Macroclasse","Descrizione Bdo","BDO","Data decorrenza","Data scadenza","RTI_cod","ROI_dip","ROI_CID","ROI_CdC","Forn_nome","ODAGMacr_txt","ODAGMacr_nbr" ]],left_on="BDO",right_on="BDO", how="left",suffixes=('', '_bdo'))
df_BDO_per=pd.merge(df_BDO_per,df_BDO[["Contratto Macroclasse","Descrizione Bdo","BDO","Data decorrenza","Data scadenza","RTI_cod","ROI_dip","ROI_CID","ROI_CdC","Fornitori","Subfornitori","ODAGMacr_txt","ODAGMacr_nbr","Valore Ordine","Importo Attestato","Importo Residuo","ROI","StatoBDO","Chius_Amm","Fornitore RTI","Numero Contratto Gara","DEC","RUP"]],left_on="BDO",right_on="BDO", how="left",suffixes=('', '_bdo'))
#Create Flag to check if Consuntivi todo or not required
###
w_inclBDO=["2024332010", "2025330717"]
mask1=(df_BDO_per['BDO'].isin(w_inclBDO)) 
dfBDO_per_filtA=df_BDO_per[mask1]



df_BDO_per["Flag_daCons"]=np.where(((df_BDO_per["ValorePDC"]>0)|(df_BDO_per["MAP_NoCons"]>0))|(~((df_BDO_per["MAP_Cons"]>0)|(df_BDO_per["Mapzero"]>"0")|(df_BDO_per["Importo Residuo"]<100))&(df_BDO_per["ID_VerbCh"]<"7")),"S","N")
#check if canone and consuntiv not done

df_BDOLin_per["FlagCan_daCons"]=np.where((~((df_BDOLin_per["ValoreCons"]>0)|(df_BDOLin_per["ValorePDC"]>0)|(df_BDOLin_per["MAP_NoCons"]>0))&(df_BDOLin_per["ID_VerbCh"]<"7")&(df_BDOLin_per["Val_Resid"]>100)&(df_BDOLin_per["Tipologia Fornitura"]=="a Canone")&(df_BDOLin_per["ResMonths"]<df_BDOLin_per["QtaRes"])),"S","N")
#df_BDO_per["Flag_daCons"]=np.where(((df_BDO_per["ValorePDC"]>0)|(df_BDO_per["MAP_NoCons"]>0))|(~((df_BDO_per["MAP_Cons"]>0)|(df_BDO_per["Mapzero"]>"0")|(df_BDO_per["ID_VerbAt"]=="9")|(df_BDO_per["Importo Residuo"]>100))),
#"S","N")
df_BDOLin_per=pd.merge(df_BDOLin_per,df_BDO_per[["BDO_periodo","Flag_daCons","ROI","StatoBDO"]],left_on="BDO_periodo",right_on="BDO_periodo", how="left",suffixes=('', '_bdo'))

df_BDO=pd.merge(df_BDO,df_BDO_riepMAP[["BDO","MAP_Cons","MAP_NoCons","MAP_noverb","ValorePDC","TotBEF"]],left_on="BDO",right_on="BDO", how="left",suffixes=('', '_bdo'))

#df_BDO=pd.merge(df_BDO,df_periodi,left_on="index",right_on="index", how="left")

#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#dfBDOLin_per["DataFine"]=pd.to_datetime(dfBDOLin_per["DataFine"]); dfBDOLin_per["DataFine"]=dfBDOLin_per["DataFine"].dt.strftime(wDtFormat)
#df_BDO=df_BDO[(df_BDO["Data decorrenza"]<=df_BDO["PeriodoFine"]) & (df_BDO["Data scadenza"]>=df_BDO["PeriodoIniz"])]

### Converte le strutture di riferimento dei BDO sulla base delle eccezioni riportate nel file read_remapDip
df_BDO=pd.merge(df_BDO,w_eccez_Bdo,on="BDO",how="left",suffixes=('', '_pos'))
df_BDO["ROI_dip"]= np.where(df_BDO['eccez_Dip'].notnull(),df_BDO['eccez_Dip'],df_BDO['ROI_dip'])

df_BDOLin=pd.merge(df_BDOLin,w_eccez_Bdo,on="BDO",how="left",suffixes=('', '_pos'))
df_BDOLin["ROI_dip"]= np.where(df_BDOLin['eccez_Dip'].notnull(),df_BDOLin['eccez_Dip'],df_BDOLin['ROI_dip'])

df_BDO_per=pd.merge(df_BDO_per,w_eccez_Bdo,on="BDO",how="left",suffixes=('', '_pos'))
df_BDO_per["ROI_dip"]= np.where(df_BDO_per['eccez_Dip'].notnull(),df_BDO_per['eccez_Dip'],df_BDO_per['ROI_dip'])
#Group by BDO
#

tempoFin0 = timeit.default_timer()

print("Durata Attivita BDO Periodo*** ",tempoFin0-tempoInz0)

#df_BDO = df_BDO.loc[(df_BDO['Data decorrenza']<w_Fine) & (df_BDO['Data scadenza']>w_Iniz)]
pass