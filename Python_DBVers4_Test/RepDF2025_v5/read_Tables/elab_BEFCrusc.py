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
#from read_Interni import df_RisInt_byCID,df_RisInt_byName
import os
import sys
from read_BefQuiet import df_BEF_Map,df_quiet
#from read_BefQuiet import 
from read_cruscQuiet import dfCrQuiet_grp
from read_BdoLin import df_BDO
from read_Bdo_Oda import df_forn_byCod 
from read_verbPass import df_VerbPassSAL,df_VerbPassChius
#from read_map import  df_MAP_bytask # df_MAP_noCons, df_MAP_zero, df_PDC, df_MAP_byMAP,
#from read_prog import df_prog_bytask # df_prog_byTipoRis, df_prog_int, df_prog_altro, df_prog_ammort, df_Rilasci_task
#from read_BefQuiet import df_FattPass_IVA,df_BEF_Map
#from read_rilasci import df_Rilasci_task
##Read working directories 
#new_dir = os.getcwd()
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task,listfilter_Incarico

from Util_logging import logger
wPythProc="elab_BEFCrusc"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items
#sort values index
df_VerbPassSAL=df_VerbPassSAL.loc[df_VerbPassSAL["ID_VerbSAL"]>1]
df_VerbPassSAL = df_VerbPassSAL.sort_values(by = ['Numero BDO',"BDO_periodo"],ignore_index=True)
#df2.loc[0, 'PeriodoNext'] = df2.loc[0, 'Periodo']
#For each row in dataframe
for i in range(1, len(df_VerbPassSAL)):
#get the next row values (.loc[n+1] next value .loc[n-1] prev value)
    df_VerbPassSAL.loc[i, 'BDO_periodo_prec'] = df_VerbPassSAL.loc[i-1, 'BDO_periodo']
    df_VerbPassSAL.loc[i, 'Numero BDO_prec'] = df_VerbPassSAL.loc[i-1, 'Numero BDO']
#clean unwanted values from list
#create mask of condition to avoid
mask = (df_VerbPassSAL['Numero BDO_prec'] != df_VerbPassSAL['Numero BDO'])|(df_VerbPassSAL['BDO_periodo'] == df_VerbPassSAL['BDO_periodo_prec'])
#Clean unwanted value based on mask conditions
df_VerbPassSAL.loc[mask,['Numero BDO_prec','BDO_periodo_prec']]=""

#wrk_dir_inp=dir_dict["inpDirProd"]
#Get BEF 
df_BEF_Map["PeriodoBEF"]="20"+df_BEF_Map["BEF"].str[0:2]+"-"+df_BEF_Map["BEF"].str[2:4]
df_BEF_Forn_grp=df_BEF_Map.groupby(["Numero BDO","BEF","PeriodoBEF","Forn_cod","Subf_cod","RTI_cod"]).agg({"TotBEF":"sum","Numero Fattura":"max","Data Fattura":"max","Data Pagamento":"max"}).reset_index()
df_BEF_Forn_grp["BEF_FornSubf"]=df_BEF_Forn_grp["BEF"]+"_"+df_BEF_Forn_grp["Forn_cod"]+"_"+df_BEF_Forn_grp["Subf_cod"]
df_BEF_Forn_grp=pd.merge(df_BEF_Forn_grp,df_forn_byCod[["Forn_nome","Forn_cod"]],left_on="Subf_cod", right_on="Forn_cod",how="left",suffixes=('', '_fsub'))
df_BEF_Forn_grp=df_BEF_Forn_grp.rename(columns={"Forn_nome":"Subf_nome","Numero BDO":"BDO","PeriodoBEF":"Periodo"})
df_BEF_Forn_grp=pd.merge(df_BEF_Forn_grp,df_forn_byCod[["Forn_nome","Forn_cod"]],left_on="Forn_cod", right_on="Forn_cod",how="left",suffixes=('', '_forn'))
#df_BEF_Forn_grp=df_BEF_Forn_grp.rename(columns={"Forn_nome":"Subf_nome","Forn_nome_forn":"Forn_nome","Numero BDO":"BDO"})
df_BEF_Forn_grp=pd.merge(df_BEF_Forn_grp,df_BDO[["ROI","BDO","ROI_dip","StatoBDO","Contratto Gara","DEC","RUP","Fornitore RTI","Numero Contratto Gara"]],left_on="BDO", right_on="BDO",how="left",suffixes=('', '_fBDO'))


df_BEF_Quiet_grp=df_quiet.groupby(["BDO","BEF","Periodo_SAL","Forn_cod","Subf_cod"]).agg({"Tot_Quiet":"sum","Stato_Quiet":"max","Numero Fattura":"max","Data Fattura":"max","Data Liquidazione":"max"}).reset_index()
df_BEF_Quiet_grp["BEF_FornSubf"]=df_BEF_Quiet_grp["BEF"]+"_"+df_BEF_Quiet_grp["Forn_cod"]+"_"+df_BEF_Quiet_grp["Subf_cod"]

df_BEF_Pagab=dfCrQuiet_grp.groupby(["BEF"]).agg({"Pagata":"max","Pagabile":"max"}).reset_index()

df_BEF_Forn_grp=pd.merge(df_BEF_Forn_grp,df_BEF_Quiet_grp[["BEF_FornSubf","Tot_Quiet","Stato_Quiet"]],left_on="BEF_FornSubf", right_on="BEF_FornSubf",how="left",suffixes=('', '_BEF'))
df_BEF_Forn_grp=pd.merge(df_BEF_Forn_grp,df_BEF_Pagab[["BEF","Pagata","Pagabile"]],left_on="BEF", right_on="BEF",how="left",suffixes=('', '_BEF'))


df_BEF_Forn_grp["Quiet_Atteso"]= np.where((df_BEF_Forn_grp["Forn_cod_fsub"]>""),df_BEF_Forn_grp["TotBEF"]*0.8,0 ) 
#Aggancia il verbale di SAL del periodo
df_BEF_Forn_grp["BDO_Periodo"]=df_BEF_Forn_grp["BDO"]+"_"+df_BEF_Forn_grp["Periodo"]
df_BEF_Forn_grp=pd.merge(df_BEF_Forn_grp,df_VerbPassSAL[["Stat_VerbSAL","BDO_periodo",]],left_on="BDO_Periodo", right_on="BDO_periodo",how="left",suffixes=('', '_effettivo'))
df_BEF_Forn_grp=df_BEF_Forn_grp.rename(columns={"Stat_VerbSAL":"Stat_VerbSAL_corr"})

#Aggancia il verbale di SAL del periodo successivo
df_BEF_Forn_grp["BDO_Periodo_prec"]=df_BEF_Forn_grp["BDO"]+"_"+df_BEF_Forn_grp["Periodo"]
df_BEF_Forn_grp=pd.merge(df_BEF_Forn_grp,df_VerbPassSAL[["BDO_periodo_prec","Stat_VerbSAL","BDO_periodo",]],left_on="BDO_Periodo_prec", right_on="BDO_periodo_prec",how="left",suffixes=('', '_succ'))
df_BEF_Forn_grp=pd.merge(df_BEF_Forn_grp,df_VerbPassChius[["Numero BDO","Stat_VerbCh"]],left_on="BDO", right_on="Numero BDO",how="left",suffixes=('', '_BEF'))

mask = (df_BEF_Forn_grp['BDO_periodo_prec'].notnull())&(df_BEF_Forn_grp['BDO_periodo_prec']>"0")
#Clean unwanted value based on mask conditions
df_BEF_Forn_grp.loc[mask,'Periodo_succ']=df_BEF_Forn_grp['BDO_periodo_succ'].str[11:18]



tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_BEF_Forn_grp.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_BEF_Forn_grp.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BEF_Forn_grp.memory_usage(deep=True)) )
pass

#Group by PPA 

pass