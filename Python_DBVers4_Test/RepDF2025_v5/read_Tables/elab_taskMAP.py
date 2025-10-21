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
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task,listfilter_Incarico

from Util_logging import logger
wPythProc="elab_taskMAP"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

#from read_Interni import df_RisInt_byCID,df_RisInt_byName
import os
import sys
from read_map import  df_MAP_bytask, df_tot_Accrual_bytask # df_MAP_noCons, df_MAP_zero, df_PDC, df_MAP_byMAP,
from read_prog import df_prog_bytask # df_prog_byTipoRis, df_prog_int, df_prog_altro, df_prog_ammort, df_Rilasci_task
from read_BefQuiet import df_FattPass_IVA,df_BEF_Map
from elab_Bdo_Periodo import df_BDOLin,df_BDO
from read_rilasci import df_Rilasci_task
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDirProd"]
wYear=parm_dict["parm_risint_year"]

#Group by task BDO Lin
df_BDOLin_bytask=df_MAP_bytask.groupby(["DocAcq","Linea","Anno competenza","Mese competenza","Task","TipoDocAcq"],observed=True).agg({"ValoreMAP":"sum"}).reset_index()
df_BDOLin_bytask=pd.merge(df_BDOLin_bytask,df_prog_bytask,left_on="Task", right_on="Codice Task",how="left",suffixes=('', '_Dip'))
df_BDOLin_bytask=pd.merge(df_BDOLin_bytask,df_BDOLin[["CUP","DocAcq","Linea","Data_MLS","Subf_nome","Forn_nome","Forn_RTI","TipLin_descr","ODAG","ODAGnbr","Macrocl_nbr","ROI_dip","Val_resid","ROI_CdC","MLS","Descrizione DocAcq","ROI","Forn_val","Importo Attestato"]],on=["DocAcq","Linea"],how="left",suffixes=('', '_Fatt'))
df_BDOLin_bytask=pd.merge(df_BDOLin_bytask,df_Rilasci_task[["Tipo Incarico", "Cod_Rilascio", "Tit_Rilascio", "Tipo_Rilascio", "LOB", "INCARICO",	"Demand","Numero Task"]],left_on="Task", right_on="Numero Task",how="left",suffixes=('', '_Ril'))
#df_BDOLin_bytask["Val_Resid"]=df_BDOLin_bytask["Forn_val"]-df_BDOLin_bytask["Importo Attestato"]
#Group by task BDO
df_BDO_bytask=df_MAP_bytask.groupby(["DocAcq","Anno competenza","Mese competenza","Task","TipoDocAcq"],observed=True).agg({"ValoreMAP":"sum"}).reset_index()
df_BDO_bytask=pd.merge(df_BDO_bytask,df_prog_bytask,left_on="Task", right_on="Codice Task",how="left",suffixes=('', '_Dip'))
df_BDO_bytask=pd.merge(df_BDO_bytask,df_BDO[["Forn_RTI","ODAG","Macrocl_nbr","Macrocl_txt","DocAcq","ROI_dip","Val_resid","Forn_val","Importo Attestato","ROI_CdC","ROI","Descrizione DocAcq"]],left_on="DocAcq", right_on="DocAcq",how="left",suffixes=('', '_Fatt'))
df_BDO_bytask=pd.merge(df_BDO_bytask,df_Rilasci_task[["Tipo Incarico", "Cod_Rilascio", "Tit_Rilascio", "Tipo_Rilascio", "LOB", "INCARICO",	"Demand","Numero Task"]],left_on="Task", right_on="Numero Task",how="left",suffixes=('', '_Ril'))
#remove all columns contains "_Ril"
df_BDO_bytask=df_BDO_bytask.loc[:,~df_BDO_bytask.columns.str.contains('_Ril', case=False)]
#df_BDO["Val_Resid"]=df_BDO["Valore Ordine"]-df_BDO["Importo Attestato"]
#Group by MAP task BDO Lin
df_MAP_bytask_fatt=pd.merge(df_MAP_bytask,df_prog_bytask,left_on="Task", right_on="Codice Task",how="left",suffixes=('', '_Dip'))

wkeepColms=["MAP","Num.Fatt.Fornitore","Data doc. Fattura","Protocollo IVA","Data di pagamento"]
df_FattPass_IVA_r=df_FattPass_IVA[wkeepColms]

#wkeepColms=["MAP", "BEF", "Data Pagamento","Numero Fattura", "Data Fattura"]
#df_BEF_Map_r=df_BEF_Map[wkeepColms]

df_FattPass_IVA_r.set_index("MAP",inplace=True)
#df_BEF_Map_r.set_index("MAP",inplace=True)
df_MAP_bytask_fatt.set_index("MAP",inplace=True)

df_MAP_bytask_fatt=df_MAP_bytask_fatt.join(df_FattPass_IVA_r,how="left")
#df_MAP_bytask_fatt=df_MAP_bytask_fatt.join(df_BEF_Map_r,how="left")

df_MAP_bytask_fatt=df_MAP_bytask_fatt.rename_axis("MAP").reset_index()

#df_MAP_bytask_fatt=pd.merge(df_MAP_bytask_fatt,df_FattPass_IVA[["MAP","Num.Fatt.Fornitore","Data doc. Fattura","Protocollo IVA","Data di pagamento"]],left_on="MAP", right_on="MAP",how="left",suffixes=('', '_Fatt'))
#df_MAP_bytask_fatt=pd.merge(df_MAP_bytask_fatt,df_BEF_Map[["MAP", "BEF", "Data Pagamento","Numero Fattura", "Data Fattura"]],left_on="MAP", right_on="MAP",how="left",suffixes=('', '_Ril'))
wkeepColms=["CUP","DocAcq","Linea","Data_MLS","Subf_nome","Forn_nome","Forn_RTI","TipLin_descr","ODAG","Macrocl_nbr","Macrocl_txt"]
df_BDOLin_r=df_BDOLin[wkeepColms]


#df_BDOLin_r.set_index("BDOLin",inplace=True)
#df_MAP_bytask_fatt.set_index("BDOLin",inplace=True)

#df_MAP_bytask_fatt=df_MAP_bytask_fatt.join(df_BDOLin_r,how="left", lsuffix='_df1', rsuffix='_df2')
#df_MAP_bytask_fatt=df_MAP_bytask_fatt.rename_axis("BDOLin").reset_index()

df_MAP_bytask_fatt=pd.merge(df_MAP_bytask_fatt,df_BDOLin[["CUP","DocAcq","Linea","Data_MLS","Subf_nome","Forn_nome","Forn_RTI","TipLin_descr","ODAG","Macrocl_txt","ODAGnbr","Macrocl_nbr"]],on=["DocAcq","Linea"],how="left",suffixes=('', '_Fatt'))

df_MAP_bytask_fatt=pd.merge(df_MAP_bytask_fatt,df_Rilasci_task[["Tipo Incarico", "Cod_Rilascio", "Tit_Rilascio", "Tipo_Rilascio", "LOB", "INCARICO",	"Demand","Numero Task"]],left_on="Task", right_on="Numero Task",how="left",suffixes=('', '_Ril'))



df_MAP_bytask_fatt["Data di pagamento"]=pd.to_datetime(df_MAP_bytask_fatt["Data di pagamento"], format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d/%m/%Y')
df_MAP_bytask_fatt["Data doc. Fattura"]=pd.to_datetime(df_MAP_bytask_fatt["Data doc. Fattura"], format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d/%m/%Y')

#Select Accrual for specifica date
df_tot_Accrual_byYr=df_tot_Accrual_bytask.loc[df_tot_Accrual_bytask["Anno competenza"]==wYear]
df_tot_Accrual_byYr=df_tot_Accrual_byYr.rename(columns={"Task":"Codice Task","ValoreMAP":"RateiMAP"})

df_prog_bytask=pd.merge(df_prog_bytask,df_tot_Accrual_byYr,left_on="Codice Task", right_on="Codice Task",how="left",suffixes=('', '_Dip'))

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_prog_bytask.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_prog_bytask.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_prog_bytask.memory_usage(deep=True)) )


pass