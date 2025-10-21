import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from elab_Bdo_Periodo import df_BDO_recent_per
from elab_Oda_Periodo import df_ODA_recent_per
from read_Bdo_Oda import df_BDOODA

from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task
#Read dummy tables
#rom read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
from Util_logging import logger
from read_verbPass import wDict_Verb
wPythProc="elab_Bdo_pivot"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()

df_BDO_recent_per["Periodo"]=df_BDO_recent_per["Anno competenza"].astype(str)+"-"+df_BDO_recent_per["Mese competenza"].astype(str).str.zfill(2)
#df_BDO_recent_per["Stat_VerbSAL"]=df_BDO_recent_per["ID_VerbSAL"].replace(wDict_Verb)

df_BDO_recent_per_daCons=df_BDO_recent_per.loc[df_BDO_recent_per["Flag_daCons"]==True]
df_BDO_recent_per_daCons["Progr"]=1
df_BDOnoCons_byDip=df_BDO_recent_per_daCons.pivot_table(index=["ROI_dip"], values=["DocAcq"] ,columns=["Periodo"],
							   aggfunc="sum")
df_BDOnoCons_byDip.columns = df_BDOnoCons_byDip.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
df_BDOnoCons_byDip=df_BDOnoCons_byDip.reset_index().rename_axis(None,axis=1)

df_BDOnoCons_byBDO=df_BDO_recent_per_daCons.pivot_table(index=["ROI","DocAcq","ID_VerbAt"], values=["Progr"] ,columns=["Periodo"],
							   aggfunc="sum")
df_BDOnoCons_byBDO.columns = df_BDOnoCons_byBDO.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
df_BDOnoCons_byBDO=df_BDOnoCons_byBDO.reset_index().rename_axis(None,axis=1)
df_BDOnoCons_byBDO=pd.merge(df_BDOnoCons_byBDO,df_BDOODA[["Descrizione DocAcq","StatoDocAcq","DocAcq","ROI_dip"]],on="DocAcq",how="left",suffixes=('', '_oda'))
df_BDOnoCons_byBDO["Stat_VerbAt"]=df_BDOnoCons_byBDO["ID_VerbAt"].replace(wDict_Verb)



df_BDOnoVerb_byBDO=df_BDO_recent_per.pivot_table(index=["DocAcq","ID_VerbAt","ID_VerbCh"], values=["MAP_noverb"] ,columns=["Periodo"],
							   aggfunc="sum")
df_BDOnoVerb_byBDO.columns = df_BDOnoVerb_byBDO.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
df_BDOnoVerb_byBDO=df_BDOnoVerb_byBDO.reset_index().rename_axis(None,axis=1)
df_BDOnoVerb_byBDO=pd.merge(df_BDOnoVerb_byBDO,df_BDOODA[["ROI_dip","ODAG","Macrocl_nbr","ROI","Descrizione DocAcq","StatoDocAcq","DataIniz","DocAcq","DataFine","Fornitori","Subfornitori","Forn_RTI"]],on="DocAcq",how="left",suffixes=('', '_oda'))
df_BDOnoVerb_byBDO["Stat_VerbAt"]=df_BDOnoVerb_byBDO["ID_VerbAt"].replace(wDict_Verb)
df_BDOnoVerb_byBDO["Stat_VerbCh"]=df_BDOnoVerb_byBDO["ID_VerbCh"].replace(wDict_Verb)


df_BDOVerb_recent_per=df_BDO_recent_per.pivot_table(index=["DocAcq","ID_VerbAt","ID_VerbCh"], values=["Stat_VerbSAL"] ,columns=["Periodo"],
							   aggfunc="max") #,"MAP_Cons","MAP_noverb"
df_BDOVerb_recent_per.columns = df_BDOVerb_recent_per.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
df_BDOVerb_recent_per=df_BDOVerb_recent_per.reset_index().rename_axis(None,axis=1)

df_BDOVerb_recent_per=pd.merge(df_BDOVerb_recent_per,df_BDOODA[["DocAcq","Descrizione DocAcq","ROI_dip","ODAG","Macrocl_nbr","ROI","StatoDocAcq","DataFine","DataIniz","Subfornitori","Forn_RTI"]],on="DocAcq",how="left",suffixes=('', '_oda'))
df_BDOVerb_recent_per["Stat_VerbAt"]=df_BDOVerb_recent_per["ID_VerbAt"].replace(wDict_Verb)
df_BDOVerb_recent_per["Stat_VerbCh"]=df_BDOVerb_recent_per["ID_VerbCh"].replace(wDict_Verb)


#ODA pivot per periodo
df_ODAnoCons_byODA=df_ODA_recent_per.groupby(["Anno competenza","Mese competenza","DocAcq","ROI_dip","ROI"],observed=True).agg(numbdo=("DocAcq","count")).reset_index()
df_ODAnoCons_byODA["Periodo"]=df_ODAnoCons_byODA["Anno competenza"].astype(str)+"-"+df_ODAnoCons_byODA["Mese competenza"].astype(str).str.zfill(2)
df_ODAnoCons_byROI=df_ODAnoCons_byODA.pivot_table(index=["ROI_dip","ROI"], columns=["Periodo"],values=["numbdo"],
							   aggfunc="sum")
#25/09/2024		df_col_Vpivot_BDO=dfBDO_per_dip.pivot_table(index=['Dip'],  
df_ODAnoCons_byROI.columns =df_ODAnoCons_byROI.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
df_ODAnoCons_byROI=df_ODAnoCons_byROI.reset_index().rename_axis(None,axis=1)

df_ODAnoCons_byDip=df_ODAnoCons_byODA.pivot_table(index=["DocAcq","ROI_dip","ROI"], columns=["Periodo"],values=["numbdo"],
							   aggfunc="sum")
#25/09/2024		df_col_Vpivot_BDO=dfBDO_per_dip.pivot_table(index=['Dip'],  
df_ODAnoCons_byDip.columns =df_ODAnoCons_byDip.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
df_ODAnoCons_byDip=df_ODAnoCons_byDip.reset_index().rename_axis(None,axis=1)
df_ODAnoCons_byDip=pd.merge(df_ODAnoCons_byDip,df_BDOODA[["DocAcq","Descrizione DocAcq"]],on="DocAcq",how="left",suffixes=('', '_oda'))


tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_BDOnoCons_byDip.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_BDOnoCons_byDip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDOnoCons_byDip.memory_usage(deep=True)) )
