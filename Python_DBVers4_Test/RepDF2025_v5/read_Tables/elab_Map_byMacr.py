#Read Map Dataframe
from read_map import df_MAP_byMAP
from read_BdoLin import df_BDOLin
import pandas as pd
from elab_ODAG_RUPDEC import df_ODAGMacr,df_ODAG
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task

import numpy as np
import sys
wDt_now = pd.to_datetime('now') # Restituisce la data di oggi
#"Documento Acquisto"
from Util_logging import logger
#wPythProc="read_verbPass"
import timeit  #funzione timeit per valutare durata estrazione
#tempoInz=timeit.default_timer()
wPythProc="elab_Map_byMacr"

tempoInz=timeit.default_timer()	

df_MAP_byMAP =pd.merge(df_MAP_byMAP,df_BDOLin[["DocAcq","Linea","ODAGnbr","Macrocl_txt","Macrocl_nbr", "TipLin_Descr" ]], on=["DocAcq","Linea"],how="left",suffixes=("","_bdo")) #|"BDOLin","ODAGMacr_nbr","Numero Contratto Gara","Contratto Macroclasse","Tipologia Fornitura"

#"DocAcq","ODAGnbr","Macrocl_txt","RDI","StatoDocAcq","ROI","DataCreaz","StartDate","DocAcq_FlagCh","Macrocl_nbr","DEC","RUP","ROI_dip","ROI_CID","ROI_CdC"

# Unisce le informazioni sui MAP
df_MAP_byMAP["mesi_daMAP"]= 12*(wDt_now.year - df_MAP_byMAP["Anno competenza"].map(int))+wDt_now.month- df_MAP_byMAP["Mese competenza"].map(int)


#Seleziona solo i MAP con macroclasse valorizzata
df_MAP_macr= df_MAP_byMAP[df_MAP_byMAP["ODAGnbr"]>2000]
#df_MAP_macr["BDOLin"]=df_MAP_macr["Documento Acquisto"]+"-"+df_MAP_macr["Posizione"].str.zfill(3)

df_MAP_macr["Periodo"]=df_MAP_macr["Anno competenza"].astype(str)+"-"+df_MAP_macr["Mese competenza"].astype(str).str.zfill(2)
#df_MAP_macr["Periodo"]=df_MAP_macr["Anno competenza"]+"-"+df_MAP_macr["Mese competenza"].str.zfill(2) 

dfMAP_Mg=df_MAP_macr.groupby(["ODAGnbr","Macrocl_nbr","TipLin_Descr"]).agg({"MAP_Cons":"sum"}).reset_index()
dfMAP_Mg_R=dfMAP_Mg
#sys.exit()
dfMAP_Mper=df_MAP_macr.loc[df_MAP_macr["mesi_daMAP"]<13]
dfMAP_Mg=dfMAP_Mper.groupby(["ODAGnbr","Macrocl_nbr","TipLin_Descr"]).agg({"MAP_Cons":"sum","Periodo":"nunique"}).reset_index()
dfMAP_Mg["AttMens_12m"]=dfMAP_Mg["MAP_Cons"]/12 #dfMAP_Mg["Periodo"] #/12
dfMAP_Mg_12m=dfMAP_Mg

dfMAP_Mper=df_MAP_macr.loc[df_MAP_macr["mesi_daMAP"]<7]
dfMAP_Mg=dfMAP_Mper.groupby(["ODAGnbr","Macrocl_nbr","TipLin_Descr"]).agg({"MAP_Cons":"sum","Periodo":"nunique"}).reset_index()
dfMAP_Mg["AttMens_06m"]=dfMAP_Mg["MAP_Cons"]/6 #dfMAP_Mg["Periodo"]
dfMAP_Mg_06m=dfMAP_Mg


dfMAP_Mper=df_MAP_macr.loc[df_MAP_macr["mesi_daMAP"]<4]
dfMAP_Mg=dfMAP_Mper.groupby(["ODAGnbr","Macrocl_nbr","TipLin_Descr"]).agg({"MAP_Cons":"sum","Periodo":"nunique"}).reset_index()
#dfMAP_Mg.columns=dfMAP_Mg.columns.droplevel(0)
dfMAP_Mg["AttMens_03m"]=dfMAP_Mg["MAP_Cons"]/3 #dfMAP_Mg["Periodo"]
dfMAP_Mg_03m=dfMAP_Mg
#Build a selection for rdtimate by macro and supplier types
dfMAP_Mg_R=pd.merge(dfMAP_Mg_R,dfMAP_Mg_03m[["ODAGnbr","Macrocl_nbr","TipLin_Descr","AttMens_03m"]], on=["ODAGnbr","Macrocl_nbr","TipLin_Descr"],how="left",suffixes=("","_y"))
dfMAP_Mg_R=pd.merge(dfMAP_Mg_R,dfMAP_Mg_06m[["ODAGnbr","Macrocl_nbr","TipLin_Descr","AttMens_06m"]], on=["ODAGnbr","Macrocl_nbr","TipLin_Descr"],how="left",suffixes=("","_y"))
dfMAP_Mg_R=pd.merge(dfMAP_Mg_R,dfMAP_Mg_12m[["ODAGnbr","Macrocl_nbr","TipLin_Descr","AttMens_12m"]], on=["ODAGnbr","Macrocl_nbr","TipLin_Descr"],how="left",suffixes=("","_y"))
#dfMAP_Mg_R=pd.merge(dfMAP_Mg_R,df_ODAGMacr[["ODAGnbr","Macrocl_nbr"]], on=["ODAGMacr_nbr"],how="left",suffixes=("","_y"))


dfMAP_M_Tot=dfMAP_Mg_R.groupby(["ODAGnbr","Macrocl_nbr"]).agg({"AttMens_03m":"sum","AttMens_06m":"sum","AttMens_12m":"sum"}).reset_index()
#dfMAP_M_Tot.columns=dfMAP_M_Tot.columns.droplevel(0)
dfMAP_M_Tot["StimaMensile"]=dfMAP_M_Tot[["AttMens_03m","AttMens_06m","AttMens_12m"]].max(axis=1)
dfMAP_M_Tot["Stima_03m"]=dfMAP_M_Tot["StimaMensile"]*3
dfMAP_M_Tot["Stima_06m"]=dfMAP_M_Tot["StimaMensile"]*6
dfMAP_M_Tot["Stima_12m"]=dfMAP_M_Tot["StimaMensile"]*12

#df_MAP_macr_mth= df_MAP_macr.groupby(["ODAGMacr_nbr","Numero Contratto Gara","Tipologia Fornitura"]).agg({"ValoreMAP":"sum"}).rename(columns={"ValoreMap":"MAP_cons"}).reset_index()


#df_MAP_macr= df_MAP_macr.groupby(["Documento Acquisto", "Posizione", "MAP", "Cod.Accett.MAP", "Mese competenza", "Anno competenza","Tipo ordine acquisto"]).agg(
#MAP_Cons= ("MAP_Cons","sum"),Qta= ("Qta","sum")).reset_index()
#Associa il valore delle stime a livello di macroclasse
df_ODAGMacr=pd.merge(df_ODAGMacr,dfMAP_M_Tot, on=["ODAGnbr","Macrocl_nbr"],how="left",suffixes=("","_y"))
#df_ODAGMacr.columns=df_ODAGMacr.columns.droplevel(0)
#calcola il residuo a livello di macroclasse con consumi proiettati a 6 mesi e 12 mesi 
df_ODAGMacr["StimaRes_03m"]=df_ODAGMacr["Libero su macroclasse"]-df_ODAGMacr["Stima_03m"]
df_ODAGMacr["StimaRes_06m"]=df_ODAGMacr["Libero su macroclasse"]-df_ODAGMacr["Stima_06m"]
df_ODAGMacr["StimaRes_12m"]=df_ODAGMacr["Libero su macroclasse"]-df_ODAGMacr["Stima_12m"]
#Select on multiple conditions
df_ODAGMacr["CheckRes"]=np.select([df_ODAGMacr["Libero su macroclasse"]<0,df_ODAGMacr["StimaRes_03m"]<0,df_ODAGMacr["StimaRes_06m"]<0,df_ODAGMacr["StimaRes_12m"]<0],["**macroclasse neg.","*res.macrocl <3 mesi","*res.macrocl <6 mesi","*res.macrocl <12 mesi"],default="")
dfMAP_ODAG=df_ODAGMacr.groupby(["ODAGnbr"]).agg({"Libero su macroclasse":"sum","StimaRes_03m":"min","StimaRes_06m":"min","StimaRes_12m":"min"}).reset_index()
dfMAP_ODAG["CheckRes"]=np.select([dfMAP_ODAG["Libero su macroclasse"]<0,dfMAP_ODAG["StimaRes_03m"]<0,dfMAP_ODAG["StimaRes_06m"]<0,dfMAP_ODAG["StimaRes_12m"]<0],["**macroclasse neg.","*res.macrocl <3 mesi","*res.macrocl <6 mesi","*res.macrocl <12 mesi"],default="")
df_ODAG=pd.merge(df_ODAG,dfMAP_ODAG, left_on=["ODAGnbr"],right_on=["ODAGnbr"],how="left",suffixes=("","_y"))

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(dfMAP_ODAG.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(dfMAP_ODAG.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(dfMAP_ODAG.memory_usage(deep=True)) )
