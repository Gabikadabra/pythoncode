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
#from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_rilasci import df_Rilasci_task
import os
import sys
import logging
import logging.config
from Util_leggeParm import parm_dict
from read_Interni import df_RisInt_byCID, df_RisInt_byName
#Read dummy tables
from read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
wrk_environm=value = parm_dict.get("parm_environment", "Develop")
logging.config.fileConfig('logging2.conf')
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
wrk_environm=value = parm_dict.get("parm_environment", "Develop")

# create logger
logger = logging.getLogger('simpleExample')

logger.info('start reading supplier list' )

df_BDOODA_full=pd.read_excel(wrk_dir_inp+"\\"+"Da BDOODA ME23N.XLSX", dtype={"Doc. acq.":str,"InizVal.":str,"FineVal.":str,"Referente Operativo.": str,"Pos..1":str,"Posiz":str, "Fornitore":str, "Contract Manager":str,"Doc. acq.":str,"Cod. Fornitore Pos.":str,"Codice Subappaltatore":str,"Contr.":str,"Pos.":str,"Codice CUP":str,"Stato di elaborazione doc. acquisti":str,"Stato elaborazione doc. acquisti":str,"CtrMgr_cod":str, "Referente Operativo":str})
df_BDOODA_full=df_BDOODA_full[~df_BDOODA_full["Doc. acq."].isnull()]

m1 = (df_BDOODA_full["Doc. acq."]!="2024331656")|(df_BDOODA_full["Pos."]!="1")  
m2= (df_BDOODA_full["Doc. acq."]!="2024331653")|((df_BDOODA_full["Pos."]!="4")&(df_BDOODA_full["Pos."]!="5")&(df_BDOODA_full["Pos."]!="6"))
#m2 = df['A'].map(s2).fillna(-np.inf).lt(5)
df_BDOODA_full=df_BDOODA_full.loc[m1&m2]
#df_BDOODA_full = df_BDOODA_full.mask(m1 & m2)

#df_BDOODA_full=df_BDOODA_full[(df_BDOODA_full["Doc. acq."]!="2024331656")|(df_BDOODA_full["Pos."]!="1")]

df_BDOODA_full.loc[df_BDOODA_full["FineVal."] > "2099-12-31 00:00:00", "FineVal."] = "2099-12-31 00:00:00"
df_BDOODA_full.loc[df_BDOODA_full["InizVal."] == "0000-00-00 00:00:00", "InizVal."] = "1980-01-01 00:00:00"
df_BDOODA_full.loc[df_BDOODA_full["FineVal."] == "0000-00-00 00:00:00", "FineVal."] = "2099-12-31 00:00:00"
df_BDOODA_full.loc[df_BDOODA_full["FineVal."] == 0, "FineVal."] = "2099-12-31 00:00:00"

df_BDOODA_full["PNRR"]=np.where((df_BDOODA_full["Rilevante PNRR"]=="SI"),"SI","" )
#df_BDOODA_full["InizVal."]=pd.to_datetime(df_BDOODA_full["InizVal."], format='%d.%m.%Y')
#df_BDOODA_full["FineVal."]=pd.to_datetime(df_BDOODA_full["FineVal."], format='%d.%m.%Y')
df_BDOODA_full["PNRR"] = pd.Categorical(df_BDOODA_full["PNRR"],categories=["","SI"], ordered=True)


df_forn_RTI= df_BDOODA_full[['Fornitore', 'Numero conto fornitore']].drop_duplicates()
df_forn_RTI=df_forn_RTI.rename(columns={"Fornitore":"Forn_cod","Numero conto fornitore":"Forn_nome" })
df_forn_pos= df_BDOODA_full[['Cod. Fornitore Pos.', 'Cod. Fornitore Pos..1']].drop_duplicates()
df_forn_pos=df_forn_pos.rename(columns={"Cod. Fornitore Pos.":"Forn_cod","Cod. Fornitore Pos..1":"Forn_nome" })

df_forn_subf= df_BDOODA_full[['Codice Subappaltatore', 'Codice Subappaltatore.1']].drop_duplicates()
df_forn_subf=df_forn_subf.rename(columns={"Codice Subappaltatore":"Forn_cod","Codice Subappaltatore.1":"Forn_nome" })

df_forn_add=pd.read_excel(wrk_dir_inp+"\\"+"Da TabFornitori_add.xlsx", dtype={"Fornitore":str,	"Nome del fornitore":str})
df_forn_add=df_forn_add.rename(columns={"Fornitore":"Forn_cod","Nome del fornitore":"Forn_nome" })
df_forn_add=df_forn_add[["Forn_cod","Forn_nome"]]
#create list suppliers
df_forn = pd.concat([df_forn_RTI, df_forn_pos, df_forn_subf,df_forn_add], ignore_index=True, sort=False)
#Remove duplicate by CID
df_forn_byCod  =  df_forn.drop_duplicates( subset = ["Forn_cod"],  keep = "first" )
#Remove duplicate by Name
df_forn_byName  = df_forn.drop_duplicates( subset = ["Forn_nome"],  keep = "first" )
#record log
logger.info('end reading supplier list ' + str(df_forn.shape[0]) )

df_BDOODA_full["Referente Operativo"]=df_BDOODA_full["Referente Operativo"].str.zfill(3)
df_BDOODA_full["Contract Manager"]=df_BDOODA_full["Contract Manager"].str.zfill(3)

df_BDOODA_full["BDOODALin"]=df_BDOODA_full["Doc. acq."]+"-"+df_BDOODA_full["Pos."].str.zfill(3) 
#Remove deleted lines
df_BDOODA_full=df_BDOODA_full[df_BDOODA_full["I"] != "L"]
df_BDOODA_full["Referente Operativo"]=df_BDOODA_full["Referente Operativo"].str.zfill(3)
df_BDOODA_full=pd.merge(df_BDOODA_full,df_RisInt_byCID[["Int_Dip","CID","CdC"]],left_on="Referente Operativo", right_on="CID",how="left",suffixes=('', '_int'))
df_BDOODA_full=pd.merge(df_BDOODA_full,df_forn_byCod,left_on="Fornitore", right_on="Forn_cod",how="left",suffixes=('', '_forn'))

df_BDOODA_full=df_BDOODA_full.rename(columns={"Fornitore":"Forn_main","Pos.":"Posiz","Stato elaborazione doc. acquisti":"Stato_cod","Stato di elaborazione doc. acquisti":"Stato","Referente Operativo":"ROI_cod", "Descr.Referente Operativo":"ROI","Milestone":"MLS","Attività":"Prestaz_cod","Tipo Linea":"TipLin_cod","Descrizione Tipo Linea":"TipLin_descr","Data Fine Milestone":"MLS_dtFine","DEC/DL":"DEC","Codice CUP":"CUP","Codice Subappaltatore":"Subf_cod","Pos..1":"Macrocl_nbr","Contract Manager":"CtrMgr_cod","Contr.":"ODAG","RdA":"RDI","Data Effettiva Attivazione":"StartDate","Forn_nome":"Forn_nome","Int_Dip":"ROI_dip","CdC":"ROI_CdC"}) #Rilevante PNRR":"PNRR" 
df_BDOODA_full=df_BDOODA_full.replace(np.nan,0)
wSelODABDO=["Doc. acq.","Posiz","Numero CIG","ODAG","Data cr.","Forn_main","InizVal.","Forn_nome","Testo breve","FineVal.","TipLin_cod","TipLin_descr","Prestaz_cod","Prestazione", "Tipo","Stato","Stato_cod","Subf_cod","MLS", "Importo Subappalto","Subappalto","TipLin_cod","Valore netto","Quantità","Prezzo lordo", "Prestazione",  "Testo testata BDO", "Contratto superiore","CtrMgr_cod","Descr.Contract Manager", "Contratto Gara","Gr. merci","DEC","Descrizione_DEC","Descrizione Milestone","Tipologia contratto",	"Versione","Prezzo lordo","RDI","Contratto superiore","Macrocl_nbr","CUP","MLS_dtFine","TipLin_descr","TipLin_cod","Prestaz_cod","ROI","ROI_cod","ROI_dip","StartDate","ROI_CdC","PNRR"] 

logger.info('start reading ODA details' )
df_BDOODA_full=df_BDOODA_full[(df_BDOODA_full["Valore netto"].notnull())]




df_BDOODALin=df_BDOODA_full[wSelODABDO]
#df_BDOLin=pd.merge(df_BDOLin,df_ODAGMacr[["ODAGMacr_txt","ODAGMacr_nbr"]],left_on="ODAGMacr_txt", right_on="ODAGMacr_txt",how="left",suffixes=('', '_verb'))
df_BDOODALin["BDOLin"]=df_BDOODALin["Doc. acq."]+"-"+df_BDOODALin["Posiz"].str.zfill(3) 



#,Forn_list=("Forn_nome",list)
logger.info('end reading ODA details ' + str(df_BDOODALin.shape[0]) )
