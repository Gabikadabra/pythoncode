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
#from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_rilasci import df_Rilasci_task
import os
import sys
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task
#Read dummy tables
#rom read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
from Util_logging import logger
wPythProc="read_BdoOda"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()

from read_Interni import df_RisInt_byCID, df_RisInt_byName
#Read dummy tables
#from read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
#wrk_environm=value = parm_dict.get("parm_environment", "Develop")



df_BDOODA_full=pd.read_excel(wrk_dir_inp+"\\"+"Da BDOODA ME23N.XLSX", dtype={"Doc. acq.":str,"InizVal.":str,"FineVal.":str,"Referente Operativo.": str,"Pos..1":str,"Posiz":str, "Fornitore":str, "Contract Manager":str,"Doc. acq.":str,"Cod. Fornitore Pos.":str,"Codice Subappaltatore":str,"Contr.":str,"Pos.":str,"Codice CUP":str,"Stato di elaborazione doc. acquisti":str,"Stato elaborazione doc. acquisti":str,"CtrMgr_cod":str, "Referente Operativo":str})
df_BDOODA_full=df_BDOODA_full[~df_BDOODA_full["Doc. acq."].isnull()]

"""#begin format columns - reduce size"""
wLi=["Doc. acq.","Pos.","Ril","Fornitore","Referente Operativo","Contract Manager","Cod. Fornitore Pos.","Contr.","RdA","DEC/DL"];wDi={value:"int64" for value in wLi}
wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["Quantità","Prezzo lordo",	"Valore netto","Importo Subappalto"];wDf={value:"float" for value in wLf}
wLc=["Tipo","I","UMO","Rilevante PNRR","CdC","Numero CIG","Subappalto",	"Codice Interno subappalto","Attività","Gr. merci"	,"Divisa Subappalto","Descr.Fornitore","Stato elaborazione doc. acquisti","Definizione 2 gr. merci",	"Tipo Linea",	"Descrizione Tipo Linea",	"Codice Subappaltatore",	"Contratto Gara",	"Descr. Subappaltatore","Tipo Ricezione","Milestone",	"Affidamento diretto",	"Divisa Subappalto","Contratto superiore","Descr.Referente Operativo","Documento chiuso","Stato di elaborazione doc. acquisti","Annullamento Residuo","Descrizione_DEC","Codice CUP","Versione","Pos..1"];wDc={value:"category" for value in wLc}
wLnbr=wLi+wLf
df_BDOODA_full[wLnbr]=df_BDOODA_full[wLnbr].fillna(0)
wList_date=["InizVal."	,"FineVal.",	"Data cr.",	"Data del documento acquisto","Data Fine Milestone",	"Data Effettiva Attivazione"]
#Date format
for d in wList_date:
    df_BDOODA_full[d]=pd.to_datetime(df_BDOODA_full[d],errors='coerce')
#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDs,wDf,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_BDOODA_full= df_BDOODA_full.astype(convert_dict)
"""#end format columns - reduce size"""


# ***Filter*** in testfilter is on and specific selection is in place
if (wrk_environm.lower()=="testfilter") & ("*all" not in listfilter_BDO):
    df_BDOODA_full=df_BDOODA_full.loc[(df_BDOODA_full["Doc. acq."].isin(listfilter_BDO))|(df_BDOODA_full["Contr."].isin(listfilter_ODAG))]

#df_BDOODA_full.loc[df_BDOODA_full["FineVal."] > "2099-12-31 00:00:00", "FineVal."] = "2099-12-31 00:00:00"
#df_BDOODA_full.loc[df_BDOODA_full["InizVal."] == "0000-00-00 00:00:00", "InizVal."] = "1980-01-01 00:00:00"
#df_BDOODA_full.loc[df_BDOODA_full["FineVal."] == "0000-00-00 00:00:00", "FineVal."] = "2099-12-31 00:00:00"
#df_BDOODA_full.loc[df_BDOODA_full["FineVal."] == 0, "FineVal."] = "2099-12-31 00:00:00"


#df_BDOODA_full["InizVal."]=pd.to_datetime(df_BDOODA_full["InizVal."], format='%d.%m.%Y')
#df_BDOODA_full["FineVal."]=pd.to_datetime(df_BDOODA_full["FineVal."], format='%d.%m.%Y')



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

df_forn=df_forn.loc[df_forn["Forn_cod"].notnull()]
df_forn["Forn_cod"]=df_forn["Forn_cod"].astype("int")
#Remove duplicate by CID
df_forn_byCod  =  df_forn.drop_duplicates( subset = ["Forn_cod"],  keep = "first" )
#Remove duplicate by Name
df_forn_byName  = df_forn.drop_duplicates( subset = ["Forn_nome"],  keep = "first" )

#df_BDOODA_full["Referente Operativo"]=df_BDOODA_full["Referente Operativo"].str.zfill(3)
#df_BDOODA_full["Contract Manager"]=df_BDOODA_full["Contract Manager"].str.zfill(3)
#Remove deleted lines
df_BDOODA_full=df_BDOODA_full.loc[df_BDOODA_full["I"] != "L"]

#df_BDOODA_full["BDOODALin"]=df_BDOODA_full["Doc. acq."].multiply(1000)+df_BDOODA_full["Pos."]
df_BDOODA_full.rename(columns={df_BDOODA_full.columns[21]: "StatoDocAcq_cod"}, inplace=True)
df_BDOODA_full.rename(columns={df_BDOODA_full.columns[60]: "StatoDocAcq"}, inplace=True)
df_BDOODA_full.rename(columns={df_BDOODA_full.columns[57]: "Forn_nome"}, inplace=True)
#df_BDOODA_full["Referente Operativo"]=df_BDOODA_full["Referente Operativo"].str.zfill(3)
df_BDOODA_full=pd.merge(df_BDOODA_full,df_RisInt_byCID[["Int_Dip","CID","CdC"]],left_on="Referente Operativo", right_on="CID",how="left",suffixes=('', '_int'))
#df_BDOODA_full=pd.merge(df_BDOODA_full,df_forn_byCod,left_on="Fornitore", right_on="Forn_cod",how="left",suffixes=('', '_forn'))
wColRename={"Descr.Referente Operativo":"ROI","Doc. acq.":"DocAcq","Cod. Fornitore Pos..1":"Forn_nome", "Pos.":"Linea","Numero CIG":"CIG","Data cr.":"DataCreaz","Descr.Fornitore":"Forn_RTI","InizVal.":"DataIniz","Testo breve":"Lin_descrizione","FineVal.":"DataFine","Prestazione":"Prestaz_nome", "Tipo":"TipoDocAcq","MLS":"MLS_cod", "Importo Subappalto":"Subf_val","Valore netto":"Forn_val","Quantità":"Forn_qta","Testo testata BDO":"Descrizione DocAcq", "Pos..1":"Macrocl_nbr","Tipo":"TipoDocAcq","Pos.":"Linea","Tipo Linea":"TipLin_cod","Descrizione Tipo Linea":"TipLin_descr","Codice Subappaltatore":"Subf_cod","Milestone":"MLS","Contr.":"ODAGnbr","Attività":"Prestaz_cod","Contratto superiore":"Macrocl_txt","Codice Subappaltatore.1":"Subf_nome","Codice CUP":"CUP","Data Fine Milestone":"Data_MLS","Data Effettiva Attivazione":"StartDate","Int_Dip":"ROI_dip","CID":"ROI_cod","CdC_int":"ROI_CdC","Rilevante PNRR":"PNRR","RdA":"RDI","DEC/DL":"DEC_cod","Descrizione_DEC":"DEC_sap","Contratto Gara":"ODAG"}
df_BDOODA_full=df_BDOODA_full.rename(columns=wColRename)
df_BDOODA_full[["Forn_nome","Subf_nome"]] = df_BDOODA_full[["Forn_nome","Subf_nome"]].replace(np.nan, "")
df_BDOODA_full=df_BDOODA_full.replace(np.nan,0)



wSelODABDO=["DocAcq","Linea","CIG","ODAG","DataCreaz","Forn_RTI","DataIniz","Forn_nome","Lin_descrizione","DataFine","TipLin_cod","TipLin_descr","Prestaz_cod","Prestaz_nome", "TipoDocAcq","StatoDocAcq","StatoDocAcq_cod","Subf_cod","MLS", "Subf_val","Subf_nome","TipLin_cod","Forn_val","Forn_qta","Descrizione DocAcq",  "ODAGnbr","MLS","RDI","Macrocl_nbr","CUP","Data_MLS","TipLin_descr","ROI","ROI_cod","ROI_dip","StartDate","ROI_CdC","Macrocl_txt"] 

df_BDOODA_full=df_BDOODA_full[(df_BDOODA_full["Forn_val"].notnull())]
###   ....

df_BDOODALin=df_BDOODA_full[wSelODABDO]
#df_BDOLin=pd.merge(df_BDOLin,df_ODAGMacr[["ODAGMacr_txt","ODAGMacr_nbr"]],left_on="ODAGMacr_txt", right_on="ODAGMacr_txt",how="left",suffixes=('', '_verb'))
#df_BDOODALin["BDOLin"]=df_BDOODALin["Doc. acq."].multiply(1000)+df_BDOODALin["Posiz"] 

wList_date=["DataIniz"	,"DataFine",	"DataCreaz","StartDate","Data_MLS"]
#Date format
for d in wList_date:
    df_BDOODALin[d]=pd.to_datetime(df_BDOODALin[d],errors='coerce')

#df_BDOODALin.loc[df_BDOODALin["DataFine"] == 0, "Datafine"] = "2099-12-31 00:00:00"
#df_BDOODALin[df_BDOODALin["DataIniz"] == 0, "DataIniz"] = "1980-01-01 00:00:00"

df_BDOODA=df_BDOODALin.groupby(["DocAcq","CIG","ODAGnbr","DataCreaz","Forn_RTI","TipoDocAcq","StatoDocAcq","StatoDocAcq_cod","Descrizione DocAcq",  "ODAG","RDI","Macrocl_nbr","ROI","ROI_cod","ROI_dip","ROI_CdC"], observed=True).agg(Fornitori= ("Forn_nome","unique"),
Subfornitori=("Subf_nome","unique"),DataIniz=("DataIniz","min"),StartDate=("StartDate","min"),DataFine=("DataFine","max"),Forn_val=("Forn_val","sum"),Subf_val=("Subf_val","sum")
     ).reset_index()

df_BDOODA['Fornitori']=[','.join(map(str, l)) for l in df_BDOODA['Fornitori']]
df_BDOODA['Subfornitori']=[','.join(map(str, l)) for l in df_BDOODA['Subfornitori']]


#df_BDOODA['Subfornitori']=[','.join(map(str, l)) for l in df_BDOODA['Subfornitori']]



tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_BDOODALin.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_BDOODALin.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDOODALin.memory_usage(deep=True)) )
pass
#,Forn_list=("Forn_nome",list)

