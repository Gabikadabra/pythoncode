#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from Util_leggeParm import parm_dict
#Read dummy tables
from read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
wrk_environm=value = parm_dict.get("parm_environment", "Develop")
wrk_risint_year = parm_dict.get("parm_risint_year", "2025")


import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
from read_rilasci import df_Rilasci_task
from read_remapDip import w_eccez_Rilasci, w_eccez_Task
import os
import sys

import logging
import logging.config
#2025-04-01 start Sub to clead non numeric chars
import re
def remove_chars(s):
#2025.04.08 begin
#  	return re.sub("[^0-9,]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
	a= re.sub("[^0-9,.]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
	a=a.strip()
	return a.replace(",",".")
#2025.04.08 end
     #2025-04-01 end
logging.config.fileConfig('logging2.conf')

# create logger
logger = logging.getLogger('simpleExample')
#Dictionary for resource types
dict_RscType = {"Fogli Ore":"INT","Prestazioni Esterne":"EST","Altri costi":"EST","Prestazioni Call Center":"ALTRO","Servizi Data Center":"ALTRO","Infrastruttura":"ALTRO","Manutenzione":"ALTRO","C/Vendita":"ALTRO","Ammortamenti":"ALTRO"}


##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]

df_prog=pd.read_excel(wrk_dir_inp+"\\"+"Da ZConsuntivi.xlsx",dtype={"Codice Risorsa": str, "CUP": str})
#2025-04-01 start - Remove all non (numeric chars and commas)
df_prog["QTA Prev. Corr."]=df_prog["QTA Prev. Corr."].apply(remove_chars)
#2025.04.08 begin
#df_prog["QTA Prev. Corr."]=df_prog["QTA Prev. Corr."].str.replace(',', '.')
#df_prog["QTA Prev. Corr."]=df_prog["QTA Prev. Corr."].apply(remove_chars)
df_prog["QTA Prev. Corr."]=df_prog["QTA Prev. Corr."].apply(pd.to_numeric,errors="coerce").fillna(0)
#2025.04.08 end
df_prog["QTA Cons. di Periodo"]=df_prog["QTA Cons. di Periodo"].apply(remove_chars)
df_prog["QTA Cons. di Periodo"]=df_prog["QTA Cons. di Periodo"].apply(pd.to_numeric,errors="coerce").fillna(0)
df_prog["QTA YTD"]=df_prog["QTA YTD"].apply(remove_chars)
#2025-04-01 end

#df_prog["QTA Prev. Corr."]=pd.to_numeric(df_prog["QTA Prev. Corr."],errors="coerce")
#df_prog["QTA Prev. Corr."]=df_prog["QTA Prev. Corr."].replace(np.nan, 0)
#df_prog["QTA Cons. di Periodo"]=pd.to_numeric(df_prog["QTA Cons. di Periodo"],errors="coerce")
#df_prog["QTA Cons. di Periodo"]=df_prog["QTA Cons. di Periodo"].replace(np.nan, 0)
df_prog["QTA YTD"]=pd.to_numeric(df_prog["QTA YTD"],errors="coerce")
df_prog["QTA YTD"]=df_prog["QTA YTD"].replace(np.nan, 0)

#Extract previous ZConsuntivi 
wpath_legacy = wrk_dir_inp+"\\ImportLegacy"


logger.info('start reading tasks')
df_prog_legacy=pd.read_excel(wpath_legacy+"\\"+"Da_ZConsuntivi_Legacy.xlsx",dtype={"Codice Risorsa": str, "CUP": str})
##filenames = [file for file in os.listdir(wpath_legacy) if file.startswith('ZCons')]
##df_prog_legacy = pd.concat([pd.read_excel(wpath_legacy +"\\"+ file) for file in filenames],dtype={"Codice Risorsa": str, "CUP": str}, ignore_index=True)
#Remove task already exist in actual file

df_prog_tasklist=df_prog["Codice Task"].drop_duplicates()

df_prog_legacy=pd.merge(df_prog_legacy,df_prog_tasklist,left_on="Codice Task", right_on="Codice Task",how="outer",indicator=True)
df_prog_legacy=df_prog_legacy[df_prog_legacy['_merge']=='left_only']

#df_prog_legacy["Data I. Corrente"]=pd.to_datetime(df_prog_legacy["Data I. Corrente"], format='%Y-%m-%d',errors="coerce").dt.strftime("%d/%m/%Y")
#df_prog_legacy["Data F. Corrente"]=pd.to_datetime(df_prog_legacy["Data F. Corrente"], format='%Y-%m-%d',errors="coerce").dt.strftime("%d/%m/%Y")
#df_prog_legacy["Data I. Contrattuale"]=pd.to_datetime(df_prog_legacy["Data I. Contrattuale"], format='%Y-%m-%d',errors="coerce").dt.strftime("%d/%m/%Y")
#df_prog_legacy["Data F. Contrattuale"]=pd.to_datetime(df_prog_legacy["Data F. Contrattuale"], format='%Y-%m-%d',errors="coerce").dt.strftime("%d/%m/%Y")


#=pd.merge(df_prog_legacy,df_prog,left_on="Codice Task", right_on="Codice Task",how="right",suffixes=('', '_pos'))
#



#Add tasks not included in actual file
df_prog = pd.concat([df_prog, df_prog_legacy], ignore_index=True, sort=False)

#Iniziatile default start date and end date
w_dft_startdate="01/01/1980";w_dft_enddate="31/12/2099"
#pd.to_datetime(w_dft_startdate, format='%Y-%m-%d') Assign the default value for date
df_prog["Data I. Corrente"].fillna(pd.to_datetime(w_dft_startdate, format='%d/%m/%Y'), inplace = True)
df_prog["Data F. Corrente"].fillna(pd.to_datetime(w_dft_enddate, format='%d/%m/%Y'), inplace = True)
df_prog["Data I. Contrattuale"].fillna(pd.to_datetime(w_dft_startdate, format='%d/%m/%Y'), inplace = True)
df_prog["Data F. Contrattuale"].fillna(pd.to_datetime(w_dft_enddate, format='%d/%m/%Y'), inplace = True)
#df_prog["Data I. Corrente"].fillna("1980-01-01 00:00:00.000", inplace = True)
#df_prog["Data F. Corrente"].fillna("2099-12-31 00:00:00.000", inplace = True)
df_prog.loc[df_prog["Data F. Corrente"] == "00/00/0000", "Data F. Corrente"] = pd.to_datetime(w_dft_enddate, format='%d/%m/%Y')
df_prog.loc[df_prog["Data I. Corrente"] == "00/00/0000", "Data I. Corrente"] = pd.to_datetime(w_dft_startdate, format='%d/%m/%Y')
df_prog.loc[df_prog["Data F. Contrattuale"] == "00/00/0000", "Data F. Contrattuale"] = pd.to_datetime(w_dft_enddate, format='%d/%m/%Y')
df_prog.loc[df_prog["Data I. Contrattuale"] == "00/00/0000", "Data I. Contrattuale"] = pd.to_datetime(w_dft_startdate, format='%d/%m/%Y')
#Extract "Anno Inizio" and "Anno Fine"
df_prog["Anno Inizio"]=pd.to_datetime(df_prog['Data I. Corrente'],errors='coerce').dt.strftime('%Y')
df_prog["Anno Fine"]=pd.to_datetime(df_prog['Data F. Corrente'],errors='coerce').dt.strftime('%Y')



#Filter if Develop environment
if wrk_environm =="Develop":
	df_prog= df_prog.loc[df_prog["Codice Task"].isin(df_Task_dummy["Task"])]

df_prog= df_prog.loc[(df_prog["Categoria Risorsa"].notna())|(df_prog["Codice Risorsa"].notna())]

#df_prog["CID"]=df_prog["Codice Risorsa"].str.slice(3, 3)
#df_progaA= df_prog.loc(df_prog["Categoria Risorsa"]=="Fogli Ore") 



df_prog["CID"]= np.where((df_prog["Categoria Risorsa"]=="Fogli Ore"),df_prog["Codice Risorsa"].str.slice(3, 6),"" ) 
#df_progaA= df_prog.loc[df_prog["Categoria Risorsa"]=="Fogli Ore"] 
#df_progaA["CID"]= df_prog["Codice Risorsa"].str.slice(3, 3) 


#Assegna il tipo risorsa sulla base della categoria
df_prog["TipoRis"]= df_prog['Categoria Risorsa'].map(dict_RscType)
df_prog=pd.merge(df_prog,df_RisInt_byName[["Cognome Nome","CdC","Int_Dip"]],left_on="Resp. Progetto", right_on="Cognome Nome",how="left",suffixes=('', '_prog'))

df_prog=pd.merge(df_prog,df_RisInt_byName[["Cognome Nome","CdC","Int_Dip"]],left_on="Resp. MP", right_on="Cognome Nome",how="left",suffixes=('', '_MP'))
df_prog=pd.merge(df_prog,df_RisInt_byCID[["CID","Cognome Nome","CdC","Int_Dip"]],left_on="CID", right_on="CID",how="left",suffixes=('', '_int'))
df_prog=df_prog.rename(columns={"CdC_MP":"MP_CdC","Int_Dip_MP":"MP_Dip" ,"CdC_int":"RisInt_CDC","Cognome Nome_int":"RisInt" ,"Resp. Progetto":"Resp. PM","Int_Dip":"PM_Dip","Int_Dip_int":"RisInt_Dip","Linea di business":"LOB" })
wCharColumns=["Cod. Area Progetto","Cod. Progetto","Descrizione Progetto","Tipologia Progetto","Tipo Rilascio","Stato Progetto","Resp. PM","Resp. MP","Cod. Area MP","Codice MP","Progetto CdF","Codice Task","Riferimento WBS","Task Type","Categoria Risorsa","Codice RL","Descrizione RL","Incarico","Data I. Corrente","Data F. Corrente","Data I. Contrattuale","Data F. Contrattuale","Descrizione Task","CUP","CID","TipoRis","CdC","MP_CdC","MP_Dip","PM_Dip"]
wNumColumns=["QTA Prev. Corr.",   "QTA Prev. Contr.",  "Imp. Prev. Corr.", "Imp. Prev. Contr.",  "QTA Cons. di Periodo", "QTA YTD", "Imp. Cons. di Period", "Imp. Cons. YTD"]
df_prog[wCharColumns]=df_prog[wCharColumns].fillna("")
df_prog[wNumColumns]=df_prog[wNumColumns].fillna(0)
df_prog_altro = df_prog[df_prog['TipoRis'] == 'ALTRO']
df_prog_ammort = df_prog[df_prog['Categoria Risorsa'] == 'Ammortamenti']

#### get Accrual for task period
from read_Accrual import df_Accrual_bytask
#remove Ammortamenti
df_prog_noammort = df_prog[df_prog['Categoria Risorsa'] != 'Ammortamenti']

#recalculate task with Ammortamenti
df_prog_ammort_grp=df_prog_ammort.groupby(
["Cod. Area Progetto","Cod. Progetto","Descrizione Progetto","Tipologia Progetto","Tipo Rilascio","Stato Progetto","Resp. PM","Resp. MP","Cod. Area MP","Codice MP","Progetto CdF","Codice Task","Task Type","Incarico","Data I. Corrente","Data F. Corrente","Data I. Contrattuale","Data F. Contrattuale","Anno Inizio","Anno Fine","Descrizione Task","CUP","CID","MP_CdC","MP_Dip","PM_Dip"]).agg(**{
"Imp. Prev. Corr.":("Imp. Prev. Corr.","sum"),
"Imp. Prev. Contr.":("Imp. Prev. Contr.","sum"),
"Imp. Cons. di Period":("Imp. Cons. di Period","sum"),
"ConsuntivoYTD":("Imp. Cons. di Period","sum")  #Attenzione il consuntivo YTD viene posto uguale a quello dello stesso anno
     }).reset_index()


##2 df_prog_ammort_grp=pd.merge(df_prog_ammort_grp,df_Accrual_bytask, how="left", left_on=["Codice Task","Anno Inizio"],right_on=["Codice Task","Anno"] )
##2 df_prog_ammort_grp=df_prog_ammort_grp.rename(columns={"Imp. Cons. di Period":"Consunt.Prec","ValoreMAP":"Imp. Cons. di Period"  })




#re-compose prog dataframe
##2df_prog = pd.concat([df_prog_noammort, df_prog_ammort_grp], ignore_index=True, sort=False)

#converte in formato stringa i campi che possono avere dati "sporchi" 
df_prog['Descrizione RL'] = df_prog['Descrizione RL'].astype(str)
df_prog['Codice RL'] = df_prog['Codice RL'].astype(str)
df_prog['Categoria Risorsa'] = df_prog['Codice RL'].astype(str)

#converte con data fittizia le date non corrette
#datetime = '1980-01-01';df_prog["Data I. Corrente"] = pd.to_datetime(df_prog["Data I. Corrente"], errors='coerce').fillna(pd.Timestamp(datetime)) 
#datetime = '1980-01-01';df_prog["Data I. Contrattuale"] = pd.to_datetime(df_prog["Data I. Contrattuale"], errors='coerce').fillna(pd.Timestamp(datetime)) 
#datetime = '2099-12-31';df_prog["Data F. Corrente"] = pd.to_datetime(df_prog["Data F. Corrente"], errors='coerce').fillna(pd.Timestamp(datetime)) 
#datetime = '2099-12-31';df_prog["Data F. Contrattuale"] = pd.to_datetime(df_prog["Data F. Contrattuale"], errors='coerce').fillna(pd.Timestamp(datetime)) 
#Filtra le righe per le risorse interne
df_prog_int = df_prog[df_prog['TipoRis'] == 'INT']
df_prog_int = df_prog_int[df_prog_int["Anno Fine"]==wrk_risint_year]
df_prog_int=df_prog_int.rename(columns={"QTA Prev. Corr.":"HPrev_Corr","QTA Cons. di Periodo":"HCons"})

df_prog_int_riep = df_prog_int.groupby(["RisInt_Dip","RisInt","Cod. Progetto","Codice Task"]).agg(
HPrev_Corr= ("HPrev_Corr","sum"),
HCons=("HCons","sum")
     ).reset_index()
df_prog_int_riep["OreResidue"]=df_prog_int_riep["HPrev_Corr"]-df_prog_int_riep["HCons"]
df_prog_int_riep["TotHPianRis"] = df_prog_int_riep.groupby('RisInt',as_index = False)['HPrev_Corr'].transform('sum')

df_prog_byResCategory=df_prog.groupby(["Cod. Area Progetto","Cod. Progetto","Descrizione Progetto","Tipologia Progetto","Tipo Rilascio","Stato Progetto","Resp. PM","Resp. MP","Cod. Area MP","Codice MP","Progetto CdF","Codice Task","Riferimento WBS","Task Type","Categoria Risorsa","Codice RL","Descrizione RL","Incarico","Data I. Corrente","Data F. Corrente","Data I. Contrattuale","Data F. Contrattuale","Descrizione Task","CUP","CID","TipoRis","CdC","MP_CdC","MP_Dip","PM_Dip"]).agg(
Pianif_Corr_QTA= ("QTA Prev. Corr.","sum"),
Pianif_Contr_QTA=("QTA Prev. Contr.","sum"),
Pianif_Corr =("Imp. Prev. Corr.","sum"),
Pianif_Contr=("Imp. Prev. Contr.","sum"),
Consuntivo_QTA=("QTA Cons. di Periodo","sum"),
Consuntivo_QTAYTD=("QTA YTD","sum"),
Consuntivo=("Imp. Cons. di Period","sum"),
ConsuntivoYTD=("Imp. Cons. YTD","sum")
     ).reset_index()
#Tot by task
#df_prog=df_prog.fillna(0)
df_prog_byTipoRis=df_prog.groupby(
["Cod. Area Progetto","Cod. Progetto","Descrizione Progetto","Tipologia Progetto","Tipo Rilascio","Stato Progetto","Resp. PM","Resp. MP","Cod. Area MP","Codice MP","Progetto CdF","Codice Task","Task Type","Incarico","Data I. Corrente","Data F. Corrente","Data I. Contrattuale","Data F. Contrattuale","Descrizione Task","CUP","CID","TipoRis","MP_CdC","MP_Dip","PM_Dip","LOB"]).agg(Pianif_Corr_QTA= ("QTA Prev. Corr.","sum"),
Pianif_Contr_QTA=("QTA Prev. Contr.","sum"),
Pianif_Corr =("Imp. Prev. Corr.","sum"),
Pianif_Contr=("Imp. Prev. Contr.","sum"),
Consuntivo_QTA=("QTA Cons. di Periodo","sum"),
Consuntivo_QTAYTD=("QTA YTD","sum"),
Consuntivo=("Imp. Cons. di Period","sum"),
ConsuntivoYTD=("Imp. Cons. YTD","sum")
     ).reset_index()
#for each resource type 
wTipoRisorsa=["INT","EST","ALTRO"]
for wRis in wTipoRisorsa:
     df_prog_byTipoRis["Pian_"+wRis]=np.where((df_prog_byTipoRis["TipoRis"]==wRis),df_prog_byTipoRis["Pianif_Corr"],0 )
     df_prog_byTipoRis["Cons_"+wRis]=np.where((df_prog_byTipoRis["TipoRis"]==wRis),df_prog_byTipoRis["Consuntivo"],0 ) 
#group resources 
df_task_byTipoRis=df_prog_byTipoRis.groupby(
["Codice Task"]).agg(Pian_Int= ("Pian_INT","sum"),
Pian_Est=("Pian_EST","sum"),
Pian_Altro =("Pian_ALTRO","sum"),
Cons_Int=("Cons_INT","sum"),
Cons_Est=("Cons_EST","sum"),
Cons_Altro=("Cons_ALTRO","sum")
     ).reset_index()

df_prog_bytask=df_prog.groupby(
["Cod. Area Progetto","Cod. Progetto","Descrizione Progetto","Tipologia Progetto","Tipo Rilascio","Stato Progetto","Resp. PM","Resp. MP","Cod. Area MP","Codice MP","Progetto CdF","Codice Task","Task Type","Incarico","Data I. Corrente","Data F. Corrente","Data I. Contrattuale","Data F. Contrattuale","Descrizione Task","CUP","MP_CdC","MP_Dip","PM_Dip","LOB"]).agg(
Pianif_Corr =("Imp. Prev. Corr.","sum"),
Pianif_Contr=("Imp. Prev. Contr.","sum"),
Consuntivo=("Imp. Cons. di Period","sum"),
ConsuntivoYTD=("Imp. Cons. YTD","sum")
     ).reset_index()
df_prog_bytask=pd.merge(df_prog_bytask,df_Rilasci_task[["Tipo Incarico", "Cod_Rilascio","Tipo_Rilascio","Cod_PPA","Cod_CUP","Numero Task"]],left_on="Codice Task", right_on="Numero Task",how="left",suffixes=('', '_Ril'))
df_prog_bytask=pd.merge(df_prog_bytask,df_task_byTipoRis,left_on="Codice Task", right_on="Codice Task",how="left",suffixes=('', '_ris'))
df_prog_bytask["Figur."]=np.where(df_prog_bytask["Codice Task"].str.contains("FIG", regex=True),"S","N" )

df_prog_bytask=pd.merge(df_prog_bytask,w_eccez_Task,on="Codice Task",how="left",suffixes=('', '_pos'))
df_prog_bytask["PM_Dip"]= np.where(df_prog_bytask['eccez_Dip'].notnull(),df_prog_bytask['eccez_Dip'],df_prog_bytask['PM_Dip'])

#Delete unwanted columns
df_prog_bytask = df_prog_bytask.drop(["Numero Task"], axis=1)
logger.info('end reading tasks ' + str(df_prog_bytask.shape[0]))