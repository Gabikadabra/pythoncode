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
wPythProc="read_Accrual"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
#--------------- Read folder for receiving goods
#for file in os.listdir("/Users/darren/Desktop/test"):
#    if (file.startswith("art"))&(file.endswith("xlsx")):
#        print(file)

df_Accrual=pd.read_excel(wrk_dir_inp+"\\"+"Da_ZAccrual.XLSX")
#Change column(s) type

df_Accrual=df_Accrual[(df_Accrual["Periodo contabile"].notnull())&(df_Accrual["OdA/BdO"].gt((2000000000)))]
# ***Filter*** in testfilter is on and specific selection is in place
#if (wrk_environm.lower()=="testfilter") & ("*all" not in listfilter_BDO):
#    df_Accrual=df_Accrual.loc[(df_Accrual["OdA/BdO"].isin(listfilter_BDO))|(df_Accrual["MAP"].isin(listfilter_MAP))]

wList_date=["DtaReg R/R"	,"Inizio validità",	"Fine validità"]
#Date format
for d in wList_date:
    df_Accrual[d]=pd.to_datetime(df_Accrual[d],errors='coerce')

"""#begin format columns - reduce size"""
wLi=["Numero oggetto ratei/risconti","Periodo contabile","Esercizio","Conto Co.Ge."];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
#wLf=["Importo reg."];wDf={value:"float" for value in wLf}
wLc=["CdC","Descrizione Conto","Nominativo Fornitore","Testo"];wDc={value:"category" for value in wLc}
wLnbr=wList_date
df_Accrual[wLnbr]=df_Accrual[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDc]:
    convert_dict.update(d)
#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
df_Accrual= df_Accrual.astype(convert_dict)
"""#end format columns - reduce size"""




# rename columns
#create dictionary for column(s) format 
#convert_dict = {"OdA/BdO":str,	"Posizione":str,	"MAP":str,	"Nominativo Fornitore":str,"Esercizio":str,"Conto Co.Ge.": str}   # {'A': int, 'C': float} "Periodo contabile"
# Convert columns using the dictionary

#df_Accrual = df_Accrual.astype(convert_dict)


wrenMAPColumns={"OdA/BdO":"Documento Acquisto","Conto Co.Ge.":"contabile","Importo reg.":"ValoreMAP","Nominativo Fornitore":"Fornitore" }
#df_Accrual["MAP"]=df_Accrual["MAP"].str[0:9]
df_Accrual=df_Accrual.rename(columns=wrenMAPColumns)
#df_Accrual =df_Accrual.loc[pd.to_numeric(df_Accrual["Periodo contabile"],errors='coerce').notnull()]
#Format Year
df_Accrual["AnnoInizio"]=df_Accrual["Inizio validità"].dt.year
#df_Accrual["Esercizio"]=df_Accrual["Esercizio"].str[0:4]
#df_Accrual["Periodo contabile"]=df_Accrual["Periodo contabile"].astype('Int64').astype('str')
#df_Accrual["Posizione"]=df_Accrual["Posizione"].str[:-2].str.title()
#df_Accrual["Documento Acquisto"]=df_Accrual["Documento Acquisto"].str[0:10]

#Anno Codice Task
df_Accrual_bymap=df_Accrual.groupby(
["Documento Acquisto","Posizione","MAP","Elemento WBS","Esercizio", "Periodo contabile","AnnoInizio"]).agg(
                                            ValoreMAP =("ValoreMAP","sum"),Tot_darateizzare=("Importo totale da rateizzare","max")).reset_index()
#df_Accrual_bymap= df_Accrual_bymap.insert(0, 'NewMAP_ID', range(209900000, 209900000 + len(df_Accrual_bymap)))
#Rename columns

wrenMAPColumns={"Esercizio":"Anno competenza","Periodo contabile":"Mese competenza", "Elemento WBS":"Task"}
df_Accrual_bymap=df_Accrual_bymap.rename(columns=wrenMAPColumns)
df_Accrual_bymap["Cod.Accett.MAP"]="X"
df_Accrual_bymap["Tipo ordine acquisto"]="Z3"

df_Accrual_bytaskODA=df_Accrual_bymap.groupby(
["Task","Anno competenza","AnnoInizio","Documento Acquisto","Posizione"]).agg(
                                            ValoreMAP =("ValoreMAP","sum"),Tot_darateizzare=("Tot_darateizzare","max") ).reset_index()
#create new column value based on condition 
df_Accrual_bytaskODA["ChgAmt"]=np.where(df_Accrual_bytaskODA["Anno competenza"]==df_Accrual_bytaskODA["AnnoInizio"],
                                        df_Accrual_bytaskODA["ValoreMAP"]-df_Accrual_bytaskODA["Tot_darateizzare"],df_Accrual_bytaskODA["ValoreMAP"]
                                         )
df_Accrual_bytask=df_Accrual_bytaskODA.groupby(
["Task","Anno competenza"]).agg(
                                            ValoreMAP =("ValoreMAP","sum"),ChgAmt=("ChgAmt","sum") ).reset_index()
#Rename columns
wrenMAPColumns={"Task":"Codice Task","Anno competenza":"Anno"}
df_Accrual_bytask=df_Accrual_bytask.rename(columns=wrenMAPColumns)

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_Accrual_bytask.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_Accrual_bytask.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_Accrual_bytask.memory_usage(deep=True)) )


pass
