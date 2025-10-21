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
#Clean non numeric values
import re
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task


from Util_logging import logger
#Read dummy tables
#rom read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 

wPythProc="read_ODAG"
import timeit  #funzione timeit per valutare durata estrazione
tempoInz=timeit.default_timer()
#Remove chars not numeric form column
def remove_chars(s):
#2025.04.08 begin
#   return re.sub("[^0-9,]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
    a= re.sub("[^0-9,.]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
    a=a.strip()
    return a.replace(",",".")

##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]

wrk_environm=value = dir_dict.get("inpEnv", "Test") #if testFilter select only few  items

df_ODAG=pd.read_excel(wrk_dir_inp+"\\"+"Da DailyRepODAG.xlsx", dtype={"Numero Contratto Gara":str,	"Data Decorrenza":str,	"Data scadenza":str,	"Data inizio fornitura":str,	"Data fine fornitura":str})
df_ODAG= df_ODAG.drop(df_ODAG.iloc[:, 25:],axis = 1)

w_colnumeric_toclean=["Valore Contratto","Valore Impegnato",	"Valore Ordinato approvato","Valore Ordinato"	,"Valore Fatturato"	,"Valore Attestato","Val. Residuo Ordinato","Valore Residuo Impegnato","Importo Subappalto","Ordinato Mensile","Importo Testata Contratto"]
for w_colnamecln in w_colnumeric_toclean:
	df_ODAG[w_colnamecln]=df_ODAG[w_colnamecln].apply(remove_chars);df_ODAG[w_colnamecln]=df_ODAG[w_colnamecln].apply(pd.to_numeric,errors="coerce").fillna(0)



df_ODAGMacr=pd.read_excel(wrk_dir_inp+"\\"+"Da DailyRepMacroclasse.xlsx") #dtype={"Numero Contratto Gara":str,	"Contratto Macroclasse":str }
#df_ODAGMacr["ODAGMacr_txt"] = df_ODAGMacr["Contratto Gara"].str.cat(df_ODAGMacr["Descrizione Macroclasse"], sep = "_")
#df_ODAGMacr["ODAGMacr_nbr"] = df_ODAGMacr["Numero Contratto Gara"].multiply(1000)+(df_ODAGMacr["Contratto Macroclasse"])

##11/04/2025 df_ODAGMacr["Libero su macroclasse"].apply(lambda x: x.replace(' ', '')).str.replace(',', '.')
#******  TEST *** 
# keep only one element
#df_ODAG=df_ODAG.loc[df_ODAG["Numero Contratto Gara"]=="2020331919"]
#df_ODAGMacr=df_ODAGMacr.loc[df_ODAGMacr["Numero Contratto Gara"]=="2020331919"]        
#******  end TEST *** 
##11/04/2025
w_colnumeric_toclean=["Valore Economico macroclasse","Valore Ordinato",	"Valore Ordinato approvato","Valore Impegnato"	,"Valore Fatturato"	,"Valore Attestato","Libero su macroclasse","RDI in attesa BdO","RDI in Composizione"]
for w_colnamecln in w_colnumeric_toclean:
	df_ODAGMacr[w_colnamecln]=df_ODAGMacr[w_colnamecln].apply(remove_chars);df_ODAGMacr[w_colnamecln]=df_ODAGMacr[w_colnamecln].apply(pd.to_numeric,errors="coerce").fillna(0)

"""#begin format columns - reduce size ODAGMacro"""
wLi=["Numero Contratto Gara"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["RDI in attesa BdO","RDI in Composizione"];wDf={value:"float" for value in wLf}
#wLc=["Tipo","I","UMO","Rilevante PNRR","CdC","Numero CIG","Subappalto",	"Codice Interno subappalto","Attività","Gr. merci"	,"Divisa Subappalto","Descr.Fornitore","Stato elaborazione doc. acquisti","Definizione 2 gr. merci",	"Tipo Linea",	"Descrizione Tipo Linea",	"Codice Subappaltatore",	"Contratto Gara",	"Descr. Subappaltatore","Tipo Ricezione","Milestone",	"Affidamento diretto",	"Divisa Subappalto","Contratto superiore","Descr.Referente Operativo","Documento chiuso","Stato di elaborazione doc. acquisti","Annullamento Residuo","Descrizione_DEC","Codice CUP","Versione"];wDc={value:"category" for value in wLc}
wLnbr=wLi+wLf
df_ODAGMacr[wLnbr]=df_ODAGMacr[wLnbr].fillna(0)

#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDf]:
    convert_dict.update(d)

df_ODAGMacr= df_ODAGMacr.astype(convert_dict)


"""#begin format columns - reduce size ODAG"""
wLi=["Numero Contratto Gara"];wDi={value:"int64" for value in wLi}
#wLs=["Prestazione"];wDs={value:"str" for value in wLs}
wLf=["Importo Testata Contratto"];wDf={value:"float" for value in wLf}
#wLc=["Tipo","I","UMO","Rilevante PNRR","CdC","Numero CIG","Subappalto",	"Codice Interno subappalto","Attività","Gr. merci"	,"Divisa Subappalto","Descr.Fornitore","Stato elaborazione doc. acquisti","Definizione 2 gr. merci",	"Tipo Linea",	"Descrizione Tipo Linea",	"Codice Subappaltatore",	"Contratto Gara",	"Descr. Subappaltatore","Tipo Ricezione","Milestone",	"Affidamento diretto",	"Divisa Subappalto","Contratto superiore","Descr.Referente Operativo","Documento chiuso","Stato di elaborazione doc. acquisti","Annullamento Residuo","Descrizione_DEC","Codice CUP","Versione"];wDc={value:"category" for value in wLc}
wLnbr=wLi+wLf
df_ODAG[wLnbr]=df_ODAG[wLnbr].fillna(0)
wList_date=["Data Decorrenza"	,"Data scadenza",	"Data inizio fornitura"]
#Date format
for d in wList_date:
    df_ODAG[d]=pd.to_datetime(df_ODAG[d],errors='coerce')
#Convert columns format from dictionary
convert_dict = {}
for d in [wDi,wDf]:
    convert_dict.update(d)

df_ODAG= df_ODAG.astype(convert_dict)

# ***Filter*** in testfilter is on and specific selection is in place
if (wrk_environm.lower()=="testfilter") & ("*all" not in listfilter_ODAG):
    df_ODAG=df_ODAG.loc[(df_ODAG["Numero Contratto Gara"].isin(listfilter_ODAG))]
    df_ODAGMacr=df_ODAGMacr.loc[(df_ODAGMacr["Numero Contratto Gara"].isin(listfilter_ODAG))]

#convert_dict = {"Tipo": "category", "Doc. acq.": int,"Pos.":int,"I":"category","Ril":int,"Quantità":float,"UMO":"category","Prezzo lordo":float,	"Valore netto":float,"InizVal.":float,	"FineVal.":float,	"Data cr.":float,	"Data del documento acquisto":float,"Fornitore":int,
#"Rilevante PNRR":"category"}   
wrename_dict={"Numero Contratto Gara": "ODAGnbr", "Contratto Gara": "ODAG","Descrizione Macroclasse":"Macrocl_txt","Contratto Macroclasse":"Macrocl_nbr"}
df_ODAG.rename(columns=wrename_dict, inplace=True)
df_ODAGMacr.rename(columns=wrename_dict, inplace=True)

from get_ctrMgmt import df_ODAG_ctrMgmt,df_ODA_ctrMgmt

#get DEC & RUP information
df_ODAG=pd.merge(df_ODAG,df_ODAG_ctrMgmt[["ODAGnbr","DEC","RUP"]],on=["ODAGnbr"], how="left",suffixes=('', '_bdo'))
df_ODAG=df_ODAG[(df_ODAG["RUP"].notnull())&(df_ODAG["RUP"].notnull())] # Eliminate ODAG without DEC AND RUP
#df_ODAG=df_ODAG.drop(["ODAGnbr"],axis=1)
df_ODAGMacr=pd.merge(df_ODAGMacr,df_ODAG_ctrMgmt[["ODAGnbr","DEC","RUP"]],on=["ODAGnbr"], how="left",suffixes=('', '_bdo'))
#df_ODAGMacr=df_ODAGMacr.drop(["ODAGnbr"],axis=1)
pass

tempoFin = timeit.default_timer()
logger.info('end reading ' + str(df_ODAGMacr.shape[0])+" Durata processo: " + wPythProc +" - dim: "+str(df_ODAGMacr.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_ODAGMacr.memory_usage(deep=True)) )

##11/04/2025 m1=df_ODAGMacr["Libero su macroclasse"].str.endswith('-')
##11/04/2025 m2=df_ODAGMacr["Libero su macroclasse"].str[:-1].str.replace(',', '.').astype(float)
##11/04/2025 df_ODAGMacr["Libero su macroclasse"]=np.where(m1,-m2,m2)
pass