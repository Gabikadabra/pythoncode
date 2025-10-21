#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from Util_leggeParm import parm_dict
from funct.Dframe_toSheet import scriveFoglio_dict 
#Read dummy tables
from read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
from read_Interni import df_RisInt_byName
wrk_environm=value = parm_dict.get("parm_environment", "Develop")
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
#from read_Interni import df_RisInt_byCID,df_RisInt_byName
import os
import sys
from read_rilasci import df_RilasciPPA_sub
#
from read_verbAtt import df_elab_RilPPA, df_elab_Imp_Spalm, df_elab_Verb_Spalm, df_elab_Verb_Fatt, df_elab_RefRil,df_RilInc, df_elab_Rilasci,df_elab_Incarichi, df_RCA_Verb, df_elab_Verbali, df_Incarichi,df_RefInc,df_RCA_Verb_FirmatoG,df_RCA_Verb_ChiusoG, df_RilasciPPA,df_RCA_Azioni
wrk_RptDir=dir_dict["rptDir"]
import time
wrk_VAT=1+.22
wDtFormat = "%d/%m/%Y" 
wTimestr = time.strftime("%Y%m%d")
df_elab_Rilasci=df_elab_Rilasci.rename(columns={"Cod Rilascio":"cod rilascio" })


#Aggancia il CUP 
df_Rilasci_byCUP=df_RilasciPPA.groupby(['Codice Rilascio']).agg(CodiceCup = ("Codice CUP","unique")
     ).reset_index()
df_Rilasci_byCUP=df_Rilasci_byCUP.rename(columns={"Codice Rilascio":"cod rilascio" })
#Convert list column in a string 
df_Rilasci_byCUP['CodiceCup']=[','.join(map(str, l)) for l in df_Rilasci_byCUP['CodiceCup']]

#Get Incarichi by Rilascio
df_RilIncarichi=pd.merge(df_elab_Rilasci,df_elab_Incarichi,left_on="cod incarico",right_on="cod incarico", how="left",suffixes=('', '_sal'))
#Get CUP by Rilascio
df_RilIncarichi=pd.merge(df_RilIncarichi,df_Rilasci_byCUP,left_on="cod rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))
#Get Valore by Verbali Firmati
df_RilIncarichi=pd.merge(df_RilIncarichi,df_RCA_Verb_FirmatoG,left_on="cod rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))
#Get Valore by Verbali Chiusi
df_RilIncarichi=pd.merge(df_RilIncarichi,df_RCA_Verb_ChiusoG,left_on="cod rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))

df_RilIncarichi["Imp rilascio"]=df_RilIncarichi["Imp rilascio"]*wrk_VAT
df_RilIncarichi["Importo totale"]=df_RilIncarichi["Importo totale"]*wrk_VAT
df_RilIncarichi["tot_Firmato"]=df_RilIncarichi["tot_Firmato"]*wrk_VAT
df_RilIncarichi["tot_Chiuso"]=df_RilIncarichi["tot_Chiuso"]*wrk_VAT
#get list of Verbali
df_elab_Verbali_bycodVerb=df_elab_Verbali.groupby(["cod verbale",	"cod rilascio",	"tipo verbale",	"stato verbale",	"dt creazione",	"dt firma",	"dt chiusura"
], dropna=False).agg(
ImportoVerbale= ("importo verbale","sum")).reset_index()
#Get Info about Rilascio
df_elab_Verbali_bycodVerb=pd.merge(df_elab_Verbali_bycodVerb,df_RilIncarichi,left_on="cod rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))
df_elab_Verbali_bycodVerb=pd.merge(df_elab_Verbali_bycodVerb,df_RCA_Azioni,left_on="cod verbale",right_on="codice sap", how="left",suffixes=('', '_sal'))

df_elab_Verbali_bycodVerb["ImportoVerbale"]=df_elab_Verbali_bycodVerb["ImportoVerbale"]*wrk_VAT


df_elab_Verb_Fatt=pd.merge(df_elab_Verb_Fatt,df_elab_Verbali_bycodVerb,left_on="Cod Rilascio",right_on="cod rilascio", how="left",suffixes=('', '_sal'))
#df_elab_Verb_Fatt["ImportoVerbale"]=df_elab_Verb_Fatt["ImportoVerbale"]*wrk_VAT
df_elab_Verb_Fatt["Imp fattura"]=df_elab_Verb_Fatt["Imp fattura"]*wrk_VAT
df_elab_Verb_Fatt["fattura liquidata"]=df_elab_Verb_Fatt["fattura liquidata"]*wrk_VAT

wDtFormat = "%d/%m/%Y" 
wTimestr = time.strftime("%Y%m%d")
pass

