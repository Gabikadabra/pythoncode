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
#from read_rilasci import df_Rilasci_resp, df_Rilasci_subt
import os
import sys
import re
#Remove chars not numeric form column
def remove_chars(s):
#2025.04.08 begin
#   return re.sub("[^0-9,]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
    a= re.sub("[^0-9,]+","",str(s))  # to remove all non alphabets & numbers & commas & dots  re.sub(r"[^a-zA-Z0-9,.]+", '', s)
    a=a.strip()
    return a.replace(",",".")
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]

wpath_fpass = wrk_dir_inp+"\\Staging\\fatture"
filenames = [file for file in os.listdir(wpath_fpass) if file.startswith('FatturePassive')]
df_fpass = pd.concat([pd.read_excel(wpath_fpass +"\\"+ file) for file in filenames], ignore_index=True)
#
filenames = [file for file in os.listdir(wpath_fpass) if file.startswith('FattureLiquidate')]
df_fliquid = pd.concat([pd.read_excel(wpath_fpass  +"\\"+ file) for file in filenames], ignore_index=True)
#--------------End


wrk_dir_inp2=wrk_dir_inp+"\\Staging\\cicloattivo"

wNamFile='Da WeeklyCA_Verb.xlsx'

#xlsx = pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=None, header=None)
#for sheet in xlsx.keys(): xlsx[sheet].to_excel(sheet+'.xlsx', header=False, index=False)

wNamSht = "RCA_INCARICHI"; df_RCA_Incarichi=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"CODICE_SAP": str})  
wNamSht = "RCA_REFERENTI_RL"; df_RCA_RefRL=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"cod incarico":str,"num versione":str, "Nome Referente":str,"Ruolo":str,"Ente":str,"Ordine approvazione":str})  
wNamSht = "RCA_AVANZAMENTI"; df_RCA_Avanzam=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"cod_rilascio": str})  
wNamSht = "RCA_SPALMATURE_IMPEGNI"; df_RCA_ImpSpalm=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"num impegno":str, "anno impegno":str, "cod rilascio":str, "num capitolo":str, "anno capitolo":str, "num decreto":str, "anno_decreto":str})  
wNamSht = "RCA_SPALMATURE_VERBALI"; df_RCA_VerbSpalm=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"cod verbale":str, "cod rilascio":str, "num impegno":str, "anno impegno":str, "num capitolo":str, "anno capitolo":str, "num decreto":str, "anno_decreto": str})  

df_RCA_ImpSpalm["Rif_Impegno"]=df_RCA_ImpSpalm["anno impegno"]+"-"+df_RCA_ImpSpalm["num impegno"]
df_RCA_ImpSpalm["Rif_Capitolo"]=df_RCA_ImpSpalm["anno capitolo"]+"-"+df_RCA_ImpSpalm["num capitolo"]
df_RCA_ImpSpalm["Rif_Decreto"]=df_RCA_ImpSpalm["anno_decreto"]+"-"+df_RCA_ImpSpalm["num decreto"]
df_RCA_ImpSpalm["Rif_Finanz"]="#I"+df_RCA_ImpSpalm["num impegno"]+"_#C"+df_RCA_ImpSpalm["num capitolo"]+"_#D"+df_RCA_ImpSpalm["num decreto"]
df_RCA_ImpSpalm["Ril_Finanz"]=df_RCA_ImpSpalm["cod rilascio"]+"_"+df_RCA_ImpSpalm["Rif_Finanz"]

df_RCA_VerbSpalm["Rif_Impegno"]=df_RCA_VerbSpalm["anno impegno"]+"-"+df_RCA_VerbSpalm["num impegno"]
df_RCA_VerbSpalm["Rif_Capitolo"]=df_RCA_VerbSpalm["anno capitolo"]+"-"+df_RCA_VerbSpalm["num capitolo"]
df_RCA_VerbSpalm["Rif_Decreto"]=df_RCA_VerbSpalm["anno_decreto"]+"-"+df_RCA_VerbSpalm["num decreto"]
df_RCA_VerbSpalm["Rif_Finanz"]="#I"+df_RCA_VerbSpalm["num impegno"]+"_#C"+df_RCA_VerbSpalm["num capitolo"]+"_#D"+df_RCA_VerbSpalm["num decreto"]
df_RCA_VerbSpalm["Ril_Finanz"]=df_RCA_VerbSpalm["cod rilascio"]+"_"+df_RCA_VerbSpalm["Rif_Finanz"]
#Add verbali corpo + canone + consumo


#df_RCA_Verb = df_RCA_VerbCorpo.append(df_RCA_VerbCons)
#df_RCA_Verb = df_RCA_Verb.append(df_RCA_VerbCan)

wNamFile='Da_ZAttivo_Fatture.xlsx'
#xlsx = pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=None, header=None)
#for sheet in xlsx.keys(): xlsx[sheet].to_excel(sheet+'.xlsx', header=False, index=False)
df_RCA_Verb_Fatt=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, dtype={"Num fattura":str, "Num Fattura Stornata":str, "Dt Fattura":str, "Codice Verbale":str, "Cod Rilascio":str, "Num impegno":str, "Anno impegno":str, "Num capitolo":str, "Anno capito":str,"Num decreto":str, "Anno decreto":str})  
df_RCA_Verb_Fatt["Rif_Finanz"]="#I"+df_RCA_Verb_Fatt["Num impegno"]+"_#C"+df_RCA_Verb_Fatt["Num capitolo"]+"_#D"+df_RCA_Verb_Fatt["Num decreto"]
df_RCA_Verb_Fatt["Ril_finanz"]=df_RCA_Verb_Fatt["Cod Rilascio"]+"_"+df_RCA_Verb_Fatt["Rif_Finanz"]
#remove Fattura Stornata dall'elenco and Fattura di storno
wFatt_Stornate=df_RCA_Verb_Fatt["Num Fattura Stornata"].unique()
df_RCA_Verb_Fatt=df_RCA_Verb_Fatt.loc[(~df_RCA_Verb_Fatt["Num fattura"].isin(wFatt_Stornate))&(df_RCA_Verb_Fatt["Num Fattura Stornata"].isnull())]

wNamFile='Da_ZAttivo_FattureLiq.xlsx'
#xlsx = pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=None, header=None)
#for sheet in xlsx.keys(): xlsx[sheet].to_excel(sheet+'.xlsx', header=False, index=False)
df_RCA_Verb_FattLiq=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, dtype={"Num fattura":str})  
df_RCA_Verb_Fatt=pd.merge(df_RCA_Verb_Fatt,df_RCA_Verb_FattLiq,left_on="Num fattura", right_on="Num fattura",how="left",suffixes=('', '_verb'))
#df_RCA_Verb_Fatt.loc[df_RCA_Verb_Fatt['Importo Liquidato'].notnull(), ['fattura liquidata']] = [4, 8]

df_RCA_Verb_Fatt["fattura liquidata"]=np.where(df_RCA_Verb_Fatt['Importo Liquidato'].notnull(),
                                        df_RCA_Verb_Fatt['Imp fattura'],0
                                         )


wNamFile='Da DailyCA_Verb.xlsx'
#xlsx = pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=None, header=None)
#for sheet in xlsx.keys(): xlsx[sheet].to_excel(sheet+'.xlsx', header=False, index=False)
wNamSht = "RCA_PERSONE_AZIONI_IN_CARICO"; df_RCA_Azioni=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"codice sap":str, "stato":str, "cognome":str, "nome":str, "email":str})  
wNamSht = "RCA_VERBALI_CORPO"; df_RCA_VerbCorpo=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"cod verbale":str,"cod rilascio":str, "tipo verbale":str, "stato verbale":str})  
wNamSht = "RCA_VERBALI_CONSUMO"; df_RCA_VerbCons=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"cod verbale":str,"cod rilascio":str, "tipo verbale":str, "stato verbale":str})  
wNamSht = "RCA_VERBALI_CANONE"; df_RCA_VerbCan=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, sheet_name=wNamSht, dtype={"cod verbale":str,"cod rilascio":str, "tipo verbale":str, "stato verbale":str})  
df_RCA_Verb = pd.concat([df_RCA_VerbCorpo, df_RCA_VerbCons, df_RCA_VerbCan], ignore_index=True, sort=False)
df_RCA_Verb=pd.merge(df_RCA_Verb,df_RCA_Azioni,left_on="cod verbale", right_on="codice sap",how="left",suffixes=('', '_verb'))
#df_RCA_Verb['importo verbale'] = df_RCA_Verb['importo verbale'].astype(float)
df_RCA_Verb['importo verbale']=df_RCA_Verb['importo verbale'].apply(remove_chars)

df_RCA_Verb['importo verbale']=df_RCA_Verb['importo verbale'].apply(pd.to_numeric,errors="coerce").fillna(0)

df_elab_Verb_Spalm=df_RCA_VerbSpalm.groupby(["cod verbale","cod rilascio", "num impegno",	"anno impegno",	"num capitolo",	"anno capitolo",	"num decreto",	"anno_decreto","mensilita"], dropna=False).agg(importo_verbxcap= ("importo_verbxcap","sum")).reset_index()
df_elab_Imp_Spalm=df_RCA_ImpSpalm.groupby(["num impegno", "anno impegno", "cod rilascio", "num capitolo", "anno capitolo", "num decreto", "anno_decreto"], dropna=False).agg(**{"importo impegno": ("importo impegno","sum")}).reset_index()


df_Elab_RCA_VerbSpalm=df_RCA_VerbSpalm.groupby(["cod verbale","cod rilascio", "Rif_Impegno", "Rif_Capitolo", "Rif_Decreto", "Ril_Finanz"], dropna=False).agg(Importo_verbxcap= ("importo_verbxcap","sum")).reset_index()
#summarize RCA Verb separate running Verb and complete status Verb
#Create total column for completed Verb
wVerbApprov_list=["Firmato","Chiuso"]
wVerbInPreparation_list=["In negoziazione","In bozza"]
wVerbNotApprov_list=["Approvazione Ente","In Firma Ente"]
df_RCA_Verb_Appr=df_RCA_Verb.loc[df_RCA_Verb["stato verbale"].isin(wVerbApprov_list)]
df_RCA_Verb_NotAppr=df_RCA_Verb[df_RCA_Verb["stato verbale"].isin(wVerbNotApprov_list)]
df_RCA_Verb_InPrep=df_RCA_Verb[df_RCA_Verb["stato verbale"].isin(wVerbInPreparation_list)]
#Firma / Chiuso
df_RCA_Verb_Firmato=df_RCA_Verb[df_RCA_Verb["stato verbale"].isin(["Firmato"])]
df_RCA_Verb_Chiuso=df_RCA_Verb[df_RCA_Verb["stato verbale"].isin(["Chiuso"])]
#.isin(lis)
df_RCA_Verb_Appr=df_RCA_Verb_Appr[["cod verbale","cod rilascio","stato verbale","importo verbale","mensilita"]]
df_RCA_Verb_InPrep=df_RCA_Verb_InPrep[["cod verbale","cod rilascio","stato verbale","importo verbale","mensilita"]]
df_RCA_Verb_NotAppr=df_RCA_Verb_NotAppr[["cod verbale","cod rilascio","stato verbale","importo verbale","mensilita","cognome","nome"]]
df_RCA_Verb_Appr["status"]=df_RCA_Verb_Appr["cod verbale"]+"("+df_RCA_Verb_Appr["stato verbale"]+")"
df_RCA_Verb_NotAppr["status"]=df_RCA_Verb_NotAppr["cod verbale"]+"("+df_RCA_Verb_NotAppr["stato verbale"]+"#"+df_RCA_Verb_NotAppr["cognome"]+" "+df_RCA_Verb_NotAppr["nome"]+")"

df_RCA_Verb_InPrep["status"]=df_RCA_Verb_InPrep["cod verbale"]+"("+df_RCA_Verb_InPrep["stato verbale"]+")"

df_RCA_Verb_InApprG=df_RCA_Verb_Appr.groupby(
["cod rilascio"], dropna=False).agg(tot_inApprov= ("importo verbale","sum"),
                                                                  statusApprov= ("status","unique")
     ).reset_index()
df_RCA_Verb_NotApprG=df_RCA_Verb_NotAppr.groupby(
["cod rilascio"], dropna=False).agg(tot_notApprov= ("importo verbale","sum"),
                                            statusPrep= ("status","unique")
     ).reset_index()
df_RCA_Verb_InPrepG=df_RCA_Verb_InPrep.groupby(
["cod rilascio"], dropna=False).agg(tot_InPrep= ("importo verbale","sum"), statusPrep= ("status","unique")
     ).reset_index()
#Firma / Chiuso
df_RCA_Verb_FirmatoG=df_RCA_Verb_Firmato.groupby(
["cod rilascio"], dropna=False).agg(tot_Firmato= ("importo verbale","sum")
     ).reset_index()
df_RCA_Verb_ChiusoG=df_RCA_Verb_Chiuso.groupby(
["cod rilascio"], dropna=False).agg(tot_Chiuso= ("importo verbale","sum")
     ).reset_index()


df_RCA_Verb_riep = pd.concat([df_RCA_Verb_InApprG, df_RCA_Verb_NotApprG, df_RCA_Verb_InPrepG], ignore_index=True, sort=False)
#Convert list column in a string 
str_cols = df_RCA_Verb_riep.columns[df_RCA_Verb_riep.dtypes==object]
df_RCA_Verb_riep[str_cols] = df_RCA_Verb_riep[str_cols].fillna('.')
df_RCA_Verb_riep = df_RCA_Verb_riep.fillna(0)


df_RCA_Verb_riep= df_RCA_Verb_riep.fillna('')



df_RCA_Verb_riep['statusPrep']=[','.join(map(str, l)) for l in df_RCA_Verb_riep['statusPrep']]
df_RCA_Verb_riep['statusApprov']=[','.join(map(str, l)) for l in df_RCA_Verb_riep['statusApprov']]

df_RCA_Verb_riep = df_RCA_Verb_riep.groupby(["cod rilascio"], dropna=False).agg(tot_InPrep= ("tot_InPrep","sum"),
                      tot_inApprov= ("tot_inApprov","sum"),tot_notApprov= ("tot_notApprov","sum"),statusApprov= ("statusApprov","unique"),statusPrep= ("statusPrep","unique")
     ).reset_index()
df_RCA_Verb_riep['statusPrep']=df_RCA_Verb_riep['statusPrep'].apply(lambda x: str(x).replace('[','').replace(']','').replace('\'', '').replace(' ',''))
df_RCA_Verb_riep['statusApprov']=df_RCA_Verb_riep['statusApprov'].apply(lambda x: str(x).replace('[','').replace(']','').replace('\'', '').replace(' ',''))

#df_RCA_Verb_riep['statusPrep']=df_RCA_Verb_riep['statusPrep'].apply(lambda x: x[1:-1])
#df_RCA_Verb_riep['statusApprov']=df_RCA_Verb_riep['statusApprov'].apply(lambda x: x[1:-1])


#### Add Rilasci
df_Rilasci_cln=pd.read_excel(wrk_dir_inp+"\\"+"da ZRilasci.xlsx",dtype={"INCARICO": str, "Codice CUPCodice CUP": str,"Codice Cliente": str})
df_Rilasci_cln.rename(columns={ df_Rilasci_cln.columns[21]: "Numero Task" }, inplace = True)
#Use legacy table
#Filter stato
wRil_Status = ['Chiusura Amministrativa', 'Concluso', 'Sospeso','Annullato']
df_Rilasci_cln=df_Rilasci_cln.loc[~df_Rilasci_cln["Stato Incarico"].isin(wRil_Status)]

#filter Profit Center MP DCL DA* WA*
df_Rilasci_cln=df_Rilasci_cln.loc[(df_Rilasci_cln["Profit Center MP"].notnull())] # tolto controllo 17/7/2025 &(df_Rilasci_cln["Codice Rilascio Contr."].notnull())
df_Rilasci_cln=df_Rilasci_cln.loc[~df_Rilasci_cln["Profit Center MP"].str.startswith(("LA","DA"))]
#distinct list of Rilasci
df_Rilasci_cln["Codice Rilascio Contr."]=df_Rilasci_cln["Codice Rilascio Contr."].str.strip() #clean blanks 

df_Rilasci_cln["INCARICO"]=df_Rilasci_cln["INCARICO"].str.strip() #clean blanks 


#filter Rilasci not "Giunta"
df_Rilasci_cln=df_Rilasci_cln.loc[df_Rilasci_cln["Codice Cliente"].str.contains("5000039", na=False)]

wIncarichi_cln_list=df_Rilasci_cln["INCARICO"].unique()

#Extract Rilasci
wNamFile='Da_ZAttivo_Rilasci.XLSX'
df_RilInc=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, dtype={"Cod Rilascio":str,"cod incarico":str,"Titolo rilascio":str,"Tipo rilascio":str,"Dt inizio":str,"Dt fine":str,"Dt validazione":str,"Periodicità":str})  
df_RilInc["Cod Rilascio"]=df_RilInc["Cod Rilascio"].str.strip() #clean blanks 

df_elab_Rilasci=df_RilInc.loc[df_RilInc["cod incarico"].isin(wIncarichi_cln_list)]
wRilasci_cln_list=df_elab_Rilasci["Cod Rilascio"].unique()



#if "24S66" in wIncarichi_cln_list:
#    if "24S66-02" not in wRilasci_cln_list: wRilasci_cln_list=np.append(wRilasci_cln_list,"24S66-02")

#if "24S11" in wIncarichi_cln_list:
#    if "24S11-01" not in wRilasci_cln_list: wRilasci_cln_list=np.append(wRilasci_cln_list,"24S11-01")
#    if "24S11-02" not in wRilasci_cln_list: wRilasci_cln_list=np.append(wRilasci_cln_list,"24S11-02")
#    if "24S11-03" not in wRilasci_cln_list: wRilasci_cln_list=np.append(wRilasci_cln_list,"24S11-03")



#Extract Incarichi
wNamFile='Da_ZAttivo_Incarichi.XLSX'
df_Incarichi=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, dtype={"cod incarico":str,"Titolo incarico":str,"Anno incarico":str,"Cod DG RL":str,"Cod Struttura ARIA":str,"Stato incarico":str,"Tipo incarico":str,"Dt inizio":str,"Dt fine":str,"Num versione":str})  
df_elab_Incarichifull=df_Incarichi
df_elab_Incarichi=df_Incarichi.loc[df_Incarichi["cod incarico"].isin(wIncarichi_cln_list)]

#Extract Ref ARIA
wNamFile='Da_ZAttivo_Ref.XLSX'
df_RefInc=pd.read_excel(wrk_dir_inp2+"\\"+wNamFile, dtype={"cod incarico":str,"Cod Rilascio":str,"Num versione":str,"Nome Referente":str,"Ruolo":str,"Ente":str,"Flag":str})  
#df_RefInc_vers=df_RefInc.drop_duplicates(subset=["cod incarico", "Num versione"], ignore_index=True)

df_RefInc_cln=df_RefInc.loc[df_RefInc["cod incarico"].isin(wIncarichi_cln_list)]
df_RefInc_cln=df_RefInc_cln.loc[(df_RefInc_cln["Ruolo"]!="DAFC")] #&(df_RefInc_cln["Ordine approvazione"]!=2)
df_RefInc_cln["Ordine approvazione"]=df_RefInc_cln["Ordine approvazione"].astype(str)
#df_RilasciPPA["Importo Totale Rilascio"].str.replace('.', '').str.replace(",",".").astype(float)
w_colnumeric_toclean=["importo verbale"]
#df_RCA_Verb[w_colnumeric_toclean]=df_RCA_Verb[w_colnumeric_toclean].apply(lambda x: x.str.replace('.', '')).apply(lambda x: x.str.replace(',', '.')).astype(float)
#for w_colnamecln in w_colnumeric_toclean:#
#	df_RCA_Verb[w_colnamecln]=df_RCA_Verb[w_colnamecln].apply(remove_chars);df_RCA_Verb[w_colnamecln]=df_RCA_Verb[w_colnamecln].apply(pd.to_numeric,errors="coerce").fillna(0)

#clean Verb 
df_elab_Verbali=df_RCA_Verb.loc[df_RCA_Verb["cod rilascio"].isin(wRilasci_cln_list)]


#Extract PPA
wNamFile='Da ZRilasci_PPA_UE.xlsx'
df_RilasciPPA=pd.read_excel(wrk_dir_inp+"\\"+wNamFile, dtype={"Codice Contratto": str, "Codice CUP":str, "Numero Capitolo":str, "Anno Capitolo":str})
df_RilasciPPA=df_RilasciPPA[df_RilasciPPA["Tipo Rilascio"].notnull()]
wCharColumns=["Codice Rilascio","Codice Attività PPA","Codice CUP","Numero Capitolo","Anno Capitolo","Classificazione del rilascio"]
wNumColumns=["Quota rilascio a valere capitolo e PPA"]
df_RilasciPPA[wCharColumns]=df_RilasciPPA[wCharColumns].fillna("")
df_RilasciPPA[wNumColumns]=df_RilasciPPA[wNumColumns].fillna(0)
df_RilasciPPA["Codice Rilascio"]=df_RilasciPPA["Codice Rilascio"].str.strip()
#Convert from string to float
df_RilasciPPA["Importo Totale Rilascio"]=df_RilasciPPA["Importo Totale Rilascio"].str.replace('.', '').str.replace(",",".").astype(float)
#w_colnumeric_toclean=["Importo Totale Rilascio"]
#for w_colnamecln in w_colnumeric_toclean:
#	df_RilasciPPA[w_colnamecln]=df_RilasciPPA[w_colnamecln].apply(remove_chars);df_RilasciPPA[w_colnamecln]=df_RilasciPPA[w_colnamecln].apply(pd.to_numeric,errors="coerce").fillna(0)


df_Rilasci_bycod=df_RilasciPPA.drop_duplicates(subset=["Codice Rilascio", "Classificazione del rilascio"], ignore_index=True)
#df_Rilasci_bycod=df_RilasciPPA.groupby(["Codice Rilascio"]).max()["Classificazione del rilascio"].reset_index()

df_Rilasci_bycod=df_Rilasci_bycod.rename(columns={"Codice Rilascio":"Cod Rilascio"})

df_elab_Rilasci=pd.merge(df_elab_Rilasci,df_Rilasci_bycod[["Cod Rilascio","Classificazione del rilascio"]],left_on="Cod Rilascio", right_on="Cod Rilascio",how="left",suffixes=('', '_ppa'))


df_elab_RilPPA=df_RilasciPPA.loc[df_RilasciPPA["Codice Rilascio"].isin(wRilasci_cln_list)]




#df_RilasciPPA_sub=df_RilasciPPA.groupby(["Codice Rilascio","Codice Attività PPA","Codice CUP","Numero Capitolo","Anno Capitolo"])[["Quota rilascio a valere capitolo e PPA"]].sum().reset_index()

#Extract Ref RL
df_RCA_RefRL=df_RCA_RefRL.rename(columns={"num versione":"Num versione"})
df_RCA_RefRL=df_RCA_RefRL.loc[df_RCA_RefRL["cod incarico"].isin(wIncarichi_cln_list)]

#Extract Spalmature Impegni
df_elab_Imp_Spalm["cod rilascio"]=df_elab_Imp_Spalm["cod rilascio"].str.strip()
#df_elab_Imp_Spalm2=df_elab_Imp_Spalm.loc[df_elab_Imp_Spalm["cod rilascio"].str.contains("20TEC")]
df_elab_Imp_Spalm=df_elab_Imp_Spalm.loc[df_elab_Imp_Spalm["cod rilascio"].isin(wRilasci_cln_list)]
#Extract Spalmature Verbali
df_elab_Verb_Spalm["cod rilascio"]=df_elab_Verb_Spalm["cod rilascio"].str.strip()
df_elab_Verb_Spalm=df_elab_Verb_Spalm.loc[df_elab_Verb_Spalm["cod rilascio"].isin(wRilasci_cln_list)]
#Extract Fatture

df_RCA_Verb_Fatt["Cod Rilascio"]=df_RCA_Verb_Fatt["Cod Rilascio"].str.strip()
df_RCA_Verb_Fatt["Mensilità"]=pd.to_datetime(df_RCA_Verb_Fatt["Mensilità"].astype(str), format='%Y-%m-%d').dt.strftime('%m/%Y')
df_elab_Verb_Fatt=df_RCA_Verb_Fatt.loc[df_RCA_Verb_Fatt["Cod Rilascio"].isin(wRilasci_cln_list)]
#Referenti Aria + RL
df_elab_RefRil = pd.concat([df_RCA_RefRL, df_RefInc_cln], ignore_index=True, sort=False)
wSelRuolo=["Referente Responsabile","Referente Operativo"]
df_elab_RefRil=df_elab_RefRil.loc[(df_elab_RefRil["cod incarico"].isin(wIncarichi_cln_list))&(df_elab_RefRil["Ruolo"].isin(wSelRuolo))]
df_elab_RefRil["Ordine approvazione"]=df_elab_RefRil["Ordine approvazione"].astype(str)
#df_elab_RefRil=df_elab_RefRil.replace("nan","",regex=True)
df_elab_RefRil["Ordine approvazione"].mask(df_elab_RefRil["Ordine approvazione"] == 'nan', '', inplace=True)
pass