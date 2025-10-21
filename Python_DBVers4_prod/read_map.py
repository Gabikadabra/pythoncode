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
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Accrual import df_Accrual, df_Accrual_bymap
from read_BefQuiet import df_BEF_Map
#from read_rilasci import df_Rilasci_task
import os
import sys
import logging
import logging.config
import timeit  #funzione timeit per valutare durata estrazione

logging.config.fileConfig('logging2.conf')
tempoInz = timeit.default_timer()
# create logger
logger = logging.getLogger('simpleExample')
##Read working directories 
#new_dir = os.getcwd()
wrk_dir_inp=dir_dict["inpDir"]
#--------------- Read folder for receiving goods
#for file in os.listdir("/Users/darren/Desktop/test"):
#    if (file.startswith("art"))&(file.endswith("xlsx")):
#        print(file)

# rename columns
wpath_fpass = wrk_dir_inp+"\\MAPCons"
#wpath_fpass=r"\Users\fcaneri\OneDrive - ARIA S.p.A\PMO Shared\ImportDati\MAPCons"

logger.info('start reading MAP')

filenames = [file for file in os.listdir(wpath_fpass) if file.endswith('.xlsx')]
#******  TEST *** 
# keep only last list's element
# keep only third element
#filenames = [filenames[0]]
#filenames = filenames[-1:]
#******  end TEST *** 
pass

df_MAP_tot = pd.concat([pd.read_excel(wpath_fpass +"\\"+ file, decimal=',', dtype=str) for file in filenames], ignore_index=True)
#Change column(s) type
#Filter columns
wFiltColumns=["Documento Acquisto","Posizione","Tipo ordine acquisto","Fornitore","Testo breve Doc.","Numero MAP","Cod.Accett.MAP","Testo breve MAP","Valore MAP","Task","Campo ut. 1 el. WBS","Milestone MAP","Mese competenza","Anno competenza","Cod. fornitore pos.","Cod.Sub.re.","Cod.Risorsa",'Tipo Fatt. MAP',"Qu.Map"]

df_MAP_tot=df_MAP_tot[wFiltColumns]
#df_MAP_tot.drop(wFiltColumns, axis=1, inplace=True)

# rename columns
wrenMAPColumns={"Numero MAP":"MAP","Conto Co.Ge.":"contabile","Valore MAP":"ValoreMAP","Campo ut. 1 el. WBS":"CUP" }
df_MAP_tot=df_MAP_tot.rename(columns=wrenMAPColumns)
df_MAP_tot['ValoreMAP'] = pd.to_numeric(df_MAP_tot['ValoreMAP'], errors='coerce')
df_MAP_tot['Qu.Map'] = pd.to_numeric(df_MAP_tot['Qu.Map'], errors='coerce')

#create dictionary for column(s) format 
df_MAP_tot["Cod.Risorsa"].fillna("0", inplace = True)
#df_MAP_tot.rename(columns={ df_MAP_tot.columns[1]: "Documento Acquisto" }, inplace = True)
convert_dict = {"Documento Acquisto": str, "Posizione": str,"Fornitore":str,"Cod.Risorsa":str,"Mese competenza":str,	"Anno competenza":str,"Cod. fornitore pos.":str,"Cod.Sub.re.":str,"Tipo Fatt. MAP":str,"MAP":str, "Cod.Risorsa":str }   # {'A': int, 'C': float}
# Convert columns using the dictionary
df_MAP_tot = df_MAP_tot.astype(convert_dict)
df_MAP_tot=df_MAP_tot.loc[(df_MAP_tot["MAP"].notna())] 
#"Numero MAP":str,
# Define the conversion dictionary
#filter for specific columns notna
#df_MAP_tot=df_MAP_tot[df_MAP_tot["MAP"].notna()]
#df_MAP_tot["MAP"]=df_MAP_tot["MAP"].str[0:9]
#df_MAP_tot["MAP"].fillna("0", inplace = True)
df_MAP_tot["Cod.Risorsa"]=df_MAP_tot["Cod.Risorsa"].str[:-2].str.title()
#### get Accrual for task period
#from read_Accrual import df_Accrual_bymap

#Exclude MAP in accrual list
#df_MAP_tot=pd.merge(df_MAP_tot,df_Accrual_bymap,left_on="MAP", right_on="MAP",how="outer",indicator=True)
#df_MAP_tot=df_MAP_tot[df_MAP_tot['_merge']=='left_only']
wpath_fpass = wrk_dir_inp+"\\MAPCons\\MAPFatture"
#wpath_fpass=r"\Users\fcaneri\OneDrive - ARIA S.p.A\PMO Shared\ImportDati\MAPCons"
filenames = [file for file in os.listdir(wpath_fpass) if file.endswith('.XLSX')]
df_pagam_tot = pd.concat([pd.read_excel(wpath_fpass +"\\"+ file, decimal=',') for file in filenames], ignore_index=True)

wColselMAP=["MAP","Num.Fatt.Fornitore","Data doc. Fattura","Protocollo IVA","Data di pagamento"]
df_pagam_tot_red=df_pagam_tot[wColselMAP] 
#******  TEST *** 
convert_dict = {"MAP":str,"Num.Fatt.Fornitore":str,"Data doc. Fattura":str,"Protocollo IVA":str,"Data di pagamento":str }   # {'A': int, 'C': float}
# Convert columns using the dictionary
df_pagam_tot_red = df_pagam_tot_red.astype(convert_dict)


df_pagam_tot_red=df_pagam_tot_red.groupby(["MAP"]).agg(FattFornit=("Num.Fatt.Fornitore","max"),DataFattura=("Data doc. Fattura","max"), ProtocolloIVA=("Protocollo IVA","max"),DataPagamento=("Data di pagamento","max")).reset_index()



###df_prog_legacy=pd.merge(df_prog_legacy,df_prog_tasklist,left_on="Codice Task", right_on="Codice Task",how="outer",indicator=True)
wColselMAP=["Documento Acquisto","Posizione","MAP","Cod.Accett.MAP","Mese competenza","Anno competenza","Task","ValoreMAP", "Tipo ordine acquisto","Cod.Risorsa","Qu.Map","CUP"]

#Get MAPzero
#, dtype={"Documento Acquisto": str, "Fornitore":str,"Numero MAP":str,"Cod.Risorsa":str,"Conto Co.Ge.":str,"Mese competenza":str,	"Anno competenza":str,"Cod. fornitore pos.":str}
df_MAP_zero=df_MAP_tot.loc[(df_MAP_tot["ValoreMAP"].values==0) & (df_MAP_tot["Tipo ordine acquisto"].values!="ZC") ]
wColzeroMAP=["Documento Acquisto","Mese competenza","Anno competenza"]
#Filter Relevant columns
df_MAP_zero=df_MAP_zero[wColzeroMAP]
df_MAP_zero = df_MAP_zero.drop_duplicates(subset=wColzeroMAP, keep="last")
df_MAP_zero["Periodo"]=df_MAP_zero["Anno competenza"]+"-"+df_MAP_zero["Mese competenza"].str.zfill(2) 
df_MAP_zero["BDO_periodo"]=df_MAP_zero["Documento Acquisto"]+"_"+df_MAP_zero["Periodo"] 

df_MAP_spost = pd.read_excel(wrk_dir_inp+"\\"+"Da MAP_SpostamentoCosti.xlsx", dtype={"Documento contabile": str,"MAP": str,"Documento Acquisto": str,"Mese competenza": str,"Anno competenza": str})
df_MAP_tot = pd.concat([df_MAP_tot, df_MAP_spost], ignore_index=True, sort=False)
# togli e selezione per MAP non accettato 
df_MAP_tot["Cod.Accett.MAP"] = df_MAP_tot["Cod.Accett.MAP"].fillna(value=" ")
#df_MAP_tot=df_MAP_tot.loc[(df_MAP_tot["Cod.Accett.MAP"].notna())] 

###TEMPorary filer
#df_MAP_tot =df_MAP_tot[df_MAP_tot["Documento Acquisto"]=="2023332167" ]


#Filter if Develop environment
if wrk_environm =="Develop":
	df_MAP_tot= df_MAP_tot.loc[df_MAP_tot["Documento Acquisto"].isin(df_BDO_dummy["BDO"])]
#read MAP
logger.info('end reading MAP ' + str(df_MAP_tot.shape[0]))

#Filter MAP Zero
df_MAP_tot_nonZero=df_MAP_tot.loc[(df_MAP_tot["ValoreMAP"].values!=0)]

#Isola i MAP utilizzati successivamente per i ratei
df_MAP_attestato_pre_accrual = df_MAP_tot_nonZero.loc[df_MAP_tot_nonZero["MAP"].isin(df_Accrual_bymap['MAP'])]
#Raggruppa i MAP per task
df_MAP_tot_pre_accrual=df_MAP_attestato_pre_accrual.groupby(["MAP","Task","Anno competenza"])["ValoreMAP"].sum().reset_index()


#Verificare se la seguente istruzione ha senso ...
df_MAP_tot_nonZero=df_MAP_tot_nonZero.loc[ (df_MAP_tot_nonZero["Tipo ordine acquisto"].values!="Z3")|(df_MAP_tot_nonZero["Tipo ordine acquisto"].values!="Z4")|(df_MAP_tot_nonZero["Tipo ordine acquisto"].values!="ZI") ]

#Exclude MAP in MAPAccrual file
df_MAP_tot_nonZero = df_MAP_tot_nonZero.loc[~df_MAP_tot_nonZero["MAP"].isin(df_Accrual_bymap['MAP'])]


#Add MAP of Accrual(s)
df_MAP_tot_nonZero = pd.concat([df_MAP_tot_nonZero, df_Accrual_bymap], ignore_index=True, sort=False)

#Riepiloga i MAP degli accrual per task
df_tot_Accrual_bymap=df_Accrual_bymap.groupby(["MAP","Task","Anno competenza"])["ValoreMAP"].sum().reset_index()
#group MAP utilizzati per attestazione iniziale (innegativo per scalare dal task)
df_MAP_tot_pre_accrual["ValoreMAP"]=-df_MAP_tot_pre_accrual["ValoreMAP"]

df_tot_Accrual_bymap = pd.concat([df_MAP_tot_pre_accrual, df_tot_Accrual_bymap], ignore_index=True, sort=False)

df_tot_Accrual_bymap.fillna({'ValoreMAP':0},inplace=True)

df_tot_Accrual_bytask=df_tot_Accrual_bymap.groupby(["Task","Anno competenza"])["ValoreMAP"].sum().reset_index()


df_MAP_accrual = df_MAP_tot_nonZero.loc[df_MAP_tot_nonZero["MAP"].isin(df_Accrual_bymap['MAP'])]


#df_MAP_byOdaLin_per

df_MAP_byOdaLin_per=df_MAP_tot_nonZero.groupby(["Documento Acquisto", "Posizione", "Tipo ordine acquisto", "Mese competenza", "Anno competenza"]).agg(
ValoreCons= ("ValoreMAP","sum"),Qta= ("Qu.Map","sum")).reset_index()
df_MAP_byOdaLin_per["Periodo"]=df_MAP_byOdaLin_per["Anno competenza"]+"-"+df_MAP_byOdaLin_per["Mese competenza"].str.zfill(2) 
df_MAP_byOdaLin_per["BDOLin"]=df_MAP_byOdaLin_per["Documento Acquisto"]+"-"+df_MAP_byOdaLin_per["Posizione"].str.zfill(3)
df_MAP_byOdaLin_per["BDOLin_periodo"]=df_MAP_byOdaLin_per["BDOLin"]+"_"+df_MAP_byOdaLin_per["Periodo"]

df_MAP_byOdaLin=df_MAP_tot_nonZero.groupby(["Documento Acquisto", "Posizione", "Tipo ordine acquisto"]).agg(
ValoreCons= ("ValoreMAP","sum"),QtaCons= ("Qu.Map","sum")).reset_index()
df_MAP_byOdaLin["BDOLin"]=df_MAP_byOdaLin["Documento Acquisto"]+"-"+df_MAP_byOdaLin["Posizione"].str.zfill(3)
# To create df_MAP_byOda_per without 
#df_MAP_byOdaLinCons=df_MAP_tot_nonZero[df_MAP_tot_nonZero["Cod.Accett.MAP"]!=" "]

df_MAP_byOda_per=df_MAP_tot_nonZero.groupby(["Documento Acquisto", "Tipo ordine acquisto", "Mese competenza", "Anno competenza"]).agg(
ValoreMAP= ("ValoreMAP","sum")).reset_index()
df_MAP_byOda_per["Periodo"]=df_MAP_byOda_per["Anno competenza"]+"-"+df_MAP_byOda_per["Mese competenza"].str.zfill(2) 
df_MAP_byOda_per["ODA_periodo"]=df_MAP_byOda_per["Documento Acquisto"]+"_"+df_MAP_byOda_per["Periodo"]
###create net df_MAP_byOda_perCons  
# 
# 
df_MAP_bytaskAll=df_MAP_tot_nonZero.groupby(["Documento Acquisto", "Posizione", "MAP", "Cod.Accett.MAP", "Mese competenza", "Anno competenza", "Task", "Tipo ordine acquisto","Cod.Risorsa"]).agg(
ValoreMAP= ("ValoreMAP","sum"),Qta= ("Qu.Map","sum")).reset_index()
df_MAP_noCons=df_MAP_bytaskAll.loc[(df_MAP_bytaskAll["Cod.Accett.MAP"]==" ") & (df_MAP_bytaskAll["Tipo ordine acquisto"]=="Z4")]
df_MAP_noCons["BDOLin"]=df_MAP_noCons["Documento Acquisto"]+"-"+df_MAP_noCons["Posizione"].str.zfill(3)
df_MAP_noCons["Periodo"]=df_MAP_noCons["Anno competenza"]+"-"+df_MAP_noCons["Mese competenza"].str.zfill(2) 
df_MAP_noCons["BDOLin_periodo"]=df_MAP_noCons["BDOLin"]+"_"+df_MAP_noCons["Periodo"] 
df_MAP_noCons=df_MAP_noCons.rename(columns={"ValoreMAP":"MAP_NoCons","Documento Acquisto":"BDO"})
df_MAP_noCons_byLin=df_MAP_noCons.groupby(["BDO", "BDOLin", "Periodo", "BDOLin_periodo"]).agg(
MAP_NoCons= ("MAP_NoCons","sum")).reset_index()

df_MAPCons=df_MAP_tot_nonZero.loc[(df_MAP_tot_nonZero["Cod.Accett.MAP"]=="X")]

#-----  inizio
df_MAPCons_byOdaLin_per=df_MAPCons.groupby(["Documento Acquisto", "Posizione", "Tipo ordine acquisto", "Mese competenza", "Anno competenza"]).agg(
ValoreCons= ("ValoreMAP","sum"),Qta= ("Qu.Map","sum")).reset_index()
df_MAPCons_byOdaLin_per["Periodo"]=df_MAPCons_byOdaLin_per["Anno competenza"]+"-"+df_MAPCons_byOdaLin_per["Mese competenza"].str.zfill(2) 
df_MAPCons_byOdaLin_per["BDOLin"]=df_MAPCons_byOdaLin_per["Documento Acquisto"]+"-"+df_MAPCons_byOdaLin_per["Posizione"].str.zfill(3)
df_MAPCons_byOdaLin_per["BDOLin_periodo"]=df_MAPCons_byOdaLin_per["BDOLin"]+"_"+df_MAPCons_byOdaLin_per["Periodo"]

df_MAPCons_byOdaLin=df_MAPCons.groupby(["Documento Acquisto", "Posizione", "Tipo ordine acquisto"]).agg(
ValoreCons= ("ValoreMAP","sum"),Qta= ("Qu.Map","sum")).reset_index()
df_MAPCons_byOdaLin["BDOLin"]=df_MAPCons_byOdaLin["Documento Acquisto"]+"-"+df_MAPCons_byOdaLin["Posizione"].str.zfill(3)
# To create df_MAP_byOda_per without 
#df_MAP_byOdaLinCons=df_MAP_tot_nonZero[df_MAP_tot_nonZero["Cod.Accett.MAP"]!=" "]

df_MAPCons_byOda_per=df_MAPCons.groupby(["Documento Acquisto", "Tipo ordine acquisto", "Mese competenza", "Anno competenza"]).agg(
ValoreMAP= ("ValoreMAP","sum")).reset_index()
df_MAPCons_byOda_per["Periodo"]=df_MAPCons_byOda_per["Anno competenza"]+"-"+df_MAPCons_byOda_per["Mese competenza"].str.zfill(2) 
df_MAPCons_byOda_per["ODA_periodo"]=df_MAPCons_byOda_per["Documento Acquisto"]+"_"+df_MAPCons_byOda_per["Periodo"]

#-----  fine
#Select all the MAP with flag cons
df_MAP_bytaskAll=df_MAP_bytaskAll.loc[df_MAP_bytaskAll["Cod.Accett.MAP"]=="X"]
df_MAP_bytask=df_MAP_bytaskAll.groupby(["Documento Acquisto", "Posizione", "MAP", "Cod.Accett.MAP", "Mese competenza", "Anno competenza", "Task","Tipo ordine acquisto","Cod.Risorsa"]).agg(
MAP_Cons= ("ValoreMAP","sum"),Qta= ("Qta","sum")).reset_index()
df_MAP_bytask["BDOLin"]=df_MAP_bytask["Documento Acquisto"]+"-"+df_MAP_bytask["Posizione"].str.zfill(3)
df_MAP_bytask["Periodo"]=df_MAP_bytask["Anno competenza"]+"-"+df_MAP_bytask["Mese competenza"].str.zfill(2) 
df_MAP_bytask["BDOLin_periodo"]=df_MAP_bytask["BDOLin"]+"_"+df_MAP_bytask["Periodo"] 
#from MAPby Task table get BEF
df_BEF_Map_r=df_BEF_Map[["MAP","BEF"]]

df_BEF_Map_r.set_index("MAP",inplace=True)
df_MAP_bytask.set_index("MAP",inplace=True)

df_MAP_bytask=df_MAP_bytask.join(df_BEF_Map_r, how='left',lsuffix='_df1', rsuffix='_df2')
df_MAP_bytask=df_MAP_bytask.rename_axis("MAP").reset_index()

#df_MAP_bytask=pd.merge(df_MAP_bytask,df_BEF_Map[["MAP","BEF"]],left_on="MAP", right_on="MAP",how="left")

df_MAP_byMAP=df_MAP_bytaskAll.groupby(["Documento Acquisto", "Posizione", "MAP", "Cod.Accett.MAP", "Mese competenza", "Anno competenza","Tipo ordine acquisto","Cod.Risorsa"]).agg(
MAP_Cons= ("ValoreMAP","sum"),Qta= ("Qta","sum")).reset_index()
df_MAP_byMAP["BDOLin"]=df_MAP_byMAP["Documento Acquisto"]+"-"+df_MAP_byMAP["Posizione"].str.zfill(3)
df_MAP_byMAP["Periodo"]=df_MAP_byMAP["Anno competenza"]+"-"+df_MAP_byMAP["Mese competenza"].str.zfill(2) 
df_MAP_byMAP["BDOLin_periodo"]=df_MAP_byMAP["BDOLin"]+"_"+df_MAP_byMAP["Periodo"] 


df_MAP_byMAP.set_index("MAP",inplace=True)
#from MAPbyMAP table get BEF
df_MAP_byMAP=df_MAP_byMAP.join(df_BEF_Map_r, how='left',lsuffix='_df1', rsuffix='_df2')
df_MAP_byMAP=df_MAP_byMAP.rename_axis("MAP").reset_index()


#df_MAP_byMAP=pd.merge(df_MAP_byMAP,df_BEF_Map[["MAP","BEF"]],left_on="MAP", right_on="MAP",how="left")

df_MAP_noverb=df_MAP_byMAP.loc[df_MAP_byMAP["Tipo ordine acquisto"].values=="Z4"]

df_MAP_noverb=pd.merge(df_MAP_noverb,df_BEF_Map,left_on="MAP", right_on="MAP",how="outer",indicator=True,suffixes=('', 'a'))
#from MAPbyMAP table get BEF
#df_MAP_noverb=df_MAP_noverb.join(df_BEF_Map, how='outer')

df_MAP_noverb=df_MAP_noverb.loc[df_MAP_noverb['_merge']=='left_only']
#Groupby per max min sum 

df_MAP_noverb_byLin=df_MAP_noverb.groupby(["Documento Acquisto", "Posizione","BDOLin","Periodo"]).agg(MAP_noverb=("MAP_Cons",np.sum),Qta=("Qta",np.sum)).reset_index()
df_MAP_noverb_byLin["BDOLin_periodo"]=df_MAP_noverb_byLin["BDOLin"]+"_"+df_MAP_noverb_byLin["Periodo"] 

wrenMAPColumns={"Documento d'acquisto":"Documento Acquisto","Numero MAP":"MAP","Conto Co.Ge.":"contabile"}
df_MAP_tot=df_MAP_tot.rename(columns=wrenMAPColumns)

logger.info('start reading df_PDC')
df_PDC=pd.read_excel(wrk_dir_inp+"\\"+"Da PDC_PortaleN.xlsx",sheet_name=1, dtype={"Numero BDO":str,"Posizione BDO":str,"Descrizione Posizione":str,"Codice PDC":str,"Periodo PDC": str,"Utente caricamento":str,"Codifica numerica documento":str,"Stato della PDC":str,"Divisione":str,"Centro di Costo":str,"Fornitore RTI":str,"ROI":str,"Fornitore Prestazione":str,"Tipo fornitura":str,"RDI":str,"Posizione RDI":str,"Subappalto":str,"Subappaltatore":str,"Descrizione Prestazione":str,"UOM":str,"Numero milestone":str,"Costo subappalto": float,"Valore netto prestazione": float,"Quantità richiesta":float,"Prezzo lordo":float,"Quantità consegnata":float,"Importo avanzamento":float,"PDC quantità da consegnare":float })
#Convert the Period Column
df_PDC["Anno competenza"]=df_PDC['Periodo PDC'].str.slice(3,7)
df_PDC["Mese competenza"]=df_PDC['Periodo PDC'].str[0:2]
#df_PDC['Periodo PDC']=df_PDC['Periodo PDC'].str.slice(3,7)+"-"+df_PDC['Periodo PDC'].str[0:2]
wColselPDC=["Numero BDO", "Posizione BDO", "Descrizione Posizione", "Fornitore RTI", "ROI", "Data invio ROI", "Tipo fornitura", "Importo avanzamento", "Anno competenza", "Mese competenza"]
#Filter Relevant columns
df_PDC=df_PDC[wColselPDC]
df_PDC["Periodo"]=df_PDC["Anno competenza"]+"-"+df_PDC["Mese competenza"].str.zfill(2) 
df_PDC["BDOLin"]=df_PDC["Numero BDO"]+"-"+df_PDC["Posizione BDO"].str.zfill(3) 
df_PDC["BDOLin_periodo"]=df_PDC["BDOLin"]+"_"+df_PDC["Periodo"] 
df_PDC=df_PDC.rename(columns={"Importo avanzamento":"ValorePDC","Numero BDO":"BDO"})
df_PDC_byLin=df_PDC.groupby(["BDO", "BDOLin", "Periodo", "BDOLin_periodo"]).agg(
ValorePDC= ("ValorePDC","sum")).reset_index()

logger.info('end reading df_PDC ' + str(df_PDC.shape[0]))
tempoFin = timeit.default_timer()

print("Durata Attivita *** ",tempoFin-tempoInz)
pass
#--------------End