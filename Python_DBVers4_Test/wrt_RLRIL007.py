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
#from read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 
from read_Interni import df_RisInt_byName
wrk_environm=value = parm_dict.get("parm_environment", "Develop")
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
#from read_Interni import df_RisInt_byCID,df_RisInt_byName
import os
import sys
from elab_monitVerb import df_RilIncarichi
from elab_taskMAP import df_MAP_bytask_fatt

from read_rilasci import df_Rilasci_subt  #df_RilasciPPA_sub, 
#
from read_verbAtt_v2 import df_elab_RilPPA, df_elab_Imp_Spalm, df_elab_Verb_Spalm, df_elab_Verb_Fatt, df_elab_RefRil,df_RilInc, df_elab_Rilasci,df_elab_Incarichi, df_RCA_Verb, df_elab_Verbali, df_Incarichi,df_RefInc,df_RCA_Verb_FirmatoG,df_RCA_Verb_ChiusoG, df_RilasciPPA,df_RCA_Azioni
wrk_RptDir=dir_dict["rptDir"]
import time
wrk_VAT=1+.22
wDtFormat = "%d/%m/%Y" 
wTimestr = time.strftime("%Y%m%d")

#print (dir_dict["rptDir"])
wrk_RptDir=dir_dict["rptDir"]
wrk_CsvDir=dir_dict["xlsDir"]
wrk_stagDir=dir_dict["stagDir"]
#"df_MAP_bytask":"read_map",

f_rptDir= wrk_RptDir

##  read Incarichi

f_reptname="RLRIL007_sintesiRilasci" #"PMRep117-test"
wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"  
wRept_fulloutput=f_rptDir+"\\full"+f_reptname+".xlsx"
df_MAP_bytask_fatt=df_MAP_bytask_fatt.rename(columns={"Cod_Rilascio":"cod rilascio" })
df_Rilasci_subt=df_Rilasci_subt.rename(columns={"Cod_Rilascio":"cod rilascio" })
#Unisce le informazioni del dataframe df_MAP_bytask_fatt (da raggruppare per rilascio) e df_RilIncarichi
df_elab_Rilasci=df_elab_Rilasci.rename(columns={"Cod Rilascio":"cod rilascio","tot_Firmato":"RL_verbFirm","tot_Chiuso":"RL_verbChiusi","CodiceCup":"CUP" })

#df_MAP_bytask_fatt["Mappato"]= np.where(df_MAP_bytask_fatt["Data doc. Fattura"].isnull(),df_MAP_bytask_fatt["MAP_Cons"],0 )
df_MAP_bytask_fatt["Forn_Fatt"]= np.where(df_MAP_bytask_fatt["Data doc. Fattura"].notnull(),df_MAP_bytask_fatt["MAP_Cons"],0 )
df_MAP_bytask_fatt["Forn_Pagato"]= np.where(df_MAP_bytask_fatt["Data di pagamento"].notnull(),df_MAP_bytask_fatt["MAP_Cons"],0 )

df_elab_Verbali["VerbInBozza"]=np.where(df_elab_Verbali["stato verbale"].isin(["In bozza","In negoziazione"]),df_elab_Verbali["importo verbale"],0 )
#df_elab_Verbali["VerbInBozza"]=np.where(df_elab_Verbali["stato verbale"]=="In Firma Ente",df_elab_Verbali["importo verbale"],0 )
df_elab_Verbali["VerbFirmaARIA"]=np.where(df_elab_Verbali["stato verbale"]=="In Firma ARIA",df_elab_Verbali["importo verbale"],0 )
df_elab_Verbali["VerbFirmaRL"]=np.where(df_elab_Verbali["stato verbale"].isin(["In Firma ENTE","Approvazione Ente"]),df_elab_Verbali["importo verbale"],0 )
df_elab_Verbali["VerbFirmato+Chiuso"]=np.where(df_elab_Verbali["stato verbale"].isin(["Chiuso","Firmato"]),df_elab_Verbali["importo verbale"],0 )

#Raccoglie le linee del dataframe df_MAP_bytask_fatt ( totalizzato per rilascio)
df_attestato_byRil=df_MAP_bytask_fatt.groupby(['cod rilascio']).agg(Mappato = ("MAP_Cons","sum"), FatturatoF = ("Forn_Fatt","sum"), PagatoF = ("Forn_Pagato","sum")
     ).reset_index()

#Raccoglie le linee del dataframe df_MAP_bytask_fatt ( totalizzato per rilascio)
df_elab_Verbali_byRil=df_elab_Verbali.groupby(['cod rilascio']).agg(VerbInBozza = ("VerbInBozza","sum"), VerbFirmaARIA = ("VerbFirmaARIA","sum"), VerbFirmaRL = ("VerbFirmaRL","sum"), VerbChiuso = ("VerbFirmato+Chiuso","sum")
     ).reset_index()

df_elab_Verb_Fatt=df_elab_Verb_Fatt.rename(columns={"Cod Rilascio":"cod rilascio" })

#Raccoglie le linee del dataframe df_MAP_bytask_fatt ( totalizzato per rilascio)
df_elab_Verb_Fatt_byRil=df_elab_Verb_Fatt.groupby(['cod rilascio']).agg(FattAttivo = ("Imp fattura","sum"), impFtAttLiq = ("fattura liquidata","sum")
     ).reset_index()

df_RilIncarichi=pd.merge(df_RilIncarichi,df_attestato_byRil,on="cod rilascio", how="left",suffixes=('', '_sal'))
df_RilIncarichi=pd.merge(df_RilIncarichi,df_elab_Verbali_byRil,on="cod rilascio", how="left",suffixes=('', '_verb'))
df_RilIncarichi=pd.merge(df_RilIncarichi,df_Rilasci_subt[["LOB","BDGCostiCorrente","BDGCostiContratt","ConsTotali","resp_ril","resp_dmd","resp_ril_dip","resp_dmd_dip","cod rilascio"]],on="cod rilascio", how="left",suffixes=('', '_ril'))
df_RilIncarichi=pd.merge(df_RilIncarichi,df_elab_Verb_Fatt_byRil,on="cod rilascio", how="left",suffixes=('', '_ril'))

wList_date=["FattAttivo"	,"impFtAttLiq"]
#Date format
for d in wList_date:
    df_RilIncarichi[d+"+Iva"]=df_RilIncarichi[d]*wrk_VAT


wListaEnti=["ALB - ALER BERGAMO","ALM - ALER MILANO","ALP - ALER PAVIA","ALR - ALER BRESCIA","ALV - ALER VARESE","ARPA - A.R.P.A.","CR - CONSIGLIO REGIONALE","ERSAF - ERSAF","POLIS - POLIS LOMBARDIA","SC - STRUTTURA COMMISSARIALE","SRH - FINLOMBARDA"]
df_RilIncarichi["TipoEnte"]=np.where(df_RilIncarichi["Cod DG RL"].isin(wListaEnti),"ENTE SIREG", np.where(df_RilIncarichi["Cod DG RL"].str.contains("ASST|ATS|IRCCS|POLICLINICO"),"ENTE SANITARIO",np.where(df_RilIncarichi["Cod DG RL"].str.contains("COMUNE DI"),"COMUNE",np.where(df_RilIncarichi["Cod DG RL"]=="G1 - WELFARE","RL - WELFARE","RL - ALTRE DG"))))
#wList_date=["Data I. Corrente","Data F. Corrente","Data I. Contrattuale","Data F. Contrattuale"]
#Date format
#for d in wList_date:
#    df_prog[d]=pd.to_datetime(df_prog[d],errors='coerce')


#    pd.to_datetime(df_RilIncarichi[d], dayfirst=True, format="%Y-%m-%d %H:%M:%S").dt.strftime('%d/%m/%Y')
#    df_RilIncarichi[d+"_"]=df_RilIncarichi[d].to_timestamp
#    df_RilIncarichi[d+"_"]=pd.to_datetime(df_RilIncarichi[d+"_"],format=wDtFormat,errors='coerce')

#	df_elab_Rilasci["Dt fine"]=pd.to_datetime(df_elab_Rilasci["Dt fine"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')


with pd.ExcelWriter(wRept_fulloutput) as writer:
    df_RilIncarichi.to_excel(writer,sheet_name="RilIncarichi", index=False, startrow=0, startcol=0)

writer.close()

writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
workbk_Out=writer.book

f_descript="Riepilogo Rilasci" #*** Personalizzare report (punto 1)
f_reptcode="Rilasci_AttPassivo" #*** Personalizzare report (punto 2)
f_sheetname="RilIncarichi" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )

#Create parameter to write sheets
rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":"Develop" }
scriveFoglio_dict(writer,workbk_Out,df_RilIncarichi,rept_dict) 		#*** Personalizzare selezione report (punto 7)

writer.close()



wDtFormat = "%d/%m/%Y"
wTimestr = time.strftime("%Y%m%d")

print("durata: "+ str(wTimestr))
pass

