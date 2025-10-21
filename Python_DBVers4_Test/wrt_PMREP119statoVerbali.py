#New Reportistica ---
###Codice da completare
#Estrae elenco dipendenti 
# #def EstraeRepGMRE(fTest,fListDip,fListRep,fNomeFileOut):
#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
from Util_leggeDir import dir_dict
from funct.Dframe_toSheet import scriveFoglio,scriveFoglio_dict 

#from Util_leggeParm import parm_dict
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
from elab_Bdo_Periodo import df_BDO_per
from read_verbPass import df_Verb_ROFI
from funct.Dframe_templSh import write_templShname
from Util_leggeDir import dir_dict
from read_ROFI import df_ROFI
from read_BdoOda import df_forn_byName
df_forn_byName["Forn_nome"]=df_forn_byName["Forn_nome"].str.upper()
df_forn_byName  = df_forn_byName.drop_duplicates( subset = ["Forn_nome"],  keep = "first" )

wrk_RptDir=dir_dict["rptDir"]

wColumns=["ROI","ROI_dip",	"Fornitore RTI","StatoBDO",	"BDO"	,"Periodo"	,"Stat_VerbAt","Stat_VerbCh",	"Stat_VerbSAL","DEC","RUP","Fornitori","Subfornitori","Numero Contratto Gara"]
df_BDO_Verb=df_BDO_per[wColumns]

df_BDO_Verb['Stat_VerbAt'] = df_BDO_Verb['Stat_VerbAt'].str.upper()
df_BDO_Verb['Stat_VerbCh'] = df_BDO_Verb['Stat_VerbCh'].str.upper()
df_BDO_Verb['Stat_VerbSAL'] = df_BDO_Verb['Stat_VerbSAL'].str.upper()

wFirmaROI=["IN FIRMA ROI"]
wFirmaROFI=["IN FIRMA ROFI"]
wAssQuiet=["ASSOCIAZIONE QUIETANZE"]
wFirmaDEC=["IN APPROVAZIONE DEC"]
wFirmaRUP=["IN APPROVAZIONE RUP"]
wFirmaForn=["IN FIRMA ROFI"]

wFirmaRich=["ASSOCIAZIONE QUIETANZE","IN FIRMA ROFI","IN APPROVAZIONE DEC","IN APPROVAZIONE RUP","IN FIRMA ROI"]

df_BDO_VerbSAL=df_BDO_Verb.loc[df_BDO_Verb["Stat_VerbSAL"].isin(wFirmaRich)]
#Carica elenco verbali in una lista
wlist_BDOSAL=df_BDO_VerbSAL['BDO'].tolist()

#20251010 predispone elenco Verbali Unici di attivazione e di chiusura da completare (vengono tolte le casistiche riferite alle evenienze di Verbali di SAL non completati)
df_BDO_VerbUnici=df_BDO_Verb.loc[df_BDO_Verb["Stat_VerbAt"].isin(wFirmaRich)|((df_BDO_Verb["Stat_VerbCh"].isin(wFirmaRich))&(~df_BDO_Verb["BDO"].isin(wlist_BDOSAL)))]
df_BDO_VerbUnici=df_BDO_VerbUnici.drop_duplicates(subset='BDO', keep="last")


#20251010 Rimuovi stato verbale di chiusura se esistono ancora verbali di SAL da completare
#df_BDO_VerbUnici=df_BDO_VerbUnici.merge(df_BDO_VerbSAL.drop_duplicates(), on=["BDO"], how="left",indicator=True )
#df_BDO_VerbUnici=df_BDO_VerbUnici[df_BDO_VerbUnici['_merge'] == 'left_only']

df_BDO_VerbSAL=pd.concat([df_BDO_VerbSAL, df_BDO_VerbUnici])

df_BDO_VerbSAL=pd.merge(df_BDO_VerbSAL,df_ROFI[["BDO","RAGIONE_SOCIALE","ROFI"]],left_on="BDO", right_on="BDO",how="left",suffixes=('', '_rofi') )

df_BDO_VerbSAL=pd.merge(df_BDO_VerbSAL,df_forn_byName,left_on="RAGIONE_SOCIALE", right_on="Forn_nome",how="left",suffixes=('', '_rofi') )

df_BDO_VerbSAL=pd.merge(df_BDO_VerbSAL,df_Verb_ROFI,left_on="BDO", right_on="Numero BDO",how="left",suffixes=('', '_rofi') )

#df_BDO_VerbSAL=pd.merge(df_BDO_VerbSAL,df_Verb_ROFI,left_on="BDO", right_on="Numero BDO",how="left",suffixes=('', '_task') )

selROI=(df_BDO_VerbSAL["Stat_VerbAt"].isin(wFirmaROI))|(df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaROI))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wFirmaROI))
selDEC=(df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaDEC))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wFirmaDEC))
selRUP=(df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaRUP))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wFirmaRUP))
selROFI= (df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaROFI))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wFirmaROFI))
selAssQ=(df_BDO_VerbSAL["Stat_VerbCh"].isin(wAssQuiet))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wAssQuiet))
df_BDO_VerbSAL.loc[selROI,"In approvaz.(Aria)"]=df_BDO_VerbSAL.loc[selROI,"ROI"]
df_BDO_VerbSAL.loc[selDEC,"In approvaz.(Aria)"]=df_BDO_VerbSAL.loc[selDEC,"DEC"]
df_BDO_VerbSAL.loc[selRUP,"In approvaz.(Aria)"]=df_BDO_VerbSAL.loc[selDEC,"RUP"]

#labels=[df_BDO_VerbSAL["ROI"],df_BDO_VerbSAL["DEC"],df_BDO_VerbSAL["RUP"]]
#labels=[df_BDO_VerbSAL["ROI"],df_BDO_VerbSAL["DEC"],df_BDO_VerbSAL["RUP"]]

df_BDO_VerbSAL.loc[selROFI,"ApprovForn"]=df_BDO_VerbSAL.loc[selROFI,"ROFI"]
#df_BDO_VerbSAL.loc[df_BDO_VerbSAL[m1],"ApprovAria"]=df_BDO_VerbSAL.loc[df_BDO_VerbSAL[m1],"ROI"]
#df_BDO_VerbSAL.loc[df_BDO_VerbSAL[m2],"ApprovAria"]=df_BDO_VerbSAL.loc[df_BDO_VerbSAL[m2],"DEC"]
#f_BDO_VerbSAL.loc[df_BDO_VerbSAL[m3],"ApprovAria"]=df_BDO_VerbSAL.loc[df_BDO_VerbSAL[m3],"RUP"]
#f_BDO_VerbSAL['ApprovAria']=np.select([m1,m2,m3],labels)

#m1=(df_BDO_VerbSAL["Stat_VerbAt"].isin(wFirmaForn))|(df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaForn))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wFirmaForn))

#labels=[df_BDO_VerbSAL["ROFI"]]

#df_BDO_VerbSAL['ApprovForm']=np.select([m1],labels)

#df_BDO_VerbSAL['ApprovAria']=np.where(((df_BDO_VerbSAL["Stat_VerbAt"].isin(wFirmaROI))|(df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaROI))|(df_BDO_per["Stat_VerbSAL"].isin(wFirmaROI)),
#                    df_BDO_VerbSAL["ROI"]),np.where(((df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaDEC))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wFirmaDEC)),
#                    df_BDO_VerbSAL["DEC"]),np.where(((df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaRUP))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wFirmaRUP)),
#                    df_BDO_VerbSAL["RUP"]),"" )))
#df_BDO_VerbSAL['ApprovForn']=np.where((df_BDO_VerbSAL["Stat_VerbAt"].isin(wFirmaForn))|(df_BDO_VerbSAL["Stat_VerbCh"].isin(wFirmaForn))|(df_BDO_VerbSAL["Stat_VerbSAL"].isin(wFirmaForn),
#                    df_BDO_VerbSAL["ROFI"],"" ))
df_BDO_VerbSAL=pd.merge(df_BDO_VerbSAL,df_RisInt_byName[["Cognome Nome","PMO","Int_Dip","E-mail"]],left_on="In approvaz.(Aria)", right_on="Cognome Nome",how="left",suffixes=('', '_task') )


df_BDO_VerbSAL=pd.merge(df_BDO_VerbSAL,df_RisInt_byCID[["CID","E-mail"]],left_on="PMO", right_on="CID",how="left",suffixes=('', '_pmo') )

f_reptname="PMREP119 VerbApprov" #"PMRep117-test"
wRept_output=wrk_RptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

df_BDO_VerbSAL["ODAGforn"]=df_BDO_VerbSAL["Numero Contratto Gara"]+"_"+df_BDO_VerbSAL["Forn_cod"]


writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
workbk_Out=writer.book

df_BDO_VerbSAL_sel=df_BDO_VerbSAL.loc[df_BDO_VerbSAL['In approvaz.(Aria)'].notnull()]

f_descript="Riepilogo Verbali passivi in approvazione Aria" #*** Personalizzare report (punto 1)
f_reptcode="VerbSAL_Approv" #*** Personalizzare report (punto 2)	
f_sheetname="VerbSAL_ApprovARIA" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
scriveFoglio_dict(writer,workbk_Out,df_BDO_VerbSAL_sel,rept_dict) 		#*** Personalizzare selezione report (punto 7)


df_BDO_VerbSAL_sel=df_BDO_VerbSAL.loc[df_BDO_VerbSAL['ApprovForn'].notnull()]



f_descript="Riepilogo Verbali passivi in approvazione ROFI" #*** Personalizzare report (punto 1)
f_reptcode="VerbSAL_Approv" #*** Personalizzare report (punto 2)	
f_sheetname="VerbSAL_ApprovForn" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
scriveFoglio_dict(writer,workbk_Out,df_BDO_VerbSAL_sel,rept_dict) 		#*** Personalizzare selezione report (punto 7)

df_BDO_VerbSAL_sel=df_BDO_VerbSAL.loc[selAssQ]

f_descript="Riepilogo Verbali passivi in associazione quietanza" #*** Personalizzare report (punto 1)
f_reptcode="VerbSAL_Approv" #*** Personalizzare report (punto 2)	
f_sheetname="VerbSAL_AssQuiet" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
scriveFoglio_dict(writer,workbk_Out,df_BDO_VerbSAL_sel,rept_dict) 		#*** Personalizzare selezione report (punto 7)



writer.close()