##PMREP116 - Report Progetti e ciclo passivo per strutture
from funct.Dframe_templSh import write_templShname
from Util_leggeDir import dir_dict
import locale
from read_rilasci import  wIncarichi_list_act 

locale.setlocale(locale.LC_ALL, 'it_IT')
#(f_reptname,f_lib_dframe,f_wrk_dframe):

def RilRL000_rpt_gen():
	#wTest=False
#	wPrefixFileOut="PMREP240_VerbAttivi"
	#wNomeFileOut=wPrefixFileOut+"_ALL" 
	import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	

	import pandas as pd
# soluzione per scrivere Excel 
	import xlsxwriter
#from read_rilasci import df_Rilasci_task, df_Rilasci_subt
#from funct.Dframe_toSheet import scriveFoglio 

	from funct.Dframe_toSheet import scriveFoglio,scriveFoglio_dict 
	from read_rilasci import wIncarichi_list_act,wRilasci_list_act
	from read_verbAtt_v2 import df_elab_RilPPA, df_elab_Imp_Spalm, df_elab_Verb_Spalm, df_elab_Verb_Fatt, df_elab_RefRil, df_elab_Rilasci,df_elab_Incarichi, df_RCA_Verb, df_elab_Verbali
	#Rep116prep_sub(wTest,"AllDip","AllRept",wNomeFileOut)
	wrk_CsvDir=dir_dict["csvDir"]
	wFmtFile="RptFormati"

#print (dir_dict["rptDir"])
	wrk_RptDir=dir_dict["rptDir"]
	wrk_CsvDir=dir_dict["xlsDir"]
	wrk_stagDir=dir_dict["stagDir"]
#"df_MAP_bytask":"read_map",

	f_rptDir= wrk_RptDir

##  read Incarichi

	f_reptname="Incarichi" #"PMRep117-test"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

	writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book

	df_elab_Incarichi["Dt fine"]=pd.to_datetime(df_elab_Incarichi["Dt fine"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
	df_elab_Incarichi["Dt inizio"]=pd.to_datetime(df_elab_Incarichi["Dt inizio"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')

	df_elab_Incarichi=df_elab_Incarichi.loc[df_elab_Incarichi["cod incarico"].isin(wIncarichi_list_act)]
	#wRilasci_list_act=df_elab_Incarichi["Cod Rilascio"].unique() 
	#df_RefInc_cln=df_elab_Incarichi.loc[df_elab_Incarichi["cod incarico"].isin(wIncarichi_list_act)]

	f_descript="Riepilogo Incarichi per RL" #*** Personalizzare report (punto 1)
	f_reptcode="Incarichi_RL" #*** Personalizzare report (punto 2)
	f_sheetname="Foglio1" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )

#Create parameter to write sheets
	rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
	scriveFoglio_dict(writer,workbk_Out,df_elab_Incarichi,rept_dict) 		#*** Personalizzare selezione report (punto 7)

	writer.close()


##  read Rilasci

	f_reptname="Rilasci" #"PMRep117-test"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

	writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book

	f_descript="Riepilogo Rilasci per RL" #*** Personalizzare report (punto 1)
	f_reptcode="Rilasci_RL" #*** Personalizzare report (punto 2)	
	f_sheetname="Foglio1" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
#Exclude rilasci dtFine not null
	df_elab_Rilasci=df_elab_Rilasci.loc[df_elab_Rilasci["Dt fine"].notnull()]
	df_elab_Rilasci["Dt fine"]=pd.to_datetime(df_elab_Rilasci["Dt fine"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
	df_elab_Rilasci["Dt inizio"]=pd.to_datetime(df_elab_Rilasci["Dt inizio"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
	df_elab_Rilasci["Dt validazione"]=pd.to_datetime(df_elab_Rilasci["Dt validazione"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')

	df_elab_Rilasci=df_elab_Rilasci.loc[df_elab_Rilasci["Cod Rilascio"].isin(wRilasci_list_act)]

#Create parameter to write sheets
	rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
	scriveFoglio_dict(writer,workbk_Out,df_elab_Rilasci,rept_dict) 		#*** Personalizzare selezione report (punto 7)

	writer.close()

##  read tabFinanziamenti
##
	f_reptname="tabFinanziamenti + PPA" #"PMRep117-test"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

	writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book

	f_descript="Finaziamenti e PPA" #*** Personalizzare report (punto 1)
	f_reptcode="tabFinanziamenti + PPA_RL" #*** Personalizzare report (punto 2)
	f_sheetname="Foglio1" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
	df_elab_RilPPA=df_elab_RilPPA.loc[df_elab_RilPPA["Codice Rilascio"].isin(wRilasci_list_act)]

	wDate="Data Rilascio";df_elab_RilPPA[wDate]=pd.to_datetime(df_elab_RilPPA[wDate].astype(str), format='%Y-%m-%d',errors='coerce');df_elab_RilPPA[wDate]=df_elab_RilPPA[wDate].dt.strftime("%d-%b-%Y")
	wDate="Data Validazione";df_elab_RilPPA[wDate]=pd.to_datetime(df_elab_RilPPA[wDate].astype(str), format='%Y-%m-%d',errors='coerce');df_elab_RilPPA[wDate]=df_elab_RilPPA[wDate].dt.strftime("%d-%b-%Y")
#Create parameter to write sheets
	rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
	scriveFoglio_dict(writer,workbk_Out,df_elab_RilPPA,rept_dict) 		#*** Personalizzare selezione report (punto 7)

	writer.close()


##  read Spalmatura Impegni
##
	f_reptname="SpalmaturaImpegni" #"PMRep117-test"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

	writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book

	f_descript="Spalmatura Impegni RL" #*** Personalizzare report (punto 1)
	f_reptcode="SpalmaturaImpegni_RL" #*** Personalizzare report (punto 2)
	f_sheetname="Foglio1" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
	df_elab_Imp_Spalm=df_elab_Imp_Spalm.loc[df_elab_Imp_Spalm["cod rilascio"].isin(wRilasci_list_act)]
#Create parameter to write sheets
	rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
	scriveFoglio_dict(writer,workbk_Out,df_elab_Imp_Spalm,rept_dict) 		#*** Personalizzare selezione report (punto 7)

	writer.close()


##  read Spalmatura Verbali
##
	f_reptname="SpalmaturaVerbali" #"PMRep117-test"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

	writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book

	f_descript="Spalmatura Verbali RL" #*** Personalizzare report (punto 1)
	f_reptcode="SpalmaturaVerbali_RL" #*** Personalizzare report (punto 2)
	f_sheetname="Foglio1" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
#	df_elab_Verb_Spalm["mensilita"]=pd.to_datetime(df_elab_Verb_Spalm["mensilita"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
	df_elab_Verb_Spalm=df_elab_Verb_Spalm.loc[df_elab_Verb_Spalm["cod rilascio"].isin(wRilasci_list_act)]
#Create parameter to write sheets
	rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
	scriveFoglio_dict(writer,workbk_Out,df_elab_Verb_Spalm,rept_dict) 		#*** Personalizzare selezione report (punto 7)

	writer.close()



##  read Fatture
##
	f_reptname="Fatture" #"PMRep117-test"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

	writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book

	f_descript="Fatture RL" #*** Personalizzare report (punto 1)
	f_reptcode="Fatture_RL" #*** Personalizzare report (punto 2)
	f_sheetname="Foglio1" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )

	df_elab_Verb_Fatt["Dt Fattura"]=pd.to_datetime(df_elab_Verb_Fatt["Dt Fattura"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
#	df_elab_Verb_Fatt["Mensilità"]=pd.to_datetime(df_elab_Verb_Fatt["Mensilità"].astype(str), format='%d-%m-%Y %H:%M:%S',errors='coerce').dt.strftime('%m/%Y')

	df_elab_Verb_Fatt['Imp fattura'] = df_elab_Verb_Fatt['Imp fattura']*1.22
	df_elab_Verb_Fatt= df_elab_Verb_Fatt.round({'Imp fattura':2})
	df_elab_Verb_Fatt=df_elab_Verb_Fatt.loc[df_elab_Verb_Fatt["Cod Rilascio"].isin(wRilasci_list_act)]

#Create parameter to write sheets
	rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
	scriveFoglio_dict(writer,workbk_Out,df_elab_Verb_Fatt,rept_dict) 		#*** Personalizzare selezione report (punto 7)

	writer.close()


##  read Referenti
##
	f_reptname="Referenti" #"PMRep117-test"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

	writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book

	f_descript="Referenti RL" #*** Personalizzare report (punto 1)
	f_reptcode="Referenti_RL" #*** Personalizzare report (punto 2)
	f_sheetname="Foglio1" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
	df_elab_RefRil=df_elab_RefRil.loc[df_elab_RefRil["cod incarico"].isin(wIncarichi_list_act)]
#Create parameter to write sheets
	rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
	scriveFoglio_dict(writer,workbk_Out,df_elab_RefRil,rept_dict) 		#*** Personalizzare selezione report (punto 7)

	writer.close()

##  read Referenti
##
	f_reptname="Verbali" #"PMRep117-test"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

	writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book

	f_descript="Verbali RL" #*** Personalizzare report (punto 1)
	f_reptcode="Verbali_RL" #*** Personalizzare report (punto 2)
	f_sheetname="Foglio1" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )

#Convert date format
	df_elab_Verbali["dt creazione"]=pd.to_datetime(df_elab_Verbali["dt creazione"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
	df_elab_Verbali["dt firma"]=pd.to_datetime(df_elab_Verbali["dt firma"].astype(str), format='%Y-%m-%d',errors='coerce').dt.strftime('%d-%b-%Y')
	df_elab_Verbali["dt chiusura"]=pd.to_datetime(df_elab_Verbali["dt chiusura"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
#	df_RCA_Verb["mensilita"]=pd.to_datetime(df_RCA_Verb["mensilita"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
#	df_RCA_Verb["dal"]=pd.to_datetime(df_RCA_Verb["dal"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
#	df_RCA_Verb["al"]=pd.to_datetime(df_RCA_Verb["al"].astype(str), format='%Y-%m-%d %H:%M:%S',errors='coerce').dt.strftime('%d-%b-%Y')
#Create parameter to write sheets
	df_elab_Verbali=df_elab_Verbali.loc[df_elab_Verbali["cod rilascio"].isin(wRilasci_list_act)]
	rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000","rpt_dftTitle":f_descript,"rpt_strRow":0,"rpt_strCol":0,"rpt_environm":"Develop" }
	scriveFoglio_dict(writer,workbk_Out,df_elab_Verbali,rept_dict) 		#*** Personalizzare selezione report (punto 7)

	writer.close()




#call function from main module
if __name__ == "__main__":
	wlist=["DIH","ATQM"]
	wDict={
#"parm1":"value1",
#"parm2":"value2",
#"parm3":"value3"
}
	RilRL000_rpt_gen()
	pass