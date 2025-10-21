##PMREP116 - Report Progetti e ciclo passivo per strutture
#from funct.Dframe_templSh import write_templShname
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task
import os
from Util_leggeDir import dir_dict
wprogzero_fmyear= dir_dict.get("parm_progzerovalue_fmyear", 2025) #Get parm from parameters in RptFormati.xlsx
wprog_fmyear= dir_dict.get('parm_prog_fmyear',2024) #Get parm from parameters in RptFormati.xlsx
wbdo_lastdays= dir_dict.get('parm_lastBDO_days',365) #Get parm from parameters in RptFormati.xlsx
from Util_logging import logger
#wPythProc="read_verbPass"
import timeit  #funzione timeit per valutare durata estrazione
#tempoInz=timeit.default_timer()

wPythPro="wrt_PMREP116_"
#(f_reptname,f_lib_dframe,f_wrk_dframe):

def Rep116_rpt_gen(fListDip,fDict):
	#wTest=False
	wPrefixFileOut="PMREP116_BDOPrgnw_byDipv5"
	wNomeFileOut=wPrefixFileOut+"_ALL" 
	
	#Rep116prep_sub(wTest,"AllDip","AllRept",wNomeFileOut)
	for wDip in fListDip:
		if wDip == "AllDip":
			wNomeFileOut=wPrefixFileOut+"_ALL" 
			Rep116_rpt_sub(wDip,wNomeFileOut,fDict)		
		else:
			wNomeFileOut=wPrefixFileOut+"_"+wDip
			Rep116_rpt_sub(wDip,wNomeFileOut,fDict)

def Rep116_rpt_sub(sListDip,sNomeFileOut,sDict):
	pass
#	from Util_logging import logger
#Read dummy tables
#rom read_dummy import df_MAP_dummy,df_BDO_dummy,df_ODAG_dummy,df_Task_dummy 

	wPythProc=wPythPro+"_"+sListDip

	tempoInz=timeit.default_timer()	
	#status create / modify read-only file 
	from stat import S_IREAD, S_IRGRP, S_IROTH, S_IWUSR, S_IWRITE, S_IWOTH, S_IRWXU 
	import contextlib
	wDebugsts="Y"
	wEnvironment="dev"

#2)	Gestione DIRECTORY CORRENTE (contestualizzazione cartella python sulla base dell’utente)
#da inserire -> selezione directory sulla base dell'utente che utilizza Python
	current_dir=os.getcwd()
	print(current_dir)
	#logging.config.fileConfig('logging2.conf')

# create logger
#	logger = logging.getLogger('simpleExample')
	import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	

	import pandas as pd
# soluzione per scrivere Excel 
	import xlsxwriter
#from read_rilasci import df_Rilasci_task, df_Rilasci_subt
#from funct.Dframe_toSheet import scriveFoglio 

	from funct.Dframe_toSheet import scriveFoglio_dict_v5 
	from elab_ril_verb import df_Rilasci_subt
	from read_verbPass import wDict_Verb
	from elab_Bdo_pivot import df_BDOnoCons_byDip,df_BDOnoCons_byBDO,df_BDOnoVerb_byBDO,df_BDOVerb_recent_per,df_ODAnoCons_byDip,df_ODAnoCons_byROI
	from elab_Bdo_Periodo  import df_BDO_recent_per
	wDtFormat = "%d/%m/%Y" 
	wrk_CsvDir=dir_dict["csvDir"]
	wFmtFile="RptFormati"

#print (dir_dict["rptDir"])
	wrk_RptDir=dir_dict["rptDir"]
	wrk_CsvDir=dir_dict["xlsDir"]
	wrk_stagDir=dir_dict["stagDir"]
#"df_MAP_bytask":"read_map",



	f_rptDir= wrk_RptDir
	f_reptname=sNomeFileOut #"PMRep117-test"
	if not sDict:
		sDict["AllRept"]="All reports"
	wRept_output=f_rptDir+"\\"+f_reptname+".xlsx"   

	#verifica se esiste il file e se esiste lo rimuove
	with contextlib.suppress(FileNotFoundError):
		os.chmod(wRept_output, S_IRWXU) #cambia le autorizzazioni sul file
		os.remove(wRept_output)

	writer= pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
	workbk_Out=writer.book
	
	#Report BDO_Cons
	if "BDO_Cons" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Report stato consuntivazioni BDO" 
		f_reptcode="BdO_Periodo"
		f_sheetname="BdO_Periodo"
	#	logger.info('start process for ' + 'f_Rilasci_subt' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"elab_Bdo_Periodo","df_BDO_per")
#		from elab_Bdo_pivot import df_BDOnoCons_byBDO

#		df_BDOnoCons_byBDO["Stato_VerbAt"]=df_BDOnoCons_byBDO["ID_VerbAt"].replace(wDict_Verb)
#		df_BDOnoCons_byBDO["Stato_VerbCh"]=df_BDOnoCons_byBDO["ID_VerbCh"].replace(wDict_Verb)
#		df_BDOnoCons_byBDO["Stato_VerbSAL"]=df_BDOnoCons_byBDO["ID_VerbSAL"].replace(wDict_Verb)
	
		if "AllDip" != sListDip: 
			df_BDO_per_dip=df_BDO_recent_per.loc[(df_BDO_recent_per["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_BDO_per_dip=df_BDO_recent_per
		#Filter if "Data decorrenza" is not in right format
#		df_BDO_per_dip=df_BDO_per_dip.loc[pd.to_datetime(df_BDO_per_dip['Data decorrenza'], errors='coerce',format='%Y-%m-%d').notnull()]
#		df_BDO_per_dip["Data decorrenza"]=(df_BDO_per_dip["Data decorrenza"]).dt.strftime(wDtFormat)
#		df_BDO_per_dip["Data scadenza"]=(df_BDO_per_dip["Data scadenza"]).dt.strftime(wDtFormat)
#select rows that requires activities
		df_BDO_per_dip_todo=df_BDO_per_dip.loc[(df_BDO_per_dip["Flag_daCons"])]	
#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_BDO_per_dip_todo,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_BDO_per_dip_todo.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_BDO_per_dip_todo.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDO_per_dip_todo.memory_usage(deep=True)) )


		#wDim= str(df_BDO_per_dip_todo.shape[0])   
		#logger.info('dFrame df_BDO_per_dip_todo' + 'dimension ' + wDim )

#Report BDO_Cons
	if "BDO_Cons" in sDict.keys() or "AllRept" in sDict.keys():
#------------------ New report 
		f_descript="Riepilogo BDO da consuntivare per Struttura e Periodo" 
		f_reptcode="RiepDip_BDO_daCons"
		f_sheetname="RiepDip_BDO_daCons"
		logger.info('start process for ' + 'f_Rilasci_subt' )
#Convert date format dd/mm/YYYY
		if "AllDip" != sListDip: 
			df_BDO_per_dip=df_BDOnoCons_byDip.loc[(df_BDOnoCons_byDip["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_BDO_per_dip=df_BDOnoCons_byDip
#		df_BDOnoCons_byDip=df_BDOnoCons_byDip
#		df_BDO_per_dip_grp=df_BDO_per_dip_todo.groupby(["ROI_dip","Anno competenza","Mese competenza" ],observed=True,as_index=False).agg({"DocAcq":"nunique"})
#		df_BDO_per_dip_grp["Periodo"]=df_BDO_per_dip_grp["Anno competenza"].astype(str)+"-"+df_BDO_per_dip_grp["Mese competenza"].astype(str).str.zfill(2)
#		df_col_Vpivot_BDO=df_BDO_per_dip_grp.pivot_table(index=["ROI_dip"], values=["DocAcq"] ,columns=["Periodo"],
#							   aggfunc="sum")
#25/09/2024		df_col_Vpivot_BDO=dfBDO_per_dip.pivot_table(index=['Dip'],  
#25/09/2024			values=['BDO'] ,
#25/09/2024			columns=['PeriodoMMM'],
#25/09/2024			aggfunc=pd.Series.nunique, margins=True, margins_name="#Totale#" )
#		df_col_Vpivot_BDO.columns = df_col_Vpivot_BDO.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
#		df_col_Vpivot_BDO=df_col_Vpivot_BDO.reset_index().rename_axis(None,axis=1)
#		df_col_Vpivot_BDO.loc["total"]=df_col_Vpivot_BDO.sum()
#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":1,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_BDOnoCons_byDip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
#		wDim= str(df_col_Vpivot_BDO.shape[0])   
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_BDOnoCons_byDip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_BDOnoCons_byDip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDOnoCons_byDip.memory_usage(deep=True)) )

#Report BDO_Cons_mensili
	if "BDO_Cons_mensili" in sDict.keys() or "AllRept" in sDict.keys():
#------------------ New report 
		f_descript="Prospetto BDO da consuntivare nei mesi per Struttura e ROI" 
		f_reptcode="RiepBDO_Cons_mensili"
		f_sheetname="RiepBDO_Cons_mensili"
		logger.info('start process for ' + 'RiepBDO_Cons_mensili' )


		if "AllDip" != sListDip: 
			df_BDO_per_dip=df_BDOnoCons_byBDO.loc[(df_BDOnoCons_byBDO["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_BDO_per_dip=df_BDOnoCons_byBDO

#		df_BDO_per_dip_todo["Stato_VerbAt"]=df_BDO_per_dip_todo["ID_VerbAt"].replace(wDict_Verb)

#		df_BDO_per_dip_grp=df_BDO_per_dip_todo.groupby(["ROI_dip","DocAcq","Anno competenza","Mese competenza","ROI","Descrizione DocAcq","StatoDocAcq","ID_VerbAt"],observed=True).agg(numbdo=("DocAcq","count")).reset_index()

#		df_BDO_per_dip_grp["Periodo"]=df_BDO_per_dip_grp["Anno competenza"].astype(str)+"-"+df_BDO_per_dip_grp["Mese competenza"].astype(str).str.zfill(2)#
#		df_BDO_per_dip_grp["Periodo"]=df_BDO_per_dip_grp["Periodo"].astype("category")

#		df_BDO_per_dip_grp = df_BDO_per_dip_grp.astype({"Periodo":"category"})

#		df_col_Vpivot_BDO=df_BDO_per_dip_grp.pivot_table(index=["ROI_dip","ROI","DocAcq","Descrizione DocAcq","StatoDocAcq","ID_VerbAt"], columns=["Periodo"],values=["numbdo"],
#							   aggfunc="sum")
#25/09/2024		df_col_Vpivot_BDO=dfBDO_per_dip.pivot_table(index=['Dip'],  
#25/09/2024			values=['BDO'] ,
#25/09/2024			columns=['PeriodoMMM'],
#25/09/2024			aggfunc=pd.Series.nunique, margins=True, margins_name="#Totale#" )
#		df_col_Vpivot_BDO.columns = df_col_Vpivot_BDO.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
#		df_col_Vpivot_BDO=df_col_Vpivot_BDO.reset_index().rename_axis(None,axis=1)

		df_BDO_per_dip["Stat_VerbAt"]=df_BDO_per_dip["ID_VerbAt"].replace(wDict_Verb)
#		df_col_Vpivot_BDO.loc["total"]=df_col_Vpivot_BDO.sum()
#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_BDO_per_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		wDim= str(df_BDO_per_dip.shape[0])   
		logger.info('dFrame df_col_Vpivot_BDO' + 'dimension ' + wDim )
#		df_col_Vpivot_BDO = df_col_Vpivot_BDO.reset_index()                #index to columns
	if "BDO_Verb" in sDict.keys() or "AllRept" in sDict.keys():
#------------------ New report 
		f_descript="Riepilogo MAP senza Verbali" 
		f_reptcode="Riep_Verbali"
		f_sheetname="Riep_Verbali"
		logger.info('start process for ' + 'Riep_Verbali' )

		if "AllDip" != sListDip: 
			df_BDOnoVerb_per_dip=df_BDOnoVerb_byBDO.loc[(df_BDOnoVerb_byBDO["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_BDOnoVerb_per_dip=df_BDOnoVerb_byBDO

		df_BDOnoVerb_per_dip["Stat_VerbAt"]=df_BDOnoVerb_per_dip["ID_VerbAt"].replace(wDict_Verb)

#		df_col_Vpivot_BDO.loc["total"]=df_col_Vpivot_BDO.sum()
#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_BDOnoVerb_per_dip,rept_dict)

		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_BDOnoVerb_per_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_BDOnoVerb_per_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDOnoVerb_per_dip.memory_usage(deep=True)) )

#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
#		wDim= str(df_col_Vpivot_BDO.shape[0])   
#		logger.info('dFrame df_col_Vpivot_BDO' + 'dimension ' + wDim )



#		df_col_Vpivot_BDO = df_col_Vpivot_BDO.reset_index()                #index to columns
#Riepilogo verbali
	if "BDO_Verb" in sDict.keys() or "AllRept" in sDict.keys():
#------------------ New report 
		f_descript="Riepilogo Verbali" 
		f_reptcode="Riep_StatoVerbali"
		f_sheetname="Riep_StatoVerbali"
		logger.info('start process for ' + 'Riep_StatoVerbali' )

		if "AllDip" != sListDip: 
			df_BDOVerb_per_dip=df_BDOVerb_recent_per.loc[(df_BDOVerb_recent_per["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_BDOVerb_per_dip=df_BDOVerb_recent_per

#		df_BDO_per_dip_grp=df_BDO_per_dip.groupby(
#			["ROI_dip","DocAcq","ODAG","Macrocl_nbr","Anno competenza","Mese competenza","ROI","Descrizione DocAcq","StatoDocAcq","Stat_VerbAt","DataFine","DataIniz","Stat_VerbCh","Fornitori","Subfornitori","Forn_RTI"],observed=True).agg(
#			Stat_VerbSAL=("ID_VerbSAL",list)
#    	).reset_index() #,MAP_Cons=("MAP_Cons","sum"),MAP_noverb=("MAP_noverb","sum")

#		df_BDOVerb_recent_per['Stat_VerbSAL']=[','.join(map(str, l)) for l in df_BDOVerb_recent_per['Stat_VerbSAL']]
#Stat_VerbSAL
#		df_BDO_per_dip_grp["Periodo"]=df_BDO_per_dip_grp["Anno competenza"].astype(str)+"-"+df_BDO_per_dip_grp["Mese competenza"].astype(str).str.zfill(2)
		#df_BDO_per_dip_grp=df_BDO_per_dip.groupby(["ROI_Dip","BDO","ODAGMacr_nbr","Periodo","ROI","Descrizione Bdo","StatoBDO","Stat_VerbAt","Data decorrenza","Data scadenza","Stat_VerbCh","MAP_noverb"],as_index=False).agg({"Documento d'acquisto":"nunique"})
#		df_col_Vpivot_BDO=df_BDO_per_dip_grp.pivot_table(index=["ROI_dip","ODAG","Macrocl_nbr","ROI","Descrizione DocAcq","StatoDocAcq","ID_VerbAt","DataFine","DocAcq","DataIniz","Stat_VerbCh","Fornitori","Subfornitori","Forn_RTI"], values=["Stat_VerbSAL"] ,columns=["Periodo"],
#							   aggfunc="max") #,"MAP_Cons","MAP_noverb"
#25/09/2024		df_col_Vpivot_BDO=dfBDO_per_dip.pivot_table(index=['Dip'],  
#25/09/2024			values=['BDO'] ,
#25/09/2024			columns=['PeriodoMMM'],
#25/09/2024			aggfunc=pd.Series.nunique, margins=True, margins_name="#Totale#" )
#		df_col_Vpivot_BDO.columns = df_col_Vpivot_BDO.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
#		df_col_Vpivot_BDO=df_col_Vpivot_BDO.reset_index().rename_axis(None,axis=1)


#		df_col_Vpivot_BDO.loc["total"]=df_col_Vpivot_BDO.sum()
#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_BDOVerb_per_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_BDOVerb_per_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_BDOVerb_per_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDOVerb_per_dip.memory_usage(deep=True)) )

#		df_col_Vpivot_BDO = df_col_Vpivot_BDO.reset_index()                #index to columns
#		df_col_Vpivot_BDO.loc[df_col_Vpivot_BDO.index[-1],"Dip"]="Totale"

#		Dettaglio Verbali
	if "BDO_Verb" in sDict.keys() or "AllRept" in sDict.keys():
#------------------ New report 
		f_descript="Dettaglio Verbali BDO" 
		f_reptcode="Dett_Verbali"
		f_sheetname="Dett_Verbali"
#		logger.info('start process for ' + 'Dett_Verbali' )

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_BDO_per_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
#		wDim= str(df_BDO_per_dip.shape[0])   
#		logger.info('df_BDO_per_dip' + 'dimension ' + wDim )

		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_BDO_per_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_BDO_per_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDO_per_dip.memory_usage(deep=True)) )


	#Report BDO_Cons
	if "List_BDO" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Report elenco BDO" 
		f_reptcode="List_BdO_Elenco"
		f_sheetname="List_BdO_Elenco"
		logger.info('start process for ' + 'List_BdO_Elenco' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"elab_Bdo_Periodo","df_BDO")
		from elab_Bdo_Periodo import df_BDO
	#estrae la data di oggi
#	now = pd.to_datetime('now')
#last2 = now - pd.DateOffset(months=2)

		if "AllDip" != sListDip: 
			df_BDO_dip=df_BDO[(df_BDO["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_BDO_dip=df_BDO
#202508		df_BDO_dip["Data scadenza"]=pd.to_datetime(df_BDO_dip["Data scadenza"])
#filtra per gli ultimi mesi
		df_BDO_dip=df_BDO_dip[(df_BDO_dip["DataFine"] >= pd.to_datetime('now')- pd.DateOffset(days=wbdo_lastdays))]
	
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	
#202508		df_BDO_dip["DataCreaz"]=pd.to_datetime(df_BDO_dip["Data Creazione Bdo"]);df_BDO_dip["Data Creazione Bdo"]=df_BDO_dip["Data Creazione Bdo"].dt.strftime(wDtFormat)
#		df_BDO["Data Approvazione"]=pd.to_datetime(df_BDO["Data Approvazione"]);df_BDO["Data Approvazione"]=df_BDO["Data Approvazione"].dt.strftime(wDtFormat)
#		df_BDO["Data decorrenza"]=df_BDO["Data decorrenza"].dt.strftime(wDtFormat)
#		df_BDO["Data scadenza"]=df_BDO["Data scadenza"].dt.strftime(wDtFormat)
#202508		df_BDO_dip["Data decorrenza"]=pd.to_datetime(df_BDO_dip["Data decorrenza"]);df_BDO_dip["Data decorrenza"]=df_BDO_dip["Data decorrenza"].dt.strftime(wDtFormat)
		#df_BDO_dip["Data scadenza"]=pd.to_datetime(df_BDO_dip["Data scadenza"]);
#202508		df_BDO_dip["Data scadenza"]=df_BDO_dip["Data scadenza"].dt.strftime(wDtFormat)
#202508		df_BDO_dip["DataEffAtt"]=pd.to_datetime(df_BDO_dip["DataEffAtt"]);df_BDO_dip["DataEffAtt"]=df_BDO_dip["DataEffAtt"].dt.strftime(wDtFormat)		
#		df_BDO_dip["DataEffAtt"]= np.where((df_BDO_dip["DataEffAtt"].isnull()),df_BDO_dip["Data decorrenza"],df_BDO_dip["DataEffAtt"] ) 

#		df_BDO["DataEffAtt"]=pd.to_datetime(df_BDO["DataEffAtt"]);df_BDO["DataEffAtt"]=df_BDO["DataEffAtt"].dt.strftime(wDtFormat)

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_BDO_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		wDim= str(df_BDO_dip.shape[0])   
		logger.info('dFrame df_BDO_dip' + 'dimension ' + wDim )

#Report BDO_Cons
	if "List_BDOLinee" in sDict.keys() or "AllRept" in sDict.keys():
#------------------ New report 
		f_descript="Dettaglio BDO Linee" 
		f_reptcode="Elenco_BDOLinee"
		f_sheetname="Elenco_BDOLinee"
		logger.info('start process for ' + 'Elenco_BDOLinee' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"elab_Bdo_Periodo","df_BDOLin")
		from elab_Bdo_Periodo import df_BDOLin


		if "AllDip" != sListDip: 
			df_BDOLin_dip=df_BDOLin.loc[(df_BDOLin["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_BDOLin_dip=df_BDOLin
#		df_BDOLin_dip["Data decorrenza"]=df_BDOLin_dip["Data decorrenza"].dt.strftime(wDtFormat)
#		df_BDOLin_dip["Data scadenza"]=df_BDOLin_dip["Data scadenza"].dt.strftime(wDtFormat)
#202508		df_BDOLin_dip["Data decorrenza"]=pd.to_datetime(df_BDOLin_dip["Data decorrenza"], format=wDtFormat);df_BDOLin_dip["Data decorrenza"]=df_BDOLin_dip["Data decorrenza"].dt.strftime(wDtFormat)
#202508	df_BDOLin_dip["Data scadenza"]=pd.to_datetime(df_BDOLin_dip["Data scadenza"], format=wDtFormat);df_BDOLin_dip["Data scadenza"]=df_BDOLin_dip["Data scadenza"].dt.strftime(wDtFormat)
#202508	df_BDOLin_dip["Data_MLS"]=pd.to_datetime(df_BDOLin_dip["Data_MLS"], format=wDtFormat);df_BDOLin_dip["Data_MLS"]=df_BDOLin_dip["Data_MLS"].dt.strftime(wDtFormat)

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_BDOLin_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_BDOLin_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_BDOLin_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDOLin_dip.memory_usage(deep=True)) )


	#Report Rilasci
	if "Prog_Task" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Report Progetti/Task" 
		f_reptcode="List_Prog_Task"
		f_sheetname="List_Prog_Task"
		logger.info('start process for ' + 'List_prog_task' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"read_prog","df_prog_bytask")
		from elab_taskMAP import df_prog_bytask
	
		if "AllDip" != sListDip: 
			df_prog_bytask_dip=df_prog_bytask.loc[(df_prog_bytask["MP_Dip"]== sListDip)|(df_prog_bytask["PM_Dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_prog_bytask_dip=df_prog_bytask
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	
#202508		df_prog_bytask_dip["Data I. Corrente"]=pd.to_datetime(df_prog_bytask_dip["Data I. Corrente"],errors='coerce');df_prog_bytask_dip["Data I. Corrente"]=df_prog_bytask_dip["Data I. Corrente"].dt.strftime(wDtFormat)
#202508		df_prog_bytask_dip["Data F. Corrente"]=pd.to_datetime(df_prog_bytask_dip["Data F. Corrente"],errors='coerce')
#Select if project has 0 as prevision and advance
		df_prog_bytask_dip=df_prog_bytask_dip[(df_prog_bytask_dip["Pianif_Corr"]!=0)|(df_prog_bytask_dip["Consuntivo"]!=0)|(df_prog_bytask_dip["Data F. Corrente"].dt.year>=wprogzero_fmyear)]
#Select if project has end date after a date specified in parameter
#202508		df_prog_bytask_dip=df_prog_bytask_dip[(df_prog_bytask_dip["Data F. Corrente"].dt.year>=wprog_fmyear)]
#202508		df_prog_bytask_dip["Data F. Corrente"]=df_prog_bytask_dip["Data F. Corrente"].dt.strftime(wDtFormat)
#202508	df_prog_bytask_dip["Data I. Contrattuale"]=pd.to_datetime(df_prog_bytask_dip["Data I. Contrattuale"],errors='coerce');df_prog_bytask_dip["Data I. Contrattuale"]=df_prog_bytask_dip["Data I. Contrattuale"].dt.strftime(wDtFormat)
#202508		df_prog_bytask_dip["Data F. Contrattuale"]=pd.to_datetime(df_prog_bytask_dip["Data F. Contrattuale"],errors='coerce');df_prog_bytask_dip["Data F. Contrattuale"]=df_prog_bytask_dip["Data F. Contrattuale"].dt.strftime(wDtFormat)
		df_prog_bytask_dip["ResiduoTask"]=df_prog_bytask_dip["Pianif_Corr"]-df_prog_bytask_dip["Consuntivo"]
#Create parameter to write sheets
#202508		df_prog_bytask_dip.loc[df_prog_bytask_dip["Data I. Corrente"] == "01/01/1980", "Data I. Corrente"] = ""
#202508		df_prog_bytask_dip.loc[df_prog_bytask_dip["Data F. Corrente"] == "31/12/2099", "Data F. Corrente"] = ""
#202508		df_prog_bytask_dip.loc[df_prog_bytask_dip["Data I. Contrattuale"] == "01/01/1980", "Data I. Contrattuale"] = ""
#202508		df_prog_bytask_dip.loc[df_prog_bytask_dip["Data F. Contrattuale"] == "31/12/2099", "Data F. Contrattuale"] = ""



		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_prog_bytask_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_prog_bytask_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_prog_bytask_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_prog_bytask_dip.memory_usage(deep=True)) )




#Report Rilasci
	if "Rilasci" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Rilasci Contrattuali" 
		f_reptcode="RilasciContratt"
		f_sheetname="RilasciContratt"
		logger.info('start process for ' + 'List_Rilasci_subt' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"read_rilasci","df_Rilasci_subt")
		from read_rilasci import df_Rilasci_subt
	
		if "AllDip" != sListDip: 
			df_Rilasci_subt_dip=df_Rilasci_subt.loc[(df_Rilasci_subt["resp_ril_dip"].str.contains(sListDip) )|(df_Rilasci_subt["resp_dmd_dip"]==sListDip )]

#			df_Rilasci_subt_dip=df_Rilasci_subt.loc[(sListDip in df_Rilasci_subt["resp_ril_dip"] )|(df_Rilasci_subt["resp_dmd_dip"]==sListDip )]
		#filtra se il controllo dei consuntivi non è OK
		else: df_Rilasci_subt_dip=df_Rilasci_subt
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_Rilasci_subt_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_Rilasci_subt_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_Rilasci_subt_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_Rilasci_subt_dip.memory_usage(deep=True)) )


#Report Rilasci
	if "RilascibyTask" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Rilasci Contrattuali per task" 
		f_reptcode="RilCont_bytask"
		f_sheetname="RilCont_bytask"
		logger.info('start process for ' + 'df_Rilasci_task' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"read_rilasci","df_Rilasci_task")
		from read_rilasci import df_Rilasci_task
	
		if "AllDip" != sListDip: 
			df_Rilasci_task_dip=df_Rilasci_task.loc[(df_Rilasci_task["resp_ril_dip"].str.contains(sListDip) )|(df_Rilasci_task["resp_dmd_dip"]==sListDip )]
		#filtra se il controllo dei consuntivi non è OK
		else: df_Rilasci_task_dip=df_Rilasci_task
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	
		df_Rilasci_task_dip["DataIniz"]=pd.to_datetime(df_Rilasci_task_dip["DataIniz"],errors='coerce');df_Rilasci_task_dip["DataIniz"]=df_Rilasci_task_dip["DataIniz"].dt.strftime(wDtFormat)
		df_Rilasci_task_dip["DataFine"]=pd.to_datetime(df_Rilasci_task_dip["DataFine"],errors='coerce');df_Rilasci_task_dip["DataFine"]=df_Rilasci_task_dip["DataFine"].dt.strftime(wDtFormat)
		df_Rilasci_task_dip["DataValidRC"]=pd.to_datetime(df_Rilasci_task_dip["DataValidRC"],errors='coerce');df_Rilasci_task_dip["DataValidRC"]=df_Rilasci_task_dip["DataValidRC"].dt.strftime(wDtFormat)

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_Rilasci_task_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_Rilasci_task_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_Rilasci_task_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_Rilasci_task_dip.memory_usage(deep=True)) )


	if "List_ODA" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Elenco Ordini di acquisto" 
		f_reptcode="Elab_OdA_Periodo"
		f_sheetname="Elab_OdA_Periodo"

	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"elab_Oda_Periodo","df_ODA_per")
		from elab_Oda_Periodo import df_ODA_recent_per
	
		if "AllDip" != sListDip: 
			df_ODA_per_dip=df_ODA_recent_per[(df_ODA_recent_per["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_ODA_per_dip=df_ODA_recent_per[(df_ODA_recent_per["ROI_dip"]!=0)]
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_ODA_per_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_ODA_per_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_ODA_per_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_ODA_per_dip.memory_usage(deep=True)) )



#Report Rilasci
	if "ODA_periodo" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Prospetto ODA da consuntivare per Struttura e Periodo" 
		f_reptcode="Riep_ODA_periodo"
		f_sheetname="Riep_ODA_periodo"
		logger.info('start process for ' + 'Riep_ODA_periodo' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"elab_Oda_Periodo","df_ODA_per")
		from elab_Oda_Periodo import df_ODA_recent_per
	
		if "AllDip" != sListDip: 
			df_ODA_per_dip=df_ODAnoCons_byROI.loc[(df_ODAnoCons_byROI["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_ODA_per_dip=df_ODAnoCons_byROI[(df_ODAnoCons_byROI["ROI_dip"]!= 0)]
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	
#		df_ODA_per_dip_grp=df_ODA_per_dip.groupby(["ROI_dip","ROI","Anno competenza","Mese competenza","DocAcq","Descrizione DocAcq"],observed=True).agg(numbdo=("DocAcq","count")).reset_index()
#		df_ODA_per_dip_grp["Periodo"]=df_ODA_per_dip_grp["Anno competenza"].astype(str)+"-"+df_ODA_per_dip_grp["Mese competenza"].astype(str).str.zfill(2)
#		df_col_Vpivot_ODA=df_ODA_per_dip_grp.pivot_table(index=["ROI_dip","ROI","DocAcq","Descrizione DocAcq"], columns=["Periodo"],values=["numbdo"],
#							   aggfunc="sum")
#25/09/2024		df_col_Vpivot_BDO=dfBDO_per_dip.pivot_table(index=['Dip'],  
#25/09/2024			values=['BDO'] ,
#25/09/2024			columns=['PeriodoMMM'],
#25/09/2024			aggfunc=pd.Series.nunique, margins=True, margins_name="#Totale#" )
#		df_col_Vpivot_ODA.columns = df_col_Vpivot_ODA.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
#		df_col_Vpivot_ODA=df_col_Vpivot_ODA.reset_index().rename_axis(None,axis=1)

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_ODA_per_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_ODA_per_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_ODA_per_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_ODA_per_dip.memory_usage(deep=True)) )

#Report Rilasci
	if "ODA_periodo" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Riepilogo ODA da consuntivare per Struttura e Periodo" 
		f_reptcode="RiepDip_ODA_daCons"
		f_sheetname="RiepDip_ODA_daCons"
		logger.info('start process for ' + 'RiepDip_ODA_daCons' )
		if "AllDip" != sListDip: 
			df_ODA_per_dip=df_ODAnoCons_byDip.loc[(df_ODAnoCons_byDip["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_ODA_per_dip=df_ODAnoCons_byDip[(df_ODAnoCons_byDip["ROI_dip"]!= 0)]
#		df_ODAnoCons_byDip=df_ODAnoCons_byDip[df_ODAnoCons_byDip]
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#stepprv		write_templShname(f_reptcode,"elab_Oda_Periodo","df_ODA_per")
#stepprv		from elab_Oda_Periodo import df_ODA_per
	
#stepprv		if "AllDip" != sListDip: 
#stepprv			df_ODA_per_dip=df_ODA_per[(df_ODA_per["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
#stepprv		else: df_ODA_per_dip=df_ODA_per
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	
		#df_BDO_per_dip_grp=df_BDO_per_dip.groupby(["ROI_Dip","BDO","ODAGMacr_nbr","Periodo","ROI","Descrizione Bdo","StatoBDO","Stat_VerbAt","Data decorrenza","Data scadenza","Stat_VerbCh","MAP_noverb"],as_index=False).agg({"Documento d'acquisto":"nunique"})
#stepprv		df_ODA_per_dip_grp=df_ODA_per_dip.groupby(["ROI_dip","ROI","Periodo","ODA","Testo testata BDO"]).agg(numbdo=("ODA","count")).reset_index()

#		df_col_Vpivot_ODA=df_ODA_per_dip_grp.pivot_table(index=["ROI_dip","ROI"], columns=["Periodo"],values=["numbdo"],
#							   aggfunc="sum")
#25/09/2024		df_col_Vpivot_BDO=dfBDO_per_dip.pivot_table(index=['Dip'],  
#25/09/2024			values=['BDO'] ,
#25/09/2024			columns=['PeriodoMMM'],
#25/09/2024			aggfunc=pd.Series.nunique, margins=True, margins_name="#Totale#" )
#		df_col_Vpivot_ODA.columns = df_col_Vpivot_ODA.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
#		df_col_Vpivot_ODA=df_col_Vpivot_ODA.reset_index().rename_axis(None,axis=1)
#		df_col_Vpivot_BDO.loc["total"]=df_col_Vpivot_BDO.sum()
#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts}
		scriveFoglio_dict_v5(writer,workbk_Out,df_ODAnoCons_byDip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_ODAnoCons_byDip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_ODAnoCons_byDip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_ODAnoCons_byDip.memory_usage(deep=True)) )

#------------------------
	#Report Risorse interne
	if "RiepRisInt" in sDict.keys() or "AllRept" in sDict.keys():
		f_descript="Risorse Interne per progetto e task" 
		f_reptcode="RiepRisInt_byProg"
		f_sheetname="RiepRisInt_byProg"
		logger.info('start process for ' + 'RiepRisInt_byProg' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"read_prog","df_prog_int_riep")
		from read_prog import df_prog_int_riep


		if "AllDip" != sListDip: 
			df_prog_int_riep_dip=df_prog_int_riep[(df_prog_int_riep["RisInt_Dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_prog_int_riep_dip=df_prog_int_riep
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts}
		scriveFoglio_dict_v5(writer,workbk_Out,df_prog_int_riep_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_prog_int_riep_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_prog_int_riep_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_prog_int_riep_dip.memory_usage(deep=True)) )

	
#-----lettura del file Periodi di Consuntivazione die BDO -> da formato csv come Data Frame (tabelle gestite da Python)
	wNow = pd.to_datetime("now")
	wLast1month = wNow - pd.DateOffset(months=1)
	wLast2month = wNow - pd.DateOffset(months=2)
	#Crea una lista dei periodi (mese corrente e due mesi precedenti)
	wMesiPeriodo = [wNow,wLast1month,wLast2month ]

	#Trascrive contenuto nel dataframe in verticale
	wdf_periodi = pd.DataFrame(wMesiPeriodo)
	wdf_periodi.columns =["mese"]
	wdf_periodi["meseAAAA-MM"]=wdf_periodi["mese"].dt.to_period("M").astype("string")
	wPeriodi2 = wdf_periodi["meseAAAA-MM"].tolist()

	#Report Rilasci
	if "BDOLin_Per" in sDict.keys() or "AllRept" in sDict.keys():

		f_descript="Report stato consuntivazioni mensili BDO per Linea" 
		f_reptcode="BdOLin_Per"
		logger.info('start process for ' + 'BdOLin_Per' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"elab_Bdo_Periodo","df_BDOLin_per")
		from elab_Bdo_Periodo import df_BDOLin_recent_per
		
		if "AllDip" != sListDip: 
			df_BDOLin_per_dip=df_BDOLin_recent_per[(df_BDOLin_recent_per["ROI_dip"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_BDOLin_per_dip=df_BDOLin_recent_per
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	
		for wPeriodSel in wPeriodi2:	
			df_BDOLin_per_dip_per = df_BDOLin_per_dip[df_BDOLin_per_dip['Periodo'] == wPeriodSel]
			f_sheetname="BdOLin_Per_"+wPeriodSel

#202508			df_BDOLin_per_dip_per["Data decorrenza"]=pd.to_datetime(df_BDOLin_per_dip_per["Data decorrenza"],errors='coerce');df_BDOLin_per_dip_per["Data decorrenza"]=df_BDOLin_per_dip_per["Data decorrenza"].dt.strftime(wDtFormat)
#202508			df_BDOLin_per_dip_per["Data scadenza"]=pd.to_datetime(df_BDOLin_per_dip_per["Data scadenza"],errors='coerce');df_BDOLin_per_dip_per["Data scadenza"]=df_BDOLin_per_dip_per["Data scadenza"].dt.strftime(wDtFormat)
#202508			df_BDOLin_per_dip_per["Data_MLS"]=pd.to_datetime(df_BDOLin_per_dip_per["Data_MLS"],errors='coerce');df_BDOLin_per_dip_per["Data_MLS"]=df_BDOLin_per_dip_per["Data_MLS"].dt.strftime(wDtFormat)

#Create parameter to write sheets
			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
			scriveFoglio_dict_v5(writer,workbk_Out,df_BDOLin_per_dip_per,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     

		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_BDOLin_per_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_BDOLin_per_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_BDOLin_per_dip.memory_usage(deep=True)) )
	
	#Report Rilasci
	if "RDI_IF" in sDict.keys() or "AllRept" in sDict.keys():

		f_descript="Report stato richieste di fornitura" 
		f_reptcode="RDI_IF"
		logger.info('start process for ' + 'RDI_IF' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#		write_templShname(f_reptcode,"elab_If_Rdi","df_RDI_port")
		from elab_If_Rdi import df_RDI_port
		
		if "AllDip" != sListDip: 
			df_RDI_port_dip=df_RDI_port.loc[(df_RDI_port["Dip RDI"]== sListDip)]
		#filtra se il controllo dei consuntivi non è OK
		else: df_RDI_port_dip=df_RDI_port
#		df_BDO_dip=df_BDO_dip[(df_BDO_dip["Flag_daCons"]=="S")]	
		f_sheetname="RDI"

#202508		df_RDI_port_dip["Data_lavorazione_byPMO-BO"]=pd.to_datetime(df_RDI_port_dip["Data_lavorazione_byPMO-BO"].astype(str), format='%Y-%m-%d',errors='coerce');df_RDI_port_dip["Data_lavorazione_byPMO-BO"]=df_RDI_port_dip["Data_lavorazione_byPMO-BO"].dt.strftime("%d.%m.%Y")
#202508		df_RDI_port_dip["DataCreaz"]=pd.to_datetime(df_RDI_port_dip["DataCreazione Bdo"].astype(str), format='%Y-%m-%d',errors='coerce');df_RDI_port_dip["Data Creazione Bdo"]=df_RDI_port_dip["Data Creazione Bdo"].dt.strftime("%d.%m.%Y")

	#	df_RDI_port_dip["Data_lavorazione_byPMO-BO"]=df_RDI_port_dip["Data_lavorazione_byPMO-BO"].dt.strftime("%d.%m.%Y")
	#	df_RDI_port_dip["Data Creazione Bdo"]=df_RDI_port_dip["Data Creazione Bdo"].dt.strftime("%d.%m.%Y")
#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":wEnvironment,"rpt_debug":wDebugsts }
		scriveFoglio_dict_v5(writer,workbk_Out,df_RDI_port_dip,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     

		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_RDI_port_dip.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_RDI_port_dip.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_RDI_port_dip.memory_usage(deep=True)) )


#scriveFoglio(workbk_Out,worksh_Out,work_dframe,rpt_Code,rpt_Title,rpt_format)
	writer.close()

#imposta il file come read only
	os.chmod(wRept_output, S_IREAD|S_IRGRP|S_IROTH)



#call function from main module
if __name__ == "__main__":
    wlist=["DIH","AllDip"]
	#"AllDip",
    wDict={
#"AllDip"
#"parm1":"value1",
#"parm2":"value2",
#"parm3":"value3"
}
    Rep116_rpt_gen(wlist,wDict)
    pass