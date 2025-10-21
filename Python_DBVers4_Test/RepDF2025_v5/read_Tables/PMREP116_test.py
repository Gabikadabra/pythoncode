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
#	from elab_ril_verb import df_Rilasci_subt
#	from read_verbPass import wDict_Verb
#	from elab_Bdo_pivot import df_BDOnoCons_byDip,df_BDOnoCons_byBDO,df_BDOnoVerb_byBDO,df_BDOVerb_recent_per,df_ODAnoCons_byDip,df_ODAnoCons_byROI
#	from elab_Bdo_Periodo  import df_BDO_recent_per
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
    wlist=["AllDip"]
	#"AllDip",
    wDict={
#"AllDip"
#"parm1":"value1",
#"parm2":"value2",
#"parm3":"value3"
}
    Rep116_rpt_gen(wlist,wDict)
    pass