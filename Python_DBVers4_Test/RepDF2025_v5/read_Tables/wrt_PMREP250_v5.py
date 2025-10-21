##PMREP116 - Report Progetti e ciclo passivo per strutture
from funct.Dframe_templSh import write_templShname
from Util_leggeParams import parm_dict,listfilter_ODAG,listfilter_BDO,listfilter_MAP,listfilter_Task
from Util_leggeDir import dir_dict
import sys
#(f_reptname,f_lib_dframe,f_wrk_dframe):
from Util_logging import logger
from read_verbPass import wDict_Verb
#wPythProc="read_verbPass"
import timeit  #funzione timeit per valutare durata estrazione
#tempoInz=timeit.default_timer()



def Rep250_rpt_gen(fListSel,fODAGnbr,fDict):
	#wTest=False
	wPrefixFileOut="PMREP250_ODAG_v5_"
	wNomeFileOut=wPrefixFileOut+"_ALL" 
	#Rep116prep_sub(wTest,"AllDip","AllRept",wNomeFileOut)
	for wSel in fListSel:
		if wSel == "AllRef":
			wNomeFileOut=wPrefixFileOut 
			Rep250_rpt_sub(wSel,fODAGnbr,wNomeFileOut,fDict)		
		else:    #wsel "AllIn" (pe tutti i referenti) "SelODAG" (selez ODAG) or []
			wNomeFileOut=wPrefixFileOut  # +"_"+wDip
			Rep250_rpt_sub(wSel,fODAGnbr,wNomeFileOut,fDict)
#sSel Lista Referenti (DEC o RUP); sODAGnbr elenco ODAG; sDict dictionary
def Rep250_rpt_sub(sSel,sODAGnbr, sNomeFileOut,sDict):
	pass
#	import logging
	import logging.config
	import os
	#status create / modify read-only file 
	from stat import S_IREAD, S_IRGRP, S_IROTH, S_IWUSR, S_IWRITE, S_IWOTH, S_IRWXU 
	import contextlib

#2)	Gestione DIRECTORY CORRENTE (contestualizzazione cartella python sulla base dell’utente)
#da inserire -> selezione directory sulla base dell'utente che utilizza Python
	current_dir=os.getcwd()
	print(current_dir)

	import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	

	import pandas as pd
# soluzione per scrivere Excel 
#	import xlsxwriter
#from read_rilasci import df_Rilasci_task, df_Rilasci_subt
#from funct.Dframe_toSheet import scriveFoglio 

	from funct.Dframe_toSheet import scriveFoglio_dict 
	
	wPythProc="wrt_PMREP250"
	tempoInz=timeit.default_timer()

	wDtFormat = "%d/%m/%Y" 
	wrk_CsvDir=dir_dict["csvDir"]
	wFmtFile="RptFormati"

#print (dir_dict["rptDir"])
	wrk_RptDir=dir_dict["rptDir"]
	wrk_CsvDir=dir_dict["xlsDir"]
	wrk_stagDir=dir_dict["stagDir"]
#"df_MAP_bytask":"read_map",

	f_rptDir= wrk_RptDir
	f_reptname=sNomeFileOut #"PMRep FileName "
#	if not sDict:
#		sDict["AllRept"]="All reports"



	#-----Imposta le librerie (vengono richiamate funzioni presenti in altri moduli)
#Legge la libreria Pandas per il DataFrame/gestione tabelle
#	import pandas as pd
#	import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
	from datetime import timedelta,datetime  
#predispone funzioni di utilità
#	from Funz.Util_ReportDate import GetRptDate,GetAlfb
	#from Funz.Util_wDir import RepInit0	#Per la scrittura su Excel si usa la libreria xlsxwriter (limite che non si possono leggere file Excel e non si possono modificare con quella libreria)
#	from Funz.Util_RptHeader import RepTitoloData, RepColumnFormat
#	from Funz.Estrae_InfoBDO import estrae_InfoBDO
# soluzione per scrivere Excel 
	import xlsxwriter
#-----Predisposizione dell'ambiente per leggere i file e scrivere su Excel
	wDt_now = pd.to_datetime('now') # Restituisce la data di oggi
	#wDt_2mthBefore = wDt_now - pd.DateOffset(months=2) #Ottiene i due mesi precedenti
	#Estrae le dir di input (wdirInp)  e quelle di output (wdirOut)
	#wdirInp,wdirOut=RepInit0()
	#estrae la data di oggi	
	#wToday = pd.to_datetime('now')
#Definire nome file output
#	df_InfoBDO=estrae_InfoBDO()
#	nome_xls= fNomeFileOut		
#2)	Gestione DIRECTORY CORRENTE (contestualizzazione cartella python sulla base dell’utente)
#da inserire -> selezione directory sulla base dell'utente che utilizza Python
	current_dir=os.getcwd()
	print(current_dir)
#	logging.config.fileConfig('logging2.conf')

# create logger
#	logger = logging.getLogger('simpleExample')
#	import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	from elab_Map_byMacr import df_ODAGMacr, df_ODAG, dfMAP_Mg_R
	from read_Interni import df_RisInt_byName
	from elab_Bdo_Periodo import df_BDO_recent_per,df_BDO 
	wDim= str(df_ODAG.shape[0]);wDescr="Num. ODAG ";logger.debug(wDescr + wDim );logger.error(wDescr + wDim );logger.warning(wDescr + wDim );logger.info(wDescr + wDim )
	wDim= str(df_ODAGMacr.shape[0]);wDescr="Num. Macrocl ";logger.debug(wDescr + wDim );logger.error(wDescr + wDim );logger.warning(wDescr + wDim );logger.info(wDescr + wDim )
	wDim= str(df_BDO.shape[0]);wDescr="Num. BDO ";logger.debug(wDescr + wDim );logger.error(wDescr + wDim );logger.warning(wDescr + wDim );logger.info(wDescr + wDim )

	wFmtFile="RptFormati"

#print (dir_dict["rptDir"])
	wrk_RptDir=dir_dict["rptDir"]
	wrk_CsvDir=dir_dict["xlsDir"]
	wrk_stagDir=dir_dict["stagDir"]
#"df_MAP_bytask":"read_map",
#Create all different reports for all references 
	#	df["Data_Iniz"]=df["Data_Iniz"].str.split(" ").str[0]
	df_ODAG["Data scadenza"]=pd.to_datetime(df_ODAG["Data scadenza"])
	df_ODAGper=df_ODAG[(df_ODAG["Data scadenza"] >= pd.to_datetime('now')- pd.DateOffset(months=2))&(df_ODAG["Valore Contratto"] >1000)]

	if sSel=="AllRef":
#get combination of two columns
#   create tuple from unique values on Data frame (DEC or RUP )
		wunique_values0 = df_ODAGper['DEC'].dropna().unique()
		wunique_values = df_ODAGper['RUP'].dropna().unique()
#aggiunge nella stessa lista le due  estrazioni
		wunique_values=np.append(wunique_values,wunique_values0)
#Elenco dei DEC e dei RUP
		wunique_values=list(dict.fromkeys(wunique_values))
		pass
		#for each reference create a single file
		for wref in list(filter(None, wunique_values)):		
#			dfODAG_byref = dfODAG[(dfODAG['ODAG_DEC'] == wref) | (dfODAG['ODAG_RUP'] == wref)]
			dfODAG_sel = df_ODAGper[(df_ODAG['DEC'] == wref)]
			dfODAG_sel.drop_duplicates(inplace=True)
			dfODAG_sel["ref"]="DEC"

			wunique_ODAGDEC = dfODAG_sel[["ODAGnbr","ref"]].values.tolist()
			
#			dfODAG_sel["ODAG_nbr"].dropna().unique()
#			wunique_ODAGDEC = [x+("DEC") for x in wunique_ODAGDEC]

#			for index in range(len(wunique_ODAGDEC)):
#				wunique_ODAGDEC[index].append("DEC")

			dfODAG_sel = df_ODAGper[(df_ODAG['RUP'] == wref)]
			dfODAG_sel.drop_duplicates(inplace=True)
			dfODAG_sel["ref"]="RUP"
			wunique_ODAGRUP = dfODAG_sel[["ODAGnbr","ref"]].values.tolist()
#			wunique_ODAGRUP = dfODAG_sel['ODAG_nbr','ref'].dropna().unique()
#			wunique_ODAGRUP = [x+ ("RUP") for x in wunique_ODAGRUP]
#			for index in range(len(wunique_ODAGRUP)):
#				wunique_ODAGRUP[index].append("RUP")			
			wunique_ODAG_nbr = wunique_ODAGRUP+wunique_ODAGDEC
#Create file Excel 
# Create Excel file
			if wref in df_RisInt_byName["RisInt_nome"].values:
				wDip= df_RisInt_byName.loc[(df_RisInt_byName["RisInt_nome"] == wref), 'Int_Dip'].iloc[0]
			else:
				wDip="none"
# Create new file for reference
		#get first 4 chars from ref name
			wref_pref=wref[0:8]
# Create an initial page with all ODAg for same references
			wRept_output=wrk_RptDir+"\\"+f_reptname+"_"+wDip+"_"+wref_pref+".xlsx"
	#verifica se esiste il file e se esiste lo rimuove
			with contextlib.suppress(FileNotFoundError):
				os.chmod(wRept_output, S_IRWXU) #cambia le autorizzazioni sul file
				os.remove(wRept_output)
	
			writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
			workbk_Out=writer.book

			money_fmt = workbk_Out.add_format({'num_format': '#,##0.00'})
			date_fmt = workbk_Out.add_format({'num_format': 'dd/mm/yyyy'})
# create GLOBAL sheet
# First Sheet - Excel ODAG Summary
			wSht_first="RiepODAG"
			worksh_first_Out = workbk_Out.add_worksheet(wSht_first)

	#registra il progressivo del numero di tabella (es. Table1)
			wNumTable = 0

		#	create an empty dataframe as the ODAG one
	#dfODAG_gen = pd.DataFrame().reindex_like(dfODAG)
## Create empty tables to collect for reference
			dfODAG_gen = df_ODAG.iloc[:0,:].copy()
			dfVerb_gen = df_BDO_recent_per.iloc[:0,:].copy()
			dfBDO_gen = df_BDO.iloc[:0,:].copy()

#Formula to write back to main sheet
			wbck_mainsheet='=HYPERLINK(CELL("address",RiepODAG!A2), "link foglio principale")'
# create Riep Sheet by DEC/RUP
# for each ODag Number for specified reference
			for (wODAGsel,wruolo) in wunique_ODAG_nbr:
			#Create Main ODAG sheet with macrodetails
			# directory, filename, shname, rept title
		#-------------------------loop for each ODAG
#				dfMAP_filt=dfMAP_M.loc[dfMAP_M["ODAG_nbr"]==wODAGsel]
				dfBDO_filt=df_BDO.loc[df_BDO["ODAGnbr"]==wODAGsel]
				dfODAG_filt=df_ODAG.loc[df_ODAG["ODAGnbr"]==wODAGsel]
#				dfBEF_filt=dfBEF.loc[dfBEF["ODAG_nbr"]==wODAGsel]
#				dfBEF_nopag_filt=dfBEF_nopag.loc[dfBEF_nopag["ODAG_nbr"]==wODAGsel]

				dfODAGMacr_filt=df_ODAGMacr.loc[df_ODAGMacr["ODAGnbr"]==wODAGsel]
				dfMAP_Mg_filt=dfMAP_Mg_R.loc[dfMAP_Mg_R["ODAGnbr"]==wODAGsel]
				dfBDO_per_filt=df_BDO_recent_per.loc[df_BDO_recent_per["ODAGnbr"]==wODAGsel]

				wmin_vResid=dfODAGMacr_filt["Libero su macroclasse"].min()
				wmin_vRes6m=dfODAGMacr_filt["StimaRes_06m"].min()
				wmin_vRes12m=dfODAGMacr_filt["StimaRes_12m"].min()
## Write report ODAG Macr 
				f_descript="Riepilogo ODAG " + str(wODAGsel)

				logger.info('start process for ' + 'ODAGMacr_riep' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#				write_templShname(f_reptcode,"elab_If_Rdi","df_RDI_port")
#				from elab_If_Rdi import df_RDI_port

				f_sheetname="Riep_"+str(wODAGsel)

				wDim= str(dfODAGMacr_filt.shape[0])   
				logger.info('foglio ODAG Macroclasse' + ' dimension ' + wDim )
				wstrTipof=int(wDim) +9 +4 #Calculate row for following table

#write data ODAG
				f_reptcode="ODAG_riep"
				dfODAG_filt= dfODAG_filt.sort_values(by=["ODAGnbr"],ascending=[True])
#inserisce i riferimenti alle pagine dei diversi ODAG
				dfODAG_filt["link ODAG"]='=HYPERLINK(CELL("address",'+'Riep_'+dfODAG_filt["ODAGnbr"].astype(str)+'!A2), "RiepContratto_'+dfODAG_filt["ODAGnbr"].astype(str)+'")'
#Convert date format to dd/mm/aaaa
				w_column="Data Decorrenza";dfODAG_filt[w_column]=pd.to_datetime(dfODAG_filt[w_column].astype(str), format='%d.%m.%Y');dfODAG_filt[w_column]=dfODAG_filt[w_column].dt.strftime('%d/%m/%Y')
				w_column="Data scadenza";dfODAG_filt[w_column]=dfODAG_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))

				rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-nofmt","rpt_dftTitle":f_descript,"rpt_strRow":5,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" ,"rpt_prvrow":"Riepilogo contratto" }
				scriveFoglio_dict(writer,workbk_Out,dfODAG_filt,rept_dict)

#write data ODAG
				f_reptcode="ODAGMacr_tipof"
				rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-nofmt","rpt_dftTitle":f_descript,"rpt_strRow":wstrTipof,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N","rpt_prvrow":"Riepilogo per Macroclasse e Tipologia fornitura"  }
				scriveFoglio_dict(writer,workbk_Out,dfMAP_Mg_filt,rept_dict)

#Create parameter to write sheets
				f_reptcode="ODAGMacr_riep"
				rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-hdr","rpt_dftTitle":f_descript,"rpt_strRow":9,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" ,"rpt_prvrow":"Riepilogo per Macroclasse" }
				scriveFoglio_dict(writer,workbk_Out,dfODAGMacr_filt,rept_dict)

#Genric format for all the columns in excel sheet
				worksh_Out=writer.sheets[f_sheetname]
				worksh_Out.set_column(4,15,14,money_fmt)

#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
# ODAGMacr_tipof ODAG_riep


				if wmin_vResid <0:
					wODAGstatus="**macroclasse neg."
				elif wmin_vRes6m <0:
					wODAGstatus="*res.macrocl <6 mesi"
				elif wmin_vRes12m <0:
					wODAGstatus="*res.macrocl <12 mesi"
				else:
					wODAGstatus=""
				dfODAG_filt["Stato macroclassi"]=wODAGstatus

			#change filter role based for reports
				wfilt="In approvazione DEC" if wruolo =="DEC" else "In approvazione RUP"

				dfBDO_per_filt_ref=dfBDO_per_filt[(dfBDO_per_filt['Stat_VerbAt'] == wfilt) | (dfBDO_per_filt['Stat_VerbCh'] == wfilt)| (dfBDO_per_filt['Stat_VerbSAL'] == wfilt)]
			#Add report(s) selected to main table for reference (Verbali)
				dfVerb_gen=pd.concat([dfVerb_gen,dfBDO_per_filt_ref], ignore_index=True)
				dfVerb_gen.info()
			#change filter role based for reports
				wfilt="In app.dal Dec/DL" if wruolo =="DEC" else "In approvazione dal RUP"			
				dfBDO_filt_ref=dfBDO_filt[(dfBDO_filt['StatoBDO'] == wfilt)]
			#Add report(s) selected to main table for reference (BDO)
				dfBDO_gen=pd.concat([dfBDO_gen,dfBDO_filt_ref], ignore_index=True)
			#Add report(s) selected to main table for reference (ODAG - contracts)
				dfODAG_gen=pd.concat([dfODAG_gen,dfODAG_filt], ignore_index=True)
	
## Write report ODAG Macr 
				f_descript="Dettaglio BDO e RDI ODAG " + str(wODAGsel)
				f_sheetname="Dett_"+str(wODAGsel)	
#			wbck_mainsheet='=HYPERLINK(CELL("address",RiepODAG!A2), "link foglio principale")'
#Create parameter to write sheets
				f_reptcode="Dett_ODAG_BDO"

				w_column="Data decorrenza";dfBDO_filt[w_column]=dfBDO_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
				w_column="Data scadenza";dfBDO_filt[w_column]=dfBDO_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))

				rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":5,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" }
				scriveFoglio_dict(writer,workbk_Out,dfBDO_filt,rept_dict)

				f_descript="Riepilogo Verbali BDO" 
				f_reptcode="Dett_ODAG_Verbali"
				f_sheetname="Riep_Verb"+str(wODAGsel)
				wDimBDO= str(dfBDO_filt.shape[0])   # get lines for BDO filtered
				w_column="Data decorrenza";dfBDO_per_filt[w_column]=dfBDO_per_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
				w_column="Data scadenza";dfBDO_per_filt[w_column]=dfBDO_per_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
#Create parameter to write sheets
				rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":5 ,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" }
				scriveFoglio_dict(writer,workbk_Out,dfBDO_per_filt,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     

##### Recap DEC / RUP
## Write report ODAG Macr 
			f_descript="Riepilogo Generale ODAG - " + wref
			f_sheetname="RiepODAG"	
#			wbck_mainsheet='=HYPERLINK(CELL("address",RiepODAG!A2), "link foglio principale")'
#Create parameter to write sheets
			f_reptcode="ODAG_riep"
			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"Rpt000hdr","rpt_dftTitle":f_descript,"rpt_strRow":5,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" }

			dfODAG_gen= dfODAG_gen.sort_values(by=["ODAGnbr"],ascending=[True])
#inserisce i riferimenti alle pagine dei diversi ODAG
			dfODAG_gen["link ODAG"]='=HYPERLINK(CELL("address",'+'Riep_'+dfODAG_gen["ODAGnbr"].astype(str)+'!A2), "RiepContratto_'+dfODAG_gen["ODAGnbr"].astype(str)+'")'

			scriveFoglio_dict(writer,workbk_Out,dfODAG_gen,rept_dict)
			wDim= str(dfODAG_gen.shape[0])   
			logger.info('foglio dfODAG_gen ' + ' dimension ' + wDim )
			wstrBDO=int(wDim) +8 #Calculate row for following table
#-----------------------------------
			f_descript="Riepilogo BDO" 
			f_reptcode="Riep_ODAG_BDO"
			f_sheetname="RiepODAG"

			w_column="Data decorrenza";dfBDO_gen[w_column]=dfBDO_gen[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
			w_column="Data scadenza";dfBDO_gen[w_column]=dfBDO_gen[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
#Create parameter to write sheets
			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-","rpt_dftTitle":f_descript,"rpt_strRow":wstrBDO,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" ,"rpt_prvrow":"BDO in approvazione"}
			scriveFoglio_dict(writer,workbk_Out,dfBDO_gen,rept_dict)

			wDim= str(dfBDO_gen.shape[0])   
			logger.info('foglio dfBDO_gen ' + ' dimension ' + wDim )
			wstrVerb=int(wDim) +wstrBDO+4 #Calculate row for following table



			f_descript="Riepilogo Verbali" 
			f_reptcode="Riep_ODAG_Verbali"
			f_sheetname="RiepODAG"

			w_column="Data decorrenza";dfVerb_gen[w_column]=dfVerb_gen[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
			w_column="Data scadenza";dfVerb_gen[w_column]=dfVerb_gen[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))

#Create parameter to write sheets
			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-","rpt_dftTitle":f_descript,"rpt_strRow":wstrVerb,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"Y","rpt_prvrow":"Verbali in approvazione" }
			scriveFoglio_dict(writer,workbk_Out,dfVerb_gen,rept_dict)

			worksh_Out=writer.sheets[f_sheetname]
			worksh_Out.set_column(0,0,14)
			worksh_Out.set_column(4,15,14,money_fmt)
#			worksh_Out.write_formula(1,5,wbck_mainsheet)


			writer.close()
#imposta il file come read only
			os.chmod(wRept_output, S_IREAD|S_IRGRP|S_IROTH)

# create Riep Sheet by contract

# Close Excel file

			#Create main sheet
			
	if sSel=="AllIn":
		#	wunique_ODAG_nbr = dfODAG["ODAG_nbr"].dropna().unique()
		dfODAG_sel = df_ODAGper
		dfODAG_sel["ref"]="all"
		wunique_ODAG_nbr = dfODAG_sel["ODAGnbr"].unique()
		wDip="ALL"
		wref_pref=""

		wRept_output=wrk_RptDir+"\\"+f_reptname+"_"+wDip+"_"+wref_pref+".xlsx"
#----------------------------------------------------
		with contextlib.suppress(FileNotFoundError):
			os.chmod(wRept_output, S_IRWXU) #cambia le autorizzazioni sul file
			os.remove(wRept_output)

		writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
		workbk_Out=writer.book
		money_fmt = workbk_Out.add_format({'num_format': '#,##0.00'})
		date_fmt = workbk_Out.add_format({'num_format': 'dd/mm/yyyy'})
# create GLOBAL sheet
# First Sheet - Excel ODAG Summary
		wSht_first="RiepODAG"
		worksh_first_Out = workbk_Out.add_worksheet(wSht_first)

		#registra il progressivo del numero di tabella (es. Table1)
		wNumTable = 0

		#	create an empty dataframe as the ODAG one
	#dfODAG_gen = pd.DataFrame().reindex_like(dfODAG)
## Create empty tables to collect for reference
		dfODAG_gen = df_ODAG.iloc[:0,:].copy()
		dfVerb_gen = df_BDO_recent_per.iloc[:0,:].copy()
		dfBDO_gen = df_BDO.iloc[:0,:].copy()
#Formula to write back to main sheet
		wbck_mainsheet='=HYPERLINK(CELL("address",RiepODAG!A2), "link foglio principale")'

# create Riep Sheet by DEC/RUP
# for each ODag Number for specified reference
		for wODAGsel in wunique_ODAG_nbr:
			#Create Main ODAG sheet with macrodetails
			# directory, filename, shname, rept title
		#-------------------------loop for each ODAG
#				dfMAP_filt=dfMAP_M.loc[dfMAP_M["ODAG_nbr"]==wODAGsel]
			dfBDO_filt=df_BDO.loc[df_BDO["ODAGnbr"]==wODAGsel]
			dfODAG_filt=df_ODAG.loc[df_ODAG["ODAGnbr"]==wODAGsel]
#				dfBEF_filt=dfBEF.loc[dfBEF["ODAG_nbr"]==wODAGsel]
#				dfBEF_nopag_filt=dfBEF_nopag.loc[dfBEF_nopag["ODAG_nbr"]==wODAGsel]

			dfODAGMacr_filt=df_ODAGMacr.loc[df_ODAGMacr["ODAGnbr"]==wODAGsel]
			dfMAP_Mg_filt=dfMAP_Mg_R.loc[dfMAP_Mg_R["ODAGnbr"]==wODAGsel]
			dfBDO_per_filt=df_BDO_recent_per.loc[df_BDO_recent_per["ODAGnbr"]==wODAGsel]

			wmin_vResid=dfODAGMacr_filt["Libero su macroclasse"].min()
			wmin_vRes6m=dfODAGMacr_filt["StimaRes_06m"].min()
			wmin_vRes12m=dfODAGMacr_filt["StimaRes_12m"].min()
## Write report ODAG Macr 
			f_descript="Riepilogo ODAG " + str(wODAGsel)

			logger.info('start process for ' + 'ODAGMacr_riep' )
	#Check if sheetname as reptcode exist, if not create dataframe columns based on dFrame structure
#				write_templShname(f_reptcode,"elab_If_Rdi","df_RDI_port")
#				from elab_If_Rdi import df_RDI_port

			f_sheetname="Riep_"+str(wODAGsel)

			wDim= str(dfODAGMacr_filt.shape[0])   
			logger.info('foglio ODAG Macroclasse' + ' dimension ' + wDim )
			wstrTipof=int(wDim) +9 +4 #Calculate row for following table

#write data ODAG
			f_reptcode="ODAG_riep"

#riformattare			w_column="Data Decorrenza";dfODAG_filt[w_column]=dfODAG_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
#riformattare			w_column="Data scadenza";dfODAG_filt[w_column]=dfODAG_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))

			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-nofmt","rpt_dftTitle":f_descript,"rpt_strRow":5,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" ,"rpt_prvrow":"Riepilogo contratto" }
			scriveFoglio_dict(writer,workbk_Out,dfODAG_filt,rept_dict)
			worksh_Out=writer.sheets[f_sheetname];worksh_Out.write_formula(1,8,wbck_mainsheet)
#write data ODAG
			f_reptcode="ODAGMacr_tipof"
			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-nofmt","rpt_dftTitle":f_descript,"rpt_strRow":wstrTipof,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N","rpt_prvrow":"Riepilogo per Macroclasse e Tipologia fornitura"  }
			scriveFoglio_dict(writer,workbk_Out,dfMAP_Mg_filt,rept_dict)

#Create parameter to write sheets
			f_reptcode="ODAGMacr_riep"
			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000nohdr","rpt_dftTitle":f_descript,"rpt_strRow":9,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" ,"rpt_prvrow":"Riepilogo per Macroclasse" }
			scriveFoglio_dict(writer,workbk_Out,dfODAGMacr_filt,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
# Link to go bak to main sheet
			worksh_Out=writer.sheets[f_sheetname]
			worksh_Out.set_column(4,15,14,money_fmt)

			if wmin_vResid <0:
				wODAGstatus="**macroclasse neg."
			elif wmin_vRes6m <0:
				wODAGstatus="*res.macrocl <6 mesi"
			elif wmin_vRes12m <0:
				wODAGstatus="*res.macrocl <12 mesi"
			else:
				wODAGstatus=""
			dfODAG_filt["Stato macroclassi"]=wODAGstatus
			#select BDO in for approval status 
			w_BDOSel_inclstat=["In approvazione DEC", "In approvazione RUP"]
			dfBDO_filt_ref=dfBDO_filt['StatoDocAcq'].isin(w_BDOSel_inclstat)
					#select Verb in for approval status
			w_VerbSel_inclstat=["In app.dal Dec/DL", "In approvazione dal RUP","In firma ROI"]
			dfBDO_per_filt["Stat_VerbCh"]=dfBDO_per_filt["ID_VerbCh"].replace(wDict_Verb)
			dfBDO_per_filt["Stat_VerbAt"]=dfBDO_per_filt["ID_VerbAt"].replace(wDict_Verb)
			dfBDO_per_filt["Stat_VerbSAL"]=dfBDO_per_filt["ID_VerbSAL"].replace(wDict_Verb)


			mask1=(dfBDO_per_filt['Stat_VerbAt'].isin(w_VerbSel_inclstat)) | (dfBDO_per_filt['Stat_VerbCh'].isin(w_VerbSel_inclstat))| (dfBDO_per_filt['Stat_VerbSAL'].isin(w_VerbSel_inclstat))
			dfBDO_per_filt_ref=dfBDO_per_filt[mask1]
					#dfBDO_per_filt_ref=dfBDO_per_filt['In approvazione DEC'].isin(w_BDOSel_inclstat)
						#Add report(s) selected to main table for reference (Verbali)
			dfVerb_gen=pd.concat([dfVerb_gen,dfBDO_per_filt_ref], ignore_index=True)
			dfVerb_gen.info()
			
			#Add report(s) selected to main table for reference (BDO)
			dfBDO_gen=pd.concat([dfBDO_gen,dfBDO_filt_ref], ignore_index=True)
			#Add report(s) selected to main table for reference (ODAG - contracts)
			dfODAG_gen=pd.concat([dfODAG_gen,dfODAG_filt], ignore_index=True)
	
## Write report ODAG Macr 
			f_descript="Dettaglio BDO e RDI ODAG " + str(wODAGsel)
			f_sheetname="Dett_"+str(wODAGsel)	
#			wbck_mainsheet='=HYPERLINK(CELL("address",RiepODAG!A2), "link foglio principale")'
#Create parameter to write sheets
			f_reptcode="Dett_ODAG_BDO"

#			w_column="Data decorrenza";dfBDO_filt[w_column]=dfBDO_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
#			w_column="Data scadenza";dfBDO_filt[w_column]=dfBDO_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))

			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":5,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" }
			scriveFoglio_dict(writer,workbk_Out,dfBDO_filt,rept_dict)


			f_descript="Riepilogo Verbali BDO" 
			f_reptcode="Dett_ODAG_Verbali"
			f_sheetname="Riep_Verb"+str(wODAGsel)
			wDimBDO= str(dfBDO_filt.shape[0])   # get lines for BDO filtered
#Create parameter to write sheets
#			w_column="Data decorrenza";dfBDO_per_filt[w_column]=dfBDO_per_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
#			w_column="Data scadenza";dfBDO_per_filt[w_column]=dfBDO_per_filt[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))

			rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":5 ,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" }
			scriveFoglio_dict(writer,workbk_Out,dfBDO_per_filt,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght    


##### Recap DEC / RUP

		wDim= str(dfODAG_gen.shape[0])   
		logger.info('foglio dfODAG_gen ' + ' dimension ' + wDim )
		wstrBDO=int(wDim) +8 #Calculate row for following table
#-----------------------------------
		f_descript="Riepilogo BDO" 
		f_reptcode="Riep_ODAG_BDO"
		f_sheetname="RiepODAG"

#Create parameter to write sheets
#		dfBDO_gen.dropna(subset=["BDO"],inplace=True)
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-","rpt_dftTitle":f_descript,"rpt_strRow":wstrBDO,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" ,"rpt_prvrow":"BDO in approvazione"}
		scriveFoglio_dict(writer,workbk_Out,dfBDO_gen,rept_dict)

		wDim= str(dfBDO_gen.shape[0])   
		logger.info('foglio dfBDO_gen ' + ' dimension ' + wDim )
		wstrVerb=int(wDim) +wstrBDO+4 #Calculate row for following table



		f_descript="Riepilogo Verbali" 
		f_reptcode="Riep_ODAG_Verbali"
		f_sheetname="RiepODAG"

#riformattare		w_column="Data decorrenza";dfVerb_gen[w_column]=dfVerb_gen[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
#riformattare		w_column="Data scadenza";dfVerb_gen[w_column]=dfVerb_gen[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))

#Create parameter to write sheets
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000-","rpt_dftTitle":f_descript,"rpt_strRow":wstrVerb,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"Y","rpt_prvrow":"Verbali in approvazione" }
		scriveFoglio_dict(writer,workbk_Out,dfVerb_gen,rept_dict)
## Write report ODAG Macr 
		f_descript="Riepilogo Generale ODAG - " #+ wref
		f_sheetname="RiepODAG"	
#			wbck_mainsheet='=HYPERLINK(CELL("address",RiepODAG!A2), "link foglio principale")'
#Create parameter to write sheets
		f_reptcode="ODAG_riep"

#Convert date format to dd/mm/aaaa
#riformattare		w_column="Data Decorrenza";dfODAG_gen[w_column]=dfODAG_gen[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
#riformattare		w_column="Data scadenza";dfODAG_gen[w_column]=dfODAG_gen[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))

		dfODAG_gen= dfODAG_gen.sort_values(by=["ODAGnbr"],ascending=[True])
#inserisce i riferimenti alle pagine dei diversi ODAG
		dfODAG_gen["link ODAG"]='=HYPERLINK(CELL("address",'+'Riep_'+dfODAG_gen["ODAGnbr"].astype(str)+'!A2), "RiepContratto_'+dfODAG_gen["ODAGnbr"].astype(str)+'")'
		#dfODAG_gen=dfODAG_gen.reindex(columns=wReindex_ODAGen)
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"Rpt000hdr","rpt_dftTitle":f_descript,"rpt_strRow":5,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" }
		scriveFoglio_dict(writer,workbk_Out,dfODAG_gen,rept_dict)

		worksh_Out=writer.sheets[f_sheetname]
		worksh_Out.set_column(4,15,14,money_fmt)

#			worksh_Out.write_formula(1,5,wbck_mainsheet)


		writer.close()
#imposta il file come read only
		os.chmod(wRept_output, S_IREAD|S_IRGRP|S_IROTH)		

#Write detail macroclasse & BDo in specific report
#202506   PMTRP251 
#	if sSel=="Allin":
		wPrefixFileOut="PMREP251_Macrocl"
		wNomeFileOut=wPrefixFileOut+"_ALL" 
#get combination of two columns
		wunique_ODAG_nbr = df_ODAGper[["ODAGnbr"]].values.tolist()
			
		wRept_output=wrk_RptDir+"\\"+wNomeFileOut+".xlsx"
	#verifica se esiste il file e se esiste lo rimuove
	#	with contextlib.suppress(FileNotFoundError):
	#		os.chmod(wRept_output, S_IRWXU) #cambia le autorizzazioni sul file
	#		os.remove(wRept_output)
	
		writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
		workbk_Out=writer.book

		money_fmt = workbk_Out.add_format({'num_format': '#,##0.00'})
		date_fmt = workbk_Out.add_format({'num_format': 'dd/mm/yyyy'})
# create GLOBAL sheet
# First Sheet - Excel ODAG Summary
		wSht_first="RiepODAG"
		worksh_first_Out = workbk_Out.add_worksheet(wSht_first)
#registra il progressivo del numero di tabella (es. Table1)

#    	wmin_vResid=df_ODAGMacr["Libero su macroclasse"].min()
#        wmin_vRes6m=df_ODAGMacr["StimaRes_06m"].min()
#        wmin_vRes12m=df_ODAGMacr["StimaRes_12m"].min()


#        if wmin_vResid <0:
#			wODAGstatus="**macroclasse neg."
#		elif wmin_vRes6m <0:
#			wODAGstatus="*res.macrocl <6 mesi"
#		elif wmin_vRes12m <0:
#			wODAGstatus="*res.macrocl <12 mesi"
#		else:
#			wODAGstatus=""
#		df_ODAGper["Stato macroclassi"]=wODAGstatus

#### Recap DEC / RUP
## Write report ODAG Macr 
		f_descript="Riepilogo Generale ODAG - "
		f_sheetname="RiepODAG"	
#			wbck_mainsheet='=HYPERLINK(CELL("address",RiepODAG!A2), "link foglio principale")'
#Create parameter to write sheets
		f_reptcode="ODAG_riep"
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"Rpt000hdr","rpt_dftTitle":f_descript,"rpt_strRow":5,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" }

#riformattare		w_column="Data scadenza";df_ODAGper[w_column]=df_ODAGper[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))
#riformattare		w_column="Data Decorrenza";df_ODAGper[w_column]=df_ODAGper[w_column].apply(lambda x: pd.to_datetime(x).strftime('%d/%m/%Y'))


		scriveFoglio_dict(writer,workbk_Out,df_ODAGper,rept_dict)
		wDim= str(df_ODAGper.shape[0])   
		logger.info('foglio dfODAG_gen ' + ' dimension ' + wDim )
		wstrBDO=int(wDim) +8 #Calculate row for following table
#-----------------------------------
#Create parameter to write sheets
		f_descript="Riepilogo Macroclassi "

		f_reptcode="ODAGMacr_riep"
		f_sheetname="RiepMacr"
		rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt000nohdr","rpt_dftTitle":f_descript,"rpt_strRow":9,"rpt_strCol":0,"rpt_environm":"dev","rpt_debug":"N" ,"rpt_prvrow":"Riepilogo per Macroclasse" }


		df_ODAGMacr=pd.merge(df_ODAGMacr,df_ODAGper[["ODAGnbr","DEC","RUP"]], on=["ODAGnbr"],how="left",suffixes=("","_y"))


		scriveFoglio_dict(writer,workbk_Out,df_ODAGMacr,rept_dict)
#scriveFoglio(writer,workbk_Out,f_descript,imported,f_descript,f_wrk_dframe,"rpt000")
#record in log dataframe lenght     
# Link to go bak to main sheet

		writer.close()

		tempoFin = timeit.default_timer()
		logger.info('end reading ' + str(df_ODAGMacr.shape[0])+" Durata processo: " + wPythProc +" " +f_reptcode +" - dim: "+str(df_ODAGMacr.memory_usage().sum())+" Durata processo: "  + str(round(tempoFin-tempoInz,3))+ " Dataframe dimensione " + str(df_ODAGMacr.memory_usage(deep=True)) )

#------------------------------------
#--------




#call function from main module
if __name__ == "__main__":
	wlist=["AllIn","AllRef"]
	#wlist=["CARDANI ANGELO MARIA"]
	wODAGnbr=[]
	#"AllDip",["AllIn"],["AllRef"],"AllRept","AllDip"
	wDict={
#"AllDip"
#"parm1":"value1",
#"parm2":"value2",
#"parm3":"value3"
}
	Rep250_rpt_gen(wlist,wODAGnbr,wDict)
	pass