#Report to create Excel sheet 
#Rpt000 = no Header
#Rpt000hdr = Header
#Rpt000- = no Header and title on previous row
#Rpt000-nofmt = no Header, title on previous row and column format
##Rpt00no0fmt = no Header and no column format
#Rpt001 = Dataframe with Header and Table
#Rpt001- = Dataframe with Header and Table title on previous row
from Util_leggeDir import dir_dict
import pandas as pd
#from Util_ReportDate import GetRptDate,GetAlfb
	#from Funz.Util_wDir import RepInit0	#Per la scrittura su Excel si usa la libreria xlsxwriter (limite che non si possono leggere file Excel e non si possono modificare con quella libreria)
from .Util_RptHeader import RepTitoloData, RepColumnFormat
from Util_leggeParams import df_rptFormats,df_titles,df_columnType,parm_dict

def scriveFoglio_dict_v5(wrk_wrtName,wrk_Book,wrk_df,wrept_dict):  #,wrk_format,wrk_colformat
	wFmtFile="RptFormati_v5"
#	wSh_repTitoli="RepTitoli"
	#Default values and Dictionary values
	w_reptShname= wrept_dict.get("SheetName", "Sheet1")
	w_reptCode= wrept_dict.get("rpt_code", "NoreportCode")
	w_reptFmtCode= wrept_dict.get("rpt_Type", "rpt000")
	w_reptdftTitle= wrept_dict.get("rpt_dftTitle", "Titolo report di default")
	w_reptstrRow= wrept_dict.get("rpt_strRow", 0)
	w_reptstrCol= wrept_dict.get("rpt_strCol", 0)
	w_reptaddLayout= wrept_dict.get("rpt_addLayout", False)
	w_reptDebug= wrept_dict.get("rpt_debug", "N")
	w_reptenvironm= wrept_dict.get("rpt_environm", " -test")
	w_reptprvrow=wrept_dict.get("rpt_prvrow", "")
	# Imposta formato intestazione colonne (sfondo giallo)
	bckgrdYellow_fmt = wrk_Book.add_format()
	bckgrdYellow_fmt.set_border(True)  #	wsheader_format.set_align("center")
	bckgrdYellow_fmt.set_bg_color("yellow")
	bckgrdYellow_fmt.set_font_color("#3d1ec7")
	# Imposta formato intestazione colonne (sfondo verde)
	bckgrdGreen_fmt = wrk_Book.add_format()
	bckgrdGreen_fmt.set_bg_color("#42f58d")
	bckgrdGreen_fmt.set_border(2)
	bckgrdGreen_fmt.set_underline()

	money_fmt = wrk_Book.add_format({'num_format': '#,##0.00'})
	date_fmt = wrk_Book.add_format({'num_format': 'dd/mm/yy'})

	wrk_CsvDir=dir_dict["csvDir"]
	df_columnType.drop_duplicates(subset=['ColumnName'],inplace=True)
### temporary instruction to remove duplicate columns 202508
	wrk_df = wrk_df.loc[:,~wrk_df.columns.duplicated()].copy()
### end temporary instruction
	#if debug create excel with all dataframe columns
	if w_reptDebug=="Y":
			#create list of all column names
		index_list = wrk_df.columns.values.tolist()
		#Create dataframe for column sequences with all df columns 
		df_colseq_all=pd.DataFrame(index_list, columns=['ColumnName'])
		#Column alias not available for full xls print
		df_colseq_all['ColumnAlias']=df_colseq_all['ColumnName']
		df_colseq_all=pd.merge(df_colseq_all,df_columnType,left_on="ColumnAlias", right_on="ColumnName",how="left",suffixes=('', '_colA') )
		df_colseq_all = df_colseq_all.loc[df_colseq_all["ColumnName"].notna()]	
#merge df_colseq with df_columnType by columnName and ColumnAlias
		df_colseq_all['ColFormat'] = df_colseq_all['ColFormat'].fillna(10)
		df_colseq_all['ColumnName_colA'] = df_colseq_all['ColumnName_colA'].fillna(df_colseq_all['ColumnAlias'])
#drop nan rows (means that field do not match with the one in pattern in excel given)
	#    from openpyxl import load_workbook
		wRept_output=wrk_CsvDir+"\\"+w_reptCode+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

		writer1=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
		workbk1_Out=writer1.book
		money1_fmt = workbk1_Out.add_format({'num_format': '#,##0.00'})
		date1_fmt = workbk1_Out.add_format({'num_format': 'dd/mm/yy'})

## Write excel Backup for test purpose
		wrk_df.to_excel(writer1, sheet_name=w_reptShname, index=False, startrow=0, startcol=0)	
		worksh1_Out=writer1.sheets[w_reptShname]
#Get column size for formatting in a list
		wrk_col_size = df_colseq_all['ColFormat'].values.tolist()
# Get the dimensions of the dataframe.
		w_reptstrCol=0
		(max_row, max_col) = wrk_df.shape
		for i in range(len(wrk_col_size)):
			wcol_size=wrk_col_size[i]
			if isinstance(wcol_size, int):
				worksh1_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, wcol_size)
				pass
			elif wcol_size=="ccy14":
				worksh1_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 14,money1_fmt)
				pass
			elif wcol_size=="ccyL":
				worksh1_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 13,money1_fmt)
			elif wcol_size=="ccy":
				worksh1_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 12,money1_fmt)
			elif wcol_size=="date":
				worksh1_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 10,date1_fmt)
			else:

				pass
		df_dict = pd.DataFrame(data=wrept_dict, index=[0])
		df_dict =(df_dict.T)
		writer1.close()
		df_dict.to_excel(wrk_CsvDir+"\\"+w_reptCode+"_dict.xlsx", sheet_name="df_dict", startrow=0) #, index=False




	#df_columnType = pd.read_excel(f, sheet_name='Formato')
#	df_colseq= pd.read_excel(f, sheet_name=w_reptCode)
	df_colseq=df_rptFormats.loc[df_rptFormats["NomeFoglio"]==w_reptCode]
#	df_colseq=df_rptFormats.loc[df_rptFormats["NomeFoglio"]==w_reptCode]

	#if Foglio excel not exist add in Excel
	if df_colseq.empty:

			#create list of all column names
		index_list = wrk_df.columns.values.tolist()
		#Create dataframe for column sequences with all df columns 
		df_colseq_all=pd.DataFrame(index_list, columns=['ColumnName'])
		if w_reptaddLayout:
			import openpyxl
#    	wrk_CsvDir=dir_dict["csvDir"]
			wFmtFile="RptFormati_v5"
			wBook = openpyxl.load_workbook(wrk_CsvDir+"\\"+wFmtFile+".xlsx")
			wBook.create_sheet(w_reptCode)
			wSheet=wBook[w_reptCode]
			wSheet.cell(1,1).value ="ColumnName"
			wSheet.cell(1,2).value ="ColumnAlias"
			wSheet.cell(1,3).value ="Riepilogo"
			i=2
			for item in index_list:
			#wSheet.append(item)
				wSheet.cell(i,1).value = item
				i=i+1	
			wBook.save(wrk_CsvDir+"\\"+wFmtFile+".xlsx")
		#if dataframe for reportcode is empty use all field for selection
		df_colseq=df_colseq_all
		#add new rows with format on main dataframe

#		df_titles=pd.read_excel(f, sheet_name=wSh_repTitoli)
#get Column ReptFormat
##2	df_columnType = pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name='Formato')
#get data frame by rept name    
##2	df_colseq= pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name=w_reptCode)
#get data frame titoli per aggiornare il titolo da impiegare per i fogli Excel   
##2	df_titles=pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name=wSh_repTitoli)
#Convert dataframe to dictionary
	dict_titles=dict(df_titles.values)
	wRept_Title= dict_titles.get(w_reptCode, w_reptdftTitle)
#### Temporar

	df_colseq=pd.merge(df_colseq,df_columnType,left_on="ColumnName", right_on="ColumnName",how="left",suffixes=('', '_colN') )
	df_colseq['ColumnAlias'] = df_colseq['ColumnAlias'].fillna(df_colseq['ColumnName'])
	df_colseq=pd.merge(df_colseq,df_columnType,left_on="ColumnAlias", right_on="ColumnName",how="left",suffixes=('', '_colA') )

#merge df_colseq with df_columnType by columnName and ColumnAlias
	df_colseq['ColFormat'] = df_colseq['ColFormat'].fillna(10)
	df_colseq['ColumnName_colA'] = df_colseq['ColumnName_colA'].fillna(df_colseq['ColumnAlias'])
	df_colseq['ColFormat_colA'] = df_colseq['ColFormat_colA'].fillna(df_colseq['ColFormat'])
#drop nan rows (means that field do not match with the one in pattern in excel given)
	df_colseq = df_colseq.loc[df_colseq["ColumnName_colA"].notna()]	

#Rename dataframe columns
#from dataframe to dataframe, replacing new lables
	wrk_df=wrk_df.rename(columns=dict(zip(df_colseq["ColumnName"], df_colseq["ColumnAlias"])))
	df_colseq['Column_nbr'] = range(0, len(df_colseq))
	wrk_df_colseq=df_colseq.loc[df_colseq["Riepilogo"].notnull()]
	if not wrk_df_colseq.empty:
		wrk_df_colseq_tot=wrk_df_colseq.loc[wrk_df_colseq["Riepilogo"].str.startswith(("t","T"))]
		wrk_df_colseq_count=wrk_df_colseq.loc[wrk_df_colseq["Riepilogo"].str.startswith(("c","C"))]
	else:
		wrk_df_colseq_tot=wrk_df_colseq; wrk_df_colseq_count=wrk_df_colseq
#Get list for Index
#extract to list element in ColumnAlias for index
	wrk_label_index = df_colseq['ColumnAlias'].values.tolist()
#20250922 Change from ColFormat_colA to ColFormat
#Get column size for formatting in a list
	wrk_col_size = df_colseq['ColFormat_colA'].values.tolist()


	wReindex=wrk_label_index
	wReindex2 = [x for x in wReindex if x in wrk_df.columns]


	if w_reptDebug=="Y":
## Write excel Backup for test purpose
		wrk_dfcolOrder = pd.DataFrame(wReindex2, columns=["Column_Name"]) 
		wrk_dfcolOrder.to_excel(wrk_CsvDir+"\\"+w_reptCode+"colOrder.xlsx", sheet_name=w_reptShname, index=False, startrow=0, startcol=0)	

	#if not wrk_df.empty:
		#Ordina le colonne
	wrk_df_sort= wrk_df.sort_values(by=wReindex2)
		#dfsort.rename(columns={'BDO_ChiusAmm': 'ChiusAmm', 'Verb_SAL_Descr': 'SAL_stato', 'Verb_Attiv_Descr': 'VerbAtt', 'Verb_Chius_Descr': 'VerbChius,"Check_Cons":"ConsOK'}, inplace=True)
	wrk_df_sort=wrk_df_sort.reindex(columns=wReindex)
#		print(dfsort.head(10))
		#Predispone il Data Frame per la scrittura su Excel
	wrk_df_sort.to_excel(wrk_wrtName, sheet_name=w_reptShname, index=False, startrow=w_reptstrRow, startcol=w_reptstrCol)	


	#wrk_df.to_excel(wrk_wrtName, sheet_name=wrk_shtName, index=False, startrow=5, startcol=0)	
		#Prepara worksheet per Periodo_BDO
	worksh_Out=wrk_wrtName.sheets[w_reptShname]
	if wrk_df.empty: worksh_Out.write(w_reptstrRow+1,w_reptstrCol ,"Non sono presenti elementi per "+ w_reptCode)
	if ("rpt001" in w_reptFmtCode.lower() and w_reptstrRow>0) or "hdr" in w_reptFmtCode.lower():
		#Scrittura intestazione del foglio
		RepTitoloData(wrk_Book,worksh_Out,wRept_Title+" "+w_reptenvironm)

#-----Scrive titolo
	if "rpt001" in w_reptFmtCode.lower():

# Get the dimensions of the dataframe.
		(max_row, max_col) = wrk_df_sort.shape
		column_settings = [{"header": column} for column in wrk_df_sort]
#create table name and subtotals required
		if max_row > 0:
			w_tabName=w_reptShname.replace("-","_")
	#Create Excel table based on dframe; columns in order as sheet field; Create table as wrk_reptCode name
			worksh_Out.add_table(w_reptstrRow, w_reptstrCol, max_row+w_reptstrRow, max_col - 1+w_reptstrCol, {"columns": column_settings,"name": w_tabName})
	#if dataframe not empty create totals / subtotals 
			if not wrk_df_colseq_tot.empty:
				for index,row in wrk_df_colseq_tot.iterrows():
					worksh_Out.write(3,row['Column_nbr']+ w_reptstrCol,"Tot"+row['ColumnAlias'])
					wFormula = "=Subtotal(9,"+w_tabName+"["+row['ColumnAlias']+"])"     
					worksh_Out.write_formula(4,row['Column_nbr']+ w_reptstrCol ,wFormula)   
			if not wrk_df_colseq_count.empty:
				for index,row in wrk_df_colseq_count.iterrows():
					worksh_Out.write(3,row['Column_nbr']+ w_reptstrCol ,"n."+row['ColumnAlias'])
					wFormula = "=Subtotal(3,"+w_tabName+"["+row['ColumnAlias']+"])"     
					worksh_Out.write_formula(4,row['Column_nbr']+ w_reptstrCol ,wFormula)   
#-----Formatting header
	#if report format contains "-" write title in previous row 
	if "-" in w_reptFmtCode.lower():worksh_Out.write(w_reptstrRow-1,w_reptstrCol,w_reptprvrow)

	if "rpt000" in w_reptFmtCode.lower():
		for col_num, value in enumerate(wrk_df_sort.columns.values):
			worksh_Out.write(w_reptstrRow, w_reptstrCol+col_num, value, bckgrdGreen_fmt)
#	worksh_Out.autofilter(wrpt_row,0,wrk_df_sort.shape[0],wrk_df_sort.shape[1])
#-----Format Columns
	if "nofmt" not in w_reptFmtCode.lower():
		for i in range(len(wrk_col_size)):
			wcol_size=wrk_col_size[i]
			if isinstance(wcol_size, int):
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, wcol_size)
				pass
			elif wcol_size=="ccy14":
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 14,money_fmt)
				pass
			elif wcol_size=="ccyL":
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 13,money_fmt)
			elif wcol_size=="ccy":
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 12,money_fmt)
			elif wcol_size=="date":
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 10,date_fmt)
			else:
				pass
	#	worksh_Out.add_table(wrpt_row, 0, max_row, max_col - 1
#                            , {'columns': wrk_label_index
#                            ,"name": "wrk_reptCode"}
#                            )	
	pass


def scriveFoglio_dict(wrk_wrtName,wrk_Book,wrk_df,wrept_dict):
	wFmtFile="RptFormati"
	wSh_repTitoli="RepTitoli"
	#Default values and Dictionary values
	w_reptShname= wrept_dict.get("SheetName", "Sheet1")
	w_reptCode= wrept_dict.get("rpt_code", "NoreportCode")
	w_reptFmtCode= wrept_dict.get("rpt_Type", "rpt000")
	w_reptdftTitle= wrept_dict.get("rpt_dftTitle", "Titolo report di default")
	w_reptstrRow= wrept_dict.get("rpt_strRow", 0)
	w_reptstrCol= wrept_dict.get("rpt_strCol", 0)
	w_reptDebug= wrept_dict.get("rpt_debug", "N")
	w_reptenvironm= wrept_dict.get("rpt_environm", " -test")
	w_reptprvrow=wrept_dict.get("rpt_prvrow", "")
#	w_reptemplShname=wrept_dict.get("rpt_templShname", w_reptShname)
	#w_shName = wrept_dict["SheetName"] if "SheetName" in wrept_dict else "Sheet1"
	#w_reptCode = wrept_dict["rpt_code"] if "rpt_code" in wrept_dict else "NoreportCode"
	#w_reptFmtCode = wrept_dict["rpt_Type"] if "rpt_Type" in wrept_dict else "rpt000"
	#w_reptdftTitle = wrept_dict["rpt_dftTitle"] if "rpt_dftTitle" in wrept_dict else "Titolo report di default"
	#w_reptstrRow = wrept_dict["rpt_strRow"] if "rpt_strRow" in wrept_dict else 0
	#w_reptstrCol = wrept_dict["rpt_strCol"] if "rpt_strCol" in wrept_dict else 0
	#w_reptenvironm = wrept_dict["rpt_environm"] if "rpt_environm" in wrept_dict else "-"

	#on type report selevt option(s) for different layout
	#if w_reptFmtCode=="rpt001":
	#	w_reptstrRow=5;w_reptstrCol=0
	#elif wrk_reptFmtCode=="rpt000":
	#	wrpt_row=0;wrpt_col=0
	#else:
	#	wrpt_row=5;wrpt_col=0
	#----- Impostazione formati delle celle
	# Imposta Formato currency
	#money_fmt = workbk_Out.add_format({'num_format': '#,##0.00'})
	# Imposta formato intestazione colonne (sfondo giallo)
	bckgrdYellow_fmt = wrk_Book.add_format()
	bckgrdYellow_fmt.set_border(True)  #	wsheader_format.set_align("center")
	bckgrdYellow_fmt.set_bg_color("yellow")
	bckgrdYellow_fmt.set_font_color("#3d1ec7")
	# Imposta formato intestazione colonne (sfondo verde)
	bckgrdGreen_fmt = wrk_Book.add_format()
	bckgrdGreen_fmt.set_bg_color("#42f58d")
	bckgrdGreen_fmt.set_border(2)
	bckgrdGreen_fmt.set_underline()

	money_fmt = wrk_Book.add_format({'num_format': '#,##0.00'})
	date_fmt = wrk_Book.add_format({'num_format': 'dd/mm/yy'})

	wrk_CsvDir=dir_dict["csvDir"]

	if w_reptDebug=="Y":
## Write excel Backup for test purpose
		wrk_df.to_excel(wrk_CsvDir+"\\"+w_reptCode+".xlsx", sheet_name=w_reptShname, index=False, startrow=0, startcol=0)	
		df_dict = pd.DataFrame(data=wrept_dict, index=[0])
		df_dict =(df_dict.T)
		df_dict.to_excel(wrk_CsvDir+"\\"+w_reptCode+"_dict.xlsx", sheet_name="df_dict", startrow=0) #, index=False

	with open(wrk_CsvDir+"\\"+wFmtFile+".xlsx", "rb") as f:
		df_columnType = pd.read_excel(f, sheet_name='Formato')
		df_colseq= pd.read_excel(f, sheet_name=w_reptCode)
		df_titles=pd.read_excel(f, sheet_name=wSh_repTitoli)
#get Column ReptFormat
##2	df_columnType = pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name='Formato')
#get data frame by rept name    
##2	df_colseq= pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name=w_reptCode)
#get data frame titoli per aggiornare il titolo da impiegare per i fogli Excel   
##2	df_titles=pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name=wSh_repTitoli)
#Convert dataframe to dictionary
	dict_titles=dict(df_titles.values)
	wRept_Title= dict_titles.get(w_reptCode, w_reptdftTitle)

	df_columnType.drop_duplicates(subset=['ColumnName'],inplace=True)

	df_colseq=pd.merge(df_colseq,df_columnType,left_on="ColumnName", right_on="ColumnName",how="left",suffixes=('', '_colN') )
	df_colseq['ColumnAlias'] = df_colseq['ColumnAlias'].fillna(df_colseq['ColumnName'])
	df_colseq=pd.merge(df_colseq,df_columnType,left_on="ColumnAlias", right_on="ColumnName",how="left",suffixes=('', '_colA') )

#merge df_colseq with df_columnType by columnName and ColumnAlias
	df_colseq['ColFormat'] = df_colseq['ColFormat'].fillna(10)
	df_colseq['ColFormat_colA'] = df_colseq['ColFormat_colA'].fillna(df_colseq['ColFormat'])
#drop nan rows (means that field do not match with the one in pattern in excel given)
	df_colseq = df_colseq[df_colseq["ColumnName"].notna()]	

#Rename dataframe columns
#from dataframe to dataframe, replacing new lables
	wrk_df=wrk_df.rename(columns=dict(zip(df_colseq["ColumnName"], df_colseq["ColumnAlias"])))
	df_colseq['Column_nbr'] = range(0, len(df_colseq))
	wrk_df_colseq=df_colseq[df_colseq["Riepilogo"].notnull()]
	if not wrk_df_colseq.empty:
		wrk_df_colseq_tot=wrk_df_colseq[wrk_df_colseq["Riepilogo"].str.startswith(("t","T"))]
		wrk_df_colseq_count=wrk_df_colseq[wrk_df_colseq["Riepilogo"].str.startswith(("c","C"))]
	else:
		wrk_df_colseq_tot=wrk_df_colseq; wrk_df_colseq_count=wrk_df_colseq
#Get list for Index
#extract to list element in ColumnAlias for index
	wrk_label_index = df_colseq['ColumnAlias'].values.tolist()

#Get column size for formatting in a list
	wrk_col_size = df_colseq['ColFormat_colA'].values.tolist()


	wReindex=wrk_label_index
	wReindex2 = [x for x in wReindex if x in wrk_df.columns]


	if w_reptDebug=="Y":
## Write excel Backup for test purpose
		df = pd.DataFrame(wReindex2, columns=["Column_Name"]) 
		wrk_df.to_excel(wrk_CsvDir+"\\"+w_reptCode+"colOrder.xlsx", sheet_name=w_reptShname, index=False, startrow=0, startcol=0)	

	#if not wrk_df.empty:
		#Ordina le colonne
	wrk_df_sort= wrk_df.sort_values(by=wReindex2)
		#dfsort.rename(columns={'BDO_ChiusAmm': 'ChiusAmm', 'Verb_SAL_Descr': 'SAL_stato', 'Verb_Attiv_Descr': 'VerbAtt', 'Verb_Chius_Descr': 'VerbChius,"Check_Cons":"ConsOK'}, inplace=True)
	wrk_df_sort=wrk_df_sort.reindex(columns=wReindex)
#		print(dfsort.head(10))
		#Predispone il Data Frame per la scrittura su Excel
	wrk_df_sort.to_excel(wrk_wrtName, sheet_name=w_reptShname, index=False, startrow=w_reptstrRow, startcol=w_reptstrCol)	


	#wrk_df.to_excel(wrk_wrtName, sheet_name=wrk_shtName, index=False, startrow=5, startcol=0)	
		#Prepara worksheet per Periodo_BDO
	worksh_Out=wrk_wrtName.sheets[w_reptShname]
	if wrk_df.empty: worksh_Out.write(w_reptstrRow+1,w_reptstrCol ,"Non sono presenti elementi per "+ w_reptCode)
	if ("rpt001" in w_reptFmtCode.lower() and w_reptstrRow>0) or "hdr" in w_reptFmtCode.lower():
		#Scrittura intestazione del foglio
		RepTitoloData(wrk_Book,worksh_Out,wRept_Title+" "+w_reptenvironm)

#-----Scrive titolo
	if "rpt001" in w_reptFmtCode.lower():

# Get the dimensions of the dataframe.
		(max_row, max_col) = wrk_df_sort.shape
		column_settings = [{"header": column} for column in wrk_df_sort]
#create table name and subtotals required
		if max_row > 0:
			w_tabName=w_reptShname.replace("-","_")
	#Create Excel table based on dframe; columns in order as sheet field; Create table as wrk_reptCode name
			worksh_Out.add_table(w_reptstrRow, w_reptstrCol, max_row+w_reptstrRow, max_col - 1+w_reptstrCol, {"columns": column_settings,"name": w_tabName})
	#if dataframe not empty create totals / subtotals 
			if not wrk_df_colseq_tot.empty:
				for index,row in wrk_df_colseq_tot.iterrows():
					worksh_Out.write(3,row['Column_nbr']+ w_reptstrCol,"Tot"+row['ColumnAlias'])
					wFormula = "=Subtotal(9,"+w_tabName+"["+row['ColumnAlias']+"])"     
					worksh_Out.write_formula(4,row['Column_nbr']+ w_reptstrCol ,wFormula)   
			if not wrk_df_colseq_count.empty:
				for index,row in wrk_df_colseq_count.iterrows():
					worksh_Out.write(3,row['Column_nbr']+ w_reptstrCol ,"n."+row['ColumnAlias'])
					wFormula = "=Subtotal(3,"+w_tabName+"["+row['ColumnAlias']+"])"     
					worksh_Out.write_formula(4,row['Column_nbr']+ w_reptstrCol ,wFormula)   
#-----Formatting header
	#if report format contains "-" write title in previous row 
	if "-" in w_reptFmtCode.lower():worksh_Out.write(w_reptstrRow-1,w_reptstrCol,w_reptprvrow)

	if "rpt000" in w_reptFmtCode.lower():
		for col_num, value in enumerate(wrk_df_sort.columns.values):
			worksh_Out.write(w_reptstrRow, w_reptstrCol+col_num, value, bckgrdGreen_fmt)
#	worksh_Out.autofilter(wrpt_row,0,wrk_df_sort.shape[0],wrk_df_sort.shape[1])
#-----Format Columns
	if "nofmt" not in w_reptFmtCode.lower():
		for i in range(len(wrk_col_size)):
			wcol_size=wrk_col_size[i]
			if isinstance(wcol_size, int):
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, wcol_size)
				pass
			elif wcol_size=="ccy14":
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 14,money_fmt)
				pass
			elif wcol_size=="ccyL":
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 13,money_fmt)
			elif wcol_size=="ccy":
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 12,money_fmt)
			elif wcol_size=="date":
				worksh_Out.set_column(i+ w_reptstrCol, i+ w_reptstrCol, 10,date_fmt)
			else:
				pass
	#	worksh_Out.add_table(wrpt_row, 0, max_row, max_col - 1
#                            , {'columns': wrk_label_index
#                            ,"name": "wrk_reptCode"}
#                            )	
	pass

#### OLD function
def scriveFoglio(wrk_wrtName,wrk_Book,wrk_shtName,wrk_df,wrk_reptName,wrk_reptCode,wrk_reptFmtCode):
	wFmtFile="RptFormati"
	wSh_repTitoli="RepTitoli"
	#on type report selevt option(s) for different layout
	if wrk_reptFmtCode=="rpt001":
		wrpt_row=5;wrpt_col=0
	elif wrk_reptFmtCode=="rpt000":
		wrpt_row=0;wrpt_col=0
	else:
		wrpt_row=5;wrpt_col=0
	#----- Impostazione formati delle celle
	# Imposta Formato currency
	#money_fmt = workbk_Out.add_format({'num_format': '#,##0.00'})
	# Imposta formato intestazione colonne (sfondo giallo)
	bckgrdYellow_fmt = wrk_Book.add_format()
	bckgrdYellow_fmt.set_border(True)  #	wsheader_format.set_align("center")
	bckgrdYellow_fmt.set_bg_color("yellow")
	bckgrdYellow_fmt.set_font_color("#3d1ec7")
	# Imposta formato intestazione colonne (sfondo verde)
	bckgrdGreen_fmt = wrk_Book.add_format()
	bckgrdGreen_fmt.set_bg_color("#42f58d")
	bckgrdGreen_fmt.set_border(2)
	bckgrdGreen_fmt.set_underline()

	money_fmt = wrk_Book.add_format({'num_format': '#,##0.00'})
	date_fmt = wrk_Book.add_format({'num_format': 'dd/mm/yyyy'})

	wrk_CsvDir=dir_dict["csvDir"]
#get Column ReptFormat
	df_columnType = pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name='Formato')
#get data frame by rept name    
	df_colseq= pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name=wrk_reptCode)
#get data frame titoli    
	df_titles=pd.read_excel(wrk_CsvDir+"\\"+wFmtFile+".xlsx", sheet_name=wSh_repTitoli)
#Convert dataframe to dictionary
	dict_titles=dict(df_titles.values)
	wRept_Title= dict_titles.get(wrk_reptCode, wrk_reptName)

	df_columnType.drop_duplicates(subset=['ColumnName'],inplace=True)

	df_colseq=pd.merge(df_colseq,df_columnType,left_on="ColumnName", right_on="ColumnName",how="left",suffixes=('', '_colN') )
	df_colseq['ColumnAlias'] = df_colseq['ColumnAlias'].fillna(df_colseq['ColumnName'])
	df_colseq=pd.merge(df_colseq,df_columnType,left_on="ColumnAlias", right_on="ColumnName",how="left",suffixes=('', '_colA') )

#merge df_colseq with df_columnType by columnName and ColumnAlias
	df_colseq['ColFormat'] = df_colseq['ColFormat'].fillna(10)
	df_colseq['ColFormat_colA'] = df_colseq['ColFormat_colA'].fillna(df_colseq['ColFormat'])
	

#Rename dataframe columns
#from dataframe to dataframe, replacing new lables
	wrk_df=wrk_df.rename(columns=dict(zip(df_colseq["ColumnName"], df_colseq["ColumnAlias"])))
	df_colseq['Column_nbr'] = range(0, len(df_colseq))
	wrk_df_colseq=df_colseq[df_colseq["Riepilogo"].notnull()]
	if not wrk_df_colseq.empty:
		wrk_df_colseq_tot=wrk_df_colseq[wrk_df_colseq["Riepilogo"].str.startswith(("t","T"))]
		wrk_df_colseq_count=wrk_df_colseq[wrk_df_colseq["Riepilogo"].str.startswith(("c","C"))]
	else:
		wrk_df_colseq_tot=wrk_df_colseq; wrk_df_colseq_count=wrk_df_colseq
#Get list for Index
#extract to list element in ColumnAlias for index
	wrk_label_index = df_colseq['ColumnAlias'].values.tolist()

#Get column size for formatting in a list
	wrk_col_size = df_colseq['ColFormat_colA'].values.tolist()


	wReindex=wrk_label_index
	wReindex2 = [x for x in wReindex if x in wrk_df.columns]
		#Ordina le colonne
	wrk_df_sort= wrk_df.sort_values(by=wReindex2)
		#dfsort.rename(columns={'BDO_ChiusAmm': 'ChiusAmm', 'Verb_SAL_Descr': 'SAL_stato', 'Verb_Attiv_Descr': 'VerbAtt', 'Verb_Chius_Descr': 'VerbChius,"Check_Cons":"ConsOK'}, inplace=True)
	wrk_df_sort=wrk_df_sort.reindex(columns=wReindex)
#		print(dfsort.head(10))
		#Predispone il Data Frame per la scrittura su Excel
	wrk_df_sort.to_excel(wrk_wrtName, sheet_name=wrk_shtName, index=False, startrow=wrpt_row, startcol=wrpt_col)	



	#wrk_df.to_excel(wrk_wrtName, sheet_name=wrk_shtName, index=False, startrow=5, startcol=0)	
		#Prepara worksheet per Periodo_BDO
	worksh_Out=wrk_wrtName.sheets[wrk_shtName]
	
#-----Scrive titolo
	if wrk_reptFmtCode!="rpt000":
		#Scrittura intestazione del foglio
		RepTitoloData(wrk_Book,worksh_Out,wRept_Title)
# Get the dimensions of the dataframe.
		(max_row, max_col) = wrk_df_sort.shape
		column_settings = [{"header": column} for column in wrk_df_sort]

	#Create Excel table based on dframe; columns in order as sheet field; Create table as wrk_reptCode name
		worksh_Out.add_table(wrpt_row, 0, max_row+wrpt_row, max_col - 1, {"columns": column_settings,"name": wrk_reptCode})

		if not wrk_df_colseq_tot.empty:
			for index,row in wrk_df_colseq_tot.iterrows():
				worksh_Out.write(2,row['Column_nbr'] ,"Tot"+row['ColumnAlias'])
				wFormula = "=Subtotal(9,"+wrk_reptCode+"["+row['ColumnAlias']+"])"     
				worksh_Out.write_formula(3,row['Column_nbr'] ,wFormula)   
		if not wrk_df_colseq_count.empty:
			for index,row in wrk_df_colseq_count.iterrows():
				worksh_Out.write(2,row['Column_nbr'] ,"n."+row['ColumnAlias'])
				wFormula = "=Subtotal(3,"+wrk_reptCode+"["+row['ColumnAlias']+"])"     
				worksh_Out.write_formula(3,row['Column_nbr'] ,wFormula)   
#-----Formatting header
#	for col_num, value in enumerate(wrk_df_sort.columns.values):
#			worksh_Out.write(wrpt_row, col_num, value, bckgrdGreen_fmt)
#	worksh_Out.autofilter(wrpt_row,0,wrk_df_sort.shape[0],wrk_df_sort.shape[1])
#-----Format Columns
	for i in range(len(wrk_col_size)):
		wcol_size=wrk_col_size[i]
		if isinstance(wcol_size, int):
			worksh_Out.set_column(i, i, wcol_size)
			pass
		elif wcol_size=="ccy14":
			worksh_Out.set_column(i, i, 14,money_fmt)
			pass
		elif wcol_size=="ccyL":
			worksh_Out.set_column(i, i, 13,money_fmt)
		elif wcol_size=="ccy":
			worksh_Out.set_column(i, i, 12,money_fmt)
		elif wcol_size=="date":
			worksh_Out.set_column(i,i, 10, date_fmt)
		else:
			pass
	#	worksh_Out.add_table(wrpt_row, 0, max_row, max_col - 1
#                            , {'columns': wrk_label_index
#                            ,"name": "wrk_reptCode"}
#                            )	
	pass
	
