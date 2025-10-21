##Write ODA Periodo

from Util_leggeDir import dir_dict
from funct.Dframe_toSheet import scriveFoglio,scriveFoglio_dict 
import pandas as pd
import numpy as np
# soluzione per scrivere Excel 
import xlsxwriter
from elab_Oda_Periodo import df_ODA_per
wDict_Tipocanone={"2023331930":"canone trimestr","2024330138":"canone trimestr","2024331660":"canone trimestr","2025330700":"canone annuale anticipato","2024331750":"a fine lavori","2025330292":"a consumo","2025330293":"a consumo","2025330294":"a consumo","2025330409":"a consumo","2025330455":"a consumo","2025330647":"a consumo","2025330833":"a consumo","2023330765":"a consumo","2024331708":"a consumo","2024331709":"a consumo","2025330382":"a consumo","2025330383":"a consumo","2025330547":"a consumo","2025330725":"a consumo","2025330726":"a consumo","2025330278":"a consumo","2025330685":"a consumo","2023330167":"a consumo","2024331506":"a consumo","2023331761":"a consumo","2024331850":"a consumo","2025330265":"a consumo","2018331021":"a consumo","2023331449":"a consumo"
                  ,"2025330024":"una tantum"
                  ,"2025331018":"una tantum"
                  ,"2024331847":"una tantum"
                  ,"2025330784":"a consumo"
                  ,"2024331089":"a consumo"
                  ,"2025330780":"canone annuale anticipato"}
#from read_verbAtt import df_RCA_Verb, df_Elab_RCA_VerbSpalm, df_RCA_Verb_Fatt
#from read_verbPass import df_VerbPassChius, df_VerbPassAtt, df_VerbPassSAL_daRep
df_ODA_per=df_ODA_per[df_ODA_per["Periodo"]>="2025-01"]

#print (dir_dict["rptDir"])
wrk_RptDir=dir_dict["rptDir"]
wrk_CsvDir=dir_dict["xlsDir"]
wrk_stagDir=dir_dict["stagDir"]



	#----- Impostazione formati delle celle
	# Imposta Formato currency
	#money_fmt = workbk_Out.add_format({'num_format': '#,##0.00'})

from read_Interni import df_RisInt_byName,df_RisInt_byCID
#List of ODA to remove
wListODA_remove = ["2023331804", "2023330204","2024330288","2024330289"]

filtered_ODA=df_ODA_per[(df_ODA_per["FineVal."].str.contains("05.2025|06.2025|07.2025|08.2025|09.2025")==False)&(~df_ODA_per["ODA"].isin(wListODA_remove))&(df_ODA_per["ROI_dip"]!=0)&(df_ODA_per["Val_Resid"]>10)]

df_col_Vpivot_ODA=filtered_ODA.fillna(0).pivot_table(index=["ROI_dip","ROI","ODA","Testo testata BDO","Forn_nome","InizVal.","FineVal.","Val_Resid"], columns=["Periodo"],values=["ValoreMAP"],
							   aggfunc="sum")
#25/09/2024		df_col_Vpivot_BDO=dfBDO_per_dip.pivot_table(index=['Dip'],  
#25/09/2024			values=['BDO'] ,
#25/09/2024			columns=['PeriodoMMM'],
#25/09/2024			aggfunc=pd.Series.nunique, margins=True, margins_name="#Totale#" )
df_col_Vpivot_ODA.columns = df_col_Vpivot_ODA.columns.droplevel(0) #remove amount			df_col_Vpivot_BDO.columns.name = None               #remove categories
#		df_col_Vpivot_BDO.drop(columns=df.columns[0], axis=1,  inplace=True)
df_col_Vpivot_ODA=df_col_Vpivot_ODA.reset_index().rename_axis(None,axis=1)

#df_col_Vpivot_ODA=df_col_Vpivot_ODA[(df_col_Vpivot_ODA["2025-01"]>0)|(df_col_Vpivot_ODA["2025-02"]>0)]
#Multiple conditions to check for "tipologia"
condition1=(df_col_Vpivot_ODA["2025-05"].round(1) == df_col_Vpivot_ODA["2025-06"].round(1))&(df_col_Vpivot_ODA["2025-06"].round(1) == df_col_Vpivot_ODA["2025-07"].round(1))&((df_col_Vpivot_ODA["2025-06"] >0))
condition2=((df_col_Vpivot_ODA["2025-05"].round(0) == df_col_Vpivot_ODA["2025-07"].round(0))|(df_col_Vpivot_ODA["2025-07"].round(0) == df_col_Vpivot_ODA["2025-05"].round(0)))&(df_col_Vpivot_ODA["2025-07"] >df_col_Vpivot_ODA["2025-06"])&(df_col_Vpivot_ODA["2025-07"] >df_col_Vpivot_ODA["2025-08"])
condition3=(df_col_Vpivot_ODA["2025-05"] > 0)&(df_col_Vpivot_ODA["2025-06"] >0)&(df_col_Vpivot_ODA["2025-07"] >0)


#Get column tipologia
df_col_Vpivot_ODA["tipologia"]=np.where(condition1, "canone cost.",np.where(condition2,"rateo",np.where(condition3,"canone var.","tbd")))
df_col_Vpivot_ODA["tipologiaRivista"] = df_col_Vpivot_ODA["ODA"].map(wDict_Tipocanone)

df_col_Vpivot_ODA["tipologia"]=np.where(df_col_Vpivot_ODA["tipologiaRivista"].notnull(), df_col_Vpivot_ODA['tipologiaRivista'], df_col_Vpivot_ODA['tipologia'])
df_col_Vpivot_ODA=pd.merge(df_col_Vpivot_ODA,df_RisInt_byName[["Cognome Nome","PMO","E-mail"]],left_on="ROI", right_on="Cognome Nome",how="left",suffixes=('', '_pos'))
df_col_Vpivot_ODA=pd.merge(df_col_Vpivot_ODA,df_RisInt_byCID[["CID","E-mail"]],left_on="PMO", right_on="CID",how="left",suffixes=('', '_pmo'))

wrenMAPColumns={"E-mail":"mail","E-mail_pmo":"pmo mail"}
df_col_Vpivot_ODA=df_col_Vpivot_ODA.rename(columns=wrenMAPColumns)
df_col_Vpivot_ODA.drop(['Cognome Nome','CID','PMO'], axis=1, inplace=True)
#df_col_Vpivot_ODA["tipologia"]=np.where((df_col_Vpivot_ODA["2025-01"] == df_col_Vpivot_ODA["2025-02"])&(df_col_Vpivot_ODA["2025-02"] == df_col_Vpivot_ODA["2025-03"]), 'canone costante', np.where((df_col_Vpivot_ODA["2025-01"] == df_col_Vpivot_ODA["2025-03"])&(df_col_Vpivot_ODA["2025-02"] >0)),"rateo", np.where((df_col_Vpivot_ODA["2025-01"] > 0)&(df_col_Vpivot_ODA["2025-02"] >0)&(df_col_Vpivot_ODA["2025-03"] >0),"canone variabile", "tbd"))    

#wRept="PMREP130_ODA_periodo" ;wRept_name=wRept+"_ALL" 
# Create a Pandas Excel writer using XlsxWriter as the engine.
#writer = pd.ExcelWriter(wrk_RptDir+"\\"+wRept_name+".xlsx", engine='xlsxwriter')


f_reptname="PMREP130_ODA_periodo_ALL" #"PMRep117-test"
wRept_output=wrk_RptDir+"\\"+f_reptname+".xlsx"   
	#verifica se esiste il file e se esiste lo rimuove

writer=pd.ExcelWriter(wRept_output, engine="xlsxwriter")	
workbk_Out=writer.book

f_descript="Riepilogo ODA consuntivati / da consuntivare" #*** Personalizzare report (punto 1)
f_reptcode="Riep_ODAPeriodo" #*** Personalizzare report (punto 2)	
f_sheetname="Riep_ODAPeriodo" #*** Personalizzare report (punto 3)
#	logger.info('start process for ' + f_reptcode + f_descript )
rept_dict={"rpt_code":f_reptcode,"SheetName":f_sheetname,"rpt_Type":"rpt001","rpt_dftTitle":f_descript,"rpt_strRow":6,"rpt_strCol":0,"rpt_environm":"Prod" }
scriveFoglio_dict(writer,workbk_Out,df_col_Vpivot_ODA,rept_dict) 	

writer.close()

# Convert the dataframe to an XlsxWriter Excel object.
#Write Excel file
#df_col_Vpivot_ODA.to_excel(writer, sheet_name='Foglio1', index=False)
"""workbook=writer.book
worksheet = writer.sheets["Foglio1"]
	# Imposta formato intestazione colonne (sfondo giallo)
bckgrdYellow_fmt = workbook.add_format()
bckgrdYellow_fmt.set_border(True)  #	wsheader_format.set_align("center")
bckgrdYellow_fmt.set_bg_color("yellow")
bckgrdYellow_fmt.set_font_color("#3d1ec7")
# Imposta formato intestazione colonne (sfondo verde)
bckgrdGreen_fmt = workbook.add_format()
bckgrdGreen_fmt.set_bg_color("#42f58d")
bckgrdGreen_fmt.set_border(2)
bckgrdGreen_fmt.set_underline()

money_fmt = workbook.add_format({'num_format': '#,##0.00'})
date_fmt = workbook.add_format({'num_format': 'dd/mm/yyyy'})

worksheet.set_column(2,2,10)
worksheet.set_column(3,4,25)
worksheet.set_column(5,6,10)
worksheet.set_column(1,1,25)
#worksheet.set_column(6,7,25)
worksheet.set_column(7,17,14,money_fmt)
worksheet.set_column(18,19,15)
worksheet.set_column(20,21,25)

#workbk_Out,worksh_Out,work_dframe,rpt_Code,rpt_Title,rpt_format)
# Get the dimensions of the dataframe.
(max_row, max_col) = df_col_Vpivot_ODA.shape
# Make the columns wider for clarity.
#worksheet.set_column(0, max_col - 1, 12)
# Set the autofilter.
worksheet.autofilter(0, 0, max_row, max_col - 1)

writer.close()
"""
