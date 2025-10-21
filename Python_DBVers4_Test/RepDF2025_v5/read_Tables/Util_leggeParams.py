#Legge da un file Excel tutte le colonne presenti e li inserisce in un unico dataframe dove è possibile identificare il NomeFoglio selezionato
import pandas as pd
import os
import xlsxwriter
from Util_leggeDir import dir_dict
wrk_CsvDir=dir_dict["csvDir"]

#def Util_readParms():
workbook = pd.ExcelFile(wrk_CsvDir+"\\"+'RptFormati_v5.xlsx')
sheets = workbook.sheet_names

with open(wrk_CsvDir+"\\"+'RptFormati_v5'+".xlsx", "rb") as f:
    dfparmcol_alls = pd.concat([pd.read_excel(f, sheet_name=s)
        .assign(NomeFoglio=s) for s in sheets])
    #dfparmcol_alls = pd.concat([pd.read_excel(wrk_CsvDir+"\\"+'RptFormati.xlsx', sheet_name=s)
    #            .assign(NomeFoglio=s) for s in sheets])

df_rptFormats=dfparmcol_alls.loc[dfparmcol_alls["ColumnName"].notnull(),["NomeFoglio","ColumnName","ColumnAlias","Riepilogo"]]
df_titles=dfparmcol_alls.loc[dfparmcol_alls["NomeFoglio"]=="RepTitoli",["ReptName",	"ReptTitle"] ]
df_columnType=dfparmcol_alls.loc[dfparmcol_alls["NomeFoglio"]=="Formato",["ColumnName","ColFormat"]]
df_rptParmValues=dfparmcol_alls.loc[dfparmcol_alls["NomeFoglio"]=="ParmValues",["Variable", "Value"]]
parm_dict=dict(zip(df_rptParmValues['Variable'], df_rptParmValues['Value']))

with open(wrk_CsvDir+"\\"+'RptSelezioni_v5'+".xlsx", "rb") as f:
    dffilter = pd.read_excel(f, sheet_name="Test_select")
listfilter_ODAG=dffilter.loc[dffilter['ODAG'].notnull(),"ODAG"].tolist()
listfilter_MAP=dffilter.loc[dffilter['MAP'].notnull(),"MAP"].tolist()
listfilter_Task=dffilter.loc[dffilter['Task'].notnull(),"Task"].tolist()
listfilter_BDO=dffilter.loc[dffilter['BDO'].notnull(),"BDO"].tolist()
listfilter_Incarico=dffilter.loc[dffilter['Incarico'].notnull(),"Incarico"].tolist()
pass
#listfilter_ODAG=dffilter["ODAG"].tolist()


#call function from main module
#if __name__ == "__main__":
#	wlist=[]#"CITIMP","TAM"
	#"AllDip",
#	wDict={
#"AllDip"
#"parm1":"value1",
#"parm2":"value2",
#"parm3":"value3"
#}
#	Util_readParms(wlist,wDict)
#	Util_readParms()
#	pass