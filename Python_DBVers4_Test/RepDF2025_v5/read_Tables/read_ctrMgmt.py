#Get file from Contract Management
import glob
import os
from Util_leggeDir import dir_dict
import pandas as pd
import numpy as np #per sostituire i valori NA	from Funz.Util_wDir import RepInit0	
	#import datetime   #importare libreria per la data del giorno
from datetime import timedelta,datetime  
from read_Interni import df_RisInt_byCID, df_RisInt_byName
#from read_rilasci import df_Rilasci_task
import sys

##Read working directories 
#new_dir = os.getcwd()
wrk_ctrMgmt_inp=dir_dict["ctrMgmt"]
wrk_stagDir=dir_dict["stagDir"]

wrk_ctrMgmtODAG_inp=wrk_ctrMgmt_inp+r"\Monitoraggio ODAG\avanzamento macroclassi - 2025 settimanale"
wrk_ctrMgmtODA_inp=wrk_ctrMgmt_inp+r"\Monitoraggio ODA\avanzamento ODA - 2025 settimanale"
#Get ODAG File
list_of_files = glob.glob(wrk_ctrMgmtODAG_inp+r"/*.xlsx") # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getmtime)

df_ODAG_ctrMgmt=pd.read_excel(latest_file, sheet_name=0, dtype={"Numero Contratto Gara": str,"Documento d'acquisto": str,"Richiesta d'acquisto":str,"Data Creazione Bdo":str,	"Data Approvazione":str,	"Data decorrenza":str,	"Data scadenza":str,"Tipologia Fornitura":str,"Stato Testata Ordine":str}, decimal=',')
df_ODAG_ctrMgmt.rename(columns={df_ODAG_ctrMgmt.columns[0]: 'ODAG',df_ODAG_ctrMgmt.columns[1]: 'ODAGnbr', df_ODAG_ctrMgmt.columns[2]: 'CIG', df_ODAG_ctrMgmt.columns[3]: 'ODAG_desc'},inplace=True)
#replace column name if contains a value
df_ODAG_ctrMgmt.columns= [col if not 'DEC' in col else 'DEC' for col in df_ODAG_ctrMgmt.columns]
#df_ODAG_ctrMgmt = df_ODAG_ctrMgmt.rename(columns=lambda c: "DEC" if "DEC" in c else c)

df_ODAG_ctrMgmt=df_ODAG_ctrMgmt[["ODAG","ODAGnbr","CIG","ODAG_desc","CTRM","DEC","RUP"]].astype(str)

df_ODAG_ctrMgmt=pd.merge(df_ODAG_ctrMgmt,df_RisInt_byName[["RisInt_nome","Int_Dip"]],left_on="DEC", right_on="RisInt_nome",how="left",suffixes=('', '_dec'))
df_ODAG_ctrMgmt=pd.merge(df_ODAG_ctrMgmt,df_RisInt_byName[["RisInt_nome","Int_Dip"]],left_on="RUP", right_on="RisInt_nome",how="left",suffixes=('', '_rup'))
df_ODAG_ctrMgmt=df_ODAG_ctrMgmt.rename(columns={"Int_Dip":"DEC_Dip","Int_Dip_rup":"RUP_Dip"})
#delete all columns begin with "Cogno"
df_ODAG_ctrMgmt= df_ODAG_ctrMgmt.loc[:,~df_ODAG_ctrMgmt.columns.str.startswith('Cogno')]

df_ODAG_ctrMgmt.to_excel(wrk_stagDir+"\\ctrmgmt\\"+"ODAG_ctrMgmt.xlsx", sheet_name="foglio1", index=False, startrow=0, startcol=0)	

#Get ODA File
list_of_files = glob.glob(wrk_ctrMgmtODA_inp+r"/*.xlsx") # * means all if need specific format then *.csv
latest_file = max(list_of_files, key=os.path.getmtime)

df_ODA_ctrMgmt=pd.read_excel(latest_file, sheet_name=0, dtype={"Numero Contratto Gara": str,"Documento d'acquisto": str,"Richiesta d'acquisto":str,"Data Creazione Bdo":str,	"Data Approvazione":str,	"Data decorrenza":str,	"Data scadenza":str,"Tipologia Fornitura":str,"Stato Testata Ordine":str}, decimal=',')
df_ODA_ctrMgmt.rename(columns={df_ODA_ctrMgmt.columns[0]: 'ODA', df_ODA_ctrMgmt.columns[1]: 'CIG', df_ODA_ctrMgmt.columns[2]: 'ODA_desc'},inplace=True)
#replace column name if contains a value
df_ODA_ctrMgmt.columns= [col if not 'DEC' in col else 'DEC' for col in df_ODA_ctrMgmt.columns]

df_ODA_ctrMgmt=df_ODA_ctrMgmt[["ODA","CIG","ODA_desc","DEC","RUP"]].astype(str)

df_ODA_ctrMgmt=pd.merge(df_ODA_ctrMgmt,df_RisInt_byName[["RisInt_nome","Int_Dip"]],left_on="DEC", right_on="RisInt_nome",how="left",suffixes=('', '_dec'))
df_ODA_ctrMgmt=pd.merge(df_ODA_ctrMgmt,df_RisInt_byName[["RisInt_nome","Int_Dip"]],left_on="RUP", right_on="RisInt_nome",how="left",suffixes=('', '_rup'))
df_ODA_ctrMgmt=df_ODA_ctrMgmt.rename(columns={"Int_Dip":"DEC_Dip","Int_Dip_rup":"RUP_Dip"})

#delete all columns begin with "Cogno"
df_ODA_ctrMgmt= df_ODA_ctrMgmt.loc[:,~df_ODA_ctrMgmt.columns.str.startswith('Cogno')]

df_ODA_ctrMgmt.to_excel(wrk_stagDir+"\\ctrmgmt\\"+"ODA_ctrMgmt.xlsx", sheet_name="foglio1", index=False, startrow=0, startcol=0)	
df_ODA_ctrMgmt=df_ODA_ctrMgmt["ODA"].astype(int)
df_ODAG_ctrMgmt=df_ODAG_ctrMgmt["ODAGnbr"].astype(int)
pass